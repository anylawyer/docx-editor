/**
 * Republish the workspace packages to GitHub Packages (npm.pkg.github.com)
 * under the owning organization's scope.
 *
 * Why a rename step exists
 * -----------------------
 * GitHub Packages only accepts an npm package whose scope matches the account
 * that owns the repository. This repo is `anylawyer/docx-editor` while the
 * packages are published to npmjs.org as `@eigenpal/*`, so a straight publish
 * is rejected. This script rewrites the whole graph to `@<owner>/*` — the
 * manifests AND the built `dist/`, because the adapters import core by package
 * name (`import ... from '@eigenpal/docx-editor-core'`), so renaming only the
 * manifests would ship packages whose imports resolve to the upstream npmjs
 * build instead of this fork's.
 *
 * Internal dependency ranges are pinned to the exact published version so a
 * consumer always gets one coherent build of the set, never a mix of a fork
 * adapter with an upstream core.
 *
 * The rewrite is destructive and meant for a throwaway CI checkout. To undo it
 * locally, restore only the tracked package.json manifests (a `git restore`
 * limited to those six paths) and re-run `bun run build:packages`. Do NOT run
 * `git checkout -- packages`: that also discards source edits in the working
 * tree.
 *
 * Usage:
 *   node scripts/publish-github-packages.mjs \
 *     --scope anylawyer --version 1.9.0-fork.42 \
 *     [--repo owner/name] [--tag latest] [--publish] [--dry-run]
 *
 * Without `--publish` it only rewrites and reports (useful to inspect a diff).
 * `--dry-run` implies `--publish` semantics but passes `--dry-run` to npm, so
 * npm packs the tarball and reports contents without uploading.
 */

import { readdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join, dirname, extname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const PACKAGES_DIR = join(ROOT, 'packages');
const GPR_REGISTRY = 'https://npm.pkg.github.com';

/** Files whose contents may carry a package specifier. Anything else is copied through. */
const TEXT_EXTENSIONS = new Set([
  '.js',
  '.mjs',
  '.cjs',
  '.jsx',
  '.ts',
  '.mts',
  '.cts',
  '.tsx',
  '.json',
  '.map',
  '.css',
  '.vue',
  '.html',
  '.txt',
  '.md',
]);

/** Manifest fields whose keys are package names a consumer will resolve. */
const DEPENDENCY_FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies'];

function parseArgs(argv) {
  const args = { tag: 'latest', publish: false, dryRun: false };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    const next = () => {
      const value = argv[++i];
      if (value === undefined) throw new Error(`Missing value for ${arg}`);
      return value;
    };
    if (arg === '--scope') args.scope = next();
    else if (arg === '--version') args.version = next();
    else if (arg === '--tag') args.tag = next();
    else if (arg === '--repo') args.repo = next();
    else if (arg === '--publish') args.publish = true;
    else if (arg === '--dry-run') args.dryRun = true;
    else throw new Error(`Unknown argument: ${arg}`);
  }
  if (!args.scope)
    throw new Error('--scope is required (the GitHub org / user that owns the repo)');
  if (!args.version) throw new Error('--version is required');
  // npm scopes are lowercase; GitHub logins can carry capitals.
  args.scope = args.scope.toLowerCase().replace(/^@/, '');
  return args;
}

function readJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'));
}

function writeJson(path, value) {
  writeFileSync(path, `${JSON.stringify(value, null, 2)}\n`);
}

/** Every publishable workspace package, as { dir, manifestPath, manifest }. */
function collectPackages() {
  return readdirSync(PACKAGES_DIR, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => join(PACKAGES_DIR, entry.name))
    .filter((dir) => existsSync(join(dir, 'package.json')))
    .map((dir) => {
      const manifestPath = join(dir, 'package.json');
      return { dir, manifestPath, manifest: readJson(manifestPath) };
    })
    .filter((pkg) => pkg.manifest.private !== true && typeof pkg.manifest.name === 'string');
}

/**
 * The single scope every workspace package shares. Bailing on a mixed set is
 * deliberate: the dist rewrite is a scope-prefix replace, which is only
 * complete if there is exactly one source scope to replace.
 */
function resolveSourceScope(packages) {
  const scopes = new Set(
    packages.map((pkg) => {
      const match = /^@([^/]+)\//.exec(pkg.manifest.name);
      if (!match) throw new Error(`Package ${pkg.manifest.name} is unscoped; cannot republish it.`);
      return match[1];
    })
  );
  if (scopes.size !== 1) {
    throw new Error(`Expected one workspace scope, found: ${[...scopes].join(', ')}`);
  }
  return [...scopes][0];
}

function* walkFiles(dir) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) yield* walkFiles(full);
    else if (entry.isFile()) yield full;
  }
}

/**
 * Rewrite `@sourceScope/` → `@targetScope/` through a package's build output.
 * Covers runtime imports, `.d.ts` type imports, sourcemaps and any CSS that
 * references a sibling package by name.
 */
function rewriteDist(pkgDir, sourceScope, targetScope) {
  const distDir = join(pkgDir, 'dist');
  if (!existsSync(distDir)) {
    throw new Error(`${pkgDir} has no dist/ — run \`bun run build:packages\` before publishing.`);
  }
  const from = `@${sourceScope}/`;
  const to = `@${targetScope}/`;
  let filesChanged = 0;
  for (const file of walkFiles(distDir)) {
    if (!TEXT_EXTENSIONS.has(extname(file))) continue;
    const original = readFileSync(file, 'utf8');
    if (!original.includes(from)) continue;
    writeFileSync(file, original.split(from).join(to));
    filesChanged++;
  }
  return filesChanged;
}

/**
 * Point the manifest at the target scope: rename the package, pin every
 * internal dependency to the exact version being published, and aim publishes
 * at GitHub Packages. `devDependencies` are left alone — npm ships them in the
 * manifest but never installs them for a consumer.
 */
function rewriteManifest({ manifest, manifestPath }, { sourceScope, targetScope, version, repo }) {
  const rename = (name) =>
    name.startsWith(`@${sourceScope}/`)
      ? `@${targetScope}/${name.slice(sourceScope.length + 2)}`
      : name;

  const previousName = manifest.name;
  manifest.name = rename(manifest.name);
  manifest.version = version;

  const internalDeps = [];
  for (const field of DEPENDENCY_FIELDS) {
    const deps = manifest[field];
    if (!deps) continue;
    manifest[field] = Object.fromEntries(
      Object.entries(deps).map(([name, range]) => {
        if (!name.startsWith(`@${sourceScope}/`)) return [name, range];
        internalDeps.push(rename(name));
        return [rename(name), version];
      })
    );
  }
  if (manifest.peerDependenciesMeta) {
    manifest.peerDependenciesMeta = Object.fromEntries(
      Object.entries(manifest.peerDependenciesMeta).map(([name, meta]) => [rename(name), meta])
    );
  }

  manifest.publishConfig = { ...manifest.publishConfig, registry: GPR_REGISTRY };
  if (repo) {
    // GitHub Packages links a package to a repository through this field;
    // pointing it at upstream would attach the publish to the wrong repo.
    manifest.repository = { type: 'git', url: `git+https://github.com/${repo}.git` };
  }

  writeJson(manifestPath, manifest);
  return { previousName, name: manifest.name, internalDeps };
}

/**
 * Order packages so a dependency is always published before its dependents.
 * If a run dies partway through, everything already uploaded is installable —
 * no published package references a version that never made it.
 */
function orderByDependencies(packages) {
  const remaining = [...packages];
  const emitted = new Set();
  const ordered = [];
  while (remaining.length > 0) {
    const index = remaining.findIndex((pkg) =>
      (pkg.internalDeps ?? []).every((dep) => emitted.has(dep))
    );
    // A cycle shouldn't happen in this workspace; publishing in the original
    // order is a better failure mode than looping forever.
    const next = index === -1 ? remaining[0] : remaining[index];
    remaining.splice(index === -1 ? 0 : index, 1);
    emitted.add(next.manifest.name);
    ordered.push(next);
  }
  return ordered;
}

/** True if this exact version is already on the registry. */
function versionExists(name, version) {
  try {
    execFileSync('npm', ['view', `${name}@${version}`, 'version'], {
      stdio: ['ignore', 'pipe', 'ignore'],
    });
    return true;
  } catch {
    return false;
  }
}

function publishPackage(pkgDir, name, version, { tag, dryRun }) {
  const args = ['publish', '--tag', tag];
  if (dryRun) args.push('--dry-run');
  try {
    execFileSync('npm', args, { cwd: pkgDir, stdio: 'inherit' });
    return 'published';
  } catch (error) {
    // GitHub Packages versions are immutable, so a retry after a run that died
    // partway through will hit an already-uploaded package. Treat that as done
    // rather than failing the whole publish.
    if (!dryRun && versionExists(name, version)) return 'already published';
    throw error;
  }
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const packages = collectPackages();
  if (packages.length === 0) throw new Error('No publishable packages found under packages/.');

  const sourceScope = resolveSourceScope(packages);
  const targetScope = args.scope;

  console.log(`Republishing ${packages.length} packages`);
  console.log(`  @${sourceScope}/*  →  @${targetScope}/*`);
  console.log(`  version:  ${args.version}`);
  console.log(`  dist-tag: ${args.tag}`);
  console.log(`  registry: ${GPR_REGISTRY}`);
  if (args.dryRun) console.log('  mode:     dry run (npm packs but does not upload)');
  else if (!args.publish) console.log('  mode:     rewrite only (no npm publish)');
  console.log('');

  for (const pkg of packages) {
    const filesChanged =
      sourceScope === targetScope ? 0 : rewriteDist(pkg.dir, sourceScope, targetScope);
    const { previousName, name, internalDeps } = rewriteManifest(pkg, {
      sourceScope,
      targetScope,
      version: args.version,
      repo: args.repo,
    });
    pkg.internalDeps = internalDeps;
    const deps = internalDeps.length > 0 ? ` deps→ ${internalDeps.join(', ')}` : '';
    console.log(`  ${previousName} → ${name} (${filesChanged} dist files rewritten)${deps}`);
  }
  console.log('');

  const ordered = orderByDependencies(packages);

  if (!args.publish && !args.dryRun) {
    console.log(
      `Rewrite complete. Publish order: ${ordered.map((p) => p.manifest.name).join(', ')}`
    );
    console.log('Re-run with --publish to upload.');
    return;
  }

  for (const pkg of ordered) {
    console.log(`\n── ${pkg.manifest.name}@${args.version}`);
    const outcome = publishPackage(pkg.dir, pkg.manifest.name, args.version, {
      tag: args.tag,
      dryRun: args.dryRun,
    });
    if (outcome === 'already published') console.log(`   already on the registry — skipped`);
  }

  console.log(`\nDone. ${ordered.length} packages ${args.dryRun ? 'packed' : 'published'}.`);
}

main();
