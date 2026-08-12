# Consuming this fork from GitHub Packages

This fork publishes its own build of every workspace package to **GitHub
Packages** (`npm.pkg.github.com`) so other repositories in the same
organization can install it without waiting for an upstream npm release.

This is the only publishing path in this fork. Upstream's release workflow
(changesets → npmjs.org) is not carried here, so nothing publishes
`@eigenpal/*` from this repo and nothing consumes the changeset queue.

## What gets published, and why the name changes

GitHub Packages only accepts a package whose **scope matches the account that
owns the repository**. This repo is owned by `anylawyer`, so the packages are
republished under that scope:

| npmjs.org (upstream)           | GitHub Packages (this fork)     |
| ------------------------------ | ------------------------------- |
| `@eigenpal/docx-editor-core`   | `@anylawyer/docx-editor-core`   |
| `@eigenpal/docx-editor-react`  | `@anylawyer/docx-editor-react`  |
| `@eigenpal/docx-editor-vue`    | `@anylawyer/docx-editor-vue`    |
| `@eigenpal/docx-editor-agents` | `@anylawyer/docx-editor-agents` |
| `@eigenpal/docx-editor-i18n`   | `@anylawyer/docx-editor-i18n`   |
| `@eigenpal/nuxt-docx-editor`   | `@anylawyer/nuxt-docx-editor`   |

The rename is applied to the built `dist/` as well as the manifests, because
the adapters import core by package name. Internal dependencies are pinned to
the **exact** version published in the same run, so a consumer can never end up
with this fork's React adapter talking to upstream's core.

Import paths change accordingly:

```ts
import { DocxEditor } from '@anylawyer/docx-editor-react';
```

## Authentication (required — even though the packages are public)

The GitHub Packages npm registry rejects anonymous reads. Every consumer needs
a token, in CI and on developer machines alike. This is the step that trips
people up.

### In the consuming repository

Add an `.npmrc` at its root (safe to commit — it contains no secret):

```
@anylawyer:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=${NODE_AUTH_TOKEN}
```

### For local development

Create a **classic** personal access token with the `read:packages` scope and
put it in `~/.npmrc` (not in the repo):

```
//npm.pkg.github.com/:_authToken=ghp_yourTokenHere
```

Then `npm install` resolves `@anylawyer/*` normally.

### For GitHub Actions in the consuming repository

The consuming repo's own `GITHUB_TOKEN` works, but only once this package
grants that repo access:

1. Open the package page (repo → **Packages** → the package)
2. **Package settings** → **Manage Actions access**
3. **Add repository** → pick the consuming repo → **Read**

Then in that repo's workflow:

```yaml
- uses: actions/setup-node@v6
  with:
    node-version: '24'
    registry-url: 'https://npm.pkg.github.com'
    scope: '@anylawyer'
- run: npm ci
  env:
    NODE_AUTH_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

Alternatively, skip the per-package access grant and use an organization secret
holding a PAT with `read:packages`, passing it as `NODE_AUTH_TOKEN` instead.

## Installing

```bash
npm install @anylawyer/docx-editor-react@latest
```

Peer dependencies (`react`, `react-dom`, and the `prosemirror-*` set) are the
consumer's to install. Match the pinned ProseMirror versions this repo uses —
see `overrides` / `resolutions` in the root `package.json`; multiple copies of
`prosemirror-model` or `prosemirror-view` break the editor at runtime.

### Choosing a version

The `latest` dist-tag always points at the newest fork build. Automatic
publishes are **prereleases** (`1.9.0-fork.42`), so a caret range like
`^1.9.0` will _not_ match them — pin exactly, or track the tag:

```jsonc
{
  "dependencies": {
    // pinned to one reviewed build (recommended for apps)
    "@anylawyer/docx-editor-react": "1.9.0-fork.42",
  },
}
```

## Publishing a new version

Two ways, both in `.github/workflows/publish-github-packages.yml`:

- **Automatic** — every push to `main` publishes
  `<current version>-fork.<run number>` and moves the `latest` tag to it.
- **Manual** — Actions → _Publish to GitHub Packages_ → **Run workflow**. Leave
  `version` blank for the same auto-numbering, or pass an exact version (for
  example `1.9.0-anylawyer.1`) to cut a build you intend to pin. `dry_run`
  builds and packs without uploading.

GitHub Packages is immutable: a version can never be republished, which is why
every run needs a fresh version number.

The workflow gates on typecheck and the unit suite before publishing. Playwright
and the parity gate are deliberately left to `ci.yml` on pull requests, to keep
publishes fast.

## Troubleshooting

| Symptom                                                     | Cause                                                                         |
| ----------------------------------------------------------- | ----------------------------------------------------------------------------- |
| `401 Unauthorized` on install                               | No token for `npm.pkg.github.com`, or the PAT lacks `read:packages`           |
| `404 Not Found` for `@anylawyer/...`                        | `.npmrc` scope mapping missing, so npm looked on npmjs.org                    |
| `403` during publish: scope does not match repository owner | The package scope must equal the repo owner; the publish script handles this  |
| `EPUBLISHCONFLICT` / "cannot publish over existing version" | That version already exists — bump it; GitHub Packages versions are immutable |
| Editor crashes with duplicate ProseMirror state             | Two copies of a `prosemirror-*` package — dedupe with overrides               |
