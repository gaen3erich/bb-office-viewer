# Office Viewer for BB

A fork of [bb-plugin-office-viewer](https://github.com/VKirill/bb-plugin-office-viewer) with two additions that make it usable on **Windows** and let you **edit** Word documents:

1. **Windows path fix** — the upstream crashed on Windows paths (`C:\...`) because it looked for `/` as the path separator. The fork resolves the document directory and file name with `path.dirname()` / `path.basename()` instead of `lastIndexOf("/")`.
2. **Word (.docx) editing** — the **Edit document** button enters a seamless, in-place (WYSIWYG) editing mode over the rendered pages, and a structured paragraph-list view is available too. Saving rewrites only `word/document.xml`; tables, images, styles and numbering are preserved.

## What it opens

`.docx`, `.docm`, `.dotx`, `.xlsx`, `.xls`, `.ods`, `.csv`, `.pptx`, `.pdf`, images, video, audio and HTML — with search, zoom, quoting into chat, and saving without losing formatting.

## Install on a new machine

Requires [BB](https://getbb.app) on the machine. On PowerShell:

```powershell
# Install everything
./setup.ps1
bb plugin list
```

or install it directly:

```powershell
bb plugin install git:https://github.com/gaen3erich/bb-office-viewer.git --yes
bb plugin list
```

## Development

```powershell
npm install
bb plugin build .
bb plugin reload office-viewer
```

## Notes

- This repository contains the plugin at its root; `node_modules/` and `dist/` are ignored in git, so BB builds them on install.
- Updates from the upstream marketplace will not arrive automatically. Pull the upstream changes into this fork, rebuild (`bb plugin build .`), and commit.
