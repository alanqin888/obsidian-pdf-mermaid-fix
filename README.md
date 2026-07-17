# Obsidian PDF Mermaid Fix

An ultra-lightweight Obsidian plugin that fixes the annoying issue where wide Mermaid diagrams (like sequence diagrams) get horizontally truncated/cut off when using Obsidian's native "Export to PDF" feature.

## How it works

When you use the `Export to PDF (Mermaid Fix)` command, this plugin temporarily injects a precise `@media print` CSS rule that forces all Mermaid SVG elements to scale responsively (`max-width: 100%`) and avoid page breaks, and then automatically invokes Obsidian's native, high-quality PDF export dialog.

This means you get the best of both worlds:
- Perfect Mermaid diagrams that fit within your A4 (or other) page size.
- The familiar, feature-rich native Obsidian PDF export dialog.

## How to use

1. Open any Markdown file containing wide Mermaid diagrams.
2. Right-click anywhere in the file explorer on the file, or right click inside the editor.
3. Select **"Export to PDF (Mermaid Fix)"**.
4. The native PDF export dialog will appear. Click Export and enjoy perfectly scaled diagrams!

## Manual Installation

1. Create a folder named `obsidian-pdf-mermaid-fix` inside your vault's `.obsidian/plugins/` directory.
2. Download `main.js` and `manifest.json` from the latest Release and place them in the folder.
3. Reload Obsidian and enable the plugin in Community Plugins.
