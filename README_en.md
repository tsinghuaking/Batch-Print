English | [简体中文](README.md)

<div align="center">
  <img src="docs/logo.png" alt="Batch Print Tool" width="260">
</div>

<div align="center">

![License](https://img.shields.io/badge/License-MIT-green.svg)
![Python](https://img.shields.io/badge/Python-3.14-blue.svg)
![OS](https://img.shields.io/badge/OS-Windows-blue.svg)
[![Release](https://img.shields.io/github/v/release/tsinghuaking/Batch-Print?label=Release)](https://github.com/tsinghuaking/Batch-Print/releases)
[![Download](https://img.shields.io/github/downloads/tsinghuaking/Batch-Print/total?label=Downloads)](https://github.com/tsinghuaking/Batch-Print/releases)

</div>

## Introduction

**Batch Print Tool** is a small local utility for Windows that batch-prints **Word / Excel / PDF** files,
with per-file control over **duplex / simplex, copies, color / monochrome, and page ranges**.
Everything runs on your own machine — no file is ever uploaded to any network.

> Why this exists: Windows' built-in right-click "Print" cannot configure these details per file. This tool puts
> them all in one web page — pick your settings once and hit "Print".

Key features:

- 📄 **PDF / Word(.doc/.docx) / Excel(.xls/.xlsx) / PPT(.ppt/.pptx)** — Word/Excel/PPT print via local Microsoft Office directly, and PDF via the system reader's "right-click print". No intermediate conversion, fastest and most reliable.
- 🖨️ **Per-file settings**: duplex (long-edge / short-edge flip), copies, color / monochrome, page range (e.g. `1-3`, `1,3,5-7`)
- 📋 **Tabular task list**: auto-sorted by filename; columns for index, file, duplex, copies, color, page range, **real-time page count**, status
- 🎯 **Global + per-file control**: set defaults in the "Global settings" row and click "Apply to all", or tweak any single row
- ⚫ **Monochrome by default**: newly added files print in black & white to avoid wasting color ink
- 💾 **Remember the default printer**: click "Set as default" and it is auto-selected next time — no more dropdown hunting
- 🔄 **One-click reprint of failures**: failed files are marked red; the "Reprint failed files" button retries just those — no need to start over
- ✅ **Real-time status feedback**: each file turns green "Done" on success, or red "Failed" with details in the log
- 🎨 **Modern UI**: Neve-inspired design, Apple system font stack, custom rounded overlay dropdowns (no layout shift, no native browser dropdowns), soft palette, with a header logo and "Batch Print 批量打印工具" title
- 🔗 **Header shortcut buttons**: WeChat, Xiaohongshu (RED) and GitHub buttons on the top-right (official brand colors, uniform size) for contacting the author and following the project
- 🪟 **No black window, opens in your browser**: built only on the Python standard library (no third-party web framework); the
  printing is done by your local Office and the system PDF reader, so no extra engine is bundled — double-click opens the page in your system's default browser (no console window),
  and the "Exit program" button on the top-right stops the service.
- ⚡ **Faster prints**: calls your local Office and system PDF reader directly (same as right-click print) — no intermediate conversion, no engine cold start, fast and reliable paper output

## Preview

<div align="center">
  <img src="docs/preview.png" alt="UI preview">
</div>

## Quick Start (regular users, recommended)

1. Download **`BatchPrint.exe`** from [Releases](https://github.com/tsinghuaking/Batch-Print/releases)
   (single file, with the local print capability built in, ~40MB).
2. **Double-click** `BatchPrint.exe`:
   - It opens the print page in your system's default browser automatically (no black box);
   - If it doesn't open, visit `http://127.0.0.1:5001` manually.
3. When done, click "Exit program" on the top-right to close the window and stop the service.

> If port `5001` is already taken (a previous instance is running), the app simply opens your browser to the
> existing service instead of starting a second one.

## How to use

1. **Pick a printer**: choose a real printer from the top dropdown; click "Refresh" to re-scan; click "Set as default" for your usual one.
2. **Upload files**: click the dashed box or drag files in — multiple PDF / Word / Excel at once; files are auto-sorted by name and the table shows the real-time page count for each.
3. **Set options**:
   - Same for all: choose in the "Global settings" row, then click "Apply to all";
   - Per file: edit each row in the table;
   - Page syntax: `1-3` (pages 1–3), `1,3,5-7` (specific pages), empty = all pages.
4. **Print**: click "Print":
   - Each file is submitted straight to your local Office / system PDF reader (same as right-click print), with no intermediate conversion;
   - Each row updates live to "Done / Failed";
   - If any fail, click "Reprint failed files" to retry only those.
5. **Inspect the log**: the log panel below the page shows the printer, settings and return for each submission; failures include the error detail for easier debugging.

## Run from source (developers)

```bash
pip install -r requirements.txt
python app.py
# open http://127.0.0.1:5001 in your browser
```

## Build the exe

```bash
pip install pyinstaller
python build_exe.py
```

Output: `dist/批量打印工具.exe`. The script uses `--add-data` to embed `templates/`
into the exe, so distribution needs only that one file (self-contained, double-click to run).

> Note: GitHub's web upload caps attachments at 25MB. Publish via the CLI instead:
> `gh release create vX.Y "dist/批量打印工具.exe"` (uses the uploads API, supports up to 2GB).

## Project layout

```
批量打印工具/
├── app.py              # web service (Python stdlib http.server, no third-party deps)
├── print_core.py       # print core: enumerate printers / local Office direct print / system PDF reader
├── templates/
│   └── index.html      # web UI
├── requirements.txt
├── build_exe.py        # build script
├── LICENSE
└── .gitignore
```

## FAQ

**1. Nothing prints / it reports failure?**
- Make sure you selected a **real printer**. "Microsoft Print to PDF / WPS PDF / Adobe PDF" are virtual printers that
  pop a "Save As" dialog and don't suit silent batch printing.
- In Windows Settings → Printers, check whether the printer is "Paused" or "Offline".
- If a previous print job is stuck in the queue, clear the queue before retrying.

**2. The "WPS PDF" printer name shows garbage characters?**
That's a bad name WPS wrote into the system at install time (not a bug here), and it's a virtual printer — don't use it for real prints.

**3. What do I need to print Word / Excel / PPT?**
Just have Microsoft Office (Word/Excel/PowerPoint) installed locally. The tool prints through Office's built-in COM interface directly — no LibreOffice or any intermediate conversion; if Office is missing it falls back to LibreOffice to export PDF first. PDF printing relies on the system's associated PDF reader (e.g. Edge, Adobe Reader) — no extra install needed.

**4. Where do uploaded files go?**
They are temporarily stored in `%USERPROFILE%/.batch_print/uploads/` and are not auto-deleted after printing — clean up manually if needed.

**5. The printer does nothing / is super slow?**
- The first time you switch printers, the app enumerates all system printers (~1–2s) and caches the list for 120 seconds;
- The tool calls your local Office / system PDF reader directly (same as right-click print), so paper usually comes out fast; if nothing happens, make sure you selected a real printer and it is not "Paused / Offline";
- Virtual printers (Microsoft Print to PDF / WPS PDF / Adobe PDF) pop a "Save As" dialog and don't suit silent batch printing.

## Third-party licenses

- This tool uses only the Python standard library (no Flask or other web framework) for the web service; printing is done by the local Office and the system PDF reader, with no bundled third-party print engine.
- Released under the [MIT License](LICENSE).

## License

Released under the [MIT License](LICENSE).
