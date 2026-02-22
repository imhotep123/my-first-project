# CLAUDE.md — AI Assistant Guide for kiyaku_viewer

## Project Overview

This repository contains a single-file Python desktop application: **kiyaku_viewer.py** (規約比較ビューワー — "Terms/Agreement Comparison Viewer").

The application is a **Japanese legal document comparison tool** that:
- Loads two Word (`.docx`) documents containing legal terms/agreements (規約)
- Parses article numbers (条文番号) from each document
- Allows the user to look up articles side-by-side for comparison
- Supports bulk search (ESC key) across both panels simultaneously
- Exports displayed articles to `.txt` files

## File Structure

```
my-first-project/
├── kiyaku_viewer.py   # Entire application — single self-contained script
├── README.md          # Brief Japanese description
├── .gitignore         # Standard Python gitignore
└── CLAUDE.md          # This file
```

There is no build system, test suite, package manager config, or additional module structure. The entire application lives in `kiyaku_viewer.py`.

## Architecture

### Module Sections (in order within `kiyaku_viewer.py`)

| Lines | Section | Purpose |
|-------|---------|---------|
| 1–43 | Dependency check | Auto-installs `python-docx` via pip if missing |
| 45 | Import | `from docx import Document` after confirming availability |
| 52–76 | Number normalization | `normalize_num()` converts full-width/kanji digits → ASCII |
| 84–154 | Word parsing | `parse_docx()` extracts articles from `.docx` files |
| 157–166 | Article lookup | `find_article()` looks up articles by key with normalization fallback |
| 173–179 | Constants/styles | Font and color constants for the UI |
| 187–376 | `DocPanel` class | A single-panel widget (search bar + text display) |
| 383–543 | `CompareViewer` class | Main window: toolbar, split-pane layout, file I/O, export |
| 550–625 | `SearchDialog` class | Modal dialog for ESC-triggered unified search |
| 632–635 | Entry point | `if __name__ == "__main__"` block |

### Key Data Structures

- **`articles: dict[str, str]`** — Maps article key (e.g. `"第3条"`, `"第5条の2"`) to the article's full text content.
- **Article key format**: `第{N}条` or `第{N}条の{M}` where N and M are normalized Arabic numerals.

### Core Logic

**`parse_docx(path)`** (`kiyaku_viewer.py:98`):
- Reads all paragraphs and table cells from a `.docx` file
- Uses `_ART_RE` regex to detect article-number lines (`第X条`, `第X条のY`)
- Uses `_NAME_RE` regex to detect article-name lines (`（目的）`, `（定義）`, etc.)
- Article names that appear immediately before an article-number line are attributed to the following article, not the preceding one
- Returns `{key: text}` dict

**`normalize_num(s)`** (`kiyaku_viewer.py:59`):
- Converts full-width digits (`０–９`) to ASCII via translation table
- Converts kanji numerals (`一`, `二`, …, `千`) using positional arithmetic
- Used both at parse time and search time so queries match stored keys

**`find_article(articles, key)`** (`kiyaku_viewer.py:157`):
- First tries direct dict lookup
- Falls back to normalizing both the query and all stored keys for a tolerant match

## Running the Application

**Requirements:**
- Python 3.9+ (uses `from __future__ import annotations` and modern type hints)
- `python-docx` — the app auto-installs this on first run if missing
- A display (requires Tkinter + a graphical environment)

**Run:**
```bash
python kiyaku_viewer.py
```

On first run without `python-docx`, a dialog will prompt to install it automatically.

## Development Conventions

### Language
- **Application UI text**: Japanese only (labels, messages, dialogs)
- **Code comments**: Japanese (inline) and English (section headers)
- **Docstrings**: Japanese

### Code Style
- Python 3 with `from __future__ import annotations` for PEP 563 style hints
- Type hints used throughout (function signatures, instance variables)
- Section separators use the pattern: `# ─────── Description ───────`
- No external linter config present; follow PEP 8

### UI Framework
- **Tkinter** (`tk`) — standard library only (except `python-docx`)
- Layout uses `.pack()` and `.grid()` geometry managers
- Color constants defined at module level (`C_LEFT_H`, `C_RIGHT_H`, `BG_LEFT`, `BG_RIGHT`)
- Font constants: `F_BODY`, `F_BOLD`, `F_SMALL` — all use `"MS Gothic"` (Japanese monospace)

### Class Design
- `DocPanel(tk.Frame)` — encapsulates one side of the comparison view; can be used independently
- `CompareViewer` — orchestrates two `DocPanel` instances and the toolbar
- `SearchDialog(tk.Toplevel)` — modal dialog; stores result in `self.result` before destroying itself

### Regex Patterns (module-level constants)
- `_ART_RE` (`kiyaku_viewer.py:84`): Matches article-number lines; named groups `art` and `sub`
- `_NAME_RE` (`kiyaku_viewer.py:90`): Matches article-name lines (fully parenthesized lines)
- `_FULLWIDTH` (`kiyaku_viewer.py:52`): `str.maketrans` table for full-width digit normalization

## Key Behaviors to Preserve

1. **Article name attribution**: The `（条文名）` line immediately before `第X条` belongs to that article, not the previous one. The parsing logic removes it from the previous article's buffer (`kiyaku_viewer.py:131–139`).

2. **Tolerant number matching**: Users may input kanji, full-width, or ASCII article numbers — all normalize to the same key. Do not break `normalize_num()` or the fallback in `find_article()`.

3. **Auto-install prompt**: The `_ensure_docx()` function at module top runs before any other imports. If you restructure imports, ensure this runs first.

4. **ESC key binding**: `<Escape>` on the root window triggers `_unified_search()` (`kiyaku_viewer.py:390`), which opens `SearchDialog` and applies results to both panels.

5. **Export folder naming**: The export folder is named after the article key being displayed, with filesystem-unsafe characters replaced by `_` (`kiyaku_viewer.py:507`).

## No Tests / No Build Steps

There are no automated tests, no Makefile, no `setup.py`, `pyproject.toml`, or any build tooling. This is a standalone script intended to be run directly.

If adding tests, use `pytest` and mock `tkinter` and `docx.Document` appropriately since the app is GUI-based.

## Common Modification Areas

- **Add a new article pattern** (e.g., `附則`): Update `_ART_RE` and `_make_key()`
- **Change font or colors**: Modify the constants block (`kiyaku_viewer.py:173–179`)
- **Add a third comparison panel**: Extend `CompareViewer._build_ui()` to add a third `DocPanel` to the `PanedWindow`
- **Support `.txt` or `.pdf` input**: Add a new parsing function alongside `parse_docx()` and update the `filetypes` in `_load1`/`_load2`
- **Add keyboard shortcuts**: Bind to `self.root` in `CompareViewer.__init__()` or per-panel in `DocPanel._build()`
