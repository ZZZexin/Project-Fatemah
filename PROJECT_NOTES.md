# Project Fatemah — Design Notes

## What This Is
WellCAD TV (televiewer) data processing automation.
Replaces manual step-by-step VBScript execution with a structured pipeline.

---

## Two Functional Modules

### Module 1 — Batch Convert (around vendor app)
- **Input:** proprietary raw OTV/ATV files (LOX or similar format from tool vendor)
- **Process:** batch wrapper that shells out to the vendor's OTV conversion application
- **Output:** HED / LGX / LAS files ready for WellCAD import
- **Status:** NOT YET BUILT — blocked on vendor app details (see open questions)

### Module 2 — WellCAD Processing Pipeline
- **Input:** HED / LGX / LAS from Module 1 (or existing files)
- **Process:** import into WellCAD → extend/slice logs → orientation (HS + TN/MN) → apply templates
- **Output:** final _hs_ / _tn_ / _mn_ WCL deliverable files
- **Status:** SCRIPTS REFACTORED — 001, 003, 004 modernised, in this repo

---

## File Structure

```
Project_Fatemah/
├── pipeline.vbs            ← orchestrator (NOT YET BUILT)
├── config.ini              ← active client config (NOT YET BUILT)
├── _lib.vbs                ← shared helpers (PARTIALLY BUILT)
├── pipeline.log            ← runtime output (generated on run)
│
├── clients/                ← one .ini per client (NOT YET BUILT)
│   ├── RTIO.ini
│   └── ClientB.ini
│
├── 001_wcRTIOtv_Las_TV_import.vbs        ← Module 2 step 1 (refactored)
├── 003_wcRTIOtv_extend and slice.vbs     ← Module 2 step 2 (refactored)
├── 004__wcRTIOtv_any_orientation.vbs     ← Module 2 step 3 (refactored)
│
└── (future) ui/
    └── launcher.py                       ← Phase 2 UI (tkinter)
```

---

## Client Config Design (clients/RTIO.ini)

```ini
[client]
name             = RTIO
company_display  = RIO TINTO IRON ORE
country          = AUSTRALIA

[paths]
template_path    = C:\Proc_TV\05_RTIO_TV_LasPrep\templates
; root_path is auto-detected from script location

[module1]
vendor_app       = (TBD — need exe path and CLI flags)
input_format     = lox
otv_output_format = hed

[module2]
otv_import_mode  = hed          ; bmp or hed
devi_preference  = auto         ; auto | nsgazi | gazi | azi
apply_template   = true

[output]
suffix_hs = hs
suffix_tn = tn
suffix_mn = mn
```

---

## Language Decision — OPEN

**VBScript path (current):** no new dependencies, but limited error handling and no good UI path.

**Python approach — two libraries, two jobs:**

| Library | How it works | Used for |
|---|---|---|
| `pywin32` (win32com.client) | Calls WellCAD's COM object directly in code | Module 2 — replaces VBScript entirely, same API |
| `pywinauto` | Drives any Windows app by clicking its actual UI | Module 1 — automates vendor OTV app if it has no CLI |

- Both work on WellCAD v5.2. The "Python in v5.7" is WellCAD's *built-in* Python console — irrelevant here.
- `pip install pywin32 pywinauto` — no other dependencies.
- pywin32: fast and reliable (API-level). pywinauto: slower and fragile if vendor updates UI, but works on any app.
- Recommended: **single `pipeline.py`** — pywinauto for Module 1, pywin32 for Module 2, tkinter for Phase 2 UI.

Decision needed before building pipeline orchestrator.

---

## Open Questions

1. **Module 1 vendor app:** What is the executable name? Does it have a CLI (command-line flags) or is it GUI-only?
2. **Language:** VBScript pipeline or Python + pywin32?
3. **Client list:** Which clients need different configs? What specifically differs per client?
4. **Module independence:** Can Module 2 run against a pre-existing WCL, or does Module 1 always run first?

---

## What Was Done So Far

- Investigated all 3 original VBS scripts for issues
- Refactored 001, 003, 004:
  - Removed ~1500 lines of duplicated code
  - Extracted shared functions (NonZeroMin, Max, InterpolateLogEnd, GetDeviationType,
    RotateImages, DeleteLogsByLookup, SliceAllLogs, etc.)
  - Fixed ATV LAS scan path bug (was scanning wrong folder)
  - Replaced nested If chains with Select Case
  - Removed dead commented-out code
- Created `_lib.vbs` shared library (partial)
- Created `C:\Project_Fatemah` folder with git remote → https://github.com/ZZZexin/Project-Fatemah.git

---

## Next Steps (in order)

1. Decide: VBScript pipeline vs Python + pywin32
2. Answer open questions above (vendor app, client list)
3. Build `clients/RTIO.ini` config structure
4. Build `pipeline.vbs` (or `pipeline.py`) orchestrator
5. Build Module 1 batch converter once vendor app details known
6. Phase 2: simple UI (HTA or tkinter)
