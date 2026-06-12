# Project Fatemah — Design Notes

## What This Is
WellCAD TV (televiewer) data processing automation.
Replaces manual step-by-step VBScript execution with a structured Python pipeline,
ending with a GUI launcher usable by field staff.

---

## Three Functional Areas

### Module 1 — Batch Convert (around vendor OTV app)
- **Input:** proprietary raw OTV/ATV files (LOX or similar) from the tool vendor
- **Process:** pywinauto drives the vendor conversion app to batch-process files
- **Output:** HED / LGX / LAS files ready for WellCAD import
- **Status:** NOT YET BUILT — blocked on vendor app details (see open questions)

### Module 2 — WellCAD Processing Pipeline
- **Input:** HED / LGX / LAS from Module 1, or dropped in manually
- **Process:** import into WellCAD → extend/slice logs → orientation (HS + TN/MN) → apply templates
- **Output:** final `_hs_` / `_tn_` / `_mn_` WCL deliverable files
- **Status:** VBS logic refactored and in repo — ready to port to Python

### Utility — Header Filler
- **Input:** open WCL in WellCAD + Well Report Excel in same folder
- **Process:** reads Excel Well Report, maps fields to WellCAD header
- **Output:** WellCAD header populated
- **Status:** standalone VBS exists (`modules/vb/Header Filling/`) — port to Python later

---

## Language Stack

| Library | Job | Where |
|---|---|---|
| `pywin32` (win32com.client) | Calls WellCAD COM API directly | Module 2, Header Filler |
| `pywinauto` | Drives vendor OTV app UI to batch convert | Module 1 |
| `tkinter` | GUI launcher — built-in Python, no install | Phase 4 |

- Works on WellCAD v5.2. COM is language-agnostic — no v5.7 requirement.
- x64: no issues. COM handles cross-bitness automatically (out-of-process server).
- Install: `pip install pywin32 pywinauto`

---

## Folder Structure

```
Project_Fatemah/
│
├── pipeline.py                        ← entry point — run this
├── PROJECT_NOTES.md                   ← this file
├── .gitignore                         ← excludes logs/, __pycache__/, *.pyc
│
├── config/
│   ├── default.ini                    ← fallback values for all clients
│   └── clients/
│       ├── RTIO.ini                   ← RTIO-specific overrides
│       └── ClientB.ini
│
├── modules/
│   ├── convert/                       ← Module 1: batch convert (pywinauto)
│   │   ├── batch_convert.py           ← loops input folder, drives vendor app per file
│   │   └── vendor_app.py             ← pywinauto wrapper: open→load→convert→close
│   │
│   ├── process/                       ← Module 2: WellCAD pipeline (pywin32)
│   │   ├── import_data.py             ← LAS + OTV/ATV image import  (from script 001)
│   │   ├── process.py                 ← extend, slice, devi logs     (from script 003)
│   │   └── orient_compile.py          ← HS + TN/MN, templates        (from script 004)
│   │
│   ├── header_filler/                 ← Utility: fill WellCAD header from Well Report
│   │   └── header_filler.py           ← port of universal_wr_reader.vbs
│   │
│   └── vb/                            ← legacy VBScript — reference only
│       ├── _lib.vbs
│       ├── 001_wcRTIOtv_Las_TV_import.vbs
│       ├── 003_wcRTIOtv_extend and slice_v2.08.vbs
│       ├── 004__wcRTIOtv_any_orientation_RTIO_STYLE_V-0.0.vbs
│       └── Header Filling/
│           ├── universal_wr_reader.vbs
│           └── universal_wr_reader_interp.vbs
│
├── lib/                               ← shared utilities across all modules
│   ├── wellcad_helpers.py             ← log_exists(), get_devi_type(), rotate_images() …
│   ├── las_parser.py                  ← read ~Well and ~Params from LAS files
│   └── logger.py                      ← timestamped log to file + console
│
├── templates/                         ← version-controlled WellCAD templates
│   ├── shared/                        ← used by all clients
│   │   ├── AutoBulkLoad.ini
│   │   ├── HEDImport.ini
│   │   ├── BMPImport.ini
│   │   └── ConvertLogTo.ini
│   └── clients/
│       ├── RTIO/                      ← RTIO .wdt and .ini files
│       │   ├── GEOPHYSICS IMPORTd2.wdt
│       │   ├── RTIO_2_BHTV_OPTV_nsg.wdt
│       │   ├── RTIO_2_BHTV_nsg.wdt
│       │   ├── RTIO_2_OPTV_nsg.wdt
│       │   ├── Template PWS - NSGAZI3.wdt
│       │   ├── Template PWS - GAZI3.wdt
│       │   ├── Template PWS - AZI3.wdt
│       │   ├── Template PWS - NOLASDEV.wdt
│       │   ├── MN_RotateConfig.ini
│       │   ├── nsgAZI_RotateConfig.ini
│       │   ├── gAZI_RotateConfig.ini
│       │   ├── AZI_RotateConfig.ini
│       │   ├── AZI_OBI_RotateConfig.ini
│       │   ├── AZI_ABI_RotateConfig.ini
│       │   ├── NormaliseImage_1D.ini
│       │   ├── NormaliseImage_Static.ini
│       │   ├── PWS_Lookup_DeleteTheseColumnsList_01.ini
│       │   └── PWS_Lookup_DeleteForHIGHSIDE_List.ini
│       └── ClientB/
│
├── logs/                              ← runtime output — gitignored
│
└── ui/                                ← Phase 4
    └── launcher.py                    ← tkinter: client picker, module toggles, log view
```

---

## Client Config (config/clients/RTIO.ini)

```ini
[client]
name             = RTIO
company_display  = RIO TINTO IRON ORE
country          = AUSTRALIA

[paths]
template_path        = templates/clients/RTIO
shared_template_path = templates/shared

[module1]
vendor_app        = (TBD — exe path and CLI/GUI mode)
input_format      = lox
otv_output_format = hed

[module2]
otv_import_mode  = hed          ; bmp | hed
devi_preference  = auto         ; auto | nsgazi | gazi | azi
apply_template   = true

[output]
suffix_hs = hs
suffix_tn = tn
suffix_mn = mn
```

---

## Build Plan

### Phase 1 — Foundation
- [ ] Step 1: Finalise and commit folder structure
- [ ] Step 2: `lib/logger.py`
- [ ] Step 3: `lib/las_parser.py`
- [ ] Step 4: `lib/wellcad_helpers.py`
- [ ] Step 5: `config/default.ini` + `config/clients/RTIO.ini`

### Phase 2 — Module 2 (WellCAD pipeline)
- [ ] Step 6: `modules/process/import_data.py`
- [ ] Step 7: `modules/process/process.py`
- [ ] Step 8: `modules/process/orient_compile.py`
- [ ] Step 9: `pipeline.py` — wire up and test Module 2 end to end

### Phase 3 — Module 1 (batch convert)
- [ ] Step 10: `modules/convert/vendor_app.py`  *(blocked: need vendor app details)*
- [ ] Step 11: `modules/convert/batch_convert.py`
- [ ] Step 12: Wire Module 1 into `pipeline.py`

### Phase 4 — GUI
- [ ] Step 13: `ui/launcher.py` — client selector, module toggles, run button, log viewer

---

## Open Questions

1. **Module 1 vendor app:** exe name? CLI flags available, or GUI-only?
2. **Client list:** which other clients, and what differs (templates, header mappings, formats)?

---

## History

| Date | What |
|---|---|
| 2026-06-12 | VBS scripts 001, 003, 004 reviewed and refactored |
| 2026-06-12 | Project structure, language stack, and build plan designed |
| 2026-06-12 | Folder structure scaffolded, PROJECT_NOTES.md committed to GitHub |
