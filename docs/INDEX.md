# WSR Generator — Documentation

Master outline for the CES PFS CSAR (Non-STLA) Weekly Status Report automation.
We expand each section in installments; this file is the map.

**Quick start for Windows users:** see root [`README.md`](../README.md).

---

## 1. Problem & goal

**Context:** Every week the team builds a 12-slide WSR PowerPoint from Scrum Excel
(+ optional Planning Excel). Doing it by hand is slow and easy to get wrong.

**Goal:** One-click Windows app that reads the workbooks and produces the PPT,
with clear errors when Excel is wrong.

**Out of scope (for now):** SharePoint auto-upload, signed binaries, GitLab CI,
auto-updating the installed GUI.

---

## 2. Stakeholder story (what users do)

**Context:** Non-developers on Windows. No Python. Download `WSR-Generator.exe`,
pick Scrum file, optionally Planning file, click Generate, open the PPT.

**Will expand with:** screenshots, SmartScreen notes, where to put Excel files,
what “success” vs “warning” vs “error” looks like.

---

## 3. Inputs & outputs

**Context:**

| Input | Typical file | Role |
| --- | --- | --- |
| Scrum workbook | `SCRUM_PFS_….xlsm` | Charts, pending DCRs, DDP, tracker, visibility |
| Planning workbook | `Book2.xlsx` | Quarterly planning slide (optional) |
| Planned quarter % | GUI field (default 90) | Slide 11: planned hours = this % of available |
| PPT template | bundled in the app | Branding / layouts |

| Output | Role |
| --- | --- |
| `WSR_….pptx` | The weekly report |
| `WSR_….log` | Run log next to the PPT (especially on failure) |

**Will expand with:** required sheet names, required columns, sample paths.

---

## 4. Slide map (auto vs manual)

**Context:** The deck is always 12 slides. Some bodies are filled from Excel;
others are headers-only for humans to edit in PowerPoint.

| # | Slide | Automation |
| --- | --- | --- |
| 1 | Title | Auto |
| 2 | Agenda | Auto |
| 3 | MOM | Header only (manual) |
| 4 | DCR status (charts + summary) | Auto |
| 5 | Pending evaluation | Auto |
| 6 | Pending implementation | Auto |
| 7 | DDP MS4-5 | Header only (manual) |
| 8 | Eval handoff | Header only (manual) |
| 9 | Discussion | Header only (manual) |
| 10 | Risks | Header + legend; body manual |
| 11 | Quarterly planning | Auto if Book2 present |
| 12 | Closing | Auto |

**Will expand with:** title wording rules, column layouts, business filters.

---

## 5. End-to-end pipeline (code flow)

**Context:** One function orchestrates everything: `wsr.report.generate_report`.

```
GUI / CLI
   → validate Scrum (hard) + Planning (soft)
   → resolve week / report date from Graph sheet
   → load tracker, visibility, DDP, planning sheets
   → build chart PNGs (eval, impl, planning)
   → assemble 12 slides from template
   → save PPTX + sanitize package
   → return ReportResult(output_path, log_path, warnings)
```

**Key files:**

| Step | Module |
| --- | --- |
| GUI | `wsr_app.py` |
| CLI | `generate_report.py` |
| Orchestration | `wsr/report/generate.py` |
| Slide order | `wsr/report/deck.py` |
| Validation | `wsr/validate.py` |
| Charts | `wsr/charts.py` + `wsr/graph.py` |
| Pending tables | `wsr/pending.py` |
| Slide builders | `wsr/slides/*` |

**Will expand with:** sequence diagram, call graph, what each return type holds.

---

## 6. Package map (where code lives)

**Context:** After the refactor, logic is split by job — not one giant script.

```
wsr/
  report/        # generate, deck, timing, workbook, assets, models
  report_data/   # Excel → row dicts (summary, ddp, handoff, …)
  slides/        # PowerPoint layout per slide
  constants/     # sheet names + inch positions
  style/         # fonts, colors, table styling (wsr_style shim at root)
  validate.py, graph.py, charts.py, pending.py, tracker.py, …
```

**Will expand with:** “if you need to change X, open Y” cheat sheet.

---

## 7. Business rules (Excel → tables)

**Context:** Stakeholders care most about *why* a DCR appears or not.

Examples already encoded:

- **Pending (slides 5–6):** `PRCRState` Evaluate/Implement, `At Risk` = On Track,
  planned completion ≤ report date.
- **Summary (slide 4):** baseline totals (e.g. 130 / 50 / 80); some labels left blank
  by design.
- **DDP (slide 7):** headers only — filled manually in PowerPoint (row builder in
  `report_data/ddp.py` kept for optional future use).
- **Columns:** matched by **name** (exact header text), not column letter — with one
  planning-sheet fallback to column index.

**Will expand with:** full filter tables, status keyword lists, date parsing quirks.

---

## 8. Validation & errors

**Context:** Wrong Scrum file should not produce a silent bad PPT.

- **Hard fail (`WsrDataError`):** missing file/sheets/required columns, empty graph weeks.
- **Soft warning:** missing Planning book → slide 11 placeholder; PPT still builds.
- GUI shows message + **Open log**.

**Will expand with:** full checklist of checks, example error texts, log format.

---

## 9. Charts

**Context:** Slide 4 uses combo charts (bars = counts, lines = %). Data comes from
`CSAR_WSR_Graph (Non-STLA)` via `wsr/graph.py`; PNGs from `wsr/charts.py`.
Slide 11 uses Book2 hours → planned bandwidth bar chart.

**Will expand with:** series meaning, week detection, label placement notes.

---

## 10. Windows GUI distribution

**Context:** Office is Windows-only. App is PyInstaller → `WSR-Generator.exe`.
Built by GitHub Actions on every push to `main` (no Windows laptop needed to build).

Users replace the old `.exe` when code changes — no auto-update.

**Will expand with:** download steps, SharePoint handoff, version naming, SmartScreen,
GitLab alternative if the team never uses GitHub.

---

## 11. Developer setup & local runs

**Context:** For people changing code (often on Mac/Linux for edit, Windows for final exe).

```bash
pip install -r requirements.txt
python wsr_app.py
python generate_report.py --data path/to/scrum.xlsm --output WSR_Report.pptx
```

Windows exe locally: `pyinstaller --noconfirm WSR-Generator.spec`.

**Will expand with:** requirements, template/assets paths, common debug tips.

---

## 12. How we change things safely

**Context:** Small checklist before touching production Excel logic.

- Prefer column-name checks in `validate.py` when adding new fields.
- Slide order only in `deck.py`.
- Business filters in `pending.py` / `report_data/*`, not in slide layout files.
- Rebuild Windows exe and replace the shared copy after merging.

**Will expand with:** PR checklist, test plan (manual), “do not push comment-only noise”.

---

## 13. Known limitations & future ideas

**Context (parked):**

- Unsigned exe (SmartScreen friction)
- No auto-update / no version badge in the window yet
- GitHub vs GitLab distribution mismatch for the team
- SharePoint versioning / upload (mentor idea — awaiting approval)
- Handoff / discussion / DDP automation exists in `report_data/` but those slides
  are left manual by design (headers only)

**Will expand when decisions land.**

---

## 14. Glossary

**Context:** Short definitions (DCR, PRCRState, Non-STLA, Graph sheet, DRB/L2, MS4-5, …).

**Will expand as a one-pager for new joiners.**

---

## Installment plan

Suggested order when we expand:

1. §4 Slide map + §7 Business rules (what the PPT must show)
2. §5 Pipeline + §6 Package map (how code is organized)
3. §8 Validation + §3 Inputs (what breaks and why)
4. §10 Distribution (how Windows users get the app)
5. §2 Stakeholder story (polished user guide)
6. §9 Charts, §11 Dev setup, §12–14 as needed

Tell me which section to expand first.
