# WSR Generator

Generates the CES PFS CSAR (Non-STLA) Weekly Status Report PowerPoint from the
weekly Excel files. Built so a non-technical stakeholder can produce the WSR in
one click, with no Python or command line.

## For stakeholders (no Python needed)

1. Download the app for your OS from the latest build:
   - **Windows:** [Windows build](../../actions/workflows/build-windows.yml) →
     newest run → `WSR-Generator-windows` artifact → unzip → `WSR-Generator.exe`.
   - **macOS:** [macOS build](../../actions/workflows/build-macos.yml) → newest
     run → `WSR-Generator-macos` artifact → unzip → `WSR-Generator.app`.
2. Open the app (double-click).
3. **Browse** to your **Scrum workbook** (the `.xlsm`/`.xlsx` with the
   `CSAR_WSR_Graph (Non-STLA)` sheet).
4. Optionally browse to the **Planning workbook** (for the quarterly planning
   slide). If it sits next to the Scrum file as `Book2.xlsx`, it's picked up
   automatically.
5. Click **Generate WSR**. The report week and date are detected automatically
   from the sheet. When it finishes you're offered to open the `.pptx`.

> **First launch warnings (unsigned app):**
> - **Windows:** SmartScreen may warn → "More info" → "Run anyway".
> - **macOS:** Gatekeeper may block it → right-click the app → **Open** →
>   **Open**, or run `xattr -dr com.apple.quarantine WSR-Generator.app`.

## Inputs

| App field         | File                          | Used for                                   |
| ----------------- | ----------------------------- | ------------------------------------------ |
| Scrum workbook    | `data.xlsm`                   | Charts, DCR tables, status, DDP, handoff   |
| Planning workbook | `Book2.xlsx`                  | Quarterly planning slide (optional)        |

The reporting **week** and **date** are read from the latest populated week in
the `CSAR_WSR_Graph (Non-STLA)` sheet — no manual entry required.

## Developers

Install and run from source:

```bash
pip install -r requirements.txt

# GUI app
python wsr_app.py

# Command line (week/date auto-detected; override with --week / --date)
python generate_report.py --data data.xlsm --output WSR_Report.pptx
```

## Building the Windows executable

You don't need a Windows machine: every push to `main` runs the
[`Build Windows App`](.github/workflows/build-windows.yml) workflow and uploads
`WSR-Generator.exe` as an artifact. Tagging a release also attaches the exe to
the release.

To build locally on Windows:

```bash
pip install -r requirements-dev.txt
pyinstaller --noconfirm WSR-Generator.spec
# -> dist/WSR-Generator.exe
```

The template (`templates/`) and closing image (`assets/`) are bundled into the
executable automatically.
