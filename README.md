# Force Automation — Windows Packaging Guide

This package turns your existing `final.py` (Calibration Report Automation
System) into a Windows desktop app that installs and runs on a machine with
no Python installed. **No calibration, calculation, certificate-generation,
or UI logic was changed.** Only deployment-related code was added.

---

## 1. What was analyzed

- **Entry point:** single script, Dash app with an embedded Flask server
  (`server = Flask(__name__)`, `app = Dash(__name__, server=server, ...)`).
- **Imports:** `dash`, `dash_bootstrap_components`, `flask`,
  `python-dateutil`, `pandas`, `numpy`, `python-docx`, `openpyxl` — all in
  `requirements.txt`.
- **Certificate templates:** loaded via relative filenames
  `"C-Certificate no-Year-PR capacity, party name.docx"` and
  `"...-LC capacity, party name.docx"`, opened with `python-docx`.
  This works only when the current working directory happens to be the
  project folder — it would break once packaged.
- **Images/assets:** `html.Img(src="/assets/csir_logo.png")` — Dash's
  built-in convention of auto-serving a folder named `assets/` next to the
  script (this also picks up `style.css` automatically).
- **Output files:** `doc.save(filename)` saved into whatever the current
  working directory was — unsafe once packaged (could land in
  Program Files, a read-only folder, or PyInstaller's temp extraction dir).
- **Server startup:** two conflicting blocks existed — an unused
  `start_server()` helper, and the actual `if __name__ == "__main__":` block
  running `app.run(debug=True, host="127.0.0.1", port=5080)` (development
  mode, no browser auto-open, single fixed port with no fallback).
- **No hardcoded developer-machine paths** (e.g. `C:\Users\Akansha\...`)
  were found in the source.

## 2. What was changed, and why

| File | Change | Why |
|---|---|---|
| `main.py` (was `final.py`) | Added `resource_path()` helper | Resolves template/asset paths correctly both in dev and inside a PyInstaller build (onefile `_MEIPASS` or onedir exe folder). |
| `main.py` | Added `get_output_dir()`, and changed both `doc.save(filename)` calls to save into `Documents/Force Automation/Certificates/` | Guarantees a writable location, regardless of install folder, without needing admin rights. |
| `main.py` | Added `get_log_dir()`, `logging.basicConfig(...)`, and a global `sys.excepthook` | A windowed (no-console) build has nowhere to show a crash — errors are now written to `%APPDATA%\Force Automation\logs\force_automation.log`. |
| `main.py` | `Dash(__name__, ..., assets_folder=resource_path("assets"))` | Makes the CSIR logo and `style.css` load correctly when bundled. |
| `main.py` | Both `TEMPLATE_PATH` assignments now use `resource_path("templates/...")` | Certificate templates are found reliably after packaging. |
| `main.py` | Certificate status message now also shows the save folder | Small UX improvement so the user immediately knows where their `.docx` went — no logic change. |
| `main.py` | Rewrote the bottom `if __name__ == "__main__":` block: removed the leftover `debug=True` dev block, added `find_free_port()`, a single `threading.Timer` that opens the browser exactly once, and try/except logging around startup | Meets the "production mode" and "auto-open browser without duplicate tabs" requirements without touching any Dash callback or business logic. |

Every other line of your original calculations, callbacks, interpolation,
classification, uncertainty, certificate numbering, and Word
formatting logic is untouched — see the diff below for verification.

## 3. Project structure

```
Force_Automation/
├── main.py                      # your original final.py, with only the deployment changes above
├── requirements.txt
├── Force_Automation.spec        # production build (windowed, no console)
├── Force_Automation_debug.spec  # debug build (console visible, for troubleshooting)
├── build_exe.bat
├── templates/
│   ├── C-Certificate no-Year-PR capacity, party name.docx
│   └── C-Certificate no-Year-LC capacity, party name.docx
├── assets/
│   ├── csir_logo.png
│   └── style.css
└── installer/
    ├── installer.iss
    └── app_icon.ico
```

After building, PyInstaller will also create `build/` (temporary) and
`dist/Force_Automation/` (the finished app — an "onedir" build, i.e. a
folder containing `Force_Automation.exe` plus everything it needs).

**Why onedir instead of onefile?** A onefile `.exe` re-extracts itself to a
temp folder on every launch, which is slower and adds another layer of path
resolution to get wrong. Onedir starts faster and makes resource paths
fully predictable. The Inno Setup installer packages the whole
`dist/Force_Automation/` folder, so the end user never sees the difference
— they just get one `Force_Automation_Setup.exe` to run.

## 4. Build instructions (run on Windows)

You need **Python 3.10+** and **Inno Setup 6** installed on the Windows
machine that builds the installer (the end user needs neither).

### Step A — Install dependencies and build the .exe

```bat
build_exe.bat
```

This creates a virtual environment, installs everything in
`requirements.txt`, and runs:

```bat
pyinstaller Force_Automation.spec
```

Result: `dist\Force_Automation\Force_Automation.exe`

Run it directly to sanity-check it on the dev machine first:

```bat
dist\Force_Automation\Force_Automation.exe
```

It should open your default browser to the calibration UI with no console
window. If something goes wrong and you need to see errors, build the debug
variant instead (shows a console window):

```bat
pyinstaller Force_Automation_debug.spec
dist\Force_Automation_Debug\Force_Automation_Debug.exe
```

or check `%APPDATA%\Force Automation\logs\force_automation.log`, which is
written by both builds.

### Step B — Build the installer

Open `installer\installer.iss` in Inno Setup Compiler and click **Compile**,
or from the command line:

```bat
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\installer.iss
```

Result: `installer\Output\Force_Automation_Setup.exe`

This is the single file you distribute.

## 5. What the installer does

- Installs into `Program Files\Force Automation\` (per-user, no admin
  required to install or run — `PrivilegesRequired=lowest`).
- Creates a Start Menu entry and an optional Desktop shortcut.
- Registers standard Windows uninstall information (visible in
  "Add or Remove Programs").
- Bundles the Python runtime and every dependency — nothing else needs to
  be installed on the target machine.

## 6. Where things live on the user's machine

| What | Where |
|---|---|
| Application files | `C:\Program Files\Force Automation\` (or wherever installed) |
| Generated certificates | `Documents\Force Automation\Certificates\` |
| Error/debug log | `%APPDATA%\Force Automation\logs\force_automation.log` |

## 7. Testing checklist

Test on a clean Windows VM/machine with **no Python, pip, or project source
code present**:

1. Run `Force_Automation_Setup.exe` → installer completes, shortcuts appear.
2. Double-click the Desktop or Start Menu shortcut → browser opens
   automatically to the Calibration UI, no console window, only one tab.
3. Load Cell workflow: enter data → Compute → Generate Certificate → confirm
   the `.docx` appears in `Documents\Force Automation\Certificates\` and
   opens correctly in Word.
4. Proving Ring workflow: same checks as above.
5. Generate a second certificate of each type → confirm no filename
   collisions or permission errors.
6. Uninstall from "Add or Remove Programs" → confirm the app is removed
   cleanly (your generated certificates in Documents are intentionally left
   in place — they're the user's data, not application files).

## 8. Icon

`installer\app_icon.ico` was generated from your uploaded `csir_logo.png`
(multi-resolution: 16–256px) and is used for the `.exe`, shortcuts, and the
installer itself. If you'd prefer a different icon, replace
`installer\app_icon.ico` with your own `.ico` file (ideally including
16, 32, 48, and 256px sizes) before building.
