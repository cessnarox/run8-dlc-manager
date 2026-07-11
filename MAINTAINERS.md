# Maintainer notes

For whoever picks this up — welcome. The design goals, in order: never lose
a user's transaction IDs, never delete anything, stay ONE file with zero
dependencies beyond the Python standard library.

## Layout

Everything is in `Run8DLCManager.pyw` (~6,000 lines):

- Top: paths (`APP_DIR`, `DATA_DIR` = `data\` next to the program, with
  one-time self-migration of pre-0.9.16 loose files), `VERSION`,
  `detect_defaults()`, `CATALOG_SEED` (the embedded 3DTS store catalog),
  fetch/fuzzy-match helpers.
- `class App`: config/catalog/ledger state and the evidence-based
  ownership scanner (installer filenames + in-game folders/files).
- `cmd_*` functions: the CLI (argparse subcommands in `main()`). The GUI
  drives long operations by spawning `[sys.executable, __file__, <cmd>]`
  and streaming output into its log — `[dl] NN%` lines drive the
  progress bar.
- `class Gui(tk.Tk)`: three tabs (Gallery / List / Settings) plus overlay
  pages (Add Purchase, Import, Transactions). The gallery is CANVAS-NATIVE:
  tiles are drawn items (`_gtiles` / `_gal_relayout`), never embedded
  widgets — that is what makes scrolling smooth; do not put widgets back
  into the canvas.
- `SetupPane(tk.Frame)`: the Settings page; also hosted in a Toplevel
  (`SetupWizard`) for first-run only.
- Embedded assets: `ICON_B64` (original icon), `THEME_ICONS` (12 pre-tinted
  per-theme icons), masthead tables (`LOCO_SCHEMES`, `CONSISTS`,
  `PALETTES`) — the toolbar train and the theme picker draw from these.

## Rules that keep users safe

1. NEVER overwrite an existing transaction ID in the ledger.
2. NEVER delete user files. "Disable" moves route folders to
   `data\uninstalled\`; only the purge command really deletes, behind a
   double warning.
3. Damaged JSON must never prevent startup — `load_json` keeps a
   `.corrupt` copy and continues with defaults so Restore/Reset stay
   reachable.
4. Destructive CLI commands dry-run unless `--yes`.
5. Don't touch `CATALOG_SEED` product descriptions by hand; the `refresh`
   command maintains the live `data\catalog.json` (new products, alternate
   URL spellings via `alt_urls`).
6. The app icon is never redrawn — per-theme versions are RECOLORED from
   the original (users are attached to it).

## Working on it

- Any Python 3.10+ with tkinter runs it: `python Run8DLCManager.pyw`.
- `--demo` gives a full simulation mode (fake statuses, every workflow
  acted out, nothing written) — use it to test UI states you don't own.
- `python Run8DLCManager.pyw --help` lists the CLI; `doctor` repairs
  ledger/catalog damage.
- Check compiles cleanly: `python -W error::SyntaxWarning -m py_compile
  Run8DLCManager.pyw`.
- Build the EXE with `build_exe.bat` (PyInstaller, one file). Antivirus
  sometimes side-eyes fresh PyInstaller builds; offering the .pyw
  alongside the EXE is deliberate.
- The store scraper (`cmd_refresh` + `SLUG_RE`) and the update checker
  (`_check_game_update`, parses "MONTH DD YYYY" headlines on
  run8studios.com) are the only things that break if 3DTS/Run8 redo their
  websites — both fail silent-safe by design.

## Versioning

0.0.1 steps. Bump `VERSION` every release; it shows in the GUI log line
and the Settings footer.
