# Run8 DLC Manager

One window that keeps your whole Run8 Train Simulator DLC collection under
control: what you own vs. the full 3DTS store, what's actually installed in
the sim, every transaction ID you've ever been issued, new purchases,
reinstalls, reversible disables, one-click backups — and it tells you when
the game itself needs updating.

## Get started (2 minutes)

1. Put **`Run8DLCManager.exe`** (or `Run8DLCManager.pyw` if you have
   Python 3) in its own folder — for example
   `Documents\Run8 DLC Manager`. It keeps everything it makes in a
   `data` folder next to itself; your folder stays tidy.
2. Double-click it. A short first-run setup asks where things live (your
   Run8 install is auto-detected), lets you pick a text size and a
   railroad color theme, and offers to import your old purchase records —
   receipt screenshots (read by Windows' built-in OCR), saved store
   emails (.eml), or a document/spreadsheet with your transaction IDs.
3. That's it. The manager scans your collection and shows everything.

A **Settings** tab covers the same choices any time, and can create
Desktop / Start Menu shortcuts for you.

## The window

Three tabs — **Gallery**, **List**, **Settings** — plus two buttons:
**Check My Collection** (rescan after you install something) and
**Add Purchase**.

**Gallery** shows every product in the store as a picture card, color-coded
by status. Click the little **▾** next to a title (or the title itself) for
the store's description. Each card carries its own buttons — **Buy** on
products you don't own, **Reinstall** / **Disable** / **Enable** on ones
you do.

**List** is the same catalog as a sortable table with a detail panel:
bigger photo, full description, the evidence for how the manager knows you
own something ("How I Know"), and the action buttons. Drag the divider
between table and panel wherever you like — it remembers.

**Status chips** above either view filter with one click:
**Installed** (verified inside the game folder), **Owned (not installed)**,
**Disabled**, **Not owned**. The bottom line always shows your totals and
the cost to complete the set.

## Buying and recording purchases

**Buy** opens the real 3DTS store page in your browser — the manager never
touches your payment. Afterwards, click **Add Purchase**:

- **Receipt page address** (easiest): copy the address bar on the
  "Transaction Approved" page — if it's on your clipboard, it's already
  filled in. The manager saves the transaction ID, downloads the
  installer, and files everything.
- **Transaction ID only**: paste the ID — or a whole download link, the ID
  is pulled out automatically. The manager builds the store's standard
  download link itself.
- **Installer file or purchase email**: point it at an installer EXE you
  already have (any add-on file works), or a saved .eml receipt email.

Every purchase lands in the ledger and in a plain `transactions.txt`.
**Transaction history** (in Settings → Tools) shows the full list.
Nothing is ever overwritten.

## Disable / Enable

**Disable** moves a route's folders into a quarantine inside the manager's
data folder — nothing is ever deleted — and **Enable** puts them straight
back. Note this works on **routes**: equipment packs share their files
with other products, so they can't be switched off one at a time.
"Permanently delete disabled items" (Settings → Tools) is the only thing
in the app that ever really deletes, and it warns you twice.

## Updates — the game and the store

After each scan the manager quietly checks run8studios.com. If the site
announces an update newer than your game files, a **"Game update
available!"** button appears and the official Run8 updater runs — its
output right in the manager's log. (The updater also updates DLC files,
which is usually required after installing new content — the manager
prompts you for the same reason when it sees something newly installed.)
If the store lists products the catalog doesn't know yet, a **"New DLC in
store!"** button adds just the new ones in seconds. Both checks can be
switched off in Settings.

**Refresh Store Prices** (Settings → Tools) re-reads every product page
for current prices — only needed occasionally; they don't do sales.

## Backups — and getting out of trouble

**Back Up Now** (Settings) packs your installers, ledger, records, and
settings into one compressed `Run8DLC_Backup.zip` in your Backups folder
(set it to another drive for real safety). Each backup replaces the last.

**Restore from backup** unpacks one over your current setup — with a
serious warning first, and your current records are snapshotted before
anything is touched. **Reset settings to defaults** wipes only settings
(never your ledger or installers) and reruns first-time setup. And if a
data file ever gets damaged, the app keeps a copy aside and starts anyway
— so the rescue tools are always reachable.

## Demo mode

Run the manager with ` --demo` added to the shortcut and you get a safe
playground: fake statuses, every button working, every workflow simulated
in the log — nothing on disk is touched. Great for seeing what the app
does before pointing it at your real collection.

## Folder layout

```
Run8 DLC Manager\
  Run8DLCManager.exe   <- the program (or .pyw)
  README.md            <- this file
  data\                <- everything the app makes: settings, catalog,
                          ledger, images, quarantined routes
  Installers\          <- your DLC installer EXEs
  Backups\             <- Run8DLC_Backup.zip lands here
```

## Command line (optional)

Everything is scriptable:
`report · installed · add · reinstall · uninstall · restore · updater ·
transactions · refresh · media · import-records · ocr-receipts · backup ·
restore-backup · doctor · snapshot`. Product names are fuzzy ("mp15 pack
3" works). `doctor` finds and repairs ledger/catalog damage; destructive
commands dry-run unless you add `--yes`.

## Building the EXE from source

`build_exe.bat` (in the project's development folder) turns
`Run8DLCManager.pyw` into a single no-Python-needed EXE via PyInstaller.
Antivirus tools sometimes side-eye freshly built PyInstaller EXEs;
sharing the .pyw avoids that entirely.
