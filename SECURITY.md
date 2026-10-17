# Security

## Reporting a vulnerability

Please report security problems privately through GitHub's
[private vulnerability reporting](https://github.com/rfastball/vbaloop/security/advisories/new)
rather than a public issue. Include the vbaloop version (`vbaloop version`), Windows and Excel versions,
and steps to reproduce. You should get a first response within a week.

## What vbaloop changes on your machine

vbaloop automates Microsoft Excel through COM. To do that it needs, and may change, the following.
Nothing here is done silently: `setup` asks first, and `doctor` only reports unless you pass a flag.

| Change | When | How to undo |
|---|---|---|
| **Trust access to the VBA project object model** (`AccessVBOM=1` under `HKCU\Software\Microsoft\Office\<ver>\Excel\Security`) | `vbaloop setup` after you agree at its prompt, `setup --enable-vbom` (non-interactive; `--yes` alone never changes it), or `vbaloop doctor --fix` | Excel → Options → Trust Center → Macro Settings, or set the value back to `0` |
| Copies `vbaloop.exe` to `%LOCALAPPDATA%\vbaloop\bin` and adds it to your **user** PATH | `vbaloop setup`, `setup --install-only`, `install.ps1`, `vbaloop update` (replaces the exe after SHA-256 verification) | Delete the folder and remove the PATH entry |
| Per-user state in `%LOCALAPPDATA%\vbaloop\`: Excel PIDs it started or quarantined (`excel-pids.json`), the project folders it has seen (`projects.json`, used by `init --all`), and the last release check (`update.json`) | Session commands; any command that reads a `vbaloop.toml`; `status`/`doctor`/`version`/`init --check` | Delete the folder |
| Project files: `vbaloop.toml`, `src\LogShim.bas`, managed blocks in `CLAUDE.md`/`AGENTS.md`/`.gitignore`, `.claude\skills\vbaloop\`, and `_dbg\` (logs, session state, workbook backups) | `vbaloop init` / `setup`, and normal use | Delete them; content outside the managed markers is never touched |
| A copy of your workbook in a new `<name>\` folder next to it (the original is never modified) | `setup --workbook`, or dropping a workbook on `vbaloop.exe` | Delete the folder |

`vbaloop setup --plan` lists every change above that a given `setup` would make, without making any of them.

Trusting VBA project access lets any macro or program running as you modify VBA projects. Enable it
only on a machine where you develop VBA, and turn it off when you are done if that matters to you.

## What vbaloop does not do

- It never sends your code or workbooks anywhere. Its only network requests go to the GitHub API
  and release downloads for this repository: at most once a day a background check of the latest
  release number (from `status`, `doctor`, `version`, `init --check`; off in CI, for builds inside the
  source repo, or with `VBALOOP_NO_UPDATE_CHECK=1`), and `vbaloop update` / `install.ps1`, which
  download the release and refuse to install it unless its SHA-256 matches `checksums.txt`.
- It does not need administrator rights and writes nothing outside your user profile and project folder.
- On its own it ends an Excel process only if it started that process itself and has confirmed no other
  workbook is open in it; otherwise it quarantines and asks you (see DESIGN §11).
- `vbaloop doctor --kill-orphans` is the one explicit exception: it ends windowless Excel instances
  that were started by COM automation in your own Windows session. Such an instance may belong to
  another automation tool, so run `vbaloop doctor` first — it lists each candidate and says whether
  vbaloop started it.

## Verifying a download

Release binaries are not Authenticode-signed yet, so Windows SmartScreen may warn on first run.
Each release ships `checksums.txt` and a GitHub build provenance attestation:

```powershell
gh attestation verify .\vbaloop.exe --repo rfastball/vbaloop
```
