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
| **Trust access to the VBA project object model** (`AccessVBOM=1` under `HKCU\Software\Microsoft\Office\<ver>\Excel\Security`) | `vbaloop setup` after you agree, or `vbaloop doctor --fix` | Excel → Options → Trust Center → Macro Settings, or set the value back to `0` |
| Copies `vbaloop.exe` to `%LOCALAPPDATA%\vbaloop\bin` and adds it to your **user** PATH | `vbaloop setup`, `setup --install-only`, `install.ps1` | Delete the folder and remove the PATH entry |
| Per-user state in `%LOCALAPPDATA%\vbaloop\` (Excel PIDs it started or quarantined) | Any session command | Delete the folder |
| Project files: `vbaloop.toml`, `src\LogShim.bas`, managed blocks in `CLAUDE.md`/`AGENTS.md`/`.gitignore`, `.claude\skills\vbaloop\`, and `_dbg\` (logs, session state, workbook backups) | `vbaloop init` / `setup`, and normal use | Delete them; content outside the managed markers is never touched |

Trusting VBA project access lets any macro or program running as you modify VBA projects. Enable it
only on a machine where you develop VBA, and turn it off when you are done if that matters to you.

## What vbaloop does not do

- It never sends your code or workbooks anywhere. It makes no network requests at runtime;
  only `install.ps1` downloads the release from GitHub.
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
