# vbaloop agent setup procedure

This document is for coding agents (Claude Code, Codex, etc.) to read and follow as written. The user only needs to give the agent the repo URL.

> Set up vbaloop in this folder using https://github.com/rfastball/vbaloop

If the user did not name a workbook or project, follow "Decide from the current folder" in step 1.
Talk to the user in the language they use with you; this document is in English only for consistency.

## Principles

- **Change nothing without consent.** AccessVBOM (an Excel security setting) and installing the exe / adding it to PATH are
  reported by `setup --plan` under `consent`. Ask the user exactly that and follow their answer.
- **Add `--enable-vbom` only when the user has explicitly agreed to that security setting change.** Never edit the registry directly.
  `--yes` only skips prompts; it does not stand in for that consent.
- **Never edit, move or delete the original workbook directly.** vbaloop makes a copy next to the original.
- **Never touch Excel directly.** Killing Excel processes and COM access such as PowerShell `GetActiveObject` / `New-Object -ComObject`
  are all forbidden. Check the environment with `vbaloop doctor --json`.
- Pass `--json` to every vbaloop command and branch only on the envelope's `ok` and `hint`. On failure, `hint` comes first.

## 1. Decide what to set up

| Situation | Starting point | What to confirm |
|---|---|---|
| The user points at a workbook | `--workbook "<absolute path to original .xlsm>"` | That it is `.xlsm`. For `.xlsx`/`.xls`, ask the user to Save As `.xlsm` in Excel |
| Starting fresh | `--new "<name>"` | The name, and where to create it (it is created under the folder this command runs in) |
| A prepared folder | `--folder "<folder>"` | That the folder has 0 `.xlsm` (one is created) or exactly 1 (it is adopted) |

### When the user only gave the repo URL: decide from the current folder

If the user did not name a target, decide from the agent's current working folder. This is only a candidate; nothing changes
until the user confirms it from the `--plan` result in steps 3-4.

| Current folder | Candidate starting point |
|---|---|
| Has `vbaloop.toml` | Already a vbaloop project. Tell the user there is nothing to set up - just work by the managed block in that folder's `CLAUDE.md`/`AGENTS.md`. If they want the managed files repaired, `--folder "<current folder>"` |
| Exactly 1 `.xlsm` | `--workbook "<current folder>\<that file>"`. A `<name>\` copy folder is created next to it |
| Several `.xlsm` | Ask the user which one. Use `--workbook` with the chosen one |
| Empty apart from entries starting with `.` (`.git` etc.) | New project here: `--folder "<current folder>"`. The workbook is named after the folder |
| Anything else (only other files) | Ask the user. For a new project, get a name and use `--new "<name>"` (created under the current folder); for an existing workbook, get its path |

A `.xlsm` starting with `~$` is an Excel lock file; skip it. If there is only a `.xlsx`, it is not a macro workbook, so ask the user
whether to Save As `.xlsm`.

## 2. Get vbaloop

```
vbaloop version
```

If `vbaloop` is not found, try `%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe version`.

If found, check that it has the features this procedure needs. Old versions lack `--plan`, `--workbook` and `--enable-vbom`.

```
vbaloop help
```

If the `setup` line of the output has no `--plan`, it is an old version, so **upgrade** with the install script below (it overwrites in place).
Use the same script when the exe is missing entirely. This downloads and runs an install script from the internet, so tell the user before proceeding.

```
powershell -NoProfile -NonInteractive -Command "irm https://raw.githubusercontent.com/rfastball/vbaloop/master/install.ps1 | iex"
```

In a human's terminal the script continues into interactive setup after installing, but in an agent shell (no input, `-NonInteractive`)
it installs without asking and then prints agent guidance (the exe's absolute path, this document, how to use `--plan`).

Installing **does not change the PATH of the shell the agent is using right now.** Run later commands with the absolute path
`%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe` (PowerShell: `& "$env:LOCALAPPDATA\vbaloop\bin\vbaloop.exe"`).
Read every `vbaloop` below as that path.

## 3. Look at the plan (changes nothing)

Run with the starting point chosen in step 1.

```
vbaloop setup --plan --json --workbook 'C:\work\Settlement.xlsm'
```

Reading the `payload`:

| Field | Meaning |
|---|---|
| `ready` | If `false`, pass `blockers` and the envelope's `hint` to the user and stop (Excel not installed, folder conflict, etc.) |
| `project` | Project folder (`folder`), workbook, `action` (copy / reuse / inplace / create / adopt), verification method (`verify`) |
| `changes` | What running it will change. A summary to show the user |
| `consent[]` | Items that need consent. `approveFlag` is the flag to add on approval (null means it is already the default), `declineFlag` the flag to add on refusal (null means refusing makes it impossible to proceed) |
| `args` | Array of run arguments without the consent flags (starting from `"setup"`). Append it as is after the exe path. The recommended form - no shell quoting issues |
| `command` | The same command as one line quoted for PowerShell. Paths are single-quoted so `$` and backticks are not interpreted. In other shells such as bash, use `args` |

If `ready` is `false` with an "inside another vbaloop project" error, that folder belongs to an outer project. Ask the user for another location.

## 4. Confirm with the user

Show the `changes` summary and the `consent` items together and ask once. For example:

> Here is the setup:
> - Make a copy of the original in `C:\work\Settlement\` (the original stays as is) and extract the code into `src\`.
> - Install vbaloop to `%LOCALAPPDATA%\vbaloop\bin` and add it to PATH.
> - **The Excel security setting "Trust access to the VBA project object model" must be turned on.** vbaloop does not work
>   while it is off. You can turn it off later in Excel Options > Trust Center > Macro Settings.
>
> Shall I go ahead?

- If `accessVbom` is refused, setup cannot proceed. Tell the user to turn it on themselves in Excel Options and try again.
- If `install` is refused, add `--no-install`. Tell the user they will have to run it by absolute path every time.

## 5. Run

Run `args` (or `command` in PowerShell) with the agreed `approveFlag`s (e.g. `--enable-vbom`) and the refused `declineFlag`s appended.
It starts Excel to verify, so it may take a few minutes. Set the timeout to 5 minutes or more.

Example after the user agreed to the AccessVBOM change (without that agreement, leave `--enable-vbom` off):

```
vbaloop setup --workbook 'C:\work\Settlement.xlsm' --yes --json --enable-vbom
```

On failure:

- `payload.consentRequired` contains `accessVbom` -> go back to step 4 and get consent. Never add the flag without consent.
- exit 3 -> follow the `hint`. If Excel is in cell edit mode or has a dialog open (busy), tell the user and retry.
  For a quarantine (`recovery`) hint, do not kill Excel; tell the user the situation.
- Anything else -> pass `errors` and `hint` to the user as they are.

## 6. Wrap up: `payload.next`

Report to the user:

- The project folder (`next.folder`). If it is a copy, that the original is untouched.
- The verification result (`verify`): for `pull`, the extracted code is in `src\`. For `push+test`, the sample test passed.

Offer two ways to continue:

- **New session (recommended)**: open the agent in the project folder using the entries in `next.newSession` (per-agent `workingDirectory` and `command`, plus a PowerShell one-liner `powershell`).
  That folder's `CLAUDE.md`/`AGENTS.md` and the skill load automatically.
- **Continue in this session**: as `next.continueHere` says, read the managed block and `SKILL.md` and follow their rules. Run vbaloop commands
  from the project folder, or point `--config` at that folder's `vbaloop.toml`. Use the absolute exe path from `next.exe`.

If `next.excelSession` is present, the Excel used for verification is still open. If you are not continuing right away, close it with
`vbaloop session stop --json` from the project folder.

## Appendix: common situations

- **The original is open in Excel**: the last saved content is copied. If there are unsaved changes, suggest the user saves first.
- **Setting up the same workbook again**: the dev folder made earlier is reused (`action: reuse`). Later changes to the original are not picked up.
  To import it fresh, after the user confirms, also pass `--folder "<new folder>"`.
- **Several `.xlsm` in the folder**: specifying one with `--workbook` imports only that one into the new folder.
- **The `--new` folder already exists**: use another name, or if you mean to continue there, point `--folder` at it.
