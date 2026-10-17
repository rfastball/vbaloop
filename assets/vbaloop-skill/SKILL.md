---
name: vbaloop
description: Autonomous development loop for Excel VBA projects. Use when editing, injecting, testing or debugging VBA source (.bas/.cls/.frm), working on .xlsm workbook macros, running vbaloop commands (session/push/pull/test/run/diff/lint/doctor/dist/rollback), building a distribution artifact (dist), or recovering from failures (poisoned/stale/busy/quarantined session, attaching to the user's Excel, lease conflicts, timeouts, compile-error modals, partial push failures, restoring backups).
---

# vbaloop workflow

vbaloop is a set of atomic commands that safely round-trip VBA between src text files and an Excel workbook.
This skill holds the procedures that chain those commands and the failure playbooks. The base rules (encoding,
do not edit managed copies, etc.) live in the CLAUDE.md managed block. Command reference: `vbaloop help` (details: https://github.com/rfastball/vbaloop).

## Symptom -> action (incident entry point)

When stuck, jump from here to the matching section. The rule is always **the envelope's `hint` comes first**;
the table below is context for interpreting the hint.

| Symptom | Action |
|------|------|
| `init --check` exits 3 + `payload.state: "exe-older"` (or a warning that the exe is older than the project) | **Do not run init**; ask the user to run `vbaloop update` -> §exe older than the project |
| Session missing/stale/poisoned (exit 3) | auto-ensure cleans up and restarts. If it repeats, `doctor` -> §Session poisoned / stale |
| Session quarantined (exit 3, hint mentions `recovery clear`/`--discard`) | Never kill Excel yourself; tell the user -> §Session quarantined / busy |
| busy (exit 3, Excel in cell edit mode or showing a dialog) | Retry the same command shortly -> §Session quarantined / busy |
| read-only refusal + hint mentions `session attach` | After the user confirms, `session attach` -> §Attaching to the user's Excel |
| test: 1 fail + many skipped | Only the first fail's `detail` (modal text) is a real signal -> §test timeout / modal |
| Fixed it but same error, push keeps returning `cached:true` | `push --force` bypasses the fingerprint cache -> §push refused |
| push refused (CP949, module mismatch, document module) | §push refused |
| push fails + `payload.compile.result: compile_error` | Read the modal text, fix, push again -> §push compile gate |
| push aborted + `payload.recovery` | §Workbook state after a failed push |
| Want the workbook back to its pre-push state | (`session stop` first if a persistent session is up) `rollback --latest` -> §Backups and rollback |
| pull blocked (unsynced src edits) | Default: `push` first; if the workbook is the truth, `pull --force` -> §pull caveats |
| lease conflict (held by another agent) | Retry / wait for auto-reclaim -> §lease conflict |
| dist unsynced, dangling, bootstrap timeout | `references\dist.md` |
| Unexplained compile/runtime error | Look up `references\vba-pitfalls.md` |
| Suspect the environment (Excel/COM/VBOM) | `doctor --json` first |

## Reading results (all commands)

- Always pass `--json`. Envelope: `{command, ok, exitCode, hint, errors[], payload}`
- Branch on two fields only: `ok` (success) and `hint` (next action)
- exit code: success 0 / environment error 3 (session missing, stale or poisoned, fatal Excel/COM, busy, quarantined - follow the hint) / otherwise 2
- On failure `hint` says what to do next. Following the hint always comes first
- Do not remember or infer state - the tool decides whether revalidation is needed (hash cache, idempotent start)
- An `update: {current, latest, command}` field in the envelope only announces a new vbaloop release (unrelated to ok/hint).
  Tell the user after finishing the task. `vbaloop update` replaces the exe on the PC, so run it only when the user wants

## Standard loop

1. `vbaloop init --check --json` - onboarding/drift check. On drift (exit 2) converge with `vbaloop init`, then continue.
   On exit 3 + `payload.state: "exe-older"` do not call init; stop (§exe older than the project)
2. Edit src\*.bas/.cls/.frm
3. `vbaloop push --json` - no-op if nothing changed. On failure errors carry file:line inline (e.g. a non-CP949 character)
4. `vbaloop test --json` - payload.tests[].detail carries the failed assertion (expected/actual) inline
5. On fail, find the cause from detail and go back to 2. When green (payload.summary.fail == 0 and
   payload.summary.skipped == 0), git commit
6. If you used a persistent session, `vbaloop session stop --json` - saves and exits. To discard changes, `--no-save`

Commands obtain Excel on their own. How is set by `[excel] session` in `vbaloop.toml`:

- `ephemeral` (default): each session command (push/pull/test/run/diff) starts a hidden Excel and closes it when done
  (envelope `session: {action:"ephemeral", pid, readOnly}`). Only push opens for write and saves; the rest open read-only -
  sheets changed by tests do not persist, and test/diff work even while the user has the canonical workbook open. No `session stop` needed
- If a session opened with `session start`/`session attach` is alive, it is used regardless of mode (for fast iteration; `session stop` when done)
- `persistent`: the old behavior. On a missing/stale/poisoned session it cleans up, starts one, and keeps it between commands
  (envelope `session: {action:"started", pid}`)

Only `Public Sub Test_*()` (no arguments) are collected as tests. Assert with `LogShim.AssertEq`.

## Failure playbooks

### exe older than the project (exe-older)

- `[vbaloop] managed_by` in `vbaloop.toml` is the vbaloop version that last wrote the managed files. If the running exe is
  older, init (and init --check) refuses with exit 3 + `payload.state: "exe-older"` and a hint asking the user to run `vbaloop update`.
  This is not drift - "converging" to an older exe's canonical files would downgrade the project
- **Do not fix this yourself**: do not retry `vbaloop init`, do not delete or edit the managed_by line, and do not
  revert managed files by hand. Tell the user the situation (managedBy and exeVersion) and ask them to run
  `vbaloop update` (installs the new exe). After the update, start again from `init --check`
- push/test etc. are not blocked; they only print the same warning on stderr. You can keep looping, but tell the user
- If `managed_by` is not semver (exit 2, `state: "invalid"`), ask the user to check the toml

### Session poisoned / stale

- Rare in normal use - commands that need a session clean up and restart on entry (auto-ensure)
- exit code 3 means the automatic start itself failed (Excel launch / fatal COM) - even without parsing JSON you can
  branch to `session start` and retry. If it repeats, check the environment with `vbaloop doctor --json`
- After a fatal COM error or timeout the session is marked poisoned. If Excel is already dead or it is confirmed that
  there is no user work, the tool cleans up and the next command restarts - otherwise it is quarantined (below)
- For diagnosis, `vbaloop status --json` shows the state (none/active/poisoned/stale/busy/quarantined)

### Session quarantined / busy

- The tool never kills Excel automatically. After a timeout/poison, if it cannot prove that Excel holds no user work
  (other workbooks open, busy, unresponsive, an attached user Excel), it leaves `_dbg\recovery.json` and
  later session commands are refused with exit 3 + hint
- **Do not lift it yourself.** Do not kill Excel with `session stop --discard` (loses unsaved changes) or
  `doctor --kill-orphans`; tell the user the situation (PID and reason from the hint):
  - If the user saves/closes that Excel themselves -> `vbaloop recovery clear --json`
  - If the user says it may be discarded -> `vbaloop session stop --discard --json`
    (only terminates an Excel the tool started; refuses if non-session workbooks are visible. An attached session is only detached)
- `vbaloop recovery status --json` shows the quarantine reason and whether the process is alive. If the recorded Excel
  has already exited, the next session command lifts the marker automatically
- `recovery clear --force` is a release that needs human judgment - only when the user tells you to
- busy (exit 3, Excel is in cell edit mode or has a dialog open): the user is using Excel.
  Retry the same command shortly; if it keeps happening, tell the user. The session is kept

### Attaching to the user's Excel

- If the user already has the canonical workbook open in their own Excel, push (write) is refused as read-only and
  the hint suggests `vbaloop session attach` (in ephemeral mode test/pull/diff are read-only and keep working)
- attach makes push saves and test/run macro execution happen inside the user's Excel, so **only after the user confirms**
  run `vbaloop session attach --json`. After that the loop (push/test) is the same as usual
- `vbaloop session stop --json` only detaches (no save/close/quit - only what push saved remains)
- A timeout/poison in an attached session is always a quarantine - the user's Excel is never killed

### lease conflict (multiple agents)

- Mutating commands (session/pull/push/test) acquire it automatically. Normally nothing to do
- If another agent holds it, the command exits 2 immediately and shows the owner/idle time
- If the idle time is short, retry shortly. If the owner died or has been idle for 1 hour, it is reclaimed automatically
- Only force-release a lease that is clearly abandoned, with `vbaloop lease break --force`

### test timeout / modal / session death

- Compile-error/runtime-error modals are dismissed by the watchdog and only that test is failed
  (detail carries the modal text - that is the actual error message)
- Dismissing a compile error can make Excel exit (session death). Then test stops immediately and
  reports the remaining tests as `result: "skipped"` with `payload.sessionDied: true`
- **With 1 fail + many skipped, the first fail's detail (modal text) is the only real signal** -
  skipped just means not run; do not use it for root-causing. Fix the code, then push + test to revalidate.
  A hint like "rerun the same command" is about the session (a fresh Excel is ready), not a reason to rerun without fixing.
  If the hint says session start, do that before rerunning
- MsgBox/infinite loops are cut by the timeout (default 120 s, `--timeout <seconds>`). An ephemeral Excel with only the
  session workbook open is cleaned up by the tool and the next command starts a new one (hint: rerun the same command). A persistent
  session's Excel is quarantined if the tool cannot prove there is no user work (§Session quarantined / busy) - tell the user and wait for the release
- Do not put MsgBox/InputBox in test code - they hang headless runs

### push refused

- Non-CP949 character: look at file:line in errors and remove the character or replace it with ChrW
- Workbook module not in src: if the deletion is intended, `--prune`; otherwise pull it back first
- Document modules (ThisWorkbook/Sheet*) can only be replaced in place - they cannot be created
- Suspect the workbook and src have diverged (fixed it but same error repeats, push keeps returning `cached:true`):
  `vbaloop push --force --json` bypasses the fingerprint cache and reinjects everything
  (no need to delete `_dbg\sync.json` by hand). To see what is really there, `vbaloop diff --json`

### push compile gate

- In ephemeral mode, push runs a VBE compile after injection and before saving. On a compile error it **does not save** and exits 2 with
  `payload.compile = {result: "compile_error", detail: <modal text>}`. The canonical workbook on disk is in its pre-push state
- Fix the cause from detail (e.g. "Variable not defined", in Excel's UI language) and push again. The modal does not give the error
  location, so start with the modules you just changed. Also check `references\vba-pitfalls.md`
- If you are blocked by a pre-existing compile error unrelated to your change (legacy workbook), **after the user confirms** use `push --skip-compile`;
  to turn it off permanently the user sets `[push] compile = false` - do not turn the gate off yourself to get around it
- `result: "check_failed"` (could not verify) is an environment problem - `doctor --json`, then retry
- In a persistent session (persistent mode, session start/attach) the gate is skipped (`result: "skipped"` + reason).
  There, compile errors show up as test fail detail (modal text)

### Workbook state after a failed push

A failed push never saves - **the canonical workbook on disk is always in its pre-push state**. The in-memory workbook
state is given by `payload.recovery.workbookState` (and the error text):

- `discarded`: the ephemeral Excel was thrown away without saving. Fix the cause and retry
- `untouched`: nothing changed. Fix the cause and retry
- `restored`: every module this push touched was restored from a snapshot and verified. Fix the cause and retry
- `poisoned`: could not roll back (form re-import, prune, integrity failure, failed restore). The session is discarded and the next command
  restarts Excel **without saving** and reopens the canonical workbook from disk - no extra action; fix the cause and retry
- If you end the session right after a failure, `session stop --no-save` is safe (plain stop saves)

### Backups and rollback

- push leaves a copy of the workbook in `_dbg\backups\<id>-push\` right before injecting (skipped for a no-op push;
  newest 10 kept, toml `[backup] keep = N`, 0 disables)
- List: `vbaloop rollback --list --json`. Restore: `vbaloop rollback --latest --json` or `--backup <id>`
- **exit 3 while a persistent session is alive** - run `vbaloop session stop` first. Use `session stop --no-save` whenever the in-memory
  workbook may be bad (e.g. after a failed push) - plain stop saves it to disk. Ephemeral mode works directly
- Right before restoring, the current file is kept as a `pre-rollback` backup - the very next `rollback --latest` undoes it
- After a restore, src may differ from the workbook. **The next push reinjects all of src and overwrites the restore.**
  To adopt the restored workbook as src, `vbaloop pull --force` (plain pull is deliberately blocked); to compare, `vbaloop diff`

## pull caveats

- If there are unsynced src edits (fingerprint matches neither the last push nor the last pull), pull is blocked.
  The default response is to apply them first with `vbaloop push`. If the workbook is the truth, overwrite with `vbaloop pull --force`
- pull is only justified after a human edited directly in the session Excel
- If you only need to look, use `vbaloop diff` (read-only), not pull - pull rewrites src
- The first pull (no sync history) passes without the guard and produces a one-time format normalization (BOM/line endings) diff - expected

## Forms (.frm/.frx)

- Code-only changes: a normal push is enough (in place)
- A human edited a form by hand in the session (especially replacing a picture): `vbaloop push --force-form-import --json`
- If a .frx was replaced directly in src, push picks it up automatically
- The .frm text and the .frx are a pair - lint catches a missing partner

## Round-trip limits (impossible by design - work around them)

- Procedure-level Attributes (VB_ProcData macro shortcut keys) cannot round-trip and are lost on the first pull.
  Move shortcut keys to runtime registration with Application.OnKey
- Identifier casing can produce a one-time diff on re-pull after a push (stable after it converges - expected)
- Form pictures do not appear in the frm text - a manual replacement cannot be detected (use --force-form-import above)

## Distribution (dist)

`vbaloop dist --out <path> --json` builds a clean distribution copy (`.xlsm`/`.xlam`) - it copies the on-disk workbook as
the base and, in a temporary Excel (no window), applies module exclusion, version stamp, LogShim sanitizing and VBE compile check, then moves it atomically.
**push must come first** (refused when unsynced; the hint says to push, then retry). For the procedure, formats (xlsm/xlam), options and the full failure
playbook, read `references\dist.md` - it is a separate task from the edit loop, so it is not always loaded.

## VBA pitfall reference (look up / record)

- **Look up**: before writing new VBA, and before debugging an unexplained compile/runtime error,
  check `references\vba-pitfalls.md`. It is a short list, so skimming is cheap
- **Record**: if a pitfall not on the list cost you a debugging cycle, report it upstream. The project-local
  copy is a managed file, so direct edits are overwritten by the next init:
  1. Check for duplicates against existing entries, then draft it following the inclusion criteria and entry format at the top of the document
  2. Create a report with `vbaloop bugreport "<title>" --json` and put the draft in the body
     (on an installed copy, report it as a GitHub issue via payload.issueUrl - paste the body as is)
  3. In a development environment with the vbaloop source repo, add it to the canonical copy
     (`assets\vbaloop-skill\references\vba-pitfalls.md`) and rebuild
     (each project's next init --check detects the drift and converges)

## Reporting vbaloop's own defects

- If a vbaloop command has a bug or design violation (broken invariant, wrong verdict, data corruption),
  do not just work around it - file a report: `vbaloop bugreport "<title>" --json`
- Open the generated file at payload.path and fill in the TODO/parenthesized slots (severity/summary/evidence/repro).
  Evidence is the commands you ran and the raw --json output in code blocks. The exe build info is filled in automatically
- Duplicate check: `vbaloop bugreport --list --json` (resolved=false are open reports - for the same symptom,
  add evidence to the existing report instead of creating a new one)
- On an installed copy (no source repo) the report is created in `_dbg\bugreports\` and payload.issueUrl gives
  the GitHub issue link (https://github.com/rfastball/vbaloop/issues) - fill in the body, then file it as an issue

## Auxiliary commands

- `vbaloop doctor --json` - checks the Excel/COM/AccessVBOM environment (independent of the session). First stop when you suspect the environment.
  It also reports the quarantine state and windowless orphan EXCEL processes (whether each may be cleaned up, and why) - cleanup
  is `--kill-orphans` only after the user confirms (do not run it first yourself)
- `vbaloop diff --json` - read-only comparison of the workbook VBProject vs src. Reports a per-module verdict
  (same/changed/onlyInWorkbook/onlyInSrc/skipped) and the first differing line (firstDiffLine) in the payload.
  Always use this to see what is really in the workbook - direct COM access to the session Excel from PowerShell etc.
  is forbidden. `--module <name>` dumps the full workbook-side body of that module as
  payload.module.workbookBody. In ephemeral mode it uses a read-only throwaway Excel; in persistent mode it only works
  with an active session (no auto-start - exit 3 if none). Forms compare code-behind only (layout/.frx excluded). Exits 0 even when there are differences -
  read payload.clean
- `vbaloop lint --json` - full convention/encoding/header check (no Excel needed). A batch pre-check before push
- `vbaloop run <Module.Proc> --json` - runs one arbitrary macro. Exposes the return value (primitive types) and span logs.
  For debugging entry points that are not tests - assertion failures do not decide success
