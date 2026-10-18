## Development loop (vbaloop)

The build/test loop runs through `vbaloop` (a single exe). If it is not on PATH, the install path is
`%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe` (if even that is missing, reinstall with `vbaloop setup`).
Onboarding/drift repair: `vbaloop init` (check only: `vbaloop init --check`).
If init --check exits 3 with `exe-older` (the managed_by version in vbaloop.toml is newer than the running exe), do not fix it
with init; ask the user to run `vbaloop update` (do not edit or delete the managed_by line).
Write project-specific content outside this managed block (init overwrites the inside of the block).

1. Edit: change src\*.bas/.cls/.frm
   New UserForm: write JSON and plain VBA outside src, then
   `vbaloop form build --spec forms/MyForm.json --code forms/MyForm.vba --json` -> push.
   No template or hand-written FRX is needed. Before form work, read
   `.claude/skills/vbaloop/references/forms.md` for design, code rules and schema.
   Rebuild with --force only to replace the generated pair from those inputs.
2. Session: ephemeral by default - push/pull/test/run/diff start a hidden Excel per command and close it (vbaloop.toml
   [excel] session). If a session opened with session start/attach is alive, it is used
3. Inject: vbaloop push (src -> in-place injection into the workbook -> saved only if the VBE compile passes; supports --prune/--force.
   Backup to _dbg\backups\ right before injecting)
4. Run: vbaloop test (runs every Test_* + reads the LogShim jsonl, --timeout <seconds>)
5. Fix: on fail, fix the cause and go back to 1; when green you are done (commit now if you use version control)

- Every command supports --json: {command, ok, exitCode, hint, errors[], payload}
- exit code: success 0 / environment error 3 (session/Excel/COM/busy/quarantined - follow the hint) / otherwise 2
- Quarantined (hint mentions recovery clear / --discard): the tool could not safely clean up Excel, so it stopped instead of killing it.
  Never kill Excel on your own initiative - that includes running --discard or doctor --kill-orphans before the user says so.
  Tell the user; if the user closes that Excel, vbaloop recovery clear; if they say it may be discarded, vbaloop session stop --discard
- busy (exit 3): Excel is in cell edit mode or showing a dialog - retry the same command shortly
- If the user already has the canonical workbook open in Excel (read-only refusal hint): after the user confirms,
  vbaloop session attach - push/test then work in that Excel and session stop only detaches
- Tests are Public Sub Test_*() (no arguments); assert with LogShim.AssertEq (no Debug.Assert)
- src\LogShim.bas is a vendored copy - do not edit directly; update with vbaloop init
- When push fails, the hint states the workbook state (untouched / fully restored / discarded / session poisoned - not saved).
  After a failed push the canonical workbook on disk is always in its pre-push state.
  On a compile error (payload.compile.result = compile_error), fix it and push again.
  Exception: if the error is pre-existing and unrelated to this change, ask the user; only after they confirm, push --skip-compile.
  If you fixed something but the same error repeats and push keeps returning cached:true, push --force bypasses the fingerprint cache.
  To revert the workbook file itself, vbaloop rollback --latest (list: rollback --list;
  exit 3 while a persistent session is alive - session stop first)
- In ephemeral mode test/run open the canonical workbook read-only - sheets changed by tests are not saved
- Distribution artifact: vbaloop dist --out <path> (clean copy build - filter/stamp/sanitize/VBE compile check/optional bootstrap).
  push must come first (refused when last push != src). Procedure/failure playbook: SKILL.md
- pull is blocked when there are unsynced src edits - by default apply them with push first;
  only when the workbook is the truth, overwrite with pull --force
- No direct COM access to the session Excel (GetActiveObject, New-Object -ComObject etc. - including reads/diagnostics).
  To see what is really in the workbook, vbaloop diff --json (read-only; --module <name> dumps a single module)
- src conventions: UTF-8 without BOM + CRLF + exactly one trailing newline. Characters outside CP949 (em dash U+2014, emoji) are refused by push
- If you find a defect in vbaloop itself: create a standard report with vbaloop bugreport "<title>" and fill it in
  (procedure: .claude\skills\vbaloop\SKILL.md)
- Detailed procedures/failure playbooks: .claude\skills\vbaloop\SKILL.md (managed by vbaloop init - do not edit directly)
