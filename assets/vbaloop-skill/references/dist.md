# vbaloop dist - building a distribution artifact

`vbaloop dist --out <path> --json` builds a clean distribution copy (`.xlsm` or `.xlam`). It copies the on-disk
workbook as the base -> in its own temporary Excel (no window) applies (module exclusion -> version stamp -> LogShim sanitizing
-> optional comment stripping -> VBE compile check -> optional bootstrap) -> produces the output by an atomic move. The session workbook
is not touched (no lease needed). dist does not change the workbook, so it is not subject to backup/rollback.

## Contract

- **push must come first**: dist uses the on-disk workbook where "last push == src" as the base. When unsynced it
  refuses (`ok:false`, hint tells you to run `vbaloop push`, then retry). **Order: (edit ->) push -> dist.**
- The format is decided by the `--out` extension: `.xlsm` (macro workbook) / `.xlam` (add-in - IsAddin set automatically, saved as
  xlam). Without `--out`: `<root>\dist\<workbook name>.xlsm`. `dist\` is ignored by the init-managed .gitignore
  (build artifact = not committed).
- Exclusions: primarily `' @dist exclude` in the module **header** (ignored if written in the body - header only),
  supplemented by toml `[dist] exclude` globs (`Test_*` etc.). Exclusion is per module only.
- Version: toml `[dist] version` (semver). The git rev is appended automatically - **outside a git repo rev=`nogit`**
  (the tool never stamps its own rev). Written to both BUILD_* constants and CustomDocumentProperties.
- Options: `--strip-comments` (removes comments, protects `'` inside strings), `--bootstrap <Module.Proc>` (runs worksheet
  setup after injection - modals isolated by the kill watchdog), `--skip-compile` (skips the compile check - escape hatch).
- Compile check: runs VBE "Compile VBAProject" before publishing. The result is `payload.compile.result`
  (`compiled` / `already_compiled` / `unverified` (published with a warning) / `compile_error` / `check_failed`).
  For the last two nothing is published (`payload.published: false`) and the existing distribution copy is left as is.

## Failure playbook

- **sync out of date (most common)**: as the hint says, `vbaloop push`, then retry. **Watch out especially right after a
  vbaloop update** - when the new exe's `init` re-vendors `src\LogShim.bas` etc., the src fingerprint no longer matches.
  Keep the order **update -> init -> push -> dist** (init is filesystem-only, so it cannot push).
- **dangling refs (compile-safety failure)**: a remaining module references a symbol of an excluded module. Look at the
  `module:line` entries in payload.danglingRefs and narrow the exclusion or remove the reference (Test modules are only safe to drop wholesale if they reference only each other).
- **LogShim IS_RELEASE missing (CliError)**: violates the distribution sanitizing contract - it fails loudly to prevent a distribution copy
  that silently has logging on. `src\LogShim.bas` has drifted from the managed copy, so restore the canonical copy with
  `vbaloop init` (if it still reproduces, it is a vbaloop defect - `vbaloop bugreport`).
- **compile error (`compile_error`)**: the error text carries the VBE dialog text (e.g. "Variable not defined",
  in Excel's UI language). Fix it in src and push -> dist. If the cause is a reference to an excluded module's symbol, handle it like dangling refs.
  If test passed but only this fails, it is an error in the body of a procedure tests do not reach - it will still blow up in the distribution copy.
- **compile check not possible (`check_failed`)**: control missing / execution failed / timeout. Check
  AccessVBOM with `doctor --json`. If it must be published without the check, `--skip-compile` (tell a human).
- **bootstrap timeout (120 s)**: the bootstrap procedure hung on a MsgBox/modal. The temporary Excel is
  reclaimed by kill (no harm to the session). Make that procedure non-interactive and compile-clean.
- **document/sheet modules cannot be excluded**: ignored with a warning (cannot be Removed). If you really need to exclude it, rethink the design.
- **No `doctor --kill-orphans` (while dist runs)**: dist's temporary Excel has no window, so doctor's orphan scan
  picks it up, and doctor does not take the mutex, so it can kill dist. Do not use `--kill-orphans` while dist is running.
