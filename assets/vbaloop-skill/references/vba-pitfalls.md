# VBA pitfall reference (for agents)

Traps agents hit when writing/debugging VBA. Skim before writing new code and when facing an unexplained
compile/runtime error. This file is a vbaloop managed copy - do not edit directly;
report additions with `vbaloop bugreport` or a GitHub issue (https://github.com/rfastball/vbaloop/issues)
(recording procedure: SKILL.md section "VBA pitfall reference").

**Inclusion criteria** (at least one, real incidents only):
(a) an LLM gets it plausibly wrong from other-language intuition
(b) it fails silently/late and burns debugging cycles
(c) it is specific to headless automation sessions
**Entry format**: `- **Title**: symptom/cause in 1-2 lines. Correct pattern.`
**Excluded**: general style advice, anything already in the CLAUDE.md managed block or SKILL.md (link only).

## 1. Specific to headless sessions

- **Event macros do not run**: tool sessions open the workbook with `EnableEvents=False` -
  initialization that relies on `Workbook_Open`/`Worksheet_Change` etc. is silently skipped.
  In tests, replace event-driven logic with explicit calls to the procedure.
- **Application state left behind after an error exit**: VBA has no finally. If you exit on an error with
  `ScreenUpdating=False`, `Calculation=xlManual` or `EnableEvents=False` still set, then because vbaloop keeps
  sessions alive the next test starts in a polluted state. Always restore in the error handler.
- **Relying on Selection/ActiveSheet**: an automation session has no selection/active state made by a user.
  An unqualified `Range("A1")` binds to ActiveSheet and writes to the wrong sheet in multi-sheet code.
  Always qualify: `Worksheets("Name").Range(...)`.
- **Debug.Print is invisible headless**: Immediate window output cannot be read.
  Logs you need to see in the terminal go through `LogShim.LogD`/`LogJ`. Assert with `LogShim.AssertEq`.
- **No breakpoints/Stop/VBE debug mode**: they need manual intervention and hang headless runs.
  Do not leave `Stop` statements in source. `Erl` returns 0 in code without line numbers.
- **File dialogs hang too**: not just MsgBox/InputBox (see SKILL) -
  `Application.GetOpenFilename`/`GetSaveAsFilename`/`FileDialog` and `UserForm.Show` are also
  forbidden on headless execution paths - they are only cut by the timeout and Excel is usually quarantined (needs a user release). Inject paths as arguments/constants.

## 2. VBA language traps (cross-language intuition betrayed)

- **And/Or do not short-circuit**: `If x Is Nothing Or x.Count = 0 Then` evaluates `x.Count` even when x is Nothing
  -> runtime error 91. Split the guard into nested Ifs: `If x Is Nothing Then Exit Sub`.
- **In Dim i, j As Long, i is a Variant**: every declarator needs its own `As <type>`.
  Use `Dim i As Long, j As Long`.
- **No C-style escapes**: `\"` and `\\` are literal characters, not escapes.
  Double a quote inside a string: `"He said ""hi"""`. Path backslashes stay as they are: `"C:\temp"`.
- **Smart quotes sneaking in**: Unicode quotes from generated/pasted text are syntax errors in VBA.
  Straight ASCII quotes only.
- **Integer is 16-bit**: overflows at 32,767 - a row count alone can exceed it.
  Default to `Long` for integers.
- **For Each control variables must be Variant/object-compatible**: `Dim c As Long` then
  `For Each c In coll` fails. Use `Dim c As Variant` (or the element's object type).
- **The parentheses in `Foo (arg)` are not call syntax**: they force the argument to be evaluated by value (defeating ByRef) or
  evaluate an object's default member. Call procedures as `Foo arg` or `Call Foo(arg)`.
- **Me does not exist in every module kind**: valid only in Sheet/ThisWorkbook/class/form modules, not in standard
  .bas modules. Check which module type the code is for first.
- **Blanket On Error Resume Next swallows errors**: failures propagate silently and the cause cannot be
  pinned down. Use a narrow recovery block followed by `On Error GoTo 0`, or a label handler
  (`Exit Sub` before the label). `Resume` only inside the handler block.
- **Without Option Explicit a typo becomes a new variable**: put `Option Explicit` at the top of every module.
  Catching undeclared assignments as compile errors is actually cheaper headless (avoids runtime debugging).
- **Line continuation `_` needs a space before it**: `arg1,_` is a syntax error; `arg1, _` is right.
- **End statements must match the block kind**: closing a Sub with `End Function` is a compile error -
  expensive headless because it goes through a compile modal. Pre-check with lint before push.
- **Do not pass objects to AssertEq**: objects like Range blow up on default-member evaluation / error 91 or compare
  differently than intended. Pass scalars: `AssertEq "name", expected, ws.Range("A1").Value2`.
- **Module-level declarations may only come before the first procedure**: putting `Private m_x As Boolean`
  between procedures to keep it next to related code is a compile error "Only comments may appear after
  End Sub, End Function, or End Property". Collect all variable/constant/Declare/Event declarations in the
  declarations section at the top of the module and mark ownership with comments only.

## 3. Covered by other documents (links only)

- Characters outside CP949, no Debug.Assert, module-qualified calls, no commits before tests are green:
  CLAUDE.md managed block
- MsgBox/InputBox timeouts, round-trip limits (VB_ProcData loss, one-time casing diff, form pictures),
  poisoned/stale recovery: SKILL.md
