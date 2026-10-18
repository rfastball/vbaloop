# VBA UserForm playbook

Read this before designing, generating or substantially editing an Excel MSForms UserForm.
Use the user's existing form language when available. These defaults improve consistency;
they cannot guarantee visual quality without opening the actual form in Excel.

## Start with an interaction contract

Before coordinates or code, decide: target data, main task, initial focus, Enter/Esc behavior,
success behavior (close or stay), failure behavior (keep input), and whether sheet interaction
must remain possible. A filter palette often needs modeless behavior; a one-shot confirmation
usually does not. Do not choose modeless merely to make an automated run return.

Sketch the reading order: target/context -> input -> options -> status -> actions. Keep a
frequent utility compact enough to leave useful worksheet context visible. Avoid a dashboard
header, ornamental cards, redundant preview pane or large empty areas for a one-input task.
Agree on control names and state before authoring; generated coordinates are an implementation
of that interaction, not the design itself.

## Layout and typography in MSForms

| Decision | Starting point, then verify in Excel |
|---|---|
| Spacing | Use a small shared scale, e.g. 6/12/18 pt. Align input edges and action baselines. |
| Text | One installed font that supports the UI language; start near 10-11 pt. Give Korean text room vertically. |
| Inputs/actions | Start near 22-26 pt high; multiline input grows with expected content, not form height. |
| Hierarchy | One primary action, quieter secondary action, persistent field labels. |
| States | Selected/disabled/error needs text or shape as well as color. Avoid faint text for essential options. |
| Long content | Reserve status height; wrap or truncate intentionally and expose the full target in ControlTipText. |

These numbers are defaults, not a universal style or a DPI conversion formula. Set font before
measuring text. AutoSize can help measure labels, but do not let it unexpectedly move aligned
columns. MSForms has no CSS flex/grid layout: keep static geometry in JSON and use one layout
routine only when runtime content actually requires it. Do not scatter Left/Top assignments
through Click, Change and Activate handlers.

JSON dimensions and MSForms control geometry use points. Exported FRM ClientWidth/ClientHeight
use a different serialization scale; do not paste them into JSON. Form Width/Height include
window chrome: check controls against the usable InsideWidth/InsideHeight. Frame children use
their parent's coordinate space, not the form's. Leave room for the frame caption and border.
[Microsoft: client dimensions](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/insideheight-insidewidth-properties).

Prefer RGB(r, g, b) in styling VBA; VBA &H color literals use BGR ordering, unlike CSS hex.
Set BackStyle deliberately for labels and avoid stacking borders, sunken effects and frames
around the same group. Font/color/TabIndex and other properties absent from the JSON schema
belong in Initialize; adding invented JSON keys will fail validation. Theme/layout changes
made at runtime do not appear in the unloaded designer, so inspect the displayed form.

## Keyboard first; custom styling second

Set TabIndex explicitly within each container and remove decorative controls from the tab
path where supported. Native CommandButtons supply focus and keyboard behavior; use them
by default. Labels used as buttons or toggles need an explicit keyboard equivalent and
visible focus/state feedback. A mouse-only toggle is not complete just because it looks good.

Choose Default and Cancel buttons deliberately. Enter can click a default button when another
CommandButton does not have focus; verify the actual input/control combination. Multiline
text needs an explicit MultiLine/EnterKeyBehavior policy so typing a new line never unexpectedly
applies changes. Do not steal Ctrl+Enter if it is needed for text entry.
[Microsoft: Click](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/click-event),
[EnterKeyBehavior](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/enterkeybehavior-property).

## Code golden rules

1. **The form owns interaction, the service owns work.** Event handlers validate, collect
   values, call the shared operation and render results. Keep worksheet algorithms, caching
   and business rules in standard/class modules usable without showing the form.
2. **State is data.** Use explicit Boolean/Enum fields; do not derive behavior from Caption,
   color or a checkmark glyph. Have one RenderState/UpdateUI routine derive presentation.
   Small Click handlers change state and call it; avoid a generic UI framework for one form.
3. **Initialize is not dependency injection.** It runs after loading and before showing.
   Set styles/defaults there; caller-supplied targets may arrive later. Use an explicit
   Configure method/properties before Show. Activate may run repeatedly: do not reset user
   input or rerun expensive work there. Refresh context deliberately when reusing a form.
   [Microsoft: Initialize](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/initialize-event).
4. **Bind the operation to its intended workbook/range.** Especially for modeless forms,
   inject the target; do not resolve ActiveWorkbook/Selection again on Apply. Revalidate
   that the target still exists and is usable when executing. A displayed target name must
   match the object passed to the engine.
5. **Own lifetime explicitly.** For modeless/reusable forms use an explicit instance and an
   owner reference; define reuse identity, close and release paths. Hide retains state;
   Unload tears down the loaded UI. Handle title-bar X as well as Cancel. Remove registry
   references during the explicit close/QueryClose path; do not rely solely on Terminate
   to break references that themselves keep an instance alive.
6. **Guard programmatic events and reentry locally.** Use an mUpdating flag around assignments
   that trigger Change handlers and an mBusy flag for Apply. Excel Application.EnableEvents
   is not an MSForms event guard. Restore guards on every error path; disabling one visible
   button alone does not guard keyboard or alternate entry points. Prefer Repaint for wait
   feedback; DoEvents can reenter handlers and needs a concrete reason and a busy guard.
7. **Restore what you changed.** Save prior Cursor/ScreenUpdating/etc. and restore on success
   and failure. Avoid changing application-wide state for simple UI updates. Preserve input
   and focus on validation failure; show actionable error text. Use On Error Resume Next
   only around a specific expected lookup/cleanup, then restore normal error handling.
8. **Favor generated design-time controls.** Names in JSON become named controls with ordinary
   event procedures. Runtime Controls.Add requires separate event wiring (typed WithEvents
   handlers retained for their lifetime); naming a Sub alone does not wire a runtime control.
   Do not add that machinery for a fixed form. Avoid control names that shadow form members.

## Lessons from UnivFilter (conditional patterns)

Source inspected: user's `UnivFIlter/src/frmFilter01.frm` and `src/modUnivFilter.bas`,
2026-10-18. These are code observations, not a visual audit of its binary FRX layout.
The external project is inspiration, not a runtime dependency or a required template.

| Observed pattern | Reuse when / caveat |
|---|---|
| pIsExcludeMode/pIsStackMode/pIsPartialMatch -> UpdateUI | Good separation of state and presentation; keep captions out of business logic. |
| SetLabelStyle + label Click/MouseMove | Compact flat segmented choices. Add keyboard equivalents; native controls are cheaper when this styling is unnecessary. |
| Visible btnApply/btnCancel moved to Left=9000 with Default/Cancel | A keyboard bridge for styled label actions. Advanced workaround, not a default: remove invisible tab stops, guard both action paths, verify Enter/Esc in every input. Visible=False is not an equivalent keyboard contract. |
| ResetHoverEffects on UserForm_MouseMove | MSForms lacks a simple universal mouse-leave styling hook. Test movement across neighboring controls/frames too; hover must never be the only state signal. Avoid logging MouseMove. |
| Short header plus full ControlTipText | Keeps a small utility readable; retain enough visible target context to avoid acting on the wrong column. |
| InjectBodyRange/InjectFilterArea -> modFilterRun.RunFilter | Share one engine between the form and shortcut path. Do not copy the engine into the form. |
| GetFilterForm registry keyed by target header address/field | Reuse the right modeless instance; reinject current ranges on reuse. Explicitly test X/Cancel/reopen and release, not just successful Apply. |
| Position/focus Win32 helpers | Useful for a worksheet-adjacent palette, but not needed for a basic form. Reuse a proven helper; verify 32/64-bit declarations, DPI/multiple monitors and focus behavior before adopting. |

Do not copy project-specific global state, APIs or cleanup code wholesale. In particular,
last-used mode persistence is a product decision; resetting the cursor to xlDefault is not
the same as restoring its prior value. The playbook recommends improving these boundaries
when adapting the pattern, without changing the reference project.

## Create a UserForm from scratch

Write two UTF-8 files outside `src` (for example `forms/frmHello.json` and
`forms/frmHello.vba`). A pre-existing form or FRX is not required. Excel and VBOM
access are required; the tool does not change security settings.

```json
{
  "name": "frmHello", "caption": "Hello", "width": 280, "height": 160,
  "controls": [
    {"type":"TextBox", "name":"txtName", "text":"World", "left":12, "top":12, "width":220, "height":24},
    {"type":"CommandButton", "name":"cmdHello", "caption":"Greet", "left":12, "top":48, "width":90, "height":24},
    {"type":"Label", "name":"lblResult", "caption":"", "left":12, "top":84, "width":220, "height":18}
  ]
}
```

```vb
Option Explicit
Private Sub cmdHello_Click()
    Me.lblResult.Caption = "Hello, " & Me.txtName.Text
End Sub
```

```powershell
vbaloop form build --spec forms/frmHello.json --code forms/frmHello.vba --json
vbaloop push --json
```

- Inputs resolve from cwd; output goes to the configured src root. `--config` is supported.
- Supported types: Label, TextBox, CommandButton, Frame. Only Frame can contain nested
  `controls`. Names must be unique across the form (case-insensitive), start with an ASCII
  letter, contain only letters/digits/underscore and be at most 40 characters.
- Form requires name/width/height; each control requires type/name/left/top/width/height.
  Optional caption is supported except on TextBox, which accepts optional text instead.
  Geometry is in points: dimensions positive, coordinates non-negative, all at most 10000.
  Unknown/duplicate fields and text outside the system ACP are rejected.
- Code is a plain form body, without VERSION/Begin export envelopes or Attribute lines.
  Event names must match control names. To display the form, add a standard-module Sub
  containing `frmHello.Show`; modal forms wait for the user to close them.
- Build uses private Excel to verify exported/imported properties and code, then writes
  UTF-8 frm + binary frx. It does not open/save the target workbook or execute the code.
  Full-project compilation occurs in push's existing compile gate.
- Rebuilding existing source requires `form build ... --force`. Both generated files are
  replaced from JSON/VBA; Excel/frm edits are not merged. Keep JSON/VBA as authoring inputs
  and rebuild before push after editing them. Push alone reads frm/frx, not JSON/VBA;
  pull does not update the authoring inputs.
- Publication errors restore replaced files. If restoration fails, the hint points to
  retained backups: restore the pair before push. The two-file update is not crash-atomic;
  after an interrupted build, rebuild successfully before pushing.
- Pictures, MultiPage, TabStrip, arbitrary ActiveX and extra properties are not supported.


## Acceptance loop

Do not confuse compilation with design review.

1. Check reading order, control bounds, labels and named event handlers before building.
   Keep authoring JSON/VBA together; engine and launcher stay in normal modules.
2. Build -> push -> focused existing tests. Build checks serialization, not event behavior;
   compilation does not check appearance. Add only missing meaningful behavior coverage,
   e.g. invalid input leaves data unchanged and valid input reaches the intended target.
3. Open the actual form using the project launcher. Inspect initial, selected, empty, error
   and success states with realistic Korean/long text. Check clipping, alignment and usable
   worksheet context. Check other DPI/monitor environments when those are required.
4. Walk Tab/Shift+Tab, Enter, Esc and title-bar X. For modeless forms, switch worksheets and
   workbooks, reactivate, close and reopen; verify target, input preservation and cleanup.
   Exercise every custom label action without a mouse.
5. Report separately what compiled, what behavior was tested and what was visually inspected.
   If UI inspection was unavailable or interrupted, state that limit. JSON, FRM text and HTML
   mockups cannot establish that the actual MSForms layout looks correct.

## Maintaining this reference

This file is managed by init. Improve the canonical
`assets/vbaloop-skill/references/forms.md` in the vbaloop repository; rebuild and use init to
propagate it. Keep project-specific design decisions outside managed files. Add a trick only
with its use case, keyboard/lifetime costs and observed verification, not as an untested
universal rule.
