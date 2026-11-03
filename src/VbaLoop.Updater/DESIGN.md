---
name: vbaloop updater
description: Korean native desktop interface for offline VBA package updates.
colors:
  primary: "#087346"
  content: "#FFFFFF"
  chrome: "#F6F8F7"
  inset: "#F3F6F4"
  text: "#192321"
  secondary-text: "#57685F"
  separator: "#E2E8E4"
  button-text: "#273D35"
  button-border: "#CCD5D0"
  focus: "#153FBD"
  drop-surface: "#FAFCFB"
  drop-border: "#B7C7BF"
  error-surface: "#FFF7ED"
  error-text: "#804312"
typography:
  headline:
    fontFamily: "Malgun Gothic"
    fontSize: "26px"
    fontWeight: 600
    lineHeight: "36px"
  version:
    fontFamily: "Segoe UI"
    fontSize: "30px"
    fontWeight: 600
  section:
    fontFamily: "Malgun Gothic"
    fontSize: "14px"
    fontWeight: 400
  body:
    fontFamily: "Malgun Gothic"
    fontSize: "14px"
    fontWeight: 400
    lineHeight: "22px"
  caption:
    fontFamily: "Malgun Gothic"
    fontSize: "13px"
    fontWeight: 400
    lineHeight: "20px"
rounded:
  control: "6px"
  drop-target: "8px"
spacing:
  compact: "12px"
  section: "22px"
  section-large: "28px"
  content-inline: "40px"
components:
  button-primary:
    backgroundColor: "{colors.primary}"
    textColor: "{colors.content}"
    rounded: "{rounded.control}"
    padding: "11px 18px"
  button-secondary:
    backgroundColor: "{colors.content}"
    textColor: "{colors.button-text}"
    rounded: "{rounded.control}"
    padding: "11px 18px"
  drop-target:
    backgroundColor: "{colors.drop-surface}"
    rounded: "{rounded.drop-target}"
    padding: "24px"
  change-details:
    backgroundColor: "{colors.inset}"
    padding: "12px"
  error-notice:
    backgroundColor: "{colors.error-surface}"
    textColor: "{colors.error-text}"
    rounded: "{rounded.control}"
    padding: "20px"
---

# Design System: vbaloop updater

## Overview

The updater uses a light native Windows window, Korean system typography, and one emerald accent.
The chosen workbook and version change receive the strongest emphasis before the user applies an update.
This scoped specification is extracted from `MainWindow.xaml` and its code-behind; root `DESIGN.md`
remains the CLI architecture contract. Token lengths use px for interchange and represent WPF
device-independent pixels, not physical display pixels.

**Key Characteristics:**

- A single window changes between selection, inspection, progress, completion, and error states.
- Generous vertical spacing separates the task from its optional details.
- One primary footer action stays visible while the content scrolls.
- Status uses explicit text as well as color.

The approved direction is the native desktop mockup `exec-09db5019-23f3-4d25-9917-a523503fe8bf.png`.
Its ready-screen composition is retained: heading, side-by-side versions, workbook, package,
expandable changes, backup notice, and the bottom-right action. The user approved the system fonts
and compact status line above the heading. No invented brand metaphor or generated design seed applies.

## Colors

The primary emerald identifies the next action, new version, positive status, and vector marks.
White content and quiet gray chrome establish structure; separators and inset details remain subdued.
Secondary text supports paths and explanations. Warm error surfaces and text distinguish recovery
instructions. Keyboard focus uses the separate blue focus token. Values in the frontmatter are normative.

## Typography

Malgun Gothic is the Korean interface face. Segoe UI distinguishes version numerals and the wordmark.
The headline, version, section, body, and caption roles are listed in the frontmatter. Both actual
file names use 16-DIP semibold text with 24-DIP leading in selection and inspection. Empty-file hints
use 14-DIP normal secondary text; emphasis appears only after selection. Progress labels use
18-DIP semibold text with 28-DIP leading. The status line is 12-DIP normal; field labels are 14-DIP
normal. Completion details use 16/26 and error explanations use the body 14/22 role. Malgun Gothic
has no dedicated installed semibold face, so hierarchy relies on size, tone, and spacing as well as weight.
Most text wraps. The workbook path is a single ellipsized line with the complete path in a tooltip.

## Layout

The window requests 800 × 720 DIP, with a 620 × 540 minimum. Startup height is limited to the
working-area height minus 48 DIP. A 58-DIP brand row sits above a vertically scrolling body and
fixed footer. Body margins are 40 horizontal, 30 top, and 28 bottom. Footer padding is 32 horizontal
and 18 vertical. The version comparison has equal side columns and a 60-DIP arrow column.
There are no web breakpoints or mobile layouts. Optional details scroll inside a 190-DIP maximum.

## Elevation & Depth

The authored interface uses tonal surfaces and one-DIP separators without custom shadows.
The operating system owns the titlebar, outside window shadow, and native file dialogs.
No authored transition animation exists; the standard WPF indeterminate progress animation is used.

## Shapes

Buttons and error notices have modest control corners; the drop target uses the larger corner token.
The brand loop, arrow, and completion check are XAML vector geometry. No raster image, remote
font, or icon-font dependency ships with the interface. The executable and window icon use the same
loop in white on the emerald rounded square in `Assets/updater.svg`. `Assets/updater.ico` contains
16, 20, 24, 32, 40, 48, 64, 128, and 256-pixel frames for Windows titlebar, taskbar, and Explorer sizes.

## Components

- **Buttons:** minimum height 42 DIP; primary footer minimum width 176 DIP; selection buttons 140 DIP wide. Secondary actions are
  outlined. Hover, pressed, and disabled surface opacity is 0.83, 0.65, and 0.4 respectively.
  Keyboard focus changes the border to two DIP in the focus color. State changes are immediate.
- **Selection:** one inset group contains separate package and workbook rows, each with a native
  file-picker button. Both inputs are visible from the start and may be chosen in either order.
  Selected filenames and parent paths are ellipsized with full-path tooltips; buttons change to
  “변경”. Drag/drop remains available across the window. The footer names the missing input and
  enables “적용 내용 확인” only when both files are present. Selection never starts inspection
  automatically; the explicit footer action opens the read-only inspection before the apply screen.
- **Inspection:** versions and file names lead. Change details expand to read-only selectable text.
  Warnings name deletions or retained sheet code. First adoption uses an unchecked confirmation box.
- **Progress:** real CLI phase labels accompany an indeterminate bar; no percentage is invented.
  The ordinary close action is blocked during work. A slow-save notice permits closing only after
  explicit confirmation that the background save continues.
- **Completion:** a restrained outlined check, target/version, workbook name, primary open-file
  action, and applicable backup-folder action. The already-applied state explicitly reports no change.
- **Errors:** a warm inline notice gives the next action. Details expand and can be copied.
  An uncertain save offers folder/backup inspection instead of an automatic retry.

## Do's and Don'ts

- **Do** keep the target workbook and version change visible before application.
- **Do** use actual stage names and distinguish confirmed success from an uncertain save.
- **Do** retain selectable details and full-path tooltips for longer values.
- **Don't** introduce a project sidebar, editor, or promotional artwork into this updater.
- **Don't** replace explicit status text with color-only indicators or invented progress percentages.

Parent visual inspection found the ready and completion views legible and fitting their window.
The reviewed timeout/save-state corrections passed source review. Screen-reader behavior, full keyboard
traversal, Windows high-contrast mode, and the complete display-scale/long-content matrix remain
unverified; the presence of focus styling and live-setting metadata is not a conformance claim.
The labeled `--preview` states use illustrative data with file-writing actions disabled.
