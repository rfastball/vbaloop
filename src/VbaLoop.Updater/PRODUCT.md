# vbaloop updater

<!-- impeccable:product-schema 1 -->

## Platform

Windows desktop, Korean UI. WPF on the repository's .NET 10 runtime.

## Users and purpose

Internal-network users receive a txt VBA package and apply it to an existing business workbook.
They choose files, review the target/version/module changes, create an updated copy, and see its location.
The original is never opened by Excel or overwritten by the GUI. Inspection and editing use a verified working copy;
Apply creates the verified working copy in a new per-run result folder before Excel opens it, then uses Save in place.
No SaveAs or post-save move/copy is needed. Completion opens the result rather than the original.
Copies are removed on failures before saving; once saving starts, they are preserved even if the outcome is uncertain.
The approved UI is the light, emerald-accented four-state mockup from this conversation.

## Constraints

- A single self-contained executable embeds the complete CLI as a managed DLL built from the same source revision.
  A separate updater process hosts the verified DLL with the updater's bundled runtime, avoiding a second runtime bundle.
- No CLI installation, PATH registration, network calls, security-setting changes, or project setup.
- Desktop Excel and existing VBA project access are required. The app explains missing prerequisites.
- The CLI owns backup, Excel, module changes, compile checks, and saving. The GUI owns interaction.
- Failed or incomplete compilation cannot save through the GUI. An uncertain save is never called success.
- No editor, project manager, downgrade override, automatic restore, or automatic retry.
- Local profile backups follow the existing CLI policy; persistent VDI storage is not yet specified.

## Accessibility

Keyboard-focusable controls, visible focus, scalable native text, scrollable content, and textual status.
