# Third-party notices

vbaloop includes or adapts material from the following projects.

## xlflow

Parts of vbaloop's Excel COM plumbing were adapted from xlflow
(https://github.com/harumiWeb/xlflow):

- `src/VbaLoop/Excel/OleMessageFilter.cs` — COM `IMessageFilter` retry for `RPC_E_CALL_REJECTED`
  (from `ExcelBridgeSupport.OleMessageFilter`)
- `src/VbaLoop/Excel/ExcelAttach.cs` — attaching to an Excel process by window handle
  (top tier of `ExcelBridgeSupport.TryGetExcelByProcessId`)
- `src/VbaLoop/Cli/CliArgs.cs` — detecting `--json` from raw argv so parse errors keep the JSON envelope
  (from `rawArgsRequestJSON`)

Other places that mention xlflow in comments describe design comparisons, not copied code.

```
MIT License

Copyright (c) 2026 harumiWeb

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## .NET runtime

Release builds of `vbaloop.exe` are self-contained and bundle the .NET runtime
(https://github.com/dotnet/runtime), licensed under the MIT License,
Copyright (c) .NET Foundation and Contributors.
