Attribute VB_Name = "LogShim"
' @desc Writes debug logs and test assertions to jsonl
Option Explicit

'============================================================
' [vbaloop-managed vendored module] Do not edit. Not self-verifying.
' If this shows up in diff/status/drift, it was not "my edit" -
' vbaloop init refreshed it to the canonical copy. Verify outside the file:
'   git diff              (against the committed copy)
'   vbaloop init --check  (byte compare with the embedded canonical copy)
' The canonical copy lives in vbaloop, not here. Once confirmed, commit as is.
'============================================================

' Self-contained version shipped by vbaloop. No add-in dependency.
' Path: <Workbook.Path>\_dbg\<workbook base name>.jsonl
' Record schema: kind D/J/A/X (unchanged). Existing call-site signatures unchanged.
' Writes are true appends (Binary Access Read Write + LOF+1).

Private Const DBG_FOLDER As String = "_dbg"
' Rolling line cap. Trimmed oldest-first only on Reset.
Private Const MAX_LINES As Long = 5000

' Release neutralization switch (design 9.3). dist flips it to True only in the injected text,
' turning the test/debug sinks into no-ops. Canonical copy and src are always False (the
' artifact is a build output) - never set it to True by hand.
' Call sites (LogShim.LogD ...) stay in place and still compile; runtime cost is one constant branch.
' Exception: error-class events (IsErrorEvent) are still recorded in release builds to keep a
' field diagnostic trail (bug report 2026-07-10). LogD/AssertEq/Reset are fully silent in release.
Private Const IS_RELEASE As Boolean = False

' @desc Records level and msg as a D record
Public Sub LogD(level As String, msg As String)
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    AppendLine "{""ts"":""" & NowTimestamp() & """,""kind"":""D"",""level"":""" & _
               JsonEscape(level) & """,""msg"":""" & JsonEscape(msg) & """}"
Done:
End Sub

' @desc Records eventName and dataJson as a J record
Public Sub LogJ(eventName As String, dataJson As String)
    If IS_RELEASE And Not IsErrorEvent(eventName) Then Exit Sub
    On Error GoTo Done
    Dim dataPart As String
    If LenB(dataJson) = 0 Then dataPart = "null" Else dataPart = dataJson
    AppendLine "{""ts"":""" & NowTimestamp() & """,""kind"":""J"",""event"":""" & _
               JsonEscape(eventName) & """,""data"":" & dataPart & "}"
Done:
End Sub

' @desc Builds JSON from event name and key/value arrays, records via LogJ
Public Sub LogEvent(eventName As String, keys As Variant, vals As Variant)
    If IS_RELEASE And Not IsErrorEvent(eventName) Then Exit Sub
    On Error GoTo Done
    Dim json As String, i As Long
    json = "{"
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then json = json & ","
        json = json & """" & JsonEscape(CStr(keys(i))) & """:" & JsonVal(vals(i))
    Next i
    json = json & "}"
    LogJ eventName, json
Done:
End Sub

' @desc Records an assertion result as an A record; never raises on failure
Public Sub AssertEq(name As String, expected As Variant, actual As Variant)
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    Dim okFlag As Boolean
    okFlag = (CStr(expected) = CStr(actual))
    AppendLine "{""ts"":""" & NowTimestamp() & """,""kind"":""A"",""name"":""" & _
               JsonEscape(name) & """,""ok"":" & IIf(okFlag, "true", "false") & _
               ",""expected"":""" & JsonEscape(CStr(expected)) & _
               """,""actual"":""" & JsonEscape(CStr(actual)) & """}"
Done:
End Sub

' @desc Trims jsonl to the cap and appends a reset marker
Public Sub Reset()
    If IS_RELEASE Then Exit Sub
    On Error GoTo Done
    Dim path As String
    path = LogPath()
    If LenB(path) = 0 Then Exit Sub
    EnsureDbgFolder path

    Dim marker As String
    marker = "{""ts"":""" & NowTimestamp() & """,""kind"":""X"",""event"":""reset""}"

    Dim existing As String
    If LenB(Dir(path)) > 0 Then existing = ReadFileUtf8(path)
    If Len(existing) = 0 Then
        WriteFileUtf8 path, marker & vbLf
        Exit Sub
    End If

    Dim lines() As String
    lines = Split(existing, vbLf)
    Dim keep As Long
    keep = MAX_LINES - 1
    If keep < 0 Then keep = 0

    Dim startIdx As Long, i As Long, sb As String
    Dim total As Long
    total = UBound(lines) - LBound(lines) + 1
    If total > keep Then
        startIdx = UBound(lines) - keep + 1
        If startIdx < LBound(lines) Then startIdx = LBound(lines)
    Else
        startIdx = LBound(lines)
    End If
    For i = startIdx To UBound(lines)
        If Len(lines(i)) > 0 Then sb = sb & lines(i) & vbLf
    Next i
    WriteFileUtf8 path, sb & marker & vbLf
Done:
End Sub

' === Internals ===

' @desc Whether an error-class event must be recorded even in release (field diagnostics)
Private Function IsErrorEvent(eventName As String) As Boolean
    ' Whitelist of operational error events only. Add run_start/run_end here if needed.
    Select Case eventName
        Case "cmd_error", "fatal_error", "param_guard"
            IsErrorEvent = True
    End Select
End Function

Private Function LogPath() As String
    ' Unsaved workbook (no Path): silently ignored
    If LenB(ThisWorkbook.Path) = 0 Then Exit Function
    LogPath = ThisWorkbook.Path & "\" & DBG_FOLDER & "\" & BaseName(ThisWorkbook.Name) & ".jsonl"
End Function

Private Function BaseName(fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then BaseName = Left$(fileName, p - 1) Else BaseName = fileName
End Function

Private Sub EnsureDbgFolder(path As String)
    Dim folder As String
    folder = Left$(path, InStrRev(path, "\") - 1)
    If LenB(Dir(folder, vbDirectory)) = 0 Then MkDir folder
End Sub

Private Sub AppendLine(line As String)
    Dim path As String
    path = LogPath()
    If LenB(path) = 0 Then Exit Sub
    EnsureDbgFolder path

    Dim bytes() As Byte
    bytes = Utf8Bytes(line & vbLf)

    Dim f As Integer
    f = FreeFile
    On Error GoTo CleanFail
    Open path For Binary Access Read Write As #f

    Dim pos As Long
    pos = LOF(f) + 1
    ' Avoid gluing onto the last line when an external tool left no trailing newline
    If pos > 1 Then
        Dim lastByte As Byte
        Get #f, pos - 1, lastByte
        If lastByte <> 10 Then
            Dim lfByte(0 To 0) As Byte
            lfByte(0) = 10
            Put #f, pos, lfByte
            pos = pos + 1
        End If
    End If
    Put #f, pos, bytes
    Close #f
    Exit Sub

CleanFail:
    Dim savedNum As Long, savedDesc As String
    savedNum = Err.Number: savedDesc = Err.Description
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Err.Raise savedNum, "LogShim.AppendLine", savedDesc
End Sub

Private Function Utf8Bytes(s As String) As Byte()
    Dim t As Object
    Set t = CreateObject("ADODB.Stream")
    t.Type = 2
    t.Charset = "utf-8"
    t.Open
    t.WriteText s
    t.Position = 0
    t.Type = 1
    t.Position = 3 ' skip the 3-byte UTF-8 BOM
    Utf8Bytes = t.Read
    t.Close
End Function

Private Function ReadFileUtf8(path As String) As String
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2
    s.Charset = "utf-8"
    s.Open
    s.LoadFromFile path
    Dim text As String
    text = s.ReadText
    s.Close
    If Len(text) > 0 Then
        If Left$(text, 1) = ChrW$(&HFEFF&) Then text = Mid$(text, 2)
    End If
    ReadFileUtf8 = text
End Function

Private Sub WriteFileUtf8(path As String, content As String)
    Dim t As Object, b As Object
    Set t = CreateObject("ADODB.Stream")
    t.Type = 2
    t.Charset = "utf-8"
    t.Open
    t.WriteText content

    Set b = CreateObject("ADODB.Stream")
    b.Type = 1
    b.Open
    t.Position = 3 ' skip the 3-byte UTF-8 BOM
    t.CopyTo b
    t.Close
    b.SaveToFile path, 2 ' adSaveCreateOverWrite
    b.Close
End Sub

Private Function NowTimestamp() As String
    Dim ms As Long
    ms = CLng((Timer - Int(Timer)) * 1000) Mod 1000
    NowTimestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format$(ms, "000")
End Function

' @desc JSON representation by value type (escape strings only; numbers/booleans raw)
Private Function JsonVal(v As Variant) As String
    Select Case VarType(v)
        Case vbBoolean
            JsonVal = IIf(v, "true", "false")
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency
            JsonVal = CStr(v)
        Case Else
            JsonVal = """" & JsonEscape(CStr(v)) & """"
    End Select
End Function

Private Function JsonEscape(s As String) As String
    Dim i As Long, ch As String, code As Long, result As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34:      result = result & "\"""
            Case 92:      result = result & "\\"
            Case 8:       result = result & "\b"
            Case 9:       result = result & "\t"
            Case 10:      result = result & "\n"
            Case 12:      result = result & "\f"
            Case 13:      result = result & "\r"
            Case 0 To 31: result = result & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else:    result = result & ch
        End Select
    Next i
    JsonEscape = result
End Function
