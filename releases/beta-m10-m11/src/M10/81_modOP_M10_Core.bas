Attribute VB_Name = "modOP_M10_Core"
Option Explicit

Private Const OP_M10_VERSION As String = "BETA-M10-1.0.0"
Private Const OP_M10_MARKER As String = "WEBARCH_OP_M10"
Private Const OP_M10_META As String = "OP_M10_META"
Private Const OP_M10_HOME As String = "OP_HOME"
Private Const OP_M10_SUMMARY As String = "OP_SINTESI"
Private Const OP_M10_REPORT As String = "OP_REPORT"
Private Const OP_M10_MARKER_CELL As String = "AA1"

Public Function OP_M10_Version() As String
    OP_M10_Version = OP_M10_VERSION
End Function

Public Function OP_M10_SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    OP_M10_SheetExists = Not ws Is Nothing
End Function

Public Function OP_M10_PreflightInstall(ByRef detail As String) As Boolean
    Dim names As Variant
    Dim i As Long
    Dim ws As Worksheet

    detail = vbNullString
    names = Array(OP_M10_HOME, OP_M10_SUMMARY, OP_M10_REPORT, OP_M10_META)

    For i = LBound(names) To UBound(names)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(names(i)))
        On Error GoTo 0

        If Not ws Is Nothing Then
            If CStr(ws.Range(OP_M10_MARKER_CELL).Value2) <> OP_M10_MARKER Then
                detail = "Collisione: il foglio '" & CStr(names(i)) & _
                         "' esiste ma non appartiene a M10. Nessuna modifica eseguita."
                Exit Function
            End If
        End If
    Next i

    OP_M10_PreflightInstall = True
End Function

Public Function OP_M10_GetOrCreateOwnedSheet(ByVal sheetName As String, _
                                              ByVal visibleState As XlSheetVisibility, _
                                              ByRef detail As String) As Worksheet
    Dim ws As Worksheet

    detail = vbNullString
    On Error GoTo EH

    If OP_M10_SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        If CStr(ws.Range(OP_M10_MARKER_CELL).Value2) <> OP_M10_MARKER Then
            detail = "Il foglio '" & sheetName & "' non e' di M10."
            Exit Function
        End If
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        ws.Range(OP_M10_MARKER_CELL).NumberFormat = "@"
        ws.Range(OP_M10_MARKER_CELL).Value2 = OP_M10_MARKER
        ws.Columns("AA:AA").Hidden = True
    End If

    ws.Visible = visibleState
    Set OP_M10_GetOrCreateOwnedSheet = ws
    Exit Function

EH:
    detail = "Creazione/verifica foglio '" & sheetName & "' non riuscita [" & _
             CStr(Err.Number) & "]: " & Err.Description
End Function

Public Sub OP_M10_WriteLiteral(ByVal target As Range, ByVal value As Variant)
    If target Is Nothing Then Exit Sub
    target.NumberFormat = "@"
    If IsError(value) Then
        target.Value2 = "#ERRORE"
    ElseIf IsEmpty(value) Then
        target.Value2 = vbNullString
    Else
        target.Value2 = value
    End If
End Sub

Private Function OP_M10_ConfigSheet() As Worksheet
    Dim ws As Worksheet
    Dim hitRun As Range
    Dim hitCommit As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set OP_M10_ConfigSheet = ws
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        Set hitRun = Nothing
        Set hitCommit = Nothing
        On Error Resume Next
        Set hitRun = ws.UsedRange.Find(What:="RunID", After:=ws.Cells(1, 1), _
                                       LookIn:=xlValues, LookAt:=xlPart, _
                                       SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                       MatchCase:=False)
        Set hitCommit = ws.UsedRange.Find(What:="CommitID", After:=ws.Cells(1, 1), _
                                          LookIn:=xlValues, LookAt:=xlPart, _
                                          SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                          MatchCase:=False)
        On Error GoTo 0

        If Not hitRun Is Nothing And Not hitCommit Is Nothing Then
            Set OP_M10_ConfigSheet = ws
            Exit Function
        End If
    Next ws
End Function

Public Function OP_M10_GetConfigText(ByVal addressText As String, _
                                     Optional ByVal defaultValue As String = "N/D") As String
    Dim ws As Worksheet
    Dim v As Variant

    On Error GoTo Fallback
    Set ws = OP_M10_ConfigSheet()
    If ws Is Nothing Then GoTo Fallback

    v = ws.Range(addressText).Value2

    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        OP_M10_GetConfigText = defaultValue
    Else
        OP_M10_GetConfigText = CStr(v)
    End If
    Exit Function

Fallback:
    OP_M10_GetConfigText = defaultValue
End Function

Private Function OP_M10_TextAtom(ByVal v As Variant) As String
    Dim s As String
    Dim n As Long

    If IsError(v) Then
        OP_M10_TextAtom = "#ERR"
        Exit Function
    End If
    If IsEmpty(v) Then
        OP_M10_TextAtom = "<EMPTY>"
        Exit Function
    End If

    s = CStr(v)
    n = Len(s)
    If n = 0 Then
        OP_M10_TextAtom = "<BLANK>"
    ElseIf n <= 12 Then
        OP_M10_TextAtom = s
    Else
        OP_M10_TextAtom = CStr(n) & ":" & Left$(s, 4) & ":" & _
                          Mid$(s, (n \ 2), 4) & ":" & Right$(s, 4)
    End If
End Function

Private Function OP_M10_AtomScore(ByVal s As String) As Double
    Dim i As Long
    Dim score As Double
    Dim codeValue As Long

    score = Len(s) * 17#
    For i = 1 To Len(s)
        codeValue = AscW(Mid$(s, i, 1))
        If codeValue < 0 Then codeValue = codeValue + 65536
        score = score + (CDbl(codeValue) * CDbl((i Mod 31) + 1))
        If score > 2000000000# Then
            score = score - Fix(score / 2147483629#) * 2147483629#
        End If
    Next i
    OP_M10_AtomScore = score
End Function

Private Function OP_M10_FingerprintRange(ByVal rng As Range) As String
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    Dim h As Double
    Dim atom As String
    Dim rowsCount As Long
    Dim colsCount As Long

    On Error GoTo EH

    rowsCount = rng.Rows.Count
    colsCount = rng.Columns.Count
    data = rng.Value2
    h = 146959#

    If rowsCount = 1 And colsCount = 1 Then
        atom = OP_M10_TextAtom(data)
        h = h * 131# + OP_M10_AtomScore(atom) + 1#
        If h > 2147483629# Then h = h - Fix(h / 2147483629#) * 2147483629#
    Else
        For r = 1 To rowsCount
            For c = 1 To colsCount
                atom = OP_M10_TextAtom(data(r, c))
                h = h * 131# + OP_M10_AtomScore(atom) + CDbl(r) * 7# + CDbl(c) * 13#
                If h > 2147483629# Then h = h - Fix(h / 2147483629#) * 2147483629#
            Next c
        Next r
    End If

    OP_M10_FingerprintRange = CStr(rowsCount) & "x" & CStr(colsCount) & ":" & Format$(h, "0")
    Exit Function

EH:
    OP_M10_FingerprintRange = "ERR:" & CStr(Err.Number)
End Function

Private Function OP_M10_FindDataRange(ByVal objectName As String) As Range
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastCell As Range

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(objectName)
    If Not ws Is Nothing Then
        Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
        If lastCell Is Nothing Then
            Set OP_M10_FindDataRange = ws.Range("A1")
        Else
            Set OP_M10_FindDataRange = ws.Range(ws.Cells(1, 1), _
                                                ws.Cells(lastCell.Row, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column))
        End If
        On Error GoTo 0
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(objectName)
        If Not lo Is Nothing Then
            Set OP_M10_FindDataRange = lo.Range
            On Error GoTo 0
            Exit Function
        End If
    Next ws

    On Error GoTo 0
End Function

Public Function OP_M10_ObjectFingerprint(ByVal objectName As String) As String
    Dim rng As Range

    Set rng = OP_M10_FindDataRange(objectName)
    If rng Is Nothing Then
        OP_M10_ObjectFingerprint = "MISSING"
    Else
        OP_M10_ObjectFingerprint = OP_M10_FingerprintRange(rng)
    End If
End Function

Public Function OP_M10_ObjectRowCount(ByVal objectName As String) As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rng As Range
    Dim lastCell As Range

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(objectName)
        If Not lo Is Nothing Then
            If lo.DataBodyRange Is Nothing Then
                OP_M10_ObjectRowCount = 0
            Else
                OP_M10_ObjectRowCount = lo.DataBodyRange.Rows.Count
            End If
            On Error GoTo 0
            Exit Function
        End If
    Next ws

    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(objectName)
    If Not ws Is Nothing Then
        Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
        If lastCell Is Nothing Or lastCell.Row <= 1 Then
            OP_M10_ObjectRowCount = 0
        Else
            OP_M10_ObjectRowCount = lastCell.Row - 1
        End If
    End If

    On Error GoTo 0
End Function

Public Function OP_M10_SourceToken() As String
    Dim runId As String
    Dim commitId As String

    runId = OP_M10_GetConfigText("B18", "N/D")
    commitId = OP_M10_GetConfigText("B136", "N/D")

    OP_M10_SourceToken = _
        "RUN=" & runId & _
        "|COMMIT=" & commitId & _
        "|PROT=" & OP_M10_ObjectFingerprint("DB_PROTOCOLLI_CORRENTI") & _
        "|DEC=" & OP_M10_ObjectFingerprint("DB_DECRETI") & _
        "|INC=" & OP_M10_ObjectFingerprint("DB_INCIDENTI") & _
        "|FSC=" & OP_M10_ObjectFingerprint("FSC_DB_CORRENTI")
End Function

Private Function OP_M10_MetaSheet(ByRef detail As String) As Worksheet
    Dim ws As Worksheet

    detail = vbNullString
    If Not OP_M10_SheetExists(OP_M10_META) Then
        detail = "M10 non installato: manca OP_M10_META."
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(OP_M10_META)
    If CStr(ws.Range(OP_M10_MARKER_CELL).Value2) <> OP_M10_MARKER Then
        detail = "OP_M10_META non appartiene a M10."
        Exit Function
    End If

    Set OP_M10_MetaSheet = ws
End Function

Public Sub OP_M10_SetMeta(ByVal keyName As String, ByVal value As Variant)
    Dim ws As Worksheet
    Dim detail As String
    Dim f As Range
    Dim targetRow As Long

    Set ws = OP_M10_MetaSheet(detail)
    If ws Is Nothing Then Err.Raise vbObjectError + 6101, "OP_M10_SetMeta", detail

    Set f = ws.Columns(1).Find(What:=keyName, After:=ws.Cells(1, 1), _
                               LookIn:=xlValues, LookAt:=xlWhole, _
                               SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                               MatchCase:=False)
    If f Is Nothing Then
        targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If targetRow < 2 Then targetRow = 2
        OP_M10_WriteLiteral ws.Cells(targetRow, 1), keyName
        OP_M10_WriteLiteral ws.Cells(targetRow, 2), value
    Else
        OP_M10_WriteLiteral ws.Cells(f.Row, 2), value
    End If
End Sub

Public Function OP_M10_GetMeta(ByVal keyName As String, _
                               Optional ByVal defaultValue As String = "") As String
    Dim ws As Worksheet
    Dim detail As String
    Dim f As Range

    Set ws = OP_M10_MetaSheet(detail)
    If ws Is Nothing Then
        OP_M10_GetMeta = defaultValue
        Exit Function
    End If

    Set f = ws.Columns(1).Find(What:=keyName, After:=ws.Cells(1, 1), _
                               LookIn:=xlValues, LookAt:=xlWhole, _
                               SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                               MatchCase:=False)
    If f Is Nothing Then
        OP_M10_GetMeta = defaultValue
    Else
        OP_M10_GetMeta = CStr(ws.Cells(f.Row, 2).Value2)
    End If
End Function

Public Sub OP_M10_MarkAligned(ByVal refreshStatus As String, ByVal refreshDetail As String)
    OP_M10_SetMeta "VERSION", OP_M10_VERSION
    OP_M10_SetMeta "ALIGNED_SOURCE_TOKEN", OP_M10_SourceToken()
    OP_M10_SetMeta "ALIGNED_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    OP_M10_SetMeta "REFRESH_STATUS", refreshStatus
    OP_M10_SetMeta "REFRESH_DETAIL", refreshDetail
    OP_M10_SetMeta "SRC_FP_DB_DECRETI", OP_M10_ObjectFingerprint("DB_DECRETI")
    OP_M10_SetMeta "SRC_FP_DB_INCIDENTI", OP_M10_ObjectFingerprint("DB_INCIDENTI")
    OP_M10_SetMeta "SRC_FP_FSC_DB_CORRENTI", OP_M10_ObjectFingerprint("FSC_DB_CORRENTI")
End Sub

Public Function OP_M10_IsCurrent(ByRef detail As String) As Boolean
    Dim currentToken As String
    Dim alignedToken As String
    Dim refreshStatus As String

    detail = vbNullString
    currentToken = OP_M10_SourceToken()
    alignedToken = OP_M10_GetMeta("ALIGNED_SOURCE_TOKEN", "")
    refreshStatus = UCase$(Trim$(OP_M10_GetMeta("REFRESH_STATUS", "")))

    If Len(alignedToken) = 0 Then
        detail = "Nessun riallineamento M10 registrato."
        Exit Function
    End If

    If StrComp(currentToken, alignedToken, vbBinaryCompare) <> 0 Then
        detail = "La base sorgente e' cambiata dopo l'ultimo riallineamento M10."
        Exit Function
    End If

    If refreshStatus <> "OK" Then
        detail = "L'ultimo riallineamento operativo non e' certificabile come OK: " & refreshStatus
        Exit Function
    End If

    OP_M10_IsCurrent = True
End Function

Private Function OP_M10_ShapeCaption(ByVal shp As Shape) As String
    Dim s As String

    On Error Resume Next
    s = shp.TextFrame.Characters.Text
    If Len(Trim$(s)) = 0 Then s = shp.TextFrame2.TextRange.Text
    If Len(Trim$(s)) = 0 Then s = shp.AlternativeText
    On Error GoTo 0

    OP_M10_ShapeCaption = Trim$(s)
End Function

Public Function OP_M10_InvokeReloadAction(ByVal sheetName As String, _
                                          ByVal required As Boolean, _
                                          ByRef detail As String) As Boolean
    Dim ws As Worksheet
    Dim shp As Shape
    Dim caption As String
    Dim actionName As String
    Dim found As Boolean
    Dim previousSheet As Object

    detail = vbNullString
    On Error GoTo EH

    If Not OP_M10_SheetExists(sheetName) Then
        If required Then
            detail = "Manca la pagina richiesta: " & sheetName
            Exit Function
        Else
            OP_M10_InvokeReloadAction = True
            detail = sheetName & ": assente, ignorata."
            Exit Function
        End If
    End If

    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set previousSheet = ActiveSheet

    For Each shp In ws.Shapes
        caption = UCase$(OP_M10_ShapeCaption(shp))
        actionName = vbNullString
        On Error Resume Next
        actionName = shp.OnAction
        On Error GoTo EH

        If Len(actionName) > 0 Then
            If InStr(1, caption, "RICARICA", vbTextCompare) > 0 Or _
               InStr(1, caption, "RIALLINEA", vbTextCompare) > 0 Then
                found = True
                Application.Run actionName
                Exit For
            End If
        End If
    Next shp

    On Error Resume Next
    If Not previousSheet Is Nothing Then previousSheet.Activate
    On Error GoTo EH

    If Not found Then
        If required Then
            detail = "Nessuna azione RICARICA/RIALLINEA verificabile in " & sheetName & _
                     ". M10 si arresta invece di chiamare macro indovinate."
            Exit Function
        Else
            detail = sheetName & ": nessuna azione di ricarica esposta; nessuna macro invocata."
        End If
    Else
        detail = sheetName & ": ricarica invocata."
    End If

    OP_M10_InvokeReloadAction = True
    Exit Function

EH:
    detail = sheetName & ": errore durante la ricarica [" & CStr(Err.Number) & "]: " & Err.Description
End Function

Public Function OP_M10_CountFlags(ByVal sheetName As String) As Long
    Dim ws As Worksheet
    Dim rng As Range
    Dim data As Variant
    Dim r As Long
    Dim c As Long
    Dim s As String
    Dim countValue As Long

    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set rng = ws.UsedRange
    data = rng.Value2

    If rng.Rows.Count = 1 And rng.Columns.Count = 1 Then
        s = UCase$(CStr(data))
        If s = "DIVERGED" Or s = "SOURCE_ABSENT" Or s = "DA_RIVEDERE" Or s = "N/D" Then countValue = 1
    Else
        For r = 1 To rng.Rows.Count
            For c = 1 To rng.Columns.Count
                If Not IsError(data(r, c)) And Not IsEmpty(data(r, c)) Then
                    s = UCase$(Trim$(CStr(data(r, c))))
                    If s = "DIVERGED" Or s = "SOURCE_ABSENT" Or _
                       s = "DA_RIVEDERE" Or s = "N/D" Then
                        countValue = countValue + 1
                    End If
                End If
            Next c
        Next r
    End If

Done:
    OP_M10_CountFlags = countValue
End Function

Public Function OP_M10_OwnedMarker() As String
    OP_M10_OwnedMarker = OP_M10_MARKER
End Function
