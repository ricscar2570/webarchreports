Attribute VB_Name = "modOP_M10_Core"
Option Explicit

Private Const OP_M10_VERSION As String = "BETA-M10.1-AUDIT-REMEDIATION-1.0.0"
Private Const OP_M10_MARKER As String = "WEBARCH_OP_M10"
Private Const OP_M10_META As String = "OP_M10_META"
Private Const OP_M10_HOME As String = "OP_HOME"
Private Const OP_M10_SUMMARY As String = "OP_SINTESI"
Private Const OP_M10_REPORT As String = "OP_REPORT"
Private Const OP_M10_MARKER_CELL As String = "AA1"

Private Const OP_M10_FP_PREFIX As String = "FP2:"
Private Const OP_M10_GATE_OK As String = "OK"
Private Const OP_M10_GATE_BLOCKED As String = "BLOCKED"

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

Public Function OP_M10_IsOwnedSheet(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Worksheets(sheetName)
    OP_M10_IsOwnedSheet = (CStr(ws.Range(OP_M10_MARKER_CELL).Value2) = OP_M10_MARKER)

SafeExit:
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
    On Error GoTo EH
    Set OP_M10_ConfigSheet = ThisWorkbook.Worksheets(SH_CONFIG)
    Exit Function
EH:
    Set OP_M10_ConfigSheet = Nothing
End Function

Public Function OP_M10_GetConfigText(ByVal addressText As String, _
                                      Optional ByVal fallback As String = "N/D") As String
    Dim ws As Worksheet
    Dim v As Variant

    On Error GoTo FallbackValue
    Set ws = OP_M10_ConfigSheet()
    If ws Is Nothing Then GoTo FallbackValue

    v = ws.Range(addressText).Value2
    If IsError(v) Then GoTo FallbackValue

    If Len(Trim$(CStr(v))) = 0 Then
        OP_M10_GetConfigText = fallback
    Else
        OP_M10_GetConfigText = Trim$(CStr(v))
    End If
    Exit Function

FallbackValue:
    OP_M10_GetConfigText = fallback
End Function

Public Function OP_M10_CheckDiscard(ByRef detail As String) As Boolean
    detail = vbNullString
    On Error GoTo EH

    If MINIMAL_DiscardRequired() Then
        detail = MINIMAL_DiscardMessage()
        Exit Function
    End If

    OP_M10_CheckDiscard = True
    Exit Function
EH:
    detail = "Impossibile verificare lo stato di sicurezza MINIMAL [" & _
             CStr(Err.Number) & "]: " & Err.Description
End Function

Private Function OP_M10_FindListObject(ByVal objectName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(objectName)
        If Not lo Is Nothing Then
            Set OP_M10_FindListObject = lo
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

Private Function OP_M10_FindDataRange(ByVal objectName As String, _
                                      ByRef rng As Range, _
                                      ByRef detail As String) As Boolean
    Dim lo As ListObject
    Dim ws As Worksheet
    Dim lastRowCell As Range
    Dim lastColCell As Range

    detail = vbNullString
    Set rng = Nothing
    On Error GoTo EH

    Set lo = OP_M10_FindListObject(objectName)
    If Not lo Is Nothing Then
        Set rng = lo.Range
        OP_M10_FindDataRange = True
        Exit Function
    End If

    If Not OP_M10_SheetExists(objectName) Then
        detail = "Oggetto '" & objectName & "' non trovato."
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(objectName)
    Set lastRowCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    Set lastColCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
                                    MatchCase:=False)

    If lastRowCell Is Nothing Then
        Set rng = ws.Range("A1")
    ElseIf lastColCell Is Nothing Then
        Set rng = ws.Range("A1")
    Else
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRowCell.Row, lastColCell.Column))
    End If

    OP_M10_FindDataRange = True
    Exit Function

EH:
    detail = "Lettura oggetto '" & objectName & "' non riuscita [" & _
             CStr(Err.Number) & "]: " & Err.Description
End Function

Private Sub OP_M10_HashText(ByVal s As String, ByRef h1 As Double, ByRef h2 As Double)
    Const P1 As Double = 2147483629#
    Const P2 As Double = 2147483587#
    Dim i As Long
    Dim codePoint As Long

    For i = 1 To Len(s)
        codePoint = AscW(Mid$(s, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536

        h1 = h1 * 257# + CDbl(codePoint + 1)
        h1 = h1 - Fix(h1 / P1) * P1

        h2 = h2 * 263# + CDbl(codePoint + 3)
        h2 = h2 - Fix(h2 / P2) * P2
    Next i

    h1 = h1 * 257# + 31#
    h1 = h1 - Fix(h1 / P1) * P1
    h2 = h2 * 263# + 37#
    h2 = h2 - Fix(h2 / P2) * P2
End Sub

Private Function OP_M10_NumberText(ByVal v As Variant) As String
    Dim s As String
    s = Format$(CDbl(v), "0.###############################")
    If Application.DecimalSeparator <> "." Then
        s = Replace(s, Application.DecimalSeparator, ".")
    End If
    OP_M10_NumberText = s
End Function

Private Function OP_M10_SerializeValue(ByVal v As Variant) As String
    Dim vt As VbVarType

    On Error GoTo EH

    If IsError(v) Then
        OP_M10_SerializeValue = "E"
        Exit Function
    End If

    If IsEmpty(v) Then
        OP_M10_SerializeValue = "Z"
        Exit Function
    End If

    vt = VarType(v)

    Select Case vt
        Case vbString
            OP_M10_SerializeValue = "S" & CStr(Len(CStr(v))) & ":" & CStr(v)
        Case vbBoolean
            If CBool(v) Then
                OP_M10_SerializeValue = "B1"
            Else
                OP_M10_SerializeValue = "B0"
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            OP_M10_SerializeValue = "N:" & OP_M10_NumberText(v)
        Case vbDate
            OP_M10_SerializeValue = "D:" & OP_M10_NumberText(CDbl(CDate(v)))
        Case vbNull
            OP_M10_SerializeValue = "NULL"
        Case Else
            OP_M10_SerializeValue = "V" & CStr(CLng(vt)) & ":" & CStr(v)
    End Select
    Exit Function

EH:
    OP_M10_SerializeValue = "SERERR:" & CStr(Err.Number)
End Function

Private Function OP_M10_FingerprintRange(ByVal rng As Range) As String
    Dim values As Variant
    Dim r As Long
    Dim c As Long
    Dim h1 As Double
    Dim h2 As Double
    Dim serial As String
    Dim rowCount As Long
    Dim colCount As Long

    On Error GoTo EH

    h1 = 146959810#
    h2 = 216613626#
    rowCount = rng.Rows.Count
    colCount = rng.Columns.Count

    OP_M10_HashText "ROWS=" & CStr(rowCount) & ";COLS=" & CStr(colCount), h1, h2

    values = rng.Value2

    If rowCount = 1 And colCount = 1 Then
        serial = OP_M10_SerializeValue(values)
        If Left$(serial, 7) = "SERERR:" Then GoTo SerializationError
        OP_M10_HashText serial, h1, h2
    Else
        For r = 1 To rowCount
            For c = 1 To colCount
                serial = OP_M10_SerializeValue(values(r, c))
                If Left$(serial, 7) = "SERERR:" Then GoTo SerializationError
                OP_M10_HashText "R" & CStr(r) & "C" & CStr(c) & ":" & serial, h1, h2
            Next c
        Next r
    End If

    OP_M10_FingerprintRange = OP_M10_FP_PREFIX & CStr(rowCount) & "x" & CStr(colCount) & ":" & _
                              Format$(h1, "0") & ":" & Format$(h2, "0")
    Exit Function

SerializationError:
    OP_M10_FingerprintRange = "ERR:SERIAL"
    Exit Function

EH:
    OP_M10_FingerprintRange = "ERR:" & CStr(Err.Number)
End Function

Public Function OP_M10_FingerprintIsValid(ByVal fingerprint As String) As Boolean
    fingerprint = Trim$(fingerprint)
    If Len(fingerprint) <= Len(OP_M10_FP_PREFIX) Then Exit Function
    If Left$(fingerprint, Len(OP_M10_FP_PREFIX)) <> OP_M10_FP_PREFIX Then Exit Function
    If InStr(1, fingerprint, "ERR:", vbTextCompare) > 0 Then Exit Function
    If InStr(1, fingerprint, "MISSING", vbTextCompare) > 0 Then Exit Function
    OP_M10_FingerprintIsValid = True
End Function

Public Function OP_M10_ObjectFingerprint(ByVal objectName As String) As String
    Dim rng As Range
    Dim detail As String

    If Not OP_M10_FindDataRange(objectName, rng, detail) Then
        OP_M10_ObjectFingerprint = "MISSING:" & objectName
        Exit Function
    End If

    OP_M10_ObjectFingerprint = OP_M10_FingerprintRange(rng)
End Function

Public Function OP_M10_ObjectRowCount(ByVal objectName As String) As Long
    Dim lo As ListObject
    Dim loSheet As ListObject
    Dim ws As Worksheet
    Dim lastCell As Range
    Dim bestRows As Long
    Dim currentRows As Long

    On Error GoTo EH

    Set lo = OP_M10_FindListObject(objectName)
    If Not lo Is Nothing Then
        If lo.DataBodyRange Is Nothing Then
            OP_M10_ObjectRowCount = 0
        Else
            OP_M10_ObjectRowCount = lo.DataBodyRange.Rows.Count
        End If
        Exit Function
    End If

    If Not OP_M10_SheetExists(objectName) Then
        OP_M10_ObjectRowCount = -1
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(objectName)

    If ws.ListObjects.Count > 0 Then
        bestRows = -1
        For Each loSheet In ws.ListObjects
            If loSheet.DataBodyRange Is Nothing Then
                currentRows = 0
            Else
                currentRows = loSheet.DataBodyRange.Rows.Count
            End If
            If currentRows > bestRows Then bestRows = currentRows
        Next loSheet

        If bestRows < 0 Then bestRows = 0
        OP_M10_ObjectRowCount = bestRows
        Exit Function
    End If

    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                 MatchCase:=False)
    If lastCell Is Nothing Then
        OP_M10_ObjectRowCount = 0
    ElseIf lastCell.Row <= 1 Then
        OP_M10_ObjectRowCount = 0
    Else
        OP_M10_ObjectRowCount = lastCell.Row - 1
    End If
    Exit Function

EH:
    OP_M10_ObjectRowCount = -1
End Function

Public Function OP_M10_SheetTableRowCount(ByVal sheetName As String, _
                                           ByRef ok As Boolean) As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rowsNow As Long
    Dim rowsBest As Long

    ok = False
    On Error GoTo EH

    If Not OP_M10_SheetExists(sheetName) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(sheetName)

    If ws.ListObjects.Count = 0 Then Exit Function

    rowsBest = -1
    For Each lo In ws.ListObjects
        If lo.DataBodyRange Is Nothing Then
            rowsNow = 0
        Else
            rowsNow = lo.DataBodyRange.Rows.Count
        End If
        If rowsNow > rowsBest Then rowsBest = rowsNow
    Next lo

    If rowsBest < 0 Then rowsBest = 0
    OP_M10_SheetTableRowCount = rowsBest
    ok = True
    Exit Function

EH:
    ok = False
End Function

Private Function OP_M10_ConfigSignature(ByRef detail As String) As String
    Dim ws As Worksheet
    Dim fp As String
    Dim sig As String
    Dim upperName As String
    Dim includeSheet As Boolean

    detail = vbNullString
    sig = "CFG2"

    On Error GoTo EH

    fp = OP_M10_ObjectFingerprint(SH_CONFIG)
    If Not OP_M10_FingerprintIsValid(fp) Then
        detail = "Configurazione alpha non verificabile: " & SH_CONFIG & " = " & fp
        Exit Function
    End If
    sig = sig & "|ALPHA_CFG=" & fp

    For Each ws In ThisWorkbook.Worksheets
        upperName = UCase$(ws.Name)
        includeSheet = False

        If upperName = "FSC_CONFIG" Then includeSheet = True
        If upperName = "OP_FIELD_DEFS" Then includeSheet = True
        If InStr(1, upperName, "_MAP", vbTextCompare) > 0 Then includeSheet = True
        If InStr(1, upperName, "MAPPING", vbTextCompare) > 0 Then includeSheet = True
        If InStr(1, upperName, "_CONFIG", vbTextCompare) > 0 Then includeSheet = True
        If InStr(1, upperName, "_DEFS", vbTextCompare) > 0 Then includeSheet = True

        If upperName = UCase$(OP_M10_META) Then includeSheet = False
        If upperName = UCase$(OP_M10_HOME) Then includeSheet = False
        If upperName = UCase$(OP_M10_SUMMARY) Then includeSheet = False
        If upperName = UCase$(OP_M10_REPORT) Then includeSheet = False

        If includeSheet Then
            fp = OP_M10_ObjectFingerprint(ws.Name)
            If Not OP_M10_FingerprintIsValid(fp) Then
                detail = "Fingerprint configurazione non valido: " & ws.Name & " = " & fp
                Exit Function
            End If
            sig = sig & "|" & ws.Name & "=" & fp
        End If
    Next ws

    OP_M10_ConfigSignature = sig
    Exit Function

EH:
    detail = "Impossibile calcolare la firma configurazioni [" & CStr(Err.Number) & "]: " & Err.Description
End Function

Public Function OP_M10_TrySourceToken(ByRef token As String, ByRef detail As String) As Boolean
    Dim fpProt As String
    Dim fpDec As String
    Dim fpInc As String
    Dim fpFsc As String
    Dim cfgSig As String
    Dim runId As String
    Dim commitId As String

    token = vbNullString
    detail = vbNullString

    runId = OP_M10_GetConfigText("B18", "N/D")
    commitId = OP_M10_GetConfigText("B136", "N/D")

    If runId = "N/D" Then
        detail = "RunID non disponibile nel foglio " & SH_CONFIG & "."
        Exit Function
    End If
    If commitId = "N/D" Then
        detail = "CommitID non disponibile nel foglio " & SH_CONFIG & "."
        Exit Function
    End If

    fpProt = OP_M10_ObjectFingerprint("DB_PROTOCOLLI_CORRENTI")
    fpDec = OP_M10_ObjectFingerprint("DB_DECRETI")
    fpInc = OP_M10_ObjectFingerprint("DB_INCIDENTI")
    fpFsc = OP_M10_ObjectFingerprint("FSC_DB_CORRENTI")

    If Not OP_M10_FingerprintIsValid(fpProt) Then
        detail = "DB_PROTOCOLLI_CORRENTI non verificabile: " & fpProt
        Exit Function
    End If
    If Not OP_M10_FingerprintIsValid(fpDec) Then
        detail = "DB_DECRETI non verificabile: " & fpDec
        Exit Function
    End If
    If Not OP_M10_FingerprintIsValid(fpInc) Then
        detail = "DB_INCIDENTI non verificabile: " & fpInc
        Exit Function
    End If
    If Not OP_M10_FingerprintIsValid(fpFsc) Then
        detail = "FSC_DB_CORRENTI non verificabile: " & fpFsc
        Exit Function
    End If

    cfgSig = OP_M10_ConfigSignature(detail)
    If Len(cfgSig) = 0 Then Exit Function

    token = "SRC2|VER=" & OP_M10_VERSION & _
            "|RUN=" & runId & _
            "|COMMIT=" & commitId & _
            "|PROT=" & fpProt & _
            "|DEC=" & fpDec & _
            "|INC=" & fpInc & _
            "|FSC=" & fpFsc & _
            "|" & cfgSig

    OP_M10_TrySourceToken = True
End Function

Public Function OP_M10_SourceToken() As String
    Dim token As String
    Dim detail As String

    If OP_M10_TrySourceToken(token, detail) Then
        OP_M10_SourceToken = token
    Else
        OP_M10_SourceToken = "INVALID:" & detail
    End If
End Function

Public Function OP_M10_TryStateToken(ByRef token As String, ByRef detail As String) As Boolean
    Dim src As String
    Dim views As Variant
    Dim i As Long
    Dim fp As String

    token = vbNullString
    detail = vbNullString

    If Not OP_M10_TrySourceToken(src, detail) Then Exit Function

    views = Array("OP_PROTOCOLLI", "OP_PROTOCOLLI_PLUS", "OP_SCADENZE", _
                  "OP_DECRETI", "OP_INCIDENTI", "OP_FASCICOLI")

    token = "STATE2|" & src

    For i = LBound(views) To UBound(views)
        fp = OP_M10_ObjectFingerprint(CStr(views(i)))
        If Not OP_M10_FingerprintIsValid(fp) Then
            detail = "Vista operativa non verificabile: " & CStr(views(i)) & " = " & fp
            token = vbNullString
            Exit Function
        End If
        token = token & "|" & CStr(views(i)) & "=" & fp
    Next i

    OP_M10_TryStateToken = True
End Function

Public Function OP_M10_StateToken() As String
    Dim token As String
    Dim detail As String

    If OP_M10_TryStateToken(token, detail) Then
        OP_M10_StateToken = token
    Else
        OP_M10_StateToken = "INVALID:" & detail
    End If
End Function

Public Function OP_M10_FscStateToken(ByRef detail As String) As String
    Dim ws As Worksheet
    Dim sig As String
    Dim fp As String
    Dim found As Boolean

    detail = vbNullString
    sig = "FSC2"

    On Error GoTo EH

    For Each ws In ThisWorkbook.Worksheets
        If Left$(UCase$(ws.Name), 4) = "FSC_" Then
            fp = OP_M10_ObjectFingerprint(ws.Name)
            If Not OP_M10_FingerprintIsValid(fp) Then
                detail = "Stato FSC non verificabile: " & ws.Name & " = " & fp
                Exit Function
            End If
            sig = sig & "|" & ws.Name & "=" & fp
            found = True
        End If
    Next ws

    If Not found Then
        detail = "Nessun foglio FSC_* disponibile."
        Exit Function
    End If

    OP_M10_FscStateToken = sig
    Exit Function

EH:
    detail = "Impossibile calcolare lo stato FSC [" & CStr(Err.Number) & "]: " & Err.Description
End Function

Private Function OP_M10_MetaSheet() As Worksheet
    On Error Resume Next
    Set OP_M10_MetaSheet = ThisWorkbook.Worksheets(OP_M10_META)
    On Error GoTo 0
End Function

Private Function OP_M10_MetaKeyRow(ByVal keyName As String) As Long
    Dim ws As Worksheet
    Dim hit As Range

    Set ws = OP_M10_MetaSheet()
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set hit = ws.Columns(1).Find(What:=keyName, After:=ws.Cells(1, 1), _
                                 LookIn:=xlValues, LookAt:=xlWhole, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                 MatchCase:=False)
    On Error GoTo 0

    If Not hit Is Nothing Then OP_M10_MetaKeyRow = hit.Row
End Function

Public Sub OP_M10_SetMeta(ByVal keyName As String, ByVal value As Variant)
    Dim ws As Worksheet
    Dim rowNo As Long

    Set ws = OP_M10_MetaSheet()
    If ws Is Nothing Then Err.Raise vbObjectError + 1091, "OP_M10_SetMeta", "OP_M10_META assente."

    rowNo = OP_M10_MetaKeyRow(keyName)
    If rowNo = 0 Then
        rowNo = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If rowNo < 2 Then rowNo = 2
        OP_M10_WriteLiteral ws.Cells(rowNo, 1), keyName
    End If

    OP_M10_WriteLiteral ws.Cells(rowNo, 2), value
End Sub

Public Function OP_M10_GetMeta(ByVal keyName As String, _
                                Optional ByVal fallback As String = vbNullString) As String
    Dim ws As Worksheet
    Dim rowNo As Long
    Dim v As Variant

    Set ws = OP_M10_MetaSheet()
    If ws Is Nothing Then
        OP_M10_GetMeta = fallback
        Exit Function
    End If

    rowNo = OP_M10_MetaKeyRow(keyName)
    If rowNo = 0 Then
        OP_M10_GetMeta = fallback
        Exit Function
    End If

    v = ws.Cells(rowNo, 2).Value2
    If IsError(v) Then
        OP_M10_GetMeta = fallback
    Else
        OP_M10_GetMeta = CStr(v)
    End If
End Function

Public Sub OP_M10_MarkAlignmentSuccess(ByVal successDetail As String)
    Dim stateToken As String
    Dim sourceToken As String
    Dim detail As String

    If Not OP_M10_CheckDiscard(detail) Then
        Err.Raise vbObjectError + 1092, "OP_M10_MarkAlignmentSuccess", detail
    End If

    If Not OP_M10_TrySourceToken(sourceToken, detail) Then
        Err.Raise vbObjectError + 1093, "OP_M10_MarkAlignmentSuccess", detail
    End If

    If Not OP_M10_TryStateToken(stateToken, detail) Then
        Err.Raise vbObjectError + 1094, "OP_M10_MarkAlignmentSuccess", detail
    End If

    OP_M10_SetMeta "VERSION", OP_M10_VERSION
    OP_M10_SetMeta "LAST_GOOD_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    OP_M10_SetMeta "LAST_GOOD_SOURCE_TOKEN", sourceToken
    OP_M10_SetMeta "LAST_GOOD_STATE_TOKEN", stateToken
    OP_M10_SetMeta "LAST_GOOD_DETAIL", successDetail
    OP_M10_SetMeta "LAST_GOOD_SRC_FP_DB_DECRETI", OP_M10_ObjectFingerprint("DB_DECRETI")
    OP_M10_SetMeta "LAST_GOOD_SRC_FP_DB_INCIDENTI", OP_M10_ObjectFingerprint("DB_INCIDENTI")
    OP_M10_SetMeta "LAST_GOOD_SRC_FP_FSC_DB_CORRENTI", OP_M10_ObjectFingerprint("FSC_DB_CORRENTI")
    OP_M10_SetMeta "LAST_GOOD_SRC_FP_DB_PROTOCOLLI_CORRENTI", OP_M10_ObjectFingerprint("DB_PROTOCOLLI_CORRENTI")

    OP_M10_SetMeta "LAST_ATTEMPT_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    OP_M10_SetMeta "LAST_ATTEMPT_STATUS", OP_M10_GATE_OK
    OP_M10_SetMeta "LAST_ATTEMPT_DETAIL", successDetail
    OP_M10_SetMeta "CURRENT_GATE", OP_M10_GATE_OK
End Sub

Public Sub OP_M10_MarkAlignmentFailure(ByVal failureDetail As String)
    OP_M10_SetMeta "VERSION", OP_M10_VERSION
    OP_M10_SetMeta "LAST_ATTEMPT_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    OP_M10_SetMeta "LAST_ATTEMPT_STATUS", "FAILED"
    OP_M10_SetMeta "LAST_ATTEMPT_DETAIL", failureDetail
    OP_M10_SetMeta "CURRENT_GATE", OP_M10_GATE_BLOCKED
End Sub

Public Function OP_M10_IsCurrent(ByRef detail As String) As Boolean
    Dim currentToken As String
    Dim lastGood As String
    Dim gateStatus As String

    detail = vbNullString

    If Not OP_M10_CheckDiscard(detail) Then Exit Function

    gateStatus = OP_M10_GetMeta("CURRENT_GATE", OP_M10_GATE_BLOCKED)
    If StrComp(gateStatus, OP_M10_GATE_OK, vbTextCompare) <> 0 Then
        detail = "Gate operativo bloccato. Ultimo tentativo: " & _
                 OP_M10_GetMeta("LAST_ATTEMPT_STATUS", "N/D") & ". " & _
                 OP_M10_GetMeta("LAST_ATTEMPT_DETAIL", vbNullString)
        Exit Function
    End If

    lastGood = OP_M10_GetMeta("LAST_GOOD_STATE_TOKEN", vbNullString)
    If Len(lastGood) = 0 Then
        detail = "Nessun allineamento M10.1 valido registrato."
        Exit Function
    End If

    If Not OP_M10_TryStateToken(currentToken, detail) Then Exit Function

    If StrComp(currentToken, lastGood, vbBinaryCompare) <> 0 Then
        detail = "La base operativa e' cambiata dopo l'ultimo allineamento valido."
        Exit Function
    End If

    OP_M10_IsCurrent = True
End Function

Public Function OP_M10_OwnedMarker() As String
    OP_M10_OwnedMarker = OP_M10_MARKER
End Function

Public Function OP_M10_OwnedMarkerCell() As String
    OP_M10_OwnedMarkerCell = OP_M10_MARKER_CELL
End Function

Public Function OP_M10_HomeSheetName() As String
    OP_M10_HomeSheetName = OP_M10_HOME
End Function

Public Function OP_M10_SummarySheetName() As String
    OP_M10_SummarySheetName = OP_M10_SUMMARY
End Function

Public Function OP_M10_ReportSheetName() As String
    OP_M10_ReportSheetName = OP_M10_REPORT
End Function

Public Function OP_M10_MetaSheetName() As String
    OP_M10_MetaSheetName = OP_M10_META
End Function
