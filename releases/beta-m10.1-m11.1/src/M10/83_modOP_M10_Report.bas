Attribute VB_Name = "modOP_M10_Report"
Option Explicit

Private Sub OP_M10_ClearGenerated(ByVal ws As Worksheet)
    Dim markerValue As String

    markerValue = CStr(ws.Range("AA1").Value2)
    ws.Cells.Clear
    ws.Range("AA1").NumberFormat = "@"
    ws.Range("AA1").Value2 = markerValue
    ws.Columns("AA:AA").Hidden = True
End Sub

Private Sub OP_M10_SetHeader(ByVal ws As Worksheet, _
                             ByVal titleText As String, _
                             ByVal subtitleText As String)
    With ws
        .Range("A1").Value2 = titleText
        .Range("A2").Value2 = subtitleText
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 18
        .Range("A2").Font.Italic = True
        .Columns("A:H").ColumnWidth = 18
        .Columns("A:A").ColumnWidth = 24
        .Rows("1:2").RowHeight = 24
    End With
End Sub

Private Function OP_M10_SourceRowsText(ByVal sourceObject As String) As Variant
    Dim fp As String
    Dim rowsCount As Long

    If Len(sourceObject) = 0 Then
        OP_M10_SourceRowsText = "-"
        Exit Function
    End If

    fp = OP_M10_ObjectFingerprint(sourceObject)
    If Not OP_M10_FingerprintIsValid(fp) Then
        OP_M10_SourceRowsText = "N/D"
        Exit Function
    End If

    rowsCount = OP_M10_ObjectRowCount(sourceObject)
    If rowsCount < 0 Then
        OP_M10_SourceRowsText = "N/D"
    Else
        OP_M10_SourceRowsText = rowsCount
    End If
End Function

Private Function OP_M10_ViewRowsText(ByVal opSheet As String, _
                                     ByRef stateText As String) As Variant
    Dim ok As Boolean
    Dim rowsCount As Long
    Dim fp As String

    stateText = "MANCANTE"

    If Not OP_M10_SheetExists(opSheet) Then
        OP_M10_ViewRowsText = "N/D"
        Exit Function
    End If

    fp = OP_M10_ObjectFingerprint(opSheet)
    If Not OP_M10_FingerprintIsValid(fp) Then
        stateText = "ERRORE"
        OP_M10_ViewRowsText = "N/D"
        Exit Function
    End If

    rowsCount = OP_M10_SheetTableRowCount(opSheet, ok)
    If ok Then
        OP_M10_ViewRowsText = rowsCount
        stateText = "DISPONIBILE"
    Else
        OP_M10_ViewRowsText = "N/D"
        stateText = "DISPONIBILE - CONTEGGIO N/D"
    End If
End Function

Private Sub OP_M10_AddDomainRow(ByVal ws As Worksheet, _
                                ByVal rowNum As Long, _
                                ByVal domainName As String, _
                                ByVal sourceObject As String, _
                                ByVal opSheet As String, _
                                ByVal provenance As String)
    Dim sourceRows As Variant
    Dim opRows As Variant
    Dim opState As String

    sourceRows = OP_M10_SourceRowsText(sourceObject)
    opRows = OP_M10_ViewRowsText(opSheet, opState)

    ws.Cells(rowNum, 1).Value2 = domainName
    ws.Cells(rowNum, 2).Value2 = sourceObject
    ws.Cells(rowNum, 3).Value2 = sourceRows
    ws.Cells(rowNum, 4).Value2 = opSheet
    ws.Cells(rowNum, 5).Value2 = opRows
    ws.Cells(rowNum, 6).Value2 = "N/D - non inferito da celle"
    ws.Cells(rowNum, 7).Value2 = opState
    ws.Cells(rowNum, 8).Value2 = provenance
End Sub

Private Sub OP_M10_AppendKpiLatest(ByVal wsOut As Worksheet, ByRef nextRow As Long)
    Dim ws As Worksheet
    Dim firstDataRow As Long
    Dim foundAny As Boolean
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long
    Dim lastRowCell As Range
    Dim lastColCell As Range
    Dim outCol As Long

    wsOut.Cells(nextRow, 1).Value2 = "ULTIMI RECORD KPI OPERATIVI"
    wsOut.Cells(nextRow, 1).Font.Bold = True
    nextRow = nextRow + 1
    firstDataRow = nextRow

    For Each ws In ThisWorkbook.Worksheets
        If UCase$(Left$(ws.Name, 3)) = "OP_" Then
            If InStr(1, UCase$(ws.Name), "KPI", vbTextCompare) > 0 Then
                Set lastRowCell = Nothing
                Set lastColCell = Nothing

                On Error Resume Next
                Set lastRowCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                                LookIn:=xlFormulas, LookAt:=xlPart, _
                                                SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                                MatchCase:=False)
                Set lastColCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                                LookIn:=xlFormulas, LookAt:=xlPart, _
                                                SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
                                                MatchCase:=False)
                On Error GoTo 0

                If Not lastRowCell Is Nothing Then
                    If Not lastColCell Is Nothing Then
                        lastRow = lastRowCell.Row
                        lastCol = lastColCell.Column
                        If lastCol > 8 Then lastCol = 8

                        foundAny = True
                        wsOut.Cells(nextRow, 1).Value2 = ws.Name
                        outCol = 2

                        For c = 1 To lastCol
                            If IsError(ws.Cells(lastRow, c).Value) Then
                                wsOut.Cells(nextRow, outCol).Value2 = "#ERRORE"
                            Else
                                wsOut.Cells(nextRow, outCol).Value2 = ws.Cells(lastRow, c).Value2
                            End If
                            outCol = outCol + 1
                        Next c

                        nextRow = nextRow + 1
                    End If
                End If
            End If
        End If
    Next ws

    If Not foundAny Then
        wsOut.Cells(firstDataRow, 1).Value2 = "Nessuna tabella/foglio OP_*KPI disponibile."
        nextRow = firstDataRow + 1
    End If
End Sub

Private Function OP_M10_FindValidityRow(ByVal ws As Worksheet) As Long
    Dim hit As Range

    On Error Resume Next
    Set hit = ws.Columns(1).Find(What:="VALIDITA' EXPORT", After:=ws.Cells(1, 1), _
                                 LookIn:=xlValues, LookAt:=xlWhole, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                 MatchCase:=False)
    On Error GoTo 0

    If Not hit Is Nothing Then OP_M10_FindValidityRow = hit.Row
End Function

Public Function OP_M10_UpdateValidityBanner(ByRef detail As String) As Boolean
    Dim wsS As Worksheet
    Dim wsR As Worksheet
    Dim stateDetail As String
    Dim isCurrent As Boolean
    Dim rowS As Long
    Dim rowR As Long

    detail = vbNullString
    On Error GoTo EH

    If Not OP_M10_CheckDiscard(detail) Then Exit Function

    If Not OP_M10_SheetExists("OP_SINTESI") Then
        detail = "OP_SINTESI non disponibile."
        Exit Function
    End If
    If Not OP_M10_SheetExists("OP_REPORT") Then
        detail = "OP_REPORT non disponibile."
        Exit Function
    End If

    Set wsS = ThisWorkbook.Worksheets("OP_SINTESI")
    Set wsR = ThisWorkbook.Worksheets("OP_REPORT")

    rowS = OP_M10_FindValidityRow(wsS)
    rowR = OP_M10_FindValidityRow(wsR)

    If rowS = 0 Then
        detail = "Riga VALIDITA' EXPORT non trovata in OP_SINTESI."
        Exit Function
    End If
    If rowR = 0 Then
        detail = "Riga VALIDITA' EXPORT non trovata in OP_REPORT."
        Exit Function
    End If

    isCurrent = OP_M10_IsCurrent(stateDetail)

    If isCurrent Then
        wsS.Cells(rowS, 2).Value2 = "CORRENTE"
        wsR.Cells(rowR, 2).Value2 = "CORRENTE"
    Else
        wsS.Cells(rowS, 2).Value2 = "BLOCCATO"
        wsR.Cells(rowR, 2).Value2 = "BLOCCATO"
    End If

    wsS.Cells(rowS, 3).Value2 = stateDetail
    wsR.Cells(rowR, 3).Value2 = stateDetail

    detail = stateDetail
    OP_M10_UpdateValidityBanner = True
    Exit Function

EH:
    detail = "Aggiornamento banner validita' non riuscito [" & CStr(Err.Number) & "]: " & Err.Description
End Function

Public Function OP_M10_RicostruisciOutput(ByRef detail As String) As Boolean
    Dim wsS As Worksheet
    Dim wsR As Worksheet
    Dim stateDetail As String
    Dim isCurrent As Boolean
    Dim nextRow As Long
    Dim lastRow As Long
    Dim validityRow As Long

    detail = vbNullString
    On Error GoTo EH

    If Not OP_M10_CheckDiscard(detail) Then Exit Function

    If Not OP_M10_SheetExists("OP_SINTESI") Then
        detail = "M10.1 non installato: OP_SINTESI assente."
        Exit Function
    End If
    If Not OP_M10_SheetExists("OP_REPORT") Then
        detail = "M10.1 non installato: OP_REPORT assente."
        Exit Function
    End If

    Set wsS = ThisWorkbook.Worksheets("OP_SINTESI")
    Set wsR = ThisWorkbook.Worksheets("OP_REPORT")

    OP_M10_ClearGenerated wsS
    OP_M10_ClearGenerated wsR

    OP_M10_SetHeader wsS, "WEBARCH BETA - SINTESI OPERATIVA M10.1", _
                     "Dati operativi separati dai risultati alpha. Nessuna nota utente entra nel PDF per default."

    wsS.Range("A4").Value2 = "Versione M10"
    wsS.Range("B4").Value2 = OP_M10_Version()
    wsS.Range("A5").Value2 = "RunID sorgente"
    wsS.Range("B5").Value2 = OP_M10_GetConfigText("B18", "N/D")
    wsS.Range("A6").Value2 = "CommitID sorgente"
    wsS.Range("B6").Value2 = OP_M10_GetConfigText("B136", "N/D")

    wsS.Range("D4").Value2 = "Ultimo GOOD"
    wsS.Range("E4").Value2 = OP_M10_GetMeta("LAST_GOOD_AT", "N/D")
    wsS.Range("D5").Value2 = "Ultimo tentativo"
    wsS.Range("E5").Value2 = OP_M10_GetMeta("LAST_ATTEMPT_STATUS", "N/D")
    wsS.Range("D6").Value2 = "Gate corrente"
    wsS.Range("E6").Value2 = OP_M10_GetMeta("CURRENT_GATE", "BLOCKED")

    wsS.Cells(8, 1).Value2 = "Dominio"
    wsS.Cells(8, 2).Value2 = "Oggetto fonte"
    wsS.Cells(8, 3).Value2 = "Righe fonte"
    wsS.Cells(8, 4).Value2 = "Vista operativa"
    wsS.Cells(8, 5).Value2 = "Righe vista"
    wsS.Cells(8, 6).Value2 = "Segnalazioni"
    wsS.Cells(8, 7).Value2 = "Stato"
    wsS.Cells(8, 8).Value2 = "Provenienza"
    wsS.Range("A8:H8").Font.Bold = True

    OP_M10_AddDomainRow wsS, 9, "PROTOCOLLI", "DB_PROTOCOLLI_CORRENTI", "OP_PROTOCOLLI", _
                        "Fonte alpha + annotazioni M2/M3"
    OP_M10_AddDomainRow wsS, 10, "PROTOCOLLI PLUS", "DB_PROTOCOLLI_CORRENTI", "OP_PROTOCOLLI_PLUS", _
                        "M4 campi configurabili; policy export separata"
    OP_M10_AddDomainRow wsS, 11, "SCADENZE", vbNullString, "OP_SCADENZE", _
                        "Registro operativo M5 + promemoria protocolli"
    OP_M10_AddDomainRow wsS, 12, "DECRETI", "DB_DECRETI", "OP_DECRETI", _
                        "Fonte alpha + rettifiche/record M6-M7"
    OP_M10_AddDomainRow wsS, 13, "INCIDENTI", "DB_INCIDENTI", "OP_INCIDENTI", _
                        "Fonte alpha + rettifiche/eventi M6-M8"
    OP_M10_AddDomainRow wsS, 14, "FASCICOLI", "FSC_DB_CORRENTI", "OP_FASCICOLI", _
                        "Fotografia FSC B1 + annotazioni/rettifiche M2/M6/M9"

    nextRow = 17
    OP_M10_AppendKpiLatest wsS, nextRow

    validityRow = nextRow + 1
    isCurrent = OP_M10_IsCurrent(stateDetail)
    wsS.Cells(validityRow, 1).Value2 = "VALIDITA' EXPORT"

    If isCurrent Then
        wsS.Cells(validityRow, 2).Value2 = "CORRENTE"
    Else
        wsS.Cells(validityRow, 2).Value2 = "BLOCCATO"
    End If

    wsS.Cells(validityRow, 3).Value2 = stateDetail
    wsS.Columns("A:H").EntireColumn.AutoFit
    If wsS.Columns("H:H").ColumnWidth > 45 Then wsS.Columns("H:H").ColumnWidth = 45

    OP_M10_SetHeader wsR, "WEBARCH BETA - REPORT OPERATIVO M10.1", _
                     "OPERATIVO - NON ISTITUZIONALE - NON SOSTITUISCE REPORT/CERTIFICAZIONI ALPHA"

    wsR.Range("A4").Value2 = "Generato"
    wsR.Range("B4").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsR.Range("D4").Value2 = "RunID"
    wsR.Range("E4").Value2 = OP_M10_GetConfigText("B18", "N/D")
    wsR.Range("D5").Value2 = "CommitID"
    wsR.Range("E5").Value2 = OP_M10_GetConfigText("B136", "N/D")
    wsR.Range("D6").Value2 = "Fonte"
    wsR.Range("E6").Value2 = "Alpha/FSC in sola lettura + layer OP_"

    With wsS.Range("A8:H" & CStr(validityRow))
        wsR.Range("A7").Resize(.Rows.Count, .Columns.Count).Value2 = .Value2
    End With
    wsR.Range("A7:H7").Font.Bold = True

    wsR.Columns("A:H").EntireColumn.AutoFit
    If wsR.Columns("H:H").ColumnWidth > 45 Then wsR.Columns("H:H").ColumnWidth = 45
    wsR.PageSetup.Orientation = xlLandscape
    wsR.PageSetup.Zoom = False
    wsR.PageSetup.FitToPagesWide = 1
    wsR.PageSetup.FitToPagesTall = False
    wsR.PageSetup.CenterHeader = "WEBARCH BETA - REPORT OPERATIVO - NON ISTITUZIONALE"
    wsR.PageSetup.CenterFooter = "M10.1 | Fonte e dato operativo distinti"

    lastRow = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    wsR.PageSetup.PrintArea = wsR.Range("A1:H" & CStr(lastRow)).Address

    wsS.Calculate
    wsR.Calculate

    detail = "OP_SINTESI e OP_REPORT M10.1 ricostruiti."
    OP_M10_RicostruisciOutput = True
    Exit Function

EH:
    detail = "Ricostruzione output M10.1 non riuscita [" & CStr(Err.Number) & "]: " & Err.Description
End Function

Public Sub OP_M10_APRI_SINTESI()
    If OP_M10_SheetExists("OP_SINTESI") Then
        ThisWorkbook.Worksheets("OP_SINTESI").Activate
    Else
        MsgBox "OP_SINTESI non disponibile. Esegui OP_M10_INSTALLA.", _
               vbExclamation, "WebArch Beta M10.1"
    End If
End Sub

Public Sub OP_M10_ESPORTA_PDF()
    Dim detail As String
    Dim outputPath As Variant
    Dim defaultPath As String
    Dim ws As Worksheet

    On Error GoTo EH

    If Not OP_M10_CheckDiscard(detail) Then
        MsgBox "EXPORT BLOCCATO: " & detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    If Not OP_M10_CheckPendingEdits(detail) Then
        MsgBox "EXPORT BLOCCATO: " & detail, vbExclamation, "WebArch Beta M10.1"
        Exit Sub
    End If

    If Not OP_M10_IsCurrent(detail) Then
        MsgBox "EXPORT BLOCCATO: " & detail & vbCrLf & vbCrLf & _
               "Esegui prima RIALLINEA OPERATIVO oppure la nuova importazione integrata.", _
               vbExclamation, "WebArch Beta M10.1"
        Exit Sub
    End If

    If Not OP_M10_RicostruisciOutput(detail) Then
        MsgBox detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    If Not OP_M10_IsCurrent(detail) Then
        MsgBox "EXPORT BLOCCATO DOPO RICALCOLO: " & detail, _
               vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    defaultPath = ThisWorkbook.Path & Application.PathSeparator & _
                  "WebArch_BETA_REPORT_OPERATIVO_M10_1_" & _
                  Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    outputPath = Application.GetSaveAsFilename(InitialFileName:=defaultPath, _
                    FileFilter:="PDF (*.pdf), *.pdf", _
                    Title:="Esporta report operativo WebArch Beta M10.1")
    If VarType(outputPath) = vbBoolean Then Exit Sub

    If LCase$(Right$(CStr(outputPath), 4)) <> ".pdf" Then
        outputPath = CStr(outputPath) & ".pdf"
    End If

    If Len(Dir$(CStr(outputPath))) > 0 Then
        If MsgBox("Sostituire il PDF esistente?", vbYesNo + vbQuestion, _
                  "WebArch Beta M10.1") <> vbYes Then Exit Sub
    End If

    If Not OP_M10_CheckDiscard(detail) Then
        MsgBox "EXPORT BLOCCATO: " & detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets("OP_REPORT")
    ws.Calculate
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(outputPath), _
                           Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, OpenAfterPublish:=False

    If Len(Dir$(CStr(outputPath))) = 0 Then
        Err.Raise vbObjectError + 6120, "OP_M10_ESPORTA_PDF", "Excel non ha creato il PDF."
    End If

    If FileLen(CStr(outputPath)) <= 0 Then
        Err.Raise vbObjectError + 6121, "OP_M10_ESPORTA_PDF", "Il PDF creato e' vuoto."
    End If

    MsgBox "REPORT OPERATIVO M10.1 ESPORTATO:" & vbCrLf & CStr(outputPath), _
           vbInformation, "WebArch Beta M10.1"
    Exit Sub

EH:
    MsgBox "Export M10.1 non riuscito [" & CStr(Err.Number) & "]: " & Err.Description, _
           vbCritical, "WebArch Beta M10.1"
End Sub
