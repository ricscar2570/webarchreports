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

Private Sub OP_M10_SetHeader(ByVal ws As Worksheet, ByVal titleText As String, ByVal subtitleText As String)
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

Private Function OP_M10_UsedDataRows(ByVal sheetName As String) As Long
    Dim ws As Worksheet
    Dim lastCell As Range

    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                 MatchCase:=False)
    If lastCell Is Nothing Then
        OP_M10_UsedDataRows = 0
    ElseIf lastCell.Row <= 1 Then
        OP_M10_UsedDataRows = 0
    Else
        OP_M10_UsedDataRows = lastCell.Row - 1
    End If
Done:
End Function

Private Sub OP_M10_AddDomainRow(ByVal ws As Worksheet, ByVal rowNum As Long, _
                                ByVal domainName As String, ByVal sourceObject As String, _
                                ByVal opSheet As String, ByVal provenance As String)
    Dim sourceRows As Variant
    Dim opRows As Variant
    Dim opState As String
    Dim flags As Long

    If Len(sourceObject) = 0 Then
        sourceRows = "-"
    ElseIf OP_M10_ObjectFingerprint(sourceObject) = "MISSING" Then
        sourceRows = "N/D"
    Else
        sourceRows = OP_M10_ObjectRowCount(sourceObject)
    End If

    If OP_M10_SheetExists(opSheet) Then
        opRows = OP_M10_UsedDataRows(opSheet)
        flags = OP_M10_CountFlags(opSheet)
        If flags > 0 Then
            opState = "ATTENZIONE"
        Else
            opState = "OK"
        End If
    Else
        opRows = "N/D"
        flags = 0
        opState = "MANCANTE"
    End If

    ws.Cells(rowNum, 1).Value2 = domainName
    ws.Cells(rowNum, 2).Value2 = sourceObject
    ws.Cells(rowNum, 3).Value2 = sourceRows
    ws.Cells(rowNum, 4).Value2 = opSheet
    ws.Cells(rowNum, 5).Value2 = opRows
    ws.Cells(rowNum, 6).Value2 = flags
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
    Dim lastCell As Range
    Dim outCol As Long

    wsOut.Cells(nextRow, 1).Value2 = "ULTIMI RECORD KPI OPERATIVI"
    wsOut.Cells(nextRow, 1).Font.Bold = True
    nextRow = nextRow + 1
    firstDataRow = nextRow

    For Each ws In ThisWorkbook.Worksheets
        If UCase$(Left$(ws.Name, 3)) = "OP_" And _
           InStr(1, UCase$(ws.Name), "KPI", vbTextCompare) > 0 Then

            Set lastCell = Nothing
            On Error Resume Next
            Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                         LookIn:=xlFormulas, LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                         MatchCase:=False)
            On Error GoTo 0

            If Not lastCell Is Nothing Then
                lastRow = lastCell.Row
                Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                             LookIn:=xlFormulas, LookAt:=xlPart, _
                                             SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
                                             MatchCase:=False)
                lastCol = lastCell.Column
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
    Next ws

    If Not foundAny Then
        wsOut.Cells(firstDataRow, 1).Value2 = "Nessuna tabella/foglio OP_*KPI disponibile."
        nextRow = firstDataRow + 1
    End If
End Sub

Public Function OP_M10_RicostruisciOutput(ByRef detail As String) As Boolean
    Dim wsS As Worksheet
    Dim wsR As Worksheet
    Dim alignedDetail As String
    Dim isCurrent As Boolean
    Dim nextRow As Long
    Dim lastRow As Long

    detail = vbNullString
    On Error GoTo EH

    If Not OP_M10_SheetExists("OP_SINTESI") Or Not OP_M10_SheetExists("OP_REPORT") Then
        detail = "M10 non installato: mancano OP_SINTESI/OP_REPORT."
        Exit Function
    End If

    Set wsS = ThisWorkbook.Worksheets("OP_SINTESI")
    Set wsR = ThisWorkbook.Worksheets("OP_REPORT")

    OP_M10_ClearGenerated wsS
    OP_M10_ClearGenerated wsR

    OP_M10_SetHeader wsS, "WEBARCH BETA - SINTESI OPERATIVA", _
                     "Valori OP_ separati dai risultati alpha. Nessuna nota utente entra nel PDF per default."

    wsS.Range("A4").Value2 = "Versione M10"
    wsS.Range("B4").Value2 = OP_M10_Version()
    wsS.Range("A5").Value2 = "RunID sorgente"
    wsS.Range("B5").Value2 = OP_M10_GetConfigText("B18", "N/D")
    wsS.Range("A6").Value2 = "CommitID sorgente"
    wsS.Range("B6").Value2 = OP_M10_GetConfigText("B136", "N/D")
    wsS.Range("D4").Value2 = "Riallineato il"
    wsS.Range("E4").Value2 = OP_M10_GetMeta("ALIGNED_AT", "N/D")
    wsS.Range("D5").Value2 = "Stato refresh"
    wsS.Range("E5").Value2 = OP_M10_GetMeta("REFRESH_STATUS", "NON ALLINEATO")
    wsS.Range("D6").Value2 = "Token sorgente"
    wsS.Range("E6").Value2 = OP_M10_SourceToken()

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
    OP_M10_AddDomainRow wsS, 10, "SCADENZE", "", "OP_SCADENZE", _
                        "Registro operativo M5"
    OP_M10_AddDomainRow wsS, 11, "DECRETI", "DB_DECRETI", "OP_DECRETI", _
                        "Fonte alpha + rettifiche/record M6-M7"
    OP_M10_AddDomainRow wsS, 12, "INCIDENTI", "DB_INCIDENTI", "OP_INCIDENTI", _
                        "Fonte alpha + rettifiche/eventi M6-M8"
    OP_M10_AddDomainRow wsS, 13, "FASCICOLI", "FSC_DB_CORRENTI", "OP_FASCICOLI", _
                        "Fotografia FSC B1 + annotazioni/rettifiche M2/M6/M9"

    nextRow = 16
    OP_M10_AppendKpiLatest wsS, nextRow

    isCurrent = OP_M10_IsCurrent(alignedDetail)
    wsS.Range("A" & CStr(nextRow + 1)).Value2 = "VALIDITA' EXPORT"
    If isCurrent Then
        wsS.Range("B" & CStr(nextRow + 1)).Value2 = "CORRENTE"
    Else
        wsS.Range("B" & CStr(nextRow + 1)).Value2 = "BLOCCATO"
    End If
    wsS.Range("C" & CStr(nextRow + 1)).Value2 = alignedDetail

    wsS.Columns("A:H").EntireColumn.AutoFit
    If wsS.Columns("H:H").ColumnWidth > 45 Then wsS.Columns("H:H").ColumnWidth = 45

    ' OP_REPORT e' una vista autonoma e distinta dal report alpha.
    OP_M10_SetHeader wsR, "WEBARCH BETA - REPORT OPERATIVO", _
                     "BOZZA OPERATIVA - NON SOSTITUISCE REPORT/CERTIFICAZIONI ALPHA"
    wsR.Range("A4").Value2 = "Generato"
    wsR.Range("B4").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsR.Range("D4").Value2 = "RunID"
    wsR.Range("E4").Value2 = OP_M10_GetConfigText("B18", "N/D")
    wsR.Range("D5").Value2 = "CommitID"
    wsR.Range("E5").Value2 = OP_M10_GetConfigText("B136", "N/D")

    wsS.Range("A8:H" & CStr(nextRow + 1)).Copy
    wsR.Range("A7").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    wsR.Columns("A:H").EntireColumn.AutoFit
    If wsR.Columns("H:H").ColumnWidth > 45 Then wsR.Columns("H:H").ColumnWidth = 45
    wsR.PageSetup.Orientation = xlLandscape
    wsR.PageSetup.Zoom = False
    wsR.PageSetup.FitToPagesWide = 1
    wsR.PageSetup.FitToPagesTall = False
    wsR.PageSetup.CenterHeader = "WEBARCH BETA - REPORT OPERATIVO - NON UFFICIALE"
    wsR.PageSetup.CenterFooter = "M10 | Fonte e dato operativo distinti"

    lastRow = wsR.Cells(wsR.Rows.Count, 1).End(xlUp).Row
    wsR.PageSetup.PrintArea = wsR.Range("A1:H" & CStr(lastRow)).Address

    wsS.Calculate
    wsR.Calculate

    detail = "OP_SINTESI e OP_REPORT ricostruiti."
    OP_M10_RicostruisciOutput = True
    Exit Function

EH:
    detail = "Ricostruzione output M10 non riuscita [" & CStr(Err.Number) & "]: " & Err.Description
End Function

Public Sub OP_M10_APRI_SINTESI()
    If OP_M10_SheetExists("OP_SINTESI") Then
        ThisWorkbook.Worksheets("OP_SINTESI").Activate
    Else
        MsgBox "OP_SINTESI non disponibile. Esegui OP_M10_INSTALLA.", vbExclamation, "WebArch Beta M10"
    End If
End Sub

Public Sub OP_M10_ESPORTA_PDF()
    Dim detail As String
    Dim outputPath As Variant
    Dim defaultPath As String
    Dim ws As Worksheet

    On Error GoTo EH

    If Not OP_M10_IsCurrent(detail) Then
        MsgBox "EXPORT BLOCCATO: " & detail & vbCrLf & vbCrLf & _
               "Esegui prima RIALLINEA OPERATIVO oppure la nuova importazione integrata.", _
               vbExclamation, "WebArch Beta M10"
        Exit Sub
    End If

    If Not OP_M10_RicostruisciOutput(detail) Then
        MsgBox detail, vbCritical, "WebArch Beta M10"
        Exit Sub
    End If

    ' Ricontrollo dopo la ricostruzione: nessun export se la base e' diventata stale.
    If Not OP_M10_IsCurrent(detail) Then
        MsgBox "EXPORT BLOCCATO DOPO RICALCOLO: " & detail, vbCritical, "WebArch Beta M10"
        Exit Sub
    End If

    defaultPath = ThisWorkbook.Path & Application.PathSeparator & _
                  "WebArch_BETA_REPORT_OPERATIVO_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"

    outputPath = Application.GetSaveAsFilename(InitialFileName:=defaultPath, _
                    FileFilter:="PDF (*.pdf), *.pdf", _
                    Title:="Esporta report operativo WebArch Beta")
    If VarType(outputPath) = vbBoolean Then Exit Sub

    If LCase$(Right$(CStr(outputPath), 4)) <> ".pdf" Then
        outputPath = CStr(outputPath) & ".pdf"
    End If

    If Len(Dir$(CStr(outputPath))) > 0 Then
        If MsgBox("Sostituire il PDF esistente?", vbYesNo + vbQuestion, _
                  "WebArch Beta M10") <> vbYes Then Exit Sub
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

    MsgBox "REPORT OPERATIVO ESPORTATO:" & vbCrLf & CStr(outputPath), _
           vbInformation, "WebArch Beta M10"
    Exit Sub

EH:
    MsgBox "Export M10 non riuscito [" & CStr(Err.Number) & "]: " & Err.Description, _
           vbCritical, "WebArch Beta M10"
End Sub
