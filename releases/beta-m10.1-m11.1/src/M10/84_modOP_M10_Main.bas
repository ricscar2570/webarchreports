Attribute VB_Name = "modOP_M10_Main"
Option Explicit

Private Sub OP_M10_DeleteOwnedButtons(ByVal ws As Worksheet)
    Dim i As Long
    Dim shp As Shape

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, 11) = "OP_M10_BTN_" Then shp.Delete
    Next i
End Sub

Private Sub OP_M10_AddButton(ByVal ws As Worksheet, _
                             ByVal btnName As String, _
                             ByVal caption As String, _
                             ByVal macroName As String, _
                             ByVal leftPos As Double, _
                             ByVal topPos As Double, _
                             ByVal widthValue As Double)
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(5, leftPos, topPos, widthValue, 28)
    shp.Name = btnName
    shp.TextFrame.Characters.Text = caption
    shp.OnAction = macroName
    shp.Placement = xlFreeFloating
End Sub

Private Sub OP_M10_BuildHome(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("AA1").NumberFormat = "@"
    ws.Range("AA1").Value2 = OP_M10_OwnedMarker()
    ws.Columns("AA:AA").Hidden = True

    ws.Range("A1").Value2 = "WEBARCH BETA - AREA OPERATIVA M10.1"
    ws.Range("A2").Value2 = "Audit Remediation. HOME alpha e vecchie macro restano invariati."
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 18

    ws.Range("A4").Value2 = "Regola"
    ws.Range("B4").Value2 = "Significato"
    ws.Range("A5").Value2 = "IMPORTA FONTI + OPERATIVO"
    ws.Range("B5").Value2 = "Pipeline alpha -> core M3/M4/M5/M7/M8/M9 -> output -> unico Save finale."
    ws.Range("A6").Value2 = "IMPORTA FSC + OPERATIVO"
    ws.Range("B6").Value2 = "FSC mantiene il proprio BUSY; M10.1 verifica tutto lo stato FSC_* e il discard prima del Save."
    ws.Range("A7").Value2 = "RIALLINEA OPERATIVO"
    ws.Range("B7").Value2 = "Chiama direttamente le API core: nessun click simulato e nessun BUSY annidato."
    ws.Range("A8").Value2 = "ESPORTA PDF"
    ws.Range("B8").Value2 = "Fail-closed: bloccato con discard, bozze non applicate o stato diverso dall'ultimo GOOD."
    ws.Range("A9").Value2 = "LAST GOOD / LAST ATTEMPT"
    ws.Range("B9").Value2 = "Un fallimento non puo' sovrascrivere la baseline dell'ultimo allineamento valido."
    ws.Range("A10").Value2 = "NOTE"
    ws.Range("B10").Value2 = "Escluse dal PDF operativo per default."

    ws.Range("A4:B4").Font.Bold = True
    ws.Columns("A:A").ColumnWidth = 30
    ws.Columns("B:B").ColumnWidth = 86

    OP_M10_DeleteOwnedButtons ws
    OP_M10_AddButton ws, "OP_M10_BTN_IMPORT", "IMPORTA FONTI + OPERATIVO", _
                     "OP_M10_IMPORTA_FONTI", 20, 230, 190
    OP_M10_AddButton ws, "OP_M10_BTN_FSC", "IMPORTA FSC + OPERATIVO", _
                     "OP_M10_IMPORTA_FASCICOLI", 225, 230, 190
    OP_M10_AddButton ws, "OP_M10_BTN_ALIGN", "RIALLINEA OPERATIVO", _
                     "OP_M10_RIALLINEA", 430, 230, 165
    OP_M10_AddButton ws, "OP_M10_BTN_SUM", "APRI SINTESI", _
                     "OP_M10_APRI_SINTESI", 20, 270, 130
    OP_M10_AddButton ws, "OP_M10_BTN_PDF", "ESPORTA PDF OPERATIVO", _
                     "OP_M10_ESPORTA_PDF", 165, 270, 190
    OP_M10_AddButton ws, "OP_M10_BTN_STATUS", "STATO M10.1", _
                     "OP_M10_STATO", 370, 270, 130
End Sub

Private Sub OP_M10_DeleteIfCreated(ByVal sheetName As String, ByVal existedBefore As Boolean)
    Dim ws As Worksheet
    Dim oldAlerts As Boolean

    If existedBefore Then Exit Sub
    If Not OP_M10_SheetExists(sheetName) Then Exit Sub
    If Not OP_M10_IsOwnedSheet(sheetName) Then Exit Sub

    oldAlerts = Application.DisplayAlerts
    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Application.DisplayAlerts = False
    ws.Delete

SafeExit:
    Application.DisplayAlerts = oldAlerts
End Sub

Public Sub OP_M10_INSTALLA()
    Dim detail As String
    Dim warningDetail As String
    Dim wsHome As Worksheet
    Dim wsSummary As Worksheet
    Dim wsReport As Worksheet
    Dim wsMeta As Worksheet

    Dim homeExisted As Boolean
    Dim summaryExisted As Boolean
    Dim reportExisted As Boolean
    Dim metaExisted As Boolean

    On Error GoTo EH

    If Not OP_M10_PreflightInstall(detail) Then
        MsgBox detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    homeExisted = OP_M10_SheetExists("OP_HOME")
    summaryExisted = OP_M10_SheetExists("OP_SINTESI")
    reportExisted = OP_M10_SheetExists("OP_REPORT")
    metaExisted = OP_M10_SheetExists("OP_M10_META")

    Set wsHome = OP_M10_GetOrCreateOwnedSheet("OP_HOME", xlSheetVisible, detail)
    If wsHome Is Nothing Then GoTo Failed

    Set wsSummary = OP_M10_GetOrCreateOwnedSheet("OP_SINTESI", xlSheetVisible, detail)
    If wsSummary Is Nothing Then GoTo Failed

    Set wsReport = OP_M10_GetOrCreateOwnedSheet("OP_REPORT", xlSheetVisible, detail)
    If wsReport Is Nothing Then GoTo Failed

    Set wsMeta = OP_M10_GetOrCreateOwnedSheet("OP_M10_META", xlSheetVeryHidden, detail)
    If wsMeta Is Nothing Then GoTo Failed

    If Len(CStr(wsMeta.Range("A1").Value2)) = 0 Then
        wsMeta.Range("A1").Value2 = "KEY"
        wsMeta.Range("B1").Value2 = "VALUE"
    End If

    OP_M10_SetMeta "VERSION", OP_M10_Version()

    If Len(OP_M10_GetMeta("CURRENT_GATE", vbNullString)) = 0 Then
        OP_M10_SetMeta "CURRENT_GATE", "BLOCKED"
    End If

    If Len(OP_M10_GetMeta("LAST_ATTEMPT_STATUS", vbNullString)) = 0 Then
        OP_M10_SetMeta "LAST_ATTEMPT_STATUS", "NOT RUN"
        OP_M10_SetMeta "LAST_ATTEMPT_DETAIL", "Eseguire RIALLINEA OPERATIVO M10.1."
    End If

    OP_M10_BuildHome wsHome

    ' La ricostruzione iniziale e' opportunistica: l'installazione non finge
    ' un allineamento e non fallisce solo perche' le fonti non sono ancora disponibili.
    If Not OP_M10_RicostruisciOutput(warningDetail) Then
        If Len(warningDetail) > 0 Then
            OP_M10_SetMeta "INSTALL_WARNING", warningDetail
        End If
    Else
        OP_M10_SetMeta "INSTALL_WARNING", vbNullString
    End If

    wsHome.Activate

    MsgBox "BETA M10.1 AUDIT REMEDIATION INSTALLATA." & vbCrLf & vbCrLf & _
           "Nessun componente alpha/FSC/M1-M9 e' stato sostituito da M10.1." & vbCrLf & _
           "Lo stato resta BLOCCATO finche' RIALLINEA OPERATIVO non termina con successo.", _
           vbInformation, "WebArch Beta M10.1"
    Exit Sub

Failed:
    On Error Resume Next
    OP_M10_DeleteIfCreated "OP_M10_META", metaExisted
    OP_M10_DeleteIfCreated "OP_REPORT", reportExisted
    OP_M10_DeleteIfCreated "OP_SINTESI", summaryExisted
    OP_M10_DeleteIfCreated "OP_HOME", homeExisted
    On Error GoTo 0

    MsgBox "Installazione M10.1 non completata." & vbCrLf & vbCrLf & detail, _
           vbCritical, "WebArch Beta M10.1"
    Exit Sub

EH:
    detail = "Installazione M10.1 non riuscita [" & CStr(Err.Number) & "]: " & Err.Description

    On Error Resume Next
    OP_M10_DeleteIfCreated "OP_M10_META", metaExisted
    OP_M10_DeleteIfCreated "OP_REPORT", reportExisted
    OP_M10_DeleteIfCreated "OP_SINTESI", summaryExisted
    OP_M10_DeleteIfCreated "OP_HOME", homeExisted
    On Error GoTo 0

    MsgBox detail, vbCritical, "WebArch Beta M10.1"
End Sub

Public Sub OP_M10_STATO()
    Dim detail As String
    Dim currentText As String

    If OP_M10_IsCurrent(detail) Then
        currentText = "CORRENTE - export operativo consentito."
    Else
        currentText = "NON CORRENTE - export operativo bloccato."
    End If

    MsgBox "WEBARCH BETA M10.1" & vbCrLf & _
           "Versione: " & OP_M10_Version() & vbCrLf & _
           "Stato: " & currentText & vbCrLf & _
           "Ultimo GOOD: " & OP_M10_GetMeta("LAST_GOOD_AT", "N/D") & vbCrLf & _
           "Ultimo tentativo: " & OP_M10_GetMeta("LAST_ATTEMPT_STATUS", "N/D") & vbCrLf & _
           "Gate: " & OP_M10_GetMeta("CURRENT_GATE", "BLOCKED") & vbCrLf & vbCrLf & _
           detail, vbInformation, "WebArch Beta M10.1"
End Sub
