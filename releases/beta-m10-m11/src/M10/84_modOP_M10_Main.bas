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

Private Sub OP_M10_AddButton(ByVal ws As Worksheet, ByVal btnName As String, _
                             ByVal caption As String, ByVal macroName As String, _
                             ByVal leftPos As Double, ByVal topPos As Double, _
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

    ws.Range("A1").Value2 = "WEBARCH BETA - AREA OPERATIVA M10"
    ws.Range("A2").Value2 = "Comandi nuovi. HOME alpha e vecchie macro restano invariati."
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 18

    ws.Range("A4").Value2 = "Regola"
    ws.Range("B4").Value2 = "Significato"
    ws.Range("A5").Value2 = "IMPORTA FONTI + OPERATIVO"
    ws.Range("B5").Value2 = "Pipeline alpha pubblica senza Save -> riallineamento OP_ -> unico Save finale."
    ws.Range("A6").Value2 = "IMPORTA FSC + OPERATIVO"
    ws.Range("B6").Value2 = "FSC_Aggiorna mantiene il proprio guard; dopo il ritorno M10 riallinea e salva una volta."
    ws.Range("A7").Value2 = "RIALLINEA OPERATIVO"
    ws.Range("B7").Value2 = "Non pubblica bozze: richiama solo azioni RICARICA/RIALLINEA realmente esposte."
    ws.Range("A8").Value2 = "ESPORTA PDF"
    ws.Range("B8").Value2 = "Bloccato se fonte e ultimo token allineato non coincidono."
    ws.Range("A9").Value2 = "NOTE"
    ws.Range("B9").Value2 = "Escluse dal PDF operativo per default."
    ws.Range("A4:B4").Font.Bold = True
    ws.Columns("A:A").ColumnWidth = 28
    ws.Columns("B:B").ColumnWidth = 78

    OP_M10_DeleteOwnedButtons ws
    OP_M10_AddButton ws, "OP_M10_BTN_IMPORT", "IMPORTA FONTI + OPERATIVO", _
                     "OP_M10_IMPORTA_FONTI", 20, 210, 190
    OP_M10_AddButton ws, "OP_M10_BTN_FSC", "IMPORTA FSC + OPERATIVO", _
                     "OP_M10_IMPORTA_FASCICOLI", 225, 210, 190
    OP_M10_AddButton ws, "OP_M10_BTN_ALIGN", "RIALLINEA OPERATIVO", _
                     "OP_M10_RIALLINEA", 430, 210, 165
    OP_M10_AddButton ws, "OP_M10_BTN_SUM", "APRI SINTESI", _
                     "OP_M10_APRI_SINTESI", 20, 250, 130
    OP_M10_AddButton ws, "OP_M10_BTN_PDF", "ESPORTA PDF OPERATIVO", _
                     "OP_M10_ESPORTA_PDF", 165, 250, 190
    OP_M10_AddButton ws, "OP_M10_BTN_STATUS", "STATO M10", _
                     "OP_M10_STATO", 370, 250, 120
End Sub

Public Sub OP_M10_INSTALLA()
    Dim detail As String
    Dim wsHome As Worksheet
    Dim wsSummary As Worksheet
    Dim wsReport As Worksheet
    Dim wsMeta As Worksheet
    Dim createdOk As Boolean

    On Error GoTo EH

    If Not OP_M10_PreflightInstall(detail) Then
        MsgBox detail, vbCritical, "WebArch Beta M10"
        Exit Sub
    End If

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
    If Len(OP_M10_GetMeta("REFRESH_STATUS", "")) = 0 Then
        OP_M10_SetMeta "REFRESH_STATUS", "NON ALLINEATO"
    End If

    OP_M10_BuildHome wsHome

    If Not OP_M10_RicostruisciOutput(detail) Then GoTo Failed

    createdOk = True
    wsHome.Activate

    MsgBox "BETA M10: COMANDI UNIFICATI, SINTESI E REPORT INSTALLATI." & vbCrLf & vbCrLf & _
           "Nessun componente alpha/FSC/M1-M9 e' stato sostituito." & vbCrLf & _
           "Prima dell'export esegui RIALLINEA OPERATIVO.", _
           vbInformation, "WebArch Beta M10"
    Exit Sub

Failed:
    MsgBox "Installazione M10 non completata." & vbCrLf & vbCrLf & detail, _
           vbCritical, "WebArch Beta M10"
    Exit Sub

EH:
    detail = "Installazione M10 non riuscita [" & CStr(Err.Number) & "]: " & Err.Description
    MsgBox detail, vbCritical, "WebArch Beta M10"
End Sub

Public Sub OP_M10_STATO()
    Dim detail As String
    Dim currentText As String

    If OP_M10_IsCurrent(detail) Then
        currentText = "CORRENTE - export operativo consentito."
    Else
        currentText = "NON CORRENTE - export operativo bloccato."
    End If

    MsgBox "WEBARCH BETA M10" & vbCrLf & _
           "Versione: " & OP_M10_Version() & vbCrLf & _
           "Stato: " & currentText & vbCrLf & _
           "Ultimo riallineamento: " & OP_M10_GetMeta("ALIGNED_AT", "N/D") & vbCrLf & _
           "Refresh: " & OP_M10_GetMeta("REFRESH_STATUS", "N/D") & vbCrLf & vbCrLf & _
           detail, vbInformation, "WebArch Beta M10"
End Sub
