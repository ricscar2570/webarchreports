Attribute VB_Name = "modOP_M10_NativeTests"
Option Explicit

Private Sub OP_M10_TestAssert(ByVal conditionValue As Boolean, ByVal messageText As String)
    If Not conditionValue Then
        Err.Raise vbObjectError + 6190, "OP_M10_TEST_COMPLETO", messageText
    End If
End Sub

Public Sub OP_M10_TEST_COMPLETO()
    Dim beforeSheets As Long
    Dim afterFirstInstall As Long
    Dim afterSecondInstall As Long
    Dim token As String
    Dim originalAligned As String
    Dim originalStatus As String
    Dim originalAt As String
    Dim originalDetail As String
    Dim detail As String
    Dim currentResult As Boolean

    On Error GoTo Failed

    If InStr(1, UCase$(ThisWorkbook.Name), "WEBARCH_TEST_OP_M10", vbTextCompare) <> 1 Then
        Err.Raise vbObjectError + 6191, "OP_M10_TEST_COMPLETO", _
                  "Usa una copia sacrificabile il cui nome inizi con WebArch_TEST_OP_M10."
    End If

    beforeSheets = ThisWorkbook.Worksheets.Count

    OP_M10_INSTALLA
    afterFirstInstall = ThisWorkbook.Worksheets.Count
    OP_M10_TestAssert OP_M10_SheetExists("OP_HOME"), "OP_HOME mancante."
    OP_M10_TestAssert OP_M10_SheetExists("OP_SINTESI"), "OP_SINTESI mancante."
    OP_M10_TestAssert OP_M10_SheetExists("OP_REPORT"), "OP_REPORT mancante."
    OP_M10_TestAssert OP_M10_SheetExists("OP_M10_META"), "OP_M10_META mancante."
    OP_M10_TestAssert afterFirstInstall >= beforeSheets, "Numero fogli diminuito durante installazione."

    OP_M10_INSTALLA
    afterSecondInstall = ThisWorkbook.Worksheets.Count
    OP_M10_TestAssert afterSecondInstall = afterFirstInstall, _
                      "Reinstallazione M10 ha duplicato fogli."

    token = OP_M10_SourceToken()
    OP_M10_TestAssert Len(token) > 20, "Token sorgente vuoto o non plausibile."

    originalAligned = OP_M10_GetMeta("ALIGNED_SOURCE_TOKEN", "")
    originalStatus = OP_M10_GetMeta("REFRESH_STATUS", "")
    originalAt = OP_M10_GetMeta("ALIGNED_AT", "")
    originalDetail = OP_M10_GetMeta("REFRESH_DETAIL", "")

    OP_M10_SetMeta "ALIGNED_SOURCE_TOKEN", token
    OP_M10_SetMeta "REFRESH_STATUS", "OK"
    currentResult = OP_M10_IsCurrent(detail)
    OP_M10_TestAssert currentResult, "M10 non riconosce una base allineata: " & detail

    OP_M10_SetMeta "ALIGNED_SOURCE_TOKEN", "TEST_STALE_TOKEN"
    currentResult = OP_M10_IsCurrent(detail)
    OP_M10_TestAssert Not currentResult, "Gate stale non ha bloccato una base differente."

    OP_M10_SetMeta "ALIGNED_SOURCE_TOKEN", token
    OP_M10_SetMeta "REFRESH_STATUS", "FAILED"
    currentResult = OP_M10_IsCurrent(detail)
    OP_M10_TestAssert Not currentResult, "Gate export non ha bloccato REFRESH_STATUS=FAILED."

    OP_M10_SetMeta "ALIGNED_SOURCE_TOKEN", token
    OP_M10_SetMeta "REFRESH_STATUS", "OK"
    OP_M10_TestAssert OP_M10_RicostruisciOutput(detail), _
                      "Ricostruzione sintesi/report fallita: " & detail

    OP_M10_TestAssert InStr(1, CStr(ThisWorkbook.Worksheets("OP_REPORT").Range("A1").Value2), _
                                   "REPORT OPERATIVO", vbTextCompare) > 0, _
                      "Titolo OP_REPORT inatteso."

    ' Ripristina i soli metadati M10 toccati dal test.
    OP_M10_SetMeta "ALIGNED_SOURCE_TOKEN", originalAligned
    OP_M10_SetMeta "REFRESH_STATUS", originalStatus
    OP_M10_SetMeta "ALIGNED_AT", originalAt
    OP_M10_SetMeta "REFRESH_DETAIL", originalDetail

    MsgBox "BETA M10 TEST: PASS" & vbCrLf & vbCrLf & _
           "Installazione idempotente, token/staleness e output generati verificati." & vbCrLf & _
           "Il test NON salva il workbook.", vbInformation, "WebArch Beta M10"
    Exit Sub

Failed:
    On Error Resume Next
    OP_M10_SetMeta "ALIGNED_SOURCE_TOKEN", originalAligned
    OP_M10_SetMeta "REFRESH_STATUS", originalStatus
    OP_M10_SetMeta "ALIGNED_AT", originalAt
    OP_M10_SetMeta "REFRESH_DETAIL", originalDetail
    On Error GoTo 0

    MsgBox "BETA M10 TEST: FAIL" & vbCrLf & vbCrLf & _
           "Errore " & CStr(Err.Number) & ": " & Err.Description & vbCrLf & _
           "Chiudi la copia di test SENZA salvare.", vbCritical, "WebArch Beta M10"
End Sub
