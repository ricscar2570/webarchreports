Attribute VB_Name = "modOP_M11_NativeAcceptance"
Option Explicit

Private Sub OP_M11_Assert(ByVal conditionValue As Boolean, ByVal messageText As String)
    If Not conditionValue Then
        Err.Raise vbObjectError + 6200, "OP_M11_TEST_ACCETTAZIONE", messageText
    End If
End Sub

Private Function OP_M11_HasSheet(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    OP_M11_HasSheet = Not ws Is Nothing
End Function

Public Sub OP_M11_TEST_ACCETTAZIONE()
    Dim requiredSheets As Variant
    Dim i As Long
    Dim detail As String
    Dim t0 As Double
    Dim elapsed As Double
    Dim token As String
    Dim sheetCount1 As Long
    Dim sheetCount2 As Long

    On Error GoTo Failed

    If InStr(1, UCase$(ThisWorkbook.Name), "WEBARCH_TEST_OP_M11", vbTextCompare) <> 1 Then
        Err.Raise vbObjectError + 6201, "OP_M11_TEST_ACCETTAZIONE", _
                  "Usa una copia sacrificabile il cui nome inizi con WebArch_TEST_OP_M11."
    End If

    requiredSheets = Array("OP_PROTOCOLLI", "OP_SCADENZE", "OP_DECRETI", _
                           "OP_INCIDENTI", "OP_FASCICOLI", _
                           "OP_HOME", "OP_SINTESI", "OP_REPORT", "OP_M10_META")

    For i = LBound(requiredSheets) To UBound(requiredSheets)
        OP_M11_Assert OP_M11_HasSheet(CStr(requiredSheets(i))), _
                      "Foglio richiesto mancante: " & CStr(requiredSheets(i))
    Next i

    sheetCount1 = ThisWorkbook.Worksheets.Count
    OP_M10_INSTALLA
    sheetCount2 = ThisWorkbook.Worksheets.Count
    OP_M11_Assert sheetCount1 = sheetCount2, "Reinstallazione M10 non idempotente."

    t0 = Timer
    token = OP_M10_SourceToken()
    elapsed = Timer - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#

    OP_M11_Assert Len(token) > 20, "Token sorgente non disponibile."
    OP_M11_Assert OP_M10_RicostruisciOutput(detail), _
                  "Ricostruzione M10 fallita: " & detail

    OP_M11_Assert OP_M10_ObjectFingerprint("DB_PROTOCOLLI_CORRENTI") <> "MISSING", _
                  "DB_PROTOCOLLI_CORRENTI non disponibile."
    OP_M11_Assert OP_M10_ObjectFingerprint("DB_DECRETI") <> "MISSING", _
                  "DB_DECRETI non disponibile."
    OP_M11_Assert OP_M10_ObjectFingerprint("DB_INCIDENTI") <> "MISSING", _
                  "DB_INCIDENTI non disponibile."
    OP_M11_Assert OP_M10_ObjectFingerprint("FSC_DB_CORRENTI") <> "MISSING", _
                  "FSC_DB_CORRENTI non disponibile."

    MsgBox "BETA M11 - GATE NATIVO PARZIALE: PASS" & vbCrLf & vbCrLf & _
           "Strutture integrate presenti; reinstallazione M10 idempotente; output ricostruibile." & vbCrLf & _
           "Tempo calcolo token sulla copia: " & Format$(elapsed, "0.000") & " s." & vbCrLf & vbCrLf & _
           "RESTANO OBBLIGATORI: Debug > Compila VBAProject, ciclo operativo reale, Save/reopen " & _
           "e prova sui volumi reali prima del congelamento RC.", _
           vbInformation, "WebArch Beta M11"
    Exit Sub

Failed:
    MsgBox "BETA M11 TEST: FAIL" & vbCrLf & vbCrLf & _
           "Errore " & CStr(Err.Number) & ": " & Err.Description & vbCrLf & _
           "Chiudi la copia di test SENZA salvare.", vbCritical, "WebArch Beta M11"
End Sub
