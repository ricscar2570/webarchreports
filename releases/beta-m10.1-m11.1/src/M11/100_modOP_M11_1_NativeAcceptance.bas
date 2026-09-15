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
    Dim sourceObjects As Variant
    Dim i As Long
    Dim detail As String
    Dim guardDetail As String
    Dim t0 As Double
    Dim elapsed As Double
    Dim token As String
    Dim stateToken As String
    Dim fp As String
    Dim sheetCount1 As Long
    Dim sheetCount2 As Long
    Dim guardHeld As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed

    If InStr(1, UCase$(ThisWorkbook.Name), "WEBARCH_TEST_OP_M11", vbTextCompare) <> 1 Then
        Err.Raise vbObjectError + 6201, "OP_M11_TEST_ACCETTAZIONE", _
                  "Usa una copia sacrificabile il cui nome inizi con WebArch_TEST_OP_M11."
    End If

    OP_M11_Assert InStr(1, OP_M10_Version(), "M10.1", vbTextCompare) > 0, _
                  "M10.1 Audit Remediation non installata."

    OP_M11_Assert Not MINIMAL_DiscardRequired(), _
                  "La copia di test e' gia' marcata da scartare."

    requiredSheets = Array("OP_PROTOCOLLI", "OP_PROTOCOLLI_PLUS", "OP_SCADENZE", _
                           "OP_DECRETI", "OP_INCIDENTI", "OP_FASCICOLI", _
                           "OP_HOME", "OP_SINTESI", "OP_REPORT", "OP_M10_META")

    For i = LBound(requiredSheets) To UBound(requiredSheets)
        OP_M11_Assert OP_M11_HasSheet(CStr(requiredSheets(i))), _
                      "Foglio richiesto mancante: " & CStr(requiredSheets(i))
    Next i

    sheetCount1 = ThisWorkbook.Worksheets.Count
    OP_M10_INSTALLA
    sheetCount2 = ThisWorkbook.Worksheets.Count
    OP_M11_Assert sheetCount1 = sheetCount2, "Reinstallazione M10.1 non idempotente."

    OP_M11_Assert OP_M10_CheckPendingEdits(detail), _
                  "La copia contiene modifiche operative non consolidate: " & detail

    t0 = Timer
    token = vbNullString
    OP_M11_Assert OP_M10_TrySourceToken(token, detail), _
                  "Source token M10.1 non disponibile: " & detail
    elapsed = Timer - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#

    OP_M11_Assert Left$(token, 5) = "SRC2|", "Formato source token non M10.1."

    sourceObjects = Array("DB_PROTOCOLLI_CORRENTI", "DB_DECRETI", _
                          "DB_INCIDENTI", "FSC_DB_CORRENTI")

    For i = LBound(sourceObjects) To UBound(sourceObjects)
        fp = OP_M10_ObjectFingerprint(CStr(sourceObjects(i)))
        OP_M11_Assert OP_M10_FingerprintIsValid(fp), _
                      CStr(sourceObjects(i)) & " non verificabile: " & fp
    Next i

    ' Prova critica M11.1: il riallineamento deve funzionare mentre AGGIORNA_TUTTO
    ' detiene gia' il BUSY. Questo e' il caso che il vecchio M11 non verificava.
    If Not EASYONE_TryBeginOperation("AGGIORNA_TUTTO", guardDetail) Then
        Err.Raise vbObjectError + 6202, "OP_M11_TEST_ACCETTAZIONE", _
                  "Impossibile acquisire BUSY di prova: " & guardDetail
    End If
    guardHeld = True

    OP_M11_Assert OP_M10_RiallineaOperativo(False, detail), _
                  "Riallineamento M10.1 sotto BUSY fallito: " & detail

    EASYONE_EndOperation "AGGIORNA_TUTTO"
    guardHeld = False

    stateToken = vbNullString
    OP_M11_Assert OP_M10_TryStateToken(stateToken, detail), _
                  "State token non valido dopo riallineamento: " & detail
    OP_M11_Assert Left$(stateToken, 7) = "STATE2|", "Formato state token non M10.1."

    OP_M11_Assert OP_M10_IsCurrent(detail), _
                  "M10.1 non e' CORRENTE dopo il riallineamento: " & detail

    OP_M11_Assert OP_M10_RicostruisciOutput(detail), _
                  "Ricostruzione output M10.1 fallita: " & detail

    OP_M11_Assert OP_M10_IsCurrent(detail), _
                  "La ricostruzione output ha reso stale il sistema: " & detail

    OP_M11_Assert Not MINIMAL_DiscardRequired(), _
                  "Il gate M11.1 ha attivato inaspettatamente MINIMAL_DiscardRequired."

    MsgBox "BETA M11.1 - GATE NATIVO STRUTTURALE: PASS" & vbCrLf & vbCrLf & _
           "Verificati: M10.1, fogli M3-M9, pending-edit gate, fingerprint stretti, " & _
           "riallineamento core sotto BUSY, state token, output e stato corrente." & vbCrLf & _
           "Tempo source token sulla copia: " & Format$(elapsed, "0.000") & " s." & vbCrLf & vbCrLf & _
           "RESTANO OBBLIGATORI prima di RC1: Debug > Compila VBAProject, importazione reale, " & _
           "Save/reopen, vecchie macro/staleness, PDF e prova prestazioni sui dati reali.", _
           vbInformation, "WebArch Beta M11.1"
    Exit Sub

Failed:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If guardHeld Then EASYONE_EndOperation "AGGIORNA_TUTTO"
    On Error GoTo 0

    MsgBox "BETA M11.1 TEST: FAIL" & vbCrLf & vbCrLf & _
           "Errore " & CStr(errNumber) & ": " & errDescription & vbCrLf & _
           "Chiudi la copia di test SENZA salvare.", _
           vbCritical, "WebArch Beta M11.1"
End Sub
