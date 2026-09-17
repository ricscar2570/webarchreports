Attribute VB_Name = "modOP_M10_NativeTests"
Option Explicit

Private Const OP_M10_TEST_SHEET As String = "OP_M10_TEST_HASH"

Private Sub OP_M10_TestAssert(ByVal condition As Boolean, ByVal messageText As String)
    If Not condition Then
        Err.Raise vbObjectError + 6190, "OP_M10_TEST_COMPLETO", messageText
    End If
End Sub

Private Function OP_M10_TestCountSheet(ByVal sheetName As String) As Long
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            OP_M10_TestCountSheet = OP_M10_TestCountSheet + 1
        End If
    Next ws
End Function

Private Sub OP_M10_TestDeleteTempSheet()
    Dim oldAlerts As Boolean

    If Not OP_M10_SheetExists(OP_M10_TEST_SHEET) Then Exit Sub

    oldAlerts = Application.DisplayAlerts
    On Error GoTo SafeExit

    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(OP_M10_TEST_SHEET).Delete

SafeExit:
    Application.DisplayAlerts = oldAlerts
End Sub

Public Sub OP_M10_TEST_COMPLETO()
    Dim detail As String
    Dim guardDetail As String
    Dim fp1 As String
    Dim fp2 As String
    Dim stateToken As String
    Dim sourceToken As String
    Dim oldLastGood As String
    Dim oldGate As String
    Dim oldAttemptStatus As String
    Dim oldAttemptDetail As String
    Dim oldAttemptAt As String
    Dim guardHeld As Boolean
    Dim ws As Worksheet
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    OP_M10_TestAssert Left$(ThisWorkbook.Name, Len("WebArch_TEST_OP_M10")) = "WebArch_TEST_OP_M10", _
                      "Il test puo' essere eseguito solo su una copia WebArch_TEST_OP_M10..."

    OP_M10_TestAssert InStr(1, OP_M10_Version(), "M10.1", vbTextCompare) > 0, _
                      "Versione M10.1 non installata."

    OP_M10_INSTALLA
    OP_M10_INSTALLA

    OP_M10_TestAssert OP_M10_TestCountSheet("OP_HOME") = 1, "OP_HOME duplicato."
    OP_M10_TestAssert OP_M10_TestCountSheet("OP_SINTESI") = 1, "OP_SINTESI duplicato."
    OP_M10_TestAssert OP_M10_TestCountSheet("OP_REPORT") = 1, "OP_REPORT duplicato."
    OP_M10_TestAssert OP_M10_TestCountSheet("OP_M10_META") = 1, "OP_M10_META duplicato."

    sourceToken = vbNullString
    detail = vbNullString
    OP_M10_TestAssert OP_M10_TrySourceToken(sourceToken, detail), _
                      "Source token M10.1 non valido: " & detail
    OP_M10_TestAssert Left$(sourceToken, 5) = "SRC2|", "Formato source token inatteso."

    OP_M10_TestAssert OP_M10_CheckPendingEdits(detail), _
                      "Preflight modifiche visibili non superato: " & detail

    OP_M10_TestDeleteTempSheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = OP_M10_TEST_SHEET

    ws.Range("A1").Value2 = "ABCDxxxxxEFGHyyyWXYZ"
    fp1 = OP_M10_ObjectFingerprint(OP_M10_TEST_SHEET)

    ws.Range("A1").Value2 = "ABCD12345EFGH___WXYZ"
    fp2 = OP_M10_ObjectFingerprint(OP_M10_TEST_SHEET)

    OP_M10_TestAssert OP_M10_FingerprintIsValid(fp1), "Fingerprint 1 non valido."
    OP_M10_TestAssert OP_M10_FingerprintIsValid(fp2), "Fingerprint 2 non valido."
    OP_M10_TestAssert StrComp(fp1, fp2, vbBinaryCompare) <> 0, _
                      "Collisione fingerprint: due testi differenti hanno lo stesso hash."

    OP_M10_TestAssert Not OP_M10_FingerprintIsValid("ERR:91"), "ERR:91 accettato come fingerprint."
    OP_M10_TestAssert Not OP_M10_FingerprintIsValid("MISSING:TEST"), _
                      "MISSING accettato come fingerprint."

    OP_M10_TestDeleteTempSheet

    oldLastGood = OP_M10_GetMeta("LAST_GOOD_STATE_TOKEN", vbNullString)
    oldGate = OP_M10_GetMeta("CURRENT_GATE", vbNullString)
    oldAttemptStatus = OP_M10_GetMeta("LAST_ATTEMPT_STATUS", vbNullString)
    oldAttemptDetail = OP_M10_GetMeta("LAST_ATTEMPT_DETAIL", vbNullString)
    oldAttemptAt = OP_M10_GetMeta("LAST_ATTEMPT_AT", vbNullString)

    OP_M10_SetMeta "LAST_GOOD_STATE_TOKEN", "TEST_SENTINEL_LAST_GOOD"
    OP_M10_MarkAlignmentFailure "TEST FAILURE - NON DEVE TOCCARE LAST GOOD"

    OP_M10_TestAssert OP_M10_GetMeta("LAST_GOOD_STATE_TOKEN", vbNullString) = _
                      "TEST_SENTINEL_LAST_GOOD", _
                      "Un fallimento ha sovrascritto LAST_GOOD_STATE_TOKEN."

    OP_M10_SetMeta "LAST_GOOD_STATE_TOKEN", oldLastGood
    OP_M10_SetMeta "CURRENT_GATE", oldGate
    OP_M10_SetMeta "LAST_ATTEMPT_STATUS", oldAttemptStatus
    OP_M10_SetMeta "LAST_ATTEMPT_DETAIL", oldAttemptDetail
    OP_M10_SetMeta "LAST_ATTEMPT_AT", oldAttemptAt

    If Not EASYONE_TryBeginOperation("AGGIORNA_TUTTO", guardDetail) Then
        Err.Raise vbObjectError + 6191, "OP_M10_TEST_COMPLETO", _
                  "Impossibile acquisire BUSY per il test annidamento: " & guardDetail
    End If
    guardHeld = True

    OP_M10_TestAssert OP_M10_RiallineaOperativo(False, detail), _
                      "Riallineamento core sotto BUSY AGGIORNA_TUTTO fallito: " & detail

    EASYONE_EndOperation "AGGIORNA_TUTTO"
    guardHeld = False

    stateToken = vbNullString
    OP_M10_TestAssert OP_M10_TryStateToken(stateToken, detail), _
                      "State token M10.1 non valido dopo riallineamento: " & detail
    OP_M10_TestAssert Left$(stateToken, 7) = "STATE2|", "Formato state token inatteso."

    OP_M10_TestAssert OP_M10_IsCurrent(detail), _
                      "M10.1 non risulta corrente dopo riallineamento riuscito: " & detail

    ' Ultimo test: il discard deve bloccare immediatamente IsCurrent.
    ' Da questo punto la copia di test deve essere chiusa SENZA salvare.
    MINIMAL_RequireDiscard

    OP_M10_TestAssert Not OP_M10_IsCurrent(detail), _
                      "M10.1 considera corrente una copia marcata MINIMAL_DiscardRequired."

    MsgBox "BETA M10.1 TEST: PASS" & vbCrLf & vbCrLf & _
           "Coperti: installazione idempotente, fingerprint anti-collisione, " & _
           "last-good protetto, pending-edit preflight, riallineamento core sotto BUSY e discard gate." & _
           vbCrLf & vbCrLf & "CHIUDI QUESTA COPIA SENZA SALVARE: il test ha attivato volontariamente il discard.", _
           vbInformation, "WebArch Beta M10.1"
    Exit Sub

EH:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    OP_M10_TestDeleteTempSheet
    If guardHeld Then EASYONE_EndOperation "AGGIORNA_TUTTO"
    On Error GoTo 0

    MsgBox "BETA M10.1 TEST: FAIL" & vbCrLf & vbCrLf & _
           "[" & CStr(errNumber) & "] " & errDescription & vbCrLf & vbCrLf & _
           "Chiudi la copia di test senza salvare.", _
           vbCritical, "WebArch Beta M10.1"
End Sub
