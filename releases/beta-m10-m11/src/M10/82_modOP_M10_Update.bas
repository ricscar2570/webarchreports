Attribute VB_Name = "modOP_M10_Update"
Option Explicit

Private Function OP_M10_PreflightOperational(ByRef detail As String) As Boolean
    detail = vbNullString

    If Not OP_M10_SheetExists("OP_HOME") Or _
       Not OP_M10_SheetExists("OP_SINTESI") Or _
       Not OP_M10_SheetExists("OP_REPORT") Then
        detail = "M10 non installato. Esegui OP_M10_INSTALLA."
        Exit Function
    End If

    If Not ThisWorkbook.Saved Then
        detail = "Il workbook contiene modifiche non salvate. M10 non avvia una nuova importazione " & _
                 "per non sovrascrivere bozze o modifiche locali. Salva o annulla consapevolmente " & _
                 "le modifiche, poi riprova."
        Exit Function
    End If

    OP_M10_PreflightOperational = True
End Function

Private Function OP_M10_ReloadDomainVerified(ByVal sheetName As String, _
                                                   ByVal sourceObject As String, _
                                                   ByVal sourceMetaKey As String, _
                                                   ByRef detail As String) As Boolean
    Dim sourceFp As String
    Dim priorSourceFp As String
    Dim beforeViewFp As String
    Dim afterViewFp As String
    Dim actionDetail As String

    detail = vbNullString

    sourceFp = OP_M10_ObjectFingerprint(sourceObject)
    If sourceFp = "MISSING" Then
        detail = sheetName & ": fonte richiesta mancante (" & sourceObject & ")."
        Exit Function
    End If

    priorSourceFp = OP_M10_GetMeta(sourceMetaKey, "")
    beforeViewFp = OP_M10_ObjectFingerprint(sheetName)

    If Not OP_M10_InvokeReloadAction(sheetName, True, actionDetail) Then
        detail = actionDetail
        Exit Function
    End If

    afterViewFp = OP_M10_ObjectFingerprint(sheetName)
    If afterViewFp = "MISSING" Then
        detail = sheetName & ": la vista non e' disponibile dopo RICARICA."
        Exit Function
    End If

    ' Se la fonte e' cambiata rispetto all'ultimo riallineamento e la vista rimane
    ' byte-logicamente identica nel fingerprint M10, non certifichiamo il refresh.
    If Len(priorSourceFp) > 0 And priorSourceFp <> "MISSING" Then
        If StrComp(sourceFp, priorSourceFp, vbBinaryCompare) <> 0 Then
            If StrComp(beforeViewFp, afterViewFp, vbBinaryCompare) = 0 Then
                detail = sheetName & ": la fonte e' cambiata ma la vista non mostra alcun cambiamento " & _
                         "verificabile dopo RICARICA. Riallineamento bloccato per sicurezza."
                Exit Function
            End If
        End If
    End If

    detail = actionDetail
    OP_M10_ReloadDomainVerified = True
End Function

Public Function OP_M10_RiallineaOperativo(ByVal requireSavedWorkbook As Boolean, _
                                           ByRef detail As String) As Boolean
    Dim oneDetail As String
    Dim allDetail As String
    Dim ok As Boolean

    detail = vbNullString
    On Error GoTo EH

    If requireSavedWorkbook Then
        If Not ThisWorkbook.Saved Then
            detail = "Riallineamento bloccato: il workbook contiene modifiche non salvate."
            Exit Function
        End If
    End If

    ' Le pagine M7-M9 espongono RICARICA nel contratto operativo.
    ' M10 usa esclusivamente l'OnAction realmente presente sul foglio.
    ok = OP_M10_ReloadDomainVerified("OP_DECRETI", "DB_DECRETI", _
                                      "SRC_FP_DB_DECRETI", oneDetail)
    allDetail = allDetail & oneDetail & vbCrLf
    If Not ok Then GoTo NotAligned

    ok = OP_M10_ReloadDomainVerified("OP_INCIDENTI", "DB_INCIDENTI", _
                                      "SRC_FP_DB_INCIDENTI", oneDetail)
    allDetail = allDetail & oneDetail & vbCrLf
    If Not ok Then GoTo NotAligned

    ok = OP_M10_ReloadDomainVerified("OP_FASCICOLI", "FSC_DB_CORRENTI", _
                                      "SRC_FP_FSC_DB_CORRENTI", oneDetail)
    allDetail = allDetail & oneDetail & vbCrLf
    If Not ok Then GoTo NotAligned

    ' Le viste M3/M5 non vengono forzate se non espongono un comando di ricarica.
    ' Le note non entrano nel PDF M10; lo scadenziario resta un registro manuale/persistente.
    Call OP_M10_InvokeReloadAction("OP_PROTOCOLLI", False, oneDetail)
    allDetail = allDetail & oneDetail & vbCrLf
    Call OP_M10_InvokeReloadAction("OP_SCADENZE", False, oneDetail)
    allDetail = allDetail & oneDetail

    OP_M10_MarkAligned "OK", allDetail

    If Not OP_M10_RicostruisciOutput(detail) Then
        allDetail = allDetail & vbCrLf & detail
        OP_M10_MarkAligned "FAILED", allDetail
        detail = allDetail
        Exit Function
    End If

    detail = allDetail
    OP_M10_RiallineaOperativo = True
    Exit Function

NotAligned:
    OP_M10_MarkAligned "FAILED", allDetail
    detail = allDetail
    Exit Function

EH:
    detail = "Riallineamento operativo M10 non riuscito [" & CStr(Err.Number) & "]: " & Err.Description
    On Error Resume Next
    OP_M10_MarkAligned "FAILED", detail
    On Error GoTo 0
End Function

Public Sub OP_M10_RIALLINEA()
    Dim detail As String

    If OP_M10_RiallineaOperativo(True, detail) Then
        MsgBox "BETA M10: RIALLINEAMENTO OPERATIVO COMPLETATO." & vbCrLf & vbCrLf & _
               "Il workbook NON e' stato salvato automaticamente." & vbCrLf & _
               "Salva con Ctrl+S dopo aver verificato le viste." & vbCrLf & vbCrLf & detail, _
               vbInformation, "WebArch Beta M10"
    Else
        MsgBox "BETA M10: RIALLINEAMENTO NON COMPLETATO." & vbCrLf & vbCrLf & detail, _
               vbCritical, "WebArch Beta M10"
    End If
End Sub

Public Sub OP_M10_IMPORTA_FONTI()
    Dim folderPath As String
    Dim detail As String
    Dim refreshDetail As String
    Dim guardHeld As Boolean
    Dim excelStateCaptured As Boolean
    Dim pipelineStarted As Boolean
    Dim dataWritten As Boolean
    Dim saveAttempted As Boolean
    Dim success As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    If Not OP_M10_PreflightOperational(detail) Then
        MsgBox detail, vbExclamation, "WebArch Beta M10"
        Exit Sub
    End If

    If Not EASYONE_TryBeginOperation("AGGIORNA_TUTTO", detail) Then
        MsgBox detail, vbExclamation, WA_APP_NAME
        Exit Sub
    End If
    guardHeld = True

    If Not MINIMAL_PrepareWorkbook(detail) Then
        MsgBox detail, vbExclamation, WA_APP_NAME
        GoTo Cleanup
    End If

    WA_ClearSourceFolderSessionSilent
    If Not ChooseSourceFolder(folderPath, True) Then GoTo Cleanup

    CaptureExcelState
    excelStateCaptured = True
    EnterControlledExecution
    ForceDisableSourceMacros
    ExecutionLogBegin "BETA M10 - IMPORTA FONTI + OPERATIVO"
    pipelineStarted = True

    If Not MINIMAL_ExecutePipeline(detail) Then GoTo Failed
    dataWritten = True

    If Not OP_M10_RiallineaOperativo(False, refreshDetail) Then
        detail = "La fonte e' stata aggiornata, ma il riallineamento operativo M10 e' fallito." & _
                 vbCrLf & refreshDetail
        MINIMAL_RequireDiscard
        GoTo Failed
    End If

    detail = detail & vbCrLf & "Riallineamento operativo M10: OK."
    SetOperationStatus "WebArch M10 - salvataggio finale..."
    ExecutionLogEnd True, "Importazione e riallineamento M10 completati; salvataggio finale da confermare.", _
                    detail, "", "", True

    Application.DisplayAlerts = True
    saveAttempted = True
    ThisWorkbook.Save
    If Not ThisWorkbook.Saved Then
        Err.Raise vbObjectError + 6110, "OP_M10_IMPORTA_FONTI", _
                  "Excel non ha confermato il salvataggio finale M10."
    End If

    success = True
    GoTo Cleanup

Failed:
    On Error Resume Next
    ExecutionLogEnd False, "BETA M10 interrotta.", detail
    On Error GoTo 0
    GoTo Cleanup

EH:
    errNumber = Err.Number
    errDescription = Err.Description
    detail = "Errore M10 [" & CStr(errNumber) & "]: " & errDescription
    If dataWritten Or saveAttempted Then MINIMAL_RequireDiscard
    On Error Resume Next
    ExecutionLogRecordFailure detail
    On Error GoTo 0

Cleanup:
    On Error Resume Next
    If excelStateCaptured Then RestoreExcelState
    If guardHeld Then EASYONE_EndOperation "AGGIORNA_TUTTO"
    On Error GoTo 0

    If success Then
        MsgBox "BETA M10: IMPORTAZIONE FONTI + OPERATIVO COMPLETATA E SALVATA." & _
               vbCrLf & vbCrLf & detail, vbInformation, "WebArch Beta M10"
    ElseIf pipelineStarted Then
        If MINIMAL_DiscardRequired() Then
            MsgBox detail & vbCrLf & vbCrLf & MINIMAL_DiscardMessage(), _
                   vbCritical, "WebArch Beta M10"
        Else
            MsgBox "BETA M10 interrotta prima del salvataggio finale." & vbCrLf & vbCrLf & detail, _
                   vbExclamation, "WebArch Beta M10"
        End If
    End If
End Sub

Public Sub OP_M10_IMPORTA_FASCICOLI()
    Dim detail As String
    Dim beforeToken As String
    Dim afterToken As String
    Dim refreshDetail As String
    Dim saveAttempted As Boolean
    Dim sourceChanged As Boolean

    On Error GoTo EH

    If Not OP_M10_PreflightOperational(detail) Then
        MsgBox detail, vbExclamation, "WebArch Beta M10"
        Exit Sub
    End If

    If Not MINIMAL_PrepareWorkbook(detail) Then
        MsgBox detail, vbExclamation, WA_APP_NAME
        Exit Sub
    End If

    beforeToken = OP_M10_SourceToken()

    ' FSC_Aggiorna gestisce autonomamente il proprio BUSY: nessun guard M10 viene annidato.
    Application.Run "FSC_Aggiorna"

    afterToken = OP_M10_SourceToken()
    sourceChanged = (StrComp(beforeToken, afterToken, vbBinaryCompare) <> 0)

    If Not sourceChanged Then
        MsgBox "BETA M10: la fotografia FSC non risulta cambiata." & vbCrLf & _
               "L'operazione puo' essere stata annullata oppure la fonte e' identica. Nessun Save M10 eseguito.", _
               vbInformation, "WebArch Beta M10"
        Exit Sub
    End If

    If Not OP_M10_ReloadDomainVerified("OP_FASCICOLI", "FSC_DB_CORRENTI", _
                                      "SRC_FP_FSC_DB_CORRENTI", refreshDetail) Then
        MINIMAL_RequireDiscard
        MsgBox "FSC aggiornato ma OP_FASCICOLI non riallineato." & vbCrLf & vbCrLf & _
               refreshDetail & vbCrLf & vbCrLf & MINIMAL_DiscardMessage(), _
               vbCritical, "WebArch Beta M10"
        Exit Sub
    End If

    OP_M10_MarkAligned "OK", "FSC: " & refreshDetail

    If Not OP_M10_RicostruisciOutput(detail) Then
        OP_M10_MarkAligned "FAILED", detail
        MINIMAL_RequireDiscard
        MsgBox detail & vbCrLf & vbCrLf & MINIMAL_DiscardMessage(), _
               vbCritical, "WebArch Beta M10"
        Exit Sub
    End If

    saveAttempted = True
    ThisWorkbook.Save
    If Not ThisWorkbook.Saved Then
        Err.Raise vbObjectError + 6111, "OP_M10_IMPORTA_FASCICOLI", _
                  "Excel non ha confermato il salvataggio finale dopo FSC."
    End If

    MsgBox "BETA M10: FSC + OP_FASCICOLI RIALLINEATI E SALVATI.", _
           vbInformation, "WebArch Beta M10"
    Exit Sub

EH:
    detail = "Importazione FSC M10 non riuscita [" & CStr(Err.Number) & "]: " & Err.Description
    If sourceChanged Or saveAttempted Then
        On Error Resume Next
        MINIMAL_RequireDiscard
        detail = detail & vbCrLf & MINIMAL_DiscardMessage()
        On Error GoTo 0
    End If
    MsgBox detail, vbCritical, "WebArch Beta M10"
End Sub
