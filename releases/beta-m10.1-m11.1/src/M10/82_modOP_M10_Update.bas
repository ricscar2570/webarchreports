Attribute VB_Name = "modOP_M10_Update"
Option Explicit

Private Function OP_M10_PreflightOperational(ByVal requireSavedWorkbook As Boolean, _
                                              ByVal checkPendingEdits As Boolean, _
                                              ByRef detail As String) As Boolean
    detail = vbNullString

    If Not OP_M10_SheetExists("OP_HOME") Then
        detail = "M10.1 non installato. Esegui OP_M10_INSTALLA."
        Exit Function
    End If
    If Not OP_M10_SheetExists("OP_SINTESI") Then
        detail = "M10.1 non installato: OP_SINTESI assente."
        Exit Function
    End If
    If Not OP_M10_SheetExists("OP_REPORT") Then
        detail = "M10.1 non installato: OP_REPORT assente."
        Exit Function
    End If
    If Not OP_M10_SheetExists("OP_M10_META") Then
        detail = "M10.1 non installato: OP_M10_META assente."
        Exit Function
    End If

    If Not OP_M10_IsOwnedSheet("OP_HOME") Then
        detail = "OP_HOME non appartiene a M10.1."
        Exit Function
    End If
    If Not OP_M10_IsOwnedSheet("OP_SINTESI") Then
        detail = "OP_SINTESI non appartiene a M10.1."
        Exit Function
    End If
    If Not OP_M10_IsOwnedSheet("OP_REPORT") Then
        detail = "OP_REPORT non appartiene a M10.1."
        Exit Function
    End If
    If Not OP_M10_IsOwnedSheet("OP_M10_META") Then
        detail = "OP_M10_META non appartiene a M10.1."
        Exit Function
    End If

    If Not OP_M10_CheckDiscard(detail) Then Exit Function

    If requireSavedWorkbook Then
        If Not ThisWorkbook.Saved Then
            detail = "Il workbook contiene modifiche non salvate. M10.1 non avvia una nuova operazione " & _
                     "prima che il file sia stato salvato consapevolmente."
            Exit Function
        End If
    End If

    If checkPendingEdits Then
        If Not OP_M10_CheckPendingEdits(detail) Then Exit Function
    End If

    OP_M10_PreflightOperational = True
End Function

Private Function OP_M10_QualifiedMacroName(ByVal macroName As String) As String
    OP_M10_QualifiedMacroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

Private Function OP_M10_ResultBoolean(ByVal result As Variant, _
                                       ByRef understood As Boolean) As Boolean
    understood = False

    If VarType(result) = vbBoolean Then
        understood = True
        OP_M10_ResultBoolean = CBool(result)
        Exit Function
    End If

    If IsNumeric(result) Then
        understood = True
        OP_M10_ResultBoolean = (CLng(result) <> 0)
    End If
End Function

Private Function OP_M10_RunCoreBoolean(ByVal macroName As String, _
                                        ByRef detail As String) As Boolean
    Dim result As Variant
    Dim localDetail As String
    Dim errOneArg As Long
    Dim errOneText As String
    Dim errNoArg As Long
    Dim errNoText As String
    Dim understood As Boolean
    Dim okResult As Boolean

    detail = vbNullString

    On Error Resume Next
    Err.Clear
    result = Application.Run(OP_M10_QualifiedMacroName(macroName), localDetail)
    errOneArg = Err.Number
    errOneText = Err.Description
    Err.Clear
    On Error GoTo 0

    If errOneArg = 0 Then
        okResult = OP_M10_ResultBoolean(result, understood)
        If understood Then
            If okResult Then
                detail = localDetail
                OP_M10_RunCoreBoolean = True
            Else
                If Len(localDetail) = 0 Then localDetail = macroName & " ha restituito False."
                detail = localDetail
            End If
            Exit Function
        End If
    End If

    localDetail = vbNullString
    result = Empty

    On Error Resume Next
    Err.Clear
    result = Application.Run(OP_M10_QualifiedMacroName(macroName))
    errNoArg = Err.Number
    errNoText = Err.Description
    Err.Clear
    On Error GoTo 0

    If errNoArg = 0 Then
        okResult = OP_M10_ResultBoolean(result, understood)
        If understood Then
            If okResult Then
                detail = macroName & ": OK."
                OP_M10_RunCoreBoolean = True
            Else
                detail = macroName & " ha restituito False."
            End If
            Exit Function
        End If
    End If

    detail = macroName & " non eseguibile/verificabile. Con dettaglio [" & _
             CStr(errOneArg) & "]: " & errOneText & _
             "; senza dettaglio [" & CStr(errNoArg) & "]: " & errNoText
End Function

Private Function OP_M10_TryRunCount(ByVal macroName As String, _
                                     ByRef countValue As Long, _
                                     ByRef detail As String) As Boolean
    Dim result As Variant

    detail = vbNullString
    countValue = 0

    On Error GoTo EH
    result = Application.Run(OP_M10_QualifiedMacroName(macroName))

    If Not IsNumeric(result) Then
        detail = macroName & " non ha restituito un conteggio numerico."
        Exit Function
    End If

    countValue = CLng(result)
    If countValue < 0 Then
        detail = macroName & " ha restituito un conteggio negativo."
        Exit Function
    End If

    OP_M10_TryRunCount = True
    Exit Function

EH:
    detail = macroName & " non disponibile [" & CStr(Err.Number) & "]: " & Err.Description
End Function

Private Function OP_M10_TryRunCountCandidates(ByVal candidates As Variant, _
                                               ByRef countValue As Long, _
                                               ByRef detail As String) As Boolean
    Dim i As Long
    Dim oneDetail As String
    Dim allDetail As String

    countValue = 0
    detail = vbNullString

    For i = LBound(candidates) To UBound(candidates)
        If OP_M10_TryRunCount(CStr(candidates(i)), countValue, oneDetail) Then
            OP_M10_TryRunCountCandidates = True
            detail = CStr(candidates(i))
            Exit Function
        End If

        If Len(allDetail) > 0 Then allDetail = allDetail & vbCrLf
        allDetail = allDetail & oneDetail
    Next i

    detail = allDetail
End Function

Private Function OP_M10_CheckOnePending(ByVal domainName As String, _
                                          ByVal candidates As Variant, _
                                          ByRef detail As String) As Boolean
    Dim countValue As Long
    Dim apiDetail As String

    detail = vbNullString

    If Not OP_M10_TryRunCountCandidates(candidates, countValue, apiDetail) Then
        detail = "Preflight bozze non verificabile per " & domainName & "." & vbCrLf & apiDetail
        Exit Function
    End If

    If countValue > 0 Then
        detail = "Operazione bloccata: " & domainName & _
                 " contiene " & CStr(countValue) & _
                 " modifica/e visibile/i non ancora consolidate."
        Exit Function
    End If

    OP_M10_CheckOnePending = True
End Function

Public Function OP_M10_CheckPendingEdits(ByRef detail As String) As Boolean
    detail = vbNullString

    If Not OP_M10_CheckOnePending("M3/Protocolli", _
            Array("OP_M3_CountVisibleEdits", "OP_M3_VisibleEditCount"), detail) Then Exit Function

    If Not OP_M10_CheckOnePending("M4/Campi", _
            Array("OP_M4_CountPlusEdits", "OP_M4_CountVisibleEdits"), detail) Then Exit Function

    If Not OP_M10_CheckOnePending("M5/Scadenze", _
            Array("OP_M5_CountVisibleEdits", "OP_M5_VisibleEditCount"), detail) Then Exit Function

    If Not OP_M10_CheckOnePending("M7/Decreti", _
            Array("OP_M7_CountVisibleEdits", "OP_M7_VisibleEditCount"), detail) Then Exit Function

    If Not OP_M10_CheckOnePending("M8/Incidenti", _
            Array("OP_M8_CountVisibleEdits", "OP_M8_VisibleEditCount"), detail) Then Exit Function

    If Not OP_M10_CheckOnePending("M9/Fascicoli", _
            Array("OP_M9_CountVisibleEdits", "OP_M9_VisibleEditCount"), detail) Then Exit Function

    OP_M10_CheckPendingEdits = True
End Function

Private Function OP_M10_VerifyView(ByVal sheetName As String, _
                                    ByRef detail As String) As Boolean
    Dim fp As String

    detail = vbNullString

    If Not OP_M10_SheetExists(sheetName) Then
        detail = "Vista richiesta assente: " & sheetName
        Exit Function
    End If

    fp = OP_M10_ObjectFingerprint(sheetName)
    If Not OP_M10_FingerprintIsValid(fp) Then
        detail = "Vista '" & sheetName & "' non verificabile dopo il refresh: " & fp
        Exit Function
    End If

    OP_M10_VerifyView = True
End Function

Private Function OP_M10_RunOneCore(ByVal macroName As String, _
                                    ByVal verifySheet As String, _
                                    ByRef allDetail As String) As Boolean
    Dim oneDetail As String
    Dim discardDetail As String

    If Not OP_M10_CheckDiscard(discardDetail) Then
        allDetail = allDetail & discardDetail
        Exit Function
    End If

    If Not OP_M10_RunCoreBoolean(macroName, oneDetail) Then
        If Len(allDetail) > 0 Then allDetail = allDetail & vbCrLf
        allDetail = allDetail & macroName & ": " & oneDetail
        Exit Function
    End If

    If Len(verifySheet) > 0 Then
        If Not OP_M10_VerifyView(verifySheet, oneDetail) Then
            If Len(allDetail) > 0 Then allDetail = allDetail & vbCrLf
            allDetail = allDetail & macroName & ": " & oneDetail
            Exit Function
        End If
    End If

    If Len(allDetail) > 0 Then allDetail = allDetail & vbCrLf
    allDetail = allDetail & macroName & ": OK"

    OP_M10_RunOneCore = True
End Function

Private Function OP_M10_RefreshAllOperational(ByRef detail As String) As Boolean
    Dim allDetail As String

    detail = vbNullString

    If Not OP_M10_RunOneCore("OP_M3_RefreshViewCore", "OP_PROTOCOLLI", allDetail) Then GoTo Failed
    If Not OP_M10_RunOneCore("OP_M4_RefreshProtocolPlusCore", "OP_PROTOCOLLI_PLUS", allDetail) Then GoTo Failed
    If Not OP_M10_RunOneCore("OP_M5_SyncProtocolRemindersCore", vbNullString, allDetail) Then GoTo Failed
    If Not OP_M10_RunOneCore("OP_M5_RefreshViewCore", "OP_SCADENZE", allDetail) Then GoTo Failed
    If Not OP_M10_RunOneCore("OP_M7_RefreshViewCore", "OP_DECRETI", allDetail) Then GoTo Failed
    If Not OP_M10_RunOneCore("OP_M8_RefreshViewCore", "OP_INCIDENTI", allDetail) Then GoTo Failed
    If Not OP_M10_RunOneCore("OP_M9_RefreshViewCore", "OP_FASCICOLI", allDetail) Then GoTo Failed

    detail = allDetail
    OP_M10_RefreshAllOperational = True
    Exit Function

Failed:
    detail = allDetail
End Function

Public Function OP_M10_RiallineaOperativo(ByVal requireSavedWorkbook As Boolean, _
                                           ByRef detail As String) As Boolean
    Dim refreshDetail As String
    Dim outputDetail As String

    detail = vbNullString
    On Error GoTo EH

    If Not OP_M10_PreflightOperational(requireSavedWorkbook, requireSavedWorkbook, detail) Then
        Exit Function
    End If

    If Not OP_M10_RefreshAllOperational(refreshDetail) Then
        detail = "Riallineamento operativo non completato." & vbCrLf & refreshDetail
        OP_M10_MarkAlignmentFailure detail
        Exit Function
    End If

    If Not OP_M10_RicostruisciOutput(outputDetail) Then
        detail = "Le viste sono state ricaricate, ma la sintesi operativa non e' stata ricostruita." & _
                 vbCrLf & outputDetail
        OP_M10_MarkAlignmentFailure detail
        Exit Function
    End If

    OP_M10_MarkAlignmentSuccess refreshDetail

    If Not OP_M10_UpdateValidityBanner(outputDetail) Then
        detail = "Allineamento completato, ma il banner di validita' non e' aggiornabile." & _
                 vbCrLf & outputDetail
        OP_M10_MarkAlignmentFailure detail
        Exit Function
    End If

    detail = refreshDetail
    OP_M10_RiallineaOperativo = True
    Exit Function

EH:
    detail = "Riallineamento operativo M10.1 non riuscito [" & CStr(Err.Number) & "]: " & Err.Description
    On Error Resume Next
    OP_M10_MarkAlignmentFailure detail
    On Error GoTo 0
End Function

Public Sub OP_M10_RIALLINEA()
    Dim detail As String

    If OP_M10_RiallineaOperativo(True, detail) Then
        MsgBox "BETA M10.1: RIALLINEAMENTO OPERATIVO COMPLETATO." & vbCrLf & vbCrLf & _
               "Il workbook NON e' stato salvato automaticamente." & vbCrLf & _
               "Verifica le viste e salva con Ctrl+S." & vbCrLf & vbCrLf & detail, _
               vbInformation, "WebArch Beta M10.1"
    Else
        MsgBox "BETA M10.1: RIALLINEAMENTO NON COMPLETATO." & vbCrLf & vbCrLf & detail, _
               vbCritical, "WebArch Beta M10.1"
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

    If Not OP_M10_PreflightOperational(True, True, detail) Then
        MsgBox detail, vbExclamation, "WebArch Beta M10.1"
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
    ExecutionLogBegin "BETA M10.1 - IMPORTA FONTI + OPERATIVO"
    pipelineStarted = True

    If Not MINIMAL_ExecutePipeline(detail) Then GoTo Failed
    dataWritten = True

    If Not OP_M10_RiallineaOperativo(False, refreshDetail) Then
        detail = "La fonte e' stata aggiornata, ma il riallineamento operativo M10.1 e' fallito." & _
                 vbCrLf & refreshDetail
        MINIMAL_RequireDiscard
        On Error Resume Next
        OP_M10_MarkAlignmentFailure detail
        On Error GoTo 0
        GoTo Failed
    End If

    If Not OP_M10_CheckDiscard(refreshDetail) Then
        detail = refreshDetail
        GoTo Failed
    End If

    detail = detail & vbCrLf & "Riallineamento operativo M10.1: OK."
    SetOperationStatus "WebArch M10.1 - salvataggio finale..."
    ExecutionLogEnd True, "Importazione e riallineamento M10.1 completati; salvataggio finale.", _
                    detail, "", "", True

    ' ExecutionLogEnd/SetOperationStatus possono aggiornare metadati alpha.
    ' Certifichiamo di nuovo lo stato finale immediatamente prima dell'unico Save.
    OP_M10_MarkAlignmentSuccess refreshDetail & vbCrLf & "Log finale alpha registrato."
    If Not OP_M10_UpdateValidityBanner(refreshDetail) Then
        detail = "Impossibile certificare il banner finale prima del Save." & vbCrLf & refreshDetail
        MINIMAL_RequireDiscard
        OP_M10_MarkAlignmentFailure detail
        GoTo Failed
    End If

    Application.DisplayAlerts = True
    saveAttempted = True
    ThisWorkbook.Save
    If Not ThisWorkbook.Saved Then
        Err.Raise vbObjectError + 6110, "OP_M10_IMPORTA_FONTI", _
                  "Excel non ha confermato il salvataggio finale M10.1."
    End If

    success = True
    GoTo Cleanup

Failed:
    On Error Resume Next
    ExecutionLogEnd False, "BETA M10.1 interrotta.", detail
    On Error GoTo 0
    GoTo Cleanup

EH:
    errNumber = Err.Number
    errDescription = Err.Description
    detail = "Errore M10.1 [" & CStr(errNumber) & "]: " & errDescription

    If dataWritten Or saveAttempted Then
        On Error Resume Next
        MINIMAL_RequireDiscard
        OP_M10_MarkAlignmentFailure detail
        On Error GoTo 0
    End If

    On Error Resume Next
    ExecutionLogRecordFailure detail
    On Error GoTo 0

Cleanup:
    On Error Resume Next
    If excelStateCaptured Then RestoreExcelState
    If guardHeld Then EASYONE_EndOperation "AGGIORNA_TUTTO"
    On Error GoTo 0

    If success Then
        MsgBox "BETA M10.1: IMPORTAZIONE FONTI + OPERATIVO COMPLETATA E SALVATA." & _
               vbCrLf & vbCrLf & detail, vbInformation, "WebArch Beta M10.1"
    ElseIf pipelineStarted Then
        If MINIMAL_DiscardRequired() Then
            MsgBox detail & vbCrLf & vbCrLf & MINIMAL_DiscardMessage(), _
                   vbCritical, "WebArch Beta M10.1"
        Else
            MsgBox "BETA M10.1 interrotta prima del salvataggio finale." & vbCrLf & vbCrLf & detail, _
                   vbExclamation, "WebArch Beta M10.1"
        End If
    End If
End Sub

Public Sub OP_M10_IMPORTA_FASCICOLI()
    Dim detail As String
    Dim beforeFsc As String
    Dim afterFsc As String
    Dim refreshDetail As String
    Dim saveAttempted As Boolean
    Dim fscChanged As Boolean

    On Error GoTo EH

    If Not OP_M10_PreflightOperational(True, True, detail) Then
        MsgBox detail, vbExclamation, "WebArch Beta M10.1"
        Exit Sub
    End If

    If Not MINIMAL_PrepareWorkbook(detail) Then
        MsgBox detail, vbExclamation, WA_APP_NAME
        Exit Sub
    End If

    beforeFsc = OP_M10_FscStateToken(detail)
    If Len(beforeFsc) = 0 Then
        MsgBox detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    ' B1 conserva il proprio comando UI e il proprio BUSY. M10.1 non annida un secondo guard.
    ' La validita' viene verificata dopo il ritorno usando tutto lo stato FSC_* e il discard MINIMAL.
    Application.Run OP_M10_QualifiedMacroName("FSC_Aggiorna")

    If Not OP_M10_CheckDiscard(detail) Then
        On Error Resume Next
        OP_M10_MarkAlignmentFailure detail
        On Error GoTo 0
        MsgBox detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    afterFsc = OP_M10_FscStateToken(detail)
    If Len(afterFsc) = 0 Then
        OP_M10_MarkAlignmentFailure detail
        MsgBox detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    fscChanged = (StrComp(beforeFsc, afterFsc, vbBinaryCompare) <> 0)

    If Not fscChanged Then
        MsgBox "BETA M10.1: nessun cambiamento durevole verificabile nei fogli FSC_*." & vbCrLf & _
               "L'operazione puo' essere stata annullata oppure la fotografia era gia' allineata. " & _
               "Nessun Save M10.1 eseguito.", vbInformation, "WebArch Beta M10.1"
        Exit Sub
    End If

    If Not OP_M10_RiallineaOperativo(False, refreshDetail) Then
        detail = "FSC ha modificato lo stato B1, ma il riallineamento operativo e' fallito." & _
                 vbCrLf & refreshDetail
        MINIMAL_RequireDiscard
        On Error Resume Next
        OP_M10_MarkAlignmentFailure detail
        On Error GoTo 0
        MsgBox detail & vbCrLf & vbCrLf & MINIMAL_DiscardMessage(), _
               vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    If Not OP_M10_CheckDiscard(detail) Then
        MsgBox detail, vbCritical, "WebArch Beta M10.1"
        Exit Sub
    End If

    saveAttempted = True
    ThisWorkbook.Save
    If Not ThisWorkbook.Saved Then
        Err.Raise vbObjectError + 6111, "OP_M10_IMPORTA_FASCICOLI", _
                  "Excel non ha confermato il salvataggio finale dopo FSC."
    End If

    MsgBox "BETA M10.1: FSC + OPERATIVO RIALLINEATI E SALVATI.", _
           vbInformation, "WebArch Beta M10.1"
    Exit Sub

EH:
    detail = "Importazione FSC M10.1 non riuscita [" & CStr(Err.Number) & "]: " & Err.Description
    If fscChanged Or saveAttempted Then
        On Error Resume Next
        MINIMAL_RequireDiscard
        OP_M10_MarkAlignmentFailure detail
        detail = detail & vbCrLf & MINIMAL_DiscardMessage()
        On Error GoTo 0
    End If
    MsgBox detail, vbCritical, "WebArch Beta M10.1"
End Sub
