param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$template = Join-Path $root "Cruscotto_Gabinetto_WebArch_v1.1_TEMPLATE.xlsx"
$output = Join-Path $root "Cruscotto_Gabinetto_WebArch_v1.1.xlsm"
$vbaFolder = Join-Path $root "VBA"
$logPath = Join-Path $root "BUILD\Build-WebArch.log"
$builderVersion = "1.1.1"

$moduleFiles = @(
    "modUtilities.bas",
    "modSourceRecognition.bas",
    "modImport.bas",
    "modCriticita.bas",
    "modProtocollo.bas",
    "modFascicoli.bas",
    "modKPI.bas",
    "modReport.bas",
    "modCollaudo.bas",
    "modExport.bas",
    "modSecurity.bas",
    "modMain.bas"
)

function Write-Log([string]$Text) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Text"
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Release-ComObject($Object) {
    if ($null -ne $Object) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
Write-Log "Avvio build WebArch v1.1 - builder $builderVersion."

if (-not (Test-Path -LiteralPath $template)) { throw "Template non trovato: $template" }
foreach ($module in $moduleFiles) {
    $modulePath = Join-Path $vbaFolder $module
    if (-not (Test-Path -LiteralPath $modulePath)) { throw "Modulo VBA mancante: $modulePath" }
}
$thisWorkbookCodePath = Join-Path $vbaFolder "ThisWorkbook.code.txt"
if (-not (Test-Path -LiteralPath $thisWorkbookCodePath)) { throw "Codice ThisWorkbook mancante: $thisWorkbookCodePath" }

if (Test-Path -LiteralPath $output) {
    if (-not $Force) {
        $backupName = "Cruscotto_Gabinetto_WebArch_v1.1_prebuild_{0}.xlsm" -f (Get-Date -Format "yyyyMMdd_HHmmss")
        $backupPath = Join-Path $root $backupName
        Copy-Item -LiteralPath $output -Destination $backupPath -Force
        Write-Log "Copia della build precedente: $backupPath"
    }
    Remove-Item -LiteralPath $output -Force
}

$excel = $null
$workbook = $null
$vbProject = $null
$vbComponents = $null
$thisWorkbookComponent = $null
$codeModule = $null
$importedComponent = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    Write-Log "Apertura del template XLSX."
    $workbook = $excel.Workbooks.Open($template, 0, $false)

    # Hotfix 1.1.1: in alcuni Excel un XLSX espone VBProject ma VBComponents e' $null.
    # Convertiamo prima in XLSM, chiudiamo e riapriamo, poi importiamo il codice VBA.
    Write-Log "Conversione preliminare del template in XLSM."
    $workbook.SaveAs($output, 52)
    $workbook.Close($true)
    Release-ComObject $workbook
    $workbook = $null
    Start-Sleep -Milliseconds 500

    Write-Log "Riapertura del workbook XLSM."
    $workbook = $excel.Workbooks.Open($output, 0, $false)

    try {
        $vbProject = $workbook.VBProject
        if ($null -eq $vbProject) { throw "VBProject non disponibile." }
        $vbComponents = $vbProject.VBComponents
        if ($null -eq $vbComponents) { throw "VBComponents non disponibile dopo la conversione XLSM." }
        $componentCount = [int]$vbComponents.Count
        if ($componentCount -lt 1) { throw "La raccolta VBComponents e' vuota dopo la conversione XLSM." }
        Write-Log "Accesso al progetto VBA disponibile. Componenti iniziali: $componentCount"
    }
    catch {
        throw @"
Excel non consente di inizializzare correttamente il progetto VBA.
Verificare:
File > Opzioni > Centro protezione > Impostazioni Centro protezione >
Impostazioni macro > Considera attendibile l'accesso al modello a oggetti dei progetti VBA.

Dopo la build l'opzione puo' essere nuovamente disabilitata.
Dettaglio: $($_.Exception.Message)
"@
    }

    foreach ($moduleFile in $moduleFiles) {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($moduleFile)
        $oldComponent = $null
        try { $oldComponent = $vbComponents.Item($moduleName) } catch { $oldComponent = $null }
        if ($null -ne $oldComponent -and $oldComponent.Type -ne 100) {
            $vbComponents.Remove($oldComponent)
            Release-ComObject $oldComponent
            $oldComponent = $null
        }

        $modulePath = Join-Path $vbaFolder $moduleFile
        try { $importedComponent = $vbComponents.Import($modulePath) }
        catch { throw "Importazione VBA fallita per '$moduleFile'. Dettaglio COM: $($_.Exception.Message)" }
        if ($null -eq $importedComponent) { throw "L'importazione di '$moduleFile' non ha restituito un componente VBA valido." }
        Write-Log "Importato: $moduleFile -> $($importedComponent.Name)"
        Release-ComObject $importedComponent
        $importedComponent = $null
    }

    $thisWorkbookComponent = $vbComponents.Item("ThisWorkbook")
    if ($null -eq $thisWorkbookComponent) { throw "Componente ThisWorkbook non trovato." }
    $codeModule = $thisWorkbookComponent.CodeModule
    if ($null -eq $codeModule) { throw "CodeModule di ThisWorkbook non disponibile." }
    if ($codeModule.CountOfLines -gt 0) { $codeModule.DeleteLines(1, $codeModule.CountOfLines) }

    $encoding = [System.Text.Encoding]::GetEncoding(1252)
    $thisWorkbookCode = [System.IO.File]::ReadAllText($thisWorkbookCodePath, $encoding)
    $codeModule.AddFromString($thisWorkbookCode)
    Write-Log "Aggiornato il modulo ThisWorkbook."
    $workbook.Save()

    $excel.EnableEvents = $true
    $excel.ScreenUpdating = $true

    Write-Log "Esecuzione inizializzazione interfaccia e protezioni."
    $macroName = "'{0}'!InitializeWebArchUI" -f $workbook.Name
    [void]$excel.Run($macroName)

    Write-Log "Esecuzione smoke test VBA su tutti i motori applicativi."
    $smokeMacro = "'{0}'!BuildSmokeTest" -f $workbook.Name
    $smokeResult = [string]$excel.Run($smokeMacro)
    if ($smokeResult -ne "OK") { throw "Smoke test VBA non superato: $smokeResult" }
    Write-Log "Smoke test VBA: OK"

    $workbook.Save()
    Write-Log "Workbook operativo salvato: $output"
    if (-not (Test-Path -LiteralPath $output)) { throw "La build non ha prodotto il file XLSM." }
    $size = (Get-Item -LiteralPath $output).Length
    if ($size -le 0) { throw "Il file XLSM prodotto e' vuoto." }
    Write-Log "Build completata. Dimensione: $size byte."
}
catch {
    Write-Log "ERRORE: $($_.Exception.Message)"
    if ($null -ne $_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber -gt 0) {
        Write-Log "RIGA POWERSHELL: $($_.InvocationInfo.ScriptLineNumber)"
        if ($null -ne $_.InvocationInfo.Line) { Write-Log "COMANDO: $($_.InvocationInfo.Line.Trim())" }
    }
    if (-not [string]::IsNullOrWhiteSpace($_.ScriptStackTrace)) {
        Write-Log "STACK: $($_.ScriptStackTrace -replace "`r?`n", " | ")"
    }

    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; $workbook = $null }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $importedComponent
    Release-ComObject $codeModule
    Release-ComObject $thisWorkbookComponent
    Release-ComObject $vbComponents
    Release-ComObject $vbProject
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if (Test-Path -LiteralPath $output) {
        $failedPath = "$output.FAILED"
        Move-Item -LiteralPath $output -Destination $failedPath -Force -ErrorAction SilentlyContinue
        Write-Log "La build incompleta e' stata isolata come: $failedPath"
    }

    Write-Host ""
    Write-Host "BUILD NON COMPLETATA" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($true) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $importedComponent
    Release-ComObject $codeModule
    Release-ComObject $thisWorkbookComponent
    Release-ComObject $vbComponents
    Release-ComObject $vbProject
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "BUILD COMPLETATA" -ForegroundColor Green
Write-Host "File creato: $output" -ForegroundColor Green
Write-Host "Aprire il file XLSM e provare prima DIAGNOSTICA FONTI, poi AGGIORNA TUTTO." -ForegroundColor Cyan
exit 0
