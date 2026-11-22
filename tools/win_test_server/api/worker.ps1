# Embedded in the Go executable. Run in the logged-on desktop, one fresh Office instance per job.
# API references (Microsoft Learn; implementation is original):
# https://learn.microsoft.com/en-us/office/vba/api/word.documents.open
# https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open
# https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.displayalerts
# https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat
param([Parameter(Mandatory=$true)][string]$Directory)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$state = @{ phase = 'starting'; opened = $false; exported = $false; officePid = 0; officeVersion = ''; officeBuild = ''; application = ''; error = '' }
function Checkpoint([string]$phase) {
    $state.phase = $phase
    [IO.File]::WriteAllText("$Directory\state.tmp", ($state | ConvertTo-Json -Depth 6), $utf8)
    Move-Item -LiteralPath "$Directory\state.tmp" -Destination "$Directory\state.json" -Force
}
$app = $null
$document = $null
$oldUpdateLinks = $null
$ext = ''
try {
    Checkpoint 'starting'
    $inputFile = @(Get-ChildItem -LiteralPath $Directory -Filter 'input.*')[0]
    $ext = $inputFile.Extension.ToLowerInvariant()
    $outputFile = "$Directory\output.pdf"
    $missing = [Type]::Missing
    switch ($ext) {
        '.docx' { $progId = 'Word.Application'; $processName = 'WINWORD'; $state.application = 'Word' }
        '.xlsx' { $progId = 'Excel.Application'; $processName = 'EXCEL'; $state.application = 'Excel' }
        '.pptx' { $progId = 'PowerPoint.Application'; $processName = 'POWERPNT'; $state.application = 'PowerPoint' }
        default { throw 'Unsupported extension' }
    }
    $before = @(Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $session = (Get-Process -Id $PID).SessionId
    $app = New-Object -ComObject $progId
    $owned = @(Get-Process -Name $processName | Where-Object { $_.SessionId -eq $session -and $_.Id -notin $before })
    if ($owned.Count -ne 1) { throw 'Cannot identify a fresh Office process; refusing an unobserved open' }
    $state.officePid = $owned[0].Id
    Checkpoint 'starting'
    # PowerPoint defaults to ppAlertsNone, even in a visible window. Always explicitly enable alerts.
    if ($ext -eq '.pptx') { $app.DisplayAlerts = 2; $app.Visible = -1 }
    elseif ($ext -eq '.docx') { $app.DisplayAlerts = -1; $app.Visible = $true }
    else { $app.DisplayAlerts = $true; $app.Visible = $true; $app.EnableEvents = $false }
    $app.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
    $state.officeVersion = [string]$app.Version
    $state.officeBuild = $owned[0].MainModule.FileVersionInfo.FileVersion
    if (!$state.officeVersion.StartsWith('16.0')) { throw "Office 2016 (16.0) required; found $($state.officeVersion)" }
    Checkpoint 'opening'
    switch ($ext) {
        '.docx' {
            $oldUpdateLinks = $app.Options.UpdateLinksAtOpen
            $app.Options.UpdateLinksAtOpen = $false
            # Normal Open, NOT OpenNoRepairDialog. Argument 14 is OpenAndRepair=false.
            $document = $app.Documents.Open($inputFile.FullName, $false, $true, $false,
                $missing, $missing, $false, $missing, $missing, $missing, $missing, $true, $missing, $false)
        }
        '.xlsx' {
            # UpdateLinks=0, ReadOnly=true; omitted CorruptLoad defaults to xlNormalLoad (0),
            # which does not attempt recovery through the object model. Never retry with repair/extract.
            $document = $app.Workbooks.Open($inputFile.FullName, 0, $true)
        }
        '.pptx' { $document = $app.Presentations.Open($inputFile.FullName, -1, 0, -1) }
    }
    if ($null -eq $document) { throw 'Office returned no document' }
    $state.opened = $true
    Checkpoint 'opened'
    Start-Sleep -Milliseconds 1500 # Let the independent window observer see delayed repair dialogs.
    Checkpoint 'exporting'
    switch ($ext) {
        '.docx' { $document.ExportAsFixedFormat($outputFile, 17) } # wdExportFormatPDF
        '.xlsx' { $document.ExportAsFixedFormat(0, $outputFile, 0, $true, $false) } # all printable sheets, respect print areas
        '.pptx' {
            # PDF, print intent, slides (not handouts), include hidden slides, all slides. A range is in slide
            # numbers, which start at the deck's first slide number (5 for a deck numbered from 5), not at 1.
            $first = $document.PageSetup.FirstSlideNumber
            $range = $document.PrintOptions.Ranges.Add($first, $first + $document.Slides.Count - 1)
            $document.ExportAsFixedFormat($outputFile, 2, 2, 0, 1, 1, -1, $range, 1, '', $true, $true, $true, $true, $false)
        }
    }
    $state.exported = $true
    Checkpoint 'exported'
    Start-Sleep -Milliseconds 750
} catch {
    $state.failedPhase = $state.phase
    $state.error = $_.Exception.ToString()
    Checkpoint 'error'
} finally {
    try {
        if ($null -ne $document) {
            if ($ext -eq '.pptx') { $document.Saved = -1; $document.Close() }
            elseif ($ext -eq '.docx') { $document.Close(0) }
            else { $document.Close($false) }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        }
        if ($null -ne $app) {
            if ($null -ne $oldUpdateLinks) { $app.Options.UpdateLinksAtOpen = $oldUpdateLinks }
            $app.Quit()
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        }
    } catch {
        $state.error += "`nCleanup: " + $_.Exception.Message
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Checkpoint 'done'
}
