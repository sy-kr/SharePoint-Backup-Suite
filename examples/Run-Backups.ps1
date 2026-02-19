<#
.SYNOPSIS
    Orchestrator script - runs multiple SharePoint backups and sends an
    email report on completion.

.DESCRIPTION
    Cross-platform orchestrator (Windows + Linux/macOS).  Runs each backup
    job sequentially, captures exit codes, timing, and per-job log files,
    then sends a single summary email via SMTP (with STARTTLS).

    Copy this file to the same directory as spbackup.ps1 and edit the
    CONFIGURATION section.  On Windows, also copy Load-Credentials.ps1
    from examples/windows/ to load credentials from Credential Manager.
    On Linux, set environment variables before running (e.g. via an env
    file or wrapper script).

    By default (interactive), job output is streamed to both the console and
    per-job log files via Tee-Object, so --verbose output is visible in real
    time.

    Use -Headless when running from Task Scheduler, cron, or other
    non-interactive hosts.  This suppresses all console output (including
    progress bars), writing only to log files.  The report is still saved
    to disk and emailed.

    Exit codes from spbackup.ps1:
      0  - success (all items backed up)
      1  - partial failure (some downloads / exports failed) or fatal error
      2  - verify mismatch

    This script's own exit code:
      0  - every job succeeded (exit code 0)
      1  - one or more jobs returned non-zero

.EXAMPLE
    # Interactive — verbose job output streams to console and log file:
    pwsh -NoProfile -File "C:\spbackup\Run-Backups.ps1"

    # Headless (Task Scheduler / cron) — all output goes to log files only:
    pwsh -NoProfile -File "C:\spbackup\Run-Backups.ps1" -Headless

.NOTES
    Edit the CONFIGURATION section below before first use.
#>
param(
    [switch]$Headless
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Suppress progress bars and console output in headless mode to prevent
# pipeline congestion under Task Scheduler or other non-interactive hosts.
# The report is still written to a log file and emailed.
if ($Headless) {
    $ProgressPreference = 'SilentlyContinue'
}

function Write-Console {
    <# Write-Host wrapper that is suppressed in headless mode. #>
    param(
        [Parameter(Position = 0)] [object]$Object = '',
        [string]$ForegroundColor
    )
    if ($Headless) { return }
    $p = @{}
    if ($ForegroundColor) { $p['ForegroundColor'] = $ForegroundColor }
    Write-Host $Object @p
}

function Get-JobFileChanges {
    <#
    .SYNOPSIS
        Parse a job log for file-level activity (downloads, exports, deletions).
        Requires --verbose on the backup job so [LEVEL] event lines are present.
    .OUTPUTS
        Ordered hashtable with Downloaded, Exported, Deleted (count), Failed lists.
    #>
    param(
        [string]$JobLog,
        [string]$Tool
    )
    $changes = [ordered]@{
        Downloaded = [System.Collections.Generic.List[string]]::new()
        Exported   = [System.Collections.Generic.List[string]]::new()
        Deleted    = 0
        Failed     = [System.Collections.Generic.List[string]]::new()
    }
    if (-not $JobLog -or -not (Test-Path -LiteralPath $JobLog)) { return $changes }

    try {
        switch ($Tool) {
            'library' {
                foreach ($line in [System.IO.File]::ReadLines($JobLog)) {
                    if ($line -match '\] file_downloaded - Downloaded: (.+) \(\d+ bytes\)') {
                        $changes.Downloaded.Add($Matches[1])
                    }
                    elseif ($line -match '\] delta_deletions - (\d+) file') {
                        $changes.Deleted += [int]$Matches[1]
                    }
                    elseif ($line -match '\] file_download_deleted - Skipping \(deleted.+?\): (.+)') {
                        $changes.Deleted++
                    }
                    elseif ($line -match '\] file_download_fail - .+?: (.+?) —') {
                        $changes.Failed.Add($Matches[1])
                    }
                }
            }
            'loop' {
                $lastExportName = ''
                foreach ($line in [System.IO.File]::ReadLines($JobLog)) {
                    if ($line -match '\] export_start - Exporting: (.+)') {
                        $lastExportName = $Matches[1].Trim()
                    }
                    elseif ($line -match '\] export_complete - Done: (.+)') {
                        $changes.Exported.Add($Matches[1].Trim())
                    }
                    elseif ($line -match '\] export_meta_fail - ' -or $line -match '\] export_item_error - ') {
                        $failName = if ($lastExportName) { $lastExportName } else { '(unknown)' }
                        if (-not $changes.Failed.Contains($failName)) {
                            $changes.Failed.Add($failName)
                        }
                    }
                }
            }
            'list' {
                foreach ($line in [System.IO.File]::ReadLines($JobLog)) {
                    if ($line -match '\] download_attachment_success - Downloaded: (.+)') {
                        $changes.Downloaded.Add($Matches[1].Trim())
                    }
                    elseif ($line -match '\] items_fetched - Found (\d+) items') {
                        $count = [int]$Matches[1]
                        if ($count -gt 0) {
                            $changes.Exported.Add("$count list item(s) exported to CSV")
                        }
                    }
                }
            }
        }
    } catch {
        # Non-fatal — changes section will be empty
    }
    return $changes
}

function Format-FileChangeLines {
    <#
    .SYNOPSIS
        Render file-change info as report lines.  Returns nothing if no activity.
    #>
    param([System.Collections.Specialized.OrderedDictionary]$Changes, [int]$MaxItems = 25)
    $lines = [System.Collections.Generic.List[string]]::new()
    $hasAny = ($Changes.Downloaded.Count + $Changes.Exported.Count + $Changes.Failed.Count + $Changes.Deleted) -gt 0
    if (-not $hasAny) { return $lines }

    if ($Changes.Downloaded.Count -gt 0) {
        $lines.Add("      Downloaded ($($Changes.Downloaded.Count)):")
        $shown = [math]::Min($Changes.Downloaded.Count, $MaxItems)
        for ($i = 0; $i -lt $shown; $i++) {
            $lines.Add("        - $($Changes.Downloaded[$i])")
        }
        if ($Changes.Downloaded.Count -gt $MaxItems) {
            $lines.Add("        ... and $($Changes.Downloaded.Count - $MaxItems) more")
        }
    }
    if ($Changes.Exported.Count -gt 0) {
        $lines.Add("      Exported ($($Changes.Exported.Count)):")
        $shown = [math]::Min($Changes.Exported.Count, $MaxItems)
        for ($i = 0; $i -lt $shown; $i++) {
            $lines.Add("        - $($Changes.Exported[$i])")
        }
        if ($Changes.Exported.Count -gt $MaxItems) {
            $lines.Add("        ... and $($Changes.Exported.Count - $MaxItems) more")
        }
    }
    if ($Changes.Deleted -gt 0) {
        $lines.Add("      Deleted on SharePoint: $($Changes.Deleted)")
    }
    if ($Changes.Failed.Count -gt 0) {
        $lines.Add("      Failed ($($Changes.Failed.Count)):")
        $shown = [math]::Min($Changes.Failed.Count, $MaxItems)
        for ($i = 0; $i -lt $shown; $i++) {
            $lines.Add("        - $($Changes.Failed[$i])")
        }
        if ($Changes.Failed.Count -gt $MaxItems) {
            $lines.Add("        ... and $($Changes.Failed.Count - $MaxItems) more")
        }
    }
    return $lines
}

# -----------------------------------------------------------------------------
# CONFIGURATION - edit these values
# -----------------------------------------------------------------------------

# Where spbackup.ps1 lives (same directory as this script by default)
$SpBackupRoot = $PSScriptRoot

# Base output directory (each job appends its own subfolder)
# Windows example: 'D:\SharePointBackups'
# Linux example:   '/var/lib/spbackup'
$BackupBase = '/var/lib/spbackup'

# SMTP settings (set $SmtpServer to '' to disable email)
$SmtpServer   = 'smtp.example.com'
$SmtpPort     = 587
$SmtpFrom     = 'backups@example.com'
$SmtpUser     = 'backups@example.com'                     # set to '' for anonymous relay
$SmtpPass     = ''                                        # set to '' for anonymous relay
$SmtpUseSsl   = $true                                     # STARTTLS
$SmtpTo       = @(
    'admin@example.com'
    # Add more recipients here
)

# Backup jobs - each entry is a hashtable with the spbackup.ps1 arguments.
# Comment out or remove jobs you don't need.
#
# NOTE: Avoid --verbose for scheduled (unattended) runs.  Verbose output is
#       useful interactively but generates huge log files and can cause
#       pipeline backpressure that dramatically slows the backup.
$Jobs = @(
    # --- Document library backups ---
    @{
        Label = 'Team Site - Documents'
        Tool  = 'library'
        Args  = @('backup',
                   '--url',     'https://contoso.sharepoint.com/sites/TeamSite',
                   '--library', 'Documents',
                   '--out',     (Join-Path $BackupBase 'TeamSite'))
    }
    @{
        Label = 'HR - Shared Documents'
        Tool  = 'library'
        Args  = @('backup',
                   '--url',     'https://contoso.sharepoint.com/sites/HR',
                   '--library', 'Shared Documents',
                   '--out',     (Join-Path $BackupBase 'HR'))
    }
    # --- Microsoft List backups ---
    @{
        Label = 'Project Tasks (List)'
        Tool  = 'list'
        Args  = @('backup',
                   '--url',     'https://contoso.sharepoint.com/sites/Projects/Lists/Tasks/AllItems.aspx',
                   '--out',     (Join-Path $BackupBase 'Lists' 'ProjectTasks'))
    }
    # --- Microsoft Loop backups ---
    # @{
    #     Label = 'Team Wiki (Loop)'
    #     Tool  = 'loop'
    #     Args  = @('backup',
    #                '--url',     'https://loop.cloud.microsoft/p/<your-loop-workspace-id>',
    #                '--out',     (Join-Path $BackupBase 'TeamWiki'))
    # }
)

# -----------------------------------------------------------------------------
# END CONFIGURATION
# -----------------------------------------------------------------------------

# -- Load credentials ---------------------------------------------------------
# Windows: load from Credential Manager via Load-Credentials.ps1
#          (copy from examples/windows/ to the same directory as this script)
# Linux:   set TENANT_ID, CLIENT_ID, CLIENT_SECRET as environment variables
#          before running (e.g. via /etc/spbackup/env or a wrapper script)
$credLoader = Join-Path $PSScriptRoot 'Load-Credentials.ps1'
if (Test-Path $credLoader) {
    . $credLoader
}

$spbackupScript = Join-Path $SpBackupRoot 'spbackup.ps1'
if (-not (Test-Path $spbackupScript)) {
    Write-Error "spbackup.ps1 not found at $spbackupScript"
    exit 1
}

$runStart   = Get-Date
$hostName   = $env:COMPUTERNAME ?? (hostname)
$dateStr    = $runStart.ToString('dd-MM-yyyy')
$results    = [System.Collections.Generic.List[object]]::new()
$anyFailure = $false

Write-Console ''
Write-Console "=== SharePoint Backup Orchestrator ===" -ForegroundColor Cyan
Write-Console "  Host:    $hostName"
Write-Console "  Date:    $dateStr"
Write-Console "  Jobs:    $($Jobs.Count)"
Write-Console "  Output:  $BackupBase"
Write-Console ''

# Ensure per-job log directory exists (used by output capture below)
$logDir = Join-Path $BackupBase 'orchestrator-logs'
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# -----------------------------------------------------------------------------
# Run each job
# -----------------------------------------------------------------------------
try {

foreach ($job in $Jobs) {
    $label    = $job.Label
    $tool     = $job.Tool
    $jobArgs  = $job.Args
    $jobStart = Get-Date

    Write-Console "[$dateStr] Starting: $label" -ForegroundColor Cyan

    # Build full argument list: tool + job-specific args
    $fullArgs = @($tool) + $jobArgs

    # Capture all output streams to a per-job log file.  In interactive mode
    # (default) Tee-Object also streams to the console so --verbose output is
    # visible in real time.  In headless mode, Out-File sends everything to
    # the log only — no console output.
    $safeLabel = $label -replace '[^a-zA-Z0-9\-]', '_'
    $jobLog    = Join-Path $logDir "job-${safeLabel}.log"

    try {
        if ($Headless) {
            & $spbackupScript @fullArgs *>&1 | Out-File -LiteralPath $jobLog -Encoding utf8
        } else {
            & $spbackupScript @fullArgs *>&1 | Tee-Object -LiteralPath $jobLog
        }
        $exitCode = $LASTEXITCODE
        if ($null -eq $exitCode) { $exitCode = 0 }
    } catch {
        $_.Exception.Message | Out-File -LiteralPath $jobLog -Append -Encoding utf8
        $exitCode = 1
    }

    $jobEnd     = Get-Date
    $duration   = $jobEnd - $jobStart
    $durationStr = '{0:hh\:mm\:ss}' -f $duration

    $status = switch ($exitCode) {
        0       { 'SUCCESS' }
        2       { 'VERIFY_MISMATCH' }
        default { 'FAILED' }
    }

    if ($exitCode -ne 0) { $anyFailure = $true }

    $color = switch ($exitCode) {
        0       { 'Green' }
        2       { 'Yellow' }
        default { 'Red' }
    }
    Write-Console "[$dateStr] $label - $status (exit $exitCode, $durationStr)" -ForegroundColor $color

    # Record result FIRST so it's captured even if summary parsing fails
    $result = [ordered]@{
        Label      = $label
        Tool       = $tool
        Status     = $status
        ExitCode   = $exitCode
        Duration   = $durationStr
        Summary    = ''
        Changes    = $null
        StartTime  = $jobStart.ToString('HH:mm:ss')
        EndTime    = $jobEnd.ToString('HH:mm:ss')
        JobLog     = $jobLog
    }
    $results.Add($result)

    # Extract summary line from job log (best-effort, non-fatal)
    try {
        if (Test-Path -LiteralPath $jobLog) {
            [array]$tailLines = @(Get-Content -LiteralPath $jobLog -Tail 30 |
                                  Where-Object { $_.Trim() -ne '' })
            if ($tailLines.Count -gt 0) {
                $bcLine = @($tailLines | Where-Object { $_ -match 'Backup complete:' }) |
                          Select-Object -Last 1
                $result.Summary = if ($bcLine) { $bcLine.Trim() } else { $tailLines[-1].Trim() }
            }
        }
    } catch {
        $result.Summary = '(could not parse output)'
    }

    # Extract file-level changes from job log (best-effort, non-fatal)
    $result.Changes = Get-JobFileChanges -JobLog $jobLog -Tool $tool
}

$runEnd      = Get-Date
$totalTime   = '{0:hh\:mm\:ss}' -f ($runEnd - $runStart)
$successCount = @($results | Where-Object { $_.ExitCode -eq 0 }).Count
$failCount    = @($results | Where-Object { $_.ExitCode -ne 0 }).Count

# -----------------------------------------------------------------------------
# Build report
# -----------------------------------------------------------------------------
$overallStatus = if ($anyFailure) { 'FAILURE' } else { 'SUCCESS' }

$reportLines = [System.Collections.Generic.List[string]]::new()
$reportLines.Add("SharePoint Backup Report - $hostName - $dateStr")
$reportLines.Add("=" * 60)
$reportLines.Add('')
$reportLines.Add("Overall:    $overallStatus")
$reportLines.Add("Jobs:       $($results.Count) total, $successCount succeeded, $failCount failed")
$reportLines.Add("Total time: $totalTime")
$reportLines.Add("Started:    $($runStart.ToString('dd-MM-yyyy HH:mm:ss'))")
$reportLines.Add("Finished:   $($runEnd.ToString('dd-MM-yyyy HH:mm:ss'))")
$reportLines.Add('')
$reportLines.Add("-" * 60)

foreach ($r in $results) {
    $icon = if ($r.ExitCode -eq 0) { '[OK]' } elseif ($r.ExitCode -eq 2) { '[!!]' } else { '[FAIL]' }
    $reportLines.Add('')
    $reportLines.Add("$icon $($r.Label)")
    $reportLines.Add("    Tool:     $($r.Tool)")
    $reportLines.Add("    Status:   $($r.Status) (exit code $($r.ExitCode))")
    $reportLines.Add("    Time:     $($r.StartTime) - $($r.EndTime) ($($r.Duration))")
    if ($r.Summary) {
        $reportLines.Add("    Summary:  $($r.Summary)")
    }
    $reportLines.Add("    Log:      $($r.JobLog)")

    # File-level changes (only when --verbose is in the job args)
    if ($r.Changes) {
        [array]$changeLines = @(Format-FileChangeLines -Changes $r.Changes)
        if ($changeLines.Count -gt 0) {
            $reportLines.Add('    Changes:')
            foreach ($cl in $changeLines) { $reportLines.Add($cl) }
        }
    }
}

$reportLines.Add('')
$reportLines.Add("-" * 60)
$reportLines.Add('')

# Add output tail for failed jobs
[array]$failedJobs = @($results | Where-Object { $_.ExitCode -ne 0 })
if ($failedJobs.Count -gt 0) {
    $reportLines.Add("DETAILED OUTPUT FOR FAILED JOBS")
    $reportLines.Add("=" * 60)

    foreach ($r in $failedJobs) {
        $reportLines.Add('')
        $reportLines.Add(">>> $($r.Label) (exit code $($r.ExitCode)) <<<")
        $reportLines.Add("    Log: $($r.JobLog)")
        $reportLines.Add('')
        if ($r.JobLog -and (Test-Path -LiteralPath $r.JobLog)) {
            [array]$logTail = @(Get-Content -LiteralPath $r.JobLog -Tail 100)
            if ($logTail.Count -ge 100) {
                $reportLines.Add('... (last 100 lines -- full log at path above) ...')
                $reportLines.Add('')
            }
            $reportLines.Add(($logTail -join "`r`n"))
        } else {
            $reportLines.Add('(no output captured)')
        }
        $reportLines.Add('')
    }
}

$reportText = $reportLines -join "`r`n"

# Print report to console
Write-Console ''
Write-Console $reportText

# -----------------------------------------------------------------------------
# Write report to file
# -----------------------------------------------------------------------------
$logFile = Join-Path $logDir "backup-report-$($runStart.ToString('ddMMyyyy-HHmmss')).txt"
$reportText | Out-File -FilePath $logFile -Encoding utf8
Write-Console "Report saved to: $logFile" -ForegroundColor DarkGray

# -----------------------------------------------------------------------------
# Send email
# -----------------------------------------------------------------------------
if ($SmtpTo.Count -gt 0 -and $SmtpServer) {
    $subject = if ($anyFailure) {
        "FAILED: SharePoint backup on $hostName - $dateStr ($failCount of $($results.Count) jobs failed)"
    } else {
        "SUCCESS: SharePoint backup on $hostName - $dateStr ($successCount jobs completed)"
    }

    try {
        # Build the email using System.Net.Mail for full STARTTLS + credential control
        $smtpClient = [System.Net.Mail.SmtpClient]::new($SmtpServer, $SmtpPort)
        $smtpClient.EnableSsl = $SmtpUseSsl

        if ($SmtpUser -and $SmtpPass) {
            $smtpClient.Credentials = [System.Net.NetworkCredential]::new($SmtpUser, $SmtpPass)
        }

        $message = [System.Net.Mail.MailMessage]::new()
        $message.From = [System.Net.Mail.MailAddress]::new($SmtpFrom)
        foreach ($to in $SmtpTo) {
            $message.To.Add($to)
        }
        $message.Subject = $subject
        $message.Body    = $reportText

        # Attach the report file
        $attachment = $null
        if (Test-Path $logFile) {
            $attachment = [System.Net.Mail.Attachment]::new($logFile)
            $message.Attachments.Add($attachment)
        }

        $smtpClient.Send($message)

        # Clean up
        if ($attachment) { $attachment.Dispose() }
        $message.Dispose()
        $smtpClient.Dispose()

        Write-Console "Email sent to: $($SmtpTo -join ', ')" -ForegroundColor Green
    } catch {
        Write-Console "WARNING: Failed to send email: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Console "The backup report is still available at: $logFile" -ForegroundColor Yellow
    }
}

# -----------------------------------------------------------------------------
# Exit
# -----------------------------------------------------------------------------
$exitCode = if ($anyFailure) { 1 } else { 0 }
Write-Console ''
Write-Console "Orchestrator finished ($overallStatus, exit code $exitCode)" -ForegroundColor $(if ($anyFailure) { 'Yellow' } else { 'Green' })
exit $exitCode

} catch {
    # -- Crash handler - the orchestrator itself hit an unhandled error -----
    $crashMsg  = $_.Exception.Message
    $crashLine = $_.InvocationInfo.PositionMessage
    Write-Console "ORCHESTRATOR ERROR: $crashMsg" -ForegroundColor Red
    Write-Console $crashLine -ForegroundColor Red

    # Try to send a crash-notification email
    if ($SmtpTo.Count -gt 0 -and $SmtpServer) {
        try {
            $crashSubject = "CRASH: SharePoint backup orchestrator on $($env:COMPUTERNAME ?? (hostname)) - $((Get-Date).ToString('dd-MM-yyyy'))"
            $crashBody    = @"
The backup orchestrator script crashed with an unhandled error.

Error:     $crashMsg
Location:  $crashLine
Time:      $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')
Host:      $($env:COMPUTERNAME ?? (hostname))

Jobs completed before crash: $($results.Count) of $($Jobs.Count)

--- Completed job results ---
$(
    ($results | ForEach-Object {
        "$($_.Status) | $($_.Label) | exit $($_.ExitCode) | $($_.Duration)"
    }) -join "`r`n"
)

--- Full error ---
$($_ | Out-String)
"@

            $client = [System.Net.Mail.SmtpClient]::new($SmtpServer, $SmtpPort)
            $client.EnableSsl = $SmtpUseSsl
            if ($SmtpUser -and $SmtpPass) {
                $client.Credentials = [System.Net.NetworkCredential]::new($SmtpUser, $SmtpPass)
            }
            $msg = [System.Net.Mail.MailMessage]::new()
            $msg.From = [System.Net.Mail.MailAddress]::new($SmtpFrom)
            foreach ($to in $SmtpTo) { $msg.To.Add($to) }
            $msg.Subject = $crashSubject
            $msg.Body    = $crashBody
            $client.Send($msg)
            $msg.Dispose()
            $client.Dispose()
            Write-Console "Crash notification email sent to: $($SmtpTo -join ', ')" -ForegroundColor Yellow
        } catch {
            Write-Console "WARNING: Failed to send crash email: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    exit 1
}
