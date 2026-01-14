<#
Compliance Search + optional SoftDelete purge (updated for EXO module v3.9.0+ SearchOnly session)

Permissions note:
Navigate to https://compliance.microsoft.com/ > Permissions > Microsoft Purview Solutions
Select "Roles" > select "eDiscovery Manager" > Add User

Source reference:
https://learn.microsoft.com/en-us/exchange/policy-and-compliance/ediscovery/compliance-search
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()][string]$Name,
    [Parameter()][string]$FromEmail,
    [Parameter()][string]$Subject,
    [Parameter()][datetime]$StartDate,
    [Parameter()][datetime]$EndDate,
    [Parameter()][string]$ExchangeLocation = "All",
    [Parameter()][int]$TimeoutSeconds = 900,
    [Parameter()][string]$LogPath,
    [Parameter()][switch]$UseCalendar
)

$script:TranscriptStarted = $false

function Start-OptionalTranscript {
    param([Parameter()][string]$Path)

    if (-not $Path) {
        return
    }

    try {
        Start-Transcript -Path $Path -Append | Out-Null
        $script:TranscriptStarted = $true
    } catch {
        Write-Warning "Unable to start transcript at '$Path'. $($_.Exception.Message)"
    }
}

function Stop-OptionalTranscript {
    if ($script:TranscriptStarted) {
        try {
            Stop-Transcript | Out-Null
        } catch {
            Write-Warning "Unable to stop transcript. $($_.Exception.Message)"
        }
    }
}

function Read-NonEmptyInput {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter()][switch]$AllowEmpty
    )

    while ($true) {
        $value = (Read-Host $Prompt).Trim()
        if ($AllowEmpty -or $value) {
            return $value
        }
        Write-Warning "Value cannot be empty."
    }
}

function Get-OneDateViaCalendar {
    param([Parameter(Mandatory)][string]$Title)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object Drawing.Size @(243,230)
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true

    $calendar = New-Object System.Windows.Forms.MonthCalendar
    $calendar.ShowTodayCircle = $false
    $calendar.MaxSelectionCount = 1
    $form.Controls.Add($calendar)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Location = New-Object System.Drawing.Point(40,165)
    $ok.Size = New-Object System.Drawing.Size(150,23)
    $ok.Text = 'OK'
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Date selection cancelled."
    }

    return $calendar.SelectionStart.Date
}

function Get-DateInput {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter()][switch]$UseCalendar
    )

    if ($UseCalendar) {
        try {
            return Get-OneDateViaCalendar -Title $Title
        } catch {
            Write-Warning "Calendar picker unavailable. Falling back to manual entry."
        }
    }

    while ($true) {
        $parsed = $null
        $value = Read-NonEmptyInput -Prompt "$Title (MM/dd/yyyy)"
        if ([datetime]::TryParse($value, [ref]$parsed)) {
            return $parsed.Date
        }
        Write-Warning "Invalid date format."
    }
}

function Escape-KqlValue {
    param([Parameter(Mandatory)][string]$Value)

    $escaped = $Value -replace '\\', '\\\\'
    $escaped = $escaped -replace '"', '\\"'
    return $escaped
}

function Confirm-AllMailboxes {
    param([Parameter(Mandatory)][string]$ExchangeLocation)

    if ($ExchangeLocation -ne 'All') {
        return
    }

    $confirmation = Read-NonEmptyInput -Prompt "ExchangeLocation is 'All'. Type 'ALL' to confirm"
    if ($confirmation -ne 'ALL') {
        throw "Search cancelled."
    }
}

function Confirm-WhatIf {
    if ($WhatIfPreference) {
        Write-Warning "-WhatIf is enabled; no changes will be performed."
    }
}

Start-OptionalTranscript -Path $LogPath

try {
    Confirm-WhatIf

    # Load module + connect (SearchOnly session is required for Compliance Search in newer EXO versions)
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
    Connect-IPPSSession -EnableSearchOnlySession | Out-Null

    # Prompts
    if (-not $Name) {
        $Name = Read-NonEmptyInput -Prompt "Compliance Search Name"
    }
    if (-not $FromEmail) {
        $FromEmail = Read-NonEmptyInput -Prompt "Enter email address this came from (wildcard: *domain.com also works)"
    }
    if (-not $Subject) {
        $Subject = Read-NonEmptyInput -Prompt "Enter the subject line of the email (wildcard: Reset Your password for* also works)"
    }

    if ($FromEmail -match '[*?]' -or $Subject -match '[*?]') {
        Write-Warning "Wildcard characters detected. Ensure the scope is as intended."
    }

    if (-not $StartDate) {
        $StartDate = Get-DateInput -Title "Select START Date" -UseCalendar:$UseCalendar
    }
    if (-not $EndDate) {
        $EndDate = Get-DateInput -Title "Select END Date" -UseCalendar:$UseCalendar
    }

    if ($EndDate -lt $StartDate) {
        throw "END date cannot be earlier than START date."
    }

    if (-not $FromEmail -or -not $Subject -or -not $Name) {
        throw "Name, FromEmail, and Subject are required."
    }

    Confirm-AllMailboxes -ExchangeLocation $ExchangeLocation

    $Startdate = $StartDate.ToString("MM/dd/yyyy")
    $Enddate = $EndDate.ToString("MM/dd/yyyy")

    $escapedFrom = Escape-KqlValue -Value $FromEmail
    $escapedSubject = Escape-KqlValue -Value $Subject

    # Query
    $query = "(sent>=$Startdate) AND (sent<=$Enddate) AND (From:`"$escapedFrom`") AND (subject:`"$escapedSubject`")"

    Write-Host "Query: $query"

    # Search - Create or Update
    Write-Host "Creating/Updating ComplianceSearch: $Name"
    $existing = $null
    try { $existing = Get-ComplianceSearch -Identity $Name -ErrorAction Stop } catch { $existing = $null }

    if ($existing) {
        if ($PSCmdlet.ShouldProcess($Name, "Update ComplianceSearch query")) {
            Set-ComplianceSearch -Identity $Name -ContentMatchQuery $query | Out-Null
        } else {
            return
        }
    } else {
        if ($PSCmdlet.ShouldProcess($Name, "Create ComplianceSearch")) {
            New-ComplianceSearch -Name $Name -ExchangeLocation $ExchangeLocation -ContentMatchQuery $query | Out-Null
        } else {
            return
        }
    }

    # Search - Start
    if ($PSCmdlet.ShouldProcess($Name, "Start ComplianceSearch")) {
        Write-Host "Starting ComplianceSearch: $Name"
        Start-ComplianceSearch -Identity $Name | Out-Null
    } else {
        return
    }

    # Search - Wait for completion
    Write-Host -NoNewline "Searching"
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        $s = Get-ComplianceSearch -Identity $Name
        if ($s.Status -eq "Completed") { break }
        if ($s.Status -match "Failed|Error") { throw "Search failed. Status: $($s.Status)" }
        if ($stopwatch.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
            throw "Search timed out after $TimeoutSeconds seconds."
        }
        Write-Host -NoNewline "."
        Start-Sleep -Seconds 2
    }
    Write-Host ""
    Write-Host "Search completed!"

    $search = Get-ComplianceSearch -Identity $Name

    # Item Count
    Write-Host "Found '$($search.Items)' items"
    Write-Host ""

    # Mailboxes (parse SuccessResults)
    $mailboxes = @()
    if ($search.SuccessResults) {
        foreach ($line in ($search.SuccessResults -split '[\r\n]+')) {
            if ($line -match 'Location:\s*([^,]+),.*Item count:\s*(\d+)' -and [int]$matches[2] -gt 0) {
                $mailboxes += $matches[1].Trim()
            }
        }
    } else {
        Write-Warning "SuccessResults was empty."
    }

    Write-Host "In mailboxes:"
    if ($mailboxes.Count) { $mailboxes | Sort-Object -Unique } else { Write-Host "(none returned)" }

    # Purge - Confirm/Skip
    $purge = Read-Host "Type the word 'purge' to purge these items (SoftDelete). If you are not purging, press Enter to end."
    if ($purge -eq "purge") {
        if ($PSCmdlet.ShouldProcess($Name, "Submit SoftDelete purge")) {
            Write-Host "Submitting purge action (SoftDelete)..."
            $action = New-ComplianceSearchAction -SearchName $Name -Purge -PurgeType SoftDelete
            Write-Host "Purge action submitted. Action: $($action.Identity)"
        }
    }
}
finally {
    # Clean disconnect (avoid deprecated manual PSSession cleanup)
    try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}
    try { Disconnect-IPPSSession | Out-Null } catch {}
    Stop-OptionalTranscript
}
