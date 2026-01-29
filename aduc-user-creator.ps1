Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:targetOuPatterns = @(
    "OU=Users,OU=Camden,OU=Mazza Demo",
    "OU=Users,OU=Tinton Falls,OU=Mazza Demo",
    "OU=Users,OU=Philadelphia,OU=Mazza Demo",
    "OU=Users,OU=Campus Parkway,OU=Mazza Demo"
)

$script:ouDisplayMap = @{}
$script:terminateUserMap = @{}
$script:reenableUserMap = @{}
$script:lastGeneratedCredentials = $null
$script:disabledOuDisplayMap = @{}
$script:terminateUserDisplayNames = @()
$script:reenableUserDisplayNames = @()
$script:resetUserMap = @{}
$script:resetUserDisplayNames = @()

function Get-SamAccountName {
    param(
        [string]$FirstName,
        [string]$LastName
    )

    $firstInitial = $FirstName.Trim().Substring(0, 1)
    $last = $LastName.Trim()
    $sam = ($firstInitial + $last).ToLowerInvariant()
    return ($sam -replace "[^a-z0-9]", "")
}

function Get-ProxyAddresses {
    param(
        [string]$MailNickname,
        [string]$Domain
    )

    $primary = "SMTP:$MailNickname@$Domain"
    return @($primary)
}

function Get-AvailableDomains {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        return $null
    }

    try {
        $forest = Get-ADForest
        $domains = @()
        if ($forest.Domains) {
            $domains += $forest.Domains
        }
        if ($forest.UPNSuffixes) {
            $domains += $forest.UPNSuffixes
        }
        return $domains | Sort-Object -Unique
    }
    catch {
        return $null
    }
}

function Update-Status {
    param(
        [System.Windows.Forms.Label]$Label,
        [string]$Message,
        [System.Drawing.Color]$Color
    )

    $Label.Text = $Message
    $Label.ForeColor = $Color
}

function Invoke-DirectorySync {
    param(
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.TextBox]$LogTextBox
    )

    try {
        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
        if ($StatusLabel) {
            Update-Status -Label $StatusLabel -Message "Sync started." -Color ([System.Drawing.Color]::DarkGreen)
        }
        if ($script:footerStatusLabel) {
            $script:footerStatusLabel.Text = "Sync started at $([DateTime]::Now.ToString('g'))"
        }
        if ($LogTextBox) {
            $LogTextBox.Text = "Sync started at $([DateTime]::Now.ToString('g'))"
        }
    }
    catch {
        if ($StatusLabel) {
            Update-Status -Label $StatusLabel -Message "Sync failed: $($_.Exception.Message)" -Color ([System.Drawing.Color]::DarkRed)
        }
        if ($script:footerStatusLabel) {
            $script:footerStatusLabel.Text = "Sync failed: $($_.Exception.Message)"
        }
        if ($LogTextBox) {
            $LogTextBox.Text = "Sync failed: $($_.Exception.Message)"
        }
    }
}

function Get-DisableOuFromUserOu {
    param(
        [string]$UserOuDn
    )

    if ($UserOuDn -match "OU=Users,") {
        return $UserOuDn -replace "OU=Users,", "OU=Disabled Users,"
    }

    return $null
}

function Get-UsersOuFromDisabledOu {
    param(
        [string]$DisabledDn
    )

    if ($DisabledDn -match "OU=Disabled Users,") {
        return $DisabledDn -replace "OU=Disabled Users,", "OU=Users,"
    }

    return $null
}

function Load-DisabledOUs {
    param(
        [System.Windows.Forms.ComboBox]$ComboBox,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $previousSelection = $ComboBox.Text
    $ComboBox.Items.Clear()
    $script:disabledOuDisplayMap = @{}

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $StatusLabel -Message "ActiveDirectory module not found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $ous = Get-ADOrganizationalUnit -Filter * | Sort-Object Name
    }
    catch {
        Update-Status -Label $StatusLabel -Message "Unable to query OUs. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $matchedItems = @()
    foreach ($ou in $ous) {
        $dn = $ou.DistinguishedName
        $matchesPattern = $false
        $matchedPattern = $null

        foreach ($pattern in $script:targetOuPatterns) {
            $disabledPattern = $pattern -replace "^OU=Users,", "OU=Disabled Users,"
            if ($dn -like "*$disabledPattern*") {
                $matchesPattern = $true
                $matchedPattern = $disabledPattern
                break
            }
        }

        if ($matchesPattern) {
            $displayName = $matchedPattern -replace "^OU=Disabled Users,OU=", ""
            $displayName = $displayName -replace ",OU=Mazza Demo", ""
            $displayName = $displayName -replace "OU=", ""
            $displayName = $displayName -replace ",", " / "

            $matchedItems += [PSCustomObject]@{
                DisplayName = $displayName
                Dn          = $dn
            }
        }
    }

    foreach ($item in $matchedItems | Sort-Object DisplayName) {
        $script:disabledOuDisplayMap[$item.DisplayName] = $item.Dn
        [void]$ComboBox.Items.Add($item.DisplayName)
    }

    if ($ComboBox.Items.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($previousSelection) -and $script:disabledOuDisplayMap.ContainsKey($previousSelection)) {
            $ComboBox.SelectedItem = $previousSelection
        }
        else {
            $ComboBox.SelectedIndex = 0
        }
        Update-Status -Label $StatusLabel -Message "Loaded disabled OUs from Active Directory." -Color ([System.Drawing.Color]::DarkGreen)
    }
    else {
        Update-Status -Label $StatusLabel -Message "No disabled OUs found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
    }
}

function Resolve-DisabledOuDn {
    param(
        [string]$Selection
    )

    if ($script:disabledOuDisplayMap.ContainsKey($Selection)) {
        return $script:disabledOuDisplayMap[$Selection]
    }

    return $Selection
}

function Show-CredentialsDialog {
    param(
        [string]$Title,
        [string]$Message
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.Size = New-Object System.Drawing.Size(520, 360)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $messageBox = New-Object System.Windows.Forms.TextBox
    $messageBox.Location = New-Object System.Drawing.Point(20, 20)
    $messageBox.Size = New-Object System.Drawing.Size(460, 230)
    $messageBox.Multiline = $true
    $messageBox.ReadOnly = $true
    $messageBox.ScrollBars = "Vertical"
    $messageBox.Text = $Message

    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy"
    $copyButton.Location = New-Object System.Drawing.Point(280, 265)
    $copyButton.Size = New-Object System.Drawing.Size(90, 30)
    $copyButton.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($Message)
    })

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(390, 265)
    $closeButton.Size = New-Object System.Drawing.Size(90, 30)
    $closeButton.Add_Click({ $dialog.Close() })

    $dialog.Controls.AddRange(@($messageBox, $copyButton, $closeButton))
    [void]$dialog.ShowDialog()
}

function Confirm-Action {
    param(
        [string]$Message,
        [string]$Title
    )

    $result = [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

function Show-AboutDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "About"
    $dialog.Size = New-Object System.Drawing.Size(360, 220)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "JUMP 1.0`r`nLast Update 1/29/2026`r`n`r`nJUMP is a PowerShell Script tool.`r`n`r`nAuthor: Roger Anderson`r`nEmail: randerson@jaydien.com"
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(300, 120)
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(240, 150)
    $closeButton.Size = New-Object System.Drawing.Size(80, 28)
    $closeButton.Add_Click({ $dialog.Close() })

    $dialog.Controls.AddRange(@($label, $closeButton))
    [void]$dialog.ShowDialog()
}

function Get-LocalIconImage {
    param(
        [string]$Path,
        [int]$Width,
        [int]$Height
    )

    try {
        if (-not (Test-Path -Path $Path)) {
            return $null
        }
        $icon = New-Object System.Drawing.Icon($Path)
        $bitmap = $icon.ToBitmap()
        $scaled = New-Object System.Drawing.Bitmap($bitmap, $Width, $Height)
        $icon.Dispose()
        $bitmap.Dispose()
        return $scaled
    }
    catch {
        return $null
    }
}

function Set-ButtonIcon {
    param(
        [System.Windows.Forms.Button]$Button,
        [string]$Path
    )

    if (-not $Button) {
        return
    }

    $targetWidth = [Math]::Max(1, $Button.ClientSize.Width)
    $targetHeight = [Math]::Max(1, $Button.ClientSize.Height)
    $targetSize = [Math]::Max(1, [Math]::Floor([Math]::Min($targetWidth, $targetHeight) * 0.6))
    $image = Get-LocalIconImage -Path $Path -Width $targetSize -Height $targetSize
    if ($image) {
        $Button.Image = $image
    }
}

function Register-ButtonIcon {
    param(
        [System.Windows.Forms.Button]$Button,
        [string]$Path
    )

    if (-not $Button) {
        return
    }

    $Button.Tag = $Path
    Set-ButtonIcon -Button $Button -Path $Path
    $Button.Add_Resize({
        $iconPath = $this.Tag
        if ($iconPath) {
            Set-ButtonIcon -Button $this -Path $iconPath
        }
    })
}

function Register-HoverHighlight {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$HoverColor
    )

    if (-not $Button) {
        return
    }

    if ($Button.Tag -isnot [hashtable]) {
        $Button.Tag = @{ IconPath = $Button.Tag }
    }

    $defaultColor = $Button.BackColor
    if ($defaultColor -eq $null) {
        $defaultColor = [System.Drawing.SystemColors]::Control
    }
    $Button.Tag.DefaultBackColor = $defaultColor

    $hover = $HoverColor
    if ($hover -eq $null) {
        $hover = [System.Drawing.Color]::LightSteelBlue
    }

    $Button.Add_MouseEnter({
        $this.BackColor = $this.Tag.HoverColor
    })

    $Button.Add_MouseLeave({
        $this.BackColor = $this.Tag.DefaultBackColor
    })

    $Button.Tag.HoverColor = $hover
}

function New-RandomPassword {
    param(
        [int]$Length = 16
    )

    $chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%&*"
    $random = New-Object System.Random
    $passwordChars = for ($i = 0; $i -lt $Length; $i++) {
        $chars[$random.Next(0, $chars.Length)]
    }
    return -join $passwordChars
}

function Load-OUs {
    param(
        [System.Windows.Forms.ComboBox]$ComboBox,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $previousSelection = $ComboBox.Text
    $ComboBox.Items.Clear()
    $script:ouDisplayMap = @{}

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $StatusLabel -Message "ActiveDirectory module not found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $ous = Get-ADOrganizationalUnit -Filter * | Sort-Object Name
    }
    catch {
        Update-Status -Label $StatusLabel -Message "Unable to query OUs. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    foreach ($ou in $ous) {
        $dn = $ou.DistinguishedName
        $matchesPattern = $false
        $matchedPattern = $null

        foreach ($pattern in $script:targetOuPatterns) {
            if ($dn -like "*$pattern*") {
                $matchesPattern = $true
                $matchedPattern = $pattern
                break
            }
        }

        if ($matchesPattern) {
            $displayName = $matchedPattern -replace "^OU=Users,OU=", ""
            $displayName = $displayName -replace ",OU=Mazza Demo", ""
            $displayName = $displayName -replace "OU=", ""
            $displayName = $displayName -replace ",", " / "

            $script:ouDisplayMap[$displayName] = $dn
            [void]$ComboBox.Items.Add($displayName)
        }
    }

    if ($ComboBox.Items.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($previousSelection) -and $script:ouDisplayMap.ContainsKey($previousSelection)) {
            $ComboBox.SelectedItem = $previousSelection
        }
        else {
            $ComboBox.SelectedIndex = 0
        }
        Update-Status -Label $StatusLabel -Message "Loaded target OUs from Active Directory." -Color ([System.Drawing.Color]::DarkGreen)
    }
    else {
        Update-Status -Label $StatusLabel -Message "No OUs found. Enter OU DN manually." -Color ([System.Drawing.Color]::DarkRed)
    }
}

function Resolve-OuDn {
    param(
        [string]$Selection
    )

    if ($script:ouDisplayMap.ContainsKey($Selection)) {
        return $script:ouDisplayMap[$Selection]
    }

    return $Selection
}

function Get-ThumbnailImage {
    param(
        [string]$Url,
        [int]$Width,
        [int]$Height
    )

    try {
        $webClient = New-Object System.Net.WebClient
        $bytes = $webClient.DownloadData($Url)
        $stream = New-Object System.IO.MemoryStream(,$bytes)
        $image = [System.Drawing.Image]::FromStream($stream)
        $scaled = New-Object System.Drawing.Bitmap($image, $Width, $Height)
        $stream.Dispose()
        $image.Dispose()
        return $scaled
    }
    catch {
        return $null
    }
}

function Show-Panel {
    param(
        [System.Windows.Forms.Panel]$PanelToShow,
        [System.Windows.Forms.Panel]$PanelToHide1,
        [System.Windows.Forms.Panel]$PanelToHide2
    )

    $PanelToShow.Visible = $true
    $PanelToShow.BringToFront()
    if ($PanelToHide1) {
        $PanelToHide1.Visible = $false
    }
    if ($PanelToHide2) {
        $PanelToHide2.Visible = $false
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Jaydien Unified Management Platform"
$form.Size = New-Object System.Drawing.Size(960, 560)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.MinimumSize = New-Object System.Drawing.Size(800, 500)

$font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)

$footerStatusLabel = New-Object System.Windows.Forms.Label
$footerStatusLabel.Text = "Jaydien Network Solutions"
$footerStatusLabel.Dock = "Bottom"
$footerStatusLabel.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)
$footerStatusLabel.Height = 26
$footerStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Regular)
$footerStatusLabel.ForeColor = [System.Drawing.Color]::DimGray
$footerStatusLabel.BackColor = [System.Drawing.Color]::WhiteSmoke
$script:footerStatusLabel = $footerStatusLabel

$mainMenuPanel = New-Object System.Windows.Forms.Panel
$mainMenuPanel.Dock = "Fill"

$createPanel = New-Object System.Windows.Forms.Panel
$createPanel.Dock = "Fill"
$createPanel.Visible = $false

$terminatePanel = New-Object System.Windows.Forms.Panel
$terminatePanel.Dock = "Fill"
$terminatePanel.Visible = $false

$reenablePanel = New-Object System.Windows.Forms.Panel
$reenablePanel.Dock = "Fill"
$reenablePanel.Visible = $false

$resetPanel = New-Object System.Windows.Forms.Panel
$resetPanel.Dock = "Fill"
$resetPanel.Visible = $false

$menuTitleLabel = New-Object System.Windows.Forms.Label
$menuTitleLabel.Text = "Select an action"
$menuTitleLabel.Dock = "Top"
$menuTitleLabel.Padding = New-Object System.Windows.Forms.Padding(20, 20, 0, 10)
$menuTitleLabel.AutoSize = $true
$menuTitleLabel.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 14, [System.Drawing.FontStyle]::Regular)

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock = "Top"
$helpMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Help")
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("About")
$aboutMenuItem.Add_Click({
    Show-AboutDialog
})
[void]$helpMenuItem.DropDownItems.Add($aboutMenuItem)
[void]$menuStrip.Items.Add($helpMenuItem)

$menuSyncButton = New-Object System.Windows.Forms.Button
$menuSyncButton.Text = "SYNC NOW"
$menuSyncButton.Location = New-Object System.Drawing.Point(830, 20)
$menuSyncButton.Size = New-Object System.Drawing.Size(110, 30)
$menuSyncButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Regular)
$menuSyncButton.Anchor = "Top,Right"
$menuSyncButton.Add_Click({
    Invoke-DirectorySync
})

$menuGrid = New-Object System.Windows.Forms.TableLayoutPanel
$menuGrid.ColumnCount = 4
$menuGrid.RowCount = 2
$menuGrid.Dock = "Fill"
$menuGrid.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$menuGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$menuGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$menuGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$menuGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$menuGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$menuGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$menuGrid.BackColor = [System.Drawing.Color]::Transparent
$menuGrid.Parent = $mainMenuPanel

$createTileButton = New-Object System.Windows.Forms.Button
$createTileButton.Text = "Create User"
$createTileButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Regular)
$createTileButton.TextImageRelation = "ImageAboveText"
$createTileButton.ImageAlign = "MiddleCenter"
$createTileButton.TextAlign = "BottomCenter"
$createTileButton.Dock = "Fill"
$createTileButton.Margin = New-Object System.Windows.Forms.Padding(4)

$terminateTileButton = New-Object System.Windows.Forms.Button
$terminateTileButton.Text = "Terminate User"
$terminateTileButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Regular)
$terminateTileButton.TextImageRelation = "ImageAboveText"
$terminateTileButton.ImageAlign = "MiddleCenter"
$terminateTileButton.TextAlign = "BottomCenter"
$terminateTileButton.Dock = "Fill"
$terminateTileButton.Margin = New-Object System.Windows.Forms.Padding(4)

$reenableTileButton = New-Object System.Windows.Forms.Button
$reenableTileButton.Text = "Re-enable User"
$reenableTileButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Regular)
$reenableTileButton.TextImageRelation = "ImageAboveText"
$reenableTileButton.ImageAlign = "MiddleCenter"
$reenableTileButton.TextAlign = "BottomCenter"
$reenableTileButton.Dock = "Fill"
$reenableTileButton.Margin = New-Object System.Windows.Forms.Padding(4)

$resetTileButton = New-Object System.Windows.Forms.Button
$resetTileButton.Text = "Reset Password"
$resetTileButton.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 10, [System.Drawing.FontStyle]::Regular)
$resetTileButton.TextImageRelation = "ImageAboveText"
$resetTileButton.ImageAlign = "MiddleCenter"
$resetTileButton.TextAlign = "BottomCenter"
$resetTileButton.Dock = "Fill"
$resetTileButton.Margin = New-Object System.Windows.Forms.Padding(4)

$dummyTileButton1 = New-Object System.Windows.Forms.Button
$dummyTileButton1.Text = "Under Development"
$dummyTileButton1.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Regular)
$dummyTileButton1.Enabled = $false
$dummyTileButton1.BackColor = [System.Drawing.Color]::LightGray
$dummyTileButton1.TextImageRelation = "ImageAboveText"
$dummyTileButton1.ImageAlign = "MiddleCenter"
$dummyTileButton1.TextAlign = "BottomCenter"
$dummyTileButton1.Dock = "Fill"
$dummyTileButton1.Margin = New-Object System.Windows.Forms.Padding(4)

$dummyTileButton2 = New-Object System.Windows.Forms.Button
$dummyTileButton2.Text = "Under Development"
$dummyTileButton2.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Regular)
$dummyTileButton2.Enabled = $false
$dummyTileButton2.BackColor = [System.Drawing.Color]::LightGray
$dummyTileButton2.TextImageRelation = "ImageAboveText"
$dummyTileButton2.ImageAlign = "MiddleCenter"
$dummyTileButton2.TextAlign = "BottomCenter"
$dummyTileButton2.Dock = "Fill"
$dummyTileButton2.Margin = New-Object System.Windows.Forms.Padding(4)

$dummyTileButton3 = New-Object System.Windows.Forms.Button
$dummyTileButton3.Text = "Under Development"
$dummyTileButton3.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Regular)
$dummyTileButton3.Enabled = $false
$dummyTileButton3.BackColor = [System.Drawing.Color]::LightGray
$dummyTileButton3.TextImageRelation = "ImageAboveText"
$dummyTileButton3.ImageAlign = "MiddleCenter"
$dummyTileButton3.TextAlign = "BottomCenter"
$dummyTileButton3.Dock = "Fill"
$dummyTileButton3.Margin = New-Object System.Windows.Forms.Padding(4)

$dummyTileButton4 = New-Object System.Windows.Forms.Button
$dummyTileButton4.Text = "Under Development"
$dummyTileButton4.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9, [System.Drawing.FontStyle]::Regular)
$dummyTileButton4.Enabled = $false
$dummyTileButton4.BackColor = [System.Drawing.Color]::LightGray
$dummyTileButton4.TextImageRelation = "ImageAboveText"
$dummyTileButton4.ImageAlign = "MiddleCenter"
$dummyTileButton4.TextAlign = "BottomCenter"
$dummyTileButton4.Dock = "Fill"
$dummyTileButton4.Margin = New-Object System.Windows.Forms.Padding(4)

$createTileButton.Add_Click({
    Show-Panel -PanelToShow $createPanel -PanelToHide1 $mainMenuPanel -PanelToHide2 $terminatePanel
    $reenablePanel.Visible = $false
    $resetPanel.Visible = $false
    Load-OUs -ComboBox $ouComboBox -StatusLabel $createStatusLabel
})

$terminateTileButton.Add_Click({
    Show-Panel -PanelToShow $terminatePanel -PanelToHide1 $mainMenuPanel -PanelToHide2 $createPanel
    $reenablePanel.Visible = $false
    $resetPanel.Visible = $false
    Load-OUs -ComboBox $terminateOuComboBox -StatusLabel $terminateStatusLabel
})

$reenableTileButton.Add_Click({
    Show-Panel -PanelToShow $reenablePanel -PanelToHide1 $mainMenuPanel -PanelToHide2 $terminatePanel
    $createPanel.Visible = $false
    $resetPanel.Visible = $false
    Load-DisabledOUs -ComboBox $reenableOuComboBox -StatusLabel $reenableStatusLabel
})

$resetTileButton.Add_Click({
    Show-Panel -PanelToShow $resetPanel -PanelToHide1 $mainMenuPanel -PanelToHide2 $reenablePanel
    $createPanel.Visible = $false
    $terminatePanel.Visible = $false
    Load-OUs -ComboBox $resetOuComboBox -StatusLabel $resetStatusLabel
})

$backToMenuFromCreate = New-Object System.Windows.Forms.Button
$backToMenuFromCreate.Text = "Back"
$backToMenuFromCreate.Location = New-Object System.Drawing.Point(20, 20)
$backToMenuFromCreate.Size = New-Object System.Drawing.Size(80, 28)
$backToMenuFromCreate.Font = $font
$backToMenuFromCreate.Add_Click({
    Show-Panel -PanelToShow $mainMenuPanel -PanelToHide1 $createPanel -PanelToHide2 $terminatePanel
})

$backToMenuFromTerminate = New-Object System.Windows.Forms.Button
$backToMenuFromTerminate.Text = "Back"
$backToMenuFromTerminate.Location = New-Object System.Drawing.Point(20, 20)
$backToMenuFromTerminate.Size = New-Object System.Drawing.Size(80, 28)
$backToMenuFromTerminate.Font = $font
$backToMenuFromTerminate.Add_Click({
    Show-Panel -PanelToShow $mainMenuPanel -PanelToHide1 $createPanel -PanelToHide2 $terminatePanel
})

$backToMenuFromReenable = New-Object System.Windows.Forms.Button
$backToMenuFromReenable.Text = "Back"
$backToMenuFromReenable.Location = New-Object System.Drawing.Point(20, 20)
$backToMenuFromReenable.Size = New-Object System.Drawing.Size(80, 28)
$backToMenuFromReenable.Font = $font
$backToMenuFromReenable.Add_Click({
    Show-Panel -PanelToShow $mainMenuPanel -PanelToHide1 $reenablePanel -PanelToHide2 $terminatePanel
})

$backToMenuFromReset = New-Object System.Windows.Forms.Button
$backToMenuFromReset.Text = "Back"
$backToMenuFromReset.Location = New-Object System.Drawing.Point(20, 20)
$backToMenuFromReset.Size = New-Object System.Drawing.Size(80, 28)
$backToMenuFromReset.Font = $font
$backToMenuFromReset.Add_Click({
    Show-Panel -PanelToShow $mainMenuPanel -PanelToHide1 $resetPanel -PanelToHide2 $reenablePanel
})

$firstNameLabel = New-Object System.Windows.Forms.Label
$firstNameLabel.Text = "First Name"
$firstNameLabel.Location = New-Object System.Drawing.Point(20, 60)
$firstNameLabel.Size = New-Object System.Drawing.Size(120, 24)
$firstNameLabel.Font = $font

$firstNameTextBox = New-Object System.Windows.Forms.TextBox
$firstNameTextBox.Location = New-Object System.Drawing.Point(180, 58)
$firstNameTextBox.Size = New-Object System.Drawing.Size(460, 24)
$firstNameTextBox.Font = $font

$lastNameLabel = New-Object System.Windows.Forms.Label
$lastNameLabel.Text = "Last Name"
$lastNameLabel.Location = New-Object System.Drawing.Point(20, 100)
$lastNameLabel.Size = New-Object System.Drawing.Size(120, 24)
$lastNameLabel.Font = $font

$lastNameTextBox = New-Object System.Windows.Forms.TextBox
$lastNameTextBox.Location = New-Object System.Drawing.Point(180, 98)
$lastNameTextBox.Size = New-Object System.Drawing.Size(460, 24)
$lastNameTextBox.Font = $font

$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Text = "Email/UPN Domain"
$domainLabel.Location = New-Object System.Drawing.Point(20, 140)
$domainLabel.Size = New-Object System.Drawing.Size(150, 24)
$domainLabel.Font = $font

$domainComboBox = New-Object System.Windows.Forms.ComboBox
$domainComboBox.Location = New-Object System.Drawing.Point(180, 138)
$domainComboBox.Size = New-Object System.Drawing.Size(460, 24)
$domainComboBox.Font = $font
$domainComboBox.DropDownStyle = "DropDownList"

$ouLabel = New-Object System.Windows.Forms.Label
$ouLabel.Text = "Target OU"
$ouLabel.Location = New-Object System.Drawing.Point(20, 180)
$ouLabel.Size = New-Object System.Drawing.Size(120, 24)
$ouLabel.Font = $font

$ouComboBox = New-Object System.Windows.Forms.ComboBox
$ouComboBox.Location = New-Object System.Drawing.Point(180, 178)
$ouComboBox.Size = New-Object System.Drawing.Size(360, 24)
$ouComboBox.Font = $font
$ouComboBox.DropDownStyle = "DropDown"

$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Text = "Password (optional)"
$passwordLabel.Location = New-Object System.Drawing.Point(20, 220)
$passwordLabel.Size = New-Object System.Drawing.Size(150, 24)
$passwordLabel.Font = $font

$passwordTextBox = New-Object System.Windows.Forms.TextBox
$passwordTextBox.Location = New-Object System.Drawing.Point(180, 218)
$passwordTextBox.Size = New-Object System.Drawing.Size(460, 24)
$passwordTextBox.Font = $font
$passwordTextBox.UseSystemPasswordChar = $true

$mustChangePasswordCheckBox = New-Object System.Windows.Forms.CheckBox
$mustChangePasswordCheckBox.Text = "Require password reset at next logon"
$mustChangePasswordCheckBox.Location = New-Object System.Drawing.Point(180, 248)
$mustChangePasswordCheckBox.Size = New-Object System.Drawing.Size(320, 24)
$mustChangePasswordCheckBox.Font = $font
$mustChangePasswordCheckBox.Checked = $true

$copyCredentialsButton = New-Object System.Windows.Forms.Button
$copyCredentialsButton.Text = "Copy Credentials"
$copyCredentialsButton.Location = New-Object System.Drawing.Point(510, 248)
$copyCredentialsButton.Size = New-Object System.Drawing.Size(130, 28)
$copyCredentialsButton.Font = $font
$copyCredentialsButton.Enabled = $false

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Text = "Logs"
$previewLabel.Location = New-Object System.Drawing.Point(20, 290)
$previewLabel.Size = New-Object System.Drawing.Size(120, 24)
$previewLabel.Font = $font

$previewTextBox = New-Object System.Windows.Forms.TextBox
$previewTextBox.Location = New-Object System.Drawing.Point(180, 288)
$previewTextBox.Size = New-Object System.Drawing.Size(460, 90)
$previewTextBox.Multiline = $true
$previewTextBox.ReadOnly = $true
$previewTextBox.Font = $font

$createButton = New-Object System.Windows.Forms.Button
$createButton.Text = "Create User"
$createButton.Location = New-Object System.Drawing.Point(180, 390)
$createButton.Size = New-Object System.Drawing.Size(140, 36)
$createButton.Font = $font
$createButton.Enabled = $false

$createSyncButton = New-Object System.Windows.Forms.Button
$createSyncButton.Text = "SYNC NOW"
$createSyncButton.Location = New-Object System.Drawing.Point(340, 390)
$createSyncButton.Size = New-Object System.Drawing.Size(120, 36)
$createSyncButton.Font = $font
$createSyncButton.Add_Click({
    Invoke-DirectorySync -StatusLabel $createStatusLabel -LogTextBox $previewTextBox
})

$createStatusLabel = New-Object System.Windows.Forms.Label
$createStatusLabel.Text = "Ready"
$createStatusLabel.Location = New-Object System.Drawing.Point(20, 440)
$createStatusLabel.Size = New-Object System.Drawing.Size(620, 24)
$createStatusLabel.Font = $font
$createStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

$terminateOuLabel = New-Object System.Windows.Forms.Label
$terminateOuLabel.Text = "Source Users OU"
$terminateOuLabel.Location = New-Object System.Drawing.Point(20, 60)
$terminateOuLabel.Size = New-Object System.Drawing.Size(150, 24)
$terminateOuLabel.Font = $font

$terminateOuComboBox = New-Object System.Windows.Forms.ComboBox
$terminateOuComboBox.Location = New-Object System.Drawing.Point(180, 58)
$terminateOuComboBox.Size = New-Object System.Drawing.Size(360, 24)
$terminateOuComboBox.Font = $font
$terminateOuComboBox.DropDownStyle = "DropDown"

$refreshTerminateOuButton = New-Object System.Windows.Forms.Button
$refreshTerminateOuButton.Text = "Refresh OUs"
$refreshTerminateOuButton.Location = New-Object System.Drawing.Point(550, 56)
$refreshTerminateOuButton.Size = New-Object System.Drawing.Size(90, 28)
$refreshTerminateOuButton.Font = $font

$terminateSearchLabel = New-Object System.Windows.Forms.Label
$terminateSearchLabel.Text = "Search"
$terminateSearchLabel.Location = New-Object System.Drawing.Point(20, 100)
$terminateSearchLabel.Size = New-Object System.Drawing.Size(120, 24)
$terminateSearchLabel.Font = $font

$terminateSearchTextBox = New-Object System.Windows.Forms.TextBox
$terminateSearchTextBox.Location = New-Object System.Drawing.Point(180, 98)
$terminateSearchTextBox.Size = New-Object System.Drawing.Size(360, 24)
$terminateSearchTextBox.Font = $font

$userListBox = New-Object System.Windows.Forms.ListBox
$userListBox.Location = New-Object System.Drawing.Point(20, 140)
$userListBox.Size = New-Object System.Drawing.Size(620, 180)
$userListBox.Font = $font
$userListBox.SelectionMode = "MultiExtended"

$terminateButton = New-Object System.Windows.Forms.Button
$terminateButton.Text = "Terminate Selected"
$terminateButton.Location = New-Object System.Drawing.Point(180, 330)
$terminateButton.Size = New-Object System.Drawing.Size(180, 36)
$terminateButton.Font = $font

$terminateSyncButton = New-Object System.Windows.Forms.Button
$terminateSyncButton.Text = "SYNC NOW"
$terminateSyncButton.Location = New-Object System.Drawing.Point(370, 330)
$terminateSyncButton.Size = New-Object System.Drawing.Size(120, 36)
$terminateSyncButton.Font = $font
$terminateSyncButton.Add_Click({
    Invoke-DirectorySync -StatusLabel $terminateStatusLabel -LogTextBox $terminatePreviewTextBox
})

$terminatePreviewLabel = New-Object System.Windows.Forms.Label
$terminatePreviewLabel.Text = "Logs"
$terminatePreviewLabel.Location = New-Object System.Drawing.Point(20, 380)
$terminatePreviewLabel.Size = New-Object System.Drawing.Size(150, 24)
$terminatePreviewLabel.Font = $font

$terminatePreviewTextBox = New-Object System.Windows.Forms.TextBox
$terminatePreviewTextBox.Location = New-Object System.Drawing.Point(180, 378)
$terminatePreviewTextBox.Size = New-Object System.Drawing.Size(460, 56)
$terminatePreviewTextBox.Multiline = $true
$terminatePreviewTextBox.ReadOnly = $true
$terminatePreviewTextBox.Font = $font

$terminateStatusLabel = New-Object System.Windows.Forms.Label
$terminateStatusLabel.Text = "Ready"
$terminateStatusLabel.Location = New-Object System.Drawing.Point(20, 440)
$terminateStatusLabel.Size = New-Object System.Drawing.Size(620, 24)
$terminateStatusLabel.Font = $font
$terminateStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

$reenableOuLabel = New-Object System.Windows.Forms.Label
$reenableOuLabel.Text = "Disabled Users OU"
$reenableOuLabel.Location = New-Object System.Drawing.Point(20, 60)
$reenableOuLabel.Size = New-Object System.Drawing.Size(150, 24)
$reenableOuLabel.Font = $font

$reenableOuComboBox = New-Object System.Windows.Forms.ComboBox
$reenableOuComboBox.Location = New-Object System.Drawing.Point(180, 58)
$reenableOuComboBox.Size = New-Object System.Drawing.Size(360, 24)
$reenableOuComboBox.Font = $font
$reenableOuComboBox.DropDownStyle = "DropDown"

$refreshReenableOuButton = New-Object System.Windows.Forms.Button
$refreshReenableOuButton.Text = "Refresh OUs"
$refreshReenableOuButton.Location = New-Object System.Drawing.Point(550, 56)
$refreshReenableOuButton.Size = New-Object System.Drawing.Size(90, 28)
$refreshReenableOuButton.Font = $font

$reenableSearchLabel = New-Object System.Windows.Forms.Label
$reenableSearchLabel.Text = "Search"
$reenableSearchLabel.Location = New-Object System.Drawing.Point(20, 100)
$reenableSearchLabel.Size = New-Object System.Drawing.Size(120, 24)
$reenableSearchLabel.Font = $font

$reenableSearchTextBox = New-Object System.Windows.Forms.TextBox
$reenableSearchTextBox.Location = New-Object System.Drawing.Point(180, 98)
$reenableSearchTextBox.Size = New-Object System.Drawing.Size(360, 24)
$reenableSearchTextBox.Font = $font

$reenableUserListBox = New-Object System.Windows.Forms.ListBox
$reenableUserListBox.Location = New-Object System.Drawing.Point(20, 140)
$reenableUserListBox.Size = New-Object System.Drawing.Size(620, 180)
$reenableUserListBox.Font = $font
$reenableUserListBox.SelectionMode = "MultiExtended"

$reenableButton = New-Object System.Windows.Forms.Button
$reenableButton.Text = "Re-enable Selected"
$reenableButton.Location = New-Object System.Drawing.Point(180, 330)
$reenableButton.Size = New-Object System.Drawing.Size(180, 36)
$reenableButton.Font = $font

$reenableSyncButton = New-Object System.Windows.Forms.Button
$reenableSyncButton.Text = "SYNC NOW"
$reenableSyncButton.Location = New-Object System.Drawing.Point(370, 330)
$reenableSyncButton.Size = New-Object System.Drawing.Size(120, 36)
$reenableSyncButton.Font = $font
$reenableSyncButton.Add_Click({
    Invoke-DirectorySync -StatusLabel $reenableStatusLabel -LogTextBox $reenableLogTextBox
})

$reenableLogLabel = New-Object System.Windows.Forms.Label
$reenableLogLabel.Text = "Logs"
$reenableLogLabel.Location = New-Object System.Drawing.Point(20, 380)
$reenableLogLabel.Size = New-Object System.Drawing.Size(150, 24)
$reenableLogLabel.Font = $font

$reenableLogTextBox = New-Object System.Windows.Forms.TextBox
$reenableLogTextBox.Location = New-Object System.Drawing.Point(180, 378)
$reenableLogTextBox.Size = New-Object System.Drawing.Size(460, 56)
$reenableLogTextBox.Multiline = $true
$reenableLogTextBox.ReadOnly = $true
$reenableLogTextBox.Font = $font

$reenableStatusLabel = New-Object System.Windows.Forms.Label
$reenableStatusLabel.Text = "Ready"
$reenableStatusLabel.Location = New-Object System.Drawing.Point(20, 440)
$reenableStatusLabel.Size = New-Object System.Drawing.Size(620, 24)
$reenableStatusLabel.Font = $font
$reenableStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

$resetOuLabel = New-Object System.Windows.Forms.Label
$resetOuLabel.Text = "Users OU"
$resetOuLabel.Location = New-Object System.Drawing.Point(20, 60)
$resetOuLabel.Size = New-Object System.Drawing.Size(150, 24)
$resetOuLabel.Font = $font

$resetOuComboBox = New-Object System.Windows.Forms.ComboBox
$resetOuComboBox.Location = New-Object System.Drawing.Point(180, 58)
$resetOuComboBox.Size = New-Object System.Drawing.Size(360, 24)
$resetOuComboBox.Font = $font
$resetOuComboBox.DropDownStyle = "DropDown"

$refreshResetOuButton = New-Object System.Windows.Forms.Button
$refreshResetOuButton.Text = "Refresh OUs"
$refreshResetOuButton.Location = New-Object System.Drawing.Point(550, 56)
$refreshResetOuButton.Size = New-Object System.Drawing.Size(90, 28)
$refreshResetOuButton.Font = $font

$resetSearchLabel = New-Object System.Windows.Forms.Label
$resetSearchLabel.Text = "Search"
$resetSearchLabel.Location = New-Object System.Drawing.Point(20, 100)
$resetSearchLabel.Size = New-Object System.Drawing.Size(120, 24)
$resetSearchLabel.Font = $font

$resetSearchTextBox = New-Object System.Windows.Forms.TextBox
$resetSearchTextBox.Location = New-Object System.Drawing.Point(180, 98)
$resetSearchTextBox.Size = New-Object System.Drawing.Size(360, 24)
$resetSearchTextBox.Font = $font

$resetUserListBox = New-Object System.Windows.Forms.ListBox
$resetUserListBox.Location = New-Object System.Drawing.Point(20, 140)
$resetUserListBox.Size = New-Object System.Drawing.Size(620, 180)
$resetUserListBox.Font = $font
$resetUserListBox.SelectionMode = "MultiExtended"

$resetPasswordButton = New-Object System.Windows.Forms.Button
$resetPasswordButton.Text = "Reset Password"
$resetPasswordButton.Location = New-Object System.Drawing.Point(180, 330)
$resetPasswordButton.Size = New-Object System.Drawing.Size(180, 36)
$resetPasswordButton.Font = $font

$resetSyncButton = New-Object System.Windows.Forms.Button
$resetSyncButton.Text = "SYNC NOW"
$resetSyncButton.Location = New-Object System.Drawing.Point(370, 330)
$resetSyncButton.Size = New-Object System.Drawing.Size(120, 36)
$resetSyncButton.Font = $font
$resetSyncButton.Add_Click({
    Invoke-DirectorySync -StatusLabel $resetStatusLabel -LogTextBox $resetLogTextBox
})

$resetLogLabel = New-Object System.Windows.Forms.Label
$resetLogLabel.Text = "Logs"
$resetLogLabel.Location = New-Object System.Drawing.Point(20, 380)
$resetLogLabel.Size = New-Object System.Drawing.Size(150, 24)
$resetLogLabel.Font = $font

$resetLogTextBox = New-Object System.Windows.Forms.TextBox
$resetLogTextBox.Location = New-Object System.Drawing.Point(180, 378)
$resetLogTextBox.Size = New-Object System.Drawing.Size(460, 56)
$resetLogTextBox.Multiline = $true
$resetLogTextBox.ReadOnly = $true
$resetLogTextBox.Font = $font

$resetStatusLabel = New-Object System.Windows.Forms.Label
$resetStatusLabel.Text = "Ready"
$resetStatusLabel.Location = New-Object System.Drawing.Point(20, 440)
$resetStatusLabel.Size = New-Object System.Drawing.Size(620, 24)
$resetStatusLabel.Font = $font
$resetStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

$updatePreview = {
    $first = $firstNameTextBox.Text
    $last = $lastNameTextBox.Text
    $domain = $domainComboBox.Text

    if ([string]::IsNullOrWhiteSpace($first) -or
        [string]::IsNullOrWhiteSpace($last) -or
        [string]::IsNullOrWhiteSpace($domain)) {
        $previewTextBox.Text = ""
        return
    }

    $sam = Get-SamAccountName -FirstName $first -LastName $last
    $mailNickname = $sam
    $upn = "$sam@$domain"
    $proxy = (Get-ProxyAddresses -MailNickname $mailNickname -Domain $domain) -join "; "

    $previewTextBox.Text =
        "samAccountName: $sam`r`n" +
        "userPrincipalName: $upn`r`n" +
        "mailNickname: $mailNickname`r`n" +
        "proxyAddresses: $proxy"
}


function Update-CreateButtonState {
    $first = $firstNameTextBox.Text.Trim()
    $last = $lastNameTextBox.Text.Trim()
    $domain = $domainComboBox.Text.Trim()
    $ouSelection = $ouComboBox.Text.Trim()
    $ouDn = Resolve-OuDn -Selection $ouSelection

    $isValid = -not ([string]::IsNullOrWhiteSpace($first) -or
        [string]::IsNullOrWhiteSpace($last) -or
        [string]::IsNullOrWhiteSpace($domain) -or
        [string]::IsNullOrWhiteSpace($ouDn))

    $createButton.Enabled = $isValid
    if (-not $isValid) {
        Update-Status -Label $createStatusLabel -Message "Fill in all required fields to enable Create User." -Color ([System.Drawing.Color]::DarkRed)
    }
    else {
        Update-Status -Label $createStatusLabel -Message "Ready" -Color ([System.Drawing.Color]::DarkSlateGray)
    }
}

function Load-UsersFromOu {
    $userListBox.Items.Clear()
    $script:terminateUserMap = @{}
    $script:terminateUserDisplayNames = @()

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $terminateStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $ouSelection = $terminateOuComboBox.Text.Trim()
    $ouDn = Resolve-OuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($ouDn)) {
        Update-Status -Label $terminateStatusLabel -Message "Select or enter an OU distinguished name." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $users = Get-ADUser -Filter * -SearchBase $ouDn -SearchScope OneLevel -Properties UserPrincipalName, SamAccountName | Sort-Object Name
    }
    catch {
        Update-Status -Label $terminateStatusLabel -Message "Unable to query users in OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    foreach ($user in $users) {
        $display = $user.Name
        $script:terminateUserMap[$display] = $user
        $script:terminateUserDisplayNames += $display
    }

    Filter-TerminateUserList
    Update-Status -Label $terminateStatusLabel -Message "Loaded $($users.Count) users from OU." -Color ([System.Drawing.Color]::DarkGreen)
    $terminatePreviewTextBox.Text = "Loaded $($users.Count) users from $ouDn"
}

function Load-ReenableUsersFromOu {
    $reenableUserListBox.Items.Clear()
    $script:reenableUserMap = @{}
    $script:reenableUserDisplayNames = @()

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $reenableStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $ouSelection = $reenableOuComboBox.Text.Trim()
    $ouDn = Resolve-DisabledOuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($ouDn)) {
        Update-Status -Label $reenableStatusLabel -Message "Select or enter a disabled users OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $users = Get-ADUser -Filter * -SearchBase $ouDn -SearchScope OneLevel -Properties UserPrincipalName, SamAccountName | Sort-Object Name
    }
    catch {
        Update-Status -Label $reenableStatusLabel -Message "Unable to query users in OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    foreach ($user in $users) {
        $display = $user.Name
        $script:reenableUserMap[$display] = $user
        $script:reenableUserDisplayNames += $display
    }

    Filter-ReenableUserList
    Update-Status -Label $reenableStatusLabel -Message "Loaded $($users.Count) users from OU." -Color ([System.Drawing.Color]::DarkGreen)
    $reenableLogTextBox.Text = "Loaded $($users.Count) users from $ouDn"
}

function Filter-TerminateUserList {
    $filterText = $terminateSearchTextBox.Text.Trim().ToLowerInvariant()
    $userListBox.Items.Clear()
    foreach ($display in $script:terminateUserDisplayNames) {
        if ([string]::IsNullOrWhiteSpace($filterText) -or $display.ToLowerInvariant().Contains($filterText)) {
            [void]$userListBox.Items.Add($display)
        }
    }
}

function Filter-ReenableUserList {
    $filterText = $reenableSearchTextBox.Text.Trim().ToLowerInvariant()
    $reenableUserListBox.Items.Clear()
    foreach ($display in $script:reenableUserDisplayNames) {
        if ([string]::IsNullOrWhiteSpace($filterText) -or $display.ToLowerInvariant().Contains($filterText)) {
            [void]$reenableUserListBox.Items.Add($display)
        }
    }
}

function Load-ResetUsersFromOu {
    $resetUserListBox.Items.Clear()
    $script:resetUserMap = @{}
    $script:resetUserDisplayNames = @()

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $resetStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $ouSelection = $resetOuComboBox.Text.Trim()
    $ouDn = Resolve-OuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($ouDn)) {
        Update-Status -Label $resetStatusLabel -Message "Select or enter an OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    try {
        $users = Get-ADUser -Filter * -SearchBase $ouDn -SearchScope OneLevel -Properties UserPrincipalName, SamAccountName | Sort-Object Name
    }
    catch {
        Update-Status -Label $resetStatusLabel -Message "Unable to query users in OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    foreach ($user in $users) {
        $display = $user.Name
        $script:resetUserMap[$display] = $user
        $script:resetUserDisplayNames += $display
    }

    Filter-ResetUserList
    Update-Status -Label $resetStatusLabel -Message "Loaded $($users.Count) users from OU." -Color ([System.Drawing.Color]::DarkGreen)
    $resetLogTextBox.Text = "Loaded $($users.Count) users from $ouDn"
}

function Filter-ResetUserList {
    $filterText = $resetSearchTextBox.Text.Trim().ToLowerInvariant()
    $resetUserListBox.Items.Clear()
    foreach ($display in $script:resetUserDisplayNames) {
        if ([string]::IsNullOrWhiteSpace($filterText) -or $display.ToLowerInvariant().Contains($filterText)) {
            [void]$resetUserListBox.Items.Add($display)
        }
    }
}

$firstNameTextBox.Add_TextChanged({
    $updatePreview.Invoke()
    Update-CreateButtonState
})
$lastNameTextBox.Add_TextChanged({
    $updatePreview.Invoke()
    Update-CreateButtonState
})
$domainComboBox.Add_SelectedIndexChanged({
    $updatePreview.Invoke()
    Update-CreateButtonState
})
$ouComboBox.Add_TextChanged({
    Update-CreateButtonState
})

$terminateOuComboBox.Add_TextChanged({
    Load-UsersFromOu
})

$terminateSearchTextBox.Add_TextChanged({
    Filter-TerminateUserList
})

$reenableOuComboBox.Add_TextChanged({
    Load-ReenableUsersFromOu
})

$reenableSearchTextBox.Add_TextChanged({
    Filter-ReenableUserList
})

$resetOuComboBox.Add_TextChanged({
    Load-ResetUsersFromOu
})

$resetSearchTextBox.Add_TextChanged({
    Filter-ResetUserList
})

$refreshTerminateOuButton.Add_Click({
    Load-OUs -ComboBox $terminateOuComboBox -StatusLabel $terminateStatusLabel
})

$refreshReenableOuButton.Add_Click({
    Load-DisabledOUs -ComboBox $reenableOuComboBox -StatusLabel $reenableStatusLabel
})

$refreshResetOuButton.Add_Click({
    Load-OUs -ComboBox $resetOuComboBox -StatusLabel $resetStatusLabel
})

$copyCredentialsButton.Add_Click({
    if ($script:lastGeneratedCredentials) {
        [System.Windows.Forms.Clipboard]::SetText($script:lastGeneratedCredentials)
        $previewTextBox.Text = "Credentials copied to clipboard."
    }
    else {
        $previewTextBox.Text = "No auto-generated credentials available."
    }
})

$createButton.Add_Click({
    $first = $firstNameTextBox.Text.Trim()
    $last  = $lastNameTextBox.Text.Trim()
    $domain = $domainComboBox.Text.Trim()

    $ouSelection = $ouComboBox.Text.Trim()
    $ouDn = Resolve-OuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($first) -or
        [string]::IsNullOrWhiteSpace($last) -or
        [string]::IsNullOrWhiteSpace($domain)) {
        Update-Status -Label $createStatusLabel -Message "First name, last name, and domain are required." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if ([string]::IsNullOrWhiteSpace($ouDn)) {
        Update-Status -Label $createStatusLabel -Message "Select or enter an OU distinguished name." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $createStatusLabel -Message "ActiveDirectory module not available. Cannot create user." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if (-not (Confirm-Action -Message "Create user $first $last in $ouDn?" -Title "Confirm Create User")) {
        Update-Status -Label $createStatusLabel -Message "Create user cancelled." -Color ([System.Drawing.Color]::DarkSlateGray)
        return
    }

    $sam = Get-SamAccountName -FirstName $first -LastName $last
    $mailNickname = $sam
    $upn = "$sam@$domain"
    $proxyAddresses = Get-ProxyAddresses -MailNickname $mailNickname -Domain $domain

    $plainPassword = $passwordTextBox.Text.Trim()
    $generatedPassword = $false
    if ([string]::IsNullOrWhiteSpace($plainPassword)) {
        $plainPassword = New-RandomPassword
        $generatedPassword = $true
    }
    $securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force

    $otherAttributes = @{
        proxyAddresses = $proxyAddresses
        mailNickname   = $mailNickname
    }

    try {
        $params = @{
            Name              = "$first $last"
            GivenName         = $first
            Surname           = $last
            SamAccountName    = $sam
            UserPrincipalName = $upn
            Path              = $ouDn
            Enabled           = $true
            OtherAttributes   = $otherAttributes
            AccountPassword   = $securePassword
        }

        if ($mustChangePasswordCheckBox.Checked) {
            $params.ChangePasswordAtLogon = $true
        }

        New-ADUser @params
        Update-Status -Label $createStatusLabel -Message "User created successfully." -Color ([System.Drawing.Color]::DarkGreen)

        if ($generatedPassword) {
            $credentialText = "Username: $sam`r`nEmail: $upn`r`nPassword: $plainPassword"
            $script:lastGeneratedCredentials = $credentialText
            $copyCredentialsButton.Enabled = $true
            [System.Windows.Forms.Clipboard]::SetText($credentialText)
            [void][System.Windows.Forms.MessageBox]::Show(
                "Auto-generated credentials copied to clipboard.`r`n`r`n$credentialText",
                "User Information",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            $previewTextBox.Text = "User created. Auto-generated credentials copied to clipboard."
        }
        else {
            $script:lastGeneratedCredentials = $null
            $copyCredentialsButton.Enabled = $false
            $previewTextBox.Text = "User created."
        }
    }
    catch {
        Update-Status -Label $createStatusLabel -Message "Failed to create user: $($_.Exception.Message)" -Color ([System.Drawing.Color]::DarkRed)
    }
})

$terminateButton.Add_Click({
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $terminateStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $ouSelection = $terminateOuComboBox.Text.Trim()
    $sourceOuDn = Resolve-OuDn -Selection $ouSelection

    if ([string]::IsNullOrWhiteSpace($sourceOuDn)) {
        Update-Status -Label $terminateStatusLabel -Message "Select or enter a source OU." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $disabledOuDn = Get-DisableOuFromUserOu -UserOuDn $sourceOuDn
    if (-not $disabledOuDn) {
        Update-Status -Label $terminateStatusLabel -Message "Unable to resolve Disabled Users OU. Expected OU=Users, in path." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if ($userListBox.SelectedItems.Count -eq 0) {
        Update-Status -Label $terminateStatusLabel -Message "Select at least one user to terminate." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $selectedUsers = $userListBox.SelectedItems -join ', '
    if (-not (Confirm-Action -Message "Terminate selected users?`r`n$selectedUsers" -Title "Confirm Termination")) {
        Update-Status -Label $terminateStatusLabel -Message "Termination cancelled." -Color ([System.Drawing.Color]::DarkSlateGray)
        return
    }

    $failed = @()
    $completed = @()
    foreach ($display in $userListBox.SelectedItems) {
        $user = $script:terminateUserMap[$display]
        if (-not $user) {
            $failed += $display
            continue
        }

        try {
            Disable-ADAccount -Identity $user.DistinguishedName
            Set-ADUser -Identity $user.DistinguishedName -Replace @{msExchHideFromAddressLists = $true}
            $newPassword = New-RandomPassword
            $securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
            Set-ADAccountPassword -Identity $user.DistinguishedName -Reset -NewPassword $securePassword
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $disabledOuDn
            $completed += $display
        }
        catch {
            $failed += $display
        }
    }

    $terminatePreviewTextBox.Text = "Steps: Disabled account, reset password, moved to Disabled Users OU.`r`n"
    if ($completed.Count -gt 0) {
        $terminatePreviewTextBox.Text += "Completed: $($completed -join ', ')"
    }
    if ($failed.Count -gt 0) {
        if ($terminatePreviewTextBox.Text.Length -gt 0) {
            $terminatePreviewTextBox.Text += "`r`n"
        }
        $terminatePreviewTextBox.Text += "Failed: $($failed -join ', ')"
    }

    if ($failed.Count -gt 0) {
        Update-Status -Label $terminateStatusLabel -Message "Completed with errors. Failed: $($failed -join ', ')" -Color ([System.Drawing.Color]::DarkRed)
    }
    else {
        Update-Status -Label $terminateStatusLabel -Message "Selected users terminated and moved to Disabled Users OU." -Color ([System.Drawing.Color]::DarkGreen)
        $terminatedUsers = $completed -join ', '
        [void][System.Windows.Forms.MessageBox]::Show(
            "Terminated users:`r`n$terminatedUsers",
            "User Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

$reenableButton.Add_Click({
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $reenableStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if ($reenableUserListBox.SelectedItems.Count -eq 0) {
        Update-Status -Label $reenableStatusLabel -Message "Select at least one user to re-enable." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $selectedUsers = $reenableUserListBox.SelectedItems -join ', '
    if (-not (Confirm-Action -Message "Re-enable selected users?`r`n$selectedUsers" -Title "Confirm Re-enable")) {
        Update-Status -Label $reenableStatusLabel -Message "Re-enable cancelled." -Color ([System.Drawing.Color]::DarkSlateGray)
        return
    }

    $failed = @()
    $completedCredentials = @()
    foreach ($display in $reenableUserListBox.SelectedItems) {
        $user = $script:reenableUserMap[$display]
        if (-not $user) {
            $failed += $display
            continue
        }

        try {
            Enable-ADAccount -Identity $user.DistinguishedName
            $newPassword = New-RandomPassword
            $securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
            Set-ADAccountPassword -Identity $user.DistinguishedName -Reset -NewPassword $securePassword
            Set-ADUser -Identity $user.DistinguishedName -Replace @{msExchHideFromAddressLists = $false}

            $userOuDn = $null
            if ($user.DistinguishedName -match "^CN=.*?,(.*)$") {
                $userOuDn = $Matches[1]
            }
            $targetOu = Get-UsersOuFromDisabledOu -DisabledDn $userOuDn
            if (-not $targetOu) {
                throw "Unable to resolve Users OU for $($user.Name)."
            }
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $targetOu
            $userUpn = $user.UserPrincipalName
            if ([string]::IsNullOrWhiteSpace($userUpn)) {
                $userUpn = $user.SamAccountName
            }
            $completedCredentials += "Username: $($user.SamAccountName)`r`nEmail: $userUpn`r`nPassword: $newPassword"
        }
        catch {
            $failed += $display
        }
    }

    if ($failed.Count -gt 0) {
        Update-Status -Label $reenableStatusLabel -Message "Completed with errors. Failed: $($failed -join ', ')" -Color ([System.Drawing.Color]::DarkRed)
        $reenableLogTextBox.Text = "Failed to re-enable: $($failed -join ', ')"
    }
    else {
        Update-Status -Label $reenableStatusLabel -Message "Selected users re-enabled." -Color ([System.Drawing.Color]::DarkGreen)
        $reenableLogTextBox.Text = "Re-enabled: $($reenableUserListBox.SelectedItems -join ', ')"
    }

    if ($completedCredentials.Count -gt 0) {
        $credentialText = $completedCredentials -join "`r`n`r`n"
        [System.Windows.Forms.Clipboard]::SetText($credentialText)
        Show-CredentialsDialog -Title "User Information" -Message $credentialText
    }
})

$resetPasswordButton.Add_Click({
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Update-Status -Label $resetStatusLabel -Message "ActiveDirectory module not available." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    if ($resetUserListBox.SelectedItems.Count -eq 0) {
        Update-Status -Label $resetStatusLabel -Message "Select at least one user to reset." -Color ([System.Drawing.Color]::DarkRed)
        return
    }

    $selectedUsers = $resetUserListBox.SelectedItems -join ', '
    if (-not (Confirm-Action -Message "Reset passwords for selected users?`r`n$selectedUsers" -Title "Confirm Password Reset")) {
        Update-Status -Label $resetStatusLabel -Message "Password reset cancelled." -Color ([System.Drawing.Color]::DarkSlateGray)
        return
    }

    $failed = @()
    $completedCredentials = @()
    foreach ($display in $resetUserListBox.SelectedItems) {
        $user = $script:resetUserMap[$display]
        if (-not $user) {
            $failed += $display
            continue
        }

        try {
            $newPassword = New-RandomPassword
            $securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
            Set-ADAccountPassword -Identity $user.DistinguishedName -Reset -NewPassword $securePassword
            $userUpn = $user.UserPrincipalName
            if ([string]::IsNullOrWhiteSpace($userUpn)) {
                $userUpn = $user.SamAccountName
            }
            $completedCredentials += "Username: $($user.SamAccountName)`r`nEmail: $userUpn`r`nPassword: $newPassword"
        }
        catch {
            $failed += $display
        }
    }

    if ($failed.Count -gt 0) {
        Update-Status -Label $resetStatusLabel -Message "Completed with errors. Failed: $($failed -join ', ')" -Color ([System.Drawing.Color]::DarkRed)
        $resetLogTextBox.Text = "Failed to reset: $($failed -join ', ')"
    }
    else {
        Update-Status -Label $resetStatusLabel -Message "Selected users reset." -Color ([System.Drawing.Color]::DarkGreen)
        $resetLogTextBox.Text = "Password reset for: $($resetUserListBox.SelectedItems -join ', ')"
    }

    if ($completedCredentials.Count -gt 0) {
        $credentialText = $completedCredentials -join "`r`n`r`n"
        [System.Windows.Forms.Clipboard]::SetText($credentialText)
        Show-CredentialsDialog -Title "User Information" -Message $credentialText
    }
})

$createPanel.Controls.AddRange(@(
    $backToMenuFromCreate,
    $firstNameLabel,
    $firstNameTextBox,
    $lastNameLabel,
    $lastNameTextBox,
    $domainLabel,
    $domainComboBox,
    $ouLabel,
    $ouComboBox,
    $passwordLabel,
    $passwordTextBox,
    $mustChangePasswordCheckBox,
    $copyCredentialsButton,
    $previewLabel,
    $previewTextBox,
    $createButton,
    $createSyncButton,
    $createStatusLabel
))

$terminatePanel.Controls.AddRange(@(
    $backToMenuFromTerminate,
    $terminateOuLabel,
    $terminateOuComboBox,
    $refreshTerminateOuButton,
    $terminateSearchLabel,
    $terminateSearchTextBox,
    $userListBox,
    $terminateButton,
    $terminateSyncButton,
    $terminatePreviewLabel,
    $terminatePreviewTextBox,
    $terminateStatusLabel
))

$reenablePanel.Controls.AddRange(@(
    $backToMenuFromReenable,
    $reenableOuLabel,
    $reenableOuComboBox,
    $refreshReenableOuButton,
    $reenableSearchLabel,
    $reenableSearchTextBox,
    $reenableUserListBox,
    $reenableButton,
    $reenableSyncButton,
    $reenableLogLabel,
    $reenableLogTextBox,
    $reenableStatusLabel
))

$resetPanel.Controls.AddRange(@(
    $backToMenuFromReset,
    $resetOuLabel,
    $resetOuComboBox,
    $refreshResetOuButton,
    $resetSearchLabel,
    $resetSearchTextBox,
    $resetUserListBox,
    $resetPasswordButton,
    $resetSyncButton,
    $resetLogLabel,
    $resetLogTextBox,
    $resetStatusLabel
))

$menuGrid.Controls.Add($createTileButton, 0, 0)
$menuGrid.Controls.Add($terminateTileButton, 1, 0)
$menuGrid.Controls.Add($reenableTileButton, 2, 0)
$menuGrid.Controls.Add($resetTileButton, 3, 0)
$menuGrid.Controls.Add($dummyTileButton1, 0, 1)
$menuGrid.Controls.Add($dummyTileButton2, 1, 1)
$menuGrid.Controls.Add($dummyTileButton3, 2, 1)
$menuGrid.Controls.Add($dummyTileButton4, 3, 1)
$mainMenuPanel.Controls.AddRange(@(
    $menuGrid,
    $menuTitleLabel,
    $menuSyncButton,
    $menuStrip
))

$form.MainMenuStrip = $menuStrip
$form.Controls.Add($footerStatusLabel)
$form.Controls.Add($mainMenuPanel)
$form.Controls.Add($createPanel)
$form.Controls.Add($terminatePanel)
$form.Controls.Add($reenablePanel)
$form.Controls.Add($resetPanel)

$form.Add_Shown({
    $form.Activate()

    $domainValues = Get-AvailableDomains
    if ($domainValues -and $domainValues.Count -gt 0) {
        foreach ($domain in $domainValues) {
            [void]$domainComboBox.Items.Add($domain)
        }
        $domainComboBox.SelectedIndex = 0
        Update-Status -Label $createStatusLabel -Message "Domains loaded from Active Directory." -Color ([System.Drawing.Color]::DarkGreen)
    }
    else {
        Update-Status -Label $createStatusLabel -Message "Unable to load domains from Active Directory." -Color ([System.Drawing.Color]::DarkRed)
    }

    Load-OUs -ComboBox $ouComboBox -StatusLabel $createStatusLabel
    Load-OUs -ComboBox $terminateOuComboBox -StatusLabel $terminateStatusLabel
    Load-DisabledOUs -ComboBox $reenableOuComboBox -StatusLabel $reenableStatusLabel
    Load-OUs -ComboBox $resetOuComboBox -StatusLabel $resetStatusLabel

    Register-ButtonIcon -Button $createTileButton -Path "C:\\JUMP\\enableuser.ico"
    Register-ButtonIcon -Button $terminateTileButton -Path "C:\\JUMP\\termuser.ico"
    Register-ButtonIcon -Button $reenableTileButton -Path "C:\\JUMP\\reenable.ico"
    Register-ButtonIcon -Button $resetTileButton -Path "C:\\JUMP\\pwreset.ico"
    Register-ButtonIcon -Button $dummyTileButton1 -Path "C:\\JUMP\\underdev.ico"
    Register-ButtonIcon -Button $dummyTileButton2 -Path "C:\\JUMP\\underdev.ico"
    Register-ButtonIcon -Button $dummyTileButton3 -Path "C:\\JUMP\\underdev.ico"
    Register-ButtonIcon -Button $dummyTileButton4 -Path "C:\\JUMP\\underdev.ico"

    $hoverColor = [System.Drawing.Color]::LightSteelBlue
    Register-HoverHighlight -Button $createTileButton -HoverColor $hoverColor
    Register-HoverHighlight -Button $terminateTileButton -HoverColor $hoverColor
    Register-HoverHighlight -Button $reenableTileButton -HoverColor $hoverColor
    Register-HoverHighlight -Button $resetTileButton -HoverColor $hoverColor
    Register-HoverHighlight -Button $dummyTileButton1 -HoverColor $hoverColor
    Register-HoverHighlight -Button $dummyTileButton2 -HoverColor $hoverColor
    Register-HoverHighlight -Button $dummyTileButton3 -HoverColor $hoverColor
    Register-HoverHighlight -Button $dummyTileButton4 -HoverColor $hoverColor
})

[void]$form.ShowDialog()
