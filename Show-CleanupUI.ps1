<#
.SYNOPSIS
    Windows Forms front-end for the Qlik private-content cleanup.

.DESCRIPTION
    - Loads all apps (name + ID) from the Repository API so you can pick one.
    - Pick a deletion window: 30, 60, 90 or 120 days.
    - Optionally include orphaned private content: sheets/bookmarks with no
      owner reference, or whose owner has been removed from its directory or
      blacklisted in QRS. These are found regardless of age, since the owner
      can no longer log in to clean them up.
    - Optionally protect specific owners: enter one or more DOMAIN\user
      entries (exact UserDirectory\UserId match) and their private content is
      excluded from the dated scan and kept. Note this protects only the
      normal date-window results; orphaned content owned by a protected user
      is still deleted, since an orphaned owner is gone/blocked anyway.
    - Preview (dry run) lists the private sheets/bookmarks that match.
    - Delete removes them via the Engine API (enigma.js) or QRS metadata.
      Orphans with no owner at all can't be impersonated in the engine, so
      those are cleaned up via QRS metadata delete only.

    Nothing is deleted until you click Delete and confirm.

.NOTES
    Run with PowerShell 7+ on Windows:
        pwsh -File .\Show-CleanupUI.ps1
    WinForms needs an STA thread; pwsh runs STA by default. If ShowDialog
    throws an apartment-state error, start pwsh, then dot-run this script.
#>

[CmdletBinding()]
param(
    [string]$Server             = 'qlik.example.com',
    [string]$CertDir            = 'C:\qlik-certs',
    [string]$AdminUserDirectory = 'INTERNAL',
    [string]$AdminUserId        = 'sa_repository',
    [string]$EngineSchema       = '12.612.0',
    [string]$DestroyScript      = (Join-Path $PSScriptRoot 'destroy-objects.cjs'),
    [string]$OutDir             = (Join-Path $PSScriptRoot 'run')
)

Import-Module (Join-Path $PSScriptRoot 'QlikCleanup.Common.psm1') -Force
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# ---- shared state --------------------------------------------------------
$script:Cert        = $null
$script:AllApps     = @()
$script:Manifest    = @()
$script:SelectedApp = $null

# ======================= form scaffold ====================================
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Qlik - Delete recent private content'
$form.Size            = New-Object System.Drawing.Size(900, 812)
$form.StartPosition   = 'CenterScreen'
$form.MinimumSize     = New-Object System.Drawing.Size(820, 732)

function New-Label($text, $x, $y, $w = 110) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text; $l.Location = "$x,$y"; $l.Size = "$w,22"
    $l.TextAlign = 'MiddleLeft'; $l
}
function New-Text($val, $x, $y, $w) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Text = $val; $t.Location = "$x,$y"; $t.Size = "$w,22"; $t
}

# ---- connection row ------------------------------------------------------
$form.Controls.Add((New-Label 'Server' 12 15))
$txtServer  = New-Text $Server 70 13 230
$form.Controls.Add($txtServer)

$form.Controls.Add((New-Label 'Cert dir' 320 15 60))
$txtCert    = New-Text $CertDir 385 13 200
$form.Controls.Add($txtCert)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = 'Load apps'; $btnLoad.Location = '600,12'; $btnLoad.Size = '110,24'
$form.Controls.Add($btnLoad)

$lblConn = New-Label "$AdminUserDirectory\$AdminUserId" 720 15 160
$lblConn.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblConn)

# ---- app filter + list ---------------------------------------------------
$form.Controls.Add((New-Label 'Filter (name or ID)' 12 48 120))
$txtFilter = New-Text '' 140 46 560
$form.Controls.Add($txtFilter)

$lst = New-Object System.Windows.Forms.ListView
$lst.Location = '12,76'; $lst.Size = '860,300'
$lst.View = 'Details'; $lst.FullRowSelect = $true; $lst.MultiSelect = $false
$lst.HideSelection = $false; $lst.GridLines = $true
$lst.Anchor = 'Top,Left,Right'
[void]$lst.Columns.Add('Name', 300)
[void]$lst.Columns.Add('Stream', 130)
[void]$lst.Columns.Add('Owner', 170)
[void]$lst.Columns.Add('Last reload', 130)
[void]$lst.Columns.Add('App ID', 280)
$form.Controls.Add($lst)

# ---- timeframe -----------------------------------------------------------
$grpDays = New-Object System.Windows.Forms.GroupBox
$grpDays.Text = 'Delete private content updated in the last...'
$grpDays.Location = '12,386'; $grpDays.Size = '430,60'
$form.Controls.Add($grpDays)
$radioDays = @{}
$i = 0
foreach ($d in 30, 60, 90, 120) {
    $r = New-Object System.Windows.Forms.RadioButton
    $r.Text = "$d days"; $r.Location = (New-Object System.Drawing.Point((12 + $i * 78), 25))
    $r.Size = '74,22'; $r.Tag = $d
    if ($d -eq 90) { $r.Checked = $true }
    $r.Add_CheckedChanged({ $btnDelete.Enabled = $false })
    $grpDays.Controls.Add($r); $radioDays[$d] = $r; $i++
}
# "All" - ignore the last-updated date and target every private sheet/bookmark.
$rAll = New-Object System.Windows.Forms.RadioButton
$rAll.Text = 'All'; $rAll.Location = (New-Object System.Drawing.Point((12 + 4 * 78), 25))
$rAll.Size = '64,22'; $rAll.Tag = 0
$rAll.Add_CheckedChanged({ $btnDelete.Enabled = $false })
$grpDays.Controls.Add($rAll); $radioDays[0] = $rAll

# ---- orphaned content ------------------------------------------------------
$chkOrphans = New-Object System.Windows.Forms.CheckBox
$chkOrphans.Text = 'Include orphaned content (deleted/removed owners, any age)'
$chkOrphans.Location = '12,448'; $chkOrphans.Size = '430,22'
$chkOrphans.AutoSize = $false
$form.Controls.Add($chkOrphans)

# ---- protected users (exclude their private content from deletion) --------
# Exact match on UserDirectory\UserId. One entry per line, or separated by ';'.
# Applies to the dated scan only - orphaned content is still deleted even if
# owned by a protected user, since an orphaned owner is gone/blocked anyway.
$lblProtect = New-Object System.Windows.Forms.Label
$lblProtect.Text = "Protect these owners (DOMAIN\user, one per line or ';'-separated) - their private content is kept:"
$lblProtect.Location = '12,474'; $lblProtect.Size = '860,18'
$form.Controls.Add($lblProtect)

$txtProtect = New-Object System.Windows.Forms.TextBox
$txtProtect.Location = '12,494'; $txtProtect.Size = '860,44'
$txtProtect.Multiline = $true; $txtProtect.ScrollBars = 'Vertical'
$txtProtect.Anchor = 'Top,Left,Right'
$form.Controls.Add($txtProtect)

# ---- method note (deletion always does both engine + QRS) ----------------
$lblMethod = New-Object System.Windows.Forms.Label
$lblMethod.Text = "Delete removes the content from" + [Environment]::NewLine + "both the Engine and the QRS catalog."
$lblMethod.Location = '446,390'; $lblMethod.Size = '236,44'
$lblMethod.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblMethod)

# ---- action buttons ------------------------------------------------------
$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = 'Preview (dry run)'; $btnPreview.Location = '688,392'; $btnPreview.Size = '180,24'
$form.Controls.Add($btnPreview)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = 'Delete selected content'; $btnDelete.Location = '688,420'; $btnDelete.Size = '180,26'
$btnDelete.Enabled = $false
$btnDelete.ForeColor = [System.Drawing.Color]::DarkRed
$form.Controls.Add($btnDelete)

# ---- status + log --------------------------------------------------------
$lblStatus = New-Label 'Load apps to begin.' 12 544 860
$lblStatus.Font = New-Object System.Drawing.Font($lblStatus.Font, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$log = New-Object System.Windows.Forms.TextBox
$log.Location = '12,570'; $log.Size = '860,190'
$log.Multiline = $true; $log.ScrollBars = 'Vertical'; $log.ReadOnly = $true
$log.Font = New-Object System.Drawing.Font('Consolas', 9)
$log.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($log)

# ======================= helpers ==========================================
function Add-Log([string]$msg) {
    $log.AppendText(('{0}  {1}{2}' -f (Get-Date -Format 'HH:mm:ss'), $msg, [Environment]::NewLine))
    [System.Windows.Forms.Application]::DoEvents()
}
function Set-Busy([bool]$busy) {
    $form.Cursor = if ($busy) { 'WaitCursor' } else { 'Default' }
    foreach ($c in @($btnLoad, $btnPreview, $btnDelete, $txtServer, $txtCert, $chkOrphans, $txtProtect)) { $c.Enabled = -not $busy }
    [System.Windows.Forms.Application]::DoEvents()
}
function Get-SelectedDays {
    ($radioDays.GetEnumerator() | Where-Object { $_.Value.Checked } | Select-Object -First 1).Key
}
function Get-WindowText([int]$d) {
    if ($d -le 0) { 'any date (ALL private content)' } else { "the last $d days" }
}
function Update-AppList {
    $lst.BeginUpdate(); $lst.Items.Clear()
    $f = $txtFilter.Text.Trim()
    $rows = if ($f) {
        $script:AllApps | Where-Object { $_.Name -like "*$f*" -or $_.Id -like "*$f*" }
    } else { $script:AllApps }
    foreach ($a in $rows) {
        $it = New-Object System.Windows.Forms.ListViewItem
        $it.Text = [string]$a.Name
        [void]$it.SubItems.Add([string]$a.Stream)
        [void]$it.SubItems.Add([string]$a.Owner)
        [void]$it.SubItems.Add([string]$a.LastReload)
        [void]$it.SubItems.Add([string]$a.Id)
        $it.Tag = $a
        [void]$lst.Items.Add($it)
    }
    $lst.EndUpdate()
}
function Get-CertOrThrow {
    if (-not $script:Cert) { $script:Cert = Get-QlikClientCertificate -CertDir $txtCert.Text.Trim() }
    $script:Cert
}

# ======================= event handlers ===================================
$btnLoad.Add_Click({
    try {
        Set-Busy $true
        $script:Cert = $null
        $cert = Get-CertOrThrow
        Add-Log "Loading apps from $($txtServer.Text) ..."
        $script:AllApps = @(Get-QlikApps -Server $txtServer.Text.Trim() -Certificate $cert `
                              -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId |
                            Sort-Object Name)
        Update-AppList
        $lblStatus.Text = "Loaded $($script:AllApps.Count) app(s). Select one, choose a timeframe, then Preview."
        Add-Log "Loaded $($script:AllApps.Count) app(s)."
    } catch {
        Add-Log "ERROR loading apps: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Load failed', 'OK', 'Error') | Out-Null
    } finally { Set-Busy $false }
})

$txtFilter.Add_TextChanged({ Update-AppList })

# Any change that affects what Preview would produce must force a re-Preview
# before Delete is allowed, so the manifest can never be out of sync with the
# current settings (e.g. editing protected owners after previewing).
$txtProtect.Add_TextChanged({ $btnDelete.Enabled = $false })
$chkOrphans.Add_CheckedChanged({ $btnDelete.Enabled = $false })

$lst.Add_SelectedIndexChanged({
    if ($lst.SelectedItems.Count -gt 0) {
        $script:SelectedApp = $lst.SelectedItems[0].Tag
        $btnDelete.Enabled = $false        # require a fresh preview before delete
        $lblStatus.Text = "Selected: $($script:SelectedApp.Name)  [$($script:SelectedApp.Id)] - click Preview."
    }
})

$btnPreview.Add_Click({
    if (-not $script:SelectedApp) {
        [System.Windows.Forms.MessageBox]::Show('Select an app first.', 'No app', 'OK', 'Information') | Out-Null
        return
    }
    try {
        Set-Busy $true
        $days = Get-SelectedDays
        $win  = Get-WindowText $days
        $cert = Get-CertOrThrow
        Add-Log "Preview: '$($script:SelectedApp.Name)' - private sheets/bookmarks updated within $win."
        $dated = @(Get-QlikPrivateContent -Server $txtServer.Text.Trim() -Certificate $cert `
                              -AppId $script:SelectedApp.Id -Days $days `
                              -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId)

        # Protected owners: exact UserDirectory\UserId match, case-insensitive.
        # Split on newlines and semicolons, trim, drop blanks. Applied to the
        # DATED scan only - orphaned content is still deleted even if owned by a
        # protected user (an orphaned owner is gone/blocked, so "protect their
        # active work" doesn't apply).
        $protectKeys = @{}
        foreach ($line in ($txtProtect.Text -split "[`r`n;]")) {
            $entry = $line.Trim()
            if ($entry) { $protectKeys[$entry.ToLowerInvariant()] = $true }
        }
        if ($protectKeys.Count -gt 0) {
            $before = $dated.Count
            $dated = @($dated | Where-Object {
                $key = ("{0}\{1}" -f $_.ownerUserDirectory, $_.ownerUserId).ToLowerInvariant()
                -not $protectKeys.ContainsKey($key)
            })
            $kept = $before - $dated.Count
            Add-Log ("Protected owners: {0} entr{1} -> excluded {2} dated object(s) from deletion." -f `
                $protectKeys.Count, $(if ($protectKeys.Count -eq 1) { 'y' } else { 'ies' }), $kept)
        }

        $orphans = @()
        if ($chkOrphans.Checked) {
            Add-Log 'Scanning for orphaned content (no owner, or owner removed/blacklisted)...'
            $orphans = @(Get-QlikOrphanedPrivateContent -Server $txtServer.Text.Trim() -Certificate $cert `
                              -AppId $script:SelectedApp.Id `
                              -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId)
            Add-Log "Orphan scan found $($orphans.Count) object(s)."
        }

        # Merge, de-duplicated by qrsId (an orphan inside the date window would
        # otherwise show up twice).
        $seen = @{}
        $merged = @()
        foreach ($m in ($dated + $orphans)) {
            if (-not $seen.ContainsKey($m.qrsId)) { $seen[$m.qrsId] = $true; $merged += $m }
        }
        $script:Manifest = $merged

        if ($script:Manifest.Count -eq 0) {
            $lblStatus.Text = "No matching private content ($win$(if ($chkOrphans.Checked) { ' + orphans' }))."
            Add-Log 'Nothing matched.'
            $btnDelete.Enabled = $false
        } else {
            foreach ($m in $script:Manifest) {
                $orphanTag = if ($m.OrphanReason) { " [orphan: $($m.OrphanReason)]" } else { '' }
                Add-Log ("  {0,-9} {1,-30} owner={2}\{3}  updated={4}{5}" -f `
                    $m.objectType, $m.name, $m.ownerUserDirectory, $m.ownerUserId, $m.modifiedDate, $orphanTag)
            }
            $orphanCount = @($script:Manifest | Where-Object { $_.OrphanReason }).Count
            $protNote = if ($protectKeys.Count -gt 0) { " ($($protectKeys.Count) owner(s) protected)" } else { '' }
            $lblStatus.Text = "$($script:Manifest.Count) object(s) would be deleted ($win$(if ($orphanCount -gt 0) { ", incl. $orphanCount orphan(s)" }))$protNote. Review the log, then Delete."
            $btnDelete.Enabled = $true
        }
    } catch {
        Add-Log "ERROR during preview: $($_.Exception.Message)"
    } finally { Set-Busy $false }
})

$btnDelete.Add_Click({
    if (-not $script:Manifest -or $script:Manifest.Count -eq 0) { return }
    $days   = Get-SelectedDays
    $win    = Get-WindowText $days
    $orphanCount = @($script:Manifest | Where-Object { $_.OrphanReason }).Count
    $orphanNote = if ($orphanCount -gt 0) { "`nIncludes $orphanCount orphaned object(s) (no owner, or owner removed/blacklisted) - those are removed via QRS metadata delete only, since there's no live owner to impersonate in the engine." } else { '' }
    $msg = "Delete $($script:Manifest.Count) private object(s) from`n" +
           "'$($script:SelectedApp.Name)'`n" +
           "updated within $win, from both the engine and the QRS?$orphanNote`n`nThis cannot be undone."
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'Confirm delete', 'YesNo', 'Warning') -ne 'Yes') {
        Add-Log 'Delete cancelled.'; return
    }
    try {
        Set-Busy $true
        $stamp        = Get-Date -Format 'yyyyMMdd-HHmmss'
        $manifestPath = Join-Path $OutDir "objects-$stamp.json"
        $backupPath   = Join-Path $OutDir "backup-$stamp.json"
        $resultsPath  = Join-Path $OutDir "results-$stamp.json"
        # -AsArray matters here: piping a SINGLE object into ConvertTo-Json
        # serializes a bare {...} instead of [{...}], and destroy-objects.cjs's
        # Array.isArray() check would then treat that as "manifest is empty"
        # and skip engine deletion entirely - exactly the case a lone orphan
        # hits most often.
        $script:Manifest | ConvertTo-Json -Depth 10 -AsArray | Set-Content $manifestPath
        $script:Manifest | ConvertTo-Json -Depth 10 -AsArray | Set-Content $backupPath
        Add-Log "Manifest + backup written to $OutDir"

        $cert = Get-CertOrThrow

        # Engine first: removes the object inside the app. Both IDs were captured
        # at preview time, so each stage works from the manifest and never has to
        # re-find an item the other stage already deleted.
        Add-Log 'Deleting via Engine API (enigma.js)...'
        Remove-QlikPrivateContentEngine -Server $txtServer.Text.Trim() -CertDir $txtCert.Text.Trim() `
            -ManifestPath $manifestPath -ResultsPath $resultsPath -DestroyScript $DestroyScript `
            -Schema $EngineSchema -Log { param($m) Add-Log $m }

        # QRS second: removes the catalog entry. If the engine sync already
        # cleared it, the delete returns 404 and is reported as "already removed".
        Add-Log 'Cleaning up catalog metadata via QRS...'
        Remove-QlikPrivateContentQrs -Server $txtServer.Text.Trim() -Certificate $cert `
            -Manifest $script:Manifest -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId `
            -Log { param($m) Add-Log $m }

        Start-Sleep -Seconds 3   # allow engine->repository sync
        $deletedIds = @($script:Manifest | ForEach-Object { $_.qrsId })
        $remaining  = @(Get-QlikPrivateContent -Server $txtServer.Text.Trim() -Certificate $cert `
                        -AppId $script:SelectedApp.Id -Days 0 `
                        -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId |
                        Where-Object { $_.qrsId -in $deletedIds })
        $lblStatus.Text = "Done. $($remaining.Count) of $($script:Manifest.Count) targeted object(s) still remain (expect 0)."
        Add-Log "Verification: $($remaining.Count) of $($script:Manifest.Count) targeted object(s) remaining."
        $btnDelete.Enabled = $false
        $script:Manifest = @()
    } catch {
        Add-Log "ERROR during delete: $($_.Exception.Message)"
    } finally { Set-Busy $false }
})

[void]$form.ShowDialog()
$form.Dispose()
