<#
    QlikCleanup.Common.psm1
    Shared helpers for the private-content cleanup tool.
    Imported by Show-CleanupUI.ps1 (and reusable from the CLI orchestrator).
    Requires PowerShell 7+.
#>

function New-QlikXrf {
    -join ((48..57) + (97..102) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
}

function Get-QlikHost {
    <#
        Normalize whatever the user typed in the Server field to a bare host.
        Tolerates a pasted scheme (https://), an explicit :port, and trailing
        slashes/paths so they don't end up doubled into the request URL.
            https://srv.dom/  -> srv.dom
            srv.dom:4242       -> srv.dom
    #>
    param([Parameter(Mandatory)][string]$Server)
    $h = $Server.Trim()
    $h = $h -replace '^[a-zA-Z][a-zA-Z0-9+.-]*://', ''   # strip scheme
    $h = $h -replace '[/?#].*$', ''                       # strip path/query/fragment
    $h = $h -replace ':\d+$', ''                          # strip :port
    if ([string]::IsNullOrWhiteSpace($h)) { throw "Server name is empty after parsing '$Server'." }
    $h
}

function Get-QlikClientCertificate {
    <#
        Returns an X509Certificate2 usable for TLS client auth on Windows.
        Prefers client.pem + client_key.pem (QMC/cURL-style export); falls back
        to client.pfx. The PEM pair is round-tripped through PFX so the private
        key is properly associated for SChannel (a straight CreateFromPemFile
        cert often fails client auth with "private key not accessible").
    #>
    param(
        [Parameter(Mandatory)][string]$CertDir,
        [string]$PfxPassword = ''
    )
    $clientPem = Join-Path $CertDir 'client.pem'
    $clientKey = Join-Path $CertDir 'client_key.pem'
    $clientPfx = Join-Path $CertDir 'client.pfx'

    if ((Test-Path $clientPem) -and (Test-Path $clientKey)) {
        $tmp = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPemFile($clientPem, $clientKey)
        $pfxBytes = $tmp.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx)
        return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($pfxBytes)
    }
    if (Test-Path $clientPfx) {
        return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($clientPfx, $PfxPassword)
    }
    throw "Expected client.pem + client_key.pem (or client.pfx) in '$CertDir'."
}

function Invoke-QlikQrs {
    <# Single entry point for QRS calls; handles xrfkey, identity header, cert. #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$Path,                       # e.g. app/full
        [ValidateSet('Get','Delete','Post','Put')][string]$Method = 'Get',
        [string]$Filter,
        [Parameter(Mandatory)]$Certificate,
        [string]$AdminUserDirectory = 'INTERNAL',
        [string]$AdminUserId        = 'sa_repository'
    )
    $xrf = New-QlikXrf
    $qs  = "xrfkey=$xrf"
    if ($Filter) { $qs += '&filter=' + [uri]::EscapeDataString($Filter) }
    $hostName = Get-QlikHost -Server $Server
    $uri = 'https://{0}:4242/qrs/{1}?{2}' -f $hostName, $Path.TrimStart('/'), $qs
    $headers = @{
        'X-Qlik-Xrfkey' = $xrf
        'X-Qlik-User'   = "UserDirectory=$AdminUserDirectory; UserId=$AdminUserId"
        'Content-Type'  = 'application/json'
    }
    Invoke-RestMethod -Uri $uri -Headers $headers -Certificate $Certificate `
        -Method $Method -SkipCertificateCheck
}

function Expand-QlikRecords {
    <#
        Some PowerShell/Invoke-RestMethod combinations return a JSON array of
        records as a SINGLE object whose properties are parallel arrays
        (e.g. .id = @(g1,g2,g3,g4)). This detects that shape (signalled by a
        scalar field like 'id' arriving as an array) and rebuilds one record
        per element. A normal multi-record response, or a genuine single
        record, passes straight through untouched.
    #>
    param($Response)
    # Invoke-RestMethod returns $null when QRS matches zero records. Wrapping
    # $null with @() produces a 1-ELEMENT array containing $null (PowerShell's
    # array-subexpression operator treats $null as "one output object", not
    # "no output"), which would otherwise fall through and get expanded into a
    # single blank/garbage record below. Catch that here so "zero matches"
    # really means zero matches.
    if ($null -eq $Response) { return @() }
    $items = @($Response)
    if ($items.Count -eq 0) { return $items }
    if ($items.Count -ne 1 -or $null -eq $items[0]) { return $items }
    $obj = $items[0]
    if (-not ($obj.id -is [System.Array])) { return $items }   # genuine single record
    $n = @($obj.id).Count
    $names = $obj.PSObject.Properties.Name
    @(for ($i = 0; $i -lt $n; $i++) {
        $h = [ordered]@{}
        foreach ($nm in $names) {
            $v = $obj.$nm
            $h[$nm] = if (($v -is [System.Array]) -and ($v.Count -eq $n)) { $v[$i] } else { $v }
        }
        [pscustomobject]$h
    })
}

function Get-QlikApps {
    <# Returns all apps with their IDs for selection. #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)]$Certificate,
        [string]$AdminUserDirectory = 'INTERNAL',
        [string]$AdminUserId        = 'sa_repository'
    )
    $apps = Expand-QlikRecords (Invoke-QlikQrs -Server $Server -Path 'app/full' -Certificate $Certificate `
                -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId)
    foreach ($a in $apps) {
        [pscustomobject]@{
            Name       = [string]$a.name
            Id         = [string]$a.id
            Owner      = "$($a.owner.userDirectory)\$($a.owner.userId)"
            Stream     = if ($a.stream) { [string]$a.stream.name } else { '<unpublished>' }
            Published  = [bool]$a.published
            LastReload = [string]$a.lastReloadTime
        }
    }
}

function Get-QlikPrivateContent {
    <#
        Private sheets+bookmarks for one app. If $Days is greater than 0, only
        objects modified within that many days are returned; if $Days is 0 (or
        less), the date filter is dropped and ALL private sheets/bookmarks for
        the app are returned regardless of last-updated date.
    #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)]$Certificate,
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][int]$Days,
        [string]$AdminUserDirectory = 'INTERNAL',
        [string]$AdminUserId        = 'sa_repository'
    )
    $filter = "app.id eq $AppId " +
              "and published eq false " +
              "and (objectType eq 'sheet' or objectType eq 'bookmark')"
    if ($Days -gt 0) {
        $cutoff = (Get-Date).ToUniversalTime().AddDays(-1 * $Days).ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        $filter += " and modifiedDate ge '$cutoff'"
    }
    $objs = Expand-QlikRecords (Invoke-QlikQrs -Server $Server -Path 'app/object/full' -Filter $filter `
                -Certificate $Certificate -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId)
    foreach ($o in $objs) {
        [pscustomobject]@{
            appId              = $AppId
            qrsId              = [string]$o.id
            engineObjectId     = [string]$o.engineObjectId
            objectType         = [string]$o.objectType
            name               = [string]$o.name
            modifiedDate       = [string]$o.modifiedDate
            ownerUserDirectory = $o.owner.userDirectory
            ownerUserId        = $o.owner.userId
        }
    }
}

function Get-QlikUsers {
    <#
        All users known to the Repository, with the flags needed to tell a
        live account from a stale one:
          - removedExternally : the directory connector (AD/etc.) no longer
                                 returns this account, but QRS kept the record.
          - blacklisted       : disabled/blacklisted in QRS.
        Either flag means the account can no longer log in, so any private
        content it still owns is effectively unreachable by its owner.
    #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)]$Certificate,
        [string]$AdminUserDirectory = 'INTERNAL',
        [string]$AdminUserId        = 'sa_repository'
    )
    $users = Expand-QlikRecords (Invoke-QlikQrs -Server $Server -Path 'user/full' -Certificate $Certificate `
                -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId)
    foreach ($u in $users) {
        [pscustomobject]@{
            UserDirectory     = [string]$u.userDirectory
            UserId            = [string]$u.userId
            Name              = [string]$u.name
            RemovedExternally = [bool]$u.removedExternally
            Blacklisted       = [bool]$u.blacklisted
        }
    }
}

function Get-QlikOrphanedPrivateContent {
    <#
        Private sheets/bookmarks for one app whose owner can no longer act on
        them - either because the App.Object has no owner reference at all,
        or because the owner account has been removed from its directory or
        blacklisted in QRS. These never get cleaned up by Get-QlikPrivateContent
        date-window logic if the object happens to be old, and the person who
        'owns' them can no longer log in to delete them themselves.

        Always ignores any date window - an orphan is an orphan regardless of
        modifiedDate, so every private sheet/bookmark in the app is checked.

        Returns the same shape as Get-QlikPrivateContent plus an OrphanReason:
            'no-owner'          - object has no owner directory/id at all
            'owner-removed'     - owner no longer exists in the User repository,
                                   or was removed externally (deleted from the
                                   source directory)
            'owner-blacklisted' - owner exists but is blacklisted in QRS
    #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)]$Certificate,
        [Parameter(Mandatory)][string]$AppId,
        [string]$AdminUserDirectory = 'INTERNAL',
        [string]$AdminUserId        = 'sa_repository'
    )
    $allPrivate = Get-QlikPrivateContent -Server $Server -Certificate $Certificate -AppId $AppId -Days 0 `
                    -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId
    $users = Get-QlikUsers -Server $Server -Certificate $Certificate `
                -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId

    # Map "directory\id" -> user record, for live (non-removed, non-blacklisted)
    # accounts only. Anything not in here is no longer a usable owner.
    $liveUsers = @{}
    foreach ($u in $users) {
        if (-not $u.RemovedExternally -and -not $u.Blacklisted) {
            $liveUsers["$($u.UserDirectory)\$($u.UserId)"] = $u
        }
    }
    # Separate lookup so we can tell "missing entirely" apart from "exists but blacklisted".
    $anyUsers = @{}
    foreach ($u in $users) { $anyUsers["$($u.UserDirectory)\$($u.UserId)"] = $u }

    foreach ($o in $allPrivate) {
        $hasOwner = -not [string]::IsNullOrWhiteSpace($o.ownerUserId)
        $key      = "$($o.ownerUserDirectory)\$($o.ownerUserId)"

        $reason = $null
        if (-not $hasOwner) {
            $reason = 'no-owner'
        } elseif (-not $liveUsers.ContainsKey($key)) {
            $reason = if ($anyUsers.ContainsKey($key) -and $anyUsers[$key].Blacklisted) {
                'owner-blacklisted'
            } else {
                'owner-removed'
            }
        }

        if ($reason) {
            $o | Add-Member -NotePropertyName OrphanReason -NotePropertyValue $reason -PassThru
        }
    }
}

function Remove-QlikPrivateContentQrs {
    <# Owner-agnostic metadata deletion via QRS. #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)]$Certificate,
        [Parameter(Mandatory)]$Manifest,
        [string]$AdminUserDirectory = 'INTERNAL',
        [string]$AdminUserId        = 'sa_repository',
        [scriptblock]$Log = { param($m) Write-Host $m }
    )
    foreach ($m in $Manifest) {
        try {
            Invoke-QlikQrs -Server $Server -Path "app/object/$($m.qrsId)" -Method Delete `
                -Certificate $Certificate -AdminUserDirectory $AdminUserDirectory -AdminUserId $AdminUserId | Out-Null
            & $Log ("QRS deleted {0} '{1}'" -f $m.objectType, $m.name)
        } catch {
            $code = $null
            try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            if ($code -eq 404) {
                # Already gone - e.g. the engine delete ran first and the
                # engine->repository sync already removed this metadata row.
                & $Log ("QRS already removed (ok) {0} '{1}'" -f $m.objectType, $m.name)
            } else {
                & $Log ("QRS delete FAILED {0}: {1}" -f $m.qrsId, $_.Exception.Message)
            }
        }
    }
}

function Remove-QlikPrivateContentEngine {
    <# Engine API deletion via enigma.js (destroy-objects.cjs), streamed to $Log.
       If -OverrideUser is set (UserDirectory\UserId), ALL objects are deleted
       under that single identity instead of impersonating each owner. Use this
       for section-access apps, supplying a service account that has ADMIN in
       the section access table so the engine opens an unreduced session. #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$CertDir,
        [Parameter(Mandatory)][string]$ManifestPath,
        [Parameter(Mandatory)][string]$ResultsPath,
        [Parameter(Mandatory)][string]$DestroyScript,
        [string]$Schema = '12.612.0',
        [string]$OverrideUser = '',
        [scriptblock]$Log = { param($m) Write-Host $m }
    )
    $node = (Get-Command node -ErrorAction Stop).Source
    $args = @($DestroyScript, '--manifest', $ManifestPath, '--host', $Server,
              '--certs', $CertDir, '--schema', $Schema, '--results', $ResultsPath, '--execute')
    if (-not [string]::IsNullOrWhiteSpace($OverrideUser)) {
        $args += @('--override-user', $OverrideUser)
    }
    & $node @args 2>&1 | ForEach-Object { & $Log ([string]$_) }
}

Export-ModuleMember -Function New-QlikXrf, Get-QlikHost, Get-QlikClientCertificate, Invoke-QlikQrs,
    Get-QlikApps, Get-QlikPrivateContent, Get-QlikUsers, Get-QlikOrphanedPrivateContent,
    Remove-QlikPrivateContentQrs, Remove-QlikPrivateContentEngine
