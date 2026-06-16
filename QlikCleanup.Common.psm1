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
    $items = @($Response)
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
    <# Engine API deletion via enigma.js (destroy-objects.cjs), streamed to $Log. #>
    param(
        [Parameter(Mandatory)][string]$Server,
        [Parameter(Mandatory)][string]$CertDir,
        [Parameter(Mandatory)][string]$ManifestPath,
        [Parameter(Mandatory)][string]$ResultsPath,
        [Parameter(Mandatory)][string]$DestroyScript,
        [string]$Schema = '12.612.0',
        [scriptblock]$Log = { param($m) Write-Host $m }
    )
    $node = (Get-Command node -ErrorAction Stop).Source
    & $node $DestroyScript --manifest $ManifestPath --host $Server --certs $CertDir `
        --schema $Schema --results $ResultsPath --execute 2>&1 |
        ForEach-Object { & $Log ([string]$_) }
}

Export-ModuleMember -Function New-QlikXrf, Get-QlikHost, Get-QlikClientCertificate, Invoke-QlikQrs,
    Get-QlikApps, Get-QlikPrivateContent, Remove-QlikPrivateContentQrs, Remove-QlikPrivateContentEngine
