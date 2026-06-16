<#
.SYNOPSIS
    Delete private sheets and bookmarks created in the last N days for one
    Qlik Sense Enterprise on Windows app.

.DESCRIPTION
    Three stages, three tools:

      1. Qlik CLI        - establishes/uses the auth context (qlik context ...).
      2. Repository API  - enumerates the private App.Objects (sheets/bookmarks)
                           for the app, filtered by modifiedDate, via
                           GET /qrs/app/object/full  (client-certificate auth).
      3. Engine API      - destroys each object in the engine via enigma.js
                           (destroy-objects.cjs), impersonating each owner.

    The script is a DRY RUN by default. Add -Execute to actually delete.
    Nothing is destroyed without -Execute.

.NOTES
    Requires PowerShell 7+ (uses -SkipCertificateCheck / -Certificate).
    Run from an account/host that has the exported Qlik client certificate.
    The acting QRS identity must be a RootAdmin or ContentAdmin.

.EXAMPLE
    # Preview only
    ./Remove-AppPrivateContent.ps1 -AppId 11111111-2222-3333-4444-555555555555

.EXAMPLE
    # Actually delete via the Engine API
    ./Remove-AppPrivateContent.ps1 -AppId <guid> -Execute -Method Engine
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $AppId,

    # Look-back window. "last 90 days" => objects with modifiedDate >= now-90d.
    # Use 0 to delete ALL private sheets/bookmarks regardless of date.
    [int] $Days = 90,

    # Qlik proxy/host DNS name (no scheme).
    [string] $Server = 'qlik.example.com',

    # Folder holding the QMC-exported PEM certs: root.pem, client.pem, client_key.pem
    [string] $CertDir = 'C:\qlik-certs',

    # QRS identity used by certificate auth.
    [string] $AdminUserDirectory = 'INTERNAL',
    [string] $AdminUserId        = 'sa_repository',

    # Engine | Qrs | Both   (how to perform the deletion; Both runs engine then QRS)
    [ValidateSet('Engine','Qrs','Both')] [string] $Method = 'Both',

    # enigma schema version available under node_modules/enigma.js/schemas
    [string] $EngineSchema = '12.612.0',

    # Working folder for the manifest / backup / logs.
    [string] $OutDir = (Join-Path $PSScriptRoot 'run'),

    # Safety switch. Without it, nothing is deleted.
    [switch] $Execute
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$stamp        = Get-Date -Format 'yyyyMMdd-HHmmss'
$manifestPath = Join-Path $OutDir "objects-to-delete-$stamp.json"
$backupPath   = Join-Path $OutDir "metadata-backup-$stamp.json"
$logPath      = Join-Path $OutDir "cleanup-$stamp.log"

function Write-Log {
    param([string]$Message)
    $line = "{0}  {1}" -f (Get-Date -Format 's'), $Message
    $line | Tee-Object -FilePath $logPath -Append
}

# --------------------------------------------------------------------------
# Stage 0: Qlik CLI context
# --------------------------------------------------------------------------
# Create this once (cert-based) and the script will reuse it:
#
#   qlik context create qseow `
#       --server https://$Server `
#       --certificates $CertDir `
#       --comment "QSEoW cleanup" --insecure
#   qlik context use qseow
#
# Verify the CLI can reach the engine before deleting anything:
Write-Log "Verifying Qlik CLI context against $Server ..."
try {
    & qlik context get | Out-Null
    Write-Log "qlik-cli context OK."
} catch {
    Write-Log "WARNING: 'qlik context get' failed ($($_.Exception.Message)). Continuing with direct REST."
}

# --------------------------------------------------------------------------
# Stage 1: Repository API (QRS) - enumerate private sheets & bookmarks
# --------------------------------------------------------------------------
# Equivalent CLI forms (the dedicated QRS group handles the xrfkey for you):
#   qlik qrs --help                      # list available QRS commands
#   qlik raw get /qrs/app/object/full --query "filter=<filter>" `
#        --query "xrfkey=0123456789abcdef" --header "X-Qlik-Xrfkey: 0123456789abcdef"
# We use Invoke-RestMethod here so the date filter and parsing are explicit.

# QRS OData-ish filter. App id is unquoted; string/date literals single-quoted.
$filter = "app.id eq $AppId " +
          "and published eq false " +
          "and (objectType eq 'sheet' or objectType eq 'bookmark')"
if ($Days -gt 0) {
    $cutoffUtc = (Get-Date).ToUniversalTime().AddDays(-1 * $Days).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    Write-Log "Cutoff (UTC): modifiedDate >= $cutoffUtc"
    $filter += " and modifiedDate ge '$cutoffUtc'"
} else {
    Write-Log "No date cutoff (-Days 0): targeting ALL private sheets/bookmarks."
}

$xrf = -join ((48..57) + (97..102) | Get-Random -Count 16 | ForEach-Object {[char]$_})

$clientPfxOrPem = Join-Path $CertDir 'client.pem'
$clientKey      = Join-Path $CertDir 'client_key.pem'

# Load the client certificate (PEM cert + key). PowerShell 7.
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPemFile(
            $clientPfxOrPem, $clientKey)

$qrsBase = "https://{0}:4242/qrs" -f $Server
$uri = "{0}/app/object/full?filter={1}&xrfkey={2}" -f $qrsBase,
        [uri]::EscapeDataString($filter), $xrf

$headers = @{
    'X-Qlik-Xrfkey' = $xrf
    'X-Qlik-User'   = "UserDirectory=$AdminUserDirectory; UserId=$AdminUserId"
    'Content-Type'  = 'application/json'
}

Write-Log "Querying QRS: $qrsBase/app/object/full (filter applied)"
$objects = Invoke-RestMethod -Uri $uri -Headers $headers -Certificate $cert `
            -Method Get -SkipCertificateCheck

if (-not $objects) { $objects = @() }
$objects = @($objects)   # force array even for single result
Write-Log ("QRS returned {0} private object(s) in window." -f $objects.Count)

if ($objects.Count -eq 0) {
    Write-Log "Nothing to delete. Exiting."
    return
}

# Normalize to the manifest the engine stage consumes, and a human report.
$manifest = foreach ($o in $objects) {
    [pscustomobject]@{
        appId               = $AppId
        qrsId               = $o.id
        engineObjectId      = $o.engineObjectId
        objectType          = $o.objectType
        name                = $o.name
        modifiedDate        = $o.modifiedDate
        ownerUserDirectory  = $o.owner.userDirectory
        ownerUserId         = $o.owner.userId
    }
}

# Save a metadata backup (so you have a record of what existed) and the manifest.
$objects  | ConvertTo-Json -Depth 10 | Set-Content -Path $backupPath
$manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath
Write-Log "Manifest -> $manifestPath"
Write-Log "Full metadata backup -> $backupPath"

Write-Host ""
Write-Host "Private content to remove (app $AppId, last $Days days):" -ForegroundColor Cyan
$manifest |
    Select-Object objectType, name,
        @{n='owner';e={"$($_.ownerUserDirectory)\$($_.ownerUserId)"}},
        modifiedDate, engineObjectId |
    Sort-Object objectType, owner |
    Format-Table -AutoSize

if (-not $Execute) {
    Write-Host ""
    Write-Host "DRY RUN. Re-run with -Execute to delete the $($manifest.Count) object(s) above." -ForegroundColor Yellow
    Write-Log  "DRY RUN complete. No changes made."
    return
}

# --------------------------------------------------------------------------
# Stage 2: Deletion  (Engine first, then QRS - both work from the manifest)
# --------------------------------------------------------------------------
if ($Method -in 'Engine','Both') {
    # Engine API via enigma.js. Each owner is impersonated by destroy-objects.cjs.
    $resultsPath = Join-Path $OutDir "engine-delete-results-$stamp.json"
    $node = (Get-Command node).Source
    $script = Join-Path $PSScriptRoot 'destroy-objects.cjs'

    Write-Log "Engine deletion via enigma.js ($script)"
    & $node $script `
        --manifest $manifestPath `
        --host     $Server `
        --certs    $CertDir `
        --schema   $EngineSchema `
        --results  $resultsPath `
        --execute
    Write-Log "Engine stage exit code: $LASTEXITCODE. Results -> $resultsPath"
}
if ($Method -in 'Qrs','Both') {
    # Repository API deletion (metadata). Owner-agnostic; admin cert can delete
    # any user's App.Object. A 404 means the engine sync already removed it.
    Write-Log "QRS deletion of $($manifest.Count) object(s)."
    foreach ($m in $manifest) {
        $delXrf = -join ((48..57)+(97..102) | Get-Random -Count 16 | ForEach-Object {[char]$_})
        $delUri = "{0}/app/object/{1}?xrfkey={2}" -f $qrsBase, $m.qrsId, $delXrf
        $delHeaders = @{
            'X-Qlik-Xrfkey' = $delXrf
            'X-Qlik-User'   = "UserDirectory=$AdminUserDirectory; UserId=$AdminUserId"
        }
        try {
            Invoke-RestMethod -Uri $delUri -Headers $delHeaders -Certificate $cert `
                -Method Delete -SkipCertificateCheck | Out-Null
            Write-Log ("QRS deleted {0} {1} ({2})" -f $m.objectType, $m.name, $m.qrsId)
        } catch {
            $code = $null
            try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            if ($code -eq 404) {
                Write-Log ("QRS already removed (ok) {0} {1}" -f $m.objectType, $m.name)
            } else {
                Write-Log ("QRS delete FAILED {0}: {1}" -f $m.qrsId, $_.Exception.Message)
            }
        }
    }
}

# --------------------------------------------------------------------------
# Stage 3: Verify
# --------------------------------------------------------------------------
$verXrf = -join ((48..57)+(97..102) | Get-Random -Count 16 | ForEach-Object {[char]$_})
$verUri = "{0}/app/object/full?filter={1}&xrfkey={2}" -f $qrsBase,
           [uri]::EscapeDataString($filter), $verXrf
$verHeaders = @{ 'X-Qlik-Xrfkey' = $verXrf; 'X-Qlik-User' = "UserDirectory=$AdminUserDirectory; UserId=$AdminUserId" }
Start-Sleep -Seconds 3   # allow engine->repository sync
$remaining = @(Invoke-RestMethod -Uri $verUri -Headers $verHeaders -Certificate $cert -Method Get -SkipCertificateCheck)
Write-Log ("Verification: {0} matching private object(s) remain." -f $remaining.Count)
Write-Host ""
Write-Host ("Done. {0} object(s) still match the filter (expect 0)." -f $remaining.Count) -ForegroundColor Green
