# Qlik-Sense-private-content-clean-up
An App with a UI to delete private content from Qlik Sense Apps


# Delete recent private content for a Qlik Sense Enterprise (Windows) app

Removes **private sheets and bookmarks created in the last N days (default 90)**
for one app, identified by app ID. Uses all three interfaces you asked for:

| Stage | Tool | What it does |
|-------|------|--------------|
| Auth/context | **Qlik CLI** | `qlik context` holds the connection + certs; used to verify reachability and as the alternative deletion path (below). |
| Enumerate | **Repository API (QRS)** | `GET /qrs/app/object/full` with a `createdDate` + `published eq false` filter returns the private sheets/bookmarks and their `engineObjectId` + owner. |
| Delete | **Engine API (QIX)** | `destroy-objects.cjs` (enigma.js) calls `DestroyObject` / `DestroyBookmark` per object, impersonating each owner. |

## Files
- `Show-CleanupUI.ps1` — **Windows Forms UI** (start here for interactive use).
- `QlikCleanup.Common.psm1` — shared QRS/engine functions used by the UI.
- `Remove-AppPrivateContent.ps1` — command-line orchestrator (scripting/scheduled use).
- `destroy-objects.cjs` — Engine API deletion stage (enigma.js), shared by both.

## Windows Forms UI
```powershell
pwsh -File .\Show-CleanupUI.ps1 -Server qlik.example.com -CertDir C:\qlik-certs
```
Workflow in the window:
1. **Load apps** — pulls every app (name + ID) from the Repository API into the list. Type in the filter box to narrow by name or paste an app ID.
2. **Select an app**, then pick a timeframe radio: **30 / 60 / 90 / 120 days**.
3. **Preview (dry run)** — lists the private sheets/bookmarks that match, with owner and created date. Nothing is changed.
4. **Delete selected content** — enabled only after a preview returns results; asks for confirmation, then deletes via the **Engine API** (default) or **QRS only**, streams progress to the log, and re-checks that 0 remain.

Each delete writes a manifest, a metadata backup, and a results file to `.\run\`.
WinForms needs an STA thread; PowerShell 7 runs STA by default.

## Prerequisites
1. **Qlik CLI** installed and a cert-based context:
   ```
   qlik context create qseow --server https://<server> --certificates C:\qlik-certs --insecure
   qlik context use qseow
   ```
2. **Client certificate** exported from the QMC (Certificate export → PEM,
   *Include secret key*). Put `root.pem`, `client.pem`, `client_key.pem` in `C:\qlik-certs`.
3. **PowerShell 7+** (for `-Certificate` + `-SkipCertificateCheck`).
4. **Node.js + enigma.js** for the engine stage:
   ```
   npm install enigma.js ws
   ```
   If schema `12.612.0` isn't present, list `node_modules/enigma.js/schemas`
   and pass an available one with `-EngineSchema`.
5. The acting QRS identity (`INTERNAL\sa_repository` by default) must be a
   **RootAdmin/ContentAdmin**.

## Run
```powershell
# 1) Preview (DRY RUN — deletes nothing)
./Remove-AppPrivateContent.ps1 -AppId <app-guid> -Server qlik.example.com -CertDir C:\qlik-certs

# 2) Delete via the Engine API
./Remove-AppPrivateContent.ps1 -AppId <app-guid> -Execute -Method Engine

# Alternative: delete only the QRS metadata (reconciled on next reload)
./Remove-AppPrivateContent.ps1 -AppId <app-guid> -Execute -Method Qrs
```
A timestamped manifest, full metadata backup, and log are written to `.\run\`.

## Why Engine API for the delete (and not just QRS)?
A QRS `DELETE` removes only the **metadata row**; the object can survive inside
the app binary until the next reload. `DestroyObject`/`DestroyBookmark` over QIX
removes it at the engine level, and the QRS metadata is then cleaned up by the
engine→repository sync. That's why the default `-Method Engine` is recommended;
`-Method Qrs` is the lighter, owner-agnostic fallback.

## The owner-impersonation detail (important)
Private objects are **owner-scoped in the engine** — you only see/destroy a
user's private sheet if connected *as that user*. `destroy-objects.cjs` connects
directly to the engine (`wss://<server>:4747`) with the **client certificate**
and sets `X-Qlik-User: UserDirectory=<dir>; UserId=<id>` per owner, opening one
session per owner. Cert auth provides the admin rights; `X-Qlik-User` scopes the
session to the owner so the destroy succeeds.

### Pure-CLI alternative for the delete stage
If you prefer to stay entirely in qlik-cli instead of Node, the engine-backed
object/bookmark commands do the same thing per object:
```
qlik app object   rm <engineObjectId> -a <appId> --headers "X-Qlik-User=UserDirectory=<dir>; UserId=<id>"
qlik app bookmark rm <engineObjectId> -a <appId> --headers "X-Qlik-User=UserDirectory=<dir>; UserId=<id>"
```
Loop the manifest and call these. (Confirm your build honors the per-call
`--headers` identity override against the engine; enigma.js gives the most
deterministic control, which is why it's the default.)

## Caveats
- **Test with `-Method Qrs` / dry run on a non-prod app first.** Deletion is
  irreversible; the metadata backup in `.\run\` is your only record.
- `published eq false` targets *private* objects. Community-published or
  approved (base) sheets are excluded by design.
- Only `objectType` `sheet` and `bookmark` are targeted — not master items,
  stories, snapshots, etc.
- `doSave()` in the engine stage may be blocked for governed/published apps;
  the destroy still applies to the session, and the repository sync reconciles.
- If `qlik` can't reach a host with self-signed certs, the `--insecure` flag and
  the script's `-SkipCertificateCheck` / `rejectUnauthorized:false` handle it.
  Tighten these once you have a proper CA chain.
