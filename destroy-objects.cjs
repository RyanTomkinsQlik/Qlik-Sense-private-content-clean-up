#!/usr/bin/env node
/*
 * destroy-objects.cjs
 * --------------------
 * Engine API (QIX) deletion stage of the private-content cleanup.
 *
 * Reads the JSON manifest produced by Remove-AppPrivateContent.ps1 (the QRS
 * enumeration stage) and destroys each private sheet / bookmark IN THE ENGINE.
 *
 * Why the Engine API and not just a QRS DELETE?
 *   - A QRS DELETE removes only the App.Object *metadata* row. The object can
 *     linger inside the app's binary until the next reload.
 *   - DestroyObject / DestroyBookmark over QIX removes it at the engine level
 *     and the QRS metadata is then cleaned up by the engine->repository sync.
 *
 * Private objects are owner-scoped in the engine: you can only see/destroy a
 * user's private sheet if you are connected AS that user. We therefore connect
 * directly to the engine on port 4747 with the Qlik server CLIENT CERTIFICATE
 * and set the X-Qlik-User header to impersonate each object's owner. Objects
 * are grouped by owner so we open one session per owner.
 *
 * Usage:
 *   node destroy-objects.cjs --manifest objects-to-delete.json \
 *        --host qlik.example.com \
 *        --certs C:\\qlik-certs \
 *        [--schema 12.612.0] [--execute] [--results results.json]
 *
 * Without --execute the script is a DRY RUN (prints what it would destroy).
 */

const fs = require('fs');
const path = require('path');
const enigma = require('enigma.js');
const WebSocket = require('ws');

// ----------------------------- arg parsing --------------------------------
function getArg(name, fallback = undefined) {
  const i = process.argv.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const next = process.argv[i + 1];
  if (next === undefined || next.startsWith('--')) return true; // boolean flag
  return next;
}

const manifestPath = getArg('manifest', 'objects-to-delete.json');
const host = getArg('host');
const certDir = getArg('certs');
const schemaVer = getArg('schema', '12.612.0');
const execute = getArg('execute', false) === true;
const resultsPath = getArg('results', 'engine-delete-results.json');
const enginePort = getArg('port', '4747');

if (!host || !certDir) {
  console.error('ERROR: --host and --certs are required.');
  process.exit(2);
}

// enigma ships several QIX schemas under enigma.js/schemas/. Pick one that is
// present in your installed version if 12.612.0 is missing.
let schema;
try {
  schema = require(`enigma.js/schemas/${schemaVer}.json`);
} catch (e) {
  console.error(`ERROR: could not load enigma schema ${schemaVer}. ` +
    `List node_modules/enigma.js/schemas and pass an available one with --schema.`);
  process.exit(2);
}

// --------------------------- certificate load -----------------------------
function readCert(file) {
  const p = path.join(certDir, file);
  if (!fs.existsSync(p)) {
    console.error(`ERROR: certificate file not found: ${p}`);
    process.exit(2);
  }
  return fs.readFileSync(p);
}
// Exported from the QMC (Certificate export, PEM, "Include secret key").
const ca = readCert('root.pem');
const cert = readCert('client.pem');
const key = readCert('client_key.pem');

// ------------------------------ manifest ----------------------------------
let manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
// PowerShell's ConvertTo-Json collapses a single-item array into a bare {...}
// unless the caller passed -AsArray. Tolerate that here instead of silently
// treating a lone object (the common shape for a single orphan) as "empty".
if (manifest && !Array.isArray(manifest)) manifest = [manifest];
if (!Array.isArray(manifest) || manifest.length === 0) {
  console.log('Manifest is empty - nothing to do.');
  fs.writeFileSync(resultsPath, JSON.stringify([], null, 2));
  process.exit(0);
}

const appId = manifest[0].appId;
if (!manifest.every(o => o.appId === appId)) {
  console.error('ERROR: manifest mixes multiple appIds; expected a single app.');
  process.exit(2);
}

// Orphans with no owner at all (deleted user, no owner reference left in QRS)
// can't be impersonated in the engine - there's no identity to open the
// session as. They're carried straight into the results as "skipped" so the
// QRS metadata-delete stage downstream still picks them up.
const noOwner = manifest.filter(o => !o.ownerUserId || !o.ownerUserDirectory);
const withOwner = manifest.filter(o => o.ownerUserId && o.ownerUserDirectory);

if (noOwner.length > 0) {
  console.log(`Objects with no owner reference (cannot impersonate, engine delete skipped): ${noOwner.length}`);
  for (const o of noOwner) {
    console.log(`  [no-owner] ${o.objectType} ${o.engineObjectId} "${o.name || ''}" (orphan: ${o.OrphanReason || 'no-owner'})`);
  }
}

// Group objects by owner (UserDirectory + UserId).
const byOwner = new Map();
for (const o of withOwner) {
  const k = `${o.ownerUserDirectory}\\${o.ownerUserId}`;
  if (!byOwner.has(k)) byOwner.set(k, []);
  byOwner.get(k).push(o);
}

console.log(`App:        ${appId}`);
console.log(`Objects:    ${manifest.length}  (${withOwner.length} across ${byOwner.size} owner(s), ${noOwner.length} ownerless orphan(s))`);
console.log(`Mode:       ${execute ? 'EXECUTE (destructive)' : 'DRY RUN'}`);
console.log('');

// --------------------------- engine session -------------------------------
function openSessionAs(userDirectory, userId) {
  const session = enigma.create({
    schema,
    url: `wss://${host}:${enginePort}/app/${encodeURIComponent(appId)}/identity/cleanup-${Date.now()}`,
    createSocket: (url) =>
      new WebSocket(url, {
        ca: [ca],
        cert,
        key,
        headers: {
          // Impersonate the object owner. Cert auth grants admin rights, and
          // X-Qlik-User narrows the session to this user's security context.
          'X-Qlik-User': `UserDirectory=${userDirectory}; UserId=${userId}`,
        },
        // Self-signed Qlik certs: we validate against the exported root above.
        // Hostname on the cert is usually the machine name, not the DNS name,
        // so identity verification is disabled here. Tighten if you have a
        // proper SAN/CA chain.
        rejectUnauthorized: false,
      }),
  });
  return session;
}

async function run() {
  const results = [];

  for (const o of noOwner) {
    console.log(`[skip] ${o.objectType} ${o.engineObjectId} "${o.name || ''}" - no owner to impersonate; QRS-only delete required`);
    results.push({ ...o, action: 'skip-no-owner', success: null });
  }

  for (const [ownerKey, objects] of byOwner) {
    const [userDirectory, userId] = ownerKey.split('\\');
    console.log(`Owner ${ownerKey}: ${objects.length} object(s)`);

    if (!execute) {
      for (const o of objects) {
        console.log(`  [dry-run] would destroy ${o.objectType} ${o.engineObjectId} "${o.name || ''}"`);
        results.push({ ...o, action: 'dry-run', success: null });
      }
      continue;
    }

    let session;
    let doc;
    try {
      session = openSessionAs(userDirectory, userId);
      const global = await session.open();
      doc = await global.openDoc(appId);
    } catch (err) {
      console.error(`  ! could not open app as ${ownerKey}: ${err.message}`);
      for (const o of objects) results.push({ ...o, action: 'destroy', success: false, error: `open failed: ${err.message}` });
      try { if (session) await session.close(); } catch (_) {}
      continue;
    }

    for (const o of objects) {
      try {
        let ok;
        if (o.objectType === 'bookmark') {
          ok = await doc.destroyBookmark(o.engineObjectId);
        } else {
          // sheets and any other generic object
          ok = await doc.destroyObject(o.engineObjectId);
        }
        console.log(`  ${ok ? 'destroyed' : 'not found '} ${o.objectType} ${o.engineObjectId} "${o.name || ''}"`);
        results.push({ ...o, action: 'destroy', success: !!ok });
      } catch (err) {
        console.error(`  ! failed ${o.objectType} ${o.engineObjectId}: ${err.message}`);
        results.push({ ...o, action: 'destroy', success: false, error: err.message });
      }
    }

    // Persist the cleared state so the deletions survive in the binary.
    try {
      await doc.doSave();
    } catch (err) {
      console.error(`  ! doSave failed for ${ownerKey}: ${err.message} (app may be governed/published)`);
    }
    try { await session.close(); } catch (_) {}
  }

  fs.writeFileSync(resultsPath, JSON.stringify(results, null, 2));
  const destroyed = results.filter(r => r.success === true).length;
  const failed = results.filter(r => r.success === false).length;
  console.log('');
  console.log(`Done. destroyed=${destroyed} failed=${failed} (results -> ${resultsPath})`);
  if (failed > 0) process.exitCode = 1;
}

run().catch(err => {
  console.error('Fatal:', err);
  process.exit(1);
});
