require('dotenv').config();
const path = require('path');
const express = require('express');
const cors = require('cors');
require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const xlsx = require('xlsx');

// ---------- env validation ----------
const REQUIRED_ENV = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'TARGET_USER_EMAIL'];
const missing = REQUIRED_ENV.filter(k => !process.env[k]);
if (missing.length) {
    console.error('❌ Missing required env vars:', missing.join(', '));
    process.exit(1);
}
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// 1. Microsoft Graph Authentication
const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
});
const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

const MASTER_FILE_PATH = ':/Metfraa_Costing_App/Master_Data/Master_Cost_DB.xlsx';

// Helper: sanitize filename for OneDrive
function sanitizeName(name) {
    return String(name || 'Untitled').replace(/[^a-zA-Z0-9_-]/g, '_').replace(/_+/g, '_');
}

// Helper: download master workbook as Buffer
async function downloadMasterBuffer() {
    const targetEmail = process.env.TARGET_USER_EMAIL;
    const fileMeta = await graphClient
        .api(`/users/${targetEmail}/drive/root${MASTER_FILE_PATH}`)
        .get();
    const downloadUrl = fileMeta['@microsoft.graph.downloadUrl'];
    const response = await fetch(downloadUrl);
    return Buffer.from(await response.arrayBuffer());
}

// Helper: pick the data sheet (skip README if present)
function pickDataSheet(workbook) {
    let sheetName = workbook.SheetNames[0];
    if (sheetName.toUpperCase() === 'README' && workbook.SheetNames.length > 1) {
        sheetName = workbook.SheetNames[1];
    }
    return sheetName;
}

// =============================================================
// BOOTSTRAP: default rows for Distance + Paint
// Auto-written to Excel on first /api/master-data call if missing
// =============================================================
const BOOTSTRAP_DISTANCE = [
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST001','Option Name':'0-100 km',  'Rate':1500,'Type':'RATE','Unit':'MT','Min':0,  'Max':100 },
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST002','Option Name':'101-200 km','Rate':2500,'Type':'RATE','Unit':'MT','Min':101,'Max':200 },
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST003','Option Name':'201-300 km','Rate':3500,'Type':'RATE','Unit':'MT','Min':201,'Max':300 },
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST004','Option Name':'301-400 km','Rate':4500,'Type':'RATE','Unit':'MT','Min':301,'Max':400 },
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST005','Option Name':'401-500 km','Rate':5000,'Type':'RATE','Unit':'MT','Min':401,'Max':500 },
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST006','Option Name':'501-600 km','Rate':5500,'Type':'RATE','Unit':'MT','Min':501,'Max':600 },
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST007','Option Name':'601-700 km','Rate':6000,'Type':'RATE','Unit':'MT','Min':601,'Max':700 },
    { 'Parameter Name':'Distance','Subcategory':'Distance','Option ID':'DIST008','Option Name':'701-800 km','Rate':6500,'Type':'RATE','Unit':'MT','Min':701,'Max':800 }
];
const BOOTSTRAP_PAINT = [
    { 'Parameter Name':'Paint','Subcategory':'Paint Type','Option ID':'PAINT001','Option Name':'Primer + Enamel','Rate':3800,'Type':'RATE','Unit':'MT','Min':'','Max':'' },
    { 'Parameter Name':'Paint','Subcategory':'Paint Type','Option ID':'PAINT002','Option Name':'Primer + Epoxy', 'Rate':4700,'Type':'RATE','Unit':'MT','Min':'','Max':'' },
    { 'Parameter Name':'Paint','Subcategory':'Paint Type','Option ID':'PAINT003','Option Name':'PU',             'Rate':0,   'Type':'RATE','Unit':'MT','Min':'','Max':'' }
];

// Required columns in the canonical order. If any are missing from the sheet,
// they get added to the header on bootstrap.
const REQUIRED_COLUMNS = ['Parameter Name','Subcategory','Option ID','Option Name','Rate','Type','Unit','Min','Max','Group','Remark','Rounding Unit'];

// Write rows back to OneDrive, preserving column order.
// Uses raw fetch (not Graph SDK) so we have full control over Content-Type — the SDK
// has been observed to override headers for binary uploads in some versions.
async function writeWorkbookToOneDrive(workbook) {
    const targetEmail = process.env.TARGET_USER_EMAIL;
    const outBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // Get a fresh access token from the credential
    const tokenResp = await credential.getToken('https://graph.microsoft.com/.default');
    if (!tokenResp || !tokenResp.token) {
        throw new Error('Failed to acquire access token for Graph API');
    }

    // Graph path syntax for path-addressed items: /drive/root:/<path>:/content
    // MASTER_FILE_PATH already starts with ':/' so we just need to add ':' before '/content'.
    const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(targetEmail)}/drive/root${MASTER_FILE_PATH}:/content`;

    const resp = await fetch(url, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${tokenResp.token}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Length': String(outBuffer.length)
        },
        body: outBuffer
    });

    if (!resp.ok) {
        const text = await resp.text();
        console.error('⚠️  writeWorkbookToOneDrive failed:', resp.status, resp.statusText);
        console.error('   URL was:', url);
        console.error('   Response body:', text.substring(0, 500));
        throw new Error(`Graph PUT failed: ${resp.status} ${resp.statusText} — ${text.substring(0, 200)}`);
    }
}

// Returns true if the workbook was modified
async function bootstrapMissingDefaults(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(sheet);
    const headerRow = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);

    // Match numbered or unnumbered param names case-insensitively
    // (e.g. "Distance", "15. DISTANCE", "16. PAINT" all count as the same conceptual param)
    function hasParam(needle) {
        const target = needle.trim().toLowerCase();
        return rawData.some(r => {
            const name = String(r['Parameter Name'] || '').trim().toLowerCase();
            // Strip leading "N. " prefix to compare
            const stripped = name.replace(/^\d+\.\s*/, '');
            return stripped === target || name === target;
        });
    }

    let changed = false;
    if (!hasParam('distance')) {
        rawData.push(...BOOTSTRAP_DISTANCE);
        changed = true;
        console.log('🔧 Bootstrapping: added 8 Distance band rows');
    }
    if (!hasParam('paint')) {
        rawData.push(...BOOTSTRAP_PAINT);
        changed = true;
        console.log('🔧 Bootstrapping: added 3 Paint type rows');
    }

    // Make sure all required columns are in the header (Min/Max especially)
    let header = headerRow.slice();
    REQUIRED_COLUMNS.forEach(col => {
        if (!header.includes(col)) {
            header.push(col);
            changed = true;
        }
    });

    if (changed) {
        const newSheet = xlsx.utils.json_to_sheet(rawData, { header });
        workbook.Sheets[sheetName] = newSheet;
    }

    return changed;
}

// Helper: simple admin gate (checks header OR body field)
function isAdmin(req) {
    if (!ADMIN_PASSWORD) return true; // open mode
    const provided = req.headers['x-admin-password'] || (req.body && req.body.adminPassword) || '';
    return provided === ADMIN_PASSWORD;
}

// =============================================================
// AUTH MODE — frontend uses this to know whether to prompt
// =============================================================
app.get('/api/auth-mode', (req, res) => {
    res.json({ adminProtected: !!ADMIN_PASSWORD });
});

app.post('/api/admin-check', (req, res) => {
    if (!ADMIN_PASSWORD) return res.json({ ok: true });
    res.json({ ok: req.body && req.body.password === ADMIN_PASSWORD });
});

// =============================================================
// 1. ROUTE: Fetch Master Data from OneDrive
// =============================================================
app.get('/api/master-data', async (req, res) => {
    try {
        // No-cache so the Refresh button always pulls latest
        res.set('Cache-Control', 'no-store, no-cache, must-revalidate');
        res.set('Pragma', 'no-cache');
        res.set('Expires', '0');

        let buffer = await downloadMasterBuffer();
        let workbook = xlsx.read(buffer, { type: 'buffer' });
        let sheetName = pickDataSheet(workbook);

        // Bootstrap defaults (Distance, Paint) if missing.
        // If write-back fails, we still serve the bootstrapped data from memory
        // so the app stays usable, and just retry on the next refresh.
        const bootstrapped = await bootstrapMissingDefaults(workbook, sheetName);
        if (bootstrapped) {
            try {
                await writeWorkbookToOneDrive(workbook);
                console.log('✅ Bootstrap rows persisted to OneDrive');
                // Re-download so subsequent reads see the persisted version
                buffer = await downloadMasterBuffer();
                workbook = xlsx.read(buffer, { type: 'buffer' });
                sheetName = pickDataSheet(workbook);
            } catch (writeErr) {
                console.error('⚠️  Bootstrap write-back failed (will retry next request):', writeErr.message);
                // Continue with the in-memory bootstrapped workbook — app stays usable
            }
        }

        const rawData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        // Hide Secondary parameter as a safety net — Excel may still have it
        const HIDDEN_PARAMETERS = new Set(['secondary']);

        const formattedData = {};
        rawData.forEach(row => {
            const paramName = row['Parameter Name'];
            const subcat = row['Subcategory'];
            if (!paramName || !subcat) return;
            if (HIDDEN_PARAMETERS.has(String(paramName).trim().toLowerCase())) return;

            if (!formattedData[paramName]) {
                formattedData[paramName] = { __unit: row['Unit'] || 'MT', subs: {} };
            }
            if (!formattedData[paramName].__unit && row['Unit']) {
                formattedData[paramName].__unit = row['Unit'];
            }
            // Rounding Unit: param-level metadata (10 or 100; blank for globals).
            // Pick up the first non-blank value seen for the parameter.
            if (formattedData[paramName].__roundingUnit == null && row['Rounding Unit'] !== undefined && row['Rounding Unit'] !== '') {
                const ru = Number(row['Rounding Unit']);
                if (ru === 10 || ru === 100) formattedData[paramName].__roundingUnit = ru;
            }
            if (!formattedData[paramName].subs[subcat]) formattedData[paramName].subs[subcat] = [];

            const rate = row['Rate'] !== undefined ? row['Rate']
                       : (row['Rate (₹)'] !== undefined ? row['Rate (₹)']
                       : (row['Price'] !== undefined ? row['Price'] : 0));

            const type = (row['Type'] || 'RATE').toString().toUpperCase().trim();

            let normType;
            if (type === 'AUTO') normType = 'AUTO';
            else if (type === 'PERCENT' || type === 'PERCENTAGE' || type === '%') normType = 'PERCENT';
            else normType = 'RATE';

            const opt = {
                id: row['Option ID'],
                name: row['Option Name'],
                rate: parseFloat(rate) || 0,
                type: normType,
                group: (row['Group'] !== undefined && row['Group'] !== null) ? String(row['Group']).trim() : '',
                remark: (row['Remark'] !== undefined && row['Remark'] !== null) ? String(row['Remark']).trim() : ''
            };

            // Optional Min/Max columns for banded options (e.g. Distance bands)
            if (row['Min'] !== undefined && row['Min'] !== '') opt.min = Number(row['Min']);
            if (row['Max'] !== undefined && row['Max'] !== '') opt.max = Number(row['Max']);

            formattedData[paramName].subs[subcat].push(opt);
        });

        res.json(formattedData);
    } catch (error) {
        console.error('Master Data Error:', error.message);
        res.status(500).json({ error: error.message });
    }
});

// =============================================================
// 2. ROUTE: Save Quotation (PDF + JSON) to OneDrive
// =============================================================
app.post('/api/save-quotation', async (req, res) => {
    try {
        const { projectName, clientName, date, pdfBase64, jsonData } = req.body;
        if (!projectName || !date) {
            return res.status(400).json({ success: false, error: 'Missing projectName or date' });
        }
        if (!pdfBase64 && !jsonData) {
            return res.status(400).json({ success: false, error: 'Provide pdfBase64 or jsonData (or both)' });
        }

        const targetEmail = process.env.TARGET_USER_EMAIL;
        // Base filename = client_project_date (client first; project + date follow)
        const fileParts = [];
        if (clientName && String(clientName).trim() && String(clientName).trim() !== '—') {
            fileParts.push(sanitizeName(clientName));
        }
        fileParts.push(sanitizeName(projectName));
        fileParts.push(sanitizeName(date));
        const baseName = fileParts.join('_');

        // Per-project folder name = client_project (no date) — groups all revisions of same project
        const projectFolderParts = [];
        if (clientName && String(clientName).trim() && String(clientName).trim() !== '—') {
            projectFolderParts.push(sanitizeName(clientName));
        }
        projectFolderParts.push(sanitizeName(projectName));
        const projectFolder = projectFolderParts.join('_');

        // Revision detection (PER-PROJECT scope): look inside <projectFolder>/PDF/ for existing baseName(_Rn).pdf
        let safeName = baseName;
        let maxRev = -1;
        try {
            const pdfFolderPath = `:/Metfraa_Costing_App/Generated_Costings/${projectFolder}/PDF:`;
            const listing = await graphClient
                .api(`/users/${targetEmail}/drive/root${pdfFolderPath}/children`)
                .top(200)
                .get();
            const existing = (listing.value || [])
                .filter(f => f.file && f.name && f.name.toLowerCase().endsWith('.pdf'))
                .map(f => f.name.replace(/\.pdf$/i, ''));
            const escBase = baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const re = new RegExp(`^${escBase}(?:_R(\\d+))?$`, 'i');
            existing.forEach(n => {
                const m = n.match(re);
                if (m) {
                    const rev = m[1] ? parseInt(m[1], 10) : 0;
                    if (rev > maxRev) maxRev = rev;
                }
            });
            if (maxRev >= 0) safeName = `${baseName}_R${maxRev + 1}`;
        } catch (listErr) {
            // Folder doesn't exist yet — first save, use baseName as-is
            console.log(`First save for project "${projectFolder}" — folder will be auto-created.`);
        }

        // Safety net against Graph API listing staleness (a fresh save's file
        // can take a few seconds to appear in folder listings). Probe the computed
        // target path directly — if it already exists, bump revision and re-probe.
        // We do at most 10 probes (well beyond any realistic stale-listing depth).
        for (let probe = 0; probe < 10; probe++) {
            const probePath = `:/Metfraa_Costing_App/Generated_Costings/${projectFolder}/PDF/${safeName}.pdf:`;
            let collision = false;
            try {
                await graphClient
                    .api(`/users/${targetEmail}/drive/root${probePath}`)
                    .select('id,name')
                    .get();
                collision = true; // 200 means file exists
            } catch (e) {
                // 404 means free — break out
                collision = false;
            }
            if (!collision) break;
            // File exists at safeName — bump to the next revision
            const m = safeName.match(/_R(\d+)$/i);
            const currentRev = m ? parseInt(m[1], 10) : 0; // baseName (no suffix) = rev 0
            const nextRev = currentRev + 1;
            safeName = `${baseName}_R${nextRev}`;
            if (nextRev > maxRev) maxRev = nextRev;
            console.log(`Collision detected — bumping to ${safeName}`);
        }

        const results = { revision: safeName === baseName ? 0 : parseInt(safeName.match(/_R(\d+)$/i)[1], 10), projectFolder };

        // Save PDF inside <projectFolder>/PDF/
        if (pdfBase64) {
            const cleanB64 = pdfBase64.replace(/^data:application\/pdf;base64,/, '');
            const pdfBuffer = Buffer.from(cleanB64, 'base64');
            const pdfPath = `:/Metfraa_Costing_App/Generated_Costings/${projectFolder}/PDF/${safeName}.pdf:`;
            const pdfResult = await graphClient
                .api(`/users/${targetEmail}/drive/root${pdfPath}/content`)
                .header('Content-Type', 'application/pdf')
                .put(pdfBuffer);
            results.pdf = { fileName: `${safeName}.pdf`, webUrl: pdfResult.webUrl || null };
        }

        // Save JSON inside <projectFolder>/JSON/
        if (jsonData) {
            const jsonString = typeof jsonData === 'string' ? jsonData : JSON.stringify(jsonData, null, 2);
            const jsonPath = `:/Metfraa_Costing_App/Generated_Costings/${projectFolder}/JSON/${safeName}.json:`;
            const jsonResult = await graphClient
                .api(`/users/${targetEmail}/drive/root${jsonPath}/content`)
                .header('Content-Type', 'application/json')
                .put(jsonString);
            results.json = { fileName: `${safeName}.json`, webUrl: jsonResult.webUrl || null };
        }

        // Append a row to the centralised analytics Excel (Build E)
        try { await appendAnalyticsRow(jsonData, safeName, projectFolder); } catch (_) {}

        // Auto-submit this revision for approval (Phase 1 — approval workflow).
        // Saving = submitting. Old pending auto-archives. Email to arasu fire-and-forget.
        try {
            const parsedJson = (() => {
                if (!jsonData) return {};
                if (typeof jsonData === 'string') { try { return JSON.parse(jsonData); } catch (_) { return {}; } }
                return jsonData;
            })();
            const det = (parsedJson && parsedJson.details) || {};
            const sels = (parsedJson && parsedJson.selections) || {};
            const paramList = Object.keys(sels);

            const approvalState = await readApprovalsState();
            if (!approvalState[projectFolder]) approvalState[projectFolder] = { latestPending: null, revisions: {} };
            archivePreviousPending(approvalState[projectFolder]);
            approvalState[projectFolder].revisions[safeName] = {
                status: 'pending',
                submittedAt: new Date().toISOString(),
                submittedBy: det.client || '',
                ref: det.ref || '',
                projectName: det.name || '',
                clientName: det.client || '',
                pdfUrl: (results.pdf && results.pdf.webUrl) || null,
                jsonUrl: (results.json && results.json.webUrl) || null,
                tonnage: det.tonnage != null ? det.tonnage : null,
                area: det.area != null ? det.area : null,
                location: det.loc || null,
                paramList,
                decidedAt: null,
                remarks: {},
                globalNote: null
            };
            approvalState[projectFolder].latestPending = safeName;
            await writeApprovalsState(approvalState);

            // Fire approval email — fire-and-forget so save isn't blocked on SMTP
            const cta = buildDeepLink('approvals', projectFolder, safeName);
            const html = buildEmailHtml({
                title: 'Costing Estimation — Approval Required',
                intro: `A new costing estimation has been submitted for your review.`,
                infoRows: [
                    ['Client', det.client || '—'],
                    ['Project', det.name || '—'],
                    ['Quote Ref', det.ref || '—'],
                    ['Revision', safeName],
                    ['Location', det.loc || '—'],
                    ['Tonnage', det.tonnage ? `${det.tonnage} MT` : '—'],
                    ['Area', det.area ? `${det.area} sqm` : '—'],
                    ['PDF', (results.pdf && results.pdf.webUrl) ? `<a href="${results.pdf.webUrl}" style="color:#0066b3;">View on OneDrive</a>` : '—']
                ],
                ctaUrl: cta,
                ctaLabel: '🔍 Review on App'
            });
            sendEmail({
                to: APPROVER_EMAIL,
                subject: `[Approval] ${det.client || ''} · ${det.name || ''} · ${safeName}`,
                htmlBody: html
            });

            results.approval = { status: 'pending', submittedAt: approvalState[projectFolder].revisions[safeName].submittedAt };
        } catch (approvalErr) {
            // Surface as much detail as possible — Graph errors have statusCode + body + code
            const detail = {
                message: approvalErr && approvalErr.message,
                code: approvalErr && (approvalErr.code || approvalErr.statusCode),
                statusCode: approvalErr && approvalErr.statusCode,
                body: approvalErr && (approvalErr.body || approvalErr.responseBody || null)
            };
            console.error('❌ AUTO-SUBMIT FOR APPROVAL FAILED:', JSON.stringify(detail, null, 2));
            results.approval = { status: 'error', ...detail };
        }

        res.json({ success: true, ...results });
    } catch (error) {
        console.error("Save Quotation Error:", error.message);
        res.status(500).json({
            success: false,
            error: error.message,
            code: error.code || error.statusCode || 'UNKNOWN'
        });
    }
});

// =============================================================
// ROUTE: List all projects (each is a folder under Generated_Costings/)
// Returns: [{ name, revisionCount, lastModified }]
// =============================================================
app.get('/api/list-projects', async (req, res) => {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const root = ':/Metfraa_Costing_App/Generated_Costings:';
        let listing;
        try {
            listing = await graphClient
                .api(`/users/${targetEmail}/drive/root${root}/children`)
                .top(500)
                .get();
        } catch (e) {
            return res.json({ success: true, projects: [] });
        }
        // Each folder is one project. Skip files (legacy flat layout) and the analytics folder.
        const folders = (listing.value || []).filter(f =>
            f.folder && f.name && !f.name.startsWith('_') && f.name.toLowerCase() !== 'pdf' && f.name.toLowerCase() !== 'json'
        );
        // For each project folder, count revisions inside PDF subfolder
        const projects = [];
        for (const folder of folders) {
            let revCount = 0;
            let lastMod = folder.lastModifiedDateTime;
            try {
                const pdfList = await graphClient
                    .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/${folder.name}/PDF:/children`)
                    .top(200)
                    .get();
                const pdfs = (pdfList.value || []).filter(f => f.file && f.name.toLowerCase().endsWith('.pdf'));
                revCount = pdfs.length;
                if (pdfs.length > 0) {
                    const latest = pdfs.reduce((a,b) => new Date(a.lastModifiedDateTime) > new Date(b.lastModifiedDateTime) ? a : b);
                    lastMod = latest.lastModifiedDateTime;
                }
            } catch (_) {
                // PDF subfolder may not exist for an empty project
            }
            projects.push({ name: folder.name, revisionCount: revCount, lastModified: lastMod });
        }
        projects.sort((a,b) => new Date(b.lastModified) - new Date(a.lastModified));
        res.json({ success: true, projects });
    } catch (error) {
        console.error('list-projects error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: List all revisions for a specific project
// Query: ?project=<folderName>
// Returns: [{ name, revisionNumber, lastModified, hasJson }]
// =============================================================
app.get('/api/list-project-revisions', async (req, res) => {
    try {
        const project = req.query.project;
        if (!project) return res.status(400).json({ success: false, error: 'Missing project query param' });
        const targetEmail = process.env.TARGET_USER_EMAIL;

        // List PDFs
        let pdfs = [];
        try {
            const pdfList = await graphClient
                .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/${project}/PDF:/children`)
                .top(200)
                .get();
            pdfs = (pdfList.value || []).filter(f => f.file && f.name.toLowerCase().endsWith('.pdf'));
        } catch (_) {}

        // List JSONs to check which revisions have a JSON
        let jsons = new Set();
        try {
            const jsonList = await graphClient
                .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/${project}/JSON:/children`)
                .top(200)
                .get();
            (jsonList.value || []).forEach(f => {
                if (f.file && f.name.toLowerCase().endsWith('.json')) {
                    jsons.add(f.name.replace(/\.json$/i, ''));
                }
            });
        } catch (_) {}

        const revisions = pdfs.map(f => {
            const baseName = f.name.replace(/\.pdf$/i, '');
            const m = baseName.match(/_R(\d+)$/i);
            return {
                name: baseName,
                fileName: f.name,
                revisionNumber: m ? parseInt(m[1], 10) : 0,
                lastModified: f.lastModifiedDateTime,
                size: f.size,
                hasJson: jsons.has(baseName),
                webUrl: f.webUrl || null
            };
        }).sort((a, b) => a.revisionNumber - b.revisionNumber);

        res.json({ success: true, project, revisions });
    } catch (error) {
        console.error('list-project-revisions error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: Get a specific revision's JSON content (for loading or viewing)
// Query: ?project=<folder>&revision=<baseName>
// =============================================================
app.get('/api/get-project-revision', async (req, res) => {
    try {
        const { project, revision } = req.query;
        if (!project || !revision) {
            return res.status(400).json({ success: false, error: 'Missing project or revision' });
        }
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const path = `:/Metfraa_Costing_App/Generated_Costings/${project}/JSON/${revision}.json:`;
        const content = await graphClient
            .api(`/users/${targetEmail}/drive/root${path}/content`)
            .get();
        // Graph SDK auto-parses JSON files (it reads Content-Type and decodes the body),
        // so `content` may already be a JS object. Otherwise treat it as text and parse.
        let parsed;
        if (content && typeof content === 'object' && !Buffer.isBuffer(content)) {
            parsed = content;
        } else {
            const text = typeof content === 'string' ? content
                       : Buffer.isBuffer(content) ? content.toString('utf8')
                       : String(content || '');
            try { parsed = JSON.parse(text); }
            catch (e) {
                return res.status(500).json({ success: false, error: 'Saved JSON is not valid: ' + e.message });
            }
        }
        res.json({ success: true, project, revision, data: parsed });
    } catch (error) {
        console.error('get-project-revision error:', error.message);
        res.status(error.statusCode || 500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: Get a specific revision's PDF (returns base64) for in-app viewing
// Query: ?project=<folder>&revision=<baseName>
// =============================================================
// =============================================================
// ROUTE: Save Customer Quote PDF inside per-project Customer Quotes folder
// Body: { projectName, clientName, date, pdfBase64 }
// =============================================================
app.post('/api/save-customer-quote', async (req, res) => {
    try {
        const { projectName, clientName, date, pdfBase64 } = req.body;
        if (!projectName || !date || !pdfBase64) {
            return res.status(400).json({ success: false, error: 'Missing projectName, date, or pdfBase64' });
        }
        const targetEmail = process.env.TARGET_USER_EMAIL;

        const projectFolderParts = [];
        if (clientName && String(clientName).trim() && String(clientName).trim() !== '—') {
            projectFolderParts.push(sanitizeName(clientName));
        }
        projectFolderParts.push(sanitizeName(projectName));
        const projectFolder = projectFolderParts.join('_');

        const dateSafe = sanitizeName(date);
        const baseName = dateSafe;

        // Revision detection inside <projectFolder>/Customer Quotes/
        let safeName = baseName;
        try {
            const folderPath = `:/Metfraa_Costing_App/Generated_Costings/${projectFolder}/Customer Quotes:`;
            const listing = await graphClient
                .api(`/users/${targetEmail}/drive/root${folderPath}/children`)
                .top(200)
                .get();
            const existing = (listing.value || [])
                .filter(f => f.file && f.name && f.name.toLowerCase().endsWith('.pdf'))
                .map(f => f.name.replace(/\.pdf$/i, ''));
            const escBase = baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const re = new RegExp(`^${escBase}(?:_R(\\d+))?$`, 'i');
            let maxRev = -1;
            existing.forEach(n => {
                const m = n.match(re);
                if (m) {
                    const rev = m[1] ? parseInt(m[1], 10) : 0;
                    if (rev > maxRev) maxRev = rev;
                }
            });
            if (maxRev >= 0) safeName = `${baseName}_R${maxRev + 1}`;
        } catch (_) {
            // First customer quote for this project
        }

        const cleanB64 = pdfBase64.replace(/^data:application\/pdf;base64,/, '');
        const pdfBuffer = Buffer.from(cleanB64, 'base64');
        const pdfPath = `:/Metfraa_Costing_App/Generated_Costings/${projectFolder}/Customer Quotes/${safeName}.pdf:`;
        const pdfResult = await graphClient
            .api(`/users/${targetEmail}/drive/root${pdfPath}/content`)
            .header('Content-Type', 'application/pdf')
            .put(pdfBuffer);

        res.json({
            success: true,
            fileName: `${safeName}.pdf`,
            projectFolder,
            webUrl: pdfResult.webUrl || null,
            revision: safeName === baseName ? 0 : parseInt(safeName.match(/_R(\d+)$/i)[1], 10)
        });
    } catch (error) {
        console.error('save-customer-quote error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.get('/api/get-project-revision-pdf', async (req, res) => {
    try {
        const { project, revision } = req.query;
        if (!project || !revision) {
            return res.status(400).json({ success: false, error: 'Missing project or revision' });
        }
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const path = `:/Metfraa_Costing_App/Generated_Costings/${project}/PDF/${revision}.pdf:`;
        const stream = await graphClient
            .api(`/users/${targetEmail}/drive/root${path}/content`)
            .responseType('arraybuffer')
            .get();
        const buf = Buffer.from(stream);
        res.json({ success: true, base64: buf.toString('base64') });
    } catch (error) {
        console.error('get-project-revision-pdf error:', error.message);
        res.status(error.statusCode || 500).json({ success: false, error: error.message });
    }
});

// =============================================================
// Centralised analytics Excel (Build E)
// Appends one row per parameter per saved quote to:
// /Metfraa_Costing_App/Generated_Costings/_Analytics/cost_history.xlsx
// =============================================================
const ANALYTICS_PATH = ':/Metfraa_Costing_App/Generated_Costings/_Analytics/cost_history.xlsx:';
const SEQ_FLOOR_PATH = ':/Metfraa_Costing_App/Generated_Costings/_Analytics/seq_config.json:';
const PROJECT_DEFAULTS_PATH = ':/Metfraa_Costing_App/Generated_Costings/_Analytics/project_defaults.json:';
const APPROVALS_PATH = ':/Metfraa_Costing_App/Generated_Costings/_Analytics/approvals.json:';
const APPROVER_EMAIL = 'arasu@metfraa.com';
const COSTING_TEAM_EMAIL = 'costing@metfraa.com';
const APP_BASE_URL = process.env.APP_BASE_URL || 'https://metfraa-costing-engine.onrender.com';

// =========================================================
// APPROVALS — read/write the central approvals.json on OneDrive.
// Schema: { "<projectFolder>": { latestPending, revisions: { "<rev>": {...} } } }
// =========================================================
async function readApprovalsState() {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const content = await graphClient
            .api(`/users/${targetEmail}/drive/root${APPROVALS_PATH}/content`)
            .get();
        let parsed;
        if (content && typeof content === 'object' && !Buffer.isBuffer(content)) {
            parsed = content;
        } else {
            const text = typeof content === 'string' ? content
                       : Buffer.isBuffer(content) ? content.toString('utf8')
                       : String(content || '');
            try { parsed = JSON.parse(text); } catch (_) { return {}; }
        }
        return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (_) {
        return {}; // No file yet → empty state
    }
}

async function writeApprovalsState(state) {
    const targetEmail = process.env.TARGET_USER_EMAIL;
    const payload = JSON.stringify(state, null, 2);
    return await graphClient
        .api(`/users/${targetEmail}/drive/root${APPROVALS_PATH}/content`)
        .header('Content-Type', 'application/json')
        .put(payload);
}

// Helper: archive ALL non-decided (pending/changes_requested) revisions for a project
// before submitting a new one. Mutates `projectEntry` in place.
function archivePreviousPending(projectEntry) {
    if (!projectEntry || !projectEntry.revisions) return;
    Object.keys(projectEntry.revisions).forEach(revName => {
        const r = projectEntry.revisions[revName];
        if (r && (r.status === 'pending' || r.status === 'changes_requested')) {
            r.status = 'archived';
            r.archivedAt = new Date().toISOString();
        }
    });
    projectEntry.latestPending = null;
}

// =========================================================
// EMAIL HELPER — sends via Microsoft Graph using the TARGET_USER_EMAIL identity.
// Uses /users/{targetEmail}/sendMail (app-only auth, no SMTP credentials needed).
// =========================================================
async function sendEmail({ to, cc, subject, htmlBody }) {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        if (!targetEmail) {
            console.warn('sendEmail skipped — TARGET_USER_EMAIL not set');
            return false;
        }
        const toList = Array.isArray(to) ? to : [to];
        const ccList = cc ? (Array.isArray(cc) ? cc : [cc]) : [];
        const message = {
            message: {
                subject: subject || '(no subject)',
                body: { contentType: 'HTML', content: htmlBody || '' },
                toRecipients: toList.filter(Boolean).map(a => ({ emailAddress: { address: a } })),
                ccRecipients: ccList.filter(Boolean).map(a => ({ emailAddress: { address: a } }))
            },
            saveToSentItems: true
        };
        await graphClient
            .api(`/users/${targetEmail}/sendMail`)
            .post(message);
        console.log(`📧 sent: "${subject}" → ${toList.join(', ')}`);
        return true;
    } catch (err) {
        const detail = {
            message: err && err.message,
            statusCode: err && err.statusCode,
            code: err && err.code,
            body: err && (err.body || err.responseBody || null)
        };
        console.error('❌ sendEmail error:', JSON.stringify(detail, null, 2));
        return false;
    }
}

// HTML email template — branded header, info table, primary CTA button, footer.
function buildEmailHtml({ title, intro, infoRows, ctaUrl, ctaLabel, paramRemarks }) {
    const rows = (infoRows || []).map(r =>
        `<tr><td style="padding:6px 12px 6px 0; color:#666; font-weight:600; vertical-align:top;">${r[0]}</td>
             <td style="padding:6px 0; color:#222;">${r[1]}</td></tr>`
    ).join('');

    let remarksBlock = '';
    if (paramRemarks && Object.keys(paramRemarks).length) {
        const remarkRows = Object.entries(paramRemarks)
            .filter(([, txt]) => txt && String(txt).trim())
            .map(([param, txt]) =>
                `<tr><td style="padding:8px 14px 8px 0; color:#002b5f; font-weight:600; vertical-align:top; width:35%;">${param}</td>
                     <td style="padding:8px 0; color:#333; white-space:pre-wrap;">${String(txt).replace(/[<>&]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]))}</td></tr>`
            ).join('');
        if (remarkRows) {
            remarksBlock = `
                <div style="margin:20px 0 6px; padding:14px 16px; background:#fff8e1; border-left:4px solid #ffb300; border-radius:4px;">
                    <div style="font-weight:700; color:#664d03; margin-bottom:8px;">Requested Changes</div>
                    <table style="width:100%; border-collapse:collapse; font-size:13px;">${remarkRows}</table>
                </div>
            `;
        }
    }

    return `<!DOCTYPE html>
<html><body style="margin:0; padding:0; background:#f4f6f9; font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;">
  <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f4f6f9; padding:24px 0;">
    <tr><td align="center">
      <table cellpadding="0" cellspacing="0" border="0" width="600" style="background:#fff; border-radius:8px; box-shadow:0 1px 4px rgba(0,0,0,0.06); overflow:hidden;">
        <tr><td style="background:#002b5f; padding:18px 24px;">
          <div style="color:#fff; font-size:18px; font-weight:700;">METFRAA · Costing Workflow</div>
        </td></tr>
        <tr><td style="padding:22px 28px;">
          <h2 style="margin:0 0 12px; color:#002b5f; font-size:20px; font-weight:700;">${title || ''}</h2>
          <p style="margin:0 0 18px; color:#555; font-size:14px; line-height:1.5;">${intro || ''}</p>
          <table style="width:100%; border-collapse:collapse; font-size:13px; margin-bottom:10px;">${rows}</table>
          ${remarksBlock}
          ${ctaUrl ? `<div style="margin:24px 0 6px;">
            <a href="${ctaUrl}" style="display:inline-block; background:#0066b3; color:#fff; text-decoration:none; padding:11px 22px; border-radius:5px; font-weight:600; font-size:14px;">${ctaLabel || 'Open'}</a>
          </div>` : ''}
        </td></tr>
        <tr><td style="background:#f9fafb; padding:14px 28px; color:#888; font-size:11px; border-top:1px solid #eee;">
          This is an automated message from the METFRAA Costing Engine.
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}

// Read project defaults (margin %, gst %) — fall back to hardcoded 12 / 18 if no file
async function readProjectDefaults() {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const content = await graphClient
            .api(`/users/${targetEmail}/drive/root${PROJECT_DEFAULTS_PATH}/content`)
            .get();
        let parsed;
        if (content && typeof content === 'object' && !Buffer.isBuffer(content)) {
            parsed = content;
        } else {
            const text = typeof content === 'string' ? content
                       : Buffer.isBuffer(content) ? content.toString('utf8')
                       : String(content || '');
            try { parsed = JSON.parse(text); } catch (_) { return { margin: 12, gst: 18 }; }
        }
        return {
            margin: Number.isFinite(parseFloat(parsed?.margin)) ? parseFloat(parsed.margin) : 12,
            gst: Number.isFinite(parseFloat(parsed?.gst)) ? parseFloat(parsed.gst) : 18
        };
    } catch (_) {
        return { margin: 12, gst: 18 };
    }
}

async function writeProjectDefaults(margin, gst) {
    const targetEmail = process.env.TARGET_USER_EMAIL;
    const payload = JSON.stringify({
        margin: parseFloat(margin),
        gst: parseFloat(gst),
        setAt: new Date().toISOString()
    });
    return await graphClient
        .api(`/users/${targetEmail}/drive/root${PROJECT_DEFAULTS_PATH}/content`)
        .header('Content-Type', 'application/json')
        .put(payload);
}

// Read the sequence floor (default 0 if no config file exists)
async function readSeqFloor() {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const content = await graphClient
            .api(`/users/${targetEmail}/drive/root${SEQ_FLOOR_PATH}/content`)
            .get();
        let parsed;
        if (content && typeof content === 'object' && !Buffer.isBuffer(content)) {
            parsed = content;
        } else {
            const text = typeof content === 'string' ? content
                       : Buffer.isBuffer(content) ? content.toString('utf8')
                       : String(content || '');
            try { parsed = JSON.parse(text); } catch (_) { return 0; }
        }
        return parseInt(parsed?.floor, 10) || 0;
    } catch (_) {
        return 0; // No config file → floor 0 (i.e. no floor enforced)
    }
}

async function writeSeqFloor(floor) {
    const targetEmail = process.env.TARGET_USER_EMAIL;
    const payload = JSON.stringify({ floor: parseInt(floor, 10), setAt: new Date().toISOString() });
    return await graphClient
        .api(`/users/${targetEmail}/drive/root${SEQ_FLOOR_PATH}/content`)
        .header('Content-Type', 'application/json')
        .put(payload);
}
const ANALYTICS_HEADERS = [
    'Saved At', 'Client', 'Project', 'Revision', 'Project Folder', 'File Name',
    'Quote Ref', 'Quote Date', 'Location', 'Distance (km)',
    'Parameter', 'Qty', 'Unit', 'Pre-Margin Unit', 'Margin %',
    'Post-Margin Unit', 'Rounded Unit', 'Rounding To', 'Param Total'
];

async function appendAnalyticsRow(jsonData, savedFileName, projectFolder) {
    if (!jsonData) return;
    const data = typeof jsonData === 'string' ? JSON.parse(jsonData) : jsonData;
    if (!data || !data.details) return;
    const d = data.details || {};
    const step3 = data.step3 || {};
    const qtys = data.qtys || {};
    const savedAt = data.savedAt || new Date().toISOString();
    const revMatch = savedFileName.match(/_R(\d+)$/i);
    const revision = revMatch ? parseInt(revMatch[1], 10) : 0;

    // Build rows: one per parameter that has qty > 0
    const rowsToAppend = [];
    Object.keys(qtys).forEach(p => {
        const q = parseFloat(qtys[p]) || 0;
        if (q <= 0) return;
        // We can only compute final values if the saved JSON has step3 overrides;
        // otherwise we record what we know.
        const ov = step3[p] || {};
        const margin = ov.margin != null ? ov.margin : (d.margin || 0);
        const roundEnabled = ov.roundEnabled !== false;
        // We don't have MASTER_DB context here so we can't re-derive unit cost.
        // Use whatever's stored in `data.step3Computed` if the client populated it.
        const comp = (data.step3Computed && data.step3Computed[p]) || {};
        rowsToAppend.push([
            savedAt,
            d.client || '',
            d.name || '',
            revision,
            projectFolder || '',
            savedFileName,
            d.ref || '',
            d.date || '',
            d.loc || '',
            d.dist || 0,
            p,
            q,
            comp.unit || '',
            comp.unitPre != null ? Math.round(comp.unitPre) : '',
            margin,
            comp.unitPostMargin != null ? Math.round(comp.unitPostMargin) : '',
            comp.unitRounded != null ? Math.round(comp.unitRounded) : '',
            roundEnabled ? (comp.roundingUnit || '') : 'off',
            comp.paramTotal != null ? Math.round(comp.paramTotal) : ''
        ]);
    });
    if (rowsToAppend.length === 0) return;

    const targetEmail = process.env.TARGET_USER_EMAIL;

    // Try to download existing analytics file; if absent, create new
    let workbook;
    let isNew = false;
    try {
        const existing = await graphClient
            .api(`/users/${targetEmail}/drive/root${ANALYTICS_PATH}/content`)
            .responseType('arraybuffer')
            .get();
        workbook = xlsx.read(Buffer.from(existing), { type: 'buffer' });
    } catch (e) {
        isNew = true;
        workbook = xlsx.utils.book_new();
    }

    let sheet = workbook.Sheets['Analytics'];
    let existingData = [];
    if (sheet) {
        existingData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    }
    if (existingData.length === 0) {
        existingData.push(ANALYTICS_HEADERS);
    }
    rowsToAppend.forEach(r => existingData.push(r));

    const newSheet = xlsx.utils.aoa_to_sheet(existingData);
    if (workbook.Sheets['Analytics']) {
        workbook.Sheets['Analytics'] = newSheet;
    } else {
        xlsx.utils.book_append_sheet(workbook, newSheet, 'Analytics');
    }
    const buf = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    await graphClient
        .api(`/users/${targetEmail}/drive/root${ANALYTICS_PATH}/content`)
        .header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        .put(buf);
}

// =============================================================
// 3. ROUTE: Load Quotation by reference
// =============================================================
app.get('/api/load-quotation', async (req, res) => {
    try {
        const ref = req.query.ref;
        if (!ref) return res.status(400).json({ success: false, error: 'Missing ref query param' });

        const targetEmail = process.env.TARGET_USER_EMAIL;
        const safeRef = sanitizeName(ref);

        // Walk per-project folders + legacy flat /JSON/ folder
        const allFiles = [];
        const rootPath = ':/Metfraa_Costing_App/Generated_Costings:';
        let rootListing;
        try {
            rootListing = await graphClient
                .api(`/users/${targetEmail}/drive/root${rootPath}/children`)
                .top(500)
                .get();
        } catch (e) {
            return res.status(404).json({ success: false, error: 'No saved quotations folder in OneDrive yet.' });
        }
        const projectFolders = (rootListing.value || []).filter(f =>
            f.folder && f.name && !f.name.startsWith('_')
        );
        for (const pf of projectFolders) {
            try {
                const jl = await graphClient
                    .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/${pf.name}/JSON:/children`)
                    .top(200)
                    .get();
                (jl.value || []).forEach(f => {
                    if (f.file && f.name && f.name.toLowerCase().endsWith('.json')) {
                        allFiles.push(f);
                    }
                });
            } catch (_) {}
        }
        // Legacy flat folder
        try {
            const flatList = await graphClient
                .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/JSON:/children`)
                .top(200)
                .get();
            (flatList.value || []).forEach(f => {
                if (f.file && f.name && f.name.toLowerCase().endsWith('.json')) {
                    allFiles.push(f);
                }
            });
        } catch (_) {}

        let matches = allFiles.filter(f =>
            f.name.toLowerCase().includes(safeRef.toLowerCase())
        );

        if (matches.length === 0) {
            return res.status(404).json({
                success: false,
                error: `No saved quotation found matching "${ref}"`,
                searchedFor: safeRef,
                available: allFiles.map(f => f.name)
            });
        }

        matches.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
        const target = matches[0];

        const fileMeta = await graphClient
            .api(`/users/${targetEmail}/drive/items/${target.id}`)
            .get();
        const downloadUrl = fileMeta['@microsoft.graph.downloadUrl'];
        const dl = await fetch(downloadUrl);
        const text = await dl.text();
        let parsed;
        try { parsed = JSON.parse(text); }
        catch (e) {
            return res.status(500).json({ success: false, error: 'Saved JSON is corrupted: ' + e.message });
        }

        res.json({
            success: true,
            fileName: target.name,
            lastModified: target.lastModifiedDateTime,
            otherMatches: matches.slice(1, 6).map(m => m.name),
            data: parsed
        });
    } catch (error) {
        console.error("Load Quotation Error:", error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// 4. ROUTE: List all saved quotations
// =============================================================
// =============================================================
// ROUTE: Auto-generate the next quote number by scanning saved JSON files
// for the highest existing Met/est/<N>/<YYYY> ref and returning N+1.
// =============================================================
app.get('/api/next-quote-number', async (req, res) => {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const now = new Date();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const yyyy = now.getFullYear();

        // List all per-project folders under Generated_Costings/
        const rootPath = ':/Metfraa_Costing_App/Generated_Costings:';
        let rootListing;
        try {
            rootListing = await graphClient
                .api(`/users/${targetEmail}/drive/root${rootPath}/children`)
                .top(500)
                .get();
        } catch (listErr) {
            // Folder doesn't exist yet — start from 01
            return res.json({ success: true, quoteRef: `Met/est/${mm}/${yyyy}/01`, sequence: 1 });
        }

        const projectFolders = (rootListing.value || []).filter(f =>
            f.folder && f.name && !f.name.startsWith('_')
        );

        // Collect all JSON files across all project folders, then peek inside for refs.
        // We scan up to ~50 recent JSON files total (across all projects) — quick but bounded.
        const allJsonFiles = [];
        for (const pf of projectFolders) {
            try {
                const jl = await graphClient
                    .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/${pf.name}/JSON:/children`)
                    .top(50)
                    .get();
                (jl.value || []).forEach(f => {
                    if (f.file && f.name && f.name.toLowerCase().endsWith('.json')) {
                        allJsonFiles.push(f);
                    }
                });
            } catch (_) { /* JSON subfolder might not exist for this project */ }
        }

        // Sort by lastModified desc and limit to 50 most recent (sequence is monotonic so
        // checking recent ones is sufficient to find the highest)
        const recent = allJsonFiles
            .sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime))
            .slice(0, 50);

        // Match BOTH the old format (Met/est/XX/YYYY) AND the new format (Met/est/MM/YYYY/XX)
        // so legacy refs don't break sequencing.
        // Old format: Met/est/01/2026   → 3 parts after Met/est, last is year
        // New format: Met/est/06/2026/01 → 4 parts, last is sequence
        const reNew = /met\/est\/(\d{1,2})\/(\d{4})\/(\d+)/i;
        const reOld = /met\/est\/(\d+)\/(\d{4})/i;

        let maxSeq = 0;
        for (const f of recent) {
            try {
                const content = await graphClient
                    .api(`/users/${targetEmail}/drive/items/${f.id}/content`)
                    .get();
                let parsed = null;
                if (content && typeof content === 'object' && !Buffer.isBuffer(content)) {
                    parsed = content; // SDK already parsed it
                } else {
                    const text = typeof content === 'string' ? content
                              : Buffer.isBuffer(content) ? content.toString('utf8')
                              : String(content || '');
                    try { parsed = JSON.parse(text); } catch (_) { continue; }
                }
                const ref = String(parsed?.details?.ref || '');
                let seq = 0;
                const mNew = ref.match(reNew);
                if (mNew) {
                    seq = parseInt(mNew[3], 10);
                } else {
                    const mOld = ref.match(reOld);
                    if (mOld) seq = parseInt(mOld[1], 10);
                }
                if (seq > maxSeq) maxSeq = seq;
            } catch (_) { continue; }
        }

        // Honour the sequence floor (set by admin endpoint /api/set-seq-floor)
        const floor = await readSeqFloor();
        const effectiveMax = Math.max(maxSeq, floor);
        const nextSeq = effectiveMax + 1;
        const padded = nextSeq < 10 ? `0${nextSeq}` : String(nextSeq);
        const quoteRef = `Met/est/${mm}/${yyyy}/${padded}`;
        res.json({ success: true, quoteRef, sequence: nextSeq, previousMax: maxSeq, floor });
    } catch (error) {
        console.error('Next Quote Number Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// Admin endpoint: set the sequence floor. The next quote will be (max(scanned, floor) + 1).
// Body: { floor: <integer>, password: <admin password> }
app.post('/api/set-seq-floor', async (req, res) => {
    try {
        const { floor, password } = req.body || {};
        if (process.env.ADMIN_PASSWORD && password !== process.env.ADMIN_PASSWORD) {
            return res.status(401).json({ success: false, error: 'Unauthorized' });
        }
        const f = parseInt(floor, 10);
        if (!Number.isFinite(f) || f < 0) {
            return res.status(400).json({ success: false, error: 'Invalid floor value' });
        }
        await writeSeqFloor(f);
        res.json({ success: true, floor: f, message: `Next quote sequence will be at least ${f + 1}` });
    } catch (error) {
        console.error('set-seq-floor error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: Get project defaults (margin %, gst %) used for NEW projects
// =============================================================
app.get('/api/get-defaults', async (req, res) => {
    try {
        const defaults = await readProjectDefaults();
        res.json({ success: true, ...defaults });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: Set project defaults (margin %, gst %). Admin password required.
// =============================================================
app.post('/api/set-defaults', async (req, res) => {
    try {
        const { margin, gst, password } = req.body || {};
        if (process.env.ADMIN_PASSWORD && password !== process.env.ADMIN_PASSWORD) {
            return res.status(401).json({ success: false, error: 'Unauthorized' });
        }
        const m = parseFloat(margin), g = parseFloat(gst);
        if (!Number.isFinite(m) || m < 0 || !Number.isFinite(g) || g < 0) {
            return res.status(400).json({ success: false, error: 'Invalid margin or gst' });
        }
        await writeProjectDefaults(m, g);
        res.json({ success: true, margin: m, gst: g });
    } catch (error) {
        console.error('set-defaults error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// APPROVAL WORKFLOW ENDPOINTS
// =============================================================

// Helper: build a deep-link to the in-app routes (used in email CTAs)
function buildDeepLink(route, project, revision) {
    const q = encodeURIComponent(project);
    const r = encodeURIComponent(revision);
    return `${APP_BASE_URL}/#${route}/${q}/${r}`;
}

// POST /api/arasu-auth  — body { password }
// Returns { success: true, token: "<opaque>" } if password matches ARASU_APPROVAL_PASSWORD.
// Token is a simple HMAC-of-current-day; frontend stashes it and sends back on every approval action.
const _arasuTokenSecret = process.env.ARASU_APPROVAL_PASSWORD || '';
function _makeArasuToken() {
    // Token = base64(YYYYMMDD + ":" + sha256(YYYYMMDD + secret)) — valid for current UTC day
    const crypto = require('crypto');
    const day = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const hash = crypto.createHash('sha256').update(day + _arasuTokenSecret).digest('hex').slice(0, 24);
    return Buffer.from(`${day}:${hash}`).toString('base64');
}
function _verifyArasuToken(token) {
    if (!token || !_arasuTokenSecret) return false;
    try {
        const decoded = Buffer.from(token, 'base64').toString('utf8');
        const [day, hash] = decoded.split(':');
        if (!day || !hash) return false;
        const crypto = require('crypto');
        const expected = crypto.createHash('sha256').update(day + _arasuTokenSecret).digest('hex').slice(0, 24);
        // Accept tokens from today or yesterday (24-48 hour validity, simpler than tracking expiry)
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const yest = new Date(Date.now() - 86400_000).toISOString().slice(0, 10).replace(/-/g, '');
        return hash === expected && (day === today || day === yest);
    } catch (_) { return false; }
}

app.post('/api/arasu-auth', async (req, res) => {
    const { password } = req.body || {};
    if (!process.env.ARASU_APPROVAL_PASSWORD) {
        return res.status(500).json({ success: false, error: 'ARASU_APPROVAL_PASSWORD env var not configured' });
    }
    if (password !== process.env.ARASU_APPROVAL_PASSWORD) {
        return res.status(401).json({ success: false, error: 'Invalid password' });
    }
    res.json({ success: true, token: _makeArasuToken() });
});

// POST /api/submit-for-approval
// Body: { project, revision, clientName, projectName, ref, pdfUrl, jsonUrl, params: [<paramName>] }
// Called automatically after a successful save. Archives any older pending for this project,
// sets the new one as pending, emails arasu.
app.post('/api/submit-for-approval', async (req, res) => {
    try {
        const { project, revision, clientName, projectName, ref, pdfUrl, jsonUrl, params, tonnage, area, location } = req.body || {};
        if (!project || !revision) {
            return res.status(400).json({ success: false, error: 'Missing project or revision' });
        }
        const state = await readApprovalsState();
        if (!state[project]) state[project] = { latestPending: null, revisions: {} };
        // Archive any non-decided revisions for this project
        archivePreviousPending(state[project]);
        // Add the new pending revision
        state[project].revisions[revision] = {
            status: 'pending',
            submittedAt: new Date().toISOString(),
            submittedBy: clientName || '',
            ref: ref || '',
            projectName: projectName || '',
            clientName: clientName || '',
            pdfUrl: pdfUrl || null,
            jsonUrl: jsonUrl || null,
            tonnage: tonnage || null,
            area: area || null,
            location: location || null,
            paramList: Array.isArray(params) ? params : [],
            decidedAt: null,
            remarks: {},
            globalNote: null
        };
        state[project].latestPending = revision;
        await writeApprovalsState(state);

        // Fire the approval email to Arasu (non-blocking — don't fail the save if email errors)
        const cta = buildDeepLink('approvals', project, revision);
        const html = buildEmailHtml({
            title: 'Costing Estimation — Approval Required',
            intro: `A new costing estimation has been submitted for your review.`,
            infoRows: [
                ['Client', clientName || '—'],
                ['Project', projectName || '—'],
                ['Quote Ref', ref || '—'],
                ['Revision', revision],
                ['Location', location || '—'],
                ['Tonnage', tonnage ? `${tonnage} MT` : '—'],
                ['Area', area ? `${area} sqm` : '—'],
                ['PDF', pdfUrl ? `<a href="${pdfUrl}" style="color:#0066b3;">View on OneDrive</a>` : '—']
            ],
            ctaUrl: cta,
            ctaLabel: '🔍 Review on App'
        });
        // Fire-and-forget: the approval state is what matters, email is secondary
        sendEmail({ to: APPROVER_EMAIL, subject: `[Approval] ${clientName || ''} · ${projectName || ''} · ${revision}`, htmlBody: html });

        res.json({ success: true, status: 'pending', project, revision });
    } catch (error) {
        console.error('submit-for-approval error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// POST /api/approve-revision
// Body: { project, revision, token }   (token from /api/arasu-auth)
app.post('/api/approve-revision', async (req, res) => {
    try {
        const { project, revision, token } = req.body || {};
        if (!_verifyArasuToken(token)) {
            return res.status(401).json({ success: false, error: 'Unauthorized — invalid or expired token' });
        }
        if (!project || !revision) {
            return res.status(400).json({ success: false, error: 'Missing project or revision' });
        }
        const state = await readApprovalsState();
        const entry = state[project] && state[project].revisions[revision];
        if (!entry) {
            return res.status(404).json({ success: false, error: 'Revision not found' });
        }
        if (entry.status !== 'pending') {
            return res.status(400).json({ success: false, error: `Cannot approve — current status is "${entry.status}"` });
        }
        entry.status = 'approved';
        entry.decidedAt = new Date().toISOString();
        entry.remarks = {};
        if (state[project].latestPending === revision) state[project].latestPending = null;
        await writeApprovalsState(state);

        // Email costing team
        const cta = buildDeepLink('projects', project, revision);
        const html = buildEmailHtml({
            title: '✓ Costing Estimation — Approved',
            intro: `The following costing estimation has been approved. You can now proceed to generate the customer quote.`,
            infoRows: [
                ['Client', entry.clientName || '—'],
                ['Project', entry.projectName || '—'],
                ['Quote Ref', entry.ref || '—'],
                ['Revision', revision],
                ['Approved at', entry.decidedAt]
            ],
            ctaUrl: cta,
            ctaLabel: '📄 Open Project'
        });
        sendEmail({ to: COSTING_TEAM_EMAIL, subject: `[Approved] ${entry.clientName || ''} · ${entry.projectName || ''} · ${revision}`, htmlBody: html });

        res.json({ success: true, status: 'approved', project, revision });
    } catch (error) {
        console.error('approve-revision error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// POST /api/request-changes
// Body: { project, revision, token, remarks: { paramName: "text", ... }, globalNote }
app.post('/api/request-changes', async (req, res) => {
    try {
        const { project, revision, token, remarks, globalNote } = req.body || {};
        if (!_verifyArasuToken(token)) {
            return res.status(401).json({ success: false, error: 'Unauthorized — invalid or expired token' });
        }
        if (!project || !revision) {
            return res.status(400).json({ success: false, error: 'Missing project or revision' });
        }
        // Require at least one non-empty remark
        const remarkObj = remarks && typeof remarks === 'object' ? remarks : {};
        const cleaned = {};
        Object.entries(remarkObj).forEach(([k, v]) => {
            if (v && String(v).trim()) cleaned[k] = String(v).trim();
        });
        if (Object.keys(cleaned).length === 0 && !(globalNote && String(globalNote).trim())) {
            return res.status(400).json({ success: false, error: 'Add at least one remark before requesting changes' });
        }
        const state = await readApprovalsState();
        const entry = state[project] && state[project].revisions[revision];
        if (!entry) {
            return res.status(404).json({ success: false, error: 'Revision not found' });
        }
        if (entry.status !== 'pending') {
            return res.status(400).json({ success: false, error: `Cannot request changes — current status is "${entry.status}"` });
        }
        entry.status = 'changes_requested';
        entry.decidedAt = new Date().toISOString();
        entry.remarks = cleaned;
        entry.globalNote = globalNote && String(globalNote).trim() ? String(globalNote).trim() : null;
        if (state[project].latestPending === revision) state[project].latestPending = null;
        await writeApprovalsState(state);

        // Email costing team with remarks
        const cta = buildDeepLink('pending', project, revision);
        const html = buildEmailHtml({
            title: '✎ Costing Estimation — Changes Requested',
            intro: `The reviewer has requested changes on the costing estimation. Please review the remarks below and open a new revision to address them.`,
            infoRows: [
                ['Client', entry.clientName || '—'],
                ['Project', entry.projectName || '—'],
                ['Quote Ref', entry.ref || '—'],
                ['Revision', revision],
                ['Decided at', entry.decidedAt]
            ],
            paramRemarks: cleaned,
            ctaUrl: cta,
            ctaLabel: '📋 Open in App'
        });
        sendEmail({ to: COSTING_TEAM_EMAIL, subject: `[Changes Requested] ${entry.clientName || ''} · ${entry.projectName || ''} · ${revision}`, htmlBody: html });

        res.json({ success: true, status: 'changes_requested', project, revision, remarks: cleaned });
    } catch (error) {
        console.error('request-changes error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// GET /api/list-pending-approvals
// Returns all revisions across all projects with status="pending" (for Arasu's page).
app.get('/api/list-pending-approvals', async (req, res) => {
    try {
        const state = await readApprovalsState();
        const items = [];
        Object.entries(state).forEach(([project, entry]) => {
            if (!entry || !entry.revisions) return;
            Object.entries(entry.revisions).forEach(([rev, data]) => {
                if (data && data.status === 'pending') {
                    items.push({ project, revision: rev, ...data });
                }
            });
        });
        items.sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
        res.json({ success: true, items });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// GET /api/list-changes-requested
// Returns all revisions across all projects with status="changes_requested" (for Pending Estimates page).
app.get('/api/list-changes-requested', async (req, res) => {
    try {
        const state = await readApprovalsState();
        const items = [];
        Object.entries(state).forEach(([project, entry]) => {
            if (!entry || !entry.revisions) return;
            Object.entries(entry.revisions).forEach(([rev, data]) => {
                if (data && data.status === 'changes_requested') {
                    items.push({ project, revision: rev, ...data });
                }
            });
        });
        items.sort((a, b) => new Date(b.decidedAt) - new Date(a.decidedAt));
        res.json({ success: true, items });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// GET /api/approval-status?project=&revision=
// Look up the approval state for a specific revision (used to gate the Customer Quote button).
app.get('/api/approval-status', async (req, res) => {
    try {
        const { project, revision } = req.query || {};
        if (!project || !revision) {
            return res.status(400).json({ success: false, error: 'Missing project or revision' });
        }
        const state = await readApprovalsState();
        const entry = state[project] && state[project].revisions[revision];
        if (!entry) {
            return res.json({ success: true, found: false, status: null });
        }
        res.json({ success: true, found: true, status: entry.status, data: entry });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// GET /api/dashboard-data
// Aggregate metrics + per-project current status list (for the Dashboard page).
app.get('/api/dashboard-data', async (req, res) => {
    try {
        const state = await readApprovalsState();
        const now = new Date();
        const thirtyDaysAgo = new Date(now.getTime() - 30 * 86400_000);

        let pendingCount = 0, approvedCount = 0, changesCount = 0, archivedCount = 0;
        let approvedLast30 = 0;
        const projects = [];

        Object.entries(state).forEach(([projectFolder, entry]) => {
            if (!entry || !entry.revisions) return;
            const revs = Object.entries(entry.revisions);
            // Latest revision = most recently submitted
            revs.sort((a, b) => new Date(b[1].submittedAt || 0) - new Date(a[1].submittedAt || 0));
            const latest = revs[0];
            let projectStatus = 'unknown';
            let lastSubmitted = null;
            let lastDecided = null;
            let clientName = '', projectName = '', ref = '';
            if (latest) {
                projectStatus = latest[1].status || 'unknown';
                lastSubmitted = latest[1].submittedAt || null;
                lastDecided = latest[1].decidedAt || null;
                clientName = latest[1].clientName || '';
                projectName = latest[1].projectName || '';
                ref = latest[1].ref || '';
            }
            revs.forEach(([_, r]) => {
                if (r.status === 'pending') pendingCount++;
                else if (r.status === 'approved') {
                    approvedCount++;
                    if (r.decidedAt && new Date(r.decidedAt) >= thirtyDaysAgo) approvedLast30++;
                }
                else if (r.status === 'changes_requested') changesCount++;
                else if (r.status === 'archived') archivedCount++;
            });
            projects.push({
                project: projectFolder,
                clientName, projectName, ref,
                latestRevision: latest ? latest[0] : null,
                status: projectStatus,
                lastSubmitted, lastDecided,
                revisionCount: revs.length
            });
        });

        projects.sort((a, b) => new Date(b.lastSubmitted || 0) - new Date(a.lastSubmitted || 0));

        res.json({
            success: true,
            metrics: {
                pending: pendingCount,
                changesRequested: changesCount,
                approved: approvedCount,
                approvedLast30Days: approvedLast30,
                archived: archivedCount,
                totalProjects: projects.length
            },
            projects
        });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// GET /api/project-history?project=
// Full revision history for one project (used by Dashboard click-through).
app.get('/api/project-history', async (req, res) => {
    try {
        const { project } = req.query || {};
        if (!project) return res.status(400).json({ success: false, error: 'Missing project' });
        const state = await readApprovalsState();
        const entry = state[project];
        if (!entry || !entry.revisions) {
            return res.json({ success: true, project, revisions: [] });
        }
        const list = Object.entries(entry.revisions)
            .map(([rev, data]) => ({ revision: rev, ...data }))
            .sort((a, b) => new Date(b.submittedAt || 0) - new Date(a.submittedAt || 0));
        res.json({ success: true, project, revisions: list });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// GET /api/approval-selftest
// Diagnostic endpoint — runs each step of the approval pipeline and reports what works/fails.
// Hit this in browser after deploy to see which capability is broken.
app.get('/api/approval-selftest', async (req, res) => {
    const results = {};
    const captureErr = (e) => ({
        message: e && e.message,
        statusCode: e && e.statusCode,
        code: e && e.code,
        body: e && (e.body || e.responseBody || null)
    });

    // 1. Env vars
    results.env = {
        TARGET_USER_EMAIL: !!process.env.TARGET_USER_EMAIL,
        ARASU_APPROVAL_PASSWORD: !!process.env.ARASU_APPROVAL_PASSWORD,
        APP_BASE_URL: process.env.APP_BASE_URL || '(default)',
        APPROVER_EMAIL: APPROVER_EMAIL,
        COSTING_TEAM_EMAIL: COSTING_TEAM_EMAIL
    };

    // 2. Can we read approvals.json (or get a clean 404 = file doesn't exist yet, which is fine)?
    try {
        const state = await readApprovalsState();
        results.readApprovalsState = { ok: true, projectCount: Object.keys(state).length };
    } catch (e) {
        results.readApprovalsState = { ok: false, error: captureErr(e) };
    }

    // 3. Can we write to approvals.json? Round-trip a probe key, then remove it.
    try {
        const state = await readApprovalsState();
        const probeKey = '__selftest_probe__';
        state[probeKey] = { latestPending: null, revisions: {}, ts: new Date().toISOString() };
        await writeApprovalsState(state);
        // Remove the probe so it doesn't pollute real data
        delete state[probeKey];
        await writeApprovalsState(state);
        results.writeApprovalsState = { ok: true };
    } catch (e) {
        results.writeApprovalsState = { ok: false, error: captureErr(e) };
    }

    // 4. Can we send a test email? (Only fires if ?sendEmail=1 to avoid spamming arasu.)
    if (req.query.sendEmail === '1') {
        // Call Graph sendMail INLINE here so we can surface the full error in the response
        // (the sendEmail() helper logs to console but swallows error details).
        try {
            const targetEmail = process.env.TARGET_USER_EMAIL;
            if (!targetEmail) {
                results.sendEmail = { ok: false, error: 'TARGET_USER_EMAIL not set' };
            } else {
                const message = {
                    message: {
                        subject: '[selftest] Approval workflow diagnostic',
                        body: { contentType: 'HTML', content: '<p>This is a diagnostic email from the approval selftest endpoint. If you received it, Graph sendMail is working.</p>' },
                        toRecipients: [{ emailAddress: { address: APPROVER_EMAIL } }]
                    },
                    saveToSentItems: true
                };
                try {
                    await graphClient
                        .api(`/users/${targetEmail}/sendMail`)
                        .post(message);
                    results.sendEmail = { ok: true, note: `email sent to ${APPROVER_EMAIL}` };
                } catch (e) {
                    results.sendEmail = {
                        ok: false,
                        sender: targetEmail,
                        recipient: APPROVER_EMAIL,
                        error: captureErr(e),
                        // Stringify the whole error to catch anything not on the standard fields
                        rawErrorString: String(e)
                    };
                }
            }
        } catch (outer) {
            results.sendEmail = { ok: false, outerError: captureErr(outer) };
        }
    } else {
        results.sendEmail = { skipped: true, note: 'add ?sendEmail=1 to actually send a test email to arasu@metfraa.com' };
    }

    res.json(results);
});

app.get('/api/list-quotations', async (req, res) => {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        // Walk per-project folders under Generated_Costings/, collecting JSON files from each
        const rootPath = ':/Metfraa_Costing_App/Generated_Costings:';
        let rootListing;
        try {
            rootListing = await graphClient
                .api(`/users/${targetEmail}/drive/root${rootPath}/children`)
                .top(500)
                .get();
        } catch (listErr) {
            return res.json({ success: true, quotations: [] });
        }
        const projectFolders = (rootListing.value || []).filter(f =>
            f.folder && f.name && !f.name.startsWith('_')
        );
        const quotations = [];
        for (const pf of projectFolders) {
            try {
                const jl = await graphClient
                    .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/${pf.name}/JSON:/children`)
                    .top(200)
                    .get();
                (jl.value || []).forEach(f => {
                    if (f.file && f.name && f.name.toLowerCase().endsWith('.json')) {
                        quotations.push({
                            name: f.name.replace(/\.json$/i, ''),
                            project: pf.name,
                            lastModified: f.lastModifiedDateTime,
                            size: f.size
                        });
                    }
                });
            } catch (_) { /* JSON subfolder might not exist for some project */ }
        }
        // Also include legacy flat /JSON/ files if any exist (back-compat)
        try {
            const flatList = await graphClient
                .api(`/users/${targetEmail}/drive/root:/Metfraa_Costing_App/Generated_Costings/JSON:/children`)
                .top(200)
                .get();
            (flatList.value || []).forEach(f => {
                if (f.file && f.name && f.name.toLowerCase().endsWith('.json')) {
                    quotations.push({
                        name: f.name.replace(/\.json$/i, ''),
                        project: null, // legacy flat
                        lastModified: f.lastModifiedDateTime,
                        size: f.size
                    });
                }
            });
        } catch (_) { /* no legacy folder, fine */ }

        quotations.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));
        res.json({ success: true, quotations });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// 5. ROUTE: Update Master Data (rate editor write-back)
// Body: { updates: [ { id, rate }, ... ], adminPassword? }
// =============================================================
app.post('/api/update-master-data', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        const { updates } = req.body;
        if (!Array.isArray(updates) || updates.length === 0) {
            return res.status(400).json({ success: false, error: 'No updates provided' });
        }

        const targetEmail = process.env.TARGET_USER_EMAIL;
        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = pickDataSheet(workbook);
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);

        const updateMap = {};
        updates.forEach(u => { updateMap[String(u.id)] = Number(u.rate); });

        // Detect which rate column the sheet uses; write back to the same one
        const sampleRow = rawData[0] || {};
        let rateCol = 'Rate';
        if (sampleRow['Rate'] !== undefined) rateCol = 'Rate';
        else if (sampleRow['Rate (₹)'] !== undefined) rateCol = 'Rate (₹)';
        else if (sampleRow['Price'] !== undefined) rateCol = 'Price';

        let changed = 0;
        rawData.forEach(row => {
            const id = String(row['Option ID']);
            if (id in updateMap) {
                row[rateCol] = updateMap[id];
                changed++;
            }
        });

        if (changed === 0) {
            return res.status(400).json({ success: false, error: 'No matching Option IDs found in master sheet' });
        }

        // Preserve original column order
        const headerRow = xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || [];
        const newSheet = xlsx.utils.json_to_sheet(rawData, { header: headerRow });
        workbook.Sheets[sheetName] = newSheet;

        await writeWorkbookToOneDrive(workbook);

        res.json({ success: true, changed });
    } catch (error) {
        console.error('Update Master Data Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// 6. Status / Health Check
// =============================================================
// =============================================================
// 6. ROUTE: Add a new Paint Type row to the master sheet
// Body: { name, rate, adminPassword? }
// =============================================================
app.post('/api/add-paint-type', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        const { name, rate } = req.body;
        if (!name || typeof name !== 'string' || !name.trim()) {
            return res.status(400).json({ success: false, error: 'Paint name required' });
        }
        const numRate = Number(rate);
        if (!Number.isFinite(numRate) || numRate < 0) {
            return res.status(400).json({ success: false, error: 'Rate must be a non-negative number' });
        }

        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = pickDataSheet(workbook);
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);

        // Check duplicate name (case-insensitive)
        const trimmedName = name.trim();
        const exists = rawData.some(r =>
            String(r['Parameter Name'] || '').trim().toLowerCase() === 'paint' &&
            String(r['Option Name'] || '').trim().toLowerCase() === trimmedName.toLowerCase()
        );
        if (exists) {
            return res.status(400).json({ success: false, error: `A paint type called "${trimmedName}" already exists` });
        }

        // Generate next Paint ID (PAINTnnn)
        const existingIds = rawData
            .filter(r => String(r['Parameter Name'] || '').trim().toLowerCase() === 'paint')
            .map(r => String(r['Option ID'] || ''));
        let n = 1;
        while (existingIds.includes(`PAINT${String(n).padStart(3, '0')}`)) n++;
        const newId = `PAINT${String(n).padStart(3, '0')}`;

        rawData.push({
            'Parameter Name': 'Paint',
            'Subcategory': 'Paint Type',
            'Option ID': newId,
            'Option Name': trimmedName,
            'Rate': numRate,
            'Type': 'RATE',
            'Unit': 'MT',
            'Min': '',
            'Max': ''
        });

        // Preserve column order
        let header = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);
        REQUIRED_COLUMNS.forEach(col => { if (!header.includes(col)) header.push(col); });
        const newSheet = xlsx.utils.json_to_sheet(rawData, { header });
        workbook.Sheets[sheetName] = newSheet;

        await writeWorkbookToOneDrive(workbook);
        res.json({ success: true, id: newId, name: trimmedName, rate: numRate });
    } catch (error) {
        console.error('Add Paint Type Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// 7. ROUTE: Delete a Paint Type row from the master sheet
// Body: { id, adminPassword? }
// =============================================================
app.post('/api/delete-paint-type', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        const { id } = req.body;
        if (!id) return res.status(400).json({ success: false, error: 'Paint Option ID required' });

        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = pickDataSheet(workbook);
        const sheet = workbook.Sheets[sheetName];
        let rawData = xlsx.utils.sheet_to_json(sheet);

        const before = rawData.length;
        rawData = rawData.filter(r => !(
            String(r['Parameter Name'] || '').trim().toLowerCase() === 'paint' &&
            String(r['Option ID'] || '') === String(id)
        ));
        if (rawData.length === before) {
            return res.status(404).json({ success: false, error: `No Paint row with Option ID "${id}" found` });
        }

        let header = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);
        REQUIRED_COLUMNS.forEach(col => { if (!header.includes(col)) header.push(col); });
        const newSheet = xlsx.utils.json_to_sheet(rawData, { header });
        workbook.Sheets[sheetName] = newSheet;

        await writeWorkbookToOneDrive(workbook);
        res.json({ success: true, deletedId: id });
    } catch (error) {
        console.error('Delete Paint Type Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// 8. ROUTE: Add a generic row to the master sheet
// Body: { paramName, subcategory, optionName, rate, type ('RATE'|'PERCENT'), unit, adminPassword? }
// Used by the Step 2 "+ Add row (save to master)" flow.
// =============================================================
app.post('/api/add-master-row', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        let { paramName, subcategory, optionName, rate, type, unit, group, remark } = req.body;
        paramName = String(paramName || '').trim();
        subcategory = String(subcategory || '').trim();
        optionName = String(optionName || '').trim();
        type = String(type || 'RATE').toUpperCase().trim();
        unit = String(unit || 'MT').trim();
        group = String(group || '').trim();
        remark = String(remark || '').trim();
        const numRate = Number(rate);

        if (!paramName) return res.status(400).json({ success: false, error: 'Parameter name required' });
        if (!subcategory) return res.status(400).json({ success: false, error: 'Subcategory required' });
        if (!optionName) return res.status(400).json({ success: false, error: 'Option name required' });
        if (!Number.isFinite(numRate)) return res.status(400).json({ success: false, error: 'Rate must be a number' });
        if (type !== 'RATE' && type !== 'PERCENT') return res.status(400).json({ success: false, error: 'Type must be RATE or PERCENT' });

        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = pickDataSheet(workbook);
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);

        // Generate a unique Option ID. Use the param's first 4 chars + disambiguated
        // 4-letter sub slug (FIXC/FIXD/FIXR/FIXT/FIXG distinguish the four Fix- subcats).
        const cleanP = paramName.replace(/^\d+\.\s*/, '').replace(/[^A-Z0-9]/gi, '').toUpperCase().slice(0, 4) || 'CUST';
        const cleanS = subSlug(subcategory);
        const allIds = new Set(rawData.map(r => String(r['Option ID'] || '')));
        let n = 1;
        let newId;
        do {
            newId = `${cleanP}-${cleanS}-${String(n).padStart(2, '0')}`;
            n++;
        } while (allIds.has(newId) && n < 1000);

        rawData.push({
            'Parameter Name': paramName,
            'Subcategory': subcategory,
            'Option ID': newId,
            'Option Name': optionName,
            'Rate': numRate,
            'Unit': unit,
            'Type': type,
            'Min': '',
            'Max': '',
            'Group': group,
            'Remark': remark
        });

        let header = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);
        REQUIRED_COLUMNS.forEach(col => { if (!header.includes(col)) header.push(col); });
        const newSheet = xlsx.utils.json_to_sheet(rawData, { header });
        workbook.Sheets[sheetName] = newSheet;

        await writeWorkbookToOneDrive(workbook);
        res.json({ success: true, id: newId, paramName, subcategory, optionName, rate: numRate, type, unit });
    } catch (error) {
        console.error('Add Master Row Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: Add a whole new parameter
// Body: { paramName, unit, subcategories: [list of subcat names], adminPassword? }
// Creates one placeholder row per chosen subcategory. AUTO subcats
// (Transport Outward, Finishing, Fixing / Erection) get Type=AUTO.
// Inserted before the first global-rate parameter so it lands at end of materials.
// =============================================================
const AUTO_SUBCATS = new Set(['TRANSPORT OUTWARD', 'FINISHING', 'FIXING / ERECTION']);
const GLOBAL_PARAM_MARKERS = ['distance', 'paint', 'erection safety', 'erection height'];

// Disambiguating 4-letter slugs so that subcategories starting with "Fix..."
// don't all collide on the same prefix. Used when generating Option IDs.
const SUB_SLUG_4 = {
    'RM COST': 'RMCO',
    'WASTAGES / OVERLAP': 'WAST',
    'TRANSPORT OUTWARD': 'TRAN',
    'FIXED CHARGES (CONVERSION)': 'FIXC',
    'FIXED CHARGES (DESIGN)':     'FIXD',
    'SHOT BLAST':                 'SHOT',
    'FINISHING':                  'FINI',
    'FIXING / ERECTION':          'FIXR',
    'FIXTURES / HARDWARE':        'FIXT',
    'FIXING CHARGES':             'FIXG',
    'ACCESSORIES':                'ACCE',
    'DISTANCE BAND':              'DBND',
    'PAINT TYPE':                 'PNT',
    'SAFETY TYPE':                'SAFE',
    'HEIGHT BAND':                'HBND',
};

function subSlug(sub) {
    const key = String(sub || '').trim().toUpperCase();
    if (SUB_SLUG_4[key]) return SUB_SLUG_4[key];
    // Fallback: first 4 alnum chars
    return (String(sub || '').replace(/[^A-Z0-9]/gi, '').toUpperCase().slice(0, 4)) || 'XSUB';
}

function isGlobalParamName(name) {
    const s = String(name || '').trim().toLowerCase().replace(/^\d+\.\s*/, '');
    return GLOBAL_PARAM_MARKERS.some(m => s === m || s.startsWith(m));
}

app.post('/api/add-parameter', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        let { paramName, unit, subcategories, options, roundingUnit } = req.body;
        paramName = String(paramName || '').trim();
        unit = String(unit || 'MT').trim();
        // Rounding Unit: must be 10 or 100. Default 100 if not specified.
        const ru = Number(roundingUnit);
        const finalRounding = (ru === 10 || ru === 100) ? ru : 100;
        if (!Array.isArray(subcategories)) subcategories = [];
        subcategories = subcategories.map(s => String(s || '').trim()).filter(Boolean);
        // options is an optional map: { "RM COST": [{name, rate, type, group, remark}, ...], ... }
        if (!options || typeof options !== 'object') options = {};

        if (!paramName) return res.status(400).json({ success: false, error: 'Parameter name required' });
        if (subcategories.length === 0) return res.status(400).json({ success: false, error: 'Pick at least one subcategory' });

        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = pickDataSheet(workbook);
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);

        const stripNum = s => String(s || '').trim().toLowerCase().replace(/^\d+\.\s*/, '');
        const existingNames = new Set(rawData.map(r => stripNum(r['Parameter Name'])));
        if (existingNames.has(stripNum(paramName))) {
            return res.status(400).json({ success: false, error: 'A parameter with that name already exists' });
        }

        let maxNum = 0;
        rawData.forEach(r => {
            const m = String(r['Parameter Name'] || '').match(/^(\d+)\./);
            if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
        });
        const newPrefix = `${maxNum + 1}. `;
        const fullParamName = newPrefix + paramName.replace(/^\d+\.\s*/, '');

        const cleanP = paramName.replace(/^\d+\.\s*/, '').replace(/[^A-Z0-9]/gi, '').toUpperCase().slice(0, 4) || 'NEWP';

        // Build new rows: one per option for each subcategory.
        // If no options provided for a subcat, write one empty placeholder row (legacy behaviour).
        const newRows = [];
        subcategories.forEach(sub => {
            const isAuto = AUTO_SUBCATS.has(sub.toUpperCase().trim());
            const s3 = subSlug(sub);
            const subOpts = Array.isArray(options[sub]) ? options[sub].filter(o => o && (o.name || o.rate)) : [];

            if (isAuto) {
                // AUTO subcats always get one marker row, options are ignored
                newRows.push({
                    'Parameter Name': fullParamName,
                    'Subcategory': sub,
                    'Option ID': `${cleanP}-${s3}-01`,
                    'Option Name': '(auto from Step 1)',
                    'Rate': 0,
                    'Unit': unit,
                    'Type': 'AUTO',
                    'Min': '', 'Max': '',
                    'Group': '',
                    'Remark': 'Filled from Step 1 selection',
                    'Rounding Unit': finalRounding
                });
                return;
            }

            if (subOpts.length === 0) {
                // Empty placeholder
                newRows.push({
                    'Parameter Name': fullParamName,
                    'Subcategory': sub,
                    'Option ID': `${cleanP}-${s3}-01`,
                    'Option Name': '',
                    'Rate': 0,
                    'Unit': unit,
                    'Type': 'RATE',
                    'Min': '', 'Max': '',
                    'Group': '',
                    'Remark': '',
                    'Rounding Unit': finalRounding
                });
                return;
            }

            subOpts.forEach((o, i) => {
                const rawType = String(o.type || 'RATE').toUpperCase().trim();
                const type = (rawType === 'PERCENT' || rawType === 'PERCENTAGE' || rawType === '%') ? 'PERCENT' : 'RATE';
                newRows.push({
                    'Parameter Name': fullParamName,
                    'Subcategory': sub,
                    'Option ID': `${cleanP}-${s3}-${String(i+1).padStart(2,'0')}`,
                    'Option Name': String(o.name || '').trim(),
                    'Rate': Number(o.rate) || 0,
                    'Unit': unit,
                    'Type': type,
                    'Min': '', 'Max': '',
                    'Group': String(o.group || '').trim(),
                    'Remark': String(o.remark || '').trim(),
                    'Rounding Unit': finalRounding
                });
            });
        });

        let insertIdx = rawData.length;
        for (let i = 0; i < rawData.length; i++) {
            if (isGlobalParamName(rawData[i]['Parameter Name'])) { insertIdx = i; break; }
        }
        const updated = [...rawData.slice(0, insertIdx), ...newRows, ...rawData.slice(insertIdx)];

        let header = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);
        REQUIRED_COLUMNS.forEach(col => { if (!header.includes(col)) header.push(col); });
        workbook.Sheets[sheetName] = xlsx.utils.json_to_sheet(updated, { header });

        await writeWorkbookToOneDrive(workbook);
        res.json({ success: true, paramName: fullParamName, subcategories, rowsAdded: newRows.length });
    } catch (error) {
        console.error('Add Parameter Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: Update an existing option (full field edit except ID)
// Body: { id, paramName?, subcategory?, optionName?, rate?, type?, unit?, group?, remark?, min?, max? }
// Only fields present in the body are updated. id is the lookup key (immutable).
// =============================================================
app.post('/api/update-option', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        const { id } = req.body;
        if (!id) return res.status(400).json({ success: false, error: 'Option ID required' });

        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = pickDataSheet(workbook);
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);

        const rowIdx = rawData.findIndex(r => String(r['Option ID']) === String(id));
        if (rowIdx === -1) return res.status(404).json({ success: false, error: `Option ID "${id}" not found` });

        const fieldMap = {
            paramName:    'Parameter Name',
            subcategory:  'Subcategory',
            optionName:   'Option Name',
            rate:         'Rate',
            type:         'Type',
            unit:         'Unit',
            group:        'Group',
            remark:       'Remark',
            min:          'Min',
            max:          'Max'
        };
        Object.entries(fieldMap).forEach(([bodyKey, colName]) => {
            if (req.body[bodyKey] === undefined) return;
            let v = req.body[bodyKey];
            if (bodyKey === 'rate') v = Number(v);
            if (bodyKey === 'type') v = String(v).toUpperCase().trim();
            if (bodyKey === 'min' || bodyKey === 'max') v = (v === '' || v === null) ? '' : Number(v);
            if (typeof v === 'string') v = v.trim();
            rawData[rowIdx][colName] = v;
        });

        let header = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);
        REQUIRED_COLUMNS.forEach(col => { if (!header.includes(col)) header.push(col); });
        workbook.Sheets[sheetName] = xlsx.utils.json_to_sheet(rawData, { header });

        await writeWorkbookToOneDrive(workbook);
        res.json({ success: true, id, row: rawData[rowIdx] });
    } catch (error) {
        console.error('Update Option Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

// =============================================================
// ROUTE: Delete an option by ID
// Body: { id }
// =============================================================
app.post('/api/delete-option', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        const { id } = req.body;
        if (!id) return res.status(400).json({ success: false, error: 'Option ID required' });

        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = pickDataSheet(workbook);
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);

        const rowIdx = rawData.findIndex(r => String(r['Option ID']) === String(id));
        if (rowIdx === -1) return res.status(404).json({ success: false, error: `Option ID "${id}" not found` });

        const removed = rawData.splice(rowIdx, 1)[0];

        let header = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);
        REQUIRED_COLUMNS.forEach(col => { if (!header.includes(col)) header.push(col); });
        workbook.Sheets[sheetName] = xlsx.utils.json_to_sheet(rawData, { header });

        await writeWorkbookToOneDrive(workbook);
        res.json({ success: true, id, removed });
    } catch (error) {
        console.error('Delete Option Error:', error.message);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.get('/api/status', (req, res) => res.send("OK"));
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

// JSON 404 catch-all for /api/* (so frontend never gets HTML for a wrong path)
app.use('/api', (req, res) => {
    res.status(404).json({ success: false, error: 'API endpoint not found', path: req.path });
});

// =============================================================
// Static frontend — serves /public, falls back to index.html
// =============================================================
app.use(express.static(path.join(__dirname, 'public')));
app.get(/.*/, (req, res, next) => {
    if (req.path.startsWith('/api/')) return next();
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log('//////////////////////////////////////////////////////////');
    console.log('🚀 METFRAA COSTING SERVER RUNNING');
    console.log(`🌐 Listening on port ${PORT}`);
    console.log(`🔐 Admin password protection: ${ADMIN_PASSWORD ? 'ON' : 'OFF (open mode)'}`);
    console.log('//////////////////////////////////////////////////////////');
});
