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
const REQUIRED_COLUMNS = ['Parameter Name','Subcategory','Option ID','Option Name','Rate','Type','Unit','Min','Max','Group','Remark'];

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
        // Filename = client_project_date (client first; project + date follow for uniqueness)
        const parts = [];
        if (clientName && String(clientName).trim() && String(clientName).trim() !== '—') {
            parts.push(sanitizeName(clientName));
        }
        parts.push(sanitizeName(projectName));
        parts.push(sanitizeName(date));
        const safeName = parts.join('_');
        const results = {};

        // Save PDF
        if (pdfBase64) {
            const cleanB64 = pdfBase64.replace(/^data:application\/pdf;base64,/, '');
            const pdfBuffer = Buffer.from(cleanB64, 'base64');
            const pdfPath = `:/Metfraa_Costing_App/Generated_Costings/PDF/${safeName}.pdf:`;
            const pdfResult = await graphClient
                .api(`/users/${targetEmail}/drive/root${pdfPath}/content`)
                .header('Content-Type', 'application/pdf')
                .put(pdfBuffer);
            results.pdf = { fileName: `${safeName}.pdf`, webUrl: pdfResult.webUrl || null };
        }

        // Save JSON
        if (jsonData) {
            const jsonString = typeof jsonData === 'string' ? jsonData : JSON.stringify(jsonData, null, 2);
            const jsonPath = `:/Metfraa_Costing_App/Generated_Costings/JSON/${safeName}.json:`;
            const jsonResult = await graphClient
                .api(`/users/${targetEmail}/drive/root${jsonPath}/content`)
                .header('Content-Type', 'application/json')
                .put(jsonString);
            results.json = { fileName: `${safeName}.json`, webUrl: jsonResult.webUrl || null };
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
// 3. ROUTE: Load Quotation by reference
// =============================================================
app.get('/api/load-quotation', async (req, res) => {
    try {
        const ref = req.query.ref;
        if (!ref) return res.status(400).json({ success: false, error: 'Missing ref query param' });

        const targetEmail = process.env.TARGET_USER_EMAIL;
        const safeRef = sanitizeName(ref);

        const folderPath = ':/Metfraa_Costing_App/Generated_Costings/JSON:';
        let listing;
        try {
            listing = await graphClient
                .api(`/users/${targetEmail}/drive/root${folderPath}/children`)
                .top(200)
                .get();
        } catch (listErr) {
            return res.status(404).json({ success: false, error: 'JSON folder not found in OneDrive. Save a quotation first.' });
        }

        const files = (listing.value || []).filter(f =>
            f.file && f.name && f.name.toLowerCase().endsWith('.json')
        );

        let matches = files.filter(f =>
            f.name.toLowerCase().includes(safeRef.toLowerCase())
        );

        if (matches.length === 0) {
            return res.status(404).json({
                success: false,
                error: `No saved quotation found matching "${ref}"`,
                searchedFor: safeRef,
                available: files.map(f => f.name)
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
app.get('/api/list-quotations', async (req, res) => {
    try {
        const targetEmail = process.env.TARGET_USER_EMAIL;
        const folderPath = ':/Metfraa_Costing_App/Generated_Costings/JSON:';
        let listing;
        try {
            listing = await graphClient
                .api(`/users/${targetEmail}/drive/root${folderPath}/children`)
                .top(200)
                .get();
        } catch (listErr) {
            return res.json({ success: true, quotations: [] });
        }
        const quotations = (listing.value || [])
            .filter(f => f.file && f.name && f.name.toLowerCase().endsWith('.json'))
            .map(f => ({
                name: f.name.replace(/\.json$/i, ''),
                lastModified: f.lastModifiedDateTime,
                size: f.size
            }))
            .sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));
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

        // Generate a unique Option ID. Use the param's first 4 chars + sub's first 2 chars + N
        const cleanP = paramName.replace(/^\d+\.\s*/, '').replace(/[^A-Z0-9]/gi, '').toUpperCase().slice(0, 4) || 'CUST';
        const cleanS = subcategory.replace(/[^A-Z0-9]/gi, '').toUpperCase().slice(0, 2) || 'XX';
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

function isGlobalParamName(name) {
    const s = String(name || '').trim().toLowerCase().replace(/^\d+\.\s*/, '');
    return GLOBAL_PARAM_MARKERS.some(m => s === m || s.startsWith(m));
}

app.post('/api/add-parameter', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ success: false, error: 'Unauthorised' });

        let { paramName, unit, subcategories, options } = req.body;
        paramName = String(paramName || '').trim();
        unit = String(unit || 'MT').trim();
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
            const s3 = sub.replace(/[^A-Z0-9]/gi, '').toUpperCase().slice(0, 3) || 'XX';
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
                    'Remark': 'Filled from Step 1 selection'
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
                    'Remark': ''
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
                    'Remark': String(o.remark || '').trim()
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
