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
const REQUIRED_COLUMNS = ['Parameter Name','Subcategory','Option ID','Option Name','Rate','Type','Unit','Min','Max'];

// Write rows back to OneDrive, preserving column order
async function writeWorkbookToOneDrive(workbook) {
    const targetEmail = process.env.TARGET_USER_EMAIL;
    const outBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    await graphClient
        .api(`/users/${targetEmail}/drive/root${MASTER_FILE_PATH}/content`)
        .put(outBuffer);
}

// Returns true if the workbook was modified
async function bootstrapMissingDefaults(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(sheet);
    const headerRow = (xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || []).map(String);

    const existingParams = new Set(rawData.map(r => String(r['Parameter Name'] || '').trim().toLowerCase()));
    let changed = false;

    if (!existingParams.has('distance')) {
        rawData.push(...BOOTSTRAP_DISTANCE);
        changed = true;
        console.log('🔧 Bootstrapping: added 8 Distance band rows');
    }
    if (!existingParams.has('paint')) {
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

        // Bootstrap defaults (Distance, Paint) if missing
        const bootstrapped = await bootstrapMissingDefaults(workbook, sheetName);
        if (bootstrapped) {
            await writeWorkbookToOneDrive(workbook);
            // Re-download so subsequent reads see the persisted version
            buffer = await downloadMasterBuffer();
            workbook = xlsx.read(buffer, { type: 'buffer' });
            sheetName = pickDataSheet(workbook);
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

            const opt = {
                id: row['Option ID'],
                name: row['Option Name'],
                rate: parseFloat(rate) || 0,
                type: (type === 'PERCENT' || type === 'PERCENTAGE' || type === '%') ? 'PERCENT' : 'RATE'
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
        const { projectName, date, pdfBase64, jsonData } = req.body;
        if (!projectName || !date) {
            return res.status(400).json({ success: false, error: 'Missing projectName or date' });
        }
        if (!pdfBase64 && !jsonData) {
            return res.status(400).json({ success: false, error: 'Provide pdfBase64 or jsonData (or both)' });
        }

        const targetEmail = process.env.TARGET_USER_EMAIL;
        const safeName = `${sanitizeName(projectName)}_${sanitizeName(date)}`;
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

        const outBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        await graphClient
            .api(`/users/${targetEmail}/drive/root${MASTER_FILE_PATH}/content`)
            .put(outBuffer);

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
