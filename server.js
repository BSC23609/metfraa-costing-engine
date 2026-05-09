/**
 * METFRAA PEB Costing Engine — backend
 * Serves the frontend from /public and exposes the OneDrive-backed API.
 */
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

// Optional: ADMIN_PASSWORD protects the rate editor. If unset, editor is open.
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

// ---------- Azure / Graph ----------
const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
});
const graphClient = Client.initWithMiddleware({ authProvider });

const MASTER_PATH = ':/Metfraa_Costing_App/Master_Data/Master_Cost_DB.xlsx';
const SAVE_FOLDER = ':/Metfraa_Costing_App/Generated_Costings/';

// ---------- helpers ----------
async function downloadMasterBuffer() {
    const targetEmail = process.env.TARGET_USER_EMAIL;
    const fileMeta = await graphClient
        .api(`/users/${targetEmail}/drive/root${MASTER_PATH}`)
        .get();
    const response = await fetch(fileMeta['@microsoft.graph.downloadUrl']);
    return Buffer.from(await response.arrayBuffer());
}

function parseRate(row) {
    return row['Rate'] || row['rate'] || row['Rate (₹)'] || row['Price'] || 0;
}

function parseRange(row) {
    const min = row['Min'] !== undefined && row['Min'] !== '' ? Number(row['Min']) : null;
    const max = row['Max'] !== undefined && row['Max'] !== '' ? Number(row['Max']) : null;
    return { min, max };
}

// Simple admin gate. Returns true if request is authorised.
function isAdmin(req) {
    if (!ADMIN_PASSWORD) return true; // open mode
    const provided = req.headers['x-admin-password'] || req.body?.adminPassword || '';
    return provided === ADMIN_PASSWORD;
}

// ---------- API ----------
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

// Lets the frontend know whether to show the password prompt
app.get('/api/auth-mode', (req, res) => {
    res.json({ adminProtected: !!ADMIN_PASSWORD });
});

app.post('/api/admin-check', (req, res) => {
    if (!ADMIN_PASSWORD) return res.json({ ok: true });
    res.json({ ok: req.body?.password === ADMIN_PASSWORD });
});

app.get('/api/master-data', async (req, res) => {
    try {
        res.set('Cache-Control', 'no-store, no-cache, must-revalidate');
        res.set('Pragma', 'no-cache');
        res.set('Expires', '0');

        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const rawData = xlsx.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

        const formattedData = {};
        rawData.forEach(row => {
            const paramName = row['Parameter Name'];
            const subcat = row['Subcategory'];
            if (!paramName || !subcat) return;

            if (!formattedData[paramName]) formattedData[paramName] = {};
            if (!formattedData[paramName][subcat]) formattedData[paramName][subcat] = [];

            const { min, max } = parseRange(row);
            const opt = {
                id: row['Option ID'],
                name: row['Option Name'],
                rate: parseRate(row)
            };
            if (min !== null) opt.min = min;
            if (max !== null) opt.max = max;

            formattedData[paramName][subcat].push(opt);
        });
        res.json(formattedData);
    } catch (error) {
        console.error('master-data error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/save-quotation', async (req, res) => {
    try {
        const { fileName, htmlContent } = req.body;
        if (!fileName || !htmlContent) return res.status(400).json({ error: 'Missing fileName or htmlContent' });

        const targetEmail = process.env.TARGET_USER_EMAIL;
        const safeName = String(fileName).replace(/[\/\\?%*:|"<>]/g, '-');
        const fullPath = `${SAVE_FOLDER}${safeName}.html`;

        await graphClient
            .api(`/users/${targetEmail}/drive/root${fullPath}/content`)
            .put(htmlContent);

        res.json({ success: true, savedAs: `${safeName}.html` });
    } catch (error) {
        console.error('save-quotation error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/update-master-data', async (req, res) => {
    try {
        if (!isAdmin(req)) return res.status(401).json({ error: 'Unauthorised' });

        const { updates } = req.body;
        if (!Array.isArray(updates) || updates.length === 0) {
            return res.status(400).json({ error: 'No updates provided' });
        }

        const targetEmail = process.env.TARGET_USER_EMAIL;
        const buffer = await downloadMasterBuffer();
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rawData = xlsx.utils.sheet_to_json(sheet);

        const updateMap = {};
        updates.forEach(u => { updateMap[String(u.id)] = Number(u.rate); });

        const sampleRow = rawData[0] || {};
        let rateCol = 'Rate';
        if (sampleRow['rate'] !== undefined) rateCol = 'rate';
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

        if (changed === 0) return res.status(400).json({ error: 'No matching Option IDs found' });

        const headerRow = xlsx.utils.sheet_to_json(sheet, { header: 1 })[0] || [];
        const newSheet = xlsx.utils.json_to_sheet(rawData, { header: headerRow });
        workbook.Sheets[sheetName] = newSheet;

        const outBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        await graphClient
            .api(`/users/${targetEmail}/drive/root${MASTER_PATH}/content`)
            .put(outBuffer);

        res.json({ success: true, changed });
    } catch (error) {
        console.error('update-master-data error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Legacy alias kept for compatibility
app.get('/api/status', (req, res) => res.send('OK'));

// ---------- static frontend ----------
app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api/')) return next();
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ---------- start ----------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`✅ Metfraa Costing Engine running on port ${PORT}`);
    console.log(`   Admin password protection: ${ADMIN_PASSWORD ? 'ON' : 'OFF (open mode)'}`);
});
