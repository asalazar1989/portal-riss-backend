const axios = require('axios');

const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const SHAREPOINT_SITE_URL = process.env.SHAREPOINT_SITE_URL;
const EXCEL_FILE_PATH = process.env.EXCEL_FILE_PATH;
const EXCEL_SHEET_NAME = process.env.EXCEL_SHEET_NAME;

async function getAccessToken() {
    const tokenEndpoint = `https://accounts.accesscontrol.windows.net/${AZURE_TENANT_ID}/tokens/OAuth/2`;
    
    const resource = `00000003-0000-0ff1-ce00-000000000000/${SHAREPOINT_SITE_URL.replace('https://', '')}@${AZURE_TENANT_ID}`;
    
    const params = new URLSearchParams();
    params.append('grant_type', 'client_credentials');
    params.append('client_id', `${AZURE_CLIENT_ID}@${AZURE_TENANT_ID}`);
    params.append('client_secret', AZURE_CLIENT_SECRET);
    params.append('resource', resource);
    
    try {
        const response = await axios.post(tokenEndpoint, params);
        return response.data.access_token;
    } catch (error) {
        console.error('Error obteniendo token:', error.response?.data || error.message);
        throw new Error('Error de autenticación con SharePoint');
    }
}

async function readExcelData(token) {
    const apiUrl = `${SHAREPOINT_SITE_URL}/_api/web/GetFileByServerRelativeUrl('${EXCEL_FILE_PATH}')/ListItemAllFields`;
    
    try {
        const response = await axios.get(apiUrl, {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Accept': 'application/json;odata=verbose'
            }
        });
        return response.data;
    } catch (error) {
        console.error('Error leyendo Excel:', error.response?.data || error.message);
        throw new Error('Error leyendo datos del Excel');
    }
}

module.exports = async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    try {
        const token = await getAccessToken();
        const { action } = req.query;

        switch (action) {
            case 'read':
                const data = await readExcelData(token);
                return res.status(200).json({ success: true, data });

            default:
                return res.status(404).json({ success: false, error: 'Acción no encontrada' });
        }
    } catch (error) {
        console.error('Error en function:', error);
        return res.status(500).json({ success: false, error: error.message });
    }
};
