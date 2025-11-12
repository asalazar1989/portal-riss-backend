const axios = require('axios');

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID = process.env.TENANT_ID;
const FILE_ID = process.env.FILE_ID;
const SHEET_NAME = process.env.SHEET_NAME;

async function getAccessToken() {
    const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    
    const params = new URLSearchParams();
    params.append('client_id', CLIENT_ID);
    params.append('client_secret', CLIENT_SECRET);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');
    
    try {
        const response = await axios.post(tokenEndpoint, params);
        return response.data.access_token;
    } catch (error) {
        console.error('Error obteniendo token:', error.response?.data || error.message);
        throw new Error('Error de autenticación');
    }
}

async function readExcelData(token) {
    const url = `https://graph.microsoft.com/v1.0/drives/b!986b0cbb-6874-48fe-9330-8d8e73408f07/items/${FILE_ID}/workbook/worksheets('${SHEET_NAME}')/usedRange`;
    
    try {
        const response = await axios.get(url, {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        return response.data.values;
    } catch (error) {
        console.error('Error leyendo Excel:', error.response?.data || error.message);
        throw new Error('Error leyendo datos del Excel');
    }
}

async function updateCell(token, rowIndex, columnIndex, value) {
    const cellAddress = columnToLetter(columnIndex) + (rowIndex + 1);
    const url = `https://graph.microsoft.com/v1.0/drives/b!986b0cbb-6874-48fe-9330-8d8e73408f07/items/${FILE_ID}/workbook/worksheets('${SHEET_NAME}')/range(address='${cellAddress}')`;
    
    try {
        await axios.patch(url, {
            values: [[value]]
        }, {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        return { success: true };
    } catch (error) {
        console.error('Error actualizando celda:', error.response?.data || error.message);
        throw new Error('Error actualizando Excel');
    }
}

function columnToLetter(column) {
    let temp, letter = '';
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
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

            case 'update':
                const { rowIndex, columnIndex, value } = req.body;
                if (rowIndex === undefined || columnIndex === undefined || value === undefined) {
                    return res.status(400).json({ 
                        success: false, 
                        error: 'Faltan parámetros: rowIndex, columnIndex, value' 
                    });
                }
                await updateCell(token, rowIndex, columnIndex, value);
                return res.status(200).json({ success: true, message: 'Celda actualizada' });

            case 'batch-update':
                const { updates } = req.body;
                if (!updates || !Array.isArray(updates)) {
                    return res.status(400).json({ 
                        success: false, 
                        error: 'Se requiere un array de updates' 
                    });
                }
                for (const update of updates) {
                    await updateCell(token, update.rowIndex, update.columnIndex, update.value);
                }
                return res.status(200).json({ success: true, message: 'Celdas actualizadas' });

            default:
                return res.status(404).json({ success: false, error: 'Acción no encontrada' });
        }
    } catch (error) {
        console.error('Error:', error);
        return res.status(500).json({ success: false, error: error.message });
    }
};
