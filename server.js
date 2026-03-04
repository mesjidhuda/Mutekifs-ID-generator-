const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const AdmZip = require('adm-zip');
const QRCode = require('qrcode');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.json());

/**
 * HELPER: Converts Excel serial dates (e.g., 46023) to YYYY-MM-DD string
 */
function formatExcelDate(serial) {
    if (!serial) return "N/A";
    if (isNaN(serial)) return String(serial).trim(); 
    
    try {
        const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    } catch (e) {
        return String(serial);
    }
}

// Main Processing Route
app.post('/process', upload.fields([{ name: 'excel' }, { name: 'zip' }]), async (req, res) => {
    try {
        if (!req.files.excel || !req.files.zip) {
            return res.status(400).json({ error: 'Missing Excel or ZIP files.' });
        }

        // 1. Parse Excel
        const workbook = xlsx.readFile(req.files.excel[0].path);
        const sheetName = workbook.SheetNames[0];
        const rawData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

        // 2. Parse ZIP Photos into memory
        const zip = new AdmZip(req.files.zip[0].path);
        const zipEntries = zip.getEntries();
        const photoMap = {};

        zipEntries.forEach(entry => {
            if (!entry.isDirectory && entry.entryName.match(/\.(jpg|jpeg|png)$/i)) {
                const fileNameWithoutExt = entry.name.split('.').slice(0, -1).join('.').toLowerCase();
                photoMap[fileNameWithoutExt] = `data:image/png;base64,${entry.getData().toString('base64')}`;
            }
        });

        // 3. Process Records and Match Photos
        const processedRecords = await Promise.all(rawData.map(async (row) => {
            const idKey = String(row.id_number || "").trim();
            const cleanDate = formatExcelDate(row.expiry_date);
            const cleanEntry = formatExcelDate(row.entry_date)

            // --- QR DATA PAYLOAD: Including all data from Excel ---
            const qrPayload = JSON.stringify({
                id: idKey,
                name: (function(name) {
    return name
      ? name
          .toLowerCase()
          .trim()
          .split(/\s+/)
          .map(word =>
            word.charAt(0).toUpperCase() + word.slice(1)
          )
          .join(" ")
      : "N/A";
  })(row.full_name),
                gender: row.gender || "N/A",  // Added
                age: row.age || "N/A",        // Added (QR only)
                role: row.role || "Mutekif",
                org: row.organization || "Mesjid Huda",
                expiry: cleanDate 
            });
            
            const qrCodeBase64 = await QRCode.toDataURL(qrPayload, {
                errorCorrectionLevel: 'M',
                margin: 1,
                width: 300
            });

            return {
                ...row,
                expiry_date: cleanDate ,
                gender: row.gender || "N/A", // Ensure gender is passed to front-end for UI
                age: row.age || "N/A",       // Age passed but not rendered in HTML
                photoBase64: photoMap[idKey.toLowerCase()] || null, 
                qrBase64: qrCodeBase64
            };
        }));

        // 4. Cleanup temp files
        fs.unlinkSync(req.files.excel[0].path);
        fs.unlinkSync(req.files.zip[0].path);

        res.json({ success: true, data: processedRecords });

    } catch (error) {
        console.error("Processing Error:", error);
        res.status(500).json({ error: 'Processing failed. Check file structures.' });
    }
});

const PORT = 8080 || 3000;
app.listen(PORT, () => {
    console.log(`\n✅ ID GENERATOR ACTIVE`);
    console.log(`🌐 URL: http://localhost:${PORT}`);
});
