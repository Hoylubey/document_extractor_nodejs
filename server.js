// Geliştirilmiş versiyon: Hem CSV hem Excel ana listeyi destekler, PDF/DOCX/XLSX/CSV içerik okur
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs-extra');
const PdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = process.env.PORT || 3000;
const UPLOAD_DIR = path.join(__dirname, 'uploads');
const MASTER_CSV_PATH = path.join(__dirname, 'Doküman Özet Listesi.csv');
const MASTER_XLSX_PATH = path.join(__dirname, 'Doküman Özet Listesi.xlsx');

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        fs.ensureDir(UPLOAD_DIR)
            .then(() => cb(null, UPLOAD_DIR))
            .catch(err => cb(err));
    },
    filename: (req, file, cb) => {
        const uniqueName = `${uuidv4()}${path.extname(file.originalname)}`;
        cb(null, uniqueName);
    }
});
const upload = multer({ storage });

app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

function parseMasterList() {
    if (fs.existsSync(MASTER_XLSX_PATH)) {
        return parseXlsxData(MASTER_XLSX_PATH);
    } else if (fs.existsSync(MASTER_CSV_PATH)) {
        const content = fs.readFileSync(MASTER_CSV_PATH, 'utf8');
        return parseCsvData(content);
    } else {
        throw new Error('Ana doküman listesi dosyası bulunamadı.');
    }
}

function parseCsvData(content) {
    const lines = content.split('\n').filter(l => l.trim());
    const header = lines.find(line => line.includes('Doküman Kodu'));
    if (!header) return {};
    const headers = header.split(';').map(h => h.trim());
    const data = {};
    lines.slice(lines.indexOf(header) + 1).forEach(line => {
        const values = line.split(';').map(v => v.trim());
        const docNo = values[headers.indexOf('Doküman Kodu')];
        if (docNo) {
            data[docNo] = {
                'Döküman No': docNo,
                'Tarih': values[headers.indexOf('Hazırlama Tarihi')] || '',
                'Revizyon Sayısı': values[headers.indexOf('Revizyon No')] || '0',
                'Revizyon Tarihi': values[headers.indexOf('Revizyon Tarihi')] || '',
                'Sorumlu Departman': values[headers.indexOf('Sorumlu Kısım')] || '',
                'Dosya İsmi': values[headers.indexOf('Doküman Adı')] || ''
            };
        }
    });
    return data;
}

async function parseXlsxData(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.worksheets[0];
    const headers = sheet.getRow(1).values;
    const data = {};
    sheet.eachRow((row, index) => {
        if (index === 1) return;
        const docNo = row.getCell(headers.indexOf('Doküman Kodu')).text.trim();
        if (docNo) {
            data[docNo] = {
                'Döküman No': docNo,
                'Tarih': row.getCell(headers.indexOf('Hazırlama Tarihi')).text.trim() || '',
                'Revizyon Sayısı': row.getCell(headers.indexOf('Revizyon No')).text.trim() || '0',
                'Revizyon Tarihi': row.getCell(headers.indexOf('Revizyon Tarihi')).text.trim() || '',
                'Sorumlu Departman': row.getCell(headers.indexOf('Sorumlu Kısım')).text.trim() || '',
                'Dosya İsmi': row.getCell(headers.indexOf('Doküman Adı')).text.trim() || ''
            };
        }
    });
    return data;
}

app.post('/upload', upload.array('files'), async (req, res) => {
    try {
        const uploadedFiles = req.files;
        if (!uploadedFiles?.length) return res.status(400).send('Dosya yok.');

        const masterList = await parseMasterList();
        const extracted = [], errors = [];

        for (const file of uploadedFiles) {
            const docData = await extractInfo(file);
            const master = masterList[docData['Döküman No']];
            if (master) {
                for (const key in master) {
                    if (!docData[key]) docData[key] = master[key];
                }
            }
            extracted.push(docData);
            fs.unlinkSync(file.path);
        }

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Belge Bilgileri');
        const headers = ['Döküman No', 'Tarih', 'Revizyon Tarihi', 'Revizyon Sayısı', 'Sorumlu Departman', 'Dosya İsmi'];
        ws.addRow(headers);
        extracted.forEach(row => ws.addRow(headers.map(h => row[h] || '')));

        const buffer = await wb.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Belge_Bilgileri.xlsx');
        res.send(buffer);
    } catch (e) {
        console.error(e);
        res.status(500).send('Sunucu hatası');
    }
});

async function extractInfo(file) {
    const ext = path.extname(file.originalname).toLowerCase();
    const base = path.parse(file.originalname).name;
    const docNo = base.split('-')[0]?.trim();
    const docName = base.split('-')[1]?.trim() || '';
    const revMatch = base.match(/_(\d+)/);
    const revNo = revMatch ? revMatch[1] : '0';
    const responsible = path.dirname(file.originalname).split(path.sep).pop() || 'Ana Klasör';

    let text = '';
    if (ext === '.pdf') {
        const data = await PdfParse(fs.readFileSync(file.path));
        text = data.text;
    } else if (ext === '.docx' || ext === '.doc') {
        const result = await mammoth.extractRawText({ path: file.path });
        text = result.value;
    } else if (ext === '.xlsx' || ext === '.xls') {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(file.path);
        wb.eachSheet(sheet => {
            sheet.eachRow(row => {
                text += row.values.join(' ') + '\n';
            });
        });
    } else if (ext === '.csv') {
        text = fs.readFileSync(file.path, 'utf8');
    }

    const tarihMatch = text.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/\-]\d{2}[.\/\-]\d{4})/);
    const revDateMatch = text.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/\-]\d{2}[.\/\-]\d{4})/);

    return {
        'Döküman No': docNo,
        'Tarih': tarihMatch?.[1] || '',
        'Revizyon Tarihi': revDateMatch?.[1] || '',
        'Revizyon Sayısı': revNo,
        'Sorumlu Departman': responsible,
        'Dosya İsmi': docName
    };
}

app.listen(PORT, () => {
    console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor.`);
    fs.ensureDirSync(UPLOAD_DIR);
});
