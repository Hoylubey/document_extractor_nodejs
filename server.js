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

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        fs.ensureDir(UPLOAD_DIR, (err) => {
            if (err) return cb(err);
            cb(null, UPLOAD_DIR);
        });
    },
    filename: (req, file, cb) => {
        const uniqueName = `${uuidv4()}${path.extname(file.originalname)}`;
        cb(null, uniqueName);
    }
});
const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

function parseCsvData(csvContent) {
    const lines = csvContent.split('\n').filter(line => line.trim() !== '');
    let headerLine = lines.find(line => line.includes('Doküman Kodu'));
    if (!headerLine) return {};

    const headers = headerLine.split(';').map(h => h.trim().replace(/"/g, ''));
    const docCodeIndex = headers.indexOf('Doküman Kodu');
    const preparationDateIndex = headers.indexOf('Hazırlama Tarihi');
    const revisionNoIndex = headers.indexOf('Revizyon No');
    const revisionDateIndex = headers.indexOf('Revizyon Tarihi');
    const responsibleDeptIndex = headers.indexOf('Sorumlu Kısım');
    const docNameIndex = headers.indexOf('Doküman Adı');

    const masterList = {};
    const dataLinesStartIndex = lines.indexOf(headerLine) + 1;
    const dataLines = lines.slice(dataLinesStartIndex);

    dataLines.forEach(line => {
        const columns = line.split(';').map(c => c.trim().replace(/"/g, ''));
        const docCode = columns[docCodeIndex];
        if (docCode) {
            masterList[docCode] = {
                'Döküman No': docCode,
                'Tarih': columns[preparationDateIndex] || '',
                'Revizyon Sayısı': columns[revisionNoIndex] || '0',
                'Revizyon Tarihi': columns[revisionDateIndex] || '',
                'Sorumlu Departman': columns[responsibleDeptIndex] || '',
                'Dosya İsmi': columns[docNameIndex] || ''
            };
        }
    });

    return masterList;
}

async function extractInfo(filePath, originalRelativePath) {
    const docInfo = {
        'Döküman No': '',
        'Tarih': '',
        'Revizyon Tarihi': '',
        'Revizyon Sayısı': '0',
        'Dosya İsmi': '',
        'Sorumlu Departman': ''
    };

    const fullFileNameWithExt = path.basename(originalRelativePath);
    const fileNameWithoutExt = path.parse(fullFileNameWithExt).name;

    try {
        const correctedFileName = Buffer.from(fileNameWithoutExt, 'latin1').toString('utf-8');
        let tempFileName = correctedFileName;

        const revNumbers = [...tempFileName.matchAll(/_(\d+)/g)]
            .map(match => parseInt(match[1]))
            .filter(num => !isNaN(num));
        if (revNumbers.length > 0) {
            const maxRev = Math.max(...revNumbers);
            docInfo['Revizyon Sayısı'] = maxRev.toString();
            tempFileName = tempFileName.replace(new RegExp(`_${maxRev}`), '');
        }

        const lastHyphenIndex = tempFileName.lastIndexOf('-');
        if (lastHyphenIndex !== -1 && lastHyphenIndex > 0) {
            docInfo['Döküman No'] = tempFileName.substring(0, lastHyphenIndex).trim();
            docInfo['Dosya İsmi'] = tempFileName.substring(lastHyphenIndex + 1).trim();
        } else {
            docInfo['Döküman No'] = tempFileName.trim();
        }

    } catch {
        docInfo['Döküman No'] = fileNameWithoutExt.trim();
    }

    const pathSegments = originalRelativePath.split(/[\\/]/);
    docInfo['Sorumlu Departman'] =
        pathSegments.length > 1 ? pathSegments[pathSegments.length - 2] : 'Ana Klasör';

    let textContent = '';
    const fileExtension = path.extname(filePath).toLowerCase();
    try {
        if (fileExtension === '.pdf') {
            const dataBuffer = fs.readFileSync(filePath);
            const data = await PdfParse(dataBuffer);
            textContent = data.text.normalize('NFC');
        } else if (fileExtension === '.docx' || fileExtension === '.doc') {
            const result = await mammoth.extractRawText({ path: filePath });
            textContent = result.value.normalize('NFC');
        }
    } catch {
        return docInfo;
    }

    let matchFromText;
    matchFromText = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/);
    if (matchFromText) docInfo['Tarih'] = matchFromText[1].trim();
    matchFromText = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/i);
    if (matchFromText) docInfo['Revizyon Tarihi'] = matchFromText[1].trim();

    return docInfo;
}

app.post('/upload', upload.array('files'), async (req, res) => {
    try {
        const uploadedFiles = req.files;
        if (!uploadedFiles || uploadedFiles.length === 0) {
            return res.status(400).send('Dosya yüklenmedi.');
        }

        let masterDocumentList = {};
        try {
            const masterCsvContent = fs.readFileSync(MASTER_CSV_PATH, 'utf-8');
            masterDocumentList = parseCsvData(masterCsvContent);
        } catch (e) {
            return res.status(500).send('Ana doküman listesi okunamadı.');
        }

        const extractedData = [];
        const extractedDocumentNumbers = new Set();
        const mismatchedData = [];

        for (const file of uploadedFiles) {
            const originalRelativePath = file.originalname;
            const data = await extractInfo(file.path, originalRelativePath);

            if (data && data['Döküman No'] && !extractedDocumentNumbers.has(data['Döküman No'])) {
                const masterDoc = masterDocumentList[data['Döküman No']];

                // 🔧 Eksik bilgileri ana listeden doldur
                if (masterDoc) {
                    for (const key of Object.keys(masterDoc)) {
                        if (!data[key] || data[key].trim() === '') {
                            data[key] = masterDoc[key];
                        }
                    }
                }

                extractedData.push(data);
                extractedDocumentNumbers.add(data['Döküman No']);

                if (masterDoc) {
                    const mismatches = [];
                    if (masterDoc['Revizyon Sayısı'] !== data['Revizyon Sayısı']) {
                        mismatches.push(`Revizyon Sayısı: Ana Liste '${masterDoc['Revizyon Sayısı']}' vs. Belge '${data['Revizyon Sayısı']}'`);
                    }
                    if (masterDoc['Revizyon Tarihi'] !== data['Revizyon Tarihi']) {
                        mismatches.push(`Revizyon Tarihi: Ana Liste '${masterDoc['Revizyon Tarihi']}' vs. Belge '${data['Revizyon Tarihi']}'`);
                    }
                    if (masterDoc['Tarih'] !== data['Tarih']) {
                        mismatches.push(`Tarih: Ana Liste '${masterDoc['Tarih']}' vs. Belge '${data['Tarih']}'`);
                    }

                    if (mismatches.length > 0) {
                        mismatchedData.push({
                            'Döküman No': data['Döküman No'],
                            'Hata': mismatches.join('; ')
                        });
                    }
                } else {
                    mismatchedData.push({
                        'Döküman No': data['Döküman No'],
                        'Hata': 'Ana listede bulunmuyor.'
                    });
                }
            }

            try {
                fs.unlinkSync(file.path);
            } catch {}
        }

        if (extractedData.length === 0) {
            return res.status(400).send('Hiç geçerli belge yok.');
        }

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Belge Bilgileri');
        const headers = ['Döküman No', 'Tarih', 'Revizyon Tarihi', 'Revizyon Sayısı', 'Sorumlu Departman', 'Dosya İsmi'];
        worksheet.addRow(headers);
        extractedData.forEach(rowData => {
            worksheet.addRow(headers.map(header => rowData[header] || ''));
        });

        if (mismatchedData.length > 0) {
            const mismatchSheet = workbook.addWorksheet('Eşleşmeyen Bilgiler');
            mismatchSheet.addRow(['Döküman No', 'Hata']);
            mismatchedData.forEach(rowData => {
                mismatchSheet.addRow([rowData['Döküman No'], rowData['Hata']]);
            });
        }

        const buffer = await workbook.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Belge_Bilgileri.xlsx');
        res.send(buffer);

    } catch (e) {
        res.status(500).send('Sunucu hatası.');
    }
});

app.listen(PORT, () => {
    console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor`);
    fs.ensureDirSync(UPLOAD_DIR);
});
