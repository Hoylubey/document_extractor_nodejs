const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs-extra');
const PdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const UPLOAD_DIR = path.join(__dirname, 'uploads');

// Multer disk storage configuration
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const relativePath = path.dirname(file.originalname);
        const destinationPath = path.join(UPLOAD_DIR, relativePath);
        fs.ensureDirSync(destinationPath);
        cb(null, destinationPath);
    },
    filename: (req, file, cb) => {
        cb(null, path.basename(file.originalname));
    }
});
const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Bilgi çıkarma fonksiyonu
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
        let tempDocNo = correctedFileName;
        let tempDosyaIsmi = '';

        // Revizyon sayısını ayır
        const revMatch = tempDocNo.match(/_(\d+)(?=[^0-9_]*$)/);
        if (revMatch) {
            docInfo['Revizyon Sayısı'] = revMatch[1];
            tempDocNo = tempDocNo.substring(0, revMatch.index);
        }

        // Dosya ismini ve Döküman No'yu ayır
        const lastHyphenIndex = tempDocNo.lastIndexOf('-');
        if (lastHyphenIndex !== -1) {
            const potentialDosyaIsmi = tempDocNo.substring(lastHyphenIndex + 1).trim();
            if (/[a-zA-Z]{3,}/.test(potentialDosyaIsmi)) {
                tempDosyaIsmi = potentialDosyaIsmi;
                tempDocNo = tempDocNo.substring(0, lastHyphenIndex);
            }
        }

        // Döküman No'dan alt tire ve sonrasındaki sayıları sil
        const lastUnderscoreIndex = tempDocNo.lastIndexOf('_');
        if (lastUnderscoreIndex !== -1 && !isNaN(tempDocNo.substring(lastUnderscoreIndex + 1))) {
             tempDocNo = tempDocNo.substring(0, lastUnderscoreIndex);
        }

        docInfo['Döküman No'] = tempDocNo.trim();
        docInfo['Dosya İsmi'] = tempDosyaIsmi || tempDocNo.trim();
        
    } catch (e) {
        console.error("Dosya adı işlenirken hata oluştu:", e);
        docInfo['Döküman No'] = fileNameWithoutExt.trim();
        docInfo['Dosya İsmi'] = '';
    }

    const pathSegments = originalRelativePath.split(/[\\/]/);
    if (pathSegments.length > 1) {
        const folderNameIndex = pathSegments.length - 2;
        docInfo['Sorumlu Departman'] = folderNameIndex >= 0 ? pathSegments[folderNameIndex] : 'Ana Klasör';
    } else {
        docInfo['Sorumlu Departman'] = 'Ana Klasör';
    }

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
    } catch (e) {
        console.error(`Dosya metni okunurken hata oluştu ${filePath}:`, e);
        return docInfo;
    }

    let matchFromText;
    matchFromText = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/);
    if (matchFromText) docInfo['Tarih'] = matchFromText[1].trim();
    matchFromText = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/i);
    if (matchFromText) docInfo['Revizyon Tarihi'] = matchFromText[1].trim();

    return docInfo;
}

// Yükleme ve işleme rotası (önceki mantık aynı)
app.post('/upload', upload.array('files'), async (req, res) => {
    const uploadedFiles = req.files;
    if (!uploadedFiles || uploadedFiles.length === 0) {
        return res.status(400).send('Dosya yüklenmedi veya klasör seçilmedi.');
    }
    const extractedData = [];
    const extractedDocumentNumbers = new Set();
    for (const file of uploadedFiles) {
        const originalRelativePath = file.originalname;
        const data = await extractInfo(file.path, originalRelativePath);
        if (data && data['Döküman No'] && !extractedDocumentNumbers.has(data['Döküman No'])) {
            extractedData.push(data);
            extractedDocumentNumbers.add(data['Döküman No']);
        }
    }
    if (extractedData.length === 0) {
        return res.status(400).send('Hiçbir geçerli belge işlenemedi veya hepsi mükerrerdi.');
    }
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Belge Bilgileri');
    const headers = ['Döküman No', 'Tarih', 'Revizyon Tarihi', 'Revizyon Sayısı', 'Sorumlu Departman', 'Dosya İsmi'];
    worksheet.addRow(headers);
    extractedData.forEach(rowData => {
        const rowValues = headers.map(header => rowData[header] || '');
        worksheet.addRow(rowValues);
    });
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Belge_Bilgileri.xlsx');
    res.send(buffer);
});

app.listen(PORT, () => {
    console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor`);
    fs.ensureDirSync(UPLOAD_DIR);
});
