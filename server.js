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
        'Yayın Tarihi': '',
        'Revizyon Tarihi': '',
        'Revizyon Sayısı': '0', // Alt tireden sayı çekilemezse varsayılan değer 0
        'Dosya İsmi': '',
        'Sorumlu Departman': ''
    };

    const fullFileNameWithExt = path.basename(originalRelativePath);
    const fileNameWithoutExt = path.parse(fullFileNameWithExt).name;

    // Dosya adını işle
    try {
        const correctedFileName = Buffer.from(fileNameWithoutExt, 'latin1').toString('utf-8');
        const lastHyphenIndex = correctedFileName.lastIndexOf('-');
        const lastUnderscoreIndex = correctedFileName.lastIndexOf('_');
        
        let fileNameForRevNo = correctedFileName;

        // Dosya isminden önce revizyon numarasını ayır
        if (lastUnderscoreIndex !== -1 && lastUnderscoreIndex < correctedFileName.length - 1) {
            const revisionPart = correctedFileName.substring(lastUnderscoreIndex + 1);
            if (!isNaN(revisionPart)) { // Gelen değerin sayı olup olmadığını kontrol et
                docInfo['Revizyon Sayısı'] = revisionPart.trim();
                fileNameForRevNo = correctedFileName.substring(0, lastUnderscoreIndex); // Revizyon sayısı olmayan dosya adı
            }
        }
        
        // Döküman No ve Dosya İsmini ayır
        const effectiveHyphenIndex = fileNameForRevNo.lastIndexOf('-');
        if (effectiveHyphenIndex !== -1) {
            docInfo['Döküman No'] = fileNameForRevNo.substring(0, effectiveHyphenIndex).trim();
            docInfo['Dosya İsmi'] = fileNameForRevNo.substring(effectiveHyphenIndex + 1).trim();
        } else {
            docInfo['Dosya İsmi'] = fileNameForRevNo.trim();
        }

    } catch {
        docInfo['Dosya İsmi'] = fileNameWithoutExt.trim();
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

    let match;

    match = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/);
    if (match) docInfo['Yayın Tarihi'] = match[1].trim();

    match = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/i);
    if (match) docInfo['Revizyon Tarihi'] = match[1].trim();

    return docInfo;
}

// Yükleme ve işleme rotası
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

    const headers = ['Döküman No', 'Yayın Tarihi', 'Revizyon Tarihi', 'Revizyon Sayısı', 'Sorumlu Departman', 'Dosya İsmi'];
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
