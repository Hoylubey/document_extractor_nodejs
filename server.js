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
        fs.ensureDirSync(destinationPath); // Klasör yoksa oluştur
        cb(null, destinationPath);
    },
    filename: (req, file, cb) => {
        cb(null, path.basename(file.originalname)); // Orijinal dosya adını kullan
    }
});

const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Information extraction function
async function extractInfo(filePath, originalRelativePath) {
    const docInfo = {
        'Döküman No': '',
        'Tarih': '',
        'Revizyon Tarihi': '',
        'Revizyon Sayısı': '',
        'Dosya İsmi': '',
        'Sorumlu Departman': ''
    };

    // --- Dosya İsmi Mantığı (Onay ve Hata Ayıklama Eklendi) ---
    // DEBUGGING: Orijinal dosya adını konsola yazdır
    console.log(`DEBUG: Original Relative Path received: ${originalRelativePath}`);

    const fullFileNameWithExt = path.basename(originalRelativePath);
    const fileNameWithoutExt = path.parse(fullFileNameWithExt).name;

    const firstHyphenIndex = fileNameWithoutExt.indexOf('-');
    if (firstHyphenIndex !== -1 && firstHyphenIndex < fileNameWithoutExt.length - 1) {
        docInfo['Dosya İsmi'] = fileNameWithoutExt.substring(firstHyphenIndex + 1).trim();
    } else {
        docInfo['Dosya İsmi'] = fileNameWithoutExt.trim();
    }
    // DEBUGGING: İşlenmiş Dosya İsmini konsola yazdır
    console.log(`DEBUG: Processed 'Dosya İsmi': ${docInfo['Dosya İsmi']}`);

    // --- Sorumlu Departman Mantığı ---
    const pathSegments = originalRelativePath.split(path.sep);
    
    if (pathSegments.length > 1) {
        const folderNameIndex = pathSegments.length - 2;
        if (folderNameIndex >= 0) {
            docInfo['Sorumlu Departman'] = pathSegments[folderNameIndex];
        } else {
            docInfo['Sorumlu Departman'] = 'Ana Klasör';
        }
    } else {
        docInfo['Sorumlu Departman'] = 'Ana Klasör';
    }

    let textContent = '';
    const fileExtension = path.extname(filePath).toLowerCase();

    try {
        if (fileExtension === '.pdf') {
            const dataBuffer = fs.readFileSync(filePath);
            const data = await PdfParse(dataBuffer);
            textContent = data.text;
            console.log(`--- PDF Metin İçeriği (${path.basename(filePath)}) ---`);
            console.log(textContent);
            console.log('--- Metin İçeriği Sonu ---');
        } else if (fileExtension === '.docx' || fileExtension === '.doc') {
            const result = await mammoth.extractRawText({ path: filePath });
            textContent = result.value;
        }
    } catch (e) {
        console.error(`Dosya metni okunurken hata oluştu ${filePath}:`, e);
        return docInfo;
    }

    // --- Bilgi Çekme ---
    let match;

    // Doküman No (Düzeltildi: Sadece belge numarası karakterlerini yakala, boşlukları değil)
    // Örn: TR.01.STD.001 gibi ifadeler için
    match = textContent.match(/Doküman No\s*[:\s]*([A-Z0-9.\-]+)/i);
    if (match) docInfo['Döküman No'] = match[1].trim();
    // DEBUGGING: Çekilen Döküman No'yu konsola yazdır
    console.log(`DEBUG: Extracted 'Döküman No': ${docInfo['Döküman No']}`);

    match = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/);
    if (match) docInfo['Tarih'] = match[1].trim();

    match = textContent.match(/Revizyon No\s*[:\s]*(\d+)/i);
    if (match) docInfo['Revizyon Sayısı'] = match[1].trim();

    match = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/i);
    if (match) docInfo['Revizyon Tarihi'] = match[1].trim();

    return docInfo;
}

app.post('/upload', upload.array('files'), async (req, res) => {
    const uploadedFiles = req.files;
    if (!uploadedFiles || uploadedFiles.length === 0) {
        return res.status(400).send('No files uploaded or no folder selected.');
    }

    const extractedData = [];

    for (const file of uploadedFiles) {
        const originalRelativePath = file.originalname;
        
        const data = await extractInfo(file.path, originalRelativePath);
        if (data) {
            extractedData.push(data);
        }

        try {
            await fs.remove(file.path);
        } catch (e) {
            console.error(`Dosya silinirken hata oluştu ${file.path}:`, e);
        }
    }

    try {
        await fs.emptyDir(UPLOAD_DIR);
    } catch (e) {
        console.error(`Geçici yükleme klasörü temizlenirken hata oluştu ${UPLOAD_DIR}:`, e);
    }

    if (extractedData.length === 0) {
        return res.status(400).send('No PDF or Word documents found or processed.');
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
