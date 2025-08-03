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
const MASTER_LIST_PATH = path.join(__dirname, 'Döküman Özet Listesi.xlsx');

// Multer storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const destinationPath = path.join(UPLOAD_DIR, path.dirname(file.originalname));
        fs.ensureDirSync(destinationPath);
        cb(null, destinationPath);
    },
    filename: (req, file, cb) => {
        cb(null, path.basename(file.originalname));
    }
});
const upload = multer({ storage });

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Bilgi çıkarımı
async function extractInfo(filePath, originalRelativePath) {
    const docInfo = {
        'Döküman No': '',
        'Tarih': '',
        'Revizyon Tarihi': '',
        'Revizyon Sayısı': '',
        'Dosya İsmi': '',
        'Sorumlu Departman': ''
    };

    const fullFileNameWithExt = path.basename(originalRelativePath);
    const fileNameWithoutExt = path.parse(fullFileNameWithExt).name;

    // Dosya ismi: ilk tireden sonrası
    try {
        const correctedFileName = Buffer.from(fileNameWithoutExt, 'latin1').toString('utf-8').normalize('NFC');
        const firstHyphenIndex = correctedFileName.indexOf('-');
        docInfo['Dosya İsmi'] = firstHyphenIndex !== -1 && firstHyphenIndex < correctedFileName.length - 1
            ? correctedFileName.substring(firstHyphenIndex + 1).trim()
            : correctedFileName.trim();
    } catch {
        docInfo['Dosya İsmi'] = fileNameWithoutExt.trim();
    }

    // Sorumlu departman klasör adı
    const pathSegments = originalRelativePath.split(path.sep);
    docInfo['Sorumlu Departman'] = pathSegments.length > 1
        ? pathSegments[pathSegments.length - 2]
        : 'Ana Klasör';

    let textContent = '';
    const fileExtension = path.extname(filePath).toLowerCase();

    try {
        if (fileExtension === '.pdf') {
            const dataBuffer = fs.readFileSync(filePath);
            const data = await PdfParse(dataBuffer);
            textContent = data.text;
        } else if (fileExtension === '.docx' || fileExtension === '.doc') {
            const result = await mammoth.extractRawText({ path: filePath });
            textContent = result.value;
        }
    } catch (e) {
        console.error(`Metin okunamadı: ${filePath}`, e);
        return docInfo;
    }

    let match;
    match = textContent.match(/Doküman No\s*[:\s]*([A-Z0-9.\-]+)/i);
    if (match) docInfo['Döküman No'] = match[1].trim();

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
        return res.status(400).send('Dosya yüklenmedi.');
    }

    const extractedData = [];
    for (const file of uploadedFiles) {
        const originalRelativePath = file.originalname;
        const data = await extractInfo(file.path, originalRelativePath);
        if (data) extractedData.push(data);
        await fs.remove(file.path).catch(e => console.error('Silinemedi:', e));
    }

    // Ana listeyi yükle veya yeni oluştur
    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(MASTER_LIST_PATH)) {
        await workbook.xlsx.readFile(MASTER_LIST_PATH);
    }
    const worksheet = workbook.getWorksheet('Belge Bilgileri') || workbook.addWorksheet('Belge Bilgileri');

    const headers = ['Döküman No', 'Tarih', 'Revizyon Tarihi', 'Revizyon Sayısı', 'Sorumlu Departman', 'Dosya İsmi'];
    if (worksheet.actualRowCount === 0) worksheet.addRow(headers);

    const existingRows = {};
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // başlık
        const fileName = row.getCell(headers.indexOf('Dosya İsmi') + 1).value;
        if (fileName) existingRows[fileName] = rowNumber;
    });

    extractedData.forEach(rowData => {
        const key = rowData['Dosya İsmi'];
        const rowValues = headers.map(header => rowData[header] || '');

        if (existingRows[key]) {
            const row = worksheet.getRow(existingRows[key]);
            headers.forEach((header, i) => {
                const newVal = rowData[header];
                const oldVal = row.getCell(i + 1).value;
                row.getCell(i + 1).value = (newVal && newVal.trim() !== '') ? newVal : oldVal;
            });
            row.commit();
        } else {
            worksheet.addRow(rowValues);
        }
    });

    await workbook.xlsx.writeFile(MASTER_LIST_PATH);
    await fs.emptyDir(UPLOAD_DIR);

    // Excel dosyasını yanıt olarak gönder
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Guncel_AnaListe.xlsx');
    res.send(buffer);
});

app.listen(PORT, () => {
    console.log(`Sunucu çalışıyor: http://localhost:${PORT}`);
    fs.ensureDirSync(UPLOAD_DIR);
});
