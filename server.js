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

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        fs.ensureDir(UPLOAD_DIR, (err) => {
            if (err) {
                console.error("HATA: Yükleme klasörü oluşturulurken hata:", err);
                return cb(err);
            }
            cb(null, UPLOAD_DIR);
        });
    },
    filename: (req, file, cb) => {
        try {
            const uniqueName = `${uuidv4()}${path.extname(file.originalname)}`;
            console.log(`LOG: Yeni dosya adı oluşturuldu: ${uniqueName}`);
            cb(null, uniqueName);
        } catch (e) {
            console.error("HATA: Dosya adı oluşturulurken hata:", e);
            cb(e);
        }
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
        console.log(`LOG: İşlenen dosya adı (orijinal): ${originalRelativePath}`);
        const correctedFileName = Buffer.from(fileNameWithoutExt, 'latin1').toString('utf-8');
        let tempFileName = correctedFileName;
        
        // Revizyon sayılarını dosya adından ayır ve en büyüğünü bul
        const revNumbers = [...tempFileName.matchAll(/_(\d+)/g)]
            .map(match => parseInt(match[1]))
            .filter(num => !isNaN(num));

        if (revNumbers.length > 0) {
            const maxRev = Math.max(...revNumbers);
            docInfo['Revizyon Sayısı'] = maxRev.toString();
            // Dosya adından revizyon numarasını temizle
            tempFileName = tempFileName.replace(new RegExp(`_${maxRev}`), '');
        }

        // Döküman No ve Dosya İsmini ayır (Yeni Mantık)
        const firstHyphenIndex = tempFileName.indexOf('-');
        // Eğer tire işareti varsa, son tire işaretini baz al
        const lastHyphenIndex = tempFileName.lastIndexOf('-');
        
        if (lastHyphenIndex !== -1) {
            // Son tire işaretine göre ayır
            docInfo['Döküman No'] = tempFileName.substring(0, lastHyphenIndex).trim();
            docInfo['Dosya İsmi'] = tempFileName.substring(lastHyphenIndex + 1).trim();
        } else {
            // Eğer hiç tire işareti yoksa, dosya adının tamamını Döküman No olarak kabul et
            docInfo['Döküman No'] = tempFileName.trim();
        }
        
        console.log(`LOG: Ayrıştırılan bilgiler: Döküman No: ${docInfo['Döküman No']}, Revizyon Sayısı: ${docInfo['Revizyon Sayısı']}, Dosya İsmi: ${docInfo['Dosya İsmi']}`);

    } catch (e) {
        console.error("HATA: Dosya adı işlenirken hata oluştu:", e);
        docInfo['Döküman No'] = fileNameWithoutExt.trim();
        docInfo['Dosya İsmi'] = '';
    }

    // Sorumlu Departmanı belirleme
    const pathSegments = originalRelativePath.split(/[\\/]/);
    if (pathSegments.length > 1) {
        const folderNameIndex = pathSegments.length - 2;
        docInfo['Sorumlu Departman'] = folderNameIndex >= 0 ? pathSegments[folderNameIndex] : 'Ana Klasör';
    } else {
        docInfo['Sorumlu Departman'] = 'Ana Klasör';
    }

    // Dosya içeriğinden tarihleri çekme
    let textContent = '';
    const fileExtension = path.extname(filePath).toLowerCase();
    try {
        console.log(`LOG: Dosya içeriği okunuyor: ${filePath}`);
        if (fileExtension === '.pdf') {
            const dataBuffer = fs.readFileSync(filePath);
            const data = await PdfParse(dataBuffer);
            textContent = data.text.normalize('NFC');
        } else if (fileExtension === '.docx' || fileExtension === '.doc') {
            const result = await mammoth.extractRawText({ path: filePath });
            textContent = result.value.normalize('NFC');
        }
        console.log("LOG: Dosya içeriği başarıyla okundu.");
    } catch (e) {
        console.error(`HATA: Dosya metni okunurken hata oluştu ${filePath}:`, e);
        return docInfo;
    }

    let matchFromText;
    matchFromText = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/);
    if (matchFromText) docInfo['Tarih'] = matchFromText[1].trim();
    matchFromText = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/i);
    if (matchFromText) docInfo['Revizyon Tarihi'] = matchFromText[1].trim();

    return docInfo;
}

// Yükleme ve işleme rotası
app.post('/upload', upload.array('files'), async (req, res) => {
    try {
        const uploadedFiles = req.files;
        if (!uploadedFiles || uploadedFiles.length === 0) {
            return res.status(400).send('Dosya yüklenmedi veya klasör seçilmedi.');
        }
        console.log(`LOG: ${uploadedFiles.length} adet dosya yüklendi.`);
        const extractedData = [];
        const extractedDocumentNumbers = new Set();
        for (const file of uploadedFiles) {
            const originalRelativePath = file.originalname;
            const data = await extractInfo(file.path, originalRelativePath);
            if (data && data['Döküman No'] && !extractedDocumentNumbers.has(data['Döküman No'])) {
                extractedData.push(data);
                extractedDocumentNumbers.add(data['Döküman No']);
            }
            try {
                fs.unlinkSync(file.path);
                console.log(`LOG: Dosya başarıyla silindi: ${file.path}`);
            } catch (e) {
                console.error(`HATA: Dosya silinirken hata oluştu: ${file.path}`, e);
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
        console.log("LOG: Excel dosyası başarıyla oluşturuldu ve gönderildi.");
    } catch (error) {
        console.error("KRİTİK HATA: Yükleme rotası işlenirken genel bir hata oluştu:", error);
        res.status(500).send('Sunucu tarafında bir hata oluştu.');
    }
});

app.listen(PORT, () => {
    console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor`);
    fs.ensureDirSync(UPLOAD_DIR);
});
