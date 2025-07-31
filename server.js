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
        // file.originalname, webkitRelativePath ile gelen tam yolu içerir (örn: 'KlasorAdi/AltKlasor/Dosya.pdf')
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

// Bilgi çekme fonksiyonları
async function extractInfo(filePath, originalRelativePath) {
    const docInfo = {
        'Döküman No': '',
        'Tarih': '',
        'Revizyon Tarihi': '',
        'Revizyon Sayısı': '',
        'Dosya İsmi': '',
        'Sorumlu Departman': ''
    };

    // --- Dosya İsmi Mantığı (Geliştirilmiş) ---
    const fullFileName = path.basename(originalRelativePath); // Örn: FR.01-BS.TL.02_0-SMS ve Maling Cevap Şablonu.pdf
    const fileNameWithoutExt = path.parse(fullFileName).name; // Uzantısız kısım: FR.01-BS.TL.02_0-SMS ve Maling Cevap Şablonu

    // İlk '-' işaretinden sonraki kısmı al
    const firstHyphenIndex = fileNameWithoutExt.indexOf('-');
    if (firstHyphenIndex !== -1 && firstHyphenIndex < fileNameWithoutExt.length - 1) {
        docInfo['Dosya İsmi'] = fileNameWithoutExt.substring(firstHyphenIndex + 1).trim();
    } else {
        docInfo['Dosya İsmi'] = fileNameWithoutExt.trim(); // '-' yoksa tüm uzantısız adı kullan
    }

    // --- Sorumlu Departman Mantığı (Geliştirilmiş) ---
    const parts = originalRelativePath.split(path.sep); // Klasör yollarını ayır
    // path.sep, işletim sistemine göre / veya \ olur
    
    // Eğer dosya doğrudan yüklenen kök klasörde değilse (yani alt klasörlerdeyse)
    if (parts.length > 1 && parts[0] !== fullFileName) { // Check if it's not just the file name
        // Sorumlu Departman: Dosyanın bulunduğu klasörün adı (bir üst klasör)
        // Eğer "KlasorAdi/AltKlasor/Dosya.pdf" ise "AltKlasor" olmalı.
        // Eğer "KlasorAdi/Dosya.pdf" ise "KlasorAdi" olmalı.
        const folderName = parts[parts.length - 2]; // Sondan ikinci eleman (klasör adı)
        if (folderName) { // Eğer bir klasör adı varsa
            docInfo['Sorumlu Departman'] = folderName;
        } else {
            docInfo['Sorumlu Departman'] = 'Ana Klasör'; // Çoklu klasör yapısı yoksa
        }
    } else {
        docInfo['Sorumlu Departman'] = 'Ana Klasör'; // Dosya direkt kök klasörde
    }


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
        console.error(`Dosya metni okunurken hata oluştu ${filePath}:`, e);
        return docInfo; // Hata durumunda boş bilgilerle dön
    }

    // --- Bilgi Çekme (Revizyon Tarihi dahil, regex'ler daha esnekleştirildi) ---
    let match;

    // Doküman No (Daha esnek regex: Boşluklar ve tireler için)
    match = textContent.match(/Doküman No\s*[:\s]*([A-Z0-9.\-\s]+)/i);
    if (match) docInfo['Döküman No'] = match[1].trim();

    // Yayın Tarihi (Daha esnek regex)
    match = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/);
    if (match) docInfo['Tarih'] = match[1].trim();

    // Revizyon No (Daha esnek regex)
    match = textContent.match(/Revizyon No\s*[:\s]*(\d+)/i);
    if (match) docInfo['Revizyon Sayısı'] = match[1].trim();

    // Revizyon Tarihi (Daha esnek regex: Boşluklar, : ve farklı tarih formatları için)
    // Örnek PDF'ten aldığımız bilgi: "Revizyon Tarihi: 30.03.2020" 
    match = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})/);
    if (match) docInfo['Revizyon Tarihi'] = match[1].trim();
    // Eğer yine gelmezse, metin içeriğini konsola yazdırıp elle kontrol etmek gerekebilir.

    return docInfo;
}

app.post('/upload', upload.array('files'), async (req, res) => {
    const uploadedFiles = req.files;
    if (!uploadedFiles || uploadedFiles.length === 0) {
        return res.status(400).send('No files uploaded or no folder selected.');
    }

    const extractedData = [];

    for (const file of uploadedFiles) {
        // originalname, webkitRelativePath'ten gelen orijinal yolu içerir (örn: 'KlasorAdi/Dosya.pdf')
        // Bu yol, dosyanın kendi içinde bulunduğu klasör yapısını belirtir.
        const originalRelativePath = file.originalname; 
        
        const data = await extractInfo(file.path, originalRelativePath);
        if (data) {
            extractedData.push(data);
        }

        // Geçici dosyayı sil
        try {
            await fs.remove(file.path);
        } catch (e) {
            console.error(`Dosya silinirken hata oluştu ${file.path}:`, e);
        }
    }

    // Yüklenen tüm geçici klasörleri temizle
    try {
        await fs.emptyDir(UPLOAD_DIR); // uploads klasörünün içini tamamen boşaltır
    } catch (e) {
        console.error(`Geçici yükleme klasörü temizlenirken hata oluştu ${UPLOAD_DIR}:`, e);
    }

    if (extractedData.length === 0) {
        return res.status(400).send('No PDF or Word documents found or processed.');
    }

    // Excel oluştur
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Belge Bilgileri');

    // Başlıkları ekle - Sorumlu Departman ve Dosya İsmi sırasını düzenledim
    const headers = ['Döküman No', 'Tarih', 'Revizyon Tarihi', 'Revizyon Sayısı', 'Sorumlu Departman', 'Dosya İsmi'];
    worksheet.addRow(headers);

    // Verileri ekle
    extractedData.forEach(rowData => {
        const rowValues = headers.map(header => rowData[header] || '');
        worksheet.addRow(rowValues);
    });

    // Excel dosyasını belleğe kaydet
    const buffer = await workbook.xlsx.writeBuffer();

    // Kullanıcıya Excel dosyasını gönder
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Belge_Bilgileri.xlsx');
    res.send(buffer);
});

app.listen(PORT, () => {
    console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor`);
    fs.ensureDirSync(UPLOAD_DIR); // Uygulama başlarken uploads klasörünü oluştur
});