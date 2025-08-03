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
const MASTER_FILE_NAME = 'Doküman Özet Listesi';

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

// CSV verisini ayrıştırmak için geliştirilmiş fonksiyon
function parseCsvData(csvContent) {
    try {
        const lines = csvContent.split('\n').filter(line => line.trim() !== '');
        if (lines.length <= 1) {
            console.error("HATA: CSV dosyası boş veya sadece bir başlık satırı içeriyor.");
            return { masterList: {}, headers: [] };
        }

        let headerLine = lines.find(line => line.includes('Doküman Kodu'));
        
        if (!headerLine) {
            console.error("HATA: CSV dosyasında 'Doküman Kodu' başlığı bulunamadı. Dosya içeriğini kontrol edin.");
            return { masterList: {}, headers: [] };
        }

        const headers = headerLine.split(';').map(h => h.trim().replace(/"/g, ''));
        const docCodeIndex = headers.indexOf('Doküman Kodu');
        
        if (docCodeIndex === -1) {
            console.error("HATA: 'Doküman Kodu' sütunu bulunamadı. Lütfen başlığı kontrol edin.");
            return { masterList: {}, headers: [] };
        }

        const masterList = {};
        const dataLinesStartIndex = lines.indexOf(headerLine) + 1;
        const dataLines = lines.slice(dataLinesStartIndex);

        dataLines.forEach((line, index) => {
            try {
                const columns = line.split(';').map(c => c.trim().replace(/"/g, ''));
                if (columns.length > docCodeIndex && columns[docCodeIndex]) {
                    const docCode = columns[docCodeIndex];
                    const docData = {};
                    headers.forEach((header, colIndex) => {
                        docData[header] = columns[colIndex] || '';
                    });
                    masterList[docCode] = docData;
                }
            } catch (e) {
                console.error(`HATA: CSV satırı işlenirken hata oluştu (satır ${index + 1}):`, e);
            }
        });

        console.log(`LOG: Ana listede ${Object.keys(masterList).length} adet belge bilgisi başarıyla yüklendi.`);
        return { masterList, headers };
    } catch (e) {
        console.error("KRİTİK HATA: CSV dosyasını ayrıştırma sırasında genel hata oluştu:", e);
        return { masterList: {}, headers: [] };
    }
}


// Dosya adından ve içeriğinden bilgileri çıkarma fonksiyonu
async function extractInfo(filePath, originalRelativePath) {
    const docInfo = {
        'Doküman Kodu': '',
        'Doküman Adı': '',
        'Hazırlama Tarihi': '',
        'Revizyon Tarihi': '',
        'Revizyon No': '0',
        'Sorumlu Kısım': ''
    };

    const fullFileNameWithExt = path.basename(originalRelativePath);
    const fileNameWithoutExt = path.parse(fullFileNameWithExt).name;

    try {
        console.log(`LOG: İşlenen dosya adı (orijinal): ${originalRelativePath}`);
        const correctedFileName = Buffer.from(fileNameWithoutExt, 'latin1').toString('utf-8');
        let tempFileName = correctedFileName;
        
        const revNumbers = [...tempFileName.matchAll(/_(\d+)/g)]
            .map(match => parseInt(match[1]))
            .filter(num => !isNaN(num));

        if (revNumbers.length > 0) {
            const maxRev = Math.max(...revNumbers);
            docInfo['Revizyon No'] = maxRev.toString();
            tempFileName = tempFileName.replace(new RegExp(`_${maxRev}`), '');
        }

        const lastHyphenIndex = tempFileName.lastIndexOf('-');
        
        if (lastHyphenIndex !== -1 && lastHyphenIndex > 0) {
            docInfo['Doküman Kodu'] = tempFileName.substring(0, lastHyphenIndex).trim();
            docInfo['Doküman Adı'] = tempFileName.substring(lastHyphenIndex + 1).trim();
        } else {
            const firstSpaceIndex = tempFileName.indexOf(' ');
            if (firstSpaceIndex !== -1) {
                docInfo['Doküman Kodu'] = tempFileName.substring(0, firstSpaceIndex).trim();
                docInfo['Doküman Adı'] = tempFileName.substring(firstSpaceIndex + 1).trim();
            } else {
                docInfo['Doküman Kodu'] = tempFileName.trim();
                docInfo['Doküman Adı'] = '';
            }
        }
        
        console.log(`LOG: Ayrıştırılan bilgiler: Doküman Kodu: ${docInfo['Doküman Kodu']}, Revizyon No: ${docInfo['Revizyon No']}, Doküman Adı: ${docInfo['Doküman Adı']}`);

    } catch (e) {
        console.error("HATA: Dosya adı işlenirken hata oluştu:", e);
        docInfo['Doküman Kodu'] = fileNameWithoutExt.trim();
        docInfo['Doküman Adı'] = '';
    }

    const pathSegments = originalRelativePath.split(/[\\/]/);
    if (pathSegments.length > 1) {
        const folderNameIndex = pathSegments.length - 2;
        docInfo['Sorumlu Kısım'] = folderNameIndex >= 0 ? pathSegments[folderNameIndex] : 'Ana Klasör';
    } else {
        docInfo['Sorumlu Kısım'] = 'Ana Klasör';
    }

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
    matchFromText = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})|(\d{4}-\d{2}-\d{2})/);
    if (matchFromText) docInfo['Hazırlama Tarihi'] = matchFromText[1] || matchFromText[2] || '';
    matchFromText = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})|(\d{4}-\d{2}-\d{2})/i);
    if (matchFromText) docInfo['Revizyon Tarihi'] = matchFromText[1] || matchFromText[2] || '';

    return docInfo;
}

// Hafızadaki verilerle yeni bir Excel dosyası oluşturma ve buffer olarak döndürme
async function createMasterListBuffer(updatedList, headers) {
    console.log(`LOG: Güncellenmiş ana liste Excel dosyası oluşturuluyor.`);
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(MASTER_FILE_NAME);
    
    worksheet.addRow(headers);
    
    const updatedRecords = Object.values(updatedList);
    updatedRecords.forEach(doc => {
        const rowData = headers.map(header => doc[header] || '');
        worksheet.addRow(rowData);
    });
    
    return workbook.xlsx.writeBuffer();
}

app.post('/upload', upload.array('files'), async (req, res) => {
    try {
        const uploadedFiles = req.files;
        if (!uploadedFiles || uploadedFiles.length === 0) {
            return res.status(400).send('Dosya yüklenmedi veya klasör seçilmedi.');
        }

        console.log("LOG: Ana doküman listesi okunuyor...");
        let masterDocumentList = {};
        let originalHeaders = [];
        try {
            const masterCsvContent = fs.readFileSync(MASTER_CSV_PATH, 'utf-8');
            const result = parseCsvData(masterCsvContent);
            masterDocumentList = result.masterList;
            originalHeaders = result.headers;
            console.log(`LOG: Ana listede ${Object.keys(masterDocumentList).length} adet belge bilgisi yüklendi.`);
        } catch (e) {
            console.error(`KRİTİK HATA: Ana doküman listesi dosyası (${MASTER_CSV_PATH}) okunamadı. Dosyanın mevcut ve doğru yerde olduğundan emin olun.`, e);
            return res.status(500).send('Sunucu hatası: Ana doküman listesi dosyası bulunamıyor veya okunamıyor. Lütfen Render loglarını kontrol edin.');
        }

        const updatedMasterList = JSON.parse(JSON.stringify(masterDocumentList));
        
        const extractedDocumentNumbers = new Set();
        
        for (const file of uploadedFiles) {
            const originalRelativePath = file.originalname;
            const data = await extractInfo(file.path, originalRelativePath);
            
            if (data && data['Doküman Kodu']) {
                if (!extractedDocumentNumbers.has(data['Doküman Kodu'])) {
                    console.log(`\nLOG: Belge ${data['Doküman Kodu']} işleniyor.`);
                    extractedDocumentNumbers.add(data['Doküman Kodu']);

                    const masterDoc = updatedMasterList[data['Doküman Kodu']];
                    if (masterDoc) {
                        console.log(`LOG: Ana listede belge bilgisi bulundu. Güncelleme yapılıyor.`);
                        
                        // Yalnızca yeni dosyadan gelen geçerli ve boş olmayan bilgileri güncelle
                        // Eğer yeni dosyadan gelen bilgi boşsa, ana listedekini koru
                        updatedMasterList[data['Doküman Kodu']]['Doküman Adı'] = data['Doküman Adı'] || masterDoc['Doküman Adı'];
                        updatedMasterList[data['Doküman Kodu']]['Sorumlu Kısım'] = data['Sorumlu Kısım'] || masterDoc['Sorumlu Kısım'];
                        updatedMasterList[data['Doküman Kodu']]['Revizyon No'] = data['Revizyon No'] || masterDoc['Revizyon No'];
                        updatedMasterList[data['Doküman Kodu']]['Hazırlama Tarihi'] = data['Hazırlama Tarihi'] || masterDoc['Hazırlama Tarihi'];
                        updatedMasterList[data['Doküman Kodu']]['Revizyon Tarihi'] = data['Revizyon Tarihi'] || masterDoc['Revizyon Tarihi'];
                        
                        // Konsola hangi alanların güncellendiğini veya korunduğunu yazdırabiliriz.
                        if (data['Doküman Adı']) console.log(`LOG: ${data['Doküman Kodu']} için 'Doküman Adı' yeni dosyadan güncellendi.`);
                        else console.log(`LOG: ${data['Doküman Kodu']} için 'Doküman Adı' ana listeden korundu.`);
                        
                        if (data['Sorumlu Kısım']) console.log(`LOG: ${data['Doküman Kodu']} için 'Sorumlu Kısım' yeni dosyadan güncellendi.`);
                        else console.log(`LOG: ${data['Doküman Kodu']} için 'Sorumlu Kısım' ana listeden korundu.`);

                        if (data['Revizyon No']) console.log(`LOG: ${data['Doküman Kodu']} için 'Revizyon No' yeni dosyadan güncellendi.`);
                        else console.log(`LOG: ${data['Doküman Kodu']} için 'Revizyon No' ana listeden korundu.`);
                        
                    } else {
                        console.log(`LOG: Belge ${data['Doküman Kodu']} ana listede bulunamadı. Yeni kayıt olarak ekleniyor.`);
                        const newDoc = {};
                        originalHeaders.forEach(header => {
                            newDoc[header] = data[header] || '';
                        });
                        updatedMasterList[data['Doküman Kodu']] = newDoc;
                    }
                }
            }

            try {
                fs.unlinkSync(file.path);
                console.log(`LOG: Dosya başarıyla silindi: ${file.path}`);
            } catch (e) {
                console.error(`HATA: Dosya silinirken hata oluştu: ${file.path}`, e);
            }
        }
        
        if (Object.keys(updatedMasterList).length === 0) {
            return res.status(400).send('Hiçbir geçerli belge işlenemedi veya ana liste oluşturulamadı.');
        }

        const updatedMasterListBuffer = await createMasterListBuffer(updatedMasterList, originalHeaders);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Güncel_Doküman_Özet_Listesi.xlsx`);
        res.send(updatedMasterListBuffer);

        console.log("LOG: Güncel ana liste Excel dosyası başarıyla oluşturuldu ve gönderildi.");

    } catch (error) {
        console.error("KRİTİK HATA: Yükleme rotası işlenirken genel bir hata oluştu:", error);
        res.status(500).send('Sunucu tarafında bir hata oluştu.');
    }
});

app.listen(PORT, () => {
    console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor`);
    fs.ensureDirSync(UPLOAD_DIR);
});
