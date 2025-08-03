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
const UPLOAD_DIR = path.join(__dirname, 'Uploads');
const MASTER_FILE_NAME = 'Doküman Özet Listesi';
const MASTER_XLSX_PATH = path.join(__dirname, `${MASTER_FILE_NAME}.xlsx`);

// Yükleme klasörünü oluşturma
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

// Tarih formatını dönüştürme fonksiyonu
function formatDate(dateString) {
    if (!dateString) return '';
    try {
        let date;
        const parts = dateString.match(/(\d{1,2})[./-](\d{1,2})[./-](\d{4})/);
        if (parts) {
            date = new Date(`${parts[3]}-${parts[2]}-${parts[1]}`);
        } else {
            date = new Date(dateString);
        }
        
        if (isNaN(date.getTime())) return dateString;

        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}.${month}.${year}`;
    } catch (e) {
        return dateString;
    }
}

// CSV veya XLSX dosyasını okuyup ayrıştırmak için fonksiyon
async function parseMasterList() {
    const DEFAULT_HEADERS = ['Doküman Kodu', 'Döküman Adı', 'Sorumlu Kısım', 'Hazırlama Tarihi', 'Revizyon Tarihi', 'Revizyon No'];
    let filePath;
    let fileExtension;

    if (fs.existsSync(MASTER_XLSX_PATH)) {
        filePath = MASTER_XLSX_PATH;
        fileExtension = 'xlsx';
    } else {
        console.warn("UYARI: Ana dosya bulunamadı. Varsayılan başlıklar ile yeni bir liste oluşturuluyor.");
        return { masterList: {}, fileExtension: 'xlsx', headers: DEFAULT_HEADERS };
    }

    const masterList = {};
    let headers = [];

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.getWorksheet(1);
        
        let headerRowIndex = -1;

        worksheet.eachRow((row, rowNumber) => {
            const rowValues = row.values.slice(1).map(cell => cell ? String(cell).trim() : '');
            if (rowValues.includes('Doküman Kodu') && headerRowIndex === -1) {
                headers = rowValues;
                headerRowIndex = rowNumber;
            }
        });

        if (headerRowIndex === -1) {
            console.error("HATA: Excel dosyasında 'Doküman Kodu' başlığı bulunamadı.");
            return { masterList: {}, fileExtension: 'xlsx', headers: DEFAULT_HEADERS };
        }
        
        const docCodeIndex = headers.indexOf('Doküman Kodu');
        
        if (docCodeIndex === -1) {
            console.error("HATA: 'Doküman Kodu' sütunu bulunamadı.");
            return { masterList: {}, fileExtension: 'xlsx', headers: DEFAULT_HEADERS };
        }

        for (let i = headerRowIndex + 1; i <= worksheet.rowCount; i++) {
            const row = worksheet.getRow(i);
            const columns = row.values.slice(1).map(cell => cell ? String(cell).trim() : '');
            const docCode = columns[docCodeIndex];
            if (docCode) {
                const docData = {};
                headers.forEach((header, index) => {
                    docData[header] = columns[index] || '';
                });
                masterList[docCode] = docData;
            }
        }

        console.log(`LOG: Ana listede ${Object.keys(masterList).length} adet belge bilgisi başarıyla yüklendi.`);
        return { masterList, fileExtension, headers };
    } catch (e) {
        console.error(`KRİTİK HATA: Ana doküman listesi dosyası işlenirken hata oluştu:`, e);
        return { masterList: {}, fileExtension: 'xlsx', headers: DEFAULT_HEADERS };
    }
}

// Dosya adından ve içeriğinden bilgileri çıkarma
async function extractInfo(filePath, originalRelativePath) {
    const docInfo = {
        'Doküman Kodu': '',
        'Hazırlama Tarihi': '',
        'Revizyon Tarihi': '',
        'Revizyon No': '',
        'Döküman Adı': '',
        'Sorumlu Kısım': ''
    };

    const fullFileNameWithExt = path.basename(originalRelativePath);
    const fileNameWithoutExt = path.parse(fullFileNameWithExt).name;

    try {
        console.log(`LOG: İşlenen dosya adı (orijinal): ${originalRelativePath}`);
        let correctedFileName = Buffer.from(fileNameWithoutExt, 'latin1').toString('utf-8');
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
            docInfo['Döküman Adı'] = tempFileName.substring(lastHyphenIndex + 1).trim();
        } else {
            docInfo['Doküman Kodu'] = tempFileName.trim();
            docInfo['Döküman Adı'] = '';
        }
        
        console.log(`LOG: Ayrıştırılan bilgiler: Doküman Kodu: ${docInfo['Doküman Kodu']}, Revizyon No: ${docInfo['Revizyon No']}, Döküman Adı: ${docInfo['Döküman Adı']}`);

    } catch (e) {
        console.error("HATA: Dosya adı işlenirken hata oluştu:", e);
        docInfo['Doküman Kodu'] = fileNameWithoutExt.trim();
        docInfo['Döküman Adı'] = '';
    }

    const pathSegments = originalRelativePath.split(/[\\/]/);
    if (pathSegments.length > 1) {
        const folderNameIndex = pathSegments.length - 2;
        docInfo['Sorumlu Kısım'] = folderNameIndex >= 0 ? pathSegments[folderNameIndex] : 'Ana Klasör';
    } else {
        docInfo['Sorumlu Kısım'] = '';
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
    matchFromText = textContent.match(/Yayın Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})|(\d{4}-\d{2}-\d{2})/i);
    if (matchFromText) docInfo['Hazırlama Tarihi'] = formatDate(matchFromText[1] || matchFromText[2] || '');
    matchFromText = textContent.match(/Revizyon Tarihi\s*[:\s]*(\d{2}[.\/]\d{2}[.\/]\d{4})|(\d{4}-\d{2}-\d{2})/i);
    if (matchFromText) docInfo['Revizyon Tarihi'] = formatDate(matchFromText[1] || matchFromText[2] || '');

    return docInfo;
}

// Hafızadaki verilerle yeni bir Excel dosyası oluşturma
async function createMasterListBuffer(updatedList, headers) {
    console.log(`LOG: Güncellenmiş ana liste Excel dosyası oluşturuluyor.`);
    const updatedRecords = Object.values(updatedList);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(MASTER_FILE_NAME);
    
    worksheet.addRow(headers);
    
    updatedRecords.forEach(doc => {
        const rowData = headers.map(header => {
            const value = doc[header] || '';
            if (header.includes('Tarihi')) {
                return formatDate(value);
            }
            return value;
        });
        worksheet.addRow(rowData);
    });
    
    return workbook.xlsx.writeBuffer();
}

// Ana dosyayı diske yazma
async function saveMasterListToDisk(updatedList, headers) {
    console.log(`LOG: Ana dosya diske yazılıyor: ${MASTER_XLSX_PATH}`);
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(MASTER_FILE_NAME);
    
    worksheet.addRow(headers);
    
    Object.values(updatedList).forEach(doc => {
        const rowData = headers.map(header => {
            const value = doc[header] || '';
            if (header.includes('Tarihi')) {
                return formatDate(value);
            }
            return value;
        });
        worksheet.addRow(rowData);
    });
    
    await workbook.xlsx.writeFile(MASTER_XLSX_PATH);
    console.log(`LOG: Ana dosya başarıyla kaydedildi: ${MASTER_XLSX_PATH}`);
}

// Dosya yükleme rotası
app.post('/upload', upload.array('files'), async (req, res) => {
    try {
        const uploadedFiles = req.files;
        if (!uploadedFiles || uploadedFiles.length === 0) {
            return res.status(400).send('Dosya yüklenmedi veya klasör seçilmedi.');
        }

        console.log("LOG: Ana doküman listesi okunuyor...");
        const { masterList, headers } = await parseMasterList();
        const updatedMasterList = JSON.parse(JSON.stringify(masterList));
        const extractedDocumentNumbers = new Set();

        for (const file of uploadedFiles) {
            const originalRelativePath = file.originalname;
            const data = await extractInfo(file.path, originalRelativePath);
            
            if (data && data['Doküman Kodu']) {
                const docCode = data['Doküman Kodu'];
                if (!extractedDocumentNumbers.has(docCode)) {
                    console.log(`\nLOG: Belge ${docCode} işleniyor.`);
                    extractedDocumentNumbers.add(docCode);

                    if (updatedMasterList[docCode]) {
                        console.log(`LOG: Ana listede belge bilgisi bulundu. Güncelleme yapılıyor.`);
                        const newDocData = updatedMasterList[docCode];
                        
                        // Güncelleme: Yalnızca boş olmayan ve geçerli verilerle güncelle
                        ['Döküman Adı', 'Sorumlu Kısım', 'Revizyon No', 'Hazırlama Tarihi', 'Revizyon Tarihi'].forEach(field => {
                            if (data[field] && data[field].trim() !== '' && data[field] !== newDocData[field]) {
                                newDocData[field] = data[field];
                                console.log(`LOG: ${docCode} için '${field}' güncellendi: ${data[field]}`);
                            }
                        });
                    } else {
                        console.log(`LOG: Belge ${docCode} ana listede bulunamadı. Yeni kayıt olarak ekleniyor.`);
                        const newDoc = {};
                        headers.forEach(header => {
                            newDoc[header] = data[header] || '';
                        });
                        updatedMasterList[docCode] = newDoc;
                    }
                }
            }

            try {
                fs.unlinkSync(file.path);
            } catch (e) {
                console.error(`HATA: Dosya silinirken hata oluştu: ${file.path}`, e);
            }
        }

        if (Object.keys(updatedMasterList).length === 0) {
            return res.status(400).send('Hiçbir geçerli belge işlenemedi veya ana liste oluşturulamadı.');
        }

        // Ana dosyayı diske kaydet
        await saveMasterListToDisk(updatedMasterList, headers);

        // İndirilebilir Excel dosyası oluştur
        const updatedMasterListBuffer = await createMasterListBuffer(updatedMasterList, headers);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Güncel_Doküman_Özet_Listesi.xlsx`);
        res.send(updatedMasterListBuffer);

        console.log("LOG: Güncel ana liste Excel dosyası başarıyla oluşturuldu ve gönderildi.");

    } catch (error) {
        console.error("KRİTİK HATA: Yükleme rotası işlenirken genel bir hata oluştu:", error);
        res.status(500).send(`Sunucu tarafında bir hata oluştu: ${error.message}`);
    }
});

app.listen(PORT, () => {
    console.log(`Sunucu http://localhost:${PORT} adresinde çalışıyor`);
    fs.ensureDirSync(UPLOAD_DIR);
});
