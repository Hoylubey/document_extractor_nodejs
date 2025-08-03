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
const MASTER_FILE_NAME = 'Doküman Özet Listesi';
const MASTER_CSV_PATH = path.join(__dirname, `${MASTER_FILE_NAME}.csv`);
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

// CSV veya XLSX dosyasını okuyup ayrıştırmak için tek bir fonksiyon
async function parseMasterList() {
    let filePath;
    let fileExtension;

    if (fs.existsSync(MASTER_XLSX_PATH)) {
        filePath = MASTER_XLSX_PATH;
        fileExtension = 'xlsx';
    } else if (fs.existsSync(MASTER_CSV_PATH)) {
        filePath = MASTER_CSV_PATH;
        fileExtension = 'csv';
    } else {
        console.error("KRİTİK HATA: Ana doküman listesi dosyası (Doküman Özet Listesi.csv veya .xlsx) bulunamadı. Boş bir liste ile devam ediliyor.");
        return { masterList: {}, fileExtension: 'xlsx' };
    }

    const masterList = {};

    try {
        if (fileExtension === 'xlsx') {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
            const worksheet = workbook.getWorksheet(1);
            
            let headers = [];
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
                return { masterList: {}, fileExtension: 'xlsx' };
            }
            
            const docCodeIndex = headers.indexOf('Doküman Kodu');
            const preparationDateIndex = headers.indexOf('Hazırlama Tarihi');
            const revisionNoIndex = headers.indexOf('Revizyon No');
            const revisionDateIndex = headers.indexOf('Revizyon Tarihi');
            const responsibleDeptIndex = headers.indexOf('Sorumlu Kısım');
            const docNameIndex = headers.indexOf('Döküman Adı');

            if (docCodeIndex === -1) {
                console.error("HATA: 'Doküman Kodu' sütunu bulunamadı.");
                return { masterList: {}, fileExtension: 'xlsx' };
            }

            for (let i = headerRowIndex + 1; i <= worksheet.rowCount; i++) {
                const row = worksheet.getRow(i);
                const columns = row.values.slice(1).map(cell => cell ? String(cell).trim() : '');
                const docCode = columns[docCodeIndex];
                if (docCode) {
                    masterList[docCode] = {
                        'Döküman No': docCode,
                        'Tarih': columns[preparationDateIndex] || '',
                        'Revizyon Sayısı': columns[revisionNoIndex] || '0',
                        'Revizyon Tarihi': columns[revisionDateIndex] || '',
                        'Sorumlu Departman': columns[responsibleDeptIndex] || '',
                        'Döküman Adı': columns[docNameIndex] || ''
                    };
                }
            }
        } else if (fileExtension === 'csv') {
            const csvContent = fs.readFileSync(filePath, 'utf-8');
            const rows = [];
            let inQuote = false;
            let currentCell = '';
            let currentRow = [];
            const delimiter = ';';

            for (let i = 0; i < csvContent.length; i++) {
                const char = csvContent[i];
                if (char === '"') {
                    inQuote = !inQuote;
                } else if (char === delimiter && !inQuote) {
                    currentRow.push(currentCell.trim().replace(/"/g, ''));
                    currentCell = '';
                } else if (char === '\n' && !inQuote) {
                    currentRow.push(currentCell.trim().replace(/"/g, ''));
                    if (currentRow.some(cell => cell.length > 0)) {
                        rows.push(currentRow);
                    }
                    currentRow = [];
                    currentCell = '';
                } else {
                    currentCell += char;
                }
            }
            if (currentCell.length > 0 || currentRow.length > 0) {
                currentRow.push(currentCell.trim().replace(/"/g, ''));
                if (currentRow.some(cell => cell.length > 0)) {
                    rows.push(currentRow);
                }
            }

            if (rows.length < 2) {
                console.error("HATA: CSV dosyası boş veya sadece bir başlık satırı içeriyor.");
                return { masterList: {}, fileExtension: 'csv' };
            }

            let headerRow = null;
            let dataStartIndex = 0;
            for (let i = 0; i < rows.length; i++) {
                if (rows[i].includes('Doküman Kodu')) {
                    headerRow = rows[i];
                    dataStartIndex = i + 1;
                    break;
                }
            }
            
            if (!headerRow) {
                console.error("HATA: CSV dosyasında 'Doküman Kodu' başlığı bulunamadı.");
                return { masterList: {}, fileExtension: 'csv' };
            }

            const headers = headerRow.map(h => h.replace(/\s+/g, ' ').trim());
            
            const docCodeIndex = headers.indexOf('Doküman Kodu');
            const preparationDateIndex = headers.indexOf('Hazırlama Tarihi');
            const revisionNoIndex = headers.indexOf('Revizyon No');
            const revisionDateIndex = headers.indexOf('Revizyon Tarihi');
            const responsibleDeptIndex = headers.indexOf('Sorumlu Kısım');
            const docNameIndex = headers.indexOf('Döküman Adı');

            if (docCodeIndex === -1) {
                console.error("HATA: 'Doküman Kodu' sütunu bulunamadı.");
                return { masterList: {}, fileExtension: 'csv' };
            }

            const dataRows = rows.slice(dataStartIndex);
            dataRows.forEach((columns, index) => {
                if (columns.length > docCodeIndex) {
                    const docCode = columns[docCodeIndex];
                    if (docCode) {
                        masterList[docCode] = {
                            'Döküman No': docCode,
                            'Tarih': columns[preparationDateIndex] || '',
                            'Revizyon Sayısı': columns[revisionNoIndex] || '0',
                            'Revizyon Tarihi': columns[revisionDateIndex] || '',
                            'Sorumlu Departman': columns[responsibleDeptIndex] || '',
                            'Döküman Adı': columns[docNameIndex] || ''
                        };
                    }
                }
            });
        }
        console.log(`LOG: Ana listede ${Object.keys(masterList).length} adet belge bilgisi başarıyla yüklendi.`);
        return { masterList, fileExtension };
    } catch (e) {
        console.error(`KRİTİK HATA: Ana doküman listesi dosyası işlenirken hata oluştu:`, e);
        return { masterList: {}, fileExtension: 'xlsx' };
    }
}

// Dosya adından ve içeriğinden bilgileri çıkarma
async function extractInfo(filePath, originalRelativePath) {
    const docInfo = {
        'Döküman No': '',
        'Tarih': '',
        'Revizyon Tarihi': '',
        'Revizyon Sayısı': '0',
        'Döküman Adı': '',
        'Sorumlu Departman': ''
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
            docInfo['Revizyon Sayısı'] = maxRev.toString();
            tempFileName = tempFileName.replace(new RegExp(`_${maxRev}`), '');
        }

        const lastHyphenIndex = tempFileName.lastIndexOf('-');
        
        if (lastHyphenIndex !== -1 && lastHyphenIndex > 0) {
            docInfo['Döküman No'] = tempFileName.substring(0, lastHyphenIndex).trim();
            docInfo['Döküman Adı'] = tempFileName.substring(lastHyphenIndex + 1).trim();
        } else {
            docInfo['Döküman No'] = tempFileName.trim();
        }
        
        console.log(`LOG: Ayrıştırılan bilgiler: Döküman No: ${docInfo['Döküman No']}, Revizyon Sayısı: ${docInfo['Revizyon Sayısı']}, Döküman Adı: ${docInfo['Döküman Adı']}`);

    } catch (e) {
        console.error("HATA: Dosya adı işlenirken hata oluştu:", e);
        docInfo['Döküman No'] = fileNameWithoutExt.trim();
        docInfo['Döküman Adı'] = '';
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

// Hafızadaki verilerle yeni bir Excel dosyası oluşturma ve buffer olarak döndürme
async function createMasterListBuffer(updatedList) {
    console.log(`LOG: Güncellenmiş ana liste Excel dosyası oluşturuluyor.`);
    const headers = ['Doküman Kodu', 'Döküman Adı', 'Doküman İngilizce Adı', 'Tipi', 'Hazırlama Tarihi', 'Revizyon No', 'Revizyon Tarihi', 'Revize Eden', 'Sorumlu Kısım', 'Revizyon Nedeni', 'Revizyon', 'Son Gözden Geçirme Tarihi', 'Gelecek Gözden Geçirme Tarihi', 'Gözden Geçirme Periyodu ', 'Onay Tarihi', 'Kullanıldığı Süreçler', 'Doküman Sorumlusu', 'Arşivleme Süresi'];
    const updatedRecords = Object.values(updatedList);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(MASTER_FILE_NAME);
    
    worksheet.addRow(headers);
    
    updatedRecords.forEach(doc => {
        const rowData = headers.map(header => {
            switch (header) {
                case 'Doküman Kodu':
                    return doc['Döküman No'] || '';
                case 'Döküman Adı':
                    return doc['Döküman Adı'] || '';
                case 'Hazırlama Tarihi':
                    return doc['Tarih'] || '';
                case 'Revizyon No':
                    return doc['Revizyon Sayısı'] || '0';
                case 'Revizyon Tarihi':
                    return doc['Revizyon Tarihi'] || '';
                case 'Sorumlu Kısım':
                    return doc['Sorumlu Departman'] || '';
                default:
                    // Diğer başlıklar için mevcut veri varsa kullan, yoksa boş bırak.
                    // parseMasterList fonksiyonunda okunan veriler headers ile eşleşmiyorsa bu kısım boş kalacaktır.
                    return '';
            }
        });
        worksheet.addRow(rowData);
    });
    
    return workbook.xlsx.writeBuffer();
}

// Dosya yükleme rotası
app.post('/upload', upload.array('files'), async (req, res) => {
    try {
        const uploadedFiles = req.files;
        if (!uploadedFiles || uploadedFiles.length === 0) {
            return res.status(400).send('Dosya yüklenmedi veya klasör seçilmedi.');
        }

        console.log("LOG: Ana doküman listesi okunuyor...");
        let masterDocumentList = {};
        try {
            const result = await parseMasterList();
            masterDocumentList = result.masterList;
            console.log(`LOG: Ana listede ${Object.keys(masterDocumentList).length} adet belge bilgisi yüklendi.`);
        } catch (e) {
            console.error(`KRİTİK HATA: Ana doküman listesi dosyası okunamadı.`, e);
            masterDocumentList = {};
        }

        // Ana listenin kopyasını oluşturuyoruz, böylece orijinal veriyi bozmadan güncelleme yapabiliriz
        const updatedMasterList = JSON.parse(JSON.stringify(masterDocumentList));
        
        const extractedDocumentNumbers = new Set();
        
        // Yeni yüklenen dosyaları işliyoruz
        for (const file of uploadedFiles) {
            const originalRelativePath = file.originalname;
            const data = await extractInfo(file.path, originalRelativePath);
            
            if (data && data['Döküman No'] && !extractedDocumentNumbers.has(data['Döküman No'])) {
                console.log(`\nLOG: Belge ${data['Döküman No']} işleniyor.`);
                extractedDocumentNumbers.add(data['Döküman No']);

                const masterDoc = masterDocumentList[data['Döküman No']];
                if (masterDoc) {
                    console.log(`LOG: Ana listede belge bilgisi bulundu. Güncelleme yapılıyor.`);
                    
                    if (data['Revizyon Sayısı']) updatedMasterList[data['Döküman No']]['Revizyon Sayısı'] = data['Revizyon Sayısı'];
                    if (data['Revizyon Tarihi']) updatedMasterList[data['Döküman No']]['Revizyon Tarihi'] = data['Revizyon Tarihi'];
                    if (data['Tarih']) updatedMasterList[data['Döküman No']]['Tarih'] = data['Tarih'];
                    if (data['Döküman Adı']) updatedMasterList[data['Döküman No']]['Döküman Adı'] = data['Döküman Adı'];
                } else {
                    console.log(`LOG: Belge ${data['Döküman No']} ana listede bulunamadı. Yeni kayıt olarak ekleniyor.`);
                    // Yeni kaydı listeye ekliyoruz
                    updatedMasterList[data['Döküman No']] = data;
                }
            }

            // Geçici olarak yüklenen dosyayı sil
            try {
                fs.unlinkSync(file.path);
            } catch (e) {
                console.error(`HATA: Dosya silinirken hata oluştu: ${file.path}`, e);
            }
        }

        if (Object.keys(updatedMasterList).length === 0) {
            return res.status(400).send('Hiçbir geçerli belge işlenemedi veya ana liste oluşturulamadı.');
        }

        const updatedMasterListBuffer = await createMasterListBuffer(updatedMasterList);
        
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
