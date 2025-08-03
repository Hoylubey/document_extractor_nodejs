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
const MASTER_EXCEL_PATH = path.join(__dirname, 'Doküman Özet Listesi.xlsx');
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

// Excel verisini daha esnek bir şekilde ayrıştırmak için yeni fonksiyon
async function parseExcelData(excelFilePath) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const worksheet = workbook.getWorksheet(1); // İlk sayfayı al

        if (!worksheet) {
            console.error("HATA: Excel dosyasında sayfa bulunamadı.");
            return { masterList: {}, headers: [] };
        }

        let headerRowNumber = -1;
        const searchLimit = 10; // İlk 10 satırı kontrol et
        
        // "Doküman Kodu" başlığını içeren satırı bul
        for (let i = 1; i <= Math.min(worksheet.rowCount, searchLimit); i++) {
            const row = worksheet.getRow(i);
            let found = false;
            row.eachCell((cell) => {
                const cellValue = String(cell.value).trim();
                if (cellValue === 'Doküman Kodu' || cellValue === 'Döküman Kodu') {
                    headerRowNumber = i;
                    found = true;
                    return false; // Döngüyü durdur
                }
            });
            if (found) break;
        }

        if (headerRowNumber === -1) {
            console.error("HATA: Excel dosyasında 'Doküman Kodu' başlığı ilk 10 satırda bulunamadı.");
            return { masterList: {}, headers: [] };
        }

        const headers = [];
        const headerRow = worksheet.getRow(headerRowNumber);
        headerRow.eachCell((cell) => {
            headers.push(String(cell.value).replace(/\n/g, ' ').trim());
        });
        
        const docCodeIndex = headers.indexOf('Doküman Kodu');
        if (docCodeIndex === -1) {
            console.error("HATA: 'Doküman Kodu' sütunu başlıklar arasında bulunamadı.");
            return { masterList: {}, headers: [] };
        }
        
        const masterList = {};
        for (let i = headerRowNumber + 1; i <= worksheet.rowCount; i++) {
            const row = worksheet.getRow(i);
            const docCodeCell = row.getCell(docCodeIndex + 1);
            const docCode = docCodeCell.value ? String(docCodeCell.value).trim() : null;

            if (docCode) {
                const docData = {};
                headers.forEach((header, colIndex) => {
                    const cellValue = row.getCell(colIndex + 1).value;
                    let value = '';
                    if (cellValue instanceof Date) {
                        const d = cellValue;
                        value = `${d.getDate().toString().padStart(2, '0')}/${(d.getMonth() + 1).toString().padStart(2, '0')}/${d.getFullYear()}`;
                    } else if (typeof cellValue === 'object' && cellValue !== null) {
                        if (cellValue.richText) {
                            value = cellValue.richText.map(t => t.text).join('');
                        } else if (cellValue.formula) {
                            value = cellValue.result;
                        } else {
                            value = cellValue.text || '';
                        }
                    } else if (cellValue !== null) {
                        value = String(cellValue);
                    }
                    docData[header] = value;
                });
                masterList[docCode] = docData;
            }
        }
        
        console.log(`LOG: Ana listede ${Object.keys(masterList).length} adet belge bilgisi başarıyla yüklendi.`);
        return { masterList, headers };
    } catch (e) {
        console.error("KRİTİK HATA: Excel dosyasını ayrıştırma sırasında genel hata oluştu:", e);
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
async function createUpdatedExcelBuffer(updatedList, headers) {
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
            const result = await parseExcelData(MASTER_EXCEL_PATH);
            masterDocumentList = result.masterList;
            originalHeaders = result.headers;
            console.log(`LOG: Ana listede ${Object.keys(masterDocumentList).length} adet belge bilgisi yüklendi.`);
        } catch (e) {
            console.error(`KRİTİK HATA: Ana doküman listesi dosyası (${MASTER_EXCEL_PATH}) okunamadı. Dosyanın mevcut ve doğru yerde olduğundan emin olun.`, e);
            return res.status(500).send('Sunucu hatası: Ana doküman listesi dosyası bulunamıyor veya okunamıyor. Lütfen Render loglarını kontrol edin.');
        }

        const updatedMasterList = JSON.parse(JSON.stringify(masterDocumentList));
        const highestRevisionDocs = {};

        for (const file of uploadedFiles) {
            const originalRelativePath = file.originalname;
            const data = await extractInfo(file.path, originalRelativePath);

            if (data && data['Doküman Kodu']) {
                const docCode = data['Doküman Kodu'];
                const revisionNo = parseInt(data['Revizyon No']) || 0;

                // Eğer bu doküman kodu daha önce görülmediyse veya daha yüksek revizyon numarası varsa
                if (!highestRevisionDocs[docCode] || revisionNo > (parseInt(highestRevisionDocs[docCode]['Revizyon No']) || 0)) {
                    highestRevisionDocs[docCode] = data;
                }
            }

            try {
                fs.unlinkSync(file.path);
                console.log(`LOG: Dosya başarıyla silindi: ${file.path}`);
            } catch (e) {
                console.error(`HATA: Dosya silinirken hata oluştu: ${file.path}`, e);
            }
        }
        
        // En yüksek revizyonlu belgelerle ana listeyi güncelle
        for (const docCode in highestRevisionDocs) {
            const data = highestRevisionDocs[docCode];
            const masterDoc = updatedMasterList[docCode];

            if (masterDoc) {
                console.log(`\nLOG: Ana listede belge bilgisi bulundu (${docCode}). En yüksek revizyonlu dosya ile güncelleme yapılıyor.`);
                
                updatedMasterList[docCode]['Doküman Adı'] = data['Doküman Adı'] || masterDoc['Doküman Adı'];
                // Sorumlu Kısım güncellenmiyor
                updatedMasterList[docCode]['Revizyon No'] = data['Revizyon No'] || masterDoc['Revizyon No'];
                updatedMasterList[docCode]['Hazırlama Tarihi'] = data['Hazırlama Tarihi'] || masterDoc['Hazırlama Tarihi'];
                updatedMasterList[docCode]['Revizyon Tarihi'] = data['Revizyon Tarihi'] || masterDoc['Revizyon Tarihi'];
                
                if (data['Doküman Adı']) console.log(`LOG: ${docCode} için 'Doküman Adı' yeni dosyadan güncellendi.`);
                else console.log(`LOG: ${docCode} için 'Doküman Adı' ana listeden korundu.`);
                
                console.log(`LOG: ${docCode} için 'Sorumlu Kısım' ana listeden korundu.`);

                if (data['Revizyon No']) console.log(`LOG: ${docCode} için 'Revizyon No' yeni dosyadan güncellendi.`);
                else console.log(`LOG: ${docCode} için 'Revizyon No' ana listeden korundu.`);
            } else {
                console.log(`\nLOG: Belge ${docCode} ana listede bulunamadı. Yeni kayıt olarak ekleniyor.`);
                const newDoc = {};
                originalHeaders.forEach(header => {
                    newDoc[header] = data[header] || '';
                });
                updatedMasterList[docCode] = newDoc;
            }
        }
        
        if (Object.keys(updatedMasterList).length === 0) {
            return res.status(400).send('Hiçbir geçerli belge işlenemedi veya ana liste oluşturulamadı.');
        }

        const updatedMasterListBuffer = await createUpdatedExcelBuffer(updatedMasterList, originalHeaders);
        
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
