// Global App State
let workbook = null;
let currentTab = 'dashboard';

// Parsed Data Structures
let uretimTakipRows = []; // Üretim Takip requirements (Col A-G)
let productionLog = [];    // Üretim Takip logs entered (Col I-K)
let dosyaTakipRows = [];   // Üretim Takip summary (Col M-Q)

let montajOtomasyonLeft = [];  // MONTAJ OTOMASYON İZLEME child rows (Col A-I)
let montajOtomasyonRight = []; // MONTAJ OTOMASYON İZLEME parent rows (Col K-R)

let finalMontajLeft = [];      // FINAL MONTAJ İZLEME child rows (Col A-I)
let finalMontajRight = [];     // FINAL MONTAJ İZLEME parent rows (Col K-R)

let rotasizRows = [];          // Rotasızlar sheet rows
let rotasizHeaders = [];

let stationSheetsMap = {};     // Map of stationName -> rows array
let stationHeadersMap = {};    // Map of stationName -> headers array
let stationList = [];          // List of station sheet names
let activeStation = '';

// Helper Maps
let codeToNameMap = {};        // Map of code -> material name
let uretimListesiMap = {};     // Map of code -> Üretilecek Miktar in ÜRETİM LİSTESİ
let uretimListesiRows = [];    // Rows of ÜRETİM LİSTESİ sheet

// Pagination States
const PAGE_SIZE = 25;
let paginationState = {
    takip: { page: 1, total: 0, filtered: [] },
    assemblyLeft: { page: 1, total: 0, filtered: [] },
    assemblyRight: { page: 1, total: 0, filtered: [] },
    station: { page: 1, total: 0, filtered: [] },
    ul: { page: 1, total: 0, filtered: [] }
};

// UI Elements
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const selectFileBtn = document.getElementById('select-file-btn');
const appContainer = document.getElementById('app-container');
const loadedFileName = document.getElementById('loaded-file-name');
const changeFileBtn = document.getElementById('change-file-btn');
const exportBtn = document.getElementById('export-btn');

// Toast notification helper
function showToast(message, type = 'info') {
    const toast = document.getElementById('toast');
    toast.className = `toast show toast-${type}`;
    toast.innerHTML = `<i class="fa-solid ${type === 'success' ? 'fa-circle-check' : type === 'error' ? 'fa-triangle-exclamation' : 'fa-circle-info'}"></i> ${message}`;
    
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3500);
}

// Drag & Drop event handlers
window.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (!workbook) dropZone.classList.add('dragover');
});

window.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
});

window.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (!workbook && e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
    }
});

selectFileBtn.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

changeFileBtn.addEventListener('click', () => {
    if (confirm("Mevcut çalışmanızı kaydetmediyseniz verileriniz kaybolabilir. Devam etmek istiyor musunuz?")) {
        workbook = null;
        appContainer.style.display = 'none';
        dropZone.style.display = 'flex';
        fileInput.value = '';
    }
});

// File processing
function handleFile(file) {
    showToast("Excel dosyası okunuyor...", "info");
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, {
                type: 'array',
                cellFormula: true,
                cellNF: true,
                cellStyles: true
            });
            
            loadedFileName.textContent = `Yüklenen Dosya: ${file.name}`;
            parseWorkbook();
            
            // Switch view
            dropZone.style.display = 'none';
            appContainer.style.display = 'flex';
            
            showToast("Excel başarıyla yüklendi!", "success");
            
            // Calculate initial state
            recalculateAll();
            
            // Render initial tab
            switchTab('dashboard');
            
        } catch (err) {
            console.error(err);
            showToast("Excel dosyası işlenirken hata oluştu! Dosyanın doğru formatta olduğundan emin olun.", "error");
        }
    };
    
    reader.onerror = function() {
        showToast("Dosya okuma hatası!", "error");
    };
    
    reader.readAsArrayBuffer(file);
}

// Parse entire workbook into local JavaScript models
function parseWorkbook() {
    // Reset states
    uretimTakipRows = [];
    productionLog = [];
    dosyaTakipRows = [];
    montajOtomasyonLeft = [];
    montajOtomasyonRight = [];
    finalMontajLeft = [];
    finalMontajRight = [];
    rotasizRows = [];
    rotasizHeaders = [];
    stationSheetsMap = {};
    stationHeadersMap = {};
    stationList = [];
    codeToNameMap = {};
    uretimListesiMap = {};
    uretimListesiRows = [];

    // First scan other sheets to extract part code names and populate maps
    for (const sName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sName];
        if (!sheet || !sheet['!ref']) continue;
        
        const range = XLSX.utils.decode_range(sheet['!ref']);
        
        if (sName.includes("MONTAJ") || sName.includes("İZLEME")) {
            for (let r = range.s.r + 1; r <= range.e.r; r++) {
                const cellD = sheet[XLSX.utils.encode_cell({ r: r, c: 3 })]; // Col D (Alt Parça Kodu)
                const cellE = sheet[XLSX.utils.encode_cell({ r: r, c: 4 })]; // Col E (Alt Parça Adı)
                if (cellD && cellE && cellD.v && cellE.v) {
                    codeToNameMap[String(cellD.v).trim().toUpperCase()] = String(cellE.v).trim();
                }
            }
        } else {
            // Read headers of this sheet to find Kod and Malzeme Adı
            let headers = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: c })];
                headers.push(cell ? String(cell.v).trim().toLowerCase() : `sütun ${c+1}`);
            }
            const kodIdx = headers.indexOf('kod');
            const matIdx = headers.findIndex(h => h.includes('malzeme') && h.includes('adı'));
            const miktarIdx = headers.findIndex(h => h.includes('üretilecek') && h.includes('miktar'));
            
            if (kodIdx !== -1) {
                for (let r = range.s.r + 1; r <= range.e.r; r++) {
                    const kodCell = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + kodIdx })];
                    if (kodCell && kodCell.v !== undefined && kodCell.v !== null) {
                        const code = String(kodCell.v).trim().toUpperCase();
                        
                        if (matIdx !== -1) {
                            const matCell = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + matIdx })];
                            if (matCell && matCell.v) {
                                codeToNameMap[code] = String(matCell.v).trim();
                            }
                        }
                        
                        // If this is ÜRETİM LİSTESİ, also populate uretimListesiMap and uretimListesiRows
                        if (sName.toUpperCase().replace(/I/g, 'İ').includes('ÜRETİM LİSTESİ') || sName.toUpperCase().includes('URETIM LISTESI')) {
                            const kaynakIdx = headers.indexOf('kaynak dosya');
                            const oncelikIdx = headers.indexOf('öncelik sırası');
                            const hKodIdx = headers.indexOf('hammadde kod');
                            const hammaddeIdx = headers.indexOf('hammadde');
                            
                            const cellA = kaynakIdx !== -1 ? sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + kaynakIdx })] : null;
                            const cellB = oncelikIdx !== -1 ? sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + oncelikIdx })] : null;
                            const cellD = matIdx !== -1 ? sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + matIdx })] : null;
                            const cellE = hKodIdx !== -1 ? sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + hKodIdx })] : null;
                            const cellF = hammaddeIdx !== -1 ? sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + hammaddeIdx })] : null;
                            const miktarCell = miktarIdx !== -1 ? sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + miktarIdx })] : null;
                            const qty = miktarCell ? parseFloat(miktarCell.v) || 0 : 0;
                            
                            if (uretimListesiMap[code] === undefined) {
                                uretimListesiMap[code] = qty;
                            }
                            
                            uretimListesiRows.push({
                                rowIndex: r + 1,
                                kaynak: cellA ? String(cellA.v).trim() : '',
                                oncelik: cellB ? parseInt(cellB.v) || 999 : 999,
                                kod: code,
                                malzeme: cellD ? String(cellD.v).trim() : '',
                                hKod: cellE ? String(cellE.v).trim() : '',
                                hammadde: cellF ? String(cellF.v).trim() : '',
                                uretilecek: qty
                            });
                        }
                    }
                }
            }
        }
    }

    // 1. Parse Üretim Takip
    const utSheet = workbook.Sheets["Üretim Takip"];
    if (utSheet && utSheet['!ref']) {
        const range = XLSX.utils.decode_range(utSheet['!ref']);
        
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            // Col A-G: Requirements
            const cellA = utSheet[XLSX.utils.encode_cell({ r: r, c: 0 })]; // Col A (KAYNAK DOSYA)
            const cellB = utSheet[XLSX.utils.encode_cell({ r: r, c: 1 })]; // Col B (Öncelik Sırası)
            const cellC = utSheet[XLSX.utils.encode_cell({ r: r, c: 2 })]; // Col C (Kod)
            const cellD = utSheet[XLSX.utils.encode_cell({ r: r, c: 3 })]; // Col D (Üretilecek Miktar)
            
            if (cellA && cellA.v && cellC && cellC.v) {
                const requiredQty = cellD ? parseFloat(cellD.v) || 0 : 0;
                uretimTakipRows.push({
                    rowIndex: r + 1, // Excel row is 1-indexed
                    kaynak: String(cellA.v).trim(),
                    oncelik: cellB ? parseInt(cellB.v) || 999 : 999,
                    kod: String(cellC.v).trim().toUpperCase(),
                    uretilecek: requiredQty,
                    uretilen: 0,
                    kalan: requiredQty,
                    tamamlanma: 0
                });
            }

            // Col I-K: Production Log
            const cellI = utSheet[XLSX.utils.encode_cell({ r: r, c: 8 })]; // Col I (ÜRETİLEN KOD)
            const cellJ = utSheet[XLSX.utils.encode_cell({ r: r, c: 9 })]; // Col J (ÜRETİM ADEDİ)
            
            if (cellI && cellI.v !== undefined && cellI.v !== null && String(cellI.v).trim() !== "") {
                const prodQty = cellJ ? parseFloat(cellJ.v) || 0 : 0;
                productionLog.push({
                    rowIndex: r + 1,
                    kod: String(cellI.v).trim().toUpperCase(),
                    adet: prodQty,
                    fazla: 0
                });
            }

            // Col M-Q: Kaynak Dosya stats
            const cellM = utSheet[XLSX.utils.encode_cell({ r: r, c: 12 })]; // Col M (DOSYA BAZLI TAKİP)
            const cellN = utSheet[XLSX.utils.encode_cell({ r: r, c: 13 })]; // Col N (TOPLAM (Kalem))
            
            if (cellM && cellM.v) {
                dosyaTakipRows.push({
                    rowIndex: r + 1,
                    kaynak: String(cellM.v).trim(),
                    toplam: cellN ? parseInt(cellN.v) || 0 : 0,
                    hazir: 0,
                    eksik: 0,
                    tamamlanma: 0
                });
            }
        }

        // Sort Kaynak Dosya rows numerically by priority prefix (e.g. 1, 2, ..., 12, ..., 100)
        dosyaTakipRows.sort((a, b) => {
            const numA = parseInt(a.kaynak.split('-')[0].trim()) || 99999;
            const numB = parseInt(b.kaynak.split('-')[0].trim()) || 99999;
            return numA - numB;
        });
    }

    // 2. Parse MONTAJ OTOMASYON İZLEME
    const moSheet = workbook.Sheets["MONTAJ OTOMASYON İZLEME"];
    if (moSheet && moSheet['!ref']) {
        const range = XLSX.utils.decode_range(moSheet['!ref']);
        
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            // Left Table: Col A-I
            const cellA = moSheet[XLSX.utils.encode_cell({ r: r, c: 0 })]; // Kaynak Dosya
            const cellB = moSheet[XLSX.utils.encode_cell({ r: r, c: 1 })]; // Üst Montaj Kodu
            const cellC = moSheet[XLSX.utils.encode_cell({ r: r, c: 2 })]; // Üst Montaj Adı
            const cellD = moSheet[XLSX.utils.encode_cell({ r: r, c: 3 })]; // Alt Parça Kodu
            const cellE = moSheet[XLSX.utils.encode_cell({ r: r, c: 4 })]; // Alt Parça Adı
            const cellF = moSheet[XLSX.utils.encode_cell({ r: r, c: 5 })]; // Gereken Miktar
            
            if (cellA && cellA.v && cellD && cellD.v) {
                const cellI = moSheet[XLSX.utils.encode_cell({ r: r, c: 8 })]; // Alt Parça Limit (Formula)
                const formulaStr = cellI ? cellI.f || '' : '';
                
                let valEldeki = 0, valIhtiyac = 0, birimIhtiyac = 1;
                if (formulaStr) {
                    const match = formulaStr.match(/MAX\(\s*0\s*,\s*(-?\d+\.?\d*)\s*\+\s*SUMIF.*-\s*(-?\d+\.?\d*)\)\s*\/\s*(-?\d+\.?\d*)/i);
                    if (match) {
                        valEldeki = parseFloat(match[1]);
                        valIhtiyac = parseFloat(match[2]);
                        birimIhtiyac = parseFloat(match[3]);
                    }
                }
                
                montajOtomasyonLeft.push({
                    rowIndex: r + 1,
                    kaynak: String(cellA.v).trim(),
                    ustKod: String(cellB ? cellB.v : '').trim(),
                    ustAd: String(cellC ? cellC.v : '').trim(),
                    altKod: String(cellD.v).trim().toUpperCase(),
                    altAd: String(cellE ? cellE.v : '').trim(),
                    gereken: cellF ? parseFloat(cellF.v) || 0 : 0,
                    uretilen: 0,
                    tamamlanma: 0,
                    limit: 0,
                    valEldeki: valEldeki,
                    valIhtiyac: valIhtiyac,
                    birimIhtiyac: birimIhtiyac
                });
            }

            // Right Table: Col K-R
            const cellK = moSheet[XLSX.utils.encode_cell({ r: r, c: 10 })]; // Kaynak Dosya
            const cellL = moSheet[XLSX.utils.encode_cell({ r: r, c: 11 })]; // Üst Montaj Kodu
            const cellM = moSheet[XLSX.utils.encode_cell({ r: r, c: 12 })]; // Üst Montaj Adı
            const cellN = moSheet[XLSX.utils.encode_cell({ r: r, c: 13 })]; // Gereken Çeşit
            const cellQ = moSheet[XLSX.utils.encode_cell({ r: r, c: 16 })]; // Toplam Parça Adedi
            
            if (cellK && cellK.v && cellL && cellL.v) {
                const cellO = moSheet[XLSX.utils.encode_cell({ r: r, c: 14 })]; // Tamamlanan Çeşit (Formula)
                const formulaO = cellO ? cellO.f || '' : '';
                let countifsConstant = 0;
                if (formulaO) {
                    const match = formulaO.match(/^=?\s*(\d+)\s*\+/);
                    if (match) countifsConstant = parseInt(match[1]);
                }

                const cellR = moSheet[XLSX.utils.encode_cell({ r: r, c: 17 })]; // Ek Toplanabilir (Set) (Formula)
                const formulaR = cellR ? cellR.f || '' : '';
                let inStockTimLimit = Infinity;
                if (formulaR) {
                    const match = formulaR.match(/MIN\(\s*(-?\d+)\s*,/i);
                    if (match) {
                        inStockTimLimit = parseInt(match[1]);
                    } else {
                        const matchFallback = formulaR.match(/,\s*(-?\d+)\s*\)\s*,\s*0\s*\)$/);
                        if (matchFallback) {
                            inStockTimLimit = parseInt(matchFallback[1]);
                        }
                    }
                }
                
                montajOtomasyonRight.push({
                    rowIndex: r + 1,
                    kaynak: String(cellK.v).trim(),
                    ustKod: String(cellL.v).trim(),
                    ustAd: String(cellM ? cellM.v : '').trim(),
                    gerekenCesit: cellN ? parseInt(cellN.v) || 0 : 0,
                    tamamlananCesit: 0,
                    tamamlanma: 0,
                    toplamAdet: cellQ ? parseInt(cellQ.v) || 0 : 0,
                    limit: 0,
                    inStockTimLimit: inStockTimLimit,
                    countifsConstant: countifsConstant
                });
            }
        }
    }

    // 3. Parse FINAL MONTAJ İZLEME
    const fmSheet = workbook.Sheets["FINAL MONTAJ İZLEME"];
    if (fmSheet && fmSheet['!ref']) {
        const range = XLSX.utils.decode_range(fmSheet['!ref']);
        
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            // Left Table: Col A-I
            const cellA = fmSheet[XLSX.utils.encode_cell({ r: r, c: 0 })];
            const cellB = fmSheet[XLSX.utils.encode_cell({ r: r, c: 1 })];
            const cellC = fmSheet[XLSX.utils.encode_cell({ r: r, c: 2 })];
            const cellD = fmSheet[XLSX.utils.encode_cell({ r: r, c: 3 })];
            const cellE = fmSheet[XLSX.utils.encode_cell({ r: r, c: 4 })];
            const cellF = fmSheet[XLSX.utils.encode_cell({ r: r, c: 5 })];
            
            if (cellA && cellA.v && cellD && cellD.v) {
                const cellI = fmSheet[XLSX.utils.encode_cell({ r: r, c: 8 })];
                const formulaStr = cellI ? cellI.f || '' : '';
                
                let valEldeki = 0, valIhtiyac = 0, birimIhtiyac = 1;
                if (formulaStr) {
                    const match = formulaStr.match(/MAX\(\s*0\s*,\s*(-?\d+\.?\d*)\s*\+\s*SUMIF.*-\s*(-?\d+\.?\d*)\)\s*\/\s*(-?\d+\.?\d*)/i);
                    if (match) {
                        valEldeki = parseFloat(match[1]);
                        valIhtiyac = parseFloat(match[2]);
                        birimIhtiyac = parseFloat(match[3]);
                    }
                }
                
                finalMontajLeft.push({
                    rowIndex: r + 1,
                    kaynak: String(cellA.v).trim(),
                    ustKod: String(cellB ? cellB.v : '').trim(),
                    ustAd: String(cellC ? cellC.v : '').trim(),
                    altKod: String(cellD.v).trim().toUpperCase(),
                    altAd: String(cellE ? cellE.v : '').trim(),
                    gereken: cellF ? parseFloat(cellF.v) || 0 : 0,
                    uretilen: 0,
                    tamamlanma: 0,
                    limit: 0,
                    valEldeki: valEldeki,
                    valIhtiyac: valIhtiyac,
                    birimIhtiyac: birimIhtiyac
                });
            }

            // Right Table: Col K-R
            const cellK = fmSheet[XLSX.utils.encode_cell({ r: r, c: 10 })];
            const cellL = fmSheet[XLSX.utils.encode_cell({ r: r, c: 11 })];
            const cellM = fmSheet[XLSX.utils.encode_cell({ r: r, c: 12 })];
            const cellN = fmSheet[XLSX.utils.encode_cell({ r: r, c: 13 })];
            const cellQ = fmSheet[XLSX.utils.encode_cell({ r: r, c: 16 })];
            
            if (cellK && cellK.v && cellL && cellL.v) {
                const cellO = fmSheet[XLSX.utils.encode_cell({ r: r, c: 14 })];
                const formulaO = cellO ? cellO.f || '' : '';
                let countifsConstant = 0;
                if (formulaO) {
                    const match = formulaO.match(/^=?\s*(\d+)\s*\+/);
                    if (match) countifsConstant = parseInt(match[1]);
                }

                const cellR = fmSheet[XLSX.utils.encode_cell({ r: r, c: 17 })];
                const formulaR = cellR ? cellR.f || '' : '';
                let inStockTimLimit = Infinity;
                if (formulaR) {
                    const match = formulaR.match(/MIN\(\s*(-?\d+)\s*,/i);
                    if (match) {
                        inStockTimLimit = parseInt(match[1]);
                    } else {
                        const matchFallback = formulaR.match(/,\s*(-?\d+)\s*\)\s*,\s*0\s*\)$/);
                        if (matchFallback) {
                            inStockTimLimit = parseInt(matchFallback[1]);
                        }
                    }
                }
                
                finalMontajRight.push({
                    rowIndex: r + 1,
                    kaynak: String(cellK.v).trim(),
                    ustKod: String(cellL.v).trim(),
                    ustAd: String(cellM ? cellM.v : '').trim(),
                    gerekenCesit: cellN ? parseInt(cellN.v) || 0 : 0,
                    tamamlananCesit: 0,
                    tamamlanma: 0,
                    toplamAdet: cellQ ? parseInt(cellQ.v) || 0 : 0,
                    limit: 0,
                    inStockTimLimit: inStockTimLimit,
                    countifsConstant: countifsConstant
                });
            }
        }
    }

    // 4. Parse Rotasızlar Sheet
    const rSheet = workbook.Sheets["Rotasızlar"];
    if (rSheet && rSheet['!ref']) {
        const range = XLSX.utils.decode_range(rSheet['!ref']);
        
        // Headers
        for (let c = range.s.c; c <= range.e.c; c++) {
            const hCell = rSheet[XLSX.utils.encode_cell({ r: range.s.r, c: c })];
            rotasizHeaders.push(hCell ? String(hCell.v).trim() : `Sütun ${c+1}`);
        }

        // Data rows
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            let rowObj = {};
            let hasValue = false;
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = rSheet[XLSX.utils.encode_cell({ r: r, c: c })];
                const headerName = rotasizHeaders[c - range.s.c];
                rowObj[headerName] = cell && cell.v !== undefined ? cell.v : '';
                if (cell && cell.v !== undefined && cell.v !== '') hasValue = true;
            }
            if (hasValue) {
                rotasizRows.push(rowObj);
            }
        }
    }

    // 5. Parse Station Sheets
    const excludedSheets = [
        "ÜRETİM LİSTESİ", "Tüm Veriler", "HAMMADDE", "HAMMADDE SİPARİŞ",
        "Üretim Takip", "Rotasızlar", "MONTAJ OTOMASYON İZLEME", "FINAL MONTAJ İZLEME"
    ];

    for (const sName of workbook.SheetNames) {
        if (excludedSheets.includes(sName)) continue;
        
        const sheet = workbook.Sheets[sName];
        if (!sheet || !sheet['!ref']) continue;
        
        const range = XLSX.utils.decode_range(sheet['!ref']);
        let headers = [];
        let rows = [];
        
        // Read Headers
        for (let c = range.s.c; c <= range.e.c; c++) {
            const hCell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: c })];
            headers.push(hCell ? String(hCell.v).trim() : `Sütun ${c+1}`);
        }

        // Read Rows
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            let rowObj = {};
            let hasValue = false;
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: r, c: c })];
                const headerName = headers[c - range.s.c];
                rowObj[headerName] = cell && cell.v !== undefined ? cell.v : '';
                if (cell && cell.v !== undefined && cell.v !== '') hasValue = true;
            }
            if (hasValue) {
                // Filter out summary/empty rows
                const code = String(rowObj['Kod'] || '').trim();
                const priorityStr = String(rowObj['Öncelik Sırası'] || '').trim().toUpperCase();
                
                // Skip if Kod is empty, or if Öncelik Sırası contains summary keywords
                if (!code || 
                    priorityStr.includes('İSTASYON') || 
                    priorityStr.includes('SAYISI') || 
                    priorityStr.includes('TOPLAM') || 
                    priorityStr.includes('SÜRE') || 
                    priorityStr.includes('GÜNÜ')) {
                    continue;
                }
                
                rowObj.rowIndex = r + 1; // Excel row is 1-indexed
                rows.push(rowObj);
            }
        }

        if (rows.length > 0) {
            stationSheetsMap[sName] = rows;
            stationHeadersMap[sName] = headers;
            stationList.push(sName);
        }
    }

    // Sort stations based on custom order in localStorage, or alphabetically as fallback
    const savedOrderStr = localStorage.getItem('customStationOrder');
    if (savedOrderStr) {
        try {
            const savedOrder = JSON.parse(savedOrderStr);
            stationList.sort((a, b) => {
                const idxA = savedOrder.indexOf(a);
                const idxB = savedOrder.indexOf(b);
                if (idxA !== -1 && idxB !== -1) return idxA - idxB;
                if (idxA !== -1) return -1;
                if (idxB !== -1) return 1;
                return a.localeCompare(b, undefined, { numeric: true });
            });
        } catch (err) {
            stationList.sort();
        }
    } else {
        stationList.sort();
    }
    if (stationList.length > 0) activeStation = stationList[0];
}

// Live spreadsheet calculation engine (SUMIF, SUMIFS, FIFO matching, limits, sets)
function recalculateAll() {
    // 1. Calculate total produced quantity for each part code from the production log
    const totalProducedMap = {};
    for (const log of productionLog) {
        const code = log.kod;
        totalProducedMap[code] = (totalProducedMap[code] || 0.0) + parseFloat(log.adet);
    }

    // 2. Allocate produced quantities in FIFO order in Üretim Takip requirements
    const allocatedRequirements = {};
    for (const row of uretimTakipRows) {
        const code = row.kod;
        const totalProd = totalProducedMap[code] || 0.0;
        const previouslyAllocated = allocatedRequirements[code] || 0.0;
        const available = Math.max(0.0, totalProd - previouslyAllocated);
        
        row.uretilen = Math.min(row.uretilecek, available);
        allocatedRequirements[code] = previouslyAllocated + row.uretilen;
        
        row.kalan = Math.max(0.0, row.uretilecek - row.uretilen);
        row.tamamlanma = row.uretilecek > 0 ? (row.uretilen / row.uretilecek) : 0.0;
    }

    // Update log's Fazla Üretim
    // Fazla Üretim = MAX(0, totalProduced(Code) - totalRequired(Code))
    const totalRequiredMap = {};
    for (const row of uretimTakipRows) {
        totalRequiredMap[row.kod] = (totalRequiredMap[row.kod] || 0.0) + row.uretilecek;
    }
    
    for (const log of productionLog) {
        const req = totalRequiredMap[log.kod] || 0.0;
        const prod = totalProducedMap[log.kod] || 0.0;
        log.fazla = Math.max(0.0, prod - req);
    }

    // 3. Update DOSYA BAZLI TAKİP statistics
    // Cache uretimTakipRows by kaynak prefix matching
    for (const dRow of dosyaTakipRows) {
        const kName = dRow.kaynak;
        // Filter rows where uretimTakipRow.kaynak contains kName
        const matched = uretimTakipRows.filter(r => r.kaynak.includes(kName));
        if (!dRow.toplam || dRow.toplam === 0) {
            dRow.toplam = matched.length;
        }
        dRow.eksik = matched.filter(r => r.kalan > 0).length;
        dRow.hazir = Math.max(0, dRow.toplam - dRow.eksik);
        
        const sumReq = matched.reduce((sum, r) => sum + r.uretilecek, 0.0);
        const sumProd = matched.reduce((sum, r) => sum + r.uretilen, 0.0);
        
        const ratio = sumReq > 0 ? (sumProd / sumReq) : (dRow.eksik > 0 ? 0.0 : 1.0);
        dRow.tamamlanma = dRow.toplam > 0 ? ((dRow.hazir + dRow.eksik * ratio) / dRow.toplam) : 0.0;
    }

    // 4. Helper function to recalculate a Montaj sheet
    function calculateMontajSheet(leftList, rightList) {
        // Calculate sum of E (Üretilen Miktar) in Üretim Takip for each code
        const sumTakipUretilen = {};
        for (const row of uretimTakipRows) {
            sumTakipUretilen[row.kod] = (sumTakipUretilen[row.kod] || 0.0) + row.uretilen;
        }

        // Allocate this sum to child rows in FIFO order
        const allocatedMontaj = {};
        for (const row of leftList) {
            const code = row.altKod;
            const totalAvailable = sumTakipUretilen[code] || 0.0;
            const previouslyAllocated = allocatedMontaj[code] || 0.0;
            const available = Math.max(0.0, totalAvailable - previouslyAllocated);
            
            row.uretilen = Math.min(row.gereken, available);
            allocatedMontaj[code] = previouslyAllocated + row.uretilen;
            row.tamamlanma = row.gereken > 0 ? (row.uretilen / row.gereken) : 0.0;

            // Alt Parça Limit calculation
            // =IF(G>=F, INT(MAX(0, eldeki + log_prod - ihtiyac) / birim_ihtiyac), 0)
            if (row.uretilen >= row.gereken && row.birimIhtiyac > 0) {
                const logProd = totalProducedMap[row.altKod] || 0.0;
                row.limit = Math.floor(Math.max(0.0, row.valEldeki + logProd - row.valIhtiyac) / row.birimIhtiyac);
            } else {
                row.limit = 0;
            }
        }

        // Update Right Table: Parent assemblies
        for (const parent of rightList) {
            // Find child rows under this parent in this source file
            const kids = leftList.filter(l => l.kaynak === parent.kaynak && l.ustKod === parent.ustKod);
            
            const completedCountInLeft = kids.filter(k => k.tamamlanma >= 1.0).length;
            parent.tamamlananCesit = parent.countifsConstant + completedCountInLeft;
            parent.tamamlanma = parent.gerekenCesit > 0 ? (parent.tamamlananCesit / parent.gerekenCesit) : 0.0;
            
            // Ek Toplanabilir (Set) limit calculation
            // =IF(P=1, IF(COUNTIFS(A:A, K, B:B, L)>0, MIN(inStockLimit, MINIFS(I:I, A:A, K, B:B, L)), inStockLimit), 0)
            if (parent.tamamlanma >= 1.0) {
                if (kids.length > 0) {
                    const minKidLimit = Math.min(...kids.map(k => k.limit));
                    parent.limit = Math.min(parent.inStockTimLimit, minKidLimit);
                } else {
                    parent.limit = parent.inStockTimLimit === Infinity ? 0 : parent.inStockTimLimit;
                }
            } else {
                parent.limit = 0;
            }
        }
    }

    calculateMontajSheet(montajOtomasyonLeft, montajOtomasyonRight);
    calculateMontajSheet(finalMontajLeft, finalMontajRight);

    // 5. Update Station Sheets values
    for (const [stName, rows] of Object.entries(stationSheetsMap)) {
        for (const row of rows) {
            const code = String(row['Kod'] || '').trim().toUpperCase();
            const uretilecekMiktarVal = uretimListesiMap[code] || 0;
            
            // Match keys case-insensitively
            let uMiktarKey = Object.keys(row).find(k => k.toLowerCase() === 'üretilecek miktar');
            let tHammaddeKey = Object.keys(row).find(k => k.toLowerCase() === 'toplam hammadde miktarı');
            let hMiktarKey = Object.keys(row).find(k => k.toLowerCase() === 'hammadde miktar');
            
            if (uMiktarKey) {
                row[uMiktarKey] = uretilecekMiktarVal;
            }
            if (tHammaddeKey && hMiktarKey) {
                const hMiktarVal = parseFloat(row[hMiktarKey]) || 0;
                row[tHammaddeKey] = uretilecekMiktarVal * hMiktarVal;
            }
        }
    }
}

// -------------------------------------------------------------
// TAB SWITCH & NAVIGATION
// -------------------------------------------------------------
function switchTab(tabId) {
    document.querySelectorAll('.sidebar-nav .nav-item').forEach(item => {
        item.classList.remove('active');
        if (item.getAttribute('data-tab') === tabId) {
            item.classList.add('active');
        }
    });

    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });

    const activeContent = document.getElementById(`tab-${tabId}`);
    if (activeContent) activeContent.classList.add('active');
    
    currentTab = tabId;
    renderTab(tabId);
}

document.querySelectorAll('.sidebar-nav .nav-item').forEach(item => {
    item.addEventListener('click', () => {
        switchTab(item.getAttribute('data-tab'));
    });
});

function renderTab(tabId) {
    if (tabId === 'dashboard') {
        renderDashboard();
    } else if (tabId === 'production') {
        renderProductionTab();
    } else if (tabId === 'uretim-listesi') {
        renderUretimListesiTab();
    } else if (tabId === 'assembly') {
        renderAssemblyTab();
    } else if (tabId === 'rotasizlar') {
        renderRotasizlarTab();
    } else if (tabId === 'stations') {
        renderStationsTab();
    }
}

// -------------------------------------------------------------
// 1. DASHBOARD VIEW RENDERING
// -------------------------------------------------------------
function renderDashboard() {
    // 1. Calculate KPIs
    const totalParts = dosyaTakipRows.reduce((sum, f) => sum + f.toplam, 0);
    const readyParts = dosyaTakipRows.reduce((sum, f) => sum + f.hazir, 0);
    const missingParts = dosyaTakipRows.reduce((sum, f) => sum + f.eksik, 0);
    const totalFiles = dosyaTakipRows.length;

    document.getElementById('kpi-total-parts').textContent = totalParts.toLocaleString();
    document.getElementById('kpi-ready-parts').textContent = readyParts.toLocaleString();
    document.getElementById('kpi-missing-parts').textContent = missingParts.toLocaleString();
    document.getElementById('kpi-total-files').textContent = totalFiles.toLocaleString();

    // 2. Filter and Render files
    const searchVal = document.getElementById('dashboard-search').value.toLowerCase().trim();
    let filteredFiles = dosyaTakipRows;
    if (searchVal) {
        filteredFiles = dosyaTakipRows.filter(f => f.kaynak.toLowerCase().includes(searchVal));
    }

    const container = document.getElementById('dashboard-files-list');
    container.innerHTML = '';

    if (filteredFiles.length === 0) {
        container.innerHTML = '<div style="grid-column: 1/-1; text-align: center; color: var(--text-dim); padding: 40px;">Kayıt bulunamadı.</div>';
        return;
    }

    filteredFiles.forEach(f => {
        const pct = Math.round(f.tamamlanma * 100);
        let colorClass = 'var(--danger)';
        let glowClass = 'var(--danger-glow)';
        if (pct >= 100) {
            colorClass = 'var(--success)';
            glowClass = 'var(--success-glow)';
        } else if (pct >= 50) {
            colorClass = 'var(--warning)';
            glowClass = 'var(--warning-glow)';
        }

        const card = document.createElement('div');
        card.className = 'file-progress-card';
        card.innerHTML = `
            <div class="card-title-row">
                <span class="file-name-label" title="${f.kaynak}">${f.kaynak}</span>
                <span class="file-pct" style="color: ${colorClass};">${pct}%</span>
            </div>
            <div class="progress-container">
                <div class="progress-fill" style="width: ${pct}%; background-color: ${colorClass}; box-shadow: 0 0 8px ${glowClass};"></div>
            </div>
            <div class="card-stats-row">
                <div class="stat-item"><i class="fa-solid fa-check text-green"></i> <span>${f.hazir} / ${f.toplam} hazır</span></div>
                <div class="stat-item"><i class="fa-solid fa-hourglass-half text-orange"></i> <span>${f.eksik} eksik</span></div>
            </div>
        `;
        
        card.style.cursor = 'pointer';
        card.addEventListener('click', () => {
            openDetailsModal(f.kaynak);
        });

        // Add double click listener to filter production log by this source file
        card.addEventListener('dblclick', (e) => {
            e.stopPropagation(); // prevent modal opening again
            document.getElementById('takip-search').value = f.kaynak;
            switchTab('production');
        });

        container.appendChild(card);
    });
}

// Details Modal State & Logic
let currentModalKaynak = null;

function openDetailsModal(kaynakName) {
    console.log("openDetailsModal called with:", kaynakName);
    currentModalKaynak = kaynakName;
    
    const modalEl = document.getElementById('details-modal');
    if (modalEl) {
        modalEl.style.display = 'flex';
    }
    
    const searchEl = document.getElementById('modal-search');
    if (searchEl) {
        searchEl.value = '';
    }
    
    try {
        renderModalData();
    } catch (err) {
        console.error("Error rendering modal data:", err);
    }
}

function renderModalData() {
    if (!currentModalKaynak) return;
    
    // Find matching dosyaTakipRow
    const dRow = dosyaTakipRows.find(f => f.kaynak === currentModalKaynak);
    if (!dRow) {
        console.warn("Could not find dosyaTakipRow for kaynak:", currentModalKaynak);
        return;
    }

    // Filter uretimTakipRows matching this source
    const matchedRows = uretimTakipRows.filter(r => r.kaynak && String(r.kaynak).includes(currentModalKaynak));
    
    // Calculate total quantities
    const totalQty = matchedRows.reduce((sum, r) => sum + (parseFloat(r.uretilecek) || 0), 0);
    const producedQty = matchedRows.reduce((sum, r) => sum + (parseFloat(r.uretilen) || 0), 0);
    
    // Set headers / stats safely
    const titleEl = document.getElementById('modal-title');
    if (titleEl) {
        titleEl.innerHTML = `<i class="fa-solid fa-file-invoice text-blue"></i> ${currentModalKaynak} - Dosya Detayları`;
    }
    
    const varietyEl = document.getElementById('modal-stat-variety');
    if (varietyEl) {
        varietyEl.textContent = `${dRow.hazir || 0} / ${dRow.toplam || 0} hazır`;
    }
    
    const qtyEl = document.getElementById('modal-stat-qty');
    if (qtyEl) {
        qtyEl.textContent = `${producedQty.toLocaleString()} / ${totalQty.toLocaleString()} adet`;
    }
    
    const missingCountEl = document.getElementById('modal-stat-missing-count');
    if (missingCountEl) {
        missingCountEl.textContent = `${dRow.eksik || 0} eksik kalem`;
    }

    // Filter missing rows
    let missingRows = matchedRows.filter(r => (parseFloat(r.kalan) || 0) > 0);
    
    // Filter by modal search value safely
    const searchInput = document.getElementById('modal-search');
    const searchVal = searchInput ? searchInput.value.toLowerCase().trim() : '';
    if (searchVal) {
        missingRows = missingRows.filter(r => {
            const code = String(r.kod || '').toLowerCase();
            const name = String(codeToNameMap[r.kod] || '').toLowerCase();
            return code.includes(searchVal) || name.includes(searchVal);
        });
    }

    const tbody = document.getElementById('modal-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (missingRows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center" style="color: var(--text-dim); padding: 20px;">Eksik parça bulunamadı.</td></tr>';
        return;
    }

    missingRows.forEach(r => {
        const name = codeToNameMap[r.kod] || '-';
        
        // Find matching stations safely
        const stations = [];
        if (stationSheetsMap) {
            for (const [stName, rows] of Object.entries(stationSheetsMap)) {
                if (Array.isArray(rows)) {
                    const found = rows.some(sr => sr && String(sr['Kod'] || '').trim().toUpperCase() === String(r.kod).trim().toUpperCase());
                    if (found) {
                        stations.push(stName);
                    }
                }
            }
        }
        
        let stationsHtml = '';
        if (stations.length === 0) {
            stationsHtml = '<span style="color: var(--text-dim); font-size: 11px;">-</span>';
        } else {
            stationsHtml = stations.map(st => `
                <span class="badge" style="background: rgba(139, 92, 246, 0.15); color: #c084fc; border: 1px solid rgba(139, 92, 246, 0.25); margin: 2px; font-size: 10px; padding: 2px 6px; font-weight: 600;">${st}</span>
            `).join('');
        }

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="font-weight: 700; color: white;">${r.kod || '-'}</td>
            <td style="color: var(--text-muted); max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${name}">${name}</td>
            <td>${stationsHtml}</td>
            <td class="text-right" style="font-weight: 600;">${r.uretilecek || 0}</td>
            <td class="text-right" style="color: ${(r.uretilen || 0) > 0 ? 'var(--success)' : 'var(--text-dim)'}; font-weight: 600;">${r.uretilen || 0}</td>
            <td class="text-right" style="color: var(--warning); font-weight: 700;">${r.kalan || 0}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Modal Event Listeners
const closeBtn = document.getElementById('modal-close-btn');
if (closeBtn) {
    closeBtn.addEventListener('click', () => {
        const modal = document.getElementById('details-modal');
        if (modal) modal.style.display = 'none';
        currentModalKaynak = null;
    });
}

const modalOverlay = document.getElementById('details-modal');
if (modalOverlay) {
    modalOverlay.addEventListener('click', (e) => {
        if (e.target === modalOverlay) {
            modalOverlay.style.display = 'none';
            currentModalKaynak = null;
        }
    });
}

const modalSearch = document.getElementById('modal-search');
if (modalSearch) {
    modalSearch.addEventListener('input', renderModalData);
}

document.getElementById('dashboard-search').addEventListener('input', renderDashboard);

// -------------------------------------------------------------
// 2. PRODUCTION VIEW & LOGIC
// -------------------------------------------------------------
let activeTabAssembly = 'otomasyon'; // otomasyon or final
let expandedParentAssembly = null;   // selected parent assembly key (ustKod___kaynak) for inline accordion

function renderProductionTab() {
    // Fill autocomplete suggestions list
    initAutocomplete();
    
    // Render Log
    renderProductionLog();

    // Filter & Paginate Requirements Table
    filterAndPaginateTakipTable();
}

function initAutocomplete() {
    const codeInput = document.getElementById('prod-code');
    const autoList = document.getElementById('autocomplete-list');
    
    // Extract unique codes
    const uniqueCodes = [...new Set(uretimTakipRows.map(r => r.kod))].sort();

    codeInput.addEventListener('input', function() {
        const val = this.value.toUpperCase().trim();
        autoList.innerHTML = '';
        if (!val) return;

        let suggestions = uniqueCodes.filter(c => c.includes(val)).slice(0, 10);
        
        suggestions.forEach(s => {
            const div = document.createElement('div');
            const name = codeToNameMap[s] || 'İsim Bilgisi Yok';
            div.innerHTML = `<span class="auto-code">${s}</span><span class="auto-name" title="${name}">${name}</span>`;
            
            div.addEventListener('click', () => {
                codeInput.value = s;
                autoList.innerHTML = '';
                showQuickPartInfo(s);
            });
            autoList.appendChild(div);
        });
    });

    document.addEventListener('click', (e) => {
        if (!autoList.contains(e.target) && e.target !== codeInput) {
            autoList.innerHTML = '';
        }
    });
}

function showQuickPartInfo(code) {
    const box = document.getElementById('quick-part-info');
    code = code.toUpperCase().trim();
    
    const reqs = uretimTakipRows.filter(r => r.kod === code);
    if (reqs.length === 0) {
        box.innerHTML = `
            <div class="part-info-details">
                <div class="part-info-title">${code}</div>
                <div style="color: var(--danger); font-size: 13px; text-align: center; padding: 20px 0;">
                    <i class="fa-solid fa-circle-xmark" style="font-size:24px; margin-bottom:8px; display:block;"></i>
                    Bu kod Üretim Takip listesinde yer almıyor!
                </div>
            </div>
        `;
        return;
    }

    const name = codeToNameMap[code] || 'Tanımsız Malzeme';
    const totalReq = reqs.reduce((sum, r) => sum + r.uretilecek, 0.0);
    const totalProd = reqs.reduce((sum, r) => sum + r.uretilen, 0.0);
    const totalRem = Math.max(0.0, totalReq - totalProd);
    const overallPct = totalReq > 0 ? Math.round((totalProd / totalReq) * 100) : 0;
    
    // Find stations requiring this code
    let stations = [];
    for (const [stName, rows] of Object.entries(stationSheetsMap)) {
        if (rows.some(r => String(r['Kod'] || '').trim().toUpperCase() === code)) {
            stations.push(stName);
        }
    }

    box.innerHTML = `
        <div class="part-info-details">
            <div class="part-info-title">${code}</div>
            <div class="part-info-row">
                <span class="part-info-label">Malzeme Adı:</span>
                <span class="part-info-val" style="max-width: 65%; text-align:right;">${name}</span>
            </div>
            <div class="part-info-row">
                <span class="part-info-label">Toplam İhtiyaç:</span>
                <span class="part-info-val text-blue">${totalReq} Adet</span>
            </div>
            <div class="part-info-row">
                <span class="part-info-label">Toplam Üretilen:</span>
                <span class="part-info-val text-green">${totalProd} Adet</span>
            </div>
            <div class="part-info-row">
                <span class="part-info-label">Kalan İhtiyaç:</span>
                <span class="part-info-val text-orange">${totalRem} Adet</span>
            </div>
            <div class="part-info-row">
                <span class="part-info-label">Genel Durum:</span>
                <span class="badge ${overallPct >= 100 ? 'badge-success' : overallPct > 0 ? 'badge-warning' : 'badge-danger'}">${overallPct}% Tamamlandı</span>
            </div>
            <div class="part-info-row" style="margin-top: 5px; flex-direction:column; gap:6px; border-top:1px solid var(--border-color); padding-top:8px;">
                <span class="part-info-label" style="font-weight:600;">İşleneceği İstasyonlar:</span>
                <div style="display:flex; flex-wrap:wrap; gap:5px; margin-top:3px;">
                    ${stations.length > 0 
                        ? stations.map(s => `<span class="badge" style="background:rgba(255,255,255,0.06); color:var(--text-main); font-size:10px;">${s}</span>`).join('') 
                        : '<span style="color:var(--text-dim); font-size:12px;">Tanımlı istasyon yok. (Büyük ihtimalle Montaj)</span>'
                    }
                </div>
            </div>
        </div>
    `;
}

// Add production log entry
document.getElementById('production-form').addEventListener('submit', function(e) {
    e.preventDefault();
    const code = document.getElementById('prod-code').value.toUpperCase().trim();
    const qty = parseFloat(document.getElementById('prod-qty').value) || 0;

    if (!code || qty <= 0) return;

    // Check if code has requirements
    const hasReq = uretimTakipRows.some(r => r.kod === code);
    if (!hasReq) {
        if (!confirm(`Dikkat: "${code}" parça kodu Üretim Takip sayfasında bulunamadı. Yine de giriş yapmak istiyor musunuz?`)) {
            return;
        }
    }

    // Append to production log
    productionLog.push({
        rowIndex: uretimTakipRows.length + productionLog.length + 5, // Arbitrary index offset for new entries
        kod: code,
        adet: qty,
        fazla: 0
    });

    showToast(`"${code}" kodu için ${qty} adet üretim girildi.`, "success");
    document.getElementById('prod-qty').value = '';
    
    // Recalculate
    recalculateAll();
    
    // Re-render
    renderProductionTab();
    showQuickPartInfo(code);
});

// Delete production log entry
function deleteLogEntry(index) {
    const deletedCode = productionLog[index].kod;
    productionLog.splice(index, 1);
    showToast(`Üretim kaydı silindi.`, "info");
    
    recalculateAll();
    renderProductionTab();
    showQuickPartInfo(deletedCode);
}

function renderProductionLog() {
    const tbody = document.getElementById('production-log-body');
    tbody.innerHTML = '';
    
    document.getElementById('prod-log-count').textContent = `${productionLog.length} Kayıt`;

    if (productionLog.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="text-center" style="color: var(--text-dim); padding:20px;">Giriş yapılmadı.</td></tr>';
        return;
    }

    // Show logs in reverse order (newest first)
    [...productionLog].reverse().forEach((log, index) => {
        // Calculate original index in productionLog array
        const origIndex = productionLog.length - 1 - index;
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${origIndex + 1}</td>
            <td style="font-weight:700; color:var(--primary); cursor:pointer;" onclick="showQuickPartInfo('${log.kod}')">${log.kod}</td>
            <td class="text-right" style="font-weight:600;">${log.adet}</td>
            <td class="text-right" style="color:${log.fazla > 0 ? 'var(--warning)' : 'var(--text-dim)'};">${log.fazla > 0 ? '+' + log.fazla : '-'}</td>
            <td>
                <button class="trash-btn" onclick="deleteLogEntry(${origIndex})" title="Kayıt Sil">
                    <i class="fa-solid fa-trash-can"></i>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function filterAndPaginateTakipTable() {
    const searchVal = document.getElementById('takip-search').value.toLowerCase().trim();
    const statusVal = document.getElementById('takip-filter-status').value;
    
    let filtered = uretimTakipRows;
    
    // Filter search
    if (searchVal) {
        filtered = filtered.filter(r => r.kod.toLowerCase().includes(searchVal) || r.kaynak.toLowerCase().includes(searchVal));
    }

    // Filter status
    if (statusVal === 'complete') {
        filtered = filtered.filter(r => r.kalan === 0);
    } else if (statusVal === 'partial') {
        filtered = filtered.filter(r => r.uretilen > 0 && r.kalan > 0);
    } else if (statusVal === 'missing') {
        filtered = filtered.filter(r => r.uretilen === 0);
    }

    paginationState.takip.filtered = filtered;
    paginationState.takip.total = filtered.length;
    
    // Clamp page
    const maxPage = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
    if (paginationState.takip.page > maxPage) paginationState.takip.page = maxPage;

    renderTakipTable();
}

function renderTakipTable() {
    const tbody = document.getElementById('takip-table-body');
    tbody.innerHTML = '';

    const pState = paginationState.takip;
    const startIdx = (pState.page - 1) * PAGE_SIZE;
    const endIdx = Math.min(pState.total, startIdx + PAGE_SIZE);

    document.getElementById('takip-page-info').textContent = `Gösterilen: ${pState.total > 0 ? startIdx + 1 : 0} - ${endIdx} / Toplam: ${pState.total}`;

    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="color: var(--text-dim); padding:20px;">Eşleşen parça bulunamadı.</td></tr>';
        return;
    }

    const pageRows = pState.filtered.slice(startIdx, endIdx);
    
    pageRows.forEach(r => {
        const pct = Math.round(r.tamamlanma * 100);
        let barColor = 'var(--danger)';
        let badge = '<span class="badge badge-danger">Eksik</span>';
        if (pct >= 100) {
            barColor = 'var(--success)';
            badge = '<span class="badge badge-success">Tamamlandı</span>';
        } else if (pct > 0) {
            barColor = 'var(--warning)';
            badge = `<span class="badge badge-warning">Kısmi (${pct}%)</span>`;
        }

        const tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        tr.addEventListener('click', () => {
            showQuickPartInfo(r.kod);
            document.getElementById('prod-code').value = r.kod;
        });
        
        tr.innerHTML = `
            <td style="font-size:12px; color:var(--text-muted);">${r.kaynak}</td>
            <td>${r.oncelik}</td>
            <td style="font-weight:700; color:white;">${r.kod}</td>
            <td class="text-right" style="font-weight:600;">${r.uretilecek}</td>
            <td class="text-right" style="color:${r.uretilen > 0 ? 'var(--success)' : 'var(--text-dim)'};">${r.uretilen}</td>
            <td class="text-right" style="color:${r.kalan > 0 ? 'var(--danger)' : 'var(--text-dim)'};">${r.kalan}</td>
            <td>
                <div class="progress-bar-small">
                    <div class="progress-bar-small-fill" style="width: ${pct}%; background-color: ${barColor};"></div>
                </div>
                <span class="table-pct-text" style="color: ${barColor};">${pct}%</span>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// Pagination controls for Takip table
document.getElementById('takip-prev-btn').addEventListener('click', () => {
    if (paginationState.takip.page > 1) {
        paginationState.takip.page--;
        renderTakipTable();
    }
});
document.getElementById('takip-next-btn').addEventListener('click', () => {
    const maxPage = Math.ceil(paginationState.takip.total / PAGE_SIZE);
    if (paginationState.takip.page < maxPage) {
        paginationState.takip.page++;
        renderTakipTable();
    }
});

document.getElementById('takip-search').addEventListener('input', () => {
    paginationState.takip.page = 1;
    filterAndPaginateTakipTable();
});
document.getElementById('takip-filter-status').addEventListener('change', () => {
    paginationState.takip.page = 1;
    filterAndPaginateTakipTable();
});

// -------------------------------------------------------------
// 3. ASSEMBLY VIEW (MONTAJ İZLEME)
// -------------------------------------------------------------
// Assembly Subtabs Click
document.querySelectorAll('.assembly-tab-btn').forEach(btn => {
    btn.addEventListener('click', function() {
        document.querySelectorAll('.assembly-tab-btn').forEach(b => b.classList.remove('active'));
        this.classList.add('active');
        activeTabAssembly = this.getAttribute('data-assembly');
        
        // Reset pages and expanded accordion
        paginationState.assemblyLeft.page = 1;
        paginationState.assemblyRight.page = 1;
        expandedParentAssembly = null;
        
        filterAndPaginateAssembly();
    });
});

function renderAssemblyTab() {
    filterAndPaginateAssembly();
}

function filterAndPaginateAssembly() {
    const rightSearch = document.getElementById('assembly-right-search').value.toLowerCase().trim();
    const rightFilterVal = document.getElementById('assembly-right-filter-status').value;
    const rightSortVal = document.getElementById('assembly-right-sort').value;

    let rightSource = activeTabAssembly === 'otomasyon' ? montajOtomasyonRight : finalMontajRight;

    // Helper to get priority number
    const getPriorityNum = (kStr) => {
        const match = kStr.split('-')[0].trim();
        return parseInt(match) || 99999;
    };

    // Filter Right
    let rightFiltered = rightSource;
    if (rightSearch) {
        rightFiltered = rightSource.filter(r => 
            r.ustKod.toLowerCase().includes(rightSearch) || 
            r.ustAd.toLowerCase().includes(rightSearch) ||
            r.kaynak.toLowerCase().includes(rightSearch)
        );
    }

    // Filter Right by Status
    if (rightFilterVal === 'missing') {
        rightFiltered = rightFiltered.filter(r => r.tamamlanma < 1.0);
    } else if (rightFilterVal === 'complete') {
        rightFiltered = rightFiltered.filter(r => r.tamamlanma >= 1.0);
    }

    // Sort Right
    if (rightSortVal === 'priority') {
        rightFiltered = [...rightFiltered].sort((a, b) => {
            const priA = getPriorityNum(a.kaynak);
            const priB = getPriorityNum(b.kaynak);
            if (priA !== priB) return priA - priB;
            return a.ustKod.localeCompare(b.ustKod);
        });
    } else if (rightSortVal === 'completion-desc') {
        rightFiltered = [...rightFiltered].sort((a, b) => {
            if (b.tamamlanma !== a.tamamlanma) return b.tamamlanma - a.tamamlanma;
            return getPriorityNum(a.kaynak) - getPriorityNum(b.kaynak); // Fallback to priority
        });
    } else if (rightSortVal === 'limit-desc') {
        rightFiltered = [...rightFiltered].sort((a, b) => {
            if (b.limit !== a.limit) return b.limit - a.limit;
            return getPriorityNum(a.kaynak) - getPriorityNum(b.kaynak); // Fallback to priority
        });
    }

    paginationState.assemblyRight.filtered = rightFiltered;
    paginationState.assemblyRight.total = rightFiltered.length;

    renderAssemblyRightTable();
}

// Right Table (Parent Assembly Limits)
function renderAssemblyRightTable() {
    const tbody = document.getElementById('assembly-right-body');
    tbody.innerHTML = '';

    const pState = paginationState.assemblyRight;
    const startIdx = (pState.page - 1) * PAGE_SIZE;
    const endIdx = Math.min(pState.total, startIdx + PAGE_SIZE);

    document.getElementById('assembly-right-page-info').textContent = `Gösterilen: ${pState.total > 0 ? startIdx + 1 : 0} - ${endIdx} / Toplam: ${pState.total}`;

    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="color: var(--text-dim); padding:20px;">Eşleşen üst montaj bulunamadı.</td></tr>';
        return;
    }

    const pageRows = pState.filtered.slice(startIdx, endIdx);
    const activeTabLeftSource = activeTabAssembly === 'otomasyon' ? montajOtomasyonLeft : finalMontajLeft;

    pageRows.forEach(parent => {
        const pct = Math.round(parent.tamamlanma * 100);
        let colorClass = 'var(--danger)';
        if (pct >= 100) colorClass = 'var(--success)';
        else if (pct > 0) colorClass = 'var(--warning)';

        const hasShortage = parent.tamamlanma < 1.0;
        const parentKey = `${parent.ustKod}___${parent.kaynak}`;
        const isExpanded = hasShortage && expandedParentAssembly === parentKey;

        const tr = document.createElement('tr');
        if (hasShortage) {
            tr.style.cursor = 'pointer';
            if (isExpanded) {
                tr.classList.add('expanded-parent-row');
            }
            tr.addEventListener('click', () => {
                if (expandedParentAssembly === parentKey) {
                    expandedParentAssembly = null;
                } else {
                    expandedParentAssembly = parentKey;
                }
                filterAndPaginateAssembly();
            });
        } else {
            tr.style.cursor = 'default';
        }

        const chevron = hasShortage
            ? (isExpanded 
                ? '<i class="fa-solid fa-chevron-down text-purple" style="margin-right: 8px; font-size:10px;"></i>' 
                : '<i class="fa-solid fa-chevron-right" style="margin-right: 8px; font-size:10px; color: var(--text-dim);"></i>')
            : '';

        tr.innerHTML = `
            <td style="font-size:11px; color:var(--text-muted);">${parent.kaynak}</td>
            <td style="font-weight:700; color:white;">
                <span style="display:inline-flex; align-items:center;">
                    ${chevron}${parent.ustKod}
                </span>
            </td>
            <td style="color:var(--text-muted); max-width:180px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${parent.ustAd}">${parent.ustAd}</td>
            <td class="text-right" style="font-weight:500;">${parent.tamamlananCesit} / ${parent.gerekenCesit}</td>
            <td style="color:${colorClass}; font-weight:800;">${pct}%</td>
            <td class="text-right" style="color:var(--text-muted);">${parent.toplamAdet}</td>
            <td class="text-right highlight-column" style="color:${parent.limit > 0 ? 'var(--success)' : 'var(--danger)'}; font-weight:800; font-size:14px; text-shadow:${parent.limit > 0 ? '0 0 10px var(--success-glow)' : 'none'}">${parent.limit} Set</td>
        `;
        tbody.appendChild(tr);

        if (isExpanded) {
            // Find child rows under this parent in this source file
            const kids = activeTabLeftSource.filter(l => l.kaynak === parent.kaynak && l.ustKod === parent.ustKod);

            const subTr = document.createElement('tr');
            subTr.className = 'sub-row-expanded';

            let kidsHtml = `
                <td colspan="7" style="padding: 12px 20px; background: rgba(0, 0, 0, 0.25); border-left: 4px solid var(--purple);">
                    <div style="padding: 12px; border-radius: 8px; background: rgba(255, 255, 255, 0.02); border: 1px solid rgba(255, 255, 255, 0.05); box-shadow: inset 0 0 10px rgba(0,0,0,0.5);">
                        <h4 style="margin: 0 0 12px 0; font-size: 13px; color: #a78bfa; font-weight: 700; display: flex; justify-content: space-between; align-items: center;">
                            <span><i class="fa-solid fa-sitemap" style="margin-right: 6px;"></i> Reçete Alt Kırılım Detayları: ${parent.ustKod} (${parent.ustAd})</span>
                            <span style="font-size: 11px; font-weight: normal; color: var(--text-muted);"><i class="fa-solid fa-file-excel" style="margin-right: 4px;"></i> Kaynak: ${parent.kaynak}</span>
                        </h4>
                        <table class="table inner-table" style="width: 100%; margin: 0; font-size: 12px; border-collapse: collapse;">
                            <thead>
                                <tr style="border-bottom: 1px solid rgba(255,255,255,0.08);">
                                    <th style="padding: 8px; text-align: left; color: var(--text-muted); font-size: 11px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Alt Parça Kodu</th>
                                    <th style="padding: 8px; text-align: left; color: var(--text-muted); font-size: 11px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Alt Parça Adı</th>
                                    <th style="padding: 8px; text-align: left; color: var(--text-muted); font-size: 11px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">İstasyonlar</th>
                                    <th style="padding: 8px; text-align: right; color: var(--text-muted); font-size: 11px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Gereken</th>
                                    <th style="padding: 8px; text-align: right; color: var(--text-muted); font-size: 11px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Üretilen</th>
                                    <th style="padding: 8px; text-align: center; color: var(--text-muted); font-size: 11px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Tamamlanma %</th>
                                    <th style="padding: 8px; text-align: right; color: var(--text-muted); font-size: 11px; text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Stok Limiti</th>
                                </tr>
                            </thead>
                            <tbody>
            `;

            if (kids.length === 0) {
                kidsHtml += `
                    <tr>
                        <td colspan="7" style="padding: 12px; text-align: center; color: var(--text-dim); font-style: italic;">Bu montaj için alt parça bulunamadı.</td>
                    </tr>
                `;
            } else {
                kids.forEach(k => {
                    const kPct = Math.round(k.tamamlanma * 100);
                    let kColorClass = 'var(--danger)';
                    if (kPct >= 100) kColorClass = 'var(--success)';
                    else if (kPct > 0) kColorClass = 'var(--warning)';

                    // Find matching stations for this child code
                    const kCode = String(k.altKod || '').trim().toUpperCase();
                    const stations = [];
                    for (const [stName, rows] of Object.entries(stationSheetsMap)) {
                        const found = rows.some(r => String(r['Kod'] || '').trim().toUpperCase() === kCode);
                        if (found) {
                            stations.push(stName);
                        }
                    }

                    let stationsHtml = '';
                    if (stations.length === 0) {
                        stationsHtml = '<span style="color: var(--text-dim); font-size: 11px;">-</span>';
                    } else {
                        stationsHtml = stations.map(st => `
                            <span class="badge" style="background: rgba(139, 92, 246, 0.15); color: #c084fc; border: 1px solid rgba(139, 92, 246, 0.25); margin: 2px; font-size: 10px; padding: 2px 6px; font-weight: 600;">${st}</span>
                        `).join('');
                    }

                    kidsHtml += `
                        <tr style="border-bottom: 1px solid rgba(255,255,255,0.04); transition: background-color 0.2s;">
                            <td style="padding: 8px; font-weight: 700; color: white;">${k.altKod}</td>
                            <td style="padding: 8px; color: var(--text-muted); max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${k.altAd}">${k.altAd}</td>
                            <td style="padding: 8px; text-align: left;">${stationsHtml}</td>
                            <td style="padding: 8px; text-align: right; font-weight: 600; color: var(--text-muted);">${k.gereken}</td>
                            <td style="padding: 8px; text-align: right; color: ${k.uretilen > 0 ? 'var(--success)' : 'var(--text-dim)'}; font-weight: 600;">${k.uretilen}</td>
                            <td style="padding: 8px; text-align: center; color: ${kColorClass}; font-weight: 700;">${kPct}%</td>
                            <td style="padding: 8px; text-align: right; color: ${k.limit > 0 ? 'var(--purple)' : 'var(--text-dim)'}; font-weight: 700;">${k.limit}</td>
                        </tr>
                    `;
                });
            }

            kidsHtml += `
                            </tbody>
                        </table>
                    </div>
                </td>
            `;
            subTr.innerHTML = kidsHtml;
            tbody.appendChild(subTr);
        }
    });
}

// Assembly search input events
document.getElementById('assembly-right-search').addEventListener('input', () => {
    paginationState.assemblyRight.page = 1;
    filterAndPaginateAssembly();
});
document.getElementById('assembly-right-filter-status').addEventListener('change', () => {
    paginationState.assemblyRight.page = 1;
    filterAndPaginateAssembly();
});
document.getElementById('assembly-right-sort').addEventListener('change', () => {
    paginationState.assemblyRight.page = 1;
    filterAndPaginateAssembly();
});

// Right pagination buttons
document.getElementById('assembly-right-prev-btn').addEventListener('click', () => {
    if (paginationState.assemblyRight.page > 1) {
        paginationState.assemblyRight.page--;
        renderAssemblyRightTable();
    }
});
document.getElementById('assembly-right-next-btn').addEventListener('click', () => {
    const maxPage = Math.ceil(paginationState.assemblyRight.total / PAGE_SIZE);
    if (paginationState.assemblyRight.page < maxPage) {
        paginationState.assemblyRight.page++;
        renderAssemblyRightTable();
    }
});

// -------------------------------------------------------------
// 4. ROTASIZLAR VIEW
// -------------------------------------------------------------
function renderRotasizlarTab() {
    const tbody = document.getElementById('rotasiz-tbody');
    const searchVal = document.getElementById('rotasiz-search').value.toLowerCase().trim();
    
    tbody.innerHTML = '';
    
    let filtered = rotasizRows;
    if (searchVal) {
        filtered = rotasizRows.filter(r => {
            const code = String(r['Kod'] || '').toLowerCase();
            const mat = String(r['Malzeme'] || '').toLowerCase();
            const source = String(r['KAYNAK DOSYA'] || '').toLowerCase();
            return code.includes(searchVal) || mat.includes(searchVal) || source.includes(searchVal);
        });
    }

    document.getElementById('rotasiz-count').textContent = `${filtered.length} Kalem`;

    if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" class="text-center" style="color:var(--text-dim); padding:20px;">Mevcut kriterlere uygun rotasız parça bulunmamaktadır.</td></tr>`;
        return;
    }

    filtered.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row['KAYNAK DOSYA'] || '-'}</td>
            <td>${row['Sıra No'] || '-'}</td>
            <td>${row['Öncelik Sırası'] || '-'}</td>
            <td style="font-weight:700; color:white;">${row['Kod'] || '-'}</td>
            <td style="color:var(--text-muted);">${row['Malzeme'] || '-'}</td>
        `;
        tbody.appendChild(tr);
    });
}

document.getElementById('rotasiz-search').addEventListener('input', renderRotasizlarTab);

// -------------------------------------------------------------
// 5. STATIONS VIEW
// -------------------------------------------------------------
function renderStationsTab() {
    // 1. Render stations sidebar list
    const sidebar = document.getElementById('stations-list');
    sidebar.innerHTML = '';

    let draggedItem = null;

    stationList.forEach(st => {
        const btn = document.createElement('button');
        btn.className = `station-item-btn ${st === activeStation ? 'active' : ''}`;
        btn.setAttribute('draggable', true);
        
        // Count how many parts in this station are fully complete
        const rows = stationSheetsMap[st] || [];
        let completedCount = 0;
        rows.forEach(r => {
            const code = String(r['Kod'] || '').trim().toUpperCase();
            // Check in Üretim Takip
            const reqs = uretimTakipRows.filter(u => u.kod === code);
            if (reqs.length > 0) {
                const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
                const totalProd = reqs.reduce((sum, u) => sum + u.uretilen, 0.0);
                if (totalReq > 0 && totalProd >= totalReq) {
                    completedCount++;
                }
            }
        });

        btn.innerHTML = `
            <span class="station-name" title="${st}">${st}</span>
            <span class="badge-pill">${completedCount} / ${rows.length}</span>
        `;
        
        btn.addEventListener('click', () => {
            activeStation = st;
            paginationState.station.page = 1;
            renderStationsTab();
        });

        // HTML5 Drag and Drop events
        btn.addEventListener('dragstart', (e) => {
            draggedItem = st;
            btn.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        });
        
        btn.addEventListener('dragend', () => {
            btn.classList.remove('dragging');
            draggedItem = null;
            // Save new order to localStorage
            localStorage.setItem('customStationOrder', JSON.stringify(stationList));
        });
        
        btn.addEventListener('dragover', (e) => {
            e.preventDefault();
        });
        
        btn.addEventListener('dragenter', (e) => {
            e.preventDefault();
            btn.classList.add('drag-over');
        });
        
        btn.addEventListener('dragleave', () => {
            btn.classList.remove('drag-over');
        });
        
        btn.addEventListener('drop', (e) => {
            e.preventDefault();
            btn.classList.remove('drag-over');
            if (draggedItem && draggedItem !== st) {
                const draggedIdx = stationList.indexOf(draggedItem);
                const targetIdx = stationList.indexOf(st);
                
                // Reorder stationList
                stationList.splice(draggedIdx, 1);
                stationList.splice(targetIdx, 0, draggedItem);
                
                renderStationsTab();
            }
        });

        sidebar.appendChild(btn);
    });

    // 2. Filter & Render Station Rows
    filterAndPaginateStationData();
}

function filterAndPaginateStationData() {
    const searchVal = document.getElementById('station-search').value.toLowerCase().trim();
    const statusFilter = document.getElementById('station-filter-status') ? document.getElementById('station-filter-status').value : 'all';
    const sortVal = document.getElementById('station-sort') ? document.getElementById('station-sort').value : 'priority-asc';
    
    const rows = stationSheetsMap[activeStation] || [];
    const headers = stationHeadersMap[activeStation] || [];

    document.getElementById('current-station-title').innerHTML = `<i class="fa-solid fa-industry text-green"></i> ${activeStation} İstasyon İş Listesi`;

    // 1. Text Search Filter
    let filtered = rows;
    if (searchVal) {
        filtered = rows.filter(r => {
            const code = String(r['Kod'] || '').toLowerCase();
            const mat = String(r['Malzeme Adı'] || '').toLowerCase();
            const hCode = String(r['Hammadde Kod'] || '').toLowerCase();
            return code.includes(searchVal) || mat.includes(searchVal) || hCode.includes(searchVal);
        });
    }

    // 2. Status Filter
    if (statusFilter !== 'all') {
        filtered = filtered.filter(row => {
            const code = String(row['Kod'] || '').trim().toUpperCase();
            const reqs = uretimTakipRows.filter(u => u.kod === code);
            let completionPct = 0;
            if (reqs.length > 0) {
                const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
                const totalProd = reqs.reduce((sum, u) => sum + u.uretilen, 0.0);
                completionPct = totalReq > 0 ? Math.round((totalProd / totalReq) * 100) : 0;
            }
            if (statusFilter === 'ready') return completionPct >= 100;
            if (statusFilter === 'production') return completionPct > 0 && completionPct < 100;
            if (statusFilter === 'missing') return completionPct === 0;
            return true;
        });
    }

    // 3. Sorting
    filtered.sort((a, b) => {
        if (sortVal === 'priority-asc' || sortVal === 'priority-desc') {
            const priorityA = parseInt(a['Öncelik Sırası']) || 99999;
            const priorityB = parseInt(b['Öncelik Sırası']) || 99999;
            return sortVal === 'priority-asc' ? priorityA - priorityB : priorityB - priorityA;
        }
        if (sortVal === 'code-asc') {
            const codeA = String(a['Kod'] || '');
            const codeB = String(b['Kod'] || '');
            return codeA.localeCompare(codeB, undefined, { numeric: true, sensitivity: 'base' });
        }
        if (sortVal === 'uretilecek-desc') {
            const qtyA = parseFloat(a['Üretilecek Miktar']) || 0;
            const qtyB = parseFloat(b['Üretilecek Miktar']) || 0;
            return qtyB - qtyA;
        }
        return 0;
    });

    paginationState.station.filtered = filtered;
    paginationState.station.total = filtered.length;

    // Clamp page
    const maxPage = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
    if (paginationState.station.page > maxPage) paginationState.station.page = maxPage;

    renderStationTable(headers);
}

function renderStationTable(headers) {
    const thead = document.getElementById('station-thead');
    const tbody = document.getElementById('station-tbody');

    thead.innerHTML = '';
    tbody.innerHTML = '';

    const pState = paginationState.station;
    const startIdx = (pState.page - 1) * PAGE_SIZE;
    const endIdx = Math.min(pState.total, startIdx + PAGE_SIZE);

    document.getElementById('station-page-info').textContent = `Gösterilen: ${pState.total > 0 ? startIdx + 1 : 0} - ${endIdx} / Toplam: ${pState.total}`;

    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="text-center" style="color:var(--text-dim); padding:20px;">Bu istasyonda eşleşen iş listesi bulunamadı.</td></tr>';
        return;
    }

    // Select the key columns to display
    const displayCols = headers.filter(h => [
        'Öncelik Sırası', 'Kod', 'Malzeme Adı', 'Hammadde Kod', 'Hammadde', 
        'Rezerve Edilecek Miktar', 'Üretilecek Miktar', 'Toplam Hammadde Miktarı', 'Durum'
    ].includes(h) || h === 'Kod' || h.includes('Miktar') || h.includes('Öncelik'));

    const colsToShow = displayCols.length > 0 ? displayCols : headers.slice(0, 8);

    // Add Status header
    const finalHeaders = [...colsToShow];
    if (!finalHeaders.includes('Durum')) finalHeaders.push('Durum');

    // Create table header cells
    const trHead = document.createElement('tr');
    finalHeaders.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        if (h.includes('Miktar') || h.includes('Adet')) th.className = 'text-right';
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    const pageRows = pState.filtered.slice(startIdx, endIdx);

    pageRows.forEach(row => {
        const code = String(row['Kod'] || '').trim().toUpperCase();
        
        // Calculate status from Üretim Takip
        const reqs = uretimTakipRows.filter(u => u.kod === code);
        let completionText = '-';
        let completionPct = 0;
        let badgeHtml = '<span class="badge badge-danger">Eksik</span>';
        
        if (reqs.length > 0) {
            const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
            const totalProd = reqs.reduce((sum, u) => sum + u.uretilen, 0.0);
            completionPct = totalReq > 0 ? Math.round((totalProd / totalReq) * 100) : 0;
            completionText = `${totalProd} / ${totalReq}`;
            
            if (completionPct >= 100) {
                badgeHtml = `<span class="badge badge-success">Hazır (${completionPct}%)</span>`;
            } else if (completionPct > 0) {
                badgeHtml = `<span class="badge badge-warning">Üretimde (${completionPct}%)</span>`;
            }
        }

        const tr = document.createElement('tr');
        finalHeaders.forEach(h => {
            const td = document.createElement('td');
            
            if (h === 'Durum') {
                td.innerHTML = badgeHtml;
            } else {
                let val = row[h];
                if (typeof val === 'number') {
                    if (!Number.isInteger(val)) {
                        val = parseFloat(val.toFixed(3));
                    }
                }
                td.textContent = (val !== undefined && val !== null) ? val : '-';
                
                // Styling specific columns
                if (h === 'Kod') {
                    td.style.fontWeight = '700';
                    td.style.color = 'white';
                } else if (h === 'Malzeme Adı' || h === 'Hammadde') {
                    td.style.color = 'var(--text-muted)';
                } else if (h.includes('Miktar') || h.includes('Adet')) {
                    td.className = 'text-right';
                    td.style.fontWeight = '600';
                }
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

// Station navigation events
document.getElementById('station-search').addEventListener('input', () => {
    paginationState.station.page = 1;
    filterAndPaginateStationData();
});



if (document.getElementById('station-filter-status')) {
    document.getElementById('station-filter-status').addEventListener('change', () => {
        paginationState.station.page = 1;
        filterAndPaginateStationData();
    });
}

if (document.getElementById('station-sort')) {
    document.getElementById('station-sort').addEventListener('change', () => {
        paginationState.station.page = 1;
        filterAndPaginateStationData();
    });
}

document.getElementById('station-prev-btn').addEventListener('click', () => {
    if (paginationState.station.page > 1) {
        paginationState.station.page--;
        filterAndPaginateStationData();
    }
});
document.getElementById('station-next-btn').addEventListener('click', () => {
    const maxPage = Math.ceil(paginationState.station.total / PAGE_SIZE);
    if (paginationState.station.page < maxPage) {
        paginationState.station.page++;
        filterAndPaginateStationData();
    }
});

// -------------------------------------------------------------
// 6. EXPORT BACK TO EXCEL WORKBOOK
// -------------------------------------------------------------
exportBtn.addEventListener('click', () => {
    if (!workbook) return;
    
    showToast("Güncel veriler Excel dosyasına yazılıyor...", "info");
    
    try {
        // 1. Write the new production log entries to columns I and J in "Üretim Takip" sheet
        const utSheet = workbook.Sheets["Üretim Takip"];
        
        // Find existing range
        const utRange = XLSX.utils.decode_range(utSheet['!ref']);
        
        // Clear all log columns from row 2 onwards to rewrite completely
        for (let r = 1; r <= utRange.e.r; r++) {
            delete utSheet[XLSX.utils.encode_cell({ r: r, c: 8 })]; // Col I
            delete utSheet[XLSX.utils.encode_cell({ r: r, c: 9 })]; // Col J
            delete utSheet[XLSX.utils.encode_cell({ r: r, c: 10 })]; // Col K
        }

        // Rewrite production log rows
        productionLog.forEach((log, idx) => {
            const r = idx + 1; // row 2 onwards
            utSheet[XLSX.utils.encode_cell({ r: r, c: 8 })] = { t: 's', v: log.kod };
            utSheet[XLSX.utils.encode_cell({ r: r, c: 9 })] = { t: 'n', v: log.adet };
            
            // Formula for Col K: =IF(I<>, MAX(0, SUMIF($I$2:$I$6377, I, $J$2:$J$6377) - SUMIF($C$2:$C$6377, I, $D$2:$D$6377)), "")
            utSheet[XLSX.utils.encode_cell({ r: r, c: 10 })] = { 
                t: 'n', 
                v: log.fazla,
                f: `IF(I${r+1}<>"",MAX(0,SUMIF($I$2:$I$6377,I${r+1},$J$2:$J$6377)-SUMIF($C$2:$C$6377,I${r+1},$D$2:$D$6377)),"")`
            };
        });

        // Ensure SheetJS knows the new range of "Üretim Takip"
        const maxLogLength = Math.max(uretimTakipRows.length, productionLog.length) + 10;
        if (maxLogLength - 1 > utRange.e.r) {
            utRange.e.r = maxLogLength - 1;
            utSheet['!ref'] = XLSX.utils.encode_range(utRange);
        }

        // 2. Update all cell values in Üretim Takip formula columns (Col E, F, G, O, P, Q)
        uretimTakipRows.forEach(row => {
            const r = row.rowIndex - 1; // 0-based
            // Col E: Üretilen Miktar
            const cellE = utSheet[XLSX.utils.encode_cell({ r: r, c: 4 })] || { t: 'n' };
            cellE.v = row.uretilen;
            utSheet[XLSX.utils.encode_cell({ r: r, c: 4 })] = cellE;

            // Col F: Kalan Miktar
            const cellF = utSheet[XLSX.utils.encode_cell({ r: r, c: 5 })] || { t: 'n' };
            cellF.v = row.kalan;
            utSheet[XLSX.utils.encode_cell({ r: r, c: 5 })] = cellF;

            // Col G: Tamamlanma (%)
            const cellG = utSheet[XLSX.utils.encode_cell({ r: r, c: 6 })] || { t: 'n' };
            cellG.v = row.tamamlanma;
            utSheet[XLSX.utils.encode_cell({ r: r, c: 6 })] = cellG;
        });

        dosyaTakipRows.forEach(row => {
            const r = row.rowIndex - 1; // 0-based
            // Col O: HAZIR (Kalem)
            const cellO = utSheet[XLSX.utils.encode_cell({ r: r, c: 14 })] || { t: 'n' };
            cellO.v = row.hazir;
            utSheet[XLSX.utils.encode_cell({ r: r, c: 14 })] = cellO;

            // Col P: EKSİK (Kalem)
            const cellP = utSheet[XLSX.utils.encode_cell({ r: r, c: 15 })] || { t: 'n' };
            cellP.v = row.eksik;
            utSheet[XLSX.utils.encode_cell({ r: r, c: 15 })] = cellP;

            // Col Q: GENEL TAMAMLANMA (%)
            const cellQ = utSheet[XLSX.utils.encode_cell({ r: r, c: 16 })] || { t: 'n' };
            cellQ.v = row.tamamlanma;
            utSheet[XLSX.utils.encode_cell({ r: r, c: 16 })] = cellQ;
        });

        // 3. Update MONTAJ OTOMASYON İZLEME cell values
        const moSheet = workbook.Sheets["MONTAJ OTOMASYON İZLEME"];
        if (moSheet) {
            montajOtomasyonLeft.forEach(row => {
                const r = row.rowIndex - 1;
                // Col G: Üretilen Miktar
                const cellG = moSheet[XLSX.utils.encode_cell({ r: r, c: 6 })] || { t: 'n' };
                cellG.v = row.uretilen;
                moSheet[XLSX.utils.encode_cell({ r: r, c: 6 })] = cellG;

                // Col H: Tamamlanma Oranı (%)
                const cellH = moSheet[XLSX.utils.encode_cell({ r: r, c: 7 })] || { t: 'n' };
                cellH.v = row.tamamlanma;
                moSheet[XLSX.utils.encode_cell({ r: r, c: 7 })] = cellH;

                // Col I: Alt Parça Limit
                const cellI = moSheet[XLSX.utils.encode_cell({ r: r, c: 8 })] || { t: 'n' };
                cellI.v = row.limit;
                moSheet[XLSX.utils.encode_cell({ r: r, c: 8 })] = cellI;
            });

            montajOtomasyonRight.forEach(row => {
                const r = row.rowIndex - 1;
                // Col O: Tamamlanan Çeşit
                const cellO = moSheet[XLSX.utils.encode_cell({ r: r, c: 14 })] || { t: 'n' };
                cellO.v = row.tamamlananCesit;
                moSheet[XLSX.utils.encode_cell({ r: r, c: 14 })] = cellO;

                // Col P: Tamamlanma Oranı (%)
                const cellP = moSheet[XLSX.utils.encode_cell({ r: r, c: 15 })] || { t: 'n' };
                cellP.v = row.tamamlanma;
                moSheet[XLSX.utils.encode_cell({ r: r, c: 15 })] = cellP;

                // Col R: Ek Toplanabilir (Set)
                const cellR = moSheet[XLSX.utils.encode_cell({ r: r, c: 17 })] || { t: 'n' };
                cellR.v = row.limit;
                moSheet[XLSX.utils.encode_cell({ r: r, c: 17 })] = cellR;
            });
        }

        // 4. Update FINAL MONTAJ İZLEME cell values
        const fmSheet = workbook.Sheets["FINAL MONTAJ İZLEME"];
        if (fmSheet) {
            finalMontajLeft.forEach(row => {
                const r = row.rowIndex - 1;
                // Col G: Üretilen Miktar
                const cellG = fmSheet[XLSX.utils.encode_cell({ r: r, c: 6 })] || { t: 'n' };
                cellG.v = row.uretilen;
                fmSheet[XLSX.utils.encode_cell({ r: r, c: 6 })] = cellG;

                // Col H: Tamamlanma Oranı (%)
                const cellH = fmSheet[XLSX.utils.encode_cell({ r: r, c: 7 })] || { t: 'n' };
                cellH.v = row.tamamlanma;
                fmSheet[XLSX.utils.encode_cell({ r: r, c: 7 })] = cellH;

                // Col I: Alt Parça Limit
                const cellI = fmSheet[XLSX.utils.encode_cell({ r: r, c: 8 })] || { t: 'n' };
                cellI.v = row.limit;
                fmSheet[XLSX.utils.encode_cell({ r: r, c: 8 })] = cellI;
            });

            finalMontajRight.forEach(row => {
                const r = row.rowIndex - 1;
                // Col O: Tamamlanan Çeşit
                const cellO = fmSheet[XLSX.utils.encode_cell({ r: r, c: 14 })] || { t: 'n' };
                cellO.v = row.tamamlananCesit;
                fmSheet[XLSX.utils.encode_cell({ r: r, c: 14 })] = cellO;

                // Col P: Tamamlanma Oranı (%)
                const cellP = fmSheet[XLSX.utils.encode_cell({ r: r, c: 15 })] || { t: 'n' };
                cellP.v = row.tamamlanma;
                fmSheet[XLSX.utils.encode_cell({ r: r, c: 15 })] = cellP;

                // Col R: Ek Toplanabilir (Set)
                const cellR = fmSheet[XLSX.utils.encode_cell({ r: r, c: 17 })] || { t: 'n' };
                cellR.v = row.limit;
                fmSheet[XLSX.utils.encode_cell({ r: r, c: 17 })] = cellR;
            });
        }

        // 5. Update Station Sheets cell values in the workbook
        for (const [stName, rows] of Object.entries(stationSheetsMap)) {
            const sheet = workbook.Sheets[stName];
            if (!sheet) continue;
            
            // Read headers of this sheet to find column indexes of Üretilecek Miktar and Toplam Hammadde Miktarı
            let headers = [];
            const ref = sheet['!ref'];
            if (!ref) continue;
            const range = XLSX.utils.decode_range(ref);
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: c })];
                headers.push(cell ? String(cell.v).trim().toLowerCase() : `sütun ${c+1}`);
            }
            const uMiktarColIdx = headers.indexOf('üretilecek miktar');
            const tHammaddeColIdx = headers.indexOf('toplam hammadde miktarı');
            
            rows.forEach(row => {
                const r = row.rowIndex - 1; // 0-based index for SheetJS
                
                if (uMiktarColIdx !== -1) {
                    const cellG = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + uMiktarColIdx })] || { t: 'n' };
                    cellG.v = row['Üretilecek Miktar'] || 0;
                    sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + uMiktarColIdx })] = cellG;
                }
                
                if (tHammaddeColIdx !== -1) {
                    const cellI = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + tHammaddeColIdx })] || { t: 'n' };
                    cellI.v = row['Toplam Hammadde Miktarı'] || 0;
                    sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + tHammaddeColIdx })] = cellI;
                }
            });
        }

        // 6. Update ÜRETİM LİSTESİ sheet cell values in the workbook
        const ulSheetName = workbook.SheetNames.find(n => 
            n.toUpperCase().replace(/I/g, 'İ').includes('ÜRETİM LİSTESİ') || 
            n.toUpperCase().includes('URETIM LISTESI')
        );
        if (ulSheetName) {
            const ulSheet = workbook.Sheets[ulSheetName];
            if (ulSheet && ulSheet['!ref']) {
                let headers = [];
                const range = XLSX.utils.decode_range(ulSheet['!ref']);
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cell = ulSheet[XLSX.utils.encode_cell({ r: range.s.r, c: c })];
                    headers.push(cell ? String(cell.v).trim().toLowerCase() : `sütun ${c+1}`);
                }
                const miktarColIdx = headers.indexOf('üretilecek miktar');
                if (miktarColIdx !== -1) {
                    uretimListesiRows.forEach(row => {
                        const r = row.rowIndex - 1; // 0-based index
                        const cellG = ulSheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + miktarColIdx })] || { t: 'n' };
                        cellG.v = row.uretilecek;
                        ulSheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + miktarColIdx })] = cellG;
                    });
                }
            }
        }

        // Export workbook to download file
        const origFileName = loadedFileName.textContent.replace("Yüklenen Dosya: ", "").replace(".xlsx", "");
        const exportName = `${origFileName}_Guncel.xlsx`;
        
        XLSX.writeFile(workbook, exportName);
        showToast(`Güncel Excel dosyası indirildi: "${exportName}"`, "success");
        
    } catch (err) {
        console.error(err);
        showToast("Excel dosyasına yazılırken hata oluştu!", "error");
    }
});

// -------------------------------------------------------------
// 7. ÜRETİM LİSTESİ VIEW
// -------------------------------------------------------------
function renderUretimListesiTab() {
    filterAndPaginateUlTable();
}

function filterAndPaginateUlTable() {
    const searchVal = document.getElementById('ul-search').value.toLowerCase().trim();
    const sortVal = document.getElementById('ul-sort').value;

    let filtered = uretimListesiRows;
    
    // 1. Text Search Filter
    if (searchVal) {
        filtered = filtered.filter(r => {
            return r.kod.toLowerCase().includes(searchVal) || 
                   r.malzeme.toLowerCase().includes(searchVal) || 
                   r.hammadde.toLowerCase().includes(searchVal) || 
                   r.kaynak.toLowerCase().includes(searchVal);
        });
    }

    // 2. Sorting
    filtered.sort((a, b) => {
        if (sortVal === 'priority-asc' || sortVal === 'priority-desc') {
            const priorityA = a.oncelik;
            const priorityB = b.oncelik;
            return sortVal === 'priority-asc' ? priorityA - priorityB : priorityB - priorityA;
        }
        if (sortVal === 'code-asc') {
            return a.kod.localeCompare(b.kod, undefined, { numeric: true, sensitivity: 'base' });
        }
        if (sortVal === 'uretilecek-desc') {
            return b.uretilecek - a.uretilecek;
        }
        return 0;
    });

    paginationState.ul.filtered = filtered;
    paginationState.ul.total = filtered.length;

    // Clamp page
    const maxPage = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
    if (paginationState.ul.page > maxPage) paginationState.ul.page = maxPage;

    renderUlTable();
}

function renderUlTable() {
    const tbody = document.getElementById('ul-table-body');
    tbody.innerHTML = '';

    const pState = paginationState.ul;
    const startIdx = (pState.page - 1) * PAGE_SIZE;
    const endIdx = Math.min(pState.total, startIdx + PAGE_SIZE);

    document.getElementById('ul-page-info').textContent = `Gösterilen: ${pState.total > 0 ? startIdx + 1 : 0} - ${endIdx} / Toplam: ${pState.total}`;

    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="color:var(--text-dim); padding:20px;">Eşleşen parça bulunamadı.</td></tr>';
        return;
    }

    const pageRows = pState.filtered.slice(startIdx, endIdx);

    pageRows.forEach(row => {
        const tr = document.createElement('tr');
        
        tr.innerHTML = `
            <td style="font-size:12px; color:var(--text-muted); max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${row.kaynak}">${row.kaynak}</td>
            <td>${row.oncelik}</td>
            <td style="font-weight:700; color:white;">${row.kod}</td>
            <td style="color:var(--text-dim);">${row.malzeme}</td>
            <td style="font-size:12px; color:var(--text-muted);">${row.hKod}</td>
            <td style="color:var(--text-dim);">${row.hammadde}</td>
            <td class="text-right">
                <input type="number" class="table-input ul-qty-input text-right" 
                       value="${row.uretilecek}" 
                       data-kod="${row.kod}" 
                       min="0"
                       style="width: 80px; background: rgba(0,0,0,0.3); border: 1px solid var(--border-color); color: white; border-radius: 4px; padding: 4px 8px; font-weight: 600; outline: none; transition: var(--transition);">
            </td>
        `;
        
        // Add listener to the input
        const input = tr.querySelector('.ul-qty-input');
        input.addEventListener('change', function() {
            const newQty = parseFloat(this.value) || 0;
            const kod = this.dataset.kod;
            
            // 1. Update in uretimListesiRows (all rows matching this code)
            uretimListesiRows.forEach(r => {
                if (r.kod === kod) {
                    r.uretilecek = newQty;
                }
            });
            
            // 2. Update in uretimListesiMap
            uretimListesiMap[kod] = newQty;
            
            // 3. Recalculate
            recalculateAll();
            
            showToast(`${kod} için yeni üretim miktarı belirlendi: ${newQty}`, "success");
        });
        
        input.addEventListener('focus', function() {
            this.style.borderColor = 'var(--primary)';
            this.style.boxShadow = '0 0 8px var(--primary-glow)';
            this.style.background = 'rgba(16, 22, 40, 0.8)';
        });
        input.addEventListener('blur', function() {
            this.style.borderColor = 'var(--border-color)';
            this.style.boxShadow = 'none';
            this.style.background = 'rgba(0,0,0,0.3)';
        });

        tbody.appendChild(tr);
    });
}

// Üretim Listesi navigation events
document.getElementById('ul-search').addEventListener('input', () => {
    paginationState.ul.page = 1;
    filterAndPaginateUlTable();
});

document.getElementById('ul-sort').addEventListener('change', () => {
    paginationState.ul.page = 1;
    filterAndPaginateUlTable();
});

document.getElementById('ul-prev-btn').addEventListener('click', () => {
    if (paginationState.ul.page > 1) {
        paginationState.ul.page--;
        renderUlTable();
    }
});

document.getElementById('ul-next-btn').addEventListener('click', () => {
    const maxPage = Math.ceil(paginationState.ul.total / PAGE_SIZE);
    if (paginationState.ul.page < maxPage) {
        paginationState.ul.page++;
        renderUlTable();
    }
});
