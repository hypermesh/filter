// Global App State
let workbook = null;
let currentTab = 'dashboard';
let loadedExcelFileName = ''; // Yüklü Excel dosya adı (localStorage anahtarı için)
let parcaReceteTuketimMap = {}; // veritabanlari/parca_recete_tuketim.json verisi
let parcaBirimHammaddeMap = {}; // parça bazlı birim hammadde boyu/ölçüsü
let parcaMakineReceteleriMap = {}; // veritabanlari/parca_makine_receteleri.json verisi
let makineYillikTahminlerMap = {}; // veritabanlari/makine_yillik_tahminler.json verisi
let partiFormuluHaricKurallar = {
    haric_hammadde_onekleri: ["150.01.01"],
    haric_hammadde_kodlari: [],
    haric_parca_kodlari: [],
    haric_malzeme_kelimeleri: [],
    sadece_noktali_hammadde_gecerli: true
};

// Parsed Data Structures
let uretimTakipRows = []; // Üretim Takip requirements (Col A-G)
let productionLog = [];    // Üretim Takip logs entered (Col I-K)
let dosyaTakipRows = [];   // Üretim Takip summary (Col M-Q)
let downtimeMap = {};      // İstasyon duruş saatleri { "ISTASYON_ADI": saat }
let hiddenStationCols = new Set(); // Kullanıcının gizlediği sütun başlıkları
let productionHistory = []; // Genel Üretim Geçmişi (veritabanlari/uretim_gecmisi.json karşılığı)
let selectedSourceFiles = new Set(); // Seçili Kaynak Dosyalar Filtresi (Boş ise hepsi)
let rawExcelArrayBuffer = null; // Orijinal Excel dosyasının ham ArrayBuffer verisi (ExcelJS ile stilleri korumak için)

// --- LocalStorage Yardımcı Fonksiyonlar ---
function getStorageKey() {
    return `prodLog_${loadedExcelFileName}`;
}
function saveProductionLogToStorage() {
    try {
        localStorage.setItem(getStorageKey(), JSON.stringify(productionLog));
    } catch(e) {
        console.warn('localStorage kayıt hatası:', e);
    }
}
function loadProductionLogFromStorage() {
    try {
        const raw = localStorage.getItem(getStorageKey());
        if (raw !== null) {
            const saved = JSON.parse(raw);
            if (Array.isArray(saved)) {
                productionLog = saved;
                console.log(`[Storage] ${saved.length} üretim kaydı geri yüklendi (${loadedExcelFileName})`);
            }
        }
    } catch(e) {
        console.warn('localStorage okuma hatası:', e);
    }
}

// Reçete & Yıllık Tahmin Veritabanlarını Yükle
function loadParcaReceteTuketimDatabase() {
    // 1. Reçete başı tüketim
    fetch('../veritabanlari/parca_recete_tuketim.json')
        .then(res => res.json())
        .then(data => {
            if (data && typeof data === 'object') parcaReceteTuketimMap = data;
        })
        .catch(() => {});

    // 2. Parça - Makine Çoklu Reçete Veritabanı
    fetch('../veritabanlari/parca_makine_receteleri.json')
        .then(res => res.json())
        .then(data => {
            if (data && typeof data === 'object') {
                parcaMakineReceteleriMap = data;
                console.log(`[DB] ${Object.keys(data).length} parça için makine reçete verisi yüklendi.`);
            }
        })
        .catch(() => {});

    // 3. Makine Yıllık Üretim Tahminleri
    fetch('../veritabanlari/makine_yillik_tahminler.json')
        .then(res => res.json())
        .then(data => {
            if (data && typeof data === 'object') {
                makineYillikTahminlerMap = data;
                console.log(`[DB] ${Object.keys(data).length} makine için yıllık tahmin verisi yüklendi.`);
            }
        })
        .catch(() => {});

    // 4. Empirik Formül Hariç Tutma Kuralları
    fetch('../veritabanlari/parti_formulu_haric_kurallar.json')
        .then(res => res.json())
        .then(data => {
            if (data && typeof data === 'object') {
                partiFormuluHaricKurallar = data;
                console.log(`[DB] Empirik parti formülü hariç tutma kuralları yüklendi.`);
            }
        })
        .catch(() => {});
}
loadParcaReceteTuketimDatabase();

// Global Üretim Geçmişi Storage Fonksiyonları
function saveProductionHistoryToStorage() {
    try {
        localStorage.setItem('filter_global_production_history', JSON.stringify(productionHistory));
    } catch(e) {
        console.warn('localStorage üretim geçmişi kayıt hatası:', e);
    }
}
function loadProductionHistoryFromStorage() {
    try {
        const raw = localStorage.getItem('filter_global_production_history');
        if (raw) {
            const saved = JSON.parse(raw);
            if (Array.isArray(saved)) {
                productionHistory = saved;
            }
        }
    } catch(e) {
        console.warn('localStorage üretim geçmişi okuma hatası:', e);
    }
}

function autoApplyHistoricalCompletions() {
    if (!productionHistory || productionHistory.length === 0) return;
    
    let appliedCount = 0;
    
    // Her bir geçmiş kayıt için kontrol et
    productionHistory.forEach(hist => {
        if (!hist.kod || hist.tamamlandi === false) return;
        const code = String(hist.kod).trim().toUpperCase();
        const histKaynak = String(hist.kaynak || '').trim().toUpperCase();
        const station = hist.istasyon || 'Tüm İstasyonlar';
        
        // Üretim Takip satırlarında bu kod var mı?
        const matchingReqs = uretimTakipRows.filter(u => {
            const matchCode = u.kod === code;
            if (!matchCode) return false;
            if (histKaynak) {
                const uKaynak = String(u.kaynak || '').trim().toUpperCase();
                return uKaynak === histKaynak || uKaynak.includes(histKaynak) || histKaynak.includes(uKaynak);
            }
            return true;
        });
        
        if (matchingReqs.length > 0) {
            // Zaten productionLog'da var mı kontrol et
            const existingStLogs = productionLog.filter(log => log.kod === code && (log.station === station || log.station === 'Tüm İstasyonlar' || !log.station));
            const existingStProd = existingStLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
            const totalNeeded = matchingReqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
            
            const neededToAdd = Math.max(0.0, totalNeeded - existingStProd);
            if (neededToAdd > 0) {
                productionLog.push({
                    rowIndex: uretimTakipRows.length + productionLog.length + 5,
                    kod: code,
                    adet: neededToAdd,
                    fazla: 0,
                    station: station,
                    autoCompleted: true
                });
                appliedCount++;
            }
        }
    });
    
    if (appliedCount > 0) {
        saveProductionLogToStorage();
        console.log(`[AutoMatch] ${appliedCount} adet geçmiş üretim kaydı yeni dosyaya otomatik uygulandı.`);
    }
}

function exportProductionHistoryJSON() {
    loadProductionHistoryFromStorage();
    // Eğer productionHistory boşsa ama productionLog varsa senkronize et
    if (productionHistory.length === 0 && productionLog.length > 0) {
        productionLog.forEach(log => {
            const req = uretimTakipRows.find(u => u.kod === log.kod);
            productionHistory.push({
                kaynak: req ? (req.kaynak || '') : '',
                kod: log.kod,
                adet: log.adet,
                istasyon: log.station || 'Tüm İstasyonlar',
                tarih: new Date().toISOString(),
                tamamlandi: true
            });
        });
        saveProductionHistoryToStorage();
    }
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(productionHistory, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute("href", dataStr);
    downloadAnchor.setAttribute("download", "uretim_gecmisi.json");
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
    showToast("Üretim geçmişi 'uretim_gecmisi.json' olarak indirildi.", "success");
}
window.exportProductionHistoryJSON = exportProductionHistoryJSON;

function importProductionHistoryJSON(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const parsed = JSON.parse(e.target.result);
            if (Array.isArray(parsed)) {
                productionHistory = parsed;
                saveProductionHistoryToStorage();
                autoApplyHistoricalCompletions();
                recalculateAll();
                renderTab(currentTab);
                showToast(`${parsed.length} adet üretim geçmişi kaydı başarıyla yüklendi ve uygulandı!`, "success");
            } else {
                showToast("Geçersiz JSON formatı! Liste formatında veri bekleniyor.", "error");
            }
        } catch(err) {
            console.error(err);
            showToast("JSON dosyası okunurken hata oluştu!", "error");
        }
    };
    reader.readAsText(file);
}
window.importProductionHistoryJSON = importProductionHistoryJSON;

function saveDowntimeMapToStorage() {
    try {
        localStorage.setItem(`downtime_${loadedExcelFileName}`, JSON.stringify(downtimeMap));
    } catch(e) {}
}
function loadDowntimeMapFromStorage() {
    try {
        const raw = localStorage.getItem(`downtime_${loadedExcelFileName}`);
        if (raw) {
            downtimeMap = JSON.parse(raw);
        } else {
            downtimeMap = {};
        }
    } catch(e) {
        downtimeMap = {};
    }
}
function saveHiddenColsToStorage() {
    try { localStorage.setItem('hiddenStationCols', JSON.stringify([...hiddenStationCols])); } catch(e) {}
}
function loadHiddenColsFromStorage() {
    try {
        const raw = localStorage.getItem('hiddenStationCols');
        hiddenStationCols = raw ? new Set(JSON.parse(raw)) : new Set();
    } catch(e) { hiddenStationCols = new Set(); }
}
function clearProductionLogStorage(fileName) {
    try {
        localStorage.removeItem(`prodLog_${fileName}`);
    } catch(e) {}
}

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
let excludedHariciKodlar = new Set(); // Codes that are excluded (Merdane, Boru vs.)
let rawMaterialsRows = [];    // Rows of Hammadde Sipariş Listesi
let hammaddeSheetRows = [];   // Rows from Excel HAMMADDE / HAMMADDE SİPARİŞ sheet

// Pagination States
const PAGE_SIZE = 12;
let paginationState = {
    takip: { page: 1, total: 0, filtered: [] },
    assemblyLeft: { page: 1, total: 0, filtered: [] },
    assemblyRight: { page: 1, total: 0, filtered: [] },
    station: { page: 1, total: 0, filtered: [] },
    ul: { page: 1, total: 0, filtered: [] },
    raw: { page: 1, total: 0, filtered: [] }
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
            rawExcelArrayBuffer = e.target.result; // Orijinal buffer'ı ExcelJS için sakla
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, {
                type: 'array',
                cellFormula: true,
                cellNF: true,
                cellStyles: true
            });
            
            loadedFileName.textContent = `Yüklenen Dosya: ${file.name}`;
            loadedExcelFileName = file.name.replace(/[^a-zA-Z0-9_\-\.]/g, '_');
            parseWorkbook();
            loadProductionHistoryFromStorage(); // Global üretim geçmişini yükle
            loadProductionLogFromStorage(); // localStorage'dan mevcut dosya logunu geri yükle
            autoApplyHistoricalCompletions(); // Geçmişte yapılmış parçaları otomatik eşleştir
            loadDowntimeMapFromStorage();
            loadHiddenColsFromStorage();
            
            // YENİ: Montaj sayfası var mı kontrol et, yoksa menüden gizle
            const assemblyBtn = document.querySelector('button[data-tab="assembly"]');
            if (assemblyBtn) {
                if (workbook.Sheets["MONTAJ OTOMASYON İZLEME"] || workbook.Sheets["FINAL MONTAJ İZLEME"]) {
                    assemblyBtn.style.display = 'flex';
                } else {
                    assemblyBtn.style.display = 'none';
                    // Eğer montaj tabındayken yeniden dosya yüklenirse ve montaj yoksa dashboarda dön
                    if (currentTab === 'assembly') switchTab('dashboard');
                }
            }
            
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
    downtimeMap = {};
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
    unitTimeMap = {}; // { 'KOD': { 'ISTASYON_ADI': süre_dk } }
    selectedSourceFiles.clear();

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
                                uretilecek: qty,
                                orijinalUretilecek: qty
                            });
                        }
                    }
                }
            }
            
            // Eğer sayfa "Tüm Veriler" ise birim işlem sürelerini çıkar
            if (sName.toUpperCase().replace(/I/g, 'İ').includes('TÜM VERİLER') || sName.toUpperCase().includes('TUM VERILER')) {
                const kodIdx = headers.indexOf('kod');
                const istasyonIdx = headers.findIndex(h => h.includes('iş istasyonu') || h.includes('is istasyonu'));
                const sureIdx = headers.findIndex(h => h.includes('birim işlem süresi') || h.includes('birim islem suresi'));
                const setupIdx = headers.findIndex(h => h.includes('hazırlık süresi') || h.includes('hazirlik suresi'));
                
                if (kodIdx !== -1 && istasyonIdx !== -1 && sureIdx !== -1) {
                    for (let r = range.s.r + 1; r <= range.e.r; r++) {
                        const kodCell = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + kodIdx })];
                        const istCell = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + istasyonIdx })];
                        const sureCell = sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + sureIdx })];
                        
                        if (kodCell && kodCell.v && istCell && istCell.v) {
                            const code = String(kodCell.v).trim().toUpperCase();
                            const istasyon = String(istCell.v).trim().toUpperCase();
                            const sure = sureCell ? parseFloat(sureCell.v) || 0 : 0;
                            const setup = (setupIdx !== -1) ? (sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + setupIdx })] ? parseFloat(sheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + setupIdx })].v) || 0 : 0) : 0;
                            
                            if (!unitTimeMap[code]) unitTimeMap[code] = {};
                            unitTimeMap[code][istasyon] = { sure: sure, setup: setup };
                        }
                    }
                }
            }
        }
    }

    // 0. Parse HARİCİ_KODLAR (Hidden Sheet)
    const hariciSheet = workbook.Sheets["HARİCİ_KODLAR"];
    if (hariciSheet && hariciSheet['!ref']) {
        const range = XLSX.utils.decode_range(hariciSheet['!ref']);
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            const cell = hariciSheet[XLSX.utils.encode_cell({ r: r, c: 0 })];
            if (cell && cell.v) {
                excludedHariciKodlar.add(String(cell.v).trim().toUpperCase());
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
            const cellL = utSheet[XLSX.utils.encode_cell({ r: r, c: 11 })]; // Col L (İSTASYON)
            
            if (cellI && cellI.v !== undefined && cellI.v !== null && String(cellI.v).trim() !== "") {
                const prodQty = cellJ ? parseFloat(cellJ.v) || 0 : 0;
                productionLog.push({
                    rowIndex: r + 1,
                    kod: String(cellI.v).trim().toUpperCase(),
                    adet: prodQty,
                    fazla: 0,
                    station: cellL && cellL.v ? String(cellL.v).trim() : 'Tüm İstasyonlar'
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
        "ÜRETİM LİSTESİ", "Tüm Veriler", "HAMMADDE SİPARİŞ",
        "Üretim Takip", "Rotasızlar", "HARİCİ_KODLAR"
    ];

    const hiddenStations = [
        "HAMMADDE", "MONTAJ OTOMASYON İZLEME", "FINAL MONTAJ İZLEME", "HARİCİ_KODLAR"
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
            let rawHeader = hCell ? String(hCell.v).trim() : `Sütun ${c+1}`;
            
            const upHeader = rawHeader.toUpperCase();
            if (upHeader === 'KOD') rawHeader = 'Kod';
            else if (upHeader === 'ÖNCELİK SIRASI' || upHeader === 'ONCELIK SIRASI') rawHeader = 'Öncelik Sırası';
            else if (upHeader === 'MALZEME ADI') rawHeader = 'Malzeme Adı';
            else if (upHeader === 'HAMMADDE KOD' || upHeader === 'HAMMADDE KODU') rawHeader = 'Hammadde Kod';
            else if (upHeader === 'HAMMADDE') rawHeader = 'Hammadde';
            else if (upHeader === 'REZERVE EDİLECEK MİKTAR') rawHeader = 'Rezerve Edilecek Miktar';
            else if (upHeader === 'ÜRETİLECEK MİKTAR' || upHeader === 'URETILECEK MIKTAR') rawHeader = 'Üretilecek Miktar';
            else if (upHeader === 'TOPLAM HAMMADDE MİKTARI') rawHeader = 'Toplam Hammadde Miktarı';
            else if (upHeader === 'DURUM') rawHeader = 'Durum';
            else if (upHeader === 'HAZIRLIK SÜRESİ' || upHeader === 'HAZIRLIK SURESI') rawHeader = 'Hazırlık Süresi';
            else if (upHeader === 'BİRİM İŞLEM SÜRESİ' || upHeader === 'BIRIM ISLEM SURESI' || upHeader === 'BİRİM İŞLEM' || upHeader === 'Birim İşlem') rawHeader = 'Birim İşlem Süresi';
            else if (upHeader === 'TOPLAM SÜRE' || upHeader === 'TOPLAM SURE') rawHeader = 'Toplam Süre';
            else if (upHeader === 'SAAT') rawHeader = 'Saat';
            else if (upHeader === 'KÜMÜLATİF SÜRE' || upHeader === 'KÜMÜLATİF' || upHeader === 'KUMULATIF SURE') rawHeader = 'Kümülatif Süre';

            headers.push(rawHeader);
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
            
            // Only add to the UI left menu if it's not a hidden background station
            if (!hiddenStations.includes(sName)) {
                stationList.push(sName);
            }
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

    // 6. Parse HAMMADDE / HAMMADDE SİPARİŞ Sheet (Strictly Raw Materials)
    hammaddeSheetRows = [];
    const hamSheet = workbook.Sheets["HAMMADDE"] || workbook.Sheets["Hammadde"];
    if (hamSheet && hamSheet['!ref']) {
        const range = XLSX.utils.decode_range(hamSheet['!ref']);
        let headers = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = hamSheet[XLSX.utils.encode_cell({ r: range.s.r, c: c })];
            headers.push(cell ? String(cell.v).trim().toLowerCase() : `sütun ${c+1}`);
        }
        const kodIdx = headers.indexOf('kod');
        const hKodIdx = headers.findIndex(h => h === 'hammadde kod' || h === 'hammadde kodu' || h.includes('hammadde kod'));
        const hAdIdx = headers.findIndex(h => h === 'hammadde' || h === 'hammadde adı' || h === 'hammadde adi');
        const hMiktarIdx = headers.findIndex(h => h === 'hammadde miktar' || h === 'hammadde miktarı');
        const uMiktarIdx = headers.findIndex(h => h.includes('üretilecek') && h.includes('miktar'));

        if (hKodIdx !== -1) {
            for (let r = range.s.r + 1; r <= range.e.r; r++) {
                const cellHKod = hamSheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + hKodIdx })];
                if (!cellHKod || cellHKod.v === undefined || cellHKod.v === null) continue;
                const hKodVal = String(cellHKod.v).trim();
                if (!hKodVal || hKodVal === '-' || hKodVal.toUpperCase() === 'NAN') continue;

                const cellKod = kodIdx !== -1 ? hamSheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + kodIdx })] : null;
                const cellHAd = hAdIdx !== -1 ? hamSheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + hAdIdx })] : null;
                const cellHMiktar = hMiktarIdx !== -1 ? hamSheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + hMiktarIdx })] : null;
                const cellUMiktar = uMiktarIdx !== -1 ? hamSheet[XLSX.utils.encode_cell({ r: r, c: range.s.c + uMiktarIdx })] : null;

                const pKod = cellKod && cellKod.v ? String(cellKod.v).trim().toUpperCase() : '';
                const hAd = cellHAd && cellHAd.v ? String(cellHAd.v).trim() : '';
                const hMiktar = cellHMiktar ? parseFloat(cellHMiktar.v) || 0 : 1.0;
                const uMiktar = cellUMiktar ? parseFloat(cellUMiktar.v) || 0 : 0;

                hammaddeSheetRows.push({
                    parcaKodu: pKod,
                    hKod: hKodVal,
                    hAd: hAd,
                    hBirimMiktar: hMiktar,
                    uretilecek: uMiktar
                });
                if (pKod) {
                    parcaBirimHammaddeMap[pKod] = {
                        birimMiktar: hMiktar,
                        hKod: hKodVal,
                        hAd: hAd
                    };
                }
            }
        }
    }

    // Eğer HAMMADDE sayfası yoksa fallback: HAMMADDE SİPARİŞ sayfasını oku
    if (hammaddeSheetRows.length === 0) {
        const hsSheet = workbook.Sheets["HAMMADDE SİPARİŞ"] || workbook.Sheets["HAMMADDE SİPARİŞİ"] || workbook.Sheets["HAMMADDE SIPARIS"];
        if (hsSheet && hsSheet['!ref']) {
            const range = XLSX.utils.decode_range(hsSheet['!ref']);
            for (let r = range.s.r + 1; r <= range.e.r; r++) {
                const cellA = hsSheet[XLSX.utils.encode_cell({ r: r, c: 0 })]; // Hammadde Kod
                const cellB = hsSheet[XLSX.utils.encode_cell({ r: r, c: 1 })]; // Hammadde Adı
                const cellC = hsSheet[XLSX.utils.encode_cell({ r: r, c: 2 })]; // Toplam Miktar
                if (cellA && cellA.v) {
                    const hKod = String(cellA.v).trim();
                    const hAd = cellB && cellB.v ? String(cellB.v).trim() : '';
                    const qty = cellC ? parseFloat(cellC.v) || 0 : 0;
                    if (hKod && hKod !== '-' && hKod.toUpperCase() !== 'NAN' && !hKod.toUpperCase().includes('HAMMADDE')) {
                        hammaddeSheetRows.push({
                            parcaKodu: '',
                            hKod: hKod,
                            hAd: hAd,
                            hBirimMiktar: 1.0,
                            uretilecek: qty,
                            isDirectOrder: true
                        });
                    }
                }
            }
        }
    }
}

// Live spreadsheet calculation engine (SUMIF, SUMIFS, FIFO matching, limits, sets)
function recalculateAll() {
    // 1. Calculate total produced quantity for each part code from the production log
    const requiredStationsMap = {};
    for (const [stName, rows] of Object.entries(stationSheetsMap)) {
        for (const row of rows) {
            const code = String(row['Kod'] || '').trim().toUpperCase();
            if (code) {
                if (!requiredStationsMap[code]) requiredStationsMap[code] = [];
                requiredStationsMap[code].push(stName);
            }
        }
    }

    const codeStationSums = {}; 
    const codeGlobalSums = {};  

    for (const log of productionLog) {
        const code = log.kod;
        const adet = parseFloat(log.adet) || 0;
        const st = log.station || 'Tüm İstasyonlar';

        if (st === 'Tüm İstasyonlar') {
            codeGlobalSums[code] = (codeGlobalSums[code] || 0) + adet;
        } else {
            if (!codeStationSums[code]) codeStationSums[code] = {};
            codeStationSums[code][st] = (codeStationSums[code][st] || 0) + adet;
        }
    }

    const totalProducedMap = {};
    const allCodes = new Set([...Object.keys(codeGlobalSums), ...Object.keys(codeStationSums)]);
    
    for (const code of allCodes) {
        const globalSum = codeGlobalSums[code] || 0;
        const reqStations = requiredStationsMap[code] || [];
        
        let minStationSum = 0;
        if (reqStations.length > 0) {
            minStationSum = Infinity;
            for (const reqSt of reqStations) {
                const sSum = (codeStationSums[code] && codeStationSums[code][reqSt]) || 0;
                if (sSum < minStationSum) minStationSum = sSum;
            }
        } else {
            let sumOther = 0;
            if (codeStationSums[code]) {
                for (const sSum of Object.values(codeStationSums[code])) {
                    sumOther += sSum;
                }
            }
            minStationSum = sumOther;
        }

        totalProducedMap[code] = globalSum + minStationSum;
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

        // Formül hesaplamaları (Üretilecek Miktar güncellendikten sonra yapılmalı)
        let kumulatif = 0;
        for (const row of rows) {
            const code = String(row['Kod'] || '').trim().toUpperCase();
            let hazirlikKey = Object.keys(row).find(k => k.toLowerCase() === 'hazırlık süresi') || 'Hazırlık Süresi';
            let birimKey = Object.keys(row).find(k => k.toLowerCase() === 'birim işlem süresi') || 'Birim İşlem Süresi';
            let miktarKey = Object.keys(row).find(k => k.toLowerCase() === 'üretilecek miktar') || 'Üretilecek Miktar';
            let toplamSureKey = Object.keys(row).find(k => k.toLowerCase() === 'toplam süre') || 'Toplam Süre';
            let saatKey = Object.keys(row).find(k => k.toLowerCase() === 'saat') || 'Saat';
            let kumulatifKey = Object.keys(row).find(k => k.toLowerCase() === 'kümülatif süre') || 'Kümülatif Süre';
            
            const hazirlik = parseFloat(row[hazirlikKey]) || 0;
            const birim = parseFloat(row[birimKey]) || 0;
            const miktar = parseFloat(row[miktarKey]) || 0;
            
            // Tamamlanma kontrolü
            const reqs = uretimTakipRows.filter(u => u.kod === code);
            let isDone = false;
            if (reqs.length > 0) {
                const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
                const stLogs = productionLog.filter(log => log.kod === code && (log.station === stName || log.station === 'Tüm İstasyonlar' || !log.station));
                const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
                if (totalReq > 0 && stProd >= totalReq) {
                    isDone = true;
                }
            }

            if (isDone) {
                row[toplamSureKey] = 0;
                row[saatKey] = 0;
            } else {
                row[toplamSureKey] = hazirlik + (birim * miktar);
                row[saatKey] = row[toplamSureKey] / 86400;
            }
            
            const saatVal = parseFloat(row[saatKey]) || 0;
            kumulatif += saatVal;
            row[kumulatifKey] = kumulatif;
        }
    }

    // 6. Update Raw Materials (Hammadde Sipariş) calculation
    const rawAgg = {};
    if (hammaddeSheetRows.length > 0) {
        hammaddeSheetRows.forEach(row => {
            const hKod = row.hKod;
            if (!hKod) return;

            let totalReq = row.uretilecek || 0;
            let totalProd = 0;

            if (row.parcaKodu) {
                const pKod = row.parcaKodu;
                // Önce kullanıcının değiştirdiği miktara bak (uretimListesiMap), yoksa orijinal Excel değeri
                if (uretimListesiMap[pKod] !== undefined) {
                    totalReq = uretimListesiMap[pKod];
                } else {
                    const reqs = uretimTakipRows.filter(u => u.kod === pKod);
                    if (reqs.length > 0) {
                        totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
                    } else {
                        totalReq = row.uretilecek || 0;
                    }
                }

                // Bu parçaya ait üretim logları (Global veya herhangi bir istasyonda girilen en yüksek miktar)
                const allLogs = productionLog.filter(log => log.kod === pKod);
                const globalSum = allLogs.filter(log => !log.station || log.station === 'Tüm İstasyonlar').reduce((sum, l) => sum + parseFloat(l.adet), 0.0);

                const stSums = {};
                allLogs.filter(log => log.station && log.station !== 'Tüm İstasyonlar').forEach(l => {
                    stSums[l.station] = (stSums[l.station] || 0) + parseFloat(l.adet);
                });
                const maxStSum = Object.values(stSums).length > 0 ? Math.max(...Object.values(stSums)) : 0;

                totalProd = Math.min(totalReq, Math.max(globalSum, maxStSum));
            } else if (row.isDirectOrder) {
                totalReq = row.uretilecek || 0;
                totalProd = 0;
            }

            let origReq = row.uretilecek || 0;
            if (row.parcaKodu) {
                const pKod = row.parcaKodu;
                const ulItem = uretimListesiRows.find(u => u.kod === pKod);
                if (ulItem && ulItem.orijinalUretilecek !== undefined) {
                    origReq = ulItem.orijinalUretilecek;
                } else {
                    const reqs = uretimTakipRows.filter(u => u.kod === pKod);
                    if (reqs.length > 0) {
                        origReq = reqs.reduce((sum, u) => sum + (u.orijinalUretilecek !== undefined ? u.orijinalUretilecek : u.uretilecek), 0.0);
                    }
                }
            } else if (row.isDirectOrder) {
                origReq = row.uretilecek || 0;
            }

            const hBirim = row.hBirimMiktar || 1.0;
            const hToplam = totalReq * hBirim;
            const hOrijinalToplam = origReq * hBirim;
            const hEkstraToplam = Math.max(0, hToplam - hOrijinalToplam);
            const extraQty = Math.max(0, totalReq - origReq);
            const hUretilen = totalProd * hBirim;
            const hKalan = Math.max(0, hToplam - hUretilen);

            if (!rawAgg[hKod]) {
                rawAgg[hKod] = {
                    kod: hKod,
                    ad: row.hAd || '-',
                    toplamGereken: 0,
                    toplamOrijinalGereken: 0,
                    toplamEkstraHammadde: 0,
                    uretilenDusulen: 0,
                    kalanSiparis: 0,
                    details: []
                };
            } else if ((!rawAgg[hKod].ad || rawAgg[hKod].ad === '-') && row.hAd) {
                rawAgg[hKod].ad = row.hAd;
            }
            rawAgg[hKod].toplamGereken += hToplam;
            rawAgg[hKod].toplamOrijinalGereken += hOrijinalToplam;
            rawAgg[hKod].toplamEkstraHammadde += hEkstraToplam;
            rawAgg[hKod].uretilenDusulen += hUretilen;
            rawAgg[hKod].kalanSiparis += hKalan;
            // Parça bazlı detay bilgisi
            if (row.parcaKodu) {
                rawAgg[hKod].details.push({
                    parcaKodu: row.parcaKodu,
                    birimMiktar: row.hBirimMiktar || 1,
                    uretilecek: totalReq,
                    orijinalUretilecek: origReq,
                    extraQty: extraQty,
                    toplamMiktar: hToplam,
                    orijinalToplamMiktar: hOrijinalToplam,
                    extraMiktar: hEkstraToplam,
                    uretilenMiktar: hUretilen,
                    kalanMiktar: hKalan
                });
            }
        });
    }
    rawMaterialsRows = Object.values(rawAgg);
}

// -------------------------------------------------------------
// TAB SWITCH & NAVIGATION
// -------------------------------------------------------------
function switchTab(tabId) {
    window.switchTab = switchTab;
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
    } else if (tabId === 'workload') {
        renderWorkloadTab();
    } else if (tabId === 'raw-materials') {
        renderRawMaterialsTab();
    } else if (tabId === 'performance') {
        renderPerformanceTab();
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

    // Update Filter Status Badge in Genel Durum
    const filterBadge = document.getElementById('source-file-filter-status-badge');
    const filterCount = document.getElementById('source-file-filter-count');
    if (filterBadge && filterCount) {
        if (selectedSourceFiles.size > 0 && selectedSourceFiles.size < dosyaTakipRows.length) {
            filterBadge.style.display = 'inline-flex';
            filterCount.textContent = selectedSourceFiles.size;
        } else {
            filterBadge.style.display = 'none';
        }
    }

    // 2. Filter and Render files
    const searchVal = document.getElementById('dashboard-search') ? document.getElementById('dashboard-search').value.toLowerCase().trim() : '';
    let filteredFiles = dosyaTakipRows;
    if (searchVal) {
        filteredFiles = dosyaTakipRows.filter(f => f.kaynak.toLowerCase().includes(searchVal));
    }

    const container = document.getElementById('dashboard-files-list');
    if (!container) return;
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

        const isSelected = selectedSourceFiles.has(f.kaynak);
        const card = document.createElement('div');
        card.className = `file-progress-card ${isSelected ? 'selected' : ''}`;
        card.setAttribute('data-kaynak', f.kaynak);
        card.title = isSelected ? 'Filtreden çıkarmak için tıklayın' : 'İstasyon ve üretim listelerinde filtrelemek için tıklayın';

        card.innerHTML = `
            <div class="card-title-row">
                <div style="display: flex; align-items: center; overflow: hidden; gap: 4px;">
                    <span class="file-progress-card-check"><i class="fa-solid fa-check"></i></span>
                    <span class="file-name-label" title="${f.kaynak}">${f.kaynak}</span>
                </div>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span class="file-pct" style="color: ${colorClass};">${pct}%</span>
                    <button type="button" class="btn btn-sm modal-open-btn" onclick="event.stopPropagation(); window.openDetailsModal('${f.kaynak}')" title="Dosya Eksik Detaylarını Aç" style="background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.15); color: #c7d2fe; padding: 2px 7px; font-size: 11px; border-radius: 4px; cursor: pointer; display: inline-flex; align-items: center; gap: 4px;">
                        <i class="fa-solid fa-expand"></i> İncele
                    </button>
                </div>
            </div>
            <div class="progress-container">
                <div class="progress-fill" style="width: ${pct}%; background-color: ${colorClass}; box-shadow: 0 0 8px ${glowClass};"></div>
            </div>
            <div class="card-stats-row">
                <div class="stat-item"><i class="fa-solid fa-check text-green"></i> <span>${f.hazir} / ${f.toplam} hazır</span></div>
                <div class="stat-item"><i class="fa-solid fa-hourglass-half text-orange"></i> <span>${f.eksik} eksik</span></div>
            </div>
        `;

        // Click to toggle filter selection
        card.addEventListener('click', () => {
            if (selectedSourceFiles.has(f.kaynak)) {
                selectedSourceFiles.delete(f.kaynak);
            } else {
                selectedSourceFiles.add(f.kaynak);
            }
            updateSourceFileFilterUI();
        });

        // Double click to filter production log by this source file
        card.addEventListener('dblclick', (e) => {
            e.stopPropagation();
            const takipSearch = document.getElementById('takip-search');
            if (takipSearch) takipSearch.value = f.kaynak;
            switchTab('production');
        });

        container.appendChild(card);
    });
}

// --- KAYNAK DOSYA ÇOKLU SEÇİM VE FİLTRE YÖNETİMİ ---
function selectAllSourceFiles(isAll) {
    if (isAll) {
        dosyaTakipRows.forEach(f => selectedSourceFiles.add(f.kaynak));
    } else {
        selectedSourceFiles.clear();
    }
    updateSourceFileFilterUI();
}

function updateSourceFileFilterUI() {
    const allSources = dosyaTakipRows.map(f => f.kaynak);
    
    // Genel Durum rozetini güncelle
    const badge = document.getElementById('source-file-filter-status-badge');
    const countEl = document.getElementById('source-file-filter-count');
    if (badge && countEl) {
        if (selectedSourceFiles.size > 0 && selectedSourceFiles.size < allSources.length) {
            badge.style.display = 'inline-flex';
            countEl.textContent = selectedSourceFiles.size;
        } else {
            badge.style.display = 'none';
        }
    }

    // İstasyon sayfasındaki buton etiketini güncelle
    const labelEl = document.getElementById('station-source-filter-label');
    if (labelEl) {
        if (selectedSourceFiles.size === 0 || selectedSourceFiles.size === allSources.length) {
            labelEl.textContent = 'Kaynak: Tümü';
        } else {
            labelEl.textContent = `Kaynak: ${selectedSourceFiles.size} Seçili`;
        }
    }

    // Genel Durum kartlarını güncelle
    renderDashboard();

    // Eğer İstasyon sekmesindeysek tabloyu ve sidebarı güncelle
    if (currentTab === 'stations') {
        paginationState.station.page = 1;
        renderStationsTab();
    } else if (currentTab === 'uretim-listesi') {
        paginationState.ul.page = 1;
        filterAndPaginateUlTable();
    }

    // Eğer dropdown açıksa checkboxları güncelle
    const dd = document.getElementById('station-source-filter-dropdown');
    if (dd && dd.style.display === 'block') {
        renderStationSourceFilterOptions();
    }
}

function toggleStationSourcePicker() {
    const dd = document.getElementById('station-source-filter-dropdown');
    if (!dd) return;
    if (dd.style.display === 'none' || !dd.style.display) {
        renderStationSourceFilterOptions();
        dd.style.display = 'block';
    } else {
        dd.style.display = 'none';
    }
}

function renderStationSourceFilterOptions() {
    const listEl = document.getElementById('station-source-filter-list');
    const labelEl = document.getElementById('station-source-filter-label');
    if (!listEl) return;

    listEl.innerHTML = '';
    const allSources = dosyaTakipRows.map(f => f.kaynak);

    if (labelEl) {
        if (selectedSourceFiles.size === 0 || selectedSourceFiles.size === allSources.length) {
            labelEl.textContent = 'Kaynak: Tümü';
        } else {
            labelEl.textContent = `Kaynak: ${selectedSourceFiles.size} Seçili`;
        }
    }

    if (allSources.length === 0) {
        listEl.innerHTML = '<div style="color:var(--text-dim); font-size:12px; padding:6px;">Kaynak dosya bulunamadı.</div>';
        return;
    }

    allSources.forEach(src => {
        const isChecked = selectedSourceFiles.has(src);
        const rowDiv = document.createElement('label');
        rowDiv.style.display = 'flex';
        rowDiv.style.alignItems = 'center';
        rowDiv.style.gap = '8px';
        rowDiv.style.fontSize = '12px';
        rowDiv.style.color = isChecked ? '#a5b4fc' : '#e2e8f0';
        rowDiv.style.fontWeight = isChecked ? '600' : 'normal';
        rowDiv.style.cursor = 'pointer';
        rowDiv.style.padding = '4px 6px';
        rowDiv.style.borderRadius = '4px';
        rowDiv.style.transition = 'background 0.15s';
        rowDiv.onmouseover = () => { rowDiv.style.background = 'rgba(255,255,255,0.06)'; };
        rowDiv.onmouseout = () => { rowDiv.style.background = 'transparent'; };

        rowDiv.innerHTML = `
            <input type="checkbox" value="${src}" ${isChecked ? 'checked' : ''} style="accent-color: #6366f1; cursor: pointer; width: 14px; height: 14px;">
            <span style="flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${src}">${src}</span>
        `;

        const chk = rowDiv.querySelector('input');
        chk.addEventListener('change', (e) => {
            if (e.target.checked) {
                selectedSourceFiles.add(src);
            } else {
                selectedSourceFiles.delete(src);
            }
            updateSourceFileFilterUI();
        });

        listEl.appendChild(rowDiv);
    });
}

// Dropdown dışına tıklandığında kapatma
document.addEventListener('click', function(e) {
    const btn = document.getElementById('station-source-filter-btn');
    const dd = document.getElementById('station-source-filter-dropdown');
    if (btn && dd && !btn.contains(e.target) && !dd.contains(e.target)) {
        dd.style.display = 'none';
    }
});

window.selectAllSourceFiles = selectAllSourceFiles;
window.toggleStationSourcePicker = toggleStationSourcePicker;
window.updateSourceFileFilterUI = updateSourceFileFilterUI;

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

// Helper to bind search inputs robustly (handles input, keyup, and Enter key)
function bindSearchInput(inputId, callback) {
    const inputEl = document.getElementById(inputId);
    if (!inputEl) return;
    let timeout;
    const trigger = () => {
        clearTimeout(timeout);
        callback();
    };
    inputEl.addEventListener('input', () => {
        clearTimeout(timeout);
        timeout = setTimeout(trigger, 300);
    });
    inputEl.addEventListener('keyup', (e) => {
        if (e.key === 'Enter') {
            trigger();
        }
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

bindSearchInput('modal-search', renderModalData);
bindSearchInput('dashboard-search', renderDashboard);

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

function updateStationOptions(code) {
    const select = document.getElementById('prod-station');
    if (!select) return;
    
    // Clear and add default option
    select.innerHTML = `
        <option value="Tüm İstasyonlar">Tüm İstasyonlar (Komple Bitti)</option>
    `;
    
    code = code.toUpperCase().trim();
    if (!code) return;
    
    // Find stations requiring this code
    let stations = [];
    for (const [stName, rows] of Object.entries(stationSheetsMap)) {
        if (rows.some(r => String(r['Kod'] || '').trim().toUpperCase() === code)) {
            stations.push(stName);
        }
    }
    
    stations.forEach(st => {
        const option = document.createElement('option');
        option.value = st;
        option.textContent = `${st} İstasyonu`;
        select.appendChild(option);
    });
}

function initAutocomplete() {
    const codeInput = document.getElementById('prod-code');
    const autoList = document.getElementById('autocomplete-list');
    
    // Extract unique codes
    const uniqueCodes = [...new Set(uretimTakipRows.map(r => r.kod))].sort();

    codeInput.addEventListener('input', function() {
        const val = this.value.toUpperCase().trim();
        autoList.innerHTML = '';
        updateStationOptions(val);
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
                updateStationOptions(s);
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
    
    // Find stations requiring this code and calculate their completion status
    let stationsStatus = [];
    for (const [stName, rows] of Object.entries(stationSheetsMap)) {
        if (rows.some(r => String(r['Kod'] || '').trim().toUpperCase() === code)) {
            const stLogs = productionLog.filter(log => log.kod === code && (log.station === stName || log.station === 'Tüm İstasyonlar' || !log.station));
            const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
            const stPct = totalReq > 0 ? Math.min(100, Math.round((stProd / totalReq) * 100)) : 0;
            stationsStatus.push({
                name: stName,
                pct: stPct
            });
        }
    }

    // Detailed distribution by Kaynak Dosya (FIFO)
    let kaynakBreakdownHtml = '';
    if (reqs.length > 0) {
        kaynakBreakdownHtml = `
            <div class="part-info-row" style="margin-top: 10px; flex-direction:column; gap:6px; border-top:1px solid var(--border-color); padding-top:8px; width: 100%; flex: 1; min-height: 0;">
                <span class="part-info-label" style="font-weight:600; color:var(--primary);">Kaynak Dosya Dağılımı (FIFO):</span>
                <div style="display:flex; flex-direction:column; gap:8px; width:100%; margin-top:5px; flex: 1; min-height: 0; overflow-y:auto; padding-right:4px;">
                    ${reqs.map(r => {
                        const pct = r.uretilecek > 0 ? Math.min(100, Math.round((r.uretilen / r.uretilecek) * 100)) : 0;
                        const pctColor = pct >= 100 ? 'var(--success)' : pct > 0 ? 'var(--warning)' : 'var(--danger)';
                        return `
                            <div style="background:rgba(255,255,255,0.02); border:1px solid rgba(255,255,255,0.05); padding:6px 10px; border-radius:6px; width: 100%; box-sizing: border-box;">
                                <div style="display:flex; justify-content:space-between; font-size:11px; margin-bottom:4px; width: 100%;">
                                    <span style="font-weight:700; color:white;">${r.kaynak}</span>
                                    <span style="font-weight:600; color:${pctColor};">${r.uretilen} / ${r.uretilecek} Adet (${pct}%)</span>
                                </div>
                                <div style="width:100%; height:4px; background:rgba(255,255,255,0.08); border-radius:2px; overflow:hidden;">
                                    <div style="width:${pct}%; height:100%; background:${pctColor}; border-radius:2px;"></div>
                                </div>
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
        `;
    }

    box.innerHTML = `
        <div class="part-info-details" style="width: 100%;">
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
            <div class="part-info-row" style="margin-top: 5px; flex-direction:column; gap:6px; border-top:1px solid var(--border-color); padding-top:8px; width: 100%;">
                <span class="part-info-label" style="font-weight:600;">İstasyon Tamamlanma Durumları:</span>
                <div style="display:flex; flex-wrap:wrap; gap:5px; margin-top:3px; width:100%;">
                    ${stationsStatus.length > 0 
                        ? stationsStatus.map(st => {
                            const badgeClass = st.pct >= 100 ? 'badge-success' : st.pct > 0 ? 'badge-warning' : 'badge-danger';
                            return `<span class="badge ${badgeClass}" style="font-size:10px;">${st.name} (${st.pct}%)</span>`;
                        }).join('') 
                        : '<span style="color:var(--text-dim); font-size:12px;">Tanımlı istasyon yok. (Büyük ihtimalle Montaj)</span>'
                    }
                </div>
            </div>
            ${kaynakBreakdownHtml}
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
    const stationSelect = document.getElementById('prod-station');
    const selectedStation = stationSelect ? stationSelect.value : 'Tüm İstasyonlar';

    productionLog.push({
        rowIndex: uretimTakipRows.length + productionLog.length + 5, // Arbitrary index offset for new entries
        kod: code,
        adet: qty,
        fazla: 0,
        station: selectedStation
    });

    saveProductionLogToStorage(); // localStorage'a kaydet
    showToast(`"${code}" kodu için ${qty} adet üretim girildi.`, "success");
    document.getElementById('prod-qty').value = '';
    if (stationSelect) {
        stationSelect.value = 'Tüm İstasyonlar';
    }
    
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
    saveProductionLogToStorage(); // localStorage güncelle
    showToast(`Üretim kaydı silindi.`, "info");
    
    recalculateAll();
    renderProductionTab();
    showQuickPartInfo(deletedCode);
}

function clearAllProductionLogs() {
    if (!productionLog || productionLog.length === 0) return;
    if (confirm("Tüm üretim giriş günlüğü kayıtlarını silmek istediğinize emin misiniz?")) {
        productionLog = [];
        saveProductionLogToStorage();
        showToast("Tüm üretim günlüğü temizlendi.", "info");
        recalculateAll();
        renderProductionTab();
    }
}
window.clearAllProductionLogs = clearAllProductionLogs;

function renderProductionLog() {
    const tbody = document.getElementById('production-log-body');
    tbody.innerHTML = '';
    
    document.getElementById('prod-log-count').textContent = `${productionLog.length} Kayıt`;

    if (productionLog.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center" style="color: var(--text-dim); padding:20px;">Giriş yapılmadı.</td></tr>';
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
            <td><span class="badge" style="background:rgba(255,255,255,0.06); color:var(--text-main); font-size:10px;">${log.station || 'Tüm İstasyonlar'}</span></td>
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

function togglePartCompletion(code, shouldComplete, stationName) {
    const reqs = uretimTakipRows.filter(u => u.kod === code);
    if (reqs.length === 0) return;
    
    const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
    const targetStation = stationName || 'Tüm İstasyonlar';
    
    if (shouldComplete) {
        // Calculate production for this station
        const stLogs = productionLog.filter(log => log.kod === code && (log.station === targetStation || log.station === 'Tüm İstasyonlar' || !log.station));
        const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
        const needed = Math.max(0.0, totalReq - stProd);
        
        if (needed > 0) {
            productionLog.push({
                rowIndex: uretimTakipRows.length + productionLog.length + 5,
                kod: code,
                adet: needed,
                fazla: 0,
                station: targetStation,
                autoCompleted: true
            });
            saveProductionLogToStorage(); // localStorage güncelle
        }

        // Global Üretim Geçmişine (uretim_gecmisi) ekle / güncelle
        reqs.forEach(r => {
            const existingIdx = productionHistory.findIndex(h => h.kod === code && h.kaynak === r.kaynak && (h.istasyon === targetStation || h.istasyon === 'Tüm İstasyonlar'));
            if (existingIdx !== -1) {
                productionHistory[existingIdx].tamamlandi = true;
                productionHistory[existingIdx].adet = r.uretilecek;
                productionHistory[existingIdx].tarih = new Date().toISOString();
            } else {
                productionHistory.push({
                    kaynak: r.kaynak || '',
                    kod: code,
                    adet: r.uretilecek,
                    istasyon: targetStation,
                    tarih: new Date().toISOString(),
                    tamamlandi: true
                });
            }
        });
        saveProductionHistoryToStorage();
        showToast(`"${code}" parçası ${stationName ? stationName + ' istasyonunda' : ''} tamamlandı olarak işaretlendi ve geçmişe kaydedildi.`, "success");
    } else {
        // Remove autoCompleted logs for this code and station
        productionLog = productionLog.filter(log => {
            const isMatch = log.kod === code && log.autoCompleted && (log.station === targetStation || !stationName || log.station === 'Tüm İstasyonlar');
            return !isMatch;
        });
        
        // Recalculate to see if it's still 100%
        recalculateAll();
        
        // Check completion for this station
        const stLogs = productionLog.filter(log => log.kod === code && (log.station === targetStation || log.station === 'Tüm İstasyonlar' || !log.station));
        const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
        const stPct = totalReq > 0 ? (stProd / totalReq) * 100 : 0;
        
        if (stPct >= 100) {
            // Delete all logs matching this station to force it under 100%
            productionLog = productionLog.filter(log => !(log.kod === code && (log.station === targetStation || !stationName || log.station === 'Tüm İstasyonlar')));
        }

        // Global Üretim Geçmişinden kaldır veya tamamlandi=false yap
        reqs.forEach(r => {
            const existingIdx = productionHistory.findIndex(h => h.kod === code && h.kaynak === r.kaynak && (h.istasyon === targetStation || h.istasyon === 'Tüm İstasyonlar'));
            if (existingIdx !== -1) {
                productionHistory.splice(existingIdx, 1);
            }
        });
        saveProductionHistoryToStorage();
        saveProductionLogToStorage();
        
        showToast(`"${code}" parçasının ${stationName ? stationName + ' istasyonundaki' : ''} tamamlandı işareti kaldırıldı.`, "info");
    }
    
    // Perform recalculation
    recalculateAll();
    
    // Render whichever tab is active
    renderTab(currentTab);
}
window.togglePartCompletion = togglePartCompletion;


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
    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="color: var(--text-dim); padding:20px;">Eşleşen parça bulunamadı.</td></tr>';
        return;
    }

    const pageRows = pState.filtered;
    
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
            <td style="white-space:nowrap; width:140px; min-width:140px; padding:6px 12px;">
                <div class="code-cell-wrapper">
                    <span class="code-cell-text" style="color:${excludedHariciKodlar.has(r.kod.trim().toUpperCase()) ? '#fda4af' : 'white'};">
                        ${excludedHariciKodlar.has(r.kod.trim().toUpperCase()) ? '<i class="fa-solid fa-triangle-exclamation" style="font-size:11px; margin-right:6px; opacity:0.9;" title="Harici İşlem / Harici Kod"></i>' : ''}
                        ${r.kod}
                    </span>
                    <button type="button" draggable="false" class="part-img-btn" onmousedown="event.stopPropagation()" onclick="event.stopPropagation(); window.openPartImageModal('${r.kod}')" onmouseenter="window.showPartHoverPreview(event, '${r.kod}')" onmousemove="window.movePartHoverPreview(event)" onmouseleave="window.hidePartHoverPreview()" title="Görseli Görüntüle">
                        <i class="fa-solid fa-image"></i>
                    </button>
                </div>
            </td>
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
document.getElementById('takip-prev-btn')?.addEventListener('click', () => {
    if (paginationState.takip.page > 1) {
        paginationState.takip.page--;
        renderTakipTable();
    }
});
document.getElementById('takip-next-btn')?.addEventListener('click', () => {
    const maxPage = Math.ceil(paginationState.takip.total / PAGE_SIZE);
    if (paginationState.takip.page < maxPage) {
        paginationState.takip.page++;
        renderTakipTable();
    }
});

bindSearchInput('takip-search', () => {
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
    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="color: var(--text-dim); padding:20px;">Eşleşen üst montaj bulunamadı.</td></tr>';
        return;
    }

    const pageRows = pState.filtered;
    const activeTabLeftSource = activeTabAssembly === 'otomasyon' ? montajOtomasyonLeft : finalMontajLeft;

    // Precompute which parents actually have children to avoid O(N^2) filtering later
    const parentsWithKids = new Set();
    activeTabLeftSource.forEach(l => {
        parentsWithKids.add(`${l.ustKod}___${l.kaynak}`);
    });

    pageRows.forEach(parent => {
        const pct = Math.round(parent.tamamlanma * 100);
        let colorClass = 'var(--danger)';
        if (pct >= 100) colorClass = 'var(--success)';
        else if (pct > 0) colorClass = 'var(--warning)';

        const parentKey = `${parent.ustKod}___${parent.kaynak}`;
        const hasKids = parentsWithKids.has(parentKey);
        const isExpandable = parent.tamamlanma < 1.0 && hasKids;
        const isExpanded = isExpandable && expandedParentAssembly === parentKey;

        const tr = document.createElement('tr');
        if (isExpandable) {
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

        const chevron = isExpandable
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
bindSearchInput('assembly-right-search', () => {
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
document.getElementById('assembly-right-prev-btn')?.addEventListener('click', () => {
    if (paginationState.assemblyRight.page > 1) {
        paginationState.assemblyRight.page--;
        renderAssemblyRightTable();
    }
});
document.getElementById('assembly-right-next-btn')?.addEventListener('click', () => {
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

bindSearchInput('rotasiz-search', renderRotasizlarTab);

// -------------------------------------------------------------
// 5. STATIONS VIEW
// -------------------------------------------------------------
function renderStationsTab() {
    // 1. Render stations sidebar list
    const sidebar = document.getElementById('stations-list');
    sidebar.innerHTML = '';

    let draggedItem = null;
    const searchInput = document.getElementById('station-search');
    const searchVal = searchInput ? searchInput.value.toLowerCase().trim() : '';

    stationList.forEach(st => {
        const btn = document.createElement('button');
        
        // Count how many parts in this station are fully complete (filtered by source file if active)
        const allStationRows = stationSheetsMap[st] || [];
        const rows = selectedSourceFiles.size > 0 
            ? allStationRows.filter(r => {
                const code = String(r['Kod'] || '').trim().toUpperCase();
                const reqs = uretimTakipRows.filter(u => u.kod === code);
                return reqs.some(u => selectedSourceFiles.has(u.kaynak));
              })
            : allStationRows;

        let completedCount = 0;
        rows.forEach(r => {
            const code = String(r['Kod'] || '').trim().toUpperCase();
            const reqs = uretimTakipRows.filter(u => u.kod === code);
            if (reqs.length > 0) {
                const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
                const stLogs = productionLog.filter(log => log.kod === code && (log.station === st || log.station === 'Tüm İstasyonlar' || !log.station));
                const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
                if (totalReq > 0 && stProd >= totalReq) {
                    completedCount++;
                }
            }
        });

        // Check if there is any search match in this station
        let isSearchMatch = false;
        if (searchVal) {
            isSearchMatch = rows.some(r => {
                const code = String(r['Kod'] || '').toLowerCase();
                const mat = String(r['Malzeme Adı'] || '').toLowerCase();
                const hCode = String(r['Hammadde Kod'] || '').toLowerCase();
                return code.includes(searchVal) || mat.includes(searchVal) || hCode.includes(searchVal);
            });
        }

        const isDimmed = selectedSourceFiles.size > 0 && rows.length === 0;
        btn.className = `station-item-btn ${st === activeStation ? 'active' : ''} ${isSearchMatch ? 'search-match-pulse' : ''}`;
        btn.setAttribute('draggable', true);
        if (isDimmed) {
            btn.style.opacity = '0.45';
        } else {
            btn.style.opacity = '1';
        }
        
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
    const searchVal = document.getElementById('station-search') ? document.getElementById('station-search').value.toLowerCase().trim() : '';
    const statusFilter = document.getElementById('station-filter-status') ? document.getElementById('station-filter-status').value : 'all';
    const sortVal = document.getElementById('station-sort') ? document.getElementById('station-sort').value : 'priority-asc';
    
    const allStationRows = stationSheetsMap[activeStation] || [];
    const headers = stationHeadersMap[activeStation] || [];

    // Apply Source File Filter if active
    const rows = selectedSourceFiles.size > 0
        ? allStationRows.filter(r => {
            const code = String(r['Kod'] || '').trim().toUpperCase();
            const reqs = uretimTakipRows.filter(u => u.kod === code);
            return reqs.some(u => selectedSourceFiles.has(u.kaynak));
          })
        : allStationRows;

    const sourceCountText = selectedSourceFiles.size > 0 ? ` <span style="font-size:12px; color:#818cf8; font-weight:normal; margin-left:6px;">(${selectedSourceFiles.size} Kaynak Dosya Filtreli)</span>` : '';
    document.getElementById('current-station-title').innerHTML = `<i class="fa-solid fa-industry text-green"></i> ${activeStation} İstasyon İş Listesi${sourceCountText}`;

    // İstasyon Kapasite Rozetini Güncelle (Filtrelenmiş kaynak dosya satırlarına göre)
    const capBadge = document.getElementById('station-capacity-badge');
    if (capBadge) {
        if (!activeStation || rows.length === 0) {
            capBadge.style.display = 'none';
        } else {
            capBadge.style.display = 'flex';
            let gunlukSaat = DEFAULT_CAPACITY.varsayilan_gunluk_saat || 9;
            let makineSayisi = 1;
            if (DEFAULT_CAPACITY.istasyonlar) {
                if (DEFAULT_CAPACITY.istasyonlar[activeStation]) {
                    gunlukSaat = DEFAULT_CAPACITY.istasyonlar[activeStation].gunluk_saat || gunlukSaat;
                    makineSayisi = DEFAULT_CAPACITY.istasyonlar[activeStation].makine_sayisi || makineSayisi;
                } else {
                    const normActive = activeStation.trim().toUpperCase().replace(/\s+/g, '');
                    for (const [k, cfg] of Object.entries(DEFAULT_CAPACITY.istasyonlar)) {
                        if (k.trim().toUpperCase().replace(/\s+/g, '') === normActive) {
                            gunlukSaat = cfg.gunluk_saat || gunlukSaat;
                            makineSayisi = cfg.makine_sayisi || makineSayisi;
                            break;
                        }
                    }
                }
            }

            let totalHours = 0;
            rows.forEach(r => {
                let saatVal = parseFloat(r['Saat']) || 0;
                if (saatVal > 0) {
                    totalHours += (saatVal * 24);
                } else {
                    let topSure = parseFloat(r['Toplam Süre']) || 0;
                    totalHours += (topSure / 3600);
                }
            });

            const gunlukKapasite = gunlukSaat * makineSayisi;
            const tahminiGun = gunlukKapasite > 0 ? (totalHours / gunlukKapasite) : 0;

            capBadge.innerHTML = `
                <span style="color: var(--text-dim); display: flex; align-items: center; gap: 6px;">
                    <i class="fa-solid fa-gears text-blue"></i> <strong style="color:var(--text-primary); font-size:13px;">${makineSayisi}</strong> İstasyon
                </span>
                <span style="color: rgba(255,255,255,0.15);">|</span>
                <span style="color: var(--text-dim); display: flex; align-items: center; gap: 6px;">
                    <i class="fa-solid fa-clock text-yellow"></i> <strong style="color:var(--text-primary); font-size:13px;">${gunlukSaat}</strong> Saat/Gün
                </span>
                <span style="color: rgba(255,255,255,0.15);">|</span>
                <span style="color: var(--text-dim); display: flex; align-items: center; gap: 6px;">
                    <i class="fa-solid fa-calendar-check text-green"></i> Tahmini: <strong style="color:#10B981; font-weight:700; font-size:14px;">${tahminiGun.toFixed(1)}</strong> İş Günü <span style="color:rgba(255,255,255,0.65); font-weight:500; font-size:12px; margin-left:2px;">(${totalHours.toFixed(1)} Saat)</span>
                </span>
            `;
        }
    }

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
                const stLogs = productionLog.filter(log => log.kod === code && (log.station === activeStation || log.station === 'Tüm İstasyonlar' || !log.station));
                const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
                completionPct = totalReq > 0 ? Math.min(100, Math.round((stProd / totalReq) * 100)) : 0;
            }
            if (statusFilter === 'ready') return completionPct >= 100;
            if (statusFilter === 'production') return completionPct > 0 && completionPct < 100;
            if (statusFilter === 'missing') return completionPct === 0;
            return true;
        });
    }

    // 3. Sorting
    if (sortVal !== 'custom') {
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
    }

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
    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="text-center" style="color:var(--text-dim); padding:20px;">Bu istasyonda eşleşen iş listesi bulunamadı.</td></tr>';
        return;
    }

    // Select the key columns to display
    const ALL_DISPLAY_COLS = [
        'Öncelik Sırası', 'Kod', 'Malzeme Adı', 'Hammadde Kod', 'Hammadde',
        'Rezerve Edilecek Miktar', 'Üretilecek Miktar', 'Hammadde Miktar', 'Toplam Hammadde Miktarı',
        'Hazırlık Süresi', 'Birim İşlem Süresi', 'Toplam Süre', 'Saat', 'Kümülatif Süre', 'Durum'
    ];

    // Sadece bu sayfanın headers'ında bulunanları al (Saat / Kümülatif Süre dahil)
    const displayCols = ALL_DISPLAY_COLS.filter(c => c === 'Durum' || headers.includes(c));

    // Kullanıcının gizlediği sütunları çıkar
    const colsToShow = displayCols.filter(c => !hiddenStationCols.has(c));

    // Sütun seçici dropdown'ı güncelle
    _updateColPickerUI(displayCols);

    // Add Status header (always visible)
    const finalHeaders = colsToShow.includes('Durum') ? colsToShow : [...colsToShow, 'Durum'];

    // Create table header cells
    const trHead = document.createElement('tr');
    finalHeaders.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        if (h.includes('Miktar') || h.includes('Adet')) th.className = 'text-right';
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    const pageRows = pState.filtered;

    pageRows.forEach(row => {
        const code = String(row['Kod'] || '').trim().toUpperCase();
        
        // Calculate status from Üretim Takip
        const reqs = uretimTakipRows.filter(u => u.kod === code);
        let completionText = '-';
        let completionPct = 0;
        let isOverproduced = false;
        let badgeHtml = '<span class="badge badge-danger">Eksik</span>';
        
        if (reqs.length > 0) {
            const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
            const allLogs = productionLog.filter(log => log.kod === code);
            const stLogs = allLogs.filter(log => log.station === activeStation || log.station === 'Tüm İstasyonlar' || !log.station);
            const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
            const rawPct = totalReq > 0 ? Math.round((stProd / totalReq) * 100) : 0;
            completionPct = Math.min(100, rawPct);
            completionText = `${stProd} / ${totalReq}`;

            if (stProd > totalReq && totalReq > 0) {
                // Fazla Üretim: üretilen > hedef
                isOverproduced = true;
                const fazla = Math.round((stProd - totalReq) * 1000) / 1000;
                badgeHtml = `<span class="badge badge-overproduced"><i class="fa-solid fa-arrow-trend-up" style="font-size:10px;"></i> Fazla Üretim (+${fazla})</span>`;
            } else if (completionPct >= 100) {
                badgeHtml = `<span class="badge badge-success"><i class="fa-solid fa-check" style="font-size:10px;"></i> Tamamlandı</span>`;
            } else if (completionPct > 0) {
                badgeHtml = `<span class="badge badge-warning">Üretimde (${completionPct}%)</span>`;
            }
        }

        const tr = document.createElement('tr');
        if (isOverproduced) {
            tr.classList.add('station-row-overproduced');
        } else if (completionPct >= 100) {
            tr.classList.add('station-row-completed');
        }
        
        tr.draggable = true;
        tr.dataset.index = pState.filtered.indexOf(row);
        if (window.handleStationRowDragStart) {
            tr.addEventListener('dragstart', window.handleStationRowDragStart);
            tr.addEventListener('dragover', window.handleStationRowDragOver);
            tr.addEventListener('dragleave', window.handleStationRowDragLeave);
            tr.addEventListener('drop', window.handleStationRowDrop);
            tr.addEventListener('dragend', window.handleStationRowDragEnd);
        }

        finalHeaders.forEach(h => {
            const td = document.createElement('td');
            
            if (h === 'Durum') {
                if (reqs.length > 0) {
                    const isDone = completionPct >= 100 || isOverproduced;
                    td.innerHTML = `
                        <div class="status-cell-container" style="display: flex; align-items: center; gap: 8px; justify-content: center;">
                            <input type="checkbox" class="station-complete-checkbox" ${isDone ? 'checked' : ''} onchange="togglePartCompletion('${code}', this.checked, '${activeStation}')" title="Bitti / Üretildi Olarak İşaretle">
                            ${badgeHtml}
                        </div>
                    `;
                } else {
                    td.innerHTML = badgeHtml;
                }

            } else {
                let val = row[h];
                
                // Saat ve Kümülatif Süre formatı
                if ((h === 'Saat' || h === 'Kümülatif Süre') && typeof val === 'number') {
                    // val is decimal fraction of a day (e.g. 0.025)
                    const totalHours = Math.floor(val * 24);
                    let totalMins = Math.round((val * 24 * 60) % 60);
                    
                    // Handle edge case where rounding minutes makes it 60
                    let displayHours = totalHours;
                    if (totalMins === 60) {
                        displayHours += 1;
                        totalMins = 0;
                    }
                    val = `${displayHours}:${totalMins.toString().padStart(2, '0')}`;
                } 
                else if (typeof val === 'number') {
                    if (!Number.isInteger(val)) {
                        val = parseFloat(val.toFixed(3));
                    }
                }
                td.textContent = (val !== undefined && val !== null) ? val : '-';
                
                // Styling specific columns
                if (h === 'Kod') {
                    td.style.whiteSpace = 'nowrap';
                    td.style.width = '140px';
                    td.style.minWidth = '140px';
                    td.style.padding = '6px 12px';
                    const isHarici = excludedHariciKodlar.has(code);
                    const hariciIcon = isHarici ? '<i class="fa-solid fa-triangle-exclamation" style="font-size:11px; margin-right:6px; color:#fda4af;" title="Harici İşlem / Harici Kod"></i>' : '';
                    const textColor = isHarici ? '#fda4af' : 'white';
                    
                    td.innerHTML = `
                        <div class="code-cell-wrapper">
                            <span class="code-cell-text" style="color:${textColor};">${hariciIcon}${code}</span>
                            <button type="button" draggable="false" class="part-img-btn" onmousedown="event.stopPropagation()" onclick="event.stopPropagation(); window.openPartImageModal('${code}')" onmouseenter="window.showPartHoverPreview(event, '${code}')" onmousemove="window.movePartHoverPreview(event)" onmouseleave="window.hidePartHoverPreview()" title="Görseli Görüntüle">
                                <i class="fa-solid fa-image"></i>
                            </button>
                        </div>
                    `;
                } else if (h === 'Malzeme Adı') {
                    td.style.color = 'var(--text-muted)';
                } else if (h === 'Hammadde') {
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
bindSearchInput('station-search', () => {
    paginationState.station.page = 1;
    renderStationsTab();
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

document.getElementById('station-prev-btn')?.addEventListener('click', () => {
    if (paginationState.station.page > 1) {
        paginationState.station.page--;
        filterAndPaginateStationData();
    }
});
document.getElementById('station-next-btn')?.addEventListener('click', () => {
    const maxPage = Math.ceil(paginationState.station.total / PAGE_SIZE);
    if (paginationState.station.page < maxPage) {
        paginationState.station.page++;
        filterAndPaginateStationData();
    }
});

// -------------------------------------------------------------
// 6. EXPORT BACK TO EXCEL WORKBOOK (STİLLER VE FORMÜLLER KORUNARAK)
// -------------------------------------------------------------
exportBtn.addEventListener('click', async () => {
    if (!workbook) return;
    
    showToast("Güncel veriler hazırlanıyor (Orijinal kenarlıklar, renkler ve formüller korunuyor)...", "info");
    
    try {
        const origFileName = loadedFileName.textContent.replace("Yüklenen Dosya: ", "").replace(".xlsx", "");
        const exportName = `${origFileName}_Guncel.xlsx`;

        // 1. ÖNCELİKLİ YÖNTEM: ExcelJS ile orijinal şablonun tüm stillerini, renklerini ve formüllerini koruyarak güncelle
        if (rawExcelArrayBuffer && typeof ExcelJS !== 'undefined') {
            const excelJsWorkbook = new ExcelJS.Workbook();
            await excelJsWorkbook.xlsx.load(rawExcelArrayBuffer.slice(0));

            // A. "Üretim Takip" Sayfası Güncellemesi
            const utSheet = excelJsWorkbook.getWorksheet("Üretim Takip");
            if (utSheet) {
                // Sütun L başlığı
                const headerRow = utSheet.getRow(1);
                headerRow.getCell(12).value = 'İSTASYON';

                // Eski üretim loglarını temizle (Sütun I=9, J=10, K=11, L=12)
                const maxRow = Math.max(utSheet.rowCount, uretimTakipRows.length + 50, productionLog.length + 50);
                for (let r = 2; r <= maxRow; r++) {
                    const row = utSheet.getRow(r);
                    if (row.getCell(9).value !== null && row.getCell(9).value !== undefined) {
                        row.getCell(9).value = null;
                        row.getCell(10).value = null;
                        row.getCell(11).value = null;
                        row.getCell(12).value = null;
                    }
                }

                // Yeni üretim kayıtlarını yaz
                productionLog.forEach((log, idx) => {
                    const r = idx + 2; // Satır 2'den itibaren
                    const row = utSheet.getRow(r);
                    row.getCell(9).value = log.kod;
                    row.getCell(10).value = parseFloat(log.adet) || 0;
                    row.getCell(11).value = {
                        formula: `IF(I${r}<>"",MAX(0,SUMIF($I$2:$I$6377,I${r},$J$2:$J$6377)-SUMIF($C$2:$C$6377,I${r},$D$2:$D$6377)),"")`,
                        result: parseFloat(log.fazla) || 0
                    };
                    row.getCell(12).value = log.station || 'Tüm İstasyonlar';
                });

                // Üretim Takip satırlarının değerlerini ve formül sonuçlarını senkronize et
                uretimTakipRows.forEach(row => {
                    const r = row.rowIndex; // 1-tabanlı Excel satırı
                    const exRow = utSheet.getRow(r);
                    
                    const cellE = exRow.getCell(5); // Üretilen Miktar
                    if (cellE.formula) cellE.result = row.uretilen;
                    else cellE.value = row.uretilen;

                    const cellF = exRow.getCell(6); // Kalan Miktar
                    if (cellF.formula) cellF.result = row.kalan;
                    else cellF.value = row.kalan;

                    const cellG = exRow.getCell(7); // Tamamlanma (%)
                    if (cellG.formula) cellG.result = row.tamamlanma;
                    else cellG.value = row.tamamlanma;
                });

                // Dosya Takip (Kaynak Dosya Özeti) sütunları (O=15, P=16, Q=17)
                dosyaTakipRows.forEach(row => {
                    const r = row.rowIndex;
                    const exRow = utSheet.getRow(r);
                    
                    const cellO = exRow.getCell(15);
                    if (cellO.formula) cellO.result = row.hazir;
                    else cellO.value = row.hazir;

                    const cellP = exRow.getCell(16);
                    if (cellP.formula) cellP.result = row.eksik;
                    else cellP.value = row.eksik;

                    const cellQ = exRow.getCell(17);
                    if (cellQ.formula) cellQ.result = row.tamamlanma;
                    else cellQ.value = row.tamamlanma;
                });
            }

            // B. "MONTAJ OTOMASYON İZLEME" Sayfası
            const moSheet = excelJsWorkbook.getWorksheet("MONTAJ OTOMASYON İZLEME");
            if (moSheet) {
                montajOtomasyonLeft.forEach(row => {
                    const exRow = moSheet.getRow(row.rowIndex);
                    const cG = exRow.getCell(7); if (cG.formula) cG.result = row.uretilen; else cG.value = row.uretilen;
                    const cH = exRow.getCell(8); if (cH.formula) cH.result = row.tamamlanma; else cH.value = row.tamamlanma;
                    const cI = exRow.getCell(9); if (cI.formula) cI.result = row.limit; else cI.value = row.limit;
                });
                montajOtomasyonRight.forEach(row => {
                    const exRow = moSheet.getRow(row.rowIndex);
                    const cO = exRow.getCell(15); if (cO.formula) cO.result = row.tamamlananCesit; else cO.value = row.tamamlananCesit;
                    const cP = exRow.getCell(16); if (cP.formula) cP.result = row.tamamlanma; else cP.value = row.tamamlanma;
                    const cR = exRow.getCell(18); if (cR.formula) cR.result = row.limit; else cR.value = row.limit;
                });
            }

            // C. "FINAL MONTAJ İZLEME" Sayfası
            const fmSheet = excelJsWorkbook.getWorksheet("FINAL MONTAJ İZLEME");
            if (fmSheet) {
                finalMontajLeft.forEach(row => {
                    const exRow = fmSheet.getRow(row.rowIndex);
                    const cG = exRow.getCell(7); if (cG.formula) cG.result = row.uretilen; else cG.value = row.uretilen;
                    const cH = exRow.getCell(8); if (cH.formula) cH.result = row.tamamlanma; else cH.value = row.tamamlanma;
                    const cI = exRow.getCell(9); if (cI.formula) cI.result = row.limit; else cI.value = row.limit;
                });
                finalMontajRight.forEach(row => {
                    const exRow = fmSheet.getRow(row.rowIndex);
                    const cO = exRow.getCell(15); if (cO.formula) cO.result = row.tamamlananCesit; else cO.value = row.tamamlananCesit;
                    const cP = exRow.getCell(16); if (cP.formula) cP.result = row.tamamlanma; else cP.value = row.tamamlanma;
                    const cR = exRow.getCell(18); if (cR.formula) cR.result = row.limit; else cR.value = row.limit;
                });
            }

            // D. İstasyon Sayfaları (Kullanıcı miktar değiştirdiyse)
            for (const [stName, rows] of Object.entries(stationSheetsMap)) {
                const wsSt = excelJsWorkbook.getWorksheet(stName);
                if (!wsSt) continue;

                const headerRow = wsSt.getRow(1);
                let uMiktarCol = -1;
                let tHammaddeCol = -1;
                headerRow.eachCell((cell, colNumber) => {
                    const val = String(cell.value || '').trim().toLowerCase();
                    if (val === 'üretilecek miktar') uMiktarCol = colNumber;
                    if (val === 'toplam hammadde miktarı') tHammaddeCol = colNumber;
                });

                rows.forEach(row => {
                    const exRow = wsSt.getRow(row.rowIndex);
                    if (uMiktarCol !== -1) {
                        const cell = exRow.getCell(uMiktarCol);
                        if (cell.formula) cell.result = row['Üretilecek Miktar'] || 0;
                        else cell.value = row['Üretilecek Miktar'] || 0;
                    }
                    if (tHammaddeCol !== -1) {
                        const cell = exRow.getCell(tHammaddeCol);
                        if (cell.formula) cell.result = row['Toplam Hammadde Miktarı'] || 0;
                        else cell.value = row['Toplam Hammadde Miktarı'] || 0;
                    }
                });
            }

            // E. "ÜRETİM LİSTESİ" Sayfası
            const wsUl = excelJsWorkbook.worksheets.find(ws => 
                ws.name.toUpperCase().replace(/I/g, 'İ').includes('ÜRETİM LİSTESİ') || 
                ws.name.toUpperCase().includes('URETIM LISTESI')
            );
            if (wsUl) {
                const headerRow = wsUl.getRow(1);
                let uMiktarCol = -1;
                headerRow.eachCell((cell, colNumber) => {
                    const val = String(cell.value || '').trim().toLowerCase();
                    if (val === 'üretilecek miktar') uMiktarCol = colNumber;
                });
                if (uMiktarCol !== -1) {
                    uretimListesiRows.forEach(row => {
                        const exRow = wsUl.getRow(row.rowIndex);
                        const cell = exRow.getCell(uMiktarCol);
                        if (cell.formula) cell.result = row.uretilecek;
                        else cell.value = row.uretilecek;
                    });
                }
            }

            // ExcelJS dosyayı yazıp indir
            const buffer = await excelJsWorkbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            if (typeof saveAs !== 'undefined') {
                saveAs(blob, exportName);
            } else {
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = exportName;
                link.click();
            }
            showToast(`Orijinal tasarım ve tüm kenarlıklar korunarak indirildi: "${exportName}"`, "success");
            return;
        }

        // 2. YEDEK YÖNTEM (SheetJS Fallback)
        const utSheet = workbook.Sheets["Üretim Takip"];
        const utRange = XLSX.utils.decode_range(utSheet['!ref']);
        for (let r = 1; r <= utRange.e.r; r++) {
            delete utSheet[XLSX.utils.encode_cell({ r: r, c: 8 })];
            delete utSheet[XLSX.utils.encode_cell({ r: r, c: 9 })];
            delete utSheet[XLSX.utils.encode_cell({ r: r, c: 10 })];
            delete utSheet[XLSX.utils.encode_cell({ r: r, c: 11 })];
        }
        utSheet[XLSX.utils.encode_cell({ r: 0, c: 11 })] = { t: 's', v: 'İSTASYON' };

        productionLog.forEach((log, idx) => {
            const r = idx + 1;
            utSheet[XLSX.utils.encode_cell({ r: r, c: 8 })] = { t: 's', v: log.kod };
            utSheet[XLSX.utils.encode_cell({ r: r, c: 9 })] = { t: 'n', v: log.adet };
            utSheet[XLSX.utils.encode_cell({ r: r, c: 10 })] = { 
                t: 'n', 
                v: log.fazla,
                f: `IF(I${r+1}<>"",MAX(0,SUMIF($I$2:$I$6377,I${r+1},$J$2:$J$6377)-SUMIF($C$2:$C$6377,I${r+1},$D$2:$D$6377)),"")`
            };
            utSheet[XLSX.utils.encode_cell({ r: r, c: 11 })] = { t: 's', v: log.station || 'Tüm İstasyonlar' };
        });

        XLSX.writeFile(workbook, exportName);
        showToast(`Güncel Excel dosyası indirildi: "${exportName}"`, "success");
        
    } catch (err) {
        console.error("Excel indirme hatası:", err);
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
    
    // 0. Source File Filter
    if (selectedSourceFiles.size > 0) {
        filtered = filtered.filter(r => {
            const code = String(r.kod || '').trim().toUpperCase();
            const reqs = uretimTakipRows.filter(u => u.kod === code);
            return reqs.some(u => selectedSourceFiles.has(u.kaynak)) || (r.kaynak && selectedSourceFiles.has(r.kaynak));
        });
    }
    
    // 1. Advanced Include/Exclude Search Filter (Google-style)
    if (searchVal) {
        const terms = searchVal.split(/\s+/);
        const includeTerms = [];
        const excludeTerms = [];
        
        terms.forEach(term => {
            if (term.startsWith('-') && term.length > 1) {
                excludeTerms.push(term.slice(1));
            } else if (term) {
                includeTerms.push(term);
            }
        });
        
        filtered = filtered.filter(r => {
            const searchString = [
                r.kod,
                r.malzeme,
                r.hammadde,
                r.kaynak,
                r.hKod
            ].join(' ').toLowerCase();
            
            // All include terms must match
            const matchesInclude = includeTerms.every(term => searchString.includes(term));
            if (!matchesInclude) return false;
            
            // None of the exclude terms must match
            const matchesExclude = excludeTerms.some(term => searchString.includes(term));
            if (matchesExclude) return false;
            
            return true;
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
        if (sortVal === 'diff-desc') {
            const diffA = (a.uretilecek || 0) - (a.orijinalUretilecek || 0);
            const diffB = (b.uretilecek || 0) - (b.orijinalUretilecek || 0);
            return diffB - diffA;
        }
        if (sortVal === 'raw-diff-desc') {
            const rawA = ((a.uretilecek || 0) - (a.orijinalUretilecek || 0)) * (calculateEmpiricalBatchQty(a).unitDim || 0.2);
            const rawB = ((b.uretilecek || 0) - (b.orijinalUretilecek || 0)) * (calculateEmpiricalBatchQty(b).unitDim || 0.2);
            return rawB - rawA;
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

// --- DİNAMİK FORMÜL HARİÇ TUTMA KONTROLÜ ---
function isPartExcludedFromBatchFormula(row) {
    const cleanCode = String(row.kod || '').trim().toUpperCase();
    const hKod = String(row.hKod || '').trim().toUpperCase();
    const hammadde = String(row.hammadde || '').trim().toUpperCase();
    const malzeme = String(row.malzeme || '').trim().toUpperCase();

    const rules = partiFormuluHaricKurallar || {};
    
    // 1. Noktalı Standart Kodlama Kontrolü (Nokta içermeyenler -> Kaynaklı / Alt Montaj)
    if (rules.sadece_noktali_hammadde_gecerli !== false) {
        const isDotted = hKod.includes('.') || hKod.startsWith('150') || hKod.startsWith('152');
        if (!isDotted) {
            return { excluded: true, reason: 'Kaynaklı / Alt Montaj', shortTag: 'Kaynaklı' };
        }
    }

    // 2. Hariç Tutulacak Hammadde Önekleri (Örn: 150.01.01 Lazer Sac)
    if (Array.isArray(rules.haric_hammadde_onekleri)) {
        for (const prefix of rules.haric_hammadde_onekleri) {
            const cleanPrefix = String(prefix || '').trim().toUpperCase();
            if (cleanPrefix && hKod.startsWith(cleanPrefix)) {
                return { excluded: true, reason: `Lazer Kesim Sac (${cleanPrefix})`, shortTag: 'Lazer Sac' };
            }
        }
    }

    // 3. Doğrudan Hariç Tutulacak Hammadde Kodları
    if (Array.isArray(rules.haric_hammadde_kodlari) && rules.haric_hammadde_kodlari.includes(hKod)) {
        return { excluded: true, reason: 'Hariç Tutulan Hammadde', shortTag: 'Hariç' };
    }

    // 4. Doğrudan Hariç Tutulacak Parça Kodları
    if (Array.isArray(rules.haric_parca_kodlari) && rules.haric_parca_kodlari.includes(cleanCode)) {
        return { excluded: true, reason: 'Hariç Tutulan Parça', shortTag: 'Hariç' };
    }

    // 5. Malzeme Adında Geçen Anahtar Kelimeler (Örn: SAC, LAZER vb.)
    if (Array.isArray(rules.haric_malzeme_kelimeleri) && rules.haric_malzeme_kelimeleri.length > 0) {
        for (const kw of rules.haric_malzeme_kelimeleri) {
            const cleanKw = String(kw || '').trim().toUpperCase();
            if (cleanKw && (hammadde.includes(cleanKw) || malzeme.includes(cleanKw))) {
                return { excluded: true, reason: `Malzeme İstisnası (${cleanKw})`, shortTag: cleanKw };
            }
        }
    }

    return { excluded: false, reason: '', shortTag: '' };
}

// --- AKILLI EMPİRİK PARTİ BOYUTLANDIRMA MOTORU (YILLIK PROJEKSİYON & İSTİSNA DESTEKLİ) ---
function calculateEmpiricalBatchQty(row, options) {
    const cleanCode = String(row.kod || '').trim().toUpperCase();
    const origQty = parseFloat(row.orijinalUretilecek) || 1;
    
    // 0. İstisna Kontrolü (Kaynaklı parçalar, 150.01.01 lazer sac vb.)
    const exclusion = isPartExcludedFromBatchFormula(row);
    
    // 1. Birim Hammadde Ölçüsü (metre / kg)
    const rawInfo = parcaBirimHammaddeMap[cleanCode];
    let unitDim = rawInfo ? parseFloat(rawInfo.birimMiktar) || 0 : 0;
    if (unitDim <= 0) unitDim = 0.20; // Varsayılan 200 mm
    
    // 2. Makine Reçete Ortaklığı & Yıllık Projeksiyon Analizi
    const machineRecInfo = parcaMakineReceteleriMap[cleanCode];
    let machineCount = machineRecInfo ? machineRecInfo.makine_sayisi || 1 : 1;
    let totalAnnualDemand = 0;
    
    if (machineRecInfo && machineRecInfo.makineler) {
        Object.entries(machineRecInfo.makineler).forEach(([mName, mQty]) => {
            const mAnnual = makineYillikTahminlerMap[mName] || 50;
            totalAnnualDemand += (parseFloat(mQty) || 1) * mAnnual;
        });
    }

    // 3. Reçete Başı Tüketim (R)
    let recipeUsage = parcaReceteTuketimMap[cleanCode] || (machineRecInfo ? Math.round(machineRecInfo.toplam_birim_adet / Math.max(1, machineCount)) : 1);
    let kRecipe = 1.0;
    if (recipeUsage >= 15) kRecipe = 2.2;
    else if (recipeUsage >= 8) kRecipe = 1.8;
    else if (recipeUsage >= 4) kRecipe = 1.4;
    else if (recipeUsage >= 2) kRecipe = 1.2;

    // 4. Kaynak Dosya / Makine Frekansı Çarpanı
    const sourceCount = Math.max((row.kaynak || '').split(',').length, machineCount);
    let kFreq = 1.0;
    if (sourceCount >= 4) kFreq = 1.6;
    else if (sourceCount >= 2) kFreq = 1.3;

    // 5. Yıllık Talep Gücü Çarpanı (D)
    let kAnnual = 1.0;
    if (totalAnnualDemand >= 800) kAnnual = 2.2;
    else if (totalAnnualDemand >= 300) kAnnual = 1.7;
    else if (totalAnnualDemand >= 100) kAnnual = 1.35;
    else if (totalAnnualDemand >= 40) kAnnual = 1.15;

    // 6. Boy Kuralı Tabanı (B)
    const lengthThreshold = options ? parseFloat(options.lengthThreshold) || 0.50 : 0.50;
    const shortMin = options ? parseFloat(options.shortMin) || 20 : 20;
    const longMultiplier = options ? parseFloat(options.longMultiplier) || 1.30 : 1.30;

    let baseQty = origQty;
    if (unitDim < lengthThreshold) {
        baseQty = Math.max(shortMin, origQty);
    } else {
        baseQty = Math.ceil(origQty * longMultiplier);
    }

    // 0.1 Eğer istisna kapsamındaysa (Lazer sac, Kaynaklı parça vb.) formül uygulanmaz, net ihtiyaç kalır!
    if (exclusion.excluded) {
        return {
            unitDim: unitDim,
            recipeUsage: recipeUsage,
            sourceCount: sourceCount,
            machineCount: machineCount,
            totalAnnualDemand: totalAnnualDemand,
            kRecipe: 1.0,
            kFreq: 1.0,
            kAnnual: 1.0,
            isExcluded: true,
            exclusionReason: exclusion.reason,
            exclusionTag: exclusion.shortTag,
            finalQty: origQty,
            extraQty: 0,
            extraRaw: 0
        };
    }

    // 7. Empirik Toplam Parti Adedi
    const dynamicMultiplier = Math.max(kRecipe * kFreq, kAnnual * (machineCount > 1 ? 1.25 : 1.0));
    const empiricalQty = Math.ceil(origQty * dynamicMultiplier);
    const finalBatchQty = Math.max(baseQty, empiricalQty);

    return {
        unitDim: unitDim,
        recipeUsage: recipeUsage,
        sourceCount: sourceCount,
        machineCount: machineCount,
        totalAnnualDemand: totalAnnualDemand,
        kRecipe: kRecipe,
        kFreq: kFreq,
        kAnnual: kAnnual,
        isExcluded: false,
        exclusionReason: '',
        exclusionTag: '',
        finalQty: finalBatchQty,
        extraQty: Math.max(0, finalBatchQty - origQty),
        extraRaw: Math.round((Math.max(0, finalBatchQty - origQty) * unitDim) * 100) / 100
    };
}

function applySmartBatchRules() {
    const lengthThreshold = parseFloat(document.getElementById('batch-length-threshold').value) || 0.50;
    const shortMin = parseFloat(document.getElementById('batch-short-min').value) || 20;
    const longMultiplier = parseFloat(document.getElementById('batch-long-multiplier').value) || 1.30;

    const options = { lengthThreshold, shortMin, longMultiplier };
    let changedCount = 0;
    let totalExtraRaw = 0;

    uretimListesiRows.forEach(row => {
        const calc = calculateEmpiricalBatchQty(row, options);
        if (calc.finalQty !== row.uretilecek) {
            row.uretilecek = calc.finalQty;
            uretimListesiMap[row.kod] = calc.finalQty;
            changedCount++;
            totalExtraRaw += calc.extraRaw;
        }
    });

    recalculateAll();
    filterAndPaginateUlTable();

    showToast(`⚡ Akıllı parti kuralları ${changedCount} parçaya uygulandı! (+${Math.round(totalExtraRaw)} m/kg hammadde)`, "success");
}
window.applySmartBatchRules = applySmartBatchRules;

function resetUlToOriginal() {
    if (confirm("Tüm parçaları orijinal reçete üretim miktarlarına geri döndürmek istediğinize emin misiniz?")) {
        uretimListesiRows.forEach(row => {
            row.uretilecek = row.orijinalUretilecek;
            uretimListesiMap[row.kod] = row.orijinalUretilecek;
        });

        recalculateAll();
        filterAndPaginateUlTable();
        showToast("Tüm parçalar orijinal ihtiyaç miktarlarına sıfırlandı.", "info");
    }
}
window.resetUlToOriginal = resetUlToOriginal;

function renderUlTable() {
    const tbody = document.getElementById('ul-table-body');
    tbody.innerHTML = '';
    
    const countBadge = document.getElementById('ul-total-count');
    if (countBadge) countBadge.textContent = `${uretimListesiRows.length} Kalem`;

    const pState = paginationState.ul;

    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="text-center" style="color:var(--text-dim); padding:20px;">Eşleşen parça bulunamadı.</td></tr>';
        return;
    }

    const pageRows = pState.filtered;

    pageRows.forEach(row => {
        const tr = document.createElement('tr');
        
        // Empirik parti ve hammadde analizini hesapla
        const calc = calculateEmpiricalBatchQty(row);
        const unitDimFormatted = calc.unitDim < 1 ? `${Math.round(calc.unitDim * 1000)} mm` : `${calc.unitDim.toFixed(2)} m`;

        // Check if quantity has been modified
        const isModified = row.uretilecek !== row.orijinalUretilecek;
        let changeBadgeHtml = '';
        let inputStyle = 'width: 86px; background: rgba(0,0,0,0.35); border: 1px solid var(--border-color); color: white; border-radius: 6px; padding: 5px 8px; font-weight: 700; font-size: 13px; outline: none; transition: var(--transition);';
        
        const currentDiff = row.uretilecek - row.orijinalUretilecek;
        const currentExtraRaw = Math.round((currentDiff * calc.unitDim) * 100) / 100;

        if (isModified) {
            const diffText = currentDiff > 0 ? `+${currentDiff} Adet` : `${currentDiff} Adet`;
            const badgeColor = currentDiff > 0 ? '#34d399' : '#f87171';
            const badgeBg = currentDiff > 0 ? 'rgba(16, 185, 129, 0.15)' : 'rgba(239, 68, 68, 0.15)';
            const borderGlow = currentDiff > 0 ? 'rgba(52, 211, 153, 0.3)' : 'rgba(239, 68, 68, 0.3)';
            
            inputStyle = `width: 86px; background: rgba(16, 22, 40, 0.95); border: 1.5px solid ${badgeColor}; color: ${badgeColor}; box-shadow: 0 0 10px ${borderGlow}; border-radius: 6px; padding: 5px 8px; font-weight: 800; font-size: 13px; outline: none; transition: var(--transition);`;
            
            changeBadgeHtml = `
                <div style="font-size: 11px; margin-top: 6px; display: flex; flex-direction: column; align-items: flex-end; gap: 3px; font-weight: 600;">
                    <div style="display: flex; align-items: center; gap: 6px; white-space: nowrap;">
                        <span style="color: rgba(255, 255, 255, 0.6); font-size: 10.5px;">Orijinal: <strong style="color: white;">${row.orijinalUretilecek}</strong></span>
                        <span class="badge" style="background: ${badgeBg}; color: ${badgeColor}; border: 1px solid ${badgeColor}40; padding: 2px 6px; font-weight: 700; font-size: 10px; border-radius: 4px;">
                            ${diffText}
                        </span>
                    </div>
                    ${currentExtraRaw > 0 ? `
                    <div style="font-size: 10px; color: #38bdf8; display: flex; align-items: center; gap: 4px; background: rgba(56, 189, 248, 0.12); padding: 1px 6px; border-radius: 3px; border: 1px solid rgba(56, 189, 248, 0.25);">
                        <i class="fa-solid fa-cube" style="font-size: 9px;"></i> Ekstra: +${currentExtraRaw} m
                    </div>` : ''}
                </div>
            `;
        }

        // Tüketim, Makine Ortaklığı & Yıllık Projeksiyon Rozeti
        let usageBadge = '';
        if (calc.isExcluded) {
            const badgeIcon = calc.exclusionTag === 'Lazer Sac' ? 'fa-bolt text-yellow' : 'fa-layer-group';
            const badgeColor = calc.exclusionTag === 'Lazer Sac' ? 'background:rgba(234,179,8,0.12); color:#fde047; border:1px solid rgba(234,179,8,0.3);' : 'background:rgba(148,163,184,0.12); color:#94a3b8; border:1px solid rgba(148,163,184,0.25);';
            usageBadge = `
                <div style="display: flex; flex-direction: column; gap: 3px; font-size: 11px;">
                    <span class="badge" style="${badgeColor} padding: 2px 7px; font-weight: 700; font-size: 10px; border-radius: 4px; display: inline-flex; align-items: center; gap: 4px; width: fit-content;">
                        <i class="fa-solid ${badgeIcon}"></i> ${calc.exclusionReason}
                    </span>
                    <span style="font-size:10px; color:var(--text-dim);"><i class="fa-solid fa-shield"></i> Net İhtiyaç (Formül Dışı)</span>
                </div>
            `;
        } else {
            const isCommon = calc.machineCount >= 2;
            const commonBadgeClass = calc.machineCount >= 4 ? 'background:rgba(239,68,68,0.15); color:#fca5a5; border:1px solid rgba(239,68,68,0.3);' : (isCommon ? 'background:rgba(245,158,11,0.15); color:#fcd34d; border:1px solid rgba(245,158,11,0.3);' : 'background:rgba(255,255,255,0.05); color:var(--text-muted); border:1px solid rgba(255,255,255,0.08);');
            
            usageBadge = `
                <div style="display: flex; flex-direction: column; gap: 4px; font-size: 11px;">
                    <span class="badge" style="${commonBadgeClass} padding: 2px 7px; font-weight: 700; font-size: 10.5px; border-radius: 4px; display: inline-flex; align-items: center; gap: 4px; width: fit-content;">
                        <i class="fa-solid ${calc.machineCount >= 4 ? 'fa-fire text-orange' : (isCommon ? 'fa-diagram-project' : 'fa-cube')}"></i> ${calc.machineCount >= 2 ? calc.machineCount + ' Makinede Ortak' : 'Özel Parça (Tek Makine)'}
                    </span>
                    <div style="display:flex; align-items:center; gap:8px; font-size:10px; color:var(--text-dim);">
                        <span><i class="fa-solid fa-wrench" style="opacity:0.6;"></i> Reçete: <b>${calc.recipeUsage} ad</b></span>
                        ${calc.totalAnnualDemand > 0 ? `<span style="color:#38bdf8;"><i class="fa-solid fa-chart-line"></i> Yıllık: <b>~${Math.round(calc.totalAnnualDemand)} ad</b></span>` : ''}
                    </div>
                </div>
            `;
        }

        tr.innerHTML = `
            <td style="font-size:12px; color:var(--text-muted); max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${row.kaynak}">${row.kaynak}</td>
            <td>${row.oncelik}</td>
            <td style="white-space:nowrap; width:140px; min-width:140px; padding:6px 12px;">
                <div class="code-cell-wrapper">
                    <span class="code-cell-text" style="color:${excludedHariciKodlar.has(row.kod.trim().toUpperCase()) ? '#fda4af' : 'white'};">
                        ${excludedHariciKodlar.has(row.kod.trim().toUpperCase()) ? '<i class="fa-solid fa-triangle-exclamation" style="font-size:11px; margin-right:6px; opacity:0.9;" title="Harici İşlem / Harici Kod"></i>' : ''}
                        ${row.kod}
                    </span>
                    <button type="button" draggable="false" class="part-img-btn" onmousedown="event.stopPropagation()" onclick="event.stopPropagation(); window.openPartImageModal('${row.kod}')" onmouseenter="window.showPartHoverPreview(event, '${row.kod}')" onmousemove="window.movePartHoverPreview(event)" onmouseleave="window.hidePartHoverPreview()" title="Görseli Görüntüle">
                        <i class="fa-solid fa-image"></i>
                    </button>
                </div>
            </td>
            <td style="color:var(--text-dim); max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${row.malzeme}">${row.malzeme}</td>
            <td style="white-space:nowrap;">
                <span class="badge" style="background: rgba(56, 189, 248, 0.12); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.25); font-weight: 700; font-size: 11px;">
                    ${unitDimFormatted}
                </span>
            </td>
            <td>${usageBadge}</td>
            <td style="color:var(--text-dim); font-size: 11.5px; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${row.hammadde}">${row.hammadde || '-'}</td>
            <td class="text-right">
                <input type="number" class="table-input ul-qty-input text-right" 
                       value="${row.uretilecek}" 
                       data-kod="${row.kod}" 
                       min="0"
                       style="${inputStyle}">
                ${changeBadgeHtml}
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
            
            // 4. Re-render list to show badge and color updates
            filterAndPaginateUlTable();
            
            showToast(`${kod} için yeni parti üretim miktarı belirlendi: ${newQty}`, "success");
        });
        
        input.addEventListener('focus', function() {
            if (!isModified) {
                this.style.borderColor = 'var(--primary)';
                this.style.boxShadow = '0 0 8px var(--primary-glow)';
                this.style.background = 'rgba(16, 22, 40, 0.8)';
            }
        });
        input.addEventListener('blur', function() {
            if (!isModified) {
                this.style.borderColor = 'var(--border-color)';
                this.style.boxShadow = 'none';
                this.style.background = 'rgba(0,0,0,0.3)';
            }
        });

        tbody.appendChild(tr);
    });
}

// Üretim Listesi navigation events
bindSearchInput('ul-search', () => {
    paginationState.ul.page = 1;
    filterAndPaginateUlTable();
});

document.getElementById('ul-sort').addEventListener('change', () => {
    paginationState.ul.page = 1;
    filterAndPaginateUlTable();
});

document.getElementById('ul-prev-btn')?.addEventListener('click', () => {
    if (paginationState.ul.page > 1) {
        paginationState.ul.page--;
        renderUlTable();
    }
});

document.getElementById('ul-next-btn')?.addEventListener('click', () => {
    const maxPage = Math.ceil(paginationState.ul.total / PAGE_SIZE);
    if (paginationState.ul.page < maxPage) {
        paginationState.ul.page++;
        renderUlTable();
    }
});

// --- EXCEL EXPORT USING EXCELJS ---
async function exportStationDataToExcel() {
    console.log("İstasyon Excel İndirme başlatıldı...");
    if (!paginationState.station.filtered || paginationState.station.filtered.length === 0) {
        alert("Dışa aktarılacak veri bulunamadı.");
        return;
    }

    try {
        const stationName = activeStation || "Tüm_İstasyonlar";
        const dateStr = new Date().toLocaleDateString('tr-TR').replace(/\./g, '-');
        const fileName = `${stationName.replace(/\s+/g, '_')}_İş_Listesi_${dateStr}.xlsx`;

        // Create a new ExcelJS workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "Antigravity Engine";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet(stationName.substring(0, 31)); // Excel sheet names max 31 chars

        // Get headers from the table currently displayed
        const thead = document.getElementById('station-thead');
        const thElements = thead.querySelectorAll('th');
        const headerNames = Array.from(thElements).map(th => th.textContent.trim());

        // Add headers
        const headerRow = worksheet.addRow(headerNames);
        headerRow.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        headerRow.height = 30;
        headerRow.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } }; // Dark gray
            cell.border = {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            };
        });

        // Sütun indekslerini önceden belirle (formüller için sütun harfi lazım)
        // headerNames is 0-indexed; Excel columns are 1-indexed
        function colLetter(idx1based) {
            // idx1based: 1 = A, 2 = B, etc.
            let col = '';
            let n = idx1based;
            while (n > 0) {
                const rem = (n - 1) % 26;
                col = String.fromCharCode(65 + rem) + col;
                n = Math.floor((n - 1) / 26);
            }
            return col;
        }
        
        const colIdxOf = {}; // header name -> 1-based column index in Excel
        headerNames.forEach((h, i) => { colIdxOf[h] = i + 1; });
        
        const hazirlikCol    = colIdxOf['Hazırlık Süresi']    ? colLetter(colIdxOf['Hazırlık Süresi'])    : null;
        const birimCol       = colIdxOf['Birim İşlem Süresi'] ? colLetter(colIdxOf['Birim İşlem Süresi']) : null;
        const miktarCol      = colIdxOf['Üretilecek Miktar']  ? colLetter(colIdxOf['Üretilecek Miktar'])  : null;
        const toplamSureCol  = colIdxOf['Toplam Süre']        ? colLetter(colIdxOf['Toplam Süre'])        : null;
        const saatCol        = colIdxOf['Saat']               ? colLetter(colIdxOf['Saat'])               : null;
        const kumulatifCol   = colIdxOf['Kümülatif Süre']     ? colLetter(colIdxOf['Kümülatif Süre'])     : null;
        const hammaddeMCol   = colIdxOf['Hammadde Miktar']    ? colLetter(colIdxOf['Hammadde Miktar'])    : null;
        const toplamHammCol  = colIdxOf['Toplam Hammadde Miktarı'] ? colLetter(colIdxOf['Toplam Hammadde Miktarı']) : null;

        // Add data rows
        let excelRowNum = 2; // row 1 is header
        paginationState.station.filtered.forEach(row => {
            const code = String(row['Kod'] || '').trim().toUpperCase();
            
            // Re-calculate status
            const reqs = uretimTakipRows.filter(u => u.kod === code);
            let completionText = 'Eksik';
            let completionPct = 0;
            
            if (reqs.length > 0) {
                const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
                const stLogs = productionLog.filter(log => log.kod === code && (log.station === activeStation || log.station === 'Tüm İstasyonlar' || !log.station));
                const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
                completionPct = totalReq > 0 ? Math.min(100, Math.round((stProd / totalReq) * 100)) : 0;
                
                if (completionPct >= 100) {
                    completionText = `Hazır (${completionPct}%)`;
                } else if (completionPct > 0) {
                    completionText = `Üretimde (${completionPct}%)`;
                } else {
                    completionText = `Eksik (0%)`;
                }
            }

            const rowData = [];
            headerNames.forEach(header => {
                if (header === 'Durum') {
                    rowData.push(completionText);
                } else if (header === 'Toplam Süre' && hazirlikCol && birimCol && miktarCol) {
                    // Canlı formül: Hazırlık + (Birim * Adet)
                    rowData.push({ formula: `=${hazirlikCol}${excelRowNum}+(${birimCol}${excelRowNum}*${miktarCol}${excelRowNum})` });
                } else if (header === 'Saat' && toplamSureCol) {
                    // Canlı formül: Toplam Süre / 86400
                    rowData.push({ formula: `=${toplamSureCol}${excelRowNum}/86400` });
                } else if (header === 'Kümülatif Süre' && saatCol) {
                    // Canlı formül: kümülatif toplam
                    if (excelRowNum === 2) {
                        rowData.push({ formula: `=${saatCol}2` });
                    } else {
                        rowData.push({ formula: `=${kumulatifCol}${excelRowNum - 1}+${saatCol}${excelRowNum}` });
                    }
                } else if (header === 'Toplam Hammadde Miktarı' && miktarCol && hammaddeMCol) {
                    // Canlı formül: Adet * Hammadde Miktar
                    rowData.push({ formula: `=${miktarCol}${excelRowNum}*${hammaddeMCol}${excelRowNum}` });
                } else {
                    let val = row[header];
                    if (typeof val === 'number' && !Number.isInteger(val)) {
                        val = parseFloat(val.toFixed(3));
                    }
                    rowData.push(val !== undefined && val !== null ? val : '-');
                }
            });

            const xlRow = worksheet.addRow(rowData);
            xlRow.alignment = { vertical: 'middle', wrapText: true };
            
            xlRow.eachCell((cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFDDDDDD' } },
                    left: { style: 'thin', color: { argb: 'FFDDDDDD' } },
                    bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } },
                    right: { style: 'thin', color: { argb: 'FFDDDDDD' } }
                };
                
                const header = headerNames[colNumber - 1];
                
                // Saat ve Kümülatif Süre hücrelerine [h]:mm formatı uygula
                if (header === 'Saat' || header === 'Kümülatif Süre') {
                    cell.numFmt = '[h]:mm';
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                }
                
                if (header === 'Durum') {
                    const statusStr = String(cell.value).toLowerCase();
                    if (statusStr.includes('hazır')) {
                        cell.font = { color: { argb: 'FF10B981' }, bold: true }; // Green
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };
                    } else if (statusStr.includes('üretimde')) {
                        cell.font = { color: { argb: 'FFF59E0B' }, bold: true }; // Orange
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFBEB' } };
                    } else if (statusStr.includes('eksik')) {
                        cell.font = { color: { argb: 'FFEF4444' }, bold: true }; // Red
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF2F2' } };
                    }
                } else if (header === 'Kod') {
                    cell.font = { bold: true };
                }
                
                // Color entire row purple if it's a harici (excluded) code, except for Status column
                if (excludedHariciKodlar.has(code) && header !== 'Durum') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2EBF9' } };
                    if (header === 'Kod') {
                        cell.font = { color: { argb: 'FF6C3483' }, bold: true }; // Purple font
                    } else {
                        cell.font = { color: { argb: 'FF6C3483' } }; // Purple font
                    }
                }
            });
            
            excelRowNum++;
        });

        // Auto-fit columns
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                let columnLength = cell.value ? cell.value.toString().length : 10;
                if (columnLength > maxLength) maxLength = columnLength;
            });
            column.width = Math.min(maxLength + 2, 40); // clamp max width
        });

        // Generate and save file
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast(`İstasyon iş listesi indirildi: "${fileName}"`, "success");
    } catch (err) {
        console.error("Excel dışa aktarma hatası:", err);
        alert("Excel oluşturulurken bir hata oluştu: " + err.message);
    }
}


// -------------------------------------------------------------
// X. PERFORMANCE TAB RENDERING
// -------------------------------------------------------------
const DEFAULT_CAPACITY = {
    "varsayilan_gunluk_saat": 9,
    "istasyonlar": {
        "CY": { "gunluk_saat": 9, "makine_sayisi": 3 },
        "QUASER": { "gunluk_saat": 9, "makine_sayisi": 3 },
        "3D PRINT": { "gunluk_saat": 22, "makine_sayisi": 15 },
        "3D YAZICI GRUBU": { "gunluk_saat": 22, "makine_sayisi": 15 }
    }
};

window.selectStationAndGo = function(stName) {
    activeStation = stName;
    switchTab('stations');
};

// -------------------------------------------------------------
// WORKLOAD & CAPACITY COMPARISON TAB
// -------------------------------------------------------------
function renderWorkloadTab() {
    const kpiContainer = document.getElementById('workload-kpi-container');
    const barsContainer = document.getElementById('workload-bars-container');
    if (!barsContainer) return;

    if (!stationList || stationList.length === 0) {
        if (kpiContainer) kpiContainer.innerHTML = '';
        barsContainer.innerHTML = '<div style="color: var(--text-dim); text-align: center; padding: 40px;">Henüz istasyon verisi yüklenmedi.</div>';
        return;
    }

    // 1. Tüm istasyonların iş yükü verilerini hesapla
    const data = [];
    let grandTotalHours = 0;

    stationList.forEach(stName => {
        const rows = stationSheetsMap[stName] || [];
        if (rows.length === 0) return;

        let gunlukSaat = DEFAULT_CAPACITY.varsayilan_gunluk_saat || 9;
        let makineSayisi = 1;
        if (DEFAULT_CAPACITY.istasyonlar) {
            if (DEFAULT_CAPACITY.istasyonlar[stName]) {
                gunlukSaat = DEFAULT_CAPACITY.istasyonlar[stName].gunluk_saat || gunlukSaat;
                makineSayisi = DEFAULT_CAPACITY.istasyonlar[stName].makine_sayisi || makineSayisi;
            } else {
                const normSt = stName.trim().toUpperCase().replace(/\s+/g, '');
                for (const [k, cfg] of Object.entries(DEFAULT_CAPACITY.istasyonlar)) {
                    if (k.trim().toUpperCase().replace(/\s+/g, '') === normSt) {
                        gunlukSaat = cfg.gunluk_saat || gunlukSaat;
                        makineSayisi = cfg.makine_sayisi || makineSayisi;
                        break;
                    }
                }
            }
        }

        let stHours = 0;
        rows.forEach(r => {
            let saatVal = parseFloat(r['Saat']) || 0;
            if (saatVal > 0) {
                stHours += (saatVal * 24);
            } else {
                let topSure = parseFloat(r['Toplam Süre']) || 0;
                stHours += (topSure / 3600);
            }
        });

        const gunlukKapasite = gunlukSaat * makineSayisi;
        const tahminiGun = gunlukKapasite > 0 ? (stHours / gunlukKapasite) : 0;
        grandTotalHours += stHours;

        data.push({
            name: stName,
            partsCount: rows.length,
            gunlukSaat: gunlukSaat,
            makineSayisi: makineSayisi,
            gunlukKapasite: gunlukKapasite,
            totalHours: stHours,
            tahminiGun: tahminiGun
        });
    });

    if (data.length === 0) {
        barsContainer.innerHTML = '<div style="color: var(--text-dim); text-align: center; padding: 40px;">İstasyonlarda iş listesi bulunamadı.</div>';
        return;
    }

    // 2. Sıralama
    const sortSelect = document.getElementById('workload-sort-select');
    const sortMode = sortSelect ? sortSelect.value : 'days-desc';

    if (sortMode === 'days-desc') {
        data.sort((a, b) => b.tahminiGun - a.tahminiGun);
    } else if (sortMode === 'days-asc') {
        data.sort((a, b) => a.tahminiGun - b.tahminiGun);
    } else if (sortMode === 'name-asc') {
        data.sort((a, b) => a.name.localeCompare(b.name, 'tr', { sensitivity: 'base' }));
    }

    // 3. KPI Değerleri
    const maxItem = [...data].sort((a, b) => b.tahminiGun - a.tahminiGun)[0];
    const minItem = [...data].sort((a, b) => a.tahminiGun - b.tahminiGun)[0];
    const grandTotalDays = (grandTotalHours / (9 * 1)).toFixed(1); // referans 9h tek istasyon eşdeğeri

    if (kpiContainer) {
        kpiContainer.innerHTML = `
            <div class="kpi-card glass" style="padding: 10px 14px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.08); display: flex; align-items: center; justify-content: space-between; min-width: 0; gap: 10px;">
                <div style="display: flex; align-items: center; gap: 10px; min-width: 0;">
                    <div style="background: rgba(59,130,246,0.15); width: 34px; height: 34px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                        <i class="fa-solid fa-industry text-blue" style="font-size: 14px;"></i>
                    </div>
                    <div style="min-width: 0;">
                        <div style="font-size: 10.5px; color: var(--text-dim); text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Aktif İstasyon</div>
                        <div style="font-size: 17px; font-weight: 800; color: white; line-height: 1.2;">${data.length} <span style="font-size: 11.5px; font-weight: 500; color: var(--text-dim);">İstasyon</span></div>
                    </div>
                </div>
                <div style="font-size: 11px; color: #38bdf8; white-space: nowrap; background: rgba(59,130,246,0.1); padding: 2px 7px; border-radius: 4px;">
                    <i class="fa-solid fa-circle-check" style="font-size: 9px;"></i> Takipte
                </div>
            </div>

            <div class="kpi-card glass" style="padding: 10px 14px; border-radius: 10px; border: 1px solid rgba(255,255,255,0.08); display: flex; align-items: center; justify-content: space-between; min-width: 0; gap: 10px;">
                <div style="display: flex; align-items: center; gap: 10px; min-width: 0;">
                    <div style="background: rgba(245,158,11,0.15); width: 34px; height: 34px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                        <i class="fa-solid fa-clock text-yellow" style="font-size: 14px;"></i>
                    </div>
                    <div style="min-width: 0;">
                        <div style="font-size: 10.5px; color: var(--text-dim); text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Toplam İş Yükü</div>
                        <div style="font-size: 17px; font-weight: 800; color: white; line-height: 1.2;">${grandTotalHours.toFixed(1)} <span style="font-size: 11.5px; font-weight: 500; color: var(--text-dim);">Saat</span></div>
                    </div>
                </div>
                <div style="font-size: 11px; color: #fbbf24; white-space: nowrap; background: rgba(245,158,11,0.1); padding: 2px 7px; border-radius: 4px;">
                    Toplam Plan
                </div>
            </div>

            <div class="kpi-card glass" style="padding: 10px 14px; border-radius: 10px; border: 1px solid rgba(239,68,68,0.3); background: rgba(239,68,68,0.05); display: flex; align-items: center; justify-content: space-between; min-width: 0; gap: 10px;">
                <div style="display: flex; align-items: center; gap: 10px; min-width: 0; overflow: hidden;">
                    <div style="background: rgba(239,68,68,0.2); width: 34px; height: 34px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                        <i class="fa-solid fa-triangle-exclamation" style="color: #ef4444; font-size: 14px;"></i>
                    </div>
                    <div style="min-width: 0; overflow: hidden;">
                        <div style="font-size: 10.5px; color: #f87171; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">🚨 En Yoğun</div>
                        <div style="font-size: 14px; font-weight: 800; color: #fca5a5; line-height: 1.2; text-overflow: ellipsis; overflow: hidden; white-space: nowrap;" title="${maxItem ? maxItem.name : ''}">
                            ${maxItem ? maxItem.name : '-'}
                        </div>
                    </div>
                </div>
                <div style="text-align: right; white-space: nowrap; flex-shrink: 0;">
                    <div style="font-size: 13.5px; font-weight: 800; color: #f87171;">${maxItem ? maxItem.tahminiGun.toFixed(1) + ' Gün' : '-'}</div>
                    <div style="font-size: 10.5px; color: #fca5a5; opacity: 0.8;">${maxItem ? maxItem.totalHours.toFixed(1) + 's' : ''}</div>
                </div>
            </div>

            <div class="kpi-card glass" style="padding: 10px 14px; border-radius: 10px; border: 1px solid rgba(16,185,129,0.3); background: rgba(16,185,129,0.05); display: flex; align-items: center; justify-content: space-between; min-width: 0; gap: 10px;">
                <div style="display: flex; align-items: center; gap: 10px; min-width: 0; overflow: hidden;">
                    <div style="background: rgba(16,185,129,0.2); width: 34px; height: 34px; border-radius: 8px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                        <i class="fa-solid fa-feather" style="color: #10b981; font-size: 14px;"></i>
                    </div>
                    <div style="min-width: 0; overflow: hidden;">
                        <div style="font-size: 10.5px; color: #34d399; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">🟢 En Müsait</div>
                        <div style="font-size: 14px; font-weight: 800; color: #6ee7b7; line-height: 1.2; text-overflow: ellipsis; overflow: hidden; white-space: nowrap;" title="${minItem ? minItem.name : ''}">
                            ${minItem ? minItem.name : '-'}
                        </div>
                    </div>
                </div>
                <div style="text-align: right; white-space: nowrap; flex-shrink: 0;">
                    <div style="font-size: 13.5px; font-weight: 800; color: #34d399;">${minItem ? minItem.tahminiGun.toFixed(1) + ' Gün' : '-'}</div>
                    <div style="font-size: 10.5px; color: #6ee7b7; opacity: 0.8;">${minItem ? minItem.totalHours.toFixed(1) + 's' : ''}</div>
                </div>
            </div>
        `;
    }

    // 4. Barları Çiz (Kompakt 2 Kolonlu Kart Düzeni)
    const maxDays = Math.max(...data.map(d => d.tahminiGun), 1);
    let html = '';

    data.forEach(d => {
        const barPct = Math.min(100, Math.max(4, (d.tahminiGun / maxDays) * 100));

        let colorGradient = 'linear-gradient(90deg, #10b981, #059669)';
        let badgeHtml = '<span class="badge" style="background:rgba(16,185,129,0.15); color:#34d399; border:1px solid rgba(16,185,129,0.3); font-size:10.5px; padding:2px 7px;"><i class="fa-solid fa-feather"></i> Müsait</span>';
        let dayColor = '#34d399';

        if (d.tahminiGun >= 10) {
            colorGradient = 'linear-gradient(90deg, #ef4444, #b91c1c)';
            badgeHtml = '<span class="badge" style="background:rgba(239,68,68,0.2); color:#f87171; border:1px solid rgba(239,68,68,0.4); font-size:10.5px; padding:2px 7px; animation: pulseExcelBtn 2s infinite;"><i class="fa-solid fa-triangle-exclamation"></i> Kritik</span>';
            dayColor = '#f87171';
        } else if (d.tahminiGun >= 5) {
            colorGradient = 'linear-gradient(90deg, #f59e0b, #d97706)';
            badgeHtml = '<span class="badge" style="background:rgba(245,158,11,0.2); color:#fbbf24; border:1px solid rgba(245,158,11,0.4); font-size:10.5px; padding:2px 7px;"><i class="fa-solid fa-fire"></i> Yüksek Yük</span>';
            dayColor = '#fbbf24';
        } else if (d.tahminiGun >= 2) {
            colorGradient = 'linear-gradient(90deg, #3b82f6, #1d4ed8)';
            badgeHtml = '<span class="badge" style="background:rgba(59,130,246,0.18); color:#60a5fa; border:1px solid rgba(59,130,246,0.35); font-size:10.5px; padding:2px 7px;"><i class="fa-solid fa-check"></i> Dengeli</span>';
            dayColor = '#60a5fa';
        }

        html += `
            <div class="workload-item" onclick="selectStationAndGo('${d.name}')" style="background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.06); border-radius: 10px; padding: 11px 14px; cursor: pointer; transition: all 0.2s; display: flex; flex-direction: column; gap: 8px;" onmouseover="this.style.background='rgba(255,255,255,0.05)'; this.style.borderColor='rgba(59,130,246,0.4)';" onmouseout="this.style.background='rgba(255,255,255,0.02)'; this.style.borderColor='rgba(255,255,255,0.06)';">
                <!-- Üst Satır: İstasyon adı, parça, gün ve süre -->
                <div style="display: flex; justify-content: space-between; align-items: center; gap: 8px;">
                    <div style="display: flex; align-items: center; gap: 8px; min-width: 0; overflow: hidden;">
                        <span style="font-size: 14px; font-weight: 700; color: white; display: flex; align-items: center; gap: 6px; white-space: nowrap;">
                            <i class="fa-solid fa-industry text-blue" style="font-size: 12px;"></i> ${d.name}
                        </span>
                        <span style="font-size: 11px; color: var(--text-dim); background: rgba(255,255,255,0.05); padding: 1px 6px; border-radius: 4px; white-space: nowrap;">
                            ${d.partsCount} Parça
                        </span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 8px; flex-shrink: 0;">
                        ${badgeHtml}
                        <span style="font-size: 14px; font-weight: 800; color: ${dayColor};">
                            ${d.tahminiGun.toFixed(1)} Gün
                        </span>
                        <span style="font-size: 11.5px; color: var(--text-dim);">
                            (${d.totalHours.toFixed(1)}s)
                        </span>
                        <i class="fa-solid fa-chevron-right" style="font-size: 10px; color: var(--text-dim);"></i>
                    </div>
                </div>
                
                <!-- Alt Satır: Kompakt İlerleme Çubuğu ve Makine Detayı -->
                <div style="display: flex; align-items: center; gap: 10px;">
                    <div style="flex: 1; height: 6px; background: rgba(255,255,255,0.06); border-radius: 3px; overflow: hidden; position: relative;">
                        <div style="width: ${barPct}%; height: 100%; background: ${colorGradient}; border-radius: 3px; transition: width 0.5s ease;"></div>
                    </div>
                    <span style="font-size: 10.5px; color: var(--text-dim); white-space: nowrap; flex-shrink: 0;">
                        ⚙️ ${d.makineSayisi} Mak • ⏱️ ${d.gunlukSaat}s/g
                    </span>
                </div>
            </div>
        `;
    });

    barsContainer.innerHTML = html;
}

function renderPerformanceTab() {
    const container = document.getElementById('performance-cards-container');
    if (!container) return;
    container.innerHTML = '';
    
    if (stationList.length === 0) {
        container.innerHTML = '<div style="color: var(--text-dim); width: 100%; text-align: center;">Excel verisi yüklenmedi.</div>';
        return;
    }
    
    // Varsayılan çalışma günü sayısı (Kullanıcı 5 güne düşürdü)
    const HAFTALIK_CALISMA_GUNU = 5;
    
    stationList.forEach(stName => {
        // Kapasite bul
        let gunlukSaat = DEFAULT_CAPACITY.varsayilan_gunluk_saat;
        let makineSayisi = 1;
        
        if (DEFAULT_CAPACITY.istasyonlar[stName]) {
            gunlukSaat = DEFAULT_CAPACITY.istasyonlar[stName].gunluk_saat || gunlukSaat;
            makineSayisi = DEFAULT_CAPACITY.istasyonlar[stName].makine_sayisi || makineSayisi;
        } else {
            // Eğer isminde Quaser, 3D vs varsa yakalamaya çalış
            const upSt = stName.toUpperCase();
            if (upSt.includes("QUASER")) { gunlukSaat = 9; makineSayisi = 3; }
            else if (upSt.includes("CY")) { gunlukSaat = 9; makineSayisi = 3; }
            else if (upSt.includes("3D")) { gunlukSaat = 22; makineSayisi = 15; }
        }
        
        const downtimeHours = downtimeMap[stName] || 0;
        // Toplam kapasiteden duruşu (dk cinsinden) düşüyoruz
        const netKapasiteDk = Math.max(0, (gunlukSaat * makineSayisi * HAFTALIK_CALISMA_GUNU * 60) - (downtimeHours * 60));
        const haftalikKapasiteDk = netKapasiteDk;
        
        // Üretilen iş (Standart Saat) hesapla
        let uretilenStandartDk = 0;
        
        // Bu istasyona giren parçaları bul
        const stRows = stationSheetsMap[stName] || [];
        const uniqueCodes = new Set(stRows.map(r => String(r.Kod || '').trim().toUpperCase()).filter(k => k));
        
        uniqueCodes.forEach(code => {
            const logs = productionLog.filter(l => l.kod === code && (l.station === stName || l.station === 'Tüm İstasyonlar'));
            const stProd = logs.reduce((sum, l) => sum + l.adet, 0);
            
            let unitTime = 0;
            let setupTime = 0;
            if (unitTimeMap[code] && unitTimeMap[code][stName]) {
                unitTime = unitTimeMap[code][stName].sure || 0;
                setupTime = unitTimeMap[code][stName].setup || 0;
            }
            
            if (stProd > 0) {
                // Eğer hiç üretildiyse 1 kere setupTime ekle, üstüne (adet * unitTime) ekle
                uretilenStandartDk += setupTime + (stProd * unitTime);
            }
        });
        
        const kapasiteSaat = Math.round(haftalikKapasiteDk / 60);
        const uretilenSaat = Math.round(uretilenStandartDk / 60);
        const yuzde = haftalikKapasiteDk > 0 ? Math.min(100, Math.round((uretilenStandartDk / haftalikKapasiteDk) * 100)) : 0;
        const overCapacity = yuzde >= 100;
        
        const cardHTML = `
            <div class="perf-card">
                <h3>
                    <span><i class="fa-solid fa-industry text-blue"></i> ${stName}</span>
                    <span style="color: ${overCapacity ? '#ef4444' : '#00f0ff'};">${yuzde}%</span>
                </h3>
                <div class="perf-bar-bg">
                    <div class="perf-bar-fill ${overCapacity ? 'over-capacity' : ''}" style="width: ${yuzde}%;"></div>
                </div>
                <div class="perf-stats">
                    <span><strong>${uretilenSaat}</strong> Saat Üretildi</span>
                    <span><strong>${kapasiteSaat}</strong> Saat Kapasite</span>
                </div>
                <div style="font-size: 11px; color: rgba(255,255,255,0.4); text-align: right;">
                    (Makine: ${makineSayisi} | Günlük: ${gunlukSaat}s)
                </div>
                <div class="perf-stats" style="margin-top: 10px; border-top: 1px solid rgba(255,255,255,0.1); padding-top: 10px;">
                    <span style="display:flex; align-items:center; gap:5px;">
                        <i class="fa-solid fa-pause-circle text-orange"></i> Duruş (Saat):
                    </span>
                    <input type="number" min="0" step="0.5" class="downtime-input" value="${downtimeHours}" 
                        onchange="updateDowntime('${stName.replace(/'/g, "\\'")}', this.value)" 
                        style="width: 60px; background: rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.2); color: white; border-radius: 4px; padding: 2px 5px; text-align: center;">
                </div>
            </div>
        `;
        
        container.insertAdjacentHTML('beforeend', cardHTML);
    });
}

function updateDowntime(stName, value) {
    const val = parseFloat(value) || 0;
    downtimeMap[stName] = val;
    saveDowntimeMapToStorage();
    renderPerformanceTab();
}

// ---- Sütun Gizle/Göster ----
function toggleColPicker() {
    const dd = document.getElementById('col-picker-dropdown');
    if (dd) dd.style.display = dd.style.display === 'none' ? 'block' : 'none';
}

// Dropdown kapanması için dışarı tıklama
document.addEventListener('click', function(e) {
    const btn = document.getElementById('col-picker-btn');
    const dd = document.getElementById('col-picker-dropdown');
    if (dd && btn && !btn.contains(e.target) && !dd.contains(e.target)) {
        dd.style.display = 'none';
    }
});

function _updateColPickerUI(displayCols) {
    const list = document.getElementById('col-picker-list');
    if (!list) return;
    // Yalnızca değiştiğinde yeniden çiz (dil değiştirme döngüsünü engelle)
    const key = displayCols.join(',');
    if (list.dataset.lastCols === key) {
        // Sadece checkbox durumlarını güncelle
        list.querySelectorAll('input[type=checkbox]').forEach(cb => {
            cb.checked = !hiddenStationCols.has(cb.dataset.col);
        });
        return;
    }
    list.dataset.lastCols = key;
    list.innerHTML = '';
    // "Durum" her zaman görünür olsun, gizlenemez
    const toggleableCols = displayCols.filter(c => c !== 'Durum');
    toggleableCols.forEach(col => {
        const label = document.createElement('label');
        label.style.cssText = 'display:flex; align-items:center; gap:8px; margin-bottom:6px; cursor:pointer; font-size:13px; color:var(--text-primary);';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.dataset.col = col;
        cb.checked = !hiddenStationCols.has(col);
        cb.style.cssText = 'width:14px; height:14px; cursor:pointer; accent-color: #00f0ff;';
        cb.onchange = () => toggleColVisibility(col, cb.checked);
        label.appendChild(cb);
        label.appendChild(document.createTextNode(col));
        list.appendChild(label);
    });
}

function toggleColVisibility(colName, isVisible) {
    if (isVisible) {
        hiddenStationCols.delete(colName);
    } else {
        hiddenStationCols.add(colName);
    }
    saveHiddenColsToStorage();
    renderStationTable(stationHeadersMap[activeStation] || []);
}

// ---- Sidebar Toggle ----
document.addEventListener('DOMContentLoaded', () => {
    const toggleBtn = document.getElementById('toggle-sidebar-btn');
    const sidebar = document.getElementById('main-sidebar');
    
    // Load state from local storage
    if (localStorage.getItem('sidebarCollapsed') === 'true') {
        sidebar.classList.add('collapsed');
    }

    if (toggleBtn && sidebar) {
        toggleBtn.addEventListener('click', () => {
            sidebar.classList.toggle('collapsed');
            localStorage.setItem('sidebarCollapsed', sidebar.classList.contains('collapsed'));
        });
    }

    // Export butonları — onclick yerine burada bağla (çift tetiklenmeyi önler)
    document.getElementById('export-station-btn')?.addEventListener('click', exportStationDataToExcel);
    document.getElementById('export-raw-excel-btn')?.addEventListener('click', exportRawMaterialsToExcel);
});


// ---- DRAG AND DROP REORDERING ----
let draggedStationRow = null;

window.handleStationRowDragStart = function(e) {
    // If the drag originated from a button, input, or interactive element, cancel drag!
    if (e.target.closest && e.target.closest('button, input, select, a, .part-img-btn')) {
        e.preventDefault();
        return;
    }
    draggedStationRow = this;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', this.dataset.index);
    this.classList.add('dragging');
};

window.handleStationRowDragOver = function(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const rect = this.getBoundingClientRect();
    const offset = e.clientY - rect.top;
    
    // Clear indicator on other rows
    document.querySelectorAll('#station-tbody tr').forEach(tr => {
        if (tr !== this) {
            tr.classList.remove('drag-over-top', 'drag-over-bottom');
        }
    });

    if (offset > rect.height / 2) {
        this.classList.remove('drag-over-top');
        this.classList.add('drag-over-bottom');
    } else {
        this.classList.remove('drag-over-bottom');
        this.classList.add('drag-over-top');
    }
};

window.handleStationRowDragLeave = function(e) {
    if (e.relatedTarget && this.contains(e.relatedTarget)) return;
    this.classList.remove('drag-over-top', 'drag-over-bottom');
};

window.handleStationRowDragEnd = function(e) {
    this.classList.remove('dragging');
    document.querySelectorAll('#station-tbody tr').forEach(tr => tr.classList.remove('drag-over-top', 'drag-over-bottom', 'dragging'));
};

window.handleStationRowDrop = function(e) {
    e.preventDefault();
    e.stopPropagation();
    this.classList.remove('drag-over-top', 'drag-over-bottom');
    
    if (draggedStationRow === this || !draggedStationRow) return;

    const tbody = this.parentNode;
    const allRows = Array.from(tbody.children);
    const dragIndex = parseInt(draggedStationRow.dataset.index);
    const dropIndex = parseInt(this.dataset.index);

    const rect = this.getBoundingClientRect();
    const offset = e.clientY - rect.top;
    const insertAfter = offset > rect.height / 2;

    // Özel Sıralamaya geç
    const sortSelect = document.getElementById('station-sort');
    if (sortSelect) sortSelect.value = 'custom';

    const originalArray = stationSheetsMap[activeStation];
    if (!originalArray) return;
    
    const dragDataRow = paginationState.station.filtered[dragIndex];
    const targetDataRow = paginationState.station.filtered[dropIndex];
    
    const originalDragIndex = originalArray.indexOf(dragDataRow);
    const originalDropIndex = originalArray.indexOf(targetDataRow);
    
    if (originalDragIndex > -1 && originalDropIndex > -1) {
        // Eski yerinden çıkar
        originalArray.splice(originalDragIndex, 1);
        
        // Yeni hedefin güncel indeksini bul (dizi kaymış olabilir)
        const newTargetIndex = originalArray.indexOf(targetDataRow);
        
        let finalInsertIndex = newTargetIndex;
        if (insertAfter) {
            finalInsertIndex = newTargetIndex + 1;
        }
        
        // Yeni yerine ekle
        originalArray.splice(finalInsertIndex, 0, dragDataRow);
        
        // Kümülatif süreyi baştan hesapla
        recalcKumulatifSureForStation(activeStation);
        
        // Yeniden render
        filterAndPaginateStationData();
    }
};

function recalcKumulatifSureForStation(stName) {
    const rows = stationSheetsMap[stName];
    if (!rows) return;
    let kumulatif = 0;
    for (const row of rows) {
        const code = String(row['Kod'] || '').trim().toUpperCase();
        let hazirlikKey = Object.keys(row).find(k => k.toLowerCase() === 'hazırlık süresi') || 'Hazırlık Süresi';
        let birimKey = Object.keys(row).find(k => k.toLowerCase() === 'birim işlem süresi') || 'Birim İşlem Süresi';
        let miktarKey = Object.keys(row).find(k => k.toLowerCase() === 'üretilecek miktar') || 'Üretilecek Miktar';
        let toplamSureKey = Object.keys(row).find(k => k.toLowerCase() === 'toplam süre') || 'Toplam Süre';
        let saatKey = Object.keys(row).find(k => k.toLowerCase() === 'saat') || 'Saat';
        let kumulatifKey = Object.keys(row).find(k => k.toLowerCase() === 'kümülatif süre') || 'Kümülatif Süre';
        
        const hazirlik = parseFloat(row[hazirlikKey]) || 0;
        const birim = parseFloat(row[birimKey]) || 0;
        const miktar = parseFloat(row[miktarKey]) || 0;
        
        const reqs = uretimTakipRows.filter(u => u.kod === code);
        let isDone = false;
        if (reqs.length > 0) {
            const totalReq = reqs.reduce((sum, u) => sum + u.uretilecek, 0.0);
            const stLogs = productionLog.filter(log => log.kod === code && (log.station === stName || log.station === 'Tüm İstasyonlar' || !log.station));
            const stProd = stLogs.reduce((sum, log) => sum + parseFloat(log.adet), 0.0);
            if (totalReq > 0 && stProd >= totalReq) {
                isDone = true;
            }
        }

        if (isDone) {
            row[toplamSureKey] = 0;
            row[saatKey] = 0;
        } else {
            row[toplamSureKey] = hazirlik + (birim * miktar);
            row[saatKey] = row[toplamSureKey] / 86400;
        }
        
        const saatVal = parseFloat(row[saatKey]) || 0;
        kumulatif += saatVal;
        row[kumulatifKey] = kumulatif;
    }
}

// -------------------------------------------------------------
// 8. HAMMADDE SİPARİŞİ VIEW
// -------------------------------------------------------------
function renderRawMaterialsTab() {
    filterAndPaginateRawMaterials();
}

function filterAndPaginateRawMaterials() {
    const searchInput = document.getElementById('raw-search');
    const searchVal = searchInput ? searchInput.value.toLowerCase().trim() : '';
    const statusSelect = document.getElementById('raw-filter-status');
    const statusVal = statusSelect ? statusSelect.value : 'all';

    let filtered = rawMaterialsRows;

    if (searchVal) {
        filtered = filtered.filter(r => {
            const kod = String(r.kod || '').toLowerCase();
            const ad = String(r.ad || '').toLowerCase();
            const matchInDetails = Array.isArray(r.details) && r.details.some(d => String(d.parcaKodu || '').toLowerCase().includes(searchVal));
            return kod.includes(searchVal) || ad.includes(searchVal) || matchInDetails;
        });
    }

    if (statusVal === 'needed') {
        filtered = filtered.filter(r => r.kalanSiparis > 0);
    } else if (statusVal === 'completed') {
        filtered = filtered.filter(r => r.kalanSiparis === 0 && r.toplamGereken > 0);
    }

    // Sıralama (Kalan sipariş miktarına göre azalan)
    filtered.sort((a, b) => b.kalanSiparis - a.kalanSiparis);

    paginationState.raw.filtered = filtered;
    paginationState.raw.total = filtered.length;

    // KPI Güncellemeleri
    const totalVariety = rawMaterialsRows.length;
    const totalQty = rawMaterialsRows.reduce((sum, r) => sum + r.toplamGereken, 0);
    const remQty = rawMaterialsRows.reduce((sum, r) => sum + r.kalanSiparis, 0);
    const compQty = rawMaterialsRows.reduce((sum, r) => sum + r.uretilenDusulen, 0);

    const kpiTotalVariety = document.getElementById('kpi-raw-total-variety');
    const kpiTotalQty = document.getElementById('kpi-raw-total-qty');
    const kpiRemQty = document.getElementById('kpi-raw-remaining-qty');
    const kpiCompQty = document.getElementById('kpi-raw-completed-qty');

    if (kpiTotalVariety) kpiTotalVariety.textContent = totalVariety.toLocaleString();
    if (kpiTotalQty) kpiTotalQty.textContent = (Math.round(totalQty * 100) / 100).toLocaleString();
    if (kpiRemQty) kpiRemQty.textContent = (Math.round(remQty * 100) / 100).toLocaleString();
    if (kpiCompQty) kpiCompQty.textContent = (Math.round(compQty * 100) / 100).toLocaleString();

    renderRawMaterialsTable();
}
window.filterAndPaginateRawMaterials = filterAndPaginateRawMaterials;

function renderRawMaterialsTable() {
    const tbody = document.getElementById('raw-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const searchInput = document.getElementById('raw-search');
    const searchVal = searchInput ? searchInput.value.toLowerCase().trim() : '';

    const pState = paginationState.raw;
    if (pState.total === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="color:var(--text-dim); padding:20px;">Hammadde kaydı bulunamadı.</td></tr>';
        return;
    }

    pState.filtered.forEach((row, idx) => {
        const tr = document.createElement('tr');
        const isDone = row.kalanSiparis === 0 && row.toplamGereken > 0;
        const isPartial = row.uretilenDusulen > 0 && row.kalanSiparis > 0;
        const hasDetails = row.details && row.details.length > 0;

        // Eğer aramada alt parça kodu eşleştiyse detay paneli otomatik açık gelsin
        const matchInDetails = searchVal && hasDetails && row.details.some(d => String(d.parcaKodu || '').toLowerCase().includes(searchVal));
        const autoOpen = !!matchInDetails;

        // Karşılanma Oranı ve İlerleme Çubuğu
        const compPct = row.toplamGereken > 0 ? Math.min(100, Math.round((row.uretilenDusulen / row.toplamGereken) * 100)) : (row.uretilenDusulen > 0 ? 100 : 0);
        let barColor = '#ef4444'; // Kırmızı (Bekliyor)
        let statusLabel = 'Bekliyor';
        if (isDone || compPct >= 100) {
            barColor = 'var(--success)';
            statusLabel = 'Tamamlandı';
            tr.classList.add('station-row-completed');
        } else if (compPct > 0) {
            barColor = 'var(--warning)';
            statusLabel = 'Karşılanıyor';
        }

        const progressHtml = `
            <div style="min-width: 120px; display: flex; flex-direction: column; gap: 4px;">
                <div style="height: 6px; width: 100%; background: rgba(255,255,255,0.08); border-radius: 3px; overflow: hidden;">
                    <div style="width: ${compPct}%; background-color: ${barColor}; height: 100%; border-radius: 3px; transition: width 0.3s ease;"></div>
                </div>
                <div style="display: flex; justify-content: space-between; font-size: 11px; font-weight: 600;">
                    <span style="color: ${barColor};">${statusLabel}</span>
                    <span style="color: var(--text-dim);">${compPct}%</span>
                </div>
            </div>
        `;

        const detailBtnHtml = hasDetails
            ? `<button class="raw-detail-btn" title="Parça detaylarını göster" onclick="toggleRawDetail(this, ${idx})" style="background:${autoOpen ? 'rgba(99,102,241,0.2)' : 'none'};border:1px solid ${autoOpen ? '#6366f1' : 'var(--border)'};border-radius:4px;padding:2px 7px;cursor:pointer;color:${autoOpen ? '#a78bfa' : 'var(--text-muted)'};font-size:11px;margin-left:6px;">${autoOpen ? '▲' : '▼'}</button>`
            : '';

        // Ekstra Parti Hammadde Rozeti (Ana Satır)
        const extraRawTotal = Math.round((row.toplamEkstraHammadde || 0) * 100) / 100;
        const origRawTotal = Math.round((row.toplamOrijinalGereken || row.toplamGereken) * 100) / 100;
        const totalReqDisplay = row.toplamGereken % 1 === 0 ? row.toplamGereken : row.toplamGereken.toFixed(2);
        
        const extraRawBadge = extraRawTotal > 0 ? `
            <div style="font-size:10.5px; margin-top:4px; display:flex; flex-direction:column; align-items:flex-end; gap:2px;">
                <span style="color:rgba(255,255,255,0.6); font-size:10px;">Orijinal: <strong style="color:white;">${origRawTotal}</strong></span>
                <span class="badge" style="background:rgba(56,189,248,0.15); color:#38bdf8; border:1px solid rgba(56,189,248,0.3); font-size:10px; font-weight:700; padding:1px 6px; border-radius:3px;">
                    <i class="fa-solid fa-cube" style="font-size:9px;"></i> Ekstra: +${extraRawTotal} m
                </span>
            </div>
        ` : '';

        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td style="font-weight:700; color:white; white-space:nowrap;">${row.kod}${detailBtnHtml}</td>
            <td style="color:var(--text-muted); max-width:280px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${row.ad}">${row.ad}</td>
            <td class="text-right" style="font-weight:600;">
                <div>${totalReqDisplay}</div>
                ${extraRawBadge}
            </td>
            <td class="text-right" style="color:var(--success); font-weight:600;">${row.uretilenDusulen % 1 === 0 ? row.uretilenDusulen : row.uretilenDusulen.toFixed(2)}</td>
            <td class="text-right" style="color:var(--warning); font-weight:700; font-size:14px;">${row.kalanSiparis % 1 === 0 ? row.kalanSiparis : row.kalanSiparis.toFixed(2)}</td>
            <td>${progressHtml}</td>
        `;
        tr.dataset.detailIdx = idx;
        tbody.appendChild(tr);

        // Detay satırı
        if (hasDetails) {
            const detailTr = document.createElement('tr');
            detailTr.className = 'raw-detail-row';
            detailTr.style.display = autoOpen ? 'table-row' : 'none';
            detailTr.dataset.parentIdx = idx;

            const detailRows = row.details.map(d => {
                const bMiktarFloat = parseFloat(d.birimMiktar) || 0;
                const mmVal = Math.round(bMiktarFloat * 1000 * 100) / 100;
                const birimStr = ` × ${mmVal} mm/adet`;
                const kalanColor = d.kalanMiktar > 0 ? 'var(--warning)' : 'var(--success)';
                const isMatchedPart = searchVal && String(d.parcaKodu || '').toLowerCase().includes(searchVal);
                const partBg = isMatchedPart ? 'rgba(99,102,241,0.18)' : 'rgba(255,255,255,0.03)';
                const partBorder = isMatchedPart ? 'border-left: 3px solid #6366f1;' : '';
                
                const extraPartQty = Math.round((d.extraQty || 0) * 100) / 100;
                const extraPartRaw = Math.round((d.extraMiktar || 0) * 100) / 100;
                
                const qtyDetailHtml = extraPartQty > 0 ? `
                    <div>
                        Adet: <b style="color:white;">${d.uretilecek}</b>${birimStr}
                        <span class="badge" style="background:rgba(16,185,129,0.15); color:#34d399; border:1px solid rgba(16,185,129,0.3); font-size:10px; font-weight:700; padding:1px 5px; margin-left:6px; border-radius:3px;">
                            +${extraPartQty} Adet
                        </span>
                        <span style="color:rgba(255,255,255,0.5); font-size:10px; margin-left:4px;">(Orijinal: ${d.orijinalUretilecek})</span>
                    </div>
                ` : `Adet: <b style="color:white;">${d.uretilecek}</b>${birimStr}`;

                const totalDetailHtml = extraPartRaw > 0 ? `
                    <div style="display:flex; flex-direction:column; align-items:flex-end; gap:2px;">
                        <span>Toplam: <b style="color:white;">${d.toplamMiktar % 1 === 0 ? d.toplamMiktar : d.toplamMiktar.toFixed(2)}</b></span>
                        <span style="color:#38bdf8; font-size:10px; font-weight:700; background:rgba(56,189,248,0.12); padding:1px 5px; border-radius:3px; border:1px solid rgba(56,189,248,0.25);">
                            +${extraPartRaw} m Ekstra
                        </span>
                    </div>
                ` : `Toplam: <b style="color:white;">${d.toplamMiktar % 1 === 0 ? d.toplamMiktar : d.toplamMiktar.toFixed(2)}</b>`;

                return `
                    <tr style="background:${partBg}; border-bottom:1px solid var(--border); ${partBorder}">
                        <td style="padding:6px 12px; font-weight:700; color:${isMatchedPart ? '#38bdf8' : '#a78bfa'}; white-space:nowrap;">
                            ${d.parcaKodu}
                            <button class="part-img-btn" onclick="openPartImageModal('${d.parcaKodu}', '')" onmouseenter="window.showPartHoverPreview(event, '${d.parcaKodu}')" onmousemove="window.movePartHoverPreview(event)" onmouseleave="window.hidePartHoverPreview()" title="Parça Görselini Görüntüle"><i class="fa-solid fa-image"></i></button>
                            ${isMatchedPart ? ' <span style="font-size:10px; background:#6366f1; color:white; padding:1px 5px; border-radius:3px; margin-left:4px;">Eşleşti</span>' : ''}
                        </td>
                        <td style="padding:6px 12px; color:var(--text-muted);">${qtyDetailHtml}</td>
                        <td style="padding:6px 12px; text-align:right;">${totalDetailHtml}</td>
                        <td style="padding:6px 12px; text-align:right;">Kullanılan: <b style="color:var(--success);">${d.uretilenMiktar % 1 === 0 ? d.uretilenMiktar : d.uretilenMiktar.toFixed(2)}</b></td>
                        <td style="padding:6px 12px; text-align:right;">Kalan: <b style="color:${kalanColor};">${d.kalanMiktar % 1 === 0 ? d.kalanMiktar : d.kalanMiktar.toFixed(2)}</b></td>
                    </tr>`;
            }).join('');

            detailTr.innerHTML = `
                <td colspan="7" style="padding:0; background:rgba(99,102,241,0.05); border-left:3px solid #6366f1;">
                    <table style="width:100%; border-collapse:collapse; font-size:12px;">
                        <thead>
                            <tr style="background:rgba(99,102,241,0.1);">
                                <th style="padding:6px 12px; text-align:left; color:#a78bfa;">Parça Kodu</th>
                                <th style="padding:6px 12px; text-align:left; color:#a78bfa;">Üretilecek Adet</th>
                                <th style="padding:6px 12px; text-align:right; color:#a78bfa;">Toplam Miktar</th>
                                <th style="padding:6px 12px; text-align:right; color:#a78bfa;">Kullanılan</th>
                                <th style="padding:6px 12px; text-align:right; color:#a78bfa;">Kalan</th>
                            </tr>
                        </thead>
                        <tbody>${detailRows}</tbody>
                    </table>
                </td>
            `;
            tbody.appendChild(detailTr);
        }
    });
}

function toggleRawDetail(btn, idx) {
    const tbody = document.getElementById('raw-table-body');
    const detailRow = tbody.querySelector(`.raw-detail-row[data-parent-idx="${idx}"]`);
    if (!detailRow) return;
    const isOpen = detailRow.style.display !== 'none';
    detailRow.style.display = isOpen ? 'none' : 'table-row';
    btn.textContent = isOpen ? '▼' : '▲';
}
window.toggleRawDetail = toggleRawDetail;



function exportRawMaterialsToExcel() {
    // Sadece sipariş ihtiyacı kalanları (kalanSiparis > 0) al (Tamamlananlar hariç)
    const neededRows = rawMaterialsRows.filter(r => r.kalanSiparis > 0);
    
    if (neededRows.length === 0) {
        showToast("Sipariş edilecek kalan hammadde bulunamadı (Tüm hammadde ihtiyaçları karşılanmış).", "info");
        return;
    }

    // Yeni Excel çalışma kitabı oluştur
    const wb = XLSX.utils.book_new();
    
    // Satır verilerini hazırla (Excel'deki HAMMADDE SİPARİŞ formatı)
    const exportData = neededRows.map(r => ({
        "Hammadde Kod": r.kod,
        "Hammadde / Malzeme Adı": r.ad,
        "Sipariş Miktarı": r.kalanSiparis % 1 === 0 ? r.kalanSiparis : parseFloat(r.kalanSiparis.toFixed(3))
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // Sütun genişlikleri ayarla
    ws['!cols'] = [
        { wch: 25 }, // Hammadde Kod
        { wch: 65 }, // Hammadde / Malzeme Adı
        { wch: 18 }  // Sipariş Miktarı
    ];

    XLSX.utils.book_append_sheet(wb, ws, "HAMMADDE SİPARİŞ");

    const origFileName = loadedFileName ? loadedFileName.textContent.replace("Yüklenen Dosya: ", "").replace(".xlsx", "") : "Hammadde";
    const exportName = `${origFileName}_Hammadde_Siparis_Listesi.xlsx`;

    XLSX.writeFile(wb, exportName);
    showToast(`Kalan hammadde sipariş listesi indirildi: "${exportName}"`, "success");
}
window.exportRawMaterialsToExcel = exportRawMaterialsToExcel;

// -------------------------------------------------------------
// 9. PARÇA GÖRSELİ MODALI & HOVER ÖNİZLEME (LIGHTBOX)
// -------------------------------------------------------------
const partImageCache = new Map(); // kod -> { url, w, h } veya null

function findPartImageUrl(kod, callback) {
    const cleanKod = String(kod || '').trim();
    if (!cleanKod) { callback(null); return; }
    if (partImageCache.has(cleanKod)) {
        const cached = partImageCache.get(cleanKod);
        if (cached) callback(cached.url, cached.w, cached.h);
        else callback(null);
        return;
    }
    
    const isHttp = window.location.protocol.startsWith('http');
    const queryParam = isHttp ? `?v=${Date.now()}` : '';

    const kodVariants = [cleanKod];
    const strippedZero = cleanKod.replace(/^0+/, '');
    if (strippedZero && strippedZero !== cleanKod) kodVariants.push(strippedZero);
    const noSpecial = cleanKod.replace(/[\/\\]/g, '_');
    if (noSpecial !== cleanKod) kodVariants.push(noSpecial);

    const extensions = ['jpg', 'png', 'jpeg', 'webp', 'JPG', 'PNG', 'JPEG', 'WEBP'];
    const searchQueue = [];
    kodVariants.forEach(k => {
        extensions.forEach(ext => {
            searchQueue.push(`images/parcalar/${k}.${ext}${queryParam}`);
        });
    });

    let queueIdx = 0;
    function testNext() {
        if (queueIdx >= searchQueue.length) {
            partImageCache.set(cleanKod, null);
            callback(null);
            return;
        }
        const src = searchQueue[queueIdx++];
        const testImg = new Image();
        testImg.onload = function() {
            partImageCache.set(cleanKod, { url: src, w: testImg.naturalWidth, h: testImg.naturalHeight });
            callback(src, testImg.naturalWidth, testImg.naturalHeight);
        };
        testImg.onerror = function() {
            testNext();
        };
        testImg.src = src;
    }
    testNext();
}

function showPartHoverPreview(e, kod) {
    const popup = document.getElementById('part-hover-popup');
    const inner = document.getElementById('part-hover-inner');
    if (!popup || !inner) return;

    findPartImageUrl(kod, (url) => {
        if (!url) {
            popup.classList.remove('active');
            return;
        }
        inner.innerHTML = `
            <img src="${url}" class="part-hover-img" alt="${kod}">
            <div class="part-hover-label"><i class="fa-solid fa-cube text-blue"></i> ${kod}</div>
        `;
        popup.classList.add('active');
        positionHoverPopup(e);
    });
}

function positionHoverPopup(e) {
    const popup = document.getElementById('part-hover-popup');
    if (!popup || !popup.classList.contains('active')) return;
    
    const popupWidth = 165;
    const popupHeight = 185;
    let finalX = e.clientX + 16;
    let finalY = e.clientY - 90;
    
    if (finalX + popupWidth > window.innerWidth) {
        finalX = e.clientX - popupWidth - 16;
    }
    if (finalY + popupHeight > window.innerHeight) {
        finalY = window.innerHeight - popupHeight - 12;
    }
    if (finalY < 12) finalY = 12;

    popup.style.left = `${finalX}px`;
    popup.style.top = `${finalY}px`;
}

function movePartHoverPreview(e) {
    positionHoverPopup(e);
}

function hidePartHoverPreview() {
    const popup = document.getElementById('part-hover-popup');
    if (popup) popup.classList.remove('active');
}

function openPartImageModal(kod, matName) {
    hidePartHoverPreview();
    const modal = document.getElementById('part-image-modal');
    const container = document.getElementById('part-img-preview-container');
    if (!modal || !container) return;

    const cleanKod = String(kod || '').trim();
    const decodedMatName = matName ? decodeURIComponent(matName) : '';

    container.innerHTML = `
        <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; gap:12px; color:var(--text-muted); padding:40px;">
            <i class="fa-solid fa-spinner fa-spin" style="font-size:36px; color:#818cf8;"></i>
            <span style="font-size:14px;">Görsel yükleniyor (${cleanKod})...</span>
        </div>
    `;

    modal.classList.add('active');

    findPartImageUrl(cleanKod, (url, w, h) => {
        if (url) {
            container.innerHTML = `
                <img src="${url}" alt="${cleanKod}" class="part-lightbox-img">
                <div class="part-lightbox-tag">
                    <i class="fa-solid fa-cube" style="color:#818cf8;"></i>
                    <span>Parça: <b style="color:white;">${cleanKod}</b>${decodedMatName ? ' - ' + decodedMatName : ''}</span>
                    <span style="opacity:0.6; font-size:11px; margin-left:4px;">(${w} × ${h} px)</span>
                </div>
            `;
        } else {
            container.innerHTML = `
                <div style="text-align:center; padding:32px 24px; max-width:440px; background:#111827; border:1px solid rgba(99,102,241,0.35); border-radius:14px; box-shadow:0 25px 60px rgba(0,0,0,0.85);" onclick="event.stopPropagation()">
                    <div style="width:58px; height:58px; border-radius:50%; background:rgba(99,102,241,0.1); border:1px solid rgba(99,102,241,0.25); display:inline-flex; align-items:center; justify-content:center; margin-bottom:12px;">
                        <i class="fa-regular fa-image" style="font-size:26px; color:#818cf8;"></i>
                    </div>
                    <h4 style="margin:0 0 6px 0; color:white; font-size:16px;">Görsel Bulunamadı</h4>
                    <p style="color:var(--text-muted); font-size:13px; line-height:1.5; margin:0 0 12px 0;">
                        <b>${cleanKod}</b> kodlu parça için klasörde görsel dosyası tespit edilemedi.
                    </p>
                    <code style="display:block; margin-bottom:14px; background:#1e1b4b; padding:6px 10px; border-radius:6px; color:#38bdf8; font-size:12px; word-break:break-all;">dashboard/images/parcalar/${cleanKod}.jpg</code>
                    <input type="file" id="part-file-browser" accept="image/*" style="display:none;" onchange="handleDirectImagePreview(this)">
                    <button class="btn btn-primary btn-sm" onclick="document.getElementById('part-file-browser').click()" style="display:inline-flex; align-items:center; gap:6px; font-size:12px; padding:6px 14px;">
                        <i class="fa-solid fa-upload"></i> Bilgisayardan Görsel Seç
                    </button>
                </div>
            `;
        }
    });
}

function handleDirectImagePreview(input) {
    if (input.files && input.files[0]) {
        const file = input.files[0];
        const reader = new FileReader();
        const container = document.getElementById('part-img-preview-container');
        reader.onload = function(e) {
            if (container) {
                container.innerHTML = `
                    <img src="${e.target.result}" alt="Önizleme" class="part-lightbox-img">
                    <div class="part-lightbox-tag">
                        <i class="fa-solid fa-check-circle" style="color:var(--success);"></i>
                        <span>Seçilen Görsel: <b style="color:white;">${file.name}</b></span>
                    </div>
                `;
            }
        };
        reader.readAsDataURL(file);
    }
}

function closePartImageModal() {
    const modal = document.getElementById('part-image-modal');
    if (modal) {
        modal.classList.remove('active');
    }
}

function handlePartModalOverlayClick(event) {
    if (event.target.id === 'part-image-modal' || event.target.classList.contains('part-lightbox-overlay')) {
        closePartImageModal();
    }
}

// ESC tuşu ile modal kapatma
document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        closePartImageModal();
        hidePartHoverPreview();
    }
});

window.openPartImageModal = openPartImageModal;
window.closePartImageModal = closePartImageModal;
window.handlePartModalOverlayClick = handlePartModalOverlayClick;
window.showPartHoverPreview = showPartHoverPreview;
window.movePartHoverPreview = movePartHoverPreview;
window.hidePartHoverPreview = hidePartHoverPreview;
window.handleDirectImagePreview = handleDirectImagePreview;

// --- REÇETE / REVİZYON EXCEL YÜKLEME VE VERİTABANI GÜNCELLEME ---
function handleRecipeDbFileUpload(input) {
    if (!input.files || input.files.length === 0) return;
    
    const files = Array.from(input.files);
    let processedCount = 0;
    let totalUpdatedParts = 0;
    
    files.forEach(file => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                const firstSheetName = wb.SheetNames[0];
                const sheet = wb.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                if (!rows || rows.length <= 1) return;
                
                const defaultMachine = file.name.replace(/\.[^/.]+$/, "");
                const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
                
                const is3Col = headers.length >= 3 && (headers[0].includes('kaynak') || headers[0].includes('reçete') || headers[0].includes('recete'));
                
                for (let r = 1; r < rows.length; r++) {
                    const row = rows[r];
                    if (!row || row.length === 0) continue;
                    
                    let src = defaultMachine;
                    let code = '';
                    let qty = 1.0;
                    
                    if (is3Col) {
                        src = String(row[0] || defaultMachine).trim();
                        code = String(row[1] || '').trim().toUpperCase();
                        qty = parseFloat(row[2]) || 0;
                    } else {
                        code = String(row[0] || '').trim().toUpperCase();
                        qty = parseFloat(row[1]) || 1.0;
                    }
                    
                    if (!code || code === 'NAN' || code === '-') continue;
                    
                    if (!parcaMakineReceteleriMap[code]) {
                        parcaMakineReceteleriMap[code] = {
                            kod: code,
                            makineler: {}
                        };
                    }
                    
                    parcaMakineReceteleriMap[code].makineler[src] = qty;
                    totalUpdatedParts++;
                }
                
                // İstatistikleri güncelle
                Object.values(parcaMakineReceteleriMap).forEach(p => {
                    p.makine_sayisi = Object.keys(p.makineler || {}).length;
                    p.toplam_birim_adet = Object.values(p.makineler || {}).reduce((a, b) => a + b, 0);
                });
                
                processedCount++;
                if (processedCount === files.length) {
                    recalculateAll();
                    filterAndPaginateUlTable();
                    showToast(`✅ ${files.length} reçete dosyası başarıyla işlendi! (${totalUpdatedParts} parça kaydı güncellendi)`, "success");
                }
            } catch (err) {
                console.error('Reçete yükleme hatası:', err);
                showToast(`"${file.name}" dosyası işlenirken hata oluştu!`, "error");
            }
        };
        reader.readAsArrayBuffer(file);
    });
    
    input.value = ''; // Reset input
}
window.handleRecipeDbFileUpload = handleRecipeDbFileUpload;
