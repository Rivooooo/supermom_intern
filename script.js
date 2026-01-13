// Column Mappings (0-based)
const COL = {
    FIRST_NAME: 13, // Nama Depan
    LAST_NAME: 14,  // Nama Belakang
    EMAIL: 15,      // Email (Recommended...)
    EDD: 17,        // EDD / HPL
    INSURANCE: 18,  // Asuransi
    HOSPITAL: 20,   // Rumah Sakit
    HOSPITAL_OTHER: 21, // RS Lainnya
    PHOTO_USG: 22,  // Foto USG
    PHOTO_BOOK: 23, // Foto Buku Kehamilan
    NOTE_ANIS: 60,  // Note (Anis) - Column BI
    MANUAL_CHECK: 62, // Manual Check - Column BK
    REASON: 63,     // Reason - Column BL
    UPLOADED: 64    // Uploaded to Audience List? - Column BM
};

// State
let state = {
    allRows: [], // Array of arrays
    filteredIndices: [], // Indices of rows to review
    currentIndex: 0, // Pointer within filteredIndices
    currentRotation: 0,
    showingUSG: true, // Toggle between USG (true) and Book (false)
    headers: [],

    // Zoom & Pan State
    scale: 1,
    panning: false,
    pointX: 0,
    pointY: 0,
    startX: 0,
    startY: 0
};

// DOM Elements
const els = {
    uploadInput: document.getElementById('csv-upload'),
    btnUpload: document.getElementById('btn-upload'),
    btnExport: document.getElementById('btn-export'),
    mainInterface: document.getElementById('main-interface'),
    rowCounter: document.getElementById('row-counter'),

    // Global Loading
    loadingOverlay: document.getElementById('loading-overlay'),
    loadingTitle: document.getElementById('loading-title'),
    loadingBar: document.getElementById('loading-bar'),
    loadingPercent: document.getElementById('loading-percent'),

    // Image
    mainImage: document.getElementById('main-image'),
    imageWrapper: document.getElementById('image-wrapper'),
    imageContainer: document.getElementById('image-container'),
    imageError: document.getElementById('image-error'),
    imageLoading: document.getElementById('image-loading'),
    errorEmail: document.getElementById('error-email'),
    btnRotateLeft: document.getElementById('btn-rotate-left'),
    btnRotateRight: document.getElementById('btn-rotate-right'),
    btnToggleImage: document.getElementById('btn-toggle-image'),

    // Info
    displayName: document.getElementById('display-name'),
    displayEmail: document.getElementById('display-email'),
    displayEdd: document.getElementById('display-edd'),

    // Calculator
    calcWeeks: document.getElementById('calc-weeks'),
    calcDays: document.getElementById('calc-days'),
    calcDate: document.getElementById('calc-date'),
    btnCalculate: document.getElementById('btn-calculate'),
    calcResult: document.getElementById('calc-result'),
    resultDate: document.getElementById('result-date'),

    // Actions
    inputReason: document.getElementById('input-reason'),
    inputReasonOther: document.getElementById('input-reason-other'),
    btnValid: document.getElementById('btn-valid'),
    btnInvalid: document.getElementById('btn-invalid'),
    btnClear: document.getElementById('btn-clear'),

    // Overlay
    overlay: document.getElementById('completion-overlay'),
    btnDownloadFinal: document.getElementById('btn-download-final')
};


// Loading Overlay Helpers
function showLoading(title = 'Processing...', options = {}) {
    els.loadingOverlay.classList.remove('hidden');
    els.loadingTitle.textContent = title;

    if (options.indeterminate) {
        els.loadingBar.style.width = '100%';
        els.loadingBar.style.animation = 'pulse 1.5s ease-in-out infinite';
        els.loadingPercent.textContent = '';
    } else {
        els.loadingBar.style.animation = 'none';
        const percent = options.percent || 0;
        setLoadingProgress(percent);
    }
}

function hideLoading() {
    els.loadingOverlay.classList.add('hidden');
}

function setLoadingProgress(percent) {
    const clamped = Math.min(100, Math.max(0, percent));
    els.loadingBar.style.width = clamped + '%';
    els.loadingPercent.textContent = Math.round(clamped) + '%';
}
// Initialization
function init() {
    els.btnUpload.addEventListener('click', () => els.uploadInput.click());
    els.uploadInput.addEventListener('change', handleFileUpload);

    els.btnRotateLeft.addEventListener('click', () => rotateImage(-90));
    els.btnRotateRight.addEventListener('click', () => rotateImage(90));
    els.btnToggleImage.addEventListener('click', toggleImageSource);

    els.btnCalculate.addEventListener('click', calculateEDD);

    els.inputReason.addEventListener('change', (e) => {
        if (e.target.value === 'Other') {
            els.inputReasonOther.classList.remove('hidden');
        } else {
            els.inputReasonOther.classList.add('hidden');
        }
    });

    els.btnValid.addEventListener('click', () => submitReview('Valid'));
    els.btnInvalid.addEventListener('click', () => submitReview('Invalid'));
    els.btnClear.addEventListener('click', () => submitReview(''));

    // Navigation
    document.getElementById('btn-prev').addEventListener('click', prevRow);
    document.getElementById('btn-next').addEventListener('click', nextRow);

    els.btnExport.addEventListener('click', exportData);
    els.btnDownloadFinal.addEventListener('click', exportData);

    // Zoom & Pan Listeners
    els.imageContainer.addEventListener('wheel', handleZoom);
    els.imageContainer.addEventListener('mousedown', startPan);
    els.imageContainer.addEventListener('mousemove', pan);
    els.imageContainer.addEventListener('mouseup', endPan);
    els.imageContainer.addEventListener('mouseleave', endPan);

    // Image Load Handler
    els.mainImage.onload = () => {
        els.imageLoading.classList.add('hidden');
        els.btnToggleImage.disabled = false;

        const containerWidth = els.imageContainer.clientWidth;
        const containerHeight = els.imageContainer.clientHeight;
        const imgWidth = els.mainImage.offsetWidth;
        const imgHeight = els.mainImage.offsetHeight;

        state.pointX = (containerWidth - imgWidth) / 2;
        state.pointY = (containerHeight - imgHeight) / 2;
        updateImageTransform();
    };
    els.mainImage.onerror = () => {
        els.imageLoading.classList.add('hidden');
        els.btnToggleImage.disabled = false;
    };
}

// Zoom & Pan Logic
function handleZoom(e) {
    e.preventDefault();

    const rect = els.imageContainer.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    const xs = (x - state.pointX) / state.scale;
    const ys = (y - state.pointY) / state.scale;

    const delta = -Math.sign(e.deltaY);
    const factor = 0.1;

    let newScale = state.scale;
    if (delta > 0) {
        newScale *= (1 + factor);
    } else {
        newScale /= (1 + factor);
    }

    if (newScale < 1) newScale = 1;
    if (newScale > 10) newScale = 10;

    state.pointX = x - (xs * newScale);
    state.pointY = y - (ys * newScale);
    state.scale = newScale;

    updateImageTransform();
}

function startPan(e) {
    e.preventDefault();
    state.startX = e.clientX - state.pointX;
    state.startY = e.clientY - state.pointY;
    state.panning = true;
    els.imageContainer.style.cursor = 'grabbing';
}

function pan(e) {
    e.preventDefault();
    if (!state.panning) return;

    state.pointX = e.clientX - state.startX;
    state.pointY = e.clientY - state.startY;

    updateImageTransform();
}

function endPan(e) {
    state.panning = false;
    els.imageContainer.style.cursor = 'grab';
}

// File Handling
function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    Papa.parse(file, {
        skipEmptyLines: true,
        complete: (results) => {
            console.log('CSV parsed, rows:', results.data?.length);
            processCSV(results.data);
        },
        error: (err) => {
            console.error('CSV Error:', err);
            alert('Error parsing CSV file.');
        }
    });
}

function processCSV(data) {
    if (!data || data.length === 0) return;

    state.headers = data[0];
    state.allRows = data;

    console.log('Mapping Verification:');
    console.log(`BI (${COL.NOTE_ANIS}): ${state.headers?.[COL.NOTE_ANIS]}`);
    console.log(`BK (${COL.MANUAL_CHECK}): ${state.headers?.[COL.MANUAL_CHECK]}`);
    console.log(`BL (${COL.REASON}): ${state.headers?.[COL.REASON]}`);
    console.log(`BM (${COL.UPLOADED}): ${state.headers?.[COL.UPLOADED]}`);

    // Pre-processing & Filtering
    state.filteredIndices = [];

    for (let i = 1; i < state.allRows.length; i++) {
        const row = state.allRows[i];
        // Only require essential columns (up to PHOTO_BOOK at index 23)
        // Columns beyond that (NOTE_ANIS, MANUAL_CHECK, REASON, UPLOADED) are optional
        if (!row || row.length <= COL.PHOTO_BOOK) continue;

        const bk = (row[COL.MANUAL_CHECK] || '').trim();
        let bi = (row[COL.NOTE_ANIS] || '').trim();

        const bm = (row[COL.UPLOADED] || '').trim();
        const bl = (row[COL.REASON] || '').trim();
        const firstName = (row[COL.FIRST_NAME] || '').trim();

        const isBkEligible = ['Need Confirmation', 'Valid', ''].includes(bk);

        if (i < 5) {
            console.log(
                `Row ${i}: BK='${bk}', BI='${bi}', BL='${bl}', BM='${bm}', FirstName='${firstName}'`
            );
        }

        if (
            bm === '' &&
            isBkEligible &&
            bi === '' &&
            bl === '' &&
            firstName !== ''
        ) {
            state.filteredIndices.push(i);
        }
    }

    console.log(
        `Total rows: ${state.allRows.length - 1}, ` +
        `Filtered rows: ${state.filteredIndices.length}`
    );

    if (state.filteredIndices.length === 0) {
        alert('No rows match the review criteria!');
        return;
    }

    // Setup UI
    els.mainInterface.classList.remove('hidden');
    els.btnExport.classList.remove('hidden');
    els.rowCounter.classList.remove('hidden');

    // Try to restore previous session
    if (loadState()) {
        console.log('Restored previous session');
    } else {
        // Fresh start
        state.currentIndex = 0;
        saveState();
        loadRow(state.filteredIndices[state.currentIndex]);
    }
}

// Row Loading & Rendering
function loadRow(rowIndex) {
    const row = state.allRows[rowIndex];

    els.rowCounter.textContent = `Row ${state.currentIndex + 1} of ${state.filteredIndices.length}`;

    state.scale = 1;
    state.pointX = 0;
    state.pointY = 0;
    state.currentRotation = 0;
    state.showingUSG = true;
    updateImageTransform();

    const firstName = row[COL.FIRST_NAME] || '';
    const lastName = row[COL.LAST_NAME] || '';
    els.displayName.textContent = `${firstName} ${lastName}`;
    els.displayEmail.textContent = row[COL.EMAIL] || '';
    els.displayEdd.textContent = row[COL.EDD] || '';

    renderImage(row);

    els.inputReason.value = row[COL.REASON] || '';
    els.inputReasonOther.value = '';
    els.inputReasonOther.classList.add('hidden');
    els.calcResult.classList.add('hidden');
    els.calcWeeks.value = '';
    els.calcDays.value = '';
    els.calcDate.value = '';

    const currentStatus = row[COL.NOTE_ANIS] || '';
    updateStatusButtons(currentStatus);
}

function updateStatusButtons(status) {
    els.btnValid.classList.remove('active-valid');
    els.btnInvalid.classList.remove('active-invalid');

    if (status === 'Valid') {
        els.btnValid.classList.add('active-valid');
    } else if (status === 'Invalid') {
        els.btnInvalid.classList.add('active-invalid');
    }
}

function renderImage(row) {
    const url = state.showingUSG ? row[COL.PHOTO_USG] : row[COL.PHOTO_BOOK];
    const email = row[COL.EMAIL] || 'Unknown';

    els.imageError.classList.add('hidden');
    els.mainImage.classList.remove('hidden');
    els.imageLoading.classList.remove('hidden');
    els.btnRotateLeft.disabled = false;
    els.btnRotateRight.disabled = false;

    if (!url || url.trim() === '') {
        showImageError(email);
        els.imageLoading.classList.add('hidden');
        return;
    }

    const containerWidth = els.imageContainer.clientWidth;
    const containerHeight = els.imageContainer.clientHeight;
    const imgWidth = els.mainImage.offsetWidth;
    const imgHeight = els.mainImage.offsetHeight;

    state.pointX = (containerWidth - imgWidth) / 2;
    state.pointY = (containerHeight - imgHeight) / 2;
    state.scale = 1;

    updateImageTransform();

    els.mainImage.src = url;
    els.mainImage.onerror = () => showImageError(email);
}

function showImageError(email) {
    els.mainImage.classList.add('hidden');
    els.imageError.classList.remove('hidden');
    els.errorEmail.textContent = email;
    els.btnRotateLeft.disabled = true;
    els.btnRotateRight.disabled = true;
    els.imageLoading.classList.add('hidden');
}

function toggleImageSource() {
    state.showingUSG = !state.showingUSG;
    els.btnToggleImage.textContent = state.showingUSG ? "Switch to Book" : "Switch to USG";
    els.btnToggleImage.disabled = true;

    const row = state.allRows[state.filteredIndices[state.currentIndex]];

    state.scale = 1;
    state.pointX = 0;
    state.pointY = 0;
    state.currentRotation = 0;

    renderImage(row);
}

function rotateImage(deg) {
    state.currentRotation += deg;
    updateImageTransform();
}

function updateImageTransform() {
    els.imageWrapper.style.transform = `translate(${state.pointX}px, ${state.pointY}px) scale(${state.scale})`;
    els.mainImage.style.transform = `rotate(${state.currentRotation}deg)`;
}

// Logic & Actions
function calculateEDD() {
    const weeks = parseInt(els.calcWeeks.value) || 0;
    const days = parseInt(els.calcDays.value) || 0;
    const usgDateStr = els.calcDate.value;

    if (!usgDateStr) {
        alert('Please select a USG Date');
        return;
    }

    const usgDate = new Date(usgDateStr);
    const totalDaysGestational = (weeks * 7) + days;
    const daysRemaining = 280 - totalDaysGestational;

    const eddDate = new Date(usgDate);
    eddDate.setDate(eddDate.getDate() + daysRemaining);

    const formattedEDD = eddDate.toISOString().split('T')[0];

    els.resultDate.textContent = formattedEDD;
    els.calcResult.classList.remove('hidden');
}

function submitReview(status) {
    const rowIndex = state.filteredIndices[state.currentIndex];
    const row = state.allRows[rowIndex];

    row[COL.NOTE_ANIS] = status;

    let reason = els.inputReason.value;
    if (reason === 'Other') {
        reason = els.inputReasonOther.value;
    }
    row[COL.REASON] = reason;

    updateStatusButtons(status);
    saveState();
}

function nextRow() {
    if (state.currentIndex < state.filteredIndices.length - 1) {
        state.currentIndex++;
        loadRow(state.filteredIndices[state.currentIndex]);
    } else {
        els.overlay.classList.remove('hidden');
    }
}

function prevRow() {
    if (state.currentIndex > 0) {
        state.currentIndex--;
        loadRow(state.filteredIndices[state.currentIndex]);
    }
}

// Export
function exportData() {
    showLoading('Generating Excel...', { indeterminate: true });
    try {
        console.log("Starting export...");

        const exportCols = [COL.EMAIL, COL.NOTE_ANIS, COL.REASON];
        const cleanHeaders = exportCols.map((idx) => {
            const h = state.headers?.[idx];
            return h === null || h === undefined ? "" : String(h);
        });

        const colCount = exportCols.length;

        const toSafeCell = (cell) => {
            if (cell === null || cell === undefined) return "";
            const t = typeof cell;
            if (t === "string") return cell;
            if (t === "number") return Number.isFinite(cell) ? cell : String(cell);
            if (t === "boolean") return cell;
            if (t === "bigint") return cell.toString();
            if (cell instanceof Date) return isNaN(cell.getTime()) ? "" : cell.toISOString();

            if (t === "object") {
                try {
                    const json = JSON.stringify(cell, (key, value) => {
                        if (typeof value === "bigint") return value.toString();
                        return value;
                    });
                    if (typeof json === "string" && json.length > 5000) return json.slice(0, 5000);
                    return json ?? "";
                } catch (_) {
                    return "";
                }
            }

            return String(cell);
        };

        const reviewedSheet = [cleanHeaders];
        const unfilteredSheet = [cleanHeaders];

        for (let i = 1; i < state.allRows.length; i++) {
            const row = state.allRows[i];
            if (!Array.isArray(row)) continue;

            const cleanRow = new Array(colCount);
            for (let c = 0; c < colCount; c++) {
                cleanRow[c] = toSafeCell(row[exportCols[c]]);
            }

            const bi = (row[COL.NOTE_ANIS] || "").trim();

            if (bi !== "") {
                reviewedSheet.push(cleanRow);
            } else {
                unfilteredSheet.push(cleanRow);
            }
        }

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(
            wb,
            XLSX.utils.aoa_to_sheet(reviewedSheet),
            "Reviewed"
        );
        XLSX.utils.book_append_sheet(
            wb,
            XLSX.utils.aoa_to_sheet(unfilteredSheet),
            "Unfiltered_Rows"
        );

        XLSX.writeFile(wb, "Supermom_Review_Result.xlsx");
        console.log("Export complete.");
    } catch (e) {
        console.error("Export failed:", e);
        alert("Export failed! Check console for details.\n" + e.message);
    } finally {
        hideLoading();
    }
}

// Persistence
function saveState() {
    try {
        const changes = {};

        for (let i = 1; i < state.allRows.length; i++) {
            const row = state.allRows[i];
            if (!Array.isArray(row)) continue;

            const bi = (row[COL.NOTE_ANIS] || "").trim();
            const bl = (row[COL.REASON] || "").trim();

            if (bi !== "" || bl !== "") {
                changes[i] = {
                    bi: bi,
                    bl: bl
                };
            }
        }

        const payload = {
            currentIndex: state.currentIndex,
            changes: changes
        };

        sessionStorage.setItem("supermom_review", JSON.stringify(payload));
    } catch (e) {
        console.warn("Failed to save state", e);
        if (e.name === 'QuotaExceededError') {
            sessionStorage.removeItem("supermom_review");
            console.warn("Storage quota exceeded - state not saved");
        }
    }
}

function loadState() {
    try {
        const serialized = sessionStorage.getItem("supermom_review");
        if (!serialized) return false;

        const loaded = JSON.parse(serialized);

        if (!state.allRows.length || !state.filteredIndices.length) {
            return false;
        }

        if (loaded.changes) {
            for (const [rowIndex, data] of Object.entries(loaded.changes)) {
                const idx = parseInt(rowIndex);
                if (state.allRows[idx]) {
                    state.allRows[idx][COL.NOTE_ANIS] = data.bi;
                    state.allRows[idx][COL.REASON] = data.bl;
                }
            }
        }

        state.currentIndex = loaded.currentIndex || 0;

        if (state.currentIndex >= state.filteredIndices.length) {
            els.overlay.classList.remove("hidden");
            loadRow(state.filteredIndices[state.filteredIndices.length - 1]);
        } else {
            loadRow(state.filteredIndices[state.currentIndex]);
        }

        return true;
    } catch (e) {
        console.error("Failed to load state", e);
        sessionStorage.removeItem("supermom_review");
    }
    return false;
}

// INIT
init();