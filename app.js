/**
 * Tracking Analysis Tool - App Logic
 */

// Tab Switching
function switchTab(tabId) {
    // Hide all contents
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

    // Show target
    document.getElementById(`tab-${tabId}`).classList.add('active');
    
    // Activate button using specialized ID or fallback
    const btn = document.getElementById(`btn-tab-${tabId}`);
    if (btn) {
        btn.classList.add('active');
    } else {
        // Fallback for any legacy buttons
        const btns = document.getElementsByClassName('tab-btn');
        for (let b of btns) {
            if (b.getAttribute('onclick')?.includes(tabId)) {
                b.classList.add('active');
            }
        }
    }
}

// Global Event Listeners (Run on load)
document.addEventListener('DOMContentLoaded', () => {
    checkAuth();
    
    // Auto-init Exception Report if visible
    if (document.getElementById('exception-table-container')) {
        console.log("DOMContentLoaded: Initializing Exception Report Features...");
        loadExceptionMeta(); 
        renderExceptionTable();
    }

    // --- Old tools consolidated into Smart Workspace ---
    checkAdminUI();

    // 4. Smart Workspace Inputs -> Enter Key
    document.getElementById('smart-main-input')?.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') unifiedMainSearch();
    });
});

function checkAdminUI() {
    const urlParams = new URLSearchParams(window.location.search);
    const isAdmin = urlParams.has('admin');

    const mainTabs = document.getElementById('main-tabs');
    const userHeader = document.getElementById('user-mode-header');
    const navbar = document.querySelector('.navbar');

    if (isAdmin) {
        // Admin View
        document.body.classList.remove('user-mode');
        if (mainTabs) mainTabs.classList.remove('hidden');
        if (userHeader) userHeader.classList.add('hidden');
        if (navbar) navbar.style.display = 'block';
        
        switchTab('smart'); // Default to Unified Workspace for Admin
    } else {
        // Subordinate View (Locked to Import)
        document.body.classList.add('user-mode');
        if (mainTabs) mainTabs.classList.add('hidden');
        if (userHeader) userHeader.classList.remove('hidden');
        if (navbar) navbar.style.display = 'none';
        
        // Hide all tabs except Import section
        document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
        document.getElementById('tab-import').classList.add('active');
        
        // Ensure container is slim for mobile-first feel
        document.querySelector('main.container').classList.add('user-container');
    }

    const uploadIcon = document.getElementById('upload-icon-display');
    const uploadTitle = document.getElementById('upload-title-display');
    const uploadDesc = document.getElementById('upload-desc-display');
    const importInput = document.getElementById('import-upload');

    if (uploadIcon && uploadTitle && uploadDesc && importInput) {
        if (isAdmin) {
            // Admin: Excel + Images
            uploadIcon.textContent = "📂 / 📸";
            uploadTitle.textContent = "แตะเพื่อเลือกไฟล์ Excel หรือ รูปภาพ";
            uploadDesc.textContent = "รองรับ .xlsx, .xls และ รูปภาพ (เลือกได้หลายรูป)";
            importInput.setAttribute('accept', '.xlsx, .xls, .jpg, .jpeg, .png, .heic');
        } else {
            // User: Excel Only
            uploadIcon.textContent = "📂";
            uploadTitle.textContent = "แตะเพื่อเลือกไฟล์ Excel";
            uploadDesc.textContent = "รองรับ .xlsx, .xls (สำหรับพนักงาน)";
            importInput.setAttribute('accept', '.xlsx, .xls'); // Restrict native file picker
        }
    }
}

// Consolidated into unifiedQuickCheck() and unifiedGenerateRange()

// --- Unified Workspace Logic (Smart Tracking) ---
let lastGeneratedRange = [];

function toggleMainRangeInputs() {
    const toggle = document.getElementById('smart-main-range-toggle');
    const options = document.getElementById('smart-main-range-options');
    const bigBtn = document.getElementById('smart-main-range-btn-container');
    
    if (toggle.checked) {
        options.classList.remove('hidden');
        if(bigBtn) bigBtn.style.display = 'block';
    } else {
        options.classList.add('hidden');
        if(bigBtn) bigBtn.style.display = 'none';
        document.getElementById('smart-range-prev').value = '0';
        document.getElementById('smart-range-next').value = '0';
    }
}

async function unifiedMainSearch() {
    const inputEl = document.getElementById('smart-main-input');
    let input = inputEl.value.trim().toUpperCase();
    const isRange = document.getElementById('smart-main-range-toggle')?.checked;
    
    if (!input) {
        alert('กรุณากรอกเลขพัสดุ');
        return;
    }
    
    if (isRange) {
        unifiedGenerateRangeNew(input, inputEl);
    } else {
        unifiedSingleCheckNew(input, inputEl);
    }
}

function unifiedGenerateRangeNew(center, centerInput) {
    const qty = parseInt(document.getElementById('smart-range-qty')?.value) || 0;
    const prev = parseInt(document.getElementById('smart-range-prev').value) || 0;
    const next = parseInt(document.getElementById('smart-range-next').value) || 0;
    // qty adds equally to both sides if prev and next are both 0
    const effectivePrev = (prev === 0 && next === 0 && qty > 0) ? Math.floor(qty / 2) : prev;
    const effectiveNext = (prev === 0 && next === 0 && qty > 0) ? (qty - Math.floor(qty / 2)) : next;
    const finalPrev = Math.max(effectivePrev, prev);
    const finalNext = Math.max(effectiveNext, next);
    const resultArea = document.getElementById('smart-unified-results');
    const summaryArea = document.getElementById('smart-result-summary');

    const validation = TrackingUtils.validateTrackingNumber(center);
    if (!validation.isValid && validation.suggestion) {
        center = validation.suggestion;
        centerInput.value = center;
    } else if (!validation.isValid) {
        alert('เลขไม่ถูกต้อง: ' + validation.error);
        return;
    }

    const list = TrackingUtils.generateTrackingRange(center, finalPrev, finalNext);
    lastGeneratedRange = list.map(item => item.number);
    
    summaryArea.innerHTML = `<span class="badge badge-primary">${list.length} รายการ</span>`;

    let html = `
        <div style="padding:10px; background:#fff; border-bottom:1px solid #eee; position:sticky; top:0; z-index:5; display:flex; justify-content:space-between; align-items:center;">
            <small>สร้างจาก: ${center}</small>
            <button class="btn" style="padding:4px 8px; font-size:0.75rem;" onclick="copyUnifiedResults()">📋 Copy All</button>
        </div>
        <table>
            <thead>
                <tr>
                    <th style="width:40px;">#</th>
                    <th>เลขพัสดุ</th>
                    <th>สังกัด / สถานะ</th>
                </tr>
            </thead>
            <tbody>
    `;

    list.forEach((item, index) => {
        const rowStyle = item.isCenter ? 'style="background:#fff9c4; font-weight:bold;"' : '';
        const owner = typeof CustomerDB !== 'undefined' ? CustomerDB.get(item.number) : null;
        let ownerHtml = '-';
        if (owner) {
            ownerHtml = `<span style="color:#0d47a1; font-size:0.8rem; display:block; margin-bottom:4px;">👤 ${owner.name}</span>`;
        }

        let displayIndex = item.isCenter ? '<span style="color:#999; font-weight:bold;">-</span>' : Math.abs(item.offset);
        let indexStyle = item.offset < 0 ? 'color:#28a745;' : (item.offset > 0 ? 'color:#d32f2f;' : '');

        html += `
            <tr ${rowStyle}>
                <td style="${indexStyle} font-weight:bold;">${displayIndex}</td>
                <td class="unified-id-cell" style="font-family:monospace; font-size:0.95rem;">${TrackingUtils.formatTrackingNumber(item.number)}</td>
                <td>
                    ${ownerHtml}
                    <div style="display:flex; gap:5px;">
                        <button class="btn btn-neutral" style="padding:2px 6px; font-size:0.7rem; background:#f8f9fa; border:1px solid #ddd;" onclick="navigator.clipboard.writeText('${item.number}').then(()=>alert('คัดลอก ${item.number} ลง Clipboard แล้ว!'))">📋 คัดลอก</button>
                        <button class="btn btn-primary" style="padding:2px 6px; font-size:0.7rem;" onclick="window.open('https://track.thailandpost.co.th/?trackNumber=${item.number}&lang=th', '_blank')">🔍 เช็คสถานะ</button>
                        <button class="btn btn-success" style="padding:2px 6px; font-size:0.7rem;" onclick="stagingQuickReport(['${item.number}'], '${owner ? owner.name : ''}')">🚩 รายงาน</button>
                    </div>
                </td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    resultArea.innerHTML = html;
    lastGeneratedRange = list.map(item => item.number);
    // Show copy-all bar
    const copyBar = document.getElementById('smart-copy-all-bar');
    if (copyBar) copyBar.classList.remove('hidden');
}

function unifiedSingleCheckNew(input, inputEl) {
    const resultArea = document.getElementById('smart-unified-results');
    const summaryArea = document.getElementById('smart-result-summary');
    
    summaryArea.innerHTML = `<span class="badge badge-primary">ผลลัพธ์ 1 รายการ</span>`;
    
    const validation = TrackingUtils.validateTrackingNumber(input);
    const owner = typeof CustomerDB !== 'undefined' ? CustomerDB.get(input) : null;
    
    // Row 1 Logic
    let row1Html = '';
    let validTarget = input;
    if (validation.isValid) {
        row1Html = `<span style="color:#28a745; font-size:1.1rem;">✅ <strong>โครงสร้างเลขถูกต้อง</strong> (Check Digit Valid)</span>`;
    } else {
        if (validation.suggestion) {
            row1Html = `
                <div style="display:flex; flex-direction:column; gap:8px;">
                    <span style="color:#d32f2f; font-weight:bold; font-size:1.1rem;">❌ โครงสร้างเลขผิด: ${validation.error}</span>
                    <div style="background:#e8f5e9; padding:8px 12px; border-left:4px solid #28a745; border-radius:4px;">
                        <span style="color:#28a745; font-weight:bold;">👉 เลขที่ถูกต้องน่าจะเป็น: ${validation.suggestion}</span>
                        <button class="btn btn-primary" style="margin-left:10px; padding:4px 8px; font-size:0.8rem;" onclick="document.getElementById('smart-main-input').value='${validation.suggestion}'; unifiedMainSearch();">ใช้เลขนี้แทน</button>
                    </div>
                </div>`;
            validTarget = validation.suggestion;
        } else {
            row1Html = `<span style="color:#d32f2f; font-weight:bold; font-size:1.1rem;">❌ โครงสร้างเลขผิด: ${validation.error}</span>`;
        }
    }

    // Row 2 Logic
    let row2Html = '';
    if (owner) {
        // Safe access if timestamp exists
        const dateStr = owner.timestamp ? `(บันทึกเมื่อ ${new Date(owner.timestamp).toLocaleDateString('th-TH')})` : '';
        row2Html = `<span style="color:#0056b3; font-size:1.05rem;">🏢 พบในฐานข้อมูลของคุณ: <strong>${owner.name}</strong> <span style="font-size:0.85rem; color:#666;">${dateStr}</span></span>`;
    } else {
        row2Html = `<span style="color:#666;">⚪ ไม่พบประวัติผู้ส่งในฐานข้อมูลของเครื่องคุณ</span>`;
    }

    // Prepare UI Container
    resultArea.innerHTML = `
        <div style="padding: 20px; border: 1px solid #ddd; border-radius: 8px; background: #fff; text-align: left; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <h4 style="margin: 0 0 15px 0; font-size: 1.2rem; display:flex; justify-content:space-between; align-items:center;">
                <span>ผลการตรวจสอบ: <span style="font-family:monospace; color:#0d47a1; background:#f0f7ff; padding:2px 6px; border-radius:4px;">${TrackingUtils.formatTrackingNumber(input)}</span></span>
            </h4>
            
            <div style="display:flex; flex-direction:column; gap: 15px;">
                <!-- Row 1: Check Digit -->
                <div style="padding: 15px; border: 1px solid #eee; border-radius: 6px; background:#fafafa; display:flex; align-items:center;">
                    ${row1Html}
                </div>
                
                <!-- Row 2: Customer DB -->
                <div style="padding: 15px; border: 1px solid #eee; border-radius: 6px; background:#fafafa; display:flex; align-items:center;">
                    ${row2Html}
                </div>

                <!-- Row 3: Copy per item -->
                <div style="padding: 12px 15px; border: 1px solid #eee; border-radius: 6px; background:#fafafa; display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap;">
                    <span style="font-size:0.95rem; font-family:monospace; color:#333; word-break:break-all;">${validTarget}</span>
                    <div style="display:flex; gap:8px; flex-wrap:wrap;">
                        <button class="btn btn-neutral" style="padding:5px 12px; font-size:0.85rem; background:#fff; border:1px solid #ccc; white-space:nowrap;" onclick="navigator.clipboard.writeText('${validTarget}').then(()=>{ this.textContent='\u2705 \u0e04\u0e31\u0e14\u0e25\u0e2d\u0e01\u0e41\u0e25\u0e49\u0e27!'; setTimeout(()=>this.textContent='\ud83d\udccb \u0e2a\u0e33\u0e40\u0e19\u0e32\u0e40\u0e25\u0e02', 1500); })">📋 สำเนาเลข</button>
                        <button class="btn btn-success" style="padding:5px 12px; font-size:0.85rem; white-space:nowrap;" onclick="window.open('https://track.thailandpost.co.th/?trackNumber=${validTarget}', '_blank')">🔍 เปิดเว็บ Track&Trace</button>
                        <button class="btn btn-success" style="padding:5px 12px; font-size:0.85rem; white-space:nowrap;" onclick="stagingQuickReport(['${validTarget}'], '${owner ? owner.name : ''}')">🚩 รายงานผลตกหล่น</button>
                    </div>
                </div>
            </div>
        </div>
    `;

    // Show copy-all bar (for single item too)
    lastGeneratedRange = [validTarget];
    const copyBar = document.getElementById('smart-copy-all-bar');
    if (copyBar) copyBar.classList.remove('hidden');
}

// Copy all numbers from lastGeneratedRange
function copyAllUnifiedNumbers() {
    if (!lastGeneratedRange || lastGeneratedRange.length === 0) {
        alert('ไม่มีข้อมูลเลขใเค๊อพ');
        return;
    }
    const text = lastGeneratedRange.join('\n');
    navigator.clipboard.writeText(text).then(() => {
        alert(`คัดลอก ${lastGeneratedRange.length} รายการเรียบร้อยแล้ว!`);
    });
}

// OCR upload handler inside the Track & Trace section
async function handleTrackOcrUpload(files) {
    if (!files || files.length === 0) return;
    const statusEl = document.getElementById('track-ocr-status');
    statusEl.textContent = `กำลังสแกน ${files.length} รูป...`;

    let combinedText = '';
    try {
        for (let i = 0; i < files.length; i++) {
            statusEl.textContent = `OCR รูป ${i + 1}/${files.length}...`;
            const worker = await Tesseract.createWorker('tha+eng');
            const { data: { text } } = await worker.recognize(files[i]);
            await worker.terminate();
            combinedText += '\n' + text;
        }

        const numbers = TrackingUtils.extractTrackingNumbers(combinedText);
        if (numbers.length === 0) {
            statusEl.textContent = '⚠️ ไม่พบเลขในภาพ';
            return;
        }

        if (numbers.length === 1) {
            // Single number -> put in main input and search
            document.getElementById('smart-main-input').value = numbers[0];
            statusEl.textContent = `✅ พบเลข: ${numbers[0]} กำลังค้นหา...`;
            unifiedMainSearch();
        } else {
            // Multiple numbers -> show results list
            lastGeneratedRange = numbers;
            const resultArea = document.getElementById('smart-unified-results');
            const summaryArea = document.getElementById('smart-result-summary');
            summaryArea.innerHTML = `<span class="badge badge-primary">${numbers.length} รายการ (OCR)</span>`;

            let html = `<div style="padding:10px; background:#fff; border-bottom:1px solid #eee; position:sticky; top:0; z-index:5;"><small>สแกนจากรูป ${files.length} รูป พบ ${numbers.length} เลข</small></div><table><thead><tr><th>#</th><th>เลขพัสดุ</th><th>สถานะ</th></tr></thead><tbody>`;
            numbers.forEach((num, i) => {
                const owner = typeof CustomerDB !== 'undefined' ? CustomerDB.get(num) : null;
                html += `<tr>
                    <td style="padding-left:10px;">${i + 1}</td>
                    <td style="font-family:monospace;">${TrackingUtils.formatTrackingNumber(num)}</td>
                    <td>
                        ${owner ? `<span style="color:#0d47a1; font-size:0.8rem;">👤 ${owner.name}</span>` : ''}
                        <div style="display:flex; gap:5px; margin-top:4px;">
                            <button class="btn btn-neutral" style="padding:2px 6px; font-size:0.7rem;" onclick="navigator.clipboard.writeText('${num}').then(()=>alert('คัดลอก ${num} แล้ว!'))">📋 คัดลอก</button>
                            <button class="btn btn-primary" style="padding:2px 6px; font-size:0.7rem;" onclick="window.open('https://track.thailandpost.co.th/?trackNumber=${num}', '_blank')">🔍 เช็ค</button>
                        </div>
                    </td>
                </tr>`;
            });
            html += '</tbody></table>';
            resultArea.innerHTML = html;
            const copyBar = document.getElementById('smart-copy-all-bar');
            if (copyBar) copyBar.classList.remove('hidden');

            // --- QMS AUTO-FILL LINK ---
            const qmsTextarea = document.getElementById('qms-import-text');
            const qmsSection = document.getElementById('qms-staging-section');
            const qmsPanel = document.getElementById('qms-staging-panel');
            if (qmsTextarea && qmsSection) {
                qmsTextarea.value = numbers.join('\n');
                
                if (typeof processQmsImport === 'function') {
                    qmsSection.style.display = 'block';
                    if (qmsPanel) qmsPanel.style.display = 'block';
                    const chevron = document.getElementById('qms-staging-chevron');
                    if (chevron) chevron.textContent = '▲ ปิด';

                    qmsSection.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    
                    qmsTextarea.style.transition = "background-color 0.5s";
                    qmsTextarea.style.backgroundColor = "#fff9c4";
                    
                    setTimeout(() => {
                        processQmsImport();
                        qmsTextarea.style.backgroundColor = "";
                        
                        // If all items fall into exactly 1 group, auto-draft it to save a click!
                        if (typeof qmsStagingGroups !== 'undefined') {
                            const groupKeys = Object.keys(qmsStagingGroups);
                            if (groupKeys.length === 1) {
                                setTimeout(() => {
                                    if (typeof draftReportFromGroup === 'function') {
                                        draftReportFromGroup(groupKeys[0]);
                                    }
                                }, 500);
                            }
                        }
                    }, 800);
                }
            }
            // --------------------------
        }
        statusEl.textContent = `✅ เสร็จ พบ ${numbers.length} เลข`;
    } catch (err) {
        console.error(err);
        statusEl.textContent = '❌ เกิดข้อผิดพลาด: ' + err.message;
    }
}

// Gap analysis features removed per user request
// --- Universal Import Logic (Excel & Image/OCR) ---

let currentImportedBatches = []; // To store analyzed data before saving
let rawTrackingData = []; // Store ALL raw items (Cumulative)
let importedFileCount = 0; // Track number of files uploaded (for limit)

function clearImportData() {
    if (rawTrackingData.length === 0) return;
    if (!confirm('ต้องการล้างข้อมูลนำเข้าทั้งหมดหรือไม่?')) return;

    rawTrackingData = [];
    currentImportedBatches = [];
    importedFileCount = 0; // Reset Limit
    document.getElementById('import-preview').classList.add('hidden');
    document.getElementById('upload-status').innerText = 'ล้างข้อมูลเรียบร้อย (Ready)';
    // Reset file input
    document.getElementById('import-upload').value = '';
}


function handleFileUpload(event) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    // --- Role-Based Limit Check ---
    const urlParams = new URLSearchParams(window.location.search);
    const isAdmin = urlParams.has('admin');

    if (!isAdmin) {
        // Normal User Limit: 2 Files Max (Cumulative)
        if ((importedFileCount + files.length) > 2) {
            alert(`⚠️ จำกัดการอัปโหลดสูงสุด 2 ไฟล์สำหรับบัญชีทั่วไป\n(คุณอัปโหลดไปแล้ว ${importedFileCount} ไฟล์, พยายามเพิ่มอีก ${files.length} ไฟล์)\n\nกรุณาล้างข้อมูลเก่าก่อนหากต้องการเริ่มใหม่`);
            document.getElementById('import-upload').value = ''; // Reset input to allow re-selection
            return;
        }
    }

    // Increment Count
    importedFileCount += files.length;

    // Check type
    if (files[0].type.includes('image')) {
        handleImageImport(files); // Pass FileList (Admin Only check is inside)
    } else {
        // Excel - Support Multiple
        Array.from(files).forEach(file => {
            handleExcelImport(file);
        });
    }
}

function handleExcelImport(file) {
    document.getElementById('upload-status').innerText = `กำลังอ่านไฟล์ Excel...`;
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const newItems = [];
        const regex = /([A-Z]{2})(\d{9})([A-Z]{2})/i;

        jsonData.forEach(row => {
            if (row.length >= 3) {
                if (row[2] && typeof row[2] === 'string') {
                    const match = row[2].match(regex);
                    if (match) {
                        const price = parseFloat(row[3]) || 0;
                        let weight = row[4] || '-';
                        let originalWeight = weight;
                        let hasDiscrepancy = false;

                        // Apply Package A3 logic if price matches
                        if (price > 0 && typeof TrackingUtils !== 'undefined' && TrackingUtils.getWeightFromPriceA3) {
                            const calcWeight = TrackingUtils.getWeightFromPriceA3(price);
                            if (calcWeight !== '-') {
                                weight = calcWeight;
                                if (originalWeight !== '-' && originalWeight.toString().trim() !== weight.toString().trim()) {
                                    hasDiscrepancy = true;
                                }
                            }
                        }

                        newItems.push({
                            number: match[0].toUpperCase(),
                            price: price,
                            weight: weight,
                            hasDiscrepancy: hasDiscrepancy,
                            originalWeight: originalWeight
                        });
                    }
                }
            }
        });

        if (newItems.length === 0) {
            alert('ไม่พบเลขพัสดุในไฟล์ Excel (ตรวจสอบคอลัมน์ C)');
            document.getElementById('upload-status').innerText = 'ไม่พบเลขพัสดุ';
            return;
        }

        // Cumulative Append
        rawTrackingData.push(...newItems);

        document.getElementById('upload-status').innerText = `อ่าน Excel สำเร็จ! เพิ่ม ${newItems.length} รายการ (รวม ${rawTrackingData.length})`;
        analyzeImportedRanges(rawTrackingData);
    };
    reader.readAsArrayBuffer(file);
}

// Updated to handle Multiple Images
async function handleImageImport(files) {
    // Check if Admin
    const urlParams = new URLSearchParams(window.location.search);
    if (!urlParams.has('admin')) {
        alert('ฟีเจอร์ "นำเข้ารูปภาพ" สงวนสิทธิ์สำหรับ Admin เท่านั้น\n(พนักงานทั่วไปกรุณาใช้ไฟล์ Excel)');
        // Clear input safely
        const input = document.getElementById('import-upload');
        if (input) input.value = '';
        document.getElementById('upload-status').innerText = 'Access Denied (Admin Only)';
        return;
    }

    const statusEl = document.getElementById('upload-status');
    statusEl.innerText = `Preparing OCR for ${files.length} images...`;

    let combinedText = "";

    try {
        for (let i = 0; i < files.length; i++) {
            statusEl.innerText = `OCR Scanning Image ${i + 1}/${files.length}...`;
            const file = files[i];

            const useEngOnly = document.getElementById('admin-ocr-eng-only') && document.getElementById('admin-ocr-eng-only').checked;
            const lang = useEngOnly ? 'eng' : 'tha+eng';
            const worker = await Tesseract.createWorker(lang);
            const { data: { text } } = await worker.recognize(file);
            await worker.terminate();

            combinedText += "\n" + text;
        }

        statusEl.innerText = "Analyzing extracted text...";
        console.log('OCR Raw:', combinedText);

        const newItems = [];
        const rawLines = combinedText.split('\n').map(l => l.trim()).filter(l => l);

        let currentPrefixBody = null;

        for (let i = 0; i < rawLines.length; i++) {
            let line = rawLines[i].toUpperCase();

            // Strategy 1: Remove all spaces and check for Standard ID
            const cleanLine = line.replace(/\s+/g, '');
            const stdMatch = cleanLine.match(/([A-Z]{2})(\d{9})([A-Z]{2})/);
            if (stdMatch) {
                newItems.push({ number: stdMatch[0], price: 0, weight: 'OCR-Std' });
                currentPrefixBody = stdMatch[1] + stdMatch[2].substring(0, 4);
                continue;
            }
        }

        // --- NEW: Use Smart Extraction (Robust) ---
        // First try the new handwritten table parser
        const tableItems = TrackingUtils.extractHandwrittenTable(combinedText);
        
        // If the table extraction found prices, use that. Otherwise fallback to basic list mapping
        const hasPrices = tableItems.some(item => item.price > 0);
        
        if (hasPrices) {
            newItems.push(...tableItems);
        } else {
             // Fallback to simple extraction
             const extractedSmart = TrackingUtils.extractTrackingNumbers(combinedText);
             const contextData = TrackingUtils.extractTrackingWithContext(combinedText);
             const contextMap = new Map();
             contextData.forEach(c => contextMap.set(c.trackingNumber, c));
             
             extractedSmart.forEach(num => {
                 const context = contextMap.get(num) || null;
                 newItems.push({ number: num, price: 0, weight: 'OCR-Smart', context: context });
             });
        }

        // De-duplicate locally (within this batch) logic if needed, 
        // but analyzeImportedRanges handles grouping.

        if (newItems.length === 0) {
            // Price Fallback...
            const prices = TrackingUtils.extractPrices(combinedText);
            if (prices.length > 0) {
                if (confirm(`ไม่พบเลขพัสดุในภาพ แต่พบยอดเงิน ${prices.length} รายการ\nต้องการสร้างรายงานแยกตามราคาสินค้าแทนหรือไม่?`)) {
                    const summary = TrackingUtils.summarizePrices(prices);
                    const priceRanges = summary.groupings.map(g => ({
                        start: 'PRICE-ONLY',
                        end: 'PRICE-ONLY',
                        count: g.count,
                        price: g.price,
                        weight: 'N/A',
                        total: g.total,
                        items: [] // No tracking IDs
                    }));

                    // Note: Price-only mode replaces everything as it's a different mode
                    rawTrackingData = [];
                    currentImportedBatches = priceRanges;
                    renderImportResult(priceRanges);
                    statusEl.innerText = `วิเคราะห์ยอดเงินสำเร็จ (${summary.totalCount} รายการ)`;
                    return;
                }
            }

            alert('ไม่พบเลขพัสดุในรูปภาพนี้');
            statusEl.innerText = 'ไม่พบข้อมูลในภาพล่าสุด';
            return;
        }

        // Cumulative Append
        rawTrackingData.push(...newItems);

        statusEl.innerText = `OCR เสร็จสิ้น! เพิ่ม ${newItems.length} รายการ (รวม ${rawTrackingData.length})`;
        analyzeImportedRanges(rawTrackingData);

    } catch (err) {
        console.error(err);
        alert('เกิดข้อผิดพลาดในการอ่านรูปภาพ: ' + err.message);
        statusEl.innerText = "Error";
    }
}

function analyzeImportedRanges(trackingList) {
    if (trackingList.length === 0) return;

    // Remove duplicates from total list
    const uniqueMap = new Map();
    const discrepancyMap = new Map();
    trackingList.forEach(item => {
        uniqueMap.set(item.number, item);
        if (item.hasDiscrepancy) {
            discrepancyMap.set(item.number, item);
        }
    });
    const uniqueList = Array.from(uniqueMap.values());
    const discrepanciesList = Array.from(discrepancyMap.values());

    // Sort logic...
    // --- STEP 1: GAP ANALYSIS (100% Integrity Check) ---
    // Sort by Number only (Independent of price) to find true gaps
    const sortedForGaps = [...uniqueList].sort((a, b) => a.number.localeCompare(b.number));
    const missingItems = []; // { start, end, count, example }

    const parse = (str) => {
        const m = str.match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
        return m ? { full: str, prefix: m[1], body: parseInt(m[2]), check: m[3], suffix: m[4] } : null;
    };

    let prevGap = parse(sortedForGaps[0].number);

    for (let i = 1; i < sortedForGaps.length; i++) {
        const currGap = parse(sortedForGaps[i].number);
        if (!currGap || !prevGap) continue;

        // Check if same series (Prefix + Suffix)
        if (currGap.prefix === prevGap.prefix && currGap.suffix === prevGap.suffix) {
            const diff = currGap.body - prevGap.body;
            if (diff > 1) {
                // FOUND GAP: diff=2 means 1 missing (e.g. 1, 3 -> missing 2)
                // Missing range: (prev+1) to (curr-1)
                const startMissing = prevGap.body + 1;
                const endMissing = currGap.body - 1;
                const countMissing = endMissing - startMissing + 1;

                // Reconstruct ID for display
                const exampleID = `${currGap.prefix}${startMissing.toString().padStart(8, '0')}X${currGap.suffix}`;

                missingItems.push({
                    prefix: currGap.prefix,
                    suffix: currGap.suffix,
                    startBody: startMissing,
                    endBody: endMissing,
                    count: countMissing,
                    example: exampleID
                });
            }
        }
        prevGap = currGap;
    }

    // --- STEP 2: GROUPING FOR REPORT (User Preference: Price) ---
    // Sort by Price (Asc) then Number (Asc)
    uniqueList.sort((a, b) => {
        if (a.price !== b.price) {
            return a.price - b.price;
        }
        return a.number.localeCompare(b.number);
    });

    const rawRanges = [];

    // Helper to parse with price (reusing structure)
    const parseFull = (item) => {
        const p = parse(item.number);
        if (p) { p.price = item.price; p.weight = item.weight; p.hasDiscrepancy = item.hasDiscrepancy; p.originalWeight = item.originalWeight; p.context = item.context; }
        return p;
    };

    let start = parseFull(uniqueList[0]);
    let prev = start;
    let currentList = [uniqueList[0]];

    for (let i = 1; i < uniqueList.length; i++) {
        const curr = parseFull(uniqueList[i]);
        if (!curr) continue;

        const isContinuous = (
            curr.prefix === prev.prefix &&
            curr.suffix === prev.suffix &&
            curr.body === prev.body + 1 &&
            curr.price === prev.price &&
            curr.weight === prev.weight
        );

        if (isContinuous) {
            currentList.push(uniqueList[i]);
            prev = curr;
        } else {
            rawRanges.push({
                start: start.full,
                end: prev.full,
                count: currentList.length,
                price: start.price,
                weight: start.weight,
                items: currentList.map(x => x.number),
                contexts: currentList.map(x => x.context).filter(c => c !== null)
            });
            start = curr;
            prev = curr;
            currentList = [uniqueList[i]];
        }
    }
    rawRanges.push({
        start: start.full,
        end: prev.full,
        count: currentList.length,
        price: start.price,
        weight: start.weight,
        items: currentList.map(x => x.number),
        contexts: currentList.map(x => x.context).filter(c => c !== null)
    });

    // Virtual Optimization
    const optimizedRanges = TrackingUtils.virtualOptimizeRanges(rawRanges);

    currentImportedBatches = optimizedRanges;
    renderImportResult(optimizedRanges, missingItems, discrepanciesList);
}

function renderImportResult(ranges, missingItems = [], discrepancies = []) {
    const preview = document.getElementById('import-preview');
    const summary = document.getElementById('import-summary');
    const details = document.getElementById('import-details');

    preview.classList.remove('hidden');

    const totalItems = ranges.reduce((acc, r) => acc + r.count, 0);
    const grandTotal = ranges.reduce((acc, r) => acc + (r.total || (r.count * r.price)), 0);

    // --- GAP ALERT SECTION ---
    let gapHtml = '';
    let gapTableRows = ''; // To show in table as well

    if (missingItems.length > 0) {
        const totalMissing = missingItems.reduce((acc, m) => acc + m.count, 0);

        // Helper to format with Check Digit
        const formatID = (prefix, body, suffix) => {
            const bodyStr = body.toString().padStart(8, '0');
            const cd = TrackingUtils.calculateS10CheckDigit(bodyStr);
            return `${prefix}${bodyStr}${cd}${suffix}`;
        };

        // Generate Alert List
        let listHtml = missingItems.map(m => {
            const startID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.startBody, m.suffix));
            const endID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.endBody, m.suffix));

            const rangeText = (m.count === 1)
                ? startID // Single ID
                : `${startID} - ${endID}`; // Range

            return `<li>${rangeText} (${m.count} รายการ)</li>`;
        }).join('');

        gapHtml = `
            <div class="result-error" style="margin-top:15px; padding:15px; border:2px solid #ff4444; background:#ffebeb;">
                <h3 style="margin-top:0; color:#cc0000;">⚠️ ตรวจพบเลขพัสดุข้ามไป (GAP DETECTED)</h3>
                <p>ระบบพบว่าข้อมูลไม่ต่อเนื่อง <strong>หายไป ${totalMissing} รายการ</strong> ดังนี้:</p>
                <ul style="margin-bottom:0;">${listHtml}</ul>
                <div style="font-size:0.9rem; color:#666; margin-top:10px;">รายการที่หายไปถูกเพิ่มลงในตารางด้านล่างเพื่ออ้างอิงแล้ว (สีแดง)</div>
            </div>
        `;

        // Generate Table Rows for Gaps
        gapTableRows = missingItems.map((m, index) => {
            const startID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.startBody, m.suffix));
            const endID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.endBody, m.suffix));

            const rangeText = (m.count === 1)
                ? startID
                : `${startID} - ${endID}`;

            return `
                <tr style="background-color:#ffebeb; color:#d32f2f; border-bottom:1px solid #ffcdd2;">
                    <td style="padding:10px; font-weight:bold;">
                        <div class="line-flex">
                            <span>❌ รายการที่หายไป (Missing)</span>
                            <span class="mobile-stats" style="color:#d32f2f;">${m.count} รายการ</span>
                        </div>
                        <span style="font-size:0.9em; font-family:monospace;">${rangeText}</span>
                    </td>
                    <td class="col-qty" style="text-align:right; padding:10px;">${m.count}</td>
                    <td class="col-price" style="text-align:right; padding:10px;">-</td>
                    <td class="col-total" style="text-align:right; padding:10px;">-</td>
                </tr>
            `;
        }).join('');
    } else {
        gapHtml = `
            <div class="result-success" style="margin-top:15px; padding:10px; border:1px solid #4caf50; background:#e8f5e9;">
                <strong>✅ ครบถ้วน 100% (No Gaps)</strong> - เลขพัสดุเรียงต่อเนื่องกันสมบูรณ์
            </div>
        `;
    }

    // --- DISCREPANCY ALERT SECTION ---
    let discrepancyHtml = '';
    if (discrepancies && discrepancies.length > 0) {
        discrepancyHtml = `
            <div class="result-error" style="margin-top:15px; padding:10px; border:2px solid #ff9800; background:#fff3e0; color:#e65100;">
                <h4 style="margin:0 0 5px 0;">⚠️ พบข้อสังเกตจากไฟล์นำเข้า (Inconsistencies)</h4>
                <p style="margin:0; font-size:0.95rem;">
                    พบข้อมูลน้ำหนักที่ระบุในไฟล์ <strong>ไม่สัมพันธ์กับราคาตามตารางค่าบริการ</strong> จำนวน <strong>${discrepancies.length} รายการ</strong><br>
                    <small>เช่น ในไฟล์ระบุ ${discrepancies[0].originalWeight} แต่ราคา ${discrepancies[0].price} บาท (ระบบได้ปรับแก้เป็น ${discrepancies[0].weight} อัตโนมัติแล้ว)</small>
                </p>
            </div>
        `;
    }

    summary.innerHTML = `
        <strong>📊 สรุปผลการวิเคราะห์ (Virtual Optimization)</strong><br>
        จำนวนทั้งหมด: ${totalItems.toLocaleString()} ชิ้น<br>
        ยอดเงินรวม: <span style="font-size:1.2rem; color:#d63384; font-weight:bold;">${grandTotal.toLocaleString()} บาท</span>
        ${gapHtml}
        ${discrepancyHtml}
    `;

    // Generate Receipt-style Table
    let html = `
        <div style="background:white; padding:20px; border:1px solid #ddd; box-shadow:0 2px 5px rgba(0,0,0,0.05); font-family:'Courier New', monospace;">
            <h4 style="text-align:center; border-bottom:1px dashed #ccc; padding-bottom:10px; margin-bottom:10px;">ใบสรุปรายการ (Optimized Report)</h4>
             <div style="font-size:0.8rem; color:red; text-align:center; margin-bottom:5px;">
                *รายการถูกจัดเรียงใหม่ตามราคาน้อย-มาก (Virtual Mapping)
            </div>
            
            <table style="width:100%; border-collapse: collapse;">
                <thead>
                    <tr style="border-bottom:2px solid #000;">
                        <th style="text-align:left; padding:5px;">รายการ (Description)</th>
                        <th class="col-qty" style="text-align:right; padding:5px; white-space:nowrap;">จำนวน (Qty)</th>
                        <th class="col-price" style="text-align:right; padding:5px; white-space:nowrap;">ราคา/ชิ้น</th>
                        <th class="col-total" style="text-align:right; padding:5px; white-space:nowrap;">รวม (Total)</th>
                    </tr>
                </thead>
                <tbody>
                    ${gapTableRows} <!-- Insert Missing Items at Top -->
    `;

    ranges.forEach((r, idx) => {
        const rowTotal = r.total || (r.count * r.price);
        const rowTotalStr = rowTotal.toLocaleString('en-US', { minimumFractionDigits: 2 });

        // Logic: Show total only if count > 1
        const displayTotal = (r.count > 1) ? ` | ${rowTotalStr}` : '';

        html += `
            <tr style="border-bottom:1px dashed #eee;">
                <td style="padding:10px 0; vertical-align:top; width:100%;">
                    <!-- Line 1: Title + Qty + Total (Mobile) -->
                    <div class="line-flex">
                        <strong>${idx + 1}. EMS ราคา ${r.price} บาท</strong>
                        <span class="mobile-stats" style="color:#d63384; font-weight:bold;">${r.count} ชิ้น${displayTotal}</span>
                    </div>
                    
                    <!-- Line 2: Range Only (Mobile) -->
                    <div class="line-flex">
                        <span style="color:#0056b3; font-weight:bold; overflow-wrap:break-word; max-width:100%;">
                            ${r.start === r.end
                ? TrackingUtils.formatTrackingNumber(r.start)
                : `${TrackingUtils.formatTrackingNumber(r.start)} - ${TrackingUtils.formatTrackingNumber(r.end)}`}
                        </span>
                    </div>

                    <small style="color:#666;">น้ำหนัก (Weight): ${r.weight}</small>
                </td>
                <td class="col-qty" style="text-align:right; vertical-align:top; padding-top:10px;">${r.count}</td>
                <td class="col-price" style="text-align:right; vertical-align:top; padding-top:10px;">@${r.price}</td>
                <td class="col-total" style="text-align:right; vertical-align:top; padding-top:10px; font-weight:bold;">${rowTotalStr}</td>
            </tr>
        `;
    });

    html += `
                    </tbody>
                    <tfoot>
                        <tr style="border-top:2px solid #000; border-bottom:2px solid #000;">
                            <td colspan="4" style="padding:10px;">
                                <div class="mobile-grand-total">
                                    <span style="font-weight:bold;">รวมทั้งสิ้น (Grand Total)</span>
                                    <span style="font-weight:bold; font-size:1.1rem;">${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                                </div>
                            </td>
                        </tr>
                    </tfoot>
                </table>
            <div style="font-size:0.8rem; color:#999; text-align:center; margin-top:5px; font-style:italic;">
                (บนมือถือ: จำนวนและยอดเงินจะแสดงที่มุมขวาของรายการ)
            </div>
        </div>
    `;
    details.innerHTML = html;

    // Auto-fill batch name suggestion
    const today = new Date().toLocaleDateString('th-TH');
    document.getElementById('import-batch-name').value = `Import ${today} (${totalItems} items)`;

    // AUTO SAVE (Triggers immediately)
    setTimeout(() => {
        saveImportedBatch(true);
    }, 500);
}

function saveImportedBatch(isAuto = false) {
    if (currentImportedBatches.length === 0) return;

    const name = document.getElementById('import-batch-name').value.trim();
    const type = document.getElementById('import-batch-type').value;

    if (!name) {
        alert('กรุณาระบุชื่อกลุ่มข้อมูล (Batch Name)');
        return;
    }

    if (!isAuto) {
        if (!confirm(`ยืนยันบันทึกข้อมูล ${currentImportedBatches.length} ช่วงรวมกันเป็น 1 ชุดข้อมูล?`)) return;
    }

    // Save Logic (Optimized)
    const allItemsToSave = [];
    currentImportedBatches.forEach(r => {
        // Since we optimized, we generate from start-end.
        // Assuming contiguous S10. 
        const list = TrackingUtils.generateTrackingRange(r.start, 0, r.count - 1);
        list.forEach(i => allItemsToSave.push(i.number));
    });

    // Metadata
    const rangesMeta = currentImportedBatches.map(r => ({
        start: r.start,
        end: r.end,
        count: r.count,
        price: r.price,
        weight: r.weight,
        total: r.total || (r.count * r.price)
    }));

    const batchInfo = {
        name: name,
        type: type,
        contract: 'Imported',
        timestamp: new Date().getTime(),
        ranges: rangesMeta
    };

    // Save
    const result = CustomerDB.addBatch(batchInfo, allItemsToSave);
    const addedCount = typeof result === 'object' ? result.count : result;
    const newBatchId = typeof result === 'object' ? result.id : null;

    if (isAuto) {
        // alert(`✅ ระบบบันทึกข้อมูลอัตโนมัติเรียบร้อย!\n(Optimized ${rangesMeta.length} Groups)`);
        // Silent or small notification? User wants to SEE it.
    } else {
        alert(`บันทึกเรียบร้อย!`);
    }

    // Reset inputs
    document.getElementById('excel-upload').value = '';
    document.getElementById('import-preview').classList.add('hidden');
    currentImportedBatches = [];

    // DIRECTLY VIEW THE REPORT
    if (newBatchId && typeof loadBatchToView === 'function') {
        loadBatchToView(newBatchId);
    } else {
        // Fallback
        switchTab('customer');
        if (typeof renderDBTable === 'function') renderDBTable();
    }
}

function toggleImportHistory() {
    const sec = document.getElementById('import-history-section');
    if (sec.classList.contains('hidden')) {
        sec.classList.remove('hidden');
        renderImportHistory();
    } else {
        sec.classList.add('hidden');
    }
}

function renderImportHistory() {
    const tbody = document.querySelector('#import-history-table tbody');
    if (!tbody) return;

    // Check Admin Mode
    const urlParams = new URLSearchParams(window.location.search);
    const isAdmin = urlParams.has('admin');

    tbody.innerHTML = '';
    const batches = CustomerDB.getBatches();
    const list = Object.values(batches).sort((a, b) => b.timestamp - a.timestamp);

    if (list.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">ไม่พบประวัติการนำเข้า</td></tr>';
        return;
    }

    list.forEach(item => {
        const dateStr = new Date(item.timestamp).toLocaleString('th-TH');
        const tr = document.createElement('tr');

        let deleteBtn = '';
        if (isAdmin) {
            deleteBtn = `
                <button class="btn btn-danger" style="padding:4px 8px; font-size:0.8rem; margin-left:5px;" 
                    onclick="deleteHistoryItem('${item.id}', '${item.name}')">🗑️ ลบ</button>
            `;
        }

        tr.innerHTML = `
            <td>${dateStr}</td>
            <td>
                <strong>${item.name}</strong><br>
                <small class="text-muted">${item.type || '-'}</small>
            </td>
            <td>${item.count.toLocaleString()}</td>
            <td>
                <button class="btn btn-primary" style="padding:4px 8px; font-size:0.8rem;" 
                    onclick="loadBatchToView('${item.id}')">🔎 ดู</button>
                ${deleteBtn}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function deleteHistoryItem(batchId, batchName) {
    if (confirm(`ยืนยันการลบข้อมูล (จากประวัติ Import)?\n\nชุดข้อมูล: ${batchName}`)) {
        CustomerDB.deleteBatch(batchId);
        // Refresh this table
        renderImportHistory();
        // Also refresh main DB table if it exists (keep in sync)
        if (typeof updateDbViews === 'function') updateDbViews();
        else if (typeof renderDBTable === 'function') renderDBTable();
    }
}

// --- Snapshot Logic ---
function toggleSnapshotList() {
    const list = document.getElementById('snapshot-list');
    if (list.classList.contains('hidden')) {
        list.classList.remove('hidden');
        renderSnapshotList();
    } else {
        list.classList.add('hidden');
    }
}

function renderSnapshotList() {
    const list = document.getElementById('snapshot-list');
    const snapshots = CustomerDB.getSnapshots();

    if (snapshots.length === 0) {
        list.innerHTML = '<li>ไม่มีจุดย้อนกลับ (No Snapshots)</li>';
        return;
    }

    let html = '';
    snapshots.forEach(s => {
        const time = new Date(s.timestamp).toLocaleString('th-TH');
        html += `
            <li style="margin-bottom:5px;">
                <strong>${time}</strong> - ${s.reason} 
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#ffc107; color:#000; margin-left:10px;"
                    onclick="restoreFromSnapshot(${s.timestamp})">Restore</button>
            </li>
        `;
    });
    list.innerHTML = html;
}

function restoreFromSnapshot(ts) {
    if (confirm('ยืนยันย้อนเวลากลับไปจุดนี้? (ข้อมูลปัจจุบันจะหายไป)')) {
        if (CustomerDB.restoreSnapshot(ts)) {
            alert('✅ ย้อนเวลาสำเร็จ (Restored)');
            renderDBTable();
            renderImportHistory(); // if visible
        } else {
            alert('❌ ไม่พบข้อมูล Snapshot นี้');
        }
    }
}

// --- Backup & Restore Glue Code ---
function confirmAndBackup() {
    if (confirm('คุณต้องการ Export ข้อมูลทั้งหมดออกมาเป็นไฟล์ใช่หรือไม่?')) {
        backupData();
    }
}

function backupData() {
    CustomerDB.exportBackup();
}

function restoreData(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (!confirm('คำเตือน: การกู้คืนข้อมูลจะทับข้อมูลปัจจุบันทั้งหมด\nคุณต้องการดำเนินการต่อหรือไม่?')) {
        event.target.value = ''; // Reset
        return;
    }

    CustomerDB.importBackup(file)
        .then(() => {
            alert('✅ กู้คืนข้อมูลสำเร็จ (Restore Complete)');
            if (typeof updateDbViews === 'function') updateDbViews();
            else if (typeof renderDBTable === 'function') renderDBTable(); // Refresh UI
        })
        .catch(err => {
            alert('❌ เกิดข้อผิดพลาด: ' + err.message);
        })
        .finally(() => {
            event.target.value = ''; // Reset
        });
}

function loadBatchToView(batchId) {
    const batches = CustomerDB.getBatches();
    const batch = batches[batchId];

    if (!batch) {
        alert('ไม่พบข้อมูลชุดนี้');
        return;
    }

    // Redirect to Smart Workspace to display
    switchTab('smart');
    const box = document.getElementById('smart-unified-results');

    // CHECK IF WE HAVE RECEIPT METADATA (ranges)
    if (batch.ranges && Array.isArray(batch.ranges)) {
        // Render Receipt Style (Directly use saved metadata which is already optimized)
        const grandTotal = batch.ranges.reduce((acc, r) => acc + (r.total || 0), 0);

        let html = `
            <div class="result-success" style="margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                <strong>📂 ข้อมูล: ${batch.name}</strong>
                <button class="btn btn-neutral" onclick="switchTab('customer')" style="padding:5px 10px; font-size:0.9rem;">⬅ กลับหน้าหลัก (Back to DB)</button>
            </div>
            
            <!-- Receipt View -->
            <div style="background:white; padding:20px; border:1px solid #ddd; box-shadow:0 2px 5px rgba(0,0,0,0.05); font-family:'Courier New', monospace; max-width:800px; margin:0 auto;">
                <h4 style="text-align:center; border-bottom:1px dashed #ccc; padding-bottom:10px; margin-bottom:10px;">ใบสรุปรายการ (Optimized Report)</h4>
                 <div style="margin-bottom:10px; font-size:0.9rem;">
                    <strong>Customer:</strong> ${batch.name}<br>
                    <strong>Type:</strong> ${batch.type}<br>
                    ${batch.requestDate ? `<strong>Request Date:</strong> ${new Date(batch.requestDate).toLocaleDateString('th-TH')}<br>` : ''}
                    <strong>Date:</strong> ${new Date(batch.timestamp).toLocaleString('th-TH')}
                </div>
                 <div style="font-size:0.8rem; color:red; text-align:center; margin-bottom:5px;">
                    *จัดเรียงตามราคาน้อย-มาก (Virtual)
                </div>
                <div class="report-list-container">
                    <div class="report-header-desktop">
                        <div style="width:40%">รายการ (Description)</div>
                        <div style="width:20%; text-align:right">จำนวน (Qty)</div>
                        <div style="width:20%; text-align:right">ราคา/ชิ้น</div>
                        <div style="width:20%; text-align:right">รวม (Total)</div>
                    </div>
        `;



        batch.ranges.forEach((r, idx) => {
            html += `
                <div class="report-card">
                    <div class="report-card-desc">
                        <strong>${idx + 1}. EMS ราคา ${r.price} บาท</strong><br>
                        <span style="color:#0056b3; font-weight:bold;">${r.start === r.end ? r.start : `${r.start} - ${r.end}`}</span><br>
                        <small>น้ำหนัก (Weight): ${r.weight}</small>
                    </div>
                    <div class="report-card-scroll">
                        <div class="stat-item">
                            <span class="stat-label">จำนวน (Qty)</span>
                            <span class="stat-value">${r.count}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">ราคา/ชิ้น</span>
                            <span class="stat-value">@${r.price}</span>
                        </div>
                        <div class="stat-item highlight">
                            <span class="stat-label">รวม (Total)</span>
                            <span class="stat-value">${(r.total || 0).toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                        </div>
                        <!-- Spacer for scroll feel -->
                        <div style="min-width:10px;"></div>
                    </div>
                </div>
            `;
        });

        html += `
                    <div class="report-footer">
                        <div class="footer-label">รวมทั้งสิ้น (Grand Total)</div>
                        <div class="footer-value">${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</div>
                    </div>
                </div>
            </div>
            
            <div style="text-align:center; margin-top:20px;">
                <button class="btn" onclick="window.print()">🖨️ Print / PDF</button>
                <button class="btn btn-neutral" onclick="switchTab('customer')" style="margin-left:10px;">⬅ กลับหน้าหลัก (Back to DB)</button>
            </div>
        `;
        box.innerHTML = html;
        return;
    }

    // FALLBACK
    // Inline Fallback for safety (Legacy View)
    const lookup = CustomerDB.getLookup();
    const items = [];
    Object.keys(lookup).forEach(key => {
        if (lookup[key].batchId === batchId) items.push(key);
    });
    items.sort();

    let fallbackHtml = `
        <div class="result-success">Viewing Legacy Batch: ${batch.name} (${items.length} items)</div>
        <ul>${items.map(x => `<li>${x}</li>`).slice(0, 50).join('')}</ul>
    `;
    box.innerHTML = fallbackHtml;
}

// --- ADMIN TOOLS ---

function adminHandleTrackInput(inputEl) {
    // Basic cleanup but allow spaces and commas for multiple items
    let val = inputEl.value.toUpperCase().replace(/[^A-Z0-9\s,]/g, '');
    
    // Auto Formatter: If user pastes a giant string without commas, we can add them to look nice.
    if (val.length >= 26 && !val.includes(',')) {
        const extracted = TrackingUtils.extractTrackingNumbers(val);
        if (extracted.length > 1) {
            val = extracted.join(', ');
        }
    }
    
    inputEl.value = val;
}

function adminOpenThpTrack() {
    let rawInput = document.getElementById('admin-track-input').value.trim();
    if (!rawInput) {
        alert("กรุณากรอกเลขพัสดุ");
        return;
    }
    
    // Extract all tracking numbers using our robust parser
    const extracted = TrackingUtils.extractTrackingNumbers(rawInput);
    
    if (extracted.length === 0) {
        alert("ไม่พบรูปแบบเลขพัสดุที่ถูกต้อง (13 หลัก)");
        return;
    }
    
    // THP supports submitting multiple numbers joined by comma
    const joinedIds = extracted.join(',');
    const url = `https://track.thailandpost.co.th/dashboard?trackNumber=${joinedIds}`;
    window.open(url, '_blank');
}

/**
 * Applies grayscale and high-contrast thresholding to an image file.
 * Returns a Data URL of the processed image.
 * This removes grey backgrounds and light table lines, helping Tesseract OCR.
 */
function preprocessImageForOCR(file) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext('2d', { willReadFrequently: true });
            
            ctx.drawImage(img, 0, 0);
            
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const data = imageData.data;
            
            // Threshold value (0-255). Lower = more things become white.
            // 150 is usually good for removing light grey shadows but keeping blue/black ink
            const threshold = 150; 
            
            for (let i = 0; i < data.length; i += 4) {
                const r = data[i];
                const g = data[i + 1];
                const b = data[i + 2];
                // Luminance formula
                const v = (0.2126 * r + 0.7152 * g + 0.0722 *b >= threshold) ? 255 : 0;
                
                data[i] = data[i + 1] = data[i + 2] = v; // Set R, G, B
            }
            
            ctx.putImageData(imageData, 0, 0);
            resolve(canvas.toDataURL('image/jpeg', 1.0));
        };
        img.onerror = reject;
        img.src = URL.createObjectURL(file);
    });
}

async function adminHandleImageOcr(files) {
    if (!files || files.length === 0) return;

    const statusEl = document.getElementById('admin-ocr-status');
    const resultEl = document.getElementById('admin-ocr-result');
    resultEl.classList.remove('hidden');
    resultEl.innerHTML = '';
    
    if (statusEl) statusEl.textContent = `กำลังเริ่มประมวลผล...`;

    let combinedText = "";
    let debugProcessedImageUrl = "";

    try {
        for (let i = 0; i < files.length; i++) {
            if (statusEl) statusEl.textContent = `กำลังแยกข้อความจากภาพ ${i + 1}/${files.length}...`;
            const file = files[i];

            // Image Preprocessing: Grayscale + Thresholding
            if (statusEl) statusEl.textContent = `กำลังปรับแต่งรูปภาพ (Preprocessing)...`;
            const processedFileUrl = await preprocessImageForOCR(file);
            debugProcessedImageUrl = processedFileUrl; // Save for debug display

            // Tesseract OCR
            if (statusEl) statusEl.textContent = `กำลังแยกข้อความจากภาพ ${i + 1}/${files.length}...`;
            const useEngOnly = document.getElementById('admin-ocr-eng-only') && document.getElementById('admin-ocr-eng-only').checked;
            const lang = useEngOnly ? 'eng' : 'tha+eng';
            const worker = typeof Tesseract !== 'undefined' ? await Tesseract.createWorker(lang) : null;
            if (worker) {
                // Adjust PSM for tabular data (6 = Assume a single uniform block of text)
                await worker.setParameters({
                    tessedit_pageseg_mode: '6',
                });
                const { data: { text } } = await worker.recognize(processedFileUrl);
                await worker.terminate();
                combinedText += "\n" + text;
            }
        }

        if (statusEl) statusEl.textContent = "กำลังวิเคราะห์ข้อมูล (Analyzing)...";

        const lines = combinedText.split('\n');
        let extractedItems = [];
        
        // 1. Line-by-line parsing to link Tracking and Price on the SAME line
        lines.forEach(line => {
            const trackMatch = line.match(/[A-Z]{2}\s*\d{9}\s*[A-Z]{2}/i);
            const priceMatch = line.match(/(?:\s|^)(\d{1,4}(?:[.,]\d{2})?)(?:\s|$)/); 
            // Thai post receipt usually has price like 42, 32, 52 or 42.00 at the end of the line
            
            if (trackMatch) {
                let track = trackMatch[0].replace(/\s/g, '').toUpperCase();
                let price = 0;
                
                // If there's a number that looks like a price on the same line
                if (priceMatch && priceMatch[1]) {
                     // Filter out numbers that are just sequence numbers like "1", "2" at the start
                     const potentialPrice = parseFloat(priceMatch[1].replace(',', '.'));
                     if (potentialPrice >= 10 && potentialPrice < 5000) { // Reasonable postal bounds
                         price = potentialPrice;
                     }
                }
                extractedItems.push({ track, price, raw: line });
            }
        });

        if (extractedItems.length === 0) {
            statusEl.textContent = "ประมวลผลเสร็จสิ้น: ไม่พบเลขพัสดุในภาพ";
            resultEl.innerHTML = `<em>ไม่พบข้อมูลที่ตรงกับรูปแบบเลขไปรษณีย์ไทย 13 หลัก</em>\n\n[ข้อความดิบจาก OCR]\n${combinedText}\n<hr><img src="${debugProcessedImageUrl}" style="max-width:100%; border:1px solid #ccc; margin-top:10px;">`;
            return;
        }

        // 2. Output Formatting & Missing Check
        let outputHtml = `<strong style="color:var(--primary-color);">📌 ตรวจพบทั้งหมด ${extractedItems.length} รายการ</strong>\n<hr>`;
        let totalPrice = 0;
        
        // Sort items conceptually if possible, but keep original order for display
        let tracksOnly = extractedItems.map(x => x.track);
        
        // Find Missing logic
        let missingReport = "";
        const seqs = extractedItems.map(item => {
            let seqStr = item.track.substring(2, 10);
            return {
                prefix: item.track.substring(0,2),
                suffix: item.track.substring(11,13),
                seq: parseInt(seqStr, 10),
                full: item.track
            };
        }).filter(x => !isNaN(x.seq));

        // Group by prefix+suffix
        const groups = {};
        seqs.forEach(s => {
            let key = s.prefix + s.suffix;
            if(!groups[key]) groups[key] = [];
            groups[key].push(s.seq);
        });

        for (const [key, seqList] of Object.entries(groups)) {
            seqList.sort((a,b) => a - b);
            let missingInGrp = [];
            for (let i = 0; i < seqList.length - 1; i++) {
                let diff = seqList[i+1] - seqList[i];
                if (diff > 1 && diff < 50) { // arbitrary bound so we don't list thousands if it's two different books
                     for (let j = seqList[i] + 1; j < seqList[i+1]; j++) {
                         let missingNum = j.toString().padStart(8, '0');
                         let cd = TrackingUtils.calculateS10CheckDigit(missingNum);
                         if(cd !== null){
                             missingInGrp.push(`${key.substring(0,2)}${missingNum}${cd}${key.substring(2,4)}`);
                         }
                     }
                }
            }
            if (missingInGrp.length > 0) {
                missingReport += `<div style="color:var(--error-color); margin-bottom:10px;">
                    <strong>🚨 แจ้งเตือน: พบช่องโหว่ (เลขที่หายไป) ในช่วงหมวด ${key}:</strong><br>
                    ${missingInGrp.join(', ')}
                </div>`;
            }
        }

        if (missingReport) {
            outputHtml += missingReport + "<hr>";
        }

        // List all with Prices
        outputHtml += `<table style="width:100%; font-size:0.9rem;">
            <tr style="background:#eee;">
                <th style="padding:5px; text-align:left;">ลำดับ</th>
                <th style="padding:5px; text-align:left;">เลขพัสดุ</th>
                <th style="padding:5px; text-align:right;">ราคา (บาท)</th>
            </tr>
        `;
        
        extractedItems.forEach((item, idx) => {
             totalPrice += item.price;
             outputHtml += `
             <tr>
                 <td style="padding:5px; border-bottom:1px solid #eee;">${idx+1}</td>
                 <td style="padding:5px; border-bottom:1px solid #eee;"><strong>${item.track}</strong></td>
                 <td style="padding:5px; border-bottom:1px solid #eee; text-align:right;">${item.price > 0 ? item.price.toFixed(2) : '-'}</td>
             </tr>`;
        });
        
        outputHtml += `
            <tr style="background:#fce8e6; font-weight:bold;">
                <td colspan="2" style="padding:10px; text-align:right;">ยอดรวมทั้งหมด (Total):</td>
                <td style="padding:10px; text-align:right;">${totalPrice.toFixed(2)}</td>
            </tr>
        </table>`;
        
        // DEBUG BLOCK: Output raw text to help us fix the regex
        outputHtml += `
            <div style="margin-top:20px; padding:10px; background:#f8f9fa; border:1px solid #ccc; border-radius:4px; max-height:200px; overflow-y:auto;">
                 <strong style="color:#666; font-size:0.8rem;">[DEBUG] RAW OCR TEXT:</strong>
                 <pre style="font-size:0.75rem; white-space:pre-wrap; margin:0;">${combinedText}</pre>
            </div>
            <div style="margin-top:10px;">
                <strong style="color:#666; font-size:0.8rem;">[DEBUG] ภาพที่ปรับแสงแล้ว:</strong><br>
                <img src="${debugProcessedImageUrl}" style="max-width:100%; border:1px solid #ccc; margin-top:5px;">
            </div>
        `;

        resultEl.innerHTML = outputHtml;
        statusEl.textContent = "ประมวลผลเสร็จสิ้น (Done)";

    } catch (err) {
        console.error(err);
        if (statusEl) statusEl.textContent = "Error: " + err.message;
        resultEl.innerHTML = `<span style="color:red;">Error processing image.</span>`;
    }

    // Reset input
    const uploadInput = document.getElementById('admin-ocr-upload');
    if (uploadInput) uploadInput.value = '';
}

function adminCrossReference() {
    const input = document.getElementById('admin-crossref-input').value;
    const extracted = TrackingUtils.extractTrackingNumbers(input);
    _performCrossRef(extracted);
}

function copyCrossRefAll() {
    const inputArea = document.getElementById('admin-crossref-input');
    if (!inputArea || !inputArea.value.trim()) {
        alert('ไม่มีข้อมูลให้คัดลอกครับ');
        return;
    }
    
    // Extract only valid tracking numbers so no junk text gets copied
    const extracted = TrackingUtils.extractTrackingNumbers(inputArea.value);
    
    if (extracted && extracted.length > 0) {
        const textToCopy = extracted.join('\n');
        navigator.clipboard.writeText(textToCopy).then(() => {
            const btn = document.querySelector('button[onclick="copyCrossRefAll()"]');
            if (btn) {
                const originalText = btn.innerHTML;
                btn.innerHTML = `✅ คัดลอก ${extracted.length} รายการแล้ว`;
                btn.classList.add('btn-success');
                setTimeout(() => {
                    btn.innerHTML = originalText;
                    btn.classList.remove('btn-success');
                }, 1500);
            }
            alert(`คัดลอกเลขพัสดุสำเร็จ ${extracted.length} รายการ`);
        }).catch(err => {
            alert('ไม่สามารถคัดลอกได้: ' + err);
        });
    } else {
        alert('ค้นหาไม่พบรูปแบบเลขพัสดุที่ถูกต้องเพื่อคัดลอกครับ');
    }
}

async function adminCrossRefImage(files) {
    if (!files || files.length === 0) return;
    
    const statusEl = document.getElementById('admin-crossref-status');
    const resultEl = document.getElementById('admin-crossref-result');
    resultEl.classList.add('hidden');
    
    if (statusEl) statusEl.textContent = `กำลังสแกนรูปภาพด้วย OCR...`;

    try {
        const worker = typeof Tesseract !== 'undefined' ? await Tesseract.createWorker('eng') : null; // 'eng' is faster for just numbers
        if (worker) {
            const { data: { text } } = await worker.recognize(files[0]);
            await worker.terminate();
            
            const extracted = TrackingUtils.extractTrackingNumbers(text);
            if(extracted.length > 0) {
                 // Push to textarea so user sees what was found
                 const textArea = document.getElementById('admin-crossref-input');
                 textArea.value = textArea.value + (textArea.value ? '\n' : '') + extracted.join('\n');
                 
                 _performCrossRef(extracted);
            } else {
                 if (statusEl) statusEl.textContent = `ไม่พบเลขพัสดุในรูปภาพ`;
            }
        }
    } catch(err) {
         if (statusEl) statusEl.textContent = `Error OCR: ` + err.message;
    }
    
    // Reset file input
    document.getElementById('admin-crossref-file').value = '';
}

function _performCrossRef(trackingArray) {
    const statusEl = document.getElementById('admin-crossref-status');
    const resultEl = document.getElementById('admin-crossref-result');
    
    if (!trackingArray || trackingArray.length === 0) {
        statusEl.textContent = "กรุณาวางข้อมูล หรือ อัปโหลดรูปภาพที่มีเลขพัสดุก่อน";
        resultEl.classList.add('hidden');
        return;
    }

    statusEl.textContent = `กำลังเทียบข้อมูล ${trackingArray.length} รายการ กับฐานข้อมูลลูกค้า...`;
    
    const lookup = CustomerDB.getLookup();
    const batches = CustomerDB.getBatches();
    
    let html = `
        <table style="width:100%; font-size:0.9rem;">
            <thead>
                <tr style="background:#eee;">
                    <th>ลำดับ</th>
                    <th>เลขพัสดุ</th>
                    <th>สังกัดบริษัทในระบบ (Company / Name)</th>
                </tr>
            </thead>
            <tbody>
    `;

    let foundCount = 0;

    trackingArray.forEach((track, idx) => {
        // ... (existing logic for dbInfoMain, displayList) ...
        let dbInfoMain = lookup[track];
        let isProbableMain = false;

        if (!dbInfoMain && track.length === 13 && typeof CustomerDB.findBookMatch === 'function') {
            const bookMatch = CustomerDB.findBookMatch(track);
            if (bookMatch) {
                dbInfoMain = bookMatch;
                isProbableMain = true;
            }
        }
        
        let displayList = [];
        if (track.length === 13 && typeof TrackingUtils !== 'undefined') {
            displayList = TrackingUtils.generateTrackingRange(track, 2, 1);
        } else {
            displayList = [{ number: track, offset: 0, isCenter: true }];
        }

        // Start Group Container
        html += `<tbody style="border: 1px solid #ddd; border-bottom: 3px solid #ccc; background: white;">`;

        displayList.forEach((item, innerIdx) => {
            let dbInfo = lookup[item.number];
            let companyName = '<span style="color:#ccc; font-style:italic;">ไม่พบข้อมูล (Not Found)</span>';
            let rowStyle = '';
            const isMain = (item.offset === 0);
            
            if (isMain) {
                companyName = '<span style="color:#999; font-style:italic;">ไม่พบข้อมูล (Not Found)</span>';
                if (dbInfoMain) {
                    if (isProbableMain) {
                        companyName = `💡 คาดว่าเป็นของ: <strong style="color:#e65100;">${dbInfoMain.name}</strong> <small style="color:#ef6c00;">(อ้างจากเล่ม)</small>`;
                        rowStyle = 'background-color:#fff3e0; border-left: 5px solid #fb8c00;'; 
                    } else {
                        companyName = `<strong style="color:var(--primary-color); underline">${dbInfoMain.name}</strong>`;
                        if(batches[dbInfoMain.batchId] && batches[dbInfoMain.batchId].requestDate) {
                             companyName += ` <small style="color:#2e7d32; font-weight:bold;"> (ขอเลข: ${new Date(batches[dbInfoMain.batchId].requestDate).toLocaleDateString('th-TH')})</small>`;
                        }
                        rowStyle = 'background-color:#f1f8e9; border-left: 5px solid #43a047;'; 
                        foundCount++;
                    }
                } else {
                    rowStyle = 'background-color: #fffde7; border-left: 5px solid #fbc02d;'; 
                }
            } else {
                rowStyle = 'background-color:#fafafa; color: #888; border-left: 3px solid #eee;'; 
                if (dbInfo) {
                    companyName = `<strong style="color:#666;">${dbInfo.name}</strong>`;
                    if(batches[dbInfo.batchId] && batches[dbInfo.batchId].requestDate) {
                         companyName += ` <small style="color:#999;">(ขอเลข: ${new Date(batches[dbInfo.batchId].requestDate).toLocaleDateString('th-TH')})</small>`;
                    }
                    rowStyle = 'background-color:#f5fcf6; color: #555; border-left: 3px solid #c8e6c9;'; 
                }
            }

            let label = '';
            let trackDisplay = `<span style="font-family:monospace; font-weight:bold; font-size: 1.1em;">${item.number}</span>`;
            let indexCol = '';
            if (!isMain) {
                if (item.offset < 0) label = `(ก่อนหน้า ${Math.abs(item.offset)})`;
                if (item.offset > 0) label = `(ถัดไป ${item.offset})`;
                trackDisplay = `<span style="font-family:monospace; margin-left:15px;">↳ ${item.number} ${label}</span>`;
            } else {
                indexCol = (idx + 1).toString();
            }

            const actionsHtml = `
                <div class="status-actions" style="margin-top:${isMain ? '5px' : '2px'}; margin-bottom: 5px; ${isMain ? '' : 'font-size: 0.8em; opacity: 0.8;'}">
                    <a href="https://track.thailandpost.co.th/?trackNumber=${item.number}&lang=th" target="_blank" class="badge badge-neutral" style="background-color:#e3f2fd; color:#0d47a1; border-color:#90caf9; ${isMain ? '' : 'padding: 2px 4px;'}" title="ติดตามพัสดุ (External Link)">📌 สถานะ</a>
                    <button class="badge badge-neutral" style="border:1px solid #999; cursor:pointer; ${isMain ? '' : 'padding: 2px 4px;'}" onclick="navigator.clipboard.writeText('${item.number}').then(() => alert('คัดลอก ${item.number} แล้ว'))" title="Copy ID">📋 Copy</button>
                </div>
            `;

            html += `
                <tr style="${rowStyle}">
                    <td style="text-align:center; vertical-align: top; padding-top: 15px; border-top: 1px solid #eee;">${indexCol}</td>
                    <td style="vertical-align: top; padding-top: ${isMain ? '15px' : '5px'}; padding-bottom: ${isMain ? '5px' : '5px'}; border-top: 1px solid #eee;">
                        ${trackDisplay}
                        <br>
                        <div style="${isMain ? '' : 'margin-left: 15px;'}">${actionsHtml}</div>
                    </td>
                    <td style="vertical-align: top; padding-top: ${isMain ? '15px' : '5px'}; border-top: 1px solid #eee;">${companyName}</td>
                </tr>
            `;
        });

        // Close Group Container and add spacer row
        html += `
            <tr style="height: 15px; background: transparent;"><td colspan="3" style="border:none;"></td></tr>
        </tbody>`;
    });

    html += `</tbody></table>`;
    
    resultEl.innerHTML = html;
    resultEl.classList.remove('hidden');
    
    statusEl.innerHTML = `ตรวจพบสังกัดตรงกัน <strong style="color:green;">${foundCount}</strong> จากทั้งหมด ${trackingArray.length} รายการ`;
}

// --- Authentication & Isolation System ---

// Initial Run with improved reliability
window.addEventListener('load', function () {
    console.log('Window loaded. Running checkAuth...');
    setTimeout(checkAuth, 100);
});

if (document.readyState === 'complete' || document.readyState === 'interactive') {
    checkAuth();
}

function checkAuth() {
    try {
        console.log('Running checkAuth...');
        const urlParams = new URLSearchParams(window.location.search);
        const isAdmin = urlParams.has('admin');

        const header = document.querySelector('header');
        const tabNav = document.querySelector('.tabs');
        const loginModal = document.getElementById('login-modal');

        if (loginModal) loginModal.style.display = 'none';

        if (isAdmin) {
            // ADMIN MODE
            document.body.classList.add('admin-mode');
            document.body.classList.remove('user-mode');
            
            if (header) header.style.display = 'block';
            if (tabNav) tabNav.style.display = 'flex';

            if (!document.querySelector('.tab-btn.active')) switchTab('smart');

            const snapSec = document.getElementById('snapshot-section');
            if (snapSec) snapSec.classList.remove('hidden');

            const uploadIcon = document.getElementById('upload-icon-display');
            const uploadTitle = document.getElementById('upload-title-display');
            const uploadDesc = document.getElementById('upload-desc-display');
            const uploadInput = document.getElementById('import-upload');

            if (uploadIcon) uploadIcon.innerText = "📂 / 📷";
            if (uploadTitle) uploadTitle.innerText = "แตะเพื่อเลือกไฟล์ Excel หรือ รูปภาพ";
            if (uploadDesc) uploadDesc.innerText = "รองรับ .xlsx, .xls และ รูปภาพ (OCR)";
            if (uploadInput) uploadInput.accept = ".xlsx, .xls, image/*";

        } else {
            // USER MODE
            document.body.classList.add('user-mode');
            document.body.classList.remove('admin-mode');
            
            // Note: CSS will handle hiding .navbar and .tabs via body.user-mode
            if (header) header.style.display = 'none';
            if (tabNav) tabNav.style.display = 'none';

            // Show ONLY Import Tab
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            
            const tabImport = document.getElementById('tab-import');
            if (tabImport) tabImport.classList.add('active');

            // Find and activate the Import tab button (even if hidden, for internal state)
            const btns = document.getElementsByClassName('tab-btn');
            for (let btn of btns) {
                if (btn.getAttribute('onclick')?.includes('import')) {
                    btn.classList.add('active');
                }
            }

            let userHeader = document.getElementById('user-mode-header');
            if (!userHeader) {
                userHeader = document.createElement('div');
                userHeader.id = 'user-mode-header';
                userHeader.innerHTML = `
                    <div style="display:flex; justify-content:center; align-items:center;">
                        <span style="font-size:1rem;">📥 ระบบนำเข้าข้อมูลพัสดุ<br><small style="font-weight:normal; font-size:0.8rem;">(Import Data Entry)</small></span>
                    </div>
                `;
                const main = document.querySelector('main');
                if (main && document.body) document.body.insertBefore(userHeader, main);
            }

            const saveBtn = document.querySelector('button[onclick="saveImportedBatch()"]');
            if (saveBtn) {
                saveBtn.innerHTML = `📤 สรุปและส่งข้อมูล (Submit Report)`;
                saveBtn.classList.remove('btn-primary');
                saveBtn.style.backgroundColor = '#28a745';
                saveBtn.style.color = 'white';
            }

            const rangeGenUI = document.getElementById('range-generator-ui');
            if (rangeGenUI) rangeGenUI.style.display = 'none';
        }
    } catch (e) {
        console.error('Error in checkAuth:', e);
    }
}

// ==========================================
// SECTION: ADMIN TOOLS
// ==========================================

function adminHandleTrackInput(el) {
    el.value = el.value.toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 13);
}

function adminOpenThpTrack() {
    const input = document.getElementById('admin-track-input').value.trim();
    if (!input || input.length !== 13) {
        alert('กรุณากรอกเลขพัสดุให้ครบ 13 หลัก');
        return;
    }
    window.open(`https://track.thailandpost.co.th/?trackNumber=${input}&lang=th`, '_blank');
}

async function adminHandleImageOcr(files) {
    if (!files || files.length === 0) return;
    const statusEl = document.getElementById('admin-ocr-status');
    const resultBox = document.getElementById('admin-ocr-result');
    
    statusEl.innerText = `Preparing OCR for ${files.length} images...`;
    resultBox.classList.add('hidden');
    
    const useEngOnly = document.getElementById('admin-ocr-eng-only')?.checked;
    const lang = useEngOnly ? 'eng' : 'tha+eng';
    
    let combinedText = "";
    try {
        for (let i = 0; i < files.length; i++) {
            statusEl.innerText = `OCR Scanning Image ${i + 1}/${files.length}...`;
            const worker = await Tesseract.createWorker(lang);
            const { data: { text } } = await worker.recognize(files[i]);
            await worker.terminate();
            combinedText += "\n" + text;
        }
        
        statusEl.innerText = "Analyzing extracted data...";
        
        // Extract numbers and prices
        const tableItems = TrackingUtils.extractHandwrittenTable(combinedText);
        
        if(tableItems.length === 0) {
            statusEl.innerText = "ไม่พบข้อมูลที่น่าจะเป็นเลขพัสดุ/ราคาในภาพ";
            return;
        }
        
        // Formulate output string
        let outputStr = `พบข้อมูล ${tableItems.length} รายการ:\n\n`;
        tableItems.forEach((item, idx) => {
             outputStr += `${idx+1}. ${item.number} -> ราคา: ${item.price} บ. (นน. ${item.weight})\n`;
        });
        
        // Check missing if sequence
        const numbersOnly = tableItems.map(t => t.number).sort((a,b) => a.localeCompare(b));
        const parse = (str) => {
            const m = str.match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
            return m ? { full: str, prefix: m[1], body: parseInt(m[2]), check: m[3], suffix: m[4] } : null;
        };
        
        if (numbersOnly.length > 1) {
             let missing = [];
             let prev = parse(numbersOnly[0]);
             for(let i=1; i<numbersOnly.length; i++){
                  const curr = parse(numbersOnly[i]);
                  if(prev && curr && prev.prefix === curr.prefix && prev.suffix === curr.suffix) {
                       const diff = curr.body - prev.body;
                       if (diff > 1) {
                           for(let m = prev.body + 1; m < curr.body; m++) {
                               let bodyStr = m.toString().padStart(8,'0');
                               let cd = TrackingUtils.calculateS10CheckDigit ? TrackingUtils.calculateS10CheckDigit(bodyStr) : 'X';
                               missing.push(`${curr.prefix}${bodyStr}${cd}${curr.suffix}`);
                           }
                       }
                  }
                  prev = curr;
             }
             if (missing.length > 0) {
                 outputStr += `\n⚠️ แจ้งเตือน: พบรายการที่หายไปในลำดับ (Gaps) ${missing.length} รายการ:\n`;
                 missing.forEach(m => outputStr += `- ${m}\n`);
             } else {
                 outputStr += `\n✅ ลำดับเลขต่อเนื่องกันดี ไม่พบรายการตกหล่น`;
             }
        }
        
        resultBox.innerText = outputStr;
        resultBox.classList.remove('hidden');
        statusEl.innerText = "OCR และการวิเคราะห์เสร็จสิ้น!";
        
        document.getElementById('admin-ocr-upload').value = '';
        
    } catch(err) {
        console.error(err);
        statusEl.innerText = "Error: " + err.message;
    }
}

function adminCrossReference() {
    const inputStr = document.getElementById('admin-crossref-input').value.trim();
    if (!inputStr) {
        alert('กรุณาวางเลขพัสดุ');
        return;
    }
    
    const regex = /[A-Z]{2}\d{9}[A-Z]{2}/ig;
    const matches = inputStr.match(regex);
    if (!matches || matches.length === 0) {
         document.getElementById('admin-crossref-status').innerText = 'ไม่พบรูปแบบเลขพัสดุ 13 หลัก';
         return;
    }
    
    renderCrossReference(matches);
}

async function adminCrossRefImage(files) {
    if (!files || files.length === 0) return;
    const statusEl = document.getElementById('admin-crossref-status');
    statusEl.innerText = `Scanning Image for Cross Reference...`;
    
    try {
        const worker = await Tesseract.createWorker('eng');
        const { data: { text } } = await worker.recognize(files[0]);
        await worker.terminate();
        
        const regex = /[A-Z]{2}\d{9}[A-Z]{2}/ig;
        const matches = text.match(regex);
        if(!matches || matches.length === 0) {
            statusEl.innerText = 'ไม่พบเลขพัสดุในรูปนี้';
            return;
        }
        renderCrossReference(matches);
    } catch(e) {
        statusEl.innerText = 'Error: ' + e.message;
    }
    document.getElementById('admin-crossref-file').value = '';
}

function renderCrossReference(trackArray) {
    const uniqueTracks = [...new Set(trackArray.map(t => t.toUpperCase()))];
    const statusEl = document.getElementById('admin-crossref-status');
    const resultBox = document.getElementById('admin-crossref-result');
    
    let html = `
        <table style="width:100%; text-align:left; border-collapse: collapse;">
            <thead>
                <tr style="background:#f1f1f1;">
                    <th style="padding:8px; border-bottom:1px solid #ccc;">#</th>
                    <th style="padding:8px; border-bottom:1px solid #ccc;">เลขพัสดุ</th>
                    <th style="padding:8px; border-bottom:1px solid #ccc;">สังกัด (Customer)</th>
                    <th style="padding:8px; border-bottom:1px solid #ccc;">ประเภท</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    let foundCount = 0;
    
    uniqueTracks.forEach((track, index) => {
        let ownerName = '- ไม่พบในระบบ -';
        let ownerType = '-';
        let rowColor = '';
        
        if (typeof CustomerDB !== 'undefined') {
             const owner = CustomerDB.get(track);
             if (owner) {
                 ownerName = `👤 ${owner.name}`;
                 ownerType = owner.type;
                 rowColor = 'background:#e3f2fd;';
                 foundCount++;
             }
        }
        
        html += `
            <tr style="border-bottom:1px solid #ddd; ${rowColor}">
                <td style="padding:8px;">${index+1}</td>
                <td class="crossref-id" style="padding:8px; font-family:monospace; font-weight:bold;">${TrackingUtils.formatTrackingNumber(track)}</td>
                <td style="padding:8px;">${ownerName}</td>
                <td style="padding:8px;">${ownerType}</td>
            </tr>
        `;
    });
    
    html += `</tbody></table>`;
    
    resultBox.innerHTML = html;
    resultBox.classList.remove('hidden');
    statusEl.innerText = `ตรวจสอบสำเร็จ ${uniqueTracks.length} หมายเลข (พบประวัติ ${foundCount} รายการ)`;
}

function copyCrossRefAll() {
    const ids = Array.from(document.querySelectorAll('.crossref-id')).map(el => el.innerText.replace(/\\s/g, '')).join('\\n');
    if(ids) {
        navigator.clipboard.writeText(ids).then(() => alert('คัดลอกรายการทั้งหมดลง Clipboard แล้ว'));
    } else {
        alert('ไม่มีข้อมูลให้คัดลอก');
    }
}

// ==========================================

// ==========================================
// SECTION: QMS STAGING & IMPORT (Admin)
// ==========================================
let qmsStagingGroups = {}; // Memory to store parsed groups

// =====================================
// DATE HELPER: BE <-> CE conversion
// =====================================

/**
 * Given a BE date string like "19/03/2569 17:03" or "20/03/2569"
 * parse it and set the date+time pickers, updating hidden field and display.
 */
function setExceptionDatePickerFromBE(beStr) {
    if (!beStr) return;
    // Match DD/MM/YYYY HH:MM or DD/MM/YYYY
    const m = beStr.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2}))/);
    if (!m) return;
    const [, dd, mm, yyyy_be, hh, min] = m;
    const yyyy_ce = parseInt(yyyy_be, 10) - 543;
    const dateStr = `${yyyy_ce}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}`;
    const datePicker = document.getElementById('exception-date-picker');
    const timePicker = document.getElementById('exception-time-picker');
    if (datePicker) datePicker.value = dateStr;
    if (timePicker && hh && min) timePicker.value = `${hh.padStart(2,'0')}:${min}`;
    updateBEDisplay();
}

/** Called by oninput on both pickers — updates hidden field and the blue BE label */
function updateBEDisplay() {
    const dateVal = document.getElementById('exception-date-picker')?.value;
    const timeVal = document.getElementById('exception-time-picker')?.value;
    const beDisplay = document.getElementById('exception-be-display');
    const hidden = document.getElementById('exception-datetime');
    if (!dateVal) {
        if (beDisplay) beDisplay.textContent = '';
        if (hidden) hidden.value = '';
        return;
    }
    // Convert CE to BE
    const [yyyy_ce, mm, dd] = dateVal.split('-');
    const yyyy_be = parseInt(yyyy_ce, 10) + 543;
    const beStr = timeVal
        ? `${dd}/${mm}/${yyyy_be} ${timeVal}`
        : `${dd}/${mm}/${yyyy_be}`;
    if (beDisplay) beDisplay.textContent = `📅 ${beStr}`;
    if (hidden) hidden.value = beStr;
}

function toggleQmsStaging() {
    const panel = document.getElementById('qms-staging-panel');
    const chevron = document.getElementById('qms-staging-chevron');
    if (!panel) return;
    const open = panel.style.display !== 'none';
    panel.style.display = open ? 'none' : 'block';
    if (chevron) chevron.textContent = open ? '▼ เปิด' : '▲ ปิด';
}

function processQmsImport() {
    const text = document.getElementById('qms-import-text').value;
    if (!text) return;

    // Extract tracking numbers and optional datetime using index distance to handle multiline grid pastes
    const trackPattern = /[a-zA-Z]{2}\d{9}[a-zA-Z]{2}/g;
    // Strict numeric date regex: matches DD/MM/YYYY or DD-MM-YYYY, with optional space HH:MM
    // NOTE: Do NOT use a pattern that can match partial tracking numbers (e.g. "04TH EQ0897")
    const datetimePattern = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})(?:[^\d]*(\d{1,2}:\d{2}))?/g;
    
    const trackMap = new Map();
    let hasAnyMatches = false;

    const tracks = [];
    let match;
    while ((match = trackPattern.exec(text)) !== null) {
        tracks.push({ track: match[0].toUpperCase(), index: match.index });
    }

    if (tracks.length > 0) {
        hasAnyMatches = true;
        
        const dates = [];
        let dMatch;
        datetimePattern.lastIndex = 0;
        while ((dMatch = datetimePattern.exec(text)) !== null) {
            let dtStr = dMatch[1];
            if (dMatch[2]) dtStr += ' ' + dMatch[2];
            dates.push({ dtStr: dtStr, index: dMatch.index });
        }
        
        if (dates.length === 0) {
            console.warn("No date found in QMS import text.");
        }
        
        tracks.forEach((t, i) => {
            const nextTrackIndex = (i < tracks.length - 1) ? tracks[i+1].index : text.length;
            let validDate = dates.find(d => d.index > t.index && d.index < nextTrackIndex);
            
            // FALLBACK: If index bounding failed because of weird clipboard formatting, grab the first available date in the document.
            if (!validDate && dates.length > 0) {
                validDate = dates[0];
            }
            
            if (!trackMap.has(t.track)) {
                trackMap.set(t.track, validDate ? validDate.dtStr : '');
            } else if (validDate && !trackMap.get(t.track)) {
                trackMap.set(t.track, validDate.dtStr);
            }
        });
    } else {
        alert('ไม่พบเลขพัสดุรูปแบบ 13 หลักในข้อมูลที่วาง');
        return;
    }

    const uniqueTracks = Array.from(trackMap.keys());
    
    qmsStagingGroups = {};
    const existingExceptions = ExceptionManager.getAll();
    const existingNumSet = new Set(
        existingExceptions.flatMap(e => e.entries ? e.entries.map(ex => ex.trackNum) : [e.trackNum])
    );

    uniqueTracks.forEach(track => {
        // Group by 2 letters + space + 4 digits: e.g. "EQ 0898"
        const groupPrefix = track.substring(0, 2) + " " + track.substring(2, 6);
        const dtStr = trackMap.get(track);
        
        if (!qmsStagingGroups[groupPrefix]) {
            qmsStagingGroups[groupPrefix] = {
                id: 'grp_' + track.substring(0, 6),
                prefix: groupPrefix,
                items: [],
                companyHint: 'กำลังค้นหา...',
                companyFound: false,
                extractedDateTime: dtStr
            };
        } else if (!qmsStagingGroups[groupPrefix].extractedDateTime && dtStr) {
            qmsStagingGroups[groupPrefix].extractedDateTime = dtStr;
        }
        
        const isDuplicate = existingNumSet.has(track);
        qmsStagingGroups[groupPrefix].items.push({ track, isDuplicate });
    });

    // Smart Lookup for each group
    Object.values(qmsStagingGroups).forEach(group => {
        const firstTrack = group.items[0].track;
        const prefix = firstTrack.substring(0, 2);
        const suffix = firstTrack.substring(11, 13);
        const bodyInt = parseInt(firstTrack.substring(2, 10), 10);
        
        let foundCompany = null;

        // Helper to check DB
        const checkDb = (bInt) => {
            if (typeof CustomerDB === 'undefined' || typeof TrackingUtils === 'undefined') return null;
            const numStr = bInt.toString().padStart(8, '0');
            const cd = TrackingUtils.calculateS10CheckDigit(numStr);
            const testNum = `${prefix}${numStr}${cd}${suffix}`;
            const info = CustomerDB.get(testNum);
            return info ? info.name : null;
        };

        // 1. Check exact first track
        foundCompany = checkDb(bodyInt);
        
        if (foundCompany) {
            group.companyHint = foundCompany;
            group.companyFound = true;
        } else {
            // 2. Smart Guess: Check -2, -1, +1
            const guessOffsets = [-1, -2, 1];
            for (let offset of guessOffsets) {
                const hint = checkDb(bodyInt + offset);
                if (hint) {
                    foundCompany = hint;
                    break;
                }
            }
            if (foundCompany) {
                group.companyHint = `(คาดเดาจากเลขใกล้เคียง) ${foundCompany}`;
                group.companyFound = true; // treat as found for UI coloring
            } else {
                group.companyHint = 'ไม่พบข้อมูลบริษัทในฐานข้อมูล';
            }
        }
    });

    renderQmsGroups();
}

function renderQmsGroups() {
    const container = document.getElementById('qms-staging-results');
    if (!container) return;

    let html = '';
    const groups = Object.values(qmsStagingGroups);
    
    if (groups.length === 0) {
        container.innerHTML = '<div style="color:#666; font-size:0.85rem;">ไม่พบข้อมูลหลังจัดกลุ่ม</div>';
        return;
    }

    // Sort groups by size descending
    groups.sort((a,b) => b.items.length - a.items.length);

    groups.forEach(g => {
        const dupCount = g.items.filter(i => i.isDuplicate).length;
        const dupWarning = dupCount > 0 ? `<span style="color:#d32f2f; font-size:0.75rem; margin-left:8px; background:#ffebee; padding:2px 6px; border-radius:10px; font-weight:bold;">⚠️ ซ้ำ/ประวัติเดิม ${dupCount} รายการ</span>` : '';
        const bgHint = g.companyFound ? '#e8f5e9' : '#fff';
        const borderHint = g.companyFound ? '#81c784' : '#ccc';
        
        const timeHint = g.extractedDateTime 
            ? `<div style="font-size:0.85rem; color:#0288d1; margin-top:6px;">🕒 เวลาสแกน (QMS): ${g.extractedDateTime}</div>`
            : `<div style="font-size:0.8rem; color:#f57c00; margin-top:6px;">⚠️ ไม่พบข้อมูลเวลาสแกนในข้อมูลที่วาง</div>`;

        html += `
            <div style="border:1px solid ${borderHint}; border-radius:6px; padding:12px; margin-bottom:12px; background:${bgHint}; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:10px; flex-wrap:wrap; gap:8px;">
                    <div>
                        <div style="display:flex; align-items:center;">
                            <strong style="font-size:1.1rem; color:#1565c0;">กลุ่ม ${g.prefix}</strong>
                            <span style="font-size:0.85rem; color:#555; margin-left:8px; padding:2px 6px; background:#e0e0e0; border-radius:10px;">${g.items.length} รายการ</span>
                            ${dupWarning}
                        </div>
                        <div style="font-size:0.9rem; color:#333; margin-top:6px;">🏢 <strong>บริษัท:</strong> ${g.companyHint}</div>
                        ${timeHint}
                    </div>
                    <button class="btn btn-primary" style="font-size:0.85rem; padding:6px 12px; background-color:#0288d1; border-color:#0288d1;" onclick="draftReportFromGroup('${g.prefix}')">
                        ➕ นำกลุ่มนี้ไปสร้างรายงาน
                    </button>
                </div>
                <div style="display:flex; flex-wrap:wrap; gap:6px; font-family:monospace; font-size:0.85rem;">
                    ${g.items.map(i => `<span style="padding:4px 6px; background:${i.isDuplicate ? '#ffebee' : '#faebd7'}; border:1px solid ${i.isDuplicate ? '#ffcdd2' : '#e0e0e0'}; border-radius:4px; color:${i.isDuplicate ? '#c62828' : '#333'}">${i.track}</span>`).join('')}
                </div>
            </div>
        `;
    });

    container.innerHTML = html;
}

function draftReportFromGroup(prefix) {
    const group = qmsStagingGroups[prefix];
    if (!group) return;

    // Open Exception Single Mode if Range is open
    const rangeToggle = document.getElementById('exception-range-toggle');
    if (rangeToggle && rangeToggle.checked) {
        rangeToggle.checked = false;
        if (typeof toggleExceptionRangeMode === 'function') toggleExceptionRangeMode();
    }

    // Clear existing
    const mainInput = document.getElementById('exception-track-input');
    if (mainInput) mainInput.value = '';
    
    document.getElementById('exception-extra-items').innerHTML = '';
    
    // Attempt to reset global counter
    if (typeof window.extraItemCount !== 'undefined') {
        window.extraItemCount = 0; 
    }

    const items = group.items.map(i => i.track);
    
    if (items.length > 0) {
        if (mainInput) mainInput.value = items[0];
        
        for (let i = 1; i < items.length; i++) {
            if (typeof addExceptionExtraItem === 'function') {
                addExceptionExtraItem();
                const extras = document.querySelectorAll('.exception-extra-track');
                if (extras.length > 0) {
                    extras[extras.length - 1].value = items[i];
                }
            }
        }
    }

    // Pre-fill reason if empty
    const reasonInput = document.getElementById('exception-reason-input');
    if (reasonInput && !reasonInput.value) {
        reasonInput.value = "รายละเอียดยังไม่เข้าระบบ/ของยังไม่มาส่ง (จาก QMS)";
    }

    // Pre-fill auto-discovered company name into Subject as a helper if Subject is empty
    let companyNameVal = group.companyFound ? group.companyHint.replace('(คาดเดาจากเลขใกล้เคียง) ', '') : "";
    const subjectInput = document.getElementById('rpt-subject');
    if (subjectInput && !subjectInput.value && companyNameVal) {
        subjectInput.value = "รายงานชิ้นงานตกหล่น: " + companyNameVal;
    }

    // NEW: Auto-fill Date/Time from QMS into the new pickers
    if (group.extractedDateTime) {
        setExceptionDatePickerFromBE(group.extractedDateTime);
    }
    const fsInput = document.getElementById('exception-first-status');
    if (fsInput) fsInput.value = 'ใส่ของลงถุง'; // Enforce default

    // Scroll down to the form smoothly
    if (mainInput) {
        mainInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
        mainInput.style.transition = "background-color 0.5s";
        mainInput.style.backgroundColor = "#fff9c4";
        setTimeout(() => mainInput.style.backgroundColor = "", 1500);
    }
}

/**
 * Pre-fills the exception form from search results.
 */
function stagingQuickReport(tracks, companyName) {
    if (!tracks || tracks.length === 0) return;

    // Reset form first
    const rangeToggle = document.getElementById('exception-range-toggle');
    if (rangeToggle && rangeToggle.checked) {
        rangeToggle.checked = false;
        if (typeof toggleExceptionRangeMode === 'function') toggleExceptionRangeMode();
    }

    const mainInput = document.getElementById('exception-track-input');
    if (mainInput) mainInput.value = tracks[0];
    
    document.getElementById('exception-extra-items').innerHTML = '';
    if (typeof window.extraItemCount !== 'undefined') {
        window.extraItemCount = 0; 
    }

    if (tracks.length > 1) {
        for (let i = 1; i < tracks.length; i++) {
            if (typeof addExceptionExtraItem === 'function') {
                addExceptionExtraItem();
                const extras = document.querySelectorAll('.exception-extra-track');
                if (extras.length > 0) {
                    extras[extras.length - 1].value = tracks[i];
                }
            }
        }
    }

    // Pre-fill metadata
    const subjectInput = document.getElementById('rpt-subject');
    if (subjectInput && !subjectInput.value && companyName) {
        subjectInput.value = "รายงานชิ้นงานตกหล่น: " + companyName;
    }

    const reasonInput = document.getElementById('exception-reason-input');
    if (reasonInput && !reasonInput.value) {
        reasonInput.value = "รายละเอียดยังไม่เข้าระบบ/ของยังไม่มาส่ง";
    }

    // Success feedback and scroll
    if (mainInput) {
        mainInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
        mainInput.style.transition = "background-color 0.8s";
        mainInput.style.backgroundColor = "#e8f5e9";
        setTimeout(() => mainInput.style.backgroundColor = "", 2000);
    }
    
    // Switch to section if needed (logic usually keeps it visible)
}

// SECTION: EXCEPTION META & IMAGE ATTACHMENT
// ==========================================

const EXCEPTION_META_KEY = 'thp_exception_meta_v1';

/** Persist and restore meta fields */
function saveExceptionMeta() {
    const meta = {
        branch:   (document.getElementById('rpt-branch')?.value   || ''),
        date:     (document.getElementById('rpt-date')?.value     || ''),
        reporter: (document.getElementById('rpt-reporter')?.value || ''),
        subject:  (document.getElementById('rpt-subject')?.value  || ''),
        note:     (document.getElementById('rpt-note')?.value     || '')
    };
    localStorage.setItem(EXCEPTION_META_KEY, JSON.stringify(meta));
    
    // Update summary hint in index.html
    const hint = document.getElementById('meta-summary-reporter');
    if (hint) {
        hint.textContent = meta.reporter || '-';
    }
}

/** Get current meta from DOM */
function getExceptionMeta() {
    return {
        branch:   (document.getElementById('rpt-branch')?.value   || ''),
        date:     (document.getElementById('rpt-date')?.value     || ''),
        reporter: (document.getElementById('rpt-reporter')?.value || ''),
        subject:  (document.getElementById('rpt-subject')?.value  || ''),
        note:     (document.getElementById('rpt-note')?.value     || '')
    };
}

function loadExceptionMeta() {
    try {
        const raw = localStorage.getItem(EXCEPTION_META_KEY);
        const meta = raw ? JSON.parse(raw) : {};
        
        const b = document.getElementById('rpt-branch');
        const r = document.getElementById('rpt-reporter');
        const s = document.getElementById('rpt-subject');
        const n = document.getElementById('rpt-note');
        const d = document.getElementById('rpt-date');

        if (b) b.value = meta.branch || '';
        if (r) r.value = meta.reporter || '';
        if (s) s.value = meta.subject || '';
        if (n) n.value = meta.note || '';
        
        if (d) {
            const now = new Date();
            const today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
            if (!d.value || d.value === "") {
                d.value = today;
            }
        }
        
        // Initial hint update
        const hint = document.getElementById('meta-summary-reporter');
        if (hint) {
            hint.textContent = meta.reporter || '-';
        }

        // Final save to persist any defaults (like today's date)
        saveExceptionMeta();
    } catch(e) { 
        console.error("loadExceptionMeta error:", e);
    }
}

/** Image attachment storage (runtime only — not persisted across reloads) */
let exceptionImages = []; // Array of { dataUrl, name }
let currentEditingSessionId = null; // Track if we are editing an existing history entry

/**
 * Compress image before storing to localStorage to avoid QuotaExceededError (5MB limit).
 */
function compressExceptionImage(dataUrl, maxWidth = 1200, quality = 0.6) {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            let width = img.width;
            let height = img.height;

            if (width > maxWidth) {
                height = (maxWidth / width) * height;
                width = maxWidth;
            }

            canvas.width = width;
            canvas.height = height;

            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, width, height);

            // Export as compressed JPEG
            resolve(canvas.toDataURL('image/jpeg', quality));
        };
        img.onerror = () => resolve(dataUrl); // Fallback to original
        img.src = dataUrl;
    });
}

function handleExceptionImageUpload(files) {
    if (!files || files.length === 0) return;
    const preview = document.getElementById('exception-img-preview');
    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            const rawDataUrl = e.target.result;
            
            // Compress BEFORE pushing to global array
            const dataUrl = await compressExceptionImage(rawDataUrl);
            
            const idx = exceptionImages.length;
            exceptionImages.push({ dataUrl, name: file.name });

            // Create preview thumb
            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'position:relative; display:inline-block;';
            wrapper.id = `exc-img-${idx}`;

            const img = document.createElement('img');
            img.src = dataUrl;
            img.style.cssText = 'max-width:120px; max-height:90px; border-radius:4px; border:1px solid #ccc; display:block;';
            img.title = file.name;

            const del = document.createElement('button');
            del.textContent = '✕';
            del.style.cssText = 'position:absolute; top:-5px; right:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#d32f2f; color:white; font-size:0.7rem; cursor:pointer; line-height:1; padding:0;';
            del.onclick = function(event) {
                const currentId = parseInt(event.target.parentElement.id.replace('exc-img-', ''), 10);
                exceptionImages.splice(currentId, 1);
                wrapper.remove();
                const allWrappers = document.getElementById('exception-img-preview').children;
                Array.from(allWrappers).forEach((w, i) => { w.id = `exc-img-${i}`; });
            };

            const editBtn = document.createElement('button');
            editBtn.textContent = '✏️';
            editBtn.style.cssText = 'position:absolute; top:-5px; left:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#0288d1; color:white; font-size:0.6rem; cursor:pointer; line-height:1; padding:0;';
            editBtn.onclick = function(event) {
                const currentId = parseInt(event.target.parentElement.id.replace('exc-img-', ''), 10);
                openImageEditor(currentId);
            };

            wrapper.appendChild(img);
            wrapper.appendChild(del);
            wrapper.appendChild(editBtn);
            preview.appendChild(wrapper);
        };
        reader.readAsDataURL(file);
    });
    // Auto-open panel after first image
    const panel = document.getElementById('exception-img-panel');
    const chevron = document.getElementById('exception-img-chevron');
    if (panel && panel.style.display === 'none') {
        panel.style.display = 'block';
        if (chevron) chevron.textContent = '▲ ซ่อน';
    }
    // Reset input so same file can be re-added
    document.getElementById('exception-img-upload').value = '';
}

function clearExceptionImages() {
    exceptionImages = [];
    const preview = document.getElementById('exception-img-preview');
    if (preview) preview.innerHTML = '';
}

function updateBEDisplay() {
    const dp = document.getElementById('exception-date-picker');
    const tp = document.getElementById('exception-time-picker');
    const hidden = document.getElementById('exception-datetime');
    if (!dp || !tp || !hidden) return;
    hidden.value = `${dp.value}T${tp.value}`;
}

function setExceptionDatePickerFromBE(dateTimeStr) {
    const dp = document.getElementById('exception-date-picker');
    const tp = document.getElementById('exception-time-picker');
    const hidden = document.getElementById('exception-datetime');
    if (!dp || !tp || !hidden) return;
    const parts = dateTimeStr.split('T');
    if (parts.length === 2) {
        dp.value = parts[0];
        tp.value = parts[1];
        hidden.value = dateTimeStr;
    }
}

// ==========================================
// SECTION: EXCEPTION LOG (ตกหล่น)
// ==========================================

/**
 * Toggle between single-entry and range mode for Exception form.
 */
function toggleExceptionRangeMode() {
    const isRange = document.getElementById('exception-range-toggle').checked;
    document.getElementById('exception-single-mode').style.display = isRange ? 'none' : 'flex';
    document.getElementById('exception-range-mode').style.display = isRange ? 'block' : 'none';
    document.getElementById('exception-range-preview').textContent = '';
}

/**
 * Add an extra (non-consecutive) tracking number row below the main form.
 */
let extraItemCount = 0;
function addExceptionExtraItem() {
    const container = document.getElementById('exception-extra-items');
    extraItemCount++;
    const div = document.createElement('div');
    div.id = `extra-item-${extraItemCount}`;
    div.style.cssText = 'display:flex; gap:10px; margin-top:10px; align-items:center;';
    div.innerHTML = `
        <input type="text" class="exception-extra-track" placeholder="เลขที่เพิ่มเติม (เช่น EQ123499999TH)" maxlength="13"
            style="flex:1; text-transform:uppercase; padding:10px; border:1px solid #ccc; border-radius:4px; box-sizing:border-box; font-family:inherit; font-size:inherit;">
        <button type="button" onclick="document.getElementById('extra-item-${extraItemCount}').remove()"
            style="padding:10px 14px; border:1px solid #ffcdd2; border-radius:4px; background:#ffebee; color:#d32f2f; cursor:pointer;">✕</button>
    `;
    container.appendChild(div);
}

/**
 * Parse a tracking number string into its components.
 */
function parseExceptionTrackNum(str) {
    str = str.trim().toUpperCase().replace(/\s+/g, '');
    const m = str.match(/^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/);
    if (!m) return null;
    return { prefix: m[1], body: m[2], cd: m[3], suffix: m[4], bodyInt: parseInt(m[2]), full: str };
}

/**
 * Build a full tracking number with auto check digit.
 */
function buildExceptionTrackNum(prefix, bodyInt, suffix) {
    const bodyStr = bodyInt.toString().padStart(8, '0');
    const cd = TrackingUtils.calculateS10CheckDigit(bodyStr);
    return `${prefix}${bodyStr}${cd}${suffix}`;
}

/**
 * Main save: collects all inputs and saves as one session.
 */
function addExceptionEntry() {
    const isRange = document.getElementById('exception-range-toggle').checked;
    const reason = document.getElementById('exception-reason-input').value.trim();
    const category = document.getElementById('exception-category').value;

    if (!reason) {
        alert('กรุณาระบุเหตุผลที่ตกหล่น');
        document.getElementById('exception-reason-input').focus();
        return;
    }

    const trackNums = [];

    if (isRange) {
        const startRaw = document.getElementById('exception-start-input').value.trim().toUpperCase().replace(/\s+/g, '');
        const endRaw   = document.getElementById('exception-end-input').value.trim().toUpperCase().replace(/\s+/g, '');

        const startParsed = parseExceptionTrackNum(startRaw);
        const endParsed   = parseExceptionTrackNum(endRaw);

        if (!startParsed) { alert('เลขเริ่มต้นไม่ถูกต้อง'); return; }
        if (!endParsed)   { alert('เลขสุดท้ายไม่ถูกต้อง'); return; }
        if (startParsed.prefix !== endParsed.prefix || startParsed.suffix !== endParsed.suffix) {
            alert('เลขเริ่มต้นและสุดท้ายต้องเป็นชุดเดียวกัน (Prefix/Suffix ตรงกัน)');
            return;
        }
        if (startParsed.bodyInt > endParsed.bodyInt) {
            alert('เลขเริ่มต้นต้องน้อยกว่าหรือเท่ากับเลขสุดท้าย');
            return;
        }

        const count = endParsed.bodyInt - startParsed.bodyInt + 1;
        if (count > 500 && !confirm(`คุณกำลังบันทึก ${count} รายการ ยืนยันหรือไม่?`)) return;

        for (let i = startParsed.bodyInt; i <= endParsed.bodyInt; i++) {
            trackNums.push(buildExceptionTrackNum(startParsed.prefix, i, startParsed.suffix));
        }
    } else {
        const single = document.getElementById('exception-track-input').value.trim().toUpperCase().replace(/\s+/g, '');
        if (!single || single.length !== 13) {
            alert('กรุณากรอกเลขพัสดุให้ครบ 13 หลัก');
            document.getElementById('exception-track-input').focus();
            return;
        }
        trackNums.push(single);
    }

    // Extra Items
    document.querySelectorAll('.exception-extra-track').forEach(inp => {
        const v = inp.value.trim().toUpperCase().replace(/\s+/g, '');
        if (v && v.length === 13 && !trackNums.includes(v)) trackNums.push(v);
    });

    if (trackNums.length === 0) { alert('ไม่พบเลขพัสดุที่ถูกต้อง'); return; }

    let companyName = '-';
    if (typeof CustomerDB !== 'undefined') {
        const info = CustomerDB.get(trackNums[0]);
        if (info) companyName = info.name;
    }

    const firstStatus = document.getElementById('exception-first-status').value.trim();
    const dateTime = document.getElementById('exception-datetime').value.trim();
    
    const metadata = {
        category: category,
        branch:   document.getElementById('rpt-branch').value.trim(),
        reporter: document.getElementById('rpt-reporter').value.trim(),
        subject:  document.getElementById('rpt-subject').value.trim(),
        note:     document.getElementById('rpt-note').value.trim()
    };

    // Pass images and current editing ID to save
    const savedId = ExceptionManager.saveSession(trackNums, companyName, reason, firstStatus, dateTime, exceptionImages, currentEditingSessionId, metadata);
    if (!savedId) return;
    
    // Clear state
    currentEditingSessionId = null;
    clearExceptionImages();

    // Reset inputs
    document.getElementById('exception-track-input').value = '';
    document.getElementById('exception-start-input').value = '';
    document.getElementById('exception-end-input').value = '';
    document.getElementById('exception-range-preview').textContent = '';
    document.getElementById('exception-extra-items').innerHTML = '';
    
    // Reset date/time pickers to empty
    const dpicker = document.getElementById('exception-date-picker');
    const tpicker = document.getElementById('exception-time-picker');
    if (dpicker) dpicker.value = '';
    if (tpicker) tpicker.value = '';
    updateBEDisplay();
    extraItemCount = 0;

    renderExceptionTable();

    // Scroll to the new entry
    const container = document.getElementById('exception-table-container');
    if (container) {
        container.scrollIntoView({ behavior: 'smooth', block: 'end' });
    }

    // Visual feedback
    const saveBtn = document.getElementById('exception-save-btn');
    if (saveBtn) {
        const originalText = saveBtn.innerHTML;
        const originalBg = saveBtn.style.background;
        saveBtn.innerHTML = "✅ บันทึกรายการสำเร็จ!";
        saveBtn.style.background = "linear-gradient(135deg,#2e7d32,#388e3c)"; 
        saveBtn.disabled = true;
        setTimeout(() => {
            saveBtn.innerHTML = originalText;
            saveBtn.style.background = originalBg;
            saveBtn.disabled = false;
        }, 1500);
    }
}

/**
 * Group entries by sessionId, show range display, embed meta & images in export target.
 */
/**
 * Render the current draft report in a clean, card-based list.
 */
function renderExceptionTable() {
    const container = document.getElementById('exception-table-container');
    const exportBar = document.getElementById('exception-export-bar');
    if (!container) return;

    const exceptions = (typeof ExceptionManager !== 'undefined') ? ExceptionManager.getAll() : [];
    
    // Toggle Export Bar visibility
    if (exportBar) {
        if (exceptions.length > 0) exportBar.classList.remove('hidden');
        else exportBar.classList.add('hidden');
    }

    if (exceptions.length === 0) {
        container.innerHTML = `
            <div style="text-align:center; padding:40px 20px; color:#999; border:2px dashed #eee; border-radius:12px; background:#fafafa;">
                <p style="margin:0; font-size:1rem;">📭 ยังไม่มีรายการในฉบับร่างนี้</p>
                <p style="margin:5px 0 0 0; font-size:0.85rem;">คุณสามารถระบุเลขด้านบน หรือกดปุ่ม "นำกลุ่มนี้ไปสร้างรายงาน" จากผลการค้นหาได้ครับ</p>
            </div>
        `;
        return;
    }

    // Group by sessionId
    const sessionMap = new Map();
    exceptions.forEach(item => {
        const sid = item.sessionId || item.id;
        if (!sessionMap.has(sid)) {
            sessionMap.set(sid, {
                sessionId: sid,
                companyName: item.companyName,
                reason: item.reason,
                timestamp: item.timestamp,
                entries: []
            });
        }
        sessionMap.get(sid).entries.push(item);
    });

    const sessions = Array.from(sessionMap.values()).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    let html = '<div style="display:flex; flex-direction:column; gap:15px;">';

    sessions.forEach((session, idx) => {
        const compressed = compressEntriesForDisplay(session.entries);
        const dateStr = session.timestamp ? new Date(session.timestamp).toLocaleString('th-TH', { hour: '2-digit', minute: '2-digit' }) : '-';
        const hasImages = session.entries[0] && session.entries[0].images && session.entries[0].images.length > 0;
        const totalCount = session.entries.length;

        html += `
            <div class="report-card" style="background:white; border:1px solid #e0e0e0; border-radius:10px; padding:15px; position:relative; box-shadow:0 2px 5px rgba(0,0,0,0.03); transition:all 0.2s; display:flex; gap:12px; align-items:flex-start;">
                <input type="checkbox" class="sess-select" value="${session.sessionId}" checked style="margin-top:5px; transform:scale(1.2); cursor:pointer;">
                <div style="flex:1;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:10px;">
                    <div>
                        <div style="font-weight:bold; color:#333; font-size:1.05rem; margin-bottom:2px;">
                            🏢 ${session.entries[0].metadata?.companyName || session.companyName || 'รายการชุดที่ ' + (sessions.length - idx)}
                        </div>
                        <div style="font-size:0.85rem; color:#666;">
                            🕒 ${dateStr} • 📦 ${totalCount} ชิ้น • 📂 ${session.entries[0].category || 'เงินสด'}
                        </div>
                    </div>
                    <div style="display:flex; gap:6px;">
                        <button class="btn btn-neutral" style="padding:5px 10px; font-size:0.75rem; border:1px solid #ddd; background:#fff; border-radius:6px;" onclick="editExceptionSession('${session.sessionId}')">✏️ แก้ไข</button>
                        <button class="btn btn-neutral" style="padding:5px 10px; font-size:0.75rem; border:1px solid #ffcdd2; color:#d32f2f; background:#fff; border-radius:6px;" onclick="deleteExceptionSession('${session.sessionId}')">🗑️ ลบ</button>
                    </div>
                </div>

                <div style="background:#f9f9f9; border-radius:8px; padding:12px; margin-bottom:12px; border:1px solid #f0f0f0;">
                    <div style="display:flex; flex-wrap:wrap; gap:6px; font-family:monospace; font-size:1rem;">
                        ${compressed.map(g => `<span style="background:#fff; border:1px solid #eee; padding:4px 10px; border-radius:6px; color:#0277bd; box-shadow:0 1px 2px rgba(0,0,0,0.02);">${g.display}</span>`).join('')}
                    </div>
                </div>

                <div style="display:flex; justify-content:space-between; align-items:flex-end;">
                    <div style="font-size:1rem; color:#d32f2f; font-weight:500;">
                        <span style="color:#666; font-size:0.85rem; font-weight:normal;">สาเหตุ:</span> ${session.reason || '-'}
                    </div>
                    ${hasImages ? `<span style="font-size:0.75rem; color:#2e7d32; background:#e8f5e9; border:1px solid #c8e6c9; padding:3px 10px; border-radius:20px; display:flex; align-items:center; gap:4px;">🖼️ มีรูปประกอบ</span>` : ''}
                </div>
            </div>
        </div>
    `;
    });

    html += '</div>';
    container.innerHTML = html;
}

/**
 * Helper to format tracking numbers consistently for UI.
 */
function formatTrackingForUI(raw) {
    if (!raw) return "";
    raw = raw.replace(/\s+/g, '');
    if (raw.length === 13)
        return `${raw.slice(0,2)} ${raw.slice(2,6)} ${raw.slice(6,10)} ${raw.slice(10,11)} ${raw.slice(11,13)}`;
    return raw;
}

/**
 * Helper to compress continuous tracking ranges for compact display in Draft view.
 */
function compressEntriesForDisplay(entries) {
    const parsed = entries.map(e => {
        const m = (e.trackNum || "").replace(/\s+/g,'').match(/^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/);
        return m ? { full: (e.trackNum||"").replace(/\s+/g,''), prefix: m[1], body: parseInt(m[2]), cd: m[3], suffix: m[4] } : { full: (e.trackNum||""), prefix: null };
    });
    const sortable = parsed.filter(p => p.prefix !== null).sort((a,b) => {
        if (a.prefix !== b.prefix || a.suffix !== b.suffix) return (a.full||"").localeCompare(b.full||"");
        return a.body - b.body;
    });
    const unsortable = parsed.filter(p => p.prefix === null);
    const groups = [];
    let i = 0;
    while (i < sortable.length) {
        let j = i;
        while (j+1 < sortable.length &&
            sortable[j+1].prefix === sortable[i].prefix &&
            sortable[j+1].suffix === sortable[i].suffix &&
            sortable[j+1].body === sortable[j].body + 1) { j++; }
        if (j - i === 0) {
            groups.push({ display: formatTrackingForUI(sortable[i].full), count: 1 });
        } else {
            groups.push({
                display: `${formatTrackingForUI(sortable[i].full)} ~ ${formatTrackingForUI(sortable[j].full)}`,
                count: j - i + 1
            });
        }
        i = j + 1;
    }
    unsortable.forEach(p => groups.push({ display: formatTrackingForUI(p.full), count: 1 }));
    return groups;
}

function deleteExceptionSession(sessionId) {
    if (confirm('ลบรายการในกลุ่มนี้ทั้งหมดใช่หรือไม่?')) {
        ExceptionManager.removeSession(sessionId);
        renderExceptionTable();
    }
}

function deleteException(id) {
    if (confirm('ยืนยันลบรายการนี้?')) {
        ExceptionManager.remove(id);
        renderExceptionTable();
    }
}

function editExceptionSession(sessionId) {
    const exceptions = ExceptionManager.getAll();
    const sessionItems = exceptions.filter(e => e.sessionId === sessionId);
    if (sessionItems.length === 0) return;

    const first = sessionItems[0];
    currentEditingSessionId = sessionId;

    // 1. Populate Range/Single
    if (sessionItems.length > 1) {
        document.getElementById('exception-range-toggle').checked = true;
        document.getElementById('exception-start-input').value = sessionItems[0].trackNum;
        document.getElementById('exception-end-input').value = sessionItems[sessionItems.length-1].trackNum;
    } else {
        document.getElementById('exception-range-toggle').checked = false;
        document.getElementById('exception-track-input').value = first.trackNum;
    }
    toggleExceptionRangeMode();

    // 2. Populate Other Fields
    document.getElementById('exception-reason-input').value = first.reason;
    document.getElementById('exception-first-status').value = first.firstStatus || 'ใส่ของลงถุง';
    
    // Date/Time Pickers
    if (first.dateTime) {
        setExceptionDatePickerFromBE(first.dateTime);
    }

    // 3. Populate Images
    clearExceptionImages();
    if (first.images && Array.isArray(first.images)) {
        // We simulate handleExceptionImageUpload behavior but with dataUrls
        const preview = document.getElementById('exception-img-preview');
        first.images.forEach(imgData => {
            const idx = exceptionImages.length;
            exceptionImages.push({ dataUrl: imgData.dataUrl, originalDataUrl: imgData.dataUrl, name: imgData.name });

            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'position:relative; display:inline-block;';
            wrapper.id = `exc-img-${idx}`;

            wrapper.innerHTML = `
                <img src="${imgData.dataUrl}" style="max-width:120px; max-height:90px; border-radius:4px; border:1px solid #ccc; display:block;">
                <button onclick="const currentId = parseInt(this.parentElement.id.replace('exc-img-', ''), 10); exceptionImages.splice(currentId, 1); this.parentElement.remove(); const allWrappers = document.getElementById('exception-img-preview').children; Array.from(allWrappers).forEach((w, i) => { w.id = 'exc-img-' + i; });" style="position:absolute; top:-5px; right:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#d32f2f; color:white; font-size:0.7rem; cursor:pointer; line-height:1; padding:0;">✕</button>
                <button onclick="const currentId = parseInt(this.parentElement.id.replace('exc-img-', ''), 10); openImageEditor(currentId);" style="position:absolute; top:-5px; left:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#0288d1; color:white; font-size:0.6rem; cursor:pointer; line-height:1; padding:0;">✏️</button>
            `;
            preview.appendChild(wrapper);
        });
        
        // Open panel
        const panel = document.getElementById('exception-img-panel');
        if (panel) {
            panel.style.display = 'block';
            const chevron = document.getElementById('exception-img-chevron');
            if (chevron) chevron.textContent = '▲ ซ่อน';
        }
    }

    // Scroll to top to see the form
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function clearAllExceptions() {
    if (confirm('ยืนยันล้างประวัติรายงานทั้งหมด?')) {
        ExceptionManager.clearAll();
        renderExceptionTable();
    }
}

async function exportExceptionImage() {
    renderExceptionTable();

    const selectedSids = Array.from(document.querySelectorAll('.sess-select:checked')).map(cb => cb.value);
    if (selectedSids.length === 0) {
        alert('กรุณาเลือกอย่างน้อย 1 รายการเพื่อออกรายงานครับ');
        return;
    }

    const exceptions = ExceptionManager.getAll();
    const selectedItems = exceptions.filter(e => selectedSids.includes(e.sessionId));
    const meta = getExceptionMeta();

    // 1. Grouping Logic
    // Structure: Category -> Company -> Sessions
    const groups = {}; 

    selectedItems.forEach(item => {
        const cat = item.category || 'อื่นๆ';
        const com = item.companyName || '-';
        const sid = item.sessionId || item.id;

        if (!groups[cat]) groups[cat] = {};
        if (!groups[cat][com]) groups[cat][com] = {};
        if (!groups[cat][com][sid]) {
            groups[cat][com][sid] = { sessionId: sid, entries: [], first: item, timestamp: item.timestamp };
        }
        groups[cat][com][sid].entries.push(item);
    });

    // 2. Prepare Summary Data
    const summaryRows = [];
    Object.keys(groups).sort().forEach(cat => {
        Object.keys(groups[cat]).sort().forEach(com => {
            let total = 0;
            Object.values(groups[cat][com]).forEach(sess => total += sess.entries.length);
            summaryRows.push({ category: cat, company: com, count: total });
        });
    });

    // 3. Prepare Flat List for Pagination (Group Headers + Session Blocks)
    const exportBlocks = [];
    Object.keys(groups).sort().forEach(cat => {
        exportBlocks.push({ type: 'categoryHeader', title: cat });
        Object.keys(groups[cat]).sort().forEach(com => {
            exportBlocks.push({ type: 'companyHeader', title: com });
            const sids = Object.values(groups[cat][com]).sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
            sids.forEach(sess => {
                exportBlocks.push({ type: 'session', data: sess });
            });
        });
    });

    // Pagination: Approx 12 items/headers per page
    const ITEMS_PER_PAGE = 12;
    const totalPages = Math.ceil(exportBlocks.length / ITEMS_PER_PAGE);

    // Provide visual feedback
    const originalBtnText = event?.target?.innerText || "สร้างรูปภาพแจ้งหัวหน้า";
    if (event?.target) {
        event.target.disabled = true;
        event.target.innerText = "⏳ กำลังเตรียมไฟล์รูปภาพ...";
    }

    const reportDateDisp = meta.date ? new Date(meta.date).toLocaleDateString('th-TH') : new Date().toLocaleDateString('th-TH');
    
    // Create temporary export container
    const exportDiv = document.createElement('div');
    exportDiv.id = 'temp-export-container';
    exportDiv.style.position = 'fixed';
    exportDiv.style.top = '0';
    exportDiv.style.left = '-9999px'; // Hide off-screen
    exportDiv.style.background = 'white';
    document.body.appendChild(exportDiv);

    try {
        for (let p = 0; p < totalPages; p++) {
            const pageBlocks = exportBlocks.slice(p * ITEMS_PER_PAGE, (p + 1) * ITEMS_PER_PAGE);
            const pageNumText = totalPages > 1 ? ` (หน้า ${p + 1}/${totalPages})` : "";
            
            let pageSessions = [];
            pageBlocks.forEach(block => {
                if (block.type === 'session') {
                    pageSessions.push(block.data);
                }
            });

            // Header (Show on every page)
            let headerHtml = `
                <div style="margin-bottom:20px; border-bottom:3px solid #0288d1; padding-bottom:12px; display:flex; justify-content:space-between; align-items:flex-end;">
                    <div>
                        <div style="font-size:1.4rem; font-weight:bold; color:#01579b;">รายงานชิ้นงานที่ไม่มีสถานะรับฝาก</div>
                        <div style="font-size:1rem; color:#0288d1; margin-top:2px;">(Exception & Missing Items Report)</div>
                    </div>
                    <div style="text-align:right;">
                        <div style="font-size:1rem; font-weight:bold; color:#333;">${reportDateDisp}</div>
                        ${totalPages > 1 ? `<div style="font-size:0.9rem; color:#0288d1; font-weight:bold; margin-top:4px;">${pageNumText.trim()}</div>` : ''}
                    </div>
                </div>
                <div style="margin-bottom:15px; font-size:0.9rem; line-height:1.6; color:#444; display:grid; grid-template-columns: 1.5fr 1fr; gap:20px;">
                    <div>
                        ${meta.subject ? `<div><strong>เรื่อง:</strong> ${meta.subject}${pageNumText}</div>` : `<div><strong>เรื่อง:</strong> รายงานชิ้นงานค้าง ${pageNumText}</div>`}
                        ${meta.branch ? `<div><strong>ที่ทำการ:</strong> ${meta.branch}</div>` : ''}
                        ${meta.reporter ? `<div><strong>ผู้รายงาน:</strong> ${meta.reporter}</div>` : ''}
                    </div>
                    <div style="padding:8px; background:#f5faff; border:1px solid #d1e3f8; border-radius:6px;">
                        <strong>หมายเหตุส่วนกลาง:</strong><br>
                        <span style="font-size:0.85rem; color:#666;">${meta.note || '-'}</span>
                    </div>
                </div>`;

            // Summary Table (ONLY Page 1)
            let summaryHtml = '';
            if (p === 0) {
                summaryHtml = `
                    <div style="margin-bottom:20px;">
                        <div style="font-size:1.1rem; font-weight:bold; color:#0d47a1; margin-bottom:8px;">📊 ตารางสรุปภาพรวม (Summary)</div>
                        <table style="width:100%; border-collapse:collapse; background:#fff; border:1px solid #eee; font-size:0.95rem;">
                            <thead>
                                <tr style="background:#e3f2fd; border-bottom:2px solid #0288d1;">
                                    <th style="padding:8px; text-align:left; border:1px solid #eee;">หมวดงาน</th>
                                    <th style="padding:8px; text-align:left; border:1px solid #eee;">ชื่อบริษัท / ลูกค้า</th>
                                    <th style="padding:8px; text-align:center; border:1px solid #eee; width:15%;">รวม (ชิ้น)</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${summaryRows.map(row => `
                                    <tr>
                                        <td style="padding:8px; border:1px solid #eee; font-weight:bold; color:#01579b;">${row.category}</td>
                                        <td style="padding:8px; border:1px solid #eee;">${row.company}</td>
                                        <td style="padding:8px; border:1px solid #eee; text-align:center; font-weight:bold; font-size:1.1rem; color:#0288d1;">${row.count}</td>
                                    </tr>
                                `).join('')}
                                <tr style="background:#f1f8e9; font-weight:bold;">
                                    <td colspan="2" style="padding:10px; text-align:right; border:1px solid #eee;">รวมทั้งสิ้น (Total)</td>
                                    <td style="padding:10px; text-align:center; border:1px solid #eee; font-size:1.2rem; color:#2e7d32;">${summaryRows.reduce((a,b) => a + b.count, 0)}</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>`;
            }

            // Detail Table
            let detailRowsHtml = '';
            let sessionInReportIdx = 0; // Total index in the entire report

            exportBlocks.forEach((block, globalIdx) => {
                const isCurrentPage = (globalIdx >= p * ITEMS_PER_PAGE && globalIdx < (p + 1) * ITEMS_PER_PAGE);
                
                if (block.type === 'session') sessionInReportIdx++;

                if (!isCurrentPage) return;

                if (block.type === 'categoryHeader') {
                    detailRowsHtml += `
                        <tr style="background:#0288d1; color:#fff; font-weight:bold;">
                            <td colspan="3" style="padding:10px 15px; font-size:1.1rem;">📁 หมวด: ${block.title}</td>
                        </tr>`;
                } else if (block.type === 'companyHeader') {
                    detailRowsHtml += `
                        <tr style="background:#e1f5fe; color:#01579b; font-weight:bold; border-bottom:1px solid #b3e5fc;">
                            <td colspan="3" style="padding:8px 15px;">🏢 บริษัท: ${block.title}</td>
                        </tr>`;
                } else if (block.type === 'session') {
                    const sess = block.data;
                    const compressed = compressSessEntries(sess.entries);
                    const totalCount = sess.entries.length;
                    const trackDisplay = compressed.map(g => `<div style="font-family:monospace; font-weight:bold; white-space:nowrap; margin-bottom:2px;">${g.display}</div>`).join('');
                    
                    const firstE = sess.entries[0];
                    let dispDT = new Date(sess.timestamp).toLocaleDateString('th-TH');
                    if (firstE.dateTime) {
                        const m = firstE.dateTime.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2})$/);
                        if(m) { let y=parseInt(m[1]); if(y<2500)y+=543; dispDT=`${m[3]}/${m[2]}/${y} ${m[4]}`; }
                    }

                    const isLast = (globalIdx === exportBlocks.length - 1);
                    const dispIdx = isLast ? '🏁' : sessionInReportIdx;

                    // Per-session images
                    let sessImagesHtml = '';
                    const sessImgs = sess.entries[0].images || [];
                    if (sessImgs.length > 0) {
                        const scaleNode = document.getElementById('exception-img-scale');
                        const scaleVal = scaleNode ? scaleNode.value : "100";
                        const imgSize = (parseInt(scaleVal) / 100 * 180) + "px"; // 180px base for inline
                        
                        sessImagesHtml = `
                            <div style="display:flex; flex-wrap:wrap; gap:8px; margin-top:10px;">
                                ${sessImgs.map(img => `<img src="${img.dataUrl}" style="width:${imgSize}; border:1px solid #ddd; border-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,0.1);">`).join('')}
                            </div>`;
                    }

                    detailRowsHtml += `
                        <tr style="border-bottom: 2px solid #eee; vertical-align:top;">
                            <td style="padding:12px 5px; text-align:center; width:8%; font-weight:bold; color:#666; font-size:1.1rem; background:#f9f9f9; border-right:1px solid #eee;">
                                ${isLast ? `<div style="font-size:0.7rem; color:#d32f2f; margin-bottom:2px;">ท้ายสุด</div>` : ''}
                                ${dispIdx}
                            </td>
                            <td style="padding:12px 15px; width:40%;">${trackDisplay}</td>
                            <td style="padding:12px 15px; width:52%;">
                                <div style="color:#d32f2f; font-weight:bold; margin-bottom:6px; font-size:1.1rem;">${firstE.reason}</div>
                                <div style="font-size:0.85rem; color:#666; margin-bottom:4px;">
                                    ${firstE.firstStatus ? `<span>สถานะ: ${firstE.firstStatus}</span> | ` : ''}
                                    <span>${dispDT}</span>
                                </div>
                                <div style="background:#fff9c4; color:#5d4037; font-size:0.8rem; padding:2px 6px; border-radius:4px; display:inline-block; margin-bottom:5px;">
                                    👥 ${sess.companyName} | 📦 ${totalCount} ชิ้น | 📂 ${block.title || (exportBlocks.find(b => b.type === 'categoryHeader')?.title) || '-'}
                                </div>
                                ${sessImagesHtml}
                            </td>
                        </tr>`;
                }
            });

            let detailTableHtml = `
                <div style="font-size:1.1rem; font-weight:bold; color:#0d47a1; margin-bottom:8px;">🔍 รายละเอียดแยกตามรายการ (Itemized Details)</div>
                <table style="width:100%; border-collapse:collapse; border:1px solid #ccc; box-shadow:0 2px 10px rgba(0,0,0,0.05);">
                    <thead>
                        <tr style="background:#01579b; border-bottom:3px solid #014175;">
                            <th style="padding:12px 5px; text-align:center; border:1px solid #0277bd; color:white; font-size:0.9rem;">ลำดับ</th>
                            <th style="padding:12px 10px; text-align:left; border:1px solid #0277bd; color:white; font-size:0.9rem;">หมายเลขพัสดุ</th>
                            <th style="padding:12px 10px; text-align:left; border:1px solid #0277bd; color:white; font-size:0.9rem;">สาเหตุ / ข้อมูลสแกน / รูปภาพประกอบ</th>
                        </tr>
                    </thead>
                    <tbody>${detailRowsHtml}</tbody>
                </table>`;

            exportDiv.innerHTML = `
                <div style="background:white; padding:40px; width:1000px; font-family:'Sarabun', sans-serif; box-sizing:border-box;">
                    ${headerHtml}
                    ${summaryHtml}
                    ${detailTableHtml}
                    <div style="margin-top:30px; text-align:center; font-size:0.85rem; color:#aaa; border-top:1px solid #f0f0f0; padding-top:10px;">
                        จัดทำโดยระบบ Tracking Analyst Helper | หน้า ${p+1} จากทั้งหมด ${totalPages} หน้า
                    </div>
                </div>`;

            // Wait a moment for images and fonts to render fully
            await new Promise(r => setTimeout(r, 600));
            // Add a tick for layout stability
            await new Promise(r => window.requestAnimationFrame(r));

            const captureNode = exportDiv.querySelector('div');
            // Capture
            const canvas = await html2canvas(captureNode, {
                scale: 2,
                backgroundColor: '#ffffff',
                logging: true,
                useCORS: true,
                allowTaint: true,
                width: captureNode.offsetWidth,
                height: captureNode.offsetHeight,
                onclone: (clonedDoc) => {
                    const clonedNode = clonedDoc.getElementById(exportDiv.id);
                    if (clonedNode) {
                        clonedNode.style.opacity = "1";
                        clonedNode.style.visibility = "visible";
                        clonedNode.style.zIndex = "1";
                        clonedNode.style.left = "0";
                    }
                }
            });

            const imgData = canvas.toDataURL('image/jpeg', 0.85); // Use JPEG for better performance
            const link = document.createElement('a');
            const suffix = totalPages > 1 ? `_Part${p + 1}` : "";
            link.download = `Exception_Report_${new Date().toISOString().slice(0, 10)}${suffix}.jpg`;
            link.href = imgData;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            if (totalPages > 1) await new Promise(r => setTimeout(r, 600));
        }
    } catch (err) {
        console.error("Error generating image:", err);
        alert("ขออภัย! เกิดข้อผิดพลาดในการสร้างรูปภาพ: " + (err.message || err));
    } finally {
        if (exportDiv && exportDiv.parentNode) {
            document.body.removeChild(exportDiv);
        }
        if (event?.target) {
            event.target.disabled = false;
            event.target.innerText = originalBtnText;
        }
    }
}

function toggleHistoryImages(sid) {
    const div = document.getElementById(`sess-imgs-${sid}`);
    if (div) div.style.display = (div.style.display === 'none') ? 'flex' : 'none';
}

function toggleAllSessions(checked) {
    document.querySelectorAll('.sess-select').forEach(cb => cb.checked = checked);
}

function compressSessEntries(entries) {
    function fNum(r) {
        r = r.replace(/\s+/g, '');
        if (r.length === 13) return `${r.slice(0,2)} ${r.slice(2,6)} ${r.slice(6,10)} ${r.slice(10,11)} ${r.slice(11,13)}`;
        return r;
    }
    const parsed = entries.map(e => {
        const m = e.trackNum.replace(/\s+/g,'').match(/^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/);
        return m ? { full: e.trackNum.replace(/\s+/g,''), prefix: m[1], body: parseInt(m[2]), cd: m[3], suffix: m[4] } : { full: e.trackNum, prefix: null };
    });
    const sortable = parsed.filter(p => p.prefix !== null).sort((a,b) => {
        if (a.prefix !== b.prefix || a.suffix !== b.suffix) return a.full.localeCompare(b.full);
        return a.body - b.body;
    });
    const unsortable = parsed.filter(p => p.prefix === null);
    const groups = [];
    let i = 0;
    while (i < sortable.length) {
        let j = i;
        while (j+1 < sortable.length && sortable[j+1].prefix === sortable[i].prefix && sortable[j+1].suffix === sortable[i].suffix && sortable[j+1].body === sortable[j].body + 1) { j++; }
        if (j - i === 0) groups.push({ display: fNum(sortable[i].full), count: 1 });
        else groups.push({ display: `${fNum(sortable[i].full)} <span style="color:#555;">ถึง</span> ${fNum(sortable[j].full)}`, count: j - i + 1 });
        i = j + 1;
    }
    unsortable.forEach(p => groups.push({ display: fNum(p.full), count: 1 }));
    return groups;
}

// Auth listener removed (handled in initial load)

function updateExceptionImageScale() {
    const scaleVal = document.getElementById('exception-img-scale') ? document.getElementById('exception-img-scale').value : "100";
    const imgSize = (parseInt(scaleVal) / 100 * 200) + "px"; // Mapping 100% to 200px base
    
    // Update live preview in upload section
    const previewImgs = document.querySelectorAll('#exception-img-preview img');
    previewImgs.forEach(img => {
        img.style.maxWidth = imgSize;
        img.style.maxHeight = imgSize;
    });
    
    // Update export Target section images (if visible)
    const exportImgs = document.querySelectorAll('#exception-img-export-area img');
    exportImgs.forEach(img => {
        img.style.width = imgSize;
    });
}

function editExceptionCompany(sessionId, trackNum) {
    const newName = prompt('ระบุชื่อบริษัทใหม่ (ระบบจะเรียนรู้และบันทึกลงฐานข้อมูล 240 รายการอัตโนมัติ):');
    if (newName === null || newName.trim() === '') return;
    
    const exceptions = ExceptionManager.getAll();
    let updated = false;
    exceptions.forEach(e => {
        if (e.sessionId === sessionId) {
            e.companyName = newName.trim();
            updated = true;
        }
    });

    if (updated) {
        localStorage.setItem(EXCEPTION_KEY, JSON.stringify(exceptions));
        
        if (typeof CustomerDB !== 'undefined' && typeof TrackingUtils !== 'undefined') {
            const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
            const match = trackNum.match(regex);
            if (match) {
                const prefix = match[1];
                const bodyNum = parseInt(match[2], 10);
                const suffix = match[4];
                
                const bookStart = Math.floor((bodyNum - 1) / 240) * 240 + 1;
                const itemsToAdd = [];
                for (let i = 0; i < 240; i++) {
                    let currentNumStr = (bookStart + i).toString().padStart(8, '0');
                    if (currentNumStr.length > 8) break;
                    let checkDigit = TrackingUtils.calculateS10CheckDigit(currentNumStr);
                    if (checkDigit !== null) {
                        itemsToAdd.push(`${prefix}${currentNumStr}${checkDigit}${suffix}`);
                    }
                }
                
                if (itemsToAdd.length > 0) {
                    CustomerDB.addBatch({
                        name: newName.trim(),
                        type: "General",
                        contract: "",
                        requestDate: "",
                        timestamp: new Date().getTime()
                    }, itemsToAdd);
                    
                    alert(`อัปเดตชื่อบริษัทเป็น "${newName.trim()}" สำเร็จ!\n(สอนบอทจดจำเลขชุดนี้ 240 รายการเรียบร้อย)`);
                }
            }
        }
        renderExceptionTable();
    }
}

function editExceptionReason(sessionId) {
    const newReason = prompt('ระบุสาเหตุ / รายละเอียดการตกหล่นใหม่:');
    if (newReason === null || newReason.trim() === '') return;
    
    const exceptions = ExceptionManager.getAll();
    exceptions.forEach(e => {
        if (e.sessionId === sessionId) {
            e.reason = newReason.trim();
        }
    });
    localStorage.setItem(EXCEPTION_KEY, JSON.stringify(exceptions));
    renderExceptionTable();
}

// ==========================================
// SECTION: INDIVIDUAL IMAGE EDITOR
// ==========================================
let editIdx = -1;
let tImg = new Image();
let tScale = 1;
let baseFitScale = 1;
let tPanX = 0;
let tPanY = 0;
let isDragging = false;
let startX = 0, startY = 0;

function openImageEditor(index) {
    const item = exceptionImages[index];
    if (!item) return;
    editIdx = index;
    
    tImg.onload = function() {
        const cvs = document.getElementById('img-editor-canvas');
        baseFitScale = Math.max(cvs.width / tImg.width, cvs.height / tImg.height);
        tScale = baseFitScale;
        tPanX = (cvs.width - (tImg.width * tScale)) / 2;
        tPanY = (cvs.height - (tImg.height * tScale)) / 2;
        
        const zoomSlider = document.getElementById('img-editor-zoom');
        if(zoomSlider) zoomSlider.value = "100";
        
        drawEditorCanvas();
        document.getElementById('img-editor-modal').style.display = 'flex';
    };
    tImg.src = item.originalDataUrl || item.dataUrl; 
}

function closeImageEditor() {
    document.getElementById('img-editor-modal').style.display = 'none';
}

function drawEditorCanvas() {
    const cvs = document.getElementById('img-editor-canvas');
    if (!cvs) return;
    const ctx = cvs.getContext('2d');
    
    // Smooth rendering
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    
    ctx.clearRect(0, 0, cvs.width, cvs.height);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, cvs.width, cvs.height);
    
    if (tImg.src) {
        ctx.drawImage(tImg, tPanX, tPanY, tImg.width * tScale, tImg.height * tScale);
    }
}

function saveImageEditor() {
    if (editIdx > -1 && exceptionImages[editIdx]) {
        const cvs = document.getElementById('img-editor-canvas');
        const croppedDataUrl = cvs.toDataURL('image/jpeg', 0.92);
        exceptionImages[editIdx].dataUrl = croppedDataUrl;
        
        const thumb = document.querySelector(`#exc-img-${editIdx} img`);
        if (thumb) thumb.src = croppedDataUrl;
    }
    closeImageEditor();
}

document.addEventListener('DOMContentLoaded', () => {
    const cvs = document.getElementById('img-editor-canvas');
    const zoomSlider = document.getElementById('img-editor-zoom');
    
    if (cvs) {
        cvs.addEventListener('mousedown', e => {
            isDragging = true;
            const rect = cvs.getBoundingClientRect();
            const ratioX = cvs.width / rect.width;
            const ratioY = cvs.height / rect.height;
            startX = (e.clientX * ratioX) - tPanX;
            startY = (e.clientY * ratioY) - tPanY;
        });
        window.addEventListener('mouseup', () => { isDragging = false; });
        window.addEventListener('mousemove', e => {
            if (isDragging) {
                const rect = cvs.getBoundingClientRect();
                const ratioX = cvs.width / rect.width;
                const ratioY = cvs.height / rect.height;
                tPanX = (e.clientX * ratioX) - startX;
                tPanY = (e.clientY * ratioY) - startY;
                drawEditorCanvas();
            }
        });
        
        cvs.addEventListener('touchstart', e => {
            if (e.touches.length === 1) {
                isDragging = true;
                const rect = cvs.getBoundingClientRect();
                const ratioX = cvs.width / rect.width;
                const ratioY = cvs.height / rect.height;
                startX = (e.touches[0].clientX * ratioX) - tPanX;
                startY = (e.touches[0].clientY * ratioY) - tPanY;
            }
        }, {passive: false});
        window.addEventListener('touchend', () => { isDragging = false; });
        window.addEventListener('touchmove', e => {
            if (isDragging && e.touches.length === 1) {
                e.preventDefault();
                const rect = cvs.getBoundingClientRect();
                const ratioX = cvs.width / rect.width;
                const ratioY = cvs.height / rect.height;
                tPanX = (e.touches[0].clientX * ratioX) - startX;
                tPanY = (e.touches[0].clientY * ratioY) - startY;
                drawEditorCanvas();
            }
        }, {passive: false});
    }
    
    if (zoomSlider) {
        zoomSlider.addEventListener('input', e => {
            const pct = parseInt(e.target.value, 10) / 100;
            const newScale = baseFitScale * pct;
            const cvsW = cvs.width, cvsH = cvs.height;
            const centerX = cvsW / 2;
            const centerY = cvsH / 2;
            
            tPanX = centerX - (centerX - tPanX) * (newScale / tScale);
            tPanY = centerY - (centerY - tPanY) * (newScale / tScale);
            
            tScale = newScale;
            drawEditorCanvas();
        });
    }
});

