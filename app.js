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

    // --- Old tools consolidated into Smart Workspace ---
    // Event listeners removed for input-check-single, rangeInputs, gapInputs as UI is merged.
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
    const prev = parseInt(document.getElementById('smart-range-prev').value) || 0;
    const next = parseInt(document.getElementById('smart-range-next').value) || 0;
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

    const list = TrackingUtils.generateTrackingRange(center, prev, next);
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
                    </div>
                </td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    resultArea.innerHTML = html;
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
                
                <!-- Row 3: Online Status -->
                <div style="padding: 15px; border: 1px solid #eee; border-radius: 6px; background:#fafafa; display:flex; flex-direction:column; gap: 10px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                        <span style="font-size:1.05rem;">🌐 สถานะไปรษณีย์ไทย: <span id="online-status-text-${input}" style="color:#f57c00; font-weight:bold;">รอการตรวจสอบ...</span></span>
                        <div style="display:flex; gap:10px;">
                            <button class="btn btn-neutral" style="padding:6px 12px; font-size:0.9rem; background:#f8f9fa; border:1px solid #ddd;" onclick="navigator.clipboard.writeText('${validTarget}').then(()=>alert('คัดลอก ${validTarget} เรียบร้อย!'))">📋 สำเนาเลข</button>
                            <button class="btn btn-success" style="padding:6px 12px; font-size:0.9rem;" onclick="window.open('https://track.thailandpost.co.th/?trackNumber=${validTarget}', '_blank')">🔍 เปิดเว็บ Track&Trace</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;

    // Attempt to fetch online status
    checkOnlineStatusMock(validTarget);
}

async function checkOnlineStatusMock(trackNumber) {
    const statusTextEl = document.getElementById(`online-status-text-${trackNumber}`);
    if (!statusTextEl) return;
    
    statusTextEl.innerText = "กำลังดึงข้อมูล...";
    statusTextEl.style.color = "#0056b3";

    try {
        // Simulate network delay to try to fetch
        await new Promise(r => setTimeout(r, 800));
        
        // This fetch is guaranteed to fail due to CORS in standard browsers without a proxy,
        // but we write the logic to show the user what would happen.
        // If they ever run it via extension/proxy, it might work!
        const res = await fetch(`https://trackapi.thailandpost.co.th/post/api/v1/track?trackNumber=${trackNumber}`, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });

        if (res.ok) {
            const data = await res.json();
             // Assuming typical status response
            statusTextEl.innerText = "ดึงข้อมูลสำเร็จ (อัปเดตล่าสุด)";
            statusTextEl.style.color = "#28a745";
        } else {
             throw new Error("HTTP " + res.status);
        }
    } catch (e) {
        // Fallback due to CORS or API Block
        statusTextEl.innerHTML = `ไม่สามารถดึงข้อมูลอัตโนมัติได้ <span style="font-size:0.8rem; color:#666; font-weight:normal; display:inline-block; margin-top:4px;">(ติดระบบป้องกัน CORS ของไปรษณีย์ กรุณากดปุ่มเปิดเว็บเบราว์เซอร์ด้านขวา) ⚠️</span>`;
        statusTextEl.style.color = "#d32f2f";
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
// SECTION: EXCEPTION LOG (ตกหล่น)
// ==========================================

function renderExceptionTable() {
    const container = document.getElementById('exception-table-container');
    if (!container) return;
    
    // Check if ExceptionManager is ready
    if(typeof ExceptionManager === 'undefined') {
        container.innerHTML = '<span style="color:red;">ExceptionManager not loaded. Check db_manager.js</span>';
        return;
    }

    const exceptions = ExceptionManager.getAll();
    
    if (exceptions.length === 0) {
        container.innerHTML = '<p style="text-align:center; color:#999; padding:20px;">ยังไม่มีข้อมูลประวัติการตกหล่น</p>';
        return;
    }

    // Sort by newest first
    exceptions.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));

    let html = `
        <div id="exception-export-target" style="background:white; padding:15px; border-radius:8px;">
            <div style="margin-bottom:10px; border-bottom:2px solid #333; padding-bottom:10px;">
                <strong style="font-size:1.1rem;">รายงานขึ้นงานที่ไม่มีสถานะรับฝาก</strong>
            </div>
            <table style="width:100%; font-size:0.9rem; border-collapse: collapse;">
                <thead>
                    <tr style="background:#f1f1f1;">
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:center;">ลำดับ</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:left;">เลขพัสดุ</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:left;">ชื่อบริษัท/สังกัด (ถ้ามี)</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:left;">เหตุผล / สถานะ</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:center;" data-html2canvas-ignore>จัดการ</th>
                    </tr>
                </thead>
                <tbody>
    `;

    exceptions.forEach((item, idx) => {
        const dateStr = new Date(item.timestamp).toLocaleDateString('th-TH');
        html += `
            <tr style="border-bottom: 1px solid #eee;">
                <td style="padding:10px 8px; text-align:center;">${idx + 1}</td>
                <td style="padding:10px 8px; font-family:monospace; font-weight:bold;">${item.trackNum}</td>
                <td style="padding:10px 8px;">${item.companyName}</td>
                <td style="padding:10px 8px; color:#d32f2f;">${item.reason}</td>
                <td style="padding:10px 8px; text-align:center;" data-html2canvas-ignore>
                    <button class="btn btn-danger" style="padding:4px 8px; font-size:0.8rem;" onclick="deleteException('${item.id}')">ลบ</button>
                </td>
            </tr>
        `;
    });

    html += `
                </tbody>
            </table>
        </div>
        <div style="margin-top:10px; text-align:right;">
             <button class="btn btn-neutral" onclick="clearAllExceptions()">🗑️ ล้างประวัติทั้งหมด</button>
        </div>
    `;

    container.innerHTML = html;
}

function addExceptionEntry() {
    const trackInput = document.getElementById('exception-track-input');
    const reasonInput = document.getElementById('exception-reason-input');
    
    const trackNum = trackInput.value.trim().toUpperCase();
    const reason = reasonInput.value.trim();
    
    if (!trackNum || trackNum.length !== 13) {
        alert('กรุณากรอกเลขพัสดุให้ครบ 13 หลัก');
        trackInput.focus();
        return;
    }
    
    if (!reason) {
        alert('กรุณาระบุเหตุผลที่ตกหล่น');
        reasonInput.focus();
        return;
    }
    
    // Attempt lookup to find company name
    let companyName = '-';
    if(typeof CustomerDB !== 'undefined') {
        const lookupInfo = CustomerDB.get(trackNum);
        if(lookupInfo) {
             companyName = lookupInfo.name;
        }
    }
    
    ExceptionManager.save(trackNum, companyName, reason);
    
    // Clear inputs
    trackInput.value = '';
    // reasonInput.value = ''; // keep reason in case of multiple similar entries
    trackInput.focus();
    
    renderExceptionTable();
}

function deleteException(id) {
    if(confirm('ยอดลบรายการนี้ใช่หรือไม่?')) {
        ExceptionManager.remove(id);
        renderExceptionTable();
    }
}

function clearAllExceptions() {
    if(ExceptionManager.clearAll()) {
         renderExceptionTable();
    }
}

function exportExceptionImage() {
    const targetNode = document.getElementById('exception-export-target');
    if (!targetNode) {
        alert('ไม่พบข้อมูลที่จะสร้างรูปภาพ กรุณาเพิ่มประวัติก่อนครับ');
        return;
    }
    
    if(typeof html2canvas === 'undefined') {
        alert('ระบบกำลังโหลดเครื่องมือสร้างภาพ หรือโหลดไม่สำเร็จ กรุณาลองใหม่ (ต้องต่อเน็ต)');
        return;
    }
    
    const originalBackground = targetNode.style.background;
    targetNode.style.background = '#ffffff'; // Ensure white background for image
    
    // Temporarily apply padding for better image borders
    const originalPadding = targetNode.style.padding;
    targetNode.style.padding = "30px";
    
    html2canvas(targetNode, {
        scale: 1.5, // balanced resolution for LINE sharing
        backgroundColor: '#ffffff',
        logging: false
    }).then(canvas => {
        // Restore styles
        targetNode.style.background = originalBackground;
        targetNode.style.padding = originalPadding;
        
        // Trigger Download
        const imgData = canvas.toDataURL('image/jpeg', 0.85);
        const link = document.createElement('a');
        link.download = `Exception_Report_${new Date().toISOString().slice(0,10)}.jpg`;
        link.href = imgData;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }).catch(err => {
        targetNode.style.background = originalBackground;
        targetNode.style.padding = originalPadding;
        console.error('Error generating image:', err);
        alert('เกิดข้อผิดพลาดในการสร้างภาพ: ' + err.message);
    });
}

// Update checkAuth hook internally to also init exception table if Admin
const originalCheckAuth = checkAuth;
checkAuth = function() {
    originalCheckAuth();
    if(document.getElementById('exception-table-container')) {
        renderExceptionTable();
    }
};

