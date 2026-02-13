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
    // Activate button (simple search based on onclick attribute for now)
    const btns = document.getElementsByClassName('tab-btn');
    for (let btn of btns) {
        if (btn.getAttribute('onclick').includes(tabId)) {
            btn.classList.add('active');
        }
    }
}

// Global Event Listeners (Run on load)
document.addEventListener('DOMContentLoaded', () => {
    checkAuth();

    // 1. Check Single Input -> Enter Key
    document.getElementById('input-check-single')?.addEventListener('keypress', function (e) {
        if (e.key === 'Enter') checkSingleNumber();
    });


    // 2. Range Gen Inputs -> Enter Key
    const rangeInputs = ['input-range-center', 'input-range-prev', 'input-range-next'];
    rangeInputs.forEach(id => {
        document.getElementById(id).addEventListener('keypress', function (e) {
            if (e.key === 'Enter') generateRange();
        });
    });

    // 3. Gap Analysis Inputs -> Enter Key
    const gapInputs = ['gap-start', 'gap-end', 'gap-actual-list'];
    gapInputs.forEach(id => {
        document.getElementById(id).addEventListener('keypress', function (e) {
            if (e.key === 'Enter' && !e.shiftKey) { // Shift+Enter for new line in Verify area
                if (id !== 'gap-actual-list') findGaps(); // Don't auto-submit actual list on simple enter as it's multiline
            }
        });
    });

    // Check Admin rights for UI adjustments
    checkAdminUI();
});

function checkAdminUI() {
    const urlParams = new URLSearchParams(window.location.search);
    const isAdmin = urlParams.has('admin');

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

// 1. Single Check Logic
function checkSingleNumber() {
    const inputField = document.getElementById('input-check-single');
    let input = inputField.value.trim();
    const resultBox = document.getElementById('result-check-single');

    if (!input) return;

    const validation = TrackingUtils.validateTrackingNumber(input);
    resultBox.classList.remove('hidden', 'result-success', 'result-error');

    if (validation.isValid) {
        // DB Lookup
        const owner = typeof CustomerDB !== 'undefined' ? CustomerDB.get(input) : null;
        let ownerHtml = '';
        if (owner) {
            ownerHtml = `
                <div style="margin-top:10px; padding:10px; background:#e3f2fd; border-radius:4px; border:1px solid #bbdefb;">
                    <strong>👤 ข้อมูลลูกค้า (Customer Info)</strong><br>
                    Name: ${owner.name}<br>
                    Type: <span class="badge ${owner.type === 'Credit' ? 'badge-primary' : 'badge-neutral'}">${owner.type}</span>
                    ${owner.contract ? ` | Contract: ${owner.contract}` : ''}
                </div>
            `;
        }

        // Check Similar
        const similars = typeof CustomerDB !== 'undefined' ? CustomerDB.findSimilarByBody(input) : [];
        let similarHtml = '';
        if (similars.length > 0) {
            const similarList = similars.map(s => `<li>${s.number} (${s.info.name})</li>`).join('');
            similarHtml = `
                <div class="result-warning" style="margin-top:10px;">
                    <strong>⚠️ แจ้งเตือน: พบรายการคล้ายกัน (Different Prefix)</strong><br>
                    พบเลขที่มีตัวเลขเหมือนกันแต่อักษรหน้าต่างกันในฐานข้อมูล:
                    <ul style="margin:5px 0; padding-left:20px;">${similarList}</ul>
                </div>
            `;
        }

        resultBox.classList.add('result-success');
        resultBox.innerHTML = `
            <strong>✅ ถูกต้อง (Valid)</strong><br>Tracking Number: ${input}
            ${ownerHtml}
            ${similarHtml}
        `;
    } else {
        if (validation.suggestion) {
            // Auto-fix
            const oldInput = input;
            const fixedInput = validation.suggestion;
            inputField.value = fixedInput;

            resultBox.innerHTML = `
                 <div class="result-warning">
                     ⚠️ Check Digit ไม่ถูกต้อง (ใส่มา ${oldInput}) ระบบค้นหาด้วยเลขที่ถูกต้องให้แล้ว
                 </div>
                 <div class="result-box result-success" style="margin-top:0;">
                     <strong>✅ ถูกต้อง (Valid)</strong><br>Tracking Number: ${fixedInput}
                 </div>
             `;
        } else {
            resultBox.classList.add('result-error');
            let html = `<strong>❌ ไม่ถูกต้อง (Invalid)</strong><br>Reason: ${validation.error}`;
            resultBox.innerHTML = html;
        }
    }
}

// 2. Range Generator Logic
function generateRange() {
    const centerInput = document.getElementById('input-range-center');
    let center = centerInput.value.trim();
    const prev = parseInt(document.getElementById('input-range-prev').value) || 0;
    const next = parseInt(document.getElementById('input-range-next').value) || 0;
    const box = document.getElementById('result-range-box');

    if (!center) {
        alert('กรุณาระบุเลขตั้งต้น');
        return;
    }

    const startValidation = TrackingUtils.validateTrackingNumber(center);
    let warningHtml = '';

    if (!startValidation.isValid) {
        if (startValidation.suggestion) {
            // Auto-correct
            warningHtml = `
                <div class="result-warning">
                    ⚠️ เลขตั้งต้นผิด (${center}) ระบบใช้เลขที่ถูกต้อง <strong>${startValidation.suggestion}</strong> แทน
                </div>
             `;
            center = startValidation.suggestion;
            centerInput.value = center;
        } else {
            alert('เลขตั้งต้นไม่ถูกต้อง: ' + startValidation.error);
            return;
        }
    }

    const list = TrackingUtils.generateTrackingRange(center, prev, next);

    // Parse Reference List if available
    const refText = document.getElementById('range-reference-list').value.trim();
    let refSet = new Set();
    const hasReference = refText.length > 0;

    if (hasReference) {
        const regex = /([A-Z]{2}\d{9}[A-Z]{2})/ig;
        const matches = refText.match(regex);
        if (matches) {
            matches.forEach(m => refSet.add(m.toUpperCase()));
        }
    }

    box.classList.remove('hidden');

    let html = warningHtml + `
        <div style="margin-bottom:10px;">
            <strong>รายการทั้งหมด: ${list.length} รายการ</strong>
            <button class="btn" style="padding:4px 8px; font-size:0.8rem; margin-left:10px;" onclick="copyRangeResults()">Copy All</button>
        </div>
        <div class="table-responsive">
        <table>
            <thead>
                <tr>
                    <th>#</th>
                    <th>Tracking Number</th>
                    <th>Status (Simulation)</th>
                </tr>
            </thead>
            <tbody>
    `;

    list.forEach((item, index) => {
        const rowClass = item.isCenter ? 'style="background-color:#fff3cd; font-weight:bold;"' : '';
        let statusHtml = '';
        const owner = typeof CustomerDB !== 'undefined' ? CustomerDB.get(item.number) : null;
        let ownerHtml = '';

        if (owner) {
            ownerHtml = `<br><small style="color:#0056b3;">👤 ${owner.name} (${owner.type})</small>`;
        } else {
            // Check similar if no direct owner
            const similars = typeof CustomerDB !== 'undefined' ? CustomerDB.findSimilarByBody(item.number) : [];
            if (similars.length > 0) {
                const simTitle = similars.map(s => `${s.number} (${s.info.name})`).join(', ');
                ownerHtml = `<br><small style="color:#856404; cursor:help;" title="พบ : ${simTitle}">⚠️ คล้ายกับ ${similars[0].number}...</small>`;
            }
        }

        if (hasReference) {
            if (refSet.has(item.number)) {
                statusHtml = `<span class="badge badge-success">ใส่ของลงถุง (Items Posted)</span>`;
            } else {
                statusHtml = `<span class="badge badge-error">ไม่พบข้อมูล (Not Found)</span>`;
            }
        } else {
            // Default with link
            statusHtml = `
                <div class="status-actions">
                    <a href="https://track.thailandpost.co.th/?trackNumber=${item.number}&lang=th" target="_blank" class="badge badge-neutral" style="background-color:#e3f2fd; color:#0d47a1; border-color:#90caf9;" title="Official Deep Link">🔗 Official</a>
                    <a href="https://www.aftership.com/track/thailand-post/${item.number}?lang=th" target="_blank" class="badge badge-neutral" style="background-color:#fff3e0; color:#e65100; border-color:#ffcc80;" title="Server 2 (AfterShip) - Backup">🚀 Server 2</a>
                    <button class="badge badge-neutral" style="border:1px solid #999; cursor:pointer;" onclick="navigator.clipboard.writeText('${item.number}').then(() => alert('คัดลอก ${item.number} แล้ว'))" title="Copy ID">📋 Copy</button>
                    <a href="https://track.thailandpost.co.th" target="_blank" class="badge badge-neutral" style="border:1px solid #ccc; color:#555;" title="Open Official Site (Manual)">🌐 Manual</a>
                </div>
            `;
        }

        html += `
                <tr ${rowClass}>
                <td>${index + 1}</td>
                <td class="tracking-id">${item.number}${ownerHtml}</td>
                <td>${statusHtml}</td>
            </tr>
                `;
    });

    html += `</tbody></table></div>`;
    box.innerHTML = html;
}

function copyRangeResults() {
    const ids = Array.from(document.querySelectorAll('.tracking-id')).map(el => el.innerText).join('\n');
    navigator.clipboard.writeText(ids).then(() => alert('Copied to clipboard!'));
}

// 3. Gap Analysis Logic
function findGaps() {
    const startObjStr = document.getElementById('gap-start').value.trim();
    const endObjStr = document.getElementById('gap-end').value.trim();
    const actualText = document.getElementById('gap-actual-list').value.trim();
    const box = document.getElementById('result-gap-box');

    if (!startObjStr || !endObjStr) {
        alert('กรุณาระบุเลขเริ่มต้นและสิ้นสุด');
        return;
    }

    // Parse Start/End to get range
    const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
    const startMatch = startObjStr.toUpperCase().match(regex);
    const endMatch = endObjStr.toUpperCase().match(regex);

    if (!startMatch || !endMatch) {
        alert('รูปแบบเลขเริ่มต้นหรือสิ้นสุดไม่ถูกต้อง');
        return;
    }

    const startVal = parseInt(startMatch[2]);
    const endVal = parseInt(endMatch[2]);
    const prefix = startMatch[1];
    const suffix = startMatch[4];

    if (prefix !== endMatch[1] || suffix !== endMatch[4]) {
        alert('Prefix หรือ Suffix ไม่ตรงกันระหว่างเริ่มและจบ');
        return;
    }

    if (endVal < startVal) {
        alert('เลขสิ้นสุดต้องมากกว่าเลขเริ่มต้น');
        return;
    }

    if ((endVal - startVal) > 1000) {
        if (!confirm('ช่วงข้อมูลกว้างกว่า 1,000 รายการ อาจใช้เวลาคำนวณนาน ยืนยันทำต่อ?')) return;
    }

    // Build Expected Set
    const rangeList = TrackingUtils.generateTrackingRange(startObjStr, 0, (endVal - startVal));
    const expectedMap = new Map(); // body -> fullString

    rangeList.forEach(item => {
        const body = item.number.substring(2, 10); // Extract 8 digits
        expectedMap.set(body, item.number);
    });

    // Parse Actual List
    const rawActual = actualText.split(/[\s,]+/);
    const actualSet = new Set();

    rawActual.forEach(raw => {
        const m = raw.toUpperCase().match(regex);
        if (m) {
            actualSet.add(m[2]); // Add body digits
        }
    });

    // diff
    const missing = [];
    expectedMap.forEach((fullStr, body) => {
        if (!actualSet.has(body)) {
            missing.push(fullStr);
        }
    });

    box.classList.remove('hidden');
    if (missing.length === 0) {
        box.innerHTML = `<div class="result-success" style="padding:10px;">✅ ครบถ้วน! ไม่พบรายการตกหล่น</div>`;
    } else {
        let html = `
                <div class="result-error" style="padding:10px; margin-bottom:10px;">
                    <strong>⚠️ พบรายการหายไป ${missing.length} รายการ</strong>
            </div>
                <textarea style="height:150px;">${missing.join('\n')}</textarea>
            `;
        box.innerHTML = html;
    }
}

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
                        const weight = row[4] || '-';
                        newItems.push({
                            number: match[0].toUpperCase(),
                            price: price,
                            weight: weight
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

            const worker = await Tesseract.createWorker('tha+eng');
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
        // Reuse the logic we just built in utils.js
        const extractedSmart = TrackingUtils.extractTrackingNumbers(combinedText);
        extractedSmart.forEach(num => {
            newItems.push({ number: num, price: 0, weight: 'OCR-Smart' });
        });

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
    trackingList.forEach(item => uniqueMap.set(item.number, item));
    const uniqueList = Array.from(uniqueMap.values());

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
        if (p) { p.price = item.price; p.weight = item.weight; }
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
                items: currentList.map(x => x.number)
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
        items: currentList.map(x => x.number)
    });

    // Virtual Optimization
    const optimizedRanges = TrackingUtils.virtualOptimizeRanges(rawRanges);

    currentImportedBatches = optimizedRanges;
    renderImportResult(optimizedRanges, missingItems);
}

function renderImportResult(ranges, missingItems = []) {
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
            const startID = formatID(m.prefix, m.startBody, m.suffix);
            const endID = formatID(m.prefix, m.endBody, m.suffix);

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
            const startID = formatID(m.prefix, m.startBody, m.suffix);
            const endID = formatID(m.prefix, m.endBody, m.suffix);

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

    summary.innerHTML = `
        <strong>📊 สรุปผลการวิเคราะห์ (Virtual Optimization)</strong><br>
        จำนวนทั้งหมด: ${totalItems.toLocaleString()} ชิ้น<br>
        ยอดเงินรวม: <span style="font-size:1.2rem; color:#d63384; font-weight:bold;">${grandTotal.toLocaleString()} บาท</span>
        ${gapHtml}
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
                            ${r.start === r.end ? r.start : `${r.start} - ${r.end}`}
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
        if (typeof renderDBTable === 'function') renderDBTable();
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
            renderDBTable(); // Refresh UI
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

    // Reuse Range Tab to display
    switchTab('range');
    const box = document.getElementById('result-range-box');
    box.classList.remove('hidden');

    // CHECK IF WE HAVE RECEIPT METADATA (ranges)
    if (batch.ranges && Array.isArray(batch.ranges)) {
        // Render Receipt Style (Directly use saved metadata which is already optimized)
        const grandTotal = batch.ranges.reduce((acc, r) => acc + (r.total || 0), 0);

        let html = `
            <div class="result-success" style="margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                <strong>📂 ข้อมูล: ${batch.name}</strong>
                <button class="btn btn-neutral" onclick="switchTab('import')" style="padding:5px 10px; font-size:0.9rem;">⬅ กลับหน้าหลัก (New Import)</button>
            </div>
            
            <!-- Receipt View -->
            <div style="background:white; padding:20px; border:1px solid #ddd; box-shadow:0 2px 5px rgba(0,0,0,0.05); font-family:'Courier New', monospace; max-width:800px; margin:0 auto;">
                <h4 style="text-align:center; border-bottom:1px dashed #ccc; padding-bottom:10px; margin-bottom:10px;">ใบสรุปรายการ (Optimized Report)</h4>
                 <div style="margin-bottom:10px; font-size:0.9rem;">
                    <strong>Customer:</strong> ${batch.name}<br>
                    <strong>Type:</strong> ${batch.type}<br>
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
                <button class="btn btn-neutral" onclick="switchTab('import')" style="margin-left:10px;">⬅ กลับหน้าหลัก (New Import)</button>
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
            if (header) header.style.display = 'block';
            if (tabNav) tabNav.style.display = 'flex';

            if (!document.querySelector('.tab-btn.active')) switchTab('check');

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
            if (header) header.style.display = 'none';
            if (tabNav) tabNav.style.display = 'none';

            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            const tabImport = document.getElementById('tab-import');
            if (tabImport) tabImport.classList.add('active');

            let userHeader = document.getElementById('user-mode-header');
            if (!userHeader) {
                userHeader = document.createElement('div');
                userHeader.id = 'user-mode-header';
                userHeader.style.cssText = `
                    background: var(--primary-color); 
                    color: white; 
                    padding: 15px 20px; 
                    text-align: center; 
                    font-size: 1.2rem; 
                    font-weight: bold;
                    box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
                    margin-bottom: 25px;
                    border-radius: 0 0 16px 16px;
                    position: sticky;
                    top: 0;
                    z-index: 1000;
                `;
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

