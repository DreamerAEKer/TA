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
    // 1. Check Single Input -> Enter Key
    document.getElementById('input-check-single').addEventListener('keypress', function (e) {
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
});

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
                < tr ${rowClass}>
                <td>${index + 1}</td>
                <td class="tracking-id">${item.number}${ownerHtml}</td>
                <td>${statusHtml}</td>
            </tr >
                `;
    });

    html += `</tbody ></table > `;
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
        box.innerHTML = `< div class="result-success" style = "padding:10px;" >✅ ครบถ้วน! ไม่พบรายการตกหล่น</div > `;
    } else {
        let html = `
                < div class="result-error" style = "padding:10px; margin-bottom:10px;" >
                    <strong>⚠️ พบรายการหายไป ${missing.length} รายการ</strong>
            </div >
                <textarea style="height:150px;">${missing.join('\n')}</textarea>
            `;
        box.innerHTML = html;
    }
}
