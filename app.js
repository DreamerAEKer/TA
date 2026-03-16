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
            uploadIcon.textContent = "้ฆๆจ / ้ฆๆ‘ณ";
            uploadTitle.textContent = "ๅ–ไฝฎ็ฌ—ๅ–”็ญ็ฎ‘ๅ–”็ง้ๅ–ๅ —่…‘ๅ–โฌๅ–”ใ ้ๅ–”๏ฟฝ็ซตๅ–ๅฆ็ฌฉๅ–”ใ ็ฎค Excel ๅ–”๏ฟฝ็ฆๅ–”็ฒช่…‘ ๅ–”๏ฝ€่…นๅ–”ๆถ็ฌญๅ–”ไพง็ฌง";
            uploadDesc.textContent = "ๅ–”๏ฝ€่…‘ๅ–”ๅ็ฆๅ–”็ผ–็ฌ .xlsx, .xls ๅ–ไฝฎๅผ—ๅ–”๏ฟฝ ๅ–”๏ฝ€่…นๅ–”ๆถ็ฌญๅ–”ไพง็ฌง (ๅ–โฌๅ–”ใ ้ๅ–”๏ฟฝ็ซตๅ–ๅฆ็ฌ–ๅ–ๅคๆ–งๅ–”ใ ่ฆๅ–”โ‘ง็ฆๅ–”็็ฌก)";
            importInput.setAttribute('accept', '.xlsx, .xls, .jpg, .jpeg, .png, .heic');
        } else {
            // User: Excel Only
            uploadIcon.textContent = "้ฆๆจ";
            uploadTitle.textContent = "ๅ–ไฝฎ็ฌ—ๅ–”็ญ็ฎ‘ๅ–”็ง้ๅ–ๅ —่…‘ๅ–โฌๅ–”ใ ้ๅ–”๏ฟฝ็ซตๅ–ๅฆ็ฌฉๅ–”ใ ็ฎค Excel";
            uploadDesc.textContent = "ๅ–”๏ฝ€่…‘ๅ–”ๅ็ฆๅ–”็ผ–็ฌ .xlsx, .xls (ๅ–”๏ฟฝ่ตๅ–”๏ฟฝ็ฆๅ–”็ผ–็ฌๅ–”็ง็ฌๅ–”็ผ–็ซตๅ–”ๅ่ฆๅ–”๏ฟฝ)";
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
                    <strong>้ฆๆ ๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”ใ ่…นๅ–”ไฝฎ็ซธๅ–ๅค่ฆ (Customer Info)</strong><br>
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
                    <strong>้ฟ็…็ฌ ๅ–ไฝฎ็ฌ€ๅ–ๅค็ซพๅ–โฌๅ–”ๆ•้ๅ–”๏ฟฝ็ฌ: ๅ–”็ง็ฌๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆๅ–”ๅฆๅผ—ๅ–ๅค่ฆๅ–”โ‘ง็ซตๅ–”็ผ–็ฌ (Different Prefix)</strong><br>
                    ๅ–”็ง็ฌๅ–โฌๅ–”ใ ็ซถๅ–”ๆคธๅ…ๅ–ๅ —ๆตฎๅ–”ๆ็ฌ—ๅ–”็ผ–ๆๅ–โฌๅ–”ใ ็ซถๅ–โฌๅ–”๏ฟฝๆตฎๅ–”็ฒช่…‘ๅ–”ๆฌ็ซตๅ–”็ผ–็ฌๅ–ไฝฎ็ฌ—ๅ–ๅ —่…‘ๅ–”็ผ–็ซตๅ–”โ”ผ็ฆๅ–”๏ฟฝ็ฌๅ–ๅค่ฆๅ–”ๆ•็ฎๅ–”ไพง็ซพๅ–”ไฝฎๅฏๅ–”ๆฌ็ฎ–ๅ–”ๆฌ็ฌๅ–”ไพง็ฌๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—:
                    <ul style="margin:5px 0; padding-left:20px;">${similarList}</ul>
                </div>
            `;
        }

        resultBox.classList.add('result-success');
        resultBox.innerHTML = `
            <strong>้๏ฟฝ ๅ–”ๆ ข่…นๅ–”ไฝฎ็ฌ—ๅ–ๅค่…‘ๅ–”๏ฟฝ (Valid)</strong><br>Tracking Number: ${TrackingUtils.formatTrackingNumber(input)}
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
                     ้ฟ็…็ฌ Check Digit ๅ–ๅฆๆตฎๅ–ๅ —็ฌๅ–”็็ซตๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพ (ๅ–ๅ้ๅ–ๅ —ๆตฎๅ–”๏ฟฝ ${oldInput}) ๅ–”๏ฝ€่ตดๅ–”ๆฐ็ฌๅ–”ๅฆ็ฎๅ–”ๆฌๆ–งๅ–”ไพง็ฌ–ๅ–ๅคๆๅ–”โ‘ง็ฎ‘ๅ–”ใ ็ซถๅ–”ๆคธๅ…ๅ–ๅ —็ฌๅ–”็็ซตๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพๅ–ๅๆ–งๅ–ๅค็ฎ’ๅ–”ใ ็ฎๅ–”๏ฟฝ
                 </div>
                 <div class="result-box result-success" style="margin-top:0;">
                     <strong>้๏ฟฝ ๅ–”ๆ ข่…นๅ–”ไฝฎ็ฌ—ๅ–ๅค่…‘ๅ–”๏ฟฝ (Valid)</strong><br>Tracking Number: ${TrackingUtils.formatTrackingNumber(fixedInput)}
                 </div>
             `;
        } else {
            resultBox.classList.add('result-error');
            let html = `<strong>้๏ฟฝ ๅ–ๅฆๆตฎๅ–ๅ —็ฌๅ–”็็ซตๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพ (Invalid)</strong><br>Reason: ${validation.error}`;
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
        alert('ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพง็ฆๅ–”็ญ็ฌๅ–”่็ฎ‘ๅ–”ใ ็ซถๅ–”ๆ•ๅฏๅ–ๅค็ซพๅ–”ๆ•็ฎๅ–”๏ฟฝ');
        return;
    }

    const startValidation = TrackingUtils.validateTrackingNumber(center);
    let warningHtml = '';

    if (!startValidation.isValid) {
        if (startValidation.suggestion) {
            // Auto-correct
            warningHtml = `
                <div class="result-warning">
                    ้ฟ็…็ฌ ๅ–โฌๅ–”ใ ็ซถๅ–”ๆ•ๅฏๅ–ๅค็ซพๅ–”ๆ•็ฎๅ–”ๆฌ็ฌขๅ–”่็ฌ– (${center}) ๅ–”๏ฝ€่ตดๅ–”ๆฐ็ฌๅ–ๅ็ฌๅ–ๅค็ฎ‘ๅ–”ใ ็ซถๅ–”ๆคธๅ…ๅ–ๅ —็ฌๅ–”็็ซตๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพ <strong>${startValidation.suggestion}</strong> ๅ–ไฝฎ็ฌๅ–”๏ฟฝ
                </div>
             `;
            center = startValidation.suggestion;
            centerInput.value = center;
        } else {
            alert('ๅ–โฌๅ–”ใ ็ซถๅ–”ๆ•ๅฏๅ–ๅค็ซพๅ–”ๆ•็ฎๅ–”ๆฌ็ฎๅ–”โด็ฎๅ–”ๆ ข่…นๅ–”ไฝฎ็ฌ—ๅ–ๅค่…‘ๅ–”๏ฟฝ: ' + startValidation.error);
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
            <strong>ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆๅ–”ๆคธๅฏๅ–ๅค็ซพๅ–”๏ฟฝๆตฎๅ–”๏ฟฝ: ${list.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ</strong>
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
            ownerHtml = `<br><small style="color:#0056b3;">้ฆๆ ${owner.name} (${owner.type})</small>`;
        } else {
            // Check similar if no direct owner
            const similars = typeof CustomerDB !== 'undefined' ? CustomerDB.findSimilarByBody(item.number) : [];
            if (similars.length > 0) {
                const simTitle = similars.map(s => `${s.number} (${s.info.name})`).join(', ');
                ownerHtml = `<br><small style="color:#856404; cursor:help;" title="ๅ–”็ง็ฌ : ${simTitle}">้ฟ็…็ฌ ๅ–”ๅฆๅผ—ๅ–ๅค่ฆๅ–”โ‘ง็ซตๅ–”็ผ–็ฌ ${similars[0].number}...</small>`;
            }
        }

        if (hasReference) {
            if (refSet.has(item.number)) {
                statusHtml = `<span class="badge badge-success">ๅ–ๅ้ๅ–ๅ —็ซถๅ–”๏ฟฝ็ซพๅ–”ใ ็ซพๅ–”ๆ ข็ถๅ–”๏ฟฝ (Items Posted)</span>`;
            } else {
                statusHtml = `<span class="badge badge-error">ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”๏ฟฝ (Not Found)</span>`;
            }
        } else {
            // Default with link
            statusHtml = `
                <div class="status-actions">
                    <a href="https://track.thailandpost.co.th/?trackNumber=${item.number}&lang=th" target="_blank" class="badge badge-neutral" style="background-color:#e3f2fd; color:#0d47a1; border-color:#90caf9;" title="Official Deep Link">้ฆๆ• Official</a>
                    <a href="https://www.aftership.com/track/thailand-post/${item.number}?lang=th" target="_blank" class="badge badge-neutral" style="background-color:#fff3e0; color:#e65100; border-color:#ffcc80;" title="Server 2 (AfterShip) - Backup">้ฆๆฎ Server 2</a>
                    <button class="badge badge-neutral" style="border:1px solid #999; cursor:pointer;" onclick="navigator.clipboard.writeText('${item.number}').then(() => alert('ๅ–”ๅฆๅฏๅ–”ๆ–ทๅผ—ๅ–”๏ฟฝ็ซต ${item.number} ๅ–ไฝฎๅผ—ๅ–ๅคๆ'))" title="Copy ID">้ฆๆต Copy</button>
                    <a href="https://track.thailandpost.co.th" target="_blank" class="badge badge-neutral" style="border:1px solid #ccc; color:#555;" title="Open Official Site (Manual)">้ฆๅฏช Manual</a>
                </div>
            `;
        }

        html += `
                <tr ${rowClass}>
                <td>${index + 1}</td>
                <td class="tracking-id">${TrackingUtils.formatTrackingNumber(item.number)}${ownerHtml}</td>
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
        alert('ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพง็ฆๅ–”็ญ็ฌๅ–”่็ฎ‘ๅ–”ใ ็ซถๅ–โฌๅ–”๏ฝ€ๅคๅ–ๅ —ๆตฎๅ–”ๆ•็ฎๅ–”ๆฌ็ฎ’ๅ–”ใ ่ตดๅ–”๏ฟฝๅคๅ–ๅค็ฌๅ–”๏ฟฝ็ถๅ–”๏ฟฝ');
        return;
    }

    // Parse Start/End to get range
    const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
    const startMatch = startObjStr.toUpperCase().match(regex);
    const endMatch = endObjStr.toUpperCase().match(regex);

    if (!startMatch || !endMatch) {
        alert('ๅ–”๏ฝ€่…นๅ–”ๆถ็ฎ’ๅ–”ๆฐ็ฌๅ–โฌๅ–”ใ ็ซถๅ–โฌๅ–”๏ฝ€ๅคๅ–ๅ —ๆตฎๅ–”ๆ•็ฎๅ–”ๆฌๆ–งๅ–”๏ฝ€้ๅ–”๏ฟฝ้ๅ–”่็ฎๅ–”ๆฌ้ๅ–”่็ฌ–ๅ–ๅฆๆตฎๅ–ๅ —็ฌๅ–”็็ซตๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพ');
        return;
    }

    const startVal = parseInt(startMatch[2]);
    const endVal = parseInt(endMatch[2]);
    const prefix = startMatch[1];
    const suffix = startMatch[4];

    if (prefix !== endMatch[1] || suffix !== endMatch[4]) {
        alert('Prefix ๅ–”๏ฟฝ็ฆๅ–”็ฒช่…‘ Suffix ๅ–ๅฆๆตฎๅ–ๅ —็ฌ—ๅ–”๏ฝ€็ซพๅ–”ไฝฎๅฏๅ–”ๆฌ็ฆๅ–”็ญๆ–งๅ–”ะพ็ฎๅ–”ไพง็ซพๅ–โฌๅ–”๏ฝ€ๅคๅ–ๅ —ๆตฎๅ–ไฝฎๅผ—ๅ–”็ญ็ฌ€ๅ–”๏ฟฝ');
        return;
    }

    if (endVal < startVal) {
        alert('ๅ–โฌๅ–”ใ ็ซถๅ–”๏ฟฝๅคๅ–ๅค็ฌๅ–”๏ฟฝ็ถๅ–”ๆ–ท็ฌ—ๅ–ๅค่…‘ๅ–”ๅๆตฎๅ–”ไพง็ซตๅ–”ไฝฎๆๅ–ๅ —่ฆๅ–โฌๅ–”ใ ็ซถๅ–โฌๅ–”๏ฝ€ๅคๅ–ๅ —ๆตฎๅ–”ๆ•็ฎๅ–”๏ฟฝ');
        return;
    }

    if ((endVal - startVal) > 1000) {
        if (!confirm('ๅ–”ๅจป็ฎๅ–”ะพ็ซพๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”ไฝฎๆๅ–ๅค่ฆๅ–”ๅ็ซตๅ–”ะพ็ฎๅ–”๏ฟฝ 1,000 ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ ๅ–”๏ฟฝ่ฆๅ–”ๅ —็ฎ–ๅ–”ๅจป็ฎๅ–โฌๅ–”ะพๅผ—ๅ–”ไพง็ซธๅ–”่ตค็ฌๅ–”ะพ็ฌ“ๅ–”ๆฌ่ฆๅ–”๏ฟฝ ๅ–”โ‘ง้ๅ–”ๆฌๆถชๅ–”็ผ–็ฌๅ–”ๆคธ่ตๅ–”ๆ•็ฎๅ–”๏ฟฝ?')) return;
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
        box.innerHTML = `<div class="result-success" style="padding:10px;">้๏ฟฝ ๅ–”ๅฆ็ฆๅ–”ๆฐ็ฌๅ–ๅคๆๅ–”๏ฟฝ! ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฆๅ–”ไพงๆถชๅ–”ไฝฎ่ฆๅ–”๏ฝ€็ฌ—ๅ–”ไฝฎๆ–งๅ–”ใ ็ฎๅ–”๏ฟฝ</div>`;
    } else {
        let html = `
                <div class="result-error" style="padding:10px; margin-bottom:10px;">
                    <strong>้ฟ็…็ฌ ๅ–”็ง็ฌๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆๅ–”๏ฟฝ่ฆๅ–”โ‘ง็ฎๅ–”๏ฟฝ ${missing.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ</strong>
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
    if (!confirm('ๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพๅ–”ไฝฎ่ฆๅ–”๏ฝ€ๅผ—ๅ–ๅค่ฆๅ–”ๅ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฌๅ–”่ตค็ฎ‘ๅ–”ๅ•็ฎๅ–”ไพง็ฌๅ–”็ผ–็ฎๅ–”ๅๆ–งๅ–”โด็ฌ–ๅ–”๏ฟฝ็ฆๅ–”็ฒช่…‘ๅ–ๅฆๆตฎๅ–๏ฟฝ?')) return;

    rawTrackingData = [];
    currentImportedBatches = [];
    importedFileCount = 0; // Reset Limit
    document.getElementById('import-preview').classList.add('hidden');
    document.getElementById('upload-status').innerText = 'ๅ–”ใ ็ฎๅ–”ไพง็ซพๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–โฌๅ–”๏ฝ€ๅ…ๅ–”โ‘ง็ฌๅ–”๏ฝ€็ฎๅ–”๏ฟฝๆถช (Ready)';
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
            alert(`้ฟ็…็ฌ ๅ–”ๅ —่ตๅ–”ไฝฎๅฏๅ–”ๆ–ท็ซตๅ–”ไพง็ฆๅ–”๏ฟฝๅฏๅ–”ๆถ็ฎ“ๅ–”๏ฟฝๅผ—ๅ–”ๆ–ท้ๅ–”็็ซพๅ–”๏ฟฝ็ถๅ–”๏ฟฝ 2 ๅ–ๅฆ็ฌฉๅ–”ใ ็ฎคๅ–”๏ฟฝ่ตๅ–”๏ฟฝ็ฆๅ–”็ผ–็ฌๅ–”ๆฐๅฏๅ–”ๅตฟ็ฌๅ–”ๆ็ฌๅ–”็ผ–็ฎๅ–”ะพ็ฎๅ–”ๆฒ‘n(ๅ–”ๅฆ็ถๅ–”ๆ’ช่…‘ๅ–”็ผ–็ฌกๅ–ๅ•ๆ–งๅ–”ใ ็ฌ–ๅ–ๅฆ็ฌกๅ–ไฝฎๅผ—ๅ–ๅคๆ ${importedFileCount} ๅ–ๅฆ็ฌฉๅ–”ใ ็ฎค, ๅ–”็งๆถชๅ–”ไพงๆถชๅ–”ไพงๆตฎๅ–โฌๅ–”็งๅคๅ–ๅ —ๆตฎๅ–”๏ฟฝๅ…ๅ–”๏ฟฝ ${files.length} ๅ–ๅฆ็ฌฉๅ–”ใ ็ฎค)\n\nๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพงๅผ—ๅ–ๅค่ฆๅ–”ๅ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฎ‘ๅ–”ไฝฎ็ฎๅ–”ไพง็ซตๅ–ๅ —่…‘ๅ–”ๆฌๆ–งๅ–”ไพง็ซตๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพๅ–”ไฝฎ่ฆๅ–”๏ฝ€็ฎ‘ๅ–”๏ฝ€ๅคๅ–ๅ —ๆตฎๅ–ๅๆ–งๅ–”โด็ฎ`);
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
    document.getElementById('upload-status').innerText = `ๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅ่…‘ๅ–ๅ —่ฆๅ–”ๆฌ็ฎๅ–”็ฐๅผ—ๅ–๏ฟฝ Excel...`;
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
            alert('ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฎ–ๅ–”ๆฌ็ฎๅ–”็ฐๅผ—ๅ–๏ฟฝ Excel (ๅ–”ๆ•็ฆๅ–”ะพ็ฌ€ๅ–”๏ฟฝ่…‘ๅ–”ๆฐ็ซธๅ–”๏ฟฝๅผ—ๅ–”็ผ–ๆตฎๅ–”ๆฌ็ฎค C)');
            document.getElementById('upload-status').innerText = 'ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”๏ฟฝ';
            return;
        }

        // Cumulative Append
        rawTrackingData.push(...newItems);

        document.getElementById('upload-status').innerText = `ๅ–”๏ฟฝ็ฎๅ–”ไพง็ฌ Excel ๅ–”๏ฟฝ่ตๅ–โฌๅ–”๏ฝ€็ฎๅ–”๏ฟฝ! ๅ–โฌๅ–”็งๅคๅ–ๅ —ๆตฎ ${newItems.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ (ๅ–”๏ฝ€ๆๅ–”๏ฟฝ ${rawTrackingData.length})`;
        analyzeImportedRanges(rawTrackingData);
    };
    reader.readAsArrayBuffer(file);
}

// Updated to handle Multiple Images
async function handleImageImport(files) {
    // Check if Admin
    const urlParams = new URLSearchParams(window.location.search);
    if (!urlParams.has('admin')) {
        alert('ๅ–”็ฐๅ…ๅ–โฌๅ–”ๅ —่…‘ๅ–”๏ฝ€็ฎค "ๅ–”ๆฌ่ตๅ–โฌๅ–”ๅ•็ฎๅ–”ไพง็ฆๅ–”็็ฌกๅ–”็ฉ่ฆๅ–”๏ฟฝ" ๅ–”๏ฟฝ็ซพๅ–”ะพ็ฌๅ–”๏ฟฝๅคๅ–”ๆคธ็ฌๅ–”่็ฎคๅ–”๏ฟฝ่ตๅ–”๏ฟฝ็ฆๅ–”็ผ–็ฌ Admin ๅ–โฌๅ–”ๆคธ็ฎๅ–”ไพง็ฌๅ–”็ผ–็ฎๅ–”ๆฉฝn(ๅ–”็ง็ฌๅ–”็ผ–็ซตๅ–”ๅ่ฆๅ–”ๆฌ็ฌๅ–”็ผ–็ฎๅ–”ะพ็ฎๅ–”ๆถ็ซตๅ–”๏ฝ€็ถๅ–”ๆ’ช่ฆๅ–ๅ็ฌๅ–ๅค็ฎๅ–”็ฐๅผ—ๅ–๏ฟฝ Excel)');
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
                if (confirm(`ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฎ–ๅ–”ๆฌ็ฌญๅ–”ไพง็ฌง ๅ–ไฝฎ็ฌ—ๅ–ๅ —็ฌงๅ–”ๆฐๆถชๅ–”๏ฟฝ็ฌ–ๅ–โฌๅ–”ๅๅคๅ–”๏ฟฝ ${prices.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ\nๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพๅ–”ไฝฎ่ฆๅ–”๏ฝ€้ๅ–”๏ฝ€็ฎๅ–”ไพง็ซพๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซพๅ–”ไพง็ฌๅ–ไฝฎๆถชๅ–”ไฝฎ็ฌ—ๅ–”ไพงๆตฎๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆๅ–”๏ฟฝๅคๅ–”ๆฌ็ซธๅ–ๅค่ฆๅ–ไฝฎ็ฌๅ–”ๆฌๆ–งๅ–”๏ฝ€้ๅ–”๏ฟฝ็ฎๅ–”โด็ฎ?`)) {
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
                    statusEl.innerText = `ๅ–”ะพๅคๅ–โฌๅ–”ๅฆ็ฆๅ–”ไพง่ตดๅ–”๏ฟฝ็ฎคๅ–”โ‘ง่…‘ๅ–”ๆ–ท็ฎ‘ๅ–”ๅๅคๅ–”ๆฌ้ๅ–”่ตค็ฎ‘ๅ–”๏ฝ€็ฎๅ–”๏ฟฝ (${summary.totalCount} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ)`;
                    return;
                }
            }

            alert('ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฎ–ๅ–”ๆฌ็ฆๅ–”็็ฌกๅ–”็ฉ่ฆๅ–”็ง็ฌๅ–”ๆ็ฎ');
            statusEl.innerText = 'ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฎ–ๅ–”ๆฌ็ฌญๅ–”ไพง็ฌงๅ–”ใ ็ฎๅ–”ไพง้ๅ–”่็ฌ–';
            return;
        }

        // Cumulative Append
        rawTrackingData.push(...newItems);

        statusEl.innerText = `OCR ๅ–โฌๅ–”๏ฟฝ็ฆๅ–ๅ็ฌ€ๅ–”๏ฟฝๅคๅ–ๅค็ฌ! ๅ–โฌๅ–”็งๅคๅ–ๅ —ๆตฎ ${newItems.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ (ๅ–”๏ฝ€ๆๅ–”๏ฟฝ ${rawTrackingData.length})`;
        analyzeImportedRanges(rawTrackingData);

    } catch (err) {
        console.error(err);
        alert('ๅ–โฌๅ–”ไฝฎๅคๅ–”ๆ–ท็ซถๅ–ๅค่…‘ๅ–”ๆบนๅคๅ–”ๆ–ท็ฌงๅ–”ใ ่ฆๅ–”ๆ–ท็ฎ–ๅ–”ๆฌ็ซตๅ–”ไพง็ฆๅ–”๏ฟฝ็ฎๅ–”ไพง็ฌๅ–”๏ฝ€่…นๅ–”ๆถ็ฌญๅ–”ไพง็ฌง: ' + err.message);
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

            return `<li>${rangeText} (${m.count} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ)</li>`;
        }).join('');

        gapHtml = `
            <div class="result-error" style="margin-top:15px; padding:15px; border:2px solid #ff4444; background:#ffebeb;">
                <h3 style="margin-top:0; color:#cc0000;">้ฟ็…็ฌ ๅ–”ๆ•็ฆๅ–”ะพ็ฌ€ๅ–”็ง็ฌๅ–โฌๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ซถๅ–ๅค่ฆๅ–”โด็ฎๅ–”๏ฟฝ (GAP DETECTED)</h3>
                <p>ๅ–”๏ฝ€่ตดๅ–”ๆฐ็ฌๅ–”็ง็ฌๅ–”ะพ็ฎๅ–”ไพง็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฎๅ–”โด็ฎๅ–”ๆ•็ฎๅ–”๏ฟฝ็ฎ‘ๅ–”ๆฌ้ๅ–ๅ —่…‘ๅ–”๏ฟฝ <strong>ๅ–”๏ฟฝ่ฆๅ–”โ‘ง็ฎๅ–”๏ฟฝ ${totalMissing} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ</strong> ๅ–”ๆ–ทๅฏๅ–”ๅ็ฌๅ–”ๆ็ฎ:</p>
                <ul style="margin-bottom:0;">${listHtml}</ul>
                <div style="font-size:0.9rem; color:#666; margin-top:10px;">ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆๅ–”ๆคธๅ…ๅ–ๅ —ๆ–งๅ–”ไพงๆถชๅ–ๅฆ็ฌกๅ–”ๆ ข่…นๅ–”ไฝฎ็ฎ‘ๅ–”็งๅคๅ–ๅ —ๆตฎๅ–”ใ ็ซพๅ–ๅ็ฌๅ–”ๆ•่ฆๅ–”๏ฝ€่ฆๅ–”ๅ็ฌ–ๅ–ๅค่ฆๅ–”ๆฌๅผ—ๅ–ๅ —่ฆๅ–”ๅ็ฎ‘ๅ–”็ง้ๅ–ๅ —่…‘ๅ–”๏ฟฝ็ฎๅ–”ไพง็ซพๅ–”๏ฟฝๅคๅ–”ๅ็ฎ’ๅ–”ใ ็ฎๅ–”๏ฟฝ (ๅ–”๏ฟฝๅ…ๅ–ไฝฎ็ฌ–ๅ–”๏ฟฝ)</div>
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
                            <span>้๏ฟฝ ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆๅ–”ๆคธๅ…ๅ–ๅ —ๆ–งๅ–”ไพงๆถชๅ–ๅฆ็ฌก (Missing)</span>
                            <span class="mobile-stats" style="color:#d32f2f;">${m.count} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ</span>
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
                <strong>้๏ฟฝ ๅ–”ๅฆ็ฆๅ–”ๆฐ็ฌๅ–ๅคๆๅ–”๏ฟฝ 100% (No Gaps)</strong> - ๅ–โฌๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฎ‘ๅ–”๏ฝ€ๅ…ๅ–”โ‘ง็ซพๅ–”ๆ•็ฎๅ–”๏ฟฝ็ฎ‘ๅ–”ๆฌ้ๅ–ๅ —่…‘ๅ–”ๅ็ซตๅ–”็ผ–็ฌๅ–”๏ฟฝๆตฎๅ–”ๆฐ่…นๅ–”๏ฝ€็ฌ“ๅ–๏ฟฝ
            </div>
        `;
    }

    // --- DISCREPANCY ALERT SECTION ---
    let discrepancyHtml = '';
    if (discrepancies && discrepancies.length > 0) {
        discrepancyHtml = `
            <div class="result-error" style="margin-top:15px; padding:10px; border:2px solid #ff9800; background:#fff3e0; color:#e65100;">
                <h4 style="margin:0 0 5px 0;">้ฟ็…็ฌ ๅ–”็ง็ฌๅ–”ๅ•็ฎๅ–”๏ฟฝ้ๅ–”็ผ–็ซพๅ–โฌๅ–”ไฝฎ็ฌ—ๅ–”ๅ —่ฆๅ–”ไฝฎ็ฎๅ–”็ฐๅผ—ๅ–ๅฑถ็ฌๅ–”่ตค็ฎ‘ๅ–”ๅ•็ฎๅ–”๏ฟฝ (Inconsistencies)</h4>
                <p style="margin:0; font-size:0.95rem;">
                    ๅ–”็ง็ฌๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”ๆฌ็ฎๅ–”่ตคๆ–งๅ–”ๆฌๅฏๅ–”ไฝฎ็ฌๅ–”ๆ็ฎๅ–”๏ฝ€่ตดๅ–”ๆฐ็ถๅ–ๅ็ฌๅ–ๅฆ็ฌฉๅ–”ใ ็ฎค <strong>ๅ–ๅฆๆตฎๅ–ๅ —้ๅ–”็ผ–ๆตฎๅ–”็งๅฏๅ–”ๆฌ็ฌๅ–ๅฑถ็ซตๅ–”็ผ–็ฌๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆๅ–”ๆ•่ฆๅ–”โด็ฌ—ๅ–”ไพง็ฆๅ–”ไพง็ซพๅ–”ๅฆ็ฎๅ–”ไพง็ฌๅ–”๏ฝ€ๅคๅ–”ไฝฎ่ฆๅ–”๏ฟฝ</strong> ๅ–”ๅ —่ตๅ–”ๆฌๆๅ–”๏ฟฝ <strong>${discrepancies.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ</strong><br>
                    <small>ๅ–โฌๅ–”ๅจป็ฎๅ–”๏ฟฝ ๅ–ๅ็ฌๅ–ๅฆ็ฌฉๅ–”ใ ็ฎคๅ–”๏ฝ€่ตดๅ–”ๆฐ็ถ ${discrepancies[0].originalWeight} ๅ–ไฝฎ็ฌ—ๅ–ๅ —็ฆๅ–”ไพง็ซธๅ–”๏ฟฝ ${discrepancies[0].price} ๅ–”ๆฐ่ฆๅ–”๏ฟฝ (ๅ–”๏ฝ€่ตดๅ–”ๆฐ็ฌๅ–ๅฆ็ฌ–ๅ–ๅค็ฌกๅ–”๏ฝ€ๅฏๅ–”ๆฐ็ฎ’ๅ–”ไฝฎ็ฎๅ–โฌๅ–”ๆถ็ฎๅ–”๏ฟฝ ${discrepancies[0].weight} ๅ–”๏ฟฝๅฏๅ–”ๆ•็ฎ“ๅ–”ๆฌๆตฎๅ–”็ผ–็ฌ—ๅ–”่็ฎ’ๅ–”ใ ็ฎๅ–”๏ฟฝ)</small>
                </p>
            </div>
        `;
    }

    summary.innerHTML = `
        <strong>้ฆๆณ ๅ–”๏ฟฝ็ฆๅ–”่็ฌกๅ–”ๆบนๅผ—ๅ–”ไฝฎ่ฆๅ–”๏ฝ€ๆๅ–”่็ฎ‘ๅ–”ๅฆ็ฆๅ–”ไพง่ตดๅ–”๏ฟฝ็ฎค (Virtual Optimization)</strong><br>
        ๅ–”ๅ —่ตๅ–”ๆฌๆๅ–”ๆฌ็ฌๅ–”็ผ–็ฎๅ–”ๅๆ–งๅ–”โด็ฌ–: ${totalItems.toLocaleString()} ๅ–”ๅจปๅคๅ–ๅค็ฌ<br>
        ๅ–”โ‘ง่…‘ๅ–”ๆ–ท็ฎ‘ๅ–”ๅๅคๅ–”ๆฌ็ฆๅ–”ะพๆตฎ: <span style="font-size:1.2rem; color:#d63384; font-weight:bold;">${grandTotal.toLocaleString()} ๅ–”ๆฐ่ฆๅ–”๏ฟฝ</span>
        ${gapHtml}
        ${discrepancyHtml}
    `;

    // Generate Receipt-style Table
    let html = `
        <div style="background:white; padding:20px; border:1px solid #ddd; box-shadow:0 2px 5px rgba(0,0,0,0.05); font-family:'Courier New', monospace;">
            <h4 style="text-align:center; border-bottom:1px dashed #ccc; padding-bottom:10px; margin-bottom:10px;">ๅ–ๅ็ฌๅ–”๏ฟฝ็ฆๅ–”่็ฌกๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ (Optimized Report)</h4>
             <div style="font-size:0.8rem; color:red; text-align:center; margin-bottom:5px;">
                *ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆๅ–”ๆ ข่…นๅ–”ไฝฎ็ฌ€ๅ–”็ผ–็ฌ–ๅ–โฌๅ–”๏ฝ€ๅ…ๅ–”โ‘ง็ซพๅ–ๅๆ–งๅ–”โด็ฎๅ–”ๆ•่ฆๅ–”โด็ฆๅ–”ไพง็ซธๅ–”ไพง็ฌๅ–ๅค่…‘ๅ–”๏ฟฝ-ๅ–”โด่ฆๅ–”๏ฟฝ (Virtual Mapping)
            </div>
            
            <table style="width:100%; border-collapse: collapse;">
                <thead>
                    <tr style="border-bottom:2px solid #000;">
                        <th style="text-align:left; padding:5px;">ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ (Description)</th>
                        <th class="col-qty" style="text-align:right; padding:5px; white-space:nowrap;">ๅ–”ๅ —่ตๅ–”ๆฌๆๅ–”๏ฟฝ (Qty)</th>
                        <th class="col-price" style="text-align:right; padding:5px; white-space:nowrap;">ๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆ/ๅ–”ๅจปๅคๅ–ๅค็ฌ</th>
                        <th class="col-total" style="text-align:right; padding:5px; white-space:nowrap;">ๅ–”๏ฝ€ๆๅ–”๏ฟฝ (Total)</th>
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
                        <strong>${idx + 1}. EMS ๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆ ${r.price} ๅ–”ๆฐ่ฆๅ–”๏ฟฝ</strong>
                        <span class="mobile-stats" style="color:#d63384; font-weight:bold;">${r.count} ๅ–”ๅจปๅคๅ–ๅค็ฌ${displayTotal}</span>
                    </div>
                    
                    <!-- Line 2: Range Only (Mobile) -->
                    <div class="line-flex">
                        <span style="color:#0056b3; font-weight:bold; overflow-wrap:break-word; max-width:100%;">
                            ${r.start === r.end
                ? TrackingUtils.formatTrackingNumber(r.start)
                : `${TrackingUtils.formatTrackingNumber(r.start)} - ${TrackingUtils.formatTrackingNumber(r.end)}`}
                        </span>
                    </div>

                    <small style="color:#666;">ๅ–”ๆฌ็ฎๅ–”่ตคๆ–งๅ–”ๆฌๅฏๅ–”๏ฟฝ (Weight): ${r.weight}</small>
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
                                    <span style="font-weight:bold;">ๅ–”๏ฝ€ๆๅ–”โด็ฌๅ–”็ผ–็ฎๅ–”ๅ้ๅ–”่็ฎๅ–”๏ฟฝ (Grand Total)</span>
                                    <span style="font-weight:bold; font-size:1.1rem;">${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                                </div>
                            </td>
                        </tr>
                    </tfoot>
                </table>
            <div style="font-size:0.8rem; color:#999; text-align:center; margin-top:5px; font-style:italic;">
                (ๅ–”ๆฐ็ฌๅ–”โด้ๅ–”๏ฟฝ็ฌๅ–”็ฒช่…‘: ๅ–”ๅ —่ตๅ–”ๆฌๆๅ–”ๆฌ็ฎ’ๅ–”ใ ่ตดๅ–”โ‘ง่…‘ๅ–”ๆ–ท็ฎ‘ๅ–”ๅๅคๅ–”ๆฌ็ฌ€ๅ–”็ญ็ฎ’ๅ–”๏ฟฝ็ฌ–ๅ–”ๅ็ฌๅ–”ๆ็ฎๅ–”โด็ถๅ–”โด็ซถๅ–”ะพ่ฆๅ–”ๅ•่…‘ๅ–”ๅ็ฆๅ–”ไพงๆถชๅ–”ไฝฎ่ฆๅ–”๏ฟฝ)
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
        alert('ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพง็ฆๅ–”็ญ็ฌๅ–”่็ฌๅ–”็ฒช็ฎๅ–”๏ฟฝ็ซตๅ–”ใ ็ถๅ–ๅ —ๆตฎๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ— (Batch Name)');
        return;
    }

    if (!isAuto) {
        if (!confirm(`ๅ–”โ‘ง้ๅ–”ๆฌๆถชๅ–”็ผ–็ฌๅ–”ๆฐๅฏๅ–”ๆฌ็ฌๅ–”ๅค็ซตๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ— ${currentImportedBatches.length} ๅ–”ๅจป็ฎๅ–”ะพ็ซพๅ–”๏ฝ€ๆๅ–”โด็ซตๅ–”็ผ–็ฌๅ–โฌๅ–”ๆถ็ฎๅ–”๏ฟฝ 1 ๅ–”ๅจป็ถๅ–”ๆ–ท็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”๏ฟฝ?`)) return;
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
        // alert(`้๏ฟฝ ๅ–”๏ฝ€่ตดๅ–”ๆฐ็ฌๅ–”ๆฐๅฏๅ–”ๆฌ็ฌๅ–”ๅค็ซตๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”๏ฟฝๅฏๅ–”ๆ•็ฎ“ๅ–”ๆฌๆตฎๅ–”็ผ–็ฌ—ๅ–”่็ฎ‘ๅ–”๏ฝ€ๅ…ๅ–”โ‘ง็ฌๅ–”๏ฝ€็ฎๅ–”๏ฟฝๆถช!\n(Optimized ${rangesMeta.length} Groups)`);
        // Silent or small notification? User wants to SEE it.
    } else {
        alert(`ๅ–”ๆฐๅฏๅ–”ๆฌ็ฌๅ–”ๅค็ซตๅ–โฌๅ–”๏ฝ€ๅ…ๅ–”โ‘ง็ฌๅ–”๏ฝ€็ฎๅ–”๏ฟฝๆถช!`);
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
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฌกๅ–”๏ฝ€่ตดๅ–”ะพๅฏๅ–”ๆ•ๅคๅ–”ไฝฎ่ฆๅ–”๏ฝ€็ฌๅ–”่ตค็ฎ‘ๅ–”ๅ•็ฎๅ–”๏ฟฝ</td></tr>';
        return;
    }

    list.forEach(item => {
        const dateStr = new Date(item.timestamp).toLocaleString('th-TH');
        const tr = document.createElement('tr');

        let deleteBtn = '';
        if (isAdmin) {
            deleteBtn = `
                <button class="btn btn-danger" style="padding:4px 8px; font-size:0.8rem; margin-left:5px;" 
                    onclick="deleteHistoryItem('${item.id}', '${item.name}')">้ฆๆฃ้””๏ฟฝ ๅ–”ใ ็ฌ</button>
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
                    onclick="loadBatchToView('${item.id}')">้ฆๆ”ท ๅ–”ๆ–ท่…น</button>
                ${deleteBtn}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function deleteHistoryItem(batchId, batchName) {
    if (confirm(`ๅ–”โ‘ง้ๅ–”ๆฌๆถชๅ–”็ผ–็ฌๅ–”ไฝฎ่ฆๅ–”๏ฝ€ๅผ—ๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”๏ฟฝ (ๅ–”ๅ —่ฆๅ–”ไฝฎ็ฌกๅ–”๏ฝ€่ตดๅ–”ะพๅฏๅ–”ๆ•ๅค Import)?\n\nๅ–”ๅจป็ถๅ–”ๆ–ท็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”๏ฟฝ: ${batchName}`)) {
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
        list.innerHTML = '<li>ๅ–ๅฆๆตฎๅ–ๅ —ๆตฎๅ–”ๆ็ฌ€ๅ–”่็ฌ–ๅ–”โ‘ง็ฎๅ–”๏ฟฝ็ฌๅ–”ไฝฎๅผ—ๅ–”็ผ–็ฌ (No Snapshots)</li>';
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
    if (confirm('ๅ–”โ‘ง้ๅ–”ๆฌๆถชๅ–”็ผ–็ฌๅ–”โ‘ง็ฎๅ–”๏ฟฝ็ฌๅ–โฌๅ–”ะพๅผ—ๅ–”ไพง็ซตๅ–”ใ ๅฏๅ–”ๆฐ็ฎๅ–”ๆถ็ฌ€ๅ–”่็ฌ–ๅ–”ๆฌๅ…ๅ–๏ฟฝ? (ๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”ๆถๅฏๅ–”ๅ —็ฌ€ๅ–”่็ฌๅ–”็ผ–็ฌๅ–”ๅ —่ตดๅ–”๏ฟฝ่ฆๅ–”โ‘ง็ฎๅ–”๏ฟฝ)')) {
        if (CustomerDB.restoreSnapshot(ts)) {
            alert('้๏ฟฝ ๅ–”โ‘ง็ฎๅ–”๏ฟฝ็ฌๅ–โฌๅ–”ะพๅผ—ๅ–”ไพง้ๅ–”่ตค็ฎ‘ๅ–”๏ฝ€็ฎๅ–”๏ฟฝ (Restored)');
            renderDBTable();
            renderImportHistory(); // if visible
        } else {
            alert('้๏ฟฝ ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”๏ฟฝ Snapshot ๅ–”ๆฌๅ…ๅ–๏ฟฝ');
        }
    }
}

// --- Backup & Restore Glue Code ---
function confirmAndBackup() {
    if (confirm('ๅ–”ๅฆ็ถๅ–”ๆ’ช็ฌ—ๅ–ๅค่…‘ๅ–”ๅ็ซตๅ–”ไพง็ฆ Export ๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”ๆคธๅฏๅ–ๅค็ซพๅ–”๏ฟฝๆตฎๅ–”ๆ–ท่…‘ๅ–”๏ฟฝ็ซตๅ–”โด่ฆๅ–โฌๅ–”ๆถ็ฎๅ–”ๆฌ็ฎๅ–”็ฐๅผ—ๅ–ๅฑถ็ฎ–ๅ–”ๅจป็ฎๅ–”๏ฟฝ็ฆๅ–”็ฒช่…‘ๅ–ๅฆๆตฎๅ–๏ฟฝ?')) {
        backupData();
    }
}

function backupData() {
    CustomerDB.exportBackup();
}

function restoreData(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (!confirm('ๅ–”ๅฆ่ตๅ–โฌๅ–”ๆ•้ๅ–”๏ฟฝ็ฌ: ๅ–”ไฝฎ่ฆๅ–”๏ฝ€็ซตๅ–”็็ฎๅ–”ๅฆ้ๅ–”ๆฌ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฌ€ๅ–”็ญ็ฌๅ–”็ผ–็ฌๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”ๆถๅฏๅ–”ๅ —็ฌ€ๅ–”่็ฌๅ–”็ผ–็ฌๅ–”ๆคธๅฏๅ–ๅค็ซพๅ–”๏ฟฝๆตฎๅ–”ๆ“ปnๅ–”ๅฆ็ถๅ–”ๆ’ช็ฌ—ๅ–ๅค่…‘ๅ–”ๅ็ซตๅ–”ไพง็ฆๅ–”ๆ–ท่ตๅ–โฌๅ–”ๆฌๅคๅ–”ๆฌ็ซตๅ–”ไพง็ฆๅ–”ๆ•็ฎๅ–”๏ฟฝๆ–งๅ–”๏ฝ€้ๅ–”๏ฟฝ็ฎๅ–”โด็ฎ?')) {
        event.target.value = ''; // Reset
        return;
    }

    CustomerDB.importBackup(file)
        .then(() => {
            alert('้๏ฟฝ ๅ–”ไฝฎ่…นๅ–ๅค็ซธๅ–”็ฒช็ฌๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”๏ฟฝ่ตๅ–โฌๅ–”๏ฝ€็ฎๅ–”๏ฟฝ (Restore Complete)');
            if (typeof updateDbViews === 'function') updateDbViews();
            else if (typeof renderDBTable === 'function') renderDBTable(); // Refresh UI
        })
        .catch(err => {
            alert('้๏ฟฝ ๅ–โฌๅ–”ไฝฎๅคๅ–”ๆ–ท็ซถๅ–ๅค่…‘ๅ–”ๆบนๅคๅ–”ๆ–ท็ฌงๅ–”ใ ่ฆๅ–”๏ฟฝ: ' + err.message);
        })
        .finally(() => {
            event.target.value = ''; // Reset
        });
}

function loadBatchToView(batchId) {
    const batches = CustomerDB.getBatches();
    const batch = batches[batchId];

    if (!batch) {
        alert('ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฌๅ–”่็ฌ–ๅ–”ๆฌๅ…ๅ–๏ฟฝ');
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
                <strong>้ฆๆจ ๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—: ${batch.name}</strong>
                <button class="btn btn-neutral" onclick="switchTab('import')" style="padding:5px 10px; font-size:0.9rem;">็ฌ๏ฟฝ ๅ–”ไฝฎๅผ—ๅ–”็ผ–็ฌๅ–”๏ฟฝ็ฌๅ–ๅค่ฆๅ–”๏ฟฝๅผ—ๅ–”็ผ–็ซต (New Import)</button>
            </div>
            
            <!-- Receipt View -->
            <div style="background:white; padding:20px; border:1px solid #ddd; box-shadow:0 2px 5px rgba(0,0,0,0.05); font-family:'Courier New', monospace; max-width:800px; margin:0 auto;">
                <h4 style="text-align:center; border-bottom:1px dashed #ccc; padding-bottom:10px; margin-bottom:10px;">ๅ–ๅ็ฌๅ–”๏ฟฝ็ฆๅ–”่็ฌกๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ (Optimized Report)</h4>
                 <div style="margin-bottom:10px; font-size:0.9rem;">
                    <strong>Customer:</strong> ${batch.name}<br>
                    <strong>Type:</strong> ${batch.type}<br>
                    ${batch.requestDate ? `<strong>Request Date:</strong> ${new Date(batch.requestDate).toLocaleDateString('th-TH')}<br>` : ''}
                    <strong>Date:</strong> ${new Date(batch.timestamp).toLocaleString('th-TH')}
                </div>
                 <div style="font-size:0.8rem; color:red; text-align:center; margin-bottom:5px;">
                    *ๅ–”ๅ —ๅฏๅ–”ๆ–ท็ฎ‘ๅ–”๏ฝ€ๅ…ๅ–”โ‘ง็ซพๅ–”ๆ•่ฆๅ–”โด็ฆๅ–”ไพง็ซธๅ–”ไพง็ฌๅ–ๅค่…‘ๅ–”๏ฟฝ-ๅ–”โด่ฆๅ–”๏ฟฝ (Virtual)
                </div>
                <div class="report-list-container">
                    <div class="report-header-desktop">
                        <div style="width:40%">ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ (Description)</div>
                        <div style="width:20%; text-align:right">ๅ–”ๅ —่ตๅ–”ๆฌๆๅ–”๏ฟฝ (Qty)</div>
                        <div style="width:20%; text-align:right">ๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆ/ๅ–”ๅจปๅคๅ–ๅค็ฌ</div>
                        <div style="width:20%; text-align:right">ๅ–”๏ฝ€ๆๅ–”๏ฟฝ (Total)</div>
                    </div>
        `;



        batch.ranges.forEach((r, idx) => {
            html += `
                <div class="report-card">
                    <div class="report-card-desc">
                        <strong>${idx + 1}. EMS ๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆ ${r.price} ๅ–”ๆฐ่ฆๅ–”๏ฟฝ</strong><br>
                        <span style="color:#0056b3; font-weight:bold;">${r.start === r.end ? r.start : `${r.start} - ${r.end}`}</span><br>
                        <small>ๅ–”ๆฌ็ฎๅ–”่ตคๆ–งๅ–”ๆฌๅฏๅ–”๏ฟฝ (Weight): ${r.weight}</small>
                    </div>
                    <div class="report-card-scroll">
                        <div class="stat-item">
                            <span class="stat-label">ๅ–”ๅ —่ตๅ–”ๆฌๆๅ–”๏ฟฝ (Qty)</span>
                            <span class="stat-value">${r.count}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">ๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆ/ๅ–”ๅจปๅคๅ–ๅค็ฌ</span>
                            <span class="stat-value">@${r.price}</span>
                        </div>
                        <div class="stat-item highlight">
                            <span class="stat-label">ๅ–”๏ฝ€ๆๅ–”๏ฟฝ (Total)</span>
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
                        <div class="footer-label">ๅ–”๏ฝ€ๆๅ–”โด็ฌๅ–”็ผ–็ฎๅ–”ๅ้ๅ–”่็ฎๅ–”๏ฟฝ (Grand Total)</div>
                        <div class="footer-value">${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</div>
                    </div>
                </div>
            </div>
            
            <div style="text-align:center; margin-top:20px;">
                <button class="btn" onclick="window.print()">้ฆๆฟ้””๏ฟฝ Print / PDF</button>
                <button class="btn btn-neutral" onclick="switchTab('import')" style="margin-left:10px;">็ฌ๏ฟฝ ๅ–”ไฝฎๅผ—ๅ–”็ผ–็ฌๅ–”๏ฟฝ็ฌๅ–ๅค่ฆๅ–”๏ฟฝๅผ—ๅ–”็ผ–็ซต (New Import)</button>
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
        alert("ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพง็ซตๅ–”๏ฝ€่…‘ๅ–”ไฝฎ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”๏ฟฝ");
        return;
    }
    
    // Extract all tracking numbers using our robust parser
    const extracted = TrackingUtils.extractTrackingNumbers(rawInput);
    
    if (extracted.length === 0) {
        alert("ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฆๅ–”็็ฌกๅ–ไฝฎ็ฌๅ–”ๆฐ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฌๅ–”ๆ็ฎๅ–”ๆ ข่…นๅ–”ไฝฎ็ฌ—ๅ–ๅค่…‘ๅ–”๏ฟฝ (13 ๅ–”๏ฟฝๅผ—ๅ–”็ผ–็ซต)");
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
    
    if (statusEl) statusEl.textContent = `ๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅ็ฎ‘ๅ–”๏ฝ€ๅคๅ–ๅ —ๆตฎๅ–”ๆถ็ฆๅ–”็ญๆตฎๅ–”ะพๅผ—ๅ–”ๆบนๅผ—...`;

    let combinedText = "";
    let debugProcessedImageUrl = "";

    try {
        for (let i = 0; i < files.length; i++) {
            if (statusEl) statusEl.textContent = `ๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅ็ฎ’ๅ–”โ‘ง็ซตๅ–”ๅ•็ฎๅ–”๏ฟฝ็ซธๅ–”ะพ่ฆๅ–”โด็ฌ€ๅ–”ไพง็ซตๅ–”็ฉ่ฆๅ–”๏ฟฝ ${i + 1}/${files.length}...`;
            const file = files[i];

            // Image Preprocessing: Grayscale + Thresholding
            if (statusEl) statusEl.textContent = `ๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅ็ฌกๅ–”๏ฝ€ๅฏๅ–”ๆฐ็ฎ’ๅ–”ๆ•็ฎๅ–”ๅ็ฆๅ–”็็ฌกๅ–”็ฉ่ฆๅ–”๏ฟฝ (Preprocessing)...`;
            const processedFileUrl = await preprocessImageForOCR(file);
            debugProcessedImageUrl = processedFileUrl; // Save for debug display

            // Tesseract OCR
            if (statusEl) statusEl.textContent = `ๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅ็ฎ’ๅ–”โ‘ง็ซตๅ–”ๅ•็ฎๅ–”๏ฟฝ็ซธๅ–”ะพ่ฆๅ–”โด็ฌ€ๅ–”ไพง็ซตๅ–”็ฉ่ฆๅ–”๏ฟฝ ${i + 1}/${files.length}...`;
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

        if (statusEl) statusEl.textContent = "ๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅๆๅ–”่็ฎ‘ๅ–”ๅฆ็ฆๅ–”ไพง่ตดๅ–”๏ฟฝ็ฎคๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ— (Analyzing)...";

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
            statusEl.textContent = "ๅ–”ๆถ็ฆๅ–”็ญๆตฎๅ–”ะพๅผ—ๅ–”ๆบนๅผ—ๅ–โฌๅ–”๏ฟฝ็ฆๅ–ๅ็ฌ€ๅ–”๏ฟฝๅคๅ–ๅค็ฌ: ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฎ–ๅ–”ๆฌ็ฌญๅ–”ไพง็ฌง";
            resultEl.innerHTML = `<em>ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฌๅ–”ๆ็ฎๅ–”ๆ•็ฆๅ–”ๅ็ซตๅ–”็ผ–็ฌๅ–”๏ฝ€่…นๅ–”ๆถ็ฎ’ๅ–”ๆฐ็ฌๅ–โฌๅ–”ใ ็ซถๅ–ๅฆ็ฌกๅ–”๏ฝ€ไฟฏๅ–”ๆ’ชๅ…ๅ–”โ‘ง็ฎคๅ–ๅฆ็ฌๅ–”๏ฟฝ 13 ๅ–”๏ฟฝๅผ—ๅ–”็ผ–็ซต</em>\n\n[ๅ–”ๅ•็ฎๅ–”๏ฟฝ็ซธๅ–”ะพ่ฆๅ–”โด็ฌ–ๅ–”่็ฌๅ–”ๅ —่ฆๅ–”๏ฟฝ OCR]\n${combinedText}\n<hr><img src="${debugProcessedImageUrl}" style="max-width:100%; border:1px solid #ccc; margin-top:10px;">`;
            return;
        }

        // 2. Output Formatting & Missing Check
        let outputHtml = `<strong style="color:var(--primary-color);">้ฆๆถ ๅ–”ๆ•็ฆๅ–”ะพ็ฌ€ๅ–”็ง็ฌๅ–”ๆคธๅฏๅ–ๅค็ซพๅ–”๏ฟฝๆตฎๅ–”๏ฟฝ ${extractedItems.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ</strong>\n<hr>`;
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
                    <strong>้ฆๆฏ ๅ–ไฝฎ็ฌ€ๅ–ๅค็ซพๅ–โฌๅ–”ๆ•้ๅ–”๏ฟฝ็ฌ: ๅ–”็ง็ฌๅ–”ๅจป็ฎๅ–”๏ฟฝ็ซพๅ–ๅ•ๆ–งๅ–”ะพ็ฎ (ๅ–โฌๅ–”ใ ็ซถๅ–”ๆคธๅ…ๅ–ๅ —ๆ–งๅ–”ไพงๆถชๅ–ๅฆ็ฌก) ๅ–ๅ็ฌๅ–”ๅจป็ฎๅ–”ะพ็ซพๅ–”๏ฟฝๆตฎๅ–”ะพ็ฌ– ${key}:</strong><br>
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
                <th style="padding:5px; text-align:left;">ๅ–”ใ ่ตๅ–”ๆ–ทๅฏๅ–”๏ฟฝ</th>
                <th style="padding:5px; text-align:left;">ๅ–โฌๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”๏ฟฝ</th>
                <th style="padding:5px; text-align:right;">ๅ–”๏ฝ€่ฆๅ–”ๅฆ่ฆ (ๅ–”ๆฐ่ฆๅ–”๏ฟฝ)</th>
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
                <td colspan="2" style="padding:10px; text-align:right;">ๅ–”โ‘ง่…‘ๅ–”ๆ–ท็ฆๅ–”ะพๆตฎๅ–”ๆคธๅฏๅ–ๅค็ซพๅ–”๏ฟฝๆตฎๅ–”๏ฟฝ (Total):</td>
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
                <strong style="color:#666; font-size:0.8rem;">[DEBUG] ๅ–”็ฉ่ฆๅ–”็ง็ฌๅ–”ๆ็ฎๅ–”ๆถ็ฆๅ–”็ผ–็ฌๅ–ไฝฎ้ๅ–”ๅ็ฎ’ๅ–”ใ ็ฎๅ–”๏ฟฝ:</strong><br>
                <img src="${debugProcessedImageUrl}" style="max-width:100%; border:1px solid #ccc; margin-top:5px;">
            </div>
        `;

        resultEl.innerHTML = outputHtml;
        statusEl.textContent = "ๅ–”ๆถ็ฆๅ–”็ญๆตฎๅ–”ะพๅผ—ๅ–”ๆบนๅผ—ๅ–โฌๅ–”๏ฟฝ็ฆๅ–ๅ็ฌ€ๅ–”๏ฟฝๅคๅ–ๅค็ฌ (Done)";

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
        alert('ๅ–ๅฆๆตฎๅ–ๅ —ๆตฎๅ–”ๆ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฎ–ๅ–”๏ฟฝ็ฎๅ–”ๅฆๅฏๅ–”ๆ–ทๅผ—ๅ–”๏ฟฝ็ซตๅ–”ๅฆ็ฆๅ–”็ผ–็ฌ');
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
                btn.innerHTML = `้๏ฟฝ ๅ–”ๅฆๅฏๅ–”ๆ–ทๅผ—ๅ–”๏ฟฝ็ซต ${extracted.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆๅ–ไฝฎๅผ—ๅ–ๅคๆ`;
                btn.classList.add('btn-success');
                setTimeout(() => {
                    btn.innerHTML = originalText;
                    btn.classList.remove('btn-success');
                }, 1500);
            }
            alert(`ๅ–”ๅฆๅฏๅ–”ๆ–ทๅผ—ๅ–”๏ฟฝ็ซตๅ–โฌๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่้ๅ–”่ตค็ฎ‘ๅ–”๏ฝ€็ฎๅ–”๏ฟฝ ${extracted.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ`);
        }).catch(err => {
            alert('ๅ–ๅฆๆตฎๅ–ๅ —้ๅ–”ไพงๆตฎๅ–”ไพง็ฆๅ–”ๆ ข็ซธๅ–”็ผ–็ฌ–ๅ–”ใ ่…‘ๅ–”ไฝฎ็ฎๅ–”ๆ–ท็ฎ: ' + err);
        });
    } else {
        alert('ๅ–”ๅฆ็ฎๅ–”ๆฌๆ–งๅ–”ไพง็ฎๅ–”โด็ฎๅ–”็ง็ฌๅ–”๏ฝ€่…นๅ–”ๆถ็ฎ’ๅ–”ๆฐ็ฌๅ–โฌๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฌๅ–”ๆ็ฎๅ–”ๆ ข่…นๅ–”ไฝฎ็ฌ—ๅ–ๅค่…‘ๅ–”ๅ็ฎ‘ๅ–”็ง้ๅ–ๅ —่…‘ๅ–”ๅฆๅฏๅ–”ๆ–ทๅผ—ๅ–”๏ฟฝ็ซตๅ–”ๅฆ็ฆๅ–”็ผ–็ฌ');
    }
}

async function adminCrossRefImage(files) {
    if (!files || files.length === 0) return;
    
    const statusEl = document.getElementById('admin-crossref-status');
    const resultEl = document.getElementById('admin-crossref-result');
    resultEl.classList.add('hidden');
    
    if (statusEl) statusEl.textContent = `ๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅ้ๅ–ไฝฎ็ซตๅ–”ๆฌ็ฆๅ–”็็ฌกๅ–”็ฉ่ฆๅ–”็ง็ฌ–ๅ–ๅคๆๅ–”๏ฟฝ OCR...`;

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
                 if (statusEl) statusEl.textContent = `ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฎ–ๅ–”ๆฌ็ฆๅ–”็็ฌกๅ–”็ฉ่ฆๅ–”็€;
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
        statusEl.textContent = "ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพงๆๅ–”ไพง็ซพๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ— ๅ–”๏ฟฝ็ฆๅ–”็ฒช่…‘ ๏ฟฝ    resultEl.innerHTML = html;
    resultEl.classList.remove('hidden');
    
    statusEl.innerHTML = `\u0e15\u0e23\u0e27\u0e08\u0e1e\u0e1a\u0e2a\u0e31\u0e07\u0e01\u0e31\u0e14\u0e15\u0e23\u0e07\u0e01\u0e31\u0e19 <strong style="color:green;">${foundCount}</strong> \u0e08\u0e32\u0e01\u0e17\u0e31\u0e49\u0e07\u0e2b\u0e21\u0e14 ${trackingArray.length} \u0e23\u0e32\u0e22\u0e01\u0e32\u0e23`;
}
ไฝฎ็ฎๅ–”๏ฟฝ็ฌๅ–”๏ฟฝ็ฌๅ–ๅค่ฆ ${Math.abs(item.offset)})</span>`;
            if (item.offset > 0) label = `<span style="color:#1976d2; font-size:0.75rem;">(ๅ–”ๆ ขๅฏๅ–”ๆ–ท็ฎๅ–”๏ฟฝ ${item.offset})</span>`;

            html += `
                <div style="display:flex; justify-content:space-between; align-items:center; padding: 5px 14px 5px 28px; background:${seqBg}; border-bottom: 1px solid #f0f0f0; flex-wrap:wrap; gap:4px;">
                    <div>
                        <span style="color:#aaa; margin-right:4px;">้ซ๏ฟฝ</span>
                        <span style="font-family:monospace; color: #555;">${item.number}</span>
                        <span style="margin-left:6px;">${label}</span>
                        <span style="margin-left:8px; font-size:0.85rem;">${companyName}</span>
                    </div>
                    <div style="display:flex; gap:4px;">
                        <a href="https://track.thailandpost.co.th/?trackNumber=${item.number}&lang=th" target="_blank" class="badge badge-neutral" style="background-color:#e3f2fd; color:#0d47a1; border-color:#90caf9; padding:2px 5px; font-size:0.78rem;">้ฆๆ•</a>
                        <button class="badge badge-neutral" style="border:1px solid #ccc; cursor:pointer; padding:2px 5px; font-size:0.78rem;" onclick="navigator.clipboard.writeText('${item.number}').then(() => alert('ๅ–”ๅฆๅฏๅ–”ๆ–ทๅผ—ๅ–”๏ฟฝ็ซต ${item.number} ๅ–ไฝฎๅผ—ๅ–ๅคๆ'))">้ฆๆต</button>
                    </div>
                </div>
            `;
        });

        html += `</div></div>`;
    });

    // Dummy forEach needed to close the old loop - replaced above
    // Close dummy
    const closeLoop = () => {};
    closeLoop();

    {
        // Old table footer replacement
        const isMain = false;
        let companyName = '';
        let rowStyle = '';
        let item = {};
        let dbInfo = {};
        let indexCol = '';
        let trackDisplay = '';
        const isMain2 = (item.offset === 0);
            
            if (isMain) {
                // Style for the main number user entered
                companyName = '<span style="color:#999; font-style:italic;">ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”๏ฟฝ (Not Found)</span>';
                if (dbInfoMain) {
                    companyName = `<strong style="color:var(--primary-color);">${dbInfoMain.name}</strong>`;
                    if(batches[dbInfoMain.batchId] && batches[dbInfoMain.batchId].requestDate) {
                         companyName += ` <small style="color:#28a745;">(ๅ–”ๅ•่…‘ๅ–โฌๅ–”ใ ็ซถ: ${new Date(batches[dbInfoMain.batchId].requestDate).toLocaleDateString('th-TH')})</small>`;
                    }
                    rowStyle = 'background-color:#f8fff9;'; 
                    foundCount++;
                }
            } else {
                // Style for the surrounding sequence numbers
                rowStyle = 'background-color:#fafafa; color: #888;'; 
                if (dbInfo) {
                    companyName = `<strong style="color:var(--primary-color);">${dbInfo.name}</strong>`;
                    if(batches[dbInfo.batchId] && batches[dbInfo.batchId].requestDate) {
                         companyName += ` <small style="color:#28a745;">(ๅ–”ๅ•่…‘ๅ–โฌๅ–”ใ ็ซถ: ${new Date(batches[dbInfo.batchId].requestDate).toLocaleDateString('th-TH')})</small>`;
                    }
                    rowStyle = 'background-color:#f0fbf2; color: #444;'; 
                }
            }

            // UI Label and Formatting Difference
            let label = '';
            let trackDisplay = `<span style="font-family:monospace; font-weight:bold; font-size: 1.1em;">${item.number}</span>`;
            let indexCol = '';
            
            if (!isMain) {
                if (item.offset < 0) label = `(ๅ–”ไฝฎ็ฎๅ–”๏ฟฝ็ฌๅ–”๏ฟฝ็ฌๅ–ๅค่ฆ ${Math.abs(item.offset)})`;
                if (item.offset > 0) label = `(ๅ–”ๆ ขๅฏๅ–”ๆ–ท็ฎๅ–”๏ฟฝ ${item.offset})`;
                trackDisplay = `<span style="font-family:monospace; margin-left:15px;">้ซ๏ฟฝ ${item.number} ${label}</span>`;
            } else {
                indexCol = (idx + 1).toString();
            }

            const actionsHtml = `
                <div class="status-actions" style="margin-top:${isMain ? '5px' : '2px'}; margin-bottom: 5px; ${isMain ? '' : 'font-size: 0.8em; opacity: 0.8;'}">
                    <a href="https://track.thailandpost.co.th/?trackNumber=${item.number}&lang=th" target="_blank" class="badge badge-neutral" style="background-color:#e3f2fd; color:#0d47a1; border-color:#90caf9; ${isMain ? '' : 'padding: 2px 4px;'}" title="Official Deep Link">้ฆๆ• Official</a>
                    <button class="badge badge-neutral" style="border:1px solid #999; cursor:pointer; ${isMain ? '' : 'padding: 2px 4px;'}" onclick="navigator.clipboard.writeText('${item.number}').then(() => alert('ๅ–”ๅฆๅฏๅ–”ๆ–ทๅผ—ๅ–”๏ฟฝ็ซต ${item.number} ๅ–ไฝฎๅผ—ๅ–ๅคๆ'))" title="Copy ID">้ฆๆต Copy</button>
                </div>
            `;

            html += `
                <tr style="${rowStyle}">
                    <td style="text-align:center; vertical-align: top; padding-top: 15px; border-top: ${isMain ? '1px solid #ddd' : 'none'};">${indexCol}</td>
                    <td style="vertical-align: top; padding-top: ${isMain ? '15px' : '5px'}; padding-bottom: ${isMain ? '5px' : '5px'}; border-top: ${isMain ? '1px solid #ddd' : 'none'};">
                        ${trackDisplay}
                        <br>
                        <div style="${isMain ? '' : 'margin-left: 15px;'}">${actionsHtml}</div>
                    </td>
                    <td style="vertical-align: top; padding-top: ${isMain ? '15px' : '5px'}; border-top: ${isMain ? '1px solid #ddd' : 'none'};">${companyName}</td>
                </tr>
            `;
        });
    });

    html += `</tbody></table>`;
    
    resultEl.innerHTML = html;
    resultEl.classList.remove('hidden');
    
    statusEl.innerHTML = `ๅ–”ๆ•็ฆๅ–”ะพ็ฌ€ๅ–”็ง็ฌๅ–”๏ฟฝๅฏๅ–”ๅ็ซตๅ–”็ผ–็ฌ–ๅ–”ๆ•็ฆๅ–”ๅ็ซตๅ–”็ผ–็ฌ <strong style="color:green;">${foundCount}</strong> ๅ–”ๅ —่ฆๅ–”ไฝฎ็ฌๅ–”็ผ–็ฎๅ–”ๅๆ–งๅ–”โด็ฌ– ${trackingArray.length} ๅ–”๏ฝ€่ฆๅ–”โ‘ง็ซตๅ–”ไพง็ฆ`;
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

            if (uploadIcon) uploadIcon.innerText = "้ฆๆจ / ้ฆๆ‘ฒ";
            if (uploadTitle) uploadTitle.innerText = "ๅ–ไฝฎ็ฌ—ๅ–”็ญ็ฎ‘ๅ–”็ง้ๅ–ๅ —่…‘ๅ–โฌๅ–”ใ ้ๅ–”๏ฟฝ็ซตๅ–ๅฆ็ฌฉๅ–”ใ ็ฎค Excel ๅ–”๏ฟฝ็ฆๅ–”็ฒช่…‘ ๅ–”๏ฝ€่…นๅ–”ๆถ็ฌญๅ–”ไพง็ฌง";
            if (uploadDesc) uploadDesc.innerText = "ๅ–”๏ฝ€่…‘ๅ–”ๅ็ฆๅ–”็ผ–็ฌ .xlsx, .xls ๅ–ไฝฎๅผ—ๅ–”๏ฟฝ ๅ–”๏ฝ€่…นๅ–”ๆถ็ฌญๅ–”ไพง็ฌง (OCR)";
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
                        <span style="font-size:1rem;">้ฆๆ‘ ๅ–”๏ฝ€่ตดๅ–”ๆฐ็ฌๅ–”ๆฌ่ตๅ–โฌๅ–”ๅ•็ฎๅ–”ไพง็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฌงๅ–”็ผ–้ๅ–”ๆ–ท็ถ<br><small style="font-weight:normal; font-size:0.8rem;">(Import Data Entry)</small></span>
                    </div>
                `;
                const main = document.querySelector('main');
                if (main && document.body) document.body.insertBefore(userHeader, main);
            }

            const saveBtn = document.querySelector('button[onclick="saveImportedBatch()"]');
            if (saveBtn) {
                saveBtn.innerHTML = `้ฆๆ‘ ๅ–”๏ฟฝ็ฆๅ–”่็ฌกๅ–ไฝฎๅผ—ๅ–”็ญ้ๅ–ๅ —็ซพๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ— (Submit Report)`;
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
// SECTION: EXCEPTION LOG (ๅ–”ๆ•็ซตๅ–”๏ฟฝๅผ—ๅ–ๅ —็ฌ)
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
        container.innerHTML = '<p style="text-align:center; color:#999; padding:20px;">ๅ–”โ‘งๅฏๅ–”ๅ็ฎๅ–”โด็ฎๅ–”โดๅ…ๅ–”ๅ•็ฎๅ–”๏ฟฝๆตฎๅ–”็ๅผ—ๅ–”ๆถ็ฆๅ–”็ญๆๅ–”็ผ–็ฌ—ๅ–”่็ซตๅ–”ไพง็ฆๅ–”ๆ•็ซตๅ–”๏ฟฝๅผ—ๅ–ๅ —็ฌ</p>';
        return;
    }

    // Sort by newest first
    exceptions.sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));

    let html = `
        <div id="exception-export-target" style="background:white; padding:15px; border-radius:8px;">
            <div style="margin-bottom:10px; border-bottom:2px solid #333; padding-bottom:10px;">
                <strong style="font-size:1.1rem;">à¸£à¸²à¸¢à¸à¸²à¸à¸à¸¶à¹à¸à¸à¸²à¸à¸à¸µà¹à¹à¸¡à¹à¸¡à¸µà¸ªà¸à¸²à¸à¸°à¸£à¸±à¸à¸à¸²à¸</strong>
            </div>
            <table style="width:100%; font-size:0.9rem; border-collapse: collapse;">
                <thead>
                    <tr style="background:#f1f1f1;">
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:center;">ๅ–”ใ ่ตๅ–”ๆ–ทๅฏๅ–”๏ฟฝ</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:left;">ๅ–โฌๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”๏ฟฝ</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:left;">ๅ–”ๅจป้ๅ–ๅ —่…‘ๅ–”ๆฐ็ฆๅ–”่ไฟฏๅ–”็ผ–็ฌ/ๅ–”๏ฟฝๅฏๅ–”ๅ็ซตๅ–”็ผ–็ฌ– (ๅ–”ๆ ข็ฎๅ–”ไพงๆตฎๅ–”๏ฟฝ)</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:left;">ๅ–โฌๅ–”๏ฟฝ็ฌ—ๅ–”่็ฌขๅ–”๏ฟฝ / ๅ–”๏ฟฝ็ฌๅ–”ไพง็ฌๅ–”๏ฟฝ</th>
                        <th style="padding:8px; border-bottom:1px solid #ccc; text-align:center;" data-html2canvas-ignore>ๅ–”ๅ —ๅฏๅ–”ๆ–ท็ซตๅ–”ไพง็ฆ</th>
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
                    <button class="btn btn-danger" style="padding:4px 8px; font-size:0.8rem;" onclick="deleteException('${item.id}')">ๅ–”ใ ็ฌ</button>
                </td>
            </tr>
        `;
    });

    html += `
                </tbody>
            </table>
        </div>
        <div style="margin-top:10px; text-align:right;">
             <button class="btn btn-neutral" onclick="clearAllExceptions()">้ฆๆฃ้””๏ฟฝ ๅ–”ใ ็ฎๅ–”ไพง็ซพๅ–”ๆถ็ฆๅ–”็ญๆๅ–”็ผ–็ฌ—ๅ–”่็ฌๅ–”็ผ–็ฎๅ–”ๅๆ–งๅ–”โด็ฌ–</button>
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
        alert('ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพง็ซตๅ–”๏ฝ€่…‘ๅ–”ไฝฎ็ฎ‘ๅ–”ใ ็ซถๅ–”็งๅฏๅ–”๏ฟฝ็ฌ–ๅ–”่็ฎ–ๅ–”๏ฟฝ็ฎๅ–”ๅฆ็ฆๅ–”๏ฟฝ 13 ๅ–”๏ฟฝๅผ—ๅ–”็ผ–็ซต');
        trackInput.focus();
        return;
    }
    
    if (!reason) {
        alert('ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพง็ฆๅ–”็ญ็ฌๅ–”่็ฎ‘ๅ–”๏ฟฝ็ฌ—ๅ–”่็ฌขๅ–”ใ ็ฌๅ–”ๆ็ฎๅ–”ๆ•็ซตๅ–”๏ฟฝๅผ—ๅ–ๅ —็ฌ');
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
    if(confirm('ๅ–”โ‘ง่…‘ๅ–”ๆ–ทๅผ—ๅ–”ๆฐ็ฆๅ–”ไพงๆถชๅ–”ไฝฎ่ฆๅ–”๏ฝ€็ฌๅ–”ๆ็ฎๅ–ๅ็ฌๅ–ๅ —ๆ–งๅ–”๏ฝ€้ๅ–”๏ฟฝ็ฎๅ–”โด็ฎ?')) {
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
        alert('ๅ–ๅฆๆตฎๅ–ๅ —็ฌงๅ–”ๆฐ็ซถๅ–ๅค่…‘ๅ–”โด่…นๅ–”ใ ็ฌๅ–”ๆ็ฎๅ–”ๅ —่ตดๅ–”๏ฟฝ็ฆๅ–ๅค่ฆๅ–”ๅ็ฆๅ–”็็ฌกๅ–”็ฉ่ฆๅ–”๏ฟฝ ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพง็ฎ‘ๅ–”็งๅคๅ–ๅ —ๆตฎๅ–”ๆถ็ฆๅ–”็ญๆๅ–”็ผ–็ฌ—ๅ–”่็ซตๅ–ๅ —่…‘ๅ–”ๆฌ็ซธๅ–”๏ฝ€ๅฏๅ–”๏ฟฝ');
        return;
    }
    
    if(typeof html2canvas === 'undefined') {
        alert('ๅ–”๏ฝ€่ตดๅ–”ๆฐ็ฌๅ–”ไฝฎ่ตๅ–”ใ ๅฏๅ–”ๅ็ฎ“ๅ–”๏ฟฝๅผ—ๅ–”ๆ–ท็ฎ‘ๅ–”ๅฆ็ฆๅ–”็ฒช็ฎๅ–”๏ฟฝ็ซพๅ–”โด้ๅ–”๏ฟฝ้ๅ–”๏ฝ€็ฎๅ–”ไพง็ซพๅ–”็ฉ่ฆๅ–”๏ฟฝ ๅ–”๏ฟฝ็ฆๅ–”็ฒช่…‘ๅ–ๅ•ๆ–งๅ–”ใ ็ฌ–ๅ–ๅฆๆตฎๅ–ๅ —้ๅ–”่ตค็ฎ‘ๅ–”๏ฝ€็ฎๅ–”๏ฟฝ ๅ–”ไฝฎ็ฆๅ–”่็ฌ“ๅ–”ไพงๅผ—ๅ–”๏ฟฝ็ซพๅ–ๅๆ–งๅ–”โด็ฎ (ๅ–”ๆ•็ฎๅ–”๏ฟฝ็ซพๅ–”ๆ•็ฎๅ–”๏ฟฝ็ฎ‘ๅ–”ๆฌ็ฎๅ–”๏ฟฝ)');
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
        
        // Trigger Download as JPEG for smaller file size (LINE-friendly)
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
        alert('ๅ–โฌๅ–”ไฝฎๅคๅ–”ๆ–ท็ซถๅ–ๅค่…‘ๅ–”ๆบนๅคๅ–”ๆ–ท็ฌงๅ–”ใ ่ฆๅ–”ๆ–ท็ฎ–ๅ–”ๆฌ็ซตๅ–”ไพง็ฆๅ–”๏ฟฝ็ฆๅ–ๅค่ฆๅ–”ๅ็ฌญๅ–”ไพง็ฌง: ' + err.message);
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



