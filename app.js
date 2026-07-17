let currentFile = null;
let currentTotalItems = 0;
let currentTrackingNumbers = [];
let currentAnalyzedData = [];
let batchesCache = [];

function togglePackageSelector() {
    const type = document.getElementById('import-batch-type').value;
    const selector = document.getElementById('import-package-selector');
    if (type === 'CustomerA') {
        selector.style.display = 'block';
    } else {
        selector.style.display = 'none';
    }
}

function calculatePrice(packageType, weightGrams) {
    if (!weightGrams || weightGrams <= 0) return 0;
    
    const match = packageType.match(/^A(\d+)$/i);
    if (!match) return 0;
    const pkgNum = parseInt(match[1]);
    
    const bases = {
        1:17, 2:18, 3:19, 4:20, 5:21, 6:22,
        7:23, 8:24, 9:25, 10:26, 11:28, 12:30
    };
    
    const basePrice = bases[pkgNum];
    if (!basePrice) return 0;
    
    const kg = Math.ceil(weightGrams / 1000);
    if (kg <= 1) return basePrice;
    if (kg <= 10) return basePrice + ((kg - 1) * 10);
    if (kg <= 20) return basePrice + (9 * 10) + ((kg - 10) * 5);
    return basePrice + (9 * 10) + (10 * 5) + ((kg - 20) * 15);
}

document.getElementById('import-file').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) return;

    currentFile = file;
    document.getElementById('upload-status').innerText = `กำลังวิเคราะห์ไฟล์ ${file.name}...`;
    
    const batchNameInput = document.getElementById('import-batch-name');
    if (!batchNameInput.value) {
        const today = new Date();
        batchNameInput.value = `รายการส่งวันที่ ${today.getDate()}/${today.getMonth()+1}/${today.getFullYear() + 543}`;
    }

    const reader = new FileReader();
    reader.onload = async function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet, {header: 1});
            
            analyzeImportData(rows, file);
        } catch(err) {
            alert('ไม่สามารถอ่านไฟล์ Excel ได้: ' + err.message);
            clearImportData();
        }
    };
    reader.readAsArrayBuffer(file);
});

async function checkDuplicateInFirebase(trackingNumbers) {
    if (!window.db || trackingNumbers.length === 0) return false;
    try {
        const snapshot = await window.db.collection('batches')
            .where('trackingNumbers', 'array-contains', trackingNumbers[0])
            .limit(1).get();
        return !snapshot.empty;
    } catch(e) {
        return false;
    }
}


function getExpectedWeightFromPriceA(packageType, price) {
    const match = packageType.match(/^A(\d+)$/i);
    if (!match) return '-';
    const pkgNum = parseInt(match[1]);
    const bases = {
        1:17, 2:18, 3:19, 4:20, 5:21, 6:22,
        7:23, 8:24, 9:25, 10:26, 11:28, 12:30
    };
    const basePrice = bases[pkgNum];
    if (!basePrice || price < basePrice) return '-';

    if (price === basePrice) return '1 กก';
    
    // Reverse calculation
    let rem = price - basePrice;
    
    // Up to 10kg (+10 baht/kg)
    if (rem <= 90) {
        let extraKg = Math.ceil(rem / 10);
        return (1 + extraKg) + ' กก';
    }
    rem -= 90;
    
    // Up to 20kg (+5 baht/kg)
    if (rem <= 50) {
        let extraKg = Math.ceil(rem / 5);
        return (10 + extraKg) + ' กก';
    }
    rem -= 50;
    
    // Over 20kg (+15 baht/kg)
    let extraKg = Math.ceil(rem / 15);
    return (20 + extraKg) + ' กก';
}

async function analyzeImportData(rows, file) {
    currentTrackingNumbers = [];
    currentAnalyzedData = [];
    const trackRegex = /[A-Z]{2}\d{9}[A-Z]{2}/i;
    
    const batchType = document.getElementById('import-batch-type').value;
    const packageRate = document.getElementById('import-package-rate').value;
    
    let totalCalculatedPrice = 0;
    let errors = 0;
    
    const uniqueMap = new Map();
    const discrepancyList = [];

    for(let i=0; i<rows.length; i++) {
        if(!rows[i]) continue;
        let trackNum = null;
        let weight = 0;
        let originalWeight = '-';
        let filePrice = 0;
        
        for(let j=0; j<rows[i].length; j++) {
            const cell = rows[i][j];
            if (cell && typeof cell === 'string') {
                const match = cell.match(trackRegex);
                if (match && !trackNum) {
                    trackNum = match[0].toUpperCase();
                }
                if (cell.includes('กก') || cell.includes('kg')) {
                    originalWeight = cell.trim();
                }
            }
            if (typeof cell === 'number' || (typeof cell === 'string' && !isNaN(parseFloat(cell)))) {
                const num = parseFloat(cell);
                if (num > 100 && num < 50000 && weight === 0) {
                    weight = num;
                    if (originalWeight === '-') originalWeight = num + ' กก';
                }
                else if (num > 0 && num <= 5000 && filePrice === 0) filePrice = num;
            }
        }
        
        if (trackNum && filePrice > 0) {
            let isError = false;
            let displayWeight = originalWeight;
            
            // Weight Check (before reallocation)
            if (batchType === 'CustomerA') {
                const expectedWeight = getExpectedWeightFromPriceA(packageRate, filePrice);
                if (expectedWeight !== '-') {
                    displayWeight = expectedWeight;
                    const origNum = parseFloat(originalWeight);
                    const expectNum = parseFloat(expectedWeight);
                    
                    if (!isNaN(origNum) && origNum !== expectNum) {
                        isError = true;
                        discrepancyList.push({ orig: originalWeight, exp: expectedWeight, price: filePrice });
                    }
                }
            }

            if (isError) errors++;
            totalCalculatedPrice += filePrice;

            uniqueMap.set(trackNum, {
                track: trackNum,
                originalWeight: originalWeight,
                filePrice: filePrice,
                isError: isError
            });
        }
    }
    
    currentAnalyzedData = Array.from(uniqueMap.values());
    currentTrackingNumbers = currentAnalyzedData.map(item => item.track);
    currentTotalItems = currentTrackingNumbers.length;

    if (currentTotalItems === 0) {
        alert('ไม่พบหมายเลขพัสดุหรือราคาในไฟล์นี้');
        clearImportData();
        return;
    }

    const parse = (str) => {
        const m = str.match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
        return m ? { full: str, prefix: m[1], body: parseInt(m[2]), check: m[3], suffix: m[4] } : null;
    };
    
    // 1. Sort All IDs Ascending
    const sortedItems = [...currentAnalyzedData].sort((a, b) => a.track.localeCompare(b.track));
    
    // 2. GAP DETECTION LOGIC (Using naturally sorted trackings)
    const missingItems = [];
    let prevGap = parse(sortedItems[0].track);
    for (let i = 1; i < sortedItems.length; i++) {
        const currGap = parse(sortedItems[i].track);
        if (!currGap || !prevGap) continue;
        if (currGap.prefix === prevGap.prefix && currGap.suffix === prevGap.suffix) {
            const diff = currGap.body - prevGap.body;
            if (diff > 1) {
                const startMissing = prevGap.body + 1;
                const countMissing = diff - 1;
                const exampleID = `${currGap.prefix}${startMissing.toString().padStart(8, '0')}X${currGap.suffix}`;
                missingItems.push({ count: countMissing, example: exampleID });
            }
        }
        prevGap = currGap;
    }

    // 3. VIRTUAL ZIPPING LOGIC
    // Group counts by Price
    const priceCounts = {};
    currentAnalyzedData.forEach(item => {
        if (!priceCounts[item.filePrice]) priceCounts[item.filePrice] = 0;
        priceCounts[item.filePrice]++;
    });
    
    // Sort Price groups (Lowest first)
    const sortedPrices = Object.keys(priceCounts).map(Number).sort((a, b) => a - b);
    
    const optimizedRanges = [];
    let currentIdx = 0;
    
    sortedPrices.forEach(price => {
        const count = priceCounts[price];
        const assignedItems = sortedItems.slice(currentIdx, currentIdx + count);
        currentIdx += count;
        
        if (assignedItems.length === 0) return;
        
        // Break into contiguous sequences
        let start = parse(assignedItems[0].track);
        let prev = start;
        let currentList = [assignedItems[0]];
        
        for (let i = 1; i < assignedItems.length; i++) {
            const currItem = assignedItems[i];
            const curr = parse(currItem.track);
            if (!curr) continue;
            
            const isContinuous = (
                curr.prefix === prev.prefix &&
                curr.suffix === prev.suffix &&
                curr.body === prev.body + 1
            );

            if (isContinuous) {
                currentList.push(currItem);
                prev = curr;
            } else {
                optimizedRanges.push({
                    start: start.full,
                    end: prev.full,
                    count: currentList.length,
                    price: price,
                    errors: currentList.filter(x => x.isError).length
                });
                start = curr;
                prev = curr;
                currentList = [currItem];
            }
        }
        if (currentList.length > 0) {
            optimizedRanges.push({
                start: start.full,
                end: prev.full,
                count: currentList.length,
                price: price,
                errors: currentList.filter(x => x.isError).length
            });
        }
    });

    document.getElementById('upload-status').innerText = 'กำลังตรวจสอบข้อมูลซ้ำ...';
    const isDuplicate = await checkDuplicateInFirebase(currentTrackingNumbers);
    if (isDuplicate) {
        const confirmUpload = confirm('⚠️ ตรวจพบหมายเลขพัสดุในไฟล์นี้ เคยถูกนำเข้าสู่ระบบแล้ว\nคุณต้องการนำเข้าซ้ำหรือไม่?');
        if (!confirmUpload) {
            clearImportData();
            return;
        }
    }

    renderAnalysisTable(totalCalculatedPrice, errors, optimizedRanges, missingItems, discrepancyList);
    uploadToFirebaseAuto(file);
}

function renderAnalysisTable(totalPrice, errors, optimizedRanges, missingItems, discrepancyList) {
    const tbody = document.getElementById('analysis-tbody');
    tbody.innerHTML = '';
    
    optimizedRanges.forEach((range, index) => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid #eee';
        
        const trackDisplay = range.count > 1 ? `${range.start} - ${range.end}` : range.start;
        const totalLinePrice = range.price * range.count;
        
        tr.innerHTML = `
            <td style="padding:10px; text-align:center;">${index + 1}</td>
            <td style="padding:10px; font-weight:bold; color:#1565c0;">${trackDisplay}</td>
            <td style="padding:10px; text-align:center;">${range.count} ชิ้น</td>
            <td style="padding:10px; text-align:center;">${range.price}</td>
            <td style="padding:10px; text-align:center; font-weight:bold;">${totalLinePrice.toLocaleString()}</td>
            <td style="padding:10px; text-align:center; font-size:0.9rem; color:#888;">จับคู่ใหม่</td>
        `;
        tbody.appendChild(tr);
    });

    let gapHtml = '';
    if (missingItems && missingItems.length > 0) {
        const totalMissing = missingItems.reduce((sum, item) => sum + item.count, 0);
        gapHtml = `
        <div style="background:#ffebee; padding:15px; border-radius:8px; border: 1px solid #f44336; margin-top:15px;">
            <h4 style="color:#c62828; margin:0 0 10px 0;"><i class="fas fa-exclamation-circle"></i> แจ้งเตือน: พบช่องโหว่ (เลขที่หายไป)</h4>
            <p style="margin:0; color:#b71c1c;">ระบบพบว่าข้อมูลไม่ต่อเนื่อง <strong>หายไปรวม ${totalMissing} รายการ</strong> (เช่น แถวๆ เลข ${missingItems[0].example})</p>
        </div>`;
    }

    let diffHtml = '';
    if (discrepancyList && discrepancyList.length > 0) {
        diffHtml = `
        <div style="background:#fff8e1; padding:15px; border-radius:8px; border: 1px solid #ff9800; margin-top:15px;">
            <h4 style="color:#ef6c00; margin:0 0 10px 0;"><i class="fas fa-exclamation-triangle"></i> แจ้งเตือนน้ำหนักไม่สัมพันธ์กับราคา</h4>
            <p style="margin:0; color:#e65100; font-size:0.9rem;">
                พบข้อมูลระบุในไฟล์ <strong>ไม่สัมพันธ์กับราคาตามตารางค่าบริการ</strong> จำนวน <strong>${discrepancyList.length} รายการ</strong><br>
                <small>เช่น ในไฟล์ระบุ ${discrepancyList[0].orig} แต่ราคา ${discrepancyList[0].price} บาท (ระบบอิงราคา ${discrepancyList[0].price} เป็นหลัก)</small>
            </p>
        </div>`;
    }

    document.getElementById('analysis-summary').innerHTML = `
        <div style="display:flex; gap:15px; margin-bottom:15px;">
            <div style="background:#e3f2fd; padding:15px; border-radius:8px; flex:1; text-align:center;">
                <div style="font-size:0.9rem; color:#1565c0;">ยอดรวมทั้งหมด</div>
                <div style="font-size:1.5rem; font-weight:bold; color:#0d47a1;">${currentTotalItems} <span style="font-size:1rem;">ชิ้น</span></div>
            </div>
            <div style="background:#e8f5e9; padding:15px; border-radius:8px; flex:1; text-align:center;">
                <div style="font-size:0.9rem; color:#2e7d32;">ราคารวม (คำนวณ)</div>
                <div style="font-size:1.5rem; font-weight:bold; color:#1b5e20;">฿${totalPrice.toLocaleString()}</div>
            </div>
            <div style="background:${errors > 0 ? '#fff3e0' : '#f5f5f5'}; padding:15px; border-radius:8px; flex:1; text-align:center;">
                <div style="font-size:0.9rem; color:${errors > 0 ? '#e65100' : '#757575' };">แจ้งเตือนน้ำหนักไม่ตรง</div>
                <div style="font-size:1.5rem; font-weight:bold; color:${errors > 0 ? '#ef6c00' : '#616161' };">${errors} <span style="font-size:1rem;">รายการ</span></div>
            </div>
        </div>
        ${gapHtml}
        ${diffHtml}
        <div style="font-size:0.8rem; color:red; text-align:center; margin-top:10px; margin-bottom:5px;">
            *รายการถูกจัดเรียงใหม่แบบมัดรวมชุด (Virtual Zipping) เรียงตามราคาจากน้อยไปมาก
        </div>
    `;

    document.getElementById('import-preview').classList.remove('hidden');
}




async function uploadToFirebaseAuto(file) {
    if (!window.db || !window.storage) {
        document.getElementById('upload-status').innerText = 'ไม่สามารถอัปโหลดได้: ขาดการเชื่อมต่อ Firebase';
        return;
    }

    const batchName = document.getElementById('import-batch-name').value.trim();
    const batchType = document.getElementById('import-batch-type').value;
    document.getElementById('upload-status').innerHTML = '<i class="fas fa-spinner fa-spin"></i> กำลังอัปโหลดข้อมูลขึ้น Cloud...';

    try {
        const fileExt = file.name.split('.').pop();
        const fileName = `excel_imports/${new Date().getTime()}_${Math.random().toString(36).substring(7)}.${fileExt}`;
        const storageRef = window.storage.ref().child(fileName);
        
        await storageRef.put(file);
        const downloadURL = await storageRef.getDownloadURL();
        
        const docData = {
            batchName: batchName,
            type: batchType,
            timestamp: new Date().toISOString(),
            totalItems: currentTotalItems,
            fileName: file.name,
            fileURL: downloadURL,
            storagePath: fileName,
            trackingNumbers: currentTrackingNumbers,
            analyzedData: currentAnalyzedData
        };
        
        await window.db.collection('batches').add(docData);
        
        document.getElementById('upload-status').innerHTML = '<span style="color:#4caf50;"><i class="fas fa-check-circle"></i> อัปโหลดและวิเคราะห์เสร็จสมบูรณ์!</span>';
        if (document.getElementById('tab-history').classList.contains('active')) {
            loadStaffHistory();
        }
    } catch(err) {
        console.error(err);
        document.getElementById('upload-status').innerHTML = `<span style="color:#f44336;"><i class="fas fa-times-circle"></i> Error: ${err.message}</span>`;
    }
}

function clearImportData() {
    currentFile = null;
    currentTotalItems = 0;
    currentTrackingNumbers = [];
    currentAnalyzedData = [];
    document.getElementById('import-file').value = '';
    document.getElementById('upload-status').innerText = '';
    document.getElementById('import-preview').classList.add('hidden');
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => {
        el.style.color = '#666';
        el.style.borderBottom = 'none';
    });
    
    document.getElementById('tab-' + tabId).classList.add('active');
    
    const activeBtn = document.getElementById('btn-' + tabId);
    if(activeBtn) {
        activeBtn.style.color = '#2196f3';
        activeBtn.style.borderBottom = '3px solid #2196f3';
    }
    
    if (tabId === 'check') {
        document.getElementById('check-input').focus();
        if (batchesCache.length === 0) loadHistorySilent();
    } else if (tabId === 'history') {
        loadStaffHistory();
    }
}

async function loadStaffHistory() {
    const list = document.getElementById('staff-history-list');
    if(!list) return;
    list.innerHTML = '<div style="text-align:center; padding:20px;"><i class="fas fa-spinner fa-spin"></i> กำลังโหลดประวัติ...</div>';
    
    if (!window.db) return;
    try {
        const snapshot = await window.db.collection('batches').orderBy('timestamp', 'desc').limit(20).get();
        list.innerHTML = '';
        if(snapshot.empty) {
            list.innerHTML = '<div style="text-align:center; padding:20px; color:#999;">ไม่มีประวัติการนำเข้าในวันนี้</div>';
            return;
        }
        
        snapshot.forEach(doc => {
            const data = doc.data();
            const dateStr = new Date(data.timestamp).toLocaleString('th-TH');
            
            const card = document.createElement('div');
            card.style.background = '#fff';
            card.style.border = '1px solid #e0e0e0';
            card.style.borderRadius = '8px';
            card.style.padding = '15px';
            card.style.display = 'flex';
            card.style.justifyContent = 'space-between';
            card.style.alignItems = 'center';
            card.style.boxShadow = '0 2px 4px rgba(0,0,0,0.02)';
            card.style.marginBottom = '10px';
            
            card.innerHTML = `
                <div>
                    <div style="font-weight:bold; color:#1e3c72; font-size:1.1rem;">${data.batchName || 'ไม่มีชื่อกลุ่ม'}</div>
                    <div style="font-size:0.9rem; color:#666; margin-top:5px;"><i class="far fa-clock"></i> ${dateStr} | ${data.type || 'N/A'}</div>
                    <div style="font-size:0.9rem; color:#666; margin-top:2px;"><i class="fas fa-file-excel"></i> ${data.fileName || '-'}</div>
                </div>
                <div style="text-align:right;">
                    <div style="font-size:1.5rem; font-weight:bold; color:#2196f3;">${data.totalItems || 0} <span style="font-size:1rem;">ชิ้น</span></div>
                </div>
            `;
            list.appendChild(card);
        });
    } catch(err) {
        list.innerHTML = `<div style="color:red;">Failed to load history: ${err.message}</div>`;
    }
}

async function loadHistorySilent() {
    if (!window.db) return;
    try {
        const snapshot = await window.db.collection('batches').orderBy('timestamp', 'desc').limit(50).get();
        batchesCache = [];
        snapshot.forEach(doc => {
            const data = doc.data();
            data.id = doc.id;
            batchesCache.push(data);
        });
    } catch (err) {}
}

function checkBarcode() {
    const input = document.getElementById('check-input').value.trim().toUpperCase();
    if (!input) return;
    
    const resultDiv = document.getElementById('check-result');
    resultDiv.classList.remove('hidden');
    
    if (batchesCache.length === 0) {
        resultDiv.innerHTML = '<div style="padding:15px; color:#f44336;">กำลังโหลดข้อมูล... ลองใหม่อีกครั้ง</div>';
        loadHistorySilent();
        return;
    }

    let foundBatches = [];
    for (const batch of batchesCache) {
        if (batch.trackingNumbers && batch.trackingNumbers.includes(input)) {
            foundBatches.push(batch);
        }
    }
    
    if (foundBatches.length > 0) {
        let html = '<div style="background: #e8f5e9; border: 1px solid #4caf50; padding: 15px; border-radius: 8px;"><h3 style="color: #2e7d32; margin-top:0;"><i class="fas fa-check-circle"></i> พบพัสดุ ' + input + '</h3><ul style="margin-bottom: 0;">';
        foundBatches.forEach(b => {
            const dateStr = new Date(b.timestamp).toLocaleString('th-TH');
            html += `<li><strong>${b.batchName || 'ไม่มีชื่อ'}</strong> (${b.type || 'N/A'}) - ${dateStr}</li>`;
        });
        html += '</ul></div>';
        resultDiv.innerHTML = html;
    } else {
        resultDiv.innerHTML = `<div style="background: #ffebee; border: 1px solid #f44336; padding: 15px; border-radius: 8px;"><h3 style="color: #c62828; margin-top:0;"><i class="fas fa-times-circle"></i> ไม่พบพัสดุ ${input}</h3></div>`;
    }
    
    document.getElementById('check-input').value = '';
    document.getElementById('check-input').focus();
}

document.getElementById('check-input').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') checkBarcode();
});

document.getElementById('check-input').addEventListener('input', function(e) {
    let val = e.target.value.toUpperCase().replace(/\s/g, '');
    const match = val.match(/^([A-Z]{2})(\d{8})([A-Z]{2})$/);
    if (match && typeof TrackingUtils !== 'undefined') {
        const cd = TrackingUtils.calculateS10CheckDigit(match[2]);
        if (cd !== null) e.target.value = `${match[1]}${match[2]}${cd}${match[3]}`;
    }
});
