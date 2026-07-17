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
            let statusHtml = '<span style="color:#4caf50;"><i class="fas fa-check-circle"></i> ปกติ</span>';
            let isError = false;
            let displayWeight = originalWeight;
            
            // Reverse Weight Logic (อิงราคาเป็นหลัก)
            if (batchType === 'CustomerA') {
                const expectedWeight = getExpectedWeightFromPriceA(packageRate, filePrice);
                if (expectedWeight !== '-') {
                    displayWeight = expectedWeight;
                    
                    // Extract just numbers for comparison
                    const origNum = parseFloat(originalWeight);
                    const expectNum = parseFloat(expectedWeight);
                    
                    if (!isNaN(origNum) && origNum !== expectNum) {
                        statusHtml = '<span style="color:#ff9800;"><i class="fas fa-exclamation-triangle"></i> น้ำหนักไม่สัมพันธ์กับราคา (แก้เป็น ' + expectedWeight + ')</span>';
                        isError = true;
                    }
                }
            }

            if (isError) errors++;
            totalCalculatedPrice += filePrice;

            uniqueMap.set(trackNum, {
                track: trackNum,
                weight: displayWeight,
                originalWeight: originalWeight,
                filePrice: filePrice,
                calcPrice: filePrice, // We use filePrice as main
                statusHtml: statusHtml,
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

    // GAP DETECTION LOGIC
    const parse = (str) => {
        const m = str.match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
        return m ? { full: str, prefix: m[1], body: parseInt(m[2]), check: m[3], suffix: m[4] } : null;
    };
    
    const sortedForGaps = [...currentAnalyzedData].sort((a, b) => a.track.localeCompare(b.track));
    const missingItems = [];
    
    let prevGap = parse(sortedForGaps[0].track);
    for (let i = 1; i < sortedForGaps.length; i++) {
        const currGap = parse(sortedForGaps[i].track);
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

    // VIRTUAL MAPPING LOGIC (Group by price)
    currentAnalyzedData.sort((a, b) => {
        if (a.filePrice !== b.filePrice) return a.filePrice - b.filePrice;
        return a.track.localeCompare(b.track);
    });

    const optimizedRanges = [];
    let start = parse(currentAnalyzedData[0].track);
    let prev = start;
    let currentList = [currentAnalyzedData[0]];

    for (let i = 1; i < currentAnalyzedData.length; i++) {
        const currItem = currentAnalyzedData[i];
        const curr = parse(currItem.track);
        if (!curr) continue;
        
        const prevItem = currentList[currentList.length - 1];

        const isContinuous = (
            curr.prefix === prev.prefix &&
            curr.suffix === prev.suffix &&
            curr.body === prev.body + 1 &&
            currItem.filePrice === prevItem.filePrice
        );

        if (isContinuous) {
            currentList.push(currItem);
            prev = curr;
        } else {
            optimizedRanges.push({
                start: start.full,
                end: prev.full,
                count: currentList.length,
                price: currentList[0].filePrice,
                statusHtml: currentList.find(x => x.isError) ? '<span style="color:#ff9800">มีแจ้งเตือนน้ำหนัก</span>' : '<span style="color:#4caf50">ปกติ</span>',
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
            price: currentList[0].filePrice,
            statusHtml: currentList.find(x => x.isError) ? '<span style="color:#ff9800">มีแจ้งเตือนน้ำหนัก</span>' : '<span style="color:#4caf50">ปกติ</span>',
            errors: currentList.filter(x => x.isError).length
        });
    }

    document.getElementById('upload-status').innerText = 'กำลังตรวจสอบข้อมูลซ้ำ...';
    const isDuplicate = await checkDuplicateInFirebase(currentTrackingNumbers);
    if (isDuplicate) {
        const confirmUpload = confirm('⚠️ ตรวจพบหมายเลขพัสดุในไฟล์นี้ เคยถูกนำเข้าสู่ระบบแล้ว\nคุณต้องการนำเข้าซ้ำหรือไม่?');
        if (!confirmUpload) {
            clearImportData();
            return;
        }
    }

    renderAnalysisTable(totalCalculatedPrice, errors, optimizedRanges, missingItems);
    uploadToFirebaseAuto(file);
}

function renderAnalysisTable(totalPrice, errors, optimizedRanges, missingItems) {
    const tbody = document.getElementById('analysis-tbody');
    tbody.innerHTML = '';
    
    optimizedRanges.forEach((range, index) => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid #eee';
        if(range.errors > 0) tr.style.backgroundColor = '#fff8e1';
        
        const trackDisplay = range.count > 1 ? `${range.start} - ${range.end}` : range.start;
        
        tr.innerHTML = `
            <td style="padding:10px; text-align:center;">${index + 1}</td>
            <td style="padding:10px; font-weight:bold; color:#1565c0;">${trackDisplay}</td>
            <td style="padding:10px; text-align:center;">${range.count} ชิ้น</td>
            <td style="padding:10px; text-align:center;">${range.price}</td>
            <td style="padding:10px; text-align:center; font-weight:bold;">${(range.price * range.count).toLocaleString()}</td>
            <td style="padding:10px; text-align:center;">${range.statusHtml}</td>
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
        <div style="font-size:0.8rem; color:red; text-align:center; margin-top:10px; margin-bottom:5px;">
            *รายการถูกจัดเรียงใหม่แบบมัดรวมชุด (Virtual Mapping) เรียงตามราคาน้อยไปมาก
        </div>
    `;

    document.getElementById('import-preview').classList.remove('hidden');
}
