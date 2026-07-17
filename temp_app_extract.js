async function analyzeImportData(rows, file) {
    currentTrackingNumbers = [];
    currentAnalyzedData = [];
    const trackRegex = /[A-Z]{2}\d{9}[A-Z]{2}/i;
    
    const batchType = document.getElementById('import-batch-type').value;
    const packageRate = document.getElementById('import-package-rate').value;
    
    let totalCalculatedPrice = 0;
    let errors = 0;

    for(let i=0; i<rows.length; i++) {
        if(!rows[i]) continue;
        let trackNum = null;
        let weight = 0;
        let filePrice = 0;
        
        for(let j=0; j<rows[i].length; j++) {
            const cell = rows[i][j];
            if (cell && typeof cell === 'string') {
                const match = cell.match(trackRegex);
                if (match && !trackNum) {
                    trackNum = match[0].toUpperCase();
                }
            }
            if (typeof cell === 'number' || (typeof cell === 'string' && !isNaN(parseFloat(cell)))) {
                const num = parseFloat(cell);
                if (num > 100 && num < 50000 && weight === 0) weight = num;
                else if (num > 0 && num <= 5000 && filePrice === 0) filePrice = num;
            }
        }
        
        if (trackNum) {
            currentTrackingNumbers.push(trackNum);
            
            let calcPrice = filePrice; 
            let statusHtml = '<span style="color:#4caf50;"><i class="fas fa-check-circle"></i> ปกติ</span>';
            let isError = false;

            if (batchType === 'CustomerA') {
                calcPrice = calculatePrice(packageRate, weight);
                if (calcPrice !== filePrice) {
                    statusHtml = '<span style="color:#f44336;"><i class="fas fa-exclamation-circle"></i> ราคาไม่ตรงเรท</span>';
                    isError = true;
                }
            }

            if (weight === 0) {
                statusHtml = '<span style="color:#ff9800;"><i class="fas fa-exclamation-triangle"></i> ไม่พบน้ำหนัก</span>';
                isError = true;
            }
            
            if (isError) errors++;
            totalCalculatedPrice += calcPrice;

            currentAnalyzedData.push({
                track: trackNum,
                weight: weight,
                filePrice: filePrice,
                calcPrice: calcPrice,
                statusHtml: statusHtml
            });
        }
    }
    
    const unique = [];
    currentAnalyzedData = currentAnalyzedData.filter(item => {
        if(unique.includes(item.track)) return false;
        unique.push(item.track);
        return true;
    });
    currentTrackingNumbers = unique;
    currentTotalItems = currentTrackingNumbers.length;

    if (currentTotalItems === 0) {
        alert('ไม่พบหมายเลขพัสดุในไฟล์นี้');
        clearImportData();
        return;
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

    renderAnalysisTable(totalCalculatedPrice, errors);
    uploadToFirebaseAuto(file);
}

function renderAnalysisTable(totalPrice, errors) {
    const tbody = document.getElementById('analysis-tbody');
    tbody.innerHTML = '';
    
    currentAnalyzedData.forEach((item, index) => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid #eee';
        if(item.statusHtml.includes('exclamation')) tr.style.backgroundColor = '#fff8e1';
        
        tr.innerHTML = `
            <td style="padding:10px; text-align:center;">${index + 1}</td>
            <td style="padding:10px; font-weight:bold; color:#1565c0;">${item.track}</td>
            <td style="padding:10px; text-align:center;">${item.weight}</td>
            <td style="padding:10px; text-align:center;">${item.filePrice}</td>
            <td style="padding:10px; text-align:center; font-weight:bold;">${item.calcPrice}</td>
            <td style="padding:10px; text-align:center;">${item.statusHtml}</td>
        `;
        tbody.appendChild(tr);
    });

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
            <div style="background:${errors > 0 ? '#ffebee' : '#f5f5f5'}; padding:15px; border-radius:8px; flex:1; text-align:center;">
                <div style="font-size:0.9rem; color:${errors > 0 ? '#c62828' : '#757575' };">ข้อควรระวัง</div>
                <div style="font-size:1.5rem; font-weight:bold; color:${errors > 0 ? '#b71c1c' : '#616161' };">${errors} <span style="font-size:1rem;">รายการ</span></div>
            </div>
        </div>
    `;

    document.getElementById('import-preview').classList.remove('hidden');
}

async function uploadToFirebaseAuto(file) {
    if (!window.db || !window.storage) {
        document.getElementById('upload-status').innerText = 'ไม่สามารถอัปโหลดได้: ขาดการเชื่อมต่อ Firebase';
        return;
    }
