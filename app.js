let currentFile = null;
let currentTotalItems = 0;
let currentTrackingNumbers = [];

document.getElementById("import-file").addEventListener("change", function(e) {
    const file = e.target.files[0];
    if (!file) return;

    currentFile = file;
    document.getElementById("upload-status").innerText = `กำลังตรวจสอบไฟล์ ${file.name}...`;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: "array"});
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, {header: 1});
            
            currentTrackingNumbers = [];
            const trackRegex = /[A-Z]{2}\d{9}[A-Z]{2}/ig;
            
            for(let i=0; i<rows.length; i++) {
                if(!rows[i]) continue;
                for(let j=0; j<rows[i].length; j++) {
                    const cell = rows[i][j];
                    if (cell && typeof cell === "string") {
                        const matches = cell.match(trackRegex);
                        if (matches) {
                            matches.forEach(m => currentTrackingNumbers.push(m.toUpperCase()));
                        }
                    }
                }
            }
            
            // Remove duplicates just in case
            currentTrackingNumbers = [...new Set(currentTrackingNumbers)];
            
            currentTotalItems = currentTrackingNumbers.length;
            document.getElementById("upload-status").innerText = `ไฟล์ ${file.name} เตรียมพร้อม (พบ ${currentTotalItems} รายการ)`;
            document.getElementById("import-preview").classList.remove("hidden");
        } catch(err) {
            alert("ไม่สามารถอ่านไฟล์ Excel ได้: " + err.message);
            clearImportData();
        }
    };
    reader.readAsArrayBuffer(file);
});

async function uploadToFirebase() {
    if (!currentFile) {
        alert("กรุณาเลือกไฟล์ Excel ก่อน");
        return;
    }
    
    if (!window.db || !window.storage) {
        alert("Firebase Config ไม่สมบูรณ์");
        return;
    }

    const batchName = document.getElementById("import-batch-name").value.trim() || `Upload_${new Date().getTime()}`;
    const batchType = document.getElementById("import-batch-type").value;
    
    const btn = document.getElementById("import-save-btn");
    btn.disabled = true;
    btn.innerHTML = `<i class="fas fa-spinner fa-spin"></i> กำลังอัปโหลด...`;

    try {
        const fileExt = currentFile.name.split(".").pop();
        const fileName = `excel_imports/${new Date().getTime()}_${Math.random().toString(36).substring(7)}.${fileExt}`;
        const storageRef = window.storage.ref().child(fileName);
        
        await storageRef.put(currentFile);
        const downloadURL = await storageRef.getDownloadURL();
        
        const docData = {
            batchName: batchName,
            type: batchType,
            timestamp: new Date().toISOString(),
            totalItems: currentTotalItems,
            fileName: currentFile.name,
            fileURL: downloadURL,
            storagePath: fileName,
            trackingNumbers: currentTrackingNumbers
        };
        
        await window.db.collection("batches").add(docData);
        
        alert("อัปโหลดขึ้นคลาวด์สำเร็จเรียบร้อย!");
        clearImportData();
    } catch(err) {
        console.error(err);
        alert("Error: " + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = `<i class="fas fa-cloud-upload-alt"></i> ยืนยันการอัปโหลด (Upload to Cloud)`;
    }
}

function clearImportData() {
    currentFile = null;
    currentTotalItems = 0;
    currentTrackingNumbers = [];
    document.getElementById("import-file").value = "";
    document.getElementById("upload-status").innerText = "";
    document.getElementById("import-batch-name").value = "";
    document.getElementById("import-preview").classList.add("hidden");
}

let batchesCache = [];

function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    
    document.getElementById('btn-import').style.color = '#666';
    document.getElementById('btn-import').style.borderBottom = 'none';
    
    document.getElementById('btn-check').style.color = '#666';
    document.getElementById('btn-check').style.borderBottom = 'none';
    
    document.getElementById('tab-' + tabId).classList.add('active');
    
    const activeBtn = document.getElementById('btn-' + tabId);
    activeBtn.style.color = '#2196f3';
    activeBtn.style.borderBottom = '3px solid #2196f3';
    
    if (tabId === 'check') {
        document.getElementById('check-input').focus();
        if (batchesCache.length === 0) loadHistorySilent();
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
    } catch (err) {
        console.error('Failed to load history for checking', err);
    }
}

function checkBarcode() {
    const input = document.getElementById('check-input').value.trim().toUpperCase();
    if (!input) return;
    
    const resultDiv = document.getElementById('check-result');
    resultDiv.classList.remove('hidden');
    
    if (batchesCache.length === 0) {
        resultDiv.innerHTML = '<div style=\"padding:15px; color:#f44336;\">กำลังโหลดข้อมูลจากฐานข้อมูล... ลองใหม่อีกครั้ง</div>';
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
        let html = '<div style=\"background: #e8f5e9; border: 1px solid #4caf50; padding: 15px; border-radius: 8px;\"><h3 style=\"color: #2e7d32; margin-top:0;\"><i class=\"fas fa-check-circle\"></i> พบพัสดุ ' + input + '</h3><p>พบในกลุ่มข้อมูลต่อไปนี้:</p><ul style=\"margin-bottom: 0;\">';
        foundBatches.forEach(b => {
            const dateStr = new Date(b.timestamp).toLocaleString('th-TH');
            html += '<li><strong>' + b.batchName + '</strong> (' + b.type + ') - นำเข้าเมื่อ: ' + dateStr + '</li>';
        });
        html += '</ul></div>';
        resultDiv.innerHTML = html;
    } else {
        resultDiv.innerHTML = '<div style=\"background: #ffebee; border: 1px solid #f44336; padding: 15px; border-radius: 8px;\"><h3 style=\"color: #c62828; margin-top:0;\"><i class=\"fas fa-times-circle\"></i> ไม่พบพัสดุ ' + input + '</h3><p style=\"margin-bottom: 0;\">กรุณาตรวจสอบเลขอีกครั้ง หรือพัสดุนี้อาจยังไม่ได้ถูกนำเข้า</p></div>';
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
    if (match) {
        const prefix = match[1];
        const body = match[2];
        const suffix = match[3];
        if (typeof TrackingUtils !== 'undefined') {
            const cd = TrackingUtils.calculateS10CheckDigit(body);
            if (cd !== null) {
                e.target.value = `${prefix}${body}${cd}${suffix}`;
            }
        }
    }
});

