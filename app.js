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
            for(let i=0; i<rows.length; i++) {
                if(rows[i] && rows[i].length > 0 && typeof rows[i][0] === "string") {
                    const cellVal = rows[i][0].trim().toUpperCase();
                    if (/^[A-Z]{2}\d{9}[A-Z]{2}$/.test(cellVal)) {
                        currentTrackingNumbers.push(cellVal);
                    }
                }
            }
            
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
