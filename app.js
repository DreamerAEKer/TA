
let batchesCache = [];

function switchTab(tabId) {
    document.querySelectorAll(".tab-content").forEach(el => el.classList.remove("active"));
    document.querySelectorAll(".tab-btn").forEach(el => el.classList.remove("active"));
    
    document.getElementById("tab-" + tabId).classList.add("active");
    event.currentTarget.classList.add("active");
    
    if (tabId === "history") loadHistory();
    if (tabId === "check") {
        document.getElementById("check-input").focus();
        if (batchesCache.length === 0) loadHistory(true); // silent load
    }
}

async function loadHistory(silent = false) {
    if (!window.db) {
        if(!silent) alert("Firebase not initialized!");
        return;
    }
    
    if (!silent) {
        document.getElementById("history-loading").style.display = "block";
        document.getElementById("history-list").innerHTML = "";
    }

    try {
        const snapshot = await window.db.collection("batches").orderBy("timestamp", "desc").get();
        batchesCache = [];
        
        let html = "";
        snapshot.forEach(doc => {
            const data = doc.data();
            data.id = doc.id;
            batchesCache.push(data);
            
            const dateStr = new Date(data.timestamp).toLocaleString("th-TH");
            html += `
                <div style="border: 1px solid #ddd; padding: 15px; border-radius: 8px; margin-bottom: 10px; background: #fff;">
                    <div style="display: flex; justify-content: space-between;">
                        <div>
                            <h3 style="margin: 0; color: #1565c0;">${data.batchName} <span style="font-size: 0.8rem; background: #e3f2fd; padding: 2px 8px; border-radius: 12px;">${data.type}</span></h3>
                            <p style="margin: 5px 0; color: #666; font-size: 0.9rem;">
                                <i class="far fa-clock"></i> ${dateStr} | 
                                <i class="fas fa-box"></i> ${data.totalItems} รายการ
                            </p>
                        </div>
                        <div style="display: flex; gap: 10px;">
                            ${data.fileURL ? `<a href="${data.fileURL}" target="_blank" class="btn btn-primary" style="text-decoration: none;"><i class="fas fa-download"></i> โหลดไฟล์</a>` : ""}
                        </div>
                    </div>
                </div>
            `;
        });
        
        if (batchesCache.length === 0) {
            html = `<div style="text-align: center; color: #999; padding: 20px;">ยังไม่มีประวัติการนำเข้า</div>`;
        }
        
        document.getElementById("history-list").innerHTML = html;
        if(!silent) document.getElementById("history-loading").style.display = "none";
        
    } catch (err) {
        console.error(err);
        if(!silent) {
            document.getElementById("history-loading").style.display = "none";
            document.getElementById("history-list").innerHTML = `<div style="color: red;">Error: ${err.message}</div>`;
        }
    }
}

function checkBarcode() {
    const input = document.getElementById("check-input").value.trim().toUpperCase();
    if (!input) return;
    
    const resultDiv = document.getElementById("check-result");
    resultDiv.classList.remove("hidden");
    
    // Search in batchesCache
    let foundBatches = [];
    for (const batch of batchesCache) {
        if (batch.trackingNumbers && batch.trackingNumbers.includes(input)) {
            foundBatches.push(batch);
        }
    }
    
    if (foundBatches.length > 0) {
        let html = `<div style="background: #e8f5e9; border: 1px solid #4caf50; padding: 15px; border-radius: 8px;">
            <h3 style="color: #2e7d32; margin-top:0;"><i class="fas fa-check-circle"></i> พบพัสดุ ${input}</h3>
            <p>พบในกลุ่มข้อมูลต่อไปนี้:</p>
            <ul style="margin-bottom: 0;">
        `;
        foundBatches.forEach(b => {
            const dateStr = new Date(b.timestamp).toLocaleString("th-TH");
            html += `<li><strong>${b.batchName}</strong> (${b.type}) - นำเข้าเมื่อ: ${dateStr}</li>`;
        });
        html += `</ul></div>`;
        resultDiv.innerHTML = html;
    } else {
        resultDiv.innerHTML = `<div style="background: #ffebee; border: 1px solid #f44336; padding: 15px; border-radius: 8px;">
            <h3 style="color: #c62828; margin-top:0;"><i class="fas fa-times-circle"></i> ไม่พบพัสดุ ${input}</h3>
            <p style="margin-bottom: 0;">กรุณาตรวจสอบเลขอีกครั้ง หรือพัสดุนี้อาจยังไม่ได้ถูกนำเข้า</p>
        </div>`;
    }
    
    document.getElementById("check-input").value = "";
    document.getElementById("check-input").focus();
}

document.getElementById("check-input").addEventListener("keypress", function(e) {
    if (e.key === "Enter") checkBarcode();
});

// Settings & Auto Cleanup
function saveSettings() {
    const ttl = document.getElementById("setting-ttl").value;
    localStorage.setItem("ta_firebase_ttl", ttl);
    alert(`บันทึกการตั้งค่าแล้ว: เก็บไฟล์ย้อนหลัง ${ttl} วัน`);
}

async function forceCleanup() {
    if(!confirm("ระบบจะตรวจสอบและลบไฟล์ที่หมดอายุตามที่ตั้งค่าไว้ทันที ยืนยันหรือไม่?")) return;
    await performCleanup();
}

async function performCleanup() {
    if (!window.db || !window.storage) return;
    
    const ttlDays = parseInt(localStorage.getItem("ta_firebase_ttl") || "30");
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - ttlDays);
    
    console.log(`Checking for files older than ${ttlDays} days (${cutoffDate.toISOString()})`);
    
    try {
        const snapshot = await window.db.collection("batches").where("timestamp", "<", cutoffDate.toISOString()).get();
        if (snapshot.empty) {
            console.log("No expired batches found.");
            return;
        }
        
        let deletedCount = 0;
        for (const doc of snapshot.docs) {
            const data = doc.data();
            try {
                // Delete file from storage
                if (data.storagePath) {
                    await window.storage.ref().child(data.storagePath).delete();
                }
                // Delete document from firestore
                await window.db.collection("batches").doc(doc.id).delete();
                deletedCount++;
                console.log(`Deleted expired batch: ${data.batchName}`);
            } catch (err) {
                console.error(`Failed to delete batch ${doc.id}`, err);
            }
        }
        
        if (deletedCount > 0) {
            alert(`ลบไฟล์เก่าที่หมดอายุแล้วจำนวน ${deletedCount} รายการ เพื่อคืนพื้นที่ว่างเรียบร้อยแล้ว`);
            loadHistory();
        }
    } catch(err) {
        console.error("Cleanup error", err);
    }
}

// Initialization
document.addEventListener("DOMContentLoaded", () => {
    // Restore settings
    const savedTtl = localStorage.getItem("ta_firebase_ttl");
    if (savedTtl) {
        document.getElementById("setting-ttl").value = savedTtl;
    }
    
    // Auto load and cleanup
    setTimeout(() => {
        loadHistory();
        performCleanup();
    }, 1000);
});

