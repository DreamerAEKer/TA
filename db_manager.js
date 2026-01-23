/**
 * Customer Database Manager (Batch System Refactor)
 * Handles local storage for Tracking Owner Data in Batches.
 */

// 1. Lookup DB: Maps Tracking ID -> Batch ID (For fast search)
const LOOKUP_KEY = 'thp_tracking_lookup_v1';

// 2. Batch DB: Stores Batch Details (The "Sets")
const BATCH_KEY = 'thp_tracking_batches_v1';

const CustomerDB = {
    // --- LOOKUP OPERATIONS ---
    getLookup: () => {
        const raw = localStorage.getItem(LOOKUP_KEY);
        return raw ? JSON.parse(raw) : {};
    },
    saveLookup: (data) => {
        localStorage.setItem(LOOKUP_KEY, JSON.stringify(data));
    },

    // --- BATCH OPERATIONS ---
    getBatches: () => {
        const raw = localStorage.getItem(BATCH_KEY);
        return raw ? JSON.parse(raw) : {}; // { batchId: { ...info } }
    },
    saveBatches: (data) => {
        localStorage.setItem(BATCH_KEY, JSON.stringify(data));
    },

    // --- MAIN ACTIONS ---

    // Compatibility for app.js
    get: (id) => {
        const lookup = CustomerDB.getLookup();
        return lookup[id] || null;
    },

    // Add a new Batch
    addBatch: (batchInfo, trackingList) => {
        const batchId = Date.now().toString(); // Use timestamp as ID
        const batches = CustomerDB.getBatches();
        const lookup = CustomerDB.getLookup();

        // 1. Save Batch Info
        batches[batchId] = {
            id: batchId,
            ...batchInfo,
            count: trackingList.length,
            // Create a short range description (e.g., "ED111... - ED222...")
            rangeDesc: trackingList.length > 1
                ? `${trackingList[0]} - ${trackingList[trackingList.length - 1]}`
                : trackingList[0] || 'No Data'
        };

        // 2. Save Lookup (ID -> Batch ID)
        trackingList.forEach(id => {
            lookup[id] = {
                batchId: batchId,
                name: batchInfo.name,
                type: batchInfo.type
            };
        });

        CustomerDB.saveBatches(batches);
        CustomerDB.saveLookup(lookup);
        return { count: trackingList.length, id: batchId };
    },

    // Delete a Batch
    deleteBatch: (batchId) => {
        const batches = CustomerDB.getBatches();
        const lookup = CustomerDB.getLookup();

        if (!batches[batchId]) return;

        // 1. Remove from Lookup (Expensive but necessary for consistency)
        // Optimization: In a real DB, we'd query by batchId. Here we iterate.
        Object.keys(lookup).forEach(key => {
            if (lookup[key].batchId === batchId) {
                delete lookup[key];
            }
        });

        // 2. Remove Batch
        delete batches[batchId];

        CustomerDB.saveBatches(batches);
        CustomerDB.saveLookup(lookup);
    },

    // --- BACKUP & RESTORE SYSTEM ---
    exportBackup: () => {
        const data = {
            batches: CustomerDB.getBatches(),
            lookup: CustomerDB.getLookup(),
            timestamp: new Date().toISOString(),
            version: "1.0"
        };
        const json = JSON.stringify(data, null, 2);
        const blob = new Blob([json], { type: "application/json" });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = `tracking_backup_${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    },

    importBackup: (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = JSON.parse(e.target.result);
                    if (!data.batches || !data.lookup) {
                        throw new Error("Invalid Backup File Format");
                    }
                    // Restore
                    CustomerDB.saveBatches(data.batches);
                    CustomerDB.saveLookup(data.lookup);
                    resolve(true);
                } catch (err) {
                    reject(err);
                }
            };
            reader.readAsText(file);
        });
    },

    clearAll: () => {
        localStorage.removeItem(BATCH_KEY);
        localStorage.removeItem(LOOKUP_KEY);
    },

    // Find similar (using Lookup DB)
    findSimilarByBody: (trackingNumber) => {
        const lookup = CustomerDB.getLookup();
        const results = [];

        // Parse input
        const regex = /^([A-Z]{2})(\d{9})([A-Z]{2})$/;
        const match = trackingNumber.toUpperCase().match(regex);
        if (!match) return [];

        const [full, prefix, digits, suffix] = match;

        // Search in Lookup keys
        Object.keys(lookup).forEach(key => {
            if (key === trackingNumber) return; // Exact match, skip

            const keyMatch = key.match(regex);
            if (keyMatch) {
                const [kFull, kPrefix, kDigits, kSuffix] = keyMatch;
                if (kDigits === digits && kSuffix === suffix && kPrefix !== prefix) {
                    results.push({
                        number: key,
                        info: lookup[key]
                    });
                }
            }
        });

        return results;
    }
};

// UI State
let currentSort = { key: 'timestamp', order: 'desc' };

function sortDB(key) {
    if (currentSort.key === key) {
        currentSort.order = currentSort.order === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.key = key;
        currentSort.order = 'desc'; // Batches usually desc timestamp
    }
    renderDBTable();
}

/**
 * Prefix & Block1 Managers (Unchanged)
 */
const PREFIX_KEY = 'thp_tracking_prefixes_v1';
const PrefixManager = {
    getAll: () => {
        const raw = localStorage.getItem(PREFIX_KEY);
        return raw ? JSON.parse(raw) : ['EB', 'RB', 'EMS'];
    },
    add: (prefix) => {
        prefix = prefix.toUpperCase().substring(0, 2);
        if (!prefix || prefix.length < 2) return;
        const list = PrefixManager.getAll();
        if (!list.includes(prefix)) {
            list.push(prefix);
            list.sort();
            localStorage.setItem(PREFIX_KEY, JSON.stringify(list));
        }
    },
    remove: (prefix) => {
        let list = PrefixManager.getAll();
        list = list.filter(p => p !== prefix);
        localStorage.setItem(PREFIX_KEY, JSON.stringify(list));
    }
};

const BLOCK1_KEY = 'thp_tracking_block1_v1';
const Block1Manager = {
    getAll: () => {
        const raw = localStorage.getItem(BLOCK1_KEY);
        return raw ? JSON.parse(raw) : ['1234', '5555'];
    },
    add: (val) => {
        val = val.replace(/\D/g, '').substring(0, 4);
        if (!val || val.length < 4) return;
        const list = Block1Manager.getAll();
        if (!list.includes(val)) {
            list.push(val);
            list.sort();
            localStorage.setItem(BLOCK1_KEY, JSON.stringify(list));
        }
    },
    remove: (val) => {
        let list = Block1Manager.getAll();
        list = list.filter(p => p !== val);
        localStorage.setItem(BLOCK1_KEY, JSON.stringify(list));
    }
};

// --- RENDER FUNCTIONS (Adapted for Batches) ---

function renderDBTable() {
    const batches = CustomerDB.getBatches();
    const tbody = document.querySelector('#db-table tbody');
    const countSpan = document.getElementById('db-count');

    if (!tbody || !countSpan) return;

    tbody.innerHTML = '';
    const keys = Object.keys(batches);
    countSpan.innerText = keys.length; // Count of batches, not total items

    // Convert to Array
    const rows = keys.map(key => batches[key]);

    // Sort
    rows.sort((a, b) => {
        let valA, valB;
        if (currentSort.key === 'id') {
            // Sort by Range Desc String?
            valA = a.rangeDesc; valB = b.rangeDesc;
        } else if (currentSort.key === 'name') {
            valA = a.name.toLowerCase(); valB = b.name.toLowerCase();
        } else if (currentSort.key === 'type') {
            valA = a.type; valB = b.type;
        } else {
            // Timestamp default
            valA = a.timestamp || 0; valB = b.timestamp || 0;
        }

        if (valA < valB) return currentSort.order === 'asc' ? -1 : 1;
        if (valA > valB) return currentSort.order === 'asc' ? 1 : -1;
        return 0;
    });

    // Update Headers
    document.querySelectorAll('#db-table th').forEach(th => {
        th.style.background = '#f8f9fa';
        if (th.dataset.sort === currentSort.key) {
            th.style.background = '#e2e6ea';
        }
    });

    rows.forEach(item => {
        const dateStr = item.timestamp ? new Date(item.timestamp).toLocaleString('th-TH') : '-';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <strong>${item.rangeDesc}</strong>
                <br>
                <small style="color:#666;">(${item.count.toLocaleString()} รายการ)</small>
            </td>
            <td>${item.name}</td>
            <td>
                <span class="badge ${item.type === 'Credit' ? 'badge-primary' : 'badge-neutral'}">
                    ${item.type}
                </span>
                ${item.contract ? `<br><small>${item.contract}</small>` : ''}
            </td>
            <td style="font-size:0.85rem; color:#666;">${dateStr}</td>
            <td>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#007bff; color:white; margin-right:5px;" 
                    onclick="loadBatchToView('${item.id}')">🔎 ดู (View)</button>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#dc3545; color:white;" 
                    onclick="deleteBatchEntry('${item.id}', '${item.name}')">ลบ (Del)</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function deleteBatchEntry(batchId, batchName) {
    if (confirm(`ยืนยันการลบข้อมูลชุดนี้?\n\nลูกค้า: ${batchName}`)) {
        CustomerDB.deleteBatch(batchId);
        renderDBTable();
    }
}

function clearAllData() {
    if (confirm('ยืนยันลบข้อมูลทั้งหมด? (ข้อมูลชุดทั้งหมดจะหายไป)')) {
        CustomerDB.clearAll();
        renderDBTable();
    }
}

function saveCustomerData() {
    const name = document.getElementById('db-name').value.trim();
    const type = document.getElementById('db-type').value;
    const contract = document.getElementById('db-contract').value.trim();
    const loadedList = document.getElementById('db-tracking-list').value.trim();

    if (!name || !loadedList) {
        alert('กรุณาระบุชื่อลูกค้า และรายการเลขพัสดุ');
        return;
    }

    const batchInfo = { name, type, contract, timestamp: new Date().getTime() };
    const numbersToAdd = [];

    // Parse list (support Ranges!)
    const lines = loadedList.split(/[\r\n,]+/);
    const regex = /([A-Z]{2})(\d{8})(\d)([A-Z]{2})/;

    lines.forEach(line => {
        line = line.trim().toUpperCase();
        if (!line) return;

        // Check for range "START-END"
        if (line.includes('-')) {
            const parts = line.split('-');
            if (parts.length === 2) {
                const start = parts[0].trim();
                const end = parts[1].trim();
                const mStart = start.match(regex);
                const mEnd = end.match(regex);

                if (mStart && mEnd) {
                    const sVal = parseInt(mStart[2]);
                    const eVal = parseInt(mEnd[2]);
                    const prefix = mStart[1];
                    const suffix = mStart[4];
                    // Generate range
                    for (let i = sVal; i <= eVal; i++) {
                        let body = i.toString().padStart(8, '0');
                        let cd = TrackingUtils.calculateS10CheckDigit(body);
                        numbersToAdd.push(`${prefix}${body}${cd}${suffix}`);
                    }
                }
            }
        } else {
            // Single ID
            const m = line.match(regex);
            if (m) numbersToAdd.push(line);
        }
    });

    if (numbersToAdd.length === 0) {
        alert('ไม่พบเลขพัสดุที่ถูกต้องในรายการ');
        return;
    }

    const count = CustomerDB.addBatch(batchInfo, numbersToAdd);
    alert(`บันทึกเรียบร้อย! เพิ่ม ${count} รายการ (เป็น 1 ชุด)`);

    // Reset inputs
    document.getElementById('db-tracking-list').value = '';
    renderDBTable();
}

// Hook into Window Load to Init Table
window.addEventListener('load', renderDBTable);
