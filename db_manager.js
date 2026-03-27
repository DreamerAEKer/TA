/**
 * Customer Database Manager (Batch System Refactor)
 * Handles local storage for Tracking Owner Data in Batches.
 */

// 1. Lookup DB: Maps Tracking ID -> Batch ID (For fast search)
const LOOKUP_KEY = 'thp_tracking_lookup_v1';

// 2. Batch DB: Stores Batch Details (The "Sets")
const BATCH_KEY = 'thp_tracking_batches_v1';

// 3. Trash DB: Deleted Batches (Soft delete)
const TRASH_KEY = 'thp_tracking_trash_v1';

/**
 * Migration from Legacy Single Entry DB to Batch System
 */
function migrateLegacyData() {
    const KEYS_TO_CHECK = ['thp_tracking_db_v1', 'thp_tracking_db'];
    
    KEYS_TO_CHECK.forEach(OLD_DB_KEY => {
        const raw = localStorage.getItem(OLD_DB_KEY);
        if (!raw) return;

        try {
            const oldDb = JSON.parse(raw);
            const keys = Object.keys(oldDb);
            if (keys.length === 0) return;

            console.log(`Found legacy data in ${OLD_DB_KEY}! Migrating...`, keys.length);

        // Group by Customer Name + Type
        const groups = {};
        keys.forEach(id => {
            const info = oldDb[id];
            const key = `${info.name}_${info.type}`;
            if (!groups[key]) {
                groups[key] = {
                    name: info.name,
                    type: info.type,
                    contract: info.contract || '',
                    timestamp: info.timestamp || 0,
                    items: []
                };
            }
            groups[key].items.push(id);
        });

        // Add as batches
        Object.values(groups).forEach(g => {
            CustomerDB.addBatch({
                name: g.name,
                type: g.type,
                contract: g.contract,
                requestDate: '', 
                timestamp: g.timestamp || new Date().getTime()
            }, g.items);
        });

        // Rename old key to avoid double migration
        localStorage.setItem(OLD_DB_KEY + '_migrated_' + Date.now(), raw);
        localStorage.removeItem(OLD_DB_KEY);
        console.log("Migration complete!");
    } catch (e) {
        console.error("Migration failed:", e);
    }
    });
}

const CustomerDB = {
    // --- LOOKUP OPERATIONS ---
    getLookup: () => {
        const raw = localStorage.getItem(LOOKUP_KEY);
        return raw ? JSON.parse(raw) : {};
    },
    saveLookup: (data) => {
        try {
            localStorage.setItem(LOOKUP_KEY, JSON.stringify(data));
        } catch (e) {
            console.error("saveLookup error:", e);
            if (e.name === 'QuotaExceededError' || e.name === 'NS_ERROR_DOM_QUOTA_REACHED') {
                alert("⚠️ พื้นที่จัดเก็บข้อมูลพัสดุเต็ม! (Storage Full)\nไม่สามารถบันทึกประวัติเพิ่มได้ กรุณา Backup และ Clear ข้อมูลเก่าออกบ้างครับ");
            }
            throw e; // Rethrow to let caller handle if needed
        }
    },

    // --- BATCH OPERATIONS ---
    getBatches: () => {
        const raw = localStorage.getItem(BATCH_KEY);
        return raw ? JSON.parse(raw) : {}; // { batchId: { ...info } }
    },
    saveBatches: (data) => {
        try {
            localStorage.setItem(BATCH_KEY, JSON.stringify(data));
        } catch (e) {
            console.error("saveBatches error:", e);
            if (e.name === 'QuotaExceededError' || e.name === 'NS_ERROR_DOM_QUOTA_REACHED') {
                alert("⚠️ พื้นที่จัดเก็บชุดข้อมูลเต็ม! (Storage Full)\nไม่สามารถบันทึกประวัติเพิ่มได้ กรุณา Backup และ Clear ข้อมูลเก่าออกบ้างครับ");
            }
            throw e;
        }
    },

    // --- TRASH OPERATIONS ---
    getTrash: () => {
        const raw = localStorage.getItem(TRASH_KEY);
        return raw ? JSON.parse(raw) : {}; // { batchId: { ...info, deletedAt, items } }
    },
    saveTrash: (data) => {
        localStorage.setItem(TRASH_KEY, JSON.stringify(data));
    },

    // --- MAIN ACTIONS ---

    // Compatibility for app.js
    get: (id) => {
        const lookup = CustomerDB.getLookup();
        return lookup[id] || null;
    },

    // Group batches by company
    getCompanySummaries: () => {
        const batches = CustomerDB.getBatches();
        const summaries = {};

        Object.values(batches).forEach(b => {
            if (!summaries[b.name]) {
                summaries[b.name] = {
                    name: b.name,
                    type: b.type,
                    totalCount: 0,
                    batches: []
                };
            }
            summaries[b.name].totalCount += b.count;
            summaries[b.name].batches.push(b);
        });

        return Object.values(summaries).sort((a, b) => a.name.localeCompare(b.name));
    },

    // Add a new Batch
    addBatch: (batchInfo, trackingList) => {
        const batches = CustomerDB.getBatches();
        const lookup = CustomerDB.getLookup();

        // --- NEW: DUPLICATE CHECK ---
        // Check if a batch with the exact same tracking numbers already exists
        const existingBatchId = CustomerDB.checkBatchExists(trackingList);
        if (existingBatchId) {
            console.warn(`Duplicate batch detected: ${existingBatchId}`);
            return { error: 'DUPLICATE', id: existingBatchId, count: trackingList.length };
        }

        const batchId = Date.now().toString(); // Use timestamp as ID

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

    // Helper to check if a batch with same items exists
    checkBatchExists: (trackingList) => {
        const lookup = CustomerDB.getLookup();
        if (trackingList.length === 0) return null;
        
        // Find batchId of the first item
        const first = lookup[trackingList[0]];
        if (!first) return null;
        
        const candidateBatchId = first.batchId;
        const batches = CustomerDB.getBatches();
        const candidateBatch = batches[candidateBatchId];
        
        if (!candidateBatch || candidateBatch.count !== trackingList.length) return null;
        
        // Verify ALL items match (Strict check)
        // We can't easily get the list of items from a batch without scanning lookup,
        // but we know it's a duplicate if the lookup for ALL items in trackingList points to candidateBatchId.
        const allMatch = trackingList.every(id => lookup[id] && lookup[id].batchId === candidateBatchId);
        return allMatch ? candidateBatchId : null;
    },

    // Automatic silent deduplication (Keep only one of identical batches)
    deduplicate: () => {
        const batches = CustomerDB.getBatches();
        const lookup = CustomerDB.getLookup();
        const batchArray = Object.values(batches);
        const seenSignatures = new Set();
        const toDeleteIds = [];

        // Sort by timestamp (keep oldest first to maintain history)
        batchArray.sort((a, b) => a.timestamp - b.timestamp);

        batchArray.forEach(batch => {
            // Create a signature of the batch content
            // We need the items. We can get them from lookup.
            const items = Object.keys(lookup).filter(k => lookup[k].batchId === batch.id).sort();
            const signature = `${batch.name}|${batch.type}|${items.join(',')}`;

            if (seenSignatures.has(signature)) {
                toDeleteIds.push(batch.id);
            } else {
                seenSignatures.add(signature);
            }
        });

        if (toDeleteIds.length > 0) {
            console.log(`Auto-Deduplicate: Removing ${toDeleteIds.length} duplicate batches...`);
            toDeleteIds.forEach(id => {
                delete batches[id];
                // Remove from lookup if it points to this batch
                Object.keys(lookup).forEach(k => {
                    if (lookup[k].batchId === id) delete lookup[k];
                });
            });
            CustomerDB.saveBatches(batches);
            CustomerDB.saveLookup(lookup);
        }
    },

    // Delete a Batch (Move to Trash)
    deleteBatch: (batchId) => {
        const batches = CustomerDB.getBatches();
        const lookup = CustomerDB.getLookup();
        const trash = CustomerDB.getTrash();

        if (!batches[batchId]) return;

        const batchToTrash = { ...batches[batchId], deletedAt: Date.now() };
        const itemsToTrash = [];

        // 1. Remove from Lookup and save items for restoring later
        Object.keys(lookup).forEach(key => {
            if (lookup[key].batchId === batchId) {
                itemsToTrash.push(key);
                delete lookup[key];
            }
        });

        // 2. Save into Trash
        batchToTrash.items = itemsToTrash;
        trash[batchId] = batchToTrash;

        // 3. Remove Batch
        delete batches[batchId];

        CustomerDB.saveBatches(batches);
        CustomerDB.saveLookup(lookup);
        CustomerDB.saveTrash(trash);
    },

    // Restore a Batch from Trash
    restoreTrash: (batchId) => {
        const trash = CustomerDB.getTrash();
        const batches = CustomerDB.getBatches();
        const lookup = CustomerDB.getLookup();

        if (!trash[batchId]) return;

        const batchToRestore = trash[batchId];
        const items = batchToRestore.items || [];

        // Clean up trash specific fields
        delete batchToRestore.deletedAt;
        delete batchToRestore.items;

        batches[batchId] = batchToRestore;

        items.forEach(id => {
            lookup[id] = {
                batchId: batchId,
                name: batchToRestore.name,
                type: batchToRestore.type
            };
        });

        delete trash[batchId];

        CustomerDB.saveBatches(batches);
        CustomerDB.saveLookup(lookup);
        CustomerDB.saveTrash(trash);
    },

    permanentlyDeleteTrash: (batchId) => {
        const trash = CustomerDB.getTrash();
        if (trash[batchId]) {
            delete trash[batchId];
            CustomerDB.saveTrash(trash);
        }
    },

    emptyTrash: () => {
        localStorage.removeItem(TRASH_KEY);
    },

    // --- BACKUP & RESTORE SYSTEM ---
    exportBackup: () => {
        const data = {
            batches: CustomerDB.getBatches(),
            lookup: CustomerDB.getLookup(),
            trash: CustomerDB.getTrash(),
            timestamp: new Date().toISOString(),
            version: "1.1"
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
                    if (data.trash) CustomerDB.saveTrash(data.trash);
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
        // Do we want to clear trash too? Yes, clearAll means factory reset-ish.
        localStorage.removeItem(TRASH_KEY);
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
    },

    // เล่มที่ (Book) Analysis: 1 book = 240 items
    findBookMatch: (trackingNumber) => {
        const lookup = CustomerDB.getLookup();
        
        // Parse input: standard XX 12345678 9 TH
        const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
        const match = trackingNumber.toUpperCase().match(regex);
        if (!match) return null;

        const [full, prefix, bodyStr, cd, suffix] = match;
        const body = parseInt(bodyStr);
        if (isNaN(body)) return null;

        // ตรรกะเล่มพัสดุ: เริ่มที่ ((เล่ม-1)*240)+1 จบที่ เล่ม*240
        const bookStart = Math.floor((body - 1) / 240) * 240 + 1;
        const bookEnd = bookStart + 239;
        
        // ฟังก์ชันช่วยหา Check Digit (ใช้ของ TrackingUtils ถ้ามี)
        const getCD = (bStr) => {
            if (typeof TrackingUtils !== 'undefined' && TrackingUtils.calculateS10CheckDigit) {
                return TrackingUtils.calculateS10CheckDigit(bStr);
            }
            // Fallback: S10 weights [8, 6, 4, 2, 3, 5, 9, 7]
            const weights = [8, 6, 4, 2, 3, 5, 9, 7];
            let sum = 0;
            for (let i = 0; i < 8; i++) sum += parseInt(bStr[i]) * weights[i];
            const rem = sum % 11;
            let d = 11 - rem;
            if (d === 10) return 0;
            if (d === 11) return 5;
            return d;
        };

        // สแกนเฉพาะตัวเลขที่อยู่ในเล่มเดียวกัน (สูงสุด 240 ครั้ง)
        for (let i = bookStart; i <= bookEnd; i++) {
            if (i === body) continue; // เลขเดียวกันข้ามไป
            
            const bStr = i.toString().padStart(8, '0');
            if (bStr.length > 8) continue;

            const checkDigit = getCD(bStr);
            const key = `${prefix}${bStr}${checkDigit}${suffix}`;
            
            if (lookup[key]) {
                return {
                    ...lookup[key],
                    refNumber: key,
                    method: 'book'
                };
            }
        }
        
        return null;
    },
    // --- SNAPSHOT SYSTEM (Auto Backup) ---
    getSnapshots: () => {
        const raw = localStorage.getItem('thp_snapshots_v1');
        return raw ? JSON.parse(raw) : [];
    },

    createSnapshot: (reason = 'Auto-Backup') => {
        const snapshots = CustomerDB.getSnapshots();

        // 1. Capture Current State
        const currentState = {
            timestamp: Date.now(),
            reason: reason,
            data: {
                batches: CustomerDB.getBatches(),
                lookup: CustomerDB.getLookup()
            }
        };

        // 2. Add to list (Limit to 5 recent snapshots)
        snapshots.unshift(currentState);
        if (snapshots.length > 5) snapshots.pop();

        // 3. Save
        localStorage.setItem('thp_snapshots_v1', JSON.stringify(snapshots));
        console.log(`Snapshot created: ${reason}`);
    },

    restoreSnapshot: (timestamp) => {
        const snapshots = CustomerDB.getSnapshots();
        const target = snapshots.find(s => s.timestamp === timestamp);

        if (target) {
            CustomerDB.saveBatches(target.data.batches);
            CustomerDB.saveLookup(target.data.lookup);
            return true;
        }
        return false;
    }
};

// --- UPDATE: Wrap critical actions with Snapshot ---

// Monkey Patch importBackup to auto-snapshot
const originalImport = CustomerDB.importBackup;
CustomerDB.importBackup = async (file) => {
    // Auto Snapshot before Restore
    CustomerDB.createSnapshot('Pre-Restore Backup');
    return originalImport(file);
};

// Monkey Patch clearAll
const originalClear = CustomerDB.clearAll;
CustomerDB.clearAll = () => {
    if (Object.keys(CustomerDB.getBatches()).length > 0) {
        CustomerDB.createSnapshot('Pre-ClearAll Backup');
    }
    originalClear();
};

// UI State
let currentSort = { key: 'timestamp', order: 'desc' };

function renderCustomerList() {
    const list = document.getElementById('customer-name-list');
    const selectBox = document.getElementById('db-name-select');
    const dbNameInput = document.getElementById('db-name');
    if (!list) return;

    const batches = CustomerDB.getBatches();
    const customers = {}; // name -> { type, contract, timestamp }

    // Extract unique customers
    Object.values(batches).forEach(b => {
        if (!customers[b.name] || b.timestamp > customers[b.name].timestamp) {
            customers[b.name] = { type: b.type, contract: b.contract, timestamp: b.timestamp || 0 };
        }
    });

    list.innerHTML = '';
    if (selectBox) selectBox.innerHTML = '<option value="">- เลือกลูกค้าประวัติ -</option>';

    Object.keys(customers).sort().forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        list.appendChild(opt);

        if (selectBox) {
            const selectOpt = document.createElement('option');
            selectOpt.value = name;
            selectOpt.textContent = name;
            selectBox.appendChild(selectOpt);
        }
    });

    // Add event listener to auto-fill (only once)
    if (dbNameInput && !dbNameInput.hasAttribute('data-listener-added')) {
        dbNameInput.setAttribute('data-listener-added', 'true');
        dbNameInput.addEventListener('change', (e) => {
            const selectedName = e.target.value;
            if (customers[selectedName]) {
                const typeSelect = document.getElementById('db-type');
                const contractInput = document.getElementById('db-contract');
                if (typeSelect) typeSelect.value = customers[selectedName].type;
                if (contractInput) contractInput.value = customers[selectedName].contract || '';
            }
        });
    }
}

window.selectCustomerFromDropdown = function(name) {
    if (!name) return;
    const dbNameInput = document.getElementById('db-name');
    if (dbNameInput) {
        dbNameInput.value = name;
        dbNameInput.dispatchEvent(new Event('change'));
    }
};

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

// --- UI View Toggles ---
let currentDbView = 'recent'; // 'recent', 'company', 'trash'
let currentCompanySort = { key: 'totalCount', order: 'desc' };

function sortCompanyDB(key) {
    if (currentCompanySort.key === key) {
        currentCompanySort.order = currentCompanySort.order === 'asc' ? 'desc' : 'asc';
    } else {
        currentCompanySort.key = key;
        currentCompanySort.order = (key === 'name' || key === 'type') ? 'asc' : 'desc';
    }
    renderCompanyTable();
}

function toggleCompanyRow(idx) {
    const row = document.getElementById(`comp-batches-${idx}`);
    if (row) {
        row.style.display = row.style.display === 'none' ? 'table-row' : 'none';
    }
}

function toggleCompanyView() {
    currentDbView = currentDbView === 'company' ? 'recent' : 'company';
    updateDbViews();
}

function toggleTrashView() {
    currentDbView = currentDbView === 'trash' ? 'recent' : 'trash';
    updateDbViews();
}

function updateDbViews() {
    const listCont = document.getElementById('db-list-container');
    const compCont = document.getElementById('db-company-container');
    const trashCont = document.getElementById('db-trash-container');
    
    if(listCont) listCont.classList.add('hidden');
    if(compCont) compCont.classList.add('hidden');
    if(trashCont) trashCont.classList.add('hidden');
    
    const btnComp = document.getElementById('btn-toggle-company');
    const btnTrash = document.getElementById('btn-toggle-trash');
    
    if(btnComp) {
        btnComp.classList.remove('active', 'btn-primary');
        btnComp.classList.add('btn-neutral');
    }
    if(btnTrash) {
        btnTrash.classList.remove('active', 'btn-danger');
        btnTrash.classList.add('btn-neutral');
    }

    if (currentDbView === 'recent') {
        if(listCont) listCont.classList.remove('hidden');
        renderDBTable();
    } else if (currentDbView === 'company') {
        if(compCont) compCont.classList.remove('hidden');
        if(btnComp) {
            btnComp.classList.remove('btn-neutral');
            btnComp.classList.add('btn-primary');
        }
        renderCompanyTable();
    } else if (currentDbView === 'trash') {
        if(trashCont) trashCont.classList.remove('hidden');
        if(btnTrash) {
            btnTrash.classList.remove('btn-neutral');
            btnTrash.classList.add('btn-danger'); // Use red for trash
        }
        renderTrashTable();
    }
    
    // Update Trash Count
    const trashSpan = document.getElementById('trash-count');
    if(trashSpan) {
        trashSpan.innerText = Object.keys(CustomerDB.getTrash()).length;
    }
}

// --- RENDER FUNCTIONS (Adapted for Batches) ---

function renderDBTable() {
    renderCustomerList(); // Automatically update dropdown when table renders
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
            valA = a.rangeDesc; valB = b.rangeDesc;
        } else if (currentSort.key === 'name') {
            valA = a.name.toLowerCase(); valB = b.name.toLowerCase();
        } else if (currentSort.key === 'type') {
            valA = a.type; valB = b.type;
        } else {
            valA = a.timestamp || 0; valB = b.timestamp || 0;
        }

        if (valA < valB) return currentSort.order === 'asc' ? -1 : 1;
        if (valA > valB) return currentSort.order === 'asc' ? 1 : -1;
        return 0;
    });

    // Update Headers
    document.querySelectorAll('#db-table th').forEach(th => {
        if(!th.dataset.sort) return;
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
                ${item.requestDate ? `<br><small style="color:#28a745;">📅 ขอเลข: ${new Date(item.requestDate).toLocaleDateString('th-TH')}</small>` : ''}
            </td>
            <td style="font-size:0.85rem; color:#666;">${dateStr}</td>
            <td>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#007bff; color:white; margin-right:5px; margin-bottom:5px;" 
                    onclick="loadBatchToView('${item.id}')">🔎 ดู (View)</button>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#ffc107; color:black; margin-right:5px; margin-bottom:5px;" 
                    onclick="loadBatchToEdit('${item.id}')">✏️ แก้ไขชุดนี้</button>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#dc3545; color:white; margin-bottom:5px;" 
                    onclick="deleteBatchEntry('${item.id}', '${item.name}')">ลบ (Del)</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderCompanyTable() {
    const tbody = document.querySelector('#company-table tbody');
    if (!tbody) return;
    
    let summaries = CustomerDB.getCompanySummaries();
    tbody.innerHTML = '';
    
    if (summaries.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">ไม่พบข้อมูลบริษัท</td></tr>';
        return;
    }

    // Sort summaries
    summaries.sort((a, b) => {
        let valA, valB;
        switch(currentCompanySort.key) {
            case 'name': valA = a.name.toLowerCase(); valB = b.name.toLowerCase(); break;
            case 'totalCount': valA = a.totalCount; valB = b.totalCount; break;
            case 'batches': valA = a.batches.length; valB = b.batches.length; break;
            case 'type': valA = a.type; valB = b.type; break;
            default: valA = a.name; valB = b.name;
        }
        if (valA < valB) return currentCompanySort.order === 'asc' ? -1 : 1;
        if (valA > valB) return currentCompanySort.order === 'asc' ? 1 : -1;
        return 0;
    });

    // Update Headers UI
    document.querySelectorAll('#company-table th').forEach(th => {
        if(!th.dataset.sortcomp) return;
        th.style.background = '#f8f9fa';
        if (th.dataset.sortcomp === currentCompanySort.key) {
            th.style.background = '#e2e6ea';
        }
    });
    
    summaries.forEach((company, idx) => {
        // Master Row
        const tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        tr.onclick = () => toggleCompanyRow(idx);
        tr.innerHTML = `
            <td><strong>${company.name}</strong></td>
            <td><strong>${company.totalCount.toLocaleString()}</strong> รายการ</td>
            <td><span style="border-bottom: 1px dashed #666;">${company.batches.length} ชุด ▼</span></td>
            <td><span class="badge ${company.type === 'Credit' ? 'badge-primary' : 'badge-neutral'}">${company.type}</span></td>
        `;
        tbody.appendChild(tr);

        // Detail Row (Hidden by default)
        const trDetail = document.createElement('tr');
        trDetail.id = `comp-batches-${idx}`;
        trDetail.style.display = 'none';
        
        let batchRowsHtml = company.batches.map((b, bidx) => {
            const reqDate = b.requestDate ? new Date(b.requestDate).toLocaleDateString('th-TH') : '-';
            return `
                <tr style="border-bottom:1px solid #eee;">
                    <td style="padding:5px;">ชุดที่ ${bidx+1}: <strong>${b.rangeDesc}</strong></td>
                    <td style="padding:5px; color:#28a745;">ขอเลขที่: ${reqDate}</td>
                    <td style="padding:5px;">จำนวน: ${b.count} รายการ</td>
                    <td style="padding:5px; text-align:right;">
                        <button class="btn" style="padding:2px 6px; font-size:0.65rem; background:#ffc107; color:black;" 
                            onclick="loadBatchToEdit('${b.id}')">✏️ แก้ไข</button>
                    </td>
                </tr>
            `;
        }).join('');

        trDetail.innerHTML = `
            <td colspan="4" style="background:#f1f3f5; padding:0;">
                <div style="padding:10px; padding-left:20px; border-left:3px solid var(--primary-color);">
                    <table style="width:100%; font-size:0.85rem; border-collapse:collapse; background:transparent;">
                        ${batchRowsHtml}
                    </table>
                </div>
            </td>
        `;
        tbody.appendChild(trDetail);
    });
}

function renderTrashTable() {
    const tbody = document.querySelector('#trash-table tbody');
    if (!tbody) return;
    
    const trash = CustomerDB.getTrash();
    tbody.innerHTML = '';
    
    const keys = Object.keys(trash);
    if (keys.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">ถังขยะว่างเปล่า</td></tr>';
        return;
    }
    
    const rows = keys.map(k => trash[k]).sort((a,b) => b.deletedAt - a.deletedAt);
    
    rows.forEach(item => {
        const deleteDate = new Date(item.deletedAt).toLocaleString('th-TH');
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <strong>${item.rangeDesc}</strong><br>
                <small>(${item.count.toLocaleString()} รายการ)</small>
            </td>
            <td>${item.name}</td>
            <td style="color:#d32f2f; font-size:0.85rem;">${deleteDate}</td>
            <td>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#007bff; color:white; margin-right:5px;" 
                    onclick="restoreTrashItem('${item.id}', '${item.name}')">♻️ กู้คืน (Restore)</button>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#dc3545; color:white;" 
                    onclick="hardDeleteTrashItem('${item.id}', '${item.name}')">ลบทิ้งถาวร</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function restoreTrashItem(batchId, batchName) {
    if (confirm(`ต้องการกู้คืนข้อมูลชุดนี้ใช่หรือไม่?\n\nลูกค้า: ${batchName}`)) {
        CustomerDB.restoreTrash(batchId);
        updateDbViews();
    }
}

function hardDeleteTrashItem(batchId, batchName) {
    if (confirm(`⚠️ ยืนยันการลบทิ้งถาวร ไม่สามารถกู้คืนได้อีก!\n\nลูกค้า: ${batchName}`)) {
        CustomerDB.permanentlyDeleteTrash(batchId);
        updateDbViews();
    }
}

function emptyTrash() {
    if (confirm('⚠️ ยืนยันการล้างถังขยะทั้งหมด ข้อมูลในถังขยะจะถูกลบทิ้งอย่างถาวร ไม่สามารถกู้คืนได้!')) {
        CustomerDB.emptyTrash();
        updateDbViews();
    }
}

function loadBatchToEdit(batchId) {
    const batches = CustomerDB.getBatches();
    const batch = batches[batchId];
    if (!batch) return;

    if (!confirm(`นำข้อมูลชุดนี้ของลูกค้า "${batch.name}" ขึ้นมาแก้ไข?\n(ข้อมูลเดิมจะถูกย้ายไปถังขยะชั่วคราว เมื่อแก้ไขและกดบันทึกจะเป็นการสร้างชุดใหม่แทนที่)`)) return;

    // Fill upper form
    document.getElementById('db-name').value = batch.name;
    document.getElementById('db-type').value = batch.type || 'Credit';
    document.getElementById('db-contract').value = batch.contract || '';
    document.getElementById('db-request-date').value = batch.requestDate || '';

    // Move to Trash (Soft Delete)
    CustomerDB.deleteBatch(batchId);

    // Get items from trash to display
    const trash = CustomerDB.getTrash();
    const allItemsToEdit = (trash[batchId] && trash[batchId].items) ? trash[batchId].items : batch.items;

    // Populate textarea
    const textArea = document.getElementById('db-tracking-list');
    textArea.value = allItemsToEdit.join('\n');
    
    // Unhide Edit Section & focus
    document.getElementById('db-edit-section').classList.remove('hidden');
    textArea.focus();
    
    // Jump to form
    window.scrollTo({ top: 0, behavior: 'smooth' });
    
    alert(`ดึงเลขพัสดุชุดนี้ จำนวน ${allItemsToEdit.length} รายการ ขึ้นมาแก้ไขแล้ว!\n(ข้อมูลชุดเดิมถูกย้ายไปที่ถังขยะชั่วคราว หากยกเลิกสามารถตามไปกู้คืนได้)`);
    updateDbViews();
}

function deleteBatchEntry(batchId, batchName) {
    if (confirm(`ยืนยันการลบข้อมูลชุดนี้?\n\nลูกค้า: ${batchName}`)) {
        CustomerDB.deleteBatch(batchId);
        updateDbViews();
    }
}

function cancelDbEdit() {
    if (confirm('ยกเลิกการแก้ไขใช่หรือไม่? (ถ้ากดยกเลิก ข้อมูลที่แก้ไขจะไม่ถูกบันทึก และข้อมูลรายบริษัทนี้จะอยู่ในถังขยะ ให้ไปกดกู้คืนจากถังขยะ)')) {
        document.getElementById('db-edit-section').classList.add('hidden');
        document.getElementById('db-tracking-list').value = '';
        document.getElementById('db-name').value = '';
        document.getElementById('db-contract').value = '';
        document.getElementById('db-request-date').value = '';
    }
}

function clearDbEditData() {
    if (confirm('ยืนยันล้างข้อความในกล่องนี้ทิ้งทั้งหมด?')) {
        document.getElementById('db-tracking-list').value = '';
    }
}

function clearAllData() {
    if (confirm('ยืนยันล้างข้อมูลทั้งหมดในฐานข้อมูลลูกค้า? (ข้อมูลชุดทั้งหมดจะหายไป)')) {
        CustomerDB.clearAll();
        updateDbViews();
    }
}

function saveCustomerData() {
    const name = document.getElementById('db-name').value.trim();
    const type = document.getElementById('db-type').value;
    const contract = document.getElementById('db-contract').value.trim();
    const requestDate = document.getElementById('db-request-date').value;
    const loadedList = document.getElementById('db-tracking-list').value.trim();

    if (!name || !loadedList) {
        alert('กรุณาระบุชื่อลูกค้า และรายการเลขพัสดุ');
        return;
    }

    const batchInfo = { name, type, contract, requestDate, timestamp: new Date().getTime() };
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

    // Save to DB
    const btn = event?.target || document.querySelector('button[onclick="saveCustomerData()"]');
    const originalHtml = btn ? btn.innerHTML : 'บันทึกข้อมูลที่แก้ไข (Save Edit)';
    
    if (btn) window.setButtonLoading(btn, true);

    setTimeout(() => {
        const result = CustomerDB.addBatch(batchInfo, numbersToAdd);
        
        if (btn) window.setButtonLoading(btn, false, originalHtml);

        if (result && result.error === 'DUPLICATE') {
            window.showToast(`ข้อมูลชุดนี้มีอยู่แล้วในระบบ (Batch: ${result.id})`, 'info');
        } else {
            const count = typeof result === 'object' ? result.count : result;
            window.showToast(`บันทึกเรียบร้อย! เพิ่ม ${count} รายการ`);
        }

        // Reset inputs & hide edit section
        document.getElementById('db-name').value = '';
        document.getElementById('db-contract').value = '';
        document.getElementById('db-request-date').value = '';
        document.getElementById('db-tracking-list').value = '';
        document.getElementById('db-edit-section').classList.add('hidden');
        
        currentDbView = 'recent';
        updateDbViews();
    }, 500); 
}

// Hook into Window Load to Init Table
window.addEventListener('load', () => {
    migrateLegacyData();
    CustomerDB.deduplicate(); // Clean existing duplicates once on load
    updateDbViews();
});

/**
 * Exception Manager (For Tracking Errors / Reasons)
 */
const EXCEPTION_KEY = 'thp_tracking_exceptions_v1';

const ExceptionManager = {
    getAll: () => {
        try {
            const raw = localStorage.getItem(EXCEPTION_KEY);
            return raw ? JSON.parse(raw) : [];
        } catch(e) {
            console.error("Error loading exceptions", e);
            return [];
        }
    },

    /**
     * Save a batch of tracking numbers in one session.
     * @param {string[]} trackNums - Array of tracking numbers
     * @param {string} companyName - Company name
     * @param {string} reason - Reason
     */
    saveSession: (trackNums, companyName, reason, firstStatus = 'ใส่ของลงถุง', dateTime = '', images = [], existingSessionId = null, metadata = {}) => {
        const exceptions = ExceptionManager.getAll();
        const sessionId = existingSessionId || Date.now().toString();

        // Remove any existing entries for these numbers (overwrite)
        const filtered = exceptions.filter(e => !trackNums.includes(e.trackNum));

        trackNums.forEach((trackNum, i) => {
            filtered.push({
                id: `${sessionId}_${i}`,
                trackNum: trackNum,
                companyName: companyName,
                reason: reason,
                firstStatus: firstStatus,
                dateTime: dateTime,
                images: images, 
                timestamp: new Date().toISOString(),
                sessionId: sessionId,
                // New metadata fields
                category: (metadata && metadata.category) || 'เงินสด',
                branch: (metadata && metadata.branch) || '',
                reporter: (metadata && metadata.reporter) || '',
                subject: (metadata && metadata.subject) || '',
                note: (metadata && metadata.note) || ''
            });
        });

        try {
            localStorage.setItem(EXCEPTION_KEY, JSON.stringify(filtered));
            return sessionId;
        } catch (e) {
            console.error("ExceptionManager.saveSession Error:", e);
            if (e.name === 'QuotaExceededError' || e.name === 'NS_ERROR_DOM_QUOTA_REACHED') {
                alert("⚠️ พื้นที่จัดเก็บข้อมูลเต็ม! (Storage Full)\n\nสาเหตุ: รูปภาพที่แนบมีขนาดใหญ่เกินกว่าที่เบราว์เซอร์จะรับได้ในระบบประวัติ\nวิธีแก้ไข: โปรดล้างรูปภาพบางส่วนออก หรือใช้รูปที่มีความละเอียดน้อยลง");
            } else {
                alert("❌ เกิดข้อผิดพลาดในการบันทึก: " + e.message);
            }
            return null;
        }
    },

    // Legacy single-save (kept for compatibility)
    save: (trackNum, companyName, reason) => {
        ExceptionManager.saveSession([trackNum], companyName, reason);
    },

    removeSession: (sessionId) => {
        const exceptions = ExceptionManager.getAll();
        const filtered = exceptions.filter(e => {
            const sid = e.sessionId || e.id;
            return sid !== sessionId;
        });
        localStorage.setItem(EXCEPTION_KEY, JSON.stringify(filtered));
    },

    remove: (id) => {
        const exceptions = ExceptionManager.getAll();
        const filtered = exceptions.filter(e => e.id !== id);
        localStorage.setItem(EXCEPTION_KEY, JSON.stringify(filtered));
    },

    clearAll: () => {
        localStorage.removeItem(EXCEPTION_KEY);
    }
};
