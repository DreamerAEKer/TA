/**
 * Customer Database Manager (Batch System Refactor)
 * v1.64: Fully Asynchronous IndexedDB Storage
 */

// --- STORAGE KEYS ---
const LOOKUP_KEY = 'thp_tracking_lookup_v1';
const BATCH_KEY = 'thp_tracking_batches_v1';
const TRASH_KEY = 'thp_tracking_trash_v1';
const EXCEPTION_KEY = 'thp_tracking_exceptions_v1';
const SNAPSHOTS_KEY = 'thp_snapshots_v1';

/** 
 * v1.64: StorageV2 (IndexedDB Wrapper)
 * Bypasses the 5MB localStorage limit.
 */
const DB_NAME = 'thp_tracking_db_v2';
const DB_VERSION = 1;
const STORE_NAME = 'key_value_store';

const StorageV2 = {
    _db: null,
    _initPromise: null,

    init: () => {
        if (StorageV2._db) return Promise.resolve(StorageV2._db);
        if (StorageV2._initPromise) return StorageV2._initPromise;

        StorageV2._initPromise = new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME);
                }
            };
            request.onsuccess = (e) => {
                StorageV2._db = e.target.result;
                resolve(StorageV2._db);
            };
            request.onerror = (e) => {
                StorageV2._initPromise = null; // Allow retry on error
                reject(e.target.error);
            };
        });
        return StorageV2._initPromise;
    },

    get: async (key) => {
        await StorageV2.init();
        return new Promise((resolve, reject) => {
            const transaction = StorageV2._db.transaction([STORE_NAME], 'readonly');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.get(key);
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    },

    set: async (key, value) => {
        await StorageV2.init();
        return new Promise((resolve, reject) => {
            const transaction = StorageV2._db.transaction([STORE_NAME], 'readwrite');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.put(value, key);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });
    },

    remove: async (key) => {
        await StorageV2.init();
        return new Promise((resolve, reject) => {
            const transaction = StorageV2._db.transaction([STORE_NAME], 'readwrite');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.delete(key);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });
    }
};

const CustomerDB = {
    // --- INITIALIZATION & MIGRATION ---
    init: async () => {
        await StorageV2.init();
        const migratedKey = 'thp_v1_to_v2_migrated';
        if (!localStorage.getItem(migratedKey)) {
            console.log("v1.64: Starting migration to IndexedDB...");
            const keysToMigrate = [LOOKUP_KEY, BATCH_KEY, TRASH_KEY, EXCEPTION_KEY, SNAPSHOTS_KEY];
            for (const key of keysToMigrate) {
                const val = localStorage.getItem(key);
                if (val) {
                    try {
                        await StorageV2.set(key, JSON.parse(val));
                        console.log(`Migrated ${key}`);
                    } catch (e) {
                        console.error(`Migration failed for ${key}`, e);
                    }
                }
            }
            localStorage.setItem(migratedKey, 'true');
            console.log("v1.64: Migration complete!");
        }
    },

    // --- DATA ACCESSORS (ASYNC) ---
    getLookup: async () => (await StorageV2.get(LOOKUP_KEY)) || {},
    saveLookup: async (data) => await StorageV2.set(LOOKUP_KEY, data),

    getBatches: async () => (await StorageV2.get(BATCH_KEY)) || {},
    saveBatches: async (data) => await StorageV2.set(BATCH_KEY, data),

    getTrash: async () => (await StorageV2.get(TRASH_KEY)) || {},
    saveTrash: async (data) => await StorageV2.set(TRASH_KEY, data),

    // --- CORE LOGIC (ASYNC) ---
    get: async (id) => {
        if (!id) return null;
        const lookup = await CustomerDB.getLookup();
        if (lookup[id]) return lookup[id];

        const batches = await CustomerDB.getBatches();
        const tid = id.replace(/\s+/g, '').toUpperCase();
        const m = tid.match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
        if (!m) return null;

        const prefix = m[1], body = parseInt(m[2], 10), suffix = m[4];
        for (const bid in batches) {
            const b = batches[bid];
            if (!b.rangeDesc || !b.rangeDesc.includes(' - ')) continue;
            const parts = b.rangeDesc.split(' - ');
            const mS = parts[0].match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
            const mE = parts[1].match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
            if (mS && mE) {
                if (prefix === mS[1] && suffix === mS[4] && body >= parseInt(mS[2], 10) && body <= parseInt(mE[2], 10)) {
                    return { batchId: b.id, name: b.name, type: b.type, isRangeMatch: true };
                }
            }
        }
        return null;
    },

    addBatch: async (batchInfo, trackingList) => {
        const batches = await CustomerDB.getBatches();
        const lookup = await CustomerDB.getLookup();
        
        // v1.93: Duplicate Prevention (Check same name + rangeDesc to prevent accidental double-clicks)
        const currentDesc = trackingList.length > 1
            ? `${trackingList[0]} - ${trackingList[trackingList.length - 1]}`
            : trackingList[0] || 'No Data';

        const possibleDup = Object.values(batches).find(b => b.name === batchInfo.name && b.rangeDesc === currentDesc);
        if (possibleDup) {
            return { error: 'DUPLICATE', id: possibleDup.id };
        }

        const batchId = Date.now().toString();

        batches[batchId] = {
            id: batchId,
            ...batchInfo,
            count: trackingList.length,
            rangeDesc: trackingList.length > 1
                ? `${trackingList[0]} - ${trackingList[trackingList.length - 1]}`
                : trackingList[0] || 'No Data'
        };

        if (trackingList.length <= 100) {
            trackingList.forEach(id => {
                lookup[id] = { batchId, name: batchInfo.name, type: batchInfo.type };
            });
            await CustomerDB.saveLookup(lookup);
        }
        await CustomerDB.saveBatches(batches);
        return { count: trackingList.length, id: batchId };
    },

    deleteBatch: async (batchId) => {
        const batches = await CustomerDB.getBatches();
        const lookup = await CustomerDB.getLookup();
        const trash = await CustomerDB.getTrash();
        if (!batches[batchId]) return;

        const batch = { ...batches[batchId], deletedAt: Date.now(), items: [] };
        Object.keys(lookup).forEach(k => {
            if (lookup[k].batchId === batchId) {
                batch.items.push(k);
                delete lookup[k];
            }
        });

        trash[batchId] = batch;
        delete batches[batchId];
        await CustomerDB.saveBatches(batches);
        await CustomerDB.saveLookup(lookup);
        await CustomerDB.saveTrash(trash);
    },

    restoreTrash: async (batchId) => {
        const trash = await CustomerDB.getTrash();
        const batches = await CustomerDB.getBatches();
        const lookup = await CustomerDB.getLookup();
        if (!trash[batchId]) return;

        const batch = trash[batchId];
        const items = batch.items || [];
        delete batch.deletedAt; delete batch.items;
        batches[batchId] = batch;
        items.forEach(id => { lookup[id] = { batchId, name: batch.name, type: batch.type }; });
        delete trash[batchId];

        await CustomerDB.saveBatches(batches);
        await CustomerDB.saveLookup(lookup);
        await CustomerDB.saveTrash(trash);
    },

    permanentDeleteTrash: async (batchId) => {
        const trash = await CustomerDB.getTrash();
        delete trash[batchId];
        await CustomerDB.saveTrash(trash);
    },

    deduplicate: async () => {
        const batches = await CustomerDB.getBatches();
        const lookup = await CustomerDB.getLookup();
        const seen = new Set();
        let changed = false;
        Object.keys(batches).forEach(bid => {
            const b = batches[bid];
            const key = `${b.name}|${b.type}|${b.rangeDesc}|${b.count}`;
            if (seen.has(key)) {
                Object.keys(lookup).forEach(k => { if (lookup[k].batchId === bid) delete lookup[k]; });
                delete batches[bid];
                changed = true;
            } else seen.add(key);
        });
        if (changed) { await CustomerDB.saveBatches(batches); await CustomerDB.saveLookup(lookup); }
    },

    optimizeStorage: async () => {
        const lookup = await CustomerDB.getLookup();
        const batches = await CustomerDB.getBatches();
        let removed = 0;
        Object.keys(lookup).forEach(id => {
            const b = batches[lookup[id].batchId];
            if (b && b.rangeDesc && b.rangeDesc.includes(' - ')) { delete lookup[id]; removed++; }
        });
        if (removed > 0) await CustomerDB.saveLookup(lookup);
        return removed;
    },

    getCompanySummaries: async () => {
        const batches = await CustomerDB.getBatches();
        const sums = {};
        Object.values(batches).forEach(b => {
            if (!sums[b.name]) sums[b.name] = { name: b.name, type: b.type, totalCount: 0, batches: [] };
            sums[b.name].totalCount += b.count; sums[b.name].batches.push(b);
        });
        return Object.values(sums).sort((a,b) => a.name.localeCompare(b.name));
    }
};

/**
 * Exception Manager (ASYNC)
 */
const ExceptionManager = {
    getAll: async () => (await StorageV2.get(EXCEPTION_KEY)) || [],
    saveSession: async (trackNums, companyName, reason, firstStatus = 'ใส่ของลงถุง', dateTime = '', images = [], existingId = null, metadata = {}) => {
        const all = await ExceptionManager.getAll();
        const sid = existingId || Date.now().toString();
        let filtered = all.filter(e => !trackNums.includes(e.trackNum));
        trackNums.forEach((trackNum, i) => {
            filtered.push({
                id: `${sid}_${i}`, trackNum, companyName, reason, firstStatus, dateTime,
                images: i === 0 ? images : [], timestamp: new Date().toISOString(), sessionId: sid,
                category: metadata.category || 'เงินสด', branch: metadata.branch || '',
                reporter: metadata.reporter || '', subject: metadata.subject || '', note: metadata.note || ''
            });
        });
        // Cap to 30 sessions
        const ids = [...new Set(filtered.map(e => e.sessionId || e.id))];
        if (ids.length > 30) {
            const times = ids.map(id => ({ id, time: new Date(filtered.find(e => (e.sessionId||e.id) === id).timestamp).getTime() }));
            times.sort((a,b) => b.time - a.time);
            const keep = times.slice(0, 30).map(t => t.id);
            filtered = filtered.filter(e => keep.includes(e.sessionId || e.id));
        }
        await StorageV2.set(EXCEPTION_KEY, filtered);
        return sid;
    },
    removeSession: async (sid) => {
        const all = await ExceptionManager.getAll();
        await StorageV2.set(EXCEPTION_KEY, all.filter(e => (e.sessionId || e.id) !== sid));
    },
    clearAll: async () => await StorageV2.remove(EXCEPTION_KEY)
};

/**
 * UI INTEGRATION (ASYNC)
 */
async function updateDbViews() {
    const container = document.getElementById('db-list-container');
    if (!container) return;
    const view = typeof currentDbView !== 'undefined' ? currentDbView : 'recent';
    if (view === 'recent') {
        const batches = await CustomerDB.getBatches();
        window.renderRecentBatches(Object.values(batches).sort((a,b)=>b.timestamp - a.timestamp));
    } else if (view === 'company') {
        window.renderCompanySummaries(await CustomerDB.getCompanySummaries());
    } else if (view === 'trash') {
        const trash = await CustomerDB.getTrash();
        window.renderTrashBatches(Object.values(trash).sort((a,b)=>b.deletedAt - a.deletedAt));
    }
}

async function saveCustomerData() {
    const nameInput = document.getElementById('db-name');
    const listInput = document.getElementById('db-tracking-list');
    const name = nameInput ? nameInput.value.trim() : '';
    const type = document.getElementById('db-type') ? document.getElementById('db-type').value : 'Credit';
    const loadedList = listInput ? listInput.value.trim() : '';

    if (!name || !loadedList) { alert('กรุณากรอกชื่อและรายการพัสดุ'); return; }

    const batchInfo = { 
        name, 
        type, 
        contract: document.getElementById('db-contract') ? document.getElementById('db-contract').value.trim() : '', 
        requestDate: document.getElementById('db-request-date') ? document.getElementById('db-request-date').value.trim() : '', 
        timestamp: Date.now() 
    };

    const numbers = [], lines = loadedList.split(/[\r\n,]+/), regex = /([A-Z]{2})(\d{8})(\d)([A-Z]{2})/;

    lines.forEach(line => {
        line = line.trim().toUpperCase(); if (!line) return;
        if (line.includes('-')) {
            const p = line.split('-');
            const mS = p[0].trim().match(regex), mE = p[1].trim().match(regex);
            if (mS && mE) {
                for (let i = parseInt(mS[2]); i <= parseInt(mE[2]); i++) {
                    let b = i.toString().padStart(8, '0');
                    numbers.push(`${mS[1]}${b}${TrackingUtils.calculateS10CheckDigit(b)}${mS[4]}`);
                }
            }
        } else if (line.match(regex)) numbers.push(line);
    });

    if (numbers.length === 0) { alert('ไม่พบเลขพัสดุที่ถูกต้อง'); return; }

    const btn = document.querySelector('button[onclick="saveCustomerData()"]');
    if (btn) window.setButtonLoading(btn, true);

    try {
        const result = await CustomerDB.addBatch(batchInfo, numbers);
        if (result && result.error === 'DUPLICATE') {
            window.showToast(`ข้อมูลชุดนี้มีอยู่แล้วในระบบ (Batch: ${result.id})`, 'info');
        } else {
             window.showToast(`บันทึกเรียบร้อย! เพิ่ม ${numbers.length} รายการ`);
        }

        if(nameInput) nameInput.value = ''; 
        if(listInput) listInput.value = '';
        
        const editSec = document.getElementById('db-edit-section');
        if (editSec) editSec.classList.add('hidden');
        
        currentDbView = 'recent'; 
        if (typeof updateDbViews === 'function') await updateDbViews();
        
        // v1.94: Auto-show in search results (Instant Feedback)
        if (typeof renderUnifiedNumbers === 'function' && numbers.length > 0) {
            const searchItems = numbers.map(num => ({
                number: num,
                isCenter: true,
                offset: 0
            }));
            const title = numbers.length > 1 
                ? `ผลการบันทึก: ${name} (${numbers.length} รายการ)`
                : `ผลการบันทึก: ${numbers[0]}`;
                
            await renderUnifiedNumbers(title, searchItems, false);
        }
    } catch (err) {
        console.error("saveCustomerData Error:", err);
        alert("เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + err.message);
    } finally {
        if (btn) window.setButtonLoading(btn, false, 'บันทึกข้อมูล (Save)');
    }
}

// Note: Initialization is now handled by app.js to ensure UI readiness.
