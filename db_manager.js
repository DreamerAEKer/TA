/**
 * Customer Database Manager
 * Handles local storage operations for Tracking Owner Data.
 */

const DB_KEY = 'thp_tracking_db_v1';

const CustomerDB = {
    // Load all data
    getAll: () => {
        const raw = localStorage.getItem(DB_KEY);
        return raw ? JSON.parse(raw) : {};
    },

    // Save map
    saveAll: (data) => {
        localStorage.setItem(DB_KEY, JSON.stringify(data));
    },

    // Add entries
    addEntries: (trackingList, customerInfo) => {
        const db = CustomerDB.getAll();
        let count = 0;
        trackingList.forEach(id => {
            db[id] = customerInfo;
            count++;
        });
        CustomerDB.saveAll(db);
        return count;
    },

    // Get info for a specific ID
    get: (id) => {
        const db = CustomerDB.getAll();
        return db[id] || null;
    },

    // Clear all
    clear: () => {
        localStorage.removeItem(DB_KEY);
    },

    // Get count
    count: () => {
        return Object.keys(CustomerDB.getAll()).length;
    },

    // Find similar (Same Body+Suffix, Different Prefix)
    findSimilarByBody: (trackingNumber) => {
        const db = CustomerDB.getAll();
        const results = [];

        // Parse input
        const regex = /^([A-Z]{2})(\d{9})([A-Z]{2})$/;
        const match = trackingNumber.toUpperCase().match(regex);
        if (!match) return [];

        const [full, prefix, digits, suffix] = match;
        const targetSearch = digits + suffix; // e.g., 123456785TH

        Object.keys(db).forEach(key => {
            if (key === trackingNumber) return; // Exact match, skip

            const keyMatch = key.match(regex);
            if (keyMatch) {
                const [kFull, kPrefix, kDigits, kSuffix] = keyMatch;
                if (kDigits === digits && kSuffix === suffix && kPrefix !== prefix) {
                    results.push({
                        number: key,
                        info: db[key]
                    });
                }
            }
        });

        return results;
    }
};

// UI State
let currentSort = { key: 'timestamp', order: 'desc' }; // Default: Newest first

// UI Functions for DB Tab
function sortDB(key) {
    if (currentSort.key === key) {
        // Toggle order
        currentSort.order = currentSort.order === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.key = key;
        currentSort.order = 'asc'; // Default new sort is Ascending
    }
    renderDBTable();
}

function renderDBTable() {
    const db = CustomerDB.getAll();
    const tbody = document.querySelector('#db-table tbody');
    const countSpan = document.getElementById('db-count');

    tbody.innerHTML = '';
    const keys = Object.keys(db);
    countSpan.innerText = keys.length;

    // Convert to Array for sorting
    const rows = keys.map(key => ({
        id: key,
        ...db[key]
    }));

    // Sort
    rows.sort((a, b) => {
        let valA, valB;
        if (currentSort.key === 'id') {
            valA = a.id; valB = b.id;
        } else if (currentSort.key === 'name') {
            valA = a.name.toLowerCase(); valB = b.name.toLowerCase();
        } else if (currentSort.key === 'type') {
            valA = a.type; valB = b.type;
        } else if (currentSort.key === 'timestamp') {
            valA = a.timestamp || 0; valB = b.timestamp || 0;
        }

        if (valA < valB) return currentSort.order === 'asc' ? -1 : 1;
        if (valA > valB) return currentSort.order === 'asc' ? 1 : -1;
        return 0;
    });

    // Update Headers (Visual Indicator)
    document.querySelectorAll('#db-table th').forEach(th => {
        th.style.background = '#f8f9fa'; // Reset
        if (th.dataset.sort === currentSort.key) {
            th.style.background = '#e2e6ea';
        }
    });

    // Show top 200 (increased limit)
    const limit = 200;

    rows.slice(0, limit).forEach(item => {
        const dateStr = item.timestamp ? new Date(item.timestamp).toLocaleString('th-TH') : '-';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${item.id}</td>
            <td>${item.name}</td>
            <td>
                <span class="badge ${item.type === 'Credit' ? 'badge-primary' : 'badge-neutral'}">
                    ${item.type}
                </span>
                ${item.contract ? `<br><small>${item.contract}</small>` : ''}
            </td>
            <td style="font-size:0.85rem; color:#666;">${dateStr}</td>
            <td>
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#dc3545; color:white;" onclick="deleteEntry('${item.id}')">Del</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function deleteEntry(key) {
    const db = CustomerDB.getAll();
    delete db[key];
    CustomerDB.saveAll(db);
    renderDBTable();
}

function clearAllData() {
    if (confirm('ยืนยันลบข้อมูลทั้งหมด?')) {
        CustomerDB.clear();
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

    const info = { name, type, contract, timestamp: new Date().getTime() };
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

    const savedCount = CustomerDB.addEntries(numbersToAdd, info);
    alert(`บันทึกเรียบร้อย! เพิ่ม ${savedCount} รายการ`);

    // Reset inputs
    document.getElementById('db-tracking-list').value = '';
    renderDBTable();
}

// Hook into Window Load to Init Table
window.addEventListener('load', renderDBTable);
