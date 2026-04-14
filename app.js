/**
 * Tracking Analysis Tool - App Logic
 */

// Tab Switching
function switchTab(tabId) {
    // Hide all contents
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

    // Show target
    document.getElementById(`tab-${tabId}`).classList.add('active');
    
    // Activate button using specialized ID or fallback
    const btn = document.getElementById(`btn-tab-${tabId}`);
    if (btn) {
        btn.classList.add('active');
    } else {
        // Fallback for any legacy buttons
        const btns = document.getElementsByClassName('tab-btn');
        for (let b of btns) {
            if (b.getAttribute('onclick')?.includes(tabId)) {
                b.classList.add('active');
            }
        }
    }
}

// Global Event Listeners (Run on load)
document.addEventListener('DOMContentLoaded', async () => {
    if (typeof CustomerDB !== 'undefined') {
        try {
            await CustomerDB.init();
            if (window.migrateLegacyData) window.migrateLegacyData();
            await CustomerDB.deduplicate();
            if (typeof updateDbViews === 'function') await updateDbViews();
        } catch(e) { console.error("Database initialization failed:", e); }
    }
    checkAuth();
    
    // Auto-init Exception Report if visible
    if (document.getElementById('exception-table-container')) {
        console.log("DOMContentLoaded: Initializing Exception Report Features...");
        loadExceptionMeta(); 
        await renderExceptionTable();
    }

    // --- Old tools consolidated into Smart Workspace ---
    // checkAdminUI was removed and consolidated into checkAuth
    
    // Initialize Fuel Surcharge Toggle (v2.0-stable)
    const surchargeToggle = document.getElementById('import-surcharge-toggle');
    if (surchargeToggle) {
        const saved = localStorage.getItem('thp_import_surcharge');
        surchargeToggle.checked = (saved === 'true');
    }

    // Initialize Import Batch Type (Remember Last)
    const batchTypeSelect = document.getElementById('import-batch-type');
    if (batchTypeSelect) {
        const lastType = localStorage.getItem('thp_last_batch_type');
        if (lastType) batchTypeSelect.value = lastType;
        
        batchTypeSelect.addEventListener('change', (e) => {
            localStorage.setItem('thp_last_batch_type', e.target.value);
        });
    }

    // Initialize Active Toggle for Surcharge (v4.0.0)
    if (surchargeToggle) {
        surchargeToggle.addEventListener('change', (e) => {
            localStorage.setItem('thp_import_surcharge', e.target.checked);
            // Re-render immediately if we have data
            if (typeof currentImportedBatches !== 'undefined' && currentImportedBatches.length > 0) {
                renderImportResult(currentImportedBatches, lastMissingItems, lastDiscrepanciesList);
            }
        });
    }

    // 4. Smart Workspace Inputs -> Enter Key
    document.getElementById('smart-main-input')?.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') unifiedMainSearch();
    });

    // Real-time Check Digit Validation for Main Input
    document.getElementById('smart-main-input')?.addEventListener('input', (e) => {
        const val = e.target.value.trim().toUpperCase().replace(/\s+/g, '');
        const warning = document.getElementById('smart-main-input-warning');
        if (!warning) return;
        
        if (val.length === 13) {
            const result = TrackingUtils.validateTrackingNumber(val);
            if (!result.isValid) {
                warning.style.display = 'block';
                warning.innerHTML = `⚠️ เลขพัสดุผิด: ${result.error}${result.suggestion ? ' (น่าจะเป็น: ' + result.suggestion + ')' : ''}`;
            } else {
                warning.style.display = 'none';
            }
        } else {
            warning.style.display = 'none';
        }
    });

    // 5. Exception Report Inputs -> DB Warning
    document.getElementById('exception-track-input')?.addEventListener('input', (e) => {
        if (typeof checkDbWarningForReport === 'function') checkDbWarningForReport(e.target);
    });

    // --- Persistence: Load last search results (v1.97: Migrate to StorageV2/IndexedDB) ---
    (async () => {
        try {
            // Check new storage first
            let lastResults = await StorageV2.get('thp_last_search_results');
            let lastTitle = await StorageV2.get('thp_last_search_title');
            let lastIsOcrStr = await StorageV2.get('thp_last_search_is_ocr');
            
            // Fallback to legacy localStorage if new storage is empty
            if (!lastResults) {
                const legacyResults = localStorage.getItem('thp_last_search_results');
                if (legacyResults) {
                    lastResults = JSON.parse(legacyResults);
                    lastTitle = localStorage.getItem('thp_last_search_title');
                    lastIsOcrStr = localStorage.getItem('thp_last_search_is_ocr');
                    
                    // Migrate and Cleanup
                    await saveLastSearchResults(lastTitle, lastResults, lastIsOcrStr === 'true');
                    localStorage.removeItem('thp_last_search_results');
                    localStorage.removeItem('thp_last_search_title');
                    localStorage.removeItem('thp_last_search_is_ocr');
                }
            }

            if (lastResults && lastTitle) {
                currentUnifiedResults = lastResults;
                currentUnifiedTitle = lastTitle;
                renderStoredUnifiedNumbers(currentUnifiedTitle, currentUnifiedResults, lastIsOcrStr === 'true');
            }
        } catch(e) { console.error("Failed to load last results", e); }
    })();
});

/**
 * Real-time warning for Report Draft inputs if number exists in DB
 */
async function checkDbWarningForReport(el) {
    const val = el.value.trim().toUpperCase().replace(/\s+/g, '');
    if (val.length === 13) {
        const owner = typeof CustomerDB !== 'undefined' ? await CustomerDB.get(val) : null;
        if (owner) {
            el.style.backgroundColor = "#ffebee";
            el.style.border = "1px solid #f44336";
            if (typeof showToast === 'function') showToast(`⚠️ เลขนี้มีในฐานข้อมูลแล้ว (${owner.name})`, 'error');
        } else {
            el.style.backgroundColor = "";
            el.style.border = "";
        }
    } else {
        el.style.backgroundColor = "";
        el.style.border = "";
    }
}

// checkAdminUI function was removed and consolidated into checkAuth

// Consolidated into unifiedQuickCheck() and unifiedGenerateRange()

// --- Unified Workspace Logic (Smart Tracking) ---
let lastGeneratedRange = [];
let currentUnifiedResults = []; // Store current results for persistence & "move" logic
let currentUnifiedTitle = "";

function toggleMainRangeInputs() {
    const toggle = document.getElementById('smart-main-range-toggle');
    const options = document.getElementById('smart-main-range-options');
    const bigBtn = document.getElementById('smart-main-range-btn-container');
    
    if (toggle.checked) {
        options.classList.remove('hidden');
        if(bigBtn) bigBtn.style.display = 'block';
    } else {
        options.classList.add('hidden');
        if(bigBtn) bigBtn.style.display = 'none';
        document.getElementById('smart-range-prev').value = '0';
        document.getElementById('smart-range-next').value = '0';
    }
}

async function unifiedMainSearch() {
    const inputEl = document.getElementById('smart-main-input');
    // v1.73: Enhanced sanitization to remove ALL types of spaces and hidden characters
    let input = inputEl.value.trim().toUpperCase().replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
    const isRange = document.getElementById('smart-main-range-toggle')?.checked;
    
    if (!input) {
        alert('กรุณากรอกเลขพัสดุ');
        return;
    }
    
    if (isRange) {
        unifiedGenerateRangeNew();
    } else {
        unifiedSingleCheckNew(input, inputEl);
    }
}

/**
 * SINGLE SEARCH: Check a single number and include neighbors.
 * v1.70: Includes 1 neighbor before and after for context.
 */
async function unifiedSingleCheckNew(input, inputEl) {
    if (inputEl) inputEl.value = input;
    
    const items = TrackingUtils.generateTrackingRange(input, 1, 1);
    await renderUnifiedNumbers(`ค้นหาเลข: ${TrackingUtils.formatTrackingNumber(input)}`, items, false);
}

/**
 * BATCH SEARCH: Generate a range of numbers to check based on UI inputs.
 * v1.70: Includes 1 neighbor before and after.
 */
async function unifiedGenerateRangeNew() {
    const startRaw = (document.getElementById('exception-start-input')?.value || "").trim().toUpperCase().replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, "");
    const endRaw   = (document.getElementById('exception-end-input')?.value || "").trim().toUpperCase().replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, "");

    if (!startRaw || !endRaw) {
        alert('กรุณาระบุเลขเริ่มต้นและสุดท้าย');
        return;
    }

    const startParsed = parseExceptionTrackNum(startRaw);
    const endParsed   = parseExceptionTrackNum(endRaw);

    if (!startParsed || !endParsed) {
        alert('รูปแบบเลขไม่ถูกต้อง');
        return;
    }

    if (startParsed.prefix !== endParsed.prefix || startParsed.suffix !== endParsed.suffix) {
        alert('Prefix/Suffix ไม่ตรงกัน');
        return;
    }

    const startStr = TrackingUtils.formatTrackingNumber(startRaw);
    const endStr = TrackingUtils.formatTrackingNumber(endRaw);

    // GENERATE RANGE with 1 before and 1 after (v1.70)
    const list = [];
    const count = endParsed.bodyInt - startParsed.bodyInt + 1;
    
    if (count > 500 && !confirm(`จะสร้างเลขชุดจำนวน ${count} รายการ ยืนยันหรือไม่?`)) return;

    for (let i = startParsed.bodyInt - 1; i <= endParsed.bodyInt + 1; i++) {
        const isMain = (i >= startParsed.bodyInt && i <= endParsed.bodyInt);
        list.push({
            number: buildExceptionTrackNum(startParsed.prefix, i, startParsed.suffix),
            isCenter: isMain,
            offset: isMain ? 0 : (i < startParsed.bodyInt ? -1 : 1)
        });
    }

    // Capture main items for "Copy All"
    lastGeneratedRange = list.filter(item => item.isCenter).map(item => item.number);
    
    await renderUnifiedNumbers(`ชุดเลขต่อเนื่อง (${startStr} - ${endStr})`, list, false);
}

/**
 * SHARED RENDERER: Renders a list of tracking items into the Unified Results area.
 * v1.70: Unified "In-Order" UI [Before -> Main (Range) -> After]
 */
async function renderUnifiedNumbers(title, items, isOcr = false) {
    const resultArea = document.getElementById('smart-unified-results');
    const summaryArea = document.getElementById('smart-summary-badge');
    
    if (!resultArea) return;
    resultArea.innerHTML = '<div style="padding:20px; text-align:center; color:#888;"><span class="spinner"></span> กำลังวิเคราะห์ข้อมูลและตรวจสอบเจ้าของ...</div>';

    // 1. Fetch Owners (Async) - CRITICAL: CustomerDB.get is async
    const enrichedItems = await Promise.all(items.map(async item => {
        let owner = null;
        if (typeof CustomerDB !== 'undefined') {
            owner = await CustomerDB.get(item.number);
        }
        
        // --- History Check (v1.79) ---
        let historyInfo = null;
        if (typeof ExceptionManager !== 'undefined') {
            const allExceptions = await ExceptionManager.getAll();
            const found = allExceptions.find(e => e.trackNum === item.number.replace(/\s/g,''));
            if (found) {
                historyInfo = { date: found.timestamp, company: found.companyName, reason: found.reason, sessionId: found.sessionId || found.id };
            }
        }

        return {
            number: item.number,
            owner: owner,
            status: item.status || (owner ? owner.remark : ''),
            datetime: item.datetime || '',
            isCenter: item.isCenter || false,
            offset: item.offset || 0,
            source: item.source || '',
            history: historyInfo
        };
    }));

    currentUnifiedResults = enrichedItems;
    currentUnifiedTitle = title;
    
    // Save for F5 persistence (v1.97: Use IndexedDB to avoid QuotaExceededError)
    await saveLastSearchResults(title, enrichedItems, isOcr);

    renderStoredUnifiedNumbers(title, enrichedItems, isOcr);
}

/**
 * v1.97: Helper to save search results to large storage (IndexedDB)
 */
async function saveLastSearchResults(title, items, isOcr) {
    try {
        await StorageV2.set('thp_last_search_results', items);
        await StorageV2.set('thp_last_search_title', title);
        await StorageV2.set('thp_last_search_is_ocr', isOcr ? 'true' : 'false');
    } catch (e) {
        console.error("Failed to save persistence to IndexedDB:", e);
    }
}

/**
 * Helper to get all tracking numbers currently in the exception report form.
 * Returns a Set of cleaned tracking numbers.
 */
function getCurrentDraftNumbers() {
    const numbers = new Set();
    const rangeToggle = document.getElementById('exception-range-toggle');
    
    // 1. Single Mode Textarea
    const mainInput = document.getElementById('exception-track-input');
    if (mainInput && mainInput.value) {
        mainInput.value.split(/[\s,]+/).forEach(n => {
            const clean = n.trim().toUpperCase().replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            if (clean) numbers.add(clean);
        });
    }

    // 2. Range Mode Inputs
    if (rangeToggle && rangeToggle.checked) {
        const start = document.getElementById('exception-start-input')?.value.trim().toUpperCase().replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
        const end = document.getElementById('exception-end-input')?.value.trim().toUpperCase().replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
        if (start && end) {
            const sP = parseExceptionTrackNum(start);
            const eP = parseExceptionTrackNum(end);
            if (sP && eP && sP.prefix === eP.prefix && sP.suffix === eP.suffix) {
                for (let i = sP.bodyInt; i <= eP.bodyInt; i++) {
                    numbers.add(buildExceptionTrackNum(sP.prefix, i, sP.suffix));
                }
            }
        }
    }

    // 3. Extra Items (v1.91 Fix: match class used in generator)
    const extras = document.querySelectorAll('.exception-extra-track');
    extras.forEach(ex => {
        const clean = ex.value.trim().toUpperCase().replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
        if (clean) numbers.add(clean);
    });

    return numbers;
}

/**
 * Renders from state (skips async fetching)
 */
function renderStoredUnifiedNumbers(title, enrichedItems, isOcr = false) {
    const resultArea = document.getElementById('smart-unified-results');
    const summaryArea = document.getElementById('smart-summary-badge');
    if (!resultArea) return;

    // v1.91: Get current drafting items to hide them
    const draftSet = getCurrentDraftNumbers();

    // 2. Group by Company
    const groups = {};
    enrichedItems.forEach(item => {
        const cleanNum = item.number.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
        if (draftSet.has(cleanNum)) return; // Skip items already in draft

        const companyName = item.owner ? item.owner.name : 'ไม่มีในฐานข้อมูล (Unknown Sender)';
        if (!groups[companyName]) groups[companyName] = [];
        groups[companyName].push(item);
    });

    // 3. Update Summary Badge (Count ONLY Main items)
    const mainItemTotal = enrichedItems.filter(i => i.isCenter).length;
    if (summaryArea) {
        summaryArea.innerHTML = `<span class="badge badge-primary">${mainItemTotal} รายการหลัก${isOcr ? ' (OCR)' : ''}</span>`;
    }

    // 4. Build HTML Header
    let html = `
        <div style="padding:10px; background:linear-gradient(to bottom, #fff, #f9f9f9); border-bottom:1px solid #ddd; position:sticky; top:0; z-index:99; display:flex; justify-content:space-between; align-items:center; box-shadow:0 2px 4px rgba(0,0,0,0.05);">
            <div style="font-size:0.8rem; color:#666; font-weight:bold; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; flex:1; display:flex; align-items:center; gap:8px;">
                <input type="checkbox" id="global-master-select-all" style="width:18px; height:18px; cursor:pointer;" onclick="toggleAllSearchCheckboxes(this.checked)" ${isOcr ? '' : 'checked'}>
                <span title="${title}">🔍 ${title} (${mainItemTotal} รายการ)</span>
            </div>
            <button class="btn" style="padding:4px 12px; font-size:0.75rem; background:#fff; border:1px solid #0288d1; color:#0288d1; border-radius:4px; font-weight:bold; flex-shrink:0; margin-left:10px;" onclick="copyAllSearchTrackings()">📋 คัดลอกเลขทั้งหมด</button>
        </div>
        <div style="padding:10px;">
    `;

    // 5. Render Each Company Group
    const sortedCompanies = Object.keys(groups).sort((a, b) => {
        if (a.includes('ไม่มีในฐานข้อมูล')) return 1;
        if (b.includes('ไม่มีในฐานข้อมูล')) return -1;
        return a.localeCompare(b);
    });

    sortedCompanies.forEach((company, companyIdx) => {
        // v1.79 Filter out moved items for this company
        const groupItems = groups[company].filter(i => !i.moved);
        if (groupItems.length === 0) return; // Skip company if all moved

        const companyEscaped = company.replace(/'/g, "\\'").replace('ไม่มีในฐานข้อมูล (Unknown Sender)', '');
        const companyId = `group-${companyIdx}`;

        // Sub-group items into clusters (Consecutive ranges)
        const mainItemsInGroup = groupItems.filter(i => i.isCenter).sort((a, b) => a.number.localeCompare(b.number));
        const seriesList = [];
        let currentSeries = [];

        mainItemsInGroup.forEach((main, mIdx) => {
            const pMain = parseExceptionTrackNum(main.number);
            const prevMain = mIdx > 0 ? parseExceptionTrackNum(mainItemsInGroup[mIdx - 1].number) : null;
            
            const isConsecutive = prevMain && 
                                  pMain && 
                                  pMain.prefix === prevMain.prefix && 
                                  pMain.suffix === prevMain.suffix && 
                                  pMain.bodyInt === prevMain.bodyInt + 1;

            if (mIdx > 0 && !isConsecutive) {
                seriesList.push(currentSeries);
                currentSeries = [];
            }
            
            // Satellites for this specific main item
            const satellites = groupItems.filter(i => {
                if (i.isCenter) return false;
                const pSat = parseExceptionTrackNum(i.number);
                if (!pSat || !pMain) return false;
                return pSat.prefix === pMain.prefix && pSat.suffix === pMain.suffix && Math.abs(pSat.bodyInt - pMain.bodyInt) <= 1;
            });

            currentSeries.push({ main, satellites });
        });
        if (currentSeries.length > 0) seriesList.push(currentSeries);

        // Render the Company Header
        const mainInCompanyCount = groupItems.filter(i => i.isCenter).length;
        html += `
            <div style="margin-bottom:20px; border:1px solid #e0e0e0; border-radius:12px; overflow:hidden; background:#fff; box-shadow:0 2px 6px rgba(0,0,0,0.06);">
                <div style="background:#f0f7ff; padding:10px 15px; border-bottom:1px solid #e1f5fe; font-weight:bold; color:#0277bd; display:flex; justify-content:space-between; align-items:center; font-size:0.95rem; gap:10px;">
                    <div style="display:flex; align-items:center; gap:8px; flex:1; min-width:0;">
                        <input type="checkbox" id="master-${companyId}" style="width:20px; height:20px; cursor:pointer;" onclick="toggleGroupCheckboxes('${companyId}', this.checked)" ${isOcr ? '' : 'checked'}>
                        <div style="flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${company}">
                            <span>🏢 ${company}</span>
                        </div>
                    </div>
                    <span style="font-size:0.8rem; font-weight:normal; color:#888;">หลัก: ${mainInCompanyCount}</span>
                </div>
                <div style="padding:12px; background:#fafafa;">
        `;

        // Render each cluster
        seriesList.forEach((series, sIdx) => {
            const clusterId = `${companyId}-s${sIdx}`;
            const firstMain = series[0].main;
            const lastMain  = series[series.length - 1].main;
            const isSingle  = series.length === 1;

            const startStr = TrackingUtils.formatTrackingNumber(firstMain.number);
            const endStr = TrackingUtils.formatTrackingNumber(lastMain.number);
            const rangeTitle = isSingle ? startStr : `${startStr} ถึง ${endStr}`;
            
            // Collect all satellites for this series
            const allSats = [];
            const seenSats = new Set();
            
            let clusterHasHistory = false;
            let reportedMainCount = 0;

            series.forEach(wrap => {
                if (wrap.main.history) {
                    clusterHasHistory = true;
                    reportedMainCount++;
                }
                wrap.satellites.forEach(s => {
                    if (s.history) clusterHasHistory = true;
                    if (!seenSats.has(s.number)) {
                        seenSats.add(s.number);
                        allSats.push(s);
                    }
                });
            });

            const allReported = series.length > 0 && reportedMainCount === series.length;

            const summaryRowData = { 
                number: rangeTitle, 
                isCenter: false, // Summary row NOT selectable for counting
                offset: 0, 
                hasHistory: clusterHasHistory,
                allReported: allReported,
                status: isSingle ? (firstMain.status || '') : `กลุ่มเลขต่อเนื่อง (${series.length} ชุด)` 
            };

            const mainNumbersInGroup = new Set(series.map(wrap => wrap.main.number));

            const leadingSats = allSats.filter(s => {
                const pS = parseExceptionTrackNum(s.number);
                const pF = parseExceptionTrackNum(firstMain.number);
                if (mainNumbersInGroup.has(s.number)) return false; 
                return pS && pF && pS.bodyInt < pF.bodyInt;
            }).sort((a,b) => a.number.localeCompare(b.number));
            
            const trailingSats = allSats.filter(s => {
                const pS = parseExceptionTrackNum(s.number);
                const pL = parseExceptionTrackNum(lastMain.number);
                if (mainNumbersInGroup.has(s.number)) return false;
                return pS && pL && pS.bodyInt > pL.bodyInt;
            }).sort((a,b) => a.number.localeCompare(b.number));

            const hasAnySats = leadingSats.length > 0 || trailingSats.length > 0 || !isSingle;

            const shouldCheckDefault = !isOcr;

            html += `
                <div style="border:1px solid #90caf9; border-radius:10px; overflow:hidden; background:white; margin-bottom:12px; box-shadow:0 3px 10px rgba(2,136,209,0.05);">
                    ${renderUnifiedRow(summaryRowData, companyId, companyEscaped, hasAnySats, clusterId, shouldCheckDefault, draftSet)}
                    <div class="satellite-wrapper">
                        <!-- 1. Before Satellite -->
                        ${leadingSats.map(s => renderUnifiedRow(s, companyId, companyEscaped, false, clusterId, shouldCheckDefault, draftSet)).join('')}

                        <!-- 2. Main List -->
                        ${series.map(wrap => renderUnifiedRow(wrap.main, companyId, companyEscaped, false, clusterId, shouldCheckDefault, draftSet)).join('')}

                        <!-- 3. After Satellite -->
                        ${trailingSats.map(s => renderUnifiedRow(s, companyId, companyEscaped, false, clusterId, shouldCheckDefault, draftSet)).join('')}
                    </div>
                </div>
            `;
        });

        html += `</div></div>`;
    });

    html += `</div>`;
    resultArea.innerHTML = html;

    // Update Copy Bar
    const copyBar = document.getElementById('smart-copy-all-bar');
    if (copyBar) {
        copyBar.innerHTML = `
            <div style="display:flex; flex-direction:column; gap:8px;">
                <button class="btn btn-success" onclick="stagingAllCheckedItems()" style="width:100%; font-size:1.16rem; padding:15px; font-weight:bold; border-radius:10px; box-shadow:0 4px 15px rgba(46,125,50,0.2);">🚩 เพิ่มรายการที่เลือกเข้าตารางร่าง (<span id="search-selection-count">${mainItemTotal}</span>)</button>
                <div style="display:flex; gap:8px;">
                    <button class="btn btn-neutral" onclick="toggleAllSearchCheckboxes(true)" style="flex:1; font-size:0.8rem; padding:8px; background:#fff; border:1px solid #ccc; font-weight:bold;">✅ เลือกทั้งหมด</button>
                    <button class="btn btn-neutral" onclick="toggleAllSearchCheckboxes(false)" style="flex:1; font-size:0.8rem; padding:8px; background:#fff; border:1px solid #ccc;">❌ ยกเลิก</button>
                    <button class="btn btn-neutral" onclick="copySelectedUnifiedNumbers()" style="flex:1.2; font-size:0.8rem; padding:8px; background:#fff; border:1px solid #ccc;">📋 คัดลอกที่เลือก</button>
                </div>
            </div>
        `;
        copyBar.classList.remove('hidden');
    }
}

/**
 * Helper to render a single row in the Results Sidebar.
 * v1.91: Improved history feedback and jump-to-report action.
 */
function renderUnifiedRow(row, groupId, companyEscaped, hasSatellites = false, clusterId = null, shouldAutoCheck = true, draftSet = new Set()) {
    const isMain = row.isCenter === true;
    const rowClass = isMain ? 'unified-row center-row' : 'unified-row satellite-row';
    const rowBg = isMain ? '#fff9c4' : '#fff';
    const opacity = isMain ? '1' : '0.75';
    
    // v1.73/v1.74/v1.75/v1.92/v1.93 Fix: Only sanitize and re-format if it's a single tracking number (approx 13-15 chars)
    // If it's a range summary title (longer, contains "ถึง"), keep it as is for the UI.
    const cleanNum = row.number.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
    const isRangeTitle = row.number.includes(' ถึง ');
    // v1.93: Be more aggressive with cleanup—if it's 13 chars, it's a single track regardless of satellites (▶ icon)
    const isSingleTrack = (cleanNum.length === 13) || (!isRangeTitle && cleanNum.length <= 15);
    
    // For actions, if it's a range title, don't strip spaces yet so stagingQuickReport can parse it
    const rawNum = isSingleTrack ? cleanNum : row.number;
    const formattedNum = isSingleTrack ? TrackingUtils.formatTrackingNumber(rawNum) : row.number;

    // Metadata encoding (v1.73: use rawNum)
    const metadataObj = { status: row.status || '', datetime: row.datetime || '' };
    const metadataJson = JSON.stringify(metadataObj).replace(/"/g, '&quot;');
    
    const trackColor = isMain ? '#333' : '#777';
    const toggleAction = hasSatellites ? `onclick="toggleSatelliteGroup(this)"` : '';

    const warningIcon = row.hasHistory || row.history ? '<span class="badge badge-danger" style="margin-left:5px; padding:2px 4px; font-size:0.7rem; cursor:help;" title="เคยแจ้งรายงานแล้ว">⚠️ แจ้งแล้ว</span>' : '';
    const hideCheckbox = (isMain && row.history) || (hasSatellites && row.allReported);

    // v1.91: Instant Filter Sync
    if (isSingleTrack && draftSet.has(rawNum)) return ''; 

    const checkedAttr = shouldAutoCheck ? 'checked' : '';

    return `
        <div class="${rowClass}" ${toggleAction} style="background:${rowBg}; opacity:${opacity}; display:flex; align-items:center; border-bottom:1px solid #f2f2f2; font-size:0.88rem; min-height:48px; cursor:${hasSatellites ? 'pointer' : 'default'};">
            <div style="width:35px; text-align:center; padding:8px 0 8px 8px;">
                ${hideCheckbox ? warningIcon : (isMain ? `
                    <input type="checkbox" class="group-checkbox-${groupId} cluster-checkbox-${clusterId}" value="${rawNum}" 
                        data-metadata="${metadataJson}"
                        style="width:18px; height:18px; cursor:pointer;" onclick="event.stopPropagation(); updateSearchSelectionCount()" ${checkedAttr}>
                ` : (hasSatellites && clusterId ? `
                    <input type="checkbox" class="cluster-master-${clusterId}" 
                        style="width:18px; height:18px; cursor:pointer;" 
                        onclick="event.stopPropagation(); toggleClusterCheckboxes('${clusterId}', this.checked)" ${checkedAttr}>
                ` : ''))}
            </div>
            <div style="width:30px; text-align:center; font-weight:bold; font-size:0.75rem; color:#bbb;">
                ${hasSatellites ? '<span class="toggle-icon">▶</span>' : ''}
                ${isMain ? '<span style="color:#2e7d32; font-size:1.2rem;">•</span>' : '<span style="color:#aaa; font-size:1rem;">•</span>'}
            </div>
            <div style="width:35px; text-align:center;">
                <button class="btn" style="padding:2px 5px; font-size:1.1rem; border:none; background:none; cursor:pointer;" title="รายงานรายการนี้" onclick="event.stopPropagation(); stagingQuickReport(['${rawNum}'], '${companyEscaped}', ${metadataJson})">🚩</button>
            </div>
            <div style="flex:1; font-family:monospace; font-weight:bold; color:${trackColor}; padding:8px 5px;">
                ${formattedNum} 
                ${!isMain && !hasSatellites ? '<span style="font-size:0.65rem; color:#999; font-weight:normal; margin-left:4px;">(เลขแวดล้อม)</span>' : ''}
                ${(hasSatellites && row.hasHistory) ? '<span style="font-size:0.75rem; color:#d32f2f; margin-left:5px; font-weight:bold;">(แจ้งรายงานแล้ว)</span>' : ''}
                ${row.history ? `
                    <div class="reported-badge" style="margin-top:4px; font-weight:normal; border-top:1px dashed #ffcdd2; padding-top:4px;">
                        <span style="color:#d32f2f;">⚠️ เมื่อ ${new Date(row.history.date).toLocaleDateString('th-TH')} - ${row.history.company}</span>
                        <button class="btn btn-neutral" style="padding:1px 6px; font-size:0.7rem; margin-left:5px; background:#fff; border:1px solid #d32f2f; color:#d32f2f; transform:translateY(-1px);" onclick="event.stopPropagation(); viewReportSession('${row.history.sessionId}')">👁️ ดูรายงาน</button>
                    </div>` : ''}
                ${row.status || row.datetime ? `
                    <div style="font-size:0.72rem; color:#0288d1; margin-top:1px; font-style:italic;">
                        ${row.status ? `[${row.status}] ` : ''}${row.datetime || ''}
                    </div>
                ` : ''}
            </div>
            <div style="padding-right:10px; display:flex; gap:4px;" onclick="event.stopPropagation()">
                <button class="btn btn-neutral" style="padding:1px 5px; font-size:0.65rem; background:#fff; border:1px solid #ddd; color:#999;" title="คัดลอกเลข" onclick="navigator.clipboard.writeText('${rawNum}').then(()=>alert('คัดลอก ${rawNum} สำเร็จ!'))">📋</button>
                <button class="btn btn-trace" style="padding:1px 5px; font-size:0.65rem;" title="เช็คสถานะ" onclick="window.open('https://track.thailandpost.co.th/?trackNumber=${rawNum}&lang=th', '_blank')">🔍</button>
            </div>
        </div>
    `;
}

/**
 * Toggle the expansion of a satellite group.
 * v1.70: Animation-friendly toggle using CSS classes.
 */
function toggleSatelliteGroup(el) {
    const parent = el.parentElement;
    if (!parent) return;
    
    const wrapper = parent.querySelector('.satellite-wrapper');
    const icon = el.querySelector('.toggle-icon');
    
    if (wrapper) {
        const isExpanded = wrapper.classList.toggle('expanded');
        if (icon) {
            icon.style.transform = isExpanded ? 'rotate(90deg)' : 'rotate(0deg)';
        }
        
        // Add a visual indicator to the trigger row
        if (isExpanded) {
            el.classList.add('expanded-trigger');
        } else {
            el.classList.remove('expanded-trigger');
        }
    }
}

/**
 * Toggle all checkboxes in a group.
 */
function toggleGroupCheckboxes(groupId, checked) {
    const selector = `.group-checkbox-${groupId}`;
    const cbks = document.querySelectorAll(selector);
    cbks.forEach(cb => cb.checked = checked);
    
    // Also toggle all cluster master checkboxes in this group
    const clusterMasters = document.querySelectorAll(`input[class^="cluster-master-${groupId}"]`);
    clusterMasters.forEach(cb => cb.checked = checked);
    updateSearchSelectionCount();
}

/**
 * Toggle all checkboxes in a specific cluster/series.
 */
function toggleClusterCheckboxes(clusterId, checked) {
    const selector = `.cluster-checkbox-${clusterId}`;
    const cbks = document.querySelectorAll(selector);
    cbks.forEach(cb => cb.checked = checked);
    updateSearchSelectionCount();
}

/**
 * Toggle all search result checkboxes at once (Global Select All)
 */
function toggleAllSearchCheckboxes(checked) {
    // 1. All individual item checkboxes
    const itemCbks = document.querySelectorAll('[class^="group-checkbox-"]');
    itemCbks.forEach(cb => cb.checked = checked);

    // 2. All company group master checkboxes (starts with master-group-)
    const groupMasters = document.querySelectorAll('input[id^="master-group-"]');
    groupMasters.forEach(cb => cb.checked = checked);

    // 3. All cluster master checkboxes
    const clusterMasters = document.querySelectorAll('input[class^="cluster-master-"]');
    clusterMasters.forEach(cb => cb.checked = checked);

    // 4. Update the global header checkbox if it exists
    const globalHeaderCb = document.getElementById('global-master-select-all');
    if (globalHeaderCb) globalHeaderCb.checked = checked;

    updateSearchSelectionCount();
}

/**
 * Update the UI count for selected search items
 */
function updateSearchSelectionCount() {
    const cbks = document.querySelectorAll('[class^="group-checkbox-"]');
    let count = 0;
    cbks.forEach(cb => { if(cb.checked) count++; });
    
    const countSpan = document.getElementById('search-selection-count');
    if (countSpan) countSpan.textContent = count;
}

/**
 * Toggle all checkboxes in a specific cluster/series.
 */
function toggleClusterCheckboxes(clusterId, checked) {
    const selector = `.cluster-checkbox-${clusterId}`;
    const cbks = document.querySelectorAll(selector);
    cbks.forEach(cb => cb.checked = checked);
}

/**
 * Collect all checked items in a group and add to report.
 */
function stagingQuickReportFromGroup(groupId, companyName) {
    const selector = `.group-checkbox-${groupId}`;
    const cbks = document.querySelectorAll(selector);
    const selectedTracks = [];
    let firstMetadata = null;

    cbks.forEach(cb => {
        if (cb.checked) {
            selectedTracks.push(cb.value);
            // Try to capture metadata from the first valid checked item
            if (!firstMetadata) {
                try {
                    const metaStr = cb.getAttribute('data-metadata');
                    if (metaStr) {
                        const decoded = metaStr.replace(/&quot;/g, '"');
                        firstMetadata = JSON.parse(decoded);
                    }
                } catch(e) {
                    console.error("DEBUG: Metadata parse error", e);
                }
            }
        }
    });

    if (selectedTracks.length === 0) {
        if (typeof showToast === 'function') showToast('กรุณาเลือกอย่างน้อย 1 รายการ', 'error');
        return;
    }

    stagingQuickReport(selectedTracks, companyName, firstMetadata || {});
}

/**
 * Collect ALL checked items from ALL groups and add them to the report draft.
 */
function stagingAllCheckedItems() {
    const cbks = document.querySelectorAll('[class^="group-checkbox-"]');
    const selectedTracks = [];
    let firstMetadata = null;

    cbks.forEach(cb => {
        if (cb.checked) {
            // v1.73: Clean all possible space characters
            const raw = cb.value.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            selectedTracks.push(raw);
            if (!firstMetadata) {
                try {
                    const metaStr = cb.getAttribute('data-metadata');
                    if (metaStr) firstMetadata = JSON.parse(metaStr.replace(/&quot;/g, '"'));
                } catch(e) {}
            }
        }
    });

    if (selectedTracks.length === 0) {
        if (typeof showToast === 'function') showToast('กรุณาติ๊กเลือกรายการอย่างน้อย 1 รายการครับ', 'error');
        else alert('กรุณาเลือกรายการอย่างน้อย 1 รายการ');
        return;
    }

    stagingQuickReport(selectedTracks, "", firstMetadata || {});
}

/**
 * Check if a list of tracking numbers is a perfectly consecutive range.
 */
function isConsecutive(nums) {
    if (nums.length < 2) return false;
    const parsed = nums.map(n => parseExceptionTrackNum(n)).filter(p => !!p);
    if (parsed.length !== nums.length) return false;
    
    // Sort by numerical part
    parsed.sort((a,b) => a.bodyInt - b.bodyInt);
    
    const prefix = parsed[0].prefix;
    const suffix = parsed[0].suffix;
    for (let i = 1; i < parsed.length; i++) {
        if (parsed[i].prefix !== prefix || parsed[i].suffix !== suffix) return false;
        if (parsed[i].bodyInt !== parsed[i-1].bodyInt + 1) return false;
    }
    return true;
}

/**
 * Copy all main numbers from current result (regardless of checkboxes)
 */
function copyAllSearchTrackings() {
    if (!currentUnifiedResults) return;
    const allTracks = currentUnifiedResults
        .filter(i => i.isCenter)
        .map(i => i.number.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, ''));

    if (allTracks.length === 0) {
        alert('ไม่พบเลขที่ต้องการคัดลอก');
        return;
    }

    navigator.clipboard.writeText(allTracks.join('\n')).then(() => {
        window.showToast(`คัดลอกทั้งหมด ${allTracks.length} รายการแล้ว`, 'success');
    });
}

/**
 * v1.73/v1.81: Copy only selected numbers as clean raw strings (no spaces)
 */
function copySelectedUnifiedNumbers() {
    const cbks = document.querySelectorAll('[class^="group-checkbox-"]');
    const selectedTracks = [];

    cbks.forEach(cb => {
        if (cb.checked) {
            selectedTracks.push(cb.value.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, ''));
        }
    });

    if (selectedTracks.length === 0) {
        alert('กรุณาเลือกรายการที่ต้องการคัดลอก');
        return;
    }

    navigator.clipboard.writeText(selectedTracks.join('\n')).then(() => {
        window.showToast(`คัดลอกเฉพาะที่เลือก ${selectedTracks.length} รายการแล้ว`, 'success');
    });
}

// Consolidated File Upload handler (Excel & OCR) inside the Track & Trace section
async function handleTrackFileUpload(files) {
    if (!files || files.length === 0) return;
    const statusEl = document.getElementById('track-ocr-status');
    if (statusEl) statusEl.textContent = `กำลังประมวลผล ${files.length} ไฟล์...`;

    // Partition files into Images (OCR) and Excel
    const images = [];
    const excelFiles = [];
    for (const f of files) {
        if (f.type.includes('image')) images.push(f);
        else if (f.name.endsWith('.xlsx') || f.name.endsWith('.xls')) excelFiles.push(f);
    }

    let allItems = []; // List of { number, status, datetime, source }

    try {
        // 1. Process Images via OCR
        if (images.length > 0) {
            if (statusEl) statusEl.textContent = `กำลังสแกน OCR ${images.length} รูป...`;
            let combinedOcrText = '';
            for (let i = 0; i < images.length; i++) {
                if (statusEl) statusEl.textContent = `OCR รูป ${i + 1}/${images.length}...`;
                const worker = await Tesseract.createWorker('tha+eng');
                const { data: { text } } = await worker.recognize(images[i]);
                await worker.terminate();
                combinedOcrText += '\n' + text;
            }
            const ocrNumbers = TrackingUtils.extractTrackingNumbers(combinedOcrText);
            ocrNumbers.forEach(num => {
                allItems.push({ number: num, source: 'OCR', status: '', datetime: '' });
            });
        }

        // 2. Process Excel Files
        if (excelFiles.length > 0) {
            if (statusEl) statusEl.textContent = `กำลังอ่านโหมด Excel ${excelFiles.length} ไฟล์...`;
            for (const file of excelFiles) {
                const excelItems = await processTrackExcelFile(file);
                allItems.push(...excelItems);
            }
        }

        if (allItems.length === 0) {
            if (statusEl) statusEl.textContent = '⚠️ ไม่พบข้อมูลเลขพัสดุ';
            return;
        }

        // 3. Render Results
        if (allItems.length === 1) {
            const item = allItems[0];
            document.getElementById('smart-main-input').value = item.number;
            if (statusEl) statusEl.textContent = `✅ พบเลข: ${item.number} กำลังค้นหา...`;
            unifiedMainSearch();
        } else {
            const expandedItems = [];
            for (const item of allItems) {
                const owner = typeof CustomerDB !== 'undefined' ? await CustomerDB.get(item.number) : null;
                if (!owner) {
                    // Expand Unknown to include neighbors (1 before, 1 after)
                    const range = TrackingUtils.generateTrackingRange(item.number, 1, 1);
                    range.forEach(r => {
                        if (r.isCenter) {
                            r.status = item.status;
                            r.datetime = item.datetime;
                            r.source = item.source;
                        }
                        expandedItems.push(r);
                    });
                } else {
                    // Keep Known as single item
                    expandedItems.push({ 
                        number: item.number, 
                        owner, 
                        offset: 0, 
                        isCenter: true,
                        status: item.status,
                        datetime: item.datetime,
                        source: item.source
                    });
                }
            }
            lastGeneratedRange = expandedItems.filter(i => i.isCenter).map(item => item.number);
            await renderUnifiedNumbers(`ข้อมูลจาก ${files.length} ไฟล์`, expandedItems, true);
        }

        if (statusEl) statusEl.textContent = `✅ สำเร็จ พบข้อมูล ${allItems.length} รายการ`;
        document.getElementById('track-ocr-upload').value = '';

    } catch (err) {
        console.error(err);
        if (statusEl) statusEl.textContent = '❌ เกิดข้อผิดพลาด: ' + err.message;
    }
}

/**
 * Specifically parses Excel for the Track & Trace search context.
 */
async function processTrackExcelFile(file) {

    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                const results = [];
                const trackingRegex = /[A-Z]{2}\d{9}[A-Z]{2}/i;

                jsonData.forEach((row, idx) => {
                    // Skip header or short rows
                    // User format: Col H is index 7
                    if (row.length < 8) return;

                    const rawTrack = row[7]; // Col H
                    if (!rawTrack) return;

                    const match = rawTrack.toString().match(trackingRegex);
                    if (match) {
                        const status = row[8] || ''; // Col I
                        const rawTime = row[9] || ''; // Col J
                        
                        // User's Excel DateTime is often a string or a legacy date
                        let formattedTime = rawTime;
                        if (typeof rawTime === 'number') {
                            // XLSX serial date conversion
                            const dateObj = new Date(Math.round((rawTime - 25569) * 86400 * 1000));
                            // Since we need BE display or at least a recognizable string
                            // If it's serial, it's relative to 1900.
                            // Better to use current local format
                            formattedTime = dateObj.toLocaleDateString('th-TH') + ' ' + dateObj.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
                        }

                        results.push({
                            number: match[0].toUpperCase(),
                            status: status.toString().trim(),
                            datetime: formattedTime.toString().trim(),
                            source: 'Excel'
                        });
                    }
                });
                resolve(results);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Gap analysis features removed per user request
// --- Universal Import Logic (Excel & Image/OCR) ---

let currentImportedBatches = []; // To store analyzed data before saving
let lastMissingItems = []; // v4.0.0 Global state for re-rendering
let lastDiscrepanciesList = []; // v4.0.0 Global state for re-rendering
let rawTrackingData = []; // Store ALL raw items (Cumulative)
let importedFileCount = 0; // Track number of files uploaded (for limit)

function clearImportData() {
    if (rawTrackingData.length === 0) return;
    if (!confirm('คุณแน่ใจหรือไม่ว่าต้องการล้างข้อมูลนำเข้าทั้งหมด? (ข้อมูลที่แสดงอยู่จะไม่ถูกบันทึก)')) return;

    rawTrackingData = [];
    currentImportedBatches = [];
    importedFileCount = 0; // Reset Limit
    document.getElementById('import-preview').classList.add('hidden');
    document.getElementById('upload-status').innerText = 'ล้างข้อมูลเรียบร้อย (Ready)';
    // Reset file input
    document.getElementById('import-upload').value = '';
}


function handleFileUpload(event) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    // --- Role-Based Limit Check ---
    const urlParams = new URLSearchParams(window.location.search);
    const isAdmin = urlParams.has('admin');

    if (!isAdmin) {
        // Normal User Limit: 2 Files Max (Cumulative)
        if ((importedFileCount + files.length) > 2) {
            alert(`⚠️ จำกัดการอัปโหลดสูงสุด 2 ไฟล์สำหรับบัญชีทั่วไป\n(คุณอัปโหลดไปแล้ว ${importedFileCount} ไฟล์, พยายามเพิ่มอีก ${files.length} ไฟล์)\n\nกรุณาล้างข้อมูลเก่าก่อนหากต้องการเริ่มใหม่`);
            document.getElementById('import-upload').value = ''; // Reset input to allow re-selection
            return;
        }
    }

    // Increment Count
    importedFileCount += files.length;

    // Check type
    if (files[0].type.includes('image')) {
        handleImageImport(files); // Pass FileList (Admin Only check is inside)
    } else {
        // Excel - Support Multiple
        Array.from(files).forEach(file => {
            handleExcelImport(file);
        });
    }
}

function handleExcelImport(file) {
    document.getElementById('upload-status').innerText = `กำลังอ่านไฟล์ Excel...`;
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const newItems = [];
        const regex = /([A-Z]{2})(\d{9})([A-Z]{2})/i;

        jsonData.forEach(row => {
            if (row.length >= 3) {
                if (row[2] && typeof row[2] === 'string') {
                    const match = row[2].match(regex);
                    if (match) {
                        const surchargeChecked = document.getElementById('import-surcharge-toggle')?.checked || false;
                        let price = parseFloat(row[3]) || 0;
                        if (surchargeChecked && price > 0) {
                            price += 3; // +3 Baht Fuel Surcharge
                        }
                        let weight = row[4] || '-';
                        let originalWeight = weight;
                        let hasDiscrepancy = false;

                        // Apply Package A3 logic if price matches
                        if (price > 0 && typeof TrackingUtils !== 'undefined' && TrackingUtils.getWeightFromPriceA3) {
                            const calcWeight = TrackingUtils.getWeightFromPriceA3(price);
                            if (calcWeight !== '-') {
                                weight = calcWeight;
                                if (originalWeight !== '-' && originalWeight.toString().trim() !== weight.toString().trim()) {
                                    hasDiscrepancy = true;
                                }
                            }
                        }

                        newItems.push({
                            number: match[0].toUpperCase(),
                            price: price,
                            weight: weight,
                            hasDiscrepancy: hasDiscrepancy,
                            originalWeight: originalWeight
                        });
                    }
                }
            }
        });

        if (newItems.length === 0) {
            alert('ไม่พบเลขพัสดุในไฟล์ Excel (ตรวจสอบคอลัมน์ C)');
            document.getElementById('upload-status').innerText = 'ไม่พบเลขพัสดุ';
            return;
        }

        // Cumulative Append
        rawTrackingData.push(...newItems);

        document.getElementById('upload-status').innerText = `อ่าน Excel สำเร็จ! เพิ่ม ${newItems.length} รายการ (รวม ${rawTrackingData.length})`;
        analyzeImportedRanges(rawTrackingData);
    };
    reader.readAsArrayBuffer(file);
}

// Updated to handle Multiple Images
async function handleImageImport(files) {
    // Check if Admin
    const urlParams = new URLSearchParams(window.location.search);
    if (!urlParams.has('admin')) {
        alert('ฟีเจอร์ "นำเข้ารูปภาพ" สงวนสิทธิ์สำหรับ Admin เท่านั้น\n(พนักงานทั่วไปกรุณาใช้ไฟล์ Excel)');
        // Clear input safely
        const input = document.getElementById('import-upload');
        if (input) input.value = '';
        document.getElementById('upload-status').innerText = 'Access Denied (Admin Only)';
        return;
    }

    const statusEl = document.getElementById('upload-status');
    statusEl.innerText = `Preparing OCR for ${files.length} images...`;

    let combinedText = "";

    try {
        for (let i = 0; i < files.length; i++) {
            statusEl.innerText = `OCR Scanning Image ${i + 1}/${files.length}...`;
            const file = files[i];

            const useEngOnly = document.getElementById('admin-ocr-eng-only') && document.getElementById('admin-ocr-eng-only').checked;
            const lang = useEngOnly ? 'eng' : 'tha+eng';
            const worker = await Tesseract.createWorker(lang);
            const { data: { text } } = await worker.recognize(file);
            await worker.terminate();

            combinedText += "\n" + text;
        }

        statusEl.innerText = "Analyzing extracted text...";
        console.log('OCR Raw:', combinedText);

        const newItems = [];
        const rawLines = combinedText.split('\n').map(l => l.trim()).filter(l => l);

        let currentPrefixBody = null;

        for (let i = 0; i < rawLines.length; i++) {
            let line = rawLines[i].toUpperCase();

            // Strategy 1: Remove all spaces and check for Standard ID
            const cleanLine = line.replace(/\s+/g, '');
            const stdMatch = cleanLine.match(/([A-Z]{2})(\d{9})([A-Z]{2})/);
            if (stdMatch) {
                newItems.push({ number: stdMatch[0], price: 0, weight: 'OCR-Std' });
                currentPrefixBody = stdMatch[1] + stdMatch[2].substring(0, 4);
                continue;
            }
        }

        // --- NEW: Use Smart Extraction (Robust) ---
        // First try the new handwritten table parser
        const tableItems = TrackingUtils.extractHandwrittenTable(combinedText);
        
        // If the table extraction found prices, use that. Otherwise fallback to basic list mapping
        const hasPrices = tableItems.some(item => item.price > 0);
        
        if (hasPrices) {
            newItems.push(...tableItems);
        } else {
             // Fallback to simple extraction
             const extractedSmart = TrackingUtils.extractTrackingNumbers(combinedText);
             const contextData = TrackingUtils.extractTrackingWithContext(combinedText);
             const contextMap = new Map();
             contextData.forEach(c => contextMap.set(c.trackingNumber, c));
             
             extractedSmart.forEach(num => {
                 const context = contextMap.get(num) || null;
                 newItems.push({ number: num, price: 0, weight: 'OCR-Smart', context: context });
             });
        }

        // De-duplicate locally (within this batch) logic if needed, 
        // but analyzeImportedRanges handles grouping.

        if (newItems.length === 0) {
            // Price Fallback...
            const prices = TrackingUtils.extractPrices(combinedText);
            if (prices.length > 0) {
                if (confirm(`ไม่พบเลขพัสดุในภาพ แต่พบยอดเงิน ${prices.length} รายการ\nต้องการสร้างรายงานแยกตามราคาสินค้าแทนหรือไม่?`)) {
                    const summary = TrackingUtils.summarizePrices(prices);
                    const priceRanges = summary.groupings.map(g => ({
                        start: 'PRICE-ONLY',
                        end: 'PRICE-ONLY',
                        count: g.count,
                        price: g.price,
                        weight: 'N/A',
                        total: g.total,
                        items: [] // No tracking IDs
                    }));

                    // Note: Price-only mode replaces everything as it's a different mode
                    rawTrackingData = [];
                    currentImportedBatches = priceRanges;
                    renderImportResult(priceRanges);
                    statusEl.innerText = `วิเคราะห์ยอดเงินสำเร็จ (${summary.totalCount} รายการ)`;
                    return;
                }
            }

            alert('ไม่พบเลขพัสดุในรูปภาพนี้');
            statusEl.innerText = 'ไม่พบข้อมูลในภาพล่าสุด';
            return;
        }

        // Cumulative Append
        rawTrackingData.push(...newItems);

        statusEl.innerText = `OCR เสร็จสิ้น! เพิ่ม ${newItems.length} รายการ (รวม ${rawTrackingData.length})`;
        analyzeImportedRanges(rawTrackingData);

    } catch (err) {
        console.error(err);
        alert('เกิดข้อผิดพลาดในการอ่านรูปภาพ: ' + err.message);
        statusEl.innerText = "Error";
    }
}

function analyzeImportedRanges(trackingList) {
    if (trackingList.length === 0) return;

    // Remove duplicates from total list
    const uniqueMap = new Map();
    const discrepancyMap = new Map();
    trackingList.forEach(item => {
        uniqueMap.set(item.number, item);
        if (item.hasDiscrepancy) {
            discrepancyMap.set(item.number, item);
        }
    });
    const uniqueList = Array.from(uniqueMap.values());
    const discrepanciesList = Array.from(discrepancyMap.values());

    // Sort logic...
    // --- STEP 1: GAP ANALYSIS (100% Integrity Check) ---
    // Sort by Number only (Independent of price) to find true gaps
    const sortedForGaps = [...uniqueList].sort((a, b) => a.number.localeCompare(b.number));
    const missingItems = []; // { start, end, count, example }

    const parse = (str) => {
        const m = str.match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
        return m ? { full: str, prefix: m[1], body: parseInt(m[2]), check: m[3], suffix: m[4] } : null;
    };

    let prevGap = parse(sortedForGaps[0].number);

    for (let i = 1; i < sortedForGaps.length; i++) {
        const currGap = parse(sortedForGaps[i].number);
        if (!currGap || !prevGap) continue;

        // Check if same series (Prefix + Suffix)
        if (currGap.prefix === prevGap.prefix && currGap.suffix === prevGap.suffix) {
            const diff = currGap.body - prevGap.body;
            if (diff > 1) {
                // FOUND GAP: diff=2 means 1 missing (e.g. 1, 3 -> missing 2)
                // Missing range: (prev+1) to (curr-1)
                const startMissing = prevGap.body + 1;
                const endMissing = currGap.body - 1;
                const countMissing = endMissing - startMissing + 1;

                // Reconstruct ID for display
                const exampleID = `${currGap.prefix}${startMissing.toString().padStart(8, '0')}X${currGap.suffix}`;

                missingItems.push({
                    prefix: currGap.prefix,
                    suffix: currGap.suffix,
                    startBody: startMissing,
                    endBody: endMissing,
                    count: countMissing,
                    example: exampleID
                });
            }
        }
        prevGap = currGap;
    }

    // --- STEP 2: GROUPING FOR REPORT (User Preference: Price) ---
    // Sort by Price (Asc) then Number (Asc)
    uniqueList.sort((a, b) => {
        if (a.price !== b.price) {
            return a.price - b.price;
        }
        return a.number.localeCompare(b.number);
    });

    const rawRanges = [];

    // Helper to parse with price (reusing structure)
    const parseFull = (item) => {
        const p = parse(item.number);
        if (p) { p.price = item.price; p.weight = item.weight; p.hasDiscrepancy = item.hasDiscrepancy; p.originalWeight = item.originalWeight; p.context = item.context; }
        return p;
    };

    let start = parseFull(uniqueList[0]);
    let prev = start;
    let currentList = [uniqueList[0]];

    for (let i = 1; i < uniqueList.length; i++) {
        const curr = parseFull(uniqueList[i]);
        if (!curr) continue;

        const isContinuous = (
            curr.prefix === prev.prefix &&
            curr.suffix === prev.suffix &&
            curr.body === prev.body + 1 &&
            curr.price === prev.price &&
            curr.weight === prev.weight
        );

        if (isContinuous) {
            currentList.push(uniqueList[i]);
            prev = curr;
        } else {
        rawRanges.push({
            start: start.full,
            end: prev.full,
            count: currentList.length,
            price: start.price,
            weight: start.weight,
            items: currentList.map(x => x.number),
            contexts: currentList.map(x => x.context).filter(c => c !== null),
            hasDiscrepancy: currentList.some(x => x.hasDiscrepancy) // v4.0.0 Inline flag
        });
        start = curr;
        prev = curr;
        currentList = [uniqueList[i]];
    }
}
rawRanges.push({
    start: start.full,
    end: prev.full,
    count: currentList.length,
    price: start.price,
    weight: start.weight,
    items: currentList.map(x => x.number),
    contexts: currentList.map(x => x.context).filter(c => c !== null),
    hasDiscrepancy: currentList.some(x => x.hasDiscrepancy) // v4.0.0 Inline flag
});

    // Virtual Optimization (Admin Only)
    const isUserMode = document.body.classList.contains('user-mode');
    if (isUserMode) {
        console.info("[v2.0-stable] Subordinate mode: Skipping virtual optimization to preserve raw sequences.");
        currentImportedBatches = rawRanges;
    } else {
        const optimizedRanges = TrackingUtils.virtualOptimizeRanges(rawRanges);
        currentImportedBatches = optimizedRanges;
    }

    lastMissingItems = missingItems;
    lastDiscrepanciesList = discrepanciesList;

    renderImportResult(currentImportedBatches, missingItems, discrepanciesList);
}

function renderImportResult(ranges, missingItems = [], discrepancies = []) {
    const preview = document.getElementById('import-preview');
    const summary = document.getElementById('import-summary');
    const details = document.getElementById('import-details');
    
    // v4.0.0 Global State checks
    const isUserMode = document.body.classList.contains('user-mode');
    let summaryTableHtml = '';
    let discrepancyHtml = '';

    // v4.0.0: Fuel Surcharge Active Calculation
    const surchargeChecked = document.getElementById('import-surcharge-toggle')?.checked || false;
    const totalItems = ranges.reduce((acc, r) => acc + r.count, 0);
    const surchargeAmount = surchargeChecked ? (totalItems * 3) : 0;
    const itemsTotal = ranges.reduce((acc, r) => acc + (r.total || (r.count * r.price)), 0);
    const grandTotal = itemsTotal + surchargeAmount;

    // --- GAP ALERT SECTION ---
    let gapHtml = '';
    let gapTableRows = ''; // To show in table as well

    if (missingItems.length > 0) {
        const totalMissing = missingItems.reduce((acc, m) => acc + m.count, 0);


        // Generate Alert List
        let listHtml = missingItems.map(m => {
            const startID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.startBody, m.suffix));
            const endID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.endBody, m.suffix));

            const rangeText = (m.count === 1)
                ? startID // Single ID
                : `${startID} - ${endID}`; // Range

            return `<li>${rangeText} (${m.count} รายการ)</li>`;
        }).join('');

        gapHtml = `
            <div class="result-error" style="margin-top:15px; padding:15px; border:2px solid #ff4444; background:#ffebeb;">
                <h3 style="margin-top:0; color:#cc0000;">⚠️ ตรวจพบเลขพัสดุข้ามไป (GAP DETECTED)</h3>
                <p>ระบบพบว่าข้อมูลไม่ต่อเนื่อง <strong>หายไป ${totalMissing} รายการ</strong> ดังนี้:</p>
                <ul style="margin-bottom:0;">${listHtml}</ul>
                <div style="font-size:0.9rem; color:#666; margin-top:10px;">รายการที่หายไปถูกเพิ่มลงในตารางด้านล่างเพื่ออ้างอิงแล้ว (สีแดง)</div>
            </div>
        `;

        // Generate Table Rows for Gaps
        gapTableRows = missingItems.map((m, index) => {
            const startID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.startBody, m.suffix));
            const endID = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.endBody, m.suffix));

            const rangeText = (m.count === 1)
                ? startID
                : `${startID} - ${endID}`;

            return `
                <tr style="background-color:#ffebeb; color:#d32f2f; border-bottom:1px solid #ffcdd2;">
                    <td style="padding:10px; font-weight:bold;">
                        <div class="line-flex">
                            <span>❌ รายการที่หายไป (Missing)</span>
                            <span class="mobile-stats" style="color:#d32f2f;">${m.count} รายการ</span>
                        </div>
                        <span style="font-size:0.9em; font-family:monospace;">${rangeText}</span>
                    </td>
                    <td class="col-qty" style="text-align:right; padding:10px;">${m.count}</td>
                    <td class="col-price" style="text-align:right; padding:10px;">-</td>
                    <td class="col-total" style="text-align:right; padding:10px;">-</td>
                </tr>
            `;
        }).join('');
    } else {
        gapHtml = `
            <div class="result-success" style="margin-top:15px; padding:10px; border:1px solid #4caf50; background:#e8f5e9;">
                <strong>✅ ครบถ้วน 100% (No Gaps)</strong> - เลขพัสดุเรียงต่อเนื่องกันสมบูรณ์
            </div>
        `;
    }

    // --- ABNORMALITY ALERT SECTION (Weight Discrepancy) ---
    if (discrepancies.length > 0) {
        discrepancyHtml = `
            <div class="result-error" style="margin-top:10px; padding:15px; border:2px solid #ef6c00; background:#fff3e0; border-radius:12px;">
                <h3 style="margin-top:0; color:#e65100; font-size:1.1rem; display:flex; align-items:center; gap:8px;">
                    <span style="font-size:1.4rem;">⚠️</span> พบข้อมูลน้ำหนักผิดปกติ (${discrepancies.length} รายการ)
                </h3>
                <p style="margin-bottom:8px; font-size:0.95rem; color:#d84315;">ระบบพบว่าน้ำหนักใน Excel ไม่ตรงกับน้ำหนักที่คำนวณจากราคา (A3 Package):</p>
                <div style="background:white; border-radius:8px; padding:10px; border:1px solid #ffe0b2; max-height: 120px; overflow-y: auto;">
                    <ul style="margin:0; padding-left:20px; font-size:0.9rem; color:#5d4037;">
                        ${discrepancies.slice(0, 5).map(d => `<li><strong>${d.number}</strong>: Excel=${d.originalWeight}kg <span style="color:#d32f2f;">vs</span> กฎราคา=${d.weight}kg</li>`).join('')}
                        ${discrepancies.length > 5 ? `<li style="list-style:none; margin-top:5px; font-style:italic;">... และรายการอื่นอีก ${discrepancies.length - 5} รายการ</li>` : ''}
                    </ul>
                </div>
                <div style="font-size:0.85rem; color:#e65100; margin-top:10px; font-weight:bold; display:flex; align-items:center; gap:5px;">
                    ℹ️ กรุณาตรวจสอบไฟล์ต้นฉบับว่ามีการระบุน้ำหนักคลาดเคลื่อนหรือไม่
                </div>
            </div>
        `;
    }

    // v3.1: Spaced Formatting Utility
    const formatSpaced = (val) => {
        if (!val || val.length < 13) return val;
        const prefix = val.slice(0, 2);
        const body1 = val.slice(2, 6);
        const body2 = val.slice(6, 10);
        const check = val.slice(10, 11);
        const suffix = val.slice(11, 13);
        // Use non-breaking space before suffix to keep it with the digits
        return `${prefix} ${body1} ${body2} ${check}&nbsp;${suffix}`;
    };

    // v2.2: Compute Price/Weight Summary for Subordinates
    if (isUserMode) {
        // --- Logic 1: Sequential Timeline (for Detailed View) ---
        const detailedTimeline = [];
        ranges.forEach(r => detailedTimeline.push({ ...r, type: 'success' }));
        if (missingItems && missingItems.length > 0) {
            missingItems.forEach(m => {
                const mStart = formatID(m.prefix, m.startBody, m.suffix);
                const mEnd = formatID(m.prefix, m.endBody, m.suffix);
                detailedTimeline.push({ start: mStart, end: mEnd, count: m.count, type: 'gap', price: 0, weight: '-' });
            });
        }
        detailedTimeline.sort((a, b) => a.start.localeCompare(b.start));

        // --- Logic 2: Simulated Receipt (Grouped by Price - Tier 1) ---
        const statsMap = {};
        ranges.forEach(r => {
            const key = `${r.price}-${r.weight}`;
            if (!statsMap[key]) {
                statsMap[key] = { 
                    price: r.price, 
                    weight: r.weight, 
                    count: 0, 
                    total: 0, 
                    minId: r.start, 
                    maxId: r.end,
                    hasDiscrepancy: r.hasDiscrepancy // v4.0.0
                };
            } else {
                if (r.start.localeCompare(statsMap[key].minId) < 0) statsMap[key].minId = r.start;
                if (r.end.localeCompare(statsMap[key].maxId) > 0) statsMap[key].maxId = r.end;
                if (r.hasDiscrepancy) statsMap[key].hasDiscrepancy = true; // v4.0.0
            }
            statsMap[key].count += r.count;
            statsMap[key].total += (r.count * r.price);
        });
        const sortedStats = Object.values(statsMap).sort((a, b) => a.price - b.price);

        // --- Logic 3: Global Range ---
        const allTrackings = [];
        ranges.forEach(r => {
            allTrackings.push(r.start);
            allTrackings.push(r.end);
        });
        allTrackings.sort();
        const globalStart = allTrackings[0] || 'N/A';
        const globalEnd = allTrackings[allTrackings.length - 1] || 'N/A';

        const globalRangeHtml = `
            <div style="margin-bottom:15px; padding:12px; border:1px solid #cce5ff; background:#e7f3ff; color:#004085; border-radius:8px; text-align:left;">
                <div style="font-weight:bold; font-size:0.9rem; margin-bottom:4px;">📦 ช่วงเลขพัสดุรวมทั้งชุด (Global Range):</div>
                <div style="font-size:1.15rem; font-weight:900; letter-spacing:0.5px; font-family:monospace; line-height:1.4;">
                    ${formatSpaced(globalStart)} ถึง<br>${formatSpaced(globalEnd)}
                </div>
            </div>
        `;

        // Tier 1: Simulated Receipt View (By Price)
        const summaryReceiptHtml = `
            <div id="receipt-summary-box" style="margin-top:15px;">
                <div style="padding:5px 0 10px 0; border-bottom:2px solid #333; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                    <h4 style="margin:0; text-transform:uppercase; letter-spacing:1px;">🧾 ใบเสร็จรับฝาก (Summary Receipt)</h4>
                    <span style="font-size:0.75rem; color:#666;">v4.2.0-PRO (Surcharge Ready)</span>
                </div>
                
                <div class="receipt-table">
                    ${sortedStats.map((s, idx) => `
                        <div class="receipt-row">
                            <div style="display:flex;">
                                <div class="receipt-seq">${idx + 1}.</div>
                                <div class="receipt-content">
                                    <span class="receipt-badge badge-success">EMS ในฯ</span>
                                    <div class="receipt-title">
                                        EMS ราคา ${s.price} บาท 
                                        ${s.weight !== '-' ? `| น้ำหนัก <span style="font-weight:bold; color:#333;">${s.weight}</span>` : ''}
                                        ${s.hasDiscrepancy ? '<span style="color:#ffffff; font-weight:bold; font-size:0.75rem; margin-left:8px; background:#d32f2f; padding:2px 8px; border-radius:12px; border:1px solid #b71c1c; box-shadow:0 2px 4px rgba(211,47,47,0.3);">⚠️ นน. ไม่ตรง</span>' : ''}
                                    </div>
                                    <div class="receipt-range" style="font-family:monospace; font-weight:900; font-size:1rem; margin-top:5px; color:#333; line-height:1.3;">
                                        ${s.minId === s.maxId ? formatSpaced(s.minId) : `${formatSpaced(s.minId)} ถึง<br>${formatSpaced(s.maxId)}`}
                                    </div>
                                    <div class="receipt-stats">${s.count} @ ${s.price.toFixed(2)}</div>
                                </div>
                                <div class="receipt-total">${s.total.toLocaleString()}</div>
                            </div>
                        </div>
                    `).join('')}
                    
                    ${surchargeChecked ? `
                        <div class="receipt-row surcharge-row" style="background:#fff9c4; border-top:1px dashed #333;">
                            <div style="display:flex;">
                                <div class="receipt-seq">+</div>
                                <div class="receipt-content">
                                    <div class="receipt-title" style="font-weight:bold;">ค่า Fuel Surcharge (Active) ⚡</div>
                                    <div class="receipt-stats">ค่าบริการเหมา ${totalItems} ชิ้น x 3.00 บาท</div>
                                </div>
                                <div class="receipt-total" style="font-weight:bold; color:#d63384;">+ ${surchargeAmount.toLocaleString()}</div>
                            </div>
                        </div>
                    ` : ''}
                </div>

                <!-- Grand Total at bottom (v3.8.2) -->
                <div style="margin-top:20px; border-top:2px solid #333; padding-top:15px; text-align:center; border-bottom:2px solid #333; padding-bottom:15px;">
                    <div style="font-size:1.1rem; color:#333;">จำนวนทั้งหมด: <strong>${totalItems.toLocaleString()}</strong> ชิ้น</div>
                    <div style="font-size:1.6rem; color:#d63384; font-weight:900; letter-spacing:0.5px;">ยอดเงินรวม: ${grandTotal.toLocaleString()} บาท</div>
                </div>
            </div>
        `;

        // Tier 2: Detailed Timeline View (Chronological Gaps - Hidden)
        const detailedTimelineHtml = `
            <div style="margin-top:30px; border-top:1px solid #eee; padding-top:20px;">
                <button onclick="this.nextElementSibling.classList.toggle('hidden')" 
                        style="width:100%; background:#f8f9fa; border:1px solid #ddd; padding:10px; border-radius:8px; font-size:0.85rem; color:#666; cursor:pointer; font-weight:bold;">
                    📜 ดูการแจงลำดับเลขพัสดุแบบละเอียด (Detailed Details)
                </button>
                <div class="hidden" style="margin-top:10px; background:#fff; border:1px solid #eee; border-radius:8px; overflow:hidden;">
                    <div class="receipt-table" style="box-shadow:none;">
                        ${detailedTimeline.map((item, idx) => {
                            if (item.type === 'gap') {
                                return `
                                    <div class="receipt-row receipt-gap-row">
                                        <div style="display:flex;">
                                            <div class="receipt-seq">-</div>
                                            <div class="receipt-content">
                                                <span class="receipt-badge badge-gap">❌ ข้ามรายการ (Missing)</span>
                                                <div class="receipt-range" style="line-height:1.3;">${item.start === item.end ? formatSpaced(item.start) : `${formatSpaced(item.start)} ถึง<br>${formatSpaced(item.end)}`}</div>
                                                <div class="receipt-stats">หายไปทั้งหมด ${item.count} รายการ</div>
                                            </div>
                                            <div class="receipt-total">-</div>
                                        </div>
                                    </div>
                                `;
                            } else {
                                const subtotal = (item.count * item.price);
                                return `
                                    <div class="receipt-row">
                                        <div style="display:flex; padding: 12px 15px;">
                                            <div class="receipt-seq" style="width:30px; font-size:0.8rem; color:#ccc;">${idx + 1}</div>
                                            <div class="receipt-content">
                                                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px;">
                                                    <span style="background:#f1f5f9; color:#475569; font-size:0.65rem; padding:1px 6px; border-radius:4px; font-weight:bold; text-transform:uppercase;">EMS รายชิ้น</span>
                                                    <span style="color:#d63384; font-weight:900; font-size:1.1rem; line-height:1;">${subtotal.toLocaleString()}</span>
                                                </div>
                                                <div style="font-size:0.95rem; font-weight:700; color:#334155; margin-bottom:4px;">
                                                    ราคา <span style="color:#3b82f6;">${item.price}</span> บาท
                                                </div>
                                                <div class="receipt-range" style="font-size:0.9rem; line-height:1.3; color:#1e293b; font-family:monospace; margin-bottom:6px;">
                                                    ${item.start === item.end ? formatSpaced(item.start) : `${formatSpaced(item.start)} ถึง<br>${formatSpaced(item.end)}`}
                                                </div>
                                                <div style="font-size:0.8rem; color:#64748b; font-weight:500;">
                                                    📦 จำนวน <strong style="color:#333;">${item.count}</strong> ชิ้น
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                `;
                            }
                        }).join('')}
                    </div>
                </div>
            </div>
        `;

        summaryTableHtml = `
            ${globalRangeHtml}
            ${summaryReceiptHtml}
            ${detailedTimelineHtml}
        `;
    }
    
    // v4.0.0: Ensure preview is visible
    if (preview) preview.classList.remove('hidden');
    
    // v4.0.0: Surcharge Line Item
    let surchargeHtml = '';
    if (surchargeChecked) {
        surchargeHtml = `
            <div style="margin-bottom:10px; padding:10px; background:#fff8e1; border:1px solid #ffe082; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="color:#795548; font-weight:bold; font-size:0.85rem;">⚡ ค่า Fuel Surcharge (+3฿)</span>
                <span style="color:#795548; font-weight:bold;">${totalItems} x 3 = ${surchargeAmount.toLocaleString()} บาท</span>
            </div>
        `;
    }

    summary.innerHTML = `
        <div style="margin-bottom:15px; text-align:center;">
            <div style="font-size:1.1rem; color:#333;">จำนวนทั้งหมด: <strong>${totalItems.toLocaleString()}</strong> ชิ้น</div>
            <div style="font-size:1.3rem; color:#666;">ยอดเบื้องต้น: ${itemsTotal.toLocaleString()} บาท</div>
            <div style="font-size:1.6rem; color:#d63384; font-weight:900; margin-top:5px;">ยอดสุทธิ: ${grandTotal.toLocaleString()} บาท</div>
        </div>
        ${surchargeHtml}
        ${discrepancyHtml}
        ${gapHtml}
        ${summaryTableHtml}
    `;

    // Generate Layout
    const tableStyle = isUserMode 
        ? "background:white; padding:15px; border:1px solid #eee; border-radius:12px; box-shadow:0 10px 30px rgba(0,0,0,0.05); font-family:'Sarabun', sans-serif;" 
        : "background:white; padding:20px; border:1px solid #ddd; box-shadow:0 2px 5px rgba(0,0,0,0.05); font-family:'Courier New', monospace;";

    let html = `
        <div style="${tableStyle}">
            ${!isUserMode ? `
            <div style="text-align:center; margin-bottom:15px;">
                <button onclick="document.getElementById('import-detailed-list').classList.toggle('hidden')" 
                        style="background:#fff3cd; border:1px solid #ffeeba; padding:10px 20px; border-radius:25px; font-size:0.9rem; color:#856404; cursor:pointer; font-weight:bold; box-shadow:0 2px 5px rgba(0,0,0,0.05);">
                    📋 ดูลำดับพัสดุต่อเนื่อง (Timeline Flow)
                </button>
            </div>
            ` : ''}
            
            <div id="import-detailed-list" class="${isUserMode ? 'hidden' : 'hidden'}">
                <table style="width:100%; border-collapse: collapse;">
                <tbody>
    `;

    if (isUserMode) {
        // v2.7: Combined Timeline (Success + Gaps)
        const timeline = [];
        ranges.forEach(r => timeline.push({ ...r, type: 'success' }));
        if (missingItems && missingItems.length > 0) {
            missingItems.forEach(m => {
                const mStart = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.startBody, m.suffix));
                const mEnd = TrackingUtils.formatTrackingNumber(formatID(m.prefix, m.endBody, m.suffix));
                timeline.push({ start: mStart, end: mEnd, count: m.count, type: 'gap' });
            });
        }
        timeline.sort((a, b) => a.start.localeCompare(b.start));

        timeline.forEach(item => {
            if (item.type === 'gap') {
                html += `
                    <tr style="background:#fff5f5; color:#c62828;">
                        <td colspan="3" style="padding:12px 10px; font-weight:bold; border-left:5px solid #c62828; border-bottom:1px solid #ffcdd2;">
                            <div style="font-size:0.8rem; opacity:0.8;">⚠️ ไม่มีเลขที่นี้ (Missing Range)</div>
                            <div style="font-size:1.05rem; letter-spacing:0.5px; font-family:monospace;">
                                ${item.start === item.end ? formatSpaced(item.start) : `${formatSpaced(item.start)} ถึง<br>${formatSpaced(item.end)}`}
                            </div>
                            <div style="font-size:0.75rem; margin-top:2px;">(หายไป ${item.count} รายการ)</div>
                        </td>
                    </tr>
                `;
            } else {
                const rowTotalStr = (item.count * item.price).toLocaleString();
                html += `
                    <tr style="border-bottom:1px solid #f0f0f0;">
                         <td style="padding:10px; vertical-align:middle; width:70%;">
                            <div style="font-size:0.7rem; color:#666; text-transform:uppercase;">
                                📦 EMS ${item.price}฿ | ${item.weight} 
                                ${item.hasDiscrepancy ? '<span style="color:#e65100; font-weight:bold; margin-left:5px;">⚠️ นน. ไม่ตรง</span>' : ''}
                            </div>
                            <div style="color:#0056b3; font-weight:bold; font-size:1.05rem; line-height:1.2; font-family:monospace;">
                                ${item.start === item.end ? formatSpaced(item.start) : `${formatSpaced(item.start)} - ${formatSpaced(item.end)}`}
                            </div>
                        </td>
                        <td style="padding:10px 0; text-align:right; font-size:0.9rem; color:#666; width:10%; vertical-align:middle;">${item.count}</td>
                        <td style="padding:10px; text-align:right; font-weight:bold; font-size:1rem; color:#333; width:20%; vertical-align:middle;">${rowTotalStr}</td>
                    </tr>
                `;
            }
        });
    } else {
        // Fallback for Admin or empty state
        ranges.forEach((r, idx) => {
            const rowTotal = r.total || (r.count * r.price);
            const rowTotalStr = rowTotal.toLocaleString('en-US', { minimumFractionDigits: 2 });
            html += `
                <tr style="border-bottom:1px solid #eee;">
                    <td style="padding:12px 0;">
                        <strong>${idx + 1}. EMS ราคา ${r.price} บาท</strong><br>
                        ${r.start === r.end ? r.start : `${r.start} - ${r.end}`}
                    </td>
                    <td style="text-align:right;">${r.count}</td>
                    <td style="text-align:right;">${r.price}</td>
                    <td style="text-align:right;">${rowTotalStr}</td>
                </tr>
            `;
        });
    }

    html += `
                    </tbody>
                    <tfoot>
                        <tr style="border-top:2px solid #000; border-bottom:2px solid #000;">
                            <td colspan="4" style="padding:10px;">
                                <div class="mobile-grand-total">
                                    <span style="font-weight:bold;">รวมทั้งสิ้น (Grand Total)</span>
                                    <span style="font-weight:bold; font-size:1.1rem;">${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                                </div>
                            </td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
    `;
    details.innerHTML = html;

    // Auto-fill batch name suggestion
    const today = new Date().toLocaleDateString('th-TH');
    document.getElementById('import-batch-name').value = `Import ${today} (${totalItems} items)`;

    // AUTO SAVE (Triggers immediately)
    setTimeout(() => {
        saveImportedBatch(true);
    }, 500);
}

function saveImportedBatch(isAuto = false) {
    if (currentImportedBatches.length === 0) return;

    const name = document.getElementById('import-batch-name').value.trim();
    const type = document.getElementById('import-batch-type').value;

    // Save choice for next time
    localStorage.setItem('thp_last_batch_type', type);

    if (!name) {
        if (isAuto) {
            const now = new Date();
            const timeStr = now.getHours().toString().padStart(2, '0') + ":" + now.getMinutes().toString().padStart(2, '0');
            const dateStr = now.toLocaleDateString('th-TH');
            document.getElementById('import-batch-name').value = `นำเข้าอัตโนมัติ ${dateStr} [${timeStr}]`;
        } else {
            alert('กรุณาระบุชื่อกลุ่มข้อมูล (Batch Name)');
            return;
        }
    }

    if (!isAuto) {
        if (!confirm(`ยืนยันบันทึกข้อมูล ${currentImportedBatches.length} ช่วงรวมกันเป็น 1 ชุดข้อมูล?`)) return;
    }

    // Save Logic (Optimized)
    const allItemsToSave = [];
    currentImportedBatches.forEach(r => {
        // Since we optimized, we generate from start-end.
        // Assuming contiguous S10. 
        const list = TrackingUtils.generateTrackingRange(r.start, 0, r.count - 1);
        list.forEach(i => allItemsToSave.push(i.number));
    });

    // Metadata
    const rangesMeta = currentImportedBatches.map(r => ({
        start: r.start,
        end: r.end,
        count: r.count,
        price: r.price,
        weight: r.weight,
        total: r.total || (r.count * r.price)
    }));

    const batchInfo = {
        name: name,
        type: type,
        contract: 'Imported',
        timestamp: new Date().getTime(),
        ranges: rangesMeta
    };

    const btn = document.querySelector('button[onclick="saveImportedBatch()"]');
    const originalHtml = btn ? btn.innerHTML : '💾 บันทึกข้อมูล';
    if (btn) window.setButtonLoading(btn, true);

    // v1.93 Refactor: Use async/await for more reliable error handling
    (async () => {
        try {
            // Give UI a moment to show loading state
            await new Promise(r => setTimeout(r, 100));

            // Save
            const result = await CustomerDB.addBatch(batchInfo, allItemsToSave);
            const addedCount = typeof result === 'object' ? result.count : result;
            const newBatchId = typeof result === 'object' ? result.id : null;

            if (result && result.error === 'DUPLICATE') {
                window.showToast(`ข้อมูลชุดนี้มีอยู่แล้วในระบบ (Batch: ${result.id})`, 'info');
            } else {
                window.showToast(`บันทึกเรียบร้อย! เพิ่ม ${addedCount} รายการ`);
                
                // v4.3.0: Legacy stats are no longer needed as Dashboard pulls from Batches directly.
                console.info("[v4.3.0] Transaction recorded to core database.");
            }

            // v2.0-stable: Handle Navigation based on role
            const isUserMode = document.body.classList.contains('user-mode');

            if (isUserMode) {
                // For Subordinates: STAY HERE, don't hide, don't switch.
                console.info("[v2.0-stable] Subordinate mode: Saved batch, keeping results visible.");
                // We do NOT clear previewSec or currentImportedBatches immediately
                // However, we might want to update the UI to show "Saved" status
                window.showToast(`บันทึกเรียบร้อย! ข้อมูลถูกเก็บไว้ในประวัติแล้ว`);
            } else {
                // Admin Flow: Reset and Switch
                const uploadBtn = document.getElementById('excel-upload');
                if (uploadBtn) uploadBtn.value = '';
                const previewSec = document.getElementById('import-preview');
                if (previewSec) previewSec.classList.add('hidden');
                currentImportedBatches = [];

                if (newBatchId && typeof loadBatchToView === 'function') {
                    loadBatchToView(newBatchId);
                } else {
                    switchTab('customer');
                    if (typeof updateDbViews === 'function') await updateDbViews();
                }
            }

            // v2.0-stable: Auto-Cleanup Search Results after Batch Save
            if (typeof currentUnifiedResults !== 'undefined' && currentUnifiedResults && currentUnifiedResults.length > 0) {
                const savedSet = new Set(allItemsToSave.map(t => t.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '')));
                console.info(`[v2.0-stable] Cleaning up ${savedSet.size} items from search results...`);
                
                currentUnifiedResults = currentUnifiedResults.filter(item => {
                    const cleanNum = item.number.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
                    return !savedSet.has(cleanNum);
                });

                // Persist new state
                await saveLastSearchResults(currentUnifiedTitle || "", currentUnifiedResults, false);
                
                // Re-render
                if (typeof renderStoredUnifiedNumbers === 'function') {
                    renderStoredUnifiedNumbers(currentUnifiedTitle || "", currentUnifiedResults);
                }
            }
        } catch (err) {
            console.error("saveImportedBatch Error:", err);
            alert("เกิดข้อผิดพลาดในการบันทึกชุดข้อมูล: " + err.message);
        } finally {
            if (btn) window.setButtonLoading(btn, false, originalHtml);
        }
    })();
}

// --- Snapshot Logic ---
function toggleSnapshotList() {
    const list = document.getElementById('snapshot-list');
    if (list.classList.contains('hidden')) {
        list.classList.remove('hidden');
        renderSnapshotList();
    } else {
        list.classList.add('hidden');
    }
}

function renderSnapshotList() {
    const list = document.getElementById('snapshot-list');
    const snapshots = CustomerDB.getSnapshots();

    if (snapshots.length === 0) {
        list.innerHTML = '<li>ไม่มีจุดย้อนกลับ (No Snapshots)</li>';
        return;
    }

    let html = '';
    snapshots.forEach(s => {
        const time = new Date(s.timestamp).toLocaleString('th-TH');
        html += `
            <li style="margin-bottom:5px;">
                <strong>${time}</strong> - ${s.reason} 
                <button class="btn" style="padding:2px 6px; font-size:0.7rem; background:#ffc107; color:#000; margin-left:10px;"
                    onclick="restoreFromSnapshot(${s.timestamp})">Restore</button>
            </li>
        `;
    });
    list.innerHTML = html;
}

function restoreFromSnapshot(ts) {
    if (confirm('ยืนยันย้อนเวลากลับไปจุดนี้? (ข้อมูลปัจจุบันจะหายไป)')) {
        if (CustomerDB.restoreSnapshot(ts)) {
            alert('✅ ย้อนเวลาสำเร็จ (Restored)');
            renderDBTable();
        } else {
            alert('❌ ไม่พบข้อมูล Snapshot นี้');
        }
    }
}

// --- Backup & Restore Glue Code ---
function confirmAndBackup() {
    if (confirm('คุณต้องการ Export ข้อมูลทั้งหมดออกมาเป็นไฟล์ใช่หรือไม่?')) {
        backupData();
    }
}

function backupData() {
    CustomerDB.exportBackup();
}

function restoreData(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (!confirm('คำเตือน: การกู้คืนข้อมูลจะทับข้อมูลปัจจุบันทั้งหมด\nคุณต้องการดำเนินการต่อหรือไม่?')) {
        event.target.value = ''; // Reset
        return;
    }

    CustomerDB.importBackup(file)
        .then(() => {
            alert('✅ กู้คืนข้อมูลสำเร็จ (Restore Complete)');
            if (typeof updateDbViews === 'function') updateDbViews();
            else if (typeof renderDBTable === 'function') renderDBTable(); // Refresh UI
        })
        .catch(err => {
            alert('❌ เกิดข้อผิดพลาด: ' + err.message);
        })
        .finally(() => {
            event.target.value = ''; // Reset
        });
}

function loadBatchToView(batchId) {
    const batches = CustomerDB.getBatches();
    const batch = batches[batchId];

    if (!batch) {
        alert('ไม่พบข้อมูลชุดนี้');
        return;
    }

    // Redirect to Smart Workspace to display
    switchTab('smart');
    const box = document.getElementById('smart-unified-results');

    // CHECK IF WE HAVE RECEIPT METADATA (ranges)
    if (batch.ranges && Array.isArray(batch.ranges)) {
        // Render Receipt Style (Directly use saved metadata which is already optimized)
        const grandTotal = batch.ranges.reduce((acc, r) => acc + (r.total || 0), 0);

        let html = `
            <div class="result-success" style="margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                <strong>📂 ข้อมูล: ${batch.name}</strong>
                <button class="btn btn-neutral" onclick="switchTab('customer')" style="padding:5px 10px; font-size:0.9rem;">⬅ กลับหน้าหลัก (Back to DB)</button>
            </div>
            
            <!-- Receipt View -->
            <div style="background:white; padding:20px; border:1px solid #ddd; box-shadow:0 2px 5px rgba(0,0,0,0.05); font-family:'Courier New', monospace; max-width:800px; margin:0 auto;">
                <h4 style="text-align:center; border-bottom:1px dashed #ccc; padding-bottom:10px; margin-bottom:10px;">ใบสรุปรายการ (Optimized Report)</h4>
                 <div style="margin-bottom:10px; font-size:0.9rem;">
                    <strong>Customer:</strong> ${batch.name}<br>
                    <strong>Type:</strong> ${batch.type}<br>
                    ${batch.requestDate ? `<strong>Request Date:</strong> ${new Date(batch.requestDate).toLocaleDateString('th-TH')}<br>` : ''}
                    <strong>Date:</strong> ${new Date(batch.timestamp).toLocaleString('th-TH')}
                </div>
                 <div style="font-size:0.8rem; color:red; text-align:center; margin-bottom:5px;">
                    *จัดเรียงตามราคาน้อย-มาก (Virtual)
                </div>
                <div class="report-list-container">
                    <div class="report-header-desktop">
                        <div style="width:40%">รายการ (Description)</div>
                        <div style="width:20%; text-align:right">จำนวน (Qty)</div>
                        <div style="width:20%; text-align:right">ราคา/ชิ้น</div>
                        <div style="width:20%; text-align:right">รวม (Total)</div>
                    </div>
        `;



        batch.ranges.forEach((r, idx) => {
            html += `
                <div class="report-card">
                    <div class="report-card-desc">
                        <strong>${idx + 1}. EMS ราคา ${r.price} บาท</strong><br>
                        <span style="color:#0056b3; font-weight:bold;">${r.start === r.end ? r.start : `${r.start} - ${r.end}`}</span><br>
                        <small>น้ำหนัก (Weight): ${r.weight}</small>
                    </div>
                    <div class="report-card-scroll">
                        <div class="stat-item">
                            <span class="stat-label">จำนวน (Qty)</span>
                            <span class="stat-value">${r.count}</span>
                        </div>
                        <div class="stat-item">
                            <span class="stat-label">ราคา/ชิ้น</span>
                            <span class="stat-value">@${r.price}</span>
                        </div>
                        <div class="stat-item highlight">
                            <span class="stat-label">รวม (Total)</span>
                            <span class="stat-value">${(r.total || 0).toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                        </div>
                        <!-- Spacer for scroll feel -->
                        <div style="min-width:10px;"></div>
                    </div>
                </div>
            `;
        });

        html += `
                    <div class="report-footer">
                        <div class="footer-label">รวมทั้งสิ้น (Grand Total)</div>
                        <div class="footer-value">${grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</div>
                    </div>
                </div>
            </div>
            
            <div style="text-align:center; margin-top:20px;" class="no-print">
                <button class="btn btn-primary" onclick="showThermalReceipt('${batchId}')">🧾 ดูใบเสร็จแบบความร้อน (Thermal Mode)</button>
                <button class="btn" onclick="window.print()" style="margin-left:10px;">🖨️ Print / PDF</button>
                <button class="btn btn-neutral" onclick="switchTab('customer')" style="margin-left:10px;">⬅ กลับหน้าหลัก (Back to DB)</button>
            </div>
        `;
        box.innerHTML = html;
        return;
    }

    // FALLBACK
    // Inline Fallback for safety (Legacy View)
    const lookup = CustomerDB.getLookup();
    const items = [];
    Object.keys(lookup).forEach(key => {
        if (lookup[key].batchId === batchId) items.push(key);
    });
    items.sort();

    let fallbackHtml = `
        <div class="result-success">Viewing Legacy Batch: ${batch.name} (${items.length} items)</div>
        <ul>${items.map(x => `<li>${x}</li>`).slice(0, 50).join('')}</ul>
    `;
    box.innerHTML = fallbackHtml;
}

// --- Thermal Receipt Mode (v2.3) ---

function showThermalReceipt(batchId) {
    const batches = CustomerDB.getBatches();
    const batch = batches[batchId];
    if (!batch) return;
    renderThermalReceiptUI(batch);
}

function showCurrentActiveThermalReceipt() {
    // Collect from current imported batches
    if (!currentImportedBatches || currentImportedBatches.length === 0) return;
    
    const name = document.getElementById('import-batch-name').value.trim() || "New Import";
    const batch = {
        name: name,
        ranges: currentImportedBatches,
        timestamp: new Date().getTime()
    };
    renderThermalReceiptUI(batch);
}

function renderThermalReceiptUI(batch) {
    const box = document.getElementById('smart-unified-results');
    if (!box) return;

    const ranges = batch.ranges || [];
    const grandTotalA3 = ranges.reduce((acc, r) => acc + (r.total || (r.count * r.price)), 0);
    
    const dateStr = new Date(batch.timestamp).toLocaleDateString('th-TH');
    const timeStr = new Date(batch.timestamp).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });

    let itemsHtml = '';
    ranges.forEach((r, idx) => {
        const stdPrice = TrackingUtils.getStandardPrice(r.price);
        const stdTotal = r.count * stdPrice;
        const discountTotal = stdTotal - (r.total || (r.count * r.price));
        const startNum = TrackingUtils.formatTrackingNumber(r.start);
        const endNum = TrackingUtils.formatTrackingNumber(r.end);

        itemsHtml += `
            <div class="thermal-item">
                <span class="thermal-item-title">${idx + 1}. EMS ในฯ</span>
                <span class="thermal-item-range">${startNum} - ${endNum}</span>
                <div class="thermal-item-row">
                    <div class="thermal-item-calc">${r.count}@${stdPrice.toFixed(2)}</div>
                    <div class="thermal-item-total">฿${stdTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</div>
                </div>
                ${discountTotal > 0 ? `
                <div class="thermal-discount-row">
                    <span>ส่วนลดส่งเสริมการขาย</span>
                    <span>฿-${discountTotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                </div>
                ` : ''}
            </div>
        `;
    });

    const html = `
        <div class="thermal-receipt-container">
            <div class="thermal-receipt">
                <div class="thermal-header">
                    <h3>บริษัท ไปรษณีย์ไทย จำกัด</h3>
                    <div style="font-size:0.8rem;">ปณ. หลัก 10501</div>
                </div>
                <div class="thermal-meta">
                    <div style="display:flex; justify-content:space-between;">
                        <span>วันที่: ${dateStr}</span>
                        <span>เวลา: ${timeStr}</span>
                    </div>
                    <div>บิลเลขที่: ${batch.name}</div>
                    <div>ผู้รับฝาก: USER</div>
                </div>
                
                <div class="thermal-separator"></div>
                
                <div class="thermal-body">
                    ${itemsHtml}
                </div>
                
                <div class="thermal-separator"></div>
                
                <div class="thermal-footer">
                    <div class="thermal-grand-total">
                        <span>ยอดรวมทั้งสิ้น</span>
                        <span>฿${grandTotalA3.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span>
                    </div>
                    <div style="text-align:center; font-size:0.8rem; margin-top:10px;">ขอบคุณที่ใช้บริการ</div>
                </div>
            </div>
            
            <div class="receipt-actions no-print">
                <button class="btn btn-primary" onclick="window.print()">🖨️ พิมพ์ใบเสร็จ (Print Receipt)</button>
                <button class="btn btn-neutral" onclick="loadBatchToView('${batch.id || ''}')">🔙 ย้อนกลับ (Back)</button>
            </div>
        </div>
    `;

    // Overwrite the view
    box.innerHTML = html;
    window.scrollTo(0, 0);
}

// --- ADMIN TOOLS ---

function adminHandleTrackInput(inputEl) {
    // Basic cleanup but allow spaces and commas for multiple items
    let val = inputEl.value.toUpperCase().replace(/[^A-Z0-9\s,]/g, '');
    
    // Auto Formatter: If user pastes a giant string without commas, we can add them to look nice.
    if (val.length >= 26 && !val.includes(',')) {
        const extracted = TrackingUtils.extractTrackingNumbers(val);
        if (extracted.length > 1) {
            val = extracted.join(', ');
        }
    }
    
    inputEl.value = val;
}

function adminOpenThpTrack() {
    let rawInput = document.getElementById('admin-track-input').value.trim();
    if (!rawInput) {
        alert("กรุณากรอกเลขพัสดุ");
        return;
    }
    
    // v1.73: Clean all possible space characters from the extracted IDs
    const extracted = TrackingUtils.extractTrackingNumbers(rawInput).map(id => id.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, ''));
    
    if (extracted.length === 0) {
        alert("ไม่พบรูปแบบเลขพัสดุที่ถูกต้อง (13 หลัก)");
        return;
    }
    
    // THP supports submitting multiple numbers joined by comma
    const joinedIds = extracted.join(',');
    const url = `https://track.thailandpost.co.th/dashboard?trackNumber=${joinedIds}`;
    window.open(url, '_blank');
}

/**
 * Applies grayscale and high-contrast thresholding to an image file.
 * Returns a Data URL of the processed image.
 * This removes grey backgrounds and light table lines, helping Tesseract OCR.
 */
function preprocessImageForOCR(file) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext('2d', { willReadFrequently: true });
            
            ctx.drawImage(img, 0, 0);
            
            const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            const data = imageData.data;
            
            // Threshold value (0-255). Lower = more things become white.
            // 150 is usually good for removing light grey shadows but keeping blue/black ink
            const threshold = 150; 
            
            for (let i = 0; i < data.length; i += 4) {
                const r = data[i];
                const g = data[i + 1];
                const b = data[i + 2];
                // Luminance formula
                const v = (0.2126 * r + 0.7152 * g + 0.0722 *b >= threshold) ? 255 : 0;
                
                data[i] = data[i + 1] = data[i + 2] = v; // Set R, G, B
            }
            
            ctx.putImageData(imageData, 0, 0);
            resolve(canvas.toDataURL('image/jpeg', 1.0));
        };
        img.onerror = reject;
        img.src = URL.createObjectURL(file);
    });
}

async function adminHandleImageOcr(files) {
    if (!files || files.length === 0) return;

    const statusEl = document.getElementById('admin-ocr-status');
    const resultEl = document.getElementById('admin-ocr-result');
    resultEl.classList.remove('hidden');
    resultEl.innerHTML = '';
    
    if (statusEl) statusEl.textContent = `กำลังเริ่มประมวลผล...`;

    let combinedText = "";
    let debugProcessedImageUrl = "";

    try {
        for (let i = 0; i < files.length; i++) {
            if (statusEl) statusEl.textContent = `กำลังแยกข้อความจากภาพ ${i + 1}/${files.length}...`;
            const file = files[i];

            // Image Preprocessing: Grayscale + Thresholding
            if (statusEl) statusEl.textContent = `กำลังปรับแต่งรูปภาพ (Preprocessing)...`;
            const processedFileUrl = await preprocessImageForOCR(file);
            debugProcessedImageUrl = processedFileUrl; // Save for debug display

            // Tesseract OCR
            if (statusEl) statusEl.textContent = `กำลังแยกข้อความจากภาพ ${i + 1}/${files.length}...`;
            const useEngOnly = document.getElementById('admin-ocr-eng-only') && document.getElementById('admin-ocr-eng-only').checked;
            const lang = useEngOnly ? 'eng' : 'tha+eng';
            const worker = typeof Tesseract !== 'undefined' ? await Tesseract.createWorker(lang) : null;
            if (worker) {
                // Adjust PSM for tabular data (6 = Assume a single uniform block of text)
                await worker.setParameters({
                    tessedit_pageseg_mode: '6',
                });
                const { data: { text } } = await worker.recognize(processedFileUrl);
                await worker.terminate();
                combinedText += "\n" + text;
            }
        }

        if (statusEl) statusEl.textContent = "กำลังวิเคราะห์ข้อมูล (Analyzing)...";

        const lines = combinedText.split('\n');
        let extractedItems = [];
        
        // 1. Line-by-line parsing to link Tracking and Price on the SAME line
        lines.forEach(line => {
            const trackMatch = line.match(/[A-Z]{2}\s*\d{9}\s*[A-Z]{2}/i);
            const priceMatch = line.match(/(?:\s|^)(\d{1,4}(?:[.,]\d{2})?)(?:\s|$)/); 
            // Thai post receipt usually has price like 42, 32, 52 or 42.00 at the end of the line
            
            if (trackMatch) {
                let track = trackMatch[0].replace(/\s/g, '').toUpperCase();
                let price = 0;
                
                // If there's a number that looks like a price on the same line
                if (priceMatch && priceMatch[1]) {
                     // Filter out numbers that are just sequence numbers like "1", "2" at the start
                     const potentialPrice = parseFloat(priceMatch[1].replace(',', '.'));
                     if (potentialPrice >= 10 && potentialPrice < 5000) { // Reasonable postal bounds
                         price = potentialPrice;
                     }
                }
                extractedItems.push({ track, price, raw: line });
            }
        });

        if (extractedItems.length === 0) {
            statusEl.textContent = "ประมวลผลเสร็จสิ้น: ไม่พบเลขพัสดุในภาพ";
            resultEl.innerHTML = `<em>ไม่พบข้อมูลที่ตรงกับรูปแบบเลขไปรษณีย์ไทย 13 หลัก</em>\n\n[ข้อความดิบจาก OCR]\n${combinedText}\n<hr><img src="${debugProcessedImageUrl}" style="max-width:100%; border:1px solid #ccc; margin-top:10px;">`;
            return;
        }

        // 2. Output Formatting & Missing Check
        let outputHtml = `<strong style="color:var(--primary-color);">📌 ตรวจพบทั้งหมด ${extractedItems.length} รายการ</strong>\n<hr>`;
        let totalPrice = 0;
        
        // Sort items conceptually if possible, but keep original order for display
        let tracksOnly = extractedItems.map(x => x.track);
        
        // Find Missing logic
        let missingReport = "";
        const seqs = extractedItems.map(item => {
            let seqStr = item.track.substring(2, 10);
            return {
                prefix: item.track.substring(0,2),
                suffix: item.track.substring(11,13),
                seq: parseInt(seqStr, 10),
                full: item.track
            };
        }).filter(x => !isNaN(x.seq));

        // Group by prefix+suffix
        const groups = {};
        seqs.forEach(s => {
            let key = s.prefix + s.suffix;
            if(!groups[key]) groups[key] = [];
            groups[key].push(s.seq);
        });

        for (const [key, seqList] of Object.entries(groups)) {
            seqList.sort((a,b) => a - b);
            let missingInGrp = [];
            for (let i = 0; i < seqList.length - 1; i++) {
                let diff = seqList[i+1] - seqList[i];
                if (diff > 1 && diff < 50) { // arbitrary bound so we don't list thousands if it's two different books
                     for (let j = seqList[i] + 1; j < seqList[i+1]; j++) {
                         let missingNum = j.toString().padStart(8, '0');
                         let cd = TrackingUtils.calculateS10CheckDigit(missingNum);
                         if(cd !== null){
                             missingInGrp.push(`${key.substring(0,2)}${missingNum}${cd}${key.substring(2,4)}`);
                         }
                     }
                }
            }
            if (missingInGrp.length > 0) {
                missingReport += `<div style="color:var(--error-color); margin-bottom:10px;">
                    <strong>🚨 แจ้งเตือน: พบช่องโหว่ (เลขที่หายไป) ในช่วงหมวด ${key}:</strong><br>
                    ${missingInGrp.join(', ')}
                </div>`;
            }
        }

        if (missingReport) {
            outputHtml += missingReport + "<hr>";
        }

        // List all with Prices
        outputHtml += `<table style="width:100%; font-size:0.9rem;">
            <tr style="background:#eee;">
                <th style="padding:5px; text-align:left;">ลำดับ</th>
                <th style="padding:5px; text-align:left;">เลขพัสดุ</th>
                <th style="padding:5px; text-align:right;">ราคา (บาท)</th>
            </tr>
        `;
        
        extractedItems.forEach((item, idx) => {
             totalPrice += item.price;
             outputHtml += `
             <tr>
                 <td style="padding:5px; border-bottom:1px solid #eee;">${idx+1}</td>
                 <td style="padding:5px; border-bottom:1px solid #eee;"><strong>${item.track}</strong></td>
                 <td style="padding:5px; border-bottom:1px solid #eee; text-align:right;">${item.price > 0 ? item.price.toFixed(2) : '-'}</td>
             </tr>`;
        });
        
        outputHtml += `
            <tr style="background:#fce8e6; font-weight:bold;">
                <td colspan="2" style="padding:10px; text-align:right;">ยอดรวมทั้งหมด (Total):</td>
                <td style="padding:10px; text-align:right;">${totalPrice.toFixed(2)}</td>
            </tr>
        </table>`;
        
        // DEBUG BLOCK: Output raw text to help us fix the regex
        outputHtml += `
            <div style="margin-top:20px; padding:10px; background:#f8f9fa; border:1px solid #ccc; border-radius:4px; max-height:200px; overflow-y:auto;">
                 <strong style="color:#666; font-size:0.8rem;">[DEBUG] RAW OCR TEXT:</strong>
                 <pre style="font-size:0.75rem; white-space:pre-wrap; margin:0;">${combinedText}</pre>
            </div>
            <div style="margin-top:10px;">
                <strong style="color:#666; font-size:0.8rem;">[DEBUG] ภาพที่ปรับแสงแล้ว:</strong><br>
                <img src="${debugProcessedImageUrl}" style="max-width:100%; border:1px solid #ccc; margin-top:5px;">
            </div>
        `;

        resultEl.innerHTML = outputHtml;
        statusEl.textContent = "ประมวลผลเสร็จสิ้น (Done)";

    } catch (err) {
        console.error(err);
        if (statusEl) statusEl.textContent = "Error: " + err.message;
        resultEl.innerHTML = `<span style="color:red;">Error processing image.</span>`;
    }

    // Reset input
    const uploadInput = document.getElementById('admin-ocr-upload');
    if (uploadInput) uploadInput.value = '';
}

function adminCrossReference() {
    const input = document.getElementById('admin-crossref-input').value;
    const extracted = TrackingUtils.extractTrackingNumbers(input);
    _performCrossRef(extracted);
}

function copyCrossRefAll() {
    const ids = Array.from(document.querySelectorAll('.crossref-id')).map(el => el.innerText.replace(/\s/g, '')).join('\n');
    if(ids) {
        navigator.clipboard.writeText(ids).then(() => alert('คัดลอกรายการทั้งหมดลง Clipboard แล้ว'));
    } else {
        alert('ไม่มีข้อมูลให้คัดลอก');
    }
}

async function adminCrossRefImage(files) {
    if (!files || files.length === 0) return;
    
    const statusEl = document.getElementById('admin-crossref-status');
    const resultEl = document.getElementById('admin-crossref-result');
    resultEl.classList.add('hidden');
    
    if (statusEl) statusEl.textContent = `กำลังสแกนรูปภาพด้วย OCR...`;

    try {
        const worker = typeof Tesseract !== 'undefined' ? await Tesseract.createWorker('eng') : null; // 'eng' is faster for just numbers
        if (worker) {
            const { data: { text } } = await worker.recognize(files[0]);
            await worker.terminate();
            
            const extracted = TrackingUtils.extractTrackingNumbers(text);
            if(extracted.length > 0) {
                 // Push to textarea so user sees what was found
                 const textArea = document.getElementById('admin-crossref-input');
                 textArea.value = textArea.value + (textArea.value ? '\n' : '') + extracted.join('\n');
                 
                 _performCrossRef(extracted);
            } else {
                 if (statusEl) statusEl.textContent = `ไม่พบเลขพัสดุในรูปภาพ`;
            }
        }
    } catch(err) {
         if (statusEl) statusEl.textContent = `Error OCR: ` + err.message;
    }
    
    // Reset file input
    document.getElementById('admin-crossref-file').value = '';
}

function _performCrossRef(trackingArray) {
    const statusEl = document.getElementById('admin-crossref-status');
    const resultEl = document.getElementById('admin-crossref-result');
    
    if (!trackingArray || trackingArray.length === 0) {
        statusEl.textContent = "กรุณาวางข้อมูล หรือ อัปโหลดรูปภาพที่มีเลขพัสดุก่อน";
        resultEl.classList.add('hidden');
        return;
    }

    statusEl.textContent = `กำลังเทียบข้อมูล ${trackingArray.length} รายการ กับฐานข้อมูลลูกค้า...`;
    
    const lookup = CustomerDB.getLookup();
    const batches = CustomerDB.getBatches();
    
    let html = `
        <table style="width:100%; font-size:0.9rem;">
            <thead>
                <tr style="background:#eee;">
                    <th>ลำดับ</th>
                    <th>เลขพัสดุ</th>
                    <th>สังกัดบริษัทในระบบ (Company / Name)</th>
                </tr>
            </thead>
            <tbody>
    `;

    let foundCount = 0;

    trackingArray.forEach((track, idx) => {
        // ... (existing logic for dbInfoMain, displayList) ...
        let dbInfoMain = lookup[track];
        let isProbableMain = false;

        if (!dbInfoMain && track.length === 13 && typeof CustomerDB.findBookMatch === 'function') {
            const bookMatch = CustomerDB.findBookMatch(track);
            if (bookMatch) {
                dbInfoMain = bookMatch;
                isProbableMain = true;
            }
        }
        
        let displayList = [];
        if (track.length === 13 && typeof TrackingUtils !== 'undefined') {
            displayList = TrackingUtils.generateTrackingRange(track, 2, 1);
        } else {
            displayList = [{ number: track, offset: 0, isCenter: true }];
        }

        // Start Group Container
        html += `<tbody style="border: 1px solid #ddd; border-bottom: 3px solid #ccc; background: white;">`;

        displayList.forEach((item, innerIdx) => {
            let dbInfo = lookup[item.number];
            let companyName = '<span style="color:#ccc; font-style:italic;">ไม่พบข้อมูล (Not Found)</span>';
            let rowStyle = '';
            const isMain = (item.offset === 0);
            
            if (isMain) {
                companyName = '<span style="color:#999; font-style:italic;">ไม่พบข้อมูล (Not Found)</span>';
                if (dbInfoMain) {
                    if (isProbableMain) {
                        companyName = `💡 คาดว่าเป็นของ: <strong style="color:#e65100;">${dbInfoMain.name}</strong> <small style="color:#ef6c00;">(อ้างจากเล่ม)</small>`;
                        rowStyle = 'background-color:#fff3e0; border-left: 5px solid #fb8c00;'; 
                    } else {
                        companyName = `<strong style="color:var(--primary-color); underline">${dbInfoMain.name}</strong>`;
                        if(batches[dbInfoMain.batchId] && batches[dbInfoMain.batchId].requestDate) {
                             companyName += ` <small style="color:#2e7d32; font-weight:bold;"> (ขอเลข: ${new Date(batches[dbInfoMain.batchId].requestDate).toLocaleDateString('th-TH')})</small>`;
                        }
                        rowStyle = 'background-color:#f1f8e9; border-left: 5px solid #43a047;'; 
                        foundCount++;
                    }
                } else {
                    rowStyle = 'background-color: #fffde7; border-left: 5px solid #fbc02d;'; 
                }
            } else {
                rowStyle = 'background-color:#fafafa; color: #888; border-left: 3px solid #eee;'; 
                if (dbInfo) {
                    companyName = `<strong style="color:#666;">${dbInfo.name}</strong>`;
                    if(batches[dbInfo.batchId] && batches[dbInfo.batchId].requestDate) {
                         companyName += ` <small style="color:#999;">(ขอเลข: ${new Date(batches[dbInfo.batchId].requestDate).toLocaleDateString('th-TH')})</small>`;
                    }
                    rowStyle = 'background-color:#f5fcf6; color: #555; border-left: 3px solid #c8e6c9;'; 
                }
            }

            let label = '';
            let trackDisplay = `<span style="font-family:monospace; font-weight:bold; font-size: 1.1em;">${item.number}</span>`;
            let indexCol = '';
            if (!isMain) {
                if (item.offset < 0) label = `(ก่อนหน้า ${Math.abs(item.offset)})`;
                if (item.offset > 0) label = `(ถัดไป ${item.offset})`;
                trackDisplay = `<span style="font-family:monospace; margin-left:15px;">↳ ${item.number} ${label}</span>`;
            } else {
                indexCol = (idx + 1).toString();
            }

            const cleanId = item.number.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            const actionsHtml = `
                <div class="status-actions" style="margin-top:${isMain ? '5px' : '2px'}; margin-bottom: 5px; ${isMain ? '' : 'font-size: 0.8em; opacity: 0.8;'}">
                    <a href="https://track.thailandpost.co.th/?trackNumber=${cleanId}&lang=th" target="_blank" class="badge badge-neutral" style="background-color:#e3f2fd; color:#0d47a1; border-color:#90caf9; ${isMain ? '' : 'padding: 2px 4px;'}" title="ติดตามพัสดุ (External Link)">📌 สถานะ</a>
                    <button class="badge badge-neutral" style="border:1px solid #999; cursor:pointer; ${isMain ? '' : 'padding: 2px 4px;'}" onclick="navigator.clipboard.writeText('${cleanId}').then(() => alert('คัดลอก ${cleanId} แล้ว'))" title="Copy ID">📋 Copy</button>
                </div>
            `;

            html += `
                <tr style="${rowStyle}">
                    <td style="text-align:center; vertical-align: top; padding-top: 15px; border-top: 1px solid #eee;">${indexCol}</td>
                    <td style="vertical-align: top; padding-top: ${isMain ? '15px' : '5px'}; padding-bottom: ${isMain ? '5px' : '5px'}; border-top: 1px solid #eee;">
                        ${trackDisplay}
                        <br>
                        <div style="${isMain ? '' : 'margin-left: 15px;'}">${actionsHtml}</div>
                    </td>
                    <td style="vertical-align: top; padding-top: ${isMain ? '15px' : '5px'}; border-top: 1px solid #eee;">${companyName}</td>
                </tr>
            `;
        });

        // Close Group Container and add spacer row
        html += `
            <tr style="height: 15px; background: transparent;"><td colspan="3" style="border:none;"></td></tr>
        </tbody>`;
    });

    html += `</tbody></table>`;
    
    resultEl.innerHTML = html;
    resultEl.classList.remove('hidden');
    
    statusEl.innerHTML = `ตรวจพบสังกัดตรงกัน <strong style="color:green;">${foundCount}</strong> จากทั้งหมด ${trackingArray.length} รายการ`;
}

// --- Authentication & Isolation System ---

// Initial Run with improved reliability
window.addEventListener('load', function () {
    console.log('Window loaded. Running checkAuth...');
    setTimeout(checkAuth, 100);
});

if (document.readyState === 'complete' || document.readyState === 'interactive') {
    checkAuth();
}

function checkAuth() {
    try {
        console.log('Running checkAuth (Strict Mode)...');
        const urlParams = new URLSearchParams(window.location.search);
        const isAdmin = urlParams.has('admin');

        const header = document.getElementById('main-header');
        const tabNav = document.getElementById('main-tabs');
        const userHeader = document.getElementById('user-mode-header');

        if (isAdmin) {
            // ==========================================
            // ADMIN MODE (FULL ACCESS)
            // ==========================================
            document.body.classList.add('admin-mode');
            document.body.classList.remove('user-mode');
            
            if (header) header.classList.remove('hidden');
            if (tabNav) tabNav.classList.remove('hidden');
            if (userHeader) userHeader.classList.add('hidden');

            // Default Admin tab
            if (!document.querySelector('.tab-btn.active')) switchTab('smart');

            // Admin Upload UI: Full features
            const uploadIcon = document.getElementById('upload-icon-display');
            const uploadTitle = document.getElementById('upload-title-display');
            const uploadDesc = document.getElementById('upload-desc-display');
            const uploadInput = document.getElementById('import-upload');

            if (uploadIcon) uploadIcon.innerText = "📂 / 📸";
            if (uploadTitle) uploadTitle.innerText = "แตะเพื่อเลือกไฟล์ Excel หรือ รูปภาพ";
            if (uploadDesc) uploadDesc.innerText = "รองรับ .xlsx, .xls และ รูปภาพ (OCR)";
            if (uploadInput) uploadInput.accept = ".xlsx, .xls, image/*, .heic";

        } else {
            // ==========================================
            // STAFF MODE (Excel Import ONLY)
            // ==========================================
            document.body.classList.add('user-mode');
            document.body.classList.remove('admin-mode');
            
            if (header) header.classList.add('hidden');
            if (tabNav) tabNav.classList.add('hidden');
            if (userHeader) userHeader.classList.remove('hidden');

            // FORCE hide all tabs and stay on Import
            document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            
            const tabImport = document.getElementById('tab-import');
            if (tabImport) tabImport.classList.add('active');

            // Staff Mode initialization (preserved premium UI)
            const uploadInput = document.getElementById('import-upload');
            if (uploadInput) uploadInput.accept = ".xlsx, .xls";

            // Ensure the main button is labeled correctly for data portability (Backup)
            const saveBtn = document.getElementById('import-save-btn');
            if (saveBtn) {
                saveBtn.innerHTML = `💾 สำรองข้อมูลระบบ (Back Up)`;
                saveBtn.style.background = 'linear-gradient(135deg, #0d47a1, #1565c0)';
                saveBtn.style.color = 'white';
            }
            
            // Re-check and build user header if missing
            if (!userHeader) {
                const newUserHeader = document.createElement('div');
                newUserHeader.id = 'user-mode-header';
                newUserHeader.innerHTML = `
                    <div style="display:flex; justify-content:center; align-items:center;">
                        <span style="font-size:1rem;">📥 ระบบนำเข้าข้อมูลพัสดุ<br><small style="font-weight:normal; font-size:0.8rem;">(Import Data Entry)</small></span>
                    </div>
                `;
                const main = document.querySelector('main');
                if (main && document.body) document.body.insertBefore(newUserHeader, main);
            }
        }
    } catch (e) {
        console.error('Error in checkAuth:', e);
    }
}

// ==========================================
// SECTION: ADMIN TOOLS
// ==========================================

function adminHandleTrackInput(el) {
    el.value = el.value.toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 13);
}

// v1.73: Removed duplicate definition of adminOpenThpTrack

async function adminHandleImageOcr(files) {
    if (!files || files.length === 0) return;
    const statusEl = document.getElementById('admin-ocr-status');
    const resultBox = document.getElementById('admin-ocr-result');
    
    statusEl.innerText = `Preparing OCR for ${files.length} images...`;
    resultBox.classList.add('hidden');
    
    const useEngOnly = document.getElementById('admin-ocr-eng-only')?.checked;
    const lang = useEngOnly ? 'eng' : 'tha+eng';
    
    let combinedText = "";
    try {
        for (let i = 0; i < files.length; i++) {
            statusEl.innerText = `OCR Scanning Image ${i + 1}/${files.length}...`;
            const worker = await Tesseract.createWorker(lang);
            const { data: { text } } = await worker.recognize(files[i]);
            await worker.terminate();
            combinedText += "\n" + text;
        }
        
        statusEl.innerText = "Analyzing extracted data...";
        
        // Extract numbers and prices
        const tableItems = TrackingUtils.extractHandwrittenTable(combinedText);
        
        if(tableItems.length === 0) {
            statusEl.innerText = "ไม่พบข้อมูลที่น่าจะเป็นเลขพัสดุ/ราคาในภาพ";
            return;
        }
        
        // Formulate output string
        let outputStr = `พบข้อมูล ${tableItems.length} รายการ:\n\n`;
        tableItems.forEach((item, idx) => {
             outputStr += `${idx+1}. ${item.number} -> ราคา: ${item.price} บ. (นน. ${item.weight})\n`;
        });
        
        // Check missing if sequence
        const numbersOnly = tableItems.map(t => t.number).sort((a,b) => a.localeCompare(b));
        const parse = (str) => {
            const m = str.match(/([A-Z]{2})(\d{8})(\d)([A-Z]{2})/);
            return m ? { full: str, prefix: m[1], body: parseInt(m[2]), check: m[3], suffix: m[4] } : null;
        };
        
        if (numbersOnly.length > 1) {
             let missing = [];
             let prev = parse(numbersOnly[0]);
             for(let i=1; i<numbersOnly.length; i++){
                  const curr = parse(numbersOnly[i]);
                  if(prev && curr && prev.prefix === curr.prefix && prev.suffix === curr.suffix) {
                       const diff = curr.body - prev.body;
                       if (diff > 1) {
                           for(let m = prev.body + 1; m < curr.body; m++) {
                               let bodyStr = m.toString().padStart(8,'0');
                               let cd = TrackingUtils.calculateS10CheckDigit ? TrackingUtils.calculateS10CheckDigit(bodyStr) : 'X';
                               missing.push(`${curr.prefix}${bodyStr}${cd}${curr.suffix}`);
                           }
                       }
                  }
                  prev = curr;
             }
             if (missing.length > 0) {
                 outputStr += `\n⚠️ แจ้งเตือน: พบรายการที่หายไปในลำดับ (Gaps) ${missing.length} รายการ:\n`;
                 missing.forEach(m => outputStr += `- ${m}\n`);
             } else {
                 outputStr += `\n✅ ลำดับเลขต่อเนื่องกันดี ไม่พบรายการตกหล่น`;
             }
        }
        
        resultBox.innerText = outputStr;
        resultBox.classList.remove('hidden');
        statusEl.innerText = "OCR และการวิเคราะห์เสร็จสิ้น!";
        
        document.getElementById('admin-ocr-upload').value = '';
        
    } catch(err) {
        console.error(err);
        statusEl.innerText = "Error: " + err.message;
    }
}

function adminCrossReference() {
    const inputStr = document.getElementById('admin-crossref-input').value.trim();
    if (!inputStr) {
        alert('กรุณาวางเลขพัสดุ');
        return;
    }
    
    const regex = /[A-Z]{2}\d{9}[A-Z]{2}/ig;
    const matches = inputStr.match(regex);
    if (!matches || matches.length === 0) {
         document.getElementById('admin-crossref-status').innerText = 'ไม่พบรูปแบบเลขพัสดุ 13 หลัก';
         return;
    }
    
    renderCrossReference(matches);
}

async function adminCrossRefImage(files) {
    if (!files || files.length === 0) return;
    const statusEl = document.getElementById('admin-crossref-status');
    statusEl.innerText = `Scanning Image for Cross Reference...`;
    
    try {
        const worker = await Tesseract.createWorker('eng');
        const { data: { text } } = await worker.recognize(files[0]);
        await worker.terminate();
        
        const regex = /[A-Z]{2}\d{9}[A-Z]{2}/ig;
        const matches = text.match(regex);
        if(!matches || matches.length === 0) {
            statusEl.innerText = 'ไม่พบเลขพัสดุในรูปนี้';
            return;
        }
        renderCrossReference(matches);
    } catch(e) {
        statusEl.innerText = 'Error: ' + e.message;
    }
    document.getElementById('admin-crossref-file').value = '';
}

async function renderCrossReference(trackArray) {
    const uniqueTracks = [...new Set(trackArray.map(t => t.toUpperCase()))];
    const statusEl = document.getElementById('admin-crossref-status');
    const resultBox = document.getElementById('admin-crossref-result');
    
    let html = `
        <table style="width:100%; text-align:left; border-collapse: collapse;">
            <thead>
                <tr style="background:#f1f1f1;">
                    <th style="padding:8px; border-bottom:1px solid #ccc;">#</th>
                    <th style="padding:8px; border-bottom:1px solid #ccc;">เลขพัสดุ</th>
                    <th style="padding:8px; border-bottom:1px solid #ccc;">สังกัด (Customer)</th>
                    <th style="padding:8px; border-bottom:1px solid #ccc;">ประเภท</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    let foundCount = 0;
    
    for (let index = 0; index < uniqueTracks.length; index++) {
        const track = uniqueTracks[index];
        let ownerName = '- ไม่พบในระบบ -';
        let ownerType = '-';
        let rowColor = '';
        
        if (typeof CustomerDB !== 'undefined') {
             const owner = await CustomerDB.get(track);
             if (owner) {
                 ownerName = `👤 ${owner.name}`;
                 ownerType = owner.type;
                 rowColor = 'background:#e3f2fd;';
                 foundCount++;
             }
        }
        
        html += `
            <tr style="border-bottom:1px solid #ddd; ${rowColor}">
                <td style="padding:8px;">${index+1}</td>
                <td class="crossref-id" style="padding:8px; font-family:monospace; font-weight:bold;">${TrackingUtils.formatTrackingNumber(track)}</td>
                <td style="padding:8px;">${ownerName}</td>
                <td style="padding:8px;">${ownerType}</td>
            </tr>
        `;
    }
    
    html += `</tbody></table>`;
    
    resultBox.innerHTML = html;
    resultBox.classList.remove('hidden');
    statusEl.innerText = `ตรวจสอบสำเร็จ ${uniqueTracks.length} หมายเลข (พบประวัติ ${foundCount} รายการ)`;
}

function copyCrossRefAll() {
    const ids = Array.from(document.querySelectorAll('.crossref-id')).map(el => el.innerText.replace(/\\s/g, '')).join('\\n');
    if(ids) {
        navigator.clipboard.writeText(ids).then(() => alert('คัดลอกรายการทั้งหมดลง Clipboard แล้ว'));
    } else {
        alert('ไม่มีข้อมูลให้คัดลอก');
    }
}

// ==========================================

// ==========================================
// SECTION: QMS STAGING & IMPORT (Admin)
// ==========================================
let qmsStagingGroups = {}; // Memory to store parsed groups

// =====================================
// DATE HELPER: BE <-> CE conversion
// =====================================

/**
 * Given a BE date string like "19/03/2569 17:03" or "20/03/2569"
 * parse it and set the date+time pickers, updating hidden field and display.
 */
function setExceptionDatePickerFromBE(beStr) {
    if (!beStr) return;
    // Match DD/MM/YYYY HH:MM or DD/MM/YYYY
    const m = beStr.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2}))/);
    if (!m) return;
    const [, dd, mm, yyyy_be, hh, min] = m;
    const yyyy_ce = parseInt(yyyy_be, 10) - 543;
    const dateStr = `${yyyy_ce}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}`;
    const datePicker = document.getElementById('exception-date-picker');
    const timePicker = document.getElementById('exception-time-picker');
    if (datePicker) datePicker.value = dateStr;
    if (timePicker && hh && min) timePicker.value = `${hh.padStart(2,'0')}:${min}`;
    updateBEDisplay();
}

/** Called by oninput on both pickers — updates hidden field and the blue BE label */
function updateBEDisplay() {
    const dateVal = document.getElementById('exception-date-picker')?.value;
    const timeVal = document.getElementById('exception-time-picker')?.value;
    const beDisplay = document.getElementById('exception-be-display');
    const hidden = document.getElementById('exception-datetime');
    if (!dateVal) {
        if (beDisplay) beDisplay.textContent = '';
        if (hidden) hidden.value = '';
        return;
    }
    // Convert CE to BE
    const [yyyy_ce, mm, dd] = dateVal.split('-');
    const yyyy_be = parseInt(yyyy_ce, 10) + 543;
    const beStr = timeVal
        ? `${dd}/${mm}/${yyyy_be} ${timeVal}`
        : `${dd}/${mm}/${yyyy_be}`;
    if (beDisplay) beDisplay.textContent = `📅 ${beStr}`;
    if (hidden) hidden.value = beStr;
}

function toggleQmsStaging() {
    const panel = document.getElementById('qms-staging-panel');
    const chevron = document.getElementById('qms-staging-chevron');
    if (!panel) return;
    const open = panel.style.display !== 'none';
    panel.style.display = open ? 'none' : 'block';
    if (chevron) chevron.textContent = open ? '▼ เปิด' : '▲ ปิด';
}

function processQmsImport() {
    const inputEl = document.getElementById('qms-import-text');
    if (!inputEl) {
        console.warn("[v4.0.0] QMS Import Textarea not found in the current UI.");
        return;
    }
    const text = inputEl.value;
    if (!text) return;

    // Extract tracking numbers and optional datetime using index distance to handle multiline grid pastes
    const trackPattern = /[a-zA-Z]{2}\d{9}[a-zA-Z]{2}/g;
    // Strict numeric date regex: matches DD/MM/YYYY or DD-MM-YYYY, with optional space HH:MM
    // NOTE: Do NOT use a pattern that can match partial tracking numbers (e.g. "04TH EQ0897")
    const datetimePattern = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})(?:[^\d]*(\d{1,2}:\d{2}))?/g;
    
    const trackMap = new Map();
    let hasAnyMatches = false;

    const tracks = [];
    let match;
    while ((match = trackPattern.exec(text)) !== null) {
        tracks.push({ track: match[0].toUpperCase(), index: match.index });
    }

    if (tracks.length > 0) {
        hasAnyMatches = true;
        
        const dates = [];
        let dMatch;
        datetimePattern.lastIndex = 0;
        while ((dMatch = datetimePattern.exec(text)) !== null) {
            let dtStr = dMatch[1];
            if (dMatch[2]) dtStr += ' ' + dMatch[2];
            dates.push({ dtStr: dtStr, index: dMatch.index });
        }
        
        if (dates.length === 0) {
            console.warn("No date found in QMS import text.");
        }
        
        tracks.forEach((t, i) => {
            const nextTrackIndex = (i < tracks.length - 1) ? tracks[i+1].index : text.length;
            let validDate = dates.find(d => d.index > t.index && d.index < nextTrackIndex);
            
            // FALLBACK: If index bounding failed because of weird clipboard formatting, grab the first available date in the document.
            if (!validDate && dates.length > 0) {
                validDate = dates[0];
            }
            
            if (!trackMap.has(t.track)) {
                trackMap.set(t.track, validDate ? validDate.dtStr : '');
            } else if (validDate && !trackMap.get(t.track)) {
                trackMap.set(t.track, validDate.dtStr);
            }
        });
    } else {
        alert('ไม่พบเลขพัสดุรูปแบบ 13 หลักในข้อมูลที่วาง');
        return;
    }

    const uniqueTracks = Array.from(trackMap.keys());
    
    qmsStagingGroups = {};
    const existingExceptions = ExceptionManager.getAll();
    const existingNumSet = new Set(
        existingExceptions.flatMap(e => e.entries ? e.entries.map(ex => ex.trackNum) : [e.trackNum])
    );

    uniqueTracks.forEach(track => {
        // Group by 2 letters + space + 4 digits: e.g. "EQ 0898"
        const groupPrefix = track.substring(0, 2) + " " + track.substring(2, 6);
        const dtStr = trackMap.get(track);
        
        if (!qmsStagingGroups[groupPrefix]) {
            qmsStagingGroups[groupPrefix] = {
                id: 'grp_' + track.substring(0, 6),
                prefix: groupPrefix,
                items: [],
                companyHint: 'กำลังค้นหา...',
                companyFound: false,
                extractedDateTime: dtStr
            };
        } else if (!qmsStagingGroups[groupPrefix].extractedDateTime && dtStr) {
            qmsStagingGroups[groupPrefix].extractedDateTime = dtStr;
        }
        
        const isDuplicate = existingNumSet.has(track);
        qmsStagingGroups[groupPrefix].items.push({ track, isDuplicate });
    });

    // Smart Lookup for each group
    Object.values(qmsStagingGroups).forEach(group => {
        const firstTrack = group.items[0].track;
        const prefix = firstTrack.substring(0, 2);
        const suffix = firstTrack.substring(11, 13);
        const bodyInt = parseInt(firstTrack.substring(2, 10), 10);
        
        let foundCompany = null;

        // Helper to check DB
        const checkDb = (bInt) => {
            if (typeof CustomerDB === 'undefined' || typeof TrackingUtils === 'undefined') return null;
            const numStr = bInt.toString().padStart(8, '0');
            const cd = TrackingUtils.calculateS10CheckDigit(numStr);
            const testNum = `${prefix}${numStr}${cd}${suffix}`;
            const info = CustomerDB.get(testNum);
            return info ? info.name : null;
        };

        // 1. Check exact first track
        foundCompany = checkDb(bodyInt);
        
        if (foundCompany) {
            group.companyHint = foundCompany;
            group.companyFound = true;
        } else {
            // 2. Smart Guess: Check -2, -1, +1
            const guessOffsets = [-1, -2, 1];
            for (let offset of guessOffsets) {
                const hint = checkDb(bodyInt + offset);
                if (hint) {
                    foundCompany = hint;
                    break;
                }
            }
            if (foundCompany) {
                group.companyHint = `(คาดเดาจากเลขใกล้เคียง) ${foundCompany}`;
                group.companyFound = true; // treat as found for UI coloring
            } else {
                group.companyHint = 'ไม่พบข้อมูลบริษัทในฐานข้อมูล';
            }
        }
    });

    renderQmsGroups();
}

function renderQmsGroups() {
    const container = document.getElementById('qms-staging-results');
    if (!container) return;

    let html = '';
    const groups = Object.values(qmsStagingGroups);
    
    if (groups.length === 0) {
        container.innerHTML = '<div style="color:#666; font-size:0.85rem;">ไม่พบข้อมูลหลังจัดกลุ่ม</div>';
        return;
    }

    // Sort groups by size descending
    groups.sort((a,b) => b.items.length - a.items.length);

    groups.forEach(g => {
        const dupCount = g.items.filter(i => i.isDuplicate).length;
        const dupWarning = dupCount > 0 ? `<span style="color:#d32f2f; font-size:0.75rem; margin-left:8px; background:#ffebee; padding:2px 6px; border-radius:10px; font-weight:bold;">⚠️ ซ้ำ/ประวัติเดิม ${dupCount} รายการ</span>` : '';
        const bgHint = g.companyFound ? '#e8f5e9' : '#fff';
        const borderHint = g.companyFound ? '#81c784' : '#ccc';
        
        const timeHint = g.extractedDateTime 
            ? `<div style="font-size:0.85rem; color:#0288d1; margin-top:6px;">🕒 เวลาสแกน (QMS): ${g.extractedDateTime}</div>`
            : `<div style="font-size:0.8rem; color:#f57c00; margin-top:6px;">⚠️ ไม่พบข้อมูลเวลาสแกนในข้อมูลที่วาง</div>`;

        html += `
            <div style="border:1px solid ${borderHint}; border-radius:6px; padding:12px; margin-bottom:12px; background:${bgHint}; box-shadow:0 1px 3px rgba(0,0,0,0.05);">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:10px; flex-wrap:wrap; gap:8px;">
                    <div>
                        <div style="display:flex; align-items:center;">
                            <strong style="font-size:1.1rem; color:#1565c0;">กลุ่ม ${g.prefix}</strong>
                            <span style="font-size:0.85rem; color:#555; margin-left:8px; padding:2px 6px; background:#e0e0e0; border-radius:10px;">${g.items.length} รายการ</span>
                            ${dupWarning}
                        </div>
                        <div style="font-size:0.9rem; color:#333; margin-top:6px;">🏢 <strong>บริษัท:</strong> ${g.companyHint}</div>
                        ${timeHint}
                    </div>
                    <button class="btn btn-primary" style="font-size:0.85rem; padding:6px 12px; background-color:#0288d1; border-color:#0288d1;" onclick="draftReportFromGroup('${g.prefix}')">
                        ➕ นำกลุ่มนี้ไปสร้างรายงาน
                    </button>
                </div>
                <div style="display:flex; flex-wrap:wrap; gap:6px; font-family:monospace; font-size:0.85rem;">
                    ${g.items.map(i => `<span style="padding:4px 6px; background:${i.isDuplicate ? '#ffebee' : '#faebd7'}; border:1px solid ${i.isDuplicate ? '#ffcdd2' : '#e0e0e0'}; border-radius:4px; color:${i.isDuplicate ? '#c62828' : '#333'}">${i.track}</span>`).join('')}
                </div>
            </div>
        `;
    });

    container.innerHTML = html;
}

function draftReportFromGroup(prefix) {
    const group = qmsStagingGroups[prefix];
    if (!group) return;

    // Open Exception Single Mode if Range is open
    const rangeToggle = document.getElementById('exception-range-toggle');
    if (rangeToggle && rangeToggle.checked) {
        rangeToggle.checked = false;
        if (typeof toggleExceptionRangeMode === 'function') toggleExceptionRangeMode();
    }

    // Clear existing
    const mainInput = document.getElementById('exception-track-input');
    if (mainInput) mainInput.value = '';
    
    document.getElementById('exception-extra-items').innerHTML = '';
    
    // Attempt to reset global counter
    if (typeof window.extraItemCount !== 'undefined') {
        window.extraItemCount = 0; 
    }

    const items = group.items.map(i => i.track);
    
    if (items.length > 0) {
    // NEW: Auto-fill Date/Time from QMS into the new pickers
    if (group.extractedDateTime) {
        setExceptionDatePickerFromBE(group.extractedDateTime);
    }
    const fsInput = document.getElementById('exception-first-status');
    if (fsInput) fsInput.value = 'ใส่ของลงถุง'; // Enforce default

    // Scroll down to the form smoothly
    if (mainInput) {
        mainInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
        mainInput.style.transition = "background-color 0.5s";
        mainInput.style.backgroundColor = "#fff9c4";
        setTimeout(() => mainInput.style.backgroundColor = "", 1500);
    }
    }
}

/**
 * Pre-fills or directly moves items into the exception report draft.
 * v1.79: Implements 'Move' logic - items added to report are hidden from results.
 */
async function stagingQuickReport(tracks, companyName, metadata = {}) {
    if (!tracks) return;
    let inputTracks = Array.isArray(tracks) ? tracks : [tracks];
    
    // 1. Convert Range Titles (e.g. "STR ถึง END") to individual numbers for processing
    if (inputTracks.length === 1 && inputTracks[0].includes(' ถึง ')) {
        const parts = inputTracks[0].split(' ถึง ');
        if (parts.length === 2) {
            const startStr = parts[0].replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            const endStr = parts[1].replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            const startP = parseExceptionTrackNum(startStr);
            const endP = parseExceptionTrackNum(endStr);
            if (startP && endP && startP.prefix === endP.prefix && startP.suffix === endP.suffix) {
                inputTracks = [];
                for (let i = startP.bodyInt; i <= endP.bodyInt; i++) {
                    inputTracks.push(buildExceptionTrackNum(startP.prefix, i, startP.suffix));
                }
            } else {
                inputTracks = [startStr, endStr];
            }
        }
    }

    const sanitizedTracks = inputTracks.map(t => t.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, ''));

    // Populate form instead of direct save (per user request v1.80)
    const mainInput = document.getElementById('exception-track-input');
    const startInput = document.getElementById('exception-start-input');
    const endInput = document.getElementById('exception-end-input');
    const rangeToggle = document.getElementById('exception-range-toggle');
    const reasonInput = document.getElementById('exception-reason-input');
    const datePicker = document.getElementById('exception-date-picker');
    const timePicker = document.getElementById('exception-time-picker');
    const statusInput = document.getElementById('exception-first-status');

    if (sanitizedTracks.length > 0) {
        let existingTracks = [];
        const isCurrentRange = rangeToggle && rangeToggle.checked;
        
        // 1. Collect Existing
        if (isCurrentRange) {
            const sV = startInput.value.trim().replace(/\s/g,'');
            const eV = endInput.value.trim().replace(/\s/g,'');
            if (sV && eV) {
                const sP = parseExceptionTrackNum(sV);
                const eP = parseExceptionTrackNum(eV);
                if (sP && eP && sP.prefix === eP.prefix && sP.suffix === eP.suffix) {
                    for (let i = sP.bodyInt; i <= eP.bodyInt; i++) {
                        existingTracks.push(buildExceptionTrackNum(sP.prefix, i, sP.suffix));
                    }
                }
            }
        } else {
            const mV = mainInput.value.trim();
            if (mV) existingTracks = mV.split(/[\s,]+/).filter(v => v.length > 0);
        }

        // 2. Merge & Deduplicate
        const combined = [...existingTracks];
        sanitizedTracks.forEach(t => {
            if (!combined.includes(t)) combined.push(t);
        });

        // 3. Determine Mode (Auto-Switch if append creates non-consecutive or multiple sets)
        const isContinuousRange = isConsecutive(combined);
        
        if (isContinuousRange && combined.length > 1) {
            if (rangeToggle && !rangeToggle.checked) {
                rangeToggle.checked = true;
                if (typeof toggleExceptionRangeMode === 'function') toggleExceptionRangeMode();
            }
            if (startInput) startInput.value = combined[0];
            if (endInput) endInput.value = combined[combined.length - 1];
        } else {
            // Single or Multiple non-consecutive
            if (rangeToggle && rangeToggle.checked) {
                rangeToggle.checked = false;
                if (typeof toggleExceptionRangeMode === 'function') toggleExceptionRangeMode();
            }
            if (mainInput) mainInput.value = combined.join('\n');
        }

        // 4. Fill Metadata (Only if empty or if meaningful)
        const currentReason = reasonInput ? reasonInput.value.trim() : "";
        const isNewGenericReason = metadata.status && metadata.status.includes('กลุ่มเลขต่อเนื่อง');
        
        if (reasonInput && (!currentReason || currentReason === "รายละเอียดยังไม่เข้าระบบ/ของยังไม่มาส่ง")) {
            // If the incoming metadata is a generic summary row status, maybe don't use it if we are merging
            if (!isNewGenericReason || combined.length === sanitizedTracks.length) {
                reasonInput.value = metadata.status || "รายละเอียดยังไม่เข้าระบบ/ของยังไม่มาส่ง";
            }
        }
        
        if (statusInput && !statusInput.value) statusInput.value = "ใส่ของลงถุง";
        
        if (metadata.datetime && datePicker && !datePicker.value) {
            const dtMatch = metadata.datetime.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}:\d{2})/);
            if (dtMatch) {
                const day = dtMatch[1];
                const month = dtMatch[2];
                let year = parseInt(dtMatch[3]);
                if (year > 2500) year -= 543;
                const time = dtMatch[4];
                if (datePicker) datePicker.value = `${year}-${month.padStart(2,'0')}-${day.padStart(2,'0')}`;
                if (timePicker) timePicker.value = time;
                if (typeof updateBEDisplay === 'function') updateBEDisplay();
            }
        }

        // 5. Scroll & Highlight
        const formSec = document.getElementById('exception-section');
        if (formSec) {
            formSec.scrollIntoView({ behavior: 'smooth', block: 'center' });
            formSec.style.boxShadow = "0 0 15px rgba(2, 136, 209, 0.4)";
            setTimeout(() => formSec.style.boxShadow = "", 2000);
        }

        // 6. Refresh Results Search Panel (v1.91 Sync)
        if (typeof renderStoredUnifiedNumbers === 'function' && currentUnifiedResults && currentUnifiedResults.length > 0) {
            sanitizedTracks.forEach(t => {
                const item = currentUnifiedResults.find(i => i.number.replace(/\s/g,'') === t);
                if (item) item.moved = true;
            });
            localStorage.setItem('thp_last_search_results', JSON.stringify(currentUnifiedResults));
            renderStoredUnifiedNumbers(currentUnifiedTitle || "", currentUnifiedResults);
        }

        window.showToast(`เพิ่มแล้ว! รวมทั้งหมด ${combined.length} รายการ`, 'info');
        return;
    }
}

// SECTION: EXCEPTION META & IMAGE ATTACHMENT
// ==========================================

const EXCEPTION_META_KEY = 'thp_exception_meta_v1';

/** Persist and restore meta fields */
function saveExceptionMeta() {
    const meta = {
        branch:   (document.getElementById('rpt-branch')?.value   || ''),
        date:     (document.getElementById('rpt-date')?.value     || ''),
        reporter: (document.getElementById('rpt-reporter')?.value || ''),
        subject:  (document.getElementById('rpt-subject')?.value  || ''),
        note:     (document.getElementById('rpt-note')?.value     || '')
    };
    localStorage.setItem(EXCEPTION_META_KEY, JSON.stringify(meta));
    
    // Update summary hint in index.html
    const hint = document.getElementById('meta-summary-reporter');
    if (hint) {
        hint.textContent = meta.reporter || '-';
    }
}

/** Get current meta from DOM */
function getExceptionMeta() {
    return {
        branch:   (document.getElementById('rpt-branch')?.value   || ''),
        date:     (document.getElementById('rpt-date')?.value     || ''),
        reporter: (document.getElementById('rpt-reporter')?.value || ''),
        subject:  (document.getElementById('rpt-subject')?.value  || ''),
        note:     (document.getElementById('rpt-note')?.value     || '')
    };
}

function loadExceptionMeta() {
    try {
        const raw = localStorage.getItem(EXCEPTION_META_KEY);
        const meta = raw ? JSON.parse(raw) : {};
        
        const b = document.getElementById('rpt-branch');
        const r = document.getElementById('rpt-reporter');
        const s = document.getElementById('rpt-subject');
        const n = document.getElementById('rpt-note');
        const d = document.getElementById('rpt-date');

        if (b) b.value = meta.branch || '';
        if (r) r.value = meta.reporter || '';
        if (s) s.value = meta.subject || '';
        if (n) n.value = meta.note || '';
        
        if (d) {
            const now = new Date();
            const today = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
            if (!d.value || d.value === "") {
                d.value = today;
            }
        }
        
        // Initial hint update
        const hint = document.getElementById('meta-summary-reporter');
        if (hint) {
            hint.textContent = meta.reporter || '-';
        }

        // Final save to persist any defaults (like today's date)
        saveExceptionMeta();
    } catch(e) { 
        console.error("loadExceptionMeta error:", e);
    }
}

/** Image attachment storage (runtime only — not persisted across reloads) */
let exceptionImages = []; // Array of { dataUrl, name }
let currentEditingSessionId = null; // Track if we are editing an existing history entry

/**
 * Compress image before storing to localStorage to avoid QuotaExceededError (5MB limit).
 */
function compressExceptionImage(dataUrl, maxWidth = 1200, quality = 0.8) {
    if (typeof TrackingUtils !== 'undefined' && typeof TrackingUtils.compressImage === 'function') {
        return TrackingUtils.compressImage(dataUrl, maxWidth, quality);
    }
    // Fallback if utils not loaded
    return Promise.resolve(dataUrl);
}

function handleExceptionImageUpload(files) {
    if (!files || files.length === 0) return;
    const preview = document.getElementById('exception-img-preview');
    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = async function(e) {
            const rawDataUrl = e.target.result;
            
            // Compress BEFORE pushing to global array
            const dataUrl = await compressExceptionImage(rawDataUrl);
            
            const idx = exceptionImages.length;
            exceptionImages.push({ dataUrl, name: file.name });

            // Create preview thumb
            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'position:relative; display:inline-block;';
            wrapper.id = `exc-img-${idx}`;

            const img = document.createElement('img');
            img.src = dataUrl;
            img.style.cssText = 'max-width:120px; max-height:90px; border-radius:4px; border:1px solid #ccc; display:block;';
            img.title = file.name;

            const del = document.createElement('button');
            del.textContent = '✕';
            del.style.cssText = 'position:absolute; top:-5px; right:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#d32f2f; color:white; font-size:0.7rem; cursor:pointer; line-height:1; padding:0;';
            del.onclick = function(event) {
                const currentId = parseInt(event.target.parentElement.id.replace('exc-img-', ''), 10);
                exceptionImages.splice(currentId, 1);
                wrapper.remove();
                const allWrappers = document.getElementById('exception-img-preview').children;
                Array.from(allWrappers).forEach((w, i) => { w.id = `exc-img-${i}`; });
            };

            const editBtn = document.createElement('button');
            editBtn.textContent = '✏️';
            editBtn.style.cssText = 'position:absolute; top:-5px; left:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#0288d1; color:white; font-size:0.6rem; cursor:pointer; line-height:1; padding:0;';
            editBtn.onclick = function(event) {
                const currentId = parseInt(event.target.parentElement.id.replace('exc-img-', ''), 10);
                openImageEditor(currentId);
            };

            wrapper.appendChild(img);
            wrapper.appendChild(del);
            wrapper.appendChild(editBtn);
            preview.appendChild(wrapper);
        };
        reader.readAsDataURL(file);
    });
    // Auto-open panel after first image
    const panel = document.getElementById('exception-img-panel');
    const chevron = document.getElementById('exception-img-chevron');
    if (panel && panel.style.display === 'none') {
        panel.style.display = 'block';
        if (chevron) chevron.textContent = '▲ ซ่อน';
    }
    // Reset input so same file can be re-added
    document.getElementById('exception-img-upload').value = '';
}

function clearExceptionImages() {
    exceptionImages = [];
    const preview = document.getElementById('exception-img-preview');
    if (preview) preview.innerHTML = '';
}

function updateBEDisplay() {
    const dp = document.getElementById('exception-date-picker');
    const tp = document.getElementById('exception-time-picker');
    const hidden = document.getElementById('exception-datetime');
    if (!dp || !tp || !hidden) return;
    hidden.value = `${dp.value}T${tp.value}`;
}

function setExceptionDatePickerFromBE(dateTimeStr) {
    const dp = document.getElementById('exception-date-picker');
    const tp = document.getElementById('exception-time-picker');
    const hidden = document.getElementById('exception-datetime');
    if (!dp || !tp || !hidden) return;
    const parts = dateTimeStr.split('T');
    if (parts.length === 2) {
        dp.value = parts[0];
        tp.value = parts[1];
        hidden.value = dateTimeStr;
    }
}

// ==========================================
// SECTION: EXCEPTION LOG (ตกหล่น)
// ==========================================

/**
 * Toggle between single-entry and range mode for Exception form.
 */
function toggleExceptionRangeMode() {
    const isRange = document.getElementById('exception-range-toggle').checked;
    document.getElementById('exception-single-mode').style.display = isRange ? 'none' : 'flex';
    document.getElementById('exception-range-mode').style.display = isRange ? 'block' : 'none';
    document.getElementById('exception-range-preview').textContent = '';
}

/**
 * Add an extra (non-consecutive) tracking number row below the main form.
 */
let extraItemCount = 0;
function addExceptionExtraItem() {
    const container = document.getElementById('exception-extra-items');
    extraItemCount++;
    const div = document.createElement('div');
    div.id = `extra-item-${extraItemCount}`;
    div.style.cssText = 'display:flex; gap:10px; margin-top:10px; align-items:center;';
    div.innerHTML = `
        <input type="text" class="exception-extra-track" placeholder="เลขที่เพิ่มเติม (เช่น EQ123499999TH)" maxlength="13"
            oninput="checkDbWarningForReport(this); if(typeof renderStoredUnifiedNumbers === 'function') renderStoredUnifiedNumbers(currentUnifiedTitle || '', currentUnifiedResults || []);"
            style="flex:1; text-transform:uppercase; padding:10px; border:1px solid #ccc; border-radius:4px; box-sizing:border-box; font-family:inherit; font-size:inherit;">
        <button type="button" onclick="document.getElementById('extra-item-${extraItemCount}').remove(); if(typeof renderStoredUnifiedNumbers === 'function') renderStoredUnifiedNumbers(currentUnifiedTitle || '', currentUnifiedResults || []);"
            style="padding:10px 14px; border:1px solid #ffcdd2; border-radius:4px; background:#ffebee; color:#d32f2f; cursor:pointer;">✕</button>
    `;
    container.appendChild(div);
}

/**
 * Parse a tracking number string into its components.
 */
function parseExceptionTrackNum(str) {
    str = str.trim().toUpperCase().replace(/\s+/g, '');
    const m = str.match(/^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/);
    if (!m) return null;
    return { prefix: m[1], body: m[2], cd: m[3], suffix: m[4], bodyInt: parseInt(m[2]), full: str };
}

/**
 * Build a full tracking number with auto check digit.
 */
function buildExceptionTrackNum(prefix, bodyInt, suffix) {
    const bodyStr = bodyInt.toString().padStart(8, '0');
    const cd = TrackingUtils.calculateS10CheckDigit(bodyStr);
    return `${prefix}${bodyStr}${cd}${suffix}`;
}

/**
 * Main save: collects all inputs and saves as one session.
 */
async function addExceptionEntry() {
    const isRange = document.getElementById('exception-range-toggle').checked;
    const reason = document.getElementById('exception-reason-input').value.trim();
    const category = document.getElementById('exception-category').value;

    if (!reason) {
        alert('กรุณาระบุเหตุผลที่ตกหล่น');
        document.getElementById('exception-reason-input').focus();
        return;
    }

    const trackNums = [];

    if (isRange) {
        const startRaw = document.getElementById('exception-start-input').value.trim().toUpperCase().replace(/\s+/g, '');
        const endRaw   = document.getElementById('exception-end-input').value.trim().toUpperCase().replace(/\s+/g, '');

        const startParsed = parseExceptionTrackNum(startRaw);
        const endParsed   = parseExceptionTrackNum(endRaw);

        if (!startParsed) { alert('เลขเริ่มต้นไม่ถูกต้อง'); return; }
        if (!endParsed)   { alert('เลขสุดท้ายไม่ถูกต้อง'); return; }
        if (startParsed.prefix !== endParsed.prefix || startParsed.suffix !== endParsed.suffix) {
            alert('เลขเริ่มต้นและสุดท้ายต้องเป็นชุดเดียวกัน (Prefix/Suffix ตรงกัน)');
            return;
        }
        if (startParsed.bodyInt > endParsed.bodyInt) {
            alert('เลขเริ่มต้นต้องน้อยกว่าหรือเท่ากับเลขสุดท้าย');
            return;
        }

        const count = endParsed.bodyInt - startParsed.bodyInt + 1;
        if (count > 500 && !confirm(`คุณกำลังบันทึก ${count} รายการ ยืนยันหรือไม่?`)) return;

        for (let i = startParsed.bodyInt; i <= endParsed.bodyInt; i++) {
            trackNums.push(buildExceptionTrackNum(startParsed.prefix, i, startParsed.suffix));
        }
    } else {
        const rawInput = document.getElementById('exception-track-input').value.trim().toUpperCase();
        // Split by comma, space, or newline
        let inputs = rawInput.split(/[\s,]+/).filter(v => v.length > 3);
        // Deduplicate and fix common OCR typos
        inputs = [...new Set(inputs.map(v => v.replace(/O/g, '0').replace(/I/g, '1')))]; 
        
        
        if (inputs.length === 0) {
            alert('กรุณากรอกเลขพัสดุ');
            document.getElementById('exception-track-input').focus();
            return;
        }

        const validTracks = inputs.filter(v => v.length === 13);
        if (validTracks.length === 0) {
            alert('กรุณากรอกเลขพัสดุให้ครบ 13 หลัก (พบข้อมูลแต่รูปแบบไม่ถูกต้อง)');
            document.getElementById('exception-track-input').focus();
            return;
        }

        validTracks.forEach(v => {
            if (!trackNums.includes(v)) trackNums.push(v);
        });
    }

    // Extra Items
    document.querySelectorAll('.exception-extra-track').forEach(inp => {
        const v = inp.value.trim().toUpperCase().replace(/\s+/g, '');
        if (v && v.length === 13 && !trackNums.includes(v)) trackNums.push(v);
    });

    if (trackNums.length === 0) { alert('ไม่พบเลขพัสดุที่ถูกต้อง'); return; }

    let companyName = '-';
    if (typeof CustomerDB !== 'undefined') {
        const info = await CustomerDB.get(trackNums[0]);
        if (info) companyName = info.name;
    }

    const firstStatus = document.getElementById('exception-first-status').value.trim();
    const dateTime = document.getElementById('exception-datetime').value.trim();
    
    const metadata = {
        category: category,
        branch:   document.getElementById('rpt-branch').value.trim(),
        reporter: document.getElementById('rpt-reporter').value.trim(),
        subject:  document.getElementById('rpt-subject').value.trim(),
        note:     document.getElementById('rpt-note').value.trim()
    };

    // Pass images and current editing ID to save
    const btn = document.getElementById('exception-save-btn');
    const originalHtml = btn ? btn.innerHTML : 'บันทึก (Save)';
    if (btn) window.setButtonLoading(btn, true);

    // v1.64: ASYNC SAVE
    const savedId = await ExceptionManager.saveSession(trackNums, companyName, reason, firstStatus, dateTime, exceptionImages, currentEditingSessionId, metadata);
    
    if (btn) window.setButtonLoading(btn, false, originalHtml);

    if (!savedId) return;
    
    window.showToast(`บันทึกเรียบร้อย! (${trackNums.length} รายการ)`);
    
    // v1.75: Refresh table
    await renderExceptionTable();
    
    // Clear state
    currentEditingSessionId = null;
    clearExceptionImages();

    // Reset inputs
    document.getElementById('exception-track-input').value = '';
    document.getElementById('exception-start-input').value = '';
    document.getElementById('exception-end-input').value = '';
    document.getElementById('exception-range-preview').textContent = '';
    document.getElementById('exception-extra-items').innerHTML = '';
    
    // Reset date/time pickers to empty
    const dpicker = document.getElementById('exception-date-picker');
    const tpicker = document.getElementById('exception-time-picker');
    if (dpicker) dpicker.value = '';
    if (tpicker) tpicker.value = '';
    updateBEDisplay();
    extraItemCount = 0;

    // v2.0-stable: Auto-Cleanup Search Results after Exception Save
    if (typeof renderStoredUnifiedNumbers === 'function' && currentUnifiedResults && currentUnifiedResults.length > 0) {
        console.info(`[v2.0-stable] Auto-cleaning ${trackNums.length} saved items from search results...`);
        const savedSet = new Set(trackNums.map(t => t.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '')));
        
        currentUnifiedResults = currentUnifiedResults.filter(item => {
            const cleanNum = item.number.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            return !savedSet.has(cleanNum);
        });

        // Persist (v1.97 migration)
        await saveLastSearchResults(currentUnifiedTitle || "", currentUnifiedResults, false);
        
        // Re-render sidebar
        renderStoredUnifiedNumbers(currentUnifiedTitle || "", currentUnifiedResults);
    }

    await renderExceptionTable();

    // Scroll to the new entry
    const container = document.getElementById('exception-table-container');
    if (container) {
        container.scrollIntoView({ behavior: 'smooth', block: 'end' });
    }

    // Visual feedback
    const saveBtn = document.getElementById('exception-save-btn');
    if (saveBtn) {
        saveBtn.innerHTML = "✅ บันทึกรายการสำเร็จ!";
        saveBtn.style.background = "linear-gradient(135deg,#2e7d32,#388e3c)"; 
        saveBtn.disabled = true;
        setTimeout(() => {
            saveBtn.innerHTML = '➕ เพิ่มรายการลงในรายงาน (Add to Draft)';
            saveBtn.style.background = "linear-gradient(135deg,#0288d1,#0277bd)";
            saveBtn.disabled = false;
        }, 1500);
    }
}

/**
 * Filter state for Draft Reports
 */
let exceptionFilterMode = 'today'; // 'today' or 'history'

function setExceptionFilter(mode) {
    exceptionFilterMode = mode;
    
    // UI Update
    const bToday = document.getElementById('filter-today');
    const bHistory = document.getElementById('filter-history');
    
    if (bToday && bHistory) {
        if (mode === 'today') {
            bToday.classList.add('active');
            bToday.style.background = '#fff';
            bToday.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
            bToday.style.color = '#333';
            
            bHistory.classList.remove('active');
            bHistory.style.background = 'none';
            bHistory.style.boxShadow = 'none';
            bHistory.style.color = '#666';
        } else {
            bHistory.classList.add('active');
            bHistory.style.background = '#fff';
            bHistory.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
            bHistory.style.color = '#333';
            
            bToday.classList.remove('active');
            bToday.style.background = 'none';
            bToday.style.boxShadow = 'none';
            bToday.style.color = '#666';
        }
    }
    
    renderExceptionTable();
}

/**
 * Clear only items visible in the current filter
 */
async function clearFilteredExceptions() {
    if (!confirm("ยืนยันการล้างรายการที่แสดงอยู่หรือไม่?")) return;
    const all = await ExceptionManager.getAll();
    const today = new Date().toISOString().split('T')[0];
    
    let toKeep = [];
    if (exceptionFilterMode === 'today') {
        toKeep = all.filter(item => {
            const itemDate = new Date(item.timestamp).toISOString().split('T')[0];
            return itemDate !== today;
        });
    } else {
        toKeep = all.filter(item => {
            const itemDate = new Date(item.timestamp).toISOString().split('T')[0];
            return itemDate === today;
        });
    }
    
    await StorageV2.set(EXCEPTION_KEY, toKeep);
    await renderExceptionTable();

    // v1.91 Sync
    if (typeof renderStoredUnifiedNumbers === 'function' && currentUnifiedResults && currentUnifiedResults.length > 0) {
        // Re-check everything since some history has been wiped
        const allExceptions = await ExceptionManager.getAll();
        currentUnifiedResults.forEach(item => {
            const trackClean = item.number.replace(/\s/g,'');
            const found = allExceptions.find(e => e.trackNum === trackClean);
            if (found) {
                item.history = { date: found.timestamp, company: found.companyName, reason: found.reason, sessionId: found.sessionId || found.id };
            } else {
                item.history = null;
            }
        });
        localStorage.setItem('thp_last_search_results', JSON.stringify(currentUnifiedResults));
        renderStoredUnifiedNumbers(currentUnifiedTitle || "", currentUnifiedResults);
    }
}

function validateCheckDigitUI(trackNum) {
    if (!trackNum || trackNum.length !== 13) return true;
    const body = trackNum.substring(2, 10);
    const cd = trackNum.substring(10, 11);
    const expected = TrackingUtils.calculateS10CheckDigit(body);
    return cd === String(expected);
}

/**
 * Render the current draft report in a clean, card-based list.
 */
async function renderExceptionTable() {
    const container = document.getElementById('exception-table-container');
    const exportBar = document.getElementById('exception-export-bar');
    if (!container) return;

    let exceptions = (typeof ExceptionManager !== 'undefined') ? await ExceptionManager.getAll() : [];
    
    // Apply Filter
    const todayStr = new Date().toISOString().split('T')[0];
    if (exceptionFilterMode === 'today') {
        exceptions = exceptions.filter(item => {
            const d = new Date(item.timestamp).toISOString().split('T')[0];
            return d === todayStr;
        });
    } else {
        exceptions = exceptions.filter(item => {
            const d = new Date(item.timestamp).toISOString().split('T')[0];
            return d !== todayStr;
        });
    }

    // Toggle Export Bar visibility
    if (exportBar) {
        if (exceptions.length > 0) exportBar.classList.remove('hidden');
        else exportBar.classList.add('hidden');
    }

    if (exceptions.length === 0) {
        const msg = exceptionFilterMode === 'today' ? '📭 ยังไม่มีรายการของวันนี้' : '📭 ไม่มีประวัติรายการเก่า';
        container.innerHTML = `
            <div style="text-align:center; padding:40px 20px; color:#999; border:2px dashed #eee; border-radius:12px; background:#fafafa;">
                <p style="margin:0; font-size:1rem;">${msg}</p>
                <p style="margin:5px 0 0 0; font-size:0.85rem;">ระบุเลขด้านบนเพื่อเริ่มรายงานใหม่ครับ</p>
            </div>
        `;
        return;
    }

    // 1. Group by sessionId into Sessions
    const sessionMap = new Map();
    exceptions.forEach(item => {
        const sid = item.sessionId || item.id;
        if (!sessionMap.has(sid)) {
            sessionMap.set(sid, {
                sessionId: sid,
                companyName: item.companyName || 'ทั่วไป',
                reason: item.reason,
                timestamp: item.timestamp,
                entries: []
            });
        }
        sessionMap.get(sid).entries.push(item);
    });

    const sessions = Array.from(sessionMap.values()).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    // 2. Group Sessions by Company Name for "Independent Selection"
    const companyGroups = new Map();
    sessions.forEach(sess => {
        const name = sess.companyName;
        if (!companyGroups.has(name)) companyGroups.set(name, []);
        companyGroups.get(name).push(sess);
    });

    const sortedCompanyNames = Array.from(companyGroups.keys()).sort((a,b) => {
        if (a === 'ทั่วไป') return 1;
        if (b === 'ทั่วไป') return -1;
        return a.localeCompare(b);
    });

    let html = `
        <div class="rpt-header-navy" style="margin-bottom:12px; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
                <input type="checkbox" id="sess-select-all" checked onclick="toggleAllSessions(this.checked)">
                <span id="sess-master-label" style="font-weight:bold;">รายการดราฟต์ ทั้งหมด</span>
            </div>
            <div style="display:flex; gap:8px; align-items:center;">
                <button class="btn btn-neutral" style="padding:2px 8px; font-size:0.7rem; background:#fff; border:1px solid #ccc; border-radius:4px; font-weight:bold; color:#666;" onclick="toggleAllSessions(false)">ยกเลิกทั้งหมด</button>
                <div style="font-size:0.8rem; font-weight:normal; opacity:0.9;">
                    ${exceptionFilterMode === 'today' ? '📅 วันนี้' : '📜 ประวัติ'}
                </div>
            </div>
        </div>
        <div style="display:flex; flex-direction:column; gap:15px;">
    `;

    sortedCompanyNames.forEach((companyName, companyIdx) => {
        const companySessions = companyGroups.get(companyName);
        const companyId = `comp-${companyIdx}`;
        // v1.95: Pre-collect SIDs for the group export button
        const companySessionSids = companySessions.map(s => s.sessionId);
        const sidsJson = JSON.stringify(companySessionSids).replace(/"/g, '&quot;');

        html += `
            <div class="company-report-group" id="group-${companyId}" style="border:1px solid #ddd; border-radius:12px; overflow:hidden; background:#fff; box-shadow:0 3px 6px rgba(0,0,0,0.04);">
                <div style="background:#eef6ff; padding:10px 15px; border-bottom:1px solid #d0e1f9; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                    <div style="display:flex; align-items:center; gap:10px; flex:1; min-width:200px;">
                        <input type="checkbox" class="comp-select" data-comp="${companyId}" checked onclick="toggleSessionsInCompany('${companyId}', this.checked)" style="width:20px; height:20px; cursor:pointer;">
                        <strong style="color:#004ba0; font-size:1.05rem;">🏢 ${companyName}</strong>
                        <span id="label-${companyId}" style="font-size:0.82rem; color:#0288d1; font-weight:bold; margin-left:5px;"></span>
                    </div>
                    <button class="btn btn-primary" style="padding:5px 15px; font-size:0.8rem; border-radius:30px; background:linear-gradient(135deg,#0288d1,#01579b); border:none; box-shadow:0 2px 4px rgba(2,136,209,0.2);" onclick="exportExceptionImage(${sidsJson})">
                        📸 สร้างรายงาน บ.นี้
                    </button>
                </div>
                <div style="display:flex; flex-direction:column;">
        `;

        companySessions.forEach((session, idx) => {
            const firstEntry = session.entries[0];
            const compressed = compressEntriesForDisplay(session.entries || []);
            const dObj = new Date(session.timestamp);
            const timeStr = dObj.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
            
            // Alternating background shades within group
            const groupBg = (idx % 2 === 1) ? '#fafafa' : '#ffffff';

            let images = [];
            for (const entry of session.entries) {
                if (entry.images && entry.images.length > 0) { images = entry.images; break; }
            }
            const hasImages = images.length > 0;
            const totalCount = session.entries.length;
            const isUnknownComp = !session.companyName || session.companyName === '-' || session.companyName === 'Unknown';

            const invalidTracks = session.entries.filter(e => !validateCheckDigitUI(e.trackNum)).map(e => e.trackNum);
            const cdWarning = invalidTracks.length > 0 
                ? `<div class="badge-invalid" style="margin-top:6px;">⚠️ Check Digit ผิด: ${invalidTracks.join(', ')}</div>` 
                : '';

            html += `
                <div class="report-card individual-session" style="background:${groupBg}; border-top:1px solid #f0f0f0; padding:12px 15px;">
                    <div style="display:flex; gap:12px; align-items:flex-start;">
                        <input type="checkbox" class="sess-select" data-comp="${companyId}" value="${session.sessionId}" checked style="margin-top:5px; transform:scale(1.1); cursor:pointer;" onclick="updateSelectAllState()">
                        <div style="flex:1;">
                            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:8px;">
                                <div style="font-size:0.8rem; color:#777;">
                                    🕒 ${timeStr} • 📦 ${totalCount} ชิ้น • 📂 ${firstEntry.category || 'เงินสด'}
                                    ${isUnknownComp ? `<button class="btn btn-neutral" style="padding:0 6px; font-size:0.65rem; border:1px solid #0288d1; color:#0288d1; background:#e1f5fe; border-radius:12px; margin-left:5px;" onclick="saveSessionAsCompany('${session.sessionId}')">💾 บันทึก บ.</button>` : ''}
                                </div>
                                <div style="display:flex; gap:4px;">
                                    <button class="btn btn-neutral" style="padding:3px 6px; font-size:0.7rem; border:1px solid #ddd; background:#fff;" title="รายงานเฉพาะชุดนี้" onclick="exportExceptionImage(['${session.sessionId}'])">📸</button>
                                    <button class="btn btn-neutral" style="padding:3px 6px; font-size:0.7rem; border:1px solid #ddd; background:#fff;" title="Edit" onclick="editExceptionSession('${session.sessionId}')">✏️</button>
                                    <button class="btn btn-neutral" style="padding:3px 6px; font-size:0.7rem; border:1px solid #ffcdd2; color:#d32f2f; background:#fff;" title="Delete" onclick="deleteExceptionSession('${session.sessionId}')">🗑️</button>
                                </div>
                            </div>

                            <div style="background:#fff; border-radius:6px; padding:8px 60px 8px 10px; margin-bottom:8px; border:1px solid #eee; position:relative; min-height:30px;">
                                <button style="position:absolute; top:8px; right:8px; background:none; border:none; color:#0288d1; font-size:0.7rem; cursor:pointer;" onclick="copySessionTracks('${session.sessionId}')">📋 คัดลอก</button>
                                <div style="display:flex; flex-wrap:wrap; gap:4px; font-family:monospace; font-size:0.9rem;">
                                    ${compressed.map(g => `<span style="background:#f5faff; border:1px solid #d1e3f8; padding:1px 6px; border-radius:4px; color:#0277bd;">${g.display}</span>`).join('')}
                                </div>
                                ${cdWarning}
                            </div>

                            <div style="display:flex; justify-content:space-between; align-items:center;">
                                <div style="font-size:0.9rem; color:#d32f2f; font-weight:500;">
                                    <span style="color:#aaa; font-size:0.75rem; font-weight:normal;">สาเหตุ:</span> ${session.reason || '-'}
                                </div>
                                ${hasImages ? `
                                    <div style="display:flex; gap:4px;" id="imgs-preview-${session.sessionId}">
                                        ${images.slice(0, 3).map(img => `<img src="${img.dataUrl}" style="width:28px; height:28px; object-fit:cover; border-radius:4px; border:1px solid #ddd; cursor:pointer;" onclick="toggleCardImages('${session.sessionId}')">`).join('')}
                                        ${images.length > 3 ? `<span style="font-size:0.65rem; color:#999;">+${images.length - 3}</span>` : ''}
                                    </div>
                                ` : ''}
                            </div>

                            <div id="imgs-full-${session.sessionId}" style="display:none; margin-top:10px; grid-template-columns:repeat(auto-fill, minmax(70px, 1fr)); gap:8px; border-top:1px dashed #eee; padding-top:10px;">
                                ${images.map(img => `<img src="${img.dataUrl}" style="width:100%; border-radius:6px; border:1px solid #eee;">`).join('')}
                            </div>
                        </div>
                    </div>
                </div>
            `;
        });

        html += `</div></div>`;
    });

    html += '</div>';
    container.innerHTML = html;
    updateSelectAllState();
}

/**
 * Updates the 'Select All' and Company-level checkbox states based on a strict hierarchy
 */
function updateSelectAllState() {
    // 1. Process each company group div for isolation
    const groupDivs = document.querySelectorAll('.company-report-group');
    let totalItems = 0;
    let totalChecked = 0;

    groupDivs.forEach(div => {
        const compCb = div.querySelector('.comp-select');
        const sessCbs = div.querySelectorAll('.sess-select');
        const checkedSess = div.querySelectorAll('.sess-select:checked');
        const compId = compCb ? compCb.getAttribute('data-comp') : null;
        
        const count = sessCbs.length;
        const checkedCount = checkedSess.length;
        totalItems += count;
        totalChecked += checkedCount;

        if (compCb) {
            compCb.checked = (count > 0 && count === checkedCount);
            compCb.indeterminate = (checkedCount > 0 && checkedCount < count);
        }

        // Update company-level label (e.g., "(2/3 Selected)")
        if (compId) {
            const labelEl = document.getElementById(`label-${compId}`);
            if (labelEl) {
                if (checkedCount === count) labelEl.textContent = `(${count})`;
                else labelEl.textContent = `(${checkedCount}/${count})`;
                labelEl.style.color = (checkedCount === count) ? '#0288d1' : (checkedCount > 0 ? '#ef6c00' : '#888');
            }
        }
    });

    // 2. Update Master Select All
    const selectAllCb = document.getElementById('sess-select-all');
    if (selectAllCb) {
        selectAllCb.checked = (totalItems > 0 && totalItems === totalChecked);
        selectAllCb.indeterminate = (totalChecked > 0 && totalChecked < totalItems);
        
        const masterLabel = document.getElementById('sess-master-label');
        if (masterLabel) {
            masterLabel.textContent = `รายการดราฟต์ (${totalChecked}/${totalItems})`;
            masterLabel.style.color = (totalChecked === totalItems) ? '#fff' : (totalChecked > 0 ? '#ffeb3b' : 'rgba(255,255,255,0.7)');
        }
    }
}

/**
 * Toggles all session checkboxes by strictly cascading settings
 */
function toggleAllSessions(checked) {
    // Update leaf nodes first
    document.querySelectorAll('.sess-select').forEach(cb => cb.checked = checked);
    // Update company nodes
    document.querySelectorAll('.comp-select').forEach(cb => cb.checked = checked);
    // Update master (if manually called from somewhere else)
    const masterCb = document.getElementById('sess-select-all');
    if (masterCb) masterCb.checked = checked;

    updateSelectAllState();
}

/**
 * Toggles all session checkboxes within a specific company group container
 */
function toggleSessionsInCompany(companyId, checked) {
    // Find sessions ONLY within the corresponding group container
    const cbs = document.querySelectorAll(`.sess-select[data-comp="${companyId}"]`);
    cbs.forEach(cb => cb.checked = checked);
    updateSelectAllState();
}

/**
 * Toggle between thumbnails and full image list in a card
 */
function toggleCardImages(sessionId) {
    const full = document.getElementById(`imgs-full-${sessionId}`);
    const preview = document.getElementById(`imgs-preview-${sessionId}`);
    if (!full) return;
    
    if (full.style.display === 'none') {
        full.style.display = 'grid';
        if (preview) preview.style.opacity = '0.3';
    } else {
        full.style.display = 'none';
        if (preview) preview.style.opacity = '1';
    }
}

/**
 * Copy all tracking numbers in a session to clipboard
 */
function copySessionTracks(sessionId) {
    const sessions = ExceptionManager.getAll();
    const tracks = sessions.filter(e => e.sessionId === sessionId).map(e => e.trackNum);
    if (tracks.length === 0) return;
    
    const text = tracks.join('\n');
    navigator.clipboard.writeText(text).then(() => {
        alert(`คัดลอก ${tracks.length} รายการแล้ว`);
    }).catch(err => {
        console.error('Copy failed:', err);
    });
}

/**
 * Save the company name of a session into the database permanently
 */
function saveSessionAsCompany(sessionId) {
    const sessionItems = ExceptionManager.getAll().filter(e => e.sessionId === sessionId);
    if (sessionItems.length === 0) return;
    
    const track = sessionItems[0].trackNum;
    const currentName = sessionItems[0].companyName || "-";
    
    const newName = prompt("ระบุชื่อบริษัท/ลูกค้า สำหรับเลขพัสดุนี้:", currentName === "-" ? "" : currentName);
    if (!newName || newName.trim() === "") return;
    
    // Save to CustomerDB
    if (typeof CustomerDB !== 'undefined') {
        CustomerDB.set(track, newName.trim());
        alert("บันทึกข้อมูลบริษัทสำเร็จ");
        
        // Update all items in this session to reflect the new name in UI
        const all = ExceptionManager.getAll();
        all.forEach(item => {
            if (item.sessionId === sessionId) {
                item.companyName = newName.trim();
                if (!item.metadata) item.metadata = {};
                item.metadata.companyName = newName.trim();
            }
        });
        localStorage.setItem('thp_exception_db_v1', JSON.stringify(all));
        
        renderExceptionTable();
    }
}

/**
 * Helper to format tracking numbers consistently for UI.
 */
function formatTrackingForUI(raw) {
    if (!raw) return "";
    raw = raw.replace(/\s+/g, '');
    if (raw.length === 13)
        return `${raw.slice(0,2)} ${raw.slice(2,6)} ${raw.slice(6,10)} ${raw.slice(10,11)} ${raw.slice(11,13)}`;
    return raw;
}

/**
 * Helper to compress continuous tracking ranges for compact display in Draft view.
 */
function compressEntriesForDisplay(entries) {
    const parsed = entries.map(e => {
        const m = (e.trackNum || "").replace(/\s+/g,'').match(/^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/);
        return m ? { full: (e.trackNum||"").replace(/\s+/g,''), prefix: m[1], body: parseInt(m[2]), cd: m[3], suffix: m[4] } : { full: (e.trackNum||""), prefix: null };
    });
    const sortable = parsed.filter(p => p.prefix !== null).sort((a,b) => {
        if (a.prefix !== b.prefix || a.suffix !== b.suffix) return (a.full||"").localeCompare(b.full||"");
        return a.body - b.body;
    });
    const unsortable = parsed.filter(p => p.prefix === null);
    const groups = [];
    let i = 0;
    while (i < sortable.length) {
        let j = i;
        while (j+1 < sortable.length &&
            sortable[j+1].prefix === sortable[i].prefix &&
            sortable[j+1].suffix === sortable[i].suffix &&
            sortable[j+1].body === sortable[j].body + 1) { j++; }
        if (j - i === 0) {
            groups.push({ display: formatTrackingForUI(sortable[i].full), count: 1 });
        } else {
            groups.push({
                display: `${formatTrackingForUI(sortable[i].full)} ~ ${formatTrackingForUI(sortable[j].full)}`,
                count: j - i + 1
            });
        }
        i = j + 1;
    }
    unsortable.forEach(p => groups.push({ display: formatTrackingForUI(p.full), count: 1 }));
    return groups;
}

async function deleteExceptionSession(sessionId) {
    if (confirm('ลบรายการในกลุ่มนี้ทั้งหมดใช่หรือไม่?')) {
        const exceptions = await ExceptionManager.getAll();
        const sessionItems = exceptions.filter(e => e.sessionId === sessionId).map(e => e.trackNum.replace(/\s/g,''));
        
        await ExceptionManager.removeSession(sessionId);
        
        // --- Bounce Back Logic (v1.79) ---
        if (currentUnifiedResults && currentUnifiedResults.length > 0) {
            currentUnifiedResults.forEach(item => {
                if (sessionItems.includes(item.number.replace(/\s/g,''))) {
                    item.moved = false;
                }
            });
            localStorage.setItem('thp_last_search_results', JSON.stringify(currentUnifiedResults));
            renderStoredUnifiedNumbers(currentUnifiedTitle, currentUnifiedResults);
        }

        await renderExceptionTable();
    }
}

async function deleteException(id) {
    if (confirm('ยืนยันลบรายการนี้?')) {
        await ExceptionManager.remove(id);
        await renderExceptionTable();
    }
}

async function editExceptionSession(sessionId) {
    const exceptions = await ExceptionManager.getAll();
    const sessionItems = exceptions.filter(e => (e.sessionId || e.id) === sessionId);
    if (sessionItems.length === 0) return;

    const first = sessionItems[0];
    currentEditingSessionId = sessionId;

    // 1. Populate Range/Single
    if (sessionItems.length > 1) {
        document.getElementById('exception-range-toggle').checked = true;
        document.getElementById('exception-start-input').value = sessionItems[0].trackNum;
        document.getElementById('exception-end-input').value = sessionItems[sessionItems.length-1].trackNum;
    } else {
        document.getElementById('exception-range-toggle').checked = false;
        document.getElementById('exception-track-input').value = first.trackNum;
    }
    toggleExceptionRangeMode();

    // 2. Populate Other Fields
    document.getElementById('exception-reason-input').value = first.reason;
    document.getElementById('exception-first-status').value = first.firstStatus || 'ใส่ของลงถุง';
    
    // Date/Time Pickers
    if (first.dateTime) {
        setExceptionDatePickerFromBE(first.dateTime);
    }

    // 3. Populate Images
    clearExceptionImages();
    if (first.images && Array.isArray(first.images)) {
        // We simulate handleExceptionImageUpload behavior but with dataUrls
        const preview = document.getElementById('exception-img-preview');
        first.images.forEach(imgData => {
            const idx = exceptionImages.length;
            exceptionImages.push({ dataUrl: imgData.dataUrl, originalDataUrl: imgData.dataUrl, name: imgData.name });

            const wrapper = document.createElement('div');
            wrapper.style.cssText = 'position:relative; display:inline-block;';
            wrapper.id = `exc-img-${idx}`;

            wrapper.innerHTML = `
                <img src="${imgData.dataUrl}" style="max-width:120px; max-height:90px; border-radius:4px; border:1px solid #ccc; display:block;">
                <button onclick="const currentId = parseInt(this.parentElement.id.replace('exc-img-', ''), 10); exceptionImages.splice(currentId, 1); this.parentElement.remove(); const allWrappers = document.getElementById('exception-img-preview').children; Array.from(allWrappers).forEach((w, i) => { w.id = 'exc-img-' + i; });" style="position:absolute; top:-5px; right:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#d32f2f; color:white; font-size:0.7rem; cursor:pointer; line-height:1; padding:0;">✕</button>
                <button onclick="const currentId = parseInt(this.parentElement.id.replace('exc-img-', ''), 10); openImageEditor(currentId);" style="position:absolute; top:-5px; left:-5px; border-radius:50%; width:18px; height:18px; border:none; background:#0288d1; color:white; font-size:0.6rem; cursor:pointer; line-height:1; padding:0;">✏️</button>
            `;
            preview.appendChild(wrapper);
        });
        
        // Open panel
        const panel = document.getElementById('exception-img-panel');
        if (panel) {
            panel.style.display = 'block';
            const chevron = document.getElementById('exception-img-chevron');
            if (chevron) chevron.textContent = '▲ ซ่อน';
        }
    }

    // Update Save Button Text for Edit Mode
    const saveBtn = document.getElementById('exception-save-btn');
    if (saveBtn) {
        saveBtn.innerHTML = '<i class="fas fa-save mr-2" style="margin-right:8px;"></i> บันทึกการแก้ไขนี้ (Save Changes)';
        saveBtn.style.background = "linear-gradient(135deg,#ef6c00,#e65100)"; // Orange for edit
    }

    // Scroll to the Form Header specifically
    const formHeader = document.getElementById('exception-form-title');
    if (formHeader) {
        formHeader.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } else {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
}

async function clearAllExceptions() {
    if (confirm('ยืนยันล้างประวัติรายงานทั้งหมด?')) {
        await ExceptionManager.clearAll();
        await renderExceptionTable();
    }
}

async function exportExceptionImage(specificSids = null) {

    const selectedSids = specificSids || Array.from(document.querySelectorAll('.sess-select:checked')).map(cb => cb.value);
    if (selectedSids.length === 0) {
        alert('กรุณาเลือกอย่างน้อย 1 รายการเพื่อออกรายงานครับ');
        return;
    }

    const exceptions = await ExceptionManager.getAll();
    const selectedItems = exceptions.filter(e => selectedSids.includes(e.sessionId || e.id));
    const meta = getExceptionMeta();

    // 1. Grouping Logic
    // Structure: Category -> Company -> Sessions
    const groups = {}; 

    selectedItems.forEach(item => {
        const cat = item.category || 'อื่นๆ';
        const com = item.companyName || '-';
        const sid = item.sessionId || item.id;

        if (!groups[cat]) groups[cat] = {};
        if (!groups[cat][com]) groups[cat][com] = {};
        if (!groups[cat][com][sid]) {
            groups[cat][com][sid] = { sessionId: sid, entries: [], first: item, timestamp: item.timestamp };
        }
        groups[cat][com][sid].entries.push(item);
    });

    // 2. Prepare Summary Data
    const summaryRows = [];
    Object.keys(groups).sort().forEach(cat => {
        Object.keys(groups[cat]).sort().forEach(com => {
            let total = 0;
            Object.values(groups[cat][com]).forEach(sess => total += sess.entries.length);
            summaryRows.push({ category: cat, company: com, count: total });
        });
    });

    // 3. Prepare Flat List for Pagination (Group Headers + Session Blocks)
    const exportBlocks = [];
    Object.keys(groups).sort().forEach(cat => {
        exportBlocks.push({ type: 'categoryHeader', title: cat });
        Object.keys(groups[cat]).sort().forEach(com => {
            exportBlocks.push({ type: 'companyHeader', title: com });
            const sids = Object.values(groups[cat][com]).sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
            sids.forEach(sess => {
                exportBlocks.push({ type: 'session', data: sess });
            });
        });
    });

    // Pagination: Approx 12 items/headers per page
    const ITEMS_PER_PAGE = 12;
    const totalPages = Math.ceil(exportBlocks.length / ITEMS_PER_PAGE);

    // Provide visual feedback
    let targetBtn = null;
    try {
        if (typeof event !== 'undefined' && event?.target && (event.target.tagName === 'BUTTON' || event.target.closest('button'))) {
            targetBtn = event.target.tagName === 'BUTTON' ? event.target : event.target.closest('button');
        }
    } catch(e) {}
    
    const originalBtnText = targetBtn ? targetBtn.innerText : "สร้างรูปภาพแจ้งหัวหน้า";
    if (targetBtn) {
        targetBtn.disabled = true;
        const oldHtml = targetBtn.innerHTML;
        targetBtn.innerText = "⏳ กำลังเตรียมไฟล์...";
    }

    const reportDateDisp = meta.date ? new Date(meta.date).toLocaleDateString('th-TH') : new Date().toLocaleDateString('th-TH');
    
    // Create temporary export container
    const exportDiv = document.createElement('div');
    exportDiv.id = 'temp-export-container';
    exportDiv.style.position = 'fixed';
    exportDiv.style.top = '0';
    exportDiv.style.left = '-9999px'; // Hide off-screen
    exportDiv.style.background = 'white';
    document.body.appendChild(exportDiv);

    try {
        for (let p = 0; p < totalPages; p++) {
            const pageBlocks = exportBlocks.slice(p * ITEMS_PER_PAGE, (p + 1) * ITEMS_PER_PAGE);
            // Label as (1/5) etc as requested
            const pageNumText = totalPages > 1 ? ` (ชุดที่ ${p + 1}/${totalPages})` : "";
            
            let pageSessions = [];
            pageBlocks.forEach(block => {
                if (block.type === 'session') {
                    pageSessions.push(block.data);
                }
            });

            let headerHtml = `
                <div style="margin-bottom:20px; border-bottom:4px solid #004ba0; padding-bottom:12px; display:flex; justify-content:space-between; align-items:flex-end;">
                    <div>
                        <div style="font-size:1.6rem; font-weight:bold; color:#01579b;">รายงานชิ้นงานที่ไม่มีสถานะรับฝาก</div>
                        <div style="font-size:1.1rem; color:#0288d1; margin-top:2px; font-weight:500;">(Exception & Missing Items Report)</div>
                    </div>
                    <div style="text-align:right;">
                        <div style="font-size:1.1rem; font-weight:bold; color:#d32f2f;">${reportDateDisp}</div>
                        ${totalPages > 1 ? `<div style="font-size:0.9rem; color:#d32f2f; font-weight:bold; margin-top:4px;">${pageNumText.trim()}</div>` : ''}
                    </div>
                </div>
                <div style="margin-bottom:15px; font-size:0.9rem; line-height:1.6; color:#444; display:grid; grid-template-columns: 1.5fr 1fr; gap:20px;">
                    <div>
                        ${meta.subject ? `<div><strong>เรื่อง:</strong> ${meta.subject}${pageNumText}</div>` : `<div><strong>เรื่อง:</strong> รายงานชิ้นงานค้าง ${pageNumText}</div>`}
                        ${meta.branch ? `<div><strong>ที่ทำการ:</strong> ${meta.branch}</div>` : ''}
                        ${meta.reporter ? `<div><strong>ผู้รายงาน:</strong> ${meta.reporter}</div>` : ''}
                    </div>
                    <div style="padding:8px; background:#f5faff; border:1px solid #d1e3f8; border-radius:6px;">
                        <strong>หมายเหตุ:</strong><br>
                        <span style="font-size:0.85rem; color:#666;">${meta.note || '-'}</span>
                    </div>
                </div>`;

            // Summary Table (ONLY Page 1)
            let summaryHtml = '';
            if (p === 0) {
                summaryHtml = `
                    <div style="margin-bottom:20px;">
                        <div style="font-size:1.1rem; font-weight:bold; color:#0d47a1; margin-bottom:8px;">📊 ตารางสรุปภาพรวม (Summary)</div>
                        <table style="width:100%; border-collapse:collapse; background:#fff; border:1px solid #eee; font-size:0.95rem;">
                            <thead>
                        <tr style="background:#003366; border-bottom:3px solid #00264d;">
                                    <th style="padding:12px 10px; text-align:left; border:1px solid #004ba0; background:#01579b !important; color:#ffffff !important; font-weight:bold; font-size:1.05rem;">หมวดงาน</th>
                                    <th style="padding:12px 10px; text-align:left; border:1px solid #004ba0; background:#01579b !important; color:#ffffff !important; font-weight:bold; font-size:1.05rem;">ชื่อบริษัท / ลูกค้า</th>
                                    <th style="padding:12px 10px; text-align:center; border:1px solid #004ba0; background:#01579b !important; color:#ffffff !important; font-weight:bold; width:15%; font-size:1.05rem;">รวม (ชิ้น)</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${summaryRows.map(row => `
                                    <tr>
                                        <td style="padding:8px; border:1px solid #eee; font-weight:bold; color:#01579b;">${row.category}</td>
                                        <td style="padding:8px; border:1px solid #eee;">${row.company}</td>
                                        <td style="padding:8px; border:1px solid #eee; text-align:center; font-weight:bold; font-size:1.1rem; color:#0288d1;">${row.count}</td>
                                    </tr>
                                `).join('')}
                                <tr style="background:#f1f8e9; font-weight:bold;">
                                    <td colspan="2" style="padding:10px; text-align:right; border:1px solid #eee;">รวมทั้งสิ้น (Total)</td>
                                    <td style="padding:10px; text-align:center; border:1px solid #eee; font-size:1.2rem; color:#2e7d32;">${summaryRows.reduce((a,b) => a + b.count, 0)}</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>`;
            }

            // Detail Table
            let detailRowsHtml = '';
            let sessionInReportIdx = 0; // Total index in the entire report

            exportBlocks.forEach((block, globalIdx) => {
                const isCurrentPage = (globalIdx >= p * ITEMS_PER_PAGE && globalIdx < (p + 1) * ITEMS_PER_PAGE);
                
                if (block.type === 'session') sessionInReportIdx++;

                if (!isCurrentPage) return;

                if (block.type === 'categoryHeader') {
                    detailRowsHtml += `
                        <tr style="background:#01579b; color:#ffffff !important; font-weight:bold; border-bottom:2px solid #014175;">
                            <td colspan="3" style="padding:12px 15px; font-size:1.2rem;">📁 หมวด: ${block.title}</td>
                        </tr>`;
                } else if (block.type === 'companyHeader') {
                    detailRowsHtml += `
                        <tr style="background:#e1f5fe; color:#01579b; font-weight:bold; border-bottom:1px solid #b3e5fc;">
                            <td colspan="3" style="padding:8px 15px;">🏢 บริษัท: ${block.title}</td>
                        </tr>`;
                } else if (block.type === 'session') {
                    const sess = block.data;
                    const compressed = compressSessEntries(sess.entries);
                    const totalCount = sess.entries.length;
                    const trackDisplay = compressed.map(g => `<div style="font-family:monospace; font-weight:bold; white-space:nowrap; margin-bottom:2px;">${g.display}</div>`).join('');
                    
                    const firstE = sess.entries[0];
                    let dispDT = new Date(sess.timestamp).toLocaleDateString('th-TH');
                    if (firstE.dateTime) {
                        const m = firstE.dateTime.match(/^(\d{4})-(\d{2})-(\d{2})/);
                        if (m) {
                            let y = parseInt(m[1]);
                            if (y < 2500) y += 543;
                            let timePart = "";
                            const tm = firstE.dateTime.match(/T(\d{2}:\d{2})/);
                            if (tm) timePart = " " + tm[1];
                            dispDT = `${m[3]}/${m[2]}/${y}${timePart}`;
                        }
                    }

                    const isLast = (globalIdx === exportBlocks.length - 1);
                    const dispIdx = isLast ? '🏁' : sessionInReportIdx;

                    // Per-session images
                    let sessImagesHtml = '';
                    const sessImgs = sess.entries[0].images || [];
                    if (sessImgs.length > 0) {
                        const scaleNode = document.getElementById('exception-img-scale');
                        const scaleVal = scaleNode ? scaleNode.value : "100";
                        const imgSize = (parseInt(scaleVal) / 100 * 180) + "px"; // 180px base for inline
                        
                        sessImagesHtml = `
                            <div style="display:flex; flex-wrap:wrap; gap:8px; margin-top:10px;">
                                ${sessImgs.map(img => `<img src="${img.dataUrl}" style="width:${imgSize}; border:1px solid #ddd; border-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,0.1);">`).join('')}
                            </div>`;
                    }

                    detailRowsHtml += `
                        <tr style="border-bottom: 2px solid #eee; vertical-align:top;">
                            <td style="padding:12px 5px; text-align:center; width:8%; font-weight:bold; color:#666; font-size:1.1rem; background:#f9f9f9; border-right:1px solid #eee;">
                                ${isLast ? `<div style="font-size:0.7rem; color:#d32f2f; margin-bottom:2px;">ท้ายสุด</div>` : ''}
                                ${dispIdx}
                            </td>
                            <td style="padding:12px 15px; width:40%;">${trackDisplay}</td>
                            <td style="padding:12px 15px; width:52%;">
                                <div style="color:#d32f2f; font-weight:bold; margin-bottom:6px; font-size:1.1rem;">${firstE.reason}</div>
                                <div style="font-size:0.85rem; color:#666; margin-bottom:4px;">
                                    ${firstE.firstStatus ? `<span>สถานะ: ${firstE.firstStatus}</span> | ` : ''}
                                    <span>${dispDT}</span>
                                </div>
                                <div style="background:#fff9c4; color:#5d4037; font-size:0.8rem; padding:2px 6px; border-radius:4px; display:inline-block; margin-bottom:5px;">
                                    👥 ${sess.companyName} | 📦 ${totalCount} ชิ้น | 📂 ${block.title || (exportBlocks.find(b => b.type === 'categoryHeader')?.title) || '-'}
                                </div>
                                ${sessImagesHtml}
                            </td>
                        </tr>`;
                }
            });

            let detailTableHtml = `
                <div style="font-size:1.1rem; font-weight:bold; color:#0d47a1; margin-bottom:8px;">🔍 รายละเอียดแยกตามรายการ (Itemized Details)</div>
                <table style="width:100%; border-collapse:collapse; border:1px solid #ccc; box-shadow:0 2px 10px rgba(0,0,0,0.05);">
                    <thead>
                        <tr style="background:#01579b; border-bottom:4px solid #014175;">
                            <th style="padding:14px 5px; text-align:center; border:1px solid #004ba0; background:#01579b !important; color:#ffffff !important; font-size:1.1rem; font-weight:bold;">ลำดับ</th>
                            <th style="padding:14px 10px; text-align:left; border:1px solid #004ba0; background:#01579b !important; color:#ffffff !important; font-size:1.1rem; font-weight:bold;">หมายเลขพัสดุ</th>
                            <th style="padding:14px 10px; text-align:left; border:1px solid #004ba0; background:#01579b !important; color:#ffffff !important; font-size:1.1rem; font-weight:bold;">สาเหตุ / ข้อมูลสแกน / รูปภาพประกอบ</th>
                        </tr>
                    </thead>
                    <tbody>${detailRowsHtml}</tbody>
                </table>`;

            exportDiv.innerHTML = `
                <div style="background:white; padding:40px; width:1000px; font-family:'Sarabun', sans-serif; box-sizing:border-box;">
                    ${headerHtml}
                    ${summaryHtml}
                    ${detailTableHtml}
                    <div style="margin-top:30px; text-align:center; font-size:0.85rem; color:#aaa; border-top:1px solid #f0f0f0; padding-top:10px;">
                        Tracking Analyst Helper | ชุดที่ ${p+1} จากทั้งหมด ${totalPages} ชุด
                    </div>
                </div>`;

            // Wait a moment for images and fonts to render fully
            await new Promise(r => setTimeout(r, 600));
            // Add a tick for layout stability
            await new Promise(r => window.requestAnimationFrame(r));

            const captureNode = exportDiv.querySelector('div');
            // Capture
            const canvas = await html2canvas(captureNode, {
                scale: 2,
                backgroundColor: '#ffffff',
                logging: true,
                useCORS: true,
                allowTaint: true,
                width: captureNode.offsetWidth,
                height: captureNode.offsetHeight,
                onclone: (clonedDoc) => {
                    const clonedNode = clonedDoc.getElementById(exportDiv.id);
                    if (clonedNode) {
                        clonedNode.style.opacity = "1";
                        clonedNode.style.visibility = "visible";
                        clonedNode.style.zIndex = "1";
                        clonedNode.style.left = "0";
                    }
                }
            });

            const imgData = canvas.toDataURL('image/jpeg', 0.85); // Use JPEG for better performance
            const link = document.createElement('a');
            const suffix = totalPages > 1 ? `_Part${p + 1}` : "";
            link.download = `Exception_Report_${new Date().toISOString().slice(0, 10)}${suffix}.jpg`;
            link.href = imgData;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            if (totalPages > 1) await new Promise(r => setTimeout(r, 600));
        }
    } catch (err) {
        console.error("Error generating image:", err);
        alert("ขออภัย! เกิดข้อผิดพลาดในการสร้างรูปภาพ: " + (err.message || err));
    } finally {
        if (exportDiv && exportDiv.parentNode) {
            document.body.removeChild(exportDiv);
        }
        if (event?.target) {
            event.target.disabled = false;
            event.target.innerText = originalBtnText;
        }
    }
}

function toggleHistoryImages(sid) {
    const div = document.getElementById(`sess-imgs-${sid}`);
    if (div) div.style.display = (div.style.display === 'none') ? 'flex' : 'none';
}

function toggleAllSessions(checked) {
    document.querySelectorAll('.sess-select').forEach(cb => cb.checked = checked);
}

function compressSessEntries(entries) {
    function fNum(r) {
        r = r.replace(/\s+/g, '');
        if (r.length === 13) return `${r.slice(0,2)} ${r.slice(2,6)} ${r.slice(6,10)} ${r.slice(10,11)} ${r.slice(11,13)}`;
        return r;
    }
    const parsed = entries.map(e => {
        const m = e.trackNum.replace(/\s+/g,'').match(/^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/);
        return m ? { full: e.trackNum.replace(/\s+/g,''), prefix: m[1], body: parseInt(m[2]), cd: m[3], suffix: m[4] } : { full: e.trackNum, prefix: null };
    });
    const sortable = parsed.filter(p => p.prefix !== null).sort((a,b) => {
        if (a.prefix !== b.prefix || a.suffix !== b.suffix) return a.full.localeCompare(b.full);
        return a.body - b.body;
    });
    const unsortable = parsed.filter(p => p.prefix === null);
    const groups = [];
    let i = 0;
    while (i < sortable.length) {
        let j = i;
        while (j+1 < sortable.length && sortable[j+1].prefix === sortable[i].prefix && sortable[j+1].suffix === sortable[i].suffix && sortable[j+1].body === sortable[j].body + 1) { j++; }
        if (j - i === 0) groups.push({ display: fNum(sortable[i].full), count: 1 });
        else groups.push({ display: `${fNum(sortable[i].full)} <span style="color:#555;">ถึง</span> ${fNum(sortable[j].full)}`, count: j - i + 1 });
        i = j + 1;
    }
    unsortable.forEach(p => groups.push({ display: fNum(p.full), count: 1 }));
    return groups;
}

// Auth listener removed (handled in initial load)

function updateExceptionImageScale() {
    const scaleVal = document.getElementById('exception-img-scale') ? document.getElementById('exception-img-scale').value : "100";
    const imgSize = (parseInt(scaleVal) / 100 * 200) + "px"; // Mapping 100% to 200px base
    
    // Update live preview in upload section
    const previewImgs = document.querySelectorAll('#exception-img-preview img');
    previewImgs.forEach(img => {
        img.style.maxWidth = imgSize;
        img.style.maxHeight = imgSize;
    });
    
    // Update export Target section images (if visible)
    const exportImgs = document.querySelectorAll('#exception-img-export-area img');
    exportImgs.forEach(img => {
        img.style.width = imgSize;
    });
}

function editExceptionCompany(sessionId, trackNum) {
    const newName = prompt('ระบุชื่อบริษัทใหม่ (ระบบจะเรียนรู้และบันทึกลงฐานข้อมูล 240 รายการอัตโนมัติ):');
    if (newName === null || newName.trim() === '') return;
    
    const exceptions = ExceptionManager.getAll();
    let updated = false;
    exceptions.forEach(e => {
        if (e.sessionId === sessionId) {
            e.companyName = newName.trim();
            updated = true;
        }
    });

    if (updated) {
        localStorage.setItem(EXCEPTION_KEY, JSON.stringify(exceptions));
        
        if (typeof CustomerDB !== 'undefined' && typeof TrackingUtils !== 'undefined') {
            const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
            const match = trackNum.match(regex);
            if (match) {
                const prefix = match[1];
                const bodyNum = parseInt(match[2], 10);
                const suffix = match[4];
                
                const bookStart = Math.floor((bodyNum - 1) / 240) * 240 + 1;
                const itemsToAdd = [];
                for (let i = 0; i < 240; i++) {
                    let currentNumStr = (bookStart + i).toString().padStart(8, '0');
                    if (currentNumStr.length > 8) break;
                    let checkDigit = TrackingUtils.calculateS10CheckDigit(currentNumStr);
                    if (checkDigit !== null) {
                        itemsToAdd.push(`${prefix}${currentNumStr}${checkDigit}${suffix}`);
                    }
                }
                
                if (itemsToAdd.length > 0) {
                    CustomerDB.addBatch({
                        name: newName.trim(),
                        type: "General",
                        contract: "",
                        requestDate: "",
                        timestamp: new Date().getTime()
                    }, itemsToAdd);
                    
                    alert(`อัปเดตชื่อบริษัทเป็น "${newName.trim()}" สำเร็จ!\n(สอนบอทจดจำเลขชุดนี้ 240 รายการเรียบร้อย)`);
                }
            }
        }
        renderExceptionTable();
    }
}

function editExceptionReason(sessionId) {
    const newReason = prompt('ระบุสาเหตุ / รายละเอียดการตกหล่นใหม่:');
    if (newReason === null || newReason.trim() === '') return;
    
    const exceptions = ExceptionManager.getAll();
    exceptions.forEach(e => {
        if (e.sessionId === sessionId) {
            e.reason = newReason.trim();
        }
    });
    localStorage.setItem(EXCEPTION_KEY, JSON.stringify(exceptions));
    renderExceptionTable();
}

// ==========================================
// SECTION: INDIVIDUAL IMAGE EDITOR
// ==========================================
let editIdx = -1;
let tImg = new Image();
let tScale = 1;
let baseFitScale = 1;
let tPanX = 0;
let tPanY = 0;
let isDragging = false;
let startX = 0, startY = 0;

function openImageEditor(index) {
    const item = exceptionImages[index];
    if (!item) return;
    editIdx = index;
    
    tImg.onload = function() {
        const cvs = document.getElementById('img-editor-canvas');
        baseFitScale = Math.max(cvs.width / tImg.width, cvs.height / tImg.height);
        tScale = baseFitScale;
        tPanX = (cvs.width - (tImg.width * tScale)) / 2;
        tPanY = (cvs.height - (tImg.height * tScale)) / 2;
        
        const zoomSlider = document.getElementById('img-editor-zoom');
        if(zoomSlider) zoomSlider.value = "100";
        
        drawEditorCanvas();
        document.getElementById('img-editor-modal').style.display = 'flex';
    };
    tImg.src = item.originalDataUrl || item.dataUrl; 
}

function closeImageEditor() {
    document.getElementById('img-editor-modal').style.display = 'none';
}

function drawEditorCanvas() {
    const cvs = document.getElementById('img-editor-canvas');
    if (!cvs) return;
    const ctx = cvs.getContext('2d');
    
    // Smooth rendering
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    
    ctx.clearRect(0, 0, cvs.width, cvs.height);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, cvs.width, cvs.height);
    
    if (tImg.src) {
        ctx.drawImage(tImg, tPanX, tPanY, tImg.width * tScale, tImg.height * tScale);
    }
}

function saveImageEditor() {
    if (editIdx > -1 && exceptionImages[editIdx]) {
        const cvs = document.getElementById('img-editor-canvas');
        const croppedDataUrl = cvs.toDataURL('image/jpeg', 0.92);
        exceptionImages[editIdx].dataUrl = croppedDataUrl;
        
        const thumb = document.querySelector(`#exc-img-${editIdx} img`);
        if (thumb) thumb.src = croppedDataUrl;
    }
    closeImageEditor();
}

document.addEventListener('DOMContentLoaded', () => {
    const cvs = document.getElementById('img-editor-canvas');
    const zoomSlider = document.getElementById('img-editor-zoom');
    
    if (cvs) {
        cvs.addEventListener('mousedown', e => {
            isDragging = true;
            const rect = cvs.getBoundingClientRect();
            const ratioX = cvs.width / rect.width;
            const ratioY = cvs.height / rect.height;
            startX = (e.clientX * ratioX) - tPanX;
            startY = (e.clientY * ratioY) - tPanY;
        });
        window.addEventListener('mouseup', () => { isDragging = false; });
        window.addEventListener('mousemove', e => {
            if (isDragging) {
                const rect = cvs.getBoundingClientRect();
                const ratioX = cvs.width / rect.width;
                const ratioY = cvs.height / rect.height;
                tPanX = (e.clientX * ratioX) - startX;
                tPanY = (e.clientY * ratioY) - startY;
                drawEditorCanvas();
            }
        });
        
        cvs.addEventListener('touchstart', e => {
            if (e.touches.length === 1) {
                isDragging = true;
                const rect = cvs.getBoundingClientRect();
                const ratioX = cvs.width / rect.width;
                const ratioY = cvs.height / rect.height;
                startX = (e.touches[0].clientX * ratioX) - tPanX;
                startY = (e.touches[0].clientY * ratioY) - tPanY;
            }
        }, {passive: false});
        window.addEventListener('touchend', () => { isDragging = false; });
        window.addEventListener('touchmove', e => {
            if (isDragging && e.touches.length === 1) {
                e.preventDefault();
                const rect = cvs.getBoundingClientRect();
                const ratioX = cvs.width / rect.width;
                const ratioY = cvs.height / rect.height;
                tPanX = (e.touches[0].clientX * ratioX) - startX;
                tPanY = (e.touches[0].clientY * ratioY) - startY;
                drawEditorCanvas();
            }
        }, {passive: false});
    }
    
    if (zoomSlider) {
        zoomSlider.addEventListener('input', e => {
            const pct = parseInt(e.target.value, 10) / 100;
            const newScale = baseFitScale * pct;
            const cvsW = cvs.width, cvsH = cvs.height;
            const centerX = cvsW / 2;
            const centerY = cvsH / 2;
            
            tPanX = centerX - (centerX - tPanX) * (newScale / tScale);
            tPanY = centerY - (centerY - tPanY) * (newScale / tScale);
            
            tScale = newScale;
            drawEditorCanvas();
        });
    }
});

/**
 * v1.69: Toggle Satellite visibility with robust DOM traversal
 */
function toggleSatelliteGroup(triggerRow) {
    console.log('[DEBUG] Toggling satellite group for:', triggerRow.innerText.trim());
    
    // First try: directly next sibling
    let wrapper = triggerRow.nextElementSibling;
    
    // Check if not the wrapper, try finding in parent (more robust)
    if (!wrapper || !wrapper.classList.contains('satellite-wrapper')) {
        const parent = triggerRow.parentElement;
        if (parent) {
            wrapper = parent.querySelector('.satellite-wrapper');
        }
    }

    if (wrapper && wrapper.classList.contains('satellite-wrapper')) {
        const isExpanding = !wrapper.classList.contains('expanded');
        wrapper.classList.toggle('expanded');
        triggerRow.classList.toggle('expanded-trigger');
        console.log('[DEBUG] Success:', isExpanding ? 'Expanded' : 'Collapsed');
    } else {
        console.error('[DEBUG] Failed: Satellite wrapper not found for this row.', triggerRow);
    }
}

/**
 * DATABASE RENDERING FUNCTIONS (v1.71 RESTORED)
 */

window.renderRecentBatches = function(batches) {
    const table = document.getElementById('db-table');
    if (!table) return;
    
    // Update Count
    updateDbCount(batches.length);
    
    if (batches.length === 0) {
        table.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:40px; color:#999;">📭 ยังไม่มีข้อมูลในระบบ</td></tr>';
        return;
    }

    table.innerHTML = batches.map(b => `
        <tr onclick="loadBatchToView('${b.id}')" style="cursor:pointer;">
            <td style="padding:15px 10px;">
                <div style="font-weight:bold; color:#0277bd;">${b.name}</div>
                <div style="font-size:0.75rem; color:#888; font-family:monospace; margin-top:2px;">${b.rangeDesc}</div>
            </td>
            <td style="text-align:center;">
                <span class="badge" style="background:#e1f5fe; color:#0277bd; border:1px solid #b3e5fc;">${b.count} ชิ้น</span>
            </td>
            <td style="text-align:center; font-size:0.85rem; color:#666;">${b.type || 'EMS'}</td>
            <td style="text-align:center; font-size:0.8rem; color:#999;">
                ${new Date(b.timestamp).toLocaleDateString('th-TH')}<br>
                ${new Date(b.timestamp).toLocaleTimeString('th-TH', {hour:'2-digit', minute:'2-digit'})}
            </td>
            <td style="text-align:center;" onclick="event.stopPropagation()">
                <button class="btn btn-neutral" style="padding:4px 8px; font-size:0.7rem; color:#d32f2f; border:1px solid #ffcdd2;" onclick="deleteBatchConfirmed('${b.id}', '${b.name}')">🗑️ ลบ</button>
            </td>
        </tr>
    `).join('');
};

window.renderCompanySummaries = function(sums) {
    const table = document.getElementById('db-table');
    if (!table) return;
    
    updateDbCount(sums.length);

    if (sums.length === 0) {
        table.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:40px; color:#999;">📭 ไม่พบข้อมูลบริษัท</td></tr>';
        return;
    }

    table.innerHTML = sums.map(s => `
        <tr style="background:#f9f9f9;">
            <td colspan="2" style="padding:15px 10px; font-weight:bold; color:#333;">🏢 ${s.name}</td>
            <td style="text-align:center;"><span class="badge badge-primary">${s.totalCount} ชิ้น</span></td>
            <td style="text-align:center; font-size:0.8rem; color:#888;">${s.batches.length} ชุดข้อมููล</td>
            <td></td>
        </tr>
        ${s.batches.sort((a,b)=>b.timestamp-a.timestamp).map(b => `
            <tr onclick="loadBatchToView('${b.id}')" style="cursor:pointer; font-size:0.85rem; border-left:4px solid #eee;">
                <td style="padding:10px 20px; color:#666;">
                    • ${b.rangeDesc}
                </td>
                <td style="text-align:center; color:#888;">${b.count} ชิ้น</td>
                <td style="text-align:center; color:#888;">${b.type}</td>
                <td style="text-align:center; color:#999; font-size:0.75rem;">${new Date(b.timestamp).toLocaleDateString('th-TH')}</td>
                <td style="text-align:center;" onclick="event.stopPropagation()">
                    <button class="btn btn-neutral" style="padding:2px 6px; font-size:0.65rem; color:#d32f2f;" onclick="deleteBatchConfirmed('${b.id}', '${b.name}')">🗑️</button>
                </td>
            </tr>
        `).join('')}
    `).join('');
};

window.renderTrashBatches = function(trash) {
    const table = document.getElementById('db-table');
    if (!table) return;

    updateDbCount(trash.length);

    if (trash.length === 0) {
        table.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:40px; color:#999;">🗑️ ถังขยะว่างเปล่า</td></tr>';
        return;
    }

    table.innerHTML = trash.map(b => `
        <tr style="opacity:0.6;">
            <td style="padding:15px 10px;">
                <div style="font-weight:bold; color:#666;">${b.name}</div>
                <div style="font-size:0.72rem; color:#999;">${b.rangeDesc}</div>
            </td>
            <td style="text-align:center;">${b.count} ชิ้น</td>
            <td style="text-align:center;">${b.type}</td>
            <td style="text-align:center; font-size:0.75rem; color:#999;">
                ลบเมื่อ: ${new Date(b.deletedAt).toLocaleDateString('th-TH')}<br>
                ${new Date(b.deletedAt).toLocaleTimeString('th-TH', {hour:'2-digit', minute:'2-digit'})}
            </td>
            <td style="text-align:center;" onclick="event.stopPropagation()">
                <div style="display:flex; flex-direction:column; gap:4px;">
                    <button class="btn btn-neutral" style="padding:2px 8px; font-size:0.65rem; color:#2e7d32; border:1px solid #c8e6c9;" onclick="CustomerDB.restoreTrash('${b.id}').then(()=>updateDbViews())">🔄 กู้คืน</button>
                    <button class="btn btn-neutral" style="padding:2px 8px; font-size:0.65rem; color:#d32f2f; border:1px solid #ffcdd2;" onclick="if(confirm('ต้องการลบข้อมูลนี้ทิ้งถาวรหรือไม่?')) CustomerDB.permanentDeleteTrash('${b.id}').then(()=>updateDbViews())">❌ ลบทิ้ง</button>
                </div>
            </td>
        </tr>
    `).join('');
};

function updateDbCount(count) {
    const countEl = document.getElementById('db-count');
    if (countEl) countEl.innerText = count;
}

/**
 * UI Wrapper for confirmed batch deletion
 */
function deleteBatchConfirmed(batchId, batchName) {
    if (confirm(`คุณต้องการย้ายชุดข้อมูล "${batchName}" ไปยังถังขยะหรือไม่?`)) {
        CustomerDB.deleteBatch(batchId).then(() => {
            refreshUI();
            if (typeof window.showToast === 'function') window.showToast('ย้ายไปยังถังขยะเรียบร้อย');
        });
    }
}

/**
 * v1.96: Comprehensive UI Refresh (Fixed Scope & Robust Sync)
 * Syncs Database views, Search Results ownership, and Draft Reports.
 */
async function refreshUI() {
    console.info('[v1.96] Global UI Refresh started...');
    
    // 1. Update Customer DB views
    if (typeof updateDbViews === 'function') await updateDbViews();

    // 2. Sync Search Results Sidebar (Remove deleted owners)
    // Access directly (let variables at top level of app.js)
    if (typeof currentUnifiedResults !== 'undefined' && currentUnifiedResults && currentUnifiedResults.length > 0) {
        console.info(`[v1.96] Syncing ${currentUnifiedResults.length} search results...`);
        const updatePromises = currentUnifiedResults.map(async (item) => {
            // Re-fetch current owner state from DB
            const cleanNum = item.number.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            const newOwner = await CustomerDB.get(cleanNum);
            
            // Critical Update: 
            item.owner = newOwner || null;
            return item;
        });
        await Promise.all(updatePromises);
        
        // Persist and re-render (v1.97: Migration to StorageV2)
        await saveLastSearchResults(currentUnifiedTitle, currentUnifiedResults, false);
        
        if (typeof renderStoredUnifiedNumbers === 'function') {
            renderStoredUnifiedNumbers(currentUnifiedTitle || "", currentUnifiedResults);
        }
        console.info('[v1.96] Search Results synced successfully.');
    }

    // 3. Sync Draft Report (Exception Table)
    if (typeof ExceptionManager !== 'undefined') {
        const draftItems = await ExceptionManager.getAll();
        let draftChanged = false;
        for (const item of draftItems) {
            const clean = item.trackNum.replace(/[\s\u200B-\u200D\uFEFF\u202F]/g, '');
            const newOwner = await CustomerDB.get(clean);
            const newName = newOwner ? newOwner.name : '-';
            if (item.companyName !== newName) {
                item.companyName = newName;
                draftChanged = true;
            }
        }
        if (draftChanged) {
            await StorageV2.set(EXCEPTION_KEY, draftItems);
            if (typeof renderExceptionTable === 'function') await renderExceptionTable();
            console.info('[v1.96] Draft Report names synced.');
        }
    }
}

/**
 * Jump to a specific report session in the drafts/history list.
 * v1.91: Handles tab switching, filtering, and animation.
 */
function viewReportSession(sessionId) {
    // 1. Switch to 'smart' tab if not there
    switchTab('smart');
    
    // 2. Set filter to 'history' or 'today'? 
    // Usually, historical reports are in 'history'.
    setExceptionFilter('history');

    // 3. Wait for render and scroll
    setTimeout(() => {
        const card = document.getElementById(`sess-card-${sessionId}`);
        if (card) {
            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
            card.style.transition = 'all 0.5s ease';
            card.style.boxShadow = '0 0 20px rgba(211, 47, 47, 0.5)';
            card.style.borderColor = '#d32f2f';
            
            // Flash effect
            let flashes = 0;
            const interval = setInterval(() => {
                card.style.backgroundColor = flashes % 2 === 0 ? '#fff5f5' : '#fff';
                flashes++;
                if (flashes > 5) {
                    clearInterval(interval);
                    card.style.backgroundColor = '#fff';
                }
            }, 300);
        } else {
            // Try 'today' if not in history
            setExceptionFilter('today');
            setTimeout(() => {
                const cardToday = document.getElementById(`sess-card-${sessionId}`);
                if (cardToday) {
                    cardToday.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    cardToday.style.boxShadow = '0 0 20px rgba(2, 136, 209, 0.5)';
                } else {
                    alert('ขออภัย! ไม่พบข้อมูลรายงานชุดนี้ในเครื่องครับ (อาจถูกลบหรือจัดเก็บลงฐานข้อมูลอื่น)');
                }
            }, 100);
        }
    }, 100);
}

// v4.3.0: Premium Multi-Company Analytics Dashboard
async function showPremiumDashboard(forceRefresh = false) {
    const batches = await CustomerDB.getBatches();
    const batchList = Object.values(batches);
    
    if (batchList.length === 0) {
        window.showToast('ยังไม่มีข้อมูลประวัติในระบบ', 'info');
        return;
    }

    // --- State & Filters ---
    const now = new Date();
    let selectedYear = now.getFullYear();
    let selectedQuarter = Math.floor(now.getMonth() / 3) + 1; // 1-4
    let selectedCompany = 'all';

    // Unique Companies for Filter
    const companies = [...new Set(batchList.map(b => b.name))].sort();

    const renderDashboard = () => {
        const data = aggregateDashboardData(batchList, selectedYear, selectedQuarter, selectedCompany);
        const prevData = aggregateDashboardData(batchList, selectedYear - 1, selectedQuarter, selectedCompany);
        
        const main = document.getElementById('db-body');
        if (!main) return;

        // KPI Calculations
        const volChange = prevData.totalVolume > 0 ? ((data.totalVolume - prevData.totalVolume) / prevData.totalVolume * 100) : null;
        const valChange = prevData.totalValue > 0 ? ((data.totalValue - prevData.totalValue) / prevData.totalValue * 100) : null;

        main.innerHTML = `
            <div class="db-grid">
                <div class="kpi-card">
                    <div class="kpi-label">ปริมาณงานรวม (Packages)</div>
                    <div class="kpi-value">${data.totalVolume.toLocaleString()}</div>
                    ${volChange !== null ? `
                        <div class="kpi-delta ${volChange >= 0 ? 'delta-up' : 'delta-down'}">
                            ${volChange >= 0 ? '↑' : '↓'} ${Math.abs(volChange).toFixed(1)}% vs ปีที่แล้ว
                        </div>
                    ` : '<div class="kpi-delta" style="background:#f1f5f9; color:#64748b;">ไม่มีข้อมูลปีที่แล้ว</div>'}
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">ค่าบริการสุทธิ (Net Revenue)</div>
                    <div class="kpi-value">${data.totalValue.toLocaleString()}</div>
                    ${valChange !== null ? `
                        <div class="kpi-delta ${valChange >= 0 ? 'delta-up' : 'delta-down'}">
                            ${valChange >= 0 ? '↑' : '↓'} ${Math.abs(valChange).toFixed(1)}% vs ปีที่แล้ว
                        </div>
                    ` : '<div class="kpi-delta" style="background:#f1f5f9; color:#64748b;">ไม่มีข้อมูลปีที่แล้ว</div>'}
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">ราคาเฉลี่ยต่อชิ้น (Avg.)</div>
                    <div class="kpi-value">${data.totalVolume > 0 ? (data.totalValue / data.totalVolume).toFixed(2) : '0.00'}</div>
                    <div style="font-size:0.8rem; color:#94a3b8;">คำนวณจากราคาฝากส่ง</div>
                </div>
            </div>

            <div class="db-section">
                <h4 class="db-section-title">การกระจายตัวตามช่วงราคา (Price Distribution)</h4>
                <div class="dist-list">
                    ${Object.entries(data.priceGroups).sort((a,b) => parseFloat(b[0]) - parseFloat(a[0])).map(([price, count]) => {
                        const percent = data.totalVolume > 0 ? (count / data.totalVolume * 100) : 0;
                        return `
                            <div class="dist-item">
                                <div class="dist-label">${parseFloat(price).toFixed(2)} บาท</div>
                                <div class="dist-bar-bg">
                                    <div class="dist-bar-fill" style="width: ${percent}%"></div>
                                </div>
                                <div class="dist-value">${count.toLocaleString()}</div>
                            </div>
                        `;
                    }).join('') || '<p style="text-align:center; color:#94a3b8; padding:20px;">ไม่มีข้อมูลในช่วงเวลานี้</p>'}
                </div>
            </div>

            <div class="db-section">
                <h4 class="db-section-title">ประวัติการนำเข้ารายละเอียด (Transaction Log)</h4>
                <div style="overflow-x:auto;">
                    <table style="width:100%; border-collapse:collapse; font-size:0.85rem;">
                        <thead style="background:#f8fafc; border-bottom:1px solid #e2e8f0;">
                            <tr>
                                <th style="padding:12px; text-align:left;">วันที่</th>
                                <th style="padding:12px; text-align:left;">ชื่อชุดข้อมูล/บริษัท</th>
                                <th style="padding:12px; text-align:right;">จำนวน (ชิ้น)</th>
                                <th style="padding:12px; text-align:right;">ยอดเงิน</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${data.transactions.map(t => `
                                <tr style="border-bottom:1px solid #f1f5f9;">
                                    <td style="padding:12px; color:#64748b;">${new Date(t.timestamp).toLocaleDateString('th-TH')}</td>
                                    <td style="padding:12px; font-weight:600; color:#1e293b;">${t.name}</td>
                                    <td style="padding:12px; text-align:right; font-weight:bold;">${t.count.toLocaleString()}</td>
                                    <td style="padding:12px; text-align:right; color:#d63384; font-weight:bold;">${t.totalValue.toLocaleString()}</td>
                                </tr>
                            `).join('') || '<tr><td colspan="4" style="padding:20px; text-align:center; color:#94a3b8;">ไม่พบข้อมูล</td></tr>'}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    };

    // --- Helper for Aggregation ---
    const aggregateDashboardData = (list, year, quarter, company) => {
        const startMonth = (quarter - 1) * 3;
        const endMonth = startMonth + 2;
        
        const filtered = list.filter(b => {
            const d = new Date(b.timestamp);
            const m = d.getMonth();
            const y = d.getFullYear();
            const matchTime = (y === year && m >= startMonth && m <= endMonth);
            const matchCompany = (company === 'all' || b.name === company);
            return matchTime && matchCompany;
        });

        const totals = { totalVolume: 0, totalValue: 0, priceGroups: {}, transactions: [] };
        
        filtered.forEach(b => {
            let bTotal = 0;
            if (b.ranges) {
                b.ranges.forEach(r => {
                    totals.totalVolume += r.count;
                    const val = r.total || (r.count * r.price);
                    totals.totalValue += val;
                    bTotal += val;
                    
                    const pKey = r.price.toString();
                    totals.priceGroups[pKey] = (totals.priceGroups[pKey] || 0) + r.count;
                });
            }
            totals.transactions.push({ ...b, totalValue: bTotal });
        });

        totals.transactions.sort((a,b) => b.timestamp - a.timestamp);
        return totals;
    };

    // --- Create Backdrop & Modal ---
    const overlay = document.createElement('div');
    overlay.className = 'dashboard-overlay';
    overlay.id = 'db-overlay';
    
    overlay.innerHTML = `
        <div class="dashboard-modal">
            <div class="dashboard-header">
                <h2 class="dashboard-title">ศูนย์วิเคราะห์ข้อมูล & แดชบอร์ด</h2>
                <button class="db-close-btn" onclick="document.getElementById('db-overlay').remove()">✕</button>
            </div>
            <div class="dashboard-body">
                <div class="dashboard-controls">
                    <select id="db-company" class="db-select">
                        <option value="all">🏢 ทุกบริษัท (All Companies)</option>
                        ${companies.map(c => `<option value="${c}">${c}</option>`).join('')}
                    </select>
                    <select id="db-year" class="db-select">
                        <option value="${selectedYear}">${selectedYear}</option>
                        <option value="${selectedYear - 1}">${selectedYear - 1}</option>
                    </select>
                    <select id="db-quarter" class="db-select">
                        <option value="1" ${selectedQuarter === 1 ? 'selected' : ''}>ไตรมาส 1 (ม.ค. - มี.ค.)</option>
                        <option value="2" ${selectedQuarter === 2 ? 'selected' : ''}>ไตรมาส 2 (เม.ย. - มิ.ย.)</option>
                        <option value="3" ${selectedQuarter === 3 ? 'selected' : ''}>ไตรมาส 3 (ก.ค. - ก.ย.)</option>
                        <option value="4" ${selectedQuarter === 4 ? 'selected' : ''}>ไตรมาส 4 (ต.ค. - ธ.ค.)</option>
                    </select>
                    <div style="flex:1"></div>
                    <span style="font-size:0.75rem; color:#94a3b8;">อัปเดตเรียลไทม์จากฐานข้อมูล</span>
                </div>
                <div id="db-body"></div>
            </div>
        </div>
    `;

    document.body.appendChild(overlay);

    // Filter Listeners
    document.getElementById('db-company').addEventListener('change', (e) => { selectedCompany = e.target.value; renderDashboard(); });
    document.getElementById('db-year').addEventListener('change', (e) => { selectedYear = parseInt(e.target.value); renderDashboard(); });
    document.getElementById('db-quarter').addEventListener('change', (e) => { selectedQuarter = parseInt(e.target.value); renderDashboard(); });

    // Initial Render
    renderDashboard();
}
