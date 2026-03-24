/**
 * Smart Input Manager
 */

document.addEventListener('DOMContentLoaded', () => {
    initSmartInput();
});

function initSmartInput() {
    renderPrefixOptions();
    renderBlock1Options();
    setupSmartPaste();
    setupSmartInputEvents();
}

function setupSmartInputEvents() {
    const prefixInput = document.getElementById('smart-prefix');
    const block1Input = document.getElementById('smart-block1');
    const block2Input = document.getElementById('smart-block2');

    if (!prefixInput || !block1Input || !block2Input) return;

    // Map Thai Keyboard to English layout (for standard tracking ranges starting with EA, EQ, etc.)
    const thENMap = {
        'ฉ': 'E', 'ฏ': 'E',
        'ๆ': 'Q',
        'ไ': 'W', 'ร': 'I', 'น': 'O', 'ย': 'P',
        'ฟ': 'A', 'ห': 'S', 'ก': 'D', 'ด': 'F',
        'เ': 'G', '้': 'H', 'า': 'J',
        'ส': 'L', 'ผ': 'Z', 'ป': 'X', 'แ': 'C',
        'อ': 'V', 'ิ': 'B', 'ท': 'M', 'ม': '?',
        'พ': 'R', 'ะ': 'T', 'ั': 'Y', 'ร': 'U',
    };

    prefixInput.addEventListener('input', (e) => {
        let val = e.target.value;
        let converted = '';
        
        for(let i = 0; i < val.length; i++) {
            let ch = val[i];
            // If it's a Thai character in our map, convert it
            if (thENMap[ch]) {
                converted += thENMap[ch];
            } else {
                converted += ch;
            }
        }
        
        // Remove non-alphanumeric just to be safe, force uppercase
        converted = converted.replace(/[^A-Za-z]/g, '').toUpperCase();
        
        // Show warning if length changed due to removing non-letters
        if (val.length > 0 && converted.length === 0) {
            // Optional: small UI tick/warning here
        }
        
        e.target.value = converted;
    }); // <--- FIX: Missing closing brace for the 'input' event listener

    // Only evaluate navigation on 'keyup' to avoid interrupting IME/Predictive Text 
    // and correctly capture the length after 'input' has processed it.
    prefixInput.addEventListener('keyup', (e) => {
        // Don't auto-advance if they are backspacing or deleting
        if (e.key === 'Backspace' || e.key === 'Delete' || e.key === 'Tab') return;
        
        let val = e.target.value;
        if (val.length >= 2) {
            block1Input.focus();
            // Optional: select all text in next input for easy overwriting
            block1Input.select();
        }
    });

    block1Input.addEventListener('input', (e) => {
        let val = e.target.value.replace(/\D/g, ''); // Numbers only
        e.target.value = val;
        calculateSmartCheckDigit();
    });

    block1Input.addEventListener('keyup', (e) => {
        if (e.key === 'Backspace' || e.key === 'Delete' || e.key === 'Tab') return;
        
        if (e.target.value.length >= 4) {
             block2Input.focus();
             block2Input.select();
        }
    });

    block2Input.addEventListener('input', (e) => {
        let val = e.target.value.replace(/\D/g, ''); // Numbers only
        e.target.value = val;
        calculateSmartCheckDigit();
    });

    block2Input.addEventListener('keyup', (e) => {
        if (e.key === 'Backspace' || e.key === 'Delete' || e.key === 'Tab') return;

        // Fast Entry: Press Enter to Save and Clear Block 2
        if (e.key === 'Enter') {
             addSmartEntryAndSave();
             return;
        }

        const isRange = document.getElementById('smart-range-enable').checked;
        if (e.target.value.length >= 4 && isRange) {
             document.getElementById('smart-qty').focus();
             document.getElementById('smart-qty').select();
        } else if (e.target.value.length === 4 && !isRange) {
            // Optional: Auto-save on 4th digit (disabled by default to prevent accidental saves)
            // addSmartEntryAndSave();
        }
    });
}

function setupSmartPaste() {
    const textArea = document.getElementById('db-tracking-list');
    if (!textArea) return;

    textArea.addEventListener('paste', (e) => {
        // Prevent default paste first to analyze content
        const pasteData = (e.clipboardData || window.clipboardData).getData('text');

        // Try to extract tracking numbers
        const extracted = TrackingUtils.extractTrackingNumbers(pasteData);

        if (extracted && extracted.length > 0) {
            e.preventDefault();

            // Insert extracted numbers
            const newContent = extracted.join('\n');

            // Insert at cursor position or append?
            // Standard paste inserts at cursor.
            const startPos = textArea.selectionStart;
            const endPos = textArea.selectionEnd;
            const textBefore = textArea.value.substring(0, startPos);
            const textAfter = textArea.value.substring(endPos, textArea.value.length);

            textArea.value = textBefore + newContent + textAfter;

            // Move cursor to end of inserted text
            textArea.selectionStart = textArea.selectionEnd = startPos + newContent.length;

            // Optional: Notify user
            console.log(`Smart Paste: Extracted ${extracted.length} valid tracking numbers.`);
        }
        // If no tracking numbers found, let default paste happen (maybe they are pasting notes?)
    });
}

function renderPrefixOptions() {
    const list = document.getElementById('prefix-list');
    const prefixes = typeof PrefixManager !== 'undefined' ? PrefixManager.getAll() : [];
    list.innerHTML = '';
    prefixes.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p;
        list.appendChild(opt);
    });
}

function renderBlock1Options() {
    const list = document.getElementById('block1-list');
    if (!list) return;
    const blocks = typeof Block1Manager !== 'undefined' ? Block1Manager.getAll() : [];
    list.innerHTML = '';
    blocks.forEach(b => {
        const opt = document.createElement('option');
        opt.value = b;
        list.appendChild(opt);
    });
}

function deleteCurrentPrefix() {
    const input = document.getElementById('smart-prefix');
    const val = input.value.trim().toUpperCase();
    if (!val) return;

    if (confirm(`ต้องการลบหมวด "${val}" ออกจากรายการที่จำไว้หรือไม่?`)) {
        PrefixManager.remove(val);
        input.value = '';
        renderPrefixOptions();
    }
}

function deleteCurrentBlock1() {
    const input = document.getElementById('smart-block1');
    const val = input.value.trim();
    if (!val) return;

    if (confirm(`ต้องการลบเลขชุดหน้า "${val}" ออกจากรายการที่จำไว้หรือไม่?`)) {
        Block1Manager.remove(val);
        input.value = '';
        renderBlock1Options();
    }
}

function calculateSmartCheckDigit() {
    const b1Input = document.getElementById('smart-block1');
    const b2Input = document.getElementById('smart-block2');
    const checkInput = document.getElementById('smart-check-digit');

    if (!b1Input || !b2Input) return;

    const b1 = b1Input.value.replace(/\D/g, '');
    const b2 = b2Input.value.replace(/\D/g, '');

    if (b1.length === 4 && b2.length === 4) {
        const val = b1 + b2;
        const cd = TrackingUtils.calculateS10CheckDigit(val);
        checkInput.value = cd;
    } else {
        checkInput.value = '-';
    }
}

function toggleSmartRange() {
    const isChecked = document.getElementById('smart-range-enable').checked;
    const box = document.getElementById('smart-range-box');
    if (isChecked) {
        box.classList.remove('hidden');
    } else {
        box.classList.add('hidden');
        document.getElementById('smart-qty').value = '';
        if(document.getElementById('smart-box')) document.getElementById('smart-box').value = '';
        if(document.getElementById('smart-book')) document.getElementById('smart-book').value = '';
    }
}

function addSmartEntryAndSave() {
    // Validate Customer Name first
    const dbNameInput = document.getElementById('db-name');
    const name = dbNameInput ? dbNameInput.value.trim() : '';
    if (!name) {
        alert('กรุณากรอก "ชื่อลูกค้า / บริษัท" ก่อนเริ่มสแกนหรือบันทึกเลขพัสดุ');
        if(dbNameInput) dbNameInput.focus();
        return;
    }

    const type = document.getElementById('db-type') ? document.getElementById('db-type').value : 'Credit';
    const contract = document.getElementById('db-contract') ? document.getElementById('db-contract').value.trim() : '';
    const requestDate = document.getElementById('db-request-date') ? document.getElementById('db-request-date').value : '';

    const prefix = document.getElementById('smart-prefix').value.trim().toUpperCase();
    const b1 = document.getElementById('smart-block1').value.replace(/\D/g, '');
    const b2 = document.getElementById('smart-block2').value.replace(/\D/g, '');
    const qtyStr = document.getElementById('smart-qty').value.replace(/\D/g, '');
    const suffix = document.getElementById('smart-suffix').value.trim().toUpperCase();
    const isRange = document.getElementById('smart-range-enable').checked;

    // Validation
    if (prefix.length !== 2) {
        alert('กรุณากรอก อักษร (Prefix) 2 ตัว'); return;
    }
    if (b1.length !== 4) {
        alert('กรุณากรอก เลขชุดหน้า (4 หลัก)'); return;
    }
    if (b2.length !== 4) {
        alert('กรุณากรอก เลขชุดหลัง (4 หลัก)'); return;
    }
    if (suffix.length !== 2) {
        alert('กรุณากรอก Suffix 2 ตัวอักษร'); return;
    }

    // Save Memories
    PrefixManager.add(prefix);
    Block1Manager.add(b1); // Save Block 1
    renderPrefixOptions();
    renderBlock1Options();

    let itemsToAdd = [];

    // --- Smart Input Logic ---
    if (isRange) {
        let qty = parseInt(qtyStr, 10);
        if (isNaN(qty) || qty <= 0) {
            alert('กรุณาระบุจำนวนที่ถูกต้อง');
            return;
        }
        
        let startNum = parseInt(b1 + b2, 10);
        for (let i = 0; i < qty; i++) {
            let currentNumStr = (startNum + i).toString().padStart(8, '0');
            if (currentNumStr.length > 8) {
                alert('เลขรันเกิน 8 หลักแล้ว ระบบหยุดการทำงาน');
                break;
            }
            let checkDigit = TrackingUtils.calculateS10CheckDigit(currentNumStr);
            if (checkDigit !== null) {
                itemsToAdd.push(`${prefix}${currentNumStr}${checkDigit}${suffix}`);
            }
        }
    } else {
        let currentNumStr = b1 + b2;
        let checkDigit = TrackingUtils.calculateS10CheckDigit(currentNumStr);
        if (checkDigit !== null) {
            itemsToAdd.push(`${prefix}${currentNumStr}${checkDigit}${suffix}`);
        } else {
            alert('ไม่สามารถคำนวณ Check Digit ได้');
            return;
        }
    }

    if(itemsToAdd.length === 0) return;

    // Save to DB directly
    const batchInfo = { name, type, contract, requestDate, timestamp: new Date().getTime() };
    const savedCount = CustomerDB.addBatch(batchInfo, itemsToAdd);

    alert(`บันทึกเรียบร้อย! เพิ่ม ${savedCount} รายการ สำเร็จ`);

    // Complete cleanup for UI
    document.getElementById('db-name').value = '';
    document.getElementById('db-contract').value = '';
    document.getElementById('db-request-date').value = '';
    document.getElementById('smart-block2').value = '';
    document.getElementById('smart-check-digit').value = '-';
    
    // Refresh tables
    if(typeof updateDbViews === 'function') updateDbViews();
}

async function handleImageSelection(files) {
    if (!files || files.length === 0) return;

    const statusEl = document.getElementById('smart-ocr-status');
    const textArea = document.getElementById('db-tracking-list');

    // UI Feedback (Generic attempt)
    if (statusEl) statusEl.textContent = `Processing ${files.length} images...`;

    let combinedText = "";

    try {
        for (let i = 0; i < files.length; i++) {
            if (statusEl) statusEl.textContent = `Scanning image ${i + 1}/${files.length}...`;
            const file = files[i];

            // Tesseract OCR
            const worker = typeof Tesseract !== 'undefined' ? await Tesseract.createWorker('tha+eng') : null;
            if (worker) {
                const { data: { text } } = await worker.recognize(file);
                await worker.terminate();
                combinedText += "\n" + text;
            }
        }

        if (statusEl) statusEl.textContent = "Analyzing...";

        // Output Logic
        if (textArea) {
            const resultMsg = processSmartInput(combinedText, textArea);
            if (statusEl) {
                statusEl.textContent = resultMsg || "Done";
                setTimeout(() => statusEl.textContent = "", 3000);
            }
        }

    } catch (err) {
        console.error(err);
        if (statusEl) statusEl.textContent = "Error: " + err.message;
        alert("OCR Error: " + err.message);
    }

    // Reset specific input if exists
    const uploadInput = document.getElementById('smart-image-upload');
    if (uploadInput) uploadInput.value = '';
}

function processSmartInput(text, textArea) {
    // 1. Try Tracking Numbers
    const extracted = TrackingUtils.extractTrackingNumbers(text);

    if (extracted && extracted.length > 0) {
        insertToTextArea(textArea, extracted.join('\n'));
        return `Found ${extracted.length} Tracking Numbers`;
    }

    // 2. Fallback: Prices
    const prices = TrackingUtils.extractPrices(text);
    if (prices && prices.length > 0) {
        const summary = TrackingUtils.summarizePrices(prices);

        if (confirm(`ไม่พบเลขพัสดุ แต่พบยอดเงิน ${summary.totalCount} รายการ\nยอดรวม: ${summary.totalValue.toFixed(2)}\n\nต้องการดูสรุปยอดหรือไม่?`)) {
            let report = `--- Price Summary ---\n`;
            summary.groupings.forEach(g => {
                report += `${g.price.toFixed(2)} x ${g.count} = ${g.total.toFixed(2)}\n`;
            });
            report += `---------------------\n`;
            report += `Total Items: ${summary.totalCount}\n`;
            report += `Total Value: ${summary.totalValue.toFixed(2)}`;

            insertToTextArea(textArea, report);
            return `Extracted Price Summary`;
        }
    }

    return "No tracking numbers found";
}

function insertToTextArea(textArea, content) {
    const startPos = textArea.selectionStart;
    const endPos = textArea.selectionEnd;
    const textBefore = textArea.value.substring(0, startPos);
    const textAfter = textArea.value.substring(endPos, textArea.value.length);

    textArea.value = textBefore + (textBefore && textBefore.slice(-1) !== '\n' ? '\n' : '') + content + (textAfter && textAfter[0] !== '\n' ? '\n' : '') + textAfter;

    textArea.selectionStart = textArea.selectionEnd = startPos + content.length + 1;
}

function calculateSmartQtyFromHelpers() {
    const boxInput = document.getElementById('smart-box');
    const bookInput = document.getElementById('smart-book');
    const qtyInput = document.getElementById('smart-qty');
    
    if (!boxInput || !bookInput || !qtyInput) return;
    
    const boxes = parseInt(boxInput.value) || 0;
    const books = parseInt(bookInput.value) || 0;
    
    // 1 Box = 50 Books, 1 Book = 240 Items
    const totalBooks = (boxes * 50) + books;
    const totalItems = totalBooks * 240;
    
    if (totalItems > 0) {
        qtyInput.value = totalItems;
    } else {
        qtyInput.value = '';
    }
}

// ==========================================
// Book Calculator / Magic Tools Logic
// ==========================================

window.applyMagicSingle = function(inputVal) {
    const input = inputVal.trim().toUpperCase();
    if(!input) return;
    
    const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
    const match = input.match(regex);
    
    if(!match) {
        alert("รูปแบบเลขพัสดุไม่ถูกต้อง กรุณาระบุเลขพัสดุเต็ม 13 หลัก (เช่น EQ183240125TH)");
        return;
    }
    
    const [, prefix, bodyStr, checkD, suffix] = match;
    const bodyNum = parseInt(bodyStr, 10);
    
    const bookStart = Math.floor((bodyNum - 1) / 240) * 240 + 1;
    const startStr = bookStart.toString().padStart(8, '0');
    
    // Auto fill inputs
    document.getElementById('smart-prefix').value = prefix;
    document.getElementById('smart-block1').value = startStr.substring(0, 4);
    document.getElementById('smart-block2').value = startStr.substring(4, 8);
    document.getElementById('smart-suffix').value = suffix;
    
    // Check Batch Mode and set to 1 book
    const rangeEnable = document.getElementById('smart-range-enable');
    if(!rangeEnable.checked) {
        rangeEnable.checked = true;
        toggleSmartRange();
    }
    document.getElementById('smart-box').value = '';
    document.getElementById('smart-book').value = '1';
    
    // Trigger tick and recalculate Qty
    calculateSmartCheckDigit();
    calculateSmartQtyFromHelpers();
    
    // Clear input
    document.getElementById('magic-single').value = '';
    
    const sampleCd = TrackingUtils.calculateS10CheckDigit(startStr);
    alert(`ดึงข้อมูลเล่มสมบูรณ์ ชี้เป้าแม่นยำ!\nเริ่มตั้งต้นที่: ${prefix}${startStr}${sampleCd !== null ? sampleCd : '-'}${suffix}\nระบบได้ตั้งค่าจำนวน "1 เล่ม" ค้างไว้ให้แล้ว สามารถกดบันทึกข้อมูลได้เลยครับ`);
};

window.applyMagicBookNo = function(bookNoStr) {
    const bookNo = parseInt(bookNoStr, 10);
    if(isNaN(bookNo) || bookNo < 1) {
        if(bookNoStr.trim() !== '') alert("โปรดระบุเลขที่เล่มให้ถูกต้อง");
        return;
    }
    
    const prefixInput = document.getElementById('smart-prefix');
    if(!prefixInput.value.trim()) {
        prefixInput.value = 'EQ'; // Default guess if empty
    }
    
    const bookStart = (bookNo - 1) * 240 + 1;
    const startStr = bookStart.toString().padStart(8, '0');
    
    document.getElementById('smart-block1').value = startStr.substring(0, 4);
    document.getElementById('smart-block2').value = startStr.substring(4, 8);
    
    // Trigger check digit
    calculateSmartCheckDigit();
    
    // Auto enable Batch Form
    const rangeEnable = document.getElementById('smart-range-enable');
    if(!rangeEnable.checked) {
        rangeEnable.checked = true;
        toggleSmartRange();
    }
    
    document.getElementById('magic-bookno').value = '';
    
    // Focus on Qty to remind them to input how many books
    const qtyInput = document.getElementById('smart-book');
    qtyInput.focus();
    
    alert(`ตั้งค่าเตรียมสร้างจากเล่มที่ ${bookNo} สำเร็จ\nเลขเริ่มต้นคือ ${startStr}\nกรุณาระบุจำนวน 'เล่ม/กล่อง' ด้านล่างต่อเลยครับ`);
};
