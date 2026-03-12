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
    }
}

function addSmartEntry() {
    // Validate Customer Name first so we don't clear inputs if it's missing
    const dbNameInput = document.getElementById('db-name');
    if (dbNameInput && !dbNameInput.value.trim()) {
        alert('กรุณากรอก "ชื่อลูกค้า / บริษัท" ก่อนเริ่มสแกนหรือบันทึกเลขพัสดุ');
        dbNameInput.focus();
        return;
    }

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

    const textArea = document.getElementById('db-tracking-list');
    const currentText = textArea.value.trim();
    let newText = '';

    // --- Smart Input Logic ---
    if (isRange) {
        let qty = parseInt(qtyStr, 10);
        if (isNaN(qty) || qty <= 0) {
            alert('กรุณาระบุจำนวนที่ถูกต้อง');
            return;
        }
        
        let startNum = parseInt(b1 + b2, 10);
        let items = [];
        for (let i = 0; i < qty; i++) {
            let currentNumStr = (startNum + i).toString().padStart(8, '0');
            if (currentNumStr.length > 8) {
                alert('เลขรันเกิน 8 หลักแล้ว ระบบหยุดการทำงาน');
                break;
            }
            let checkDigit = TrackingUtils.calculateS10CheckDigit(currentNumStr);
            if (checkDigit !== null) {
                items.push(`${prefix}${currentNumStr}${checkDigit}${suffix}`);
            }
        }
        newText = items.join('\n');
    } else {
        let currentNumStr = b1 + b2;
        let checkDigit = TrackingUtils.calculateS10CheckDigit(currentNumStr);
        if (checkDigit !== null) {
            newText = `${prefix}${currentNumStr}${checkDigit}${suffix}`;
        } else {
            alert('ไม่สามารถคำนวณ Check Digit ได้');
            return;
        }
    }

    // Append
    if (currentText) {
        textArea.value = currentText + '\n' + newText;
    } else {
        textArea.value = newText;
    }

    // Clear Input
    document.getElementById('smart-block2').value = '';
    // document.getElementById('smart-qty').value = ''; // Keep qty for convenience
    document.getElementById('smart-check-digit').value = '-';
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
