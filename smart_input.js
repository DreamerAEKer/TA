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

    if (isRange && qtyStr && parseInt(qtyStr) > 0) {
        // Quantity Generation
        const qty = parseInt(qtyStr);
        const startNum = parseInt(b2);

        // Safety Limit
        if (qty > 1000) {
            if (!confirm('จำนวนมากกว่า 1,000 รายการ อาจทำให้เครื่องช้า ยืนยันทำต่อ?')) return;
        }

        const lines = [];
        for (let i = 0; i < qty; i++) {
            const currentNum = startNum + i;
            let currentB2 = currentNum.toString().padStart(4, '0');

            // Handle overflow if 9999 -> 10000. 
            // Standard tracking number logic usually implies overflow affects upper digits, 
            // but for this manual tool, we'll just stop if it breaks 4-digit structure to keep it safe.
            if (currentB2.length > 4) {
                break;
            }

            const body = b1 + currentB2;
            const cd = TrackingUtils.calculateS10CheckDigit(body);
            lines.push(`${prefix}${body}${cd}${suffix}`);
        }
        newText = lines.join('\n');

    } else {
        // Single Entry
        const body = b1 + b2;
        const cd = TrackingUtils.calculateS10CheckDigit(body);
        newText = `${prefix}${body}${cd}${suffix}`;
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
