/**
 * Tracking Analysis Tool - Utility Functions
 * Core logic for Thailand Post (S10 Standard) tracking numbers.
 */

// S10 Standard Weights
const WEIGHTS = [8, 6, 4, 2, 3, 5, 9, 7];

/**
 * Calculates the S10 Check Digit for a given 8-digit sequence.
 * @param {string} digitsStr - String of 8 digits (e.g., "12345678")
 * @returns {number|null} - The calculated check digit (0-9) or null if invalid input.
 */
function calculateS10CheckDigit(digitsStr) {
    if (!/^\d{8}$/.test(digitsStr)) {
        return null;
    }

    let sum = 0;
    for (let i = 0; i < 8; i++) {
        sum += parseInt(digitsStr[i]) * WEIGHTS[i];
    }

    const remainder = sum % 11;
    let checkDigit = 11 - remainder;

    if (checkDigit === 10) return 0;
    if (checkDigit === 11) return 5;

    return checkDigit;
}

/**
 * Validates a full tracking number (e.g., "XX123456789TH")
 * @param {string} trackingNumber 
 * @returns {object} - { isValid: boolean, error: string|null, suggestion: string|null }
 */
function validateTrackingNumber(trackingNumber) {
    // Basic format check: 2 letters, 9 digits, 2 letters
    const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
    const match = trackingNumber.toUpperCase().match(regex);

    if (!match) {
        return { isValid: false, error: "Invalid Format (Must be XX123456789TH)", suggestion: null };
    }

    const [full, prefix, body, currentCheckDigit, suffix] = match;
    const computedCheckDigit = calculateS10CheckDigit(body);

    if (computedCheckDigit === null) {
        return { isValid: false, error: "Internal Error", suggestion: null };
    }

    if (parseInt(currentCheckDigit) !== computedCheckDigit) {
        return {
            isValid: false,
            error: `Invalid Check Digit (Expected ${computedCheckDigit}, Found ${currentCheckDigit})`,
            suggestion: `${prefix}${body}${computedCheckDigit}${suffix}`
        };
    }

    return { isValid: true, error: null, suggestion: null };
}

/**
 * Generates a range of tracking numbers.
 * @param {string} centerNumber - Full tracking number to start from or center on.
 * @param {number} countBefore - How many numbers before (decrement).
 * @param {number} countAfter - How many numbers after (increment).
 * @param {boolean} matchPrefix - If true, only keeps same prefix.
 * @returns {Array} - Array of tracking objects.
 */
function generateTrackingRange(centerNumber, countBefore, countAfter) {
    const regex = /^([A-Z]{2})(\d{8})(\d)([A-Z]{2})$/;
    const match = centerNumber.toUpperCase().match(regex);

    if (!match) return [];

    const [full, prefix, bodyStr, cd, suffix] = match;
    const startNumObj = parseInt(bodyStr); // e.g. 12345678

    const results = [];

    // Range: from (start - before) to (start + after)
    const min = startNumObj - countBefore;
    const max = startNumObj + countAfter;

    for (let i = min; i <= max; i++) {
        if (i < 0) continue; // Should not happen for valid ranges usually

        let newBody = i.toString().padStart(8, '0');
        // Handle overflow > 99999999 ? For now just clamp or wrap. 
        // Realistically tracking numbers wrap or change prefix. 
        // We will just allow 8 digits. If length > 8, skip.
        if (newBody.length > 8) continue;

        let newCd = calculateS10CheckDigit(newBody);
        let validNumber = `${prefix}${newBody}${newCd}${suffix}`;

        results.push({
            number: validNumber,
            isCenter: (i === startNumObj),
            offset: i - startNumObj
        });
    }

    return results;
}

/**
 * Groups a list of Ranges by Price.
 * 
 * @param {Array} ranges - List of metadata ranges {price, count, start, end, ...}
 * @returns {Array} - Sorted Groups [{ price, subTotal, subCount, subRanges: [...] }]
 */
function groupRangesByPrice(ranges) {
    const groups = {}; // priceKey -> { price, subTotal, subCount, subRanges }

    ranges.forEach(r => {
        const price = r.price || 0;
        const key = price.toFixed(2); // Use string-fixed price as key

        if (!groups[key]) {
            groups[key] = {
                price: price,
                subTotal: 0,
                subCount: 0,
                subRanges: []
            };
        }

        groups[key].subTotal += (r.count * price);
        groups[key].subCount += r.count;
        groups[key].subRanges.push(r);
    });

    // Convert to Array
    const groupList = Object.values(groups);

    // Sort by Price (Ascending to match user req "น้อยไปมาก")
    groupList.sort((a, b) => a.price - b.price);

    return groupList;
}



/**
 * OPTIMIZED (VIRTUAL) GROUPING
 * Redistributes IDs to match the sorted Price Groups sequentially.
 * 
 * @param {Array} ranges - Original ranges from analysis
 * @returns {Array} - Virtual Groups [{ price, count, weight, virtualRanges: [{start, end, count}] }]
 */
function virtualOptimizeRanges(ranges) {
    // 1. Pool All Items (Sorted)
    let allItems = [];
    ranges.forEach(r => {
        // We assume r.items contains full tracking numbers (or we generate them if missing)
        // In current app.js, r.items IS populated.
        if (r.items) {
            allItems.push(...r.items);
        }
    });

    // Sort All IDs Ascending
    allItems.sort();

    // 2. Pool Price/Weight Counts
    // We group by "Price-Weight" signature to ensure weight consistency check
    const priceGroups = {};
    ranges.forEach(r => {
        const key = `${r.price.toFixed(2)}-${r.weight}`; // Composite key
        if (!priceGroups[key]) {
            priceGroups[key] = {
                price: r.price,
                weight: r.weight,
                totalCount: 0,
                originalRanges: []
            };
        }
        priceGroups[key].totalCount += r.count;
        priceGroups[key].originalRanges.push(r);
    });

    // Convert to Array & Sort by Price Asc
    const sortedGroups = Object.values(priceGroups);
    sortedGroups.sort((a, b) => a.price - b.price);

    // 3. Sequential Assignment (Remap)
    let currentIdx = 0;

    const virtualResults = sortedGroups.map(group => {
        // Must take 'totalCount' items from the sorted pooled list
        const myItems = allItems.slice(currentIdx, currentIdx + group.totalCount);
        currentIdx += group.totalCount;

        // Create Virtual Range(s) for these items
        // Even in virtual mode, if there are massive gaps in the pool, we technically should show them?
        // But user asked for "1ชุดเลขที่" (One set).
        // Let's try to condense to Min-Max.
        if (myItems.length === 0) return null;

        return {
            price: group.price,
            weight: group.weight,
            count: group.totalCount,
            total: group.totalCount * group.price,
            start: myItems[0],
            end: myItems[myItems.length - 1],
            isVirtual: true,
            // Consistency Check: Did we just assign items that shouldn't be here?
            // (Strictly speaking we are fabricating the link, so 'correctness' relies on the pool being complete)
            check: "Optimized"
        };
    }).filter(g => g !== null);

    return virtualResults;
}

// Export for usage in app.js (if using modules) or window global
/**
 * Cleans tracking number text by removing spaces and normalizing.
 * @param {string} text 
 * @returns {string}
 */
function cleanTrackingText(text) {
    if (!text) return "";
    return text.replace(/[^A-Za-z0-9]/g, "").toUpperCase();
}

/**
 * Extracts valid tracking numbers from a broad text block.
 * Supports:
 * - Standard 13 chars: XX123456789TH
 * - Spaced formats: XX 1234 5678 9 TH
 * - 11 chars (Missing Suffix, inferred as TH): XX123456789 -> XX123456789TH
 * 
 * Validates candidates using S10 Check Digit.
 * 
 * @param {string} text - The raw text input (e.g. from OCR or Paste)
 * @returns {Array} - Array of unique valid tracking strings.
 */
function extractTrackingNumbers(text) {
    const validNumbers = new Set();

    // Regex explanation:
    // ([A-Z]{2})       : Group 1 - Prefix (2 letters)
    // \s*              : Optional spaces
    // ([0-9]+)         : Group 2 - Digits (Greedy, we'll slice/parse later or enforce length in stricter regex)
    //                    Let's use a more structured regex to catch the specific user format "4436 2813 9"
    // ([0-9]{8})       : Body of 8 digits? No, user saw "4436 2813 9". That's 4+4+1 = 9 digits. Good.
    // Let's try to capture the pattern loosely then validate.

    // Strategy: Look for Sequence: 2 letters + many digits + optional 2 letters
    // We expect 9 digits total.

    // Regex: 
    // ([A-Za-z]{2}) : Prefix
    // \s*           : Optional spaces before digits
    // ((?:\d\s*){9}): 9 digits with optional spaces between them (e.g., "6 0 5 4  4 5 0 2  5")
    // \s*           : Optional spaces before suffix
    // ([A-Za-z]{0,2}): Optional 2-letter suffix

    const regex = /([A-Za-z]{2})\s*((?:\d\s*){9})\s*([A-Za-z]{0,2})/g;

    let match;
    while ((match = regex.exec(text)) !== null) {
        const fullMatch = match[0];
        const prefix = match[1].toUpperCase();
        // Remove spaces inside the digit group to get the 9 digits
        const rawDigits = match[2].replace(/\s+/g, ""); 
        const suffix = match[3] ? match[3].toUpperCase() : "";

        // Should theoretically be exactly 9 digits by the regex
        if (rawDigits.length !== 9) continue;

        let candidateSuffix = suffix;

        // Logic: If suffix is missing, assume 'TH'
        if (candidateSuffix.length === 0) {
            candidateSuffix = "TH";
        } else if (candidateSuffix.length === 1) {
            candidateSuffix += "H"; // Rough fallback, assume ends in TH
        }

        // Must have 2-char suffix now
        if (candidateSuffix.length !== 2) {
            continue;
        }

        // Construct candidate
        const candidateDateCheck = `${prefix}${rawDigits}${candidateSuffix}`;

        // Validate Check Digit
        const validation = validateTrackingNumber(candidateDateCheck);
        if (validation.isValid) {
            validNumbers.add(candidateDateCheck);
        }
    }

    return Array.from(validNumbers);
}

/**
 * Extracts tracking numbers along with contextual data from the surrounding text (e.g. QMS screenshot).
 * Captures Zipcode (pre2), Branch (pre1), and Status (after1).
 * 
 * @param {string} text - Raw OCR text
 * @returns {Array} - Array of objects containing { trackingNumber, pre2, pre1, after1, rawLine }
 */
function extractTrackingWithContext(text) {
    const lines = text.split('\n');
    const results = [];
    
    for (const line of lines) {
        const cleanLine = line.trim();
        if (!cleanLine) continue;
        
        // Flexible regex for Tracking ID
        const trackingRegex = /[A-Za-z]{2}\s*([0-9\s]{9,})\s*[A-Za-z]{0,2}/;
        const strictTrackingRegex = /[A-Za-z]{2}\d{9}[A-Za-z]{2}/;
        
        // Try to match strict first, if not try loose mapping like extractTrackingNumbers
        let match = cleanLine.match(strictTrackingRegex);
        let trackingNum = '';
        
        if (match) {
            trackingNum = match[0];
        } else {
             // Fallback to finding something that looks like tracking in the line and cleaning it up to see if valid
            const looseMatch = cleanLine.match(trackingRegex);
            if (looseMatch) {
                const candidates = extractTrackingNumbers(looseMatch[0]);
                if (candidates.length > 0) {
                     trackingNum = candidates[0]; // Take first valid
                     // To slice properly, we need the raw string that mapped to this candidate
                     match = looseMatch; 
                }
            }
        }
        
        if (trackingNum && match) {
            const trackIndex = cleanLine.indexOf(match[0]);
            
            const beforeStr = cleanLine.substring(0, trackIndex).trim();
            const afterStr = cleanLine.substring(trackIndex + match[0].length).trim();
            
            const beforeTokens = beforeStr.split(/\s+/).filter(t => t);
            let pre2 = '-', pre1 = '-';
            if (beforeTokens.length >= 2) {
                pre1 = beforeTokens[beforeTokens.length - 1];
                pre2 = beforeTokens[beforeTokens.length - 2];
            } else if (beforeTokens.length === 1) {
                pre1 = beforeTokens[0];
            }
            
            let after1 = '-';
            // Find the date DD/MM/YYYY
            const dateMatch = afterStr.match(/\d{2}\/\d{2}(\/\d{2,4})?/); // Matches DD/MM or DD/MM/YYYY
            if (dateMatch) {
                after1 = afterStr.substring(0, dateMatch.index).trim();
            } else {
                // fallback to just next word
                const afterTokens = afterStr.split(/\s+/).filter(t => t);
                if (afterTokens.length >= 1) {
                    after1 = afterTokens[0];
                }
            }

            // Optional: If after1 is too long (contains multiple words but not date), we might want to truncate, but let's keep it as is.
            if(after1.length === 0) after1 = '-';

            results.push({
                trackingNumber: trackingNum.toUpperCase(),
                pre2: pre2,
                pre1: pre1,
                after1: after1,
                rawLine: cleanLine
            });
        }
    }
    
    return results;
}

/**
 * Extracts price-like values (e.g. 32.00, 474.00) from text.
 * @param {string} text 
 * @returns {Array} Array of numbers (floats)
 */
function extractPrices(text) {
    if (!text) return [];
    // Regex for money: 1-5 digits, dot, 2 digits. e.g. 32.00, 1500.50
    const regex = /\b(\d{1,5}\.\d{2})\b/g;
    const matches = text.match(regex);
    if (!matches) return [];

    // Convert to numbers
    return matches.map(m => parseFloat(m));
}

/**
 * Parses lines specifically formatted like a Thai Post sender summary table.
 * Looks for Tracking IDs and optionally a price or weight on the same line.
 * @param {string} text 
 * @returns {Array} List of { number, price, weight, hasDiscrepancy, originalWeight }
 */
function extractHandwrittenTable(text) {
    const lines = text.split('\n');
    const results = [];
    
    // Looser regex to find tracking ID and a following number that could be a price (2-4 digits mostly)
    // The table structure often has tracking ID in the middle and price at the far right.
    // Example line: "2 มหาวิทยาลัยขอนแก่น WMD 40002 EQ021239785TH 42"
    // Or weird spaced OCR: "ET 6054 4998 8 TH L. nsu Ny"
    
    for (const line of lines) {
        const cleanLine = line.trim().toUpperCase();
        if (!cleanLine) continue;
        
        // Find tracking number securely (our updated extractTrackingNumbers handles spaces)
        let trackingCandidates = extractTrackingNumbers(cleanLine);
        
        if (trackingCandidates.length > 0) {
            const trackNum = trackingCandidates[0];
            
            // Because OCR spacing on the tracking ID is unpredictable in the original string (e.g., "ET 6054 4998 8 TH"),
            // we split the line around the *last 4 digits* instead of the exact string.
            // 9 digits total, check dig is 9th. Let's find index by sweeping digits.
            // A safer way is to find the index of the last digit in the original line.
            
            // Alternatively, extract Prices from the ENTIRE line and assume the LAST sensible price belongs to it.
            // Because standard lines look like: "... ID ... PRICE"
            const matches = cleanLine.match(/\d{2,4}(?:\.\d{2})?/g);
            let price = 0;
            
            if (matches && matches.length > 0) {
                 // Scan backwards for a reasonable price value
                 for (let i = matches.length - 1; i >= 0; i--) {
                     const possiblePriceStr = matches[i];
                     const possiblePrice = parseInt(possiblePriceStr, 10);
                     
                     // Postal price bounds on handwritten receipts usually 15 - 5000 
                     // Ignore digits block that might just be postal codes like "10110", or "40002" if it's over 5000
                     if (possiblePrice >= 15 && possiblePrice <= 5000) {
                         // Extra check: prevent snagging parts of the tracking number (which are usually spaced like "6054" or "4502")
                         // It's unlikely a 4-digit tracking number block is the *last* thing on a line, but just in case:
                         price = possiblePrice;
                         break; // Found the rightmost price
                     }
                 }
            }
            
            // Calculate weight based on price logic (A3 package logic)
            let weight = '-';
            let hasDiscrepancy = false;
            let originalWeight = '-';
            
            if (price > 0 && typeof getWeightFromPriceA3 !== 'undefined') {
                 weight = getWeightFromPriceA3(price);
            }
            
            results.push({
                number: trackNum,
                price: price,
                weight: weight,
                hasDiscrepancy: hasDiscrepancy,
                originalWeight: originalWeight
            });
        }
    }
    
    // Sort array so that identical prices group together nicely? 
    // Usually keep original order, let analyzeImportedRanges sort it.
    return results;
}

/**
 * Summarizes a list of prices into groups.
 * @param {Array} prices 
 * @returns {object} { groupings: [], totalCount, totalValue }
 */
function summarizePrices(prices) {
    const groups = {};
    let totalValue = 0;

    prices.forEach(p => {
        const key = p.toFixed(2);
        if (!groups[key]) {
            groups[key] = { price: p, count: 0, total: 0 };
        }
        groups[key].count++;
        groups[key].total += p;
        totalValue += p;
    });

    const groupings = Object.values(groups).sort((a, b) => a.price - b.price);

    return {
        groupings,
        totalCount: prices.length,
        totalValue
    };
}

/**
 * Formats tracking number as XX 0000 0000 0 TH
 * @param {string} trackingNumber 
 * @returns {string} Formatted tracking number
 */
function formatTrackingNumber(trackingNumber) {
    if (!trackingNumber || trackingNumber.length !== 13) return trackingNumber;
    return `${trackingNumber.substring(0, 2)}\u202F${trackingNumber.substring(2, 6)}\u202F${trackingNumber.substring(6, 10)}\u202F${trackingNumber.substring(10, 11)}\u202F${trackingNumber.substring(11, 13)}`;
}

/**
 * Calculates Package A3 weight based on price.
 * @param {number} price 
 * @returns {string} Weight string e.g. "2 กก" or "-" if unknown
 */
function getWeightFromPriceA3(price) {
    if (!price || price <= 0) return '-';

    // Up to 10kg
    if (price >= 19 && price <= 109) {
        let w = (price - 9) / 10;
        if (Number.isInteger(w)) return `${w} กก`;
    }
    // 11kg to 20kg
    if (price >= 114 && price <= 159) {
        let w = (price - 59) / 5;
        if (Number.isInteger(w)) return `${w} กก`;
    }
    // > 20kg
    if (price > 159) {
        let excess = (price - 159) / 15;
        if (Number.isInteger(excess)) {
            return `${20 + excess} กก`;
        }
    }

    return '-'; // Fallback
}

/**
 * Compresses an image based on dataUrl.
 * @param {string} dataUrl - Base64 data URL.
 * @param {number} maxWidth - Maximum width (or height) in pixels.
 * @param {number} quality - JPEG quality (0.0 to 1.0).
 * @returns {Promise<string>} - Promise resolving to compressed data URL.
 */
async function compressImage(dataUrl, maxWidth = 1000, quality = 0.7) {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            let width = img.width;
            let height = img.height;

            if (width > height) {
                if (width > maxWidth) {
                    height *= maxWidth / width;
                    width = maxWidth;
                }
            } else {
                if (height > maxWidth) {
                    width *= maxWidth / height;
                    height = maxWidth;
                }
            }

            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, width, height);
            // Export as compressed JPEG
            resolve(canvas.toDataURL('image/jpeg', quality));
        };
        img.onerror = () => resolve(dataUrl); // Fallback to original
        img.src = dataUrl;
    });
}

window.TrackingUtils = {
    calculateS10CheckDigit,
    validateTrackingNumber,
    formatTrackingNumber,
    generateTrackingRange,
    groupRangesByPrice,
    virtualOptimizeRanges,
    cleanTrackingText,
    extractTrackingNumbers,
    extractTrackingWithContext,
    extractPrices,
    summarizePrices,
    getWeightFromPriceA3,
    extractHandwrittenTable,
    compressImage, // Added
    parseThaiDateBE // Added
};

/**
 * Parses a Thai Buddhist Era date string (DD/MM/YYYY) into a CE Date object.
 * @param {string} dateStr - e.g. "28/03/2569 11:03"
 * @returns {Date|null}
 */
function parseThaiDateBE(dateStr) {
    if (!dateStr) return null;
    // Match DD/MM/YYYY HH:MM or DD/MM/YYYY
    // Group 1: Day, 2: Month, 3: Year (BE), 4: Hour (optional), 5: Minute (optional)
    const m = dateStr.match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
    if (!m) return null;
    
    const dd = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10) - 1; // 0-indexed
    const yyyy_be = parseInt(m[3], 10);
    const yyyy_ce = yyyy_be - 543;
    
    const hh = m[4] ? parseInt(m[4], 10) : 0;
    const min = m[5] ? parseInt(m[5], 10) : 0;
    
    const date = new Date(yyyy_ce, mm, dd, hh, min);
    return isNaN(date.getTime()) ? null : date;
}

/**
 * Prefix & Block1 Memory Managers
 */
const PrefixManager = {
    key: 'thp_prefix_memories',
    getAll: () => JSON.parse(localStorage.getItem(PrefixManager.key) || '["EA", "EB", "EF", "EG", "EH", "EI", "EK", "EM", "EO", "EQ", "ER", "ET", "EU", "EW", "EX", "EY", "EZ"]'),
    add: (p) => {
        const all = PrefixManager.getAll();
        if (!all.includes(p)) {
            all.push(p);
            localStorage.setItem(PrefixManager.key, JSON.stringify(all.sort()));
        }
    },
    remove: (p) => {
        const all = PrefixManager.getAll().filter(x => x !== p);
        localStorage.setItem(PrefixManager.key, JSON.stringify(all));
    }
};

const Block1Manager = {
    key: 'thp_block1_memories',
    getAll: () => JSON.parse(localStorage.getItem(Block1Manager.key) || '[]'),
    add: (b) => {
        const all = Block1Manager.getAll();
        if (!all.includes(b)) {
            all.push(b);
            localStorage.setItem(Block1Manager.key, JSON.stringify(all.sort()));
        }
    },
    remove: (b) => {
        const all = Block1Manager.getAll().filter(x => x !== b);
        localStorage.setItem(Block1Manager.key, JSON.stringify(all));
    }
};

window.PrefixManager = PrefixManager;
window.Block1Manager = Block1Manager;

/**
 * UI Utilities
 */

// Simple Toast Notification System
window.showToast = function(message, type = 'success') {
    const toast = document.createElement('div');
    toast.className = `toast-notification toast-${type}`;
    toast.style.cssText = `
        position: fixed;
        bottom: 30px;
        left: 50%;
        transform: translateX(-50%);
        padding: 12px 24px;
        background: ${type === 'success' ? '#28a745' : type === 'error' ? '#dc3545' : '#f39c12'};
        color: white;
        border-radius: 50px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 10000;
        font-weight: bold;
        display: flex;
        align-items: center;
        gap: 10px;
        animation: toastFadeIn 0.3s ease;
    `;
    
    const icon = type === 'success' ? '✅' : type === 'error' ? '❌' : '⚠️';
    toast.innerHTML = `<span>${icon}</span> ${message}`;
    
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(-50%) translateY(20px)';
        toast.style.transition = 'all 0.5s ease';
        setTimeout(() => toast.remove(), 500);
    }, 3000);
};

// Button Loading State Helper
window.setButtonLoading = function(btn, isLoading, originalHtml) {
    if (!btn) return;
    if (isLoading) {
        btn.disabled = true;
        btn.dataset.originalHtml = btn.innerHTML;
        btn.innerHTML = `<span class="spinner"></span> กำลังบันทึก...`;
        btn.style.opacity = '0.7';
    } else {
        btn.disabled = false;
        btn.innerHTML = originalHtml || btn.dataset.originalHtml || 'บันทึก';
        btn.style.opacity = '1';
    }
};
