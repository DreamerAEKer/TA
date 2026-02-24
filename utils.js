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
            isCenter: (i === startNumObj)
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

    // Pattern: 
    // Prefix (2 letters)
    // Optional Spaces
    // Digits (Sequence of 9 digits, potentially separated by spaces)
    // Optional Spaces
    // Suffix (2 letters, Optional)

    // We will find all candidates that look like tracking numbers.
    const regex = /([A-Za-z]{2})\s*([0-9\s]{9,})\s*([A-Za-z]{0,2})/g;

    let match;
    while ((match = regex.exec(text)) !== null) {
        const fullMatch = match[0];
        const prefix = match[1].toUpperCase();
        const rawDigits = match[2].replace(/\s/g, ""); // Remove inner spaces from digits
        const suffix = match[3] ? match[3].toUpperCase() : "";

        // We only care if we have exactly 9 digits for S10
        if (rawDigits.length !== 9) {
            continue;
        }

        let candidateSuffix = suffix;

        // Logic: If suffix is missing (empty), assume 'TH' (Thailand) as per user requirements for 11-char codes
        if (candidateSuffix.length === 0) {
            candidateSuffix = "TH";
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

window.TrackingUtils = {
    calculateS10CheckDigit,
    validateTrackingNumber,
    formatTrackingNumber,
    generateTrackingRange,
    groupRangesByPrice,
    virtualOptimizeRanges,
    cleanTrackingText,
    extractTrackingNumbers,
    extractPrices,
    summarizePrices,
    getWeightFromPriceA3
};
