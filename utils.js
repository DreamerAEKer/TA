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

// Export for usage in app.js (if using modules) or window global
window.TrackingUtils = {
    calculateS10CheckDigit,
    validateTrackingNumber,
    generateTrackingRange
};
