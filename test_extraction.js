
// Import logic (Shim for Node.js environment since utils.js is client-side)
const fs = require('fs');
const path = require('path');

// Read utils.js content
const utilsPath = path.join(__dirname, 'utils.js');
const utilsContent = fs.readFileSync(utilsPath, 'utf8');

// Mock window object
const window = {};

// Eval the utils code to load functions into window.TrackingUtils
// We need to strip the specific export if it conflicts, but the file assigns to window.TrackingUtils
eval(utilsContent);

const { extractTrackingNumbers } = window.TrackingUtils;

console.log("=== Testing Smart Tracking Extraction ===\n");

const testCases = [
    {
        name: "Standard Clear",
        input: "ET443628139TH",
        expected: ["ET443628139TH"]
    },
    {
        name: "Spaced (User Format)",
        input: "ET 4436 2813 9 TH",
        expected: ["ET443628139TH"]
    },
    {
        name: "Partial 11-Digit (Missing Suffix)",
        input: "ET 4436 2814 2", // Should auto-append TH
        expected: ["ET443628142TH"]
    },
    {
        name: "Mixed Bulk Text (From Image)",
        input: `
            1 วินิต ท้วมลี้ 21000 ระยอง ET 4436 2813 9 TH 1 32.00
            3 อาภรณ์ วิไลรัตน์ 90110 หาดใหญ่ ET 4436 2815 6 TH
            4 จิรเดช คารพานนท์ ET 4436 2816 0 (Missing TH in OCR maybe?)
        `,
        expected: ["ET443628139TH", "ET443628156TH", "ET443628160TH"]
    },
    {
        name: "Bad Check Digit",
        input: "ET 4436 2813 8 TH", // Digit is 9, so 8 is wrong
        expected: []
    }
];

testCases.forEach(test => {
    console.log(`Test: ${test.name}`);
    console.log(`Input: "${test.input.replace(/\n\s+/g, ' ')}"`);
    const result = extractTrackingNumbers(test.input);
    console.log(`Output:`, result);

    // Simple verification
    const passed = JSON.stringify(result) === JSON.stringify(test.expected);
    console.log(`Status: ${passed ? "✅ PASS" : "❌ FAIL"}`);
    if (!passed) {
        console.log(`Expected: ${JSON.stringify(test.expected)}`);
    }
    console.log("---------------------------------------------------");
});
