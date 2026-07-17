const fs = require('fs');
const oldCode = fs.readFileSync('app.js', 'utf8');
const newLogic = fs.readFileSync('temp_logic.js', 'utf8');

const lines = oldCode.split('\n');
let start = -1;
let end = -1;
for(let i=0; i<lines.length; i++) {
    if(lines[i].includes('async function analyzeImportData(rows, file) {')) {
        start = i;
    }
    if(start !== -1 && lines[i].includes('uploadToFirebaseAuto(file);')) {
        end = i + 2; // include closing brace
        break;
    }
}

if (start !== -1 && end !== -1) {
    let before = lines.slice(0, start).join('\n');
    let after = lines.slice(end).join('\n');
    // find renderAnalysisTable
    let rStart = -1;
    let rEnd = -1;
    const afterLines = after.split('\n');
    for(let i=0; i<afterLines.length; i++) {
        if(afterLines[i].includes('function renderAnalysisTable')) {
            rStart = i;
        }
        if(rStart !== -1 && afterLines[i].includes("document.getElementById('import-preview').classList.remove('hidden');")) {
            rEnd = i + 2;
            break;
        }
    }
    
    if (rStart !== -1 && rEnd !== -1) {
        let afterBeforeRender = afterLines.slice(0, rStart).join('\n');
        let afterAfterRender = afterLines.slice(rEnd).join('\n');
        
        const finalCode = before + '\n' + newLogic + '\n' + afterBeforeRender + '\n' + afterAfterRender;
        fs.writeFileSync('app.js', finalCode);
        console.log('Successfully replaced logic.');
    } else {
        console.log('Could not find renderAnalysisTable');
    }
} else {
    console.log('Could not find analyzeImportData');
}
