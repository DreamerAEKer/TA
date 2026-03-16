const text = `
1 11/03/2569 ปณ.3 E (E) - 10501 ปณก. ET571304088TH ใส่ของลงถุง 09/03/2569 17:03 kosit.su
2 12/03/2569 ปณ.3 E (E) - 10501 ปณก. ET605444974TH ใส่ของลงถุง 11/03/2569 11:03 kanitha.po
3 13/03/2569 ปณ.3 E (E) - 10501 ปณก. EQ089840383TH ใส่ของลงถุง 11/03/2569 16:03 piyanat.di
`;

function extractTrackingWithContext(text) {
    const lines = text.split('\n');
    const results = [];
    
    for (const line of lines) {
        const cleanLine = line.trim();
        if (!cleanLine) continue;
        
        const trackingRegex = /[A-Za-z]{2}\d{9}[A-Za-z]{2}/;
        const match = cleanLine.match(trackingRegex);
        
        if (match) {
            const trackingNum = match[0];
            const trackIndex = cleanLine.indexOf(trackingNum);
            
            const beforeStr = cleanLine.substring(0, trackIndex).trim();
            const afterStr = cleanLine.substring(trackIndex + trackingNum.length).trim();
            
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
            const dateMatch = afterStr.match(/\d{2}\/\d{2}\/\d{4}/);
            if (dateMatch) {
                after1 = afterStr.substring(0, dateMatch.index).trim();
            } else {
                // fallback to just next word
                const afterTokens = afterStr.split(/\s+/).filter(t => t);
                if (afterTokens.length >= 1) {
                    after1 = afterTokens[0];
                }
            }

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

console.log(extractTrackingWithContext(text));
