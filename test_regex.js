const text = `24/03/2569 ปณ.3 E (E) - 10501 ปณก. EQ089801455TH ใส่ของลงถุง 19/03/2569 17:03 piyanat.di
24/03/2569 ปณ.3 E (E) - 10501 ปณก. EQ089790404TH ใส่ของลงถุง 20/03/2569 15:03 wanchana.th`;

const trackPattern = /[a-zA-Z]{2}\d{9}[a-zA-Z]{2}/g;
const datetimePattern = /(\d{2}\/\d{2}\/\d{4})\s+(\d{2}:\d{2})/g;
const trackMap = new Map();
const tracks = [];
let match;
while ((match = trackPattern.exec(text)) !== null) tracks.push({track: match[0].toUpperCase(), index: match.index});
const dates = [];
let dMatch;
while ((dMatch = datetimePattern.exec(text)) !== null) dates.push({dtStr: dMatch[1] + ' ' + dMatch[2], index: dMatch.index});

tracks.forEach((t, i) => {
    const nextIdx = (i < tracks.length - 1) ? tracks[i+1].index : text.length;
    const vDate = dates.find(d => d.index > t.index && d.index < nextIdx);
    trackMap.set(t.track, vDate ? vDate.dtStr : 'FAIL');
});
console.log(Array.from(trackMap.entries()));
