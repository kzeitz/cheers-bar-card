// node test_calc.js  -> checks the handicap math against known workbook results
const fs=require('fs'),path=require('path');
const t=fs.readFileSync(path.join(__dirname,'bar_card_template.html'),'utf8');
const data=JSON.parse(fs.readFileSync(path.join(__dirname,'data.json'),'utf8'));
eval(t.split('/* ---------- pure calculation')[1].split('/* ---------- end pure')[0].replace(/^[^\n]*\n/,''));
const S=data.settings, ps=playerStats(snapshotFor(data,'2026-11-30'));
const L=['Zeitz, K','Dirks, J','Cordrey, S','Rensch, S'].map(n=>resolve(n,ps,null));
const r=computeMatch(L,S), got=r.games.map(g=>g.start).join(',');
console.log(got==='433,485,413,461'&&r.spot===0&&r.marks===0?'PASS':'FAIL', got);
