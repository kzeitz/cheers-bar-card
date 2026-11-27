// Weekly handicap audit: compares the handicaps recorded in DartConnect match recaps
// with what the Bar Card rules say they should have been.
// Usage: node source/audit/audit.js source/card/bar_card_template.html source/card/data.json source/data/recaps/<date>.json > reports/audit_<date>.md
const fs=require('fs');
const [tpl,dataF,recF]=process.argv.slice(2);
const t=fs.readFileSync(tpl,'utf8'); const data=JSON.parse(fs.readFileSync(dataF,'utf8')); const recs=JSON.parse(fs.readFileSync(recF,'utf8'));
eval(t.split('/* ---------- pure calculation')[1].split('/* ---------- end pure')[0].replace(/^[^\n]*\n/,''));
const S=data.settings, key={}; for(const [k,v] of Object.entries(data.full)) key[v.trim().toLowerCase()]=k;
const K=n=>key[n.trim().replace(/\s+/g,' ').toLowerCase()]||null;
let out=[], tot={ok:0,bad:0,skip:0};
for(const m of recs){
  const ps=playerStats(snapshotFor(data,m.date)); const notes={};
  for(const s of m.notes){const mm=s.match(/^Leg 1\.(\d+): (.*)$/); if(mm) notes[+mm[1]]=mm[2];}
  out.push(`\n### ${m.home} (home) v ${m.away}\n[DartConnect recap](${m.recap})\n\n| Game | Players | Recorded | Should be | |\n|---|---|---|---|---|`);
  for(const g of m.games){
    const H=g.home.map(K), A=g.away.map(K), rec=notes[g.n]||'none';
    const R=[...H,...A].map(k=>k?resolve(k,ps,null):{basis:'missing'});
    const who=`${g.home.join(' & ')} v ${g.away.join(' & ')}`;
    let exp;
    if(R.some(p=>p.basis==='missing')){ out.push(`| ${g.n} ${g.type} | ${who} | ${rec} | can't check (no league stats yet) | – |`); tot.skip++; continue; }
    if(/^701/.test(g.type)) exp='none';
    else if(/Cricket/.test(g.type)){ const h=R[0].mpr+R[1].mpr,a=R[2].mpr+R[3].mpr; const mk=band(S.marks,Math.abs(h-a)); exp=mk?`${h<a?'Home':'Away'} handicap: ${mk}M`:'none'; }
    else if(H.length===2){ const h=(R[0].tda+R[1].tda)/2,a=(R[2].tda+R[3].tda)/2; const sp=band(S.spots,Math.abs(h-a)); exp=sp?`Starting score: ${h<a?(501-sp)+'-501':'501-'+(501-sp)}`:'none'; }
    else { const h=R[0].tda,a=R[1].tda,st=Math.round(501*Math.min(h,a)/Math.max(h,a)); exp=st<501?`Starting score: ${h<a?st+'-501':'501-'+st}`:'none'; }
    const ok=exp===rec; ok?tot.ok++:tot.bad++;
    out.push(`| ${g.n} ${g.type} | ${who} | ${rec} | ${exp} | ${ok?'✓':'**✗**'} |`);
  }
}
console.log(`# Handicap audit – ${recs[0].date}\n\nRecorded in DartConnect vs the Bar Card rules (stats through ${snapshotFor(data,recs[0].date).through}).\n\n**${tot.ok} correct, ${tot.bad} different, ${tot.skip} not checkable.**`+out.join('\n'));
