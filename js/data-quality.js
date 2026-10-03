/* =============================================================================
 * data-quality.js  —  QBR.audit(model)
 * Offline, deterministic self-audit of the uploaded workbook. Reads the raw sheet
 * rows stashed on model.raw (by excel-loader captureRaw) plus the parsed master
 * dimension, and returns a structured findings object the Data Quality tab renders.
 * No network, no LLM. Mirrors the checks in the standalone audit.
 * ===========================================================================*/
var QBR = (window.QBR = window.QBR || {});

(function () {
  var COL = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P"];
  function clean(v){ return v==null ? "" : String(v).replace(/[​-‏‪-‮﻿]/g,"").trim(); }
  function key(n){ return clean(n).toUpperCase().replace(/[^A-Z0-9 ]/g," ").replace(/\s+/g," ").trim(); }
  function toNum(v){ if(v==null||v===""||v==="-")return null; var n=typeof v==="number"?v:parseFloat(clean(v).replace(/[^0-9.\-]/g,"")); return isNaN(n)?null:n; }
  function toGB(v){ if(v==null||v==="")return null; var s=clean(v).replace(/,/g,""); var m=s.match(/([\d.]+)\s*(TB|GB|MB)?/i); if(!m)return null; var n=parseFloat(m[1]); if(isNaN(n))return null; var u=(m[2]||"GB").toUpperCase(); if(u==="TB")n*=1024; else if(u==="MB")n/=1024; return n; }
  function median(a){ a=a.filter(function(x){return x!=null&&!isNaN(x)&&x>=0;}).sort(function(x,y){return x-y;}); if(!a.length)return null; var m=a.length>>1; return a.length%2?a[m]:(a[m-1]+a[m])/2; }
  function fmtNum(n){ return n==null?"—":Math.round(n).toLocaleString(); }

  // distinct schools (key -> Set of raw names) from a raw sheet at a school column
  function distinctSchools(rows, scol){ var m=new Map(); (rows||[]).slice(1).forEach(function(r){ var s=clean(r[scol]); if(!s)return; var k=key(s); if(!m.has(k))m.set(k,new Set()); m.get(k).add(s); }); return m; }

  QBR.audit = function (model) {
    var raw = model.raw || {};
    // The audit reads SOURCE cells, which only exist when the loader stashed them.
    // Without model.raw every check below silently reports "clean" — say so instead.
    var rawMissing = !model.raw || !Object.keys(model.raw).length;
    var out = { generatedAt: new Date(), counts:{}, banners:[], coverage:null, badCells:[], vocab:[], sentinels:[], statusInNumber:[], orgGaps:null, nameVariants:[], score:null };

    if (rawMissing) out.banners.push({level:"warn",title:"Audit incomplete — raw source rows missing",detail:"The loader did not stash the workbook's raw rows (model.raw), so the cell-level checks below (bad cells, vocabulary, sentinels, encoding) could not run. Their results are not evidence that the source is clean. Re-upload the workbook; if it persists, the loader's captureRaw step needs attention."});

    /* ---- coverage matrix ---- */
    var SH = {
      risky:    { rows: raw.risky,    scol:1, label:"Risky / Domain" },
      security: { rows: raw.security, scol:0, label:"Security" },
      storage:  { rows: raw.storage,  scol:1, label:"Storage" },
      usage:    { rows: raw.usage,    scol:1, label:"Usage" },
      canva:    { rows: raw.canva,    scol:1, label:"Canva" },
      postmaster:{rows: raw.postmaster,scol:1, label:"Postmaster" },
      product:  { rows: raw.product,  scol:0, label:"Product/Services" }
    };
    var sets={}; Object.keys(SH).forEach(function(k){ sets[k]=distinctSchools(SH[k].rows, SH[k].scol); });
    var masterSize = model.master ? model.master.size : 0;
    var inRisky = sets.risky;
    function missingFrom(setObj){ var miss=[]; inRisky.forEach(function(v,k){ if(!setObj.has(k)) miss.push([].concat(Array.from(v))[0]); }); return miss; }
    out.coverage = {
      master: masterSize,
      sheets: Object.keys(SH).map(function(k){ return { key:k, label:SH[k].label, schools: sets[k]?sets[k].size:0 }; }),
      missing: {
        security: missingFrom(sets.security),
        storage:  missingFrom(sets.storage),
        usage:    missingFrom(sets.usage)
      }
    };
    // Canva-only (no MS footprint) — reported for transparency (already excluded from Domain Status)
    var canvaOnly=[]; if(sets.canva) sets.canva.forEach(function(v,k){ if(!sets.risky.has(k)&&!sets.security.has(k)&&!sets.storage.has(k)&&!sets.usage.has(k)) canvaOnly.push(Array.from(v)[0]); });
    out.coverage.canvaOnly = canvaOnly;

    /* ---- STORAGE bad-cell register (units / magnitude / capacity) ---- */
    if (raw.storage) {
      var S={q:0,school:1,od:3,ex:4,sp:5,usedTxt:6,usage:7,usedGB:8,totalGB:9};
      // sane capacity per school (from any quarter) for suggesting a fix
      var capBySchool={};
      raw.storage.slice(1).forEach(function(r){ var s=key(r[S.school]); if(!s)return; var t=toNum(r[S.totalGB]); if(t!=null&&t>0&&t<=10240000){ (capBySchool[s]=capBySchool[s]||[]).push(t/1000); } });
      function schoolCap(s){ var a=capBySchool[s]; if(!a||!a.length)return null; return median(a); }

      raw.storage.slice(1).forEach(function(r, i){
        var xl=i+2, s=clean(r[S.school]); if(!s)return; var k=key(s), q=clean(r[S.q]);
        var od=toGB(r[S.od]), ex=toGB(r[S.ex]), sp=toGB(r[S.sp]);
        var svcTB = (od!=null||ex!=null||sp!=null) ? ((od||0)+(ex||0)+(sp||0))/1024 : null;
        var uT = (function(){ var g=toGB(r[S.usedTxt]); return g==null?null:g/1024; })();
        var uC = toNum(r[S.usedGB])!=null ? toNum(r[S.usedGB])/1000 : null;
        var capCol = toNum(r[S.totalGB])!=null ? toNum(r[S.totalGB])/1000 : null;
        var M = median([uT,uC,svcTB]);
        // 1) per-service value physically exceeds capacity -> unit mislabel (TB should be GB)
        [["od",S.od,od],["ex",S.ex,ex],["sp",S.sp,sp]].forEach(function(t){
          var v=t[2]; if(v==null)return; var vTB=v/1024;
          if(capCol!=null && vTB>capCol*1.05){
            out.badCells.push({sheet:"STORAGE_DATA",cell:COL[t[1]]+xl,tenant:s,quarter:q,column:t[0]==="ex"?"Exchange":t[0]==="od"?"OneDrive":"SharePoint",
              current:clean(r[t[1]]), fix:vTB.toFixed(2)+" GB  ("+(vTB).toFixed(3)+" TB) — relabel unit TB→GB", type:"Unit mislabel"});
          }
        });
        // 2) Used-GB column magnitude typo (disagrees with text+services which agree)
        if(uC!=null && M!=null && Math.abs(uC-M)>Math.max(1,M*0.25) && (svcTB!=null||uT!=null)){
          out.badCells.push({sheet:"STORAGE_DATA",cell:COL[S.usedGB]+xl,tenant:s,quarter:q,column:"Used Storage(GB)",
            current:clean(r[S.usedGB]), fix:Math.round(M*1000)+"  ("+M.toFixed(2)+" TB)", type:"Magnitude typo"});
        }
        // 3) impossible capacity (no managed pool here exceeds ~150 TB; flag >250 TB
        //    or >3x the school's own typical capacity)
        if(capCol!=null && (capCol>250 || (schoolCap(k)!=null && capCol>schoolCap(k)*3))){
          var sc=schoolCap(k); var tt=(clean(r[S.usage]).match(/of\s*([\d.,]+)\s*TB/i)||[])[1];
          var fix = sc!=null ? Math.round(sc*1000)+"  ("+sc.toFixed(2)+" TB, from other quarter)" : (tt?Math.round(parseFloat(tt.replace(/,/g,""))*1000)+" ("+tt+" TB, from USAGE text)":"verify capacity");
          out.badCells.push({sheet:"STORAGE_DATA",cell:COL[S.totalGB]+xl,tenant:s,quarter:q,column:"Total Storage(GB)",
            current:clean(r[S.totalGB]), fix:fix, type:"Impossible capacity"});
        }
        // 4) service split unreliable (a service dwarfs the consensus used)
        if(M!=null && svcTB!=null && svcTB>M*1.5+1 && !(capCol!=null && ((od||0)/1024>capCol||(ex||0)/1024>capCol||(sp||0)/1024>capCol))){
          out.badCells.push({sheet:"STORAGE_DATA",cell:COL[S.od]+xl,tenant:s,quarter:q,column:"OneDrive/Exchange/SharePoint",
            current:"OD="+clean(r[S.od])+" EX="+clean(r[S.ex])+" SP="+clean(r[S.sp]), fix:"verify — service sum "+svcTB.toFixed(2)+" TB vs used "+M.toFixed(2)+" TB", type:"Service check"});
        }
        // 5) missing capacity but a text total exists
        if(capCol==null){ var tt2=(clean(r[S.usage]).match(/of\s*([\d.,]+)\s*TB/i)||[])[1]; if(tt2){ out.badCells.push({sheet:"STORAGE_DATA",cell:COL[S.totalGB]+xl,tenant:s,quarter:q,column:"Total Storage(GB)",current:"(blank)",fix:Math.round(parseFloat(tt2.replace(/,/g,""))*1000)+"  ("+tt2+" TB)",type:"Missing capacity"}); } }
      });
      // Q3 capacity trust banner
      var q3bad = out.badCells.filter(function(b){return b.quarter==="Q3"&&/capacity/i.test(b.type);}).length;
      if(q3bad>=3) out.banners.push({level:"warn",title:"Q3 storage capacity is unreliable",detail:"Multiple Q3 tenants have impossible Total-Storage values (and the matching “of X TB” text is also corrupt). Q3 Pooled / Utilization are suppressed — treat Q3 storage as a used-only snapshot until the source is fixed."});
    }

    /* ---- vocabulary violations (values outside the controlled set) ---- */
    function vocabCheck(rows, col, sheet, colName, okSet){
      if(!rows)return; var bad=new Map();
      rows.slice(1).forEach(function(r){ var v=clean(r[col]); if(!v)return; var u=v.toUpperCase(); var ok=false; okSet.forEach(function(x){ if(u===x||u.indexOf(x)>=0||x.indexOf(u)>=0)ok=true; }); if(!ok)bad.set(v,(bad.get(v)||0)+1); });
      bad.forEach(function(c,val){ out.vocab.push({sheet:sheet,column:colName,value:val,count:c}); });
    }
    vocabCheck(raw.risky,4,"RISKY_USERS_AND_DOMAIN","DOMAIN_HEALTH",["HEALTHY","POSSIBLE SERVICE ISSUES","INCOMPLETE SETUP","NO SERVICES SELECTED","NO SERVICE SELECTED","NOT MANAGED","NO ACCESS","NO ADMIN ACCESS","NOT CONNECTED","NOT APPLICABLE","N/A","END CONTRACT"]);
    vocabCheck(raw.security,3,"SECURITY_DATA","Security Default",["ENABLED","DISABLED","CONDITIONAL ACCESS","NOT MANAGED","NO ACCESS","N/A"]);

    /* ---- status words in the numeric risky column ---- */
    if(raw.risky){ var sc=new Map(),tot=0; raw.risky.slice(1).forEach(function(r){ var v=r[3]; if(v==null||v==="")return; tot++; if(toNum(v)==null){ var c=clean(v); sc.set(c,(sc.get(c)||0)+1); } });
      if(sc.size){ var kinds={}; sc.forEach(function(c,k){kinds[k]=c;}); var n=0; sc.forEach(function(c){n+=c;}); out.statusInNumber.push({sheet:"RISKY_USERS_AND_DOMAIN",column:"TOTAL RISKY USERS",total:n,kinds:kinds}); } }

    /* ---- sentinel / placeholder rows ---- */
    var PH=/no data to display|no access|inactive client|not applicable|verify to see health|not enough data|not authorized|^-$/i;
    [["risky","RISKY_USERS_AND_DOMAIN"],["security","SECURITY_DATA"],["storage","STORAGE_DATA"],["usage","USAGE_REPORT"],["canva","CANVA_STATUS"],["postmaster","GOOGLE_POSTMASTERTOOLS"]].forEach(function(t){
      var rows=raw[t[0]]; if(!rows)return; var c=0; rows.slice(1).forEach(function(r){ if(r.some(function(cell){return PH.test(clean(cell));}))c++; }); if(c)out.sentinels.push({sheet:t[1],count:c,ofRows:rows.length-1});
    });

    /* ---- Postmaster reputation source present? ---- */
    if(raw.postmaster){
      var rated=0, repReal=0;
      raw.postmaster.slice(1).forEach(function(r){
        var u=clean(r[4]).toUpperCase();
        if(["HIGH","MEDIUM","LOW","BAD"].indexOf(u)>=0){ rated++; repReal++; }
        else if(/ISSUE|VERIFY TO SEE|NOT ENOUGH/.test(u)) repReal++;
      });
      // Only warn when NO reputation signal exists at all. The combined tracker
      // reports Google's live states (Issues detected / Verify to see health /
      // Not enough data) rather than HIGH/MED/LOW/BAD tiers — that IS valid data.
      if(repReal===0) out.banners.push({level:"warn",title:"Postmaster reputation source not loaded",detail:"No reputation values found in GOOGLE_POSTMASTERTOOLS. Upload the tracker (or the standalone GOOGLE POSTMASTERTOOLS.xlsx) to populate Email Reputation."});
      out.counts.pmRated = rated;
      out.counts.pmReputation = repReal;
    }

    /* ---- unattributed quarters (blank/unparseable quarter cell) ---- */
    // These rows are excluded from any specific-quarter view (they cannot be
    // attributed), so a high count means quarter-filtered totals under-report.
    var noQ = 0, noQTot = 0;
    [["storage",model.storage],["usage",model.usage],["canva",model.canva],["risky",model.risky]].forEach(function(t){
      (t[1]||[]).forEach(function(r){ noQTot++; if(!r.quarter) noQ++; });
    });
    out.counts.noQuarterRows = noQ;
    if (noQ) out.banners.push({level:"warn",title:fmtNum(noQ)+" row(s) have no quarter",detail:"Their Quarter cell is blank or unreadable, so they cannot be attributed to Q1–Q4. They are counted only while the Quarter filter is “All” and are excluded from any single-quarter view — fix the source cells so these tenants appear in quarterly totals."});

    /* ---- unrecognised Postmaster reputation values ---- */
    // normReputation preserves any value it does not recognise; those bypass every
    // colour / severity map downstream, so list them as vocabulary violations.
    var repUnk = QBR._repUnknown || {};
    Object.keys(repUnk).forEach(function(v){
      out.vocab.push({sheet:"GOOGLE_POSTMASTERTOOLS",column:"Reputation",value:v,count:repUnk[v]});
    });

    /* ---- organization gaps (schools with no org after the join) ---- */
    var orgGap=[]; if(model.master){ model.master.forEach(function(ms,k){ var org=(model.orgByKey&&model.orgByKey[k])||null; if(!org) orgGap.push(ms.name); }); }
    out.orgGaps = { count: orgGap.length, list: orgGap.sort() };

    /* ---- name variants (same key, >1 raw spelling in a sheet) ---- */
    Object.keys(sets).forEach(function(k){ sets[k].forEach(function(v){ if(v.size>1) out.nameVariants.push({sheet:SH[k].label, variants:Array.from(v)}); }); });

    /* ---- source name encoding (mojibake: UTF-8 read as Latin-1, e.g. "Ã±") ---- */
    // The loader repairs these for display, but the SOURCE cells are still wrong —
    // flag them so the tracker itself can be cleaned.
    var MOJI=/[ÃÂ][-¿]/, mojiSet={};
    ["risky","security","storage","usage","canva","postmaster"].forEach(function(k){
      var rows=raw[k]; if(!rows)return;
      rows.slice(1).forEach(function(r){ r.forEach(function(cell){ if(typeof cell==="string" && MOJI.test(cell)) mojiSet[cell.trim()]=1; }); });
    });
    out.nameEncoding = Object.keys(mojiSet);
    out.counts.nameEncoding = out.nameEncoding.length;

    /* ---- score ---- */
    var deductions = 0;
    deductions += Math.min(30, out.badCells.filter(function(b){return b.type!=="Service check"&&!/cosmetic/i.test(b.type);}).length * 2);
    deductions += out.banners.length * 8;
    deductions += Math.min(10, out.statusInNumber.reduce(function(a,s){return a+(s.total>0?4:0);},0));
    deductions += Math.min(10, Math.round(out.orgGaps.count/5));
    deductions += Math.min(8, out.vocab.length);
    var value = Math.max(0, 100 - deductions);
    // Bands describe SOURCE hygiene (the dashboard already corrects these at read time).
    var grade = value>=85?"A":value>=72?"B":value>=58?"C":value>=42?"D":"E";
    var label = value>=85?"Clean":value>=72?"Good":value>=58?"Fair — cleanup advised":value>=42?"Needs cleanup":"Poor";
    out.score = { value: value, grade: grade, label: label };

    out.counts.badCells = out.badCells.length;
    out.counts.vocab = out.vocab.length;
    out.counts.orgGaps = out.orgGaps.count;
    out.counts.sentinelSheets = out.sentinels.length;
    return out;
  };

  // Build rows for the downloadable "Source Fix List" from an audit result.
  QBR.auditFixRows = function (audit) {
    var rows = [["Sheet","Cell","Tenant","Quarter","Column","Current","Suggested fix","Type"]];
    audit.badCells.forEach(function(b){ rows.push([b.sheet,b.cell,b.tenant,b.quarter,b.column,b.current,b.fix,b.type]); });
    return rows;
  };

  if (typeof module !== "undefined" && module.exports) module.exports = QBR;
})();
