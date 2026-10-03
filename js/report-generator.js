/* =============================================================================
 * report-generator.js
 * Rule-based (offline) executive narrative generator. No LLM — deterministic
 * thresholds turn the parsed model into QBR-ready prose + recommendations.
 * ===========================================================================*/

var QBR = (window.QBR = window.QBR || {});

QBR.THRESH = { HIGH_RISK: 50, CRITICAL: 100, USAGE_LOW: 40 };

// Shared 5-band risk classification for risky-user counts — used app-wide so
// the Risky tab, Overview and Executive Report all read consistently.
//   0 → No Risk · 1-15 Low · 16-50 Medium · 51-100 High · >100 Critical
QBR.riskBand = function (v) {
  v = v || 0;
  if (v > 100) return { key: "critical", label: "Critical Risk", dot: "🔴", action: "Escalate + email POC immediately" };
  if (v > 50)  return { key: "high",     label: "High Risk",     dot: "🟠", action: "Email POC — remediate this week" };
  if (v > 15)  return { key: "medium",   label: "Medium Risk",   dot: "🟡", action: "Notify POC — monitor" };
  if (v > 0)   return { key: "low",      label: "Low Risk",      dot: "🔵", action: "Monitor" };
  return         { key: "none",     label: "No Risk / Secure", dot: "🟢", action: "None — secured" };
};
QBR.riskBadge = function (v) { const b = QBR.riskBand(v); return `<span class="risk-badge risk-${b.key}">${b.dot} ${b.label}</span>`; };

function pct(n, d) { return d ? Math.round((n / d) * 100) : 0; }
function tb(gb) { return gb == null ? "—" : (gb / 1024).toFixed(1) + " TB"; }

// agg is the object produced by app.js computeAggregates(model, filters)
QBR.generateReport = function (model, agg) {
  const t = QBR.THRESH;
  const s = [];       // narrative sections {title, body}
  const recs = [];    // {level, text}

  /* ---- Executive summary ---- */
  const healthy = agg.health["Healthy"] || 0;
  const totalDom = Object.values(agg.health).reduce((a, b) => a + b, 0);
  const noAccess = agg.health["Not Managed"] || 0;
  const disabled = agg.secDefault["DISABLED"] || 0;
  s.push({ title: "Executive Summary", body:
    `Across ${agg.totalSchools} managed tenants, ${healthy} of ${totalDom} schools reporting domain health are healthy (${pct(healthy, totalDom)}%). ` +
    `Security posture shows ${agg.secDefault["ENABLED"] || 0} tenants with Security Defaults enabled and ${disabled} with them not enabled. ` +
    `Microsoft 365 adoption averages ${agg.usageAvg == null ? "—" : agg.usageAvg.toFixed(1) + "%"}, ` +
    `total storage under management is ${tb(agg.storageTotalGB)}, and Canva Education reaches ${(agg.canvaUsers || 0).toLocaleString()} users across ${agg.canvaActiveSchools} schools. ` +
    `${agg.topRisk ? `Identity risk is concentrated at ${agg.topRisk.schoolRaw} (${agg.topRisk.risky} risky users).` : ""}`
  });

  /* ---- Security ---- */
  s.push({ title: "Security Overview", body:
    `${agg.secDefault["ENABLED"] || 0} schools operate with Security Defaults enabled and ${agg.secDefault["CONDITIONAL ACCESS"] || 0} use Conditional Access policies. ` +
    `${disabled} tenants run with Security Defaults not enabled and ${agg.secDefault["NOT MANAGED"] || 0} report no administrative access. ` +
    `${agg.mfaEnabled} tenants have MFA enabled and ${agg.ssprEnabled} have SSPR configured.`
  });

  /* ---- Risk ---- */
  const crit = agg.riskySchools.filter(r => r.risky > t.CRITICAL);
  const high = agg.riskySchools.filter(r => r.risky > t.HIGH_RISK && r.risky <= t.CRITICAL);
  s.push({ title: "Risky Sign-in Analysis", body:
    `${crit.length} school(s) exceed the CRITICAL threshold (>${t.CRITICAL} risky users)` +
    `${crit.length ? ": " + crit.slice(0,3).map(r => `${r.schoolRaw} (${r.risky})`).join(", ") : ""}. ` +
    `${high.length} school(s) fall in the HIGH RISK band (>${t.HIGH_RISK}). ` +
    `${agg.topRisk ? `${agg.topRisk.schoolRaw} represents the highest concentration of identity risk in the selected period.` : ""}`
  });

  /* ---- Usage ---- */
  s.push({ title: "Microsoft 365 Adoption", body:
    `Average utilization is ${agg.usageAvg == null ? "—" : agg.usageAvg.toFixed(1) + "%"}` +
    `${agg.usageAvg != null && agg.usageAvg < t.USAGE_LOW ? ", below the 40% value-realization benchmark, indicating licensing optimization opportunities" : ""}. ` +
    `${agg.topUsage ? `${agg.topUsage.schoolRaw} leads adoption at ${agg.topUsage.usagePct.toFixed(1)}%` : ""}` +
    `${agg.lowUsage ? ` while ${agg.lowUsage.schoolRaw} trails at ${agg.lowUsage.usagePct.toFixed(1)}%` : ""}.`
  });

  /* ---- Storage ---- */
  s.push({ title: "Storage Analysis", body:
    `Total storage consumed is ${tb(agg.storageTotalGB)} at an average utilization of ${agg.storageAvgPct == null ? "—" : agg.storageAvgPct.toFixed(1) + "%"}. ` +
    `${agg.topStorage ? `${agg.topStorage.schoolRaw} is the largest consumer at ${tb(agg.topStorage.usedGB)}.` : ""}`
  });

  /* ---- Canva ---- */
  s.push({ title: "Canva Adoption Analysis", body:
    `${(agg.canvaUsers || 0).toLocaleString()} Canva users are active across ${agg.canvaActiveSchools} partner schools. ` +
    `${agg.topCanva ? `${agg.topCanva.schoolRaw} hosts the largest community (${(agg.topCanva.users || 0).toLocaleString()} users).` : ""}`
  });

  /* ---- User Management readiness ---- */
  if (agg.umTotal != null) {
    s.push({ title: "User Management & School-Year Readiness", body:
      `${agg.umUpdated} of ${agg.umTotal} schools completed grade-level and user-information updates for ${agg.umSY} ` +
      `(${agg.umCompletion.toFixed(1)}% readiness). ` +
      `${agg.umCompletion < 50 ? "The majority remain pending and should be prioritized before the next deployment cycle." : "Rollover progress is on track across the managed portfolio."}`
    });
  }

  /* ---- Email reputation (Postmaster) ---- */
  if (agg.pmReported != null) {
    const flagged = (agg.pmIssues || 0) + (agg.pmBad || 0);   // Issues detected + BAD
    s.push({ title: "Email Reputation (Google Postmaster)", body:
      `Across ${agg.pmReported} school domain(s) reporting a reputation status, ${flagged} show deliverability problems ` +
      `(Google "Issues detected"${agg.pmBad ? " / BAD" : ""})` +
      `${flagged ? " — e.g. " + agg.pmBadSchools.slice(0, 3).join(", ") + (agg.pmBadSchools.length > 3 ? ", …" : "") : ""}. ` +
      `Flagged domains risk mail deliverability and should review SPF/DKIM/DMARC and outbound sending practices; ` +
      `domains reading "Verify to see health" must first be verified in Postmaster before Google will report their reputation.`
    });
  }

  /* ---- Recommendations (priority-ranked) ---- */
  const pmFlagged = (agg.pmIssues || 0) + (agg.pmBad || 0);
  if (pmFlagged) recs.push({ level: "High", text: `Remediate email authentication for ${pmFlagged} domain(s) with deliverability issues.` });
  if (agg.umTotal != null && agg.umCompletion < 100) recs.push({ level: "Medium", text: `Complete user-management rollover for ${agg.umTotal - agg.umUpdated} pending school(s) in ${agg.umSY}.` });
  if (disabled > 0) recs.push({ level: "Critical", text: `Enable Security Defaults for ${disabled} tenant(s) where they are not enabled.` });
  if (crit.length) recs.push({ level: "Critical", text: `Investigate ${crit.length} CRITICAL-risk school(s) exceeding ${t.CRITICAL} risky sign-ins.` });
  if (agg.secDefault["NOT MANAGED"]) recs.push({ level: "High", text: `Restore administrative access to ${agg.secDefault["NOT MANAGED"]} Not-Managed tenant(s).` });
  if (high.length) recs.push({ level: "High", text: `Remediate ${high.length} HIGH-risk school(s) (>${t.HIGH_RISK} risky users).` });
  if (agg.ssprGap > 0) recs.push({ level: "Medium", text: `Deploy SSPR to ${agg.ssprGap} tenant(s) without self-service password reset.` });
  if (agg.usageAvg != null && agg.usageAvg < t.USAGE_LOW) recs.push({ level: "Medium", text: `Launch Microsoft 365 adoption enablement to lift utilization above 40%.` });
  if (agg.topStorage) recs.push({ level: "Low", text: `Review storage lifecycle policies for top consumers (e.g. ${agg.topStorage.schoolRaw}).` });
  if (!recs.length) recs.push({ level: "Low", text: "No material risks detected in the selected scope." });

  return { sections: s, recommendations: recs };
};

// Per-chart slide-ready content: Headline / Key Finding / Business Impact / Recommendation
QBR.slideNote = function (headline, finding, impact, recommendation) {
  return { headline, finding, impact, recommendation };
};

if (typeof module !== "undefined" && module.exports) module.exports = QBR;
