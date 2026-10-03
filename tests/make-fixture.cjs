#!/usr/bin/env node
/* =============================================================================
 * make-fixture.cjs — generate tests/fixture.xlsx, a SYNTHETIC tenant tracker.
 *
 * Purpose: let anyone run the regression harness (run-tests.cjs) with ZERO
 * sensitive data. The real "ALL TENANT AUTOMATED TRACKER.xlsx" is confidential;
 * this fabricates a tiny workbook with the same sheet/column shapes and, on
 * purpose, every dirty-data edge case the engine defends against:
 *   - a school name that arrives mojibake'd ("Ã±") and must repair to "ñ"
 *   - risky counts that are null ("-") vs 0 vs positive (must stay distinct)
 *   - a corrupt storage cell (Exchange in TB that should be GB) the median must out-vote
 *   - a GB-denominated "current storage" that must NOT inflate ~1000x
 *   - live Postmaster states ("Issues detected" / "Verify to see health") not tiers
 *   - a Verified/Unverified Status column that synthesizes verification records
 * The tenants are fictional. "Fixture Control Tenant" is the marker the harness
 * uses to select the fixture snapshot profile.
 *
 * USAGE:  node tests/make-fixture.cjs        (writes ./tests/fixture.xlsx)
 * ===========================================================================*/
const fs = require("fs"), path = require("path");
function findFile(cands) { for (const c of cands) { try { fs.accessSync(c); return c; } catch (e) {} } return null; }
const base = __dirname;
const XLSXP = findFile([base + "/../qbr-app/libs/xlsx.full.min.js", base + "/../app/libs/xlsx.full.min.js", base + "/../libs/xlsx.full.min.js"]);
if (!XLSXP) { console.error("Could not locate vendored xlsx.full.min.js next to this script."); process.exit(2); }
const XLSX = require(path.resolve(XLSXP));

// A school name mis-decoded as Latin-1 (the classic "Ã±" for "ñ"). fixMojibake
// in the loader must turn this back into "…- Dasmariñas".
const MOJI = "St. Vincent University - DasmariÃ±as";

const risky = [
  ["Month", "School_Rakso", "Organization", "Total Risky Users", "Domain_Health", "Error Cause"],
  ["JANUARY",  "Alpha Academy",        "ALPHAORG", 0,   "Healthy",                  ""],
  ["FEBRUARY", "Alpha Academy",        "ALPHAORG", 5,   "Healhty",                  ""],   // typo -> Healthy; positive
  ["MARCH",    "Alpha Academy",        "ALPHAORG", "-", "Healthy",                  ""],   // null (not 0)
  ["APRIL",    "Beta College Inc.",    "ALPHAORG", 60,  "Possible Service Issues",  "MFA gaps"],  // HIGH (>50)
  ["MAY",      "Beta College Inc.",    "ALPHAORG", 120, "Incomplete Setup",         "many"],      // CRITICAL (>100)
  ["JANUARY",  "Gamma High School",    "BETAORG",  3,   "Healthy",                  ""],
  ["FEBRUARY", "Delta Institute",      "BETAORG",  12,  "No Access",                ""],   // -> Not Managed
  ["JANUARY",  "Epsilon School",       "",         8,   "Not Connected",            ""],   // orgless
  ["JANUARY",  "Zeta Learning Center", "",         0,   "No Services Selected",     ""],
  ["JANUARY",  MOJI,                   "ALPHAORG", 2,   "Healthy",                  ""],
  ["JANUARY",  "Fixture Control Tenant","FIXORG",  1,   "Healthy",                  ""],
];

const security = [
  ["School", "GDAP", "Domain Status", "Security Default", "Authentication Methods", "MFA", "Email OTP", "SMS", "SSPR"],
  // "Authentication Methods" is the dirty free-text list the tokenizer must handle
  // (mixed delimiters/spacing, "FID02" typo, "NO ACCESS" sentinel).
  ["Alpha Academy",        "guid-1",        "Healthy",                 "Enabled",             "Passkey (FIDO2), MFA, Email OTP",                       "Yes", "Yes", "No",  "Yes"],
  ["Beta College Inc.",    "None",          "Possible Service Issues", "Disabled",            "Email OTP",                                            "No",  "Yes", "No",  "No"],
  ["Gamma High School",    "guid-2;guid-3", "Healthy",                 "Conditional Access",  "MFA, Temporary Access Pass, Software OATH Tokens, Email OTP", "Yes", "Yes", "No", "No"],
  ["Delta Institute",      "None",          "No Access",               "No Access",           "NO ACCESS",                                            "-",   "-",   "-",   "-"],  // secdef -> NOT MANAGED
  ["Epsilon School",       "None",          "",                        "",                    "SMS, FID02, Email OTP",                                "",    "Yes", "Yes", ""],   // secdef -> UNKNOWN; FID02 typo -> FIDO2
  ["Zeta Learning Center", "None",          "Healthy",                 "Enabled",             "MFA,SMS,Email OTP",                                    "Yes", "Yes", "Yes", "No"],
  [MOJI,                   "guid-4",        "Healthy",                 "Enabled",             "MFA, SMS, Voice call, Email OTP",                      "Yes", "Yes", "Yes", "Yes"],
  ["Fixture Control Tenant","guid-5",       "Healthy",                 "Enabled",             "Email OTP",                                            "Yes", "No",  "No",  "Yes"],
];

// Storage: binary-GB math. Used Storage(GB)/Total Storage(GB) are DECIMAL GB
// (loader multiplies by 1024/1000). Per-service + "current storage"/"usage" text
// carry their own TB/GB unit.
const storage = [
  ["Quarter", "School", "Organization", "OneDrive", "Exchange", "SharePoint", "Current Storage", "Used Storage(GB)", "Total Storage(GB)", "Percentage%", "Usage"],
  ["Q1", "Alpha Academy",     "ALPHAORG", "1.5 TB", "0.5 TB",   "0.2 TB", "2.2 TB",    2200, 10000, 22, "2.2 TB of 10 TB used"],
  // Corrupt Exchange cell (168.83 TB should be GB): the 4-estimate median must out-vote it.
  ["Q1", "Beta College Inc.", "ALPHAORG", "1.0 TB", "168.83 TB","0.5 TB", "3.9 TB",    3808, 10000, 39, "3.9 TB of 100 TB used"],
  ["Q1", "Gamma High School", "BETAORG",  "3 TB",   "1 TB",     "1 TB",   "5 TB",      5000, 20000, 25, "5 TB of 20 TB used"],
  // GB-denominated "current storage" must NOT inflate ~1000x.
  ["Q2", "Delta Institute",   "BETAORG",  "",       "",         "",       "440.62 GB", "",   "",    "",  "440.62 GB of 1 TB used"],
  ["Q1", MOJI,                "ALPHAORG", "0.8 TB", "0.3 TB",   "0.1 TB", "1.2 TB",    1200, 5000,  24, "1.2 TB of 5 TB used"],
  ["Q1", "Fixture Control Tenant","FIXORG","0.5 TB","0.2 TB",   "0.1 TB", "0.8 TB",    800,  4000,  20, "0.8 TB of 4 TB used"],
];

const usage = [
  ["Quarter", "Client", "Organization", "Exchange Active", "OneDrive Active", "SharePoint Active", "Teams Active", "Office 365 Active", "Total of Office 365", "% of Usage Report"],
  ["Q1", "Alpha Academy",        "ALPHAORG", 90, 85, 80, 88, 95, 100, "85%"],
  ["Q1", "Beta College Inc.",    "ALPHAORG", 20, 25, 15, 30, 28, 100, "30%"],   // low adoption
  ["Q1", "Gamma High School",    "BETAORG",  50, 55, 45, 60, 58, 100, "58%"],
  ["Q2", "Delta Institute",      "BETAORG",  40, 42, 38, 45, 44, 100, "44%"],
  ["Q1", "Epsilon School",       "",         10, 12, 8,  15, 13, 100, "13%"],   // orgless, low
  ["Q1", MOJI,                   "ALPHAORG", 70, 72, 68, 75, 74, 100, "74%"],
  ["Q1", "Fixture Control Tenant","FIXORG",  60, 62, 58, 65, 63, 100, "63%"],
];

const canva = [
  ["Quarter", "School", "Certificate Status", "Certificate Update", "Total User"],
  ["Q1", "Alpha Academy",      "Certified",     "2026-12-31", 500],
  ["Q2", "Alpha Academy",      "Certified",     "2026-12-31", 600],
  ["Q1", "Beta College Inc.",  "Not Certified", "",           0],     // 0 users -> not "active"
  ["Q2", "Gamma High School",  "Certified",     "",           1200],
  ["Q1", MOJI,                 "Certified",     "",           300],
  ["Q2", "Fixture Control Tenant", "",          "",           50],
];

const postmaster = [
  ["Month", "School", "Organization", "Domain", "Domain Reputation", "Spam Rate", "Status", "Reference"],
  ["JANUARY",  "Alpha Academy",        "ALPHAORG", "alpha.edu",   "HIGH",                 "0.1%", "Verified",   "ref"],
  ["FEBRUARY", "Alpha Academy",        "ALPHAORG", "alpha.edu",   "HIGH",                 "0.1%", "Verified",   "ref"],
  ["JANUARY",  "Beta College Inc.",    "ALPHAORG", "beta.edu",    "Issues detected",      "5%",   "Unverified", "ref"],
  ["JANUARY",  "Gamma High School",    "BETAORG",  "gamma.edu",   "Verify to see health", "-",    "Verified",   ""],
  ["JANUARY",  "Delta Institute",      "BETAORG",  "delta.edu",   "MEDIUM",               "1%",   "Verified",   ""],
  ["JANUARY",  "Epsilon School",       "",         "epsilon.edu", "Not enough data",      "-",    "Unverified", ""],
  ["JANUARY",  MOJI,                   "ALPHAORG", "sv.edu",      "Issues detected",      "3%",   "Verified",   ""],
  ["JANUARY",  "Fixture Control Tenant","FIXORG",  "fix.edu",     "HIGH",                 "0%",   "Verified",   ""],
];

// PRODUCT_AND_SERVICES is grabbed raw (org/account portfolio); not parsed into a
// per-school array today, but included so the schema-contract check can see it.
const product = [
  ["School", "Organization", "Products", "Contact", "Customer Type"],
  ["Alpha Academy",     "ALPHAORG", "Microsoft;Canva", "a@example.edu", "Education"],
  ["Beta College Inc.", "ALPHAORG", "Microsoft",       "b@example.edu", "Education"],
];

// Domain registration — per-quarter export shape (sheet named like "Q3", header
// "Registration"). Registration holds a timestamp OR a sentinel (Error/PENDING);
// Remaining Days encodes direction; Remarks carry the operational action. Uses
// the existing fixture schools so no other snapshot shifts. Exercises the full
// status vocabulary through QBR.domainStatus.
const domainreg = [
  ["Schools", "DOMAIN", "Registration", "ValidPD", "Validy in Year", "Expiration", "Remaining Days", "Status", "Remarks", "Conditional formatting"],
  ["Alpha Academy",             "alpha.edu.ph",   "2026-07-25-09-37-38", "VALIDPD: 01", "01", "2027-07-25-09-37-38", "304 days remaining",       "Current",     "",                                                              304],
  ["Beta College Inc.",         "beta.edu.ph",    "2025-09-30-10-00-00", "VALIDPD: 01", "01", "2026-11-19-10-00-00", "55 days remaining",        "Current",     "",                                                              55],
  ["Gamma High School",         "gamma.edu.ph",   "2025-10-15-10-00-00", "VALIDPD: 01", "01", "2026-10-15-10-00-00", "20 days remaining",        "Current",     "",                                                              20],
  ["Delta Institute",           "delta.edu.ph",   "2025-10-10-09-41-16", "VALIDPD: 01", "01", "2026-10-10-09-41-16", "16 days remaining",        "For Renewal", "Registration status: FOR RENEWAL. Please pay the renewal fee.",  16],
  ["Epsilon School",            "epsilon.edu.ph", "2025-09-13-13-49-14", "VALIDPD: 01", "01", "2026-09-13-13-49-14", "11 days since expiration", "Expired",     "Registration status: FOR DELETION. Please pay the renewal fee.",-11],
  ["Zeta Learning Center",      "zeta.edu.ph",    "2024-09-09-19-09-43", "VALIDPD: 01", "01", "2025-09-09-19-09-43", "380 days since expiration","Expired",     "The subscription has expired. It has already been deleted from the EDU.PH database.", -380],
  [MOJI,                        "sv.edu.ph",      "Error",               "",            "",   "",                    "",                         "",            "",                                                              null],
  ["Fixture Control Tenant",    "fix.edu.ph",     "PENDING",             "",            "",   "",                    "",                         "",            "",                                                              null],
];

const wb = XLSX.utils.book_new();
const add = (name, aoa) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
add("Lists", [["Quarters", "Months"], ["Q1", "JANUARY"]]);   // decorative, ignored by parsers
add("RISKY_USERS_AND_DOMAIN", risky);
add("SECURITY_DATA", security);
add("STORAGE_DATA", storage);
add("USAGE_REPORT", usage);
add("CANVA_STATUS", canva);
add("GOOGLE_POSTMASTERTOOLS", postmaster);
add("PRODUCT_AND_SERVICES", product);
add("Q3", domainreg);   // per-quarter domain-registration export (detected by header shape)

const out = path.join(base, "fixture.xlsx");
// The vendored SheetJS is the browser UMD build; XLSX.writeFile (which needs a
// Node fs binding) is unreliable here, so serialize to a buffer and write it.
const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
fs.writeFileSync(out, buf);
console.log("wrote " + out + " (" + fs.statSync(out).size + " bytes)");
