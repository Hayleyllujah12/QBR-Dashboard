// Synthetic RISKY_USERS_AND_DOMAIN-style audit workbook (fake schools, contoso links). Step 1 of 2;
// tests/make-rich-fixture.py then styles it into SAMPLE_AUDIT_RICH.xlsx.
const X=require(require("path").join(__dirname, "..", require("fs").existsSync(require("path").join(__dirname,"../qbr-app")) ? "qbr-app" : "", "libs/xlsx.full.min.js"));
const M=["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
const wb=X.utils.book_new();
const schools=[["Alpha Test School","ORG-A"],["Beta Test Academy","ORG-A"],["Gamma Test College","ORG-B"]];
M.forEach((m,i)=>{const ws=X.utils.aoa_to_sheet([["SCHOOL","ORGANIZATION","TOTAL RISKY USERS","DOMAIN HEALTH","REFERENCES"],...schools.map((s,j)=>[s[0],s[1],i<9?(j+1)*i:null,i<9?"Healthy":null,i<9?"Open":null])]);
 if(i<9){ws["E2"].l={Target:"https://contoso.sharepoint.com/:x:/r/sites/Test/Shared%20Documents/a.xlsx?d=w1&csf=1"};ws["E3"].l={Target:"../../../../:x:/r/sites/Test/Shared%20Documents/b.xlsx?d=w2&csf=1&web=1"};}
 X.utils.book_append_sheet(wb,ws,m);});
X.utils.book_append_sheet(wb,X.utils.aoa_to_sheet([["DOMAIN HEALTH"],["Healthy"],["No Access"]]),"Drop-Down");
try{require("fs").writeFileSync(process.argv[2] || require("path").join(__dirname,"_audit_raw.xlsx"),X.write(wb,{bookType:"xlsx",type:"buffer"}));}catch(e){console.log("ERR",e.message);process.exit(1)}console.log("ok");
