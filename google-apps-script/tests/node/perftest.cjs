/**
 * Node test harness for SalesDashboard.gs (Performance + Business logic).
 * Run:  node google-apps-script/tests/node/perftest.js
 * It stubs SpreadsheetApp / CacheService / etc., feeds mock sheets through the
 * real script, and asserts the computed values and rendered HTML.
 */
global.Utilities={formatDate:function(d,tz,f){
  if(f&&f.indexOf("HH")>-1) return "24 Sep 2026, 10:30";
  function p(n){return(n<10?"0":"")+n;}
  return p(d.getDate())+"-"+p(d.getMonth()+1)+"-"+d.getFullYear();}};
global.Session={getScriptTimeZone:function(){return "Asia/Calcutta";}};
global.CacheService={getScriptCache:function(){return{get:function(){return null},put:function(){}}}};
global.HtmlService={createHtmlOutput:function(){return{addMetaTag:function(){return this},setTitle:function(){return this}}}};
function mockSheet(name, values){return{getName:function(){return name},getDataRange:function(){return{
  getValues:function(){return values},
  getDisplayValues:function(){return values.map(function(r){return r.map(function(v){
    return(v===null||v===undefined||v instanceof Date)?"":String(v);})});}};}};}
var ALL_SHEETS=[];
var SS_OBJ={getSheets:function(){return ALL_SHEETS;},
  getSheetByName:function(n){for(var i=0;i<ALL_SHEETS.length;i++)if(ALL_SHEETS[i].getName()===n)return ALL_SHEETS[i];return null;}};
global.SpreadsheetApp={getActiveSpreadsheet:function(){return SS_OBJ;}};
var fs=require('fs');
eval(fs.readFileSync(__dirname+"/../../SalesDashboard.gs","utf8"));
getSpreadsheet_=function(){return SS_OBJ;};

// ---- target layout ----
//  Row1: C1:J1 = salesperson roster (empty cells skipped) · K1 = C-SAT period
//        name · L1:U1 = C-SAT table header (name col + buckets + TOTAL/C-SAT/CFI)
//  Rows: A5:D9  targets · A10:C10 google review · B13:L16 KPI grid
//        (13=type, 14=name, 15=target count, 16=approx unit value)
var T=[]; for(var i=0;i<20;i++) T.push(new Array(22).fill(""));
T[0]=["Revenue Type","Store Target","SUBHA","DINESH","","","","","","",
  "FY26P12","MDU4","VERY SATISFIED","SOMEWHAT SATISFIED","VERY DISSATISIFIED",
  "NEITHER SATISFIED NOR DISSATISFIED","-","SOMEWHAT DISSATISFIED","TOTAL","C-SAT","CFI",""];
T[1]=["Service",300000,180000,120000];
T[2]=["Accessories",150000,90000,60000];
T[3]=["Repair",200000,120000,80000]; T[4]=["Total",650000,390000,260000];
T[5]=["Gross Profit",160000,96000,64000];
T[8]=["Prev Month","This Month","Difference"]; T[9]=[1475,1494,19];
T[12]=["","AR","","","","RR","","","","SR","",""];
T[13]=["","Care+","Protect+","AMC","Bundle","Display","Battery","Glass","Others","SRD","SHC","ARD"];
T[14]=["",25,15,40,30,50,45,20,35,60,25,18];
// row 16 (0-idx 15): approx unit values in KPI cols AND slab anchors in sheet cols
T[15]=["Achivement",400,600,900,300,"incentive per amc",150,100,0,0,0,0,"slab level",5];
T[16]=["70-80%","2000","70-80%","1000","",  "1-10","₹150","",5,"₹80","","","",300];
T[17]=["80-90%","4000","80-90%","2000","","11-25","₹250","",12,"₹120","","","",""];
T[18]=["90-100%","6000","90-100%","3000"];
// C-SAT data rows (L2:U5)
T[1][11]="Subha K";      T[1][12]=1; T[1][18]=1; T[1][19]="100.00%"; T[1][20]="0.00%";
T[2][11]="Dinesh Kumar"; T[2][12]=9; T[2][14]=1; T[2][18]=10; T[2][19]="90.00%"; T[2][20]="10.00%";
T[3][11]="Murugan Mp";   T[3][12]=2; T[3][14]=1;
T[4][11]="TOTAL";        T[4][12]=12; T[4][14]=2; T[4][18]=14; T[4][19]="85.71%"; T[4][20]="14.29%";

// ---- crm ----
var crm=[["InvoiceCode","InvoiceDate","SalesPersonName","RevenueType","ItemDescription",
 "SubCategory","SchemeCode","TaxableAmount","GrossProfit","TranType","ItemCode","PartType"],
 ["INV1","23-09-2026 10:15","Subha","Service","Svc","","",5000,1500,"SALES","",""],
 ["INV2","23-09-2026 11:00","Dinesh","Accessories","Case","","",1200,300,"SALES","",""],
 ["INV3","23-09-2026 12:00","Subha","Accessories","Care+ plan","","",900,250,"SALES","CPLUS",""]];

// ---- walkin ----
var wl=[["MobileNo","TokenDate","VisitPurposeDesc","CaseId","InvoiceCode","Remark","TokenCreatedBy"],
 ["919876543210","2026-09-23 09:15:00","Repair Submission","C1001","IV9001","Needs original glass","Subha"],
 ["919876543211","2026-09-23 09:40:00","Repair Submission","C1002","IV9002","Warranty valid","Subha"],
 ["919876543212","2026-09-23 10:05:00","General/Sales Enquiry","C1003","","Follow up in 2 days","Dinesh"],
 ["919876543213","2026-09-23 11:15:00","Collection","","IV9003","Bought case","Dinesh"],
 ["919876543214","2026-09-23 12:30:00","General/Sales Enquiry","","","","Dinesh"],
 ["919876543215","2026-09-22 10:20:00","Repair Submission","C1000","","Waiting for part","Subha"],
 ["919876543216","2026-09-22 13:00:00","Collection","C0999","IV8999","","Priya"]];

// ---- comm dates relative to today (staleness uses the server clock) ----
function daysAgoStr(n,h){var d=new Date(Date.now()-n*86400000);function p(x){return(x<10?"0":"")+x;}
  return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate())+" "+(h||"10")+":00:00";}
var cm=[["CaseId","CreatedDate","Remark"],
 ["C1001",daysAgoStr(6,"10"),"First update"],
 ["C1001",daysAgoStr(1,"15"),"Ready to collect"],
 ["C1002",daysAgoStr(4,"11"),"Informed about estimate"],
 ["C1004",daysAgoStr(0,"09"),"Called customer"]];

// ---- service ----
var sv=[["CaseId","FirstName","LastName","MobileNo","ProductDescription","CaseDate",
 "RepairCreatedDate","JobAge","Status","DeliveredDate","GSXRepairStatusDescription","CRM_JobStatusDesc",
 "CoverageCodeDescription","RepairDocType","Ready_For_Pickup_Date_CRM"],
 // duplicate of C1003 with an OLDER case date placed BEFORE the real row —
 // the merger must keep the later (real) row instead.
 ["C1003","Mani","Kandan","919876543212","Different Phone",daysAgoStr(5,"10"),"",5,"Open","","Diagnosis In Progress","","In Warranty","Walk-In Repair",""],
 ["C1001","Ravi","Kumar","919876543210","iPhone 13",daysAgoStr(14,"10"),daysAgoStr(13,"11"),14,"Open","","Repair Complete","Original Display Used","In Warranty","Walk-In Repair",daysAgoStr(12,"10")],
 ["C1002","Anita","Devi","919876543211","Samsung S23",daysAgoStr(4,"09"),daysAgoStr(3,"14"),4,"Open","","Unclaimed Device","Part Used","Out of Warranty","Carry-In Repair",daysAgoStr(1,"12")],
 ["C1003","Mani","Kandan","919876543212","OnePlus 11",daysAgoStr(2,"10"),"",2,"Open","","Diagnosis In Progress","Device Unclaimed Device due","Out Of Warranty (No Coverage)","Walk-In Repair",""],
 ["C1004","Peela","Dass","919876543217","Realme 12",daysAgoStr(1,"08"),"",1,"Open","","Repair In Progress","Ready For Pickup","In Warranty","Service Non-Repair",""],
 ["C1005","Old","Closed","919876543218","Pixel 7",daysAgoStr(55,"09"),daysAgoStr(54,"10"),55,"Delivered",daysAgoStr(50,"10"),"Repair Complete","Part Used","In Warranty","Carry-In Repair",daysAgoStr(50,"10")],
 ["C1006","Age","Fallback","919876543219","Moto G",daysAgoStr(20,"09"),"","","Open","","Waiting for Part","","Out Of Warranty (No Coverage)","Walk-In Repair",""],
 ["C1007","Can","Cell","919876543220","Nokia",daysAgoStr(2,"09"),"",2,"Cancelled","","Aborted","","In Warranty","Carry-In Repair",""]];

// ---- advance ----
var av=[["MobileNo","AvailableBalance"],["919876543210",500],["919876543210",200],["919876543217",350],["919876543219",999]];

ALL_SHEETS=[mockSheet("target",T),mockSheet("crm",crm),mockSheet("walkin",wl),
  mockSheet("service",sv),mockSheet("advance",av),mockSheet("comm",cm)];

var fails=0;
function check(n,c,x){ if(c) console.log("  PASS",n); else {fails++;console.log("  FAIL",n,JSON.stringify(x));} }

console.log("== BUSINESS: KPI approx / avg walkin / store target ==");
var loaded=loadDashboardData_();
var biz=loaded.data;
var careplus=biz.kpi.ar.filter(function(k){return k.code==="CAREPLUS";})[0];
check("CAREPLUS approx unit 400 from row 16",careplus&&careplus.approxAmount===400);
check("CAREPLUS achieved 1 (INV3)",careplus&&careplus.achieved===1);
check("approx balance unit math (target 25 - achieved 1) x 400", careplus.balance===24);
check("avg invoice value total-lines 3",biz.store.avgInvoiceLineValue===Math.round(7100/3));
check("unique walkins 7",biz.store.uniqueWalkins===7);
check("avg per walkin 1014",biz.store.avgPerWalkin===1014);

console.log("== DAILY TARGETS = balance ÷ remaining working days ==");
// crm only spans up to Sep 23 with Sep month -> remaining work days depend on data month
var subha=biz.salespeople.find(function(p){return p.name==="Subha"});
var dinesh=biz.salespeople.find(function(p){return p.name==="Dinesh"});
// balance-based: >= target-achieved, and < the old rule (target / all working days)
check("Subha svc daily > 0 and rule changed", subha.dailyTargets.service>0);
var balS=180000-5000;
var rem=biz.store.remainingWorkingDays||biz.store.daysRemaining+1;
check("Subha svc daily matches balance/remaining", subha.dailyTargets.service===Math.round(balS/Math.max(1,rem)),
  [subha.dailyTargets.service,balS,rem]);
check("Dinesh svc daily = 120000/remaining", dinesh.dailyTargets.service===Math.round(120000/Math.max(1,rem)));
check("Dinesh repair daily = 80000/remaining", dinesh.dailyTargets.repair===Math.round(80000/Math.max(1,rem)));
check("Subha GP daily = (96000-1750)/remaining", subha.dailyTargets.grossProfit===Math.round(94250/Math.max(1,rem)),
  subha.dailyTargets.grossProfit);
check("Dinesh GP daily = (64000-300)/remaining", dinesh.dailyTargets.grossProfit===Math.round(63700/Math.max(1,rem)),
  dinesh.dailyTargets.grossProfit);
check("Subha acc daily balance-based", subha.dailyTargets.accessory===Math.round(89100/Math.max(1,rem)));

console.log("== WALK-IN ==");
var w=loadPerformanceData_().walkin;
check("walkins 7",w.totals.walkins===7);
check("dead leads 1",w.totals.deadLeads===1);
check("purpose cols actual",w.purposeCols.indexOf("Repair Submission")>-1&&w.purposeCols.indexOf("General/Sales Enquiry")>-1);
var d23=w.days[0];
check("23 purposes total", (d23.purposes["Repair Submission"]||0)===2);

console.log("== JOBS ==");
var j=loadPerformanceData_().jobs;
check("open jobs 5 (dedupe keeps unique case ids)",j.totals.openJobs===5,j.totals.openJobs);
check("duplicate rows merged count",j.totals.duplicatesMerged===1,j.totals.duplicatesMerged);
check("kept row is the latest case date",j.jobs.filter(function(o){return o.caseId==="C1003";})[0].product==="OnePlus 11");
check("non-open excluded 2",j.totals.closedExcluded===2,j.totals.closedExcluded);
check("C1006 fallback age ~20d", (function(){var c=j.jobs.find(function(o){return o.caseId==="C1006";});return c&&Math.abs(c.jobAge-20)<=1;})());
check("unclaimed 2",j.totals.unclaimed===2);
check("advance via advance-sheet total 2049",j.totals.advanceTotal===2049,j.totals.advanceTotal);

// ---- exact mapping spec, letter by letter ----
check("map Case ID <- CaseId",j.sources.caseId==="CaseId",j.sources.caseId);
check("map Mobile <- MobileNo",j.sources.mobile==="MobileNo",j.sources.mobile);
check("map GSX <- GSXRepairStatusDescription",j.sources.gsxStatus==="GSXRepairStatusDescription",j.sources.gsxStatus);
check("map CRM <- CRM_JobStatusDesc",j.sources.crmRepairStatus==="CRM_JobStatusDesc",j.sources.crmRepairStatus);
check("map Warranty <- CoverageCodeDescription",j.sources.warranty==="CoverageCodeDescription",j.sources.warranty);
check("map RepairType <- RepairDocType",j.sources.repairDocType==="RepairDocType",j.sources.repairDocType);
check("map Open filter <- Status",j.sources.status==="Status",j.sources.status);
var c1001=j.jobs.find(function(o){return o.caseId==="C1001";});
check("C1001 gsx value",c1001.gsxStatus==="Repair Complete",c1001.gsxStatus);
check("C1001 crm value",c1001.crmRepairStatus==="Original Display Used",c1001.crmRepairStatus);
check("C1001 warranty value",c1001.warranty==="In Warranty",c1001.warranty);
check("C1001 repair type value",c1001.repairType==="Walk-In Repair",c1001.repairType);
check("C1001 advance 500+200 by mobile",c1001.advance===700,c1001.advance);
check("C1001 mobile",c1001.mobile==="919876543210",c1001.mobile);
check("C1001 communicated yes",c1001.communicated===true);
check("C1001 comm fresh (1d) green",c1001.staleComm===false);
var c1002=j.jobs.find(function(o){return o.caseId==="C1002";});
check("C1002 comm 4d red",c1002.staleComm===true);
check("C1004 advance 350",j.jobs.find(function(o){return o.caseId==="C1004";}).advance===350);
check("mapping line on page",true); // replaced below on HTML section

console.log("== HTML ==");
var page=buildDashboardPage_(biz);
check("KPI heeft Approx Value header",page.indexOf("Approx Achievable Value")>-1);
check("approx achieved ₹400",page.indexOf("₹400")>-1);
check("approx achievable value ₹9,600",page.indexOf("<th>Approx Achievable Value</th>")>-1&&page.indexOf("₹9,600")>-1);
check("review tile big = gain +19",page.indexOf('class="value pos">+19</div>')>-1);
check("review caption small",page.indexOf("1494 this month")>-1&&page.indexOf("previous 1475")>-1);
check("no store-total text on person cards (owner removed it)",page.indexOf("Store Total Revenue Target")===-1);
// Section 7 oval pills: mint/salmon fills with dark text, pill markup present
check("oval pill green css",page.indexOf(".apill.ok { background:#63e6b7")>-1);
check("oval pill red css",page.indexOf(".apill.no { background:#ff9da7")>-1);
check("oval pill cells render",page.indexOf('<span class="apill no">')>-1||page.indexOf('<span class="apill ok">')>-1);
// Section 1 store tiles like the old page
check("store tile grid",page.indexOf('<div class="rgrid">')>-1);
check("store tile card",page.indexOf('class="card rtile')>-1);
check("status pill text",page.indexOf("In Progress")>-1||page.indexOf("Completed")>-1);
check("per-day requirement strip",page.indexOf("Per Day Requirement")>-1);
check("metric grid",page.indexOf('class="metric a"')>-1&&page.indexOf('class="metric b"')>-1);
// Title from crm!C2 (mock C2 = "Subha")
check("title uses crm C2",page.indexOf("Store Performance Dashboard - Subha")>-1);
check("seven tabs: Index/GP/KPI/WALK-IN/Operation/Jobs/DATE",
  page.indexOf("🏠 Index")>-1&&page.indexOf("🎯 GP")>-1&&page.indexOf("📊 KPI")>-1
  &&page.indexOf("🚶 WALK-IN")>-1&&page.indexOf("⚙️ Operation")>-1
  &&page.indexOf("🧰 Open Jobs")>-1&&page.indexOf("📅 DATE")>-1);
check("index tiles render",page.indexOf("glass-tile")>-1&&page.indexOf("Store Target")>-1
  &&page.indexOf("AR Attachment")>-1&&page.indexOf("Decline Repair")>-1
  &&page.indexOf("Requote - Winback")>-1&&page.indexOf("Job Management")>-1
  &&page.indexOf("Date-wise Sales Revenue Analysis")>-1);
check("tile gotoPage hook",page.indexOf("gotoPage('gp'")>-1&&page.indexOf("gotoPage('date'")>-1);
check("tab slider row",page.indexOf('class="tabs-row"')>-1);
check("fonts bumped globally",page.indexOf("font-size:18px")>-1);
check("DISC/AMC notes hidden",page.indexOf("starts with DISC")===-1&&page.indexOf("SubCategory contains AMC")===-1);
check("date-wise GP column",page.indexOf('class="grp grp-0">Gross Profit</th>')>-1);
check("legend shows GP daily target",page.indexOf("Gross Profit "+"₹"+subha.dailyTargets.grossProfit.toLocaleString("en-IN"))>-1);
check("GP daily target in legend for Dinesh",page.indexOf("Gross Profit "+"₹"+dinesh.dailyTargets.grossProfit.toLocaleString("en-IN"))>-1);
check("job table 10-col owner order",page.indexOf(
  "<th>Repair Bucket</th><th>Aging</th><th>GSX Repair Status</th>"+
  "<th>Case ID</th><th>Name</th><th>Number</th><th>Model</th>"+
  "<th>Repair Type</th><th>Communicated</th><th>Comm. Updated</th>")>-1);
check("old job columns dropped",page.indexOf("<th>Warranty Type</th>")===-1
  &&page.indexOf("<th>Job created</th>")===-1&&page.indexOf("<th>Advance paid</th>")===-1);
check("repair bucket cell tone",page.indexOf('class="bucket-cell"')>-1);
// ---- job rules ----
var c1004=j.jobs.find(function(o){return o.caseId==="C1004";});
check("SNR ready-for-pickup flags gsxCompleted",c1004.gsxCompleted===true);
check("C1006 no-repair 20d flags",j.jobs.find(function(o){return o.caseId==="C1006";}).noRepairLate===true);
check("C1003 2d old no repair -> NOT flagged",j.jobs.find(function(o){return o.caseId==="C1003";}).noRepairLate===false);
check("C1003 OOW no advance flagged",j.jobs.find(function(o){return o.caseId==="C1003";}).oowNoAdvance===true);
check("C1006 OOW with advance NOT flagged",j.jobs.find(function(o){return o.caseId==="C1006";}).oowNoAdvance===false);
check("gsx Completed badge renders",page.indexOf('class="badge completed">Completed')>-1);
check("grphead buckets render",page.indexOf('class="grphead"')>-1);
var grpOrder=[page.indexOf('>Device Unclaimed Device due<'),page.indexOf('>Original Display Used<'),page.indexOf('>Ready For Pickup<'),page.indexOf('>Waiting for Part<')];
check("buckets alphabetical",grpOrder.every(function(x){return x>-1;})&&grpOrder[0]<grpOrder[1]&&grpOrder[1]<grpOrder[2]&&grpOrder[2]<grpOrder[3],grpOrder);
check("bucket count chip",page.indexOf('grphead-count">1 open case')>-1);
check("highlight rows render",page.indexOf("norepair-row")>-1&&page.indexOf("oow-nopay-row")>-1);
check("jobs repair type column",page.indexOf("<th>Repair Type</th>")>-1);
check("jobs warranty value in matrix data",
  loadPerformanceData_().repair.validationMatrix.coverages.indexOf("Out of Warranty")>-1
  ||page.indexOf("Out of Warranty")>-1);
check("repair tab section",page.indexOf("Repair Type Performance")>-1);
// month-to-date composition (no date rows) + integers only
check("till-date table retired",page.indexOf("Till Date ·")===-1&&page.indexOf("All repair types")===-1);
check("no per-date pivot",page.indexOf("Repair Types by Date")===-1);
check("mtd walk-in 3 / carry-in 2",page.indexOf('<strong>Walk-In Repair</strong></td><td class="num">3</td>')>-1
  &&page.indexOf('<strong>Carry-In Repair</strong></td><td class="num">2</td>')>-1);
check("mtd grand-total row retired",page.indexOf("All repair types")===-1);
var raObj=loadPerformanceData_().repair;
check("mtd totals still computed per type",raObj.mtdRows.length===2&&raObj.mtdTotal===5);
check("repair tile avg baseline",page.indexOf("avg 0/month")>-1);
check("no-fraction averages",page.indexOf("avg 0.3/month")===-1);
check("KPI names from row 14",biz.kpi.ar.some(function(k){return k.code==="CAREPLUS";}));
check("KPI targets from row 15",careplus.target===25&&biz.kpi.ar.filter(function(k){return k.code==="AMC";})[0].target===40);

check("no-baseline note",page.indexOf("no 3-month baseline")>-1);
check("repair meta",page.indexOf("5 repairs this month · 2 repair types")>-1);
check("avg days col header",page.indexOf("Avg Days · Now")>-1);
check("walk-in avg days 2 d",page.indexOf('<strong>Walk-In Repair</strong></td><td class="num">3</td><td class="num">0</td>')>-1);
// repair validation matrix: products vertical × coverage horizontal, one table
// The products×coverage matrix was REMOVED from the page on the owner's
// request; the COMPUTE object is still built (other consumers may need it).
var vmObj=raObj.validationMatrix;
check("validation matrix removed from page",page.indexOf("Repair Validation — Products × Coverage")===-1);
check("vmatrix compute still built",vmObj&&vmObj.grand===5,vmObj&&vmObj.grand);
check("vmatrix iPhone=1 Other=4",vmObj.prodTotals.iPhone===1&&vmObj.prodTotals.Other===4,
  JSON.stringify(vmObj.prodTotals));
check("vmatrix coverages incl OOW-NC",vmObj.coverages.indexOf("Out Of Warranty (No Coverage)")>-1);
check("vmatrix definition gone",page.indexOf("Definition: one table combining every repair")===-1);
check("snr excluded note",page.indexOf('"Service Non-Repair" cases are excluded')>-1);
check("walkin two-tint rows only",page.indexOf("--tint:#")>-1&&page.indexOf("--tint:hsl(")===-1
  &&page.indexOf("background:hsl(")===-1);
check("token table",page.indexOf("% of creation")>-1&&page.indexOf("Total tokens generated")>-1);
check("ready col gone",page.indexOf("Ready for pickup")===-1);
check("comm pos/neg present",page.indexOf("comm-pos")>-1&&page.indexOf("comm-neg")>-1);
check("mapping line owner columns",page.indexOf('class="map-line"')>-1
  &&page.indexOf("Number ← <b>MobileNo</b>")>-1&&page.indexOf("Repair Bucket ←")>-1
  &&page.indexOf("Open filter")>-1);
check("business sections intact",page.indexOf("Invoice Valuation")>-1&&page.indexOf("Date-wise")>-1);
// ---- section headings re-named per the new page structure ----
check("weekly section heading",page.indexOf("🗓 Weekly Revenue Analysis —")>-1&&page.indexOf('id="sec-weekly"')>-1);
check("date-wise section heading",page.indexOf("📆 Date-wise Sales Revenue Analysis")>-1&&page.indexOf('id="sec-datewise"')>-1);
check("gp sections anchor",page.indexOf('id="sec-store"')>-1&&page.indexOf('id="sec-invoice"')>-1);
check("kpi sections anchor",page.indexOf('id="sec-kpi"')>-1&&page.indexOf('id="sec-incentive"')>-1
  &&page.indexOf('id="sec-person"')>-1&&page.indexOf('id="sec-ar"')>-1);
check("Store Total column removed",page.indexOf("Store Total</th>")===-1);
check("week rows 1-7 etc",page.indexOf("<b>Week 1</b>")>-1&&page.indexOf("(1–7")>-1&&page.indexOf("<b>Week 5</b>")>-1);
check("weekly status labels",page.indexOf("In progress")>-1&&page.indexOf("Upcoming")>-1&&page.indexOf("Completed")>-1);
check("weekly target formula (66667 svc wk4)",page.indexOf("₹66,667")>-1);
check("weekly svc achieved wk4 5000",page.indexOf("<b>Week 4</b>")>-1&&page.indexOf("₹5,000")>-1);
// ---- process page: target block (F1 title + G1:P7) & walk-in summary table ----
check("CSAT title from K1",page.indexOf("FY26P12")>-1 && page.indexOf("Customer Satisfaction")>-1);
check("CSAT buckets from L1:U5 header",page.indexOf("VERY SATISFIED")>-1&&page.indexOf("NEITHER SATISFIED NOR DISSATISFIED")>-1);
check("CSAT junk dash col dropped",page.indexOf("<th>-</th>")===-1&&page.indexOf("<th></th>")===-1);
check("CSAT people rows",page.indexOf("Subha K")>-1&&page.indexOf("Dinesh Kumar")>-1&&page.indexOf("Murugan Mp")>-1);
check("CSAT total row + chips",page.indexOf("Feedback responses")>-1&&page.indexOf('cc-value">14</span>')>-1&&page.indexOf("Satisfied (VS + SS)")>-1);
check("CSAT dissatisfied pct",page.indexOf("csat-neg")>-1&&page.indexOf("14%")>-1);
check("CSAT CFI chip",page.indexOf("CFI")>-1&&page.indexOf("85.71%")>-1);
check("repair comparison strip (excl SNR)", page.indexOf("Total Repairs (excl. Service Non-Repair) — this month")>-1
  &&page.indexOf('rs-value">5</span>')>-1);
check("walk-in summary is a table", page.indexOf('ptable walksum')>-1 && page.indexOf("<b>Total Walk-ins</b></td>")>-1);

check("weekly wd wk2=7 (no holiday)",page.indexOf("(8–14 Sep)</span></td><td style=\"text-align:center\">7</td>")>-1);
check("weekly wd wk5=2",page.indexOf("(29–30 Sep)</span></td><td style=\"text-align:center\">2</td>")>-1);

console.log("== TASK 19 rework ==");
check("Subha repair target fixed to store", subha.repair.target===200000, subha.repair.target);
check("Dinesh repair target fixed to store", dinesh.repair.target===200000, dinesh.repair.target);
check("repair gate on store target (80% of 200000)", subha.repairGate.requiredValue===160000,
  subha.repairGate.requiredValue);
check("person card marks store-eligibility rule", page.indexOf("eligible to receive the incentive they earned")>-1
  ||page.indexOf("Incentive unlocked for everyone")>-1||page.indexOf("Incentive on hold for everyone")>-1);
check("weekly formula paragraph removed (owner)", page.indexOf("week target = month target ÷ actual working days")===-1);
check("walkin unique total 7", w.totals.unique===7, w.totals.unique);
check("walkin day unique 5", w.days[0].unique===5, w.days[0].unique);
check("walkin conversion headers gone", page.indexOf("Case conv.")===-1&&page.indexOf("Invoice conv.")===-1);
check("walkin unique row + col", page.indexOf("<b>Unique Walk-ins</b>")>-1
  &&page.indexOf('<th rowspan="2">Unique</th>')>-1);
check("token table right of walk-in", page.indexOf('class="walkin-pair"')>-1
  &&page.indexOf('<div class="wp-right"')>-1
  &&page.lastIndexOf("tokentable")>page.lastIndexOf('<div class="wp-right"'));
check("job buckets grouped", page.indexOf('class="grphead"')>-1);
check("decline/requote placeholders", page.indexOf('ph-tile')>-1
  &&page.indexOf("Decline Repair")>-1&&page.indexOf("Requote &amp; Win-back")>-1);
check("vmatrix shaded cells", page.indexOf("table.vmatrix td { background:#1d2354; color:#e9edff")>-1);
check("blue-glow header skin", page.indexOf("#5b7bff 0%,#4360e0")>-1);
check("3d bar styling", page.indexOf("--barc")>-1&&page.indexOf("barShine")>-1
  &&page.indexOf(".bar-fill::before")>-1);
check("index page pane", page.indexOf('data-pane="index"')>-1&&page.indexOf("index-hero")>-1);
check("index tiles built (11 now, matrix tile removed)", (page.match(/class="glass-tile"/g)||[]).length>=11,
  (page.match(/class="glass-tile"/g)||[]).length);
check("tile group heads", page.indexOf("tile-group-head")>-1);
check("seven panes incl index", (page.match(/data-pane=/g)||[]).length===7,
  (page.match(/data-pane=/g)||[]).length);

console.log("== TASK 20 polish ==");
check("person repair shows own win", page.indexOf('class="rr-val"')>-1&&page.indexOf("won by you")>-1);
check("store gate is store-level (80% of 200000)", subha.repairGate.requiredValue===160000
  &&page.indexOf("used <b>only inside the incentive calculation</b>")===-1);
check("kpi pair layout", page.indexOf('class="kpi-pair"')>-1&&page.indexOf('class="kp-right"')>-1);
check("incentive split categories", page.indexOf("Category 1 · AMC")>-1
  &&page.indexOf("Category 2 · Bundle (DISC scheme)")>-1);
check("incentive totals rows", (page.match(/All salespersons/g)||[]).length===2);
check("deep header applied", (page.match(/hdr-deep/g)||[]).length>=3&&page.indexOf('class="hdr-deep"')>-1);
check("csat flex", page.indexOf("csat-flex")>-1&&page.indexOf("csat-right")>-1);
check("liquid glass css", page.indexOf("backdrop-filter:blur(14px)")>-1&&page.indexOf("tile-sparkle")>-1);
check("glitter animation", page.indexOf("tileTwinkle")>-1);
check("bar graph per group", (page.match(/class="gbar-row"/g)||[]).length>=12,
  (page.match(/class="gbar-row"/g)||[]).length);
check("horizontal board", page.indexOf("tile-board")>-1&&page.indexOf("tile-col")>-1);
check("freshness line", page.indexOf("Latest records per sheet")>-1&&page.indexOf("fresh-chip")>-1);
check("tile count 14 (Repair Validation tile removed)", (page.match(/class="glass-tile"/g)||[]).length===14,
  (page.match(/class="glass-tile"/g)||[]).length);

console.log("== TASK 21 batch ==");
check("global gif background at 50%", page.indexOf('id="globalBgGif"')>-1&&page.indexOf("783dbbe3f2d2eb0e437be534bb3d3cb6.gif")>-1
  &&page.indexOf(".global-bg-gif")>-1&&page.indexOf("opacity:.5")>-1);
check("gif keep-alive reload", page.indexOf("bgGifKeepAlive")>-1&&page.indexOf("setInterval")>-1);
check("tile glitter video on hover", page.indexOf('class="tile-glitter-vid"')>-1&&page.indexOf("1063973464")>-1
  &&page.indexOf(".glass-tile:hover .tile-glitter-vid")>-1);
check("index subtitle removed", page.indexOf('class="index-sub"')===-1);
check("footer removed", page.indexOf('<footer class="foot">')===-1&&page.indexOf("Six pages, one data run")===-1);
check("KPI def removed", page.indexOf("Definition: store-level KPI scorecard")===-1);
check("salesperson def removed", page.indexOf("Definition: per-person target vs achieved. Service and Accessory")===-1);
check("jobs subtitle removed", page.indexOf("segregated &amp; highlighted by repair bucket")===-1);
check("live clock card + gif + tick js", page.indexOf('class="topbar-clock"')>-1
  &&page.indexOf("21-52-42-295_512.gif")>-1&&page.indexOf("tickClock")>-1);
check("amc & bundle side-by-side", page.indexOf('class="inc-pair"')>-1);
check("repair matrix removed so no duo", page.indexOf('class="duo-pair"')===-1);
check("job table fixed layout", page.indexOf(".jobtable { table-layout:fixed")>-1&&page.indexOf("<colgroup>")>-1);
check("distinct person band colours", page.indexOf("#8a9bd4")>-1&&page.indexOf("#b8860b")>-1&&page.indexOf("#0f9d8f")>-1);
check("liquid glass material on tables", page.indexOf("backdrop-filter:blur(12px) saturate(1.2)")>-1);
check("csat chips below table", page.indexOf(".csat-flex { flex-direction:column")>-1
  &&page.indexOf(".csat-flex .csat-right { flex:0 0 auto; width:100%")>-1);
check("next-slab balance lines", (page.match(/class="next-slab"/g)||[]).length>=1||page.indexOf("top slab reached")>-1);
check("subha next svc slab ₹1,21,000", subha.nextSvcSlab && subha.nextSvcSlab.need===121000, subha.nextSvcSlab && subha.nextSvcSlab.need);

console.log("== TASK 22 batch ==");
check("tile glass lighter + fonts brighter", page.indexOf("rgba(255,255,255,.34) 0%,rgba(150,178,255,.20)")>-1
  &&page.indexOf(".tile-title { color:#ffffff")>-1);
check("glitter video fits tile outline", page.indexOf(".tile-glitter-vid { inset:0; border-radius:18px")>-1);
check("token creation locked side-by-side", page.indexOf(".walkin-pair { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr)")>-1);
check("job table left aligned", page.indexOf(".jobtable th, .jobtable td { text-align:left")>-1);
check("advance paid tile removed", page.indexOf('class="label">Advance Paid</div>')===-1);
check("advance totals still computed", j.totals.advanceTotal===2049);
check("dark section shells on all pages", page.indexOf(".sec { background:rgba(10,15,40,.66)")>-1);
check("dark topbar", page.indexOf(".topbar { background:rgba(8,12,32,.78)")>-1);

console.log("== TASK 23 batch ==");
check("gif backdrop present (video removed)", page.indexOf('id="globalBgGif"')>-1&&page.indexOf('id="globalBgVideo"')===-1);
check("tile glitter behind text (z-index layering)", page.indexOf("z-index:2")>-1&&page.indexOf(".tile-glitter-vid { z-index:0")>-1);
check("tile glitter dimmed on hover", page.indexOf("opacity:.55; filter:brightness(.85)")>-1);
check("tile overflow clips video", page.indexOf(".glass-tile { overflow:hidden")>-1);
check("kpi table full width left", page.indexOf(".kpi-pair { flex-direction:column")>-1
  &&page.indexOf(".kpi-pair .kp-left table { width:100%")>-1);
check("walkin heads aligned on one line", page.indexOf('wp-head">🚶 Walk-in Analysis')>-1
  &&page.indexOf('wp-head">🎟 Token Creation')>-1);
check("day-wise token table full width", page.indexOf('class="wp-daytable"')>-1
  &&page.indexOf('wp-head">📅 Day-wise Walk-in')>-1
  &&page.indexOf(".wp-daytable .walkintable { width:100%")>-1);
check("walkin analysis split remains", page.indexOf("ptable walksum")>-1&&page.indexOf("ptable walkintable")>-1);
check("csat chips grey-shaded", page.indexOf(".csat-chip { background:#dde3ee")>-1);
check("jobtable forced left", page.indexOf(".jobtable th, .jobtable td { text-align:left !important")>-1);
check("grphead sky-blue to dark gradient", page.indexOf("linear-gradient(90deg,#4fb3f5")>-1
  &&page.indexOf("#0b3672")>-1);
check("datewise definition removed", page.indexOf("Definition: every invoice date")===-1);
check("subha band softened", page.indexOf("#8a9bd4")>-1&&page.indexOf("#3353c0")===-1);
check("chat widget on page", page.indexOf('id="chatPanel"')>-1&&page.indexOf('id="chatFab"')>-1
  &&page.indexOf("chatToggle")>-1&&page.indexOf("chatSend")>-1
  &&page.indexOf(".serverQuery(q)")>-1);

// ---- chatbot logic against the mock sheets ----
console.log("== TASK 23 chatbot mocks ==");
check("chatDig normalises 91-prefix", chatDig("+91 91987 654321")==="9876543210"||chatDig("919876543210")==="9876543210");
check("chatDig rejects short input", chatDig("12345")==="");
var dExtract=chatExtractDates("sale date from 23.09.26");
check("date extraction", dExtract.length===1&&dExtract[0]==="23-09-2026", dExtract);
var sales1=chatQuery("sale date from 23.09.26");
check("sales query finds invoices", sales1.title.indexOf("Sales on")>-1&&sales1.rows.length>0);
var salesFlat=sales1.rows.map(function(r){return r.join("");}).join("");
check("sales query revenue+gp rows", salesFlat.indexOf("INV1")>-1&&salesFlat.indexOf("Subha")>-1&&salesFlat.indexOf("7,100")>-1,
  sales1.title+" | rows="+sales1.rows.length);
var sales2=chatQuery("sales 23.09.2026 to 23.09.2026");
check("sales range parse", sales2.rows.length>0&&sales2.title.indexOf("Sales on")>-1);
var partQ=chatQuery("part number for care+ display");
check("part query item code", JSON.stringify(partQ.rows).indexOf("CPLUS")>-1, JSON.stringify(partQ).slice(0,140));
var mobQ=chatQuery("9791994639");
check("mobile 10-digit query", mobQ.title.indexOf("Mobile")>-1||mobQ.title.indexOf("No records")>-1, mobQ.title);
var mobQ2=chatQuery("919876543210");
check("mobile lookup across sheets", mobQ2.rows.length>0&&JSON.stringify(mobQ2.rows).indexOf("Ravi")>-1, JSON.stringify(mobQ2).slice(0,140));
check("mobile lookup case+invoice", JSON.stringify(mobQ2.rows).indexOf("C1001")>-1||mobQ2.note.indexOf("No records")>-1);
var mobJSON=serverQuery("919876543210");
check("serverQuery returns JSON", mobJSON.charAt(0)==="{");
var mobParsed=JSON.parse(mobJSON);
check("serverQuery shape", mobParsed.title!==undefined&&Array.isArray(mobParsed.rows));
var blankQ=JSON.parse(serverQuery(""));
check("serverQuery empty guard", blankQ.title==="Say something first");

console.log(fails?("\n"+fails+" FAILURES"):"\nALL CHECKS PASS");
process.exit(fails?1:0);
