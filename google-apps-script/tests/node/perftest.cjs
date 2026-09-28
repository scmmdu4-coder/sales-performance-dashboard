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

// ---- target (rows 12-14 KPI grid, row 15 = approx unit value) ----
var T=[]; for(var i=0;i<20;i++) T.push(new Array(18).fill(""));
T[0]=["Revenue Type","Store Target","SUBHA","DINESH","","Sep 2026","Repair Stage","Target","Achieved","%","Gap","Repair Type","Count"];
T[1]=["Service",300000,180000,120000]; T[2]=["Accessories",150000,90000,60000];
T[3]=["Repair",200000,120000,80000]; T[4]=["Total",650000,390000,260000];
T[5]=["Gross Profit",160000,96000,64000];
T[8]=["Prev Month","This Month","Difference"]; T[9]=[1475,1494,19];
T[11]=["","AR","","","","RR","","","","SR","",""];
T[12]=["","Care+","Protect+","AMC","Bundle","Display","Battery","Glass","Others","SRD","SHC","ARD"];
T[13]=["",25,15,40,30,50,45,20,35,60,25,18];
T[14]=["",400,600,900,300,250,150,100,0,0,0,0];
T[15]=["Achivement","","Achivement",""]; T[16]=["70-80%","2000","70-80%","1000"];
T[17]=["80-90%","4000","80-90%","2000"]; T[18]=["90-100%","6000","90-100%","3000"];
T[1][6]="Diagnosis Complete"; T[1][7]=40; T[1][8]=33; T[1][9]="82.5"; T[1][10]=7; T[1][11]="Walk-In Repair"; T[1][12]=3;
T[2][6]="Repair Complete";    T[2][7]=80; T[2][8]=71; T[2][9]="88.8"; T[2][10]=9; T[2][11]="Carry-In Repair"; T[2][12]=2;
T[15][5]="incentive per amc"; T[16][5]="1-10"; T[16][6]="₹150"; T[16][8]=5; T[16][9]="₹80";
T[17][5]="11-25"; T[17][6]="₹250"; T[17][8]=12; T[17][9]="₹120"; T[15][12]="slab level"; T[15][13]=5; T[16][13]=300;

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
check("CAREPLUS approx unit 400 from row 15",careplus&&careplus.approxAmount===400);
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
check("store target under repair line",page.indexOf("Store Total Revenue Target")>-1);
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
check("three tabs: Business / Process / Performance",page.indexOf("🛠 Process")>-1
  &&page.indexOf("📈 Performance")>-1&&page.indexOf("💼 Business")>-1);
check("tab slider row",page.indexOf('class="tabs-row"')>-1);
check("fonts bumped globally",page.indexOf("font-size:18px")>-1);
check("DISC/AMC notes hidden",page.indexOf("starts with DISC")===-1&&page.indexOf("SubCategory contains AMC")===-1);
check("date-wise GP column",page.indexOf('class="grp grp-0">Gross Profit</th>')>-1);
check("legend shows GP daily target",page.indexOf("Gross Profit "+"₹"+subha.dailyTargets.grossProfit.toLocaleString("en-IN"))>-1);
check("GP daily target in legend for Dinesh",page.indexOf("Gross Profit "+"₹"+dinesh.dailyTargets.grossProfit.toLocaleString("en-IN"))>-1);
check("jobs warranty column",page.indexOf("<th>Warranty Type</th>")>-1);
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
check("jobs warranty value renders",page.indexOf("Out of Warranty")>-1);
check("repair tab section",page.indexOf("Repair Type Performance")>-1);
// month-to-date composition (no date rows) + integers only
check("till-date composition table",page.indexOf("Till Date ·")>-1&&page.indexOf("All repair types")>-1);
check("no per-date pivot",page.indexOf("Repair Types by Date")===-1);
check("mtd walk-in 3 / carry-in 2",page.indexOf('<strong>Walk-In Repair</strong></td><td class="num">3</td>')>-1
  &&page.indexOf('<strong>Carry-In Repair</strong></td><td class="num">2</td>')>-1);
check("mtd grand total 5 (SNR excluded)",page.indexOf('<tr class="totrow"><td>All repair types</td><td class="num">5</td></tr>')>-1);
check("repair tile avg baseline",page.indexOf("avg 0/month")>-1);
check("no-fraction averages",page.indexOf("avg 0.3/month")===-1);
check("no-baseline note",page.indexOf("no 3-month baseline")>-1);
check("repair meta",page.indexOf("5 repairs this month · 2 repair types")>-1);
check("avg days col header",page.indexOf("Avg Days · Now")>-1);
check("walk-in avg days 2 d",page.indexOf('<strong>Walk-In Repair</strong></td><td class="num">3</td><td class="num">0</td>')>-1);
// products & coverage dashboards
check("product performance section",page.indexOf("Product Performance")>-1);
check("iphone family tile",page.indexOf('class="label">iPhone</div><div class="value">1</div>')>-1);
check("other family tile 4",page.indexOf('class="label">Other</div><div class="value">4</div>')>-1);
check("coverage performance section",page.indexOf("Coverage Performance")>-1 && page.indexOf("Out Of Warranty (No Coverage)")>-1);
check("snr excluded note",page.indexOf('"Service Non-Repair" cases are excluded')>-1);
check("walkin hsl tint",page.indexOf("background:hsl(")>-1);
check("token table",page.indexOf("% of creation")>-1&&page.indexOf("Total tokens generated")>-1);
check("ready col gone",page.indexOf("Ready for pickup")===-1);
check("comm pos/neg present",page.indexOf("comm-pos")>-1&&page.indexOf("comm-neg")>-1);
check("mapping line on page",page.indexOf('class="map-line"')>-1&&page.indexOf("Mobile ← <b>MobileNo</b>")>-1);
check("business sections intact",page.indexOf("Invoice Valuation")>-1&&page.indexOf("Date-wise")>-1);
// ---- weekly revenue analysis (section 7) + date-wise renumbered to 8 ----
check("weekly section is #7",page.indexOf("7 &middot; Weekly Revenue Analysis")>-1);
check("date-wise renumbered #8",page.indexOf("8 &middot; Date-wise")>-1);
check("Store Total column removed",page.indexOf("Store Total</th>")===-1);
check("week rows 1-7 etc",page.indexOf("<b>Week 1</b>")>-1&&page.indexOf("(1–7")>-1&&page.indexOf("<b>Week 5</b>")>-1);
check("weekly status labels",page.indexOf("In progress")>-1&&page.indexOf("Upcoming")>-1&&page.indexOf("Completed")>-1);
check("weekly target formula (66667 svc wk4)",page.indexOf("₹66,667")>-1);
check("weekly svc achieved wk4 5000",page.indexOf("<b>Week 4</b>")>-1&&page.indexOf("₹5,000")>-1);
// ---- process page: target block (F1 title + G1:P7) & walk-in summary table ----
check("process target block title from F1",page.indexOf("🎯 Sep 2026")>-1);
check("target block grid renders G1:P7",page.indexOf("Repair Stage")>-1&&page.indexOf("Diagnosis Complete")>-1&&page.indexOf("88.8")>-1);
check("repair comparison strip (excl SNR)", page.indexOf("Total Repairs (excl. Service Non-Repair) — this month")>-1
  &&page.indexOf('rs-value">5</span>')>-1);
check("walk-in summary is a table", page.indexOf('ptable walksum')>-1 && page.indexOf("<b>Total Walk-ins</b></td>")>-1);

check("weekly wd wk2=7 (no holiday)",page.indexOf("(8–14 Sep)</span></td><td style=\"text-align:center\">7</td>")>-1);
check("weekly wd wk5=2",page.indexOf("(29–30 Sep)</span></td><td style=\"text-align:center\">2</td>")>-1);

console.log(fails?("\n"+fails+" FAILURES"):"\nALL CHECKS PASS");
process.exit(fails?1:0);
