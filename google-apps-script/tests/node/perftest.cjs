/**
 * Node test harness for SalesDashboard.gs (Performance + Business logic).
 * Run:  node google-apps-script/tests/node/perftest.js
 * It stubs SpreadsheetApp / CacheService / etc., feeds mock sheets through the
 * real script, and asserts the computed values and rendered HTML.
 */
// Freeze "today" so month-relative assertions stay stable year-round: the
// mock sheets are authored as September-2026 data (clock mock says
// "24 Sep 2026, 10:30"), so the harness pins the clock to that day.
var __REAL_DATE__=Date;
(function(){
  var FROZEN_TS=new __REAL_DATE__(2026,8,24,10,30,0).getTime();
  function FakeDate(){
    if(!(this instanceof FakeDate)) return __REAL_DATE__.apply(null,arguments).toString();
    if(arguments.length===0) return new __REAL_DATE__(FROZEN_TS);
    if(arguments.length===1) return new __REAL_DATE__(arguments[0]);
    var a=Array.prototype.slice.call(arguments);
    return new (Function.prototype.bind.apply(__REAL_DATE__,[null].concat(a)))();
  }
  FakeDate.prototype=__REAL_DATE__.prototype;
  FakeDate.now=function(){return FROZEN_TS;};
  FakeDate.UTC=__REAL_DATE__.UTC; FakeDate.parse=__REAL_DATE__.parse;
  global.Date=FakeDate;
})();
global.Utilities={formatDate:function(d,tz,f){
  if(f&&f.indexOf("HH")>-1) return "24 Sep 2026, 10:30";
  function p(n){return(n<10?"0":"")+n;}
  return p(d.getDate())+"-"+p(d.getMonth()+1)+"-"+d.getFullYear();}};
global.Session={getScriptTimeZone:function(){return "Asia/Calcutta";}};
global.CacheService={getScriptCache:function(){return{get:function(){return null},put:function(){}}}};
global.HtmlService={createHtmlOutput:function(){return{addMetaTag:function(){return this},setTitle:function(){return this}}}};
function mockSheet(name, values){var sh={getName:function(){return name},
  getLastColumn:function(){return values[0]?values[0].length:1;},
  getLastRow:function(){return values.length;},
  __writes:[],
  getRange:function(r,c,nr,nc){return{getValue:function(){return values[r-1]&&values[r-1][c-1]||""},
    getValues:function(){var out=[];for(var i=0;i<(nr||1);i++){var row=[];
      for(var j=0;j<(nc||1);j++){row.push((values[r-1+i]&&values[r-1+i][c-1+j])||"");}
      out.push(row);}return out;},
    setValue:function(v){sh.__writes.push({row:r,col:c,ncols:nc||1,value:v});
      if(!values[r-1])values[r-1]=[];values[r-1][c-1]=v;},
    setValues:function(vs){vs.forEach(function(row,i){row.forEach(function(v,j){
      if(!values[r-1+i])values[r-1+i]=[];values[r-1+i][c-1+j]=v;});});
      sh.__writes.push({row:r,col:c,value:vs});}}},
  getDataRange:function(){return{
  getValues:function(){return values},
  getDisplayValues:function(){return values.map(function(r){return r.map(function(v){
    return(v===null||v===undefined||v instanceof Date)?"":String(v);})});}};}};
  return sh;}
var ALL_SHEETS=[];
var SS_OBJ={getSheets:function(){return ALL_SHEETS;},
  getSheetByName:function(n){for(var i=0;i<ALL_SHEETS.length;i++)if(ALL_SHEETS[i].getName()===n)return ALL_SHEETS[i];return null;}};
global.SpreadsheetApp={getActiveSpreadsheet:function(){return SS_OBJ;}};
global.__WA_PROPS={};
global.PropertiesService={getScriptProperties:function(){return{
  getProperty:function(k){return (k in global.__WA_PROPS)?global.__WA_PROPS[k]:null;}};}};
global.__WA_CALLS=[];
global.ContentService={createTextOutput:function(t){var o={t:t,setMimeType:function(m){o.m=m;return o;}};
  return o;},MimeType:{JSON:"application/json"}};
global.UrlFetchApp={fetch:function(url,opts){
  global.__WA_CALLS.push({url:url,opts:opts});
  var payload={}; try{payload=JSON.parse(opts.payload||"{}");}catch(e){}
  return{getResponseCode:function(){return 200;},
    getContentText:function(){return JSON.stringify({messageId:"wamid.mock.1",echo:payload});}};}};
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
 "SubCategory","SchemeCode","TaxableAmount","GrossProfit","TranType","ItemCode","PartType","OrderCode"],
 ["INV1","23-09-2026 10:15","Subha","Service","Svc","","",5000,1500,"SALES","","","C1001"],
 ["INV2","23-09-2026 11:00","Dinesh","Accessories","Case","","",1200,300,"SALES","","",""],
 ["INV3","23-09-2026 12:00","Subha","Accessories","Care+ plan","","",900,250,"SALES","CPLUS","",""]];

// ---- walkin ----
var wl=[["MobileNo","TokenDate","VisitPurposeDesc","CaseId","InvoiceCode","Remark","TokenCreatedBy","CustomerVisitSource"],
 ["919876543210","2026-09-23 09:15:00","Repair Submission","C1001","IV9001","Needs original glass","Subha","Walk-by"],
 ["919876543211","2026-09-23 09:40:00","Repair Submission","C1002","IV9002","Warranty valid","Subha","Walk-by"],
 ["919876543212","2026-09-23 10:05:00","General/Sales Enquiry","C1003","","Follow up in 2 days","Dinesh","Google"],
 ["919876543213","2026-09-23 11:15:00","Collection","","IV9003","Bought case","Dinesh","Google"],
 ["919876543214","2026-09-23 12:30:00","General/Sales Enquiry","","","","Dinesh","Reference"],
 ["919876543215","2026-09-22 10:20:00","Repair Submission","C1000","","Waiting for part","Subha","Walk-by"],
 ["919876543216","2026-09-22 13:00:00","Collection","C0999","IV8999","","Priya","Reference"]];

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
 "CoverageCodeDescription","RepairDocType","Ready_For_Pickup_Date_CRM","GSXCode","SerialNo1","PartDescription"],
 // duplicate of C1003 with an OLDER case date placed BEFORE the real row —
 // the merger must keep the later (real) row instead.
 ["C1003","Mani","Kandan","919876543212","Different Phone",daysAgoStr(5,"10"),"",5,"Open","","Diagnosis In Progress","","In Warranty","Walk-In Repair","","GSX6003-OLD","SN-OLD",""],
 ["C1001","Ravi","Kumar","919876543210","iPhone 13",daysAgoStr(14,"10"),daysAgoStr(13,"11"),14,"Open","","Repair Complete","Original Display Used","In Warranty","Walk-In Repair",daysAgoStr(12,"10"),"GSX6001","SN-C1001","Battery"],
 ["C1002","Anita","Devi","919876543211","Samsung S23",daysAgoStr(4,"09"),daysAgoStr(3,"14"),4,"Open","","Unclaimed Device","Part Used","Out of Warranty","Carry-In Repair",daysAgoStr(1,"12"),"GSX6002","SN-C1002",""],
 ["C1003","Mani","Kandan","919876543212","OnePlus 11",daysAgoStr(2,"10"),"",2,"Open","","Diagnosis In Progress","Device Unclaimed Device due","Out Of Warranty (No Coverage)","Walk-In Repair","","GSX6003","SN-C1003",""],
 ["C1004","Peela","Dass","919876543217","Realme 12",daysAgoStr(1,"08"),"",1,"Open","","Repair In Progress","Ready For Pickup","In Warranty","Service Non-Repair","","GSX6004","SN-C1004",""],
 ["C1005","Old","Closed","919876543218","Pixel 7",daysAgoStr(55,"09"),daysAgoStr(54,"10"),55,"Delivered",daysAgoStr(50,"10"),"Repair Complete","Part Used","In Warranty","Carry-In Repair",daysAgoStr(50,"10"),"GSX6005","SN-C1005",""],
 ["C1006","Age","Fallback","919876543219","Moto G",daysAgoStr(20,"09"),"","","Open","","Waiting for Part","","Out Of Warranty (No Coverage)","Walk-In Repair","","GSX6006","SN-C1006",""],
 ["C1007","Can","Cell","919876543220","Nokia",daysAgoStr(2,"09"),"",2,"Cancelled","","Aborted","","In Warranty","Carry-In Repair","","GSX6007","SN-C1007",""]];
// ---- advance ----
var av=[["MobileNo","AvailableBalance"],["919876543210",500],["919876543210",200],["919876543217",350],["919876543219",999]];

// ---- decline / quote (Task 35) ----
var dl=[["CaseId","GSXRepairStatusDescription","Product"],
 ["D1001","Declined - Rejected by Local Approver","iPhone 13"],
 ["D1002","Declined - Rejected by Local Approver","Galaxy S22"],
 ["D1002","Declined - Rejected by Local Approver","Galaxy S22"],
 ["D1003","Cancelled","Pixel 7"],
 ["D1001","Declined - Rejected by Local Approver","iPhone 13"]];
var qt=[["CaseID","RevenueType","lastAction"],
 ["Q2001","Repair Revenue","approved"],
 ["Q2002","Repair Revenue","approved"],
 ["Q2002","Repair Revenue","approved"],
 ["Q2003","Repair Revenue","cancelled"],
 ["Q2001","Repair Revenue","approved"],
 ["Q2004","Accessory Revenue","approved"]];
ALL_SHEETS=[mockSheet("target",T),mockSheet("crm",crm),mockSheet("walkin",wl),
  mockSheet("service",sv),mockSheet("advance",av),mockSheet("comm",cm),
  mockSheet("decline",dl),mockSheet("quote",qt)];

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
// Title from crm!D2 (mock D2 = "Service" — the mock's 4th column carries it)
check("title uses crm D2 (Task 39)",page.indexOf("iPlanetcare Service — Store Performance Dashboard")>-1);
check("two-line brand: iPlanetcare + Store Performance Dashboard (Task 35)",
  page.indexOf('<div class="brand">')>-1&&page.indexOf("iPlanetcare Service</div>")>-1
  &&page.indexOf('class="brand-sub"')>-1&&page.indexOf("Store Performance Dashboard</div>")>-1);
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
check("job table 11-col owner order (Task 38: CRM Status column removed)",page.indexOf(
  "<th>Repair Bucket</th><th>Aging</th><th>GSX Repair Status</th>"+
  "<th>Repair Type</th>"+
  "<th>Case ID</th><th>Name</th><th>Number</th>"+
  "<th>GSX ID</th><th>S.No</th><th>Model</th>"+
  "<th>Comm. Updated</th>")>-1&&page.indexOf("<th>CRM Status</th>")===-1);
check("old job columns dropped",page.indexOf("<th>Warranty Type</th>")===-1
  &&page.indexOf("<th>Job created</th>")===-1&&page.indexOf("<th>Advance paid</th>")===-1
  &&page.indexOf("<th>Communicated</th>")===-1);
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
check("walkin rows: single flat navy (Task 31)",page.indexOf(".walkintable tbody td { background-color: rgba(34,43,90,.92) !important; }")>-1
  &&page.indexOf("--tint")===-1&&page.indexOf("background:hsl(")===-1);
check("token table",page.indexOf("% of creation")>-1&&page.indexOf("Total tokens generated")>-1);
check("ready col gone",page.indexOf("Ready for pickup")===-1);
check("comm pos/neg present",page.indexOf("comm-pos")>-1&&page.indexOf("comm-neg")>-1);
check("comm cells soft palette (Task 35)", page.indexOf(".comm-pos { color:#0f5a37 !important; background:#b9f2cf")>-1
  &&page.indexOf(".comm-neg { color:#8a1d28 !important; background:#ffd6da")>-1
  &&page.indexOf("background:#0f9d59")===-1&&page.indexOf("background:#c62a33")===-1);
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
check("decline/requote live cards (Task 35)",
  page.indexOf("🚫 Decline Repair")>-1&&page.indexOf("🔁 Requote &amp; Win-back")>-1
  &&page.indexOf('<div class="label">Declined repairs</div><div class="value">2</div>')>-1
  &&page.indexOf('<div class="label">Win-backs</div><div class="value">2</div>')>-1
  &&page.indexOf('ph-tile')===-1);
check("vmatrix cells one navy (Task 31)", page.indexOf("table.vmatrix td { background:rgba(34,43,90,.92); color:#e9edff")>-1);
check("dark-blue header skin throughout (Task 35)", page.indexOf("#5b7bff 0%,#4360e0")===-1
  &&page.indexOf("linear-gradient(180deg,#5b7bff")===-1
  &&page.indexOf(".ptable thead th { background:#1c2b52")>-1
  &&(page.match(/#1c2b52/g)||[]).length>=8
  &&(page.match(/#101a38/g)||[]).length>=5);
check("3d bar styling (single colour, Task 31)", page.indexOf("linear-gradient(90deg,#2f8dff,#3ce3b4)")>-1
  &&page.indexOf("--barc")===-1&&page.indexOf("barShine")>-1&&page.indexOf(".bar-fill::before")>-1
  &&page.indexOf(".bar-fill.good")===-1&&page.indexOf(".bar-fill.bad")===-1);
check("index page pane", page.indexOf('data-pane="index"')>-1&&page.indexOf("index-hero")>-1);
check("index tiles built (11 now, matrix tile removed)", (page.match(/class="glass-tile"/g)||[]).length>=11,
  (page.match(/class="glass-tile"/g)||[]).length);
check("tile group heads", page.indexOf("tile-group-head")>-1);
check("eight panes incl deck (Task 24)", (page.match(/data-pane=/g)||[]).length===8,
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
check("lighter plain background for all pages (Task 31)", page.indexOf("body { background:#3f444d")>-1
  &&page.indexOf('id="globalBgGif"')===-1&&page.indexOf(".global-bg-gif {")===-1);
check("static background (no keep-alive needed)", page.indexOf("bgGifKeepAlive")===-1);
check("tile glitter video on hover", page.indexOf('class="tile-glitter-vid"')>-1&&page.indexOf("1064124616")>-1
  &&page.indexOf(".glass-tile:hover .tile-glitter-vid")>-1&&page.indexOf("shining-gold-dust-particles")>-1);
check("index subtitle removed", page.indexOf('class="index-sub"')===-1);
check("footer removed", page.indexOf('<footer class="foot">')===-1&&page.indexOf("Six pages, one data run")===-1);
check("KPI def removed", page.indexOf("Definition: store-level KPI scorecard")===-1);
check("salesperson def removed", page.indexOf("Definition: per-person target vs achieved. Service and Accessory")===-1);
check("jobs subtitle removed", page.indexOf("segregated &amp; highlighted by repair bucket")===-1);
check("live clock card, no gif + tick js (Task 30)", page.indexOf('class="topbar-clock"')>-1
  &&page.indexOf("AC40")===-1&&page.indexOf("clock-gif")===-1&&page.indexOf("tickClock")>-1);
check("amc & bundle side-by-side", page.indexOf('class="inc-pair"')>-1);
check("repair matrix removed so no duo", page.indexOf('class="duo-pair"')===-1);
check("job table auto layout (Task 38)", page.indexOf(".jobtable { table-layout:auto")>-1&&page.indexOf("<colgroup>")>-1);

console.log("== TASK 39 ==");
check("weekly revenue-type colours on groups", page.indexOf('grp grp-0 wt-svc')>-1
  &&page.indexOf('wt-acc')>-1&&page.indexOf('wt-rep')>-1&&page.indexOf('wt-gp')>-1);
check("weekly wt css", page.indexOf("#sec-weekly td.wt-svc { background:rgba(90,150,255,.14)")>-1
  &&page.indexOf("#sec-weekly th.wt-gp  { color:#b9f08c")>-1);
var opsPair=page.slice(page.indexOf('class="ops-pair"'),page.indexOf('class="ops-pair"')+6000);
check("decline & win-back side by side (ops-pair)",
  page.indexOf('class="ops-pair"')>-1
  &&page.indexOf(".ops-pair { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr)")>-1
  &&opsPair.indexOf("Decline Repair</h2>")>-1&&opsPair.indexOf("Win-back</h2>")>-1
  &&opsPair.indexOf("Decline Repair</h2>")<opsPair.indexOf("Win-back</h2>"));
check("repair type renders N-repair", page.indexOf("<td>N-repair</td>")>-1);
check("repairTypeShort_ mappings", repairTypeShort_("Service Non-Repair")==="N-repair"
  &&repairTypeShort_("Whole Unit Mail-In Return to Service Location")==="mail-in"
  &&repairTypeShort_("Walk-In Repair")==="Walk-In Repair");
check("case badges stacked under the case id", page.indexOf('class="casebadges"')>-1
  &&page.indexOf(".casebadges { display:flex; flex-direction:column")>-1);
check("deck: 8 tables as 4 side-by-side pairs",
  (page.match(/class="deck-pair"/g)||[]).length===4
  &&page.indexOf(".deck-pair { display:grid")>-1);
check("freshness chips carry date AND time", page.indexOf("23-09-2026 12:00")>-1);
check("person band colours removed: one table color set (Task 31)",
  page.indexOf("#8a9bd4")===-1&&page.indexOf("#b8860b")===-1&&page.indexOf("#0f9d8f")===-1
  &&page.indexOf("#7a3fd4")===-1&&page.indexOf("#c04a68")===-1
  &&page.indexOf("td.cell-0")===-1&&page.indexOf("th.grp-3")===-1);
check("liquid glass material on tables", page.indexOf("backdrop-filter:blur(12px) saturate(1.2)")>-1);
check("csat table-left / chips-right (Task 35)", page.indexOf(".csat-flex { flex-direction:column")===-1
  &&page.indexOf(".csat-flex { flex-direction:row")>-1
  &&page.indexOf(".csat-flex table.csat-table { width:auto; min-width:0")>-1
  &&page.indexOf(".csat-flex .csat-right { flex:0 0 240px")>-1);
check("next-slab balance lines", (page.match(/class="next-slab"/g)||[]).length>=1||page.indexOf("top slab reached")>-1);
check("subha next svc slab ₹1,22,800 (71% tier, Task 33)", subha.nextSvcSlab && subha.nextSvcSlab.need===122800, subha.nextSvcSlab && subha.nextSvcSlab.need);
check("subha next acc slab 51% tier (Task 33)", subha.nextAccSlab && subha.nextAccSlab.pct===51
  &&subha.nextAccSlab.need===Math.ceil(0.51*90000)-900, subha.nextAccSlab);

console.log("== TASK 22 batch ==");
check("tile glass flat navy + fonts brighter", page.indexOf("background:rgba(28,43,82,.98)")>-1
  &&page.indexOf(".tile-title { color:#ffffff")>-1);
check("glitter video fits tile outline", page.indexOf(".tile-glitter-vid { inset:0; border-radius:18px")>-1);
check("token creation locked side-by-side", page.indexOf(".walkin-pair { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr)")>-1);
check("job table left aligned", page.indexOf(".jobtable th, .jobtable td { text-align:left")>-1);
check("advance paid tile removed", page.indexOf('class="label">Advance Paid</div>')===-1);
check("advance totals still computed", j.totals.advanceTotal===2049);
check("section shells use the one navy surface (Task 31)", page.indexOf(".sec { background:rgba(28,43,82,.98)")>-1);
check("dark topbar", page.indexOf(".topbar { background:rgba(8,12,32,.78)")>-1);

console.log("== TASK 23 batch ==");
check("no gif/video backdrop (Task 30 plain ash)", page.indexOf('id="globalBgGif"')===-1&&page.indexOf('id="globalBgVideo"')===-1);
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

console.log("== TASK 37: customer source analysis + layout fits ==");
check("source pivot pivot present (srcstable)", page.indexOf("ptable srcstable")>-1);
check("source table sits between walkin analysis & tokens",
  page.indexOf('<div class="wp-left">')<page.indexOf('<div class="wp-mid">')&&page.indexOf('<div class="wp-mid">')<page.indexOf('<div class="wp-right">'));
check("source pivot numbers",
  page.indexOf("<td><b>Walk-by</b></td><td class=\"num\">3</td><td class=\"num\">42.9%</td>")>-1
  &&page.indexOf("<td><b>Google</b></td><td class=\"num\">2</td><td class=\"num\">28.6%</td>")>-1
  &&page.indexOf("<td><b>Reference</b></td><td class=\"num\">2</td><td class=\"num\">28.6%</td>")>-1);
check("source total row", page.indexOf("Total walk-ins</b></td><td class=\"num\"><b>7</b></td><td class=\"num\"><b>100%</b></td>")>-1);
check("walkin pair is three panes now", page.indexOf(".walkin-pair { grid-template-columns:minmax(0,1fr) minmax(0,1fr) minmax(0,1fr); }")>-1);
check("jobs table single-line, nothing hidden (Task 38)", page.indexOf(".jobtable th, .jobtable td { white-space:nowrap; overflow:visible; text-overflow:clip; }")>-1);
check("csat table fills its lane", page.indexOf(".csat-flex .csat-left table.csat-table { width:100%; table-layout:fixed; }")>-1);
check("csat chips flat navy + light text", page.indexOf(".csat-chip { background:rgba(28,43,82,.98)")
  >-1&&page.indexOf(".csat-chip .cc-value { font-size:1.5rem; font-weight:800; color:#f4f8ff")>-1);
check("jobtable forced left", page.indexOf(".jobtable th, .jobtable td { text-align:left !important")>-1);
check("grphead dark-blue default header (Task 35)", page.indexOf(".grphead td { background:#1c2b52")>-1
  &&page.indexOf("linear-gradient(180deg,#5b7bff 0%,#4360e0 55%,#3350c2 100%)")===-1
  &&page.indexOf("linear-gradient(90deg,#4fb3f5")===-1);
check("datewise definition removed", page.indexOf("Definition: every invoice date")===-1);
check("one surface everywhere (Task 31)", page.indexOf("card-unclaimed { background:rgba(28,43,82,.98)")>-1
  &&page.indexOf(".rtile.done { background:rgba(28,43,82,.98)")>-1
  &&page.indexOf("#57f2bd")===-1
  &&page.indexOf("linear-gradient(145deg,rgba(11,93,71")===-1
  &&page.indexOf("linear-gradient(145deg,rgba(90,62,8")===-1);
check("chat widget on page", page.indexOf('id="chatPanel"')>-1&&page.indexOf('id="chatFab"')>-1
  &&page.indexOf("chatToggle")>-1&&page.indexOf("chatSend")>-1
  &&page.indexOf(".serverQuery(q)")>-1);

// ---- chatbot logic against the mock sheets (Task 30: limited bot) ----
console.log("== TASK 30 chatbot mocks ==");
check("chatDig normalises 91-prefix", chatDig("+91 91987 654321")==="9876543210"||chatDig("919876543210")==="9876543210");
check("chatDig rejects short input", chatDig("12345")==="");
// mobile → name / case ID / CRM repair status / product from service sheet
var mobQ2=chatQuery("919876543210");
var mobFlat=JSON.stringify(mobQ2.rows);
check("mobile lookup across sheets", mobQ2.rows.length>0&&mobFlat.indexOf("Ravi")>-1, JSON.stringify(mobQ2).slice(0,140));
check("mobile lookup gives case+crm status+product", mobFlat.indexOf("C1001")>-1
  &&mobFlat.indexOf("Original Display Used")>-1&&mobFlat.indexOf("iPhone 13")>-1,
  JSON.stringify(mobQ2.rows).slice(0,220));
check("mobile 10-digit query", chatQuery("9791994639").title.indexOf("Mobile")>-1
  ||chatQuery("9791994639").title.indexOf("No records")>-1);
// case ID → direct service!CaseId fetch
var caseQ=chatQuery("C1001");
var caseFlat=JSON.stringify(caseQ.rows);
check("case-id query hits service row", caseQ.title.indexOf("C1001")>-1
  &&caseFlat.indexOf("Ravi")>-1&&caseFlat.indexOf("Original Display Used")>-1
  &&caseFlat.indexOf("Repair Complete")>-1, JSON.stringify(caseQ).slice(0,200));
check("case-id query inside sentence", chatQuery("what is status of C1002 please").title.indexOf("C1002")>-1);
// Task 35: everything else → universal cross-sheet exact search
var off1=chatQuery("sale date from 23.09.26");
check("universal search returns what exists, no similar-fallback", off1.title.indexOf("No matches for")===0
  ||off1.title.indexOf("🔍")>-1, off1.title);
check("universal search no-guidance", off1.title.indexOf("Give me a mobile number")===-1);
var subhQ=chatQuery("subha");
check("universal name search finds rows", subhQ.title.indexOf("🔍")>-1, subhQ.title);
check("universal name search groups by sheet", JSON.stringify(subhQ.rows).indexOf("📄")>-1);
check("universal part search exact-only", chatQuery("part number for care+").title.indexOf("Give me a mobile number")===-1);
var mobJSON=serverQuery("919876543210");
check("serverQuery returns JSON", mobJSON.charAt(0)==="{");
var mobParsed=JSON.parse(mobJSON);
check("serverQuery shape", mobParsed.title!==undefined&&Array.isArray(mobParsed.rows));
var blankQ=JSON.parse(serverQuery(""));
check("serverQuery empty guard", blankQ.title==="Say something first");

console.log("== TASK 33 batch ==");
// --- hardcoded IFS incentive ladders ---
check("service IFS 95.2% pays 11000-pro-rata", Math.round(serviceIncentiveForPct33_(95.2))===10368,
  serviceIncentiveForPct33_(95.2));
check("service IFS 105% still 11000-ladder", Math.round(serviceIncentiveForPct33_(105))===Math.round(11000*1.05/1.01));
check("service IFS 109.5% pays 10000-pro-rata", Math.round(serviceIncentiveForPct33_(109.5))===Math.round(10000*1.095));
check("service IFS 110% pays 10000-floor edge", Math.round(serviceIncentiveForPct33_(110))===Math.round(10000*1.1));
check("service IFS 111% pays 12000-pro-rata", Math.round(serviceIncentiveForPct33_(111))===Math.round(12000*1.11/1.10));
check("service IFS upside grows past 110", serviceIncentiveForPct33_(135)>serviceIncentiveForPct33_(120));
check("service IFS 50% no floor cut on ladder", Math.round(serviceIncentiveForPct33_(50))===Math.round(11000*0.5/1.01));
check("accessory IFS 95.2% lands 6000-pro-rata", Math.round(accessoryIncentiveForPct33_(95.2))===5712,
  accessoryIncentiveForPct33_(95.2));
check("accessory IFS 80% 4000-pro-rata", Math.round(accessoryIncentiveForPct33_(80))===Math.round(4000*0.8/0.8));
check("accessory IFS 65% 3000-pro-rata", Math.round(accessoryIncentiveForPct33_(65))===Math.round(3000*0.65/0.7));
check("accessory IFS 55% 2500-pro-rata", Math.round(accessoryIncentiveForPct33_(55))===Math.round(2500*0.55/0.6));
check("accessory IFS 50% none", accessoryIncentiveForPct33_(50)===0);
check("accessory IFS 111% 8000-pro-rata", Math.round(accessoryIncentiveForPct33_(111))===Math.round(8000*1.11/1.1));
var gsSrc33=fs.readFileSync(__dirname+"/../../SalesDashboard.gs","utf8");
check("incentive helpers hardcoded, not sheet-driven", gsSrc33.indexOf("svcInc = serviceIncentiveForPct33_")>-1
  &&gsSrc33.indexOf("accInc = accessoryIncentiveForPct33_")>-1);
check("person card incentive footer no longer cites A18:D25", page.indexOf("target!A18:D25")===-1);
// --- chatbot: half-page results + per-case diagnostic block ---
check("chat panel half-width css", page.indexOf(".chat-panel.half { width:50vw")>-1, "chat-panel.half css");
check("chat panel widens only on mobile/case answers", page.indexOf("classList.toggle('half', !!res.wide)")>-1);
check("mobile + case answers flagged wide", mobQ2.wide===true
  &&chatQuery("C1001").wide===true&&chatQuery("919876543299").wide===true);
check("universal hits render wide, misses compact (Task 35)", subhQ.wide===true
  &&(off1.wide===undefined||off1.wide===false));
check("universal search note pinned to exact matches",
  subhQ.note&&subhQ.note.indexOf("Exact matches only")>-1);

console.log("== TASK 34 batch ==");
var c1001j=j.jobs.find(function(o){return o.caseId==="C1001";});
check("C1001 carries gsx id + serial from service sheet", c1001j.gsxId==="GSX6001"
  &&c1001j.serialNo==="SN-C1001", [c1001j.gsxId,c1001j.serialNo]);
check("case collapse keeps latest row's gsx id", j.jobs.find(function(o){return o.caseId==="C1003";}).gsxId==="GSX6003",
  j.jobs.find(function(o){return o.caseId==="C1003";}).gsxId);
check("gsx id + serial cells sit between number and model",
  page.indexOf('<td class="num">919876543210</td><td>GSX6001</td><td>SN-C1001</td><td>iPhone 13</td>')>-1);
check("Communicated cells gone from rows", page.indexOf('class="comm-yes"')===-1
  &&page.indexOf('class="comm-no"')===-1);
check("map line shows gsx id + serial sources", page.indexOf("GSX ID ← <b>GSXCode</b>")>-1
  &&page.indexOf("Device S.No ← <b>SerialNo1</b>")>-1);
check("bucket accent css defined for all 4 owner buckets",
  page.indexOf(".grphead.bk-estimate td { background:linear-gradient(180deg,#f6a21c")>-1
  &&page.indexOf(".grphead.bk-ready td { background:linear-gradient(180deg,#2fd475")>-1
  &&page.indexOf(".grphead.bk-process td { background:linear-gradient(180deg,#22c2e3")>-1
  &&page.indexOf(".grphead.bk-unclaimed td { background:linear-gradient(180deg,#f05858")>-1);
check("unclaimed group band coloured", page.indexOf('class="grphead bk-unclaimed"')>-1);
check("ready-for-pickup group band coloured", page.indexOf('class="grphead bk-ready"')>-1);
check("row bucket cells carry the group accent",
  page.indexOf('bucket-cell bk-unclaimed')>-1&&page.indexOf('bucket-cell bk-ready')>-1);
check("bucket-cell accent edge styling (Task 35)", page.indexOf(".jobtable td.bucket-cell.bk-unclaimed")>-1
  &&page.indexOf("box-shadow:inset 3px 0 0 0 #f05858")>-1
  // whole data rows carry a light tint of the bucket colour
  &&page.indexOf(".jobtable tr.bk-unclaimed td { background:rgba(240,88,88,.16) !important")>-1
  &&page.indexOf(".jobtable tr.bk-estimate td { background:rgba(246,162,28,.16) !important")>-1
  &&page.indexOf(".jobtable tr.bk-process td { background:rgba(34,194,227,.15) !important")>-1
  &&page.indexOf(".jobtable tr.bk-ready td { background:rgba(47,212,117,.15) !important")>-1);
check("mobile case block fields", mobFlat.indexOf("CRM case status")>-1
  &&mobFlat.indexOf("GSX case status")>-1&&mobFlat.indexOf("CRM bucket")>-1
  &&mobFlat.indexOf("Case date")>-1&&mobFlat.indexOf("Comm updated")>-1,
  mobFlat.slice(0,300));
var d1=new Date(Date.now()-86400000);function p2(x){return(x<10?"0":"")+x;}
var c1001Comm=p2(d1.getDate())+"-"+p2(d1.getMonth()+1)+"-"+d1.getFullYear();
check("mobile case block carries case + customer", mobFlat.indexOf("C1001")>-1&&mobFlat.indexOf("Ravi")>-1);
check("comm updated = latest for C1001 (1 day ago)", mobFlat.indexOf(c1001Comm)>-1, [c1001Comm,mobFlat.slice(0,400)]);
check("chatCommDateForCase_ helper exists", typeof chatCommDateForCase_==="function"
  &&chatCommDateForCase_(getSpreadsheet_(),"C1009")==="");
check("case query adds comm-updated row", caseFlat.indexOf("Comm updated")>-1
  &&caseFlat.indexOf(c1001Comm)>-1, caseFlat.slice(0,300));
check("case query for C1002 comm = 4 days ago", chatQuery("C1002").rows.some(function(r){
    var d4=new Date(Date.now()-4*86400000);
    var k=p2(d4.getDate())+"-"+p2(d4.getMonth()+1)+"-"+d4.getFullYear();
    return JSON.stringify(r).indexOf(k)>-1; })===true);

console.log("== TASK 24 batch ==");
check("tile video golden-frame asset swap", page.indexOf("1064124616/thumb/8.jpg")>-1
  &&page.indexOf("1063973464")===-1
  &&page.indexOf("1064124616/preview/stock-footage-shining-gold-dust-particles")>-1);
check("chat bot renamed ask me anything", page.indexOf("ask me anything")>-1&&page.indexOf("Ask the sheets")===-1);
check("clock has no gif asset", page.indexOf("AC40-01-10-02-R")===-1
  &&page.indexOf("tgsclc821blu_thumbnail.gif")===-1&&page.indexOf("animation.gif")===-1);
check("clock fits a bigger frame (Task 31)", page.indexOf("width:400px; height:96px")>-1
  &&page.indexOf("font-size:1.296rem")>-1);
check("deck tab present", page.indexOf('data-page="deck"')>-1&&page.indexOf('id="pageDeck"')>-1);
check("deck NOT on index tiles", (page.match(/glass-tile[^>]*data-page="deck"/g)||[]).length===0);
console.log("== TASK 36: deck rework (8 tables, current month) ==");
var dk=biz.deckAnalysis;
check("deck month label", dk.monthLabel==="SEP-2026", dk.monthLabel);
check("8 deck sections in owner order",
  page.indexOf('id="deck-sales"')<page.indexOf('id="deck-walkin"')
  &&page.indexOf('id="deck-walkin"')<page.indexOf('id="deck-repair"')
  &&page.indexOf('id="deck-repair"')<page.indexOf('id="deck-accessories"')
  &&page.indexOf('id="deck-accessories"')<page.indexOf('id="deck-service"')
  &&page.indexOf('id="deck-service"')<page.indexOf('id="deck-gp"')
  &&page.indexOf('id="deck-gp"')<page.indexOf('id="deck-walkin-analysis"')
  &&page.indexOf('id="deck-walkin-analysis"')<page.indexOf('id="deck-coverage"'));
check("sales revenue rows RR/AR/SR", dk.salesRev.rows[0].label==="Repair Revenue"&&dk.salesRev.rows[0].count===0
  &&dk.salesRev.rows[1].count===2&&dk.salesRev.rows[1].value===2100
  &&dk.salesRev.rows[2].count===1&&dk.salesRev.rows[2].value===5000
  &&dk.salesRev.totalCount===3&&dk.salesRev.totalValue===7100);
check("walkin details 7/3", dk.walkins.unique===7&&dk.walkins.uniqBills===3);
check("repair analysis iPhone 1 @5000 via orderCode link",
  dk.repairRev.rows[0].label==="iPhone"&&dk.repairRev.rows[0].count===1&&dk.repairRev.rows[0].value===5000, dk.repairRev.rows[0]);
check("repair analysis Others 4 (SNR C1004 out, Jul C1005 out)",
  dk.repairRev.rows[3].label==="Others"&&dk.repairRev.rows[3].count===4&&dk.repairRev.totalCount===5);
check("accessories buckets Case 1/1200, AC+ 1/900",
  dk.accessoriesRev.rows[1].label==="Case"&&dk.accessoriesRev.rows[1].count===1&&dk.accessoriesRev.rows[1].value===1200
  &&dk.accessoriesRev.rows[3].label==="AC+"&&dk.accessoriesRev.rows[3].value===900
  &&dk.accessoriesRev.totalCount===2&&dk.accessoriesRev.totalValue===2100, dk.accessoriesRev.rows);
check("service pivot blank item code", dk.serviceRev.rows.length===1
  &&dk.serviceRev.rows[0].label==="(blank item code)"&&dk.serviceRev.rows[0].value===5000
  &&dk.serviceRev.totalValue===5000);
check("gp value rows + GP%", dk.gpValue.rows[0].label==="Accessories Revenue"
  &&dk.gpValue.rows[0].pct===26.2&&dk.gpValue.rows[2].pct===30.0
  &&dk.gpValue.total.value===7100&&dk.gpValue.total.gp===2050&&dk.gpValue.total.pct===28.9, dk.gpValue);
var waVals=dk.walkinAnalysis.map(function(r){return r.value;});
check("walkin analysis 7,3,3,2,0,1,43%,4",
  waVals[0]===7&&waVals[1]===3&&waVals[2]===3&&waVals[3]===2&&waVals[4]===0&&waVals[5]===1
  &&waVals[6]==="43%"&&waVals[7]===4, waVals);
check("coverage iPhone Battery W1 (part=Batt), Others W2/NW3",
  dk.coverage.rows[1].label==="iPhone Battery"&&dk.coverage.rows[1].count===1&&dk.coverage.rows[1].w===1
  &&dk.coverage.rows[4].count===5&&dk.coverage.rows[4].w===2&&dk.coverage.rows[4].nw===3
  &&dk.coverage.total.count===6&&dk.coverage.total.w===3&&dk.coverage.total.nw===3, dk.coverage);
check("deck renders key numbers", page.indexOf("Walkin to Bills %")>-1&&page.indexOf("iPhone Battery")>-1
  &&page.indexOf("(blank item code)")>-1);
check("index tiles: decline 2 + winback 2 live",
  page.indexOf("Declined repairs</span><b>2</b>")>-1&&page.indexOf("Win-backs</span><b>2</b>")>-1);
check("comm contrast lifted vs bucket tints (Task 36-C)",
  page.indexOf(".jobtable tbody tr td.comm-neg")>-1&&page.indexOf("#ffd3d7")>-1&&page.indexOf("#bbf0d0")>-1);
check("csat right chips compact (Task 36-B)",
  page.indexOf(".csat-right .csat-chip { flex:1 1 0; padding:6px 10px")>-1
  &&page.indexOf(".csat-right .cc-value { font-size:1.05rem")>-1);
var gsSrc=fs.readFileSync(__dirname+"/../../SalesDashboard.gs","utf8");
check("refresh includes deck", page.indexOf("deck:'pageDeck'")>-1&&gsSrc.indexOf("deck: renderDeckHtml_(")>-1);

console.log("== TASK 25 batch ==");
var NAVY="rgba(28,43,82,.98)";
check("clock: frameless, no card bg (Task 35)", page.indexOf(".topbar-clock")>-1
  &&page.indexOf("background:rgba(8,14,36,.92)")===-1
  &&page.indexOf(".topbar-clock { position:relative")>-1
  &&page.indexOf("background:transparent")>-1
  &&page.indexOf("brgprecision")===-1);
check("tables all navy", page.indexOf("table { background:"+NAVY)>-1
  &&page.indexOf("rgba(21,27,68,.84)")===-1);
check("index tiles navy", page.indexOf(".glass-tile { background:"+NAVY)>-1);
check("ph-tile css gone (Task 35 live cards replaced placeholders)", page.indexOf("ph-tile")===-1);
check("gpie-wrap + fresh chips navy", page.indexOf("#pageIndex .gpie-wrap { background:"+NAVY)>-1
  &&page.indexOf("#pageIndex .fresh-chip { background:"+NAVY)>-1);
check("chip text readable on navy", page.indexOf(".csat-chip .cc-label { font-size:11.5px; font-weight:700; color:#9fb4e8")>-1
  &&page.indexOf(".csat-chip.pos .cc-value { color:#4fe0a8")>-1
  &&page.indexOf(".csat-chip.neg .cc-value { color:#ff8fa1")>-1
  &&page.indexOf("#dde3ee")===-1);

console.log("== TASK 26 batch ==");
check("no background image assets left", page.indexOf("2c7ccfba")===-1
  &&page.indexOf("ideas-action-innovation")===-1&&page.indexOf("783dbbe3")===-1);
check("all tiles share one flat navy (Task 30)", (page.match(/rgba\(23,49,84/g)||[]).length===0
  &&(page.match(/rgba\(28,53,99/g)||[]).length===0
  &&(page.match(/linear-gradient\(145deg,rgba\(28/g)||[]).length===0
  &&(page.match(/rgba\(28,43,82,.98\)/g)||[]).length>=8);
check("glowing golden-yellow clock digits (Task 35)", page.indexOf("color:#ffd34d")>-1
  &&page.indexOf("rgba(255,196,46,.95)")>-1
  &&page.indexOf("color:#b7dcff")===-1&&page.indexOf("#ff5c66")===-1);
check("wp-right panel same flat navy", page.indexOf(".wp-right { background:rgba(28,43,82,.98)")>-1);

console.log("== TASK 27 batch ==");
check("background is flat lighter-ash everywhere (Task 31)", page.indexOf("body { background:#3f444d; color:#dbe4f6; }")>-1
  &&page.indexOf("2c7ccfba6490b77f6688ae0d429d3a3a")===-1);
console.log("== TASK 29 batch ==");
// client JS must PARSE — a syntax error here silently kills every onclick.
var m=page.match(/<script>([\s\S]*?)<\/script>/);
var clientJS=m?m[1]:"";
var jsErr=null; try{ new Function(clientJS); }catch(e){ jsErr=e.message; }
check("client script block parses (clicks work)", jsErr===null, jsErr||"");
check("all click handlers defined", clientJS.indexOf("function switchPage(")>-1
  &&clientJS.indexOf("function chatSend(")>-1&&clientJS.indexOf("function chatToggle(")>-1);
check("WhatsApp integration removed", page.indexOf("waSendLast")===-1
  &&page.indexOf("#25d366")===-1&&page.indexOf("wa_chatid")===-1);

console.log("== TASK 30 batch ==");
check("jobs table: CRM Status dropped, Repair Type still after GSX (Task 38)",
  page.indexOf("<th>GSX Repair Status</th><th>Repair Type</th><th>Case ID</th>")>-1);
check("repair bucket still carries the crm status", page.indexOf(">Original Display Used</td>")>-1);

console.log("== TASK 38: case popup + single-line rows + deck fallback ==");
check("case ids clickable (caselink)", page.indexOf('class="caselink"')>-1
  &&page.indexOf('onclick="openCaseDetails')>-1);
check("case modal markup + data island", page.indexOf('id="caseModal"')>-1
  &&page.indexOf('id="caseDetailData"')>-1&&page.indexOf('id="cmTable"')>-1);
check("popup client handlers", clientJS.indexOf("function openCaseDetails(")>-1
  &&clientJS.indexOf("function closeCaseDetails(")>-1&&clientJS.indexOf("caseDetailsMap_")>-1);
check("popup carries every requested field", page.indexOf('"readyDate"')>-1
  &&page.indexOf('"repairDate"')>-1&&page.indexOf('"bucket"')>-1&&page.indexOf('"gsxStatus"')>-1);
check("ready pickup date parsed onto jobs", j.jobs.filter(function(o){return o.readyDate;}).length===5);
var c1001d=j.jobs.find(function(o){return o.caseId==="C1001";});
check("C1001 popup snapshot complete", c1001d.readyDate.length>0&&c1001d.repairDate.length>0
  &&c1001d.crmRepairStatus==="Original Display Used"&&c1001d.gsxStatus==="Repair Complete");
var gsSrc2=gsSrc;
console.log("== TASK 38: deck stale-month fallback ==");
// Push every service-case date back to July (= older than the frozen month),
// then the case-driven deck tables must still produce numbers via fallback.
var svStale=sv.map(function(row,i){ if(i===0)return row;
  return row.map(function(v){ return typeof v==="string"?v.replace(/2026-09-/g,"2026-07-"):v; }); });
var staleSheets=[mockSheet("target",T),mockSheet("walkin",wl),mockSheet("service",svStale)];
var ssStale={getSheets:function(){return staleSheets;},
  getSheetByName:function(n){for(var i=0;i<staleSheets.length;i++)if(staleSheets[i].getName()===n)return staleSheets[i];return null;}};
var dkStale=computeDeckAnalysis_(ssStale,parseCrmRows(crm));
check("stale service sheet: repair table still valued via fallback",
  dkStale.repairRev.totalCount===6&&dkStale.repairRev.rows[0].count===1
  &&dkStale.repairRev.rows[0].value===5000&&dkStale.repairRev.rows[3].count===5,dkStale.repairRev);
check("stale service sheet: coverage still valued via fallback",
  dkStale.coverage.total.count===7&&dkStale.coverage.rows[1].count===1
  &&dkStale.coverage.rows[4].w===3)
check("server-side scan/date/part/fuzzy dead code removed",
  gsSrc2.indexOf("function chatAllSheetsScan_")===-1
  &&gsSrc2.indexOf("function chatSalesQuery")===-1
  &&gsSrc2.indexOf("function chatPartQuery(")===-1
  &&gsSrc2.indexOf("function chatFuzzyQuery(")===-1
  &&gsSrc2.indexOf("function chatExtractDates")===-1);
check("brighter index subtitles", page.indexOf("color:#c9d4f5")>-1
  &&page.indexOf("color:#e3eafd")>-1&&page.indexOf(".gpie-leg-val { color:#ffffff")>-1);

console.log("== TASK 35 batch ==");
// --- item 3: every sheet is stamped with the refresh time ---
(function(){
  var ns=biz.nsUsb?biz.nsUsb():null;
})();
var sheetNames=ALL_SHEETS.map(function(x){return x.getName();});
var stampedCount=0, allStamped=true;
ALL_SHEETS.forEach(function(sh){
  var w=sh.__writes||[];
  var hit=w.some(function(x){return typeof x.value==="string"&&x.value.indexOf("Dashboard refresh")>-1;});
  if(hit)stampedCount++;
  else allStamped=false;
});
check("refresh stamp written into EVERY sheet (Task 35)", allStamped&&stampedCount===8,
  stampedCount+"/8:"+sheetNames.join(","));
check("refresh stamp text surfaced in freshness line", page.indexOf("fresh-stamp")>-1
  &&page.indexOf("updated · all sheets")===-1, "");
var pdChk=loadPerformanceData_();
// --- items 8/9: decline + quote counts from the mock sheets ---
check("opsExtras decline count = unique 2 (dedupe)", pdChk.opsExtras.decline.count===2,
  pdChk.opsExtras.decline.count);
check("opsExtras winback count = unique 2 (dedupe)", pdChk.opsExtras.winback.count===2,
  pdChk.opsExtras.winback.count);
check("opsExtras decline sheetName surfaced", pdChk.opsExtras.decline.sheetName==="decline");
// --- item 12: universal chat search ---
var uni1=chatQuery("iPhone 13");
check("universal chat exact match across sheets", uni1&&uni1.title.indexOf("row")>-1,
  uni1&&uni1.title);
check("universal chat groups by sheet tab", JSON.stringify(uni1.rows).indexOf("📄")>-1);
check("universal chat flags wide results", uni1.wide===true);
var uniMiss=chatQuery("zzzqqq nothere");
check("universal chat miss → No matches, no fallback", uniMiss.title.indexOf("No matches")===0,
  uniMiss.title);
check("universal chat miss stays compact", uniMiss.wide!==true);
var uniMob=chatQuery("919876543210");
check("mobile fast-path still first", uniMob.title.indexOf("Mobile")>-1||uniMob.title.indexOf("mobile")>-1
  ||uniMob.title.indexOf("C1001")>-1, uniMob.title);
var uniCase=chatQuery("C1001");
check("case fast-path still first", uniCase.title.indexOf("C1001")>-1, uniCase.title);
var uni2=chatQuery("Subha");
check("universal chat multi-token/name search works", uni2.title.indexOf("🔍")>-1, uni2.title);
check("universal chat note pinned 'Exact matches only'", JSON.stringify(uni2).indexOf("Exact matches")>-1);
console.log(fails?("\n"+fails+" FAILURES"):"\nALL CHECKS PASS");
process.exit(fails?1:0);
