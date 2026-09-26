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
var T=[]; for(var i=0;i<20;i++) T.push(new Array(14).fill(""));
T[0]=["Revenue Type","Store Target","SUBHA","DINESH"];
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
 "RepairCreatedDate","JobAge","Status","DeliveredDate","GSXRepairStatusDescription","CRM_JobStatusDesc"],
 ["C1001","Ravi","Kumar","919876543210","iPhone 13",daysAgoStr(14,"10"),daysAgoStr(13,"11"),14,"Open","","Repair Complete","Original Display Used"],
 ["C1002","Anita","Devi","919876543211","Samsung S23",daysAgoStr(4,"09"),daysAgoStr(3,"14"),4,"Open","","Unclaimed Device","Part Used"],
 ["C1003","Mani","Kandan","919876543212","OnePlus 11",daysAgoStr(2,"10"),"",2,"Open","","Diagnosis In Progress","Device Unclaimed Device due"],
 ["C1004","Peela","Dass","919876543217","Realme 12",daysAgoStr(1,"08"),"",1,"Open","","Repair In Progress",""],
 ["C1005","Old","Closed","919876543218","Pixel 7",daysAgoStr(55,"09"),daysAgoStr(54,"10"),55,"Delivered",daysAgoStr(50,"10"),"Repair Complete","Part Used"],
 ["C1006","Age","Fallback","919876543219","Moto G",daysAgoStr(20,"09"),"","","Open","","Waiting for Part",""],
 ["C1007","Can","Cell","919876543220","Nokia",daysAgoStr(2,"09"),"",2,"Cancelled","","Aborted",""]];

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
check("open jobs 5",j.totals.openJobs===5,j.totals.openJobs);
check("non-open excluded 2",j.totals.closedExcluded===2,j.totals.closedExcluded);
check("C1006 fallback age ~20d", (function(){var c=j.jobs.find(function(o){return o.caseId==="C1006";});return c&&Math.abs(c.jobAge-20)<=1;})());
check("unclaimed 2",j.totals.unclaimed===2);
check("advance via advance-sheet total 2049",j.totals.advanceTotal===2049,j.totals.advanceTotal);

console.log("== HTML ==");
var page=buildDashboardPage_(biz);
check("KPI heeft Approx Value header",page.indexOf("Approx Value")>-1);
check("approx achieved ₹400",page.indexOf("₹400")>-1);
check("approx-to-target guidance ₹9,600",page.indexOf("₹9,600 to target")>-1);
check("review tile big = gain +19",page.indexOf('class="value pos">+19</div>')>-1);
check("review caption small",page.indexOf("1494 this month")>-1&&page.indexOf("previous 1475")>-1);
check("avg tile 3 details",page.indexOf("gross profit ÷")>-1);
check("avg GP 683",page.indexOf("₹683")>-1);
check("invoice ÷ total invoices label",page.indexOf("invoice value ÷ 3 total invoices")>-1);
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
check("Process tab",page.indexOf("Process")>-1&&page.indexOf("📈 Performance")===-1);
check("walkin hsl tint",page.indexOf("background:hsl(")>-1);
check("token table",page.indexOf("% of creation")>-1&&page.indexOf("Total tokens generated")>-1);
check("ready col gone",page.indexOf("Ready for pickup")===-1);
check("comm pos/neg present",page.indexOf("comm-pos")>-1&&page.indexOf("comm-neg")>-1);
check("business sections intact",page.indexOf("Invoice Valuation")>-1&&page.indexOf("Date-wise")>-1);

console.log(fails?("\n"+fails+" FAILURES"):"\nALL CHECKS PASS");
process.exit(fails?1:0);
