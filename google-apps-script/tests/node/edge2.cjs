/**
 * Edge-case harness for SalesDashboard.gs. Run: node google-apps-script/tests/node/edge2.cjs
 */
global.Utilities={formatDate:function(d,tz,f){function p(n){return(n<10?"0":"")+n;}
  return p(d.getDate())+"-"+p(d.getMonth()+1)+"-"+d.getFullYear();}};
global.Session={getScriptTimeZone:function(){return "Asia/Calcutta";}};
global.CacheService={getScriptCache:function(){return{get:function(){return null},put:function(){}}}};
global.HtmlService={createHtmlOutput:function(){return{addMetaTag:function(){return this},setTitle:function(){return this}}}};
function mockSheet(name,values){return{getName:function(){return name},getDataRange:function(){return{
  getValues:function(){return values},
  getDisplayValues:function(){return values.map(function(r){return r.map(function(v){return(v==null||v instanceof Date)?"":String(v);})});}};}};}
var ALL=[];
var SS={getSheets:function(){return ALL;},getSheetByName:function(n){for(var i=0;i<ALL.length;i++)if(ALL[i].getName()===n)return ALL[i];return null;}};
global.SpreadsheetApp={getActiveSpreadsheet:function(){return SS;}};
var fs=require("fs");
eval(fs.readFileSync(__dirname+"/../../SalesDashboard.gs","utf8"));
getSpreadsheet_=function(){return SS;};
function daysAgoStr(n,h){var d=new Date(Date.now()-n*86400000);function p(x){return(x<10?"0":"")+x;}
  return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate())+" "+(h||"10")+":00:00";}
var fails=0;
function check(n,c,x){if(c)console.log("  PASS",n);else{fails++;console.log("  FAIL",n,JSON.stringify(x));}}

// --- scenario 1: over-achieved salesperson -> daily target clamps to 0; negative review diff; approx unit 0 ---
var T1=[]; for(var i=0;i<20;i++)T1.push(new Array(22).fill(""));
T1[0]=["Revenue Type","Store Target","SUBHA"];
T1[1]=["Service",100000,50000]; T1[2]=["Accessories",10000,10000]; T1[3]=["Repair",20000,10000];
T1[8]=["Prev Month","This Month","Difference"]; T1[9]=[40,35,-5];
T1[12]=["","AR"]; T1[13]=["","Care+"]; T1[14]=["",10]; T1[15]=["",0];
var crm1=[["InvoiceCode","InvoiceDate","SalesPersonName","RevenueType","ItemDescription","SubCategory","SchemeCode","TaxableAmount","GrossProfit","TranType"],
 ["A1","23-09-2026 10:00","Subha","Service","x","","","60000","2000","SALES"],
 ["A2","23-09-2026 11:00","Subha","Repair","y","","","6000","500","SALES"],
 ["A3","23-09-2026 12:00","Subha","Accessories","z","","CPLUS","9000","800","SALES"]];
var sv1=[["CaseId","FirstName","LastName","MobileNo","CaseDate","JobAge","Status","GSXRepairStatusDescription","AvailableBalance"],
 ["C1","A","B","9000000001",daysAgoStr(3),"3","Open","Repair In Progress","1200"],
 ["C2","C","D","9000000002",daysAgoStr(2),"2","Open","Diagnosis In Progress","500"]];
var cm1=[["CaseId","CreatedDate"],["C1",daysAgoStr(3,"9")]]; // exactly >2 days old => red
ALL=[mockSheet("target",T1),mockSheet("crm",crm1),mockSheet("service",sv1),mockSheet("comm",cm1)];
console.log("== edge: balance clamps / zero approx unit / no advance tab ==");
var biz=loadDashboardData_().data;
var sub=biz.salespeople.find(function(p){return p.name==="Subha"});
check("svc daily clamps to 0 when target achieved (60k>50k)",sub.dailyTargets.service===0,sub.dailyTargets.service);
check("repair daily still positive (bal 4000)",sub.dailyTargets.repair>0,sub.dailyTargets.repair);
var cp=biz.kpi.ar.filter(function(k){return k.code==="CAREPLUS";})[0];
check("approx unit 0 -> no approx",cp.approxAmount===0);
var j=loadPerformanceData_().jobs;
check("no advance tab -> direct column fallback total 1700",j.totals.advanceTotal===1700,j.totals.advanceTotal);
check("flag advanceViaColumn set",j.advanceViaColumn===true,j.advanceViaColumn);
var page=buildDashboardPage_(biz);
check("negative review diff tile",page.indexOf('class="value neg">-5</div>')>-1);
check("review caption negative",page.indexOf("35 this month")>-1&&page.indexOf("previous 40")>-1);
check("approx cell em-dash when unit 0",page.indexOf(">—<")>-1);
check("avg invoice value tile is gone",page.indexOf("Avg Invoice Value")===-1);
check("comm C1 red (older than 2 days)",page.indexOf("comm-neg")>-1);
check("no process/perf regression",page.indexOf("Process")>-1);

// --- scenario 2: comm within 2 days => green; review tile pos clap; no review row ---
var cm2=[["CaseId","CreatedDate"],["C1",daysAgoStr(1,"9")]];
ALL=[mockSheet("target",T1),mockSheet("crm",crm1),mockSheet("service",sv1),mockSheet("comm",cm2)];
var j2=loadPerformanceData_().jobs;
var c1=j2.jobs.find(function(o){return o.caseId==="C1";});
console.log("== edge: comm green window / flags ==");
check("C1 comm within 2 days -> not stale (green)",c1&&c1.staleComm===false,c1&&c1.staleComm);
var page2=buildDashboardPage_(loadDashboardData_().data);
check("comm-pos present",page2.indexOf("comm-pos")>-1);

console.log(fails?("\n"+fails+" FAILURES"):"\nALL EDGE CHECKS PASS");
process.exit(fails?1:0);
