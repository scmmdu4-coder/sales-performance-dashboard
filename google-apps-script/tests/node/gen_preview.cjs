// Preview generator: renders the real dashboard HTML with mock data.
global.Utilities={formatDate:function(d,tz,f){function p(n){return(n<10?"0":"")+n;}
  return p(d.getDate())+"-"+p(d.getMonth()+1)+"-"+d.getFullYear();}};
global.Session={getScriptTimeZone:function(){return "Asia/Calcutta";}};
global.CacheService={getScriptCache:function(){return{get:function(){return null},put:function(){}}}};
global.HtmlService={createHtmlOutput:function(){return{addMetaTag:function(){return this},setTitle:function(){return this}}}};
function mockSheet(name,values){var sh={getName:function(){return name},
  getLastColumn:function(){return values[0]?values[0].length:1;},
  getLastRow:function(){return values.length;},
  getRange:function(r,c,nr,nc){return{
    getValue:function(){return values[r-1]&&values[r-1][c-1]||""},
    getValues:function(){var out=[];for(var i=0;i<(nr||1);i++){var row=[];
      for(var j=0;j<(nc||1);j++){row.push((values[r-1+i]&&values[r-1+i][c-1+j])||"");}
      out.push(row);}return out;},
    setValue:function(v){if(!values[r-1])values[r-1]=[];values[r-1][c-1]=v;}
  };
  },
  getDataRange:function(){return{
  getValues:function(){return values},
  getDisplayValues:function(){return values.map(function(r){return r.map(function(v){return(v==null||v instanceof Date)?"":String(v);})});}};}};
  return sh;}
var ALL=[];
var SS={getSheets:function(){return ALL;},getSheetByName:function(n){for(var i=0;i<ALL.length;i++)if(ALL[i].getName()===n)return ALL[i];return null;}};
global.SpreadsheetApp={getActiveSpreadsheet:function(){return SS;}};
var fs=require("fs"),path=require("path");
eval(fs.readFileSync("/home/user/sales-performance-dashboard/google-apps-script/SalesDashboard.gs","utf8"));
getSpreadsheet_=function(){return SS;};
function daysAgoStr(n,h){var d=new Date(Date.now()-n*86400000);function p(x){return(x<10?"0":"")+x;}
  return d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate())+" "+(h||"10")+":00:00";}
var CODES=["CPLUS","PROT","AMCPLAN","BND","DISP","BATT","GLS"];
var NAMES=["Ravi","Anita","Mani","Peela","Kumar","Divya","Suresh","Rekha","Arjun","Meena","Vijay","Latha","Karthik","Nisha","Deepak","Sona"];
var LAST=["Kumar","Devi","Kandan","Dass"];
var GSX=["Diagnosis In Progress","Repair In Progress","Repair Complete","Waiting for Part","Quality Check"];
var CRMS=["Part Used","Estimation Given","Original Display Used","Device Unclaimed Device due","Repair In Process"];
var RTYPES=["Walk-In Repair","Carry-In Repair","Service Non-Repair","Screen Replacement"];
var WARR=["In Warranty","Out Of Warranty (No Coverage)","Consumer Warranty","AppleCare Coverage"];
// Task 40: coverage v2 reads service!IssueDescription + service!ProductWarrantyDesc.
var ISSUES=["Weak battery backup","Cracked display glass","Screen flickers","","Charging port fault","Kernel panic",""];
function PW(v){return (/out|applecare/i.test(v))?"Out Of Warranty (No Coverage)":"Product Warranty";}
var SVC=[];
for(var i=0;i<26;i++){var age=(i%28)+1;
  // every third case is already ready for pickup (2-6 day turnaround)
  var ready=(i%3===0)?daysAgoStr(Math.max(0,age-(1+(i%5))),"11"):"";
  var rt=RTYPES[i%4];
  var cms=rt==="Service Non-Repair"?"Ready For Pickup":CRMS[i%5];
  // every 5th case has no repair-created date yet (rule-b highlight when older than 2 days)
  var repCreated=(i%5===0)?"":daysAgoStr(Math.max(0,age-1),"12");
  // 16 fields, aligned to the 16-column header below.
  SVC.push(["C100"+i,NAMES[i%16],LAST[i%4],"9198765432"+("000000"+i).slice(-3),"iPhone "+(10+i%7),daysAgoStr(age,"10"),repCreated,age,"Open","",GSX[i%5],cms,(i%3===0?200:(i*50)%900),WARR[i%4],rt,ready,"GSX70"+("00"+i).slice(-3),"SN-A"+("00000000"+i).slice(-8),ISSUES[i%7],PW(WARR[i%4])]);}
// a few older cases across the previous months for the 3-month comparison
for(var mo=1;mo<=3;mo++){for(var j=0;j<3+mo;j++){var a=30*mo+(j*6);
  SVC.push(["C5"+mo+""+j,"Old"+j,LAST[j%4],"919876640000"+mo+j,"Pixel 7",daysAgoStr(a,"10"),a,a,"Delivered",daysAgoStr(a-4,"10"),"Repair Complete","Part Used",0,WARR[j%4],RTYPES[(j+mo)%4],daysAgoStr(a-4,"10"),"GSX55"+mo+j,"SN-OLD"+mo+j,"",PW(WARR[j%4])]);}}
// one duplicated case row (older case date) — the page must merge it away
SVC.unshift(["C1005","Kumar","Devi","919876543205","Different Phone",daysAgoStr(9,"10"),9,9,"Open","",GSX[0],"Part Used",0,"In Warranty","Walk-In Repair","","GSX70-OLD","SN-OLD","","","Product Warranty"]);

SVC.push(["C4001","Old","Closed","919876640001","Pixel 7",daysAgoStr(55,"10"),54,55,"Delivered",daysAgoStr(50,"10"),"Repair Complete","Part Used",0,"","","","","",""]);
var sv=[["CaseId","FirstName","LastName","MobileNo","ProductDescription","CaseDate","RepairCreatedDate","JobAge","Status","DeliveredDate","GSXRepairStatusDescription","CRM_JobStatusDesc","AvailableBalance","CoverageCodeDescription","RepairDocType","Ready_For_Pickup_Date_CRM","GSXCode","SerialNo1","IssueDescription","ProductWarrantyDesc"]].concat(SVC);
var RND=mulberry(42);function mulberry(a){return function(){a|=0;a=a+0x6D2B79F5|0;var t=Math.imul(a^a>>>15,1|a);t=t+Math.imul(t^t>>>7,61|t)^t;return((t^t>>>14)>>>0)/4294967296;}}
var kpiProb=[4,2,1,0,2,1,3];
var rev=["Service","Accessories","Repair"];
var crm=[["InvoiceCode","InvoiceDate","SalesPersonName","RevenueType","ItemDescription","SubCategory","SchemeCode","TaxableAmount","GrossProfit","TranType","ItemCode","PartType","MobileNo","DiscountAmount"]];
for(var day=1;day<=25;day++){if(day%7===0)continue;for(var x=0;x<18;x++){var sp=RND()<0.55?"Subha":"Dinesh";var rt=rev[Math.floor(RND()*3)];var amt=Math.round(500+RND()*4500);var k=RND()<0.2?CODES[Math.floor(RND()*kpiProb.length)]:"";
 crm.push(["IV"+day+x,("0"+day).slice(-2)+"-09-2026 "+("0"+(9+Math.floor(RND()*9))).slice(-2)+":30",sp,rt,k?"Device plan":"",k," ",amt,Math.round(amt*0.25),"SALES","PART",k,"850000"+day+x,(RND()<0.18?Math.round(amt*(0.04+0.10*RND())):0)]);}}
// Task 43: crm!F2 carries the store name, shown verbatim as the page brand.
crm[1][5]="iPlanet Care - Sulur";
var wl=[["MobileNo","TokenDate","VisitPurposeDesc","CaseId","InvoiceCode","Remark","TokenCreatedBy","CustomerVisitSource"]];
for(var i2=0;i2<180;i2++){var dd=(i2%25)+1;if(dd%7===0)dd=1;var m="91999"+String(100000+i2).slice(-6);var pur=["Repair Submission","Collection","General/Sales Enquiry"][i2%3];var hasCase=i2%4!==0;wl.push([m,"2026-09-"+("0"+dd).slice(-2)+" 10:00:00",pur,hasCase?"C100"+(i2%26):"",i2%2?"IV"+(i2%25)+"0":"",i2%6===0?"Follow up":"",i2%3?"Subha":"Dinesh",["Walk-by","Google","Reference","Repeat Customer","Instagram"][i2%5]]);}
var cm=[["CaseId","CreatedDate","Remark"]];for(var i3=0;i3<26;i3++){if(i3%3!==0)cm.push(["C100"+i3,daysAgoStr(i3%6,"11"),"update"]);}
var av=[["MobileNo","AvailableBalance"]];for(var i4=0;i4<26;i4++)av.push(["9198765432"+("000000"+i4).slice(-3),(i4*137)%1500]);
var T=[];for(var i5=0;i5<20;i5++)T.push(new Array(18).fill(""));
T[0]=["Revenue Type","Store Target","SUBHA","DINESH","","","","","","",
"FY26P12","MDU4","VERY SATISFIED","SOMEWHAT SATISFIED","VERY DISSATISIFIED","NEITHER SATISFIED NOR DISSATISFIED","-","SOMEWHAT DISSATISFIED","TOTAL","C-SAT","CFI",""];
T[1]=["Service",900000,520000,380000];T[2]=["Accessories",430000,250000,180000];T[3]=["Repair",500000,300000,200000];T[4]=["Total",1830000,1070000,760000];T[5]=["Gross Profit",450000,270000,180000];
T[8]=["Prev Month","This Month","Difference"];T[9]=[1475,1494,19];
T[12]=["","AR","","","","RR","","","","SR","",""];T[13]=["","Care+","Protect+","AMC","Bundle","Display","Battery","Glass","Others","SRD","SHC","ARD"];
T[14]=["",25,15,40,30,50,45,20,35,60,25,18];
T[15]=["Achivement",400,600,900,300,"",250,150,100,0,0,0,0];T[16]=["70-80%","2000","70-80%","1000"];
T[17]=["80-90%","4000","80-90%","2000"];T[18]=["90-100%","6000","90-100%","3000"];
T[1][11]="Subha K";T[1][12]=9;T[1][13]=2;T[1][14]=0;T[1][15]=1;T[1][18]=12;T[1][19]="75.00%";T[1][20]="0.00%";
T[2][11]="Dinesh Kumar";T[2][12]=14;T[2][13]=2;T[2][14]=1;T[2][15]=0;T[2][17]=1;T[2][18]=18;T[2][19]="88.89%";T[2][20]="5.56%";
T[3][11]="Murugan Mp";T[3][12]=4;T[3][13]=1;T[3][14]=0;T[3][15]=1;T[3][18]=6;T[3][19]="83.33%";T[3][20]="0.00%";
T[4][11]="TOTAL";T[4][12]=27;T[4][13]=5;T[4][14]=1;T[4][15]=2;T[4][17]=1;T[4][18]=36;T[4][19]="88.89%";T[4][20]="5.56%";
T[15][5]="incentive per amc";T[16][5]="1-10";T[16][6]="₹150";T[16][8]=5;T[16][9]="₹80";T[17][5]="11-25";T[17][6]="₹250";T[17][8]=12;T[17][9]="₹120";T[15][12]="slab level";T[15][13]=5;T[16][13]=300;
// decline / quote mocks (Task 35)
var dl=[["CaseId","GSXRepairStatusDescription","Product"],
 ["D1001","Declined - Rejected by Local Approver","iPhone 13"],
 ["D1002","Declined - Rejected by Local Approver","Galaxy S22"],
 ["D1002","Declined - Rejected by Local Approver","Galaxy S22"],
 ["D1003","Cancelled","Pixel 7"],
 ["D1004","Declined - Rejected by Local Approver","OnePlus 11"],
 ["D1001","Declined - Rejected by Local Approver","iPhone 13"]];
var qt=[["CaseID","RevenueType","lastAction"],
 ["Q2001","Repair Revenue","approved"],
 ["Q2002","Repair Revenue","approved"],
 ["Q2002","Repair Revenue","approved"],
 ["Q2003","Repair Revenue","cancelled"],
 ["Q2004","Repair Revenue","approved"],
 ["Q2001","Accessory Revenue","approved"]];
ALL=[mockSheet("target",T),mockSheet("crm",crm),mockSheet("walkin",wl),mockSheet("service",sv),mockSheet("advance",av),mockSheet("comm",cm),
  mockSheet("decline",dl),mockSheet("quote",qt)];
var loaded=loadDashboardData_();
var page=buildDashboardPage_(loaded.data);
// /tmp/preview occasionally vanishes between Node invocations, so re-make it.
fs.mkdirSync("/tmp/preview",{recursive:true});
// Also mirror into the workspace so the artifact survives sandbox tmp cleanup.
try{ fs.mkdirSync("/home/user/_preview",{recursive:true}); }catch(e){}
fs.writeFileSync("/tmp/preview/index.html",page);
try{ fs.writeFileSync("/home/user/_preview/index.html",page); }catch(e){}
console.log("preview written:",page.length,"chars");
