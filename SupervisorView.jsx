import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ============================================================
// SUPABASE CONNECTION
// ============================================================
const SUPABASE_URL="https://jkxjitaoqsnxuayhawlu.supabase.co";
const SUPABASE_ANON_KEY="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpreGppdGFvcXNueHVheWhhd2x1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODYzNDc0NjMsImV4cCI6MjEwMTkyMzQ2M30.faTOFWW_sYpJm_J2EgiPbM4cQVWsZ6Fzlw3j3TMf9QY";

async function sbQuery(table,params){
  const url=SUPABASE_URL+"/rest/v1/"+table+(params?("?"+params):"");
  const res=await fetch(url,{headers:{
    "apikey":SUPABASE_ANON_KEY,
    "Authorization":"Bearer "+SUPABASE_ANON_KEY,
  }});
  if(!res.ok)throw new Error("Query failed: "+res.status);
  return res.json();
}

async function sbInsert(table,row){
  const url=SUPABASE_URL+"/rest/v1/"+table;
  const res=await fetch(url,{method:"POST",headers:{
    "apikey":SUPABASE_ANON_KEY,
    "Authorization":"Bearer "+SUPABASE_ANON_KEY,
    "Content-Type":"application/json",
    "Prefer":"return=representation",
  },body:JSON.stringify(row)});
  if(!res.ok)throw new Error("Insert failed: "+res.status);
  return res.json();
}

async function sbUpdate(table,params,row){
  const url=SUPABASE_URL+"/rest/v1/"+table+"?"+params;
  const res=await fetch(url,{method:"PATCH",headers:{
    "apikey":SUPABASE_ANON_KEY,
    "Authorization":"Bearer "+SUPABASE_ANON_KEY,
    "Content-Type":"application/json",
    "Prefer":"return=representation",
  },body:JSON.stringify(row)});
  if(!res.ok)throw new Error("Update failed: "+res.status);
  return res.json();
}


const BG="#F4F6FB",PNL="#fff",INK="#1A2340",SFT="#5B6580",LN="#DDE3F0";
const GRN="#1A8C4E",GBG="#E6F5ED",RED="#C0392B",RBG="#FBEAE8";
const BLU="#1565C0",BBG="#E3EEF9",AMB="#B45A00",ABG="#FDF3E7";
const PUR="#6B21A8",PBG="#F5EEFF",CB="#1B3A8A",DARK="#0F2461",CY="#F5C400";
const ORANGE="#CC4E00",OBG="#FFF0E5",YBG="#FFFBE0",YFG="#A07700";

const lbl={display:"block",fontSize:14,fontWeight:700,color:SFT,margin:"8px 0 3px"};
const inp={width:"100%",boxSizing:"border-box",padding:"8px 10px",borderRadius:8,border:"1px solid "+LN,fontSize:16,background:PNL};
const PBtn={width:"100%",marginTop:10,padding:"10px",background:CB,color:"#fff",border:"none",borderRadius:10,fontWeight:700,fontSize:16,cursor:"pointer"};
const AB={fontSize:14,fontWeight:700,padding:"6px 12px",borderRadius:8,border:"2px solid "+GRN,color:GRN,background:GBG,cursor:"pointer"};
const RB={fontSize:14,fontWeight:700,padding:"6px 12px",borderRadius:8,border:"2px solid "+RED,color:RED,background:RBG,cursor:"pointer"};

function Pill({c,bg,fg}){return <span style={{background:bg,color:fg,fontSize:13,fontWeight:700,padding:"2px 8px",borderRadius:12,display:"inline-block"}}>{c}</span>;}
function Card({children,s}){return <div style={{background:PNL,border:"1px solid "+LN,borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px rgba(27,58,138,.07)",...s}}>{children}</div>;}
function SHdr({title,sub}){return <div
  style={{marginBottom:10}}><div style={{fontSize:16,fontWeight:800,color:CB,textTransform:"uppercase",letterSpacing:.4}}>{title}</div>{sub&&<div style={{fontSize:14,color:SFT,marginTop:2}}>{sub}</div>}</div>;}

function monthsWorked(d){if(!d)return 0;const j=new Date(d),t=new Date();return Math.max(0,(t.getFullYear()-j.getFullYear())*12+(t.getMonth()-j.getMonth()));}
const CLIENTS_DATA_BY_MONTH={
  "Jun 2026":[
    {name:"Al Futtaim Construction",sites:[{site:"Site A",present:38,absent:3,medical:2,idle:1,total:44},{site:"Site B",present:22,absent:4,medical:0,idle:2,total:28}]},
    {name:"Emaar Properties",sites:[{site:"Site C",present:18,absent:2,medical:1,idle:0,total:21}]},
  ],
  "May 2026":[
    {name:"Al Futtaim Construction",sites:[{site:"Site A",present:41,absent:1,medical:1,idle:0,total:43},{site:"Site B",present:26,absent:1,medical:0,idle:1,total:28}]},
    {name:"Emaar Properties",sites:[{site:"Site C",present:20,absent:0,medical:0,idle:1,total:21}]},
  ],
  "Apr 2026":[
    {name:"Al Futtaim Construction",sites:[{site:"Site A",present:39,absent:2,medical:2,idle:1,total:44},{site:"Site B",present:24,absent:2,medical:1,idle:1,total:28}]},
    {name:"Emaar Properties",sites:[{site:"Site C",present:19,absent:1,medical:0,idle:1,total:21}]},
  ],
};
const buildKPI=(month)=>{
  const clients=CLIENTS_DATA_BY_MONTH[month]||CLIENTS_DATA_BY_MONTH["Jun 2026"];
  let total=0,present=0,idle=0,absent=0;
  clients.forEach(c=>c.sites.forEach(s=>{total+=s.total;present+=s.present;idle+=s.idle;absent+=s.absent;}));
  const attRate=total>0?Math.round(present/total*100):0;
  const idleScore=total>0?Math.round((1-idle/total)*100):100;
  const absScore=total>0?Math.round((1-absent/total)*100):100;
  const col=(s,t)=>s>=t?GRN:s>=t-10?AMB:RED;
  return[
    {label:"Idle Man-Day Control",        weight:35,score:idleScore,target:85,col:col(idleScore,85)},
    {label:"Resignation & Termination",   weight:30,score:96,       target:90,col:col(96,90)},
    {label:"Unplanned Absence MD Control",weight:15,score:absScore, target:85,col:col(absScore,85)},
    {label:"Requirement Fulfillment",     weight:10,score:attRate,  target:95,col:col(attRate,95)},
    {label:"Reporting Accuracy",          weight:10,score:88,       target:90,col:col(88,90)},
  ];
};

const today=new Date();

const ALL_WORKERS=[
  {id:1,name:"Ramesh Kumar",  empCode:"EMP-1042",passport:"N1234567",client:"Al Futtaim Construction",site:"Camp 3 - Site A",trade:"Mason",      stampDate:"2024-01-15",visaExpiry:"2026-07-14",passportExpiry:"2028-03-14",medicalInsExpiry:"2026-12-31",wcExpiry:"2026-12-31",siteChanges:1,status:"P",nationality:"Indian",  payType:"fixed", basicJoin:1300,foodJoin:150,otherAllowJoin:0,basic:1500,food:200,otherAllow:0,advancePending:200},
  {id:2,name:"Suresh Pillai", empCode:"EMP-1043",passport:"P7654321",client:"Al Futtaim Construction",site:"Camp 3 - Site A",trade:"Carpenter",  stampDate:"2025-03-01",visaExpiry:"2027-02-28",passportExpiry:"2029-06-01",medicalInsExpiry:"2027-02-28",wcExpiry:"2027-02-28",siteChanges:0,status:"A",nationality:"Indian",  payType:"fixed", basicJoin:1500,foodJoin:200,otherAllowJoin:0,basic:1500,food:200,otherAllow:0,advancePending:0},
  {id:3,name:"Faisal Ahmed",  empCode:"EMP-1044",passport:"A9988776",client:"Emaar Properties",        site:"Camp 1 - Site C",trade:"Electrician",stampDate:"2024-06-10",visaExpiry:"2026-07-15",passportExpiry:"2026-08-20",medicalInsExpiry:"2026-06-09",wcExpiry:"2026-06-09",siteChanges:2,status:"M",nationality:"Pakistani",payType:"hourly",rateJoin:80,rate:95,food:0,otherAllow:0,advancePending:0},
  {id:4,name:"Bikash Thapa",  empCode:"EMP-1045",passport:"B4455667",client:"Al Futtaim Construction",site:"Camp 3 - Site A",trade:"Helper",     stampDate:"2025-04-20",visaExpiry:"2027-04-19",passportExpiry:"2027-10-20",medicalInsExpiry:"2027-04-19",wcExpiry:"2027-04-19",siteChanges:0,status:"WO",nationality:"Nepali",  payType:"fixed", basicJoin:1100,foodJoin:200,otherAllowJoin:0,basic:1200,food:200,otherAllow:0,advancePending:0},
  {id:5,name:"Sanjay Gupta",  empCode:"EMP-1046",passport:"G7788990",client:"Emaar Properties",        site:"Camp 1 - Site C",trade:"Welder",    stampDate:"2023-11-05",visaExpiry:"2025-11-04",passportExpiry:"2026-09-05",medicalInsExpiry:"2025-11-04",wcExpiry:"2025-11-04",siteChanges:3,status:"P",nationality:"Indian",  payType:"hourly",rateJoin:75,rate:85,food:0,otherAllow:0,advancePending:500},
  {id:6,name:"Arjun Das",     empCode:"EMP-1047",passport:"D1122334",client:"Al Futtaim Construction",site:"Camp 3 - Site A",trade:"Plumber",    stampDate:"2024-08-22",visaExpiry:"2026-08-21",passportExpiry:"2027-02-22",medicalInsExpiry:"2026-08-21",wcExpiry:"2026-08-21",siteChanges:1,status:"I",nationality:"Indian",  payType:"fixed", basicJoin:1300,foodJoin:200,otherAllowJoin:0,basic:1300,food:200,otherAllow:0,advancePending:0},
];
const WORKER_ATT={
  "EMP-1042":{1:"P",2:"P",3:"A",4:"P",5:"P",8:"P",9:"P",10:"M",11:"M",15:"P",16:"P",17:"P",18:"WO",19:"P",22:"P",23:"P",24:"P",25:"A",26:"P"},
  "EMP-1043":{1:"P",2:"P",3:"P",4:"P",5:"P",8:"P",9:"P",10:"P",11:"P",15:"P",16:"P",17:"WO",18:"P",19:"P",22:"P",23:"P",24:"P",25:"P",26:"P"},
  "EMP-1044":{1:"P",2:"A",3:"P",4:"P",8:"P",9:"P",10:"P",15:"P",16:"P",17:"P",18:"P",19:"M",22:"P",23:"P",24:"A",25:"P",26:"P"},
  "EMP-1045":{1:"P",2:"P",3:"P",4:"P",8:"P",9:"P",10:"P",11:"P",15:"P",16:"A",17:"P",18:"P",22:"P",23:"P",24:"P",25:"P",26:"P"},
  "EMP-1046":{1:"P",2:"P",3:"P",5:"P",8:"P",9:"A",10:"P",15:"P",16:"P",17:"P",22:"P",23:"P"},
  "EMP-1047":{1:"P",2:"P",3:"P",4:"P",5:"P",8:"A",9:"P",10:"P",15:"P",16:"P",17:"P",18:"WO",22:"P",23:"I",24:"P"},
};
const SUPS=[
  {id:"s1",name:"Ahmed Al Rashidi",empCode:"SUP-001",clients:["Al Futtaim Construction","Emaar Properties"]},
  {id:"s2",name:"Vijay Menon",     empCode:"SUP-002",clients:["DEWA Infrastructure"]},
];
const CLIENTS_DATA=[
  {name:"Al Futtaim Construction",sites:[{site:"Site A",present:38,absent:3,medical:2,idle:1,total:44},{site:"Site B",present:22,absent:4,medical:0,idle:2,total:28}]},
  {name:"Emaar Properties",sites:[{site:"Site C",present:18,absent:2,medical:1,idle:0,total:21}]},
];
const SUPER_ADMIN_APPROVAL=["Medical W Paper","Medical W/O Paper","Idle","Emergency Leave","Annual Leave","Absconder","Resignation / Under Cancellation","On Training","No Salary No Fine (NSF)","Fine - Waiver Request","Advance > AED 200"];
const ADMIN_ONLY_APPROVAL=["Annual Leave","Grade Override","Advance <= AED 200"];
const ADMIN_APPROVAL=[...SUPER_ADMIN_APPROVAL,...ADMIN_ONLY_APPROVAL];
const STATUS_LIST=["Present","Absent","Idle","Medical W Paper","Medical W/O Paper","Weekly Off","Public Holiday","Emergency Leave","Annual Leave","Absconder","Refusal to work","Resignation / Under Cancellation","On Training","No Salary No Fine (NSF)"];
const STATUS_STYLE={P:{bg:GBG,fg:GRN},A:{bg:RBG,fg:RED},M:{bg:BBG,fg:BLU},I:{bg:YBG,fg:YFG},WO:{bg:PBG,fg:PUR},PH:{bg:OBG,fg:ORANGE}};

const WDATA=(()=>{
  const genDays=(dim,sunOff)=>{
    const d={};
    for(let i=1;i<=dim;i++){
      const day=new Date(today.getFullYear(),today.getMonth(),i).getDay();
      if(day===0){d[i]="WO";}
      else if(i>today.getDate()){}
      else{const r=Math.random();d[i]=r>0.88?"A":r>0.83?"M":r>0.79?"I":"P";}
    }
    return d;
  };
  const genCd=(wDays)=>{
    const cd={};
    Object.entries(wDays).forEach(([d,st])=>{if(st==="P")cd[Number(d)]=8+Math.round(Math.random()*2);});
    return cd;
  };
  const dim=new Date(today.getFullYear(),today.getMonth()+1,0).getDate();
  const result={};
  ALL_WORKERS.forEach(w=>{
    const wDays0=w.empCode==="EMP-1042"?{2:"P",3:"P",4:"A",5:"P",8:"P",9:"P",10:"M",11:"M",15:"P",16:"P",17:"P",18:"WO",19:"P",22:"P",23:"P"}:genDays(dim);
    const wDays1=genDays(28);const wDays2=genDays(30);
    result[w.empCode]={months:[
      {lb:"Jun 2026",wDays:wDays0,cdUp:true,cd:genCd(wDays0)},
      {lb:"May 2026",wDays:wDays1,cdUp:true,cd:genCd(wDays1)},
      {lb:"Apr 2026",wDays:wDays2,cdUp:true,cd:genCd(wDays2)},
      {lb:"Mar 2026",wDays:{},cdUp:false,cd:{}},
    ]};
  });
  return result;
})()

function buildMonthCal(yr,mo){
  const dim=new Date(yr,mo+1,0).getDate();
  return Array.from({length:dim},(_,i)=>{
    const d=new Date(yr,mo,i+1);
    return{date:d,sun:d.getDay()===0,fut:d>today,day:i+1};
  });
}

const SS={P:{bg:GBG,fg:GRN,lb:"P"},A:{bg:RBG,fg:RED,lb:"A"},M:{bg:BBG,fg:BLU,lb:"M"},I:{bg:YBG,fg:YFG,lb:"I"},WO:{bg:PBG,fg:PUR,lb:"WO"},H:{bg:OBG,fg:ORANGE,lb:"H"},future:{bg:"#F0F4FF",fg:"#9AABCF",lb:""}};
const STATUS_COLS={P:{bg:GBG,fg:GRN},A:{bg:RBG,fg:RED},M:{bg:BBG,fg:BLU},I:{bg:YBG,fg:YFG},WO:{bg:PBG,fg:PUR},PH:{bg:OBG,fg:ORANGE},MWP:{bg:BBG,fg:BLU},MWOP:{bg:ABG,fg:AMB},NSF:{bg:RBG,fg:RED},EL:{bg:RBG,fg:RED},AL:{bg:BBG,fg:BLU},AB:{bg:RBG,fg:RED},REF:{bg:RBG,fg:RED},RES:{bg:ABG,fg:AMB},OT:{bg:GBG,fg:GRN}};
const STATUS_CODE_MAP=[["P","Present"],["A","Absent"],["I","Idle"],["MWP","Medical W Paper"],["MWOP","Medical W/O Paper"],["WO","Weekly Off"],["PH","Public Holiday"],["EL","Emergency Leave"],["AL","Annual Leave"],["AB","Absconder"],["REF","Refusal to work"],["RES","Resignation"],["OT","On Training"],["NSF","No Salary No Fine"]];

function Block({label,val,fg,bg}){return(<div
  style={{background:bg,borderRadius:9,padding:"8px 10px",display:"flex",alignItems:"center",gap:8,border:"1px solid "+LN}}><span style={{fontSize:23,fontWeight:900,color:fg||CB,minWidth:26,textAlign:"center"}}>{val}</span><span style={{fontSize:12.5,fontWeight:700,color:SFT,lineHeight:1.3}}>{label}</span></div>);}

/* REPORT TAB */
function workerGrade(empCode,gradeOverrides,workerStatus){
  if(gradeOverrides&&gradeOverrides[empCode]) return gradeOverrides[empCode];
  // Auto-red statuses regardless of attendance
  if(["Absconder","Refusal to work","Resignation / Under Cancellation"].includes(workerStatus)) return "red";
  const wdata=typeof WDATA!=="undefined"?WDATA[empCode]:null;
  if(!wdata||!wdata.months) return "yellow";
  // Statuses excluded from working day & absent calculation
  const EXCUSED=new Set(["WO","PH","Emergency Leave","Annual Leave","On Training","No Salary No Fine (NSF)"]);
  let totalWorking=0,totalAbsent=0;
  wdata.months.forEach(m=>{
    const days=Object.values(m.wDays||{});
    let mPresent=0,mAbsent=0,mMedical=0,mIdle=0;
    days.forEach(st=>{
      if(EXCUSED.has(st)) return; // excluded - not counted at all
      if(st==="P")      mPresent++;
      else if(st==="I") mIdle++;   // Idle counts as present
      else if(st==="M") mMedical++;
      else if(st==="A") mAbsent++;
      // Absconder/Refusal/Resignation handled above at function level
    });
    // Medical: if >3 days in month ALL medical days count as absent
    const medAbsent=mMedical>3?mMedical:0;
    const monthWorking=mPresent+mIdle+mAbsent+mMedical;
    const monthAbsent=mAbsent+medAbsent;
    totalWorking+=monthWorking;
    totalAbsent+=monthAbsent;
  });
  if(totalWorking===0) return "green";
  const pct=((totalWorking-totalAbsent)/totalWorking)*100;
  if(pct>90) return "green";
  if(pct>=75) return "yellow";
  return "red";
}
function GradeBadge({grade,size}){
  const bg=grade==="green"?"#00C851":grade==="yellow"?"#FFD600":"#C0392B";
  const sz=size||10;
  const t=grade==="green"?"Attendance above 90% - Excellent":grade==="yellow"?"Attendance 75-90% - Good":"Attendance below 75% or automatic red status";
  return <div title={t} style={{width:sz,height:sz,borderRadius:"50%",background:bg,flexShrink:0,boxShadow:"0 1px 4px rgba(0,0,0,0.25)"}}/>;
}


function SupRequestForm(){
  const [reqType,setReqType]=useState("Cash Advance for Myself");
  const [reqAmount,setReqAmount]=useState("");
  const [reqReason,setReqReason]=useState("");
  const [reqRemarks,setReqRemarks]=useState("");
  const [reqSent,setReqSent]=useState(false);
  const submitReq=()=>{
    if(!reqReason)return;
    setReqSent(true);
    setTimeout(()=>{setReqSent(false);setReqAmount("");setReqReason("");setReqRemarks("");},3000);
  };
  if(reqSent) return <div style={{background:GBG,borderRadius:8,padding:"10px",fontSize:15,color:GRN,fontWeight:700,textAlign:"center"}}>Request submitted to Admin for approval.</div>;
  return(
    <>
      <label style={lbl}>Request Type</label>
      <select style={inp} value={reqType} onChange={e=>setReqType(e.target.value)}>
        {["Cash Advance for Myself","Petty Cash for Site Operations","Transport or Fuel Expense","Emergency Site Expense","Tool or Equipment Purchase","Other Operational Requirement"].map(t=><option key={t}>{t}</option>)}
      </select>
      {(reqType.includes("Cash")||reqType.includes("Expense")||reqType.includes("Purchase"))&&(
        <>
          <label style={lbl}>Amount (AED)</label>
          <input type="number" style={inp} value={reqAmount} onChange={e=>setReqAmount(e.target.value)} placeholder="e.g. 500"/>
          <div style={{fontSize:13,color:SFT,marginTop:3,marginBottom:4}}>Amount above AED 200 will require Super Admin approval.</div>
        </>
      )}
      <label style={lbl}>Reason (required)</label>
      <input style={inp} value={reqReason} onChange={e=>setReqReason(e.target.value)} placeholder="e.g. Need petty cash for emergency camp supplies"/>
      <label style={lbl}>Remarks / Additional Details</label>
      <textarea style={{...inp,height:60}} value={reqRemarks}
        onChange={e=>setReqRemarks(e.target.value)} placeholder="e.g. Will submit receipts within 3 days. Amount to be settled from camp budget."/>
      <button style={{...PBtn,background:"linear-gradient(135deg,"+DARK+","+CB+")",opacity:!reqReason?0.5:1,marginTop:10}}
        onClick={submitReq} disabled={!reqReason}>Submit Request for Approval</button>
    </>
  );
}
function SubmitTimesheetsPanel({sup}){
  const [open,setOpen]=useState(false);
  const [rClient,setRClient]=useState("");
  const [rSite,setRSite]=useState("");
  const [rMonth,setRMonth]=useState("Jun 2026");
  const [submitted,setSubmitted]=useState([]);
  const allW=sup.workers;
  const clients=[...new Set(allW.map(w=>w.client))];
  const sites=[...new Set(allW.filter(w=>!rClient||w.client===rClient).map(w=>w.site))];
  const filteredW=allW.filter(w=>(!rClient||w.client===rClient)&&(!rSite||w.site===rSite));
  const SUPER_PEND_STATUSES=["MWP","MWOP","I","NSF","RES"];
  const getWorkerStatus=(w)=>{
    if(submitted.includes(w.empCode)) return "submitted";
    const wDays=WORKER_ATT[w.empCode]||{};
    const hasPending=Object.values(wDays).some(st=>SUPER_PEND_STATUSES.includes(st));
    const hasData=Object.keys(wDays).length>0;
    if(!hasData) return "incomplete";
    if(hasPending) return "pending";
    return "ready";
  };
  const readyWorkers=filteredW.filter(w=>getWorkerStatus(w)==="ready");
  const pendingWorkers=filteredW.filter(w=>getWorkerStatus(w)==="pending");
  const incompleteWorkers=filteredW.filter(w=>getWorkerStatus(w)==="incomplete");
  const submittedWorkers=filteredW.filter(w=>getWorkerStatus(w)==="submitted");
  const submitAll=()=>{
    const ids=readyWorkers.map(w=>w.empCode);
    setSubmitted([...submitted,...ids]);
  };
  const statusConfig={
    ready:{bg:GBG,fg:GRN,label:"Ready to Submit",icon:"checkmark"},
    pending:{bg:ABG,fg:AMB,label:"Override Pending Super Admin"},
    incomplete:{bg:BG,fg:SFT,label:"Client data not yet entered"},
    submitted:{bg:BBG,fg:BLU,label:"Submitted - visible to worker"},
  };
  if(!open) return(
    <>
      <button
        onClick={()=>setOpen(true)} style={{...PBtn,background:"linear-gradient(135deg,#B8960A,"+CY+")",color:DARK,marginBottom:6,marginTop:0,display:"flex",alignItems:"center",justifyContent:"center",gap:8,fontWeight:900,fontSize:16,boxShadow:"0 3px 10px rgba(245,196,0,0.4)"}}>
        Submit Client Timesheets
      </button>
      <div style={{fontSize:13,color:SFT,textAlign:"center",marginBottom:14,padding:"0 4px"}}>
        Only submit after completing worker reconciliation, verifying client attendance, and ensuring all override and approval requests have been resolved.
      </div>
    </>
  );
  return(
    <div style={{marginBottom:14}}>
      <div style={{background:"linear-gradient(135deg,"+DARK+","+CB+")",borderRadius:12,padding:"12px 16px",marginBottom:10}}>
        <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:10}}>
          <div style={{fontSize:16,fontWeight:800,color:"#fff"}}>Submit Client Timesheets</div>
          <button
            onClick={()=>setOpen(false)} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",borderRadius:8,padding:"4px 10px",cursor:"pointer",fontSize:14,fontWeight:700}}>Close</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
          <div>
            <label style={{...lbl,marginTop:0,color:"rgba(255,255,255,0.6)"}}>Month</label>
            <select style={{...inp,background:"rgba(255,255,255,0.1)",color:"#fff",border:"1px solid rgba(255,255,255,0.2)"}} value={rMonth} onChange={e=>setRMonth(e.target.value)}>
              {["Jun 2026","May 2026","Apr 2026","Mar 2026"].map(m=><option key={m} style={{color:INK}}>{m}</option>)}
            </select>
          </div>
          <div>
            <label style={{...lbl,marginTop:0,color:"rgba(255,255,255,0.6)"}}>Client</label>
            <select style={{...inp,background:"rgba(255,255,255,0.1)",color:"#fff",border:"1px solid rgba(255,255,255,0.2)"}} value={rClient} onChange={e=>{setRClient(e.target.value);setRSite("");}}>
              <option value="" style={{color:INK}}>All clients</option>
              {clients.map(c=><option key={c} style={{color:INK}}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={{...lbl,marginTop:0,color:"rgba(255,255,255,0.6)"}}>Site</label>
            <select style={{...inp,background:"rgba(255,255,255,0.1)",color:"#fff",border:"1px solid rgba(255,255,255,0.2)"}} value={rSite} onChange={e=>setRSite(e.target.value)}>
              <option value="" style={{color:INK}}>All sites</option>
              {sites.map(s=><option key={s} style={{color:INK}}>{s}</option>)}
            </select>
          </div>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:6,marginBottom:10}}>
        {[["Ready",readyWorkers.length,GRN,GBG],["Pending SA",pendingWorkers.length,AMB,ABG],["Incomplete",incompleteWorkers.length,SFT,BG],["Submitted",submittedWorkers.length,BLU,BBG]].map(([l,v,fg,bg])=>(
          <div key={l} style={{background:bg,borderRadius:9,padding:"8px",textAlign:"center",border:"1px solid "+LN}}>
            <div style={{fontSize:23,fontWeight:900,color:fg}}>{v}</div>
            <div style={{fontSize:12,fontWeight:700,color:fg,marginTop:2}}>{l}</div>
          </div>
        ))}
      </div>
      {readyWorkers.length>0&&(
        <button onClick={submitAll} style={{...PBtn,background:GRN,marginBottom:10,marginTop:0}}>
          Submit All {readyWorkers.length} Ready Worker(s) - Publish to Worker Calendar
        </button>
      )}
      <div style={{display:"flex",flexDirection:"column",gap:6}}>
        {filteredW.map((w,i)=>{
          const wst=getWorkerStatus(w);
          const cfg=statusConfig[wst];
          const wDays=WORKER_ATT[w.empCode]||{};
          const pendDays=Object.entries(wDays).filter(([d,st])=>SUPER_PEND_STATUSES.includes(st)).map(([d,st])=>d+" "+st);
          return(
            <div key={w.id}
              style={{background:PNL,border:"1px solid "+LN,borderRadius:10,padding:"10px 12px",display:"flex",justifyContent:"space-between",alignItems:"flex-start",borderLeft:"4px solid "+cfg.fg}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:INK}}>{w.name} <span style={{fontSize:13,color:SFT}}>{w.empCode}</span></div>
                <div style={{fontSize:13,color:SFT,marginTop:2}}>{w.site}</div>
                {wst==="pending"&&pendDays.length>0&&(
                  <div style={{fontSize:13,color:AMB,marginTop:3,fontWeight:600}}>Pending Super Admin: {pendDays.join(", ")}</div>
                )}
                {wst==="submitted"&&(
                  <div style={{fontSize:13,color:BLU,marginTop:3,fontWeight:600}}>Timesheet published - visible to worker</div>
                )}
              </div>
              <span style={{fontSize:12,fontWeight:700,padding:"3px 8px",borderRadius:8,background:cfg.bg,color:cfg.fg,whiteSpace:"nowrap"}}>{cfg.label}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}


function MyRequestsTab(){
  const [items]=useState([
    {id:1,type:"Cash Advance for Myself",amount:300,reason:"Camp supplies emergency",status:"Pending from Management",date:"Today"},
    {id:2,type:"Transport or Fuel Expense",amount:150,reason:"Site visit fuel reimbursement",status:"Approved",date:"25 Jun"},
  ]);
  const [showForm,setShowForm]=useState(false);
  const [reqType,setReqType]=useState("Cash Advance for Myself");
  const [reqAmount,setReqAmount]=useState("");
  const [reqReason,setReqReason]=useState("");
  const [reqRemarks,setReqRemarks]=useState("");
  const [sent,setSent]=useState(false);
  const submit=()=>{
    if(!reqReason)return;
    setSent(true);setShowForm(false);
    setTimeout(()=>setSent(false),3000);
    setReqAmount("");setReqReason("");setReqRemarks("");
  };
  const stColor=(st)=>st==="Approved"?{bg:GBG,fg:GRN}:st==="Rejected"?{bg:RBG,fg:RED}:{bg:ABG,fg:AMB};
  return(
    <>
      {sent&&<div style={{background:GBG,borderRadius:8,padding:"9px 12px",marginBottom:10,fontSize:15,color:GRN,fontWeight:700}}>Request submitted to Admin for approval.</div>}
      {!showForm?(
        <button onClick={()=>setShowForm(true)} style={{...PBtn,background:"linear-gradient(135deg,"+DARK+","+CB+")",marginBottom:12,marginTop:0}}>+ New Operational Request</button>
      ):(
        <Card s={{marginBottom:12,border:"2px solid "+CB}}>
          <SHdr title="New Supervisor Operational Request" sub="Submitted to Super Admin for approval"/>
          <label style={lbl}>Request Type</label>
          <select style={inp} value={reqType} onChange={e=>setReqType(e.target.value)}>
            {["Cash Advance for Myself","Petty Cash for Site Operations","Transport or Fuel Expense","Emergency Site Expense","Tool or Equipment Purchase","Other Operational Requirement"].map(t=><option key={t}>{t}</option>)}
          </select>
          {(reqType.includes("Cash")||reqType.includes("Expense")||reqType.includes("Purchase"))&&(
            <>
              <label style={lbl}>Amount (AED)</label>
              <input type="number" style={inp} value={reqAmount} onChange={e=>setReqAmount(e.target.value)} placeholder="e.g. 300"/>
              {reqAmount&&<div
                style={{fontSize:13,color:Number(reqAmount)>200?AMB:BLU,marginTop:3,fontWeight:600}}>{Number(reqAmount)<=200?"Admin approval only":"Super Admin approval required (above AED 200)"}</div>}
            </>
          )}
          <label style={lbl}>Reason (required)</label>
          <input style={inp} value={reqReason} onChange={e=>setReqReason(e.target.value)} placeholder="e.g. Emergency camp supply purchase"/>
          <label style={lbl}>Remarks / Details</label>
          <textarea style={{...inp,height:60}} value={reqRemarks} onChange={e=>setReqRemarks(e.target.value)} placeholder="e.g. Receipts will be submitted within 3 days"/>
          <div style={{display:"flex",gap:8,marginTop:10}}>
            <button style={{...PBtn,margin:0,flex:2,background:GRN,opacity:!reqReason?0.5:1}} onClick={submit} disabled={!reqReason}>Submit to Admin</button>
            <button
              onClick={()=>setShowForm(false)} style={{flex:1,padding:"10px",border:"1px solid "+LN,borderRadius:10,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:15}}>Cancel</button>
          </div>
        </Card>
      )}
      {items.map((r,i)=>{
        const sc=stColor(r.status);
        return(
          <Card key={r.id} s={{marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div>
                <div style={{fontSize:16,fontWeight:700}}>{r.type}</div>
                <div style={{fontSize:14,color:SFT}}>{r.date}{r.amount?" | AED "+r.amount:""}</div>
                <div style={{fontSize:14,color:SFT,marginTop:2}}>{r.reason}</div>
              </div>
              <Pill c={r.status} bg={sc.bg} fg={sc.fg}/>
            </div>
          </Card>
        );
      })}
    </>
  );
}


function ApprovalsTab({advPending=2,medPending=2,sup}){
  const compResolved=3,compInProgress=2;
  const [search,setSearch]=useState("");
  const [stageFilter,setStageFilter]=useState("");
  const [typeFilter,setTypeFilter]=useState("");
  const ITEMS=[
    {id:1,worker:"Bikash Thapa",empCode:"EMP-1045",client:"Al Futtaim Construction",site:"Camp 3 - Site A",type:"Medical - No Paper",stage:"Pending Supervisor",date:"Today",month:"Jun 2026"},
    {id:2,worker:"Sanjay Gupta",empCode:"EMP-1046",client:"Emaar Properties",site:"Camp 1 - Site C",type:"Idle",stage:"Pending Super Admin",date:"Today",month:"Jun 2026"},
    {id:3,worker:"Arjun Das",empCode:"EMP-1047",client:"Al Futtaim Construction",site:"Camp 3 - Site A",type:"Emergency Leave",stage:"Pending Super Admin",date:"Yesterday",month:"Jun 2026"},
    {id:4,worker:"Faisal Ahmed",empCode:"EMP-1044",client:"Emaar Properties",site:"Camp 1 - Site C",type:"Medical - No Paper",stage:"Pending Super Admin",date:"20 Jun",month:"Jun 2026"},
    {id:5,worker:"Suresh Pillai",empCode:"EMP-1043",client:"Al Futtaim Construction",site:"Camp 3 - Site A",type:"Annual Leave",stage:"Pending Admin",date:"22 Jun",month:"Jun 2026"},
    {id:6,worker:"Ramesh Kumar",empCode:"EMP-1042",client:"Al Futtaim Construction",site:"Camp 3 - Site A",type:"Fine - Waiver Request",stage:"Pending Super Admin",date:"10 Jun",month:"Jun 2026"},
    {id:7,worker:"Arjun Das",empCode:"EMP-1047",client:"Al Futtaim Construction",site:"Camp 3 - Site A",type:"Grade Override",stage:"Pending Admin",date:"08 Jun",month:"Jun 2026"},
  ];
  const STAGE_COL={"Pending Supervisor":{bg:ABG,fg:AMB},"Pending Admin":{bg:BBG,fg:BLU},"Pending Super Admin":{bg:PBG,fg:PUR},"Approved":{bg:GBG,fg:GRN},"Rejected":{bg:RBG,fg:RED}};
  const filtered=ITEMS.filter(a=>{
    const q=search.toLowerCase();
    return(!search||a.worker.toLowerCase().includes(q)||a.empCode.toLowerCase().includes(q))&&(!stageFilter||a.stage===stageFilter)&&(!typeFilter||a.type===typeFilter);
  });

  const StatRow=({label,items})=>(
    <div style={{marginBottom:6}}>
      <div style={{fontSize:12,fontWeight:700,color:"rgba(255,255,255,0.7)",marginBottom:3}}>{label}</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:4}}>
        {items.map(([l,v,fg,bg])=>(
          <div key={l} style={{background:bg,borderRadius:7,padding:"4px 7px",display:"flex",alignItems:"center",gap:5,border:"1px solid "+LN,opacity:v===0?0.4:1}}>
            <span style={{fontSize:16,fontWeight:900,color:fg,minWidth:16}}>{v}</span>
            <span style={{fontSize:11.5,fontWeight:700,color:SFT,lineHeight:1.3}}>{l}</span>
          </div>
        ))}
      </div>
    </div>
  );

  return(
    <>
      <div style={{background:"linear-gradient(135deg,"+GRN+",#1aad60)",borderRadius:10,padding:"8px 12px",marginBottom:10}}>
        <div style={{fontSize:14,fontWeight:800,color:"#fff",marginBottom:2}}>Supervisor Actions Only</div>
        <div style={{fontSize:12,color:"rgba(255,255,255,0.7)",marginBottom:8}}>Requests from workers that require supervisor action before going to Admin / Super Admin</div>
        {[
          {label:"Medical With Paper (MWP)",pending:2,completed:5,note:"Verify certificate, forward to Super Admin"},
          {label:"Medical Without Paper (MWOP)",pending:1,completed:3,note:"Verify reason, forward to Super Admin"},
          {label:"Idle Requests",pending:3,completed:4,note:"Confirm with client, forward to Super Admin"},
          {label:"PPE Requests",pending:2,completed:6,note:"Check stock, forward to Super Admin"},
          {label:"Advances (<=AED 200)",pending:advPending,completed:3,note:"Verify, forward to Admin + Super Admin"},
          {label:"Complaints",pending:compInProgress,completed:compResolved,note:"Investigate and resolve"},
          {label:"Timesheet Batches",pending:1,completed:2,note:"Review and submit for Admin approval"},
        ].map(({label,pending,completed,note})=>(
          <div key={label} style={{marginBottom:6,background:"rgba(255,255,255,0.08)",borderRadius:8,padding:"6px 8px"}}>
            <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:3}}>
              <div style={{fontSize:13,fontWeight:700,color:"#fff"}}>{label}</div>
              <div style={{fontSize:11,color:"rgba(255,255,255,0.5)",fontStyle:"italic"}}>{note}</div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
              <div style={{background:pending>0?ABG:GBG,borderRadius:6,padding:"4px 8px",display:"flex",alignItems:"center",gap:5,border:"1px solid "+LN,opacity:pending===0?0.5:1}}>
                <span style={{fontSize:16,fontWeight:900,color:pending>0?AMB:GRN,minWidth:16}}>{pending}</span>
                <span style={{fontSize:11.5,fontWeight:700,color:SFT}}>Pending Action</span>
              </div>
              <div style={{background:GBG,borderRadius:6,padding:"4px 8px",display:"flex",alignItems:"center",gap:5,border:"1px solid "+LN}}>
                <span style={{fontSize:16,fontWeight:900,color:GRN,minWidth:16}}>{completed}</span>
                <span style={{fontSize:11.5,fontWeight:700,color:SFT}}>Actioned / Forwarded</span>
              </div>
            </div>
          </div>
        ))}
      </div>

      <div style={{background:"linear-gradient(135deg,"+BLU+",#1976d2)",borderRadius:10,padding:"8px 12px",marginBottom:10}}>
        <div style={{fontSize:14,fontWeight:800,color:"#fff",marginBottom:6}}>Admin Approval Only</div>
        <StatRow label="Annual Leave" items={[["Pending",1,AMB,ABG],["Approved",4,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Grade Override" items={[["Pending",1,AMB,ABG],["Approved",2,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Advance (<=AED 200)" items={[["Pending",2,AMB,ABG],["Approved",3,GRN,GBG],["Rejected",0,RED,RBG]]}/>
      </div>

      <div style={{background:"linear-gradient(135deg,"+DARK+","+CB+")",borderRadius:10,padding:"8px 12px",marginBottom:14}}>
        <div style={{fontSize:14,fontWeight:800,color:"#fff",marginBottom:6}}>Super Admin Approval Only</div>
        <StatRow label="Medical With Paper (MWP)" items={[["Pending",2,AMB,ABG],["Approved",5,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Medical Without Paper (MWOP)" items={[["Pending",1,AMB,ABG],["Approved",3,GRN,GBG],["Rejected",1,RED,RBG]]}/>
        <StatRow label="Idle" items={[["Pending",2,YFG,YBG],["Approved",3,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Emergency Leave" items={[["Pending",1,AMB,ABG],["Approved",1,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Absconder" items={[["Pending",1,AMB,ABG],["Approved",1,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Resignation / Under Cancellation" items={[["Pending",0,AMB,ABG],["Approved",1,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="On Training" items={[["Pending",0,AMB,ABG],["Approved",2,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="No Salary No Fine (NSF)" items={[["Pending",1,AMB,ABG],["Approved",1,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Fine - Waiver Request" items={[["Pending",1,AMB,ABG],["Approved",2,GRN,GBG],["Rejected",0,RED,RBG]]}/>
        <StatRow label="Advance > AED 200" items={[["Pending",1,AMB,ABG],["Approved",2,GRN,GBG],["Rejected",0,RED,RBG]]}/>
      </div>

      <div style={{borderTop:"2px solid "+LN,paddingTop:14}}>
        <div style={{fontSize:15,fontWeight:800,color:CB,marginBottom:10}}>APPROVAL REQUESTS</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
          <input style={inp} value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search worker or code..."/>
          <select style={inp} value={stageFilter} onChange={e=>setStageFilter(e.target.value)}>
            <option value="">All stages</option>
            {["Pending Supervisor","Pending Admin","Pending Super Admin","Approved","Rejected"].map(s=><option key={s}>{s}</option>)}
          </select>
        </div>
        <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
          <div style={{fontSize:14,color:SFT}}>{filtered.length} request(s){stageFilter?" - "+stageFilter:""}</div>
          <button onClick={()=>{
            const rows=filtered.map(a=>`<tr style="border-bottom:1px solid #eee"><td style="padding:6px 8px;font-weight:600">${a.worker}<br/><span style="color:#888;font-size:11px">${a.empCode}</span></td><td style="padding:6px 8px;color:#555">${a.type}</td><td style="padding:6px 8px;color:#555">${a.client} / ${a.site}</td><td style="padding:6px 8px;color:#555">${a.date}</td><td style="padding:6px 8px;font-weight:700;color:#6B21A8">${a.stage}</td></tr>`).join("");
            const w=window.open("","_blank");
            w.document.write(`<html><head><title>Approval Requests Report</title><style>body{font-family:Arial,sans-serif;margin:24px;color:#1a1a1a}h2{color:#1B3A8A}table{width:100%;border-collapse:collapse}th{background:#1B3A8A;color:#fff;padding:7px 8px;text-align:left;font-size:12px}tr:nth-child(even){background:#f5f5f5}@media print{body{margin:0}}</style></head><body><h2>FIXIT MANPOWER SUPPLY L.L.C</h2><h3>Approval Requests Report${stageFilter?" - "+stageFilter:""}</h3><p style="color:#888;font-size:12px">Generated: ${new Date().toLocaleString()}</p><table><thead><tr><th>Worker</th><th>Request Type</th><th>Client / Site</th><th>Date</th><th>Stage</th></tr></thead><tbody>${rows}</tbody></table></body></html>`);
            w.document.close();setTimeout(()=>w.print(),400);
          }} style={{fontSize:14,fontWeight:700,padding:"6px 14px",borderRadius:8,border:"1.5px solid "+CB,color:CB,background:BBG,cursor:"pointer"}}>Print / Save as PDF</button>
        </div>
        {filtered.map((a,i)=>{
          const stCol=STAGE_COL[a.stage]||{bg:BG,fg:SFT};
          return(
            <div key={a.id} style={{padding:"8px 10px",borderRadius:8,border:"1px solid "+LN,marginBottom:6,background:PNL}}>
              <div style={{display:"flex",justifyContent:"space-between",
                      alignItems:"center",marginBottom:4}}>
                <div style={{fontSize:15,fontWeight:700}}>{a.worker} <span style={{fontSize:13,color:SFT}}>{a.empCode}</span></div>
                <Pill c={a.stage} bg={stCol.bg} fg={stCol.fg}/>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                <span style={{fontSize:13,color:SFT}}>{a.type}</span>
                <span style={{fontSize:13,color:SFT}}>{a.client} / {a.site}</span>
                <span style={{fontSize:13,color:SFT}}>{a.date}</span>
              </div>
            </div>
          );
        })}
        {filtered.length===0&&<div style={{textAlign:"center",color:SFT,fontSize:15,padding:"20px 0"}}>No requests found.</div>}
      </div>
    </>
  );
}


function StatusLegend(){
  const [open,setOpen]=useState(false);
  const legendRef=useRef(null);
  const printLegend=()=>{
    if(!legendRef.current)return;
    const printWindow=window.open("","_blank");
    printWindow.document.write(
      "<html><head><title>Status Legend</title><style>"+
      "body{font-family:Arial,sans-serif;margin:20px;color:#1a1a1a}"+
      "table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:6px 8px;font-size:12px}"+
      "@media print{body{margin:0}}"+
      "</style></head><body>"+legendRef.current.innerHTML+"</body></html>"
    );
    printWindow.document.close();
    printWindow.focus();
    setTimeout(()=>{printWindow.print();},400);
  };
  const sections=[
    {title:"Attendance Status",items:[
      ["P","Present","Worker present on site",GRN,GBG],
      ["A","Absent","Worker absent without leave",RED,RBG],
      ["WO","Weekly Off","Scheduled weekly rest day",PUR,PBG],
      ["PH","Public Holiday","Official public holiday",ORANGE,OBG],
      ["I","Idle","Worker present but no work assigned",YFG,YBG],
      ["MWP","Medical With Paper","Medical leave with certificate",BLU,BBG],
      ["MWOP","Medical W/O Paper","Medical leave without certificate",AMB,ABG],
      ["EL","Emergency Leave","Unplanned personal emergency",RED,RBG],
      ["AL","Annual Leave","Scheduled annual leave",BLU,BBG],
      ["OT","On Training","Attending approved training",GRN,GBG],
      ["AB","Absconder","Worker disappeared without notice",RED,RBG],
      ["RES","Resignation","Resignation or under cancellation",AMB,ABG],
      ["NSF","No Salary No Fine","Excluded from salary and fines",SFT,BG],
      ["REF","Refusal to Work","Worker refused assigned work",RED,RBG],
    ]},
    {title:"Approval Stages",items:[
      ["PA","Pending from Management","Awaiting Admin approval",AMB,ABG],
      ["PSA","Pending from Management","Awaiting Super Admin approval",BLU,BBG],
      ["APR","Approved","Fully approved and applied",GRN,GBG],
      ["REJ","Rejected","Rejected - no change applied",RED,RBG],
    ]},
    {title:"Financial",items:[
      ["ADV","Advance","Salary advance request",ORANGE,OBG],
      ["FIN","Fine Applied","Fine deducted from salary",RED,RBG],
      ["FWR","Fine Waiver","Fine waiver request pending",AMB,ABG],
    ]},
  ];
  return(
    <div style={{border:"1px solid "+LN,borderRadius:12,overflow:"hidden",marginBottom:12}}>
      <div onClick={()=>setOpen(!open)} style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",padding:"12px 16px",background:open?DARK:PNL,cursor:"pointer"}}>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:24,height:24,borderRadius:"50%",background:CB,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
            <span style={{color:"#fff",fontWeight:900,fontSize:16}}>?</span>
          </div>
          <div>
            <div style={{fontSize:16,fontWeight:800,color:open?"#fff":CB}}>Status Legend</div>
            <div style={{fontSize:13,color:open?"rgba(255,255,255,0.6)":SFT}}>All abbreviations used in this system</div>
          </div>
        </div>
        <span style={{fontSize:21,color:open?"#fff":CB,fontWeight:700}}>{open?"-":"+"}</span>
      </div>
      {open&&(
        <div ref={legendRef} style={{padding:"12px 16px",background:BG}}>
          {sections.map((sec,si)=>(
            <div key={si} style={{marginBottom:si<sections.length-1?14:0}}>
              <div style={{fontSize:14,fontWeight:800,color:CB,textTransform:"uppercase",letterSpacing:.5,marginBottom:8,paddingBottom:4,borderBottom:"2px solid "+LN}}>{sec.title}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
                {sec.items.map(([abbr,name,desc,fg,bg])=>(
                  <div key={abbr} style={{background:PNL,border:"1px solid "+LN,borderRadius:8,padding:"6px 8px",display:"flex",gap:8,alignItems:"flex-start"}}>
                    <span style={{background:bg,color:fg,fontWeight:800,fontSize:13,padding:"2px 6px",borderRadius:6,flexShrink:0,minWidth:32,textAlign:"center"}}>{abbr}</span>
                    <div>
                      <div style={{fontSize:14,fontWeight:700,lineHeight:1.2}}>{name}</div>
                      <div style={{fontSize:12,color:SFT,lineHeight:1.3,marginTop:1}}>{desc}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}
          <div style={{marginTop:14,borderTop:"2px solid "+LN,paddingTop:12}}>
            <div style={{fontSize:14,fontWeight:800,color:CB,textTransform:"uppercase",letterSpacing:.5,marginBottom:10}}>Approval Authority -- Status by Status</div>
            {[
              {title:"Supervisor Approves Directly",color:GRN,bg:GBG,items:[
                ["P","Present","Mark directly -- no approval needed"],
                ["A","Absent","Mark directly -- no approval needed"],
                ["WO","Weekly Off","Mark directly -- no approval needed"],
                ["PH","Public Holiday","Mark directly -- no approval needed"],
                ["REF","Refusal to Work","Mark directly -- triggers fine recommendation"],
              ]},
              {title:"Admin Approves Only",color:BLU,bg:BBG,items:[
                ["EL","Emergency Leave","Supervisor submits -- Admin approves or rejects"],
                ["AB","Absconder","Supervisor submits -- Admin approves or rejects"],
                ["AL","Annual Leave","Supervisor submits -- Admin approves or rejects"],
                ["OT","On Training","Supervisor submits -- Admin approves or rejects"],
                ["ADV","Advance AED 200 or below","Supervisor raises -- Admin approves or rejects"],
                ["GO","Grade Override","Supervisor suggests -- Admin approves or rejects"],
                ["FIN","Fine Recommendation","Supervisor recommends -- Admin applies or escalates waiver"],
              ]},
              {title:"Super Admin Approves Only -- No Admin step",color:PUR,bg:PBG,items:[
                ["MWP","Medical With Paper","Supervisor submits -- Super Admin approves directly"],
                ["MWOP","Medical Without Paper","Supervisor submits -- Super Admin approves directly"],
                ["I","Idle","Supervisor submits with reason -- Super Admin approves directly"],
                ["RES","Resignation","Supervisor submits -- Super Admin approves directly"],
                ["NSF","No Salary No Fine","Supervisor submits -- Super Admin approves directly"],
                ["FWR","Fine Waiver Request","Admin escalates -- Super Admin approves or rejects"],
                ["ADV","Advance above AED 200","Supervisor raises -- Super Admin approves directly"],
              ]},
              {title:"Admin Sets -- Super Admin Approves",color:ORANGE,bg:OBG,items:[
                ["SIA","Site Incentive Allowance","Admin sets value -- Super Admin must approve before it takes effect"],
              ]},
              {title:"No Approval Needed",color:GRN,bg:GBG,items:[
                ["CTS","Client Timesheet","Supervisor enters and publishes directly -- no approval chain"],
                ["SAL","Salary Upload","Finance team uploads directly -- visible to worker immediately"],
              ]},
            ].map((sec,si)=>(
              <div key={si} style={{marginBottom:12}}>
                <div style={{fontSize:13,fontWeight:800,color:sec.color,background:sec.bg,padding:"4px 8px",borderRadius:6,marginBottom:6,display:"inline-block"}}>{sec.title}</div>
                {sec.items.map(([abbr,name,desc])=>(
                  <div key={abbr} style={{display:"flex",gap:8,alignItems:"flex-start",padding:"4px 0",borderBottom:"1px solid "+LN}}>
                    <span style={{background:sec.bg,color:sec.color,fontWeight:800,fontSize:12,padding:"2px 6px",borderRadius:5,flexShrink:0,minWidth:32,textAlign:"center"}}>{abbr}</span>
                    <div>
                      <span style={{fontSize:13,fontWeight:700,color:INK}}>{name}</span>
                      <span style={{fontSize:12,color:SFT}}> -- {desc}</span>
                    </div>
                  </div>
                ))}
              </div>
            ))}
          </div>
          <div style={{display:"flex",justifyContent:"flex-end",marginTop:14}}>
            <button
              onClick={printLegend} style={{fontSize:14,fontWeight:700,padding:"6px 14px",borderRadius:8,border:"1px solid "+LN,color:CB,background:PNL,cursor:"pointer"}}>Print / Save as PDF</button>
          </div>
        </div>
      )}
    </div>
  );
}


function ReportTab({sup,inactive,medPending,advPending}){
  const [kpiMonth,setKpiMonth]=useState("Jun 2026");
  const KPI_DATA=buildKPI(kpiMonth);
  const clientsForMonth=(CLIENTS_DATA_BY_MONTH[kpiMonth]||CLIENTS_DATA_BY_MONTH["Jun 2026"]).filter(c=>sup.clients.includes(c.name));
  const kpiRef=useRef(null);
  const casRef=useRef(null);
  const printSection=(ref,title)=>{
    if(!ref.current)return;
    const printWindow=window.open("","_blank");
    printWindow.document.write(
      "<html><head><title>"+title+"</title><style>"+
      "body{font-family:Arial,sans-serif;margin:20px;color:#1a1a1a}"+
      "table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:6px 8px;font-size:12px}"+
      "@media print{body{margin:0}}"+
      "</style></head><body>"+ref.current.innerHTML+"</body></html>"
    );
    printWindow.document.close();
    printWindow.focus();
    setTimeout(()=>{printWindow.print();},400);
  };
  const myW=sup.workers;
  const pCount=myW.filter(w=>w.status==="P").length;
  const aCount=myW.filter(w=>w.status==="A").length;
  const mCount=myW.filter(w=>w.status==="M").length;
  const iCount=myW.filter(w=>w.status==="I").length;
  const woCount=myW.filter(w=>w.status==="WO").length;
  const compResolved=3,compInProgress=2;
  const inactiveTotal=Object.values(inactive).reduce((s,v)=>s+v,0);
  const overall=Math.round(KPI_DATA.reduce((s,k)=>s+k.score*k.weight/100,0));
  return(
    <>
      <div ref={kpiRef} style={{background:PNL,border:"2px solid "+CB,borderLeft:"5px solid "+DARK,borderRadius:12,padding:"14px 16px 10px",marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:10}}>
          <div>
            <div style={{color:DARK,fontSize:14,fontWeight:800,letterSpacing:.3}}>KPI SCORECARD (MD = Man Days)</div>
            <select value={kpiMonth}
              onChange={e=>setKpiMonth(e.target.value)} style={{fontSize:13,fontWeight:700,padding:"3px 8px",borderRadius:6,border:"1px solid "+LN,background:PNL,color:DARK,marginTop:4,cursor:"pointer"}}>
              {["Jun 2026","May 2026","Apr 2026"].map(m=><option key={m}>{m}</option>)}
            </select>
          </div>
          <div><span style={{fontSize:31,fontWeight:900,color:overall>=90?GRN:overall>=80?CY:RED}}>{overall}</span><span style={{fontSize:15,color:SFT}}>/100</span></div>
        </div>
        {KPI_DATA.map((k,i)=>(
          <div key={i} style={{marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
              <span style={{fontSize:13,color:DARK,fontWeight:600}}>{k.label} <span style={{color:SFT}}>({k.weight}%)</span></span>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <span style={{fontSize:12,color:SFT}}>Target: {k.target}%</span>
                <span style={{fontSize:14,fontWeight:800,color:k.col}}>{k.score}%</span>
              </div>
            </div>
            <div style={{background:"#E8EDF5",borderRadius:4,height:7}}>
              <div style={{width:k.score+"%",height:"100%",background:k.col,borderRadius:4}}/>
            </div>
          </div>
        ))}
        <div style={{display:"flex",justifyContent:"flex-end",marginTop:8}}>
          <button
            onClick={()=>printSection(kpiRef,"KPI Scorecard - "+kpiMonth)} style={{fontSize:14,fontWeight:700,padding:"6px 14px",borderRadius:8,border:"1.5px solid "+CB,color:CB,background:BBG,cursor:"pointer"}}>Print / Save as PDF</button>
        </div>
      </div>
      <Card s={{padding:0}}>
      <div ref={casRef} style={{padding:16,background:"#fff"}}>
        <SHdr title="Client Attendance Summary" sub={kpiMonth}/>
        {clientsForMonth.map((cl,ci)=>(
          <div key={ci} style={{marginBottom:10}}>
            <div style={{background:BBG,borderRadius:8,padding:"5px 10px",marginBottom:6,fontSize:15,fontWeight:800,color:BLU}}>{cl.name}</div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:14}}>
              <thead><tr
                style={{background:DARK}}>{["Site","Total","P","A","M","I","P%"].map(h=><th key={h} style={{color:"#fff",padding:"4px 6px",fontWeight:700,textAlign:"left",fontSize:13}}>{h}</th>)}</tr></thead>
              <tbody>{cl.sites.map((s,si)=>{const pct=Math.round(s.present/s.total*100),low=pct<95;return(
                <tr key={si} style={{background:si%2?BG:PNL}}>
                  <td style={{padding:"4px 6px",color:SFT,fontSize:13}}>{s.site}</td>
                  <td style={{padding:"4px 6px",fontWeight:700}}>{s.total}</td>
                  <td style={{padding:"4px 6px",color:GRN,fontWeight:700}}>{s.present}</td>
                  <td style={{padding:"4px 6px",color:RED,fontWeight:700}}>{s.absent}</td>
                  <td style={{padding:"4px 6px",color:BLU}}>{s.medical}</td>
                  <td style={{padding:"4px 6px",color:YFG}}>{s.idle}</td>
                  <td style={{padding:"4px 6px",fontWeight:800,color:low?RED:GRN}}>{pct}%{low&&" !"}</td>
                </tr>
              );})}</tbody>
            </table>
          </div>
        ))}
        <div style={{fontSize:13,color:SFT,fontStyle:"italic"}}>Sites below 95% are marked with !</div>
      </div>
      <div style={{display:"flex",justifyContent:"flex-end",padding:"0 16px 14px"}}>
        <button
          onClick={()=>printSection(casRef,"Client Attendance Summary - "+kpiMonth)} style={{fontSize:14,fontWeight:700,padding:"6px 14px",borderRadius:8,border:"1px solid "+LN,color:CB,background:BG,cursor:"pointer"}}>Print / Save as PDF</button>
      </div>
      </Card>
      <StatusLegend/>
    </>
  );
}

/* ATTENDANCE TAB */
function AttendanceTab({sup,gradeOverrides={},inactive={emergencyLeave:0,annualLeave:0,absconder:0,refusal:0,resignation:0,training:0}}){
  const [search,setSearch]=useState("");
  const [searchBy,setSearchBy]=useState("name");
  const [statusFilter,setStatusFilter]=useState("All");
  const [viewMode,setViewMode]=useState("today");
  const [selDate,setSelDate]=useState(null);
  const [showCal,setShowCal]=useState(false);
  const [selected,setSelected]=useState([]);
  const [expandedWorker,setExpandedWorker]=useState(null);
  const [showExport,setShowExport]=useState(false);
  const activeRef=useRef(null);
  const inactiveRef=useRef(null);
  const printSection=(ref,title)=>{
    if(!ref.current)return;
    const printWindow=window.open("","_blank");
    printWindow.document.write(
      "<html><head><title>"+title+"</title><style>"+
      "body{font-family:Arial,sans-serif;margin:20px;color:#1a1a1a}"+
      "table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:6px 8px;font-size:12px}"+
      "@media print{body{margin:0}}"+
      "</style></head><body>"+ref.current.innerHTML+"</body></html>"
    );
    printWindow.document.close();
    printWindow.focus();
    setTimeout(()=>{printWindow.print();},400);
  };
  const myW=sup.workers;
  const filtered=myW.filter(w=>{
    const q=search.toLowerCase();
    const sq=!search||(searchBy==="name"?w.name.toLowerCase().includes(q):searchBy==="code"?w.empCode.toLowerCase().includes(q):searchBy==="client"?w.client.toLowerCase().includes(q):w.site.toLowerCase().includes(q));
    const sf=statusFilter==="All"||w.status===statusFilter.slice(0,1);
    return sq&&sf;
  });
  const toggleW=(id)=>setSelected(s=>s.includes(id)?s.filter(x=>x!==id):[...s,id]);
  const visaBadge=(e)=>{const d=new Date(e),diff=Math.round((d-today)/86400000);if(diff<0)return{l:"Expired",bg:RBG,fg:RED};if(diff<90)return{l:"Expires "+diff+"d",bg:ABG,fg:AMB};return{l:e,bg:GBG,fg:GRN};};
  const myWAll=sup.workers;
  const pCount=myWAll.filter(w=>w.status==="P").length;
  const aCount=myWAll.filter(w=>w.status==="A").length;
  const mCount=myWAll.filter(w=>w.status==="M").length;
  const iCount=myWAll.filter(w=>w.status==="I").length;
  const woCount=myWAll.filter(w=>w.status==="WO").length;
  const inactiveTotal=Object.values(inactive).reduce((s,v)=>s+v,0);
  const MO_NAMES=["Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"];
  const MONTHS=(()=>{
    const arr=[];
    for(let k=0;k<5;k++){
      const d=new Date(today.getFullYear(),today.getMonth()-k,1);
      arr.push(MO_NAMES[d.getMonth()]+" "+d.getFullYear());
    }
    return arr;
  })();
  const yesterday=new Date(today);yesterday.setDate(today.getDate()-1);
  const viewLabel=viewMode==="today"?"Today":viewMode==="yesterday"?yesterday.toDateString():selDate?viewMode+" - Day "+selDate:viewMode;

  const MonthCal=({monthLabel})=>{
    const idx=MONTHS.indexOf(monthLabel);
    const yr=today.getFullYear(),mo=today.getMonth()-idx;
    const dim=new Date(yr,mo+1,0).getDate();
    const fd=new Date(yr,mo,1).getDay();
    return(
      <div style={{background:PNL,border:"2px solid "+CB,borderRadius:12,padding:"12px",marginBottom:10}}>
        <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
          <span style={{fontSize:15,fontWeight:800,color:CB}}>{monthLabel} - Select a date</span>
          <button onClick={()=>setShowCal(false)} style={{fontSize:14,color:SFT,background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Close</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2,marginBottom:2}}>
          {["S","M","T","W","T","F","S"].map((d,i)=><div key={i} style={{textAlign:"center",fontSize:12,fontWeight:700,color:SFT}}>{d}</div>)}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
          {Array.from({length:fd},(_,i)=><div key={"b"+i}/>)}
          {Array.from({length:dim},(_,i)=>{
            const d=i+1,isSun=new Date(yr,mo,d).getDay()===0,isSel=selDate===d&&viewMode===monthLabel;
            return <div key={d}
              onClick={()=>{setSelDate(d);setViewMode(monthLabel);setShowCal(false);}} style={{borderRadius:5,display:"flex",alignItems:"center",justifyContent:"center",minHeight:26,cursor:"pointer",background:isSel?CB:isSun?PBG:"#F4F6FB",color:isSel?"#fff":isSun?PUR:INK,fontWeight:isSel?700:400,fontSize:13}}>{d}</div>;
          })}
        </div>
      </div>
    );
  };

  return(
    <>

      <Card s={{marginBottom:12,padding:0,position:"relative"}}>
      <div ref={activeRef} style={{padding:16,background:"#fff"}}>
        <SHdr title="Active Workers" sub="From worker daily reporting"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
          <Block label="Total Active" val={myW.length} fg={CB} bg={BBG}/>
          <Block label="Present" val={pCount} fg={GRN} bg={GBG}/>
          <Block label="Absent" val={aCount} fg={RED} bg={RBG}/>
          <Block label="Medical Leave" val={mCount} fg={BLU} bg={BBG}/>
          <Block label="Idle" val={iCount} fg={YFG} bg={YBG}/>
          <Block label="Weekly Off" val={woCount} fg={PUR} bg={PBG}/>
          <Block label="Public Holiday" val={0} fg={ORANGE} bg={OBG}/>
        </div>
      </div>
      <div style={{display:"flex",justifyContent:"flex-end",padding:"0 12px 12px"}}>
        <button
          onClick={()=>printSection(activeRef,"Active Workers")} style={{fontSize:14,fontWeight:700,padding:"6px 14px",borderRadius:8,border:"none",color:"#fff",background:"linear-gradient(135deg,"+DARK+","+CB+")",cursor:"pointer"}}>Print / Save as PDF</button>
      </div>
      </Card>
      <Card s={{marginBottom:12,padding:0,position:"relative"}}>
      <div ref={inactiveRef} style={{padding:16,background:"#fff"}}>
        <SHdr title="Inactive Workers" sub="Updated via Override tab"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
          {[["Emergency Leave","emergencyLeave",RED,RBG],["Annual Leave","annualLeave",BLU,BBG],["Absconder","absconder","#7B0000","#FFE8E8"],["Refusal to Work","refusal","#7A4000",ABG],["Resignation","resignation",AMB,ABG],["On Training","training",GRN,GBG]].map(([l,k,fg,bg])=>(
            <div key={k} style={{background:bg,borderRadius:9,padding:"7px 10px",border:"1px solid "+LN,display:"flex",alignItems:"center",gap:8,opacity:inactive[k]===0?0.45:1}}>
              <span style={{fontSize:21,fontWeight:900,color:fg,minWidth:20,textAlign:"center"}}>{inactive[k]}</span>
              <span style={{fontSize:12,fontWeight:700,color:fg,lineHeight:1.3}}>{l}</span>
            </div>
          ))}
        </div>
        {inactiveTotal===0&&<div style={{fontSize:13,color:SFT,marginTop:8,fontStyle:"italic"}}>All workers currently active. Use Override tab to update these counts.</div>}
      </div>
      <div style={{display:"flex",justifyContent:"flex-end",padding:"0 12px 12px"}}>
        <button
          onClick={()=>printSection(inactiveRef,"Inactive Workers")} style={{fontSize:14,fontWeight:700,padding:"6px 14px",borderRadius:8,border:"none",color:"#fff",background:"linear-gradient(135deg,"+DARK+","+CB+")",cursor:"pointer"}}>Print / Save as PDF</button>
      </div>
      </Card>
      {selected.length>0&&(
        <div style={{background:BBG,borderRadius:10,padding:"10px 12px",marginBottom:10,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:15,fontWeight:700,color:CB}}>{selected.length} selected</span>
          <div style={{display:"flex",gap:8,alignItems:"center"}}><div style={{fontSize:13,color:BLU,fontWeight:600}}>Worker transfers are managed by Admin only.</div><button
            onClick={()=>setSelected([])} style={{fontSize:14,color:SFT,background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Clear</button></div>
        </div>
      )}

      <Card s={{marginBottom:10}}>
        <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap"}}>
          {["today","yesterday",...MONTHS].map((m,mi)=>(
            <button key={m}
              onClick={()=>{setViewMode(m);setSelDate(null);}} style={{fontSize:14,fontWeight:700,padding:"5px 12px",borderRadius:20,cursor:"pointer",border:"2px solid "+(viewMode===m?CB:LN),background:viewMode===m?CB:PNL,color:viewMode===m?"#fff":SFT,whiteSpace:"nowrap",textTransform:"capitalize"}}>{m===MONTHS[0]?m+" (Current)":m}</button>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div>
            <label style={{...lbl,marginTop:0,marginBottom:3,color:CB}}>Client</label>
            <select style={inp} value={searchBy==="client"?search:""} onChange={e=>{setSearchBy("client");setSearch(e.target.value);}}>
              <option value="">All clients</option>
              {[...new Set(myW.map(w=>w.client))].map(c=><option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={{...lbl,marginTop:0,marginBottom:3,color:ORANGE}}>Site</label>
            <select style={inp} value={searchBy==="site"?search:""} onChange={e=>{setSearchBy("site");setSearch(e.target.value);}}>
              <option value="">All sites</option>
              {[...new Set(myW.map(w=>w.site))].map(s=><option key={s}>{s}</option>)}
            </select>
          </div>
        </div>
        <input style={{...inp,marginBottom:8}} value={searchBy==="name"||searchBy==="code"?search:""}
          onChange={e=>{setSearchBy("name");setSearch(e.target.value);}} placeholder="Search by name or employee code..."/>
        <select style={inp} value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}>
          <option>All</option>
          {STATUS_LIST.map(s=><option key={s}>{s}</option>)}
        </select>
      </Card>

      <Card>
        {filtered.length>0&&(
          <div style={{display:"flex",justifyContent:"center",marginBottom:12}}>
            <button
              onClick={()=>setShowExport(!showExport)} style={{fontSize:15,fontWeight:700,padding:"9px 20px",borderRadius:10,border:"1.5px solid "+CB,color:showExport?"#fff":CB,background:showExport?CB:BBG,cursor:"pointer"}}>
              {showExport?"Hide Download Options":"Download Results (Excel / CSV)"}
            </button>
          </div>
        )}
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
          <span style={{fontSize:14,color:SFT}}>{filtered.length} workers - {viewLabel}</span>
          <button onClick={()=>setSelected(filtered.map(w=>w.id))} style={{fontSize:14,fontWeight:700,color:CB,background:"none",border:"none",cursor:"pointer"}}>Select all</button>
        </div>
        {showExport&&(
          <div style={{background:BG,borderRadius:8,padding:"10px",marginBottom:10}}>
            <div style={{fontSize:14,color:SFT,marginBottom:8}}>Exports the {filtered.length} worker(s) currently shown above, using your Client/Site/Status filters.</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>{
                const rows=filtered.flatMap(w=>Object.entries(WORKER_ATT[w.empCode]||{}).map(([d,st])=>({worker:w.name,empCode:w.empCode,client:w.client,site:w.site,day:d,status:st})));
                const hdr="Worker,Code,Client,Site,Day,Status";
                const body=rows.map(r=>r.worker+","+r.empCode+","+r.client+","+r.site+","+r.day+","+r.status).join("\n");
                const blob=new Blob([hdr+"\n"+body],{type:"text/csv"});
                const url=URL.createObjectURL(blob);
                const a=document.createElement("a");a.href=url;a.download="attendance_search.csv";a.click();
              }} style={{...PBtn,margin:0,flex:1,background:GRN}}>Download CSV</button>
              <button onClick={()=>{
                const rows=filtered.flatMap(w=>Object.entries(WORKER_ATT[w.empCode]||{}).map(([d,st])=>({worker:w.name,empCode:w.empCode,client:w.client,site:w.site,day:d,status:st})));
                const wb=XLSX.utils.book_new();
                const header=["Worker","Emp Code","Client","Site","Day","Status"];
                const data=rows.map(r=>[r.worker,r.empCode,r.client,r.site,r.day,r.status]);
                const titleRows=[["FIXIT MANPOWER SUPPLY L.L.C",...Array(5).fill("")],["Attendance Search Export - "+viewLabel,...Array(5).fill("")],Array(6).fill("")];
                const ws=XLSX.utils.aoa_to_sheet([...titleRows,header,...data]);
                ws["!cols"]=[{wch:16},{wch:10},{wch:20},{wch:16},{wch:6},{wch:8}];
                XLSX.utils.book_append_sheet(wb,ws,"Attendance");
                XLSX.writeFile(wb,"attendance_search.xlsx");
              }} style={{...PBtn,margin:0,flex:1,background:"linear-gradient(135deg,"+DARK+","+CB+")"}}>Download Excel</button>
            </div>
            <div style={{fontSize:13,color:SFT,marginTop:6}}>{filtered.length} workers in current search</div>
          </div>
        )}
        {filtered.map((w,i)=>{
          const sc=STATUS_STYLE[w.status]||{bg:BG,fg:SFT};
          const vb=visaBadge(w.visaExpiry);
          return(
            <div key={w.id} style={{padding:"9px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div onClick={()=>setExpandedWorker(expandedWorker===w.id?null:w.id)} style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer"}}>
                <div style={{flex:1}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{width:7,height:7,borderRadius:4,background:sc.fg,display:"inline-block"}}/>
                    <span style={{fontSize:16,fontWeight:700}}>{w.name}</span>
                  </div>
                  <div style={{fontSize:13,color:SFT,marginTop:2}}>{w.empCode} | {w.trade}</div>
                  <div style={{fontSize:13,color:SFT}}>{w.client} | {w.site}</div>
                  <div style={{fontSize:13,color:SFT}}>Passport: {w.passport} | {monthsWorked(w.stampDate)} months | {w.siteChanges} site changes</div>
                  <span style={{fontSize:12,fontWeight:700,padding:"2px 7px",borderRadius:8,background:vb.bg,color:vb.fg,marginTop:3,display:"inline-block"}}>Visa: {vb.l}</span>
                </div>
                <span style={{fontSize:13,fontWeight:700,padding:"3px 9px",borderRadius:10,background:sc.bg,color:sc.fg,whiteSpace:"nowrap"}}>{w.status}</span>
              </div>
              {expandedWorker===w.id&&(()=>{
                const watt=WORKER_ATT[w.empCode]||{};
                const moLabel=["today","yesterday"].includes(viewMode)?MONTHS[0]:viewMode.replace(/ - Day \d+/,"");
                const moIdx=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(moLabel.split(" ")[0]);
                const yr=Number(moLabel.split(" ")[1]);
                const dim=moIdx>=0&&yr?new Date(yr,moIdx+1,0).getDate():30;
                const fd=moIdx>=0&&yr?new Date(yr,moIdx,1).getDay():0;
                return(
                  <div style={{background:BG,borderRadius:8,padding:"8px",marginTop:6,border:"1px solid "+LN}}>
                    <div style={{fontSize:13,fontWeight:700,color:CB,marginBottom:6}}>{moLabel} Attendance</div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2,marginBottom:2}}>
                      {["Su","Mo","Tu","We","Th","Fr","Sa"].map(d=><div key={d} style={{textAlign:"center",fontSize:10,fontWeight:700,color:SFT}}>{d}</div>)}
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
                      {Array.from({length:fd},(_,i)=><div key={"b"+i}/>)}
                      {Array.from({length:dim},(_,i)=>{
                        const d=i+1;const st=watt[d];const dsc=STATUS_COLS[st]||{bg:"#eee",fg:"#999"};
                        return(<div key={d}
                          style={{background:dsc.bg,borderRadius:4,padding:"2px 0",textAlign:"center",minHeight:22}}><div style={{fontSize:10,color:dsc.fg}}>{d}</div><div style={{fontSize:10,fontWeight:800,color:dsc.fg}}>{st||""}</div></div>);
                      })}
                    </div>
                  </div>
                );
              })()}
            </div>        );
        })}
      </Card>
      {showCal&&<MonthCal monthLabel={viewMode}/>}
    </>
  );
}

/* WORKER PROFILE TAB */
function WorkerCalendarTab({sup,gradeOverrides}){
  const [search,setSearch]=useState("");
  const [selWorker,setSelWorker]=useState(null);
  const [showDrop,setShowDrop]=useState(false);
  const [rClient,setRClient]=useState("");
  const [rSite,setRSite]=useState("");
  const [selMonth,setSelMonth]=useState("Jun 2026");
  const [clientEntries,setClientEntries]=useState({});
  const [tsStatus,setTsStatus]=useState("draft");
  const [selDates,setSelDates]=useState([]);
  const [bulkStatus,setBulkStatus]=useState("P");
  const [bulkHrs,setBulkHrs]=useState(8);
  const [bulkOT,setBulkOT]=useState(2);
  const [showOverrideForm,setShowOverrideForm]=useState(false);
  const [overrideForm,setOverrideForm]=useState({status:"MWP",reason:"",file:null});
  const [overrides,setOverrides]=useState([]);
  const [saveConfirmed,setSaveConfirmed]=useState(false);
  const [rptMonth,setRptMonth]=useState("Jun 2026");
  const [rptClient,setRptClient]=useState("");
  const [rptSite,setRptSite]=useState("");
  const [showRptPreview,setShowRptPreview]=useState(false);
  const myW=sup.workers;
  const clients=[...new Set(myW.map(w=>w.client))];
  const sites=[...new Set(myW.filter(w=>!rClient||w.client===rClient).map(w=>w.site))];
  const searchRes=showDrop?myW.filter(w=>(!rClient||w.client===rClient)&&(!rSite||w.site===rSite)&&search.length>0&&(w.name.toLowerCase().includes(search.toLowerCase())||w.empCode.toLowerCase().includes(search.toLowerCase()))):[];
  const MONTHS_LIST=[{lb:"Jun 2026"},{lb:"May 2026"},{lb:"Apr 2026"},{lb:"Mar 2026"}];
  const mo=MONTHS_LIST.find(m=>m.lb===selMonth)||MONTHS_LIST[0];
  const moIdx=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(selMonth.split(" ")[0]);
  const yr=Number(selMonth.split(" ")[1]);
  const dim=moIdx>=0&&yr?new Date(yr,moIdx+1,0).getDate():30;
  const firstDay=moIdx>=0&&yr?new Date(yr,moIdx,1).getDay():0;
  const wDays=selWorker?(WORKER_ATT[selWorker.empCode]||{}):{};
  const entKey=selWorker?selWorker.empCode+"_"+selMonth:"";
  const dayEntries=clientEntries[entKey]||{};
  const isPublished=tsStatus==="published";
  const isApproved=tsStatus==="approved"||isPublished;
  const DIRECT=["P","A"];
  const pendingOverrides=overrides.filter(o=>o.approvalStatus!=="Approved"&&o.approvalStatus!=="Rejected");
  const totalBasic=Object.values(dayEntries).reduce((s,e)=>s+(e.hrs||0),0);
  const totalOT=Object.values(dayEntries).reduce((s,e)=>s+(e.otHrs||0),0);
  const sc2Status={draft:{bg:ABG,fg:AMB,label:"Draft"},pending_admin:{bg:BBG,fg:BLU,label:"Pending Approval"},approved:{bg:GBG,fg:GRN,label:"Approved"},published:{bg:GBG,fg:GRN,label:"Published"}};
  const sc2=sc2Status[tsStatus]||sc2Status.draft;
  const s1vals=Array.from({length:dim},(_,i)=>wDays[i+1]).filter(Boolean);
  const s1pC=s1vals.filter(v=>v==="P").length;
  const s1aC=s1vals.filter(v=>v==="A").length;
  const s1idleC=s1vals.filter(v=>v==="I").length;
  const s1medC=s1vals.filter(v=>v==="MWP"||v==="MWOP"||v==="M").length;
  const s2vals=Object.values(dayEntries);
  const s2pC=s2vals.filter(e=>e.status==="P").length;
  const s2aC=s2vals.filter(e=>e.status==="A").length;
  const s2idleC=s2vals.filter(e=>e.status==="I").length;
  const s2medC=s2vals.filter(e=>e.status==="MWP"||e.status==="MWOP").length;
  const s2woC=s2vals.filter(e=>e.status==="WO").length;
  const s2phC=s2vals.filter(e=>e.status==="PH").length;
  const rptClients=[...new Set(myW.map(w=>w.client))];
  const rptSites=[...new Set(myW.filter(w=>!rptClient||w.client===rptClient).map(w=>w.site))];
  const rptWorkers=myW.filter(w=>(!rptClient||w.client===rptClient)&&(!rptSite||w.site===rptSite));
  const rptMoIdx=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(rptMonth.split(" ")[0]);
  const rptYr=Number(rptMonth.split(" ")[1]);
  const rptDim=rptMoIdx>=0&&rptYr?new Date(rptYr,rptMoIdx+1,0).getDate():30;

  const STATUS_HRS={"P":{hrs:8,otHrs:2},"A":{hrs:0,otHrs:0},"WO":{hrs:0,otHrs:0},"PH":{hrs:8,otHrs:0},"REF":{hrs:0,otHrs:0},"MWP":{hrs:8,otHrs:0},"MWOP":{hrs:8,otHrs:0},"I":{hrs:8,otHrs:0},"EL":{hrs:0,otHrs:0},"AL":{hrs:0,otHrs:0},"OT":{hrs:8,otHrs:0},"AB":{hrs:0,otHrs:0},"RES":{hrs:0,otHrs:0},"NSF":{hrs:0,otHrs:0}};
  const getDefHrs=(st)=>STATUS_HRS[st]||{hrs:0,otHrs:0};
  const toggleSel=(d)=>{if(isPublished)return;setSelDates(s=>s.includes(d)?s.filter(x=>x!==d):[...s,d]);};
  const applyDirect=(status)=>{
    const {hrs,otHrs}=getDefHrs(status);
    const filled={...dayEntries};
    selDates.forEach(d=>{filled[d]={status,hrs,otHrs,override:false};});
    setClientEntries({...clientEntries,[entKey]:filled});setSelDates([]);
  };
  const applyBulk=()=>{
    const filled={...dayEntries};
    selDates.forEach(d=>{filled[d]={status:bulkStatus,hrs:bulkHrs,otHrs:bulkOT,override:!DIRECT.includes(bulkStatus)};});
    setClientEntries({...clientEntries,[entKey]:filled});setSelDates([]);
  };
  const requestOverride=()=>{
    if(selDates.length===0||!overrideForm.reason)return;
    const newOv=selDates.map(d=>({id:Date.now()+d,day:d,month:selMonth,worker:selWorker.name,empCode:selWorker.empCode,status:bulkStatus,hrs:bulkHrs,otHrs:bulkOT,reason:overrideForm.reason,file:(overrideForm.file?overrideForm.file.name:null),approvalStatus:"Pending Super Admin"}));
    setOverrides([...overrides,...newOv]);
    const filled={...dayEntries};
    selDates.forEach(d=>{filled[d]={status:bulkStatus,hrs:bulkHrs,otHrs:bulkOT,override:true,overridePending:true};});
    setClientEntries({...clientEntries,[entKey]:filled});
    setSelDates([]);setShowOverrideForm(false);
    setOverrideForm({status:"MWP",reason:"",file:null});
  };
   const fillFromSection1=()=>{
    if(isPublished)return;
    const filled={...dayEntries};
    for(let d=1;d<=dim;d++){
      const wSt=wDays[d];if(!wSt)continue;
      const {hrs,otHrs}=getDefHrs(wSt);
      if(DIRECT.includes(wSt)){filled[d]={status:wSt,hrs,otHrs,override:false};}
      else if(!filled[d]){filled[d]={status:wSt,hrs,otHrs,override:true,overridePending:true};}
    }
    setClientEntries({...clientEntries,[entKey]:filled});
  };
  const saveEntries=()=>{setSaveConfirmed(true);setTimeout(()=>setSaveConfirmed(false),2500);};
  const exportRptExcel=()=>{
    const rows=[["Employee Name","Employee Code","Client","Site","Date","Status","Basic Hrs","OT Hrs","Total Hrs"]];
    rptWorkers.forEach(w=>{
      const key=w.empCode+"_"+rptMonth;
      const entries=clientEntries[key]||{};
      for(let d=1;d<=rptDim;d++){
        const e=entries[d];if(!e)continue;
        rows.push([w.name,w.empCode,w.client,w.site,String(d).padStart(2,"0")+"/"+rptMonth,e.status||"",e.hrs||0,e.otHrs||0,(e.hrs||0)+(e.otHrs||0)]);
      }
    });
    const wb=XLSX.utils.book_new();
    const titleRows=[["FIXIT MANPOWER SUPPLY L.L.C","","","","","","","",""],["Client Timesheet Report - "+rptMonth,"","","","","","","",""],["Client: "+(rptClient||"All")+" | Site: "+(rptSite||"All"),"","","","","","","",""],Array(9).fill("")];
    const ws=XLSX.utils.aoa_to_sheet([...titleRows,rows[0],...rows.slice(1)]);
    ws["!cols"]=[{wch:18},{wch:12},{wch:20},{wch:16},{wch:10},{wch:8},{wch:10},{wch:8},{wch:10}];
    XLSX.utils.book_append_sheet(wb,ws,"Timesheet");
    XLSX.writeFile(wb,"Timesheet_"+(rptSite||rptClient||"All").split(" ").join("_")+"_"+rptMonth.split(" ").join("_")+".xlsx");
  };

  return(
    <>
      <Card s={{marginBottom:10}}>
        <SHdr title="Worker Attendance & Timesheet" sub="Select client/site to see workers, or search by name"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div>
            <label style={{...lbl,marginTop:0,color:CB}}>Client</label>
            <select style={inp} value={rClient} onChange={e=>{setRClient(e.target.value);setRSite("");setSelWorker(null);}}>
              <option value="">All clients</option>
              {clients.map(c=><option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={{...lbl,marginTop:0,color:ORANGE}}>Site</label>
            <select style={inp} value={rSite} onChange={e=>{setRSite(e.target.value);setSelWorker(null);}}>
              <option value="">All sites</option>
              {sites.map(s=><option key={s}>{s}</option>)}
            </select>
          </div>
        </div>
        <input style={inp} value={search}
          onChange={e=>{setSearch(e.target.value);setShowDrop(true);if(!e.target.value){setShowDrop(false);}}} onFocus={()=>setShowDrop(true)} placeholder="Or search by name or employee code..."/>
        {searchRes.map((w,i)=>(
          <div key={w.id}
            onClick={()=>{setSelWorker(w);setSearch(w.name+" ("+w.empCode+")");setShowDrop(false);}} style={{display:"flex",justifyContent:"space-between",
              alignItems:"center",padding:"7px 0",
              borderTop:i?"1px solid "+LN:"none",
              cursor:"pointer",marginTop:4}}>
            <div><div style={{fontSize:16,fontWeight:700}}>{w.name}</div><div style={{fontSize:14,color:SFT}}>{w.empCode} | {w.trade} | {w.site}</div></div>
            <Pill c={w.payType==="hourly"?"Hourly":"Fixed"} bg={w.payType==="hourly"?GBG:BBG} fg={w.payType==="hourly"?GRN:BLU}/>
          </div>
        ))}
        {(rClient||rSite)&&(()=>{
          const listed=myW.filter(w=>(!rClient||w.client===rClient)&&(!rSite||w.site===rSite));
          if(listed.length===0) return null;
          return(
            <div style={{marginTop:10,border:"1px solid "+LN,borderRadius:10,overflow:"hidden"}}>
              <div style={{background:BBG,padding:"8px 12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{fontSize:14,fontWeight:700,color:CB}}>{listed.length} worker(s) - {rSite||rClient}</div>
                <div style={{fontSize:13,color:SFT}}>Tap to open attendance</div>
              </div>
              <div style={{maxHeight:280,overflowY:"auto"}}>
                {listed.map((w,i)=>{
                  const sc=STATUS_COLS[w.status]||{bg:"#eee",fg:"#888"};
                  const isActive=selWorker&&selWorker.id===w.id;
                  return(
                    <div key={w.id}
                      onClick={()=>{setSelWorker(w);setSearch("");setShowDrop(false);}} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderTop:i?"1px solid "+LN:"none",cursor:"pointer",background:isActive?"linear-gradient(135deg,"+DARK+","+CB+")":PNL,transition:"background 0.15s"}}>
                      <div style={{width:8,height:8,borderRadius:"50%",background:isActive?"#fff":sc.fg,flexShrink:0}}/>
                      <div style={{flex:1}}>
                        <div style={{fontSize:15,fontWeight:700,color:isActive?"#fff":INK}}>{w.name}</div>
                        <div style={{fontSize:13,color:isActive?"rgba(255,255,255,0.7)":SFT}}>{w.empCode} | {w.trade}</div>
                      </div>
                      <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:3}}>
                        <span style={{fontSize:12,fontWeight:800,padding:"2px 7px",borderRadius:6,background:isActive?"rgba(255,255,255,0.2)":sc.bg,color:isActive?"#fff":sc.fg}}>{w.status}</span>
                        <span style={{fontSize:11,color:isActive?"rgba(255,255,255,0.6)":SFT}}>{w.site}</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })()}
      </Card>

      {selWorker&&(
        <>
          <Card s={{marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center"}}>
              <div style={{flex:1}}>
                <label style={{...lbl,marginTop:0}}>Month</label>
                <select style={{...inp,marginBottom:0}} value={selMonth} onChange={e=>{setSelMonth(e.target.value);setSelDates([]);}}>
                  {MONTHS_LIST.map(m=><option key={m.lb}>{m.lb}</option>)}
                </select>
              </div>
              <div style={{marginLeft:12,marginTop:16}}><Pill c={sc2.label} bg={sc2.bg} fg={sc2.fg}/></div>
            </div>
          </Card>

          <Card s={{marginBottom:10}}>
            <SHdr title="Section 1 - Worker Reported Attendance"/>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginBottom:2}}>
              {["Su","Mo","Tu","We","Th","Fr","Sa"].map(d=><div key={d} style={{textAlign:"center",fontSize:11,fontWeight:700,color:SFT}}>{d}</div>)}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3}}>
              {Array.from({length:firstDay},(_,i)=><div key={"b"+i}/>)}
              {Array.from({length:dim},(_,i)=>{
                const d=i+1;const st=wDays[d];const isPendS1=st&&!["P","A","WO","PH","REF"].includes(st);const sc=isPendS1?{bg:"#FFF8E1",fg:AMB}:(STATUS_COLS[st]||{bg:"#eee",fg:"#888"});
                return(<div key={d} style={{background:sc.bg,borderRadius:5,
                                      padding:"4px 2px",textAlign:"center",minHeight:26,position:"relative"}}>
                  <div style={{fontSize:11,color:sc.fg}}>{d}</div>
                  <div style={{fontSize:12,fontWeight:800,color:sc.fg}}>{st||""}</div>
                  {isPendS1&&<div style={{position:"absolute",top:2,right:2,width:5,height:5,borderRadius:"50%",background:RED}}/>}
                </div>);
              })}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:5,marginTop:10}}>
              <div
                style={{background:GBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:GRN}}>{s1pC}</div><div style={{fontSize:11,color:GRN}}>P</div></div>
              <div
                style={{background:RBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:RED}}>{s1aC}</div><div style={{fontSize:11,color:RED}}>A</div></div>
              <div
                style={{background:YBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:YFG}}>{s1idleC}</div><div style={{fontSize:11,color:YFG}}>Idle</div></div>
              <div
                style={{background:BBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:BLU}}>{s1medC}</div><div style={{fontSize:11,color:BLU}}>Medical</div></div>
            </div>
          </Card>

          <Card s={{marginBottom:10,border:"2px solid "+GRN}}>
            <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
              <SHdr title="Section 2 - Client Timesheet Entry" sub="Tap days to select. P/A direct. Others need override."/>
              {!isPublished&&<button
                onClick={fillFromSection1} style={{fontSize:14,fontWeight:700,padding:"5px 12px",borderRadius:8,border:"1.5px solid "+CB,color:CB,background:BBG,cursor:"pointer"}}>Auto-fill from Section 1</button>}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginBottom:2}}>
              {["Su","Mo","Tu","We","Th","Fr","Sa"].map(d=><div key={d} style={{textAlign:"center",fontSize:11,fontWeight:700,color:SFT}}>{d}</div>)}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginBottom:8}}>
              {Array.from({length:firstDay},(_,i)=><div key={"b"+i}/>)}
              {Array.from({length:dim},(_,i)=>{
                const d=i+1;
                const entry=dayEntries[d];
                const st=entry?entry.status:"";
                const hrs=entry?entry.hrs:0;
                const otHrs=entry?entry.otHrs:0;
                const totalH=hrs+otHrs;
                const isSel=selDates.includes(d);
                const isPending=entry&&entry.overridePending;
                const sc=isPending?{bg:"#FFF8E1",fg:AMB}:st?(STATUS_COLS[st]||{bg:"#eee",fg:"#888"}):{bg:"#F5F5F5",fg:"#CCC"};
                return(
                  <div key={d}
                    onClick={()=>toggleSel(d)} style={{background:isSel?CB:sc.bg,borderRadius:5,padding:"3px 1px",textAlign:"center",minHeight:50,cursor:isPublished?"default":"pointer",border:isSel?"2px solid "+DARK:"1px solid "+LN,position:"relative"}}>
                    <div style={{fontSize:12,fontWeight:700,color:isSel?"#fff":sc.fg}}>{d}</div>
                    {st&&<div style={{fontSize:12,fontWeight:800,color:isSel?"#fff":sc.fg}}>{st}</div>}
                    {totalH>0&&<div style={{fontSize:14,color:isSel?"rgba(255,255,255,0.95)":GRN,fontWeight:900}}>{totalH}h</div>}
                    {isPending&&!isSel&&<div style={{position:"absolute",top:3,right:3,width:6,height:6,borderRadius:"50%",background:RED}}/>}
                  </div>
                );
              })}
            </div>

            {!isPublished&&selDates.length>0&&(
              <div style={{background:BBG,borderRadius:10,padding:"10px",marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
                  <div style={{fontSize:14,color:BLU,fontWeight:700}}>{selDates.length} day(s): {[...selDates].sort((a,b)=>a-b).join(", ")}</div>
                  <button
                    onClick={()=>setSelDates([])} style={{fontSize:14,fontWeight:700,padding:"4px 10px",borderRadius:7,border:"1px solid "+LN,background:PNL,color:SFT,cursor:"pointer"}}>Clear</button>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
                  <button onClick={()=>applyDirect("P")} style={{...PBtn,margin:0,background:GRN,fontSize:14,padding:"10px"}}>Present (8h+2OT)</button>
                  <button onClick={()=>applyDirect("A")} style={{...PBtn,margin:0,background:RED,fontSize:14,padding:"10px"}}>Absent (0h)</button>
                </div>
                <div style={{background:BG,borderRadius:8,padding:"10px"}}>
                  <div style={{fontSize:14,fontWeight:700,color:CB,marginBottom:8}}>Other Status</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
                    <div>
                      <label style={{...lbl,marginTop:0,fontSize:13}}>Status</label>
                      <select value={bulkStatus} onChange={e=>setBulkStatus(e.target.value)} style={inp}>
                        {STATUS_CODE_MAP.map(([code,label])=><option key={code} value={code}>{code} - {label}</option>)}
                      </select>
                    </div>
                    <div>
                      <label style={{...lbl,marginTop:0,fontSize:13,color:GRN}}>Basic Hrs</label>
                      <input type="number" min="0" max="8" value={bulkHrs} onChange={e=>setBulkHrs(Number(e.target.value))} style={inp}/>
                    </div>
                    <div>
                      <label style={{...lbl,marginTop:0,fontSize:13,color:ORANGE}}>OT Hrs</label>
                      <input type="number" min="0" max="12" value={bulkOT} onChange={e=>setBulkOT(Number(e.target.value))} style={inp}/>
                    </div>
                  </div>
                  {DIRECT.includes(bulkStatus)?(
                    <button onClick={applyBulk} style={{...PBtn,margin:0,background:GRN}}>Apply (no approval needed)</button>
                  ):(
                    <>
                      <button
                        onClick={()=>setShowOverrideForm(!showOverrideForm)} style={{...PBtn,margin:0,background:AMB}}>{showOverrideForm?"Cancel Override":"Request Override (Super Admin approval)"}</button>
                      {showOverrideForm&&(
                        <div style={{background:PNL,borderRadius:8,padding:"10px",marginTop:8,border:"1.5px solid "+AMB}}>
                          <label style={lbl}>Supervisor Remarks</label>
                          <input style={inp} value={overrideForm.reason} onChange={e=>setOverrideForm({...overrideForm,reason:e.target.value})} placeholder="Reason for override request..."/>
                          {(bulkStatus==="MWP"||bulkStatus==="MWOP")&&(
                            <>
                              <label style={lbl}>Attach Medical Paper (PDF/JPG)</label>
                              <input type="file" accept=".pdf,.jpg,.jpeg,.png" onChange={e=>setOverrideForm({...overrideForm,file:e.target.files[0]||null})} style={{...inp,padding:"5px"}}/>
                              {overrideForm.file&&<div style={{fontSize:13,color:GRN,marginTop:3}}>Attached: {overrideForm.file.name}</div>}
                              {!overrideForm.file&&<div style={{fontSize:13,color:AMB,marginTop:3}}>Please attach the medical certificate.</div>}
                            </>
                          )}
                          <div style={{display:"flex",gap:8,marginTop:8}}>
                            <button
                              onClick={requestOverride} style={{...PBtn,margin:0,flex:2,background:AMB,opacity:!overrideForm.reason?0.5:1}} disabled={!overrideForm.reason}>Submit to Super Admin</button>
                            <button
                              onClick={()=>setShowOverrideForm(false)} style={{flex:1,padding:"10px",border:"1px solid "+LN,borderRadius:10,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:15}}>Cancel</button>
                          </div>
                        </div>
                      )}
                    </>
                  )}
                </div>
              </div>
            )}

            <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:5,marginBottom:10}}>
              <div
                style={{background:GBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:GRN}}>{s2pC}</div><div style={{fontSize:11,color:GRN}}>P</div></div>
              <div
                style={{background:RBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:RED}}>{s2aC}</div><div style={{fontSize:11,color:RED}}>A</div></div>
              <div
                style={{background:YBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:YFG}}>{s2idleC}</div><div style={{fontSize:11,color:YFG}}>Idle</div></div>
              <div
                style={{background:BBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:BLU}}>{s2medC}</div><div style={{fontSize:11,color:BLU}}>Med</div></div>
              <div
                style={{background:PBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:PUR}}>{s2woC}</div><div style={{fontSize:11,color:PUR}}>WO</div></div>
              <div
                style={{background:OBG,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:16,fontWeight:900,color:ORANGE}}>{s2phC}</div><div style={{fontSize:11,color:ORANGE}}>PH</div></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:10}}>
              <div
                style={{background:GBG,borderRadius:8,padding:"8px",textAlign:"center"}}><div style={{fontSize:21,fontWeight:900,color:GRN}}>{totalBasic}h</div><div style={{fontSize:12,color:GRN,fontWeight:700}}>Basic Hours</div></div>
              <div
                style={{background:OBG,borderRadius:8,padding:"8px",textAlign:"center"}}><div style={{fontSize:21,fontWeight:900,color:ORANGE}}>{totalOT}h</div><div style={{fontSize:12,color:ORANGE,fontWeight:700}}>OT Hours</div></div>
              <div
                style={{background:BBG,borderRadius:8,padding:"8px",textAlign:"center"}}><div style={{fontSize:21,fontWeight:900,color:BLU}}>{totalBasic+totalOT}h</div><div style={{fontSize:12,color:BLU,fontWeight:700}}>Total Hours</div></div>
            </div>
            {saveConfirmed&&<div
              style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:8,fontSize:14,color:GRN,fontWeight:700}}>Changes saved for {selWorker.name} - {selMonth}</div>}
            {!isPublished&&<button onClick={saveEntries} style={{...PBtn,margin:"0 0 10px",background:"linear-gradient(135deg,"+DARK+","+CB+")"}}>Save Changes</button>}
            {pendingOverrides.length>0&&<div
              style={{background:RBG,border:"1px solid "+RED,borderRadius:8,padding:"8px 10px",fontSize:14,color:RED,fontWeight:700,marginBottom:10}}>{pendingOverrides.length} override(s) pending Super Admin approval before submission.</div>}
          </Card>
        </>
      )}

      
    </>
  );
}


function WorkerProfileTab({sup,gradeOverrides}){
  const myW=sup.workers;
  const [selWorker,setSelWorker]=useState(null);
  const [search,setSearch]=useState("");
  const [showDrop,setShowDrop]=useState(false);
  const [selMonth,setSelMonth]=useState("");
  const q=search.toLowerCase();
  const searchRes=(showDrop&&q.length>0)
    ?myW.filter(w=>
      w.name.toLowerCase().includes(q)||
      w.empCode.toLowerCase().includes(q))
    :[];
  const buildMonthList=(worker)=>{
    if(!worker||!worker.stampDate) return [];
    const MO=["Jan","Feb","Mar","Apr","May","Jun",
      "Jul","Aug","Sep","Oct","Nov","Dec"];
    const start=new Date(worker.stampDate);
    const now=new Date();
    const months=[];
    let d=new Date(start.getFullYear(),start.getMonth(),1);
    while(d<=now){
      months.push({
        lb:MO[d.getMonth()]+" "+d.getFullYear(),
        mo:d.getMonth(),yr:d.getFullYear()
      });
      d=new Date(d.getFullYear(),d.getMonth()+1,1);
    }
    return months.reverse();
  };
  const wMonths=selWorker?buildMonthList(selWorker):[];
  const wAtt=selWorker?(WORKER_ATT[selWorker.empCode]||{}):{}; 
  const curMonth=wMonths.find(m=>m.lb===selMonth)||wMonths[0]||null;
  return(
    <>
      <Card s={{marginBottom:10}}>
        <SHdr title="Worker Profile"
          sub="Search workers to view full profile"/>
        <div style={{position:"relative",marginBottom:8}}>
          <input style={inp}
            value={search}
            placeholder="Search by name or emp code..."
            onChange={e=>{
              const v=e.target.value;
              setSearch(v);setShowDrop(!!v);
              if(!v){setSelWorker(null);}
            }}
            onFocus={()=>search&&setShowDrop(true)}/>
          {showDrop&&searchRes.length>0&&(
            <div style={{position:"absolute",top:"100%",left:0,right:0,
              background:PNL,border:"1.5px solid "+LN,borderRadius:8,
              zIndex:99,maxHeight:180,overflowY:"auto",
              boxShadow:"0 4px 12px rgba(0,0,0,.1)"}}>
              {searchRes.map((w,i)=>(
                <div key={w.id}
                  onClick={()=>{
                    setSelWorker(w);
                    setSearch(w.name);
                    setShowDrop(false);
                    setSelMonth("");
                  }}
                  style={{display:"flex",
                    justifyContent:"space-between",
                    alignItems:"center",padding:"7px 10px",
                    borderTop:i?"1px solid "+LN:"none",
                    cursor:"pointer"}}>
                  <div>
                    <div style={{fontSize:16,fontWeight:700}}>
                      {w.name}
                    </div>
                    <div style={{fontSize:14,color:SFT}}>
                      {w.empCode} | {w.trade}
                    </div>
                  </div>
                  <span style={{fontSize:13,fontWeight:700,
                    padding:"2px 7px",borderRadius:6,
                    background:(STATUS_COLS[w.status]||{}).bg||"#eee",
                    color:(STATUS_COLS[w.status]||{}).fg||"#888"}}>
                    {w.status}
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>
      </Card>
      {selWorker&&(
        <>
          <div style={{background:"linear-gradient(135deg,"+DARK+","+CB+")",
            borderRadius:12,padding:16,color:"#fff",marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",
              alignItems:"flex-start",marginBottom:12}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <GradeBadge
                  grade={workerGrade(
                    selWorker.empCode,gradeOverrides,
                    selWorker.status)}
                  size={40}/>
                <div>
                  <div style={{fontSize:20,fontWeight:900}}>
                    {selWorker.name}
                  </div>
                  <div style={{fontSize:15,
                    color:"rgba(255,255,255,0.7)"}}>
                    {selWorker.trade} | {selWorker.nationality}
                  </div>
                </div>
              </div>
              <span style={{fontSize:14,fontWeight:800,
                padding:"4px 10px",borderRadius:8,
                background:"rgba(255,255,255,0.2)"}}>
                {selWorker.status}
              </span>
            </div>
            <div style={{display:"grid",
              gridTemplateColumns:"1fr 1fr",gap:6,marginBottom:12}}>
              {(()=>{
                const w=selWorker;
                const exDays=(dt)=>dt
                  ?Math.round((new Date(dt)-new Date())/86400000)
                  :null;
                const exCol=(d)=>d===null?"#fff"
                  :d<90?"#FF8A80":d<180?"#FFD180":"#6FCF97";
                const vDays=exDays(w.visaExpiry);
                const pDays=exDays(w.passportExpiry);
                const mDays=exDays(w.medicalInsExpiry);
                const wcDays=exDays(w.wcExpiry);
                const vExp=(w.visaExpiry||"-")
                  +(vDays!==null?" ("+vDays+"d)":"");
                const pExp=(w.passportExpiry||"-")
                  +(pDays!==null?" ("+pDays+"d)":"");
                const mExp=(w.medicalInsExpiry||"-")
                  +(mDays!==null?" ("+mDays+"d)":"");
                const wcExp=(w.wcExpiry||"-")
                  +(wcDays!==null?" ("+wcDays+"d)":"");
                const supObj=SUPS.find(
                  sv=>sv.clients&&sv.clients.includes(w.client));
                const rows=[
                  ["Emp Code",w.empCode,"#fff"],
                  ["Nationality",w.nationality,"#fff"],
                  ["Trade",w.trade,"#fff"],
                  ["Passport No.",w.passport||"-","#fff"],
                  ["Client",w.client,"#fff"],
                  ["Site",w.site,"#fff"],
                  ["Camp Name",w.campName||"-","#fff"],
                  ["Camp Location",w.campLocation||"-","#fff"],
                  ["Camp Boss",w.campBoss||"-","#fff"],
                  ["Camp Boss Phone",w.campBossPhone||"-","#fff"],
                  ["Supervisor",
                    supObj?supObj.name:"-","#F5C400"],
                  ["Months Working",
                    monthsWorked(w.stampDate)+" months","#fff"],
                  ["Stamp Date",w.stampDate||"-","#fff"],
                  ["Visa Type",w.visaType||"Fixit","#fff"],
                  ["Visa Expiry",vExp,exCol(vDays)],
                  ["Passport Expiry",pExp,exCol(pDays)],
                  ["Medical Ins Expiry",mExp,exCol(mDays)],
                  ["WC Expiry",wcExp,exCol(wcDays)],
                  ["Advance Pending",
                    "AED "+(w.advancePending||0),
                    w.advancePending>0?"#FFD180":"#fff"],
                  ["Total Advance",
                    "AED "+((WORKER_EXTRA_INFO[w.empCode]
                      ||{}).totalAdvanceTaken||0)
                      +" (lifetime)","#fff"],
                  ["Site Changes",
                    (w.siteChanges||0)+" times","#fff"],
                ];
                return rows.map(([l,v,col])=>(
                  <div key={l}
                    style={{background:"rgba(255,255,255,0.1)",
                      borderRadius:8,padding:"6px 10px"}}>
                    <div style={{fontSize:14,
                      color:"rgba(255,255,255,0.6)"}}>
                      {l}
                    </div>
                    <div style={{fontSize:16,fontWeight:700,
                      color:col}}>
                      {v}
                    </div>
                  </div>
                ));
              })()}
            </div>
            <div style={{background:"rgba(255,255,255,0.08)",
              borderRadius:10,padding:"10px 12px"}}>
              <div style={{fontSize:14,fontWeight:800,
                color:"rgba(255,255,255,0.7)",
                textTransform:"uppercase",letterSpacing:.5,
                marginBottom:10}}>
                Salary Details
              </div>
              {selWorker.payType==="fixed"?(
                <>
                  <div style={{display:"grid",
                    gridTemplateColumns:"repeat(4,1fr)",
                    gap:5,marginBottom:10}}>
                    {[
                      ["Basic",selWorker.basic||0],
                      ["Food",selWorker.food||0],
                      ["Other",selWorker.otherAllow||0],
                      ["Total",
                        (selWorker.basic||0)+
                        (selWorker.food||0)+
                        (selWorker.otherAllow||0)],
                    ].map(([l,v])=>(
                      <div key={l}
                        style={{background:"rgba(255,255,255,0.1)",
                          borderRadius:8,padding:"6px 8px",
                          textAlign:"center"}}>
                        <div style={{fontSize:12,
                          color:"rgba(255,255,255,0.5)",
                          marginBottom:2}}>{l}</div>
                        <div style={{fontSize:17,fontWeight:900,
                          color:l==="Total"?"#6FCF97":"#fff"}}>
                          AED {v}
                        </div>
                      </div>
                    ))}
                  </div>
                  <div style={{background:"rgba(0,0,0,0.2)",
                    borderRadius:8,padding:"8px 10px",
                    marginBottom:8}}>
                    <div style={{fontSize:13,fontWeight:700,
                      color:"rgba(255,255,255,0.6)",
                      marginBottom:6}}>
                      On Joining vs Current
                    </div>
                    <div style={{display:"grid",
                      gridTemplateColumns:"repeat(4,1fr)",gap:4}}>
                      {[
                        ["Basic",
                          selWorker.basicJoin||selWorker.basic||0,
                          selWorker.basic||0],
                        ["Food",
                          selWorker.foodJoin||selWorker.food||0,
                          selWorker.food||0],
                        ["Other",
                          selWorker.otherAllowJoin||
                            selWorker.otherAllow||0,
                          selWorker.otherAllow||0],
                        ["Total",
                          (selWorker.basicJoin||
                            selWorker.basic||0)+
                          (selWorker.foodJoin||
                            selWorker.food||0)+
                          (selWorker.otherAllowJoin||
                            selWorker.otherAllow||0),
                          (selWorker.basic||0)+
                          (selWorker.food||0)+
                          (selWorker.otherAllow||0)],
                      ].map(([l,join,cur])=>{
                        const inc=cur-join;
                        return(
                          <div key={l}
                            style={{background:
                              "rgba(255,255,255,0.06)",
                              borderRadius:7,padding:"6px 8px"}}>
                            <div style={{fontSize:12,
                              color:"rgba(255,255,255,0.5)",
                              marginBottom:2}}>{l}</div>
                            <div style={{fontSize:12,
                              color:"rgba(255,255,255,0.5)"}}>
                              Join: {join}
                            </div>
                            <div style={{fontSize:15,
                              fontWeight:700,color:"#fff"}}>
                              Now: {cur}
                            </div>
                            <div style={{fontSize:13,
                              fontWeight:800,
                              color:inc>0?"#6FCF97":
                                inc<0?"#FF8A80":"#888"}}>
                              {inc>0?"+":""}{inc}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                  <div style={{display:"grid",
                    gridTemplateColumns:"1fr 1fr",gap:6}}>
                    {[
                      ["Normal OT",
                        "AED "+(selWorker.otRate||"-")+"/hr"],
                      ["PH OT",
                        "AED "+(selWorker.phOtRate||"-")+"/hr"],
                    ].map(([l,v])=>(
                      <div key={l}
                        style={{background:"rgba(255,255,255,0.1)",
                          borderRadius:8,padding:"6px 8px"}}>
                        <div style={{fontSize:12,
                          color:"rgba(255,255,255,0.5)"}}>{l}
                        </div>
                        <div style={{fontSize:15,fontWeight:700,
                          color:"#fff"}}>{v}</div>
                      </div>
                    ))}
                  </div>
                </>
              ):(
                <div style={{display:"grid",
                  gridTemplateColumns:"1fr 1fr",gap:6}}>
                  <div style={{background:"rgba(255,255,255,0.1)",
                    borderRadius:8,padding:"8px 10px"}}>
                    <div style={{fontSize:12,
                      color:"rgba(255,255,255,0.5)"}}>
                      Hourly Rate (Join)
                    </div>
                    <div style={{fontSize:17,fontWeight:900,
                      color:"#fff"}}>
                      AED {selWorker.rateJoin||"-"}/hr
                    </div>
                  </div>
                  <div style={{background:"rgba(255,255,255,0.1)",
                    borderRadius:8,padding:"8px 10px"}}>
                    <div style={{fontSize:12,
                      color:"rgba(255,255,255,0.5)"}}>
                      Hourly Rate (Current)
                    </div>
                    <div style={{fontSize:17,fontWeight:900,
                      color:"#6FCF97"}}>
                      AED {selWorker.rate||"-"}/hr
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
          <Card s={{marginBottom:10}}>
            <div style={{display:"flex",
              justifyContent:"space-between",
              alignItems:"center",marginBottom:10}}>
              <div style={{fontSize:17,fontWeight:800,color:CB}}>
                Monthly Client Attendance
              </div>
              <select
                value={selMonth}
                onChange={e=>setSelMonth(e.target.value)}
                style={{...inp,width:"auto",
                  minWidth:130,marginBottom:0}}>
                {wMonths.map(m=>(
                  <option key={m.lb} value={m.lb}>{m.lb}</option>
                ))}
              </select>
            </div>
            {curMonth&&(()=>{
              const dim=new Date(
                curMonth.yr,curMonth.mo+1,0).getDate();
              const firstDay=new Date(
                curMonth.yr,curMonth.mo,1).getDay();
              const dayNums=Object.keys(wAtt).map(Number);
              const p=dayNums.filter(d=>wAtt[d]==="P").length;
              const a=dayNums.filter(d=>wAtt[d]==="A").length;
              const idl=dayNums.filter(d=>wAtt[d]==="I").length;
              const pct=dim>0?Math.round(p/dim*100):0;
              return(
                <>
                  <div style={{display:"flex",
                    justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
                    <span style={{fontSize:15,fontWeight:700,
                      color:CB}}>
                      {curMonth.lb}
                    </span>
                    <span style={{fontSize:17,fontWeight:900,
                      color:pct>90?GRN:pct>75?AMB:RED}}>
                      {pct}%
                    </span>
                  </div>
                  <div style={{display:"grid",
                    gridTemplateColumns:"repeat(3,1fr)",
                    gap:6,marginBottom:12}}>
                    <div style={{background:GBG,borderRadius:8,
                      padding:"8px",textAlign:"center"}}>
                      <div style={{fontSize:23,fontWeight:900,
                        color:GRN}}>{p}</div>
                      <div style={{fontSize:13,color:GRN}}>
                        Present
                      </div>
                    </div>
                    <div style={{background:RBG,borderRadius:8,
                      padding:"8px",textAlign:"center"}}>
                      <div style={{fontSize:23,fontWeight:900,
                        color:RED}}>{a}</div>
                      <div style={{fontSize:13,color:RED}}>
                        Absent
                      </div>
                    </div>
                    <div style={{background:YBG,borderRadius:8,
                      padding:"8px",textAlign:"center"}}>
                      <div style={{fontSize:23,fontWeight:900,
                        color:YFG}}>{idl}</div>
                      <div style={{fontSize:13,color:YFG}}>
                        Idle
                      </div>
                    </div>
                  </div>
                  <div style={{display:"grid",
                    gridTemplateColumns:"repeat(7,1fr)",gap:3}}>
                    {["Su","Mo","Tu","We","Th","Fr","Sa"].map(d=>(
                      <div key={d}
                        style={{textAlign:"center",fontSize:11,
                          fontWeight:700,color:SFT,padding:"2px"}}>
                        {d}
                      </div>
                    ))}
                    {Array.from({length:firstDay},(_,bi)=>(
                      <div key={"b"+bi}/>
                    ))}
                    {Array.from({length:dim},(_,di)=>{
                      const day=di+1;
                      const st=wAtt[day]||"";
                      const sc=SS[st]||{bg:"#f0f0f0",fg:"#bbb"};
                      return(
                        <div key={day}
                          style={{background:sc.bg,
                            borderRadius:4,padding:"3px 2px",
                            textAlign:"center",minHeight:30,
                            display:"flex",flexDirection:"column",
                            alignItems:"center",
                            justifyContent:"center"}}>
                          {st?(
                            <>
                              <span style={{fontSize:13,
                                fontWeight:800,color:sc.fg,
                                lineHeight:1.1}}>
                                {st}
                              </span>
                              <span style={{fontSize:8,
                                color:sc.fg,opacity:0.7,
                                lineHeight:1.2}}>
                                {day}
                              </span>
                            </>
                          ):(
                            <span style={{fontSize:13,
                              fontWeight:700,color:"#bbb"}}>
                              {day}
                            </span>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </>
              );
            })()}
            {!curMonth&&(
              <div style={{fontSize:14,color:SFT,
                fontStyle:"italic"}}>
                No attendance data available.
              </div>
            )}
          </Card>
          <Card s={{marginBottom:10}}>
            <SHdr title="PPE Issued"
              sub="Personal Protective Equipment"/>
            <div style={{fontSize:14,color:SFT,marginBottom:8}}>
              Total: {(PPE_ISSUED_DATA[selWorker.empCode]||[])
                .reduce((acc,it)=>acc+it.qty,0)} items
            </div>
            {(PPE_ISSUED_DATA[selWorker.empCode]||[]).map((p,pi)=>{
              const rawDate=p.issued||p.date||"";
              const parsePPEDate=(ds)=>{
                if(!ds) return new Date("2026-01-01");
                const pts=ds.split("/");
                if(pts.length===3&&pts[0].length<=2)
                  return new Date(pts[2]+"-"+pts[1]+"-"+pts[0]);
                return new Date(ds);
              };
              const mAgo=Math.round(
                (new Date()-parsePPEDate(rawDate))
                /(1000*60*60*24*30));
              return(
                <div key={pi}
                  style={{display:"flex",
                    justifyContent:"space-between",
                    alignItems:"center",padding:"6px 0",
                    borderTop:pi?"1px solid "+LN:"none"}}>
                  <div>
                    <div style={{fontSize:15,fontWeight:700}}>
                      {p.item}
                    </div>
                    <div style={{fontSize:13,color:SFT}}>
                      {p.size?("Size: "+p.size+" | "):""}
                      Qty: {p.qty} | {p.issued||p.date}
                    </div>
                  </div>
                  <span style={{fontSize:13,fontWeight:700,
                    padding:"2px 7px",borderRadius:6,
                    background:mAgo>12?RBG:mAgo>6?ABG:GBG,
                    color:mAgo>12?RED:mAgo>6?AMB:GRN}}>
                    {mAgo}mo ago
                  </span>
                </div>
              );
            })}
            {!(PPE_ISSUED_DATA[selWorker.empCode]||[]).length&&(
              <div style={{color:SFT,fontSize:14}}>
                No PPE issued.
              </div>
            )}
          </Card>
          <Card>
            <SHdr title="Pending Approvals"
              sub="Active requests for this worker"/>
            {((WORKER_EXTRA_INFO[selWorker.empCode]
              ||{}).pendingApprovals||
              WORKER_PENDING_ITEMS[selWorker.empCode]||[])
              .length>0
              ?((WORKER_EXTRA_INFO[selWorker.empCode]
                ||{}).pendingApprovals||
                WORKER_PENDING_ITEMS[selWorker.empCode]||[])
                .map((p,pi)=>{
                  const isSA=p.stage.includes("Super Admin");
                  const sc=isSA
                    ?{bg:BBG,fg:BLU}:{bg:ABG,fg:AMB};
                  return(
                    <div key={pi}
                      style={{padding:"6px 0",
                        borderTop:pi?"1px solid "+LN:"none"}}>
                      <div style={{display:"flex",
                        justifyContent:"space-between",
                        alignItems:"center"}}>
                        <div style={{fontSize:15,fontWeight:700}}>
                          {p.type}
                        </div>
                        <Pill c={p.stage} bg={sc.bg} fg={sc.fg}/>
                      </div>
                      <div style={{fontSize:13,color:SFT}}>
                        {p.date}
                        {p.note&&" - "+p.note}
                      </div>
                    </div>
                  );
              })
              :<div style={{color:SFT,fontSize:14}}>
                No pending approvals.
              </div>
            }
            <div style={{marginTop:10,paddingTop:10,
              borderTop:"1px solid "+LN}}>
              <div style={{fontSize:14,fontWeight:700,
                color:CB,marginBottom:6}}>
                Emergency Contact
              </div>
              <div style={{fontSize:15,fontWeight:600}}>
                {(WORKER_EXTRA_INFO[selWorker.empCode]
                  ||{}).emgName||"-"}
                {" - "}
                {(WORKER_EXTRA_INFO[selWorker.empCode]
                  ||{}).emgRelation||"-"}
              </div>
              <div style={{fontSize:14,color:SFT,marginTop:2}}>
                {(WORKER_EXTRA_INFO[selWorker.empCode]
                  ||{}).emgPhone||"-"}
              </div>
            </div>
          </Card>
        </>
      )}
      {!selWorker&&(
        <div style={{textAlign:"center",color:SFT,
          fontSize:18,padding:"30px 0",fontStyle:"italic"}}>
          Search for a worker to view their full profile.
        </div>
      )}
    </>
  );
}


function WorkerSalaryTab({sup}){
  const [search,setSearch]=useState("");
  const [selWorker,setSelWorker]=useState(null);
  const [showDrop,setShowDrop]=useState(false);
  const myW=sup.workers;
  const q=search.toLowerCase();
  const searchRes=showDrop&&q.length>0
    ?myW.filter(w=>w.name.toLowerCase().includes(q)||w.empCode.toLowerCase().includes(q)):[];
  const MONTHS=[
    {lb:"Jun 2026",present:18,absent:2,idle:1,ot:3.5,uploaded:false},
    {lb:"May 2026",present:21,absent:1,idle:1,ot:5.0,uploaded:true,basic:1500,food:200,otPay:100,bonus:100,advance:200,fine:0,net:1700},
    {lb:"Apr 2026",present:20,absent:2,idle:0,ot:0,uploaded:true,basic:1500,food:200,otPay:0,bonus:0,advance:0,fine:0,net:1700},
  ];
  return(
    <>
      <Card s={{marginBottom:10}}>
        <SHdr title="Worker Salary Calculation" sub="View salary breakdown for any worker"/>
        <input style={inp} value={search}
          onChange={e=>{
            const v=e.target.value;
            setSearch(v);setShowDrop(!!v);
            if(!v){setSelWorker(null);}
          }} onFocus={()=>setShowDrop(true)} placeholder="Search by name or emp code..."/>
        {searchRes.map((w,i)=>(
          <div key={w.id}
            onClick={()=>{
              setSelWorker(w);setSearch(w.name);setShowDrop(false);
            }} style={{display:"flex",justifyContent:"space-between",
              alignItems:"center",padding:"7px 0",
              borderTop:i?"1px solid "+LN:"none",
              cursor:"pointer",marginTop:4}}>
            <div><div style={{fontSize:16,fontWeight:700}}>{w.name}</div><div style={{fontSize:14,color:SFT}}>{w.empCode}</div></div>
          </div>
        ))}
      </Card>
      {selWorker&&MONTHS.map((mo,i)=>(
        <Card key={i} s={{marginBottom:10,border:mo.uploaded?"1px solid "+GRN:"1px solid "+LN}}>
          <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
            <div style={{fontSize:16,fontWeight:800,color:CB}}>{mo.lb}</div>
            <Pill c={mo.uploaded?"Salary Published":"Pending Upload"} bg={mo.uploaded?GBG:ABG} fg={mo.uploaded?GRN:AMB}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:5,marginBottom:10}}>
            {[["Present",mo.present,GRN,GBG],["Absent",mo.absent,RED,RBG],["Idle",mo.idle,YFG,YBG],["OT Hrs",mo.ot+"h",ORANGE,OBG]].map(([l,v,fg,bg])=>(
              <div key={l}
                style={{background:bg,borderRadius:7,padding:"5px",textAlign:"center"}}><div style={{fontSize:17,fontWeight:900,color:fg}}>{v}</div><div style={{fontSize:12,color:fg}}>{l}</div></div>
            ))}
          </div>
          {mo.uploaded&&(
            <>
              {[["Basic","AED "+mo.basic,GRN],["Food Allowance","AED "+mo.food,GRN],["OT Pay","AED "+mo.otPay,ORANGE],["Bonus","AED "+mo.bonus,GRN],["Advance Deduction","- AED "+mo.advance,RED],["Fine Deduction","- AED "+mo.fine,RED]].map(([l,v,c])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid "+LN}}>
                  <span style={{fontSize:14,color:SFT}}>{l}</span><span style={{fontSize:14,fontWeight:600,color:c}}>{v}</span>
                </div>
              ))}
              <div style={{borderTop:"2px solid "+CB,marginTop:6,paddingTop:6,display:"flex",justifyContent:"space-between"}}>
                <span style={{fontSize:16,fontWeight:800,color:CB}}>Net Payable</span>
                <span style={{fontSize:16,fontWeight:900,color:CB}}>AED {mo.net}</span>
              </div>
            </>
          )}
        </Card>
      ))}
      {!selWorker&&<div style={{textAlign:"center",color:SFT,fontSize:16,padding:"30px 0",fontStyle:"italic"}}>Search a worker to view salary breakdown.</div>}
    </>
  );
}

const WORKER_EXTRA_INFO={
  "EMP-1042":{agency:"Bin Sulaiman Manpower Agency",agencyContact:"Mr. Rashid Al Bin Sulaiman",agencyPhone:"+971 4 234 5678",agencyEmail:"rashid@binsulaimanmp.ae",localPhone:"+971 55 123 4567",homePhone:"+91 98765 43210",email:"ramesh.kumar@fixitmail.com",emgName:"Sita Kumar",emgPhone:"+91 98765 11111",emgRelation:"Wife",totalAdvanceTaken:1450,previousClients:["DEWA Infrastructure","Nakheel Properties"]},
  "EMP-1043":{agency:"Al Noor Recruitment Services",agencyContact:"Ms. Fatima Al Noor",agencyPhone:"+971 4 345 6789",agencyEmail:"fatima@alnoorrecruit.ae",localPhone:"+971 55 234 5678",homePhone:"+91 98765 22222",email:"suresh.pillai@fixitmail.com",emgName:"Lakshmi Pillai",emgPhone:"+91 98765 33333",emgRelation:"Mother",totalAdvanceTaken:600,previousClients:["Emaar Properties"]},
  "EMP-1044":{agency:"Gulf Star Manpower Supply",agencyContact:"Mr. Khalid Gulf Star",agencyPhone:"+971 4 456 7890",agencyEmail:"khalid@gulfstarmp.ae",localPhone:"+971 55 345 6789",homePhone:"+971 50 111 2222",email:"faisal.ahmed@fixitmail.com",emgName:"Ahmed Hassan",emgPhone:"+971 50 444 5555",emgRelation:"Brother",totalAdvanceTaken:0,previousClients:[]},
  "EMP-1045":{agency:"Bin Sulaiman Manpower Agency",agencyContact:"Mr. Rashid Al Bin Sulaiman",agencyPhone:"+971 4 234 5678",agencyEmail:"rashid@binsulaimanmp.ae",localPhone:"+971 55 456 7890",homePhone:"+977 98 1234 5678",email:"bikash.thapa@fixitmail.com",emgName:"Pemba Thapa",emgPhone:"+977 98 8765 4321",emgRelation:"Father",totalAdvanceTaken:850,previousClients:["Al Futtaim Construction"]},
};
const PPE_ISSUED_DATA={
  "EMP-1042":[
    {item:"Helmet",size:"",qty:1,condition:"Good",issued:"15/01/2026"},
    {item:"Safety Jacket",size:"",qty:1,condition:"Worn",issued:"15/01/2026"},
    {item:"Gloves",size:"",qty:2,condition:"Good",issued:"01/03/2026"},
    {item:"Safety Glasses",size:"",qty:1,condition:"Good",issued:"15/01/2026"},
    {item:"Shoes",size:"42",qty:1,condition:"Good",issued:"15/01/2026"},
    {item:"Fixit T-Shirt",size:"Large",qty:2,condition:"Good",issued:"15/01/2026"},
    {item:"Tools",size:"",qty:1,condition:"Good",issued:"15/01/2026",details:"Screwdriver set, measuring tape"},
  ],
  "EMP-1043":[
    {item:"Helmet",size:"",qty:1,condition:"Good",issued:"01/03/2025"},
    {item:"Safety Jacket",size:"",qty:1,condition:"Good",issued:"01/03/2025"},
    {item:"Gloves",size:"",qty:2,condition:"Needs Replacement",issued:"01/03/2025"},
  ],
  "EMP-1044":[
    {item:"Helmet",size:"",qty:1,condition:"New",issued:"20/04/2025"},
    {item:"Fixit T-Shirt",size:"Medium",qty:1,condition:"Good",issued:"20/04/2025"},
  ],
  "EMP-1045":[
    {item:"Helmet",size:"",qty:1,condition:"Good",issued:"10/02/2026"},
    {item:"Safety Jacket",size:"",qty:1,condition:"Worn",issued:"10/02/2026"},
  ],
};
const WORKER_PENDING_ITEMS={
  "EMP-1042":[
    {type:"Medical (MWP)",date:"10 Jun 2026",stage:"Pending Super Admin",note:"Medical certificate submitted"},
    {type:"PPE Request - Coverall",date:"01 Jul 2026",stage:"Pending Supervisor",note:"Current coverall torn"},
  ],
  "EMP-1043":[
    {type:"Idle",date:"14 May 2026",stage:"Pending Super Admin",note:"No work assigned by client"},
    {type:"PPE Request - Shoes",date:"02 Jul 2026",stage:"Pending Supervisor",note:"Sole detached"},
  ],
  "EMP-1044":[
    {type:"PPE Request - Fixit T-Shirt",date:"25 Jun 2026",stage:"Approved - Ready to Issue",note:"New joining kit"},
    {type:"Advance - AED 180",date:"28 Jun 2026",stage:"Pending Admin",note:"Family emergency"},
  ],
  "EMP-1045":[
    {type:"Emergency Leave",date:"18 Jun 2026",stage:"Pending Super Admin",note:"Family emergency"},
    {type:"Advance - AED 350",date:"27 Jun 2026",stage:"Pending Super Admin",note:"Medical expense"},
    {type:"PPE Request - Helmet",date:"28 Jun 2026",stage:"Pending Super Admin",note:"Lost on site"},
  ],
};
const PPE_REQUESTS=[
  {id:"PPE-001",empCode:"EMP-1042",worker:"Ramesh Kumar",site:"Camp 3 - Site A",client:"Al Futtaim Construction",item:"Coverall",size:"Large",qty:1,reason:"Current coverall is torn",requestedOn:"01 Jul 2026",status:"Pending Supervisor"},
  {id:"PPE-002",empCode:"EMP-1043",worker:"Suresh Pillai",site:"Camp 3 - Site A",client:"Al Futtaim Construction",item:"Shoes",size:"42",qty:1,reason:"Sole detached",requestedOn:"02 Jul 2026",status:"Pending Supervisor"},
  {id:"PPE-003",empCode:"EMP-1045",worker:"Bikash Thapa",site:"DEWA Block 7",client:"DEWA Infrastructure",item:"Helmet",size:"",qty:1,reason:"Lost on site",requestedOn:"28 Jun 2026",status:"Pending Super Admin"},
  {id:"PPE-004",empCode:"EMP-1044",worker:"Vikram Nair",site:"Emaar Site B",client:"Emaar Properties",item:"Fixit T-Shirt",size:"Medium",qty:2,reason:"New joining kit",requestedOn:"25 Jun 2026",status:"Approved"},
];

function WorkerPPETab({sup}){
  const [search,setSearch]=useState("");
  const [selWorker,setSelWorker]=useState(null);
  const [showDrop,setShowDrop]=useState(false);
  const [ppeRequests,setPpeRequests]=useState(PPE_REQUESTS);
  const [showSupReq,setShowSupReq]=useState(false);
  const [supReqForm,setSupReqForm]=useState({item:"Helmet",size:"",qty:1,reason:"",details:""});
  const [supReqSent,setSupReqSent]=useState(false);
  const [supReqList,setSupReqList]=useState([]);
  const [issuingItem,setIssuingItem]=useState(null);
  const [issueDate,setIssueDate]=useState(new Date().toISOString().slice(0,10));
  const [issueCondition,setIssueCondition]=useState("New");
  const [issueNote,setIssueNote]=useState("");
  const [issueSent,setIssueSent]=useState(false);
  const myW=sup.workers;
  const q=search.toLowerCase();
  const searchRes=showDrop&&q.length>0
    ?myW.filter(w=>w.name.toLowerCase().includes(q)||w.empCode.toLowerCase().includes(q)):[];

  const PPE_ITEMS=["Helmet","Safety Glasses","Safety Jacket","Coverall","Shoes","Gloves","Fixit T-Shirt","Tools","Handyman Pant","Others"];
  const CLOTH_SIZES=["Small","Medium","Large","Extra Large"];
  const SHOE_SIZES=["38","39","40","41","42","43","44","45","46","47","48"];
  const NEEDS_CLOTH=["Coverall","Fixit T-Shirt","Handyman Pant"];
  const NEEDS_SHOE=["Shoes"];
  const NEEDS_DETAILS=["Tools","Others"];
  const getSizes=(item)=>{if(NEEDS_SHOE.includes(item))return SHOE_SIZES;if(NEEDS_CLOTH.includes(item))return CLOTH_SIZES;return [];};

  const monthsElapsed=(dateStr)=>{
    if(!dateStr)return 0;
    const parts=dateStr.split("/");
    if(parts.length!==3)return 0;
    const issued=new Date(parts[2],parts[1]-1,parts[0]);
    return Math.max(0,Math.round((new Date()-issued)/2628000000));
  };
  const workerRequests=selWorker?ppeRequests.filter(r=>r.empCode===selWorker.empCode):[];
  const allPending=ppeRequests.filter(r=>r.status==="Pending Supervisor"&&sup.clients.includes(r.client));

  const forwardToSuperAdmin=(id)=>setPpeRequests(ppeRequests.map(r=>r.id===id?{...r,status:"Pending Super Admin"}:r));
  const openIssueForm=(r)=>{setIssuingItem(r);setIssueDate(new Date().toISOString().slice(0,10));setIssueCondition("New");setIssueNote("");};
  const confirmIssue=()=>{
    if(!issuingItem)return;
    setPpeRequests(ppeRequests.map(r=>r.id===issuingItem.id?{...r,status:"Issued",issuedOn:issueDate,condition:issueCondition,issueNote:issueNote}:r));
    setIssuingItem(null);setIssueSent(true);setTimeout(()=>setIssueSent(false),3000);
  };

  const statusColor=(st)=>st==="Pending Supervisor"?{bg:ABG,fg:AMB}:st==="Pending Super Admin"?{bg:BBG,fg:BLU}:st==="Approved"?{bg:GBG,fg:GRN}:st==="Issued"?{bg:GBG,fg:GRN}:{bg:RBG,fg:RED};

  const submitSupReq=()=>{
    if(!selWorker||!supReqForm.reason)return;
    setSupReqList([...supReqList,{...supReqForm,worker:selWorker.name,empCode:selWorker.empCode,status:"Pending Super Admin",date:"Today",id:"SUP-"+Date.now()}]);
    setSupReqSent(true);setShowSupReq(false);
    setTimeout(()=>setSupReqSent(false),3000);
    setSupReqForm({item:"Helmet",size:"",qty:1,reason:"",details:""});
  };

  const [rptClient,setRptClient]=useState("");
  const [rptSite,setRptSite]=useState("");
  const [showRptPreview,setShowRptPreview]=useState(false);
  const rptRef=useRef(null);
  const rptClients=[...new Set(myW.map(w=>w.client))];
  const rptSites=[...new Set(myW.filter(w=>!rptClient||w.client===rptClient).map(w=>w.site))];
  const rptWorkers=myW.filter(w=>(!rptClient||w.client===rptClient)&&(!rptSite||w.site===rptSite));
  const rptRows=rptWorkers.flatMap(w=>(PPE_ISSUED_DATA[w.empCode]||[]).map(it=>({worker:w.name,empCode:w.empCode,client:w.client,site:w.site,item:it.item,size:it.size||"-",qty:it.qty||1,condition:it.condition,issued:it.issued,details:it.details||""})));
  const downloadPPEExcel=()=>{
    const wb=XLSX.utils.book_new();
    const header=["Worker","Emp Code","Client","Site","PPE Item","Size","Qty","Condition","Date Issued","Details"];
    const data=rptRows.map(r=>[r.worker,r.empCode,r.client,r.site,r.item,r.size,r.qty,r.condition,r.issued,r.details]);
    const itemCounts=Object.entries(rptRows.reduce((acc,r)=>{acc[r.item]=(acc[r.item]||0)+(Number(r.qty)||1);return acc;},{})).sort((a,b)=>b[1]-a[1]);
    const titleRows=[
      ["FIXIT MANPOWER SUPPLY L.L.C",...Array(9).fill("")],
      ["PPE Issued Report",...Array(9).fill("")],
      [(rptClient||"All Clients")+" | "+(rptSite||"All Sites")+" | Generated: "+new Date().toDateString(),...Array(9).fill("")],
      Array(10).fill(""),
      ["COUNT ISSUED BY ITEM",...Array(9).fill("")],
      ["Item","Total Qty Issued",...Array(8).fill("")],
      ...itemCounts.map(([item,count])=>[item,count,...Array(8).fill("")]),
      Array(10).fill(""),
      ["FULL DETAIL",...Array(9).fill("")],
    ];
    const ws=XLSX.utils.aoa_to_sheet([...titleRows,header,...data]);
    ws["!cols"]=[{wch:16},{wch:10},{wch:20},{wch:16},{wch:14},{wch:8},{wch:5},{wch:12},{wch:12},{wch:24}];
    XLSX.utils.book_append_sheet(wb,ws,"PPE Report");
    XLSX.writeFile(wb,"PPE_Report_"+(rptClient||"AllClients").split(" ").join("_")+".xlsx");
  };
  const printPPEReport=()=>{
    if(!rptRef.current)return;
    const printWindow=window.open("","_blank");
    printWindow.document.write(
      "<html><head><title>PPE Issued Report</title><style>"+
      "body{font-family:Arial,sans-serif;margin:20px;color:#1a1a1a}"+
      "table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:6px 8px;font-size:12px}"+
      "@media print{body{margin:0}}"+
      "</style></head><body>"+rptRef.current.innerHTML+"</body></html>"
    );
    printWindow.document.close();
    printWindow.focus();
    setTimeout(()=>{printWindow.print();},400);
  };

  return(
    <>
      {issuingItem&&(
        <div style={{background:GBG,border:"2px solid "+GRN,borderRadius:12,padding:"14px 16px",marginBottom:12}}>
          <div style={{fontSize:16,fontWeight:800,color:GRN,marginBottom:4}}>Confirm PPE Issue</div>
          <div style={{fontSize:14,color:SFT,marginBottom:10}}>{issuingItem.worker} | {issuingItem.item}{issuingItem.size?" - "+issuingItem.size:""} x{issuingItem.qty}</div>
          <label style={lbl}>Issue Date</label>
          <input type="date" style={inp} value={issueDate} onChange={e=>setIssueDate(e.target.value)}/>
          <label style={lbl}>Condition at Issue</label>
          <select style={inp} value={issueCondition} onChange={e=>setIssueCondition(e.target.value)}>
            {["New","Good"].map(c=><option key={c}>{c}</option>)}
          </select>
          <label style={lbl}>Notes (optional)</label>
          <input style={inp} value={issueNote} onChange={e=>setIssueNote(e.target.value)} placeholder="e.g. Issued at site office, worker signed receipt"/>
          <div style={{display:"flex",gap:8,marginTop:10}}>
            <button style={{...PBtn,margin:0,flex:2,background:GRN}} onClick={confirmIssue}>Confirm Issue and Record</button>
            <button
              onClick={()=>setIssuingItem(null)} style={{flex:1,padding:"10px",border:"1px solid "+LN,borderRadius:10,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:15}}>Cancel</button>
          </div>
        </div>
      )}

      <Card s={{marginBottom:10}}>
        <div style={{fontSize:15,fontWeight:800,color:CB,marginBottom:8}}>Worker PPE Records</div>
        <input style={inp} value={search}
          onChange={e=>{
            const v=e.target.value;
            setSearch(v);setShowDrop(!!v);
            if(!v){setSelWorker(null);}
          }} onFocus={()=>setShowDrop(true)} placeholder="Search worker to view PPE record..."/>
        {searchRes.map((w,i)=>(
          <div key={w.id}
            onClick={()=>{
              setSelWorker(w);setSearch(w.name);setShowDrop(false);
            }} style={{display:"flex",justifyContent:"space-between",
              alignItems:"center",padding:"7px 0",
              borderTop:i?"1px solid "+LN:"none",
              cursor:"pointer",marginTop:4}}>
            <div><div style={{fontSize:16,fontWeight:700}}>{w.name}</div><div style={{fontSize:14,color:SFT}}>{w.empCode} | {w.site}</div></div>
            <Pill c={ppeRequests.filter(r=>r.empCode===w.empCode&&r.status==="Pending Supervisor").length>0?"Has pending request":"View"} bg={ppeRequests.filter(r=>r.empCode===w.empCode&&r.status==="Pending Supervisor").length>0?ABG:BBG} fg={ppeRequests.filter(r=>r.empCode===w.empCode&&r.status==="Pending Supervisor").length>0?AMB:BLU}/>
          </div>
        ))}
      </Card>

      {selWorker&&(
        <>
          <div style={{background:"linear-gradient(135deg,"+DARK+","+CB+")",borderRadius:12,padding:"12px 16px",marginBottom:10,color:"#fff"}}>
            <div style={{fontSize:18,fontWeight:800}}>{selWorker.name}</div>
            <div style={{fontSize:14,color:"rgba(255,255,255,0.7)",marginTop:2,marginBottom:10}}>{selWorker.empCode} | {selWorker.trade} | {selWorker.site}</div>
            {(PPE_ISSUED_DATA[selWorker.empCode]||[]).length>0&&(
              <>
                <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:6}}>
                  <div style={{fontSize:13,fontWeight:700,color:"rgba(255,255,255,0.5)",textTransform:"uppercase",letterSpacing:.4}}>PPE Currently Issued</div>
                  <div style={{fontSize:13,fontWeight:700,color:"rgba(255,255,255,0.6)"}}>{(PPE_ISSUED_DATA[selWorker.empCode]||[]).reduce((s,it)=>s+(Number(it.qty)||1),0)} items total</div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:5}}>
                  {(PPE_ISSUED_DATA[selWorker.empCode]||[]).map((it,i)=>(
                    <div key={i} style={{background:"rgba(255,255,255,0.08)",borderRadius:8,padding:"6px 8px"}}>
                      <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center"}}>
                        <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{it.item}{it.size?" - "+it.size:""}</div>
                        <div style={{fontSize:13,fontWeight:800,color:CY,background:"rgba(255,255,255,0.15)",borderRadius:6,padding:"1px 6px"}}>x{it.qty||1}</div>
                      </div>
                      <div style={{fontSize:12,color:"rgba(255,255,255,0.6)"}}>Issued: {it.issued}</div>
                      <div style={{fontSize:12,fontWeight:700,color:monthsElapsed(it.issued)>12?"#FF6B6B":monthsElapsed(it.issued)>6?"#F5C400":"#6FCF97"}}>{monthsElapsed(it.issued)} months ago</div>
                      {it.details&&<div style={{fontSize:12,color:"rgba(255,255,255,0.5)"}}>{it.details}</div>}
                    </div>
                  ))}
                </div>
              </>
            )}
            {(PPE_ISSUED_DATA[selWorker.empCode]||[]).length===0&&(
              <div style={{fontSize:14,color:"rgba(255,255,255,0.5)",fontStyle:"italic"}}>No PPE items on record for this worker.</div>
            )}
          </div>
          {supReqSent&&<div
            style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:GRN,fontWeight:700}}>PPE request forwarded to Super Admin for approval.</div>}
          {!showSupReq?(
            <button onClick={()=>setShowSupReq(true)} style={{...PBtn,background:ORANGE,marginBottom:10,marginTop:0}}>+ Raise PPE Request for {selWorker.name}</button>
          ):(
            <Card s={{marginBottom:10,border:"2px solid "+ORANGE}}>
              <div style={{fontSize:15,fontWeight:800,color:ORANGE,marginBottom:4}}>PPE Request - {selWorker.name}</div>
              <div style={{fontSize:13,color:SFT,marginBottom:8}}>Request will go directly to Super Admin for approval. Once approved you can issue the item.</div>
              <label style={lbl}>PPE Item</label>
              <select style={inp} value={supReqForm.item} onChange={e=>setSupReqForm({...supReqForm,item:e.target.value,size:"",details:""})}>
                {PPE_ITEMS.map(p=><option key={p}>{p}</option>)}
              </select>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                <div>
                  <label style={lbl}>Size</label>
                  {getSizes(supReqForm.item).length>0?(
                    <select style={inp} value={supReqForm.size} onChange={e=>setSupReqForm({...supReqForm,size:e.target.value})}>
                      <option value="">Select size</option>
                      {getSizes(supReqForm.item).map(s=><option key={s}>{s}</option>)}
                    </select>
                  ):(
                    <div style={{...inp,color:SFT,fontSize:14,display:"flex",alignItems:"center"}}>No size required</div>
                  )}
                </div>
                <div>
                  <label style={lbl}>Quantity</label>
                  <input type="number" min="1" max="10" style={inp} value={supReqForm.qty} onChange={e=>setSupReqForm({...supReqForm,qty:parseInt(e.target.value)||1})}/>
                </div>
              </div>
              {NEEDS_DETAILS.includes(supReqForm.item)&&(
                <>
                  <label style={lbl}>Details</label>
                  <textarea style={{...inp,height:52}} value={supReqForm.details} onChange={e=>setSupReqForm({...supReqForm,details:e.target.value})} placeholder="Specify items required in detail"/>
                </>
              )}
              <label style={lbl}>Reason for Request</label>
              <input style={inp} value={supReqForm.reason} onChange={e=>setSupReqForm({...supReqForm,reason:e.target.value})} placeholder="e.g. Item worn out / damaged / new joiner kit"/>
              {(PPE_ISSUED_DATA[selWorker.empCode]||[]).filter(it=>it.item===supReqForm.item).length>0&&(
                <div style={{background:ABG,border:"1px solid #F2C97E",borderRadius:8,padding:"7px 10px",marginTop:6,fontSize:13,color:AMB,fontWeight:600}}>
                  Last issued: {(PPE_ISSUED_DATA[selWorker.empCode]||[]).find(it=>it.item===supReqForm.item).issued} ({monthsElapsed((PPE_ISSUED_DATA[selWorker.empCode]||[]).find(it=>it.item===supReqForm.item).issued)} months ago)
                </div>
              )}
              <div style={{display:"flex",gap:8,marginTop:10}}>
                <button style={{...PBtn,margin:0,flex:2,background:ORANGE,opacity:!supReqForm.reason?0.5:1}} onClick={submitSupReq} disabled={!supReqForm.reason}>Send to Super Admin</button>
                <button
                  onClick={()=>setShowSupReq(false)} style={{flex:1,padding:"10px",border:"1px solid "+LN,borderRadius:10,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:15}}>Cancel</button>
              </div>
            </Card>
          )}
          {issueSent&&<div
            style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"9px 12px",marginBottom:10,fontSize:15,color:GRN,fontWeight:700}}>PPE issued and recorded successfully.</div>}
          {workerRequests.length>0&&(
            <Card s={{marginBottom:10}}>
              <div style={{fontSize:14,fontWeight:800,color:CB,marginBottom:8,textTransform:"uppercase",letterSpacing:.4}}>PPE Request History</div>
              {workerRequests.map((r,i)=>{
                const sc=statusColor(r.status);
                return(
                  <div key={r.id} style={{padding:"8px 0",borderTop:i?"1px solid "+LN:"none"}}>
                    <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center"}}>
                      <div>
                        <div style={{fontSize:15,fontWeight:700}}>{r.item}{r.size?" - "+r.size:""} x{r.qty}</div>
                        <div style={{fontSize:13,color:SFT}}>{r.requestedOn} | {r.reason}</div>
                      </div>
                      <Pill c={r.status} bg={sc.bg} fg={sc.fg}/>
                    </div>
                    {r.status==="Pending Supervisor"&&(
                      <button
                        onClick={()=>forwardToSuperAdmin(r.id)} style={{...PBtn,margin:"6px 0 0",padding:"6px 12px",width:"auto",fontSize:14,background:"linear-gradient(135deg,"+DARK+","+CB+")"}}>Forward to Super Admin</button>
                    )}
                    {r.status==="Approved"&&(
                      <button onClick={()=>openIssueForm(r)} style={{...PBtn,margin:"6px 0 0",padding:"6px 12px",width:"auto",fontSize:14,background:GRN}}>Confirm Issued to Worker</button>
                    )}
                  </div>
                );
              })}
            </Card>
          )}
        </>
      )}
      <Card s={{marginBottom:10,border:"2px solid "+CB}}>
        <div style={{fontSize:15,fontWeight:800,color:CB,marginBottom:8}}>Generate PPE Issued Report</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div>
            <label style={{...lbl,marginTop:0,color:CB}}>Client</label>
            <select style={inp} value={rptClient} onChange={e=>{setRptClient(e.target.value);setRptSite("");}}>
              <option value="">All clients</option>
              {rptClients.map(c=><option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={{...lbl,marginTop:0,color:ORANGE}}>Site</label>
            <select style={inp} value={rptSite} onChange={e=>setRptSite(e.target.value)}>
              <option value="">All sites</option>
              {rptSites.map(s=><option key={s}>{s}</option>)}
            </select>
          </div>
        </div>
        <div style={{background:BG,borderRadius:8,padding:"7px 10px",marginBottom:10,fontSize:14,color:SFT}}>
          {rptWorkers.length} workers | {rptRows.length} PPE items | {rptClient||"All clients"} | {rptSite||"All sites"}
        </div>
        <div style={{display:"flex",gap:8,marginBottom:showRptPreview?10:0}}>
          <button
            onClick={()=>setShowRptPreview(!showRptPreview)} style={{flex:1,padding:"9px",border:"2px solid "+CB,borderRadius:10,color:CB,background:showRptPreview?BBG:PNL,cursor:"pointer",fontWeight:700,fontSize:15}}>{showRptPreview?"Hide Preview":"Preview"}</button>
          <button onClick={downloadPPEExcel} style={{...PBtn,margin:0,flex:1,background:GRN,padding:"9px"}}>Excel</button>
          <button onClick={printPPEReport} style={{...PBtn,margin:0,flex:1,background:"linear-gradient(135deg,"+DARK+","+CB+")",padding:"9px"}}>PDF</button>
        </div>
        {showRptPreview&&(
          <div ref={rptRef} style={{marginTop:10,overflowX:"auto",border:"1px solid "+LN,borderRadius:8,background:"#fff"}}>
            <div style={{padding:"10px 12px"}}>
              <div style={{fontWeight:800,fontSize:17,color:DARK}}>FIXIT MANPOWER SUPPLY L.L.C</div>
              <div style={{fontWeight:700,fontSize:15,color:CB}}>PPE Issued Report</div>
              <div style={{fontSize:14,color:SFT,marginBottom:8}}>{rptClient||"All Clients"} | {rptSite||"All Sites"} | Generated: {new Date().toDateString()}</div>
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:14,minWidth:600}}>
              <thead><tr
                style={{background:DARK}}>{["Worker","Code","Client","Site","Item","Size","Qty","Condition","Issued"].map(h=><th key={h} style={{color:"#fff",padding:"5px 6px",textAlign:"left",fontSize:13,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
              <tbody>{rptRows.map((r,i)=>(
                <tr key={i} style={{background:i%2?BG:"#fff"}}>
                  <td style={{padding:"4px 6px",fontWeight:600}}>{r.worker}</td>
                  <td style={{padding:"4px 6px",color:BLU}}>{r.empCode}</td>
                  <td style={{padding:"4px 6px",color:SFT}}>{r.client}</td>
                  <td style={{padding:"4px 6px",color:SFT}}>{r.site}</td>
                  <td style={{padding:"4px 6px"}}>{r.item}</td>
                  <td style={{padding:"4px 6px",textAlign:"center"}}>{r.size}</td>
                  <td style={{padding:"4px 6px",textAlign:"center"}}>{r.qty}</td>
                  <td style={{padding:"4px 6px"}}>{r.condition}</td>
                  <td style={{padding:"4px 6px",color:SFT}}>{r.issued}</td>
                </tr>
              ))}</tbody>
            </table>
            {rptRows.length===0&&<div style={{padding:"12px",textAlign:"center",color:SFT,fontSize:14}}>No PPE records for selected filters.</div>}
          </div>
        )}
      </Card>

      {allPending.length>0&&(
        <Card s={{marginBottom:12,border:"2px solid "+AMB}}>
          <div style={{fontSize:15,fontWeight:800,color:AMB,marginBottom:8}}>Pending PPE Requests - Action Required</div>
          {allPending.map((r,i)=>(
            <div key={r.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                <div>
                  <div style={{fontSize:15,fontWeight:700}}>{r.worker} <span style={{fontSize:13,color:SFT}}>{r.empCode}</span></div>
                  <div style={{fontSize:14,color:SFT}}>{r.item}{r.size?" - "+r.size:""} x{r.qty} | {r.site}</div>
                  <div style={{fontSize:13,color:SFT,fontStyle:"italic",marginTop:2}}>Reason: {r.reason}</div>
                  <div style={{fontSize:13,color:SFT}}>Requested: {r.requestedOn}</div>
                </div>
                <Pill c={r.status} bg={statusColor(r.status).bg} fg={statusColor(r.status).fg}/>
              </div>
              <button
                onClick={()=>forwardToSuperAdmin(r.id)} style={{...PBtn,margin:0,padding:"7px 14px",width:"auto",fontSize:14,background:"linear-gradient(135deg,"+DARK+","+CB+")"}}>Forward to Super Admin for Approval</button>
            </div>
          ))}
        </Card>
      )}

      {ppeRequests.filter(r=>r.status==="Approved"&&sup.clients.includes(r.client)).length>0&&(
        <div style={{background:"#fff3cd",border:"2px solid "+GRN,borderRadius:8,padding:"8px 12px",marginBottom:8,fontSize:14,color:"#333",fontWeight:700}}>
          Super Admin has approved {ppeRequests.filter(r=>r.status==="Approved"&&sup.clients.includes(r.client)).length} PPE request(s). Please issue to workers below and record the date.
        </div>
      )}
      {ppeRequests.filter(r=>r.status==="Approved"&&sup.clients.includes(r.client)).length>0&&(
        <Card s={{marginBottom:12,border:"2px solid "+GRN}}>
          <div style={{fontSize:15,fontWeight:800,color:GRN,marginBottom:8}}>Approved by Super Admin - Ready to Issue</div>
          {ppeRequests.filter(r=>r.status==="Approved"&&sup.clients.includes(r.client)).map((r,i)=>(
            <div key={r.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                <div>
                  <div style={{fontSize:15,fontWeight:700}}>{r.worker} <span style={{fontSize:13,color:SFT}}>{r.empCode}</span></div>
                  <div style={{fontSize:14,color:SFT}}>{r.item}{r.size?" - "+r.size:""} x{r.qty}</div>
                </div>
                <Pill c="Approved" bg={GBG} fg={GRN}/>
              </div>
              <button onClick={()=>openIssueForm(r)} style={{...PBtn,margin:0,padding:"7px 14px",width:"auto",fontSize:14,background:GRN}}>Issue to Worker - Record Date</button>
            </div>
          ))}
        </Card>
      )}
      {!selWorker&&allPending.length===0&&<div
        style={{textAlign:"center",color:SFT,fontSize:16,padding:"20px 0",fontStyle:"italic"}}>No pending PPE requests. Search a worker to view their PPE record.</div>}
    </>
  );
}

function ComplaintsTab({sup}){
  const [complaints,setComplaints]=useState([]);
  const [replyDraft,setReplyDraft]=useState({});
  const [openReply,setOpenReply]=useState(null);
  const wById={};
  (sup.workers||[]).forEach(w=>{wById[w.dbId||w.id]=w;});
  useEffect(()=>{
    const ids=(sup.workers||[]).map(w=>w.dbId||w.id);
    if(ids.length===0)return;
    sbQuery("complaints","worker_id=in.("+ids.join(",")
      +")&select=*&order=request_date.desc")
      .then(rows=>{
        setComplaints((rows||[]).map(r=>{
          const w=wById[r.worker_id]||{};
          return {id:r.id,worker:w.name||"Worker",
            empCode:w.empCode||"",text:r.details,
            status:r.status||"Open",date:r.request_date,
            reply:r.reply||""};
        }));
      }).catch(()=>{});
  },[]);
  const resolve=async(id)=>{
    try{
      await sbUpdate("complaints","id=eq."+id,{status:"Resolved"});
    }catch(e){}
    setComplaints(complaints.map(c=>c.id===id?{...c,status:"Resolved",reply:replyDraft[id]||c.reply}:c));
    setOpenReply(null);
  };
  const openCount=complaints.filter(c=>c.status==="Open").length;
  return(
    <Card>
      <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:10}}>
        <SHdr title="Worker Complaints" sub="Submitted via worker app"/>
        {openCount>0&&<span style={{background:RBG,color:RED,fontWeight:800,fontSize:14,padding:"3px 10px",borderRadius:20}}>{openCount} open</span>}
      </div>
      {complaints.map((c,i)=>(
        <div key={c.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <div style={{fontSize:16,fontWeight:700}}>{c.worker} <span style={{fontSize:13,color:SFT}}>{c.empCode}</span></div>
              <div style={{fontSize:15,color:INK,marginTop:2}}>{c.text}</div>
              <div style={{fontSize:13,color:SFT,marginTop:2}}>{c.date}</div>
              {c.reply&&<div style={{fontSize:14,color:GRN,fontStyle:"italic",marginTop:4}}>Response: {c.reply}</div>}
            </div>
            <Pill c={c.status} bg={c.status==="Open"?RBG:GBG} fg={c.status==="Open"?RED:GRN}/>
          </div>
          {c.status==="Open"&&(
            openReply===c.id?(
              <div style={{marginTop:8}}>
                <textarea style={{...inp,height:52}} value={replyDraft[c.id]||""} onChange={e=>setReplyDraft({...replyDraft,[c.id]:e.target.value})} placeholder="Response to worker..."/>
                <div style={{display:"flex",gap:8,marginTop:6}}>
                  <button style={{...PBtn,margin:0,flex:2,padding:"7px",fontSize:14,background:GRN}} onClick={()=>resolve(c.id)}>Mark Resolved</button>
                  <button
                    onClick={()=>setOpenReply(null)} style={{flex:1,padding:"7px",border:"1px solid "+LN,borderRadius:8,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:14}}>Cancel</button>
                </div>
              </div>
            ):(
              <button onClick={()=>setOpenReply(c.id)} style={{...PBtn,margin:"8px 0 0",padding:"7px 14px",width:"auto",fontSize:14,background:GRN}}>Respond and Resolve</button>
            )
          )}
        </div>
      ))}
    </Card>
  );
}

function MedicalTab({sup}){
  const [search,setSearch]=useState("");
  const [selWorker,setSelWorker]=useState(null);
  const [showDrop,setShowDrop]=useState(false);
  const [medType,setMedType]=useState("MWP");
  const [medReason,setMedReason]=useState("");
  const [medDate,setMedDate]=useState(new Date().toISOString().slice(0,10));
  const [medFile,setMedFile]=useState(null);
  const [showForm,setShowForm]=useState(false);
  const [sent,setSent]=useState(false);
  const [items,setItems]=useState([
    {id:1,worker:"Bikash Thapa",empCode:"EMP-1045",type:"MWP",date:"10 Jun",workerReason:"Medical certificate submitted",supReason:"Medical certificate submitted",status:"pending",fileName:"medical_cert_bikash.pdf"},
    {id:2,worker:"Suresh Pillai",empCode:"EMP-1043",type:"MWOP",date:"12 Jun",workerReason:"No paper - sick day",supReason:"No paper - sick day",status:"pending",fileName:null},
  ]);
  const [fwdEdit,setFwdEdit]=useState(null);
  const [fwdReason,setFwdReason]=useState("");
  const myW=sup.workers;
  const q=search.toLowerCase();
  const searchRes=showDrop&&q.length>0
    ?myW.filter(w=>w.name.toLowerCase().includes(q)||w.empCode.toLowerCase().includes(q)):[];
  const forward=(id)=>{
    const item=items.find(m=>m.id===id);
    setFwdEdit(id);
    setFwdReason(item?(item.supReason||item.workerReason||""):"");
  };
  const doForward=(id)=>{
    setItems(items.map(m=>m.id===id?{...m,
      supReason:fwdReason,
      workerReason:m.workerReason||m.supReason,
      status:"forwarded"}:m));
    setFwdEdit(null);setFwdReason("");
  };
  const reject=(id)=>{setItems(items.map(m=>m.id===id?{...m,status:"rejected"}:m));};
  const submitNew=async()=>{
    if(!selWorker||!medReason)return;
    try{
      await sbInsert("leave_requests",{
        company_id:sup.companyId,
        worker_id:selWorker.dbId||selWorker.id,
        leave_type:medType==="MWP"?"Medical (with paper)":"Medical (no paper)",
        from_date:medDate,
        to_date:medDate,
        reason:medReason,
        status:"Pending",
        stage:"Pending Super Admin",
        request_date:new Date().toISOString().slice(0,10),
      });
    }catch(e){/* keep UI responsive */}
    setItems([{id:Date.now(),worker:selWorker.name,empCode:selWorker.empCode,type:medType,date:medDate,supReason:medReason,status:"pending",fileName:medFile?medFile.name:null},...items]);
    setSent(true);setShowForm(false);setMedReason("");setSelWorker(null);setSearch("");setMedFile(null);
    setTimeout(()=>setSent(false),3000);
  };
  return(
    <>
      <Card s={{marginBottom:10,border:"2px solid "+BLU}}>
        <SHdr title="Medical" sub="MWP/MWOP submissions - Super Admin approves directly"/>
        {sent&&<div
          style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:GRN,fontWeight:700}}>Medical submission sent to Super Admin.</div>}
        {!showForm?(
          <button onClick={()=>setShowForm(true)} style={{...PBtn,background:BLU,marginTop:0}}>+ New Medical Submission</button>
        ):(
          <>
            <label style={lbl}>Search Worker</label>
            <input style={inp} value={search}
              onChange={e=>{
            setSearch(e.target.value);setShowDrop(true);
            if(!e.target.value)setSelWorker(null);
          }} onFocus={()=>setShowDrop(true)} placeholder="Type name or emp code..."/>
            {searchRes.map((w,i)=>(
              <div key={w.id}
                onClick={()=>{
              setSelWorker(w);setSearch(w.name);setShowDrop(false);
            }} style={{padding:"6px 0",borderTop:i?"1px solid "+LN:"none",cursor:"pointer",fontSize:15,fontWeight:600}}>{w.name} <span style={{color:SFT,fontSize:13}}>{w.empCode}</span></div>
            ))}
            <label style={lbl}>Type</label>
            <select style={inp} value={medType} onChange={e=>{setMedType(e.target.value);if(e.target.value!=="MWP")setMedFile(null);}}>
              <option value="MWP">Medical With Paper (MWP)</option>
              <option value="MWOP">Medical Without Paper (MWOP)</option>
            </select>
            <label style={lbl}>Date</label>
            <input type="date" style={inp} value={medDate} onChange={e=>setMedDate(e.target.value)}/>
            {medType==="MWP"&&(
              <>
                <label style={lbl}>Attach Medical Paper</label>
                <input type="file" accept=".pdf,.jpg,.jpeg,.png" onChange={e=>setMedFile(e.target.files[0]||null)} style={{...inp,padding:"6px 10px"}}/>
                {medFile&&<div style={{fontSize:14,color:GRN,fontWeight:600,marginTop:4}}>Attached: {medFile.name}</div>}
                {!medFile&&<div style={{fontSize:13,color:SFT,marginTop:4,fontStyle:"italic"}}>Upload the worker's medical certificate - visible to Super Admin during approval.</div>}
              </>
            )}
            <label style={lbl}>Reason</label>
            <input style={inp} value={medReason} onChange={e=>setMedReason(e.target.value)} placeholder="e.g. Medical certificate submitted"/>
            <div style={{display:"flex",gap:8,marginTop:10}}>
              <button style={{...PBtn,margin:0,flex:2,background:BLU,opacity:(!selWorker||!medReason)?0.5:1}} onClick={submitNew} disabled={!selWorker||!medReason}>Submit</button>
              <button
                onClick={()=>setShowForm(false)} style={{flex:1,padding:"10px",border:"1px solid "+LN,borderRadius:10,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:15}}>Cancel</button>
            </div>
          </>
        )}
      </Card>
      {items.filter(m=>m.status==="pending").length>0&&(
        <Card s={{marginBottom:10,border:"2px solid "+AMB}}>
          <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
            <div style={{fontSize:15,fontWeight:800,color:AMB}}>Pending Approvals</div>
            <span style={{background:ABG,color:AMB,fontWeight:800,fontSize:14,padding:"3px 10px",borderRadius:20}}>{items.filter(m=>m.status==="pending").length}</span>
          </div>
          {items.filter(m=>m.status==="pending").map((m,i)=>(
            <div key={m.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{fontSize:16,fontWeight:700}}>{m.worker} <span style={{fontSize:13,color:SFT}}>{m.empCode}</span></div>
              <div style={{fontSize:14,color:BLU,fontWeight:600}}>{m.type} | {m.date}</div>
              <div style={{fontSize:13,color:SFT,marginTop:4}}>Worker's reason:</div>
              <div style={{fontSize:14,color:INK,fontStyle:"italic"}}>"{m.workerReason||m.supReason}"</div>
              {m.fileName?(
                <div style={{display:"flex",alignItems:"center",gap:6,marginTop:6,background:BBG,borderRadius:7,padding:"5px 8px"}}>
                  <span style={{fontSize:17}}>[PDF]</span>
                  <span style={{fontSize:14,color:BLU,fontWeight:600}}>{m.fileName}</span>
                </div>
              ):m.type==="MWP"&&(
                <div style={{fontSize:13,color:RED,fontWeight:600,marginTop:4}}>No medical paper attached</div>
              )}
              {fwdEdit===m.id?(
                <div style={{marginTop:8,background:BBG,borderRadius:8,padding:"10px"}}>
                  <label style={lbl}>Reason to forward (you can edit / clarify)</label>
                  <textarea style={{...inp,minHeight:56}} value={fwdReason} onChange={e=>setFwdReason(e.target.value)} placeholder="Review or clarify the reason before forwarding..."/>
                  <div style={{fontSize:12,color:SFT,marginBottom:6}}>The worker's original reason is kept on record. Super Admin will see this reason.</div>
                  <div style={{display:"flex",gap:8}}>
                    <button style={{...PBtn,margin:0,flex:2,padding:"7px",fontSize:14,background:BLU,opacity:fwdReason.trim()?1:0.5}} disabled={!fwdReason.trim()} onClick={()=>doForward(m.id)}>Confirm & Forward</button>
                    <button style={{flex:1,padding:"7px",border:"1px solid "+LN,borderRadius:8,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:14}} onClick={()=>{setFwdEdit(null);setFwdReason("");}}>Cancel</button>
                  </div>
                </div>
              ):(
                <div style={{display:"flex",gap:8,marginTop:8}}>
                  <button style={{...PBtn,margin:0,flex:2,padding:"7px",fontSize:14,background:BLU}} onClick={()=>forward(m.id)}>Review & Forward to Super Admin</button>
                  <button style={{flex:1,padding:"7px",border:"1px solid "+RED,borderRadius:8,background:RBG,color:RED,cursor:"pointer",fontWeight:700,fontSize:14}}
                    onClick={()=>reject(m.id)}>Reject</button>
                </div>
              )}
            </div>
          ))}
        </Card>
      )}
      {items.filter(m=>m.status!=="pending").length>0&&(
        <Card>
          <div style={{fontSize:15,fontWeight:800,color:SFT,marginBottom:8}}>Past Submissions</div>
          {items.filter(m=>m.status!=="pending").map((m,i)=>(
            <div key={m.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{fontSize:16,fontWeight:700}}>{m.worker} <span style={{fontSize:13,color:SFT}}>{m.empCode}</span></div>
                  <div style={{fontSize:14,color:BLU,fontWeight:600}}>{m.type} | {m.date}</div>
                  {m.workerReason&&m.workerReason!==m.supReason&&(
                    <div style={{fontSize:13,color:SFT,marginTop:2}}>Worker: <span style={{fontStyle:"italic"}}>"{m.workerReason}"</span></div>
                  )}
                  <div style={{fontSize:14,color:SFT,fontStyle:"italic",marginTop:2}}>{m.status==="forwarded"?"Forwarded":"Reason"}: "{m.supReason}"</div>
                  {m.fileName&&<div style={{fontSize:14,color:BLU,marginTop:4}}>[PDF] {m.fileName}</div>}
                </div>
                <Pill c={m.status==="forwarded"?"Sent to Super Admin":"Rejected"} bg={m.status==="forwarded"?BBG:RBG} fg={m.status==="forwarded"?BLU:RED}/>
              </div>
            </div>
          ))}
        </Card>
      )}
    </>
  );
}


function IdleTab({sup,goToOverride=()=>{}}){
  const [search,setSearch]=useState("");
  const [selWorker,setSelWorker]=useState(null);
  const [showDrop,setShowDrop]=useState(false);
  const [idleReason,setIdleReason]=useState("");
  const [idleDate,setIdleDate]=useState(new Date().toISOString().slice(0,10));
  const [showForm,setShowForm]=useState(false);
  const [sent,setSent]=useState(false);
  const [items,setItems]=useState([
    {id:1,worker:"Suresh Pillai",empCode:"EMP-1043",date:"14 May",workerReason:"No work assigned by client",supReason:"No work assigned by client",status:"pending"},
  ]);
  const [fwdEdit,setFwdEdit]=useState(null);
  const [fwdReason,setFwdReason]=useState("");
  const myW=sup.workers;
  const q=search.toLowerCase();
  const searchRes=showDrop&&q.length>0
    ?myW.filter(w=>w.name.toLowerCase().includes(q)||w.empCode.toLowerCase().includes(q)):[];
  const forward=(id)=>{
    const item=items.find(m=>m.id===id);
    setFwdEdit(id);
    setFwdReason(item?(item.supReason||item.workerReason||""):"");
  };
  const doForward=(id)=>{
    setItems(items.map(m=>m.id===id?{...m,
      supReason:fwdReason,
      workerReason:m.workerReason||m.supReason,
      status:"forwarded"}:m));
    setFwdEdit(null);setFwdReason("");
  };
  const reject=(id)=>{setItems(items.map(m=>m.id===id?{...m,status:"rejected"}:m));};
  const submitNew=async()=>{
    if(!selWorker||!idleReason)return;
    try{
      await sbInsert("attendance",{
        company_id:sup.companyId,
        worker_id:selWorker.dbId||selWorker.id,
        att_date:idleDate,
        sup_status:"I",
        override_reason:idleReason,
        override_state:"pending",
        override_by:sup.id,
      });
    }catch(e){/* keep UI responsive */}
    setItems([{id:Date.now(),worker:selWorker.name,empCode:selWorker.empCode,date:idleDate,supReason:idleReason,status:"pending"},...items]);
    setSent(true);setShowForm(false);setIdleReason("");setSelWorker(null);setSearch("");
    setTimeout(()=>setSent(false),3000);
  };
  return(
    <>
      <Card s={{marginBottom:10,border:"2px solid "+YFG}}>
        <SHdr title="Idle" sub="Idle day submissions - Super Admin approves directly"/>
        <div style={{background:BBG,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:BLU}}>
          Marking idle for one worker? Use the form below.{" "}
          <span onClick={goToOverride} style={{fontWeight:700,textDecoration:"underline",cursor:"pointer"}}>Need to mark idle for multiple workers at once? Use Override instead.</span>
        </div>
        {sent&&<div style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:GRN,fontWeight:700}}>Idle submission sent to Super Admin.</div>}
        {!showForm?(
          <button onClick={()=>setShowForm(true)} style={{...PBtn,background:YFG,marginTop:0}}>+ New Idle Submission</button>
        ):(
          <>
            <label style={lbl}>Search Worker</label>
            <input style={inp} value={search}
              onChange={e=>{
            setSearch(e.target.value);setShowDrop(true);
            if(!e.target.value)setSelWorker(null);
          }} onFocus={()=>setShowDrop(true)} placeholder="Type name or emp code..."/>
            {searchRes.map((w,i)=>(
              <div key={w.id}
                onClick={()=>{
              setSelWorker(w);setSearch(w.name);setShowDrop(false);
            }} style={{padding:"6px 0",borderTop:i?"1px solid "+LN:"none",cursor:"pointer",fontSize:15,fontWeight:600}}>{w.name} <span style={{color:SFT,fontSize:13}}>{w.empCode}</span></div>
            ))}
            <label style={lbl}>Date</label>
            <input type="date" style={inp} value={idleDate} onChange={e=>setIdleDate(e.target.value)}/>
            <label style={lbl}>Reason</label>
            <input style={inp} value={idleReason} onChange={e=>setIdleReason(e.target.value)} placeholder="e.g. No work assigned by client"/>
            <div style={{display:"flex",gap:8,marginTop:10}}>
              <button style={{...PBtn,margin:0,flex:2,background:YFG,opacity:(!selWorker||!idleReason)?0.5:1}} onClick={submitNew} disabled={!selWorker||!idleReason}>Submit</button>
              <button
                onClick={()=>setShowForm(false)} style={{flex:1,padding:"10px",border:"1px solid "+LN,borderRadius:10,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:15}}>Cancel</button>
            </div>
          </>
        )}
      </Card>
      {items.filter(m=>m.status==="pending").length>0&&(
        <Card s={{marginBottom:10,border:"2px solid "+AMB}}>
          <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
            <div style={{fontSize:15,fontWeight:800,color:AMB}}>Pending Approvals</div>
            <span style={{background:ABG,color:AMB,fontWeight:800,fontSize:14,padding:"3px 10px",borderRadius:20}}>{items.filter(m=>m.status==="pending").length}</span>
          </div>
          {items.filter(m=>m.status==="pending").map((m,i)=>(
            <div key={m.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{fontSize:16,fontWeight:700}}>{m.worker} <span style={{fontSize:13,color:SFT}}>{m.empCode}</span></div>
              <div style={{fontSize:14,color:SFT}}>{m.date}</div>
              <div style={{fontSize:13,color:SFT,marginTop:4}}>Worker's reason:</div>
              <div style={{fontSize:14,color:INK,fontStyle:"italic"}}>"{m.workerReason||m.supReason}"</div>
              {fwdEdit===m.id?(
                <div style={{marginTop:8,background:"#FFF9E6",borderRadius:8,padding:"10px"}}>
                  <label style={lbl}>Reason to forward (you can edit / clarify)</label>
                  <textarea style={{...inp,minHeight:56}} value={fwdReason} onChange={e=>setFwdReason(e.target.value)} placeholder="Review or clarify the reason before forwarding..."/>
                  <div style={{fontSize:12,color:SFT,marginBottom:6}}>The worker's original reason is kept on record. Super Admin will see this reason.</div>
                  <div style={{display:"flex",gap:8}}>
                    <button style={{...PBtn,margin:0,flex:2,padding:"7px",fontSize:14,background:YFG,opacity:fwdReason.trim()?1:0.5}} disabled={!fwdReason.trim()} onClick={()=>doForward(m.id)}>Confirm & Forward</button>
                    <button style={{flex:1,padding:"7px",border:"1px solid "+LN,borderRadius:8,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:14}} onClick={()=>{setFwdEdit(null);setFwdReason("");}}>Cancel</button>
                  </div>
                </div>
              ):(
                <div style={{display:"flex",gap:8,marginTop:8}}>
                  <button style={{...PBtn,margin:0,flex:2,padding:"7px",fontSize:14,background:YFG}} onClick={()=>forward(m.id)}>Review & Send to Super Admin</button>
                  <button style={{flex:1,padding:"7px",border:"1px solid "+RED,borderRadius:8,background:RBG,color:RED,cursor:"pointer",fontWeight:700,fontSize:14}}
                    onClick={()=>reject(m.id)}>Reject</button>
                </div>
              )}
            </div>
          ))}
        </Card>
      )}
      {items.filter(m=>m.status!=="pending").length>0&&(
        <Card>
          <div style={{fontSize:15,fontWeight:800,color:SFT,marginBottom:8}}>Past Submissions</div>
          {items.filter(m=>m.status!=="pending").map((m,i)=>(
            <div key={m.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{fontSize:16,fontWeight:700}}>{m.worker} <span style={{fontSize:13,color:SFT}}>{m.empCode}</span></div>
                  <div style={{fontSize:14,color:SFT}}>{m.date}</div>
                  <div style={{fontSize:14,color:SFT,fontStyle:"italic",marginTop:2}}>"{m.supReason}"</div>
                </div>
                <Pill c={m.status==="forwarded"?"Sent to Super Admin":"Rejected"} bg={m.status==="forwarded"?BBG:RBG} fg={m.status==="forwarded"?BLU:RED}/>
              </div>
            </div>
          ))}
        </Card>
      )}
    </>
  );
}

function AdvanceTab({sup}){
  const [search,setSearch]=useState("");
  const [selWorker,setSelWorker]=useState(null);
  const [showDrop,setShowDrop]=useState(false);
  const [advAmount,setAdvAmount]=useState("");
  const [advReason,setAdvReason]=useState("");
  const [showForm,setShowForm]=useState(false);
  const [sent,setSent]=useState(false);
  const [requests,setRequests]=useState([
    {id:1,worker:"Vikram Nair",empCode:"EMP-1044",requested:180,revised:180,reason:"Family emergency",status:"draft"},
    {id:2,worker:"Bikash Thapa",empCode:"EMP-1045",requested:350,revised:350,reason:"Medical expense",status:"draft"},
  ]);
  const myW=sup.workers;
  const q=search.toLowerCase();
  const searchRes=showDrop&&q.length>0
    ?myW.filter(w=>w.name.toLowerCase().includes(q)||w.empCode.toLowerCase().includes(q)):[];
  const getRouting=(amt)=>amt<=200?"Admin approval only":"Super Admin approval required";
  const setRevised=(i,val)=>setRequests(requests.map((r,j)=>j===i?{...r,revised:Number(val)||0}:r));
  const forward=(i)=>setRequests(requests.map((r,j)=>j===i?{...r,status:"Sent to Admin"}:r));
  const reject=(i)=>setRequests(requests.map((r,j)=>j===i?{...r,status:"Rejected"}:r));
  const submitNew=async()=>{
    if(!selWorker||!advAmount||!advReason)return;
    try{
      await sbInsert("advances",{
        company_id:sup.companyId,
        worker_id:selWorker.dbId||selWorker.id,
        amount:Number(advAmount),
        reason:advReason,
        status:"Pending",
        requested_by:sup.name,
        request_date:new Date().toISOString().slice(0,10),
      });
    }catch(e){/* keep UI responsive even if save fails */}
    setRequests([{id:Date.now(),worker:selWorker.name,empCode:selWorker.empCode,requested:Number(advAmount),revised:Number(advAmount),reason:advReason,status:"draft"},...requests]);
    setSent(true);setShowForm(false);setAdvAmount("");setAdvReason("");setSelWorker(null);setSearch("");
    setTimeout(()=>setSent(false),3000);
  };
  return(
    <>
      <Card s={{marginBottom:10,border:"2px solid "+ORANGE}}>
        <SHdr title="Advances" sub="AED 200 or below routes to Admin. Above AED 200 routes to Super Admin."/>
        {sent&&<div style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:GRN,fontWeight:700}}>Advance request raised.</div>}
        {!showForm?(
          <button onClick={()=>setShowForm(true)} style={{...PBtn,background:ORANGE,marginTop:0}}>+ Raise Advance Request</button>
        ):(
          <>
            <label style={lbl}>Search Worker</label>
            <input style={inp} value={search}
              onChange={e=>{
            setSearch(e.target.value);setShowDrop(true);
            if(!e.target.value)setSelWorker(null);
          }} onFocus={()=>setShowDrop(true)} placeholder="Type name or emp code..."/>
            {searchRes.map((w,i)=>(
              <div key={w.id}
                onClick={()=>{
              setSelWorker(w);setSearch(w.name);setShowDrop(false);
            }} style={{padding:"6px 0",borderTop:i?"1px solid "+LN:"none",cursor:"pointer",fontSize:15,fontWeight:600}}>{w.name} <span style={{color:SFT,fontSize:13}}>{w.empCode}</span></div>
            ))}
            <label style={lbl}>Amount (AED)</label>
            <input type="number" style={inp} value={advAmount} onChange={e=>setAdvAmount(e.target.value)} placeholder="e.g. 200"/>
            {advAmount&&<div style={{fontSize:13,color:AMB,fontWeight:600,marginTop:2}}>{getRouting(Number(advAmount))}</div>}
            <label style={lbl}>Reason</label>
            <input style={inp} value={advReason} onChange={e=>setAdvReason(e.target.value)} placeholder="e.g. Family emergency"/>
            <div style={{display:"flex",gap:8,marginTop:10}}>
              <button style={{...PBtn,margin:0,flex:2,background:ORANGE,opacity:(!selWorker||!advAmount||!advReason)?0.5:1}}
                onClick={submitNew} disabled={!selWorker||!advAmount||!advReason}>Submit</button>
              <button
                onClick={()=>setShowForm(false)} style={{flex:1,padding:"10px",border:"1px solid "+LN,borderRadius:10,background:PNL,color:SFT,cursor:"pointer",fontWeight:700,fontSize:15}}>Cancel</button>
            </div>
          </>
        )}
      </Card>
      {requests.filter(r=>r.status==="draft").length>0&&(
        <Card s={{marginBottom:10,border:"2px solid "+AMB}}>
          <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginBottom:8}}>
            <div style={{fontSize:15,fontWeight:800,color:AMB}}>Pending Approvals</div>
            <span style={{background:ABG,color:AMB,fontWeight:800,fontSize:14,padding:"3px 10px",borderRadius:20}}>{requests.filter(r=>r.status==="draft").length}</span>
          </div>
          {requests.filter(r=>r.status==="draft").map((r,i)=>(
            <div key={r.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{fontSize:16,fontWeight:700}}>{r.worker} <span style={{fontSize:13,color:SFT}}>{r.empCode}</span></div>
              <div style={{fontSize:14,color:SFT}}>Requested: AED {r.requested} | {r.reason}</div>
              <div style={{fontSize:13,color:AMB,fontWeight:600,marginTop:2}}>{getRouting(r.revised)}</div>
              <div style={{display:"flex",alignItems:"center",gap:6,marginTop:6}}>
                <span style={{fontSize:14,color:SFT}}>Revised amount:</span>
                <input type="number" value={r.revised} onChange={e=>setRevised(requests.indexOf(r),e.target.value)} style={{...inp,width:90,padding:"4px 8px",fontSize:15}}/>
              </div>
              <div style={{display:"flex",gap:8,marginTop:8}}>
                <button style={{...PBtn,margin:0,flex:2,padding:"7px",fontSize:14,background:ORANGE}}
                  onClick={()=>forward(requests.indexOf(r))}>Forward to {r.revised<=200?"Admin":"Super Admin"} - AED {r.revised}</button>
                <button style={{flex:1,padding:"7px",border:"1px solid "+RED,borderRadius:8,background:RBG,color:RED,cursor:"pointer",fontWeight:700,fontSize:14}}
                  onClick={()=>reject(requests.indexOf(r))}>Reject</button>
              </div>
            </div>
          ))}
        </Card>
      )}
      {requests.filter(r=>r.status!=="draft").length>0&&(
        <Card>
          <div style={{fontSize:15,fontWeight:800,color:SFT,marginBottom:8}}>Past Requests</div>
          {requests.filter(r=>r.status!=="draft").map((r,i)=>(
            <div key={r.id} style={{padding:"10px 0",borderTop:i?"1px solid "+LN:"none"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{fontSize:16,fontWeight:700}}>{r.worker} <span style={{fontSize:13,color:SFT}}>{r.empCode}</span></div>
                  <div style={{fontSize:14,color:SFT}}>Requested: AED {r.requested} | {r.reason}</div>
                </div>
                <Pill c={r.status==="Sent to Admin"?"Sent to "+(r.revised<=200?"Admin":"Super Admin"):"Rejected"} bg={r.status==="Sent to Admin"?BBG:RBG} fg={r.status==="Sent to Admin"?BLU:RED}/>
              </div>
            </div>
          ))}
        </Card>
      )}
    </>
  );
}


function TempStatusPanel({sup}){
  const [search,setSearch]=useState("");
  const [selWorker,setSelWorker]=useState(null);
  const [showDrop,setShowDrop]=useState(false);
  const [tempStatus,setTempStatus]=useState("Emergency Leave");
  const [startDate,setStartDate]=useState(new Date().toISOString().slice(0,10));
  const [endDate,setEndDate]=useState("");
  const [note,setNote]=useState("");
  const [saved,setSaved]=useState(false);
  const [savedWorkers,setSavedWorkers]=useState([
    {empCode:"EMP-1044",name:"Vikram Nair",status:"Idle",since:"01 Jul 2026",note:"Completed DEWA site. Awaiting new assignment.",active:true},
  ]);
  const myW=sup.workers;
  const q=search.toLowerCase();
  const searchRes=showDrop&&q.length>0
    ?myW.filter(w=>w.name.toLowerCase().includes(q)||w.empCode.toLowerCase().includes(q)):[];
  const TEMP_STATUSES=["Emergency Leave","On Training","Idle","Absconder","Medical Leave","Annual Leave","Refuse to Work","Resignation"];
  const submit=()=>{
    if(!selWorker) return;
    setSavedWorkers([...savedWorkers.filter(w=>w.empCode!==selWorker.empCode),{
      empCode:selWorker.empCode,name:selWorker.name,
      status:tempStatus,since:startDate,endDate:endDate||"Until reassigned",note:note,active:true
    }]);
    setSaved(true);
    setTimeout(()=>{setSaved(false);setSelWorker(null);setSearch("");setNote("");setEndDate("");},3000);
  };
  const clearStatus=(empCode)=>{
    setSavedWorkers(savedWorkers.map(w=>w.empCode===empCode?{...w,active:false}:w));
  };
  return(
    <Card s={{marginTop:12,border:"2px solid "+ORANGE}}>
      <SHdr title="Temporary Worker Status" sub="Set holding status - Admin assigns new client or site when ready"/>
      <div style={{background:ABG,border:"1px solid #F2C97E",borderRadius:8,padding:"7px 10px",marginBottom:10,fontSize:13,color:AMB,fontWeight:600}}>
        This status is temporary. Only Admin can assign workers to a new client or site. Status clears automatically when Admin makes a new assignment.
      </div>
      {saved&&<div style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:GRN,fontWeight:700}}>Temporary status saved.</div>}
      <label style={lbl}>Search Worker</label>
      <input style={inp} value={search}
        onChange={e=>{
            const v=e.target.value;
            setSearch(v);setShowDrop(!!v);
            if(!v){setSelWorker(null);}
          }} onFocus={()=>setShowDrop(true)} placeholder="Type name or emp code..."/>
      {searchRes.map((w,i)=>(
        <div key={w.id}
          onClick={()=>{setSelWorker(w);setSearch(w.name+" ("+w.empCode+")");setShowDrop(false);}} style={{display:"flex",justifyContent:"space-between",
              alignItems:"center",padding:"7px 0",
              borderTop:i?"1px solid "+LN:"none",
              cursor:"pointer",marginTop:4}}>
          <div><div style={{fontSize:16,fontWeight:700}}>{w.name}</div><div style={{fontSize:14,color:SFT}}>{w.empCode} | {w.client} | {w.site}</div></div>
          <Pill c="Select" bg={BBG} fg={BLU}/>
        </div>
      ))}
      {selWorker&&(
        <>
          <div style={{background:BBG,borderRadius:8,padding:"8px 10px",marginBottom:10,marginTop:6,fontSize:14,color:BLU,fontWeight:600}}>
            {selWorker.name} | {selWorker.empCode} | Current site: {selWorker.site}
          </div>
          <label style={lbl}>Temporary Status</label>
          <select style={inp} value={tempStatus} onChange={e=>setTempStatus(e.target.value)}>
            {TEMP_STATUSES.map(s=><option key={s}>{s}</option>)}
          </select>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <div>
              <label style={lbl}>Effective From</label>
              <input type="date" style={inp} value={startDate} onChange={e=>setStartDate(e.target.value)}/>
            </div>
            <div>
              <label style={lbl}>Expected Until (optional)</label>
              <input type="date" style={inp} value={endDate} onChange={e=>setEndDate(e.target.value)}/>
            </div>
          </div>
          <label style={lbl}>Reason / Notes</label>
          <input style={inp} value={note} onChange={e=>setNote(e.target.value)} placeholder="e.g. Completed BKC site. Awaiting Al Futtaim new allocation."/>
          <button style={{...PBtn,background:ORANGE,marginTop:10,marginBottom:0}} onClick={submit}>Save Temporary Status</button>
        </>
      )}
      {savedWorkers.filter(w=>w.active).length>0&&(
        <div style={{marginTop:12,paddingTop:12,borderTop:"1px solid "+LN}}>
          <div style={{fontSize:14,fontWeight:800,color:ORANGE,marginBottom:8,textTransform:"uppercase",letterSpacing:.4}}>Workers on Temporary Status</div>
          {savedWorkers.filter(w=>w.active).map((w,i)=>(
            <div key={i} style={{background:OBG,border:"1px solid "+ORANGE,borderRadius:10,padding:"10px 12px",marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:INK}}>{w.name} <span style={{fontSize:13,color:SFT}}>{w.empCode}</span></div>
                <div style={{fontSize:14,color:ORANGE,fontWeight:600,marginTop:2}}>{w.status}</div>
                <div style={{fontSize:13,color:SFT,marginTop:2}}>Since: {w.since} | Until: {w.endDate}</div>
                {w.note&&<div style={{fontSize:13,color:SFT,fontStyle:"italic",marginTop:2}}>{w.note}</div>}
              </div>
              <button
                onClick={()=>clearStatus(w.empCode)} style={{fontSize:13,fontWeight:700,padding:"4px 8px",borderRadius:7,border:"1px solid "+RED,color:RED,background:RBG,cursor:"pointer",whiteSpace:"nowrap"}}>Clear Status</button>
            </div>
          ))}
          <div style={{fontSize:13,color:SFT,fontStyle:"italic",marginTop:4}}>All statuses above will be cleared when Admin assigns the worker to a new client or site.</div>
        </div>
      )}
    </Card>
  );
}

function OverrideTab({sup,onInactiveChange,gradeOverrides,setGradeOverrides}){
  const [goSearch,setGoSearch]=useState("");
  const [goSelWorker,setGoSelWorker]=useState(null);
  const [goShowDrop,setGoShowDrop]=useState(false);
  const [goGrade,setGoGrade]=useState("green");
  const [goReason,setGoReason]=useState("");
  const [goSent,setGoSent]=useState(false);
  const [search,setSearch]=useState("");
  const [rClient,setRClient]=useState("");
  const [rSite,setRSite]=useState("");
  const [selectedIds,setSelectedIds]=useState([]);
  const [selMonth,setSelMonth]=useState("Jun 2026");
  const [selDates,setSelDates]=useState([]);
  const [newStatus,setNewStatus]=useState("Idle");
  const [reason,setReason]=useState("");
  const [pending,setPending]=useState([]);
  const [sent,setSent]=useState(false);
  const myW=sup.workers;
  const goSearchRes=goShowDrop?myW.filter(w=>goSearch.length>0&&(w.name.toLowerCase().includes(goSearch.toLowerCase())||w.empCode.toLowerCase().includes(goSearch.toLowerCase()))):[];
  const submitGradeOverride=()=>{
    if(!goSelWorker||!goReason)return;
    setGradeOverrides({...gradeOverrides,[goSelWorker.empCode]:goGrade});
    setGoSent(true);setGoSelWorker(null);setGoSearch("");setGoReason("");
    setTimeout(()=>setGoSent(false),3000);
  };
  const clients=[...new Set(myW.map(w=>w.client))];
  const sites=[...new Set(myW.filter(w=>!rClient||w.client===rClient).map(w=>w.site))];
  const listedWorkers=myW.filter(w=>
    (!rClient||w.client===rClient)&&
    (!rSite||w.site===rSite)&&
    (!search||w.name.toLowerCase().includes(search.toLowerCase())||w.empCode.toLowerCase().includes(search.toLowerCase()))
  );
  const toggleWorker=(id)=>setSelectedIds(s=>s.includes(id)?s.filter(x=>x!==id):[...s,id]);
  const selectAll=()=>setSelectedIds(listedWorkers.map(w=>w.id));
  const clearSelection=()=>setSelectedIds([]);
  const selectedWorkers=myW.filter(w=>selectedIds.includes(w.id));
  const toggleDate=(d)=>setSelDates(s=>s.includes(d)?s.filter(x=>x!==d):[...s,d]);
  const MONTHS_LIST=["Jun 2026","May 2026","Apr 2026","Mar 2026"];
  const moIdx=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(selMonth.split(" ")[0]);
  const yr=Number(selMonth.split(" ")[1]);
  const dim=moIdx>=0&&yr?new Date(yr,moIdx+1,0).getDate():30;
  const firstDay=moIdx>=0&&yr?new Date(yr,moIdx,1).getDay():0;
  const submitOverride=async()=>{
    if(selectedWorkers.length===0||selDates.length===0)return;
    // Write one attendance override row per worker per selected date
    const stCode=newStatus==="Idle"?"I":newStatus==="Absent"?"A":
      newStatus==="Present"?"P":newStatus==="Medical"?"M":
      newStatus==="Holiday"?"H":newStatus==="Weekly Off"?"WO":"I";
    try{
      const rows=[];
      selectedWorkers.forEach(w=>{
        selDates.forEach(d=>{
          const dd=String(d).padStart(2,"0");
          const mm=String(moIdx+1).padStart(2,"0");
          rows.push({
            company_id:sup.companyId,
            worker_id:w.dbId||w.id,
            att_date:yr+"-"+mm+"-"+dd,
            override_status:stCode,
            override_state:"pending",
            override_reason:reason||newStatus,
            override_by:sup.id,
          });
        });
      });
      if(rows.length>0)await sbInsert("attendance",rows);
    }catch(e){/* keep UI responsive */}
    const newItems=selectedWorkers.map(w=>({worker:w.name,empCode:w.empCode,month:selMonth,dates:selDates,status:newStatus,reason}));
    setPending([...pending,...newItems]);
    setSent(true);
    setSelDates([]);setReason("");setSelectedIds([]);
    setTimeout(()=>setSent(false),3000);
    if(["Emergency Leave","Absconder","Annual Leave","On Training","Refusal to work","Resignation / Under Cancellation"].includes(newStatus)&&onInactiveChange){
      const map={"Emergency Leave":"emergencyLeave","Absconder":"absconder","Annual Leave":"annualLeave","On Training":"training","Refusal to work":"refusal","Resignation / Under Cancellation":"resignation"};
      onInactiveChange(map[newStatus],selDates.length*selectedWorkers.length);
    }
  };
  return(
    <>
      <Card s={{marginBottom:10}}>
        <SHdr title="Step 1: Select Worker(s)" sub="Choose client and site to list workers, or search directly"/>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
          <div>
            <label style={{...lbl,marginTop:0,color:CB}}>Client</label>
            <select style={inp} value={rClient} onChange={e=>{setRClient(e.target.value);setRSite("");setSelectedIds([]);}}>
              <option value="">All clients</option>
              {clients.map(c=><option key={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={{...lbl,marginTop:0,color:ORANGE}}>Site</label>
            <select style={inp} value={rSite} onChange={e=>{setRSite(e.target.value);setSelectedIds([]);}}>
              <option value="">All sites</option>
              {sites.map(s=><option key={s}>{s}</option>)}
            </select>
          </div>
        </div>
        <input style={inp} value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by name or employee code (optional)..."/>
        {listedWorkers.length>0&&(
          <div style={{display:"flex",justifyContent:"space-between",
                    alignItems:"center",marginTop:10,marginBottom:4}}>
            <span style={{fontSize:14,color:SFT}}>{listedWorkers.length} worker(s) | {selectedIds.length} selected</span>
            <div style={{display:"flex",gap:10}}>
              <button onClick={selectAll} style={{fontSize:14,fontWeight:700,color:CB,background:"none",border:"none",cursor:"pointer"}}>Select all</button>
              {selectedIds.length>0&&<button onClick={clearSelection} style={{fontSize:14,fontWeight:700,color:SFT,background:"none",border:"none",cursor:"pointer"}}>Clear</button>}
            </div>
          </div>
        )}
        <div style={{maxHeight:280,overflowY:"auto"}}>
          {listedWorkers.map((w,i)=>(
            <div key={w.id} onClick={()=>toggleWorker(w.id)} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderTop:i?"1px solid "+LN:"none",cursor:"pointer"}}>
              <input type="checkbox" checked={selectedIds.includes(w.id)} onChange={()=>toggleWorker(w.id)} onClick={e=>e.stopPropagation()}/>
              <div style={{flex:1}}>
                <div style={{fontSize:16,fontWeight:700}}>{w.name} <span style={{fontSize:13,color:SFT,fontWeight:600}}>{w.empCode}</span></div>
                <div style={{fontSize:13,color:SFT}}>{w.trade} | {w.site}</div>
              </div>
            </div>
          ))}
        </div>
        {rClient===""&&rSite===""&&!search&&<div style={{fontSize:14,color:SFT,fontStyle:"italic",marginTop:8}}>Select a client/site or type to search for workers.</div>}
      </Card>
      {selectedIds.length>0&&(
        <div style={{background:BBG,borderRadius:10,padding:"8px 12px",marginBottom:10,fontSize:15,color:BLU,fontWeight:700}}>
          {selectedIds.length} worker(s) selected: {selectedWorkers.map(w=>w.name).join(", ")}
        </div>
      )}
      <Card s={{marginBottom:10}}>
        <SHdr title="Step 2: Select Dates" sub="Click to toggle multiple days - applies to all selected workers"/>
        <label style={lbl}>Month</label>
        <select style={inp} value={selMonth} onChange={e=>{setSelMonth(e.target.value);setSelDates([]);}}>
          {MONTHS_LIST.map(m=><option key={m}>{m}</option>)}
        </select>
        {selectedWorkers.length===1&&(
          <div style={{background:BBG,borderRadius:8,padding:"7px 10px",marginTop:8,marginBottom:2,fontSize:14,color:BLU,fontWeight:600}}>
            Showing {selectedWorkers[0].name}'s current attendance. Click a day to mark it for override.
          </div>
        )}
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginTop:8,marginBottom:2}}>
          {["Su","Mo","Tu","We","Th","Fr","Sa"].map(d=><div key={d} style={{textAlign:"center",fontSize:12,fontWeight:700,color:SFT}}>{d}</div>)}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:4}}>
          {Array.from({length:firstDay},(_,i)=><div key={"b"+i}/>)}
          {Array.from({length:dim},(_,i)=>i+1).map(d=>{
            const isSel=selDates.includes(d);
            const wSt=selectedWorkers.length===1?(WORKER_ATT[selectedWorkers[0].empCode]||{})[d]:null;
            const sc=wSt?(STATUS_COLS[wSt]||{bg:BG,fg:INK}):{bg:BG,fg:INK};
            return(
              <div key={d}
                onClick={()=>toggleDate(d)} style={{padding:"5px 0",textAlign:"center",borderRadius:6,cursor:"pointer",background:isSel?CB:sc.bg,color:isSel?"#fff":sc.fg,border:isSel?"2px solid "+DARK:wSt?"1px solid "+LN:"1px solid transparent",fontSize:13,fontWeight:700}}>
                <div>{d}</div>
                {wSt&&!isSel&&<div style={{fontSize:11,fontWeight:800}}>{wSt}</div>}
              </div>
            );
          })}
        </div>
        {selDates.length>0&&<div style={{fontSize:14,color:CB,fontWeight:700,marginTop:8}}>{selDates.length} day(s) selected: {selDates.sort((a,b)=>a-b).join(", ")}</div>}
      </Card>
      <Card>
        <SHdr title="Step 3: Set Status"/>
        {sent&&<div
          style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:GRN,fontWeight:700}}>Override submitted for {selectedWorkers.length} worker(s).</div>}
        <label style={lbl}>Set status to</label>
        <select style={inp} value={newStatus} onChange={e=>setNewStatus(e.target.value)}>{STATUS_LIST.map(s=><option key={s}>{s}</option>)}</select>
        <label style={lbl}>Reason</label>
        <input style={inp} value={reason} onChange={e=>setReason(e.target.value)} placeholder="Reason for override..."/>
        <button style={{...PBtn,marginTop:10,opacity:(selDates.length===0||selectedIds.length===0)?0.5:1}}
          onClick={submitOverride} disabled={selDates.length===0||selectedIds.length===0}>Submit Override ({selectedIds.length} workers x {selDates.length} days)</button>
        {selectedIds.length===0&&<div style={{fontSize:13,color:SFT,marginTop:6,fontStyle:"italic"}}>Select at least one worker in Step 1 to submit.</div>}
        {pending.length>0&&(
          <div style={{marginTop:12,paddingTop:12,borderTop:"1px solid "+LN}}>
            <div style={{fontSize:14,fontWeight:800,color:CB,marginBottom:6}}>Pending Super Admin approval: {pending.length} items</div>
            {pending.map((p,i)=>(
              <div key={i} style={{fontSize:14,color:SFT,padding:"4px 0"}}>{p.worker} ({p.empCode}) - {p.status} - {p.month} - {p.dates.length} days</div>
            ))}
          </div>
        )}
        <div
          style={{background:ABG,border:"1px solid #F2C97E",borderRadius:8,padding:"8px 10px",marginTop:10,fontSize:14,color:AMB,fontWeight:600}}>Grade overrides are sent to Admin for final approval. Auto-calculated grades update daily.</div>
      </Card>
      <Card s={{marginBottom:10,border:"2px solid "+PUR}}>
        <SHdr title="Grade Override" sub="Manually set a worker's performance grade badge - sent to Admin for final approval"/>
        {goSent&&<div
          style={{background:GBG,border:"1px solid "+GRN,borderRadius:8,padding:"8px 10px",marginBottom:10,fontSize:14,color:GRN,fontWeight:700}}>Grade override applied and sent to Admin.</div>}
        <label style={lbl}>Search Worker</label>
        <input style={inp} value={goSearch}
          onChange={e=>{setGoSearch(e.target.value);setGoShowDrop(true);if(!e.target.value)setGoSelWorker(null);}} onFocus={()=>setGoShowDrop(true)} placeholder="Type name or emp code..."/>
        {goSearchRes.map((w,i)=>(
          <div key={w.id}
            onClick={()=>{setGoSelWorker(w);setGoSearch(w.name);setGoShowDrop(false);}} style={{padding:"6px 0",borderTop:i?"1px solid "+LN:"none",cursor:"pointer",fontSize:15,fontWeight:600}}>{w.name} <span style={{color:SFT,fontSize:13}}>{w.empCode}</span></div>
        ))}
        {goSelWorker&&(
          <>
            <div style={{background:PBG,borderRadius:8,padding:"7px 10px",marginTop:8,marginBottom:8,fontSize:15,color:PUR,fontWeight:600,display:"flex",alignItems:"center",gap:8}}>
              <GradeBadge grade={workerGrade(goSelWorker.empCode,gradeOverrides,goSelWorker.status)} size={16}/>
              {goSelWorker.name} - Current grade: {(gradeOverrides&&gradeOverrides[goSelWorker.empCode])||"auto-calculated"}
            </div>
            <label style={lbl}>Set Grade To</label>
            <div style={{display:"flex",gap:8,marginBottom:8}}>
              {["green","yellow","red"].map(g=>(
                <div key={g}
                  onClick={()=>setGoGrade(g)} style={{flex:1,padding:"10px 0",textAlign:"center",borderRadius:8,cursor:"pointer",border:"2px solid "+(goGrade===g?PUR:LN),background:goGrade===g?PBG:PNL,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                  <GradeBadge grade={g} size={18}/>
                  <span style={{fontSize:13,fontWeight:700,color:goGrade===g?PUR:SFT,textTransform:"capitalize"}}>{g}</span>
                </div>
              ))}
            </div>
            <label style={lbl}>Reason for Override</label>
            <input style={inp} value={goReason} onChange={e=>setGoReason(e.target.value)} placeholder="e.g. Client feedback, exceptional performance this month"/>
            <button style={{...PBtn,marginTop:10,background:PUR,opacity:!goReason?0.5:1}} onClick={submitGradeOverride} disabled={!goReason}>Apply Grade Override</button>
          </>
        )}
      </Card>
      <TempStatusPanel sup={sup}/>
    </>
  );
}


function SupervisorApp({sup,onLogout}){
  const [tab,setTab]=useState("report");
  const [actionTab,setActionTab]=useState("medical");
  const [workerTab,setWorkerTab]=useState("profile");
  const [reqTab,setReqTab]=useState("approvals");
  const [timesheetTab,setTimesheetTab]=useState("calendar");
  const [gradeOverrides,setGradeOverrides]=useState({});
  const [inactive,setInactive]=useState({emergencyLeave:0,annualLeave:0,absconder:0,refusal:0,resignation:0,training:0});
  const handleInactiveChange=(key,count)=>setInactive(prev=>({...prev,[key]:prev[key]+count}));
  const medPending=2,advPending=2;

  const TABS=[
    {key:"report",     label:"Report"},
    {key:"attendance", label:"Attendance"},
    {key:"actions",    label:"Actions"},
    {key:"timesheet",  label:"Timesheet"},
    {key:"worker",     label:"Worker"},
    {key:"requests",   label:"Approvals"},
  ];
  const ACTION_TABS=["medical","idle","advances","override"];
  const WORKER_TABS=["profile","salary","ppe","complaints"];
  const REQUEST_TABS=["approvals","myreqs"];

  return(
    <div style={{maxWidth:800,margin:"0 auto",minHeight:"100vh",background:BG}}>
      <div style={{background:"linear-gradient(135deg,"+DARK+","+CB+")",padding:"14px 16px",color:"#fff",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div>
          <div style={{fontSize:14,fontWeight:700,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",letterSpacing:.5}}>Supervisor Dashboard</div>
          <div style={{fontSize:20,fontWeight:900,marginTop:2}}>{sup.name}</div>
          <div style={{fontSize:14,color:"rgba(255,255,255,0.6)",marginTop:2}}>Clients: {sup.clients.join(", ")}</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{position:"relative",width:38,height:38,borderRadius:19,background:"rgba(255,255,255,0.18)",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer"}}>
            <svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/>
              <path d="M13.73 21a2 2 0 0 1-3.46 0"/>
            </svg>
            <div style={{position:"absolute",top:7,right:9,width:8,height:8,borderRadius:4,background:CY,border:"1.5px solid "+DARK}}/>
          </div>
          <button onClick={onLogout} title="Log out" style={{width:38,height:38,borderRadius:19,background:"rgba(255,255,255,0.18)",border:"none",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",padding:0}}>
            <svg width="19" height="19" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
              <polyline points="16 17 21 12 16 7"/>
              <line x1="21" y1="12" x2="9" y2="12"/>
            </svg>
          </button>
        </div>
      </div>

      <div style={{display:"flex",gap:2,padding:"10px 10px 0",overflowX:"auto",borderBottom:"2px solid "+LN,background:PNL}}>
        {TABS.map(t=>(
          <button key={t.key}
            onClick={()=>setTab(t.key)} style={{fontSize:15,fontWeight:700,padding:"9px 14px",background:"none",border:"none",borderBottom:tab===t.key?"3px solid "+CB:"3px solid transparent",color:tab===t.key?CB:SFT,cursor:"pointer",whiteSpace:"nowrap"}}>{t.label}</button>
        ))}
      </div>

      <div style={{padding:12}}>
        {tab==="report"&&<ReportTab sup={sup} inactive={inactive} medPending={medPending} advPending={advPending}/>}
        {tab==="attendance"&&<AttendanceTab sup={sup} gradeOverrides={gradeOverrides} inactive={inactive}/>}

        {tab==="actions"&&(
          <>
            <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
              {[["medical","Medical ("+medPending+")"],["idle","Idle"],["advances","Advances"],["override","Override"]].map(([k,l])=>(
                <button key={k}
                  onClick={()=>setActionTab(k)} style={{fontSize:14,fontWeight:700,padding:"5px 12px",borderRadius:20,cursor:"pointer",border:"2px solid "+(actionTab===k?CB:LN),background:actionTab===k?CB:PNL,color:actionTab===k?"#fff":SFT,whiteSpace:"nowrap"}}>{l}</button>
              ))}
            </div>
            {actionTab==="medical"  &&<MedicalTab sup={sup}/>}
            {actionTab==="idle"     &&<IdleTab sup={sup} goToOverride={()=>setActionTab("override")}/>}
            {actionTab==="advances" &&<AdvanceTab sup={sup}/>}
            {actionTab==="override" &&<OverrideTab sup={sup} onInactiveChange={handleInactiveChange} gradeOverrides={gradeOverrides} setGradeOverrides={setGradeOverrides}/>}
          </>
        )}

        {tab==="timesheet"&&(
          <>
            <SubmitTimesheetsPanel sup={sup}/>
            <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
              {[["calendar","Reconciliation & Approvals"],["sitecomp","Attendance vs Timesheet Comparison"],["clienttimesheet","Client Queries Report"]].map(([k,l])=>(
                <button key={k}
                  onClick={()=>setTimesheetTab(k)} style={{fontSize:14,fontWeight:700,padding:"5px 12px",borderRadius:20,cursor:"pointer",border:"2px solid "+(timesheetTab===k?CB:LN),background:timesheetTab===k?CB:PNL,color:timesheetTab===k?"#fff":SFT,whiteSpace:"nowrap"}}>{l}</button>
              ))}
            </div>
            {timesheetTab==="calendar"        &&<WorkerCalendarTab sup={sup} gradeOverrides={gradeOverrides}/>}
            {timesheetTab==="sitecomp"        &&<SiteComparisonReport sup={sup}/>}
            {timesheetTab==="clienttimesheet" &&<ClientTimesheetReport sup={sup}/>}
          </>
        )}

        {tab==="worker"&&(
          <>
            <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
              {[["profile","Profile"],["salary","Salary Calculation"],["ppe","PPE"],["complaints","Complaints"]].map(([k,l])=>(
                <button key={k}
                  onClick={()=>setWorkerTab(k)} style={{fontSize:14,fontWeight:700,padding:"5px 12px",borderRadius:20,cursor:"pointer",border:"2px solid "+(workerTab===k?CB:LN),background:workerTab===k?CB:PNL,color:workerTab===k?"#fff":SFT,whiteSpace:"nowrap"}}>{l}</button>
              ))}
            </div>
            {workerTab==="profile"     &&<WorkerProfileTab sup={sup} gradeOverrides={gradeOverrides}/>}
            {workerTab==="salary"      &&<WorkerSalaryTab sup={sup}/>}
            {workerTab==="ppe"         &&<WorkerPPETab sup={sup}/>}
            {workerTab==="complaints"  &&<ComplaintsTab sup={sup}/>}
          </>
        )}

        {tab==="requests"&&(
          <>
            <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
              {[["approvals","Approvals & Actions View"],["myreqs","Supervisor Requests"]].map(([k,l])=>(
                <button key={k}
                  onClick={()=>setReqTab(k)} style={{fontSize:14,fontWeight:700,padding:"5px 12px",borderRadius:20,cursor:"pointer",border:"2px solid "+(reqTab===k?CB:LN),background:reqTab===k?CB:PNL,color:reqTab===k?"#fff":SFT,whiteSpace:"nowrap"}}>{l}</button>
              ))}
            </div>
            {reqTab==="approvals" &&<ApprovalsTab sup={sup} medPending={medPending} advPending={advPending}/>}
            {reqTab==="myreqs"    &&<MyRequestsTab/>}
          </>
        )}
      </div>
    </div>
  );
}

function SiteComparisonReport({sup}){
  const [rClient,setRClient]=useState("");
  const [rSite,setRSite]=useState("");
  const [rMonth,setRMonth]=useState("Jun 2026");
  const [rReceived,setRReceived]=useState("");
  const [empSearch,setEmpSearch]=useState("");
  const [preview,setPreview]=useState(false);
  const today2=new Date();
  const dayNames=["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  const allW=sup.workers;
  const clients=[...new Set(allW.map(w=>w.client))];
  const sites=[...new Set(allW.filter(w=>!rClient||w.client===rClient).map(w=>w.site))];
  const filteredW=allW.filter(w=>
    (!rClient||w.client===rClient)&&
    (!rSite||w.site===rSite)&&
    (!empSearch||w.empCode.toLowerCase().includes(empSearch.toLowerCase())||w.name.toLowerCase().includes(empSearch.toLowerCase()))
  );
  const moIdx=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(rMonth.split(" ")[0]);
  const yr2=Number(rMonth.split(" ")[1]);
  const dim=moIdx>=0&&yr2?new Date(yr2,moIdx+1,0).getDate():31;

  const buildRows=()=>{
    const rows=[];
    filteredW.forEach(w=>{
      const wDays=WORKER_ATT[w.empCode]||{};
      const allDays=new Set(Object.keys(wDays).map(Number));
      [...allDays].sort((a,b)=>a-b).forEach(d=>{
        const wSt=wDays[d]||"";
        const dateStr=String(d).padStart(2,"0")+"/"+String(today2.getMonth()+1).padStart(2,"0")+"/"+today2.getFullYear();
        const dayName=dayNames[new Date(today2.getFullYear(),today2.getMonth(),d).getDay()];
        const hrs=["P"].includes(wSt)?8:["MWP","MWOP","I","PH","OT"].includes(wSt)?8:0;
        const otHrs=wSt==="P"?2:0;
        rows.push({empCode:w.empCode,worker:w.name,date:dateStr,dayName,dayNum:d,wSt,cSt:wSt,hrs,otHrs,match:true});
      });
    });
    return rows;
  };

  // Hours per status per spec: P=8+2OT, PH=8+0OT, MWP/MWOP/I/OT=8+0OT, all others=0+0
  const STATUS_HOURS={"P":{h:8,ot:2},"A":{h:0,ot:0},"WO":{h:0,ot:0},"PH":{h:8,ot:0},"REF":{h:0,ot:0},"MWP":{h:8,ot:0},"MWOP":{h:8,ot:0},"I":{h:8,ot:0},"EL":{h:0,ot:0},"AL":{h:0,ot:0},"OT":{h:8,ot:0},"AB":{h:0,ot:0},"RES":{h:0,ot:0},"NSF":{h:0,ot:0},"TS":{h:0,ot:0}};
  const getH=(st)=>STATUS_HOURS[st]||{h:0,ot:0};
  const downloadExcel=()=>{
    const wb=XLSX.utils.book_new();
    const clientName=rClient||"All Clients";
    const siteName=rSite||"All Sites";
    const days=Array.from({length:dim},(_,i)=>i+1);
    const SUMMARY=["P","A","I","M","WO","PH","WH"];
    const titleRows=[
      ["FIXIT MANPOWER SUPPLY L.L.C - Attendance vs Client Timesheet Comparison",...Array(2+dim+SUMMARY.length).fill("")],
      ["Client: "+clientName+" | Site: "+siteName+" | Month: "+rMonth+" | Sheet Received: "+(rReceived||"N/A")+" | Supervisor: "+sup.name,...Array(2+dim+SUMMARY.length).fill("")],
      Array(3+dim+SUMMARY.length).fill(""),
      ["Employee Name","Row Type",...days,...SUMMARY,"Std Hours","OT Hours","Total Hrs"],
    ];
    const dataRows=[];
    filteredW.forEach(w=>{
      const wDays=WORKER_ATT[w.empCode]||{};
      const wRow=[w.name,"Worker Attendance"];
      const cRow=["","Client Hours (Std+OT)"];
      const wCount={P:0,A:0,I:0,M:0,WO:0,PH:0,WH:0};
      let stdTotal=0,otTotal=0;
      days.forEach(d=>{
        const wSt=wDays[d]||"";
        wRow.push(wSt);
        const {h,ot}=getH(wSt);
        cRow.push(wSt?(h>0||ot>0?h+"h"+(ot>0?"+"+ot+"OT":""):"0h"):"");
        if(wSt==="P") wCount.P++;
        else if(wSt==="A") wCount.A++;
        else if(wSt==="I") wCount.I++;
        else if(wSt==="MWP"||wSt==="MWOP") wCount.M++;
        else if(wSt==="WO") wCount.WO++;
        else if(wSt==="PH") wCount.PH++;
        stdTotal+=h; otTotal+=ot;
      });
      SUMMARY.forEach(s=>{wRow.push(wCount[s]||0);cRow.push("");});
      wRow.push(stdTotal+"h"); wRow.push(otTotal+"h"); wRow.push((stdTotal+otTotal)+"h");
      cRow.push(stdTotal+"h"); cRow.push("(Actual OT)"); cRow.push((stdTotal+otTotal)+"h");
      dataRows.push(wRow);
      dataRows.push(cRow);
      dataRows.push(Array(4+dim+SUMMARY.length).fill(""));
    });
    const allRows=[...titleRows,...dataRows,
      Array(4+dim+SUMMARY.length).fill(""),
      ["Prepared by: "+sup.name,...Array(3+dim+SUMMARY.length).fill("")],
      ["Note: OT hours are assigned per actual client timesheet by supervisor",...Array(3+dim+SUMMARY.length).fill("")],
      ["For internal use only - Fixit Manpower Supply L.L.C",...Array(3+dim+SUMMARY.length).fill("")],
    ];
    const ws=XLSX.utils.aoa_to_sheet(allRows);
    ws["!cols"]=[{wch:18},{wch:20},...Array(dim).fill({wch:4.5}),...Array(SUMMARY.length).fill({wch:5}),{wch:9},{wch:9},{wch:9}];
    ws["!freeze"]={xSplit:2,ySplit:4};
    XLSX.utils.book_append_sheet(wb,ws,"Comparison");
    XLSX.writeFile(wb,"AttendanceComparison_"+clientName.split(" ").join("_")+"_"+siteName.split(" ").join("_")+"_"+rMonth.split(" ").join("_")+".xlsx");
  };

  const previewRows=buildRows().slice(0,15);

  return(
    <Card s={{marginBottom:12,border:"2px solid "+GRN}}>
      <SHdr title="Attendance vs Timesheet Comparison Report" sub="All workers - client sheet vs worker attendance - for site visits and dispute discussions"/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div>
          <label style={{...lbl,marginTop:0}}>Month</label>
          <select style={inp} value={rMonth} onChange={e=>setRMonth(e.target.value)}>
            {["Jun 2026","May 2026","Apr 2026","Mar 2026"].map(m=><option key={m}>{m}</option>)}
          </select>
        </div>
        <div>
          <label style={{...lbl,marginTop:0,color:CB}}>Client</label>
          <select style={inp} value={rClient} onChange={e=>{setRClient(e.target.value);setRSite("");}}>
            <option value="">All clients</option>
            {clients.map(c=><option key={c}>{c}</option>)}
          </select>
        </div>
        <div>
          <label style={{...lbl,marginTop:0,color:ORANGE}}>Site</label>
          <select style={inp} value={rSite} onChange={e=>setRSite(e.target.value)}>
            <option value="">All sites</option>
            {sites.map(s=><option key={s}>{s}</option>)}
          </select>
        </div>
        <div>
          <label style={{...lbl,marginTop:0,color:BLU}}>Client Sheet Received Date</label>
          <input type="date" style={inp} value={rReceived} onChange={e=>setRReceived(e.target.value)}/>
        </div>
      </div>
      <input style={{...inp,marginBottom:6}} value={empSearch} onChange={e=>setEmpSearch(e.target.value)} placeholder="Search by employee name or code (optional)..."/>
      <div style={{background:BG,borderRadius:8,padding:"7px 10px",marginBottom:10,fontSize:14,color:SFT}}>
        {filteredW.length} workers | {rClient||"All clients"} | {rSite||"All sites"} | {rMonth}{rReceived?" | Received: "+rReceived:""}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:preview?10:0}}>
        <button
          onClick={()=>setPreview(!preview)} style={{flex:1,padding:"9px",border:"2px solid "+CB,borderRadius:10,color:CB,background:preview?BBG:PNL,cursor:"pointer",fontWeight:700,fontSize:15}}>
          {preview?"Hide Preview":"Preview"}
        </button>
        <button onClick={downloadExcel} style={{...PBtn,margin:0,flex:2,background:"linear-gradient(135deg,"+DARK+","+CB+")",padding:"9px"}}>
          Download Comparison Excel
        </button>
      </div>
      {preview&&(
        <div style={{marginTop:10,overflowX:"auto",border:"1px solid "+LN,borderRadius:8}}>
          <table style={{borderCollapse:"collapse",fontSize:12,whiteSpace:"nowrap"}}>
            <thead>
              <tr style={{background:DARK}}>
                <th style={{color:"#fff",padding:"5px 8px",textAlign:"left",minWidth:120,position:"sticky",left:0,background:DARK,zIndex:2}}>Employee Name</th>
                <th style={{color:"#fff",padding:"5px 8px",textAlign:"left",minWidth:110,position:"sticky",left:120,background:DARK,zIndex:2}}>Row Type</th>
                {Array.from({length:dim},(_,i)=>(
                  <th key={i+1} style={{color:"#fff",padding:"5px 4px",textAlign:"center",minWidth:28}}>{i+1}</th>
                ))}
                {["P","A","I","M","WO","PH","WH"].map(s=>(
                  <th key={s} style={{color:"#fff",padding:"5px 4px",textAlign:"center",minWidth:28}}>{s}</th>
                ))}
                <th style={{color:"#fff",padding:"5px 6px",textAlign:"center",minWidth:52}}>Std Hrs</th>
                <th style={{color:"#fff",padding:"5px 6px",textAlign:"center",minWidth:48}}>OT Hrs</th>
                <th style={{color:"#fff",padding:"5px 6px",textAlign:"center",minWidth:52}}>Total</th>
              </tr>
            </thead>
            <tbody>
              {filteredW.map((w,wi)=>{
                const wDays2=WORKER_ATT[w.empCode]||{};
                const wCount={P:0,A:0,I:0,M:0,WO:0,PH:0,WH:0};
                let stdTotal=0,otTotal=0;
                Array.from({length:dim},(_,i)=>{
                  const st=wDays2[i+1]||"";
                  if(st==="P") wCount.P++;
                  else if(st==="A") wCount.A++;
                  else if(st==="I") wCount.I++;
                  else if(st==="MWP"||st==="MWOP") wCount.M++;
                  else if(st==="WO") wCount.WO++;
                  else if(st==="PH") wCount.PH++;
                  const hs={"P":{h:8,ot:2},"PH":{h:8,ot:0},"MWP":{h:8,ot:0},"MWOP":{h:8,ot:0},"I":{h:8,ot:0},"OT":{h:8,ot:0}}[st]||{h:0,ot:0};
                  stdTotal+=hs.h; otTotal+=hs.ot;
                });
                return([
                  <tr key={w.id+"w"} style={{background:wi%2?"#F0F4FF":PNL}}>
                    <td style={{padding:"4px 8px",fontWeight:700,color:DARK,position:"sticky",left:0,background:wi%2?"#F0F4FF":PNL,borderRight:"1px solid "+LN}}>{w.name}</td>
                    <td style={{padding:"4px 8px",fontSize:11,color:BLU,fontWeight:600,position:"sticky",left:120,background:wi%2?"#F0F4FF":PNL,borderRight:"1px solid "+LN}}>Worker Attendance</td>
                    {Array.from({length:dim},(_,i)=>{
                      const d=i+1;const st=wDays2[d]||"";
                      const sc=STATUS_COLS[st]||{bg:"#F5F5F5",fg:"#CCC"};
                      return(<td key={d} style={{padding:"3px 2px",textAlign:"center",background:st?sc.bg:"#F9F9F9"}}>
                        {st&&<span style={{fontSize:11,fontWeight:800,color:sc.fg}}>{st}</span>}
                      </td>);
                    })}
                    {["P","A","I","M","WO","PH","WH"].map(sk=>(
                      <td key={sk} style={{padding:"3px 4px",textAlign:"center",fontWeight:700,color:wCount[sk]>0?GRN:SFT,fontSize:12}}>{wCount[sk]||0}</td>
                    ))}
                    <td style={{padding:"3px 6px",textAlign:"center",color:GRN,fontWeight:700}}>{stdTotal}h</td>
                    <td style={{padding:"3px 6px",textAlign:"center",color:ORANGE,fontWeight:700}}>{otTotal}h</td>
                    <td style={{padding:"3px 6px",textAlign:"center",color:BLU,fontWeight:700}}>{stdTotal+otTotal}h</td>
                  </tr>,
                  <tr key={w.id+"c"} style={{background:wi%2?"#F0FFF4":"#F9FFF9"}}>
                    <td style={{padding:"4px 8px",position:"sticky",left:0,background:wi%2?"#F0FFF4":"#F9FFF9",borderRight:"1px solid "+LN}}></td>
                    <td style={{padding:"4px 8px",fontSize:11,color:GRN,fontWeight:600,position:"sticky",left:120,background:wi%2?"#F0FFF4":"#F9FFF9",borderRight:"1px solid "+LN}}>Client Hours</td>
                    {Array.from({length:dim},(_,i)=>{
                      const d=i+1;const st=wDays2[d]||"";
                      const hs={"P":{h:8,ot:2},"PH":{h:8,ot:0},"MWP":{h:8,ot:0},"MWOP":{h:8,ot:0},"I":{h:8,ot:0},"OT":{h:8,ot:0}}[st]||{h:0,ot:0};
                      const total=hs.h+hs.ot;
                      return(<td key={d} style={{padding:"3px 2px",textAlign:"center",background:st?(total>0?GBG:"#FFF5F5"):"#F9F9F9"}}>
                        {st&&<span style={{fontSize:11,fontWeight:700,color:total>0?GRN:RED}}>{total>0?total+"h":"0h"}</span>}
                      </td>);
                    })}
                    {["P","A","I","M","WO","PH","WH"].map(sk=>(
                      <td key={sk} style={{padding:"3px 4px",textAlign:"center",color:SFT,fontSize:11}}>-</td>
                    ))}
                    <td style={{padding:"3px 6px",textAlign:"center",color:GRN,fontWeight:700}}>{stdTotal}h</td>
                    <td style={{padding:"3px 6px",textAlign:"center",color:AMB,fontWeight:600,fontSize:11}}>Actual</td>
                    <td style={{padding:"3px 6px",textAlign:"center",color:BLU,fontWeight:700}}>{stdTotal}h+</td>
                  </tr>,
                  <tr key={w.id+"sp"}><td colSpan={3+dim+10} style={{height:4,background:"#FAFAFA"}}/></tr>
                ]);
              })}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

function ClientTimesheetReport({sup}){
  const [rClient,setRClient]=useState("");
  const [rSite,setRSite]=useState("");
  const [rMonth,setRMonth]=useState("Jun 2026");
  const [showPreview,setShowPreview]=useState(false);
  const allW=sup.workers;
  const clients=[...new Set(allW.map(w=>w.client))];
  const sites=[...new Set(allW.filter(w=>!rClient||w.client===rClient).map(w=>w.site))];
  const filteredW=allW.filter(w=>(!rClient||w.client===rClient)&&(!rSite||w.site===rSite));
  const moIdx=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"].indexOf(rMonth.split(" ")[0]);
  const yr=Number(rMonth.split(" ")[1]);
  const dim=moIdx>=0&&yr?new Date(yr,moIdx+1,0).getDate():30;
  const downloadExcel=()=>{
    const wb=XLSX.utils.book_new();
    const header=["Emp Code","Worker","Client","Site","Date","Status","Basic Hrs","OT Hrs","Total Hrs","Query / Remark"];
    const rows=[];
    filteredW.forEach(w=>{
      const watt=WORKER_ATT[w.empCode]||{};
      for(let d=1;d<=dim;d++){
        const st=watt[d];if(!st)continue;
        rows.push([w.empCode,w.name,w.client,w.site,String(d).padStart(2,"0")+"/"+rMonth,st,st==="P"?8:0,st==="P"?2:0,st==="P"?10:0,""]);
      }
    });
    const titleRows=[["FIXIT MANPOWER SUPPLY L.L.C",...Array(9).fill("")],["Client Queries Report - "+rMonth,...Array(9).fill("")],["Client: "+(rClient||"All")+" | Site: "+(rSite||"All"),...Array(9).fill("")],Array(10).fill("")];
    const ws=XLSX.utils.aoa_to_sheet([...titleRows,header,...rows]);
    ws["!cols"]=[{wch:10},{wch:18},{wch:20},{wch:16},{wch:10},{wch:8},{wch:10},{wch:8},{wch:10},{wch:24}];
    XLSX.utils.book_append_sheet(wb,ws,"Client Queries");
    XLSX.writeFile(wb,"ClientQueries_"+(rSite||rClient||"All").split(" ").join("_")+"_"+rMonth.split(" ").join("_")+".xlsx");
  };
  return(
    <Card s={{marginBottom:12}}>
      <SHdr title="Client Queries Report" sub="Search, preview and export client attendance data for reconciliation"/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
        <div>
          <label style={{...lbl,marginTop:0}}>Month</label>
          <select style={inp} value={rMonth} onChange={e=>setRMonth(e.target.value)}>
            {["Jun 2026","May 2026","Apr 2026","Mar 2026"].map(m=><option key={m}>{m}</option>)}
          </select>
        </div>
        <div>
          <label style={{...lbl,marginTop:0,color:CB}}>Client</label>
          <select style={inp} value={rClient} onChange={e=>{setRClient(e.target.value);setRSite("");}}>
            <option value="">All clients</option>
            {clients.map(c=><option key={c}>{c}</option>)}
          </select>
        </div>
        <div>
          <label style={{...lbl,marginTop:0,color:ORANGE}}>Site</label>
          <select style={inp} value={rSite} onChange={e=>setRSite(e.target.value)}>
            <option value="">All sites</option>
            {sites.map(s=><option key={s}>{s}</option>)}
          </select>
        </div>
      </div>
      <div style={{background:BG,borderRadius:8,padding:"7px 10px",marginBottom:10,fontSize:14,color:SFT}}>
        {filteredW.length} worker(s) | {rClient||"All clients"} | {rSite||"All sites"} | {rMonth}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:showPreview?10:0}}>
        <button
          onClick={()=>setShowPreview(!showPreview)} style={{flex:1,padding:"9px",border:"2px solid "+CB,borderRadius:10,color:showPreview?"#fff":CB,background:showPreview?CB:PNL,cursor:"pointer",fontWeight:700,fontSize:15}}>{showPreview?"Hide Preview":"Preview"}</button>
        <button onClick={downloadExcel} style={{...PBtn,margin:0,flex:1,background:"linear-gradient(135deg,"+DARK+","+CB+")",padding:"9px"}}>Export Excel</button>
      </div>
      {showPreview&&(
        <div style={{overflowX:"auto",border:"1px solid "+LN,borderRadius:8,marginTop:4}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:560}}>
            <thead>
              <tr style={{background:DARK}}>
                {["Worker","Code","Date","Status","Basic","OT","Total"].map(h=>(
                  <th key={h} style={{color:"#fff",padding:"5px 6px",textAlign:"left",fontSize:12}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredW.map((w,wi)=>{
                const watt=WORKER_ATT[w.empCode]||{};
                return Array.from({length:dim},(_,i)=>{
                  const d=i+1;const st=watt[d];if(!st)return null;
                  const sc=STATUS_COLS[st]||{bg:BG,fg:SFT};
                  const basic=st==="P"?8:0;const ot=st==="P"?2:0;
                  return(
                    <tr key={w.id+"-"+d} style={{background:(wi+d)%2?BG:PNL}}>
                      <td style={{padding:"3px 6px",fontWeight:600,fontSize:13}}>{d===1?w.name:""}</td>
                      <td style={{padding:"3px 6px",color:BLU,fontSize:12}}>{d===1?w.empCode:""}</td>
                      <td style={{padding:"3px 6px",color:SFT}}>{String(d).padStart(2,"0")}/{rMonth}</td>
                      <td style={{padding:"3px 6px"}}><span style={{background:sc.bg,color:sc.fg,padding:"1px 5px",borderRadius:4,fontWeight:700,fontSize:12}}>{st}</span></td>
                      <td style={{padding:"3px 6px",color:GRN,fontWeight:700}}>{basic}h</td>
                      <td style={{padding:"3px 6px",color:ORANGE,fontWeight:700}}>{ot}h</td>
                      <td style={{padding:"3px 6px",color:BLU,fontWeight:700}}>{basic+ot}h</td>
                    </tr>
                  );
                }).filter(Boolean);
              })}
            </tbody>
          </table>
          {filteredW.length===0&&<div style={{padding:"12px",textAlign:"center",color:SFT,fontSize:14}}>No workers found for selected filters.</div>}
        </div>
      )}
    </Card>
  );
}


export default function App(){
  const [loggedIn,setLoggedIn]=useState(null);
  const [pin,setPin]=useState("");
  const [error,setError]=useState("");
  const [busy,setBusy]=useState(false);
  const doLogin=async(p)=>{
    setBusy(true);setError("");
    try{
      const rows=await sbQuery("supervisors",
        "login_pin=eq."+encodeURIComponent(p)+"&active=eq.true&select=*");
      if(rows&&rows.length>0){
        const s=rows[0];
        // Load this supervisor's team (their workers)
        const rawWorkers=await sbQuery("workers",
          "supervisor_id=eq."+s.id+"&select=*&order=full_name");
        // Resolve client + site names
        const clientIds=[...new Set((rawWorkers||[])
          .map(w=>w.client_id).filter(Boolean))];
        const siteIds=[...new Set((rawWorkers||[])
          .map(w=>w.site_id).filter(Boolean))];
        let clientMap={},siteMap={},clientNames=[];
        if(clientIds.length>0){
          const cl=await sbQuery("clients",
            "id=in.("+clientIds.join(",")+")&select=id,name");
          (cl||[]).forEach(c=>{clientMap[c.id]=c.name;});
          clientNames=(cl||[]).map(c=>c.name);
        }
        if(siteIds.length>0){
          const st=await sbQuery("sites",
            "id=in.("+siteIds.join(",")+")&select=id,site_name");
          (st||[]).forEach(x=>{siteMap[x.id]=x.site_name;});
        }
        // Normalize DB fields -> the shape the app expects
        const workers=(rawWorkers||[]).map((w,i)=>({
          id:w.id,
          dbId:w.id,
          name:w.full_name,
          empCode:w.emp_code,
          passport:w.passport_no,
          client:clientMap[w.client_id]||"",
          site:siteMap[w.site_id]||"",
          trade:w.trade,
          stampDate:w.stamp_date,
          visaExpiry:w.visa_expiry,
          passportExpiry:w.passport_expiry,
          medicalInsExpiry:w.medical_ins_expiry,
          wcExpiry:w.wc_expiry,
          siteChanges:w.site_changes||0,
          status:w.status||"P",
          nationality:w.nationality,
          payType:w.pay_type,
          basic:w.basic||0,food:w.food||0,otherAllow:w.other_allow||0,
          rate:w.rate||0,
          basicJoin:w.basic_join||0,foodJoin:w.food_join||0,
          otherAllowJoin:w.other_allow_join||0,rateJoin:w.rate_join||0,
          otRate:w.ot_rate||0,phOtRate:w.ph_ot_rate||0,
          advancePending:w.advance_pending||0,
          weekOff:w.week_off,grade:w.grade,
        }));
        const sup={
          id:s.id,
          name:s.full_name,
          empCode:s.emp_code,
          phone:s.phone,
          companyId:s.company_id,
          clients:clientNames,
          workers:workers,
        };
        setLoggedIn(sup);setError("");
      }else{
        setError("Invalid PIN.");
        setTimeout(()=>setPin(""),400);
      }
    }catch(e){
      setError("Connection error. Try again.");
      setTimeout(()=>setPin(""),400);
    }
    setBusy(false);
  };
  const pressDigit=(d)=>{
    if(pin.length>=4||busy)return;
    const next=pin+d;
    setPin(next);
    if(next.length===4) doLogin(next);
  };
  const backspace=()=>setPin(pin.slice(0,-1));
  if(loggedIn) return <SupervisorApp sup={loggedIn} onLogout={()=>{setLoggedIn(null);setPin("");}}/>;
  return(
    <div style={{maxWidth:340,margin:"60px auto",padding:24,background:PNL,borderRadius:16,boxShadow:"0 4px 20px rgba(27,58,138,.12)"}}>
      <div style={{textAlign:"center",marginBottom:24}}>
        <div style={{fontSize:23,fontWeight:900,color:DARK}}>FIXIT MANPOWER</div>
        <div style={{fontSize:15,color:SFT,marginTop:2}}>{busy?"Signing in...":"Supervisor Login - Enter PIN"}</div>
      </div>
      <div style={{display:"flex",justifyContent:"center",gap:10,marginBottom:20}}>
        {[0,1,2,3].map(i=>(
          <div key={i}
            style={{width:44,height:52,borderRadius:10,border:"2px solid "+(pin.length>i?CB:LN),background:pin.length>i?BBG:BG,display:"flex",alignItems:"center",justifyContent:"center",fontSize:25,fontWeight:900,color:CB}}>
            {pin.length>i?"*":""}
          </div>
        ))}
      </div>
      {error&&<div style={{color:RED,fontSize:15,marginBottom:12,fontWeight:600,textAlign:"center"}}>{error}</div>}
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
        {[1,2,3,4,5,6,7,8,9].map(d=>(
          <button key={d}
            onClick={()=>pressDigit(String(d))} style={{padding:"16px 0",fontSize:21,fontWeight:800,color:INK,background:BG,border:"1px solid "+LN,borderRadius:10,cursor:"pointer"}}>{d}</button>
        ))}
        <button onClick={()=>setPin("")} style={{padding:"16px 0",fontSize:15,fontWeight:700,color:SFT,background:BG,border:"1px solid "+LN,borderRadius:10,cursor:"pointer"}}>Clear</button>
        <button onClick={()=>pressDigit("0")} style={{padding:"16px 0",fontSize:21,fontWeight:800,color:INK,background:BG,border:"1px solid "+LN,borderRadius:10,cursor:"pointer"}}>0</button>
        <button onClick={backspace} style={{padding:"16px 0",fontSize:19,fontWeight:700,color:SFT,background:BG,border:"1px solid "+LN,borderRadius:10,cursor:"pointer"}}>&lt;-</button>
      </div>
      <div style={{fontSize:14,color:SFT,marginTop:16,textAlign:"center"}}>Demo PINs: 1042 (Ahmed) / 1043 (Vijay)</div>
    </div>
  );
}