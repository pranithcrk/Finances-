import { useEffect, useRef, useState, useCallback } from 'react';
import Head from 'next/head';

// ── defaults (used only on first launch before Excel is seeded) ──
const DEF_BUDGET = {
  committed: [
    {emoji:'🏠',name:'Rent',amount:1007,recurring:true,locked:false},
    {emoji:'⚡',name:'Utilities',amount:89,recurring:true,locked:true,feed:'utils'},
    {emoji:'📱',name:'Subscriptions',amount:76,recurring:true,locked:true,feed:'subs'},
    {emoji:'🛡',name:'Insurance',amount:80,recurring:true,locked:false},
    {emoji:'🚲',name:'Cycle',amount:263,recurring:true,locked:false},
  ],
  living: [
    {emoji:'🛒',name:'Groceries',amount:500,recurring:true,cat:'Groceries'},
    {emoji:'🍱',name:'Lunches',amount:120,recurring:true,cat:'Lunches'},
    {emoji:'🍜',name:'Eating Out',amount:150,recurring:false,cat:'Eating Out'},
    {emoji:'🚌',name:'Transport',amount:60,recurring:true,cat:'Transport'},
  ],
  save: [
    {emoji:'💵',name:'Savings (incl. India remit)',amount:1100,recurring:true},
    {emoji:'🏦',name:'Roth 401K',amount:415,recurring:true},
    {emoji:'📋',name:'Voluntary 401K',amount:208,recurring:true},
    {emoji:'🌧',name:'Rainy Day Fund',amount:620,recurring:true},
    {emoji:'📈',name:'ETFs (existing)',amount:250,recurring:true},
    {emoji:'⚡',name:'QQQ + VWO',amount:300,recurring:true},
    {emoji:'📊',name:'Individual Stocks',amount:200,recurring:true},
    {emoji:'🇮🇳',name:'Extra India Prepay',amount:400,recurring:true},
  ]
};
const DEF_SUBS = [
  {name:'YouTube Premium',amount:21,cycle:'monthly',ess:'essential'},
  {name:'Amazon Prime',amount:16,cycle:'monthly',ess:'essential'},
  {name:'Mobile Plan',amount:25,cycle:'monthly',ess:'essential'},
  {name:'iCloud+',amount:2.99,cycle:'monthly',ess:'essential'},
  {name:'FitHub Cal Track',amount:2.67,cycle:'monthly',ess:'essential'},
  {name:'Nibble',amount:8.13,cycle:'monthly',ess:'non'},
];
const DEF_UTILS = [
  {name:'Electricity',amount:18.86,split:'Rohan'},
  {name:'Gas',amount:19.33,split:'Rohan'},
  {name:'Internet',amount:50,split:''},
];
const DEF_CARDS = [{name:'Default Card',issuer:'•••• 0000',color:'#1a1916'}];
const DEF_TRIPS = [{name:'Tetons + Yellowstone',when:'Aug 2026',budget:2200,saved:400,expenses:[]}];
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const PAL = ['#1a1916','#2d6a4f','#1d3557','#b08968','#c1440e','#5a189a','#9a9690','#6a040f'];
const CARD_COLORS = ['#1a1916','#2d6a4f','#1d3557','#6a040f','#5a189a'];

const fmt = n => (n<0?'-$':'$')+Math.abs(parseFloat(n)||0).toLocaleString('en-US',{maximumFractionDigits:0});
const fmtd = n => (n<0?'-$':'$')+Math.abs(parseFloat(n)||0).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
const sumA = a => a.reduce((s,i)=>s+(parseFloat(i.amount)||0),0);
const esc = s => String(s||'').replace(/</g,'&lt;').replace(/"/g,'&quot;');
const shade = (hex,pct) => { const n=parseInt(hex.slice(1),16)||0; const r=Math.max(0,Math.min(255,(n>>16)+pct)); const g=Math.max(0,Math.min(255,(n>>8&0xff)+pct)); const b=Math.max(0,Math.min(255,(n&0xff)+pct)); return '#'+(r<<16|g<<8|b).toString(16).padStart(6,'0'); };

export default function App() {
  const [page, setPage] = useState('dashboard');
  const [curMonth, setCurMonth] = useState(new Date().getMonth());
  const [curYear, setCurYear] = useState(new Date().getFullYear());
  const [income, setIncome] = useState(5222.38);
  const [loan, setLoan] = useState(38550);
  const [budget, setBudget] = useState(DEF_BUDGET);
  const [subs, setSubs] = useState(DEF_SUBS);
  const [utils, setUtils] = useState(DEF_UTILS);
  const [expenses, setExpenses] = useState({});
  const [flex, setFlex] = useState({});
  const [trips, setTrips] = useState(DEF_TRIPS);
  const [goals, setGoals] = useState([]);
  const [cards, setCards] = useState(DEF_CARDS);
  const [notes, setNotes] = useState({});
  const [toast, setToast] = useState('');
  const [toastVisible, setToastVisible] = useState(false);
  const [saveStatus, setSaveStatus] = useState('saved'); // 'saved' | 'saving' | 'error'
  const [loaded, setLoaded] = useState(false);
  const [cardColor, setCardColor] = useState('#1a1916');
  const [showCardForm, setShowCardForm] = useState(false);
  const [showTripForm, setShowTripForm] = useState(false);
  const [aiOutput, setAiOutput] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [retVal, setRetVal] = useState(10);
  const [contribVal, setContribVal] = useState(1873);
  const saveTimer = useRef(null);

  const monthKey = useCallback((m=curMonth,y=curYear) => `${y}-${String(m+1).padStart(2,'0')}`, [curMonth,curYear]);

  // ── Load from Excel on mount ──
  useEffect(() => {
    fetch('/api/budget')
      .then(r => r.json())
      .then(({ data, firstRun }) => {
        if (firstRun || !data) {
          // seed Excel with defaults immediately
          persistData({ income, loan, budget, subs, utils, expenses, flex, trips, goals, cards });
        } else {
          if (data.income) setIncome(data.income);
          if (data.loan) setLoan(data.loan);
          if (data.budget?.committed?.length || data.budget?.save?.length) setBudget(data.budget);
          if (data.subs?.length) setSubs(data.subs);
          if (data.utils?.length) setUtils(data.utils);
          if (data.expenses) setExpenses(data.expenses);
          if (data.flex) setFlex(data.flex);
          if (data.trips?.length) setTrips(data.trips);
          if (data.goals?.length) setGoals(data.goals);
          if (data.cards?.length) setCards(data.cards);
        }
        setLoaded(true);
      })
      .catch(() => setLoaded(true));
  }, []);

  // ── Auto-save to Excel (debounced 1.5s) ──
  const persistData = useCallback((overrides={}) => {
    const payload = {
      income: overrides.income ?? income,
      loan: overrides.loan ?? loan,
      budget: overrides.budget ?? budget,
      subs: overrides.subs ?? subs,
      utils: overrides.utils ?? utils,
      expenses: overrides.expenses ?? expenses,
      flex: overrides.flex ?? flex,
      trips: overrides.trips ?? trips,
      goals: overrides.goals ?? goals,
      cards: overrides.cards ?? cards,
    };
    setSaveStatus('saving');
    fetch('/api/budget', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: payload }),
    })
      .then(r => r.json())
      .then(() => setSaveStatus('saved'))
      .catch(() => setSaveStatus('error'));
  }, [income, loan, budget, subs, utils, expenses, flex, trips, goals, cards]);

  const scheduleSave = useCallback((overrides={}) => {
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => persistData(overrides), 1500);
  }, [persistData]);

  const showToast = msg => { setToast(msg); setToastVisible(true); setTimeout(()=>setToastVisible(false),1800); };

  // helpers that update state + schedule save
  const setIncomeSave = v => { setIncome(v); scheduleSave({income:v}); };
  const setLoanSave = v => { setLoan(v); scheduleSave({loan:v}); };
  const setBudgetSave = b => { setBudget(b); scheduleSave({budget:b}); };
  const setSubsSave = s => { setSubs(s); scheduleSave({subs:s}); };
  const setUtilsSave = u => { setUtils(u); scheduleSave({utils:u}); };
  const setExpensesSave = e => { setExpenses(e); scheduleSave({expenses:e}); };
  const setFlexSave = f => { setFlex(f); scheduleSave({flex:f}); };
  const setTripsSave = t => { setTrips(t); scheduleSave({trips:t}); };
  const setGoalsSave = g => { setGoals(g); scheduleSave({goals:g}); };
  const setCardsSave = c => { setCards(c); scheduleSave({cards:c}); };

  // ── feeds ──
  const subsMonthlyTotal = useCallback(() => subs.reduce((s,x)=>s+(x.cycle==='annual'?x.amount/12:x.amount),0),[subs]);
  const utilsMonthlyTotal = useCallback(() => utils.reduce((s,x)=>s+x.amount,0),[utils]);

  const syncedBudget = useCallback(() => {
    const b = JSON.parse(JSON.stringify(budget));
    b.committed.forEach(r => {
      if (r.feed==='subs') r.amount = subsMonthlyTotal();
      if (r.feed==='utils') r.amount = utilsMonthlyTotal();
    });
    return b;
  }, [budget, subsMonthlyTotal, utilsMonthlyTotal]);

  const livingActualByCat = useCallback((mk=monthKey()) => {
    const out = {};
    (expenses[mk]||[]).forEach(i => out[i.cat]=(out[i.cat]||0)+i.amount);
    return out;
  }, [expenses, monthKey]);

  const flexLoggedTotal = useCallback((mk=monthKey()) => (flex[mk]||[]).reduce((s,i)=>s+i.amount,0), [flex, monthKey]);

  const travelSpentThisMonth = useCallback((mk=monthKey()) => {
    let tot=0; trips.forEach(t=>t.expenses.forEach(e=>{if(e.month===mk)tot+=e.amt;})); return tot;
  }, [trips, monthKey]);

  const monthSnapshot = useCallback((mk) => {
    const b = syncedBudget();
    const committed = sumA(b.committed);
    const expItems = (expenses[mk]||[]);
    const livingActual = expItems.reduce((s,i)=>s+i.amount,0);
    const livingBudget = sumA(b.living);
    const livingUsed = livingActual>0?livingActual:livingBudget;
    const tFixed = committed+livingUsed;
    const tSave = sumA(b.save);
    const flexItems = (flex[mk]||[]);
    const flexLog = flexItems.reduce((s,i)=>s+i.amount,0);
    let travel=0; trips.forEach(t=>t.expenses.forEach(e=>{if(e.month===mk)travel+=e.amt;}));
    const flexSpent=flexLog+travel;
    const leftover=income-tFixed-tSave-flexSpent;
    return {income,committed,livingActual,livingBudget,livingUsed,tFixed,tSave,flexLog,travel,flexSpent,leftover,expItems,flexItems};
  }, [syncedBudget, expenses, flex, trips, income]);

  if (!loaded) return <div style={{display:'flex',alignItems:'center',justifyContent:'center',height:'100vh',fontFamily:'Geist Mono, monospace',fontSize:'13px',color:'#9a9690'}}>Loading fin. ···</div>;

  // ── render helpers ──
  const mk = monthKey();
  const s = monthSnapshot(mk);
  const actualCat = livingActualByCat(mk);

  return (
    <>
      <Head>
        <title>fin.</title>
        <link href="https://fonts.googleapis.com/css2?family=Geist+Mono:wght@300;400;500&family=Instrument+Serif:ital@0;1&display=swap" rel="stylesheet" />
      </Head>
      <style>{CSS}</style>

      <nav>
        <div className="nav-brand">fin.</div>
        <div className="nav-tabs">
          {['dashboard','subs','utils','invest','travel','tracker','cards','charts','suggest','reports'].map(p=>(
            <button key={p} className={`tab-btn${page===p?' active':''}`} onClick={()=>setPage(p)}>
              {{'dashboard':'Overview','subs':'Subscriptions','utils':'Utilities','invest':'Savings & Invest','travel':'Travel','tracker':'Expenses','cards':'Cards','charts':'Charts','suggest':'Suggestions','reports':'Reports'}[p]}
            </button>
          ))}
        </div>
        <div className="save-indicator" title="Excel sync">
          {saveStatus==='saving'&&<span style={{color:'#9a9690'}}>⟳ saving…</span>}
          {saveStatus==='saved'&&<span style={{color:'#2d6a4f'}}>✓ saved</span>}
          {saveStatus==='error'&&<span style={{color:'#c1440e'}}>⚠ sync error</span>}
        </div>
      </nav>

      {page==='dashboard'&&<Dashboard {...{s,budget:syncedBudget(),income,actualCat,flex,trips,curMonth,curYear,setCurMonth,setCurYear,monthKey,fmt,fmtd,setBudgetSave,setFlexSave,setIncomeSave,showToast,mk}} />}
      {page==='subs'&&<Subs {...{subs,subsMonthlyTotal,setSubsSave,showToast}} />}
      {page==='utils'&&<Utils {...{utils,utilsMonthlyTotal,setUtilsSave,showToast}} />}
      {page==='invest'&&<Invest {...{budget:syncedBudget(),income,loan,goals,retVal,contribVal,setRetVal,setContribVal,setGoalsSave,setLoanSave,subs,showToast,fmt,fmtd}} />}
      {page==='travel'&&<Travel {...{trips,showTripForm,setShowTripForm,setTripsSave,showToast,monthKey:mk,fmt,fmtd}} />}
      {page==='tracker'&&<Tracker {...{expenses,cards,budget:syncedBudget(),curMonth,curYear,setCurMonth,setCurYear,monthKey,setExpensesSave,notes,setNotes,showToast,fmt,fmtd}} />}
      {page==='cards'&&<Cards {...{cards,expenses,curMonth,curYear,setCurMonth,setCurYear,monthKey,showCardForm,setShowCardForm,cardColor,setCardColor,setCardsSave,showToast,fmt,fmtd}} />}
      {page==='charts'&&<Charts {...{monthSnapshot,curMonth,curYear,setCurMonth,setCurYear,monthKey,budget:syncedBudget(),cards,income,trips,flex,expenses,fmt,fmtd}} />}
      {page==='suggest'&&<Suggest {...{monthSnapshot,mk,budget:syncedBudget(),subs,cards,expenses,flex,income,trips,goals,aiOutput,setAiOutput,aiLoading,setAiLoading,curMonth,curYear,fmt,fmtd}} />}
      {page==='reports'&&<Reports {...{monthSnapshot,curMonth,curYear,budget:syncedBudget(),expenses,flex,trips,fmt,fmtd}} />}

      <div className={`toast${toastVisible?' show':''}`}>{toast}</div>
    </>
  );
}

// ═══════════════ DASHBOARD ═══════════════
function Dashboard({s,budget,income,actualCat,flex,trips,curMonth,curYear,setCurMonth,setCurYear,monthKey,fmt,fmtd,setBudgetSave,setFlexSave,setIncomeSave,showToast,mk}) {
  const [incomeEdit, setIncomeEdit] = useState(false);
  const [incomeInput, setIncomeInput] = useState(String(income));

  const changeMonth = d => {
    let m=curMonth+d, y=curYear;
    if(m>11){m=0;y++;}if(m<0){m=11;y--;}
    setCurMonth(m);setCurYear(y);
  };

  const editBudget = (sec,idx,field,raw) => {
    const b = JSON.parse(JSON.stringify(budget));
    if(field==='amount'){const n=parseFloat(raw.replace(/[$,]/g,''));if(isNaN(n)||n<0)return;b[sec][idx].amount=n;}
    else b[sec][idx][field]=raw.trim()||b[sec][idx][field];
    setBudgetSave(b);
  };
  const toggleRecur = (sec,idx) => {
    const b=JSON.parse(JSON.stringify(budget));b[sec][idx].recurring=!b[sec][idx].recurring;setBudgetSave(b);
  };
  const delBudget = (sec,idx) => {
    const b=JSON.parse(JSON.stringify(budget));b[sec].splice(idx,1);setBudgetSave(b);showToast('Removed');
  };
  const addBudgetLine = sec => {
    const e={committed:'📌',living:'🍽',save:'💰'}[sec];
    const b=JSON.parse(JSON.stringify(budget));
    const item={emoji:e,name:'New item',amount:0,recurring:true};
    if(sec==='living')item.cat='Misc';
    b[sec].push(item);setBudgetSave(b);showToast('Added');
  };

  const addFlex = () => {
    const desc=document.getElementById('flx-desc').value.trim();
    const amount=parseFloat(document.getElementById('flx-amount').value);
    const cat=document.getElementById('flx-cat').value;
    if(!desc||isNaN(amount)||amount<=0){showToast('Add item + amount');return;}
    const all=JSON.parse(JSON.stringify(flex));
    if(!all[mk])all[mk]=[];
    const t=new Date();all[mk].push({desc,amount,cat,date:`${t.getMonth()+1}/${t.getDate()}`});
    setFlexSave(all);
    document.getElementById('flx-desc').value='';document.getElementById('flx-amount').value='';
    showToast('Logged');
  };
  const delFlex = i => {
    const all=JSON.parse(JSON.stringify(flex));all[mk].splice(i,1);setFlexSave(all);showToast('Removed');
  };

  const flexItems = flex[mk]||[];
  const flexSpent = s.flexSpent;

  const segs = [
    {label:'Fixed',v:s.tFixed,c:'var(--fixed)'},
    {label:'Savings & Inv',v:s.tSave,c:'var(--save)'},
    {label:'Flex spent',v:flexSpent,c:'var(--flex)'},
    {label:'Free',v:Math.max(0,s.leftover),c:'var(--border2)'},
  ];

  return (
    <div className="page">
      <div className="container">
        <div className="ph ph-flex">
          <div><div className="pt">Where the money goes</div><div className="psub">Fixed + living + savings first — what's left is yours</div></div>
          <div className="month-nav">
            <button className="month-btn" onClick={()=>changeMonth(-1)}>←</button>
            <div className="month-label">{MONTHS[curMonth]} {curYear}</div>
            <button className="month-btn" onClick={()=>changeMonth(1)}>→</button>
          </div>
        </div>

        {/* INCOME */}
        <div className="income-hero">
          <div>
            <div className="income-label">Monthly Take-Home</div>
            {incomeEdit
              ? <input className="income-edit-input" autoFocus value={incomeInput} onChange={e=>setIncomeInput(e.target.value)}
                  onBlur={()=>{const n=parseFloat(incomeInput.replace(/[$,]/g,''));if(!isNaN(n)&&n>0){setIncomeSave(n);}setIncomeEdit(false);showToast('Income updated');}}
                  onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} />
              : <div className="income-val" onClick={()=>{setIncomeInput(String(income));setIncomeEdit(true);}}>{fmt(income)}</div>
            }
          </div>
          <div className="income-breakdown">
            <div className="ib-item"><div className="ib-label">Fixed</div><div className="ib-val neg">{fmt(s.tFixed)}</div></div>
            <div className="ib-item"><div className="ib-label">Savings/Inv</div><div className="ib-val neg">{fmt(s.tSave)}</div></div>
            <div className="ib-item"><div className="ib-label">Committed Out</div><div className="ib-val neg">{fmt(s.tFixed+s.tSave)}</div></div>
          </div>
        </div>

        {/* FLEX HERO */}
        <div className="flex-hero">
          <div className="fh-left">
            <div className="fh-label">Flex — free to spend or save</div>
            <div className={`fh-val${s.leftover<0?' neg':''}`}>{fmt(s.leftover)}</div>
            <div className="fh-sub">{s.leftover>=0?`You can spend up to ${fmt(s.leftover)} more this month — or move it to savings.`:`You're ${fmt(-s.leftover)} over your free budget this month.`}</div>
          </div>
          <div className="fh-right">
            <div className="fh-stat"><div className="fh-stat-label">Flex pool</div><div className="fh-stat-val">{fmt(income-s.tFixed-s.tSave)}</div></div>
            <div className="fh-stat"><div className="fh-stat-label">Spent (flex+travel)</div><div className="fh-stat-val">{fmt(flexSpent)}</div></div>
          </div>
        </div>

        {/* alloc bar */}
        <div className="alloc-bar-wrap">
          <div className="alloc-track">
            {segs.map(seg=><div key={seg.label} className="alloc-seg" style={{width:Math.max(0,(seg.v/income)*100)+'%',background:seg.c}} />)}
          </div>
          <div className="alloc-legend">
            {segs.map(seg=><div key={seg.label} className="al-item"><span className="al-dot" style={{background:seg.c}} /><span className="al-text"><b>{seg.label}</b> {fmt(seg.v)} · {Math.round((seg.v/income)*100)}%</span></div>)}
          </div>
        </div>

        {/* COMMITTED */}
        <div className="section">
          <div className="section-head">
            <div className="section-head-left"><span className="section-dot" style={{background:'var(--fixed)'}} /><span className="section-label">Fixed Outflows</span></div>
            <div className="section-total">{fmt(s.tFixed)}</div>
          </div>
          <div className="subgroup-label">Committed</div>
          <div className="col-headers" style={{gridTemplateColumns:'1fr 130px 120px 36px'}}>
            <div className="col-h">Item</div><div className="col-h">Amount</div><div className="col-h">Frequency</div><div className="col-h" />
          </div>
          <div className="row-list">
            {budget.committed.map((it,idx)=>(
              <div key={idx} className={`budget-row${it.locked?' locked-row':''}`} style={{gridTemplateColumns:'1fr 130px 120px 36px'}}>
                <div className="row-name">
                  <span className="row-emoji">{it.emoji}</span>
                  <input className="row-name-text" defaultValue={it.name} disabled={it.locked}
                    onBlur={e=>editBudget('committed',idx,'name',e.target.value)}
                    onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} />
                </div>
                <div className="row-cell right">
                  <input className="editable-val" defaultValue={fmtd(it.amount)} disabled={it.locked}
                    onBlur={e=>editBudget('committed',idx,'amount',e.target.value)}
                    onFocus={e=>e.target.select()} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} />
                </div>
                <div className="row-cell right">
                  {it.locked
                    ? <span className="locked-tag">auto · {it.feed}</span>
                    : <span className={`recur-toggle${it.recurring?' recurring':''}`} onClick={()=>toggleRecur('committed',idx)}>{it.recurring?'Recurring':'One-time'}</span>}
                </div>
                <div className="row-cell" style={{textAlign:'center'}}>
                  {!it.locked&&<button className="del-row-btn" onClick={()=>delBudget('committed',idx)}>×</button>}
                </div>
              </div>
            ))}
          </div>
          <button className="add-line-btn" onClick={()=>addBudgetLine('committed')}>+ Add committed expense</button>

          <div className="subgroup-label">Living Expenses · budget vs actual</div>
          <div className="col-headers" style={{gridTemplateColumns:'1fr 110px 110px 90px 36px'}}>
            <div className="col-h">Item</div><div className="col-h">Budget</div><div className="col-h">Actual</div><div className="col-h">Freq</div><div className="col-h" />
          </div>
          <div className="row-list">
            {budget.living.map((it,idx)=>{
              const actual=actualCat[it.cat]||0;
              const cls=actual>it.amount?'over':(actual>0?'under':'');
              return (
                <div key={idx} className="budget-row" style={{gridTemplateColumns:'1fr 110px 110px 90px 36px'}}>
                  <div className="row-name">
                    <span className="row-emoji">{it.emoji}</span>
                    <input className="row-name-text" defaultValue={it.name}
                      onBlur={e=>editBudget('living',idx,'name',e.target.value)}
                      onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} />
                  </div>
                  <div className="row-cell right">
                    <input className="editable-val" defaultValue={fmtd(it.amount)}
                      onBlur={e=>editBudget('living',idx,'amount',e.target.value)}
                      onFocus={e=>e.target.select()} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} />
                  </div>
                  <div className={`actual-val ${cls}`}>{fmtd(actual)}</div>
                  <div className="row-cell right">
                    <span className={`recur-toggle${it.recurring?' recurring':''}`} onClick={()=>toggleRecur('living',idx)}>{it.recurring?'Rec':'1×'}</span>
                  </div>
                  <div className="row-cell" style={{textAlign:'center'}}>
                    <button className="del-row-btn" onClick={()=>delBudget('living',idx)}>×</button>
                  </div>
                </div>
              );
            })}
          </div>
          <div className="feed-note">↳ "Actual" pulls from matching categories you log on the Expenses tab this month.</div>
          <button className="add-line-btn" onClick={()=>addBudgetLine('living')}>+ Add living expense</button>
        </div>

        {/* SAVINGS */}
        <div className="section">
          <div className="section-head">
            <div className="section-head-left"><span className="section-dot" style={{background:'var(--save)'}} /><span className="section-label">Savings & Investments</span></div>
            <div className="section-total">{fmt(s.tSave)}</div>
          </div>
          <div className="col-headers" style={{gridTemplateColumns:'1fr 130px 120px 36px'}}>
            <div className="col-h">Item</div><div className="col-h">Amount</div><div className="col-h">Frequency</div><div className="col-h" />
          </div>
          <div className="row-list">
            {budget.save.map((it,idx)=>(
              <div key={idx} className="budget-row" style={{gridTemplateColumns:'1fr 130px 120px 36px'}}>
                <div className="row-name">
                  <span className="row-emoji">{it.emoji}</span>
                  <input className="row-name-text" defaultValue={it.name}
                    onBlur={e=>editBudget('save',idx,'name',e.target.value)}
                    onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} />
                </div>
                <div className="row-cell right">
                  <input className="editable-val" defaultValue={fmtd(it.amount)}
                    onBlur={e=>editBudget('save',idx,'amount',e.target.value)}
                    onFocus={e=>e.target.select()} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} />
                </div>
                <div className="row-cell right">
                  <span className={`recur-toggle${it.recurring?' recurring':''}`} onClick={()=>toggleRecur('save',idx)}>{it.recurring?'Recurring':'One-time'}</span>
                </div>
                <div className="row-cell" style={{textAlign:'center'}}>
                  <button className="del-row-btn" onClick={()=>delBudget('save',idx)}>×</button>
                </div>
              </div>
            ))}
          </div>
          <div className="feed-note">↳ India remittance is drawn FROM the Savings line — not double-counted.</div>
          <button className="add-line-btn" onClick={()=>addBudgetLine('save')}>+ Add savings / investment</button>
        </div>

        {/* FLEX LOGGER */}
        <div className="section">
          <div className="section-head">
            <div className="section-head-left"><span className="section-dot" style={{background:'var(--flex)'}} /><span className="section-label">Flex Spending — log buys here</span></div>
            <div className="section-total">{fmt(flexSpent)}</div>
          </div>
          <div className="flex-logger">
            <div className="fl-form">
              <div className="form-group"><div className="form-label">What did you buy?</div><input className="form-input" id="flx-desc" placeholder="e.g. New headphones" onKeyDown={e=>{if(e.key==='Enter')addFlex();}} /></div>
              <div className="form-group"><div className="form-label">Amount $</div><input className="form-input" id="flx-amount" type="number" step="0.01" placeholder="0.00" onKeyDown={e=>{if(e.key==='Enter')addFlex();}} /></div>
              <div className="form-group"><div className="form-label">Type</div><select className="form-input form-select" id="flx-cat"><option>Shopping</option><option>Gadgets</option><option>Experience</option><option>Gift</option><option>Health</option><option>Other</option></select></div>
              <button className="add-btn" onClick={addFlex}>+ Log</button>
            </div>
            <div className="fl-list">
              {flexItems.length===0&&trips.every(t=>t.expenses.every(e=>e.month!==mk))
                ? <div className="fl-empty">No flex purchases logged this month.</div>
                : flexItems.map((it,idx)=>({it,idx})).reverse().map(({it,idx})=>(
                    <div key={idx} className="fl-item">
                      <span>{it.desc} <span className="chip">{it.cat}</span> <span className="exp-note">{it.date}</span></span>
                      <span>{fmtd(it.amount)} <button className="del-btn" onClick={()=>delFlex(idx)}>×</button></span>
                    </div>
                  ))
              }
              {trips.flatMap(t=>t.expenses.filter(e=>e.month===mk).map((e,i)=>(
                <div key={'t'+i} className="fl-item fl-travel">
                  <span>✈ {t.name}: {e.desc} <span className="chip">Travel</span></span>
                  <span>{fmtd(e.amt)}</span>
                </div>
              )))}
            </div>
          </div>
          <div className="feed-note">↳ Travel spend logged on the Travel tab also draws from your flex pool.</div>
        </div>

        {/* SUMMARY */}
        <div className="summary-bar">
          <div className="summary-item"><div className="summary-label">Take-Home</div><div className="summary-value green">{fmt(income)}</div></div>
          <div className="summary-div" />
          <div className="summary-item"><div className="summary-label">Fixed</div><div className="summary-value">{fmt(s.tFixed)}</div></div>
          <div className="summary-div" />
          <div className="summary-item"><div className="summary-label">Savings/Inv</div><div className="summary-value">{fmt(s.tSave)}</div></div>
          <div className="summary-div" />
          <div className="summary-item"><div className="summary-label">Flex Spent</div><div className="summary-value">{fmt(flexSpent)}</div></div>
          <div className="summary-div" />
          <div className="summary-item"><div className="summary-label">Free / Save More</div><div className={`summary-value${s.leftover>=0?' green':' red'}`}>{fmt(s.leftover)}</div></div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════ SUBSCRIPTIONS ═══════════════
function Subs({subs,subsMonthlyTotal,setSubsSave,showToast}) {
  const add = () => {
    const name=document.getElementById('sub-name').value.trim();
    const amount=parseFloat(document.getElementById('sub-amount').value);
    const cycle=document.getElementById('sub-cycle').value;
    const ess=document.getElementById('sub-ess').value;
    if(!name||isNaN(amount)||amount<=0){showToast('Add service + amount');return;}
    setSubsSave([...subs,{name,amount,cycle,ess}]);
    document.getElementById('sub-name').value='';document.getElementById('sub-amount').value='';
    showToast('Added');
  };
  const del = i => { const s=[...subs];s.splice(i,1);setSubsSave(s);showToast('Removed'); };
  const edit = (i,f,v) => {
    const s=JSON.parse(JSON.stringify(subs));
    if(f==='amount'){const n=parseFloat(v.replace(/[$,]/g,''));if(!isNaN(n)&&n>=0)s[i].amount=n;}
    else s[i][f]=v;setSubsSave(s);
  };
  const toggleCycle = i => { const s=JSON.parse(JSON.stringify(subs));s[i].cycle=s[i].cycle==='monthly'?'annual':'monthly';setSubsSave(s); };
  const total = subsMonthlyTotal();
  const nonEss = subs.filter(x=>x.ess==='non');
  const nonTot = nonEss.reduce((s,x)=>s+(x.cycle==='annual'?x.amount/12:x.amount),0);

  return (
    <div className="page"><div className="container">
      <div className="ph ph-flex">
        <div><div className="pt">Subscriptions</div><div className="psub">Monthly total feeds into Overview → Fixed → Committed</div></div>
        <div style={{textAlign:'right'}}><div style={{fontSize:'10px',textTransform:'uppercase',letterSpacing:'0.08em',color:'var(--muted)'}}>Monthly total</div><div style={{fontFamily:'Instrument Serif, serif',fontSize:'30px',letterSpacing:'-0.5px'}}>${total.toFixed(2)}</div></div>
      </div>
      <div className="add-form">
        <div className="aform-grid" style={{gridTemplateColumns:'1.4fr 100px 1fr 1fr auto'}}>
          <div className="form-group"><div className="form-label">Service</div><input className="form-input" id="sub-name" placeholder="e.g. Netflix" onKeyDown={e=>{if(e.key==='Enter')add();}} /></div>
          <div className="form-group"><div className="form-label">Amount $</div><input className="form-input" id="sub-amount" type="number" step="0.01" placeholder="0.00" onKeyDown={e=>{if(e.key==='Enter')add();}} /></div>
          <div className="form-group"><div className="form-label">Billing</div><select className="form-input form-select" id="sub-cycle"><option value="monthly">Monthly</option><option value="annual">Annual</option></select></div>
          <div className="form-group"><div className="form-label">Necessity</div><select className="form-input form-select" id="sub-ess"><option value="essential">Essential</option><option value="non">Non-essential</option></select></div>
          <button className="add-btn" onClick={add}>+ Add</button>
        </div>
      </div>
      <div className="item-list">
        <div className="li-row header-row" style={{gridTemplateColumns:'1.4fr 1fr 1fr 1fr 110px 32px'}}>
          <div className="li-cell head">Service</div><div className="li-cell head right">Amount</div><div className="li-cell head center">Billing</div><div className="li-cell head right">Monthly Eq.</div><div className="li-cell head center">Necessity</div><div className="li-cell" />
        </div>
        {subs.map((x,i)=>{
          const me=x.cycle==='annual'?x.amount/12:x.amount;
          return (
            <div key={i} className="li-row" style={{gridTemplateColumns:'1.4fr 1fr 1fr 1fr 110px 32px'}}>
              <div className="li-cell"><input className="edit-inline" defaultValue={x.name} onBlur={e=>edit(i,'name',e.target.value)} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} /></div>
              <div className="li-cell right"><input className="edit-inline" style={{textAlign:'right'}} defaultValue={`$${x.amount}`} onBlur={e=>edit(i,'amount',e.target.value)} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} /></div>
              <div className="li-cell center"><span className={`cycle-toggle${x.cycle==='annual'?' annual':''}`} onClick={()=>toggleCycle(i)}>{x.cycle==='annual'?'Annual':'Monthly'}</span></div>
              <div className="li-cell right">${me.toFixed(2)}</div>
              <div className="li-cell center"><span className={`chip${x.ess==='essential'?' ess':' non'}`} style={{cursor:'pointer'}} onClick={()=>edit(i,'ess',x.ess==='essential'?'non':'essential')}>{x.ess==='essential'?'Essential':'Non-ess'}</span></div>
              <div className="li-cell center"><button className="del-btn" onClick={()=>del(i)}>×</button></div>
            </div>
          );
        })}
      </div>
      {nonEss.length>0&&<div className="insight-box" style={{marginTop:'24px'}}><div className="insight-title">Trim opportunity</div><div className="insight-text">You have <b>{nonEss.length}</b> non-essential subscription{nonEss.length>1?'s':''} costing <b>${nonTot.toFixed(2)}/mo</b> (${(nonTot*12).toFixed(0)}/yr). Consider cancelling: {nonEss.map(x=>x.name).join(', ')}.</div></div>}
    </div></div>
  );
}

// ═══════════════ UTILITIES ═══════════════
function Utils({utils,utilsMonthlyTotal,setUtilsSave,showToast}) {
  const add = () => {
    const name=document.getElementById('util-name').value.trim();
    const amount=parseFloat(document.getElementById('util-amount').value);
    const split=document.getElementById('util-split').value.trim();
    if(!name||isNaN(amount)||amount<=0){showToast('Add utility + amount');return;}
    setUtilsSave([...utils,{name,amount,split}]);
    document.getElementById('util-name').value='';document.getElementById('util-amount').value='';document.getElementById('util-split').value='';
    showToast('Added');
  };
  const del = i => { const u=[...utils];u.splice(i,1);setUtilsSave(u);showToast('Removed'); };
  const edit = (i,f,v) => {
    const u=JSON.parse(JSON.stringify(utils));
    if(f==='amount'){const n=parseFloat(v.replace(/[$,]/g,''));if(!isNaN(n)&&n>=0)u[i].amount=n;}
    else u[i][f]=v;setUtilsSave(u);
  };
  return (
    <div className="page"><div className="container">
      <div className="ph ph-flex">
        <div><div className="pt">Utilities</div><div className="psub">Monthly total feeds into Overview → Fixed → Committed</div></div>
        <div style={{textAlign:'right'}}><div style={{fontSize:'10px',textTransform:'uppercase',letterSpacing:'0.08em',color:'var(--muted)'}}>Monthly total</div><div style={{fontFamily:'Instrument Serif, serif',fontSize:'30px',letterSpacing:'-0.5px'}}>${utilsMonthlyTotal().toFixed(2)}</div></div>
      </div>
      <div className="add-form">
        <div className="aform-grid" style={{gridTemplateColumns:'1.4fr 100px 1fr auto'}}>
          <div className="form-group"><div className="form-label">Utility</div><input className="form-input" id="util-name" placeholder="e.g. Electricity" onKeyDown={e=>{if(e.key==='Enter')add();}} /></div>
          <div className="form-group"><div className="form-label">Amount $/mo</div><input className="form-input" id="util-amount" type="number" step="0.01" placeholder="0.00" onKeyDown={e=>{if(e.key==='Enter')add();}} /></div>
          <div className="form-group"><div className="form-label">Split with</div><input className="form-input" id="util-split" placeholder="optional" /></div>
          <button className="add-btn" onClick={add}>+ Add</button>
        </div>
      </div>
      <div className="item-list">
        <div className="li-row header-row" style={{gridTemplateColumns:'1.4fr 1fr 1fr 32px'}}>
          <div className="li-cell head">Utility</div><div className="li-cell head right">My Share /mo</div><div className="li-cell head center">Split</div><div className="li-cell" />
        </div>
        {utils.map((x,i)=>(
          <div key={i} className="li-row" style={{gridTemplateColumns:'1.4fr 1fr 1fr 32px'}}>
            <div className="li-cell"><input className="edit-inline" defaultValue={x.name} onBlur={e=>edit(i,'name',e.target.value)} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} /></div>
            <div className="li-cell right"><input className="edit-inline" style={{textAlign:'right'}} defaultValue={`$${x.amount}`} onBlur={e=>edit(i,'amount',e.target.value)} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} /></div>
            <div className="li-cell center exp-note">{x.split?'w/ '+x.split:'—'}</div>
            <div className="li-cell center"><button className="del-btn" onClick={()=>del(i)}>×</button></div>
          </div>
        ))}
      </div>
    </div></div>
  );
}

// ═══════════════ INVEST ═══════════════
function Invest({budget,income,loan,goals,retVal,contribVal,setRetVal,setContribVal,setGoalsSave,setLoanSave,subs,showToast,fmt,fmtd}) {
  const invItems = budget.save;
  const investMonthly = sumA(invItems);
  const rate = Math.round(investMonthly/income*100);
  const projFV = (m,ar,y) => { const r=ar/12,n=y*12;if(r===0)return m*n;return m*((Math.pow(1+r,n)-1)/r); };
  const years=[1,2,3,5,7,10];
  const maxFV = Math.max(...years.map(y=>projFV(contribVal,retVal/100,y)));
  const nonEssAnnual=subs.filter(x=>x.ess==='non').reduce((s,x)=>s+(x.cycle==='annual'?x.amount:x.amount*12),0);

  const addGoal = () => {
    const name=document.getElementById('goal-name').value.trim();
    const target=parseFloat(document.getElementById('goal-target').value);
    const saved=parseFloat(document.getElementById('goal-saved').value)||0;
    const monthly=parseFloat(document.getElementById('goal-monthly').value)||0;
    if(!name||isNaN(target)||target<=0){showToast('Add goal + target');return;}
    setGoalsSave([...goals,{name,target,saved,monthly}]);
    ['goal-name','goal-target','goal-saved','goal-monthly'].forEach(id=>document.getElementById(id).value='');
    showToast('Goal added');
  };
  const delGoal = i => { const g=[...goals];g.splice(i,1);setGoalsSave(g);showToast('Removed'); };

  const palette=['#1a1916','#2d6a4f','#1d3557','#b08968','#9a9690','#c1440e','#5a189a','#6a040f'];
  const r=70,circ=2*Math.PI*r;
  let off=0;
  const donutPaths = invItems.map((it,i)=>{const pct=investMonthly>0?it.amount/investMonthly:0;const dash=pct*circ;const el=<circle key={i} r={r} cx={100} cy={100} fill="none" stroke={palette[i%palette.length]} strokeWidth="22" strokeDasharray={`${dash} ${circ-dash}`} strokeDashoffset={-off} />;off+=dash;return el;});

  return (
    <div className="page"><div className="container">
      <div className="ph"><div className="pt">Savings & Investments</div><div className="psub">Plan allocation, project growth, track goals</div></div>
      <div className="kpi-strip">
        <div className="kpi-c"><div className="kpi-c-label">Invested / mo</div><div className="kpi-c-val">{fmt(investMonthly)}</div><div className="kpi-c-sub">From Overview</div></div>
        <div className="kpi-c"><div className="kpi-c-label">Savings Rate</div><div className="kpi-c-val">{rate}%</div><div className="kpi-c-sub">of take-home</div></div>
        <div className="kpi-c"><div className="kpi-c-label">India Loan</div>
          <div className="kpi-c-val" contentEditable suppressContentEditableWarning onBlur={e=>{const n=parseFloat(e.target.innerText.replace(/[$,]/g,''));if(!isNaN(n))setLoanSave(n);}}>{fmt(loan)}</div>
          <div className="kpi-c-sub">9.8% · click to edit</div></div>
        <div className="kpi-c"><div className="kpi-c-label">10yr Net Worth</div><div className="kpi-c-val">{fmt(projFV(contribVal,retVal/100,10))}</div><div className="kpi-c-sub">projected</div></div>
      </div>
      <div className="insight-box"><div className="insight-title">Plan snapshot</div><div className="insight-text">You're investing <b>{fmt(investMonthly)}/mo</b> (~<b>{rate}%</b> of take-home). Your India loan at <b>9.8%</b> is effectively a guaranteed return, so the prepay is well-placed. {nonEssAnnual>0&&`Redirecting $${(nonEssAnnual/12).toFixed(0)}/mo of non-essential subscriptions into ETFs would compound to ~${fmt(projFV(nonEssAnnual/12,0.10,10))} over 10 years at 10%.`}</div></div>

      <div className="subsec-title">Allocation</div>
      <div className="alloc-flex">
        <div style={{position:'relative',width:'200px',height:'200px'}}>
          <svg style={{transform:'rotate(-90deg)'}} width="200" height="200" viewBox="0 0 200 200">{donutPaths}</svg>
          <div style={{position:'absolute',top:'50%',left:'50%',transform:'translate(-50%,-50%)',textAlign:'center'}}>
            <div style={{fontFamily:'Instrument Serif, serif',fontSize:'22px',letterSpacing:'-0.5px'}}>{fmt(investMonthly)}</div>
            <div style={{fontSize:'9px',color:'var(--muted)',textTransform:'uppercase',letterSpacing:'0.08em'}}>per month</div>
          </div>
        </div>
        <div className="ring-legend">
          {invItems.map((it,i)=>{const pct=investMonthly>0?Math.round(it.amount/investMonthly*100):0;return(
            <div key={i} className="legend-row"><span className="legend-dot" style={{background:palette[i%palette.length]}} /><span className="legend-name">{it.emoji} {it.name}</span><span className="legend-pct">{pct}%</span><span className="legend-amt">{fmt(it.amount)}</span></div>
          );})}
        </div>
      </div>

      <div className="subsec-title">Growth Projection</div>
      <div className="proj-controls">
        <div className="slider-group"><label>Annual return <b>{retVal}%</b></label><input type="range" min="4" max="15" value={retVal} onChange={e=>setRetVal(+e.target.value)} /></div>
        <div className="slider-group"><label>Monthly invested <b>{fmt(contribVal)}</b></label><input type="range" min="0" max="4000" step="50" value={contribVal} onChange={e=>setContribVal(+e.target.value)} /></div>
      </div>
      <div className="proj-chart">
        <div className="proj-bars">
          {years.map(y=>{const fv=projFV(contribVal,retVal/100,y);const contributed=contribVal*12*y;const stackH=maxFV>0?(fv/maxFV)*100:0;const contribH=fv>0?(contributed/fv)*100:100;const growthH=100-contribH;return(
            <div key={y} className="pbar-col"><div className="pbar-val">{fmt(fv)}</div><div className="pbar-stack" style={{height:stackH+'%'}}><div className="pbar-growth" style={{height:growthH+'%'}} /><div className="pbar-contrib" style={{height:contribH+'%'}} /></div><div className="pbar-yr">Yr {y}</div></div>
          );})}
        </div>
        <div style={{display:'flex',gap:'20px',marginTop:'18px',justifyContent:'center'}}>
          <div style={{display:'flex',alignItems:'center',gap:'6px'}}><span style={{width:'10px',height:'10px',borderRadius:'2px',background:'var(--border2)',display:'inline-block'}} /><span style={{fontSize:'11px',color:'var(--muted)'}}>Contributed</span></div>
          <div style={{display:'flex',alignItems:'center',gap:'6px'}}><span style={{width:'10px',height:'10px',borderRadius:'2px',background:'var(--save)',display:'inline-block'}} /><span style={{fontSize:'11px',color:'var(--muted)'}}>Market growth</span></div>
        </div>
      </div>

      <div className="subsec-title">Goals</div>
      <div className="add-form">
        <div className="aform-grid" style={{gridTemplateColumns:'1.4fr 1fr 1fr 1fr auto'}}>
          <div className="form-group"><div className="form-label">Goal</div><input className="form-input" id="goal-name" placeholder="e.g. Emergency 6mo" /></div>
          <div className="form-group"><div className="form-label">Target $</div><input className="form-input" id="goal-target" type="number" placeholder="15000" /></div>
          <div className="form-group"><div className="form-label">Saved $</div><input className="form-input" id="goal-saved" type="number" placeholder="0" /></div>
          <div className="form-group"><div className="form-label">Monthly $</div><input className="form-input" id="goal-monthly" type="number" placeholder="500" /></div>
          <button className="add-btn" onClick={addGoal}>+ Add</button>
        </div>
      </div>
      {goals.length===0?<div className="empty-state"><div className="empty-state-icon">◎</div><div>No goals yet. Add one above.</div></div>:goals.map((x,i)=>{
        const pct=Math.min(100,Math.round(x.saved/x.target*100));const remaining=Math.max(0,x.target-x.saved);const months=x.monthly>0?Math.ceil(remaining/x.monthly):null;
        return (<div key={i} className="goal-card"><div className="goal-top"><div><div className="goal-name">{x.name}</div><div className="goal-meta">Target {fmt(x.target)} · {fmt(x.monthly)}/mo</div></div><button className="del-btn" onClick={()=>delGoal(i)}>×</button></div><div className="goal-prog-track"><div className="goal-prog-fill" style={{width:pct+'%'}} /></div><div className="goal-stats"><span><b>{fmt(x.saved)}</b> saved</span><span>{pct}%</span><span><b>{fmt(remaining)}</b> to go</span></div>{months!==null&&<div className="goal-eta">≈ {months} month{months!==1?'s':''} to reach goal</div>}</div>);
      })}
    </div></div>
  );
}

// ═══════════════ TRAVEL ═══════════════
function Travel({trips,showTripForm,setShowTripForm,setTripsSave,showToast,monthKey,fmt,fmtd}) {
  const addTrip = () => {
    const name=document.getElementById('trip-name').value.trim();
    const when=document.getElementById('trip-when').value.trim();
    const budget=parseFloat(document.getElementById('trip-budget').value)||0;
    const saved=parseFloat(document.getElementById('trip-saved').value)||0;
    if(!name){showToast('Add destination');return;}
    setTripsSave([...trips,{name,when,budget,saved,expenses:[]}]);
    ['trip-name','trip-when','trip-budget','trip-saved'].forEach(id=>document.getElementById(id).value='');
    setShowTripForm(false);showToast('Trip added');
  };
  const delTrip = i => { if(!confirm('Delete this trip?'))return;const t=[...trips];t.splice(i,1);setTripsSave(t);showToast('Deleted'); };
  const editTrip = (i,f,v) => {
    const t=JSON.parse(JSON.stringify(trips));
    const n=parseFloat(v.replace(/[$,]/g,''));
    if(f==='budget'||f==='saved'){if(!isNaN(n))t[i][f]=n;}else t[i][f]=v;
    setTripsSave(t);
  };
  const addTripExpense = i => {
    const desc=document.getElementById('te-desc-'+i).value.trim();
    const amt=parseFloat(document.getElementById('te-amt-'+i).value);
    const cat=document.getElementById('te-cat-'+i).value;
    if(!desc||isNaN(amt)||amt<=0){showToast('Add item + amount');return;}
    const t=JSON.parse(JSON.stringify(trips));
    t[i].expenses.push({desc,amt,cat,month:monthKey});
    setTripsSave(t);showToast('Logged');
  };
  const delTripExp = (ti,ei) => { const t=JSON.parse(JSON.stringify(trips));t[ti].expenses.splice(ei,1);setTripsSave(t); };
  return (
    <div className="page"><div className="container">
      <div className="ph ph-flex">
        <div><div className="pt">Travel</div><div className="psub">Plan trips, set budgets, save up, and track spend</div></div>
        <button className="add-btn" onClick={()=>setShowTripForm(!showTripForm)}>+ New trip</button>
      </div>
      {showTripForm&&<div className="add-form">
        <div className="aform-grid" style={{gridTemplateColumns:'1.4fr 1fr 1fr 1fr'}}>
          <div className="form-group"><div className="form-label">Destination</div><input className="form-input" id="trip-name" placeholder="e.g. Tetons" /></div>
          <div className="form-group"><div className="form-label">When</div><input className="form-input" id="trip-when" placeholder="e.g. Aug 2026" /></div>
          <div className="form-group"><div className="form-label">Budget $</div><input className="form-input" id="trip-budget" type="number" placeholder="2000" /></div>
          <div className="form-group"><div className="form-label">Saved so far $</div><input className="form-input" id="trip-saved" type="number" placeholder="0" /></div>
        </div>
        <div style={{marginTop:'12px',display:'flex',gap:'8px'}}>
          <button className="add-btn" onClick={addTrip}>Create trip</button>
          <button className="add-btn btn-ghost" onClick={()=>setShowTripForm(false)}>Cancel</button>
        </div>
      </div>}
      {trips.length===0&&<div className="empty-state"><div className="empty-state-icon">✈</div><div>No trips planned yet.</div></div>}
      {trips.map((t,i)=>{
        const spent=t.expenses.reduce((s,e)=>s+e.amt,0);
        const budgetLeft=t.budget-spent;
        const savePct=t.budget>0?Math.min(100,Math.round(t.saved/t.budget*100)):0;
        const thisMonthSpend=t.expenses.filter(e=>e.month===monthKey).reduce((s,e)=>s+e.amt,0);
        return (
          <div key={i} className="trip-card">
            <div className="trip-head">
              <div><div className="trip-name">{t.name}</div>
                <div className="trip-dates"><input className="edit-inline" style={{fontSize:'11px',color:'var(--muted)'}} defaultValue={t.when||''} placeholder="When?" onBlur={e=>editTrip(i,'when',e.target.value)} /></div>
              </div>
              <div style={{display:'flex',gap:'8px',alignItems:'center'}}>
                <span className={`trip-status${spent>0?' done':t.saved>=t.budget&&t.budget>0?' saving':' planning'}`}>{spent>0?'In progress':t.saved>=t.budget&&t.budget>0?'Funded':'Planning'}</span>
                <button className="del-btn" onClick={()=>delTrip(i)}>×</button>
              </div>
            </div>
            <div className="trip-budget-grid">
              <div><div className="tbg-label">Budget</div><div className="tbg-val"><input className="edit-inline" defaultValue={fmtd(t.budget)} style={{fontFamily:'Instrument Serif, serif',fontSize:'20px'}} onBlur={e=>editTrip(i,'budget',e.target.value)} /></div></div>
              <div><div className="tbg-label">Spent</div><div className="tbg-val">{fmtd(spent)}</div></div>
              <div><div className="tbg-label">{spent>0?'Budget left':'Saved'}</div>
                <div className={`tbg-val${spent>0?(budgetLeft<0?' over':' under'):''}`}>
                  {spent>0?fmtd(budgetLeft):<input className="edit-inline" defaultValue={fmtd(t.saved)} style={{fontFamily:'Instrument Serif, serif',fontSize:'20px'}} onBlur={e=>editTrip(i,'saved',e.target.value)} />}
                </div>
              </div>
            </div>
            {spent===0?<><div className="trip-prog-track"><div className="trip-prog-fill" style={{width:savePct+'%',background:'var(--gold)'}} /></div><div style={{fontSize:'11px',color:'var(--muted)',marginBottom:'6px'}}>{savePct}% saved toward this trip</div></>
              :<div className="trip-prog-track"><div className="trip-prog-fill" style={{width:Math.min(100,Math.round(spent/t.budget*100))+'%',background:budgetLeft<0?'var(--red)':'var(--save)'}} /></div>}
            {thisMonthSpend>0&&<div className="trip-month-note">{fmtd(thisMonthSpend)} of this spent in {monthKey} — counted in your Overview flex pool</div>}
            {t.expenses.map((e,ei)=>(
              <div key={ei} className="trip-line"><span>{e.desc} <span className="chip">{e.cat}</span> {e.month&&<span className="exp-note">{e.month}</span>}</span><span>{fmtd(e.amt)} <button className="del-btn" onClick={()=>delTripExp(i,ei)}>×</button></span></div>
            ))}
            <div className="trip-expense-form">
              <input className="form-input" id={'te-desc-'+i} placeholder="Add expense..." />
              <select className="form-input form-select" id={'te-cat-'+i}><option>Flights</option><option>Lodging</option><option>Food</option><option>Activities</option><option>Transport</option><option>Misc</option></select>
              <input className="form-input" id={'te-amt-'+i} type="number" placeholder="$" />
              <button className="add-btn" onClick={()=>addTripExpense(i)}>+</button>
            </div>
          </div>
        );
      })}
    </div></div>
  );
}

// ═══════════════ TRACKER ═══════════════
function Tracker({expenses,cards,budget,curMonth,curYear,setCurMonth,setCurYear,monthKey,setExpensesSave,notes,setNotes,showToast,fmt,fmtd}) {
  const changeMonth = d => { let m=curMonth+d,y=curYear;if(m>11){m=0;y++;}if(m<0){m=11;y--;}setCurMonth(m);setCurYear(y); };
  const mk = monthKey(curMonth,curYear);
  const items = expenses[mk]||[];
  const total = items.reduce((s,i)=>s+i.amount,0);
  const livingBudget = sumA(budget.living);
  const rem = livingBudget-total;

  const addExpense = () => {
    const desc=document.getElementById('exp-desc').value.trim();
    const amount=parseFloat(document.getElementById('exp-amount').value);
    const cat=document.getElementById('exp-cat').value;
    const card=document.getElementById('exp-card').value;
    if(!desc||isNaN(amount)||amount<=0){showToast('Add description + amount');return;}
    const all=JSON.parse(JSON.stringify(expenses));
    if(!all[mk])all[mk]=[];
    const t=new Date();all[mk].push({desc,amount,cat,card,date:`${t.getMonth()+1}/${t.getDate()}`});
    setExpensesSave(all);
    document.getElementById('exp-desc').value='';document.getElementById('exp-amount').value='';
    showToast('Added');
  };
  const delExp = i => { const all=JSON.parse(JSON.stringify(expenses));all[mk].splice(i,1);setExpensesSave(all);showToast('Deleted'); };
  const editExp = (i,f,v) => {
    const all=JSON.parse(JSON.stringify(expenses));if(!all[mk]||!all[mk][i])return;
    if(f==='amount'){const n=parseFloat(v.replace(/[$,]/g,''));if(!isNaN(n)&&n>=0)all[mk][i].amount=n;}else all[mk][i][f]=v;
    setExpensesSave(all);
  };

  const cats={};items.forEach(i=>cats[i.cat]=(cats[i.cat]||0)+i.amount);
  const sortedCats=Object.entries(cats).sort((a,b)=>b[1]-a[1]);
  const maxCat=sortedCats[0]?.[1]||1;

  return (
    <div className="page"><div className="container">
      <div className="ph ph-flex">
        <div className="pt">Living Expenses</div>
        <div className="month-nav">
          <button className="month-btn" onClick={()=>changeMonth(-1)}>←</button>
          <div className="month-label">{MONTHS[curMonth]} {curYear}</div>
          <button className="month-btn" onClick={()=>changeMonth(1)}>→</button>
        </div>
      </div>
      <div className="month-summary">
        <div className="msm-card"><div className="msm-label">Spent</div><div className="msm-val red">{fmtd(total)}</div></div>
        <div className="msm-card"><div className="msm-label">Living Budget</div><div className="msm-val">{fmt(livingBudget)}</div></div>
        <div className="msm-card"><div className="msm-label">Remaining</div><div className={`msm-val${rem>=0?' green':' red'}`}>{fmt(rem)}</div></div>
      </div>
      <div className="add-form">
        <div className="aform-grid" style={{gridTemplateColumns:'1.4fr 100px 1fr 1fr auto'}}>
          <div className="form-group"><div className="form-label">Description</div><input className="form-input" id="exp-desc" placeholder="e.g. Walmart run" onKeyDown={e=>{if(e.key==='Enter')addExpense();}} /></div>
          <div className="form-group"><div className="form-label">Amount $</div><input className="form-input" id="exp-amount" type="number" step="0.01" placeholder="0.00" onKeyDown={e=>{if(e.key==='Enter')addExpense();}} /></div>
          <div className="form-group"><div className="form-label">Category</div><select className="form-input form-select" id="exp-cat"><option>Groceries</option><option>Eating Out</option><option>Lunches</option><option>Transport</option><option>Health</option><option>Misc</option></select></div>
          <div className="form-group"><div className="form-label">Card</div><select className="form-input form-select" id="exp-card">{cards.map((c,i)=><option key={i}>{c.name}</option>)}<option>Cash</option></select></div>
          <button className="add-btn" onClick={addExpense}>+ Add</button>
        </div>
        <div className="feed-note" style={{marginTop:'10px'}}>↳ These appear as "Actual" under Fixed → Living on the Overview.</div>
      </div>
      <div className="item-list">
        <div className="expense-item header-row">
          <div className="exp-cell head">Description</div><div className="exp-cell head center">Category</div><div className="exp-cell head center hide-mobile">Card</div><div className="exp-cell head right">Amount</div><div className="exp-cell head center hide-mobile">Date</div><div className="exp-cell" />
        </div>
        {items.map((it,idx)=>({it,idx})).reverse().map(({it,idx})=>(
            <div key={idx} className="expense-item">
              <div className="exp-cell"><input className="edit-inline" defaultValue={it.desc} onBlur={e=>editExp(idx,'desc',e.target.value)} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} /></div>
              <div className="exp-cell center"><span className="chip">{it.cat}</span></div>
              <div className="exp-cell center hide-mobile"><span className="chip" style={{background:'rgba(0,0,0,0.03)'}}>{it.card||'Cash'}</span></div>
              <div className="exp-cell right"><input className="edit-inline" style={{textAlign:'right'}} defaultValue={fmtd(it.amount)} onBlur={e=>editExp(idx,'amount',e.target.value)} onKeyDown={e=>{if(e.key==='Enter')e.target.blur();}} /></div>
              <div className="exp-cell center hide-mobile exp-note">{it.date}</div>
              <div className="exp-cell center"><button className="del-btn" onClick={()=>delExp(idx)}>×</button></div>
            </div>
        ))}
      </div>
      {items.length===0&&<div className="empty-state"><div className="empty-state-icon">◦</div><div>No expenses logged this month.</div></div>}
      {sortedCats.length>0&&<div className="cat-breakdown">
        <div style={{fontSize:'10px',textTransform:'uppercase',letterSpacing:'0.08em',color:'var(--muted)',marginBottom:'14px'}}>By Category</div>
        {sortedCats.map(([cat,amt])=>{const pct=Math.round(amt/total*100),w=Math.round(amt/maxCat*100);return(
          <div key={cat} className="cat-row"><div className="cat-name">{cat}</div><div className="cat-bar-wrap"><div className="cat-bar" style={{width:w+'%'}} /></div><div className="cat-amt">{fmtd(amt)} <span style={{color:'var(--muted)',fontSize:'10px'}}>{pct}%</span></div></div>
        );})}
      </div>}
      <div className="notes-section">
        <div className="notes-label">Monthly Notes</div>
        <textarea className="notes-area" placeholder="Unusual expenses, reminders, goals..." defaultValue={notes[mk]||''}
          onBlur={e=>{const n={...notes};n[mk]=e.target.value;setNotes(n);}} />
      </div>
    </div></div>
  );
}

// ═══════════════ CARDS ═══════════════
function Cards({cards,expenses,curMonth,curYear,setCurMonth,setCurYear,monthKey,showCardForm,setShowCardForm,cardColor,setCardColor,setCardsSave,showToast,fmt,fmtd}) {
  const changeMonth = d => { let m=curMonth+d,y=curYear;if(m>11){m=0;y++;}if(m<0){m=11;y--;}setCurMonth(m);setCurYear(y); };
  const mk = monthKey(curMonth,curYear);
  const spend={};(expenses[mk]||[]).forEach(i=>{const c=i.card||'Cash';spend[c]=(spend[c]||0)+i.amount;});
  const saveCard = () => {
    const name=document.getElementById('card-name').value.trim();
    const issuer=document.getElementById('card-issuer').value.trim()||'••••';
    if(!name){showToast('Card name required');return;}
    setCardsSave([...cards,{name,issuer,color:cardColor}]);
    document.getElementById('card-name').value='';document.getElementById('card-issuer').value='';
    setShowCardForm(false);showToast('Card added');
  };
  const delCard = i => { const c=[...cards];c.splice(i,1);setCardsSave(c);showToast('Card removed'); };
  const cardEntries=Object.entries(spend).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);
  const totalSpend=cardEntries.reduce((s,[,v])=>s+v,0);
  const maxSpend=cardEntries[0]?.[1]||1;
  return (
    <div className="page"><div className="container">
      <div className="ph ph-flex">
        <div className="pt">Cards</div>
        <div className="month-nav">
          <button className="month-btn" onClick={()=>changeMonth(-1)}>←</button>
          <div className="month-label">{MONTHS[curMonth]} {curYear}</div>
          <button className="month-btn" onClick={()=>changeMonth(1)}>→</button>
        </div>
      </div>
      {showCardForm&&<div className="add-form">
        <div className="aform-grid" style={{gridTemplateColumns:'1.5fr 1fr 1fr auto'}}>
          <div className="form-group"><div className="form-label">Card name</div><input className="form-input" id="card-name" placeholder="e.g. Amex Gold" /></div>
          <div className="form-group"><div className="form-label">Issuer / last 4</div><input className="form-input" id="card-issuer" placeholder="•••• 1234" /></div>
          <div className="form-group"><div className="form-label">Color</div>
            <div className="color-pick">{CARD_COLORS.map(c=><div key={c} className={`color-dot${cardColor===c?' sel':''}`} style={{background:c}} onClick={()=>setCardColor(c)} />)}</div>
          </div>
          <button className="add-btn" onClick={saveCard}>Add card</button>
        </div>
      </div>}
      <div className="cards-grid">
        {cards.map((c,i)=>{const s=spend[c.name]||0;const cnt=(expenses[mk]||[]).filter(e=>(e.card||'Cash')===c.name).length;return(
          <div key={i} className="credit-card" style={{background:`linear-gradient(135deg,${c.color},${shade(c.color,-25)})`}}>
            <div className="cc-top"><div><div className="cc-name">{c.name}</div><div className="cc-issuer">{c.issuer}</div></div><button className="cc-del" onClick={()=>delCard(i)}>×</button></div>
            <div className="cc-bottom"><div className="cc-spend-label">Spent this month</div><div className="cc-spend">{fmtd(s)}</div><div className="cc-count">{cnt} transaction{cnt!==1?'s':''}</div></div>
          </div>
        );})}
        <div className="add-card-tile" onClick={()=>setShowCardForm(!showCardForm)}><div className="act-icon">+</div><div style={{fontSize:'11px'}}>Add a card</div></div>
      </div>
      {cardEntries.length>0&&<div><div className="cd-title">Spend by Card</div><div className="cd-bars">{cardEntries.map(([name,amt])=>{const card=cards.find(c=>c.name===name);const color=card?card.color:'#9a9690';const w=Math.round(amt/maxSpend*100);const pct=Math.round(amt/totalSpend*100);return(
        <div key={name} className="cd-bar-row"><div className="cd-bar-name"><span style={{width:'10px',height:'10px',borderRadius:'3px',background:color,display:'inline-block'}} />{name}</div><div className="cd-bar-track"><div className="cd-bar-fill" style={{width:w+'%',background:color}} /></div><div className="cd-bar-amt">{fmtd(amt)} <span style={{color:'var(--muted)',fontSize:'10px'}}>{pct}%</span></div></div>
      );})}</div></div>}
    </div></div>
  );
}

// ═══════════════ CHARTS ═══════════════
function Charts({monthSnapshot,curMonth,curYear,setCurMonth,setCurYear,monthKey,budget,cards,fmt,fmtd}) {
  const changeMonth = d => { let m=curMonth+d,y=curYear;if(m>11){m=0;y++;}if(m<0){m=11;y--;}setCurMonth(m);setCurYear(y); };
  const mk = monthKey(curMonth,curYear);
  const s = monthSnapshot(mk);

  const flowSegs=[{l:'Income',v:s.income,c:'#2d6a4f'},{l:'Fixed',v:s.tFixed,c:'#1a1916'},{l:'Savings & Inv',v:s.tSave,c:'#1d3557'},{l:'Flex spent',v:s.flexSpent,c:'#c1440e'},{l:'Free',v:Math.max(0,s.leftover),c:'#b08968'}];
  const donutSegs=[{l:'Fixed',v:s.tFixed,c:'#1a1916'},{l:'Savings & Inv',v:s.tSave,c:'#2d6a4f'},{l:'Flex spent',v:s.flexSpent,c:'#c1440e'},{l:'Free',v:Math.max(0,s.leftover),c:'#d4d0c8'}];

  const cats={};s.expItems.forEach(i=>cats[i.cat]=(cats[i.cat]||0)+i.amount);
  const sortedCats=Object.entries(cats).sort((a,b)=>b[1]-a[1]);
  const maxCat=sortedCats[0]?.[1]||1;

  const actualCat={};s.expItems.forEach(i=>actualCat[i.cat]=(actualCat[i.cat]||0)+i.amount);
  const maxBva=Math.max(...budget.living.map(l=>Math.max(l.amount,actualCat[l.cat]||0)),1);

  // 6-month trend
  const months6=[];for(let i=5;i>=0;i--){let m=curMonth-i,y=curYear;while(m<0){m+=12;y--;}months6.push({mk:`${y}-${String(m+1).padStart(2,'0')}`,label:MONTHS[m].slice(0,3)});}
  const trendData=months6.map(mo=>{const snap=monthSnapshot(mo.mk);return{label:mo.label,total:snap.livingActual+snap.flexSpent};});
  const maxT=Math.max(...trendData.map(d=>d.total),1);

  const cardSpend={};s.expItems.forEach(i=>{const c=i.card||'Cash';cardSpend[c]=(cardSpend[c]||0)+i.amount;});
  const cardEntries=Object.entries(cardSpend).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1]);
  const maxCard=cardEntries[0]?.[1]||1;

  // mini donut
  const r=70,circ=2*Math.PI*r;let off=0;
  const donutPaths=donutSegs.map((d,i)=>{const pct=s.income>0?Math.max(0,d.v)/s.income:0;const dash=pct*circ;const el=<circle key={i} r={r} cx={100} cy={100} fill="none" stroke={d.c} strokeWidth="24" strokeDasharray={`${dash} ${circ-dash}`} strokeDashoffset={-off} />;off+=dash;return el;});

  return (
    <div className="page"><div className="container">
      <div className="ph ph-flex">
        <div><div className="pt">Charts</div><div className="psub">Your money, visualized</div></div>
        <div className="month-nav">
          <button className="month-btn" onClick={()=>changeMonth(-1)}>←</button>
          <div className="month-label">{MONTHS[curMonth]} {curYear}</div>
          <button className="month-btn" onClick={()=>changeMonth(1)}>→</button>
        </div>
      </div>
      <div className="subsec-title" style={{marginTop:0}}>Income → Outflow</div>
      <div className="chart-card"><div className="flow-wrap">{flowSegs.map(f=><div key={f.l} className="flow-bar" style={{width:Math.max(8,(f.v/s.income)*100)+'%',background:f.c}}><span>{f.l}</span><span>{fmt(f.v)}</span></div>)}</div></div>
      <div className="subsec-title">Where the money went</div>
      <div className="alloc-flex">
        <div style={{position:'relative',width:'200px',height:'200px'}}>
          <svg style={{transform:'rotate(-90deg)'}} width="200" height="200" viewBox="0 0 200 200">{donutPaths}</svg>
          <div style={{position:'absolute',top:'50%',left:'50%',transform:'translate(-50%,-50%)',textAlign:'center'}}>
            <div style={{fontFamily:'Instrument Serif, serif',fontSize:'20px'}}>{fmt(s.income)}</div>
            <div style={{fontSize:'9px',color:'var(--muted)',textTransform:'uppercase'}}>income</div>
          </div>
        </div>
        <div className="ring-legend">{donutSegs.map(d=><div key={d.l} className="legend-row"><span className="legend-dot" style={{background:d.c}} /><span className="legend-name">{d.l}</span><span className="legend-pct">{Math.round(d.v/s.income*100)}%</span><span className="legend-amt">{fmt(d.v)}</span></div>)}</div>
      </div>
      <div className="subsec-title">Living spend by category</div>
      <div className="chart-card">{sortedCats.length===0?<div className="fl-empty">No living expenses logged this month.</div>:sortedCats.map(([c,v],i)=><div key={c} className="bar-h-row"><div className="bar-h-label">{c}</div><div className="bar-h-track"><div className="bar-h-fill" style={{width:Math.round(v/maxCat*100)+'%',background:PAL[i%PAL.length]}} /></div><div className="bar-h-amt">{fmtd(v)}</div></div>)}</div>
      <div className="subsec-title">Budget vs actual — living</div>
      <div className="chart-card">{budget.living.map(l=>{const act=actualCat[l.cat]||0;const over=act>l.amount;return(<div key={l.name} className="bva-row"><div className="bar-h-label">{l.emoji} {l.name}</div><div className="bva-bars"><div className="bva-bar" style={{width:Math.max(10,l.amount/maxBva*100)+'%',background:'#d4d0c8',color:'#1a1916'}}>budget {fmt(l.amount)}</div><div className="bva-bar" style={{width:Math.max(10,act/maxBva*100)+'%',background:over?'#c1440e':'#2d6a4f'}}>actual {fmt(act)}</div></div></div>);})}</div>
      <div className="subsec-title">6-month spending trend</div>
      <div className="chart-card"><div className="trend-wrap">{trendData.map(d=><div key={d.label} className="trend-col"><div className="trend-val">{d.total>0?fmt(d.total):''}</div><div className="trend-bar" style={{height:Math.round(d.total/maxT*100)+'%'}} /><div className="trend-mo">{d.label}</div></div>)}</div></div>
      <div className="subsec-title">Spend by card</div>
      <div className="chart-card">{cardEntries.length===0?<div className="fl-empty">No card spend logged this month.</div>:cardEntries.map(([name,v])=>{const card=cards.find(c=>c.name===name);const col=card?card.color:'#9a9690';return<div key={name} className="bar-h-row"><div className="bar-h-label">{name}</div><div className="bar-h-track"><div className="bar-h-fill" style={{width:Math.round(v/maxCard*100)+'%',background:col}} /></div><div className="bar-h-amt">{fmtd(v)}</div></div>;})}</div>
    </div></div>
  );
}

// ═══════════════ SUGGESTIONS ═══════════════
function Suggest({monthSnapshot,mk,budget,subs,cards,expenses,flex,income,trips,goals,aiOutput,setAiOutput,aiLoading,setAiLoading,curMonth,curYear,fmt,fmtd}) {
  const s = monthSnapshot(mk);
  const actualCat={};s.expItems.forEach(i=>actualCat[i.cat]=(actualCat[i.cat]||0)+i.amount);
  const projFV=(m,ar,y)=>{const r=ar/12,n=y*12;if(r===0)return m*n;return m*((Math.pow(1+r,n)-1)/r);};
  const nonEss=subs.filter(x=>x.ess==='non');const nonTot=nonEss.reduce((a,x)=>a+(x.cycle==='annual'?x.amount/12:x.amount),0);
  const rate=Math.round(s.tSave/income*100);
  const sugs=[];
  if(rate>=30)sugs.push({icon:'🏆',tag:'win',title:`Strong ${rate}% savings rate`,text:`You're directing <b>${fmt(s.tSave)}/mo</b> to savings. That's above the typical 20% target.`});
  else if(rate<15)sugs.push({icon:'⚠️',tag:'watch',title:`Savings rate is ${rate}%`,text:`Aim for at least 20% (<b>${fmt(income*0.2)}/mo</b>). You're at <b>${fmt(s.tSave)}</b>.`});
  budget.living.forEach(l=>{const act=actualCat[l.cat]||0;if(act>l.amount*1.1)sugs.push({icon:'📈',tag:'watch',title:`${l.name} over budget`,text:`Spent <b>${fmt(act)}</b> vs <b>${fmt(l.amount)}</b> budgeted (${Math.round((act/l.amount-1)*100)}% over).`});});
  if(nonTot>0)sugs.push({icon:'✂️',tag:'win',title:'Trim subscriptions',text:`<b>${fmt(nonTot)}/mo</b> goes to non-essential subs: ${nonEss.map(x=>x.name).join(', ')}. Cancelling and investing could grow to ~<b>${fmt(projFV(nonTot,0.10,10))}</b> in 10yrs.`});
  if(s.leftover>200)sugs.push({icon:'💰',tag:'tip',title:`${fmt(s.leftover)} unspent this month`,text:`Sweep it into ETFs or India loan prepayment (9.8% guaranteed) rather than letting it sit.`});
  else if(s.leftover<0)sugs.push({icon:'🛑',tag:'watch',title:'Over your free budget',text:`You're <b>${fmt(-s.leftover)}</b> past your discretionary pool. Hold off on new flex buys.`});
  if(!sugs.length)sugs.push({icon:'✅',tag:'win',title:'All looking healthy',text:'No flags this month. Log more expenses to get sharper suggestions.'});

  const getAI = async () => {
    setAiLoading(true);
    const cardSpend={};s.expItems.forEach(i=>{const c=i.card||'Cash';cardSpend[c]=(cardSpend[c]||0)+i.amount;});
    const payload={month:`${MONTHS[curMonth]} ${curYear}`,income,fixed:s.tFixed,savings:s.tSave,living_budget:s.livingBudget,living_actual:s.livingActual,living_by_category:actualCat,flex_spent:s.flexSpent,travel:s.travel,leftover:s.leftover,subscriptions:subs,card_spend:cardSpend,india_loan:38550,savings_breakdown:budget.save.map(x=>({name:x.name,amount:x.amount}))};
    try {
      const res=await fetch('/api/suggest',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
      const data=await res.json();
      setAiOutput(data.text||data.error||'No response.');
    } catch { setAiOutput('Could not reach the AI API. Make sure ANTHROPIC_API_KEY is set in Vercel environment variables.'); }
    setAiLoading(false);
  };

  return (
    <div className="page"><div className="container">
      <div className="ph ph-flex">
        <div><div className="pt">Suggestions</div><div className="psub">Personalized advice on spending, savings, and card use</div></div>
        <button className="add-btn" onClick={getAI} disabled={aiLoading}>{aiLoading?'Analyzing…':'✦ Get AI analysis'}</button>
      </div>
      {aiLoading&&<div className="ai-loading"><div className="spinner" />Analyzing your finances…</div>}
      {aiOutput&&!aiLoading&&<div className="ai-card" style={{whiteSpace:'pre-wrap'}}>{aiOutput}</div>}
      <div className="subsec-title">Quick wins <span style={{fontSize:'11px',color:'var(--muted2)',fontWeight:'normal'}}>· rule-based</span></div>
      {sugs.map((s,i)=>(
        <div key={i} className="sug-card"><div className="sug-icon">{s.icon}</div><div className="sug-body"><div className="sug-title">{s.title}<span className={`sug-tag${s.tag==='win'?' win':s.tag==='watch'?' watch':' tip'}`}>{s.tag==='win'?'Quick win':s.tag==='watch'?'Watch':'Tip'}</span></div><div className="sug-text" dangerouslySetInnerHTML={{__html:s.text}} /></div></div>
      ))}
    </div></div>
  );
}

// ═══════════════ REPORTS ═══════════════
function Reports({monthSnapshot,curMonth,curYear,budget,expenses,flex,trips,fmt,fmtd}) {
  const q=Math.floor(curMonth/3)+1;
  const reportMonths=scope=>{
    if(scope==='month')return[{mk:`${curYear}-${String(curMonth+1).padStart(2,'0')}`,label:`${MONTHS[curMonth]} ${curYear}`}];
    if(scope==='quarter'){const arr=[];for(let m=(q-1)*3;m<(q-1)*3+3;m++)arr.push({mk:`${curYear}-${String(m+1).padStart(2,'0')}`,label:MONTHS[m]});return arr;}
    const arr=[];for(let m=0;m<12;m++)arr.push({mk:`${curYear}-${String(m+1).padStart(2,'0')}`,label:MONTHS[m]});return arr;
  };
  const buildReport=scope=>{
    const months=reportMonths(scope);
    const periodLabel=scope==='month'?`${MONTHS[curMonth]} ${curYear}`:scope==='quarter'?`Q${q} ${curYear}`:`${curYear}`;
    let agg={income:0,fixed:0,save:0,living:0,flex:0,travel:0,leftover:0};const catTotals={};
    months.forEach(mo=>{const s=monthSnapshot(mo.mk);agg.income+=s.income;agg.fixed+=s.tFixed;agg.save+=s.tSave;agg.living+=s.livingActual;agg.flex+=s.flexLog;agg.travel+=s.travel;agg.leftover+=s.leftover;(expenses[mo.mk]||[]).forEach(i=>catTotals[i.cat]=(catTotals[i.cat]||0)+i.amount);});
    const catLines=Object.entries(catTotals).sort((a,b)=>b[1]-a[1]).map(([c,v])=>`<div class="rpt-line"><span>${c}</span><b>${fmtd(v)}</b></div>`).join('')||'<div class="rpt-line"><span>No expenses logged</span><b>$0</b></div>';
    document.getElementById('print-area').innerHTML=`<div class="rpt"><div class="rpt-header"><div class="rpt-title">Financial Summary</div><div class="rpt-period">${periodLabel} · generated ${new Date().toLocaleDateString()}</div></div><div class="rpt-kpis"><div class="rpt-kpi"><div class="rpt-kpi-label">Take-Home</div><div class="rpt-kpi-val">${fmt(agg.income)}</div></div><div class="rpt-kpi"><div class="rpt-kpi-label">Fixed</div><div class="rpt-kpi-val">${fmt(agg.fixed)}</div></div><div class="rpt-kpi"><div class="rpt-kpi-label">Saved/Invested</div><div class="rpt-kpi-val">${fmt(agg.save)}</div></div><div class="rpt-kpi"><div class="rpt-kpi-label">Leftover</div><div class="rpt-kpi-val ${agg.leftover>=0?'pos':'neg'}">${fmt(agg.leftover)}</div></div></div><div class="rpt-section"><div class="rpt-h">Outflow Summary</div><div class="rpt-line"><span>Fixed + living</span><b>${fmtd(agg.fixed)}</b></div><div class="rpt-line"><span>Savings & investments</span><b>${fmtd(agg.save)}</b></div><div class="rpt-line"><span>Flex purchases</span><b>${fmtd(agg.flex)}</b></div><div class="rpt-line"><span>Travel</span><b>${fmtd(agg.travel)}</b></div><div class="rpt-line"><span>Leftover</span><b class="${agg.leftover>=0?'pos':'neg'}">${fmtd(agg.leftover)}</b></div></div><div class="rpt-section"><div class="rpt-h">Living by Category</div>${catLines}</div><div class="rpt-foot">Generated by fin. · Not financial advice.</div></div>`;
    setTimeout(()=>window.print(),100);
  };
  return (
    <div className="page"><div className="container">
      <div className="ph"><div className="pt">Reports</div><div className="psub">Download a clean PDF — choose "Save as PDF" in the print dialog</div></div>
      <div className="report-btns">
        <button className="report-card-btn" onClick={()=>buildReport('month')}><div className="rcb-icon">📅</div><div className="rcb-title">Monthly Summary</div><div className="rcb-sub">{MONTHS[curMonth]} {curYear}</div></button>
        <button className="report-card-btn" onClick={()=>buildReport('quarter')}><div className="rcb-icon">📊</div><div className="rcb-title">Quarterly Summary</div><div className="rcb-sub">Q{q} {curYear}</div></button>
        <button className="report-card-btn" onClick={()=>buildReport('year')}><div className="rcb-icon">📈</div><div className="rcb-title">Yearly Summary</div><div className="rcb-sub">{curYear}</div></button>
      </div>
      <div className="feed-note">↳ In the print dialog set "Destination" to Save as PDF, enable "Background graphics".</div>
      <div id="print-area" style={{display:'none'}} />
    </div></div>
  );
}

// ════════════ CSS ════════════
const CSS = `
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg:#f7f6f3;--surface:#ffffff;--border:#e8e6e0;--border2:#d4d0c8;
  --text:#1a1916;--muted:#9a9690;--muted2:#b8b4ac;
  --green:#2d6a4f;--red:#c1440e;--ink:#1a1916;--blue:#1d3557;--gold:#b08968;
  --fixed:#1a1916;--save:#2d6a4f;--flex:#c1440e;--trav:#1d3557;--living:#7a5c3e;
}
html{scroll-behavior:smooth;}
body{background:var(--bg);color:var(--text);font-family:'Geist Mono',monospace;font-size:13px;line-height:1.5;min-height:100vh;}
nav{position:fixed;top:0;left:0;right:0;z-index:200;background:var(--bg);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:12px;padding:0 24px;height:52px;}
.nav-brand{font-family:'Instrument Serif',serif;font-style:italic;font-size:22px;letter-spacing:-0.5px;color:var(--ink);white-space:nowrap;}
.nav-tabs{display:flex;gap:1px;overflow-x:auto;flex:1;}
.tab-btn{background:none;border:none;cursor:pointer;padding:6px 13px;font-family:'Geist Mono',monospace;font-size:11.5px;color:var(--muted);border-radius:6px;transition:all 0.15s;white-space:nowrap;}
.tab-btn.active{background:var(--text);color:var(--bg);}
.tab-btn:hover:not(.active){color:var(--text);background:var(--border);}
.save-indicator{font-size:10px;white-space:nowrap;min-width:70px;text-align:right;}
.page{padding-top:52px;}
.container{max-width:940px;margin:0 auto;padding:44px 36px;}
.ph{margin-bottom:28px;}.pt{font-family:'Instrument Serif',serif;font-size:32px;letter-spacing:-0.5px;color:var(--ink);}.psub{color:var(--muted);font-size:12px;margin-top:4px;}
.ph-flex{display:flex;align-items:baseline;justify-content:space-between;flex-wrap:wrap;gap:14px;}
.income-hero{background:var(--text);color:var(--bg);border-radius:14px;padding:26px 30px;margin-bottom:24px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:18px;}
.income-label{font-size:10px;letter-spacing:0.1em;text-transform:uppercase;opacity:0.6;margin-bottom:6px;}
.income-val{font-family:'Instrument Serif',serif;font-size:42px;letter-spacing:-1px;line-height:1;cursor:pointer;border-radius:6px;padding:2px 6px;transition:background 0.15s;}
.income-val:hover{background:rgba(255,255,255,0.1);}
.income-edit-input{font-family:'Instrument Serif',serif;font-size:42px;letter-spacing:-1px;background:rgba(255,255,255,0.15);border:none;color:#fff;border-radius:6px;padding:2px 6px;width:200px;outline:none;}
.income-breakdown{display:flex;gap:24px;}.ib-item{text-align:right;}.ib-label{font-size:10px;opacity:0.55;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px;}.ib-val{font-family:'Instrument Serif',serif;font-size:21px;letter-spacing:-0.5px;}.ib-val.pos{color:#7ee0b0;}.ib-val.neg{color:#ff9b6b;}
.flex-hero{background:linear-gradient(135deg,#2d6a4f,#1d3557);color:#fff;border-radius:14px;padding:24px 30px;margin-bottom:26px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:18px;}
.fh-label{font-size:10px;letter-spacing:0.1em;text-transform:uppercase;opacity:0.7;margin-bottom:6px;}
.fh-val{font-family:'Instrument Serif',serif;font-size:40px;letter-spacing:-1px;line-height:1;}.fh-val.neg{color:#ff9b6b;}
.fh-sub{font-size:11px;opacity:0.7;margin-top:8px;}
.fh-right{display:flex;gap:22px;}.fh-stat{text-align:right;}.fh-stat-label{font-size:9px;opacity:0.6;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:3px;}.fh-stat-val{font-family:'Instrument Serif',serif;font-size:18px;}
.alloc-bar-wrap{margin-bottom:34px;}.alloc-track{display:flex;height:14px;border-radius:8px;overflow:hidden;background:var(--border);}.alloc-seg{transition:width 0.6s cubic-bezier(0.16,1,0.3,1);}
.alloc-legend{display:flex;gap:20px;margin-top:14px;flex-wrap:wrap;}.al-item{display:flex;align-items:center;gap:8px;}.al-dot{width:10px;height:10px;border-radius:3px;}.al-text{font-size:11px;color:var(--muted);}.al-text b{color:var(--text);}
.section{margin-bottom:30px;}.section-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;border-bottom:1px solid var(--border);padding-bottom:10px;}.section-head-left{display:flex;align-items:center;gap:10px;}.section-dot{width:8px;height:8px;border-radius:2px;}.section-label{font-size:11px;letter-spacing:0.08em;text-transform:uppercase;color:var(--text);}.section-total{font-family:'Instrument Serif',serif;font-size:20px;color:var(--ink);}
.subgroup-label{font-size:10px;letter-spacing:0.08em;text-transform:uppercase;color:var(--muted2);margin:14px 0 4px;display:flex;align-items:center;gap:8px;}.subgroup-label::after{content:'';flex:1;height:1px;background:var(--border);}
.col-headers{display:grid;padding:0 0 4px;}.col-h{font-size:10px;letter-spacing:0.06em;text-transform:uppercase;color:var(--muted2);padding:0 8px;text-align:right;}.col-h:first-child{text-align:left;padding-left:0;}
.row-list{display:flex;flex-direction:column;}.budget-row{display:grid;align-items:center;border-bottom:1px solid var(--border);transition:background 0.1s;}.budget-row:last-child{border-bottom:none;}.budget-row:hover{background:rgba(0,0,0,0.015);}
.row-name{padding:12px 0;font-size:12px;display:flex;align-items:center;gap:8px;}.row-emoji{width:20px;text-align:center;}
.row-name-text{border:none;background:none;font-family:'Geist Mono',monospace;font-size:12px;color:var(--text);padding:2px 4px;border-radius:4px;border-bottom:1px dashed transparent;transition:0.15s;max-width:200px;}.row-name-text:hover{border-bottom-color:var(--muted2);}.row-name-text:focus{outline:none;border-bottom-color:var(--text);}
.row-cell{padding:8px;}.row-cell.right{text-align:right;}
.editable-val{font-family:'Geist Mono',monospace;font-size:13px;color:var(--text);background:none;border:none;width:100%;text-align:right;padding:4px 6px;border-radius:4px;transition:background 0.15s;}.editable-val:hover{background:var(--border);}.editable-val:focus{background:var(--text);color:var(--bg);outline:none;}
.actual-val{font-size:12px;text-align:right;padding:4px 6px;}.actual-val.over{color:var(--red);}.actual-val.under{color:var(--green);}
.recur-toggle{font-size:9px;padding:3px 8px;border-radius:20px;border:1px solid var(--border2);color:var(--muted);cursor:pointer;background:var(--surface);transition:0.15s;white-space:nowrap;}.recur-toggle.recurring{background:var(--text);color:var(--bg);border-color:var(--text);}
.del-row-btn{background:none;border:none;cursor:pointer;color:var(--muted2);font-size:15px;padding:2px;transition:0.15s;border-radius:4px;}.del-row-btn:hover{color:var(--red);background:rgba(193,68,14,0.08);}
.add-line-btn{margin-top:8px;background:none;border:1px dashed var(--border2);color:var(--muted);font-family:'Geist Mono',monospace;font-size:11px;padding:8px 14px;border-radius:8px;cursor:pointer;transition:0.15s;width:100%;}.add-line-btn:hover{border-color:var(--text);color:var(--text);}
.feed-note{font-size:10px;color:var(--muted2);padding:6px 0;font-style:italic;}.locked-row .editable-val{color:var(--muted);cursor:not-allowed;}.locked-tag{font-size:9px;color:var(--blue);border:1px solid var(--blue);padding:1px 6px;border-radius:10px;opacity:0.6;}
.flex-logger{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px 20px;}.fl-form{display:grid;grid-template-columns:1.5fr 110px 1fr auto;gap:10px;align-items:end;}.fl-list{margin-top:14px;}.fl-item{display:flex;justify-content:space-between;align-items:center;padding:9px 0;border-bottom:1px solid var(--border);font-size:12px;}.fl-item:last-child{border-bottom:none;}.fl-empty{font-size:11px;color:var(--muted2);padding:10px 0;font-style:italic;}.fl-travel{color:var(--trav);}
.summary-bar{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:20px 24px;display:flex;justify-content:space-between;align-items:center;margin-top:22px;flex-wrap:wrap;gap:14px;}.summary-item{text-align:center;flex:1;min-width:90px;}.summary-label{font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);margin-bottom:6px;}.summary-value{font-family:'Instrument Serif',serif;font-size:21px;letter-spacing:-0.3px;}.summary-value.red{color:var(--red);}.summary-value.green{color:var(--green);}.summary-div{width:1px;background:var(--border);align-self:stretch;}
.add-form{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px 20px;margin-bottom:22px;}.aform-grid{display:grid;gap:10px;align-items:end;}.form-group{display:flex;flex-direction:column;gap:6px;}.form-label{font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);}
.form-input{background:var(--bg);border:1px solid var(--border);padding:9px 11px;font-family:'Geist Mono',monospace;font-size:13px;color:var(--text);border-radius:7px;width:100%;transition:border-color 0.15s;}.form-input:focus{outline:none;border-color:var(--text);}.form-select{appearance:none;cursor:pointer;}
.add-btn{background:var(--text);color:var(--bg);border:none;padding:9px 20px;font-family:'Geist Mono',monospace;font-size:12px;border-radius:7px;cursor:pointer;white-space:nowrap;transition:opacity 0.15s;}.add-btn:hover{opacity:0.8;}.add-btn:disabled{opacity:0.4;cursor:not-allowed;}.btn-ghost{background:var(--surface);color:var(--text);border:1px solid var(--border);}
.month-nav{display:flex;align-items:center;gap:12px;}.month-btn{background:none;border:1px solid var(--border);padding:5px 12px;font-family:'Geist Mono',monospace;font-size:12px;color:var(--text);cursor:pointer;border-radius:6px;transition:all 0.15s;}.month-btn:hover{background:var(--text);color:var(--bg);}.month-label{font-size:14px;min-width:130px;text-align:center;}
.month-summary{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;background:var(--border);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:22px;}.msm-card{background:var(--surface);padding:18px 20px;}.msm-label{font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);margin-bottom:6px;}.msm-val{font-family:'Instrument Serif',serif;font-size:24px;letter-spacing:-0.5px;}.msm-val.red{color:var(--red);}.msm-val.green{color:var(--green);}
.item-list{display:flex;flex-direction:column;gap:1px;background:var(--border);border:1px solid var(--border);border-radius:12px;overflow:hidden;}.li-row{background:var(--surface);display:grid;align-items:center;padding:0 16px;transition:background 0.1s;}.li-row:hover{background:#fafaf8;}.li-row.header-row{background:#f0efe9;}.li-cell{padding:12px 6px;font-size:12px;}.li-cell.right{text-align:right;}.li-cell.center{text-align:center;}.li-cell.head{font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);}
.chip{display:inline-block;font-size:9px;padding:2px 8px;border-radius:20px;border:1px solid var(--border2);color:var(--muted);}.chip.ess{background:rgba(45,106,79,0.1);color:var(--green);border-color:transparent;}.chip.non{background:rgba(193,68,14,0.1);color:var(--red);border-color:transparent;}
.cycle-toggle{font-size:10px;padding:4px 10px;border-radius:6px;border:1px solid var(--border);cursor:pointer;background:var(--surface);transition:0.15s;}.cycle-toggle.annual{background:var(--blue);color:#fff;border-color:var(--blue);}
.del-btn{background:none;border:none;cursor:pointer;color:var(--muted2);font-size:14px;padding:4px;border-radius:4px;transition:all 0.15s;}.del-btn:hover{color:var(--red);background:rgba(193,68,14,0.08);}
.edit-inline{background:none;border:none;font-family:'Geist Mono',monospace;font-size:12px;color:var(--text);cursor:pointer;width:100%;padding:2px 0;border-bottom:1px dashed transparent;transition:border-color 0.15s;}.edit-inline:hover{border-bottom-color:var(--muted2);}.edit-inline:focus{outline:none;border-bottom-color:var(--text);}
.expense-item{background:var(--surface);display:grid;grid-template-columns:1.4fr 1fr 0.9fr 90px 70px 32px;align-items:center;padding:0 14px;transition:background 0.1s;}.expense-item:hover{background:#fafaf8;}.expense-item.header-row{background:#f0efe9;}.exp-cell{padding:11px 6px;font-size:12px;}.exp-cell.right{text-align:right;}.exp-cell.center{text-align:center;}.exp-cell.head{font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);}.exp-note{font-size:11px;color:var(--muted);}
.cat-breakdown{margin-top:26px;}.cat-row{display:grid;grid-template-columns:130px 1fr 110px;gap:12px;align-items:center;margin-bottom:10px;}.cat-name{font-size:11px;color:var(--muted);}.cat-bar-wrap{background:var(--border);border-radius:2px;height:5px;}.cat-bar{height:100%;border-radius:2px;background:var(--ink);transition:width 0.5s ease;}.cat-amt{font-size:12px;text-align:right;}
.notes-section{margin-top:30px;}.notes-label{font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);margin-bottom:8px;}.notes-area{width:100%;background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px;font-family:'Geist Mono',monospace;font-size:13px;color:var(--text);resize:vertical;min-height:90px;line-height:1.6;}.notes-area:focus{outline:none;border-color:var(--text);}
.empty-state{text-align:center;padding:42px;color:var(--muted2);}.empty-state-icon{font-size:30px;margin-bottom:10px;opacity:0.4;}
.cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:16px;margin-bottom:30px;}.credit-card{border-radius:16px;padding:22px;color:#fff;position:relative;overflow:hidden;min-height:160px;display:flex;flex-direction:column;justify-content:space-between;box-shadow:0 4px 20px rgba(0,0,0,0.08);}.credit-card::after{content:'';position:absolute;top:-40%;right:-20%;width:200px;height:200px;border-radius:50%;background:rgba(255,255,255,0.08);}.cc-top{display:flex;justify-content:space-between;align-items:flex-start;}.cc-name{font-size:14px;font-weight:500;}.cc-issuer{font-size:10px;opacity:0.7;margin-top:2px;}.cc-del{background:rgba(255,255,255,0.2);border:none;color:#fff;width:22px;height:22px;border-radius:50%;cursor:pointer;font-size:13px;line-height:1;transition:0.15s;z-index:2;}.cc-del:hover{background:rgba(255,255,255,0.35);}.cc-bottom{z-index:2;}.cc-spend-label{font-size:9px;opacity:0.65;text-transform:uppercase;letter-spacing:0.08em;}.cc-spend{font-family:'Instrument Serif',serif;font-size:30px;letter-spacing:-0.5px;margin-top:2px;}.cc-count{font-size:10px;opacity:0.7;margin-top:4px;}.add-card-tile{border:1px dashed var(--border2);border-radius:16px;min-height:160px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;cursor:pointer;color:var(--muted);transition:0.15s;}.add-card-tile:hover{border-color:var(--text);color:var(--text);}.act-icon{font-size:28px;font-weight:300;}.color-pick{display:flex;gap:6px;}.color-dot{width:26px;height:26px;border-radius:7px;cursor:pointer;border:2px solid transparent;transition:0.15s;}.color-dot.sel{border-color:var(--text);transform:scale(1.1);}
.cd-title{font-family:'Instrument Serif',serif;font-size:22px;margin-bottom:16px;}.cd-bars{display:flex;flex-direction:column;gap:12px;}.cd-bar-row{display:grid;grid-template-columns:160px 1fr 110px;gap:14px;align-items:center;}.cd-bar-name{font-size:12px;display:flex;align-items:center;gap:8px;}.cd-bar-track{background:var(--border);border-radius:3px;height:8px;overflow:hidden;}.cd-bar-fill{height:100%;border-radius:3px;transition:width 0.6s ease;}.cd-bar-amt{font-size:13px;text-align:right;font-family:'Instrument Serif',serif;}
.kpi-strip{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--border);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:30px;}.kpi-c{background:var(--surface);padding:18px 20px;}.kpi-c-label{font-size:10px;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);margin-bottom:6px;}.kpi-c-val{font-family:'Instrument Serif',serif;font-size:24px;letter-spacing:-0.5px;}.kpi-c-sub{font-size:10px;color:var(--muted2);margin-top:4px;}.subsec-title{font-family:'Instrument Serif',serif;font-size:24px;margin:36px 0 16px;letter-spacing:-0.3px;}
.alloc-flex{display:grid;grid-template-columns:200px 1fr;gap:36px;align-items:center;}.ring-legend{display:flex;flex-direction:column;gap:9px;}.legend-row{display:flex;align-items:center;gap:10px;}.legend-dot{width:9px;height:9px;border-radius:2px;}.legend-name{font-size:11.5px;flex:1;}.legend-pct{font-size:11px;color:var(--muted);}.legend-amt{font-size:11.5px;min-width:56px;text-align:right;}
.proj-controls{display:flex;gap:20px;align-items:center;flex-wrap:wrap;margin-bottom:20px;background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px 20px;}.slider-group{flex:1;min-width:180px;}.slider-group label{font-size:11px;color:var(--muted);display:flex;justify-content:space-between;margin-bottom:8px;}.slider-group label b{color:var(--text);font-family:'Instrument Serif',serif;font-size:15px;}input[type=range]{width:100%;accent-color:var(--text);}
.proj-chart{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:24px;margin-bottom:24px;}.proj-bars{display:flex;align-items:flex-end;gap:10px;height:220px;padding-top:20px;}.pbar-col{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;gap:6px;}.pbar-stack{width:100%;max-width:60px;display:flex;flex-direction:column;justify-content:flex-end;height:100%;border-radius:6px 6px 0 0;overflow:hidden;transition:0.4s;position:relative;}.pbar-contrib{background:var(--border2);}.pbar-growth{background:var(--save);}.pbar-val{font-size:10px;font-family:'Instrument Serif',serif;color:var(--text);}.pbar-yr{font-size:10px;color:var(--muted);}
.goal-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px 20px;margin-bottom:12px;}.goal-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;}.goal-name{font-size:14px;font-weight:500;}.goal-meta{font-size:11px;color:var(--muted);margin-top:2px;}.goal-prog-track{background:var(--border);border-radius:4px;height:8px;overflow:hidden;margin:10px 0 8px;}.goal-prog-fill{height:100%;background:var(--save);border-radius:4px;transition:width 0.6s ease;}.goal-stats{display:flex;justify-content:space-between;font-size:11px;color:var(--muted);}.goal-stats b{color:var(--text);}.goal-eta{font-size:11px;color:var(--blue);margin-top:6px;}
.trip-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px 22px;margin-bottom:14px;}.trip-head{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;}.trip-name{font-family:'Instrument Serif',serif;font-size:20px;letter-spacing:-0.3px;}.trip-dates{font-size:11px;color:var(--muted);margin-top:2px;}.trip-status{font-size:9px;padding:3px 10px;border-radius:20px;}.trip-status.planning{background:rgba(29,53,87,0.1);color:var(--blue);}.trip-status.saving{background:rgba(176,137,104,0.15);color:var(--gold);}.trip-status.done{background:rgba(45,106,79,0.12);color:var(--green);}.trip-budget-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:14px;}.tbg-label{font-size:10px;text-transform:uppercase;letter-spacing:0.06em;color:var(--muted);margin-bottom:4px;}.tbg-val{font-family:'Instrument Serif',serif;font-size:20px;}.tbg-val.over{color:var(--red);}.tbg-val.under{color:var(--green);}.trip-prog-track{background:var(--border);border-radius:4px;height:7px;overflow:hidden;margin-bottom:6px;}.trip-prog-fill{height:100%;border-radius:4px;transition:width 0.6s ease;}.trip-line{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--border);font-size:12px;}.trip-line:last-child{border-bottom:none;}.trip-expense-form{display:grid;grid-template-columns:1.5fr 1fr 100px auto;gap:8px;margin-top:12px;align-items:end;padding-top:14px;border-top:1px solid var(--border);}.trip-month-note{font-size:10px;color:var(--muted2);margin-top:4px;}
.insight-box{background:linear-gradient(135deg,rgba(45,106,79,0.06),rgba(29,53,87,0.04));border:1px solid var(--border);border-radius:12px;padding:18px 22px;margin-bottom:24px;}.insight-title{font-size:11px;text-transform:uppercase;letter-spacing:0.08em;color:var(--save);margin-bottom:10px;}.insight-text{font-size:12.5px;line-height:1.7;}.insight-text b{color:var(--ink);}
.chart-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:22px;margin-bottom:8px;}.bar-h-row{display:grid;grid-template-columns:130px 1fr 96px;gap:12px;align-items:center;margin-bottom:11px;}.bar-h-label{font-size:11px;color:var(--muted);}.bar-h-track{background:var(--border);border-radius:3px;height:18px;position:relative;overflow:hidden;}.bar-h-fill{height:100%;border-radius:3px;transition:width 0.6s ease;}.bar-h-amt{font-size:12px;text-align:right;}.bva-row{display:grid;grid-template-columns:120px 1fr;gap:12px;align-items:center;margin-bottom:14px;}.bva-bars{display:flex;flex-direction:column;gap:4px;}.bva-bar{height:12px;border-radius:3px;display:flex;align-items:center;padding-left:6px;font-size:9px;color:#fff;min-width:fit-content;transition:width 0.6s ease;white-space:nowrap;}.trend-wrap{display:flex;align-items:flex-end;gap:12px;height:200px;padding-top:16px;}.trend-col{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;gap:6px;}.trend-bar{width:100%;max-width:50px;background:var(--flex);border-radius:5px 5px 0 0;transition:height 0.5s ease;min-height:2px;}.trend-val{font-size:10px;font-family:'Instrument Serif',serif;}.trend-mo{font-size:10px;color:var(--muted);}.flow-wrap{display:flex;flex-direction:column;gap:10px;}.flow-bar{display:flex;align-items:center;height:42px;border-radius:8px;padding:0 16px;color:#fff;font-size:12px;justify-content:space-between;transition:width 0.6s ease;}
.sug-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px 20px;margin-bottom:10px;display:flex;gap:14px;align-items:flex-start;}.sug-icon{font-size:18px;flex-shrink:0;margin-top:1px;}.sug-body{flex:1;}.sug-title{font-size:13px;font-weight:500;margin-bottom:3px;}.sug-text{font-size:12px;color:var(--muted);line-height:1.6;}.sug-text b{color:var(--text);}.sug-tag{font-size:9px;padding:2px 8px;border-radius:20px;margin-left:8px;}.sug-tag.win{background:rgba(45,106,79,0.12);color:var(--green);}.sug-tag.watch{background:rgba(193,68,14,0.12);color:var(--red);}.sug-tag.tip{background:rgba(29,53,87,0.1);color:var(--blue);}
.ai-card{background:linear-gradient(135deg,rgba(90,24,154,0.05),rgba(29,53,87,0.04));border:1px solid var(--border);border-radius:12px;padding:22px;margin-bottom:20px;font-size:13px;line-height:1.75;}.ai-loading{display:flex;align-items:center;gap:10px;color:var(--muted);font-size:12px;padding:20px;}.spinner{width:16px;height:16px;border:2px solid var(--border);border-top-color:var(--text);border-radius:50%;animation:spin 0.7s linear infinite;}@keyframes spin{to{transform:rotate(360deg);}}
.report-btns{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:18px;}.report-card-btn{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:26px 20px;cursor:pointer;text-align:center;transition:0.15s;font-family:'Geist Mono',monospace;}.report-card-btn:hover{border-color:var(--text);transform:translateY(-2px);}.rcb-icon{font-size:28px;margin-bottom:10px;}.rcb-title{font-size:14px;font-weight:500;color:var(--text);margin-bottom:4px;}.rcb-sub{font-size:11px;color:var(--muted);}
.toast{position:fixed;bottom:24px;right:24px;z-index:999;background:var(--text);color:var(--bg);padding:10px 18px;border-radius:8px;font-size:12px;opacity:0;transform:translateY(8px);transition:all 0.2s;pointer-events:none;}.toast.show{opacity:1;transform:translateY(0);}
@media(max-width:780px){
  nav{padding:0 14px;gap:8px;}.container{padding:30px 16px;}
  .income-hero,.flex-hero{flex-direction:column;align-items:flex-start;}
  .income-breakdown,.fh-right{width:100%;justify-content:space-between;gap:12px;}
  .kpi-strip{grid-template-columns:1fr 1fr;}
  .alloc-flex{grid-template-columns:1fr;justify-items:center;}
  .expense-item{grid-template-columns:1.2fr 0.8fr 60px 28px;}
  .hide-mobile{display:none;}
  .trip-budget-grid{grid-template-columns:1fr;}
  .trip-expense-form,.fl-form{grid-template-columns:1fr 1fr;}
  .report-btns{grid-template-columns:1fr;}
  .aform-grid{grid-template-columns:1fr!important;}
}
@media print{
  nav,.page,.toast{display:none!important;}
  #print-area{display:block!important;padding:0;}
  body{background:#fff;}
  .rpt{font-family:'Geist Mono',monospace;color:#1a1916;max-width:720px;margin:0 auto;}
  .rpt-header{background:#1a1916;color:#fff;padding:28px 32px;margin-bottom:24px;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  .rpt-title{font-family:'Instrument Serif',serif;font-size:30px;}.rpt-period{font-size:12px;opacity:0.8;margin-top:4px;}
  .rpt-section{margin:0 32px 22px;page-break-inside:avoid;}
  .rpt-h{font-size:11px;text-transform:uppercase;letter-spacing:0.08em;color:#9a9690;border-bottom:1px solid #e8e6e0;padding-bottom:6px;margin-bottom:10px;}
  .rpt-line{display:flex;justify-content:space-between;padding:5px 0;font-size:12px;border-bottom:1px solid #f0efe9;}.rpt-line b{font-weight:500;}
  .rpt-kpis{display:flex;margin:0 32px 24px;border:1px solid #e8e6e0;border-radius:8px;overflow:hidden;}
  .rpt-kpi{flex:1;padding:14px 16px;border-right:1px solid #e8e6e0;}.rpt-kpi:last-child{border-right:none;}
  .rpt-kpi-label{font-size:9px;text-transform:uppercase;color:#9a9690;}.rpt-kpi-val{font-family:'Instrument Serif',serif;font-size:20px;margin-top:3px;}
  .rpt-foot{margin:24px 32px 0;font-size:10px;color:#b8b4ac;border-top:1px solid #e8e6e0;padding-top:12px;}
  .pos{color:#2d6a4f;}.neg{color:#c1440e;}
}
`;
