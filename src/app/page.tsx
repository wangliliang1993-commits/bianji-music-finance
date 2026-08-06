"use client";

import { createContext, useContext, useEffect, useMemo, useState } from "react";
import {
  ArrowDownLeft, ArrowUpRight, BarChart3, Bell, BookOpen, Boxes, ChevronDown,
  CircleHelp, FileDown, FileText, Landmark, LayoutDashboard, Menu, MoreHorizontal,
  Package, Plus, Search, Settings, ShoppingBag, TrendingUp, Users, Wallet,
  X, ArrowRightLeft, CheckCircle2, Target, Radio, Save, CalendarClock, Trash2
} from "lucide-react";
import { Area, AreaChart, CartesianGrid, Cell, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { formatMoney } from "@/lib/money";
import * as XLSX from "xlsx";

type NavKey = "dashboard" | "reports" | "budget" | "finance" | "course_flow" | "live_flow" | "sales" | "inventory" | "fixed_expense" | "settings";

const trend = [
  { month: "2月", income: 0, expense: 0 }, { month: "3月", income: 0, expense: 0 },
  { month: "4月", income: 0, expense: 0 }, { month: "5月", income: 0, expense: 0 },
  { month: "6月", income: 0, expense: 0 }, { month: "7月", income: 0, expense: 0 },
  { month: "8月", income: 0, expense: 0 }
];

const courseData = [
  { name: "架子鼓", value: 0, amount: "¥0", color: "#7568e8" },
  { name: "吉他", value: 0, amount: "¥0", color: "#f3b559" },
  { name: "直播课", value: 0, amount: "¥0", color: "#48bba5" }
];

const transactions: Array<{ id:string; title:string; meta:string; account:string; amount:number; time:string; type:string }> = [];

const products: Array<{sku:string;name:string;category:string;stock:number;unit:string;cost:number;price:number;status:string}> = [];

const nav = [
  { key: "dashboard", label: "经营概览", icon: LayoutDashboard },
  { key: "reports", label: "经营报表", icon: BarChart3 },
  { key: "budget", label: "预算管理", icon: Target },
  { key: "finance", label: "资金流水", icon: Wallet },
  { key: "course_flow", label: "课费流水", icon: BookOpen },
  { key: "live_flow", label: "直播流水", icon: Radio },
  { key: "sales", label: "销售开单", icon: ShoppingBag },
  { key: "fixed_expense", label: "固定支出", icon: CalendarClock }
] as const;

const monthOptions = ["2026年8月", "2026年7月", "2026年6月", "2026年5月", "2026年4月", "2026年3月", "2026年2月", "2026年1月"];
const monthlyStats: Record<string, { income: number; expense: number }> = {
  "2026年8月": { income: 0, expense: 0 }, "2026年7月": { income: 0, expense: 0 },
  "2026年6月": { income: 0, expense: 0 }, "2026年5月": { income: 0, expense: 0 },
  "2026年4月": { income: 0, expense: 0 }, "2026年3月": { income: 0, expense: 0 },
  "2026年2月": { income: 0, expense: 0 }, "2026年1月": { income: 0, expense: 0 }
};

type LinkedEntry = { id:string; title:string; category:string; amountFen:number; account:string; date:string; source:NavKey };
type FixedItem = { name:string; category:string; cycle:string; amount:number; next:string; status:string; account?:string; note?:string };
type AppData = {
  entries: LinkedEntry[];
  addEntry: (entry: LinkedEntry) => void;
  fixedItems: FixedItem[];
  setFixedItems: React.Dispatch<React.SetStateAction<FixedItem[]>>;
  stockDeltas: Record<string,number>;
  adjustStock: (sku:string,delta:number)=>void;
  lastDate: string;
  setLastDate: (date:string)=>void;
  lastFixedTemplate: FixedItem | null;
  setLastFixedTemplate: (item:FixedItem)=>void;
  currentMonth: string;
  setCurrentMonth: (month:string)=>void;
};
const DataContext = createContext<AppData | null>(null);
const useData = () => { const value=useContext(DataContext); if(!value) throw new Error("DataContext missing"); return value; };
const monthKeyFromDate = (date:string) => { const [year,month]=date.split("-"); return `${year}年${Number(month)}月`; };
const formatFlowDate = (value:string, month:string) => /^\d{4}-\d{2}-\d{2}$/.test(value) ? value : `${month.replace("年","-").replace("月","")}-${value.includes("昨天") ? "04" : "05"}`;

function useLinkedStats(month:string) {
  const {entries,fixedItems}=useData();
  const base=monthlyStats[month] ?? {income:0,expense:0};
  const additions=entries.filter(item=>monthKeyFromDate(item.date)===month).reduce((sum,item)=>item.amountFen>=0?{...sum,income:sum.income+item.amountFen}:{...sum,expense:sum.expense+Math.abs(item.amountFen)},{income:0,expense:0});
  const fixed=fixedItems.filter(item=>item.status==="已支付"&&monthKeyFromDate(item.cycle)===month).reduce((sum,item)=>sum+item.amount*100,0);
  return {income:base.income+additions.income,expense:base.expense+additions.expense+fixed};
}

function MonthSelect({ value, onChange, compact = false }: { value: string; onChange: (month: string) => void; compact?: boolean }) {
  const match = value.match(/(\d{4})年(\d+)月/);
  const year = match?.[1] ?? "2026";
  const month = match?.[2] ?? "8";
  const update = (nextYear: string, nextMonth: string) => onChange(`${nextYear}年${Number(nextMonth)}月`);
  return <div className={compact ? "date-selects compact" : "date-selects"}><label className="month-select year"><select value={year} onChange={e=>update(e.target.value,month)} aria-label="选择年份">{["2026","2025","2024","2023","2022"].map(x=><option key={x}>{x}</option>)}</select><ChevronDown size={14}/></label><label className="month-select month"><select value={month} onChange={e=>update(year,e.target.value)} aria-label="选择月份">{Array.from({length:12},(_,i)=>String(i+1)).map(x=><option key={x} value={x}>{x}月</option>)}</select><ChevronDown size={14}/></label></div>;
}

function exportExcel(fileName: string, rows: Record<string, string | number>[]) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  sheet["!cols"] = Object.keys(rows[0] ?? {}).map(key => ({ wch: Math.max(12, key.length * 2 + 4) }));
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "边际音乐");
  XLSX.writeFile(book, `${fileName}.xlsx`);
}

function MetricCard({ title, value, note, icon: Icon, tone }: { title: string; value: string; note: string; icon: typeof Wallet; tone: string }) {
  return <div className="metric-card">
    <div className={`metric-icon ${tone}`}><Icon size={19}/></div>
    <div className="metric-copy"><span>{title}</span><strong>{value}</strong><small>{note}</small></div>
    <button className="ghost"><MoreHorizontal size={18}/></button>
  </div>;
}

function Dashboard({ onCreate }: { onCreate: () => void }) {
  const {entries,currentMonth:month,setCurrentMonth:setMonth}=useData();
  const stats = useLinkedStats(month);
  const monthEntries=entries.filter(item=>monthKeyFromDate(item.date)===month);
  const courseIncome=monthEntries.filter(item=>item.category.includes("课时费")).reduce((sum,item)=>sum+Math.max(0,item.amountFen),0);
  const salesIncome=monthEntries.filter(item=>item.category.includes("销售")||item.category.includes("直播商品")).reduce((sum,item)=>sum+Math.max(0,item.amountFen),0);
  return <>
    <div className="page-heading">
      <div><p className="eyebrow">2026 年 8 月 5 日 · 星期三</p><h1>早上好，王利凉</h1><p>这是边际音乐本月最新的经营情况。</p></div>
      <div className="heading-actions"><MonthSelect value={month} onChange={setMonth}/><button className="button secondary" onClick={()=>exportExcel(`边际音乐-${month}-经营概览`,[{月份:month,收入:stats.income/100,支出:stats.expense/100,现金结余:(stats.income-stats.expense)/100,课时费收入:164130,销售收入:122290}])}><FileDown size={17}/> 导出报表</button><button className="button primary" onClick={onCreate}><Plus size={18}/> 记一笔</button></div>
    </div>
    <div className="notice"><div><CheckCircle2 size={17}/><span>7 月账目已核对完成</span><small>所有账户余额与流水一致</small></div><button>查看月报 <ArrowUpRight size={14}/></button></div>
    <section className="metric-grid">
      <MetricCard title="现金结余" value={formatMoney(stats.income-stats.expense)} note={`结余率 ${Math.round((stats.income-stats.expense)/stats.income*100)}%`} icon={Wallet} tone="green"/>
      <MetricCard title="本月收入" value={formatMoney(stats.income)} note="较上月 +12.8%" icon={ArrowDownLeft} tone="purple"/>
      <MetricCard title="本月支出" value={formatMoney(stats.expense)} note="较上月 +3.2%" icon={ArrowUpRight} tone="orange"/>
      <MetricCard title="课时费收入" value={formatMoney(courseIncome)} note={courseIncome?"已同步课程流水":"暂无数据"} icon={BookOpen} tone="pink"/>
      <MetricCard title="销售收入" value={formatMoney(salesIncome)} note={salesIncome?"已同步销售流水":"暂无数据"} icon={ShoppingBag} tone="purple"/>
    </section>
    <section className="dashboard-grid">
      <article className="panel chart-panel">
        <div className="panel-title"><div><h2>收支趋势</h2><p>近 7 个月现金收支变化</p></div><MonthSelect value={month} onChange={setMonth} compact/></div>
        <div className="legend"><span><i className="income-dot"/>收入</span><span><i className="expense-dot"/>支出</span></div>
        <div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><AreaChart data={trend} margin={{ left: -22, right: 8, top: 12 }}><defs><linearGradient id="income" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stopColor="#7568e8" stopOpacity={.28}/><stop offset="1" stopColor="#7568e8" stopOpacity={0}/></linearGradient></defs><CartesianGrid vertical={false} stroke="#eceaf3" strokeDasharray="3 4"/><XAxis dataKey="month" axisLine={false} tickLine={false}/><YAxis axisLine={false} tickLine={false} tickFormatter={v => `${v}万`}/><Tooltip formatter={(v) => [`¥${v}万`]}/><Area type="monotone" dataKey="income" stroke="#7568e8" strokeWidth={2.5} fill="url(#income)"/><Area type="monotone" dataKey="expense" stroke="#f0a656" strokeWidth={2} fill="transparent" strokeDasharray="5 5"/></AreaChart></ResponsiveContainer></div>
      </article>
      <article className="panel course-panel">
        <div className="panel-title"><div><h2>课程收入构成</h2><p>本月各课程类别</p></div><button className="ghost"><MoreHorizontal size={18}/></button></div>
        <div className="donut-row"><div className="donut"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={courseData} dataKey="value" innerRadius={58} outerRadius={76} paddingAngle={3}>{courseData.map((x) => <Cell key={x.name} fill={x.color}/>)}</Pie></PieChart></ResponsiveContainer><div className="donut-label"><b>¥0.00</b><span>课时费收入</span></div></div>
          <div className="course-list">{courseData.map(x => <div key={x.name}><span><i style={{background:x.color}}/>{x.name}</span><b>{x.amount}</b><small>{x.value}%</small></div>)}</div></div>
      </article>
    </section>
    <section className="bottom-grid">
      <article className="panel transactions-panel"><div className="panel-title"><div><h2>近期流水</h2><p>最近发生的资金变动</p></div><button className="link-button">查看全部 <ArrowUpRight size={14}/></button></div><div className="transaction-list">{transactions.map(t => <div className="transaction" key={t.id}><div className={`tx-icon ${t.type}`}><span>{t.type === "income" ? "收" : t.type === "sale" ? "售" : "支"}</span></div><div className="tx-main"><b>{t.title}</b><span>{t.meta}</span></div><div className="tx-account"><span>{t.account}</span><small>{t.time}</small></div><strong className={t.amount > 0 ? "positive" : "negative"}>{t.amount > 0 ? "+" : "−"}{formatMoney(Math.abs(t.amount))}</strong></div>)}</div></article>
      <article className="panel store-panel"><div className="panel-title"><div><h2>账户余额</h2><p>当前资金分布</p></div><button className="ghost"><MoreHorizontal size={18}/></button></div><div className="store-card"><div className="store-head"><div className="store-avatar purple">微</div><div><b>微信商户</b><span>暂无流水</span></div><strong>¥0.00</strong></div><div className="progress"><i style={{width:"0%"}}/></div></div><div className="store-card"><div className="store-head"><div className="store-avatar orange">银</div><div><b>银行账户</b><span>暂无流水</span></div><strong>¥0.00</strong></div><div className="progress orange"><i style={{width:"0%"}}/></div></div><button className="store-report">查看资金账户 <ArrowUpRight size={14}/></button></article>
    </section>
  </>;
}

function BudgetPage({ month, setMonth }: { month: string; setMonth: (value: string) => void }) {
  const handleImport = (_file: File) => {};
  const [selected, setSelected] = useState<string | null>(null);
  const {fixedItems}=useData();
  const paidByCategory=(category:string)=>fixedItems.filter(item=>item.status==="已支付"&&item.category===category).reduce((sum,item)=>sum+item.amount,0);
  const budgets = [
    { name: "场租水电", used: paidByCategory("场租水电"), total: 0, color: "#7568e8" },
    { name: "人员工资", used: paidByCategory("人员工资"), total: 0, color: "#48bba5" },
    { name: "直播投流", used: 0, total: 0, color: "#f3b559" },
    { name: "日常经营", used: 0, total: 0, color: "#ed8095" }
  ];
  const budgetTotal=budgets.reduce((sum,item)=>sum+item.total,0);const usedTotal=budgets.reduce((sum,item)=>sum+item.used,0);
  return <><div className="page-heading"><div><p className="eyebrow">资金计划</p><h1>预算管理</h1><p>设置月度预算，随时掌握剩余额度。</p></div><div className="heading-actions"><MonthSelect value={month} onChange={setMonth}/><button className="button secondary" onClick={()=>exportExcel(`边际音乐-${month}-预算`,budgets.map(x=>({预算项目:x.name,预算金额:x.total,已使用:x.used,剩余:x.total-x.used})))}><FileDown size={16}/> 导出</button><label className="button secondary import-button">导入<input type="file" accept=".xlsx,.xls,.csv" hidden onChange={e=>e.target.files?.[0]&&handleImport(e.target.files[0])}/></label><button className="button primary" onClick={()=>setSelected("新增预算")}><Plus size={18}/> 新增预算</button></div></div><section className="metric-grid report"><MetricCard title="预算总额" value={formatMoney(budgetTotal*100)} note={month} icon={Target} tone="purple"/><MetricCard title="已经使用" value={formatMoney(usedTotal*100)} note={`使用率 ${budgetTotal?Math.round(usedTotal/budgetTotal*100):0}%`} icon={ArrowUpRight} tone="orange"/><MetricCard title="剩余额度" value={formatMoney((budgetTotal-usedTotal)*100)} note="暂无预算数据" icon={Wallet} tone="green"/></section><div className="budget-grid">{budgets.map(item => { const rate = item.total?Math.round(item.used/item.total*100):0; return <button className="panel budget-card" key={item.name} onClick={()=>setSelected(item.name)}><div><span>{item.name}</span><b>{formatMoney(item.used*100)} <small>/ {formatMoney(item.total*100)}</small></b></div><strong>{rate}%</strong><div className="budget-progress"><i style={{width:`${rate}%`,background:item.color}}/></div><small>剩余 {formatMoney((item.total-item.used)*100)}</small><ArrowUpRight className="budget-open" size={16}/></button>})}</div><BudgetModal name={selected} month={month} onClose={()=>setSelected(null)}/></>;
}

function BudgetModal({ name, month, onClose }: { name: string | null; month: string; onClose: () => void }) {
  if (!name) return null;
  const creating = name === "新增预算";
  return <div className="modal-backdrop" onMouseDown={e=>e.target===e.currentTarget&&onClose()}><div className="modal settings-modal"><div className="modal-title"><div><h2>{creating ? "新增预算" : `${name}预算明细`}</h2><p>{month} · 设置额度并查看使用记录</p></div><button className="close" onClick={onClose}><X size={20}/></button></div><div className="settings-body"><label>预算项目<select defaultValue={creating ? "场地与物业" : name}><option>场地与物业</option><option>人员工资</option><option>直播投流</option><option>日常经营</option></select></label><div className="form-grid"><label>预算金额（元）<input type="number" min="0" step="0.01" defaultValue={creating ? "" : "22000.00"} placeholder="请输入金额"/></label><label>预警比例<select><option>80%</option><option>90%</option><option>100%</option></select></label></div><label>预算说明<textarea placeholder="填写预算用途及说明"/></label>{!creating&&<div className="budget-detail"><b>本月使用记录</b><span>08-01 场地租金 <strong>−¥18,500.00</strong></span></div>}<div className="modal-actions"><button className="button secondary" onClick={onClose}>取消</button><button className="button primary" onClick={onClose}><Save size={16}/>保存预算</button></div></div></div></div>;
}

function FixedExpensePage({ month, setMonth }: { month: string; setMonth: (value: string) => void }) {
  const [editing,setEditing] = useState<string|null>(null);
  const {fixedItems:items,setFixedItems:setItems}=useData();
  const total = items.reduce((sum,item)=>sum+item.amount,0);
  const paidTotal = items.filter(item=>item.status==="已支付").reduce((sum,item)=>sum+item.amount,0);
  const paidCount = items.filter(item=>item.status==="已支付").length;
  const categoryColors: Record<string,string> = { "人员工资":"#7568e8", "日常支出":"#48bba5", "社保公积金":"#f3b559", "场租水电":"#ed8095" };
  const categoryStats = Object.keys(categoryColors).map(category=>({category,amount:items.filter(item=>item.category===category).reduce((sum,item)=>sum+item.amount,0),color:categoryColors[category]}));
  return <><div className="page-heading"><div><p className="eyebrow">周期费用</p><h1>固定支出</h1><p>统一管理每月重复发生的经营支出。</p></div><div className="heading-actions"><MonthSelect value={month} onChange={setMonth}/><button className="button secondary" onClick={()=>exportExcel(`边际音乐-${month}-固定支出`,items.map(x=>({支出项目:x.name,分类:x.category,到期日:x.cycle,金额:x.amount,下次支付:x.next,状态:x.status})))}><FileDown size={16}/>导出</button><button className="button primary" onClick={()=>setEditing("新增固定支出")}><Plus size={18}/>新增固定支出</button></div></div><div className="summary-strip fixed-summary"><div><span>固定支出项目</span><b>{items.length}</b></div><div><span>{month}预计支出</span><b>{formatMoney(total*100)}</b></div><div><span>已支付</span><b>{paidCount}</b></div><div><span>待支付</span><b className="warn">{items.length-paidCount}</b></div></div><article className="panel table-panel"><div className="data-table"><div className="table-row fixed-row header"><span>支出项目</span><span>费用分类</span><span>到期日</span><span>金额</span><span>下次支付</span><span>状态</span><span>操作</span></div>{items.map(item=><div className="table-row fixed-row clickable-row" key={item.name} onClick={()=>setEditing(item.name)}><span><b>{item.name}</b></span><span>{item.category}</span><span>{item.cycle}</span><strong>{formatMoney(item.amount*100)}</strong><span>{item.next}</span><span><em className={item.status==="已支付"?"status ok":"status low"}>{item.status}</em></span><button className="delete-row" title={`删除${item.name}`} onClick={event=>{event.stopPropagation();setItems(current=>current.filter(x=>x.name!==item.name))}}><Trash2 size={15}/>删除</button></div>)}</div></article><section className="panel category-stats"><div className="panel-title"><div><h2>分类统计</h2><p>{month}固定支出分类构成</p></div><strong>{formatMoney(total*100)}</strong></div><div className="category-stat-grid">{categoryStats.map(item=>{const rate=total?Math.round(item.amount/total*100):0;return <div key={item.category}><span><i style={{background:item.color}}/>{item.category}</span><b>{formatMoney(item.amount*100)}</b><div><i style={{width:`${rate}%`,background:item.color}}/></div><small>{rate}%</small></div>})}</div></section><div className="amount-summary fixed-total"><div><span>{month}固定支出总计</span><b>{formatMoney(total*100)}</b></div><div><span>已支付金额</span><b>{formatMoney(paidTotal*100)}</b></div><div className="net"><span>待支付金额</span><b>{formatMoney((total-paidTotal)*100)}</b></div></div><FixedExpenseModal name={editing} item={items.find(item=>item.name===editing)??null} onClose={()=>setEditing(null)} onSave={item=>{setItems(current=>{const exists=current.some(x=>x.name===editing&&editing!=="新增固定支出");return exists?current.map(x=>x.name===editing?item:x):[...current,item]});setEditing(null)}}/></>;
}

function FixedExpenseModal({ name, item, onClose, onSave }: { name: string|null; item:FixedItem|null; onClose:()=>void; onSave:(item:FixedItem)=>void }) {
  const {lastDate,setLastDate,lastFixedTemplate,setLastFixedTemplate}=useData();
  if(!name) return null;
  const creating=name==="新增固定支出";
  const defaults=item??lastFixedTemplate;
  return <div className="modal-backdrop" onMouseDown={e=>e.target===e.currentTarget&&onClose()}><div className="modal settings-modal"><div className="modal-title"><div><h2>{creating?"新增固定支出":`编辑${name}`}</h2><p>设置到期日与金额，跟踪支付状态</p></div><button className="close" onClick={onClose}><X size={20}/></button></div><form onSubmit={e=>{e.preventDefault();const data=new FormData(e.currentTarget);const due=String(data.get("due"));const nextItem={name:String(data.get("name")),category:String(data.get("category")),cycle:due,amount:Number(data.get("amount")),next:due,status:String(data.get("status")),account:String(data.get("account")),note:String(data.get("note"))};setLastDate(due);setLastFixedTemplate(nextItem);onSave(nextItem)}}><label>支出项目<input name="name" required defaultValue={defaults?.name??""} placeholder="手动输入支出项目"/></label><div className="form-grid"><label>费用分类<select name="category" defaultValue={defaults?.category??"日常支出"}><option>人员工资</option><option>日常支出</option><option>社保公积金</option><option>场租水电</option></select></label><label>到期日<input name="due" required type="date" defaultValue={defaults?.cycle??lastDate} onChange={e=>setLastDate(e.target.value)}/></label><label>固定金额（元）<input name="amount" required min="0" step="0.01" type="number" defaultValue={defaults?defaults.amount.toFixed(2):""} placeholder="请输入金额"/></label><label>资金账户<select name="account" defaultValue={defaults?.account??"银行账户"}><option>银行账户</option><option>微信商户</option><option>现金账户</option></select></label></div><label>支付状态<select name="status" defaultValue={defaults?.status??"待支付"}><option>待支付</option><option>已支付</option></select></label><label>备注（选填）<textarea name="note" defaultValue={defaults?.note??""} placeholder="填写合同周期、付款说明等"/></label><div className="modal-actions"><button type="button" className="button secondary" onClick={onClose}>取消</button><button type="submit" className="button primary"><Save size={16}/>保存固定支出</button></div></form></div></div>;
}

function DataPage({ page, onCreate }: { page: Exclude<NavKey, "dashboard">; onCreate: () => void }) {
  const {entries,addEntry,stockDeltas,currentMonth:month,setCurrentMonth:setMonth}=useData();
  const configs: Record<Exclude<NavKey, "dashboard">, [string, string, string]> = {
    budget: ["预算管理", "设置每月经营预算并跟踪执行进度", "新增预算"],
    finance: ["资金流水", "管理琴行的每一笔收支和账户转账", "新增流水"],
    course_flow: ["课费流水", "区分续费与新报收入，按次核销课程卡", "新增课费流水"],
    live_flow: ["直播流水", "单独记录直播课收入和直播经营支出", "新增直播流水"],
    sales: ["销售开单", "记录商品销售、优惠、收款与销售成本", "新建销售单"],
    inventory: ["商品库存", "查看实时库存、成本与预警", "入库 / 出库"],
    fixed_expense: ["固定支出", "管理每月自动发生的周期性经营费用", "新增固定支出"],
    reports: ["经营报表", "从现金、课时费和商品视角分析经营表现", "导出报表"],
    settings: ["系统设置", "管理账户、课程分类和团队成员", "邀请成员"]
  };
  const config = configs[page];
  const [query, setQuery] = useState("");
  const handleImport = async (file: File) => {
    const workbook = XLSX.read(await file.arrayBuffer(), {type:"array"});
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[workbook.SheetNames[0]]);
    rows.forEach((row,index) => {
      const amount = Number(row["金额"] ?? row["金额（元）"] ?? 0);
      if (!amount) return;
      addEntry({id:`IMP-${Date.now()}-${index}`, date:String(row["日期"] ?? row["记账日期"] ?? new Date().toISOString().slice(0,10)), amountFen:Math.round(amount*100), account:String(row["资金账户"] ?? "现金账户"), category:String(row["收支分类"] ?? row["收入类别"] ?? "日常支出"), title:String(row["摘要"] ?? row["流水信息"] ?? "导入流水"), source:page});
    });
  };
  const stats = useLinkedStats(month);
  const linkedProducts = products.map(product=>({...product,stock:product.stock+(stockDeltas[product.sku]??0)}));
  const filtered = useMemo(() => linkedProducts.filter(x => `${x.name}${x.sku}${x.category}`.toLowerCase().includes(query.toLowerCase())), [query,stockDeltas]);
  if (page === "budget") return <BudgetPage month={month} setMonth={setMonth}/>;
  if (page === "reports") return <Reports/>;
  if (page === "fixed_expense") return <FixedExpensePage month={month} setMonth={setMonth}/>;
  if (page === "settings") return <SettingsPage/>;
  const inventory = page === "inventory";
  const liveFlow = page === "live_flow";
  const allTransactions = [...transactions,...entries.map(item=>({id:item.id,title:item.title,meta:item.category,account:item.account,amount:item.amountFen,time:item.date,type:item.amountFen>=0?"income":"expense"}))];
  const visibleTransactions = page === "course_flow" ? allTransactions.filter(t => t.meta.includes("续费") || t.meta.includes("新报")) : page === "live_flow" ? allTransactions.filter(t => t.meta.includes("直播")) : page === "sales" ? allTransactions.filter(t=>t.type==="sale"||t.meta.includes("销售")) : allTransactions;
  const renewalItems=visibleTransactions.filter(t=>t.meta.includes("续费"));
  const newReportItems=visibleTransactions.filter(t=>t.meta.includes("新报"));
  const renewalCount=renewalItems.length;
  const renewalAmount=renewalItems.reduce((sum,t)=>sum+t.amount,0);
  const newReportCount=newReportItems.length;
  const newReportAmount=newReportItems.reduce((sum,t)=>sum+t.amount,0);
  const courseTotal=renewalAmount+newReportAmount;
  const liveIncome=visibleTransactions.filter(t=>t.amount>0).reduce((sum,t)=>sum+t.amount,0);
  const inventoryValue=linkedProducts.reduce((sum,p)=>sum+p.stock*p.cost,0);
  const exportRows = inventory ? linkedProducts.map(p=>({商品名称:p.name,SKU:p.sku,商品类别:p.category,库存:`${p.stock}${p.unit}`,零售价:p.price/100,状态:p.status})) : visibleTransactions.map(t=>({商品名称:t.title,订单号:t.id,金额:t.amount/100,资金账户:t.account,发生时间:t.time,分类:t.meta}));
  return <>
    <div className="page-heading"><div><p className="eyebrow">经营管理</p><h1>{config[0]}</h1><p>{config[1]}</p></div><div className="heading-actions"><MonthSelect value={month} onChange={setMonth}/><button className="button secondary" onClick={()=>exportExcel(`边际音乐-${month}-${config[0]}`,exportRows)}><FileDown size={16}/> 导出</button><button className="button primary" onClick={onCreate}><Plus size={18}/>{config[2]}</button></div></div>
    <div className={liveFlow ? "summary-strip live-summary" : "summary-strip"}><div><span>{month}续费人数</span><b>{renewalCount} 人</b><small>{formatMoney(renewalAmount)}</small></div><div><span>{month}新报人数</span><b>{newReportCount} 人</b><small>{formatMoney(newReportAmount)}</small></div></div>
    <article className="panel table-panel"><div className="toolbar"><div className="search-field"><Search size={17}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder={inventory ? "搜索商品名称、SKU 或类别" : "搜索摘要、单号或账户"}/></div>{page!=="course_flow"&&<><MonthSelect value={month} onChange={setMonth}/><button className="button secondary" onClick={()=>exportExcel(`边际音乐-${month}-${config[0]}`,exportRows)}><FileDown size={16}/> 导出</button></>}</div>
      {inventory ? <div className="data-table"><div className="table-row inventory-simple header"><span>商品信息</span><span>商品类别</span><span>库存</span><span>零售价</span><span>状态</span></div>{filtered.map(p => <div className="table-row inventory-simple" key={p.sku}><span className="product-cell"><i><Package size={18}/></i><b>{p.name}<small>{p.sku}</small></b></span><span>{p.category}</span><strong>{p.stock} {p.unit}</strong><span>{formatMoney(p.price)}</span><span><em className={p.status === "正常" ? "status ok" : "status low"}>{p.status}</em></span></div>)}</div>
      : page === "sales" || liveFlow ? <div className="data-table"><div className="table-row sales-simple header"><span>商品名称</span><span>订单号</span><span>售价</span><span>发生时间</span></div>{(liveFlow ? visibleTransactions : visibleTransactions.filter(t=>t.type==="sale")).map(t=><div className="table-row sales-simple clickable-row" key={t.id} onClick={onCreate}><span><b>{t.title}<small>{t.meta}</small></b></span><span>{t.id}</span><strong>{formatMoney(t.amount)}</strong><span>{formatFlowDate(t.time,month)}</span></div>)}</div>
      : <div className="data-table"><div className="table-row finance single-store header"><span>流水信息</span><span>收入类别</span><span>资金账户</span><span>记账日期</span><span>金额</span></div>{visibleTransactions.map(t => <div className="table-row finance single-store clickable-row" key={t.id} onClick={onCreate}><span><b>{t.title}</b></span><span>{t.meta.includes("续费") ? "续费" : t.meta.includes("新报") ? "新报" : t.meta.includes("销售") ? "销售" : t.meta}</span><span>{t.account}</span><span>{formatFlowDate(t.time,month)}</span><strong className={t.amount > 0 ? "positive" : "negative"}>{t.amount > 0 ? "+" : "−"}{formatMoney(Math.abs(t.amount))}</strong></div>)}</div>}
    </article>{!inventory&&page!=="course_flow"&&<div className={liveFlow ? "amount-summary live-total" : "amount-summary"}><div><span>{month}收入合计</span><b>{liveFlow ? formatMoney(liveIncome) : formatMoney(stats.income)}</b></div>{!liveFlow&&<><div><span>{month}支出合计</span><b>{formatMoney(stats.expense)}</b></div><div className="net"><span>本月净额</span><b>{formatMoney(stats.income-stats.expense)}</b></div></>}</div>}
  </>;
}

function Reports() {
  const {currentMonth:month,setCurrentMonth:setMonth}=useData();
  const [rangeStart,setRangeStart]=useState("2026年1月");
  const [rangeEnd,setRangeEnd]=useState(month);
  const stats = useLinkedStats(month);
  const reportTrend=trend.map(item=>({...item,balance:item.income-item.expense}));
  const charts=[['现金收入','income','#7568e8','#eeeafd'],['现金支出','expense','#f3b559','#fff4df'],['经营现金结余','balance','#48bba5','#e8faf5']] as const;
  return <><div className="page-heading"><div><p className="eyebrow">经营分析</p><h1>经营报表</h1><p>区分现金视角、课费收入和商品经营视角。</p></div><div className="heading-actions"><div className="period-range"><MonthSelect value={rangeStart} onChange={setRangeStart}/><span>-</span><MonthSelect value={rangeEnd} onChange={value=>{setRangeEnd(value);setMonth(value)}}/></div><button className="button secondary" onClick={()=>exportExcel(`边际音乐-${rangeStart}-${rangeEnd}-经营报表`,[{时间段:`${rangeStart}-${rangeEnd}`,现金收入:stats.income/100,现金支出:stats.expense/100,经营现金结余:(stats.income-stats.expense)/100}])}><FileDown size={17}/> 导出 Excel</button></div></div><div className="report-tabs"><button className="active">课费收入</button><button>新报收入</button><button>销售收入</button><button>商品毛利</button><button>库存报表</button></div><section className="metric-grid report"><MetricCard title="现金收入" value={formatMoney(stats.income)} note={`${rangeStart}—${rangeEnd}`} icon={ArrowDownLeft} tone="purple"/><MetricCard title="现金支出" value={formatMoney(stats.expense)} note={`${rangeStart}—${rangeEnd}`} icon={ArrowUpRight} tone="orange"/><MetricCard title="经营现金结余" value={formatMoney(stats.income-stats.expense)} note={`${rangeStart}—${rangeEnd}`} icon={Wallet} tone="green"/></section><div className="report-trend-grid">{charts.map(([title,key,stroke,fill])=><article className="panel report-chart mini" key={key}><div className="panel-title"><div><h2>{title}趋势</h2><p>{rangeStart} 至 {rangeEnd}</p></div></div><div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><AreaChart data={reportTrend}><CartesianGrid vertical={false} stroke="#eceaf3"/><XAxis dataKey="month" axisLine={false}/><YAxis axisLine={false} tickFormatter={v=>`${v}万`}/><Tooltip/><Area dataKey={key} name={title} stroke={stroke} fill={fill} strokeWidth={3}/></AreaChart></ResponsiveContainer></div></article>)}</div></>;
}

function SettingsPage() {
  const [section, setSection] = useState<string | null>(null);
  const settings = [[Landmark,"资金账户","4 个启用账户","现金、银行与支付平台"],[BookOpen,"课程类别","3 个启用类别","架子鼓、吉他、直播课"],[Users,"团队成员","6 位成员","管理员与普通成员权限"],[FileText,"财务分类","18 个收支分类","课时费、销售及日常经营"],[Settings,"企业信息","边际音乐","名称、时区与货币设置"]] as const;
  return <><div className="page-heading"><div><p className="eyebrow">企业配置</p><h1>系统设置</h1><p>维护边际音乐基础资料与团队权限。</p></div><button className="button primary" onClick={()=>setSection("邀请成员")}><Users size={17}/> 邀请成员</button></div><div className="settings-grid">{settings.map(([Icon,title,count,desc]) => <button className="setting-card" key={title} onClick={()=>setSection(title)}><i><Icon size={21}/></i><div><b>{title}</b><span>{desc}</span></div><strong>{count}</strong><ArrowUpRight size={16}/></button>)}</div><SettingsModal section={section} onClose={()=>setSection(null)}/></>;
}

function SettingsModal({ section, onClose }: { section: string | null; onClose: () => void }) {
  const [adding, setAdding] = useState(false);
  const [editing, setEditing] = useState<string | null>(null);
  if (!section) return null;
  const lists: Record<string, string[]> = {
    "资金账户": ["微信商户", "支付宝", "银行账户", "现金账户"],
    "课程类别": ["架子鼓", "吉他", "直播课"],
    "团队成员": ["王利凉 · 管理员", "李老师 · 普通成员", "张老师 · 普通成员"],
    "财务分类": ["课时费收入", "销售收入", "直播课收入", "房租物业", "工资支出", "日常费用"]
  };
  const invite = section === "邀请成员";
  const company = section === "企业信息";
  const subForm = adding || editing;
  return <div className="modal-backdrop" onMouseDown={e=>e.target===e.currentTarget&&onClose()}><div className="modal settings-modal"><div className="modal-title"><div><h2>{subForm ? `${editing ? "编辑" : "新增"}${section.replace("团队","").replace("财务","")}` : section}</h2><p>{subForm ? `填写${section}的完整信息` : invite ? "通过邮箱邀请成员加入边际音乐" : `管理${section}的详细资料`}</p></div><button className="close" onClick={subForm ? ()=>{setAdding(false);setEditing(null)} : onClose}><X size={20}/></button></div><div className="settings-body">{subForm ? <SettingsSubForm section={section} value={editing}/> : invite ? <><label>成员姓名<input placeholder="请输入姓名"/></label><label>邮箱地址<input type="email" placeholder="name@example.com"/></label><label>成员角色<select><option>普通成员</option><option>管理员</option></select></label></> : company ? <><label>企业名称<input defaultValue="边际音乐"/></label><label>默认货币<select><option>人民币 CNY</option></select></label><label>时区<select><option>Asia/Shanghai</option></select></label></> : <div className="settings-list">{(lists[section] ?? []).map((item,index)=><div key={item}><span><i>{index+1}</i>{item}</span><button onClick={()=>setEditing(item)}>编辑</button></div>)}<button className="add-setting" onClick={()=>setAdding(true)}><Plus size={16}/> 新增{section.replace("团队","").replace("财务","")}</button></div>}<div className="modal-actions"><button className="button secondary" onClick={subForm ? ()=>{setAdding(false);setEditing(null)} : onClose}>取消</button><button className="button primary" onClick={subForm ? ()=>{setAdding(false);setEditing(null)} : onClose}><Save size={16}/>{invite ? "发送邀请" : "保存设置"}</button></div></div></div></div>;
}

function SettingsSubForm({ section, value }: { section: string; value: string | null }) {
  if (section === "团队成员") return <><label>成员姓名<input defaultValue={value?.split(" · ")[0] ?? ""} placeholder="请输入姓名"/></label><label>邮箱<input type="email" placeholder="name@example.com"/></label><label>权限角色<select defaultValue={value?.includes("管理员") ? "管理员" : "普通成员"}><option>普通成员</option><option>管理员</option></select></label></>;
  if (section === "资金账户") return <><label>账户名称<input defaultValue={value ?? ""} placeholder="例如：银行账户"/></label><label>账户类型<select><option>银行</option><option>支付平台</option><option>现金</option></select></label><label>期初余额（元）<input type="number" min="0" step="0.01" placeholder="请输入期初余额"/></label></>;
  if (section === "课程类别") return <><label>课程名称<input defaultValue={value ?? ""} placeholder="例如：架子鼓"/></label><label>课程类型<select><option>线下课程</option><option>直播课程</option></select></label><label>状态<select><option>启用</option><option>停用</option></select></label></>;
  return <><label>分类名称<input defaultValue={value ?? ""}/></label><label>收支方向<select><option>收入</option><option>支出</option></select></label><label>状态<select><option>启用</option><option>停用</option></select></label></>;
}

function EntryModal({ open, onClose, mode }: { open: boolean; onClose: () => void; mode: NavKey }) {
  const {addEntry,lastDate,setLastDate}=useData();
  const [done, setDone] = useState(false);
  const [entryType, setEntryType] = useState<"income" | "expense">("income");
  useEffect(()=>{ if (open) { setDone(false); setEntryType("income"); } },[open,mode]);
  if (!open) return null;
  const isLive = mode === "live_flow";
  if (mode === "inventory") return <InventoryEntryModal onClose={onClose}/>;
  const isCourse = mode === "course_flow";
  return <div className="modal-backdrop" onMouseDown={e => e.target === e.currentTarget && onClose()}><div className="modal"><div className="modal-title"><div><h2>{done ? "记录成功" : isLive ? "新增直播流水" : isCourse ? "新增课程流水" : "新增资金流水"}</h2><p>{done ? "流水已计入本月经营报表" : isLive ? "记录直播课或直播商品收入" : isCourse ? "记录课程收费或退费" : "记录一笔边际音乐收入或支出"}</p></div><button className="close" onClick={onClose}><X size={20}/></button></div>{done ? <div className="success-state"><div><CheckCircle2 size={38}/></div><h3>已保存 ¥1,280.00 {isLive ? "直播收入" : isCourse ? "续费收入" : "收入"}</h3><p>{isLive ? "直播商品 · 微信商户" : "架子鼓课程 · 微信商户"}</p><button className="button primary" onClick={onClose}>完成</button></div> : <form onSubmit={e => {e.preventDefault();const data=new FormData(e.currentTarget);const amount=Math.round(Number(data.get("amount"))*100);addEntry({id:crypto.randomUUID(),title:String(data.get("summary")) + (mode==="sales" ? " · "+String(data.get("packageType")) : ""),category:String(data.get("category")),amountFen:amount,account:String(data.get("account")),date:String(data.get("date")),source:mode});setDone(true)}}><div className="type-toggle two"><button type="button" className="active">收入</button></div><div className="form-grid"><label>金额（元）<div className="amount-field"><span>¥</span><input name="amount" required type="number" min="0" step="0.01" defaultValue="1280.00" inputMode="decimal"/></div></label><label>记账日期<input name="date" required type="date" defaultValue={lastDate} onChange={e=>setLastDate(e.target.value)}/></label></div><div className="form-grid"><label>资金账户<select name="account" defaultValue={isLive || entryType === "income" ? "微信商户" : "银行账户"} key={entryType}><option>微信商户</option><option>支付宝</option><option>银行账户</option><option>现金账户</option></select></label><label>{isLive ? "直播分类" : "收支分类"}<select name="category" key={`category-`}>{isLive ? <><option>直播课</option><option>直播商品</option></> : isCourse ? <><option>续费收入</option><option>新报收入</option></> : <><option>销售收入</option><option>直播商品</option></>}</select></label>{isCourse && <label>课程类别<select><option>架子鼓</option><option>吉他</option><option>声乐</option></select></label>}{mode === "sales" && <label>卡类型<select name="packageType"><option>4次卡</option><option>单次卡</option></select></label>}</div><label>摘要<input name="summary" required key={`summary-`} defaultValue={isLive ? "直播商品销售" : isCourse ? "架子鼓续费" : "课程新报"}/></label><label>备注（选填）<textarea placeholder="添加更多说明…"/></label><div className="modal-actions"><button type="button" className="button secondary" onClick={onClose}>取消</button><button className="button primary">保存流水</button></div></form>}</div></div>;
}

function InventoryEntryModal({ onClose }: { onClose: () => void }) {
  const {adjustStock,lastDate,setLastDate}=useData();
  const [direction,setDirection] = useState<"in"|"out">("in");
  const [done,setDone] = useState(false);
  return <div className="modal-backdrop" onMouseDown={e=>e.target===e.currentTarget&&onClose()}><div className="modal"><div className="modal-title"><div><h2>{done ? "库存已更新" : "商品入库与出库"}</h2><p>{done ? "库存流水已生成" : "记录商品数量的增加或减少"}</p></div><button className="close" onClick={onClose}><X size={20}/></button></div>{done?<div className="success-state"><div><CheckCircle2 size={38}/></div><h3>{direction==="in"?"入库":"出库"}记录保存成功</h3><p>Roland TD-17KVX 电子鼓 · 1 套</p><button className="button primary" onClick={onClose}>完成</button></div>:<form onSubmit={e=>{e.preventDefault();const data=new FormData(e.currentTarget);const quantity=Number(data.get("quantity"));adjustStock(String(data.get("product")),direction==="in"?quantity:-quantity);setDone(true)}}><div className="type-toggle two"><button type="button" className={direction==="in"?"active":""} onClick={()=>setDirection("in")}>入库</button><button type="button" className={direction==="out"?"active":""} onClick={()=>setDirection("out")}>出库</button></div><label>商品名称<select name="product"><option value="DR-RL-TD17">Roland TD-17KVX 电子鼓</option><option value="GT-FD-CD60">Fender CD-60S 木吉他</option><option value="DR-TM-5A">TAMA 5A 胡桃木鼓棒</option><option value="AC-DAD-10">D&apos;Addario 民谣吉他弦</option></select></label><div className="form-grid"><label>{direction==="in"?"入库":"出库"}数量<input name="quantity" type="number" min="1" defaultValue="1"/></label><label>发生日期<input type="date" defaultValue={lastDate} onChange={e=>setLastDate(e.target.value)}/></label><label>{direction==="in"?"进货单价":"销售单价"}（元）<input type="number" min="0" step="0.01" defaultValue={direction==="in"?"7860.00":"9999.00"}/></label><label>业务类型<select>{direction==="in"?<><option>采购入库</option><option>退货入库</option><option>库存调整</option></>:<><option>销售出库</option><option>退货出库</option><option>报损出库</option></>}</select></label></div><label>单据编号<input defaultValue={direction==="in"?"RK20260805001":"CK20260805001"}/></label><label>备注（选填）<textarea placeholder="填写供应来源、出库用途等说明"/></label><div className="modal-actions"><button type="button" className="button secondary" onClick={onClose}>取消</button><button className="button primary">保存{direction==="in"?"入库":"出库"}</button></div></form>}</div></div>;
}

export default function Home() {
  const [active, setActive] = useState<NavKey>("dashboard");
  const [modal, setModal] = useState(false);
  const [mobile, setMobile] = useState(false);
  const [avatarUrl,setAvatarUrl]=useState("");
  const [searchQuery,setSearchQuery]=useState("");
  const [entries,setEntries]=useState<LinkedEntry[]>([]);
  const [fixedItems,setFixedItems]=useState<FixedItem[]>([]);
  const [stockDeltas,setStockDeltas]=useState<Record<string,number>>({});
  const [lastDate,setLastDate]=useState("2026-08-05");
  const [lastFixedTemplate,setLastFixedTemplate]=useState<FixedItem|null>(null);
  const [currentMonth,setCurrentMonth]=useState("2026年8月");
  const [storageReady,setStorageReady]=useState(false);
  useEffect(()=>{
    try {
      const saved=window.localStorage.getItem("bianji-finance-data-v1");
      if(saved){
        const parsed=JSON.parse(saved) as Partial<Pick<AppData,"entries"|"fixedItems"|"stockDeltas"|"lastDate"|"lastFixedTemplate">>;
        if(Array.isArray(parsed.entries)) setEntries(parsed.entries);
        if(Array.isArray(parsed.fixedItems)) setFixedItems(parsed.fixedItems);
        if(parsed.stockDeltas && typeof parsed.stockDeltas==="object") setStockDeltas(parsed.stockDeltas);
        if(typeof parsed.lastDate==="string") setLastDate(parsed.lastDate);
        if(typeof (parsed as {currentMonth?:unknown}).currentMonth==="string") setCurrentMonth((parsed as {currentMonth:string}).currentMonth);
        if(parsed.lastFixedTemplate) setLastFixedTemplate(parsed.lastFixedTemplate);
      }
    } catch { /* 损坏的本地缓存不影响页面使用 */ }
    setStorageReady(true);
  },[]);
  useEffect(()=>{
    if(!storageReady) return;
    window.localStorage.setItem("bianji-finance-data-v1",JSON.stringify({entries,fixedItems,stockDeltas,lastDate,lastFixedTemplate,currentMonth}));
  },[storageReady,entries,fixedItems,stockDeltas,lastDate,lastFixedTemplate,currentMonth]);
  const data:AppData={entries,addEntry:entry=>setEntries(current=>[...current,entry]),fixedItems,setFixedItems,stockDeltas,adjustStock:(sku,delta)=>setStockDeltas(current=>({...current,[sku]:(current[sku]??0)+delta})),lastDate,setLastDate,lastFixedTemplate,setLastFixedTemplate,currentMonth,setCurrentMonth};
  return <DataContext.Provider value={data}><div className="app-shell">
    <aside className={mobile ? "sidebar open" : "sidebar"}><div className="brand"><div className="brand-mark">边</div><div><b>边际音乐</b><span>财务经营管理</span></div><button className="mobile-close" onClick={()=>setMobile(false)}><X/></button></div><nav>{nav.map(item => <button key={item.key} className={active === item.key ? "active" : ""} onClick={()=>{setActive(item.key);setMobile(false)}}><item.icon size={19}/><span>{item.label}</span>{item.key === "live_flow" ? <em className="nav-badge">3</em> : item.key === "course_flow" ? <em className="nav-badge">1</em> : null}</button>)}<div className="nav-label">管理</div><button className={active === "settings" ? "active" : ""} onClick={()=>setActive("settings")}><Settings size={19}/><span>系统设置</span></button></nav><div className="sidebar-help"><CircleHelp size={19}/><div><b>需要帮助？</b><span>查看使用指南</span></div><ArrowUpRight size={15}/></div><label className="user-card avatar-picker"><input type="file" accept="image/*" hidden onChange={e=>{const file=e.target.files?.[0];if(file){const reader=new FileReader();reader.onload=()=>setAvatarUrl(String(reader.result));reader.readAsDataURL(file)}}}/><div><b>王利凉</b><span>管理员</span><small>点击更换头像</small></div><MoreHorizontal size={18}/></label></aside>
    {mobile && <div className="side-overlay" onClick={()=>setMobile(false)}/>}<main><header><button className="menu-button" onClick={()=>setMobile(true)}><Menu size={21}/></button><label className="global-search"><Search size={17}/><input aria-label="搜索流水、商品或单据" value={searchQuery} onChange={e=>setSearchQuery(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&searchQuery.trim()){setActive("finance")}}} placeholder="搜索流水、商品或单据"/><kbd>⌘ K</kbd></label><div className="header-right"/></header><div className="page-content">{active === "dashboard" ? <Dashboard onCreate={()=>setModal(true)}/> : <DataPage page={active as Exclude<NavKey, "dashboard">} onCreate={()=>setModal(true)}/>}<footer>边际音乐 · 数据更新于刚刚</footer></div></main><EntryModal open={modal} mode={active} onClose={()=>setModal(false)}/>
  </div></DataContext.Provider>;
}
