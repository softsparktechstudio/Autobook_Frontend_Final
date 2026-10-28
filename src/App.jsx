import React, { useEffect, useMemo, useState } from "react";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import {
  Activity, ArrowDownLeft, ArrowUpRight, BookOpen, Building2, CalendarDays, Truck, FileBarChart,
  CheckCircle2, ChevronDown, CircleHelp, FileText, LayoutDashboard, LogOut,
  Menu, Plus, RefreshCw, Search, ShieldCheck, Wallet, X, ReceiptText, Scale, Users, FilePlus2, IndianRupee, MessageCircle, Send, Bot, Package, Landmark, ClipboardCheck, FileSpreadsheet, UserCog
} from "lucide-react";

const API = import.meta.env.VITE_API_URL || "http://localhost:5000/api";

async function request(path, token, options = {}) {
  // Production VITE_API_URL must end with /api.
  const url = `${API.replace(/\/+$/, "")}/${path.replace(/^\/+/, "")}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      ...(options.body != null ? { "Content-Type": "application/json" } : {}),
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(options.headers || {})
    }
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || `Request failed (${response.status})`);
  return data;
}

function Login({ onLogin }) {
  const [email, setEmail] = useState("admin@softspark.local");
  const [password, setPassword] = useState("LocalAdmin123!");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function submit(e) {
    e.preventDefault();
    setBusy(true); setError("");
    try {
      const data = await request("/auth/login", null, {
        method: "POST", body: JSON.stringify({ email, password })
      });
      onLogin(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="login-shell">
      <div className="login-brand">
        <div className="brand-mark large">S</div>
        <div>
          <div className="brand-name">SoftSpark <span>AutoBooks Enterprise</span></div>
          <div className="brand-subtitle">Accounting, reconciliation, inventory and business controls</div>
        </div>
      </div>
      <div className="login-card">
        <div className="eyebrow">AUTOBOOKS ENTERPRISE · RELEASE 3</div>
        <h1>Welcome back</h1>
        <p className="muted">Sign in to manage your organization and accounting setup.</p>
        <form onSubmit={submit} className="form-stack">
          <label>Email address<input type="email" value={email} onChange={e => setEmail(e.target.value)} required autoComplete="username" /></label>
          <label>Password<input type="password" value={password} onChange={e => setPassword(e.target.value)} required autoComplete="current-password" /></label>
          {error && <div className="alert error">{error}</div>}
          <button className="button primary full" disabled={busy}>{busy ? "Signing in…" : "Sign in securely"}</button>
        </form>
        <div className="login-note"><ShieldCheck size={16}/> Local development build · API authentication enabled</div>
      </div>
      <div className="login-footer">Built for businesses of every size · India-first foundation</div>
    </div>
  );
}

const nav = [
  { label: "Overview", icon: LayoutDashboard, id: "overview" },
  { label: "Company setup", icon: Building2, id: "company" },
  { label: "Financial years", icon: CalendarDays, id: "years" },
  { label: "Chart of accounts", icon: BookOpen, id: "accounts" },
  { label: "Journal entries", icon: ReceiptText, id: "journals" },
  { label: "Trial balance", icon: Scale, id: "trial-balance" },
  { label: "Financial reports", icon: FileBarChart, id: "financial-reports" },
  { label: "Customers & suppliers", icon: Users, id: "parties" },
  { label: "Sales invoices", icon: FilePlus2, id: "sales-invoices" },
  { label: "Purchase bills", icon: Truck, id: "purchase-bills" },
  { label: "AI insights", icon: Activity, id: "ai-insights" },
  { label: "AI assistant", icon: MessageCircle, id: "ai-assistant" },
  { label: "Inventory & stock", icon: Package, id: "inventory" },
  { label: "Bank reconciliation", icon: Landmark, id: "bank-reconciliation" },
  { label: "Approval centre", icon: ClipboardCheck, id: "approvals" },
  { label: "Tax registers", icon: FileSpreadsheet, id: "tax-registers" },
  { label: "Team & accountant access", icon: UserCog, id: "team" },
  { label: "Enterprise centre", icon: Building2, id: "enterprise" }
];

function BusinessWorkspace({ page, session, data, setToast, load }) {
  const [items,setItems]=useState([]),[movements,setMovements]=useState([]),[historyItem,setHistoryItem]=useState(''),[bankAccounts,setBankAccounts]=useState([]),[statement,setStatement]=useState([]),[bookLines,setBookLines]=useState([]),[approvals,setApprovals]=useState([]),[tax,setTax]=useState({sales:[],purchases:[]}),[team,setTeam]=useState([]);
  const [selectedAccount,setSelectedAccount]=useState(''),[selectedStatement,setSelectedStatement]=useState(''),[selectedBookLine,setSelectedBookLine]=useState('');
  const [busy,setBusy]=useState(false),[error,setError]=useState(''),[csv,setCsv]=useState('date,description,reference,amount\n2026-10-01,Customer deposit,UTR123,1000.00\n2026-10-02,Supplier payment,UTR124,-500.00');
  const [itemForm,setItemForm]=useState({sku:'',name:'',description:'',unit:'unit',openingQuantity:'0',reorderLevel:'0',unitCost:'0'});
  const [movementForm,setMovementForm]=useState({itemId:'',movementDate:new Date().toISOString().slice(0,10),movementType:'receipt',quantity:'1',unitCost:'0',reference:'',notes:''});
  const [approvalForm,setApprovalForm]=useState({entryDate:new Date().toISOString().slice(0,10),description:'',reference:'',debitAccount:'',creditAccount:'',amount:'',note:''});
  const [teamForm,setTeamForm]=useState({fullName:'',email:'',temporaryPassword:'',role:'data_entry'});
  const money=v=>`₹${Number(v||0).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
  function exportBusinessRows(rows, filename, title=filename){const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),title.slice(0,31));XLSX.writeFile(wb,`${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);}
  function exportBusinessPdf(rows, filename, title){const doc=new jsPDF({orientation:'landscape'});doc.setFontSize(15);doc.text(`AutoBooks Business — ${title}`,14,16);const cols=Object.keys(rows[0]||{Information:'No records'});autoTable(doc,{startY:23,head:[cols],body:rows.map(r=>cols.map(c=>String(r[c]??''))),styles:{fontSize:7,cellPadding:2}});doc.save(`${filename}_${new Date().toISOString().slice(0,10)}.pdf`);}
  async function refresh(){setBusy(true);setError('');try{
    if(page==='inventory'){const r=await request('/business/inventory',session.token);setItems(r.items||[]);const chosen=historyItem||r.items?.[0]?.id||'';if(chosen){setHistoryItem(chosen);const h=await request(`/business/inventory/${chosen}/movements`,session.token);setMovements(h.movements||[]);}}
    if(page==='bank-reconciliation'){const r=await request('/business/bank/accounts',session.token);setBankAccounts(r.accounts||[]);const chosen=selectedAccount||r.accounts?.[0]?.id||'';if(chosen){setSelectedAccount(chosen);const [s,b]=await Promise.all([request(`/business/bank/${chosen}/lines`,session.token),request(`/business/bank/${chosen}/book-lines`,session.token)]);setStatement(s.lines||[]);setBookLines(b.lines||[]);}}
    if(page==='approvals'){const r=await request('/business/approvals',session.token);setApprovals(r.requests||[]);}
    if(page==='tax-registers'){setTax(await request('/business/tax-registers',session.token));}
    if(page==='team'){const r=await request('/business/team',session.token);setTeam(r.members||[]);}
  }catch(e){setError(e.message);}finally{setBusy(false);}}
  useEffect(()=>{refresh();},[page,session?.token]);
  useEffect(()=>{if(page==='inventory'&&historyItem){(async()=>{try{const h=await request(`/business/inventory/${historyItem}/movements`,session.token);setMovements(h.movements||[]);}catch(e){setError(e.message);}})();}},[historyItem]);
  useEffect(()=>{if(page==='bank-reconciliation'&&selectedAccount){(async()=>{try{const [s,b]=await Promise.all([request(`/business/bank/${selectedAccount}/lines`,session.token),request(`/business/bank/${selectedAccount}/book-lines`,session.token)]);setStatement(s.lines||[]);setBookLines(b.lines||[]);}catch(e){setError(e.message);}})();}},[selectedAccount]);
  async function submit(e,path,body,success){e.preventDefault();setBusy(true);setError('');try{const r=await request(path,session.token,{method:'POST',body:JSON.stringify(body)});setToast(success||r.message||'Saved successfully.');await refresh();if(load)await load();return r;}catch(err){setError(err.message);return null;}finally{setBusy(false);}}
  async function importStatement(e){e.preventDefault();try{const rows=csv.split(/\r?\n/).map(x=>x.trim()).filter(Boolean).slice(1).map(line=>{const p=line.split(',');return {date:p[0]?.trim(),description:p[1]?.trim(),reference:p[2]?.trim()||'',amount:Number(p[3]?.trim())};});if(!rows.length)throw new Error('Paste at least one statement data row.');const r=await request(`/business/bank/${selectedAccount}/import`,session.token,{method:'POST',body:JSON.stringify({lines:rows})});setToast(`Imported ${r.imported} statement rows.`);await refresh();}catch(e){setError(e.message+' Use CSV columns: date,description,reference,amount. Amounts: deposits positive, withdrawals negative. Do not include commas inside descriptions.');}}
  async function matchLines(){try{await request(`/business/bank/${selectedAccount}/match`,session.token,{method:'POST',body:JSON.stringify({statementLineId:selectedStatement,journalLineId:selectedBookLine})});setToast('Statement row matched to ledger.');setSelectedStatement('');setSelectedBookLine('');await refresh();}catch(e){setError(e.message);}}
  const panelTitle=({inventory:'Inventory & stock control','bank-reconciliation':'Bank reconciliation',approvals:'Approval centre','tax-registers':'Tax registers & exports',team:'Accountant collaboration'})[page]||'Business workspace';
  return <section className="panel business-workspace"><div className="panel-header"><div><h2>{panelTitle}</h2><p>AutoBooks Business · Release 2 · Company-scoped operational controls.</p></div><button className="button secondary" onClick={refresh} disabled={busy}><RefreshCw size={15}/> Refresh</button></div>
    {error&&<div className="alert error">{error}</div>}{busy&&<div className="loading-line">Loading or saving business data…</div>}
    {page==='inventory'&&<>
      <form className="form-stack business-form" onSubmit={e=>submit(e,'/business/inventory',itemForm,'Inventory item created.').then(r=>{if(r)setItemForm({sku:'',name:'',description:'',unit:'unit',openingQuantity:'0',reorderLevel:'0',unitCost:'0'});})}><h3>Add inventory item</h3><div className="business-grid"><label>SKU / Item code<input value={itemForm.sku} onChange={e=>setItemForm({...itemForm,sku:e.target.value})} required maxLength="80"/></label><label>Item name<input value={itemForm.name} onChange={e=>setItemForm({...itemForm,name:e.target.value})} required maxLength="200"/></label><label>Unit<input value={itemForm.unit} onChange={e=>setItemForm({...itemForm,unit:e.target.value})} required/></label><label>Opening quantity<input type="number" min="0" step="0.001" value={itemForm.openingQuantity} onChange={e=>setItemForm({...itemForm,openingQuantity:e.target.value})}/></label><label>Reorder alert level<input type="number" min="0" step="0.001" value={itemForm.reorderLevel} onChange={e=>setItemForm({...itemForm,reorderLevel:e.target.value})}/></label><label>Unit cost (₹)<input type="number" min="0" step="0.01" value={itemForm.unitCost} onChange={e=>setItemForm({...itemForm,unitCost:e.target.value})}/></label><label className="wide-field">Description<input value={itemForm.description} onChange={e=>setItemForm({...itemForm,description:e.target.value})}/></label></div><button className="button primary" disabled={busy}><Plus size={15}/> Create item</button></form>
      <form className="form-stack business-form" onSubmit={e=>submit(e,`/business/inventory/${movementForm.itemId}/movements`,{...movementForm,quantity:Number(movementForm.quantity),unitCost:Number(movementForm.unitCost)},'Stock movement recorded.').then(r=>{if(r)setMovementForm({...movementForm,quantity:'1',reference:'',notes:''});})}><h3>Record stock movement</h3><div className="business-grid"><label>Item<select value={movementForm.itemId} onChange={e=>setMovementForm({...movementForm,itemId:e.target.value})} required><option value="">Select item</option>{items.map(i=><option key={i.id} value={i.id}>{i.sku} — {i.name} (on hand: {Number(i.on_hand).toFixed(3)})</option>)}</select></label><label>Date<input type="date" value={movementForm.movementDate} onChange={e=>setMovementForm({...movementForm,movementDate:e.target.value})} required/></label><label>Movement<select value={movementForm.movementType} onChange={e=>setMovementForm({...movementForm,movementType:e.target.value})}><option value="receipt">Stock receipt</option><option value="issue">Stock issue</option><option value="adjustment_in">Adjustment in</option><option value="adjustment_out">Adjustment out</option></select></label><label>Quantity<input type="number" min="0.001" step="0.001" value={movementForm.quantity} onChange={e=>setMovementForm({...movementForm,quantity:e.target.value})} required/></label><label>Unit cost (₹)<input type="number" min="0" step="0.01" value={movementForm.unitCost} onChange={e=>setMovementForm({...movementForm,unitCost:e.target.value})}/></label><label>Reference<input value={movementForm.reference} onChange={e=>setMovementForm({...movementForm,reference:e.target.value})}/></label></div><button className="button secondary" disabled={busy||!movementForm.itemId}>Record movement</button></form>
      <div className="section-heading"><h3>Stock summary</h3><button className="button secondary small" onClick={()=>exportBusinessRows(items.map(i=>({SKU:i.sku,Item:i.name,Unit:i.unit,OnHand:Number(i.on_hand),ReorderLevel:Number(i.reorder_level),UnitCost:Number(i.unit_cost),StockValue:Number(i.on_hand)*Number(i.unit_cost)})),'AutoBooks_Inventory','Inventory')}><FileSpreadsheet size={15}/> Export stock Excel</button></div><div className="table-wrap"><table><thead><tr><th>SKU</th><th>Item</th><th>Unit</th><th>On hand</th><th>Reorder level</th><th>Unit cost</th><th>Stock flag</th></tr></thead><tbody>{items.map(i=><tr key={i.id}><td>{i.sku}</td><td><strong>{i.name}</strong></td><td>{i.unit}</td><td>{Number(i.on_hand).toFixed(3)}</td><td>{Number(i.reorder_level).toFixed(3)}</td><td>{money(i.unit_cost)}</td><td><span className={`status-pill ${Number(i.on_hand)<=Number(i.reorder_level)?'warning':''}`}>{Number(i.on_hand)<=Number(i.reorder_level)?'Reorder':'In stock'}</span></td></tr>)}{!items.length&&<tr><td colSpan="7">No items yet. Add an item above.</td></tr>}</tbody></table></div><div className="business-form"><h3>Stock movement history</h3><label>Choose item<select value={historyItem} onChange={e=>setHistoryItem(e.target.value)}><option value="">Select item</option>{items.map(i=><option key={i.id} value={i.id}>{i.sku} — {i.name}</option>)}</select></label><div className="table-wrap"><table><thead><tr><th>Date</th><th>Type</th><th>Quantity</th><th>Unit cost</th><th>Reference</th><th>Notes</th></tr></thead><tbody>{movements.map(m=><tr key={m.id}><td>{String(m.movement_date).slice(0,10)}</td><td>{m.movement_type}</td><td>{Number(m.quantity).toFixed(3)}</td><td>{money(m.unit_cost)}</td><td>{m.reference||'—'}</td><td>{m.notes||'—'}</td></tr>)}{!movements.length&&<tr><td colSpan="6">No stock movements recorded for this item.</td></tr>}</tbody></table></div></div><div className="info-footer">Stock movements update the quantity ledger and prevent issuing more stock than available. Inventory is not yet automatically posted to accounting journals or linked to invoice line items; use controlled financial journals for valuation until that integration is added.</div>
    </>}
    {page==='bank-reconciliation'&&<>
      <div className="business-form"><label>Bank / cash ledger account<select value={selectedAccount} onChange={e=>setSelectedAccount(e.target.value)}><option value="">Choose ledger account</option>{bankAccounts.map(a=><option key={a.id} value={a.id}>{a.code} — {a.name} · Book balance {money(a.book_balance)}</option>)}</select></label></div>
      <form className="form-stack business-form" onSubmit={importStatement}><h3>Import bank statement (CSV)</h3><p className="muted">Use four columns in this order: date,description,reference,amount. Positive amounts are deposits; negative amounts are withdrawals. Dates must be YYYY-MM-DD.</p><textarea rows="5" value={csv} onChange={e=>setCsv(e.target.value)} required/><button className="button primary" disabled={!selectedAccount||busy}>Import statement rows</button></form>
      <div className="business-grid"><div className="metric-card"><span>Statement rows</span><strong>{statement.length}</strong></div><div className="metric-card"><span>Matched rows</span><strong>{statement.filter(x=>x.matched_journal_line_id).length}</strong></div><div className="metric-card"><span>Unmatched rows</span><strong>{statement.filter(x=>!x.matched_journal_line_id).length}</strong></div><div className="metric-card"><span>Unmatched ledger lines</span><strong>{bookLines.length}</strong></div></div>
      <div className="business-grid"><label>Unmatched statement row<select value={selectedStatement} onChange={e=>setSelectedStatement(e.target.value)}><option value="">Choose statement row</option>{statement.filter(x=>!x.matched_journal_line_id).map(x=><option key={x.id} value={x.id}>{String(x.statement_date).slice(0,10)} · {x.description} · {money(x.amount)}</option>)}</select></label><label>Matching ledger transaction<select value={selectedBookLine} onChange={e=>setSelectedBookLine(e.target.value)}><option value="">Choose ledger line</option>{bookLines.map(x=><option key={x.journal_line_id} value={x.journal_line_id}>{String(x.entry_date).slice(0,10)} · JV-{x.entry_number} · {x.description} · {money(x.signed_amount)}</option>)}</select></label></div><button className="button primary" disabled={!selectedAccount||!selectedStatement||!selectedBookLine||busy} onClick={matchLines}><CheckCircle2 size={15}/> Match selected transactions</button>
      <div className="section-heading"><h3>Imported bank statement</h3><div className="inline-actions"><button className="button secondary small" onClick={()=>exportBusinessRows(statement.map(x=>({Date:String(x.statement_date).slice(0,10),Description:x.description,Reference:x.reference||'',Amount:Number(x.amount),Status:x.matched_journal_line_id?'Matched':'Unmatched'})),'AutoBooks_Bank_Reconciliation','Bank statement')}><FileSpreadsheet size={15}/> Export Excel</button><button className="button secondary small" onClick={()=>exportBusinessPdf(statement.map(x=>({Date:String(x.statement_date).slice(0,10),Description:x.description,Reference:x.reference||'',Amount:Number(x.amount),Status:x.matched_journal_line_id?'Matched':'Unmatched'})),'AutoBooks_Bank_Reconciliation','Bank reconciliation')}><FileText size={15}/> Export PDF</button></div></div><div className="table-wrap"><table><thead><tr><th>Date</th><th>Bank description</th><th>Reference</th><th>Amount</th><th>Match status</th><th>Action</th></tr></thead><tbody>{statement.map(x=><tr key={x.id}><td>{String(x.statement_date).slice(0,10)}</td><td>{x.description}</td><td>{x.reference||'—'}</td><td>{money(x.amount)}</td><td>{x.matched_journal_line_id?'Matched':'Unmatched'}</td><td>{x.matched_journal_line_id&&<button className="button secondary small" onClick={async()=>{try{await request(`/business/bank/${selectedAccount}/unmatch`,session.token,{method:'POST',body:JSON.stringify({statementLineId:x.id})});setToast('Statement row unmatched.');await refresh();}catch(e){setError(e.message);}}}>Unmatch</button>}</td></tr>)}{!statement.length&&<tr><td colSpan="6">No statement rows imported for this account.</td></tr>}</tbody></table></div><div className="info-footer">Matching requires an exact signed amount and same ledger account. Date proximity and reference are for your review; this module does not connect directly to bank APIs or certify a final reconciliation statement.</div>
    </>}
    {page==='approvals'&&<>
      <form className="form-stack business-form" onSubmit={async e=>{e.preventDefault();const amount=Number(approvalForm.amount);if(!approvalForm.debitAccount||!approvalForm.creditAccount||approvalForm.debitAccount===approvalForm.creditAccount||amount<=0){setError('Choose two different accounts and enter an amount above zero.');return;}const r=await submit(e,'/business/approvals/journals',{entryDate:approvalForm.entryDate,description:approvalForm.description,reference:approvalForm.reference,lines:[{accountId:approvalForm.debitAccount,description:approvalForm.note,debit:amount,credit:0},{accountId:approvalForm.creditAccount,description:approvalForm.note,debit:0,credit:amount}]},'Journal submitted for approval.');if(r)setApprovalForm({...approvalForm,description:'',reference:'',amount:'',note:''});}}><h3>Submit journal for independent approval</h3><div className="business-grid"><label>Entry date<input type="date" value={approvalForm.entryDate} onChange={e=>setApprovalForm({...approvalForm,entryDate:e.target.value})} required/></label><label>Description<input value={approvalForm.description} onChange={e=>setApprovalForm({...approvalForm,description:e.target.value})} required minLength="3"/></label><label>Debit account<select value={approvalForm.debitAccount} onChange={e=>setApprovalForm({...approvalForm,debitAccount:e.target.value})} required><option value="">Select account</option>{data.accounts.filter(a=>a.is_active).map(a=><option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></label><label>Credit account<select value={approvalForm.creditAccount} onChange={e=>setApprovalForm({...approvalForm,creditAccount:e.target.value})} required><option value="">Select account</option>{data.accounts.filter(a=>a.is_active).map(a=><option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></label><label>Amount (₹)<input type="number" min="0.01" step="0.01" value={approvalForm.amount} onChange={e=>setApprovalForm({...approvalForm,amount:e.target.value})} required/></label><label>Reference<input value={approvalForm.reference} onChange={e=>setApprovalForm({...approvalForm,reference:e.target.value})}/></label></div><button className="button primary" disabled={busy}><Send size={15}/> Submit for approval</button><p className="small-note">This simple workflow creates a balanced two-line journal request. It is not posted to the ledger unless a different authorized reviewer approves it.</p></form>
      <h3>Approval queue</h3><div className="table-wrap"><table><thead><tr><th>Request</th><th>Requested by</th><th>Created</th><th>Status</th><th>Review note</th><th>Action</th></tr></thead><tbody>{approvals.map(a=><tr key={a.id}><td><strong>{a.title}</strong><div className="small-note">{a.payload?.entryDate} · {money((a.payload?.lines||[]).reduce((n,l)=>n+Number(l.debit||0),0))}</div></td><td>{a.requester_name}</td><td>{String(a.created_at).slice(0,10)}</td><td>{a.status}</td><td>{a.reviewer_note||'—'}</td><td>{a.status==='pending'&&a.requested_by!==session.user.id&&['owner','admin','accountant'].includes(session.user.role)?<div className="inline-actions"><button className="button primary small" onClick={async()=>{const note=window.prompt('Optional approval note','Approved');if(note===null)return;try{await request(`/business/approvals/${a.id}/review`,session.token,{method:'POST',body:JSON.stringify({decision:'approve',note:note||''})});setToast('Approved and posted to ledger.');await refresh();if(load)await load();}catch(e){setError(e.message);}}}>Approve & post</button><button className="button secondary small" onClick={async()=>{const note=window.prompt('Reason for rejection','Please revise and resubmit');if(note===null)return;try{await request(`/business/approvals/${a.id}/review`,session.token,{method:'POST',body:JSON.stringify({decision:'reject',note})});setToast('Request rejected.');await refresh();}catch(e){setError(e.message);}}}>Reject</button></div>:a.journal_entry_id?`Journal ${a.journal_entry_id.slice(0,8)}`:'—'}</td></tr>)}{!approvals.length&&<tr><td colSpan="6">No approval requests yet.</td></tr>}</tbody></table></div><div className="info-footer">Requesters cannot approve their own requests. Approvals post the journal atomically after rechecking company accounts and the open financial year. Existing Release 1 direct journal posting remains available to its existing roles.</div>
    </>}
    {page==='tax-registers'&&<><div className="section-heading"><h3>Tax register exports</h3><div className="inline-actions"><button className="button secondary small" onClick={()=>exportBusinessRows((tax.sales||[]).map(r=>({Invoice:r.document_number,Date:String(r.document_date).slice(0,10),Customer:r.party_name,GSTIN:r.gstin||'',Taxable:Number(r.taxable_amount),Rate:Number(r.gst_rate),CGST:Number(r.cgst_amount),SGST:Number(r.sgst_amount),IGST:Number(r.igst_amount),Total:Number(r.total_amount),Status:r.status})),'AutoBooks_Sales_Tax_Register','Sales register')}><FileSpreadsheet size={15}/> Sales Excel</button><button className="button secondary small" onClick={()=>exportBusinessPdf((tax.sales||[]).map(r=>({Invoice:r.document_number,Date:String(r.document_date).slice(0,10),Customer:r.party_name,GSTIN:r.gstin||'',Taxable:Number(r.taxable_amount),CGST:Number(r.cgst_amount),SGST:Number(r.sgst_amount),IGST:Number(r.igst_amount),Total:Number(r.total_amount)})),'AutoBooks_Sales_Tax_Register','Sales / output GST register')}><FileText size={15}/> Sales PDF</button><button className="button secondary small" onClick={()=>exportBusinessRows((tax.purchases||[]).map(r=>({Bill:r.document_number,Date:String(r.document_date).slice(0,10),Supplier:r.party_name,GSTIN:r.gstin||'',Taxable:Number(r.taxable_amount),Rate:Number(r.gst_rate),CGST:Number(r.cgst_amount),SGST:Number(r.sgst_amount),IGST:Number(r.igst_amount),Total:Number(r.total_amount),Status:r.status})),'AutoBooks_Purchase_Tax_Register','Purchase register')}><FileSpreadsheet size={15}/> Purchases Excel</button><button className="button secondary small" onClick={()=>exportBusinessPdf((tax.purchases||[]).map(r=>({Bill:r.document_number,Date:String(r.document_date).slice(0,10),Supplier:r.party_name,GSTIN:r.gstin||'',Taxable:Number(r.taxable_amount),CGST:Number(r.cgst_amount),SGST:Number(r.sgst_amount),IGST:Number(r.igst_amount),Total:Number(r.total_amount)})),'AutoBooks_Purchase_Tax_Register','Purchase / input GST register')}><FileText size={15}/> Purchases PDF</button></div></div><div className="info-footer"><ShieldCheck size={16}/> {tax.disclaimer||'Review with an accountant before statutory filing.'}</div><h3>Sales / output GST register</h3><div className="table-wrap"><table><thead><tr><th>Invoice</th><th>Date</th><th>Customer</th><th>GSTIN</th><th>Taxable</th><th>CGST</th><th>SGST</th><th>IGST</th><th>Total</th></tr></thead><tbody>{(tax.sales||[]).map((r,i)=><tr key={r.document_number+i}><td>{r.document_number}</td><td>{String(r.document_date).slice(0,10)}</td><td>{r.party_name}</td><td>{r.gstin||'—'}</td><td>{money(r.taxable_amount)}</td><td>{money(r.cgst_amount)}</td><td>{money(r.sgst_amount)}</td><td>{money(r.igst_amount)}</td><td>{money(r.total_amount)}</td></tr>)}{!tax.sales?.length&&<tr><td colSpan="9">No sales tax records.</td></tr>}</tbody></table></div><h3>Purchase / input GST register</h3><div className="table-wrap"><table><thead><tr><th>Bill</th><th>Date</th><th>Supplier</th><th>GSTIN</th><th>Taxable</th><th>CGST</th><th>SGST</th><th>IGST</th><th>Total</th></tr></thead><tbody>{(tax.purchases||[]).map((r,i)=><tr key={r.document_number+i}><td>{r.document_number}</td><td>{String(r.document_date).slice(0,10)}</td><td>{r.party_name}</td><td>{r.gstin||'—'}</td><td>{money(r.taxable_amount)}</td><td>{money(r.cgst_amount)}</td><td>{money(r.sgst_amount)}</td><td>{money(r.igst_amount)}</td><td>{money(r.total_amount)}</td></tr>)}{!tax.purchases?.length&&<tr><td colSpan="9">No purchase tax records.</td></tr>}</tbody></table></div><div className="info-footer">Use the page-level Export Excel / Export PDF controls for the visible register data where supported. This is not a GST portal filing integration and should not be described as filing-ready.</div></>}
    {page==='team'&&<><form className="form-stack business-form" onSubmit={e=>submit(e,'/business/team',teamForm,'Team member created. Share the temporary password securely.').then(r=>{if(r)setTeamForm({fullName:'',email:'',temporaryPassword:'',role:'data_entry'});})}><h3>Add team member</h3><div className="business-grid"><label>Full name<input value={teamForm.fullName} onChange={e=>setTeamForm({...teamForm,fullName:e.target.value})} required/></label><label>Email<input type="email" value={teamForm.email} onChange={e=>setTeamForm({...teamForm,email:e.target.value})} required/></label><label>Temporary password (12+ characters)<input type="password" minLength="12" value={teamForm.temporaryPassword} onChange={e=>setTeamForm({...teamForm,temporaryPassword:e.target.value})} required/></label><label>Role<select value={teamForm.role} onChange={e=>setTeamForm({...teamForm,role:e.target.value})}><option value="accountant">Accountant</option><option value="data_entry">Data entry / preparer</option><option value="read_only">Read only</option></select></label></div><button className="button primary" disabled={busy||!['owner','admin'].includes(session.user.role)}><Plus size={15}/> Add member</button>{!['owner','admin'].includes(session.user.role)&&<p className="small-note">Only owners and administrators can add team members.</p>}</form><h3>Organization members</h3><div className="table-wrap"><table><thead><tr><th>Name</th><th>Email</th><th>Status</th><th>Role</th><th>Change role</th></tr></thead><tbody>{team.map(m=><tr key={m.id}><td>{m.full_name}</td><td>{m.email}</td><td>{m.status}</td><td>{m.role}</td><td>{m.id!==session.user.id&&['owner','admin'].includes(session.user.role)?<select value={m.role} onChange={async e=>{try{await request(`/business/team/${m.id}/role`,session.token,{method:'PATCH',body:JSON.stringify({role:e.target.value})});setToast('Role updated.');await refresh();}catch(err){setError(err.message);}}}><option value="accountant">Accountant</option><option value="data_entry">Data entry</option><option value="read_only">Read only</option></select>:'—'}</td></tr>)}{!team.length&&<tr><td colSpan="5">No members found.</td></tr>}</tbody></table></div><div className="info-footer">New accounts use a temporary password that must be shared securely. This build does not yet include invitation email, self-service password reset, MFA, or a forced password-change screen. Do not use shared credentials.</div></>}
  </section>;
}


function EnterpriseWorkspace({session,data,setToast}) {
  const today = new Date().toISOString().slice(0,10);
  const [tab,setTab]=useState('overview');
  const [entities,setEntities]=useState([]),[branches,setBranches]=useState([]),[centres,setCentres]=useState([]),[budgets,setBudgets]=useState([]),[transactions,setTransactions]=useState([]),[eliminations,setEliminations]=useState([]),[summary,setSummary]=useState({});
  const [busy,setBusy]=useState(false),[error,setError]=useState('');
  const [entityForm,setEntityForm]=useState({legalName:'',entityCode:'',currency:'INR',countryCode:'IN',taxIdentifier:''});
  const [branchForm,setBranchForm]=useState({entityId:'',branchCode:'',name:'',address:'',managerName:''});
  const [centreForm,setCentreForm]=useState({centreCode:'',name:'',department:'',ownerName:''});
  const [budgetForm,setBudgetForm]=useState({budgetName:'',periodStart:today.slice(0,4)+'-04-01',periodEnd:(Number(today.slice(0,4))+(Number(today.slice(5,7))>=4?1:0))+'-03-31'});
  const [lineForm,setLineForm]=useState({budgetId:'',accountId:'',costCentreId:'',plannedAmount:'',notes:''});
  const [icForm,setIcForm]=useState({sourceEntityId:'',targetEntityId:'',transactionDate:today,reference:'',description:'',amount:'',currency:'INR'});
  const [elimForm,setElimForm]=useState({transactionId:'',eliminationDate:today,description:'',amount:''});
  async function loadAll(){setError('');try{const paths=['/enterprise/summary','/enterprise/entities','/enterprise/branches','/enterprise/cost-centres','/enterprise/budgets','/enterprise/intercompany','/enterprise/eliminations'];const r=await Promise.all(paths.map(p=>request(p,session.token)));setSummary(r[0]);setEntities(r[1].entities||[]);setBranches(r[2].branches||[]);setCentres(r[3].costCentres||[]);setBudgets(r[4].budgets||[]);setTransactions(r[5].transactions||[]);setEliminations(r[6].eliminations||[]);}catch(e){setError(e.message);}}
  useEffect(()=>{loadAll();},[]);
  async function save(path,body,success,reset){setBusy(true);setError('');try{await request(path,session.token,{method:'POST',body:JSON.stringify(body)});if(reset)reset();setToast(success);await loadAll();}catch(e){setError(e.message);}finally{setBusy(false);}}
  const fmt=v=>'₹'+Number(v||0).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
  const tabs=[['overview','Overview'],['entities','Legal entities & branches'],['centres','Cost centres'],['budgets','Budgets'],['intercompany','Intercompany'],['consolidation','Consolidation workpapers']];
  return <section className="panel">
    <div className="panel-header"><div><h2>AutoBooks Enterprise</h2><p>Multi-entity master data, departmental planning and consolidation workpapers.</p></div><button className="button secondary" onClick={loadAll}><RefreshCw size={15}/> Refresh</button></div>
    <div className="inline-actions" style={{flexWrap:'wrap',marginBottom:18}}>{tabs.map(([id,label])=><button key={id} className={`button ${tab===id?'primary':'secondary'} small`} onClick={()=>setTab(id)}>{label}</button>)}</div>
    {error&&<div className="alert error">{error}</div>}
    {tab==='overview'&&<><div className="metric-grid"><Metric icon={Building2} label="Active entities" value={summary.entities||0} sub="Enterprise directory"/><Metric icon={Truck} label="Active branches" value={summary.branches||0} sub="Registered locations"/><Metric icon={Users} label="Cost centres" value={summary.costCentres||0} sub="Departments / projects"/><Metric icon={Scale} label="Budgets" value={summary.budgets?.n||0} sub={`Planned ${fmt(summary.budgets?.total)}`}/></div><div className="metric-grid" style={{marginTop:12}}><Metric icon={Activity} label="Intercompany records" value={summary.intercompany?.n||0} sub={`Recorded ${fmt(summary.intercompany?.total)}`}/><Metric icon={ShieldCheck} label="Approved eliminations" value={summary.approvedEliminations?.n||0} sub={`Workpaper total ${fmt(summary.approvedEliminations?.total)}`}/></div><div className="info-footer"><ShieldCheck size={16}/> Release 3 provides enterprise directories and reviewable workpapers. See README for boundaries: these registers do not automatically post cross-company journals or produce a certified consolidated financial statement.</div></>}
    {tab==='entities'&&<><form className="form-stack" onSubmit={e=>{e.preventDefault();save('/enterprise/entities',entityForm,'Legal entity added.',()=>setEntityForm({legalName:'',entityCode:'',currency:'INR',countryCode:'IN',taxIdentifier:''}));}}><h3>Add legal entity record</h3><div className="journal-form-grid"><label>Legal name<input value={entityForm.legalName} onChange={e=>setEntityForm({...entityForm,legalName:e.target.value})} required/></label><label>Entity code<input value={entityForm.entityCode} onChange={e=>setEntityForm({...entityForm,entityCode:e.target.value})} required placeholder="e.g. SSPL-IN"/></label><label>Currency<input maxLength="3" value={entityForm.currency} onChange={e=>setEntityForm({...entityForm,currency:e.target.value.toUpperCase()})} required/></label><label>Country code<input maxLength="2" value={entityForm.countryCode} onChange={e=>setEntityForm({...entityForm,countryCode:e.target.value.toUpperCase()})} required/></label><label>Tax identifier<input value={entityForm.taxIdentifier} onChange={e=>setEntityForm({...entityForm,taxIdentifier:e.target.value})}/></label></div><button className="button primary" disabled={busy}><Plus size={15}/> Add entity</button></form><hr/><form className="form-stack" onSubmit={e=>{e.preventDefault();save('/enterprise/branches',branchForm,'Branch added.',()=>setBranchForm({entityId:'',branchCode:'',name:'',address:'',managerName:''}));}}><h3>Add branch</h3><div className="journal-form-grid"><label>Legal entity<select value={branchForm.entityId} onChange={e=>setBranchForm({...branchForm,entityId:e.target.value})} required><option value="">Select entity</option>{entities.filter(x=>x.status==='active').map(x=><option key={x.id} value={x.id}>{x.entity_code} — {x.legal_name}</option>)}</select></label><label>Branch code<input value={branchForm.branchCode} onChange={e=>setBranchForm({...branchForm,branchCode:e.target.value})} required/></label><label>Branch name<input value={branchForm.name} onChange={e=>setBranchForm({...branchForm,name:e.target.value})} required/></label><label>Manager<input value={branchForm.managerName} onChange={e=>setBranchForm({...branchForm,managerName:e.target.value})}/></label><label>Address<input value={branchForm.address} onChange={e=>setBranchForm({...branchForm,address:e.target.value})}/></label></div><button className="button primary" disabled={busy}><Plus size={15}/> Add branch</button></form><h3>Legal entity directory</h3><div className="table-wrap"><table><thead><tr><th>Code</th><th>Legal name</th><th>Currency</th><th>Country</th><th>Tax identifier</th><th>Branches</th><th>Status</th></tr></thead><tbody>{entities.map(e=><tr key={e.id}><td>{e.entity_code}</td><td>{e.legal_name}</td><td>{e.currency}</td><td>{e.country_code}</td><td>{e.tax_identifier||'—'}</td><td>{e.branch_count||0}</td><td>{e.status}</td></tr>)}{!entities.length&&<tr><td colSpan="7">No legal entities added.</td></tr>}</tbody></table></div><h3>Branches</h3><div className="table-wrap"><table><thead><tr><th>Entity</th><th>Code</th><th>Branch</th><th>Manager</th><th>Address</th><th>Status</th></tr></thead><tbody>{branches.map(b=><tr key={b.id}><td>{b.entity_code} — {b.legal_name}</td><td>{b.branch_code}</td><td>{b.name}</td><td>{b.manager_name||'—'}</td><td>{b.address||'—'}</td><td>{b.status}</td></tr>)}{!branches.length&&<tr><td colSpan="6">No branches added.</td></tr>}</tbody></table></div></>}
    {tab==='centres'&&<><form className="form-stack" onSubmit={e=>{e.preventDefault();save('/enterprise/cost-centres',centreForm,'Cost centre created.',()=>setCentreForm({centreCode:'',name:'',department:'',ownerName:''}));}}><h3>Create cost centre / department / project</h3><div className="journal-form-grid"><label>Code<input value={centreForm.centreCode} onChange={e=>setCentreForm({...centreForm,centreCode:e.target.value})} required/></label><label>Name<input value={centreForm.name} onChange={e=>setCentreForm({...centreForm,name:e.target.value})} required/></label><label>Department / project<input value={centreForm.department} onChange={e=>setCentreForm({...centreForm,department:e.target.value})}/></label><label>Owner<input value={centreForm.ownerName} onChange={e=>setCentreForm({...centreForm,ownerName:e.target.value})}/></label></div><button className="button primary" disabled={busy}><Plus size={15}/> Save cost centre</button></form><div className="table-wrap"><table><thead><tr><th>Code</th><th>Name</th><th>Department / project</th><th>Owner</th><th>Status</th></tr></thead><tbody>{centres.map(c=><tr key={c.id}><td>{c.centre_code}</td><td>{c.name}</td><td>{c.department||'—'}</td><td>{c.owner_name||'—'}</td><td>{c.status}</td></tr>)}{!centres.length&&<tr><td colSpan="5">No cost centres created.</td></tr>}</tbody></table></div><div className="info-footer">Cost centres are maintained as enterprise master data. Existing journal forms do not yet tag every posted line to a cost centre; see README for the current reporting boundary.</div></>}
    {tab==='budgets'&&<><form className="form-stack" onSubmit={e=>{e.preventDefault();save('/enterprise/budgets',budgetForm,'Budget created as draft.',()=>setBudgetForm({...budgetForm,budgetName:''}));}}><h3>Create budget period</h3><div className="journal-form-grid"><label>Budget name<input value={budgetForm.budgetName} onChange={e=>setBudgetForm({...budgetForm,budgetName:e.target.value})} required/></label><label>Start date<input type="date" value={budgetForm.periodStart} onChange={e=>setBudgetForm({...budgetForm,periodStart:e.target.value})} required/></label><label>End date<input type="date" value={budgetForm.periodEnd} onChange={e=>setBudgetForm({...budgetForm,periodEnd:e.target.value})} required/></label></div><button className="button primary" disabled={busy}><Plus size={15}/> Create budget</button></form><form className="form-stack" onSubmit={e=>{e.preventDefault();save(`/enterprise/budgets/${lineForm.budgetId}/lines`,{accountId:lineForm.accountId,costCentreId:lineForm.costCentreId||null,plannedAmount:lineForm.plannedAmount,notes:lineForm.notes},'Budget line saved.',()=>setLineForm({...lineForm,plannedAmount:'',notes:''}));}}><h3>Add / update budget line</h3><div className="journal-form-grid"><label>Draft budget<select value={lineForm.budgetId} onChange={e=>setLineForm({...lineForm,budgetId:e.target.value})} required><option value="">Select draft budget</option>{budgets.filter(b=>b.status==='draft').map(b=><option key={b.id} value={b.id}>{b.budget_name}</option>)}</select></label><label>Expense / ledger account<select value={lineForm.accountId} onChange={e=>setLineForm({...lineForm,accountId:e.target.value})} required><option value="">Select account</option>{data.accounts.map(a=><option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></label><label>Cost centre<select value={lineForm.costCentreId} onChange={e=>setLineForm({...lineForm,costCentreId:e.target.value})}><option value="">No cost centre</option>{centres.map(c=><option key={c.id} value={c.id}>{c.centre_code} — {c.name}</option>)}</select></label><label>Planned amount (₹)<input type="number" min="0" step="0.01" value={lineForm.plannedAmount} onChange={e=>setLineForm({...lineForm,plannedAmount:e.target.value})} required/></label><label>Notes<input value={lineForm.notes} onChange={e=>setLineForm({...lineForm,notes:e.target.value})}/></label></div><button className="button primary" disabled={busy||!lineForm.budgetId}><Plus size={15}/> Save budget line</button></form><div className="table-wrap"><table><thead><tr><th>Budget</th><th>Period</th><th>Status</th><th>Lines</th><th>Planned total</th><th>Action</th></tr></thead><tbody>{budgets.map(b=><tr key={b.id}><td>{b.budget_name}</td><td>{String(b.period_start).slice(0,10)} — {String(b.period_end).slice(0,10)}</td><td>{b.status}</td><td>{b.line_count}</td><td>{fmt(b.planned_total)}</td><td>{b.status==='draft'&&['owner','admin'].includes(session.user.role)?<button className="button secondary small" onClick={async()=>{if(!confirm('Approve this budget? Approved budgets cannot be edited through this screen.'))return;try{await request(`/enterprise/budgets/${b.id}/approve`,session.token,{method:'POST',body:'{}'});setToast('Budget approved.');await loadAll();}catch(e){setError(e.message);}}}>Approve</button>:'—'}</td></tr>)}{!budgets.length&&<tr><td colSpan="6">No budgets created.</td></tr>}</tbody></table></div><div className="info-footer">Budget amounts are planning values. This version does not automatically enforce spending limits at transaction posting or produce cost-centre actual-vs-budget figures.</div></>}
    {tab==='intercompany'&&<><form className="form-stack" onSubmit={e=>{e.preventDefault();save('/enterprise/intercompany',icForm,'Intercompany transaction recorded.',()=>setIcForm({sourceEntityId:'',targetEntityId:'',transactionDate:today,reference:'',description:'',amount:'',currency:'INR'}));}}><h3>Record intercompany transaction</h3><div className="journal-form-grid"><label>From entity<select value={icForm.sourceEntityId} onChange={e=>setIcForm({...icForm,sourceEntityId:e.target.value})} required><option value="">Select source</option>{entities.map(x=><option key={x.id} value={x.id}>{x.entity_code} — {x.legal_name}</option>)}</select></label><label>To entity<select value={icForm.targetEntityId} onChange={e=>setIcForm({...icForm,targetEntityId:e.target.value})} required><option value="">Select target</option>{entities.map(x=><option key={x.id} value={x.id}>{x.entity_code} — {x.legal_name}</option>)}</select></label><label>Date<input type="date" value={icForm.transactionDate} onChange={e=>setIcForm({...icForm,transactionDate:e.target.value})} required/></label><label>Reference<input value={icForm.reference} onChange={e=>setIcForm({...icForm,reference:e.target.value})}/></label><label>Description<input value={icForm.description} onChange={e=>setIcForm({...icForm,description:e.target.value})} required/></label><label>Amount<input type="number" min="0.01" step="0.01" value={icForm.amount} onChange={e=>setIcForm({...icForm,amount:e.target.value})} required/></label><label>Currency<input maxLength="3" value={icForm.currency} onChange={e=>setIcForm({...icForm,currency:e.target.value.toUpperCase()})} required/></label></div><button className="button primary" disabled={busy}><Plus size={15}/> Record transaction</button></form><div className="table-wrap"><table><thead><tr><th>Date</th><th>Reference</th><th>From</th><th>To</th><th>Description</th><th>Amount</th><th>Status / action</th></tr></thead><tbody>{transactions.map(t=><tr key={t.id}><td>{String(t.transaction_date).slice(0,10)}</td><td>{t.reference||'—'}</td><td>{t.source_name}</td><td>{t.target_name}</td><td>{t.description}</td><td>{t.currency} {Number(t.amount).toFixed(2)}</td><td>{t.status==='draft'&&['owner','admin','accountant'].includes(session.user.role)?<div className="inline-actions"><button className="button secondary small" onClick={async()=>{try{await request(`/enterprise/intercompany/${t.id}/review`,session.token,{method:'POST',body:JSON.stringify({decision:'reviewed'})});await loadAll();}catch(e){setError(e.message);}}}>Mark reviewed</button><button className="button secondary small" onClick={()=>{setElimForm({...elimForm,transactionId:t.id,description:`Eliminate ${t.reference||t.description}`,amount:String(t.amount)});setTab('consolidation');}}>Prepare elimination</button></div>:t.status}</td></tr>)}{!transactions.length&&<tr><td colSpan="7">No intercompany records.</td></tr>}</tbody></table></div><div className="info-footer">These are intercompany control records only. Recording here does not create accounting journals in either legal entity; post corresponding entries in the respective books after review.</div></>}
    {tab==='consolidation'&&<><form className="form-stack" onSubmit={e=>{e.preventDefault();save('/enterprise/eliminations',elimForm,'Elimination workpaper created as draft.',()=>setElimForm({transactionId:'',eliminationDate:today,description:'',amount:''}));}}><h3>Create consolidation elimination workpaper</h3><div className="journal-form-grid"><label>Linked intercompany transaction<select value={elimForm.transactionId} onChange={e=>{const t=transactions.find(x=>x.id===e.target.value);setElimForm({...elimForm,transactionId:e.target.value,description:t?`Eliminate ${t.reference||t.description}`:elimForm.description,amount:t?String(t.amount):elimForm.amount});}}><option value="">Manual / unlinked adjustment</option>{transactions.map(t=><option key={t.id} value={t.id}>{t.reference||t.description} · {t.source_name} → {t.target_name}</option>)}</select></label><label>Elimination date<input type="date" value={elimForm.eliminationDate} onChange={e=>setElimForm({...elimForm,eliminationDate:e.target.value})} required/></label><label>Description<input value={elimForm.description} onChange={e=>setElimForm({...elimForm,description:e.target.value})} required/></label><label>Amount<input type="number" min="0.01" step="0.01" value={elimForm.amount} onChange={e=>setElimForm({...elimForm,amount:e.target.value})} required/></label></div><button className="button primary" disabled={busy}><Plus size={15}/> Save elimination workpaper</button></form><div className="table-wrap"><table><thead><tr><th>Date</th><th>Reference</th><th>Description</th><th>Amount</th><th>Status / action</th></tr></thead><tbody>{eliminations.map(x=><tr key={x.id}><td>{String(x.elimination_date).slice(0,10)}</td><td>{x.transaction_reference||'Manual'}</td><td>{x.description}</td><td>{fmt(x.amount)}</td><td>{x.status==='draft'&&['owner','admin'].includes(session.user.role)?<button className="button secondary small" onClick={async()=>{try{await request(`/enterprise/eliminations/${x.id}/approve`,session.token,{method:'POST',body:'{}'});setToast('Elimination workpaper approved.');await loadAll();}catch(e){setError(e.message);}}}>Approve workpaper</button>:x.status}</td></tr>)}{!eliminations.length&&<tr><td colSpan="5">No elimination workpapers.</td></tr>}</tbody></table></div><div className="info-footer">Approval records a reviewed workpaper only. It does not post consolidation journals or calculate a consolidated P&amp;L/balance sheet. Verify all eliminations and FX translation with your accountant.</div></>}
  </section>;
}


function App() {
  const [session, setSession] = useState(null);
  const [page, setPage] = useState("overview");
  const [mobileNav, setMobileNav] = useState(false);
  const [data, setData] = useState({ organization: null, company: null, years: [], accounts: [], journals: [], trialBalance: null, parties: [], invoices: [], purchaseBills: [] });
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState("");
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [accountForm, setAccountForm] = useState({ code: "", name: "", accountType: "expense" });
  const [formError, setFormError] = useState("");
  const [showJournalForm, setShowJournalForm] = useState(false);
  const [journalError, setJournalError] = useState("");
  const [journalBusy, setJournalBusy] = useState(false);
  const [journalForm, setJournalForm] = useState({
    entryDate: new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10), description: "", reference: "",
    lines: [
      { accountId: "", description: "", debit: "", credit: "" },
      { accountId: "", description: "", debit: "", credit: "" }
    ]
  });
  const today = new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  const [partyForm, setPartyForm] = useState({ partyType: "customer", name: "", gstin: "", email: "", phone: "", stateName: "Gujarat", address: "" });
  const [invoiceForm, setInvoiceForm] = useState({ customerId: "", invoiceDate: today, dueDate: "", description: "", taxableAmount: "", gstRate: "18" });
  const [partyError, setPartyError] = useState("");
  const [invoiceError, setInvoiceError] = useState("");
  const [showPartyForm, setShowPartyForm] = useState(false);
  const [showInvoiceForm, setShowInvoiceForm] = useState(false);
  const [receiptInvoice, setReceiptInvoice] = useState(null);
  const [receiptForm, setReceiptForm] = useState({ receiptDate: today, accountId: "", amount: "", reference: "" });
  const [receiptError, setReceiptError] = useState("");
  const [purchaseForm, setPurchaseForm] = useState({ supplierId: "", billNumber: "", billDate: today, dueDate: "", description: "", expenseAccountId: "", taxableAmount: "", gstRate: "18" });
  const [purchaseError, setPurchaseError] = useState("");
  const [showPurchaseForm, setShowPurchaseForm] = useState(false);
  const [paymentBill, setPaymentBill] = useState(null);
  const [supplierPaymentForm, setSupplierPaymentForm] = useState({ paymentDate: today, accountId: "", amount: "", reference: "" });
  const [supplierPaymentError, setSupplierPaymentError] = useState("");
  const [pageNumbers, setPageNumbers] = useState({});
  const [aiInsights, setAiInsights] = useState(null);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState("");
  const [assistantQuestion, setAssistantQuestion] = useState("");
  const [assistantMessages, setAssistantMessages] = useState([]);
  const [assistantLoading, setAssistantLoading] = useState(false);
  const [assistantError, setAssistantError] = useState("");
  const [financialReports, setFinancialReports] = useState(null);
  const [reportsLoading, setReportsLoading] = useState(false);
  const [reportsError, setReportsError] = useState("");
  const pageSize = 8;
  function paginated(rows, key) { const pageNo = Math.min(pageNumbers[key] || 1, Math.max(1, Math.ceil(rows.length / pageSize))); return rows.slice((pageNo - 1) * pageSize, pageNo * pageSize); }
  function pagination(key, total) { const pages = Math.max(1, Math.ceil(total / pageSize)); const current = Math.min(pageNumbers[key] || 1, pages); return total > pageSize ? <div className="pagination-bar"><span>{(current - 1) * pageSize + 1}–{Math.min(current * pageSize, total)} of {total}</span><div><button className="button secondary small" disabled={current <= 1} onClick={() => setPageNumbers(p => ({...p, [key]: current - 1}))}>Previous</button><span className="pagination-page">Page {current} / {pages}</span><button className="button secondary small" disabled={current >= pages} onClick={() => setPageNumbers(p => ({...p, [key]: current + 1}))}>Next</button></div></div> : null; }
  function exportRows() {
    const money = v => Number(v || 0);
    const rows = {
      overview: [{Metric:"Organization",Value:data.organization?.name || ""},{Metric:"Company",Value:company?.legal_name || ""},{Metric:"Accounts",Value:data.accounts.length},{Metric:"Financial years",Value:data.years.length}],
      company: Object.entries({Organization:data.organization?.name, "Legal company name":company?.legal_name, "Trade name":company?.trade_name, Currency:company?.base_currency, Country:company?.country_code, State:company?.state_name, GSTIN:company?.gstin}).map(([Field,Value])=>({Field,Value:Value || ""})),
      years: data.years.map(y=>({"Financial year":y.name,"Start date":String(y.start_date).slice(0,10),"End date":String(y.end_date).slice(0,10),Status:y.status})),
      accounts: filteredAccounts.map(a=>({Code:a.code,Account:a.name,Type:a.account_type,"Normal balance":a.normal_balance,Status:a.is_active?"Active":"Inactive"})),
      journals: data.journals.map(j=>({Entry:`JV-${j.entry_number}`,Date:String(j.entry_date).slice(0,10),Description:j.description,Reference:j.reference||"",Debit:money(j.total_debit),Credit:money(j.total_credit),Status:j.status})),
      parties: data.parties.map(p=>({Name:p.name,Type:p.party_type,GSTIN:p.gstin||"",Email:p.email||"",Phone:p.phone||"",State:p.state_name||"",Status:p.is_active?"Active":"Inactive"})),
      "sales-invoices": data.invoices.map(i=>({Invoice:i.invoice_number,Date:String(i.invoice_date).slice(0,10),Customer:i.customer_name,Taxable:money(i.taxable_amount),GST:`${i.gst_rate}%`,Total:money(i.total_amount),Received:money(i.amount_received),Balance:money(i.total_amount)-money(i.amount_received),Status:i.status})),
      "purchase-bills": data.purchaseBills.map(b=>({Bill:b.bill_number,Date:String(b.bill_date).slice(0,10),Supplier:b.supplier_name,Account:`${b.expense_code} — ${b.expense_account_name}`,Taxable:money(b.taxable_amount),GST:`${b.gst_rate}%`,Total:money(b.total_amount),Paid:money(b.amount_paid),Balance:money(b.total_amount)-money(b.amount_paid),Status:b.status})),
      "trial-balance": (data.trialBalance?.rows||[]).map(r=>({Code:r.code,Account:r.name,Type:r.account_type,Debit:money(r.total_debit),Credit:money(r.total_credit),Balance:money(r.balance),Side:r.balance_side})),
      "financial-reports": (financialReports?.incomeStatement||[]).map(r=>({Code:r.code,Account:r.name,Type:r.type,Amount:Number(r.amount)})),
      "ai-insights": (aiInsights?.insights||[]).map((x,i)=>({No:i+1,Insight:typeof x === "string"?x:(x.title||x.insight||JSON.stringify(x)),Detail:typeof x === "object"?(x.detail||x.recommendation||""):""})),
      "ai-assistant": assistantMessages.map((m,i)=>({No:i+1,Role:m.role,Message:m.content,Engine:m.engine||""}))
    };
    return rows[page] || [{Information:"No exportable rows on this page"}];
  }
  function exportExcel() { const ws=XLSX.utils.json_to_sheet(exportRows()); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,(currentPage||"Report").slice(0,31)); XLSX.writeFile(wb, `SoftSpark_${page}_${today}.xlsx`); }
  function exportPDF() { const doc=new jsPDF({orientation:"landscape"}); doc.setFontSize(16); doc.text(`SoftSpark AutoBooks Business - ${currentPage}`,14,16); doc.setFontSize(9); doc.text(`Company: ${company?.legal_name || "SoftSpark AutoBooks Business"}   Generated: ${today}`,14,23); const rows=exportRows(); const columns=Object.keys(rows[0]||{Information:""}); autoTable(doc,{startY:29,head:[columns],body:rows.map(r=>columns.map(c=>String(r[c]??""))),styles:{fontSize:7,cellPadding:2},headStyles:{fillColor:[31,67,116]}}); doc.save(`AutoBooks_Business_${page}_${today}.pdf`); }
  async function loadFinancialReports(financialYearId = "") { setReportsLoading(true); setReportsError(""); try { const result=await request(`/reports/financial-statements${financialYearId ? `?financialYearId=${encodeURIComponent(financialYearId)}` : ""}`,session.token); setFinancialReports(result); } catch(err) { setReportsError(err.message); } finally { setReportsLoading(false); } }
  async function loadAiInsights() { setAiLoading(true); setAiError(""); try { const result=await request("/ai/insights",session.token); setAiInsights(result); } catch(err) { setAiError(err.message); } finally { setAiLoading(false); } }
  async function askAssistant(e) {
    e?.preventDefault?.();
    const question = assistantQuestion.trim();
    if (!question || assistantLoading) return;
    const prior = assistantMessages.slice(-6).map(m=>({role:m.role,content:m.content}));
    setAssistantMessages(messages=>[...messages,{role:"user",content:question}]);
    setAssistantQuestion(""); setAssistantError(""); setAssistantLoading(true);
    try {
      const result = await request("/ai/ask", session.token, {method:"POST",body:JSON.stringify({question,history:prior})});
      setAssistantMessages(messages=>[...messages,{role:"assistant",content:result.answer,engine:result.engine,generatedAt:result.generatedAt}]);
    } catch(err) {
      setAssistantError(err.message);
      setAssistantMessages(messages=>messages.filter((m,i)=>!(i===messages.length-1 && m.role==="user" && m.content===question)));
      setAssistantQuestion(question);
    } finally { setAssistantLoading(false); }
  }

  function logout() {
    setSession(null);
    setData({ organization: null, company: null, years: [], accounts: [], journals: [], trialBalance: null, parties: [], invoices: [], purchaseBills: [] });
    setPage("overview");
  }

  async function load(token = session?.token) {
    if (!token) return;
    setLoading(true);
    try {
      const [org, company, years, accounts, journals, trialBalance, parties, invoices, purchaseBills] = await Promise.all([
        request("/organization", token),
        request("/company", token),
        request("/financial-years", token),
        request("/accounts", token),
        request("/journals", token),
        request("/reports/trial-balance", token),
        request("/parties", token),
        request("/sales-invoices", token),
        request("/purchase-bills", token)
      ]);
      setData({ organization: org.organization, company: company.company, years: years.financialYears, accounts: accounts.accounts, journals: journals.journals, trialBalance, parties: parties.parties, invoices: invoices.invoices, purchaseBills: purchaseBills.purchaseBills });
    } catch (err) {
      setToast(err.message);
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { if (session?.token) load(session.token); }, [session?.token]);
  useEffect(() => { if (session?.token && page === "financial-reports") loadFinancialReports(); }, [session?.token, page]);

  async function createAccount(e) {
    e.preventDefault(); setFormError("");
    try {
      await request("/accounts", session.token, { method: "POST", body: JSON.stringify(accountForm) });
      setShowAdd(false);
      setAccountForm({ code: "", name: "", accountType: "expense" });
      await load();
      setToast("Account created successfully.");
    } catch (err) {
      setFormError(err.message);
    }
  }

  async function createJournal(e) {
    e.preventDefault(); setJournalError(""); setJournalBusy(true);
    const lines = journalForm.lines.map(line => ({
      accountId: line.accountId,
      description: line.description,
      debit: Number(line.debit || 0),
      credit: Number(line.credit || 0)
    }));
    try {
      await request("/journals", session.token, { method: "POST", body: JSON.stringify({ ...journalForm, lines }) });
      setShowJournalForm(false);
      setJournalForm({ entryDate: new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10), description: "", reference: "", lines: [
        { accountId: "", description: "", debit: "", credit: "" },
        { accountId: "", description: "", debit: "", credit: "" }
      ] });
      await load();
      setToast("Journal posted successfully. Ledger and trial balance have been updated.");
    } catch (err) { setJournalError(err.message); }
    finally { setJournalBusy(false); }
  }

  async function createParty(e) {
    e.preventDefault(); setPartyError("");
    try {
      await request("/parties", session.token, { method: "POST", body: JSON.stringify(partyForm) });
      setPartyForm({ partyType: "customer", name: "", gstin: "", email: "", phone: "", stateName: "Gujarat", address: "" });
      setShowPartyForm(false); await load(); setToast("Customer/supplier saved successfully.");
    } catch (err) { setPartyError(err.message); }
  }

  async function createInvoice(e) {
    e.preventDefault(); setInvoiceError("");
    try {
      const result = await request("/sales-invoices", session.token, { method: "POST", body: JSON.stringify(invoiceForm) });
      setShowInvoiceForm(false); setInvoiceForm({ customerId: "", invoiceDate: today, dueDate: "", description: "", taxableAmount: "", gstRate: "18" });
      await load(); setToast(`Invoice ${result.invoice.invoice_number} posted. Total ₹${Number(result.invoice.total_amount).toFixed(2)}.`);
    } catch (err) { setInvoiceError(err.message); }
  }

  async function postReceipt(e) {
    e.preventDefault(); setReceiptError("");
    try {
      const result = await request(`/sales-invoices/${receiptInvoice.id}/receipts`, session.token, { method: "POST", body: JSON.stringify(receiptForm) });
      setReceiptInvoice(null); setReceiptForm({ receiptDate: today, accountId: "", amount: "", reference: "" }); await load(); setToast(`Receipt ${result.receipt.receipt_number} posted successfully.`);
    } catch (err) { setReceiptError(err.message); }
  }

  async function createPurchaseBill(e) {
    e.preventDefault(); setPurchaseError("");
    try {
      const result = await request("/purchase-bills", session.token, { method: "POST", body: JSON.stringify(purchaseForm) });
      setShowPurchaseForm(false); setPurchaseForm({ supplierId: "", billNumber: "", billDate: today, dueDate: "", description: "", expenseAccountId: "", taxableAmount: "", gstRate: "18" });
      await load(); setToast(`Purchase bill ${result.purchaseBill.bill_number} posted. Total ₹${Number(result.purchaseBill.total_amount).toFixed(2)}.`);
    } catch (err) { setPurchaseError(err.message); }
  }

  async function postSupplierPayment(e) {
    e.preventDefault(); setSupplierPaymentError("");
    try {
      const result = await request(`/purchase-bills/${paymentBill.id}/payments`, session.token, { method: "POST", body: JSON.stringify(supplierPaymentForm) });
      setPaymentBill(null); setSupplierPaymentForm({ paymentDate: today, accountId: "", amount: "", reference: "" }); await load(); setToast(`Supplier payment ${result.payment.payment_number} posted successfully.`);
    } catch (err) { setSupplierPaymentError(err.message); }
  }

  function updateJournalLine(index, field, value) {
    setJournalForm(current => ({ ...current, lines: current.lines.map((line, i) => i === index ? { ...line, [field]: value } : line) }));
  }

  const filteredAccounts = useMemo(() => {
    const term = search.toLowerCase();
    return data.accounts.filter(a => `${a.code} ${a.name} ${a.account_type}`.toLowerCase().includes(term));
  }, [data.accounts, search]);

  if (!session) return <Login onLogin={setSession} />;

  const company = data.company;
  const user = session.user;
  const currentPage = nav.find(n => n.id === page)?.label || "Overview";

  return (
    <div className="app-shell">
      {mobileNav && <button className="mobile-scrim" onClick={() => setMobileNav(false)} aria-label="Close navigation" />}
      <aside className={`sidebar ${mobileNav ? "sidebar-open" : ""}`}>
        <div className="brand-row">
          <div className="brand-mark">S</div>
          <div className="brand-name">SoftSpark <span>AutoBooks Enterprise</span></div>
          <button className="icon-button mobile-close" onClick={() => setMobileNav(false)}><X size={18}/></button>
        </div>
        <div className="workspace">
          <div className="workspace-icon"><Building2 size={18}/></div>
          <div className="workspace-text"><strong>{data.organization?.name || "Loading workspace"}</strong><span>India · INR</span></div>
          <ChevronDown size={15} className="muted-icon"/>
        </div>
        <div className="nav-caption">WORKSPACE</div>
        <nav>
          {nav.filter(item => item.id !== "team" || ["owner","admin","accountant"].includes(session.user.role)).map(item => {
            const Icon = item.icon;
            return <button key={item.id} className={`nav-item ${page === item.id ? "active" : ""}`} onClick={() => { setPage(item.id); setMobileNav(false); }}>
              <Icon size={18}/><span>{item.label}</span>{item.id === "accounts" && <span className="nav-count">{data.accounts.length}</span>}
            </button>;
          })}
        </nav>
        <div className="sidebar-bottom">
          <button className="nav-item" onClick={() => setToast("Open README.md in the project folder for the complete AutoBooks Business setup and user guide.")}><CircleHelp size={18}/><span>Help & guidance</span></button>
          <div className="user-panel">
            <div className="avatar">{(user.fullName || "S").slice(0,1).toUpperCase()}</div>
            <div className="user-details"><strong>{user.fullName}</strong><span>{user.role.replace("_"," ")}</span></div>
            <button className="icon-button" onClick={logout} title="Log out"><LogOut size={17}/></button>
          </div>
        </div>
      </aside>

      <main className="main-area">
        <header className="topbar">
          <div className="topbar-left">
            <button className="icon-button menu-button" onClick={() => setMobileNav(true)}><Menu size={20}/></button>
            <div><div className="breadcrumb">Workspace <span>/</span> {currentPage}</div><div className="page-title">{currentPage}</div></div>
          </div>
          <div className="topbar-actions">
            <div className="environment"><span/> Release 3 · Enterprise</div>
            <button className="icon-button" title="Refresh data" onClick={() => load()}><RefreshCw size={17}/></button>
            <div className="top-avatar">{(user.fullName || "S").slice(0,1).toUpperCase()}</div>
          </div>
        </header>

        <div className="content">
          {loading && <div className="loading-line">Refreshing workspace data…</div>}
          {toast && <div className="toast"><CheckCircle2 size={17}/><span>{toast}</span><button onClick={() => setToast("")}><X size={16}/></button></div>}
          <div className="export-toolbar"><span>{currentPage} report</span><div><button className="button secondary small" onClick={exportPDF}><FileText size={15}/> Export PDF</button><button className="button secondary small" onClick={exportExcel}><FileText size={15}/> Export Excel</button></div></div>

          {page === "overview" && <>
            <section className="welcome-row">
              <div><div className="eyebrow">YOUR ACCOUNTING WORKSPACE</div><h1>Good work starts with clarity.</h1><p className="muted">Manage company records, transactions, reports and financial insights.</p></div>
              <button className="button secondary" onClick={() => setPage("company")}><Building2 size={17}/> Review company setup</button>
            </section>
            <section className="metric-grid">
              <Metric icon={Building2} label="Organization" value={data.organization ? "Active" : "—"} sub={data.organization?.name || "Loading"} tone="blue"/>
              <Metric icon={BookOpen} label="Accounts configured" value={data.accounts.length} sub="Default and custom accounts" tone="violet"/>
              <Metric icon={CalendarDays} label="Financial years" value={data.years.length} sub={data.years[0]?.name || "No financial year"} tone="green"/>
              <Metric icon={Wallet} label="Base currency" value={company?.base_currency || "INR"} sub="Company reporting currency" tone="amber"/>
            </section>
            <div className="section-heading"><div><h2>Setup checklist</h2><p>Review the foundation before proceeding to transaction accounting.</p></div><span className="status-pill"><CheckCircle2 size={14}/> Workspace loaded</span></div>
            <section className="setup-grid">
              <SetupCard icon={Building2} number="01" title="Company profile" description="Confirm your legal name, trade name, state and GST registration details." status={company ? "Configured" : "Needs setup"} action={() => setPage("company")}/>
              <SetupCard icon={CalendarDays} number="02" title="Financial year" description="Check the financial year dates and open/closed status." status={data.years.length ? "Configured" : "Needs setup"} action={() => setPage("years")}/>
              <SetupCard icon={BookOpen} number="03" title="Chart of accounts" description="Review the account categories that future journals will use." status={data.accounts.length ? `${data.accounts.length} accounts` : "Needs setup"} action={() => setPage("accounts")}/>
            </section>
          </>}

          {page === "company" && <section className="panel">
            <div className="panel-header"><div><h2>Company profile</h2><p>Organization and company information seeded for local testing.</p></div><span className="status-pill"><CheckCircle2 size={14}/> Company details</span></div>
            {company ? <div className="detail-grid">
              <Detail label="Organization" value={data.organization?.name}/>
              <Detail label="Legal company name" value={company.legal_name}/>
              <Detail label="Trade name" value={company.trade_name || "Not specified"}/>
              <Detail label="Base currency" value={company.base_currency}/>
              <Detail label="Country" value={company.country_code === "IN" ? "India" : company.country_code}/>
              <Detail label="State" value={company.state_name || "Not specified"}/>
              <Detail label="GSTIN" value={company.gstin || "Not configured"}/>
              <Detail label="Company ID" value={company.id}/>
            </div> : <div className="empty-state">No company found. Run the seed command and refresh.</div>}
          </section>}

          {page === "years" && <section className="panel">
            <div className="panel-header"><div><h2>Financial years</h2><p>Financial-year boundaries for Indian business reporting.</p></div><span className="status-pill">{data.years.length} configured</span></div>
            {data.years.length ? <><div className="table-wrap"><table><thead><tr><th>Financial year</th><th>Start date</th><th>End date</th><th>Status</th></tr></thead><tbody>{paginated(data.years, "years").map(y => <tr key={y.id}><td><strong>{y.name}</strong></td><td>{String(y.start_date).slice(0,10)}</td><td>{String(y.end_date).slice(0,10)}</td><td><span className={`tag ${y.status}`}>{y.status}</span></td></tr>)}</tbody></table></div>{pagination("years", data.years.length)}</> : <div className="empty-state">No financial years configured.</div>}
          </section>}

          {page === "accounts" && <section className="panel">
            <div className="panel-header"><div><h2>Chart of accounts</h2><p>Default account categories for the company ledger.</p></div><button className="button primary" onClick={() => { setShowAdd(true); setFormError(""); }}><Plus size={17}/> Add account</button></div>
            <div className="table-toolbar"><div className="search-field"><Search size={17}/><input placeholder="Search code, account name or type…" value={search} onChange={e => setSearch(e.target.value)}/></div><span>{filteredAccounts.length} accounts</span></div>
            <div className="table-wrap"><table><thead><tr><th>Code</th><th>Account name</th><th>Account type</th><th>Normal balance</th><th>Status</th></tr></thead><tbody>
              {paginated(filteredAccounts, "accounts").map(a => <tr key={a.id}><td className="code-cell">{a.code}</td><td><strong>{a.name}</strong>{a.is_system && <span className="system-tag">DEFAULT</span>}</td><td><span className={`type-dot ${a.account_type}`}/>{capitalize(a.account_type)}</td><td>{capitalize(a.normal_balance)}</td><td><span className={`tag ${a.is_active ? "open" : "closed"}`}>{a.is_active ? "Active" : "Inactive"}</span></td></tr>)}
            </tbody></table>{!filteredAccounts.length && <div className="empty-state">No accounts match your search.</div>}</div>{pagination("accounts", filteredAccounts.length)}
            <div className="info-footer"><BookOpen size={17}/> Account balances are calculated from posted journal entries.</div>
          </section>}
          {page === "journals" && <section className="panel">
            <div className="panel-header"><div><h2>Journal entries</h2><p>Every posted entry must balance and fall inside an open financial year.</p></div><button className="button primary" onClick={() => { setShowJournalForm(v => !v); setJournalError(""); }}><Plus size={17}/>{showJournalForm ? "Close form" : "New journal"}</button></div>
            {showJournalForm && <form className="journal-form form-stack" onSubmit={createJournal}>
              <div className="journal-form-grid">
                <label>Entry date<input type="date" value={journalForm.entryDate} onChange={e => setJournalForm({ ...journalForm, entryDate: e.target.value })} required /></label>
                <label>Reference (optional)<input value={journalForm.reference} onChange={e => setJournalForm({ ...journalForm, reference: e.target.value })} maxLength={120} placeholder="Invoice / voucher no." /></label>
                <label className="journal-description">Description<input value={journalForm.description} onChange={e => setJournalForm({ ...journalForm, description: e.target.value })} required minLength={3} maxLength={500} placeholder="Describe the transaction" /></label>
              </div>
              <div className="section-heading"><div><h3>Journal lines</h3><p>Each line must have either a debit or a credit amount.</p></div><button type="button" className="button secondary" onClick={() => setJournalForm({ ...journalForm, lines: [...journalForm.lines, { accountId: "", description: "", debit: "", credit: "" }] })} disabled={journalForm.lines.length >= 100}><Plus size={15}/> Add line</button></div>
              <div className="table-wrap"><table><thead><tr><th>Account</th><th>Line description</th><th>Debit (₹)</th><th>Credit (₹)</th><th></th></tr></thead><tbody>
                {journalForm.lines.map((line, index) => <tr key={index}><td><select aria-label={`Account line ${index + 1}`} value={line.accountId} onChange={e => updateJournalLine(index, "accountId", e.target.value)} required><option value="">Select account</option>{data.accounts.filter(a => a.is_active).map(a => <option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></td><td><input value={line.description} onChange={e => updateJournalLine(index, "description", e.target.value)} placeholder="Optional" maxLength={300}/></td><td><input type="number" min="0" step="0.01" value={line.debit} onChange={e => updateJournalLine(index, "debit", e.target.value)} placeholder="0.00" /></td><td><input type="number" min="0" step="0.01" value={line.credit} onChange={e => updateJournalLine(index, "credit", e.target.value)} placeholder="0.00" /></td><td><button type="button" className="icon-button" title="Remove line" disabled={journalForm.lines.length <= 2} onClick={() => setJournalForm({ ...journalForm, lines: journalForm.lines.filter((_, i) => i !== index) })}><X size={16}/></button></td></tr>)}
              </tbody></table></div>
              <div className="journal-totals"><span>Debits: ₹{journalForm.lines.reduce((sum, line) => sum + Number(line.debit || 0), 0).toFixed(2)}</span><span>Credits: ₹{journalForm.lines.reduce((sum, line) => sum + Number(line.credit || 0), 0).toFixed(2)}</span></div>
              {journalError && <div className="alert error">{journalError}</div>}
              <div className="modal-actions"><button type="button" className="button secondary" onClick={() => setShowJournalForm(false)}>Cancel</button><button className="button primary" disabled={journalBusy}>{journalBusy ? "Posting…" : "Post journal"}</button></div>
              <div className="small-note">Posting is final. Verify the date, accounts and amounts before posting.</div>
            </form>}
            <div className="table-toolbar"><span>{data.journals.length} recent entries</span><button className="button secondary" onClick={() => load()}><RefreshCw size={15}/> Refresh</button></div>
            {data.journals.length ? <><div className="table-wrap"><table><thead><tr><th>Entry #</th><th>Date</th><th>Description</th><th>Reference</th><th>Debit</th><th>Credit</th><th>Status</th></tr></thead><tbody>{paginated(data.journals, "journals").map(j => <tr key={j.id}><td className="code-cell">JV-{j.entry_number}</td><td>{String(j.entry_date).slice(0,10)}</td><td><strong>{j.description}</strong></td><td>{j.reference || "—"}</td><td>₹{Number(j.total_debit).toFixed(2)}</td><td>₹{Number(j.total_credit).toFixed(2)}</td><td><span className={`tag ${j.status === "posted" ? "open" : "closed"}`}>{j.status}</span></td></tr>)}</tbody></table></div>{pagination("journals", data.journals.length)}</> : <div className="empty-state">No journal entries yet. Select “New journal” to post your first balanced transaction.</div>}
          </section>}

          {page === "parties" && <section className="panel">
            <div className="panel-header"><div><h2>Customers & suppliers</h2><p>Maintain business contacts for sales and purchasing workflows.</p></div><button className="button primary" onClick={() => { setShowPartyForm(v => !v); setPartyError(""); }}><Plus size={17}/>{showPartyForm ? "Close form" : "Add contact"}</button></div>
            {showPartyForm && <form className="journal-form form-stack" onSubmit={createParty}>
              <div className="journal-form-grid"><label>Contact type<select value={partyForm.partyType} onChange={e => setPartyForm({...partyForm,partyType:e.target.value})}><option value="customer">Customer</option><option value="supplier">Supplier</option><option value="both">Customer & supplier</option></select></label><label>Business / person name<input value={partyForm.name} onChange={e => setPartyForm({...partyForm,name:e.target.value})} required minLength={2} maxLength={200}/></label><label>GSTIN (optional)<input value={partyForm.gstin} onChange={e => setPartyForm({...partyForm,gstin:e.target.value.toUpperCase()})} maxLength={15} placeholder="15-character GSTIN"/></label><label>Email<input type="email" value={partyForm.email} onChange={e => setPartyForm({...partyForm,email:e.target.value})}/></label><label>Phone<input value={partyForm.phone} onChange={e => setPartyForm({...partyForm,phone:e.target.value})}/></label><label>State<input value={partyForm.stateName} onChange={e => setPartyForm({...partyForm,stateName:e.target.value})} placeholder="e.g. Gujarat"/></label><label className="journal-description">Address<input value={partyForm.address} onChange={e => setPartyForm({...partyForm,address:e.target.value})}/></label></div>
              {partyError && <div className="alert error">{partyError}</div>}<div className="modal-actions"><button type="button" className="button secondary" onClick={() => setShowPartyForm(false)}>Cancel</button><button className="button primary"><Plus size={16}/> Save contact</button></div>
            </form>}
            <div className="table-toolbar"><span>{data.parties.length} contacts</span><button className="button secondary" onClick={() => load()}><RefreshCw size={15}/> Refresh</button></div>
            {data.parties.length ? <><div className="table-wrap"><table><thead><tr><th>Name</th><th>Type</th><th>GSTIN</th><th>Email</th><th>Phone</th><th>State</th><th>Status</th></tr></thead><tbody>{paginated(data.parties, "parties").map(p => <tr key={p.id}><td><strong>{p.name}</strong></td><td>{capitalize(p.party_type)}</td><td>{p.gstin || "—"}</td><td>{p.email || "—"}</td><td>{p.phone || "—"}</td><td>{p.state_name || "—"}</td><td><span className={`tag ${p.is_active ? "open" : "closed"}`}>{p.is_active ? "Active" : "Inactive"}</span></td></tr>)}</tbody></table></div>{pagination("parties", data.parties.length)}</> : <div className="empty-state">No contacts yet. Add a customer to create your first sales invoice.</div>}
            <div className="info-footer"><Users size={17}/> GSTIN format is checked for basic 15-character structure only; this is not a GST portal validation.</div>
          </section>}

          {page === "sales-invoices" && <section className="panel">
            <div className="panel-header"><div><h2>Sales invoices</h2><p>Post sales with supported GST rates and automatic accounting entries.</p></div><button className="button primary" onClick={() => { setShowInvoiceForm(v => !v); setInvoiceError(""); }}><Plus size={17}/>{showInvoiceForm ? "Close form" : "New invoice"}</button></div>
            {showInvoiceForm && <form className="journal-form form-stack" onSubmit={createInvoice}>
              <div className="journal-form-grid"><label>Customer<select value={invoiceForm.customerId} onChange={e => setInvoiceForm({...invoiceForm,customerId:e.target.value})} required><option value="">Select customer</option>{data.parties.filter(p => p.is_active && ["customer","both"].includes(p.party_type)).map(p => <option key={p.id} value={p.id}>{p.name}{p.state_name ? ` — ${p.state_name}` : ""}</option>)}</select></label><label>Invoice date<input type="date" value={invoiceForm.invoiceDate} onChange={e => setInvoiceForm({...invoiceForm,invoiceDate:e.target.value})} required/></label><label>Due date<input type="date" value={invoiceForm.dueDate} onChange={e => setInvoiceForm({...invoiceForm,dueDate:e.target.value})}/></label><label>Taxable amount (₹)<input type="number" min="0.01" step="0.01" value={invoiceForm.taxableAmount} onChange={e => setInvoiceForm({...invoiceForm,taxableAmount:e.target.value})} required/></label><label>GST rate<select value={invoiceForm.gstRate} onChange={e => setInvoiceForm({...invoiceForm,gstRate:e.target.value})}><option value="0">0%</option><option value="5">5%</option><option value="12">12%</option><option value="18">18%</option><option value="28">28%</option></select></label><label className="journal-description">Description / service<input value={invoiceForm.description} onChange={e => setInvoiceForm({...invoiceForm,description:e.target.value})} required minLength={3} maxLength={500} placeholder="e.g. Software development services"/></label></div>
              <div className="small-note">Tax is split into CGST + SGST when customer and company states match (or state is unspecified); otherwise IGST is applied. Verify GST treatment before posting. Invoice posting is final in this build.</div>
              {invoiceError && <div className="alert error">{invoiceError}</div>}<div className="modal-actions"><button type="button" className="button secondary" onClick={() => setShowInvoiceForm(false)}>Cancel</button><button className="button primary"><FilePlus2 size={16}/> Post invoice</button></div>
            </form>}
            <div className="table-toolbar"><span>{data.invoices.length} invoices</span><button className="button secondary" onClick={() => load()}><RefreshCw size={15}/> Refresh</button></div>
            {data.invoices.length ? <><div className="table-wrap"><table><thead><tr><th>Invoice</th><th>Date</th><th>Customer</th><th>Taxable</th><th>GST</th><th>Total</th><th>Received</th><th>Balance</th><th>Status</th><th></th></tr></thead><tbody>{paginated(data.invoices, "sales-invoices").map(i => <tr key={i.id}><td className="code-cell">{i.invoice_number}</td><td>{String(i.invoice_date).slice(0,10)}</td><td><strong>{i.customer_name}</strong></td><td>₹{Number(i.taxable_amount).toFixed(2)}</td><td>{i.gst_rate}%</td><td>₹{Number(i.total_amount).toFixed(2)}</td><td>₹{Number(i.amount_received).toFixed(2)}</td><td>₹{(Number(i.total_amount)-Number(i.amount_received)).toFixed(2)}</td><td><span className={`tag ${i.status === "paid" ? "open" : "closed"}`}>{i.status.replace("_"," ")}</span></td><td>{i.status !== "paid" && i.status !== "cancelled" && <button className="button secondary" onClick={() => { setReceiptInvoice(i); setReceiptForm({receiptDate:today,accountId:data.accounts.find(a => a.is_active && ["1000","1010"].includes(a.code))?.id || "",amount:(Number(i.total_amount)-Number(i.amount_received)).toFixed(2),reference:""}); setReceiptError(""); }}>Record receipt</button>}</td></tr>)}</tbody></table></div>{pagination("sales-invoices", data.invoices.length)}</> : <div className="empty-state">No sales invoices yet. Add a customer first, then create an invoice.</div>}
            <div className="info-footer"><IndianRupee size={17}/> Invoice posting automatically creates a journal entry. Receipts debit the selected cash/bank account and credit Accounts Receivable.</div>
          </section>}

          {page === "purchase-bills" && <section className="panel">
            <div className="panel-header"><div><h2>Purchase bills</h2><p>Record supplier bills, supported GST rates, expense allocation and supplier payments.</p></div><button className="button primary" onClick={() => { setShowPurchaseForm(v => !v); setPurchaseError(""); }}><Plus size={17}/>{showPurchaseForm ? "Close form" : "New purchase bill"}</button></div>
            {showPurchaseForm && <form className="journal-form form-stack" onSubmit={createPurchaseBill}>
              <div className="journal-form-grid"><label>Supplier<select value={purchaseForm.supplierId} onChange={e => setPurchaseForm({...purchaseForm,supplierId:e.target.value})} required><option value="">Select supplier</option>{data.parties.filter(p => p.is_active && ["supplier","both"].includes(p.party_type)).map(p => <option key={p.id} value={p.id}>{p.name}{p.state_name ? ` — ${p.state_name}` : ""}</option>)}</select></label><label>Supplier bill number (optional)<input value={purchaseForm.billNumber} onChange={e => setPurchaseForm({...purchaseForm,billNumber:e.target.value})} maxLength={60} placeholder="Supplier's invoice number"/></label><label>Bill date<input type="date" value={purchaseForm.billDate} onChange={e => setPurchaseForm({...purchaseForm,billDate:e.target.value})} required/></label><label>Due date<input type="date" value={purchaseForm.dueDate} onChange={e => setPurchaseForm({...purchaseForm,dueDate:e.target.value})}/></label><label>Expense account<select value={purchaseForm.expenseAccountId} onChange={e => setPurchaseForm({...purchaseForm,expenseAccountId:e.target.value})} required><option value="">Select expense account</option>{data.accounts.filter(a => a.is_active && a.account_type === "expense").map(a => <option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></label><label>Taxable amount (₹)<input type="number" min="0.01" step="0.01" value={purchaseForm.taxableAmount} onChange={e => setPurchaseForm({...purchaseForm,taxableAmount:e.target.value})} required/></label><label>GST rate<select value={purchaseForm.gstRate} onChange={e => setPurchaseForm({...purchaseForm,gstRate:e.target.value})}><option value="0">0%</option><option value="5">5%</option><option value="12">12%</option><option value="18">18%</option><option value="28">28%</option></select></label><label className="journal-description">Description<input value={purchaseForm.description} onChange={e => setPurchaseForm({...purchaseForm,description:e.target.value})} required minLength={3} maxLength={500} placeholder="e.g. Office supplies or software subscription"/></label></div>
              <div className="small-note">Posting debits the selected expense account and GST Input Tax Credit account (1150), and credits Accounts Payable (2000). Verify input-tax eligibility and GST treatment with your accountant before posting.</div>
              {purchaseError && <div className="alert error">{purchaseError}</div>}<div className="modal-actions"><button type="button" className="button secondary" onClick={() => setShowPurchaseForm(false)}>Cancel</button><button className="button primary"><FilePlus2 size={16}/> Post purchase bill</button></div>
            </form>}
            <div className="table-toolbar"><span>{data.purchaseBills.length} purchase bills</span><button className="button secondary" onClick={() => load()}><RefreshCw size={15}/> Refresh</button></div>
            {data.purchaseBills.length ? <><div className="table-wrap"><table><thead><tr><th>Bill</th><th>Date</th><th>Supplier</th><th>Expense account</th><th>Taxable</th><th>GST</th><th>Total</th><th>Paid</th><th>Balance</th><th>Status</th><th></th></tr></thead><tbody>{paginated(data.purchaseBills, "purchase-bills").map(b => <tr key={b.id}><td className="code-cell">{b.bill_number}</td><td>{String(b.bill_date).slice(0,10)}</td><td><strong>{b.supplier_name}</strong></td><td>{b.expense_code} — {b.expense_account_name}</td><td>₹{Number(b.taxable_amount).toFixed(2)}</td><td>{b.gst_rate}%</td><td>₹{Number(b.total_amount).toFixed(2)}</td><td>₹{Number(b.amount_paid).toFixed(2)}</td><td>₹{(Number(b.total_amount)-Number(b.amount_paid)).toFixed(2)}</td><td><span className={`tag ${b.status === "paid" ? "open" : "closed"}`}>{b.status.replace("_"," ")}</span></td><td>{b.status !== "paid" && b.status !== "cancelled" && <button className="button secondary" onClick={() => { setPaymentBill(b); setSupplierPaymentForm({paymentDate:today,accountId:data.accounts.find(a => a.is_active && a.account_type === "asset" && ["1000","1010"].includes(a.code))?.id || "",amount:(Number(b.total_amount)-Number(b.amount_paid)).toFixed(2),reference:""}); setSupplierPaymentError(""); }}>Pay supplier</button>}</td></tr>)}</tbody></table></div>{pagination("purchase-bills", data.purchaseBills.length)}</> : <div className="empty-state">No purchase bills yet. Add a supplier in Customers & suppliers, then record the supplier's bill.</div>}
            <div className="info-footer"><Truck size={17}/> Supplier payments debit Accounts Payable and credit the selected cash/bank account. All postings are recorded in the journal.</div>
          </section>}

          {page === "ai-insights" && <section className="panel">
            <div className="panel-header"><div><h2>AI-assisted financial insights</h2><p>Review accounting trends and potential follow-up items based on your current records.</p></div><button className="button primary" onClick={loadAiInsights} disabled={aiLoading}><Activity size={16}/>{aiLoading ? "Analysing…" : "Generate insights"}</button></div>
            <div className="notice"><div className="notice-icon"><ShieldCheck size={19}/></div><div><strong>Decision support, not a substitute for accounting review</strong><p>Insights are based on recorded transactions and may require verification before tax, payment or compliance decisions.</p></div></div>
            {aiError && <div className="alert error">{aiError}</div>}
            {aiInsights ? <><div className="metric-grid"><Metric icon={IndianRupee} label="Sales invoiced" value={`₹${Number(aiInsights.metrics?.sales || 0).toLocaleString("en-IN",{maximumFractionDigits:2})}`} sub="Recorded sales invoices" tone="blue"/><Metric icon={Wallet} label="Purchases recorded" value={`₹${Number(aiInsights.metrics?.purchases || 0).toLocaleString("en-IN",{maximumFractionDigits:2})}`} sub="Recorded purchase bills" tone="violet"/><Metric icon={ArrowUpRight} label="Customer dues" value={`₹${Number(aiInsights.metrics?.receivables || 0).toLocaleString("en-IN",{maximumFractionDigits:2})}`} sub="Outstanding invoice balances" tone="amber"/><Metric icon={ArrowDownLeft} label="Supplier dues" value={`₹${Number(aiInsights.metrics?.payables || 0).toLocaleString("en-IN",{maximumFractionDigits:2})}`} sub="Outstanding bill balances" tone="green"/></div><div className="insight-list">{(aiInsights.insights||[]).map((item,i)=><article className="insight-card" key={i}><div className="insight-number">{String(i+1).padStart(2,"0")}</div><div><h3>{typeof item === "string" ? item : item.title || item.insight}</h3><p>{typeof item === "object" ? item.detail || item.recommendation || "Review the underlying transaction records before taking action." : "Review the underlying transaction records before taking action."}</p></div></article>)}</div><p className="small-note">Engine: {aiInsights.engine || "Accounting rules"}. {aiInsights.generatedAt ? `Generated ${new Date(aiInsights.generatedAt).toLocaleString()}.` : ""}</p></> : <div className="empty-state">Select “Generate insights” to analyse the available sales, purchase and balance data.</div>}
          </section>}

          {page === "ai-assistant" && <section className="panel assistant-panel">
            <div className="panel-header"><div><h2><Bot size={20}/> AI accounting assistant</h2><p>Ask questions about the figures recorded in your current company workspace.</p></div><button className="button secondary" onClick={()=>setAssistantMessages([])} disabled={!assistantMessages.length || assistantLoading}>Clear chat</button></div>
            <div className="notice"><div className="notice-icon"><ShieldCheck size={19}/></div><div><strong>Read-only and company-scoped</strong><p>The assistant reads a limited accounting summary for your authenticated company. It cannot create or modify transactions. Verify figures in the source reports before acting.</p></div></div>
            <div className="assistant-suggestions"><span>Try asking:</span>{["How much do customers still owe us?","Which customers have the highest outstanding balances?","How much is due to suppliers?","Summarise recent journal entries.","What should I review before making supplier payments?"] .map(q=><button key={q} className="suggestion-chip" onClick={()=>setAssistantQuestion(q)} disabled={assistantLoading}>{q}</button>)}</div>
            <div className="assistant-conversation" aria-live="polite">
              {!assistantMessages.length && <div className="assistant-empty"><div className="assistant-empty-icon"><MessageCircle size={25}/></div><h3>Ask your accounting question</h3><p>Responses use the latest accounting snapshot available when you submit each question.</p></div>}
              {assistantMessages.map((m,i)=><article key={i} className={`chat-message ${m.role === "user" ? "chat-user" : "chat-assistant"}`}><div className="chat-avatar">{m.role === "user" ? (user.fullName||"U").slice(0,1).toUpperCase() : <Bot size={17}/>}</div><div className="chat-bubble"><div className="chat-role">{m.role === "user" ? "You" : "SoftSpark AI"}</div><p>{m.content}</p>{m.engine && <div className="chat-meta">{m.engine}{m.generatedAt ? ` · ${new Date(m.generatedAt).toLocaleTimeString()}` : ""}</div>}</div></article>)}
              {assistantLoading && <article className="chat-message chat-assistant"><div className="chat-avatar"><Bot size={17}/></div><div className="chat-bubble"><div className="chat-role">SoftSpark AI</div><p className="typing-indicator">Analysing your question…</p></div></article>}
            </div>
            {assistantError && <div className="alert error">{assistantError}</div>}
            <form className="assistant-composer" onSubmit={askAssistant}><textarea value={assistantQuestion} onChange={e=>setAssistantQuestion(e.target.value)} maxLength={1200} rows={3} placeholder="Ask about receivables, payables, recorded sales, purchases or journals…" required disabled={assistantLoading}/><div className="composer-footer"><span>{assistantQuestion.length}/1200 characters · Groq {import.meta.env.VITE_GROQ_MODEL || "openai/gpt-oss-120b"}</span><button className="button primary" disabled={assistantLoading || assistantQuestion.trim().length<3}><Send size={16}/>{assistantLoading ? "Thinking…" : "Ask assistant"}</button></div></form>
          </section>}

          {page === "financial-reports" && <section className="panel">
            <div className="panel-header"><div><h2>Financial reports</h2><p>Profit & loss, balance sheet balances and GST transaction summary for the selected financial year.</p></div><button className="button secondary" onClick={()=>loadFinancialReports()} disabled={reportsLoading}><RefreshCw size={15}/>{reportsLoading?"Refreshing…":"Refresh reports"}</button></div>
            {reportsError && <div className="alert error">{reportsError}</div>}
            {reportsLoading && <div className="empty-state">Calculating financial reports…</div>}
            {!reportsLoading && financialReports?.financialYear && <>
              <div className="notice"><div className="notice-icon"><FileBarChart size={19}/></div><div><strong>{financialReports.financialYear.name}</strong><p>{String(financialReports.financialYear.start_date).slice(0,10)} to {String(financialReports.financialYear.end_date).slice(0,10)} · Based on posted journal entries and recorded GST documents</p></div></div>
              <div className="metric-grid"><div className="metric-card"><span>Total revenue</span><strong>₹{Number(financialReports.incomeTotals?.revenue||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</strong></div><div className="metric-card"><span>Total expenses</span><strong>₹{Number(financialReports.incomeTotals?.expenses||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</strong></div><div className="metric-card"><span>Net profit / (loss)</span><strong>₹{Number(financialReports.incomeTotals?.netProfit||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</strong></div><div className="metric-card"><span>Indicative net GST payable</span><strong>₹{Number(financialReports.gstSummary?.netGstPayable||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</strong></div></div>
              <h3 className="report-subheading">Profit & loss</h3><div className="table-wrap"><table><thead><tr><th>Code</th><th>Account</th><th>Type</th><th>Amount</th></tr></thead><tbody>{(financialReports.incomeStatement||[]).map((r,i)=><tr key={`${r.code}-${i}`}><td className="code-cell">{r.code}</td><td><strong>{r.name}</strong></td><td>{capitalize(r.type)}</td><td>₹{Number(r.amount).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr>)}{!financialReports.incomeStatement?.length&&<tr><td colSpan="4">No revenue or expense postings for this financial year.</td></tr>}</tbody><tfoot><tr><th colSpan="3">Net profit / (loss)</th><th>₹{Number(financialReports.incomeTotals?.netProfit||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</th></tr></tfoot></table></div>
              <h3 className="report-subheading">Balance sheet balances</h3><div className="table-wrap"><table><thead><tr><th>Code</th><th>Account</th><th>Category</th><th>Balance</th><th>Side</th></tr></thead><tbody>{(financialReports.balanceSheet||[]).map((r,i)=><tr key={`${r.code}-${i}`}><td className="code-cell">{r.code}</td><td><strong>{r.name}</strong></td><td>{capitalize(r.type)}</td><td>₹{Number(r.balance).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>{capitalize(r.side)}</td></tr>)}{!financialReports.balanceSheet?.length&&<tr><td colSpan="5">No balance sheet balances are available.</td></tr>}</tbody></table></div>
              <h3 className="report-subheading">GST transaction summary (not a filed return)</h3><div className="table-wrap"><table><thead><tr><th>Metric</th><th>Sales / Output GST</th><th>Purchases / Input GST</th></tr></thead><tbody><tr><td>Taxable value</td><td>₹{Number(financialReports.gstSummary.salesTaxable).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>₹{Number(financialReports.gstSummary.purchaseTaxable).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr><tr><td>CGST</td><td>₹{Number(financialReports.gstSummary.outputCgst).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>₹{Number(financialReports.gstSummary.inputCgst).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr><tr><td>SGST</td><td>₹{Number(financialReports.gstSummary.outputSgst).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>₹{Number(financialReports.gstSummary.inputSgst).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr><tr><td>IGST</td><td>₹{Number(financialReports.gstSummary.outputIgst).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td><td>₹{Number(financialReports.gstSummary.inputIgst).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</td></tr><tr><th>Total GST</th><th>₹{(Number(financialReports.gstSummary.outputCgst)+Number(financialReports.gstSummary.outputSgst)+Number(financialReports.gstSummary.outputIgst)).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</th><th>₹{(Number(financialReports.gstSummary.inputCgst)+Number(financialReports.gstSummary.inputSgst)+Number(financialReports.gstSummary.inputIgst)).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}</th></tr></tbody></table></div>
              <div className="info-footer"><ShieldCheck size={17}/> Management reports only. Verify ledger mappings, opening balances, GST eligibility, reversals and statutory rules with your accountant before relying on or filing these figures.</div>
            </>}
            {!reportsLoading && !financialReports?.financialYear && !reportsError && <div className="empty-state">Create a financial year before generating reports.</div>}
          </section>}

          {page === "trial-balance" && <section className="panel">
            <div className="panel-header"><div><h2>Trial balance</h2><p>Account balances are calculated from posted journal lines in the selected financial year.</p></div><button className="button secondary" onClick={() => load()}><RefreshCw size={15}/> Refresh</button></div>
            {data.trialBalance?.financialYear && <div className="notice"><div className="notice-icon"><Scale size={19}/></div><div><strong>{data.trialBalance.financialYear.name}</strong><p>{String(data.trialBalance.financialYear.start_date).slice(0,10)} to {String(data.trialBalance.financialYear.end_date).slice(0,10)} · Status: {data.trialBalance.financialYear.status}</p></div></div>}
            {data.trialBalance?.rows?.length ? <><div className="table-wrap"><table><thead><tr><th>Code</th><th>Account</th><th>Type</th><th>Total debit</th><th>Total credit</th><th>Closing balance</th><th>Side</th></tr></thead><tbody>{paginated(data.trialBalance.rows, "trial-balance").map(r => <tr key={r.id}><td className="code-cell">{r.code}</td><td><strong>{r.name}</strong></td><td>{capitalize(r.account_type)}</td><td>₹{Number(r.total_debit).toFixed(2)}</td><td>₹{Number(r.total_credit).toFixed(2)}</td><td>₹{Number(r.balance).toFixed(2)}</td><td>{capitalize(r.balance_side)}</td></tr>)}</tbody><tfoot><tr><th colSpan="3">Totals</th><th>₹{Number(data.trialBalance.totals.debit).toFixed(2)}</th><th>₹{Number(data.trialBalance.totals.credit).toFixed(2)}</th><th colSpan="2">{Math.abs(Number(data.trialBalance.totals.debit) - Number(data.trialBalance.totals.credit)) < 0.005 ? "Balanced" : "Out of balance"}</th></tr></tfoot></table></div>{pagination("trial-balance", data.trialBalance.rows.length)}</> : <div className="empty-state">No financial year found. Configure a financial year before reviewing the trial balance.</div>}
            <div className="info-footer"><ShieldCheck size={17}/> This is an operational trial balance for development and testing, not a certified statutory report.</div>
          </section>}

          {["inventory","bank-reconciliation","approvals","tax-registers","team"].includes(page) && <BusinessWorkspace key={page} page={page} session={session} data={data} setToast={setToast} load={load} />}
          {page === "enterprise" && <EnterpriseWorkspace session={session} data={data} setToast={setToast} />}
          <footer className="app-footer"><span>© SoftSpark Tech Studio · AutoBooks Enterprise</span><span>Release 3 · Enterprise entities, cost centres, budgets and consolidation workpapers</span></footer>
        </div>
      </main>

      {receiptInvoice && <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) setReceiptInvoice(null); }}><div className="modal"><div className="modal-head"><div><h2>Record customer receipt</h2><p>{receiptInvoice.invoice_number} · {receiptInvoice.customer_name}</p></div><button className="icon-button" onClick={() => setReceiptInvoice(null)}><X size={18}/></button></div><form className="form-stack" onSubmit={postReceipt}><label>Receipt date<input type="date" value={receiptForm.receiptDate} onChange={e => setReceiptForm({...receiptForm,receiptDate:e.target.value})} required/></label><label>Deposit account<select value={receiptForm.accountId} onChange={e => setReceiptForm({...receiptForm,accountId:e.target.value})} required><option value="">Select cash/bank account</option>{data.accounts.filter(a => a.is_active && a.account_type === "asset").map(a => <option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></label><label>Amount received (₹)<input type="number" min="0.01" step="0.01" max={(Number(receiptInvoice.total_amount)-Number(receiptInvoice.amount_received)).toFixed(2)} value={receiptForm.amount} onChange={e => setReceiptForm({...receiptForm,amount:e.target.value})} required/></label><label>Reference (optional)<input value={receiptForm.reference} onChange={e => setReceiptForm({...receiptForm,reference:e.target.value})} maxLength={120} placeholder="Bank reference / UTR"/></label>{receiptError && <div className="alert error">{receiptError}</div>}<div className="modal-actions"><button type="button" className="button secondary" onClick={() => setReceiptInvoice(null)}>Cancel</button><button className="button primary">Post receipt</button></div><div className="small-note">Receipt posting is final and updates the invoice balance and ledger atomically.</div></form></div></div>}

      {paymentBill && <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) setPaymentBill(null); }}><div className="modal"><div className="modal-head"><div><h2>Pay supplier</h2><p>{paymentBill.bill_number} · {paymentBill.supplier_name}</p></div><button className="icon-button" onClick={() => setPaymentBill(null)}><X size={18}/></button></div><form className="form-stack" onSubmit={postSupplierPayment}><label>Payment date<input type="date" value={supplierPaymentForm.paymentDate} onChange={e => setSupplierPaymentForm({...supplierPaymentForm,paymentDate:e.target.value})} required/></label><label>Pay from<select value={supplierPaymentForm.accountId} onChange={e => setSupplierPaymentForm({...supplierPaymentForm,accountId:e.target.value})} required><option value="">Select cash/bank account</option>{data.accounts.filter(a => a.is_active && a.account_type === "asset").map(a => <option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></label><label>Amount paid (₹)<input type="number" min="0.01" step="0.01" max={(Number(paymentBill.total_amount)-Number(paymentBill.amount_paid)).toFixed(2)} value={supplierPaymentForm.amount} onChange={e => setSupplierPaymentForm({...supplierPaymentForm,amount:e.target.value})} required/></label><label>Reference (optional)<input value={supplierPaymentForm.reference} onChange={e => setSupplierPaymentForm({...supplierPaymentForm,reference:e.target.value})} maxLength={120} placeholder="Bank reference / UTR"/></label>{supplierPaymentError && <div className="alert error">{supplierPaymentError}</div>}<div className="modal-actions"><button type="button" className="button secondary" onClick={() => setPaymentBill(null)}>Cancel</button><button className="button primary">Post supplier payment</button></div><div className="small-note">Payments are final in this build and update the payable balance and ledger atomically.</div></form></div></div>}

      {showAdd && <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) setShowAdd(false); }}>
        <div className="modal">
          <div className="modal-head"><div><h2>Add account</h2><p>Create a chart-of-account record.</p></div><button className="icon-button" onClick={() => setShowAdd(false)}><X size={18}/></button></div>
          <form className="form-stack" onSubmit={createAccount}>
            <label>Account code<input value={accountForm.code} onChange={e => setAccountForm({...accountForm, code:e.target.value})} placeholder="e.g. 6400" required maxLength={30}/></label>
            <label>Account name<input value={accountForm.name} onChange={e => setAccountForm({...accountForm, name:e.target.value})} placeholder="e.g. Internet Expense" required maxLength={160}/></label>
            <label>Account type<select value={accountForm.accountType} onChange={e => setAccountForm({...accountForm, accountType:e.target.value})}><option value="asset">Asset</option><option value="liability">Liability</option><option value="equity">Equity</option><option value="revenue">Revenue</option><option value="expense">Expense</option></select></label>
            <div className="small-note">Normal balance is assigned from the selected account type. You can only add accounts as an owner, administrator or accountant.</div>
            {formError && <div className="alert error">{formError}</div>}
            <div className="modal-actions"><button type="button" className="button secondary" onClick={() => setShowAdd(false)}>Cancel</button><button className="button primary"><Plus size={16}/> Create account</button></div>
          </form>
        </div>
      </div>}
    </div>
  );
}

function Metric({icon:Icon,label,value,sub,tone}) {
  return <div className="metric-card"><div className={`metric-icon ${tone}`}><Icon size={19}/></div><div className="metric-label">{label}</div><div className="metric-value">{value}</div><div className="metric-sub">{sub}</div></div>;
}
function SetupCard({icon:Icon,number,title,description,status,action}) {
  return <button className="setup-card" onClick={action}><div className="setup-card-top"><div className="setup-icon"><Icon size={19}/></div><span>{number}</span></div><h3>{title}</h3><p>{description}</p><div className="setup-card-bottom"><span>{status}</span><span className="arrow">↗</span></div></button>;
}
function Coming({icon:Icon,title,desc}) {
  return <div className="coming-card"><div className="coming-icon"><Icon size={18}/></div><div><strong>{title}</strong><p>{desc}</p><span>Planned</span></div></div>;
}
function Detail({label,value}) {
  return <div className="detail-item"><span>{label}</span><strong>{value || "—"}</strong></div>;
}
function capitalize(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : "—"; }

export default App;
