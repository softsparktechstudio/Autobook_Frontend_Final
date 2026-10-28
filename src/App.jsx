import React, { useEffect, useMemo, useState } from "react";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as XLSX from "xlsx";
import {
  Activity, ArrowDownLeft, ArrowUpRight, BookOpen, Building2, CalendarDays, Truck, FileBarChart,
  CheckCircle2, ChevronDown, CircleHelp, FileText, LayoutDashboard, LogOut,
  Menu, Plus, RefreshCw, Search, ShieldCheck, Wallet, X, ReceiptText, Scale, Users, FilePlus2, IndianRupee, MessageCircle, Send, Bot
} from "lucide-react";

const API = import.meta.env.VITE_API_URL || "http://localhost:5000/api";

async function request(path, token, options = {}) {
  const response = await fetch(`${API}${path}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
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
      const data = await request("/api/auth/login", null, {
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
          <div className="brand-name">SoftSpark <span>AutoBooks Core</span></div>
          <div className="brand-subtitle">Accounting, reporting and insights</div>
        </div>
      </div>
      <div className="login-card">
        <div className="eyebrow">AUTOBOOKS CORE · RELEASE 1</div>
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
  { label: "AI assistant", icon: MessageCircle, id: "ai-assistant" }
];

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
  function exportPDF() { const doc=new jsPDF({orientation:"landscape"}); doc.setFontSize(16); doc.text(`SoftSpark AutoBooks Core - ${currentPage}`,14,16); doc.setFontSize(9); doc.text(`Company: ${company?.legal_name || "SoftSpark AutoBooks Core"}   Generated: ${today}`,14,23); const rows=exportRows(); const columns=Object.keys(rows[0]||{Information:""}); autoTable(doc,{startY:29,head:[columns],body:rows.map(r=>columns.map(c=>String(r[c]??""))),styles:{fontSize:7,cellPadding:2},headStyles:{fillColor:[31,67,116]}}); doc.save(`AutoBooks_Core_${page}_${today}.pdf`); }
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
        request("api/organization", token),
        request("api/company", token),
        request("api/financial-years", token),
        request("api/accounts", token),
        request("api/journals", token),
        request("api/reports/trial-balance", token),
        request("api/parties", token),
        request("api/sales-invoices", token),
        request("api/purchase-bills", token)
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
          <div className="brand-name">SoftSpark <span>AutoBooks Core</span></div>
          <button className="icon-button mobile-close" onClick={() => setMobileNav(false)}><X size={18}/></button>
        </div>
        <div className="workspace">
          <div className="workspace-icon"><Building2 size={18}/></div>
          <div className="workspace-text"><strong>{data.organization?.name || "Loading workspace"}</strong><span>India · INR</span></div>
          <ChevronDown size={15} className="muted-icon"/>
        </div>
        <div className="nav-caption">WORKSPACE</div>
        <nav>
          {nav.map(item => {
            const Icon = item.icon;
            return <button key={item.id} className={`nav-item ${page === item.id ? "active" : ""}`} onClick={() => { setPage(item.id); setMobileNav(false); }}>
              <Icon size={18}/><span>{item.label}</span>{item.id === "accounts" && <span className="nav-count">{data.accounts.length}</span>}
            </button>;
          })}
        </nav>
        <div className="sidebar-bottom">
          <button className="nav-item" onClick={() => setToast("Open README.md in the project folder for the complete AutoBooks Core setup and user guide.")}><CircleHelp size={18}/><span>Help & guidance</span></button>
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
            <div className="environment"><span/> Release 1 · Core</div>
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

          <footer className="app-footer"><span>© SoftSpark Tech Studio · AutoBooks Core</span><span>Release 1 · Accounting, financial reports and AI-assisted insights</span></footer>
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
