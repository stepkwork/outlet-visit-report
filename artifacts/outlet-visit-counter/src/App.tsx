import { useEffect, useMemo, useState } from 'react';
import { ArrowDownToLine, Check, ChevronDown, ClipboardList, FileText, MapPin, Minus, Plus, Printer, Search, Trash2 } from 'lucide-react';
import templateUrl from './assets/Outlet_Visit_Feedback_Template.xlsx?url';
import { buildVisitWorkbook } from './lib/visit-workbook';

const BRANDS = [
  'TOSHIBA', 'LG', 'PANASONIC', 'SAMSUNG', 'HAIER', 'MIDEA', 'HISENSE',
  'SHARP', 'HITACHI', 'TCL', 'KHIND', 'FABER', 'ELBA', 'PENSONIC',
  'SKYWORTH', 'ELECTROLUX', 'CHIQ', 'BEKO', 'MORGAN', 'ISONIC', 'ECOZY', 'HITEC',
];
const CATEGORIES = ['SD', '2D', 'MD', 'SBS', 'CF', 'FFC', 'FLW', 'FLD', 'FLWD', 'TLW', 'TT'];
const CATEGORY_TITLES: Record<string, string> = {
  SD: 'Single Door', '2D': '2 Door', MD: 'Multi Door', SBS: 'Side by Side',
  CF: 'Chest Freezer', FFC: 'FFC', FLW: 'Front Load Washer',
  FLD: 'Front Load Dryer', FLWD: 'Washer Dryer', TLW: 'Top Load Washer', TT: 'Twin Tub',
};
const SAVED_KEY = 'outlet-visit-counter:saved-v1';
const DRAFT_KEY = 'outlet-visit-counter:draft-v1';

type VisitFields = {
  managerName: string;
  shopName: string;
  visitDate: string;
  outletCount: string;
  outletVisitAddress: string;
};
type Visit = VisitFields & { id: string; savedAt: string; counts: Record<string, Record<string, number>> };
const today = () => {
  const date = new Date();
  const localDate = new Date(date.getTime() - date.getTimezoneOffset() * 60_000);
  return localDate.toISOString().slice(0, 10);
};
const blankCounts = () => Object.fromEntries(BRANDS.map((brand) => [brand, Object.fromEntries(CATEGORIES.map((category) => [category, 0]))]));
const blankFields = (): VisitFields => ({ managerName: '', shopName: '', visitDate: today(), outletCount: '', outletVisitAddress: '' });
const blankDraft = () => ({ fields: blankFields(), counts: blankCounts() });
function readStored<T>(key: string, fallback: T): T {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) as T : fallback;
  } catch {
    return fallback;
  }
}
function safeFilePart(value: string) {
  return (value || 'outlet-visit').trim().replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '').toLowerCase();
}

function App() {
  const initialDraft = useMemo(() => readStored(DRAFT_KEY, blankDraft()), []);
  const [fields, setFields] = useState<VisitFields>(() => ({ ...blankFields(), ...initialDraft.fields }));
  const [counts, setCounts] = useState<Record<string, Record<string, number>>>(() => {
    const defaults = blankCounts();
    Object.entries(initialDraft.counts || {}).forEach(([brand, row]) => {
      if (defaults[brand]) Object.entries(row).forEach(([category, value]) => {
        if (CATEGORIES.includes(category)) defaults[brand][category] = Math.max(0, Number(value) || 0);
      });
    });
    return defaults;
  });
  const [savedVisits, setSavedVisits] = useState<Visit[]>(() => readStored(SAVED_KEY, []));
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [query, setQuery] = useState('');
  const [toast, setToast] = useState('');
  const [mobileSavedOpen, setMobileSavedOpen] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<string | null>(null);
  const [lastSavedAt, setLastSavedAt] = useState('');

  useEffect(() => {
    try {
      localStorage.setItem(DRAFT_KEY, JSON.stringify({ fields, counts }));
      setLastSavedAt(new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }));
    } catch {
      setToast('Could not save draft on this device');
    }
  }, [fields, counts]);
  useEffect(() => {
    try { localStorage.setItem(SAVED_KEY, JSON.stringify(savedVisits)); }
    catch { showToast('Storage is full. Export or remove a saved visit.'); }
  }, [savedVisits]);
  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(''), 2200);
    return () => window.clearTimeout(timer);
  }, [toast]);

  function showToast(message: string) {
    setToast(message);
  }
  function updateField<K extends keyof VisitFields>(key: K, value: VisitFields[K]) {
    setFields((current) => ({ ...current, [key]: value }));
  }
  function changeCount(brand: string, category: string, amount: number) {
    setCounts((current) => ({
      ...current,
      [brand]: { ...current[brand], [category]: Math.max(0, (current[brand]?.[category] || 0) + amount) },
    }));
  }
  function saveVisit() {
    const record: Visit = {
      ...fields,
      id: selectedId || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      savedAt: new Date().toISOString(),
      counts: JSON.parse(JSON.stringify(counts)),
    };
    setSavedVisits((current) => selectedId
      ? current.map((visit) => visit.id === selectedId ? record : visit)
      : [record, ...current]);
    setSelectedId(record.id);
    showToast(selectedId ? 'Visit updated and saved' : 'Visit saved on this device');
  }
  function loadVisit(visit: Visit) {
    setFields({ managerName: visit.managerName, shopName: visit.shopName, visitDate: visit.visitDate, outletCount: visit.outletCount, outletVisitAddress: visit.outletVisitAddress });
    const next = blankCounts();
    Object.entries(visit.counts || {}).forEach(([brand, row]) => {
      if (next[brand]) Object.entries(row).forEach(([category, value]) => {
        if (CATEGORIES.includes(category)) next[brand][category] = Math.max(0, Number(value) || 0);
      });
    });
    setCounts(next);
    setSelectedId(visit.id);
    showToast(`Opened ${visit.shopName || 'saved visit'}`);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  function startNewVisit() {
    setFields(blankFields());
    setCounts(blankCounts());
    setSelectedId(null);
    showToast('New visit started');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  async function exportVisit(visit: Visit) {
    try {
      const templateResponse = await fetch(templateUrl);
      if (!templateResponse.ok) throw new Error('The Excel template could not be loaded.');
      const templateBytes = await templateResponse.arrayBuffer();
      const workbookBytes = await buildVisitWorkbook(templateBytes, visit, BRANDS, CATEGORIES);
      const workbookBuffer = new ArrayBuffer(workbookBytes.byteLength);
      new Uint8Array(workbookBuffer).set(workbookBytes);
      const blob = new Blob([workbookBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${safeFilePart(visit.shopName)}-${visit.visitDate || 'visit'}.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      showToast('Excel workbook downloaded');
    } catch {
      showToast('Could not create the Excel workbook. Please try again.');
    }
  }
  function printCurrent() {
    window.print();
  }
  const visibleBrands = BRANDS.filter((brand) => brand.toLowerCase().includes(query.trim().toLowerCase()));
  const totalUnits = useMemo(() => Object.values(counts).reduce((sum, row) => sum + Object.values(row).reduce((subtotal, count) => subtotal + count, 0), 0), [counts]);
  const populatedBrands = useMemo(() => Object.values(counts).filter((row) => Object.values(row).some((count) => count > 0)).length, [counts]);
  const isEditingSaved = Boolean(selectedId);

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="brand-lockup">
          <div className="brand-mark"><ClipboardList size={18} strokeWidth={1.8} /></div>
          <div><div className="brand-name">Outlet Visit Counter</div><div className="brand-kicker">Field inventory / MY</div></div>
        </div>
        <div className="top-meta">
          <span className="sync-dot" />
          <span>Saved locally</span>
          <button className="secondary-button new-visit" type="button" onClick={startNewVisit}><Plus size={14} /> New visit</button>
        </div>
      </header>

      <main className="page-wrap">
        <div className="page-heading">
          <div>
            <div className="eyebrow">Field record / {fields.visitDate || today()}</div>
            <h1>{fields.shopName || 'Outlet visit'}</h1>
            <p>Record the display, one count at a time.</p>
          </div>
          <div className="draft-status"><span className="draft-indicator" />{isEditingSaved ? 'Editing saved visit' : `Draft saved${lastSavedAt ? ` at ${lastSavedAt}` : ''}`}</div>
        </div>

        <div className="workspace">
          <div className="main-column">
            <section className="panel visit-panel" aria-labelledby="visit-details-heading">
              <div className="section-head">
                <div className="section-title"><span className="section-number">01</span><h2 id="visit-details-heading">Visit details</h2></div>
                <span className="section-note">Saved automatically as a draft</span>
              </div>
              <div className="fields-grid">
                <div className="field">
                  <label htmlFor="managerName">Manager name</label>
                  <input id="managerName" value={fields.managerName} onChange={(event) => updateField('managerName', event.target.value)} placeholder="Name of manager" autoComplete="name" />
                </div>
                <div className="field">
                  <label htmlFor="shopName">Shop name</label>
                  <input id="shopName" value={fields.shopName} onChange={(event) => updateField('shopName', event.target.value)} placeholder="Outlet / retailer" />
                </div>
                <div className="field">
                  <label htmlFor="visitDate">Visit date</label>
                  <input id="visitDate" type="date" value={fields.visitDate} onChange={(event) => updateField('visitDate', event.target.value)} />
                </div>
                <div className="field">
                  <label htmlFor="outletCount">Outlet count</label>
                  <input id="outletCount" type="number" min="0" inputMode="numeric" value={fields.outletCount} onChange={(event) => updateField('outletCount', event.target.value)} placeholder="0" />
                </div>
                <div className="field address">
                  <label htmlFor="outletVisitAddress">Outlet visit address</label>
                  <textarea id="outletVisitAddress" rows={1} value={fields.outletVisitAddress} onChange={(event) => updateField('outletVisitAddress', event.target.value)} placeholder="Street, town, or outlet location" />
                </div>
              </div>
            </section>

            <section className="panel inventory-panel" aria-labelledby="inventory-heading">
              <div className="inventory-heading">
                <div className="inventory-title">
                  <div className="section-title"><span className="section-number">02</span><h2 id="inventory-heading">Display count</h2></div>
                  <p>Tap + or − as you check each brand. Counts stay on this device.</p>
                </div>
                <div className="totals">
                  <div className="total-block"><div className="total-num">{totalUnits}</div><div className="total-caption">Units counted</div></div>
                  <div className="total-block"><div className="total-num">{populatedBrands}<span style={{ fontSize: 12, color: '#829087' }}> / 22</span></div><div className="total-caption">Brands seen</div></div>
                </div>
                <div className="search-wrap">
                  <Search size={14} />
                  <input className="search-input" type="search" aria-label="Find a brand" placeholder="Find a brand" value={query} onChange={(event) => setQuery(event.target.value)} />
                </div>
              </div>
              <div className="print-only print-meta">
                <strong>{fields.shopName || 'Outlet visit'}</strong> · {fields.visitDate} · Manager: {fields.managerName || '—'} · Address: {fields.outletVisitAddress || '—'} · Outlet count: {fields.outletCount || '—'}
              </div>
              <div className="matrix-scroll">
                <table className="matrix">
                  <thead><tr>
                    <th scope="col">BRAND</th>
                    {CATEGORIES.map((category) => <th scope="col" key={category} title={CATEGORY_TITLES[category]}>{category}</th>)}
                  </tr></thead>
                  <tbody>
                    {visibleBrands.map((brand) => (
                      <tr key={brand}>
                        <td><div className="brand-cell"><span className="brand-index">{String(BRANDS.indexOf(brand) + 1).padStart(2, '0')}</span><span className="brand-label">{brand}</span></div></td>
                        {CATEGORIES.map((category) => {
                          const value = counts[brand]?.[category] ?? 0;
                          return <td key={category}>
                            <div className="count-control">
                              <button type="button" aria-label={`Decrease ${brand} ${category}${value === 0 ? ' (zero)' : ''}`} title={`Decrease ${brand} ${CATEGORY_TITLES[category]}`} disabled={value === 0} onClick={() => changeCount(brand, category, -1)}><Minus size={13} /></button>
                              <span className="count-value" aria-live="polite">{value}</span>
                              <button type="button" aria-label={`Increase ${brand} ${category}`} title={`Increase ${brand} ${CATEGORY_TITLES[category]}`} onClick={() => changeCount(brand, category, 1)}><Plus size={13} /></button>
                            </div>
                          </td>;
                        })}
                      </tr>
                    ))}
                    {visibleBrands.length === 0 && <tr><td colSpan={CATEGORIES.length + 1}><div className="matrix-empty">No brand matches “{query}”.</div></td></tr>}
                  </tbody>
                </table>
              </div>
            </section>
          </div>

          <aside className="sidebar">
            <section className="save-card">
              <div className="save-card-top"><h3>{isEditingSaved ? 'Update this visit' : 'Ready to save?'}</h3><span className="save-status"><Check size={12} style={{ verticalAlign: '-2px' }} /> On device</span></div>
              <button type="button" className="save-button" onClick={saveVisit}><FileText size={16} />{isEditingSaved ? 'Update saved visit' : 'Save visit record'}</button>
              <div className="draft-actions">
                <button type="button" className="secondary-button" style={{ flex: 1 }} onClick={printCurrent}><Printer size={13} /> Print sheet</button>
                <button type="button" className="secondary-button" style={{ flex: 1 }} onClick={() => { if (selectedId) { const visit = savedVisits.find((item) => item.id === selectedId); if (visit) void exportVisit({ ...visit, ...fields, counts }); } else { const temporary: Visit = { ...fields, id: 'draft', savedAt: new Date().toISOString(), counts }; void exportVisit(temporary); } }}><ArrowDownToLine size={13} /> Export Excel</button>
              </div>
              <p className="save-hint">Your working draft is saved automatically.</p>
            </section>
            <button className="mobile-saved-toggle" type="button" onClick={() => setMobileSavedOpen((open) => !open)} aria-expanded={mobileSavedOpen}>
              <span>Saved visits · {savedVisits.length}</span><ChevronDown size={16} style={{ transform: mobileSavedOpen ? 'rotate(180deg)' : undefined, transition: 'transform .16s' }} />
            </button>
            <section className={`panel side-panel${mobileSavedOpen ? ' mobile-open' : ''}`}>
              <div className="side-heading"><h3>Saved visits</h3><span className="saved-badge">{savedVisits.length}</span></div>
              <div className="visit-list">
                {savedVisits.length === 0
                  ? <div className="empty-saved"><div className="empty-icon"><MapPin size={16} /></div>Saved outlet visits will appear here.<br />Your records stay on this device.</div>
                  : savedVisits.map((visit) => (
                    <div className="visit-item" key={visit.id}>
                      <div className="visit-main" role="button" tabIndex={0} onClick={() => loadVisit(visit)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') loadVisit(visit); }}>
                        <div className="visit-shop">{visit.shopName || 'Untitled outlet'}</div>
                        <div className="visit-meta">{visit.visitDate || 'No date'} · {visit.managerName || 'No manager'}</div>
                      </div>
                      <div className="visit-actions">
                        <button type="button" className="icon-button" title="Export this visit as an Excel workbook" aria-label={`Export ${visit.shopName || 'visit'} as Excel workbook`} onClick={() => void exportVisit(visit)}><ArrowDownToLine size={14} /></button>
                        <button type="button" className="icon-button danger" title="Delete saved visit" aria-label={`Delete ${visit.shopName || 'visit'}`} onClick={() => setDeleteTarget(visit.id)}><Trash2 size={14} /></button>
                      </div>
                    </div>
                  ))}
              </div>
            </section>
          </aside>
        </div>
      </main>

      <div className={`toast-message${toast ? ' visible' : ''}`} role="status" aria-live="polite">{toast}</div>
      {deleteTarget && <div className="confirm-overlay" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) setDeleteTarget(null); }}>
        <div className="confirm-box" role="dialog" aria-modal="true" aria-labelledby="delete-heading">
          <h3 id="delete-heading">Delete this visit?</h3>
          <p>This saved record will be removed from this device. This cannot be undone.</p>
          <div className="confirm-actions">
            <button type="button" onClick={() => setDeleteTarget(null)}>Keep visit</button>
            <button type="button" className="delete-confirm" onClick={() => { setSavedVisits((current) => current.filter((visit) => visit.id !== deleteTarget)); if (selectedId === deleteTarget) startNewVisit(); setDeleteTarget(null); showToast('Saved visit deleted'); }}>Delete record</button>
          </div>
        </div>
      </div>}
    </div>
  );
}

export default App;
