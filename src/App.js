import { useState, useEffect, useCallback } from "react";
import { format, startOfWeek } from "date-fns";

// ─────────────────────────────────────────────
// CONFIG — fill these in before deploying
// ─────────────────────────────────────────────
const CONFIG = {
  // Your Cloudflare Worker URL
  WORKER_URL: "https://sweet-king-342d.mperkinslewis.workers.dev",

  // Azure AD / MSAL
  CLIENT_ID: "YOUR-AZURE-APP-CLIENT-ID",
  TENANT_ID: "YOUR-TENANT-ID",
  SHAREPOINT_SITE: "https://YOUR-TENANT.sharepoint.com/sites/YOUR-SITE",
  LIST_NAME: "QuoteRequests",

  // Business rules
  MIN_REVENUE_PER_DELIVERY: 150,
  CUSTOM_MARGIN: 0.36,
};

// ─────────────────────────────────────────────
// PRICE LIST — add your ~80 lines here
// Format: { code, description, tierA, tierB, tierC }
// Tiers are per kg unless noted
// ─────────────────────────────────────────────
const PRICE_LIST = [
  { code: "SB140", description: "Sea Bass Fillet 140-160g", tierA: 12.5, tierB: 13.0, tierC: 13.8 },
  { code: "SAL200", description: "Salmon Fillet 180-220g", tierA: 8.2, tierB: 8.8, tierC: 9.4 },
  { code: "COD180", description: "Cod Loin 180-220g", tierA: 11.0, tierB: 11.8, tierC: 12.5 },
  { code: "HAL200", description: "Halibut Steak 200g", tierA: 18.0, tierB: 19.2, tierC: 20.5 },
  { code: "BREAM", description: "Sea Bream Whole 400-600g", tierA: 6.8, tierB: 7.4, tierC: 8.0 },
  { code: "TURB", description: "Turbot Fillet", tierA: 28.0, tierB: 30.0, tierC: 32.0 },
  { code: "DOVER", description: "Dover Sole Fillet", tierA: 32.0, tierB: 34.5, tierC: 37.0 },
  { code: "MACWH", description: "Mackerel Whole", tierA: 2.8, tierB: 3.2, tierC: 3.6 },
  { code: "SCAMPI", description: "Scampi Tails", tierA: 14.0, tierB: 15.2, tierC: 16.5 },
  { code: "PRAWNT", description: "King Prawns 16/20", tierA: 9.5, tierB: 10.2, tierC: 11.0 },
  // ... add remaining lines
];

// ─────────────────────────────────────────────
// MOCK DATA for development (remove when live)
// ─────────────────────────────────────────────
const MOCK_REQUESTS = [
  {
    id: "1",
    CustomerName: "The Anchor Inn",
    FishTypes: "sea bass fillet 140-160, salmon portion 200g x20, cod loin x15",
    EstMonthlySpend: "2400",
    DeliveriesPerWeek: "2",
    CurrentSuppliers: "Bidfood",
    Notes: "Head chef wants sustainable sourcing certs",
    IsRequote: false,
    Status: "Pending",
    SubmittedDate: new Date().toISOString(),
    WeekCommencing: startOfWeek(new Date(), { weekStartsOn: 1 }).toISOString(),
  },
  {
    id: "2",
    CustomerName: "Harbour Lights Restaurant",
    FishTypes: "bass fillet 1 forty to 1 sixty x30, dover sole fillet x8, king prawns large x5kg",
    EstMonthlySpend: "3800",
    DeliveriesPerWeek: "3",
    CurrentSuppliers: "Local fishmonger, M&J Seafood",
    Notes: "Currently paying £13.50 for bass — wants to beat that",
    IsRequote: true,
    Status: "Pending",
    SubmittedDate: new Date().toISOString(),
    WeekCommencing: startOfWeek(new Date(), { weekStartsOn: 1 }).toISOString(),
  },
];

// ─────────────────────────────────────────────
// ANTHROPIC API (via Cloudflare Worker proxy)
// ─────────────────────────────────────────────
async function callClaude(prompt) {
  const res = await fetch(CONFIG.WORKER_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await res.json();
  return data.content?.[0]?.text || "";
}

async function normaliseFishDescriptions(rawText) {
  const priceListDescriptions = PRICE_LIST.map(
    (p) => `${p.code}: ${p.description}`
  ).join("\n");

  const prompt = `You are a seafood product matcher for a UK wholesale fish supplier.

Given this raw fish order text from a sales rep, extract each line item and match it to our standard price list where possible.

RAW TEXT:
${rawText}

STANDARD PRICE LIST:
${priceListDescriptions}

Return ONLY a JSON array. Each element should be:
{
  "rawDescription": "the rep's original wording",
  "matchedCode": "the price list code, or null if no match",
  "matchedDescription": "the matched description, or null",
  "quantity": "numeric quantity if stated, else null",
  "unit": "kg/portions/pieces/etc if stated, else null",
  "confidence": "high/medium/low",
  "customPricing": true or false (true if not on price list)
}

Match liberally — "bass fillet 1 forty to 1 sixty" should match "Sea Bass Fillet 140-160g". Do not include any text outside the JSON array.`;

  const raw = await callClaude(prompt);
  try {
    return JSON.parse(raw.replace(/```json|```/g, "").trim());
  } catch {
    return [];
  }
}

// ─────────────────────────────────────────────
// QUOTE CALCULATION
// ─────────────────────────────────────────────
function calculateQuote(lineItems, deliveriesPerWeek, tier = "B") {
  const lines = lineItems.map((item) => {
    if (item.customPricing || !item.matchedCode) {
      // Custom pricing: flag for manual cost lookup
      return { ...item, unitPrice: null, lineTotal: null, isCustom: true };
    }
    const priceEntry = PRICE_LIST.find((p) => p.code === item.matchedCode);
    if (!priceEntry) return { ...item, unitPrice: null, lineTotal: null, isCustom: true };

    const price = priceEntry[`tier${tier}`];
    const qty = parseFloat(item.quantity) || 1;
    return {
      ...item,
      unitPrice: price,
      lineTotal: price * qty,
      isCustom: false,
    };
  });

  const knownTotal = lines
    .filter((l) => !l.isCustom && l.lineTotal)
    .reduce((sum, l) => sum + l.lineTotal, 0);

  const revenuePerDelivery = deliveriesPerWeek > 0 ? knownTotal / deliveriesPerWeek : 0;
  const belowMinimum = revenuePerDelivery > 0 && revenuePerDelivery < CONFIG.MIN_REVENUE_PER_DELIVERY;
  const hasCustomItems = lines.some((l) => l.isCustom);

  return { lines, knownTotal, revenuePerDelivery, belowMinimum, hasCustomItems, tier };
}

// ─────────────────────────────────────────────
// COMPONENTS
// ─────────────────────────────────────────────

function StatusBadge({ status }) {
  const styles = {
    Pending: { bg: "#fef3c7", color: "#92400e", border: "#fcd34d" },
    Approved: { bg: "#d1fae5", color: "#065f46", border: "#6ee7b7" },
    Bounced: { bg: "#fee2e2", color: "#991b1b", border: "#fca5a5" },
    Processing: { bg: "#e0f2fe", color: "#0c4a6e", border: "#7dd3fc" },
  };
  const s = styles[status] || styles.Pending;
  return (
    <span
      style={{
        padding: "2px 10px",
        borderRadius: 20,
        fontSize: 11,
        fontWeight: 700,
        letterSpacing: "0.05em",
        textTransform: "uppercase",
        background: s.bg,
        color: s.color,
        border: `1px solid ${s.border}`,
      }}
    >
      {status}
    </span>
  );
}

function QuoteLineTable({ lines, tier, onTierChange }) {
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
        <span style={{ fontSize: 12, fontWeight: 600, color: "#6b7280", textTransform: "uppercase", letterSpacing: "0.06em" }}>
          Pricing Tier
        </span>
        {["A", "B", "C"].map((t) => (
          <button
            key={t}
            onClick={() => onTierChange(t)}
            style={{
              padding: "3px 14px",
              borderRadius: 6,
              border: "1.5px solid",
              borderColor: tier === t ? "#0ea5e9" : "#e5e7eb",
              background: tier === t ? "#0ea5e9" : "#fff",
              color: tier === t ? "#fff" : "#374151",
              fontWeight: 600,
              cursor: "pointer",
              fontSize: 13,
            }}
          >
            {t}
          </button>
        ))}
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead>
          <tr style={{ background: "#f8fafc", borderBottom: "2px solid #e2e8f0" }}>
            {["Raw Description", "Matched To", "Qty/Unit", "Unit Price", "Line Total", "Notes"].map((h) => (
              <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 600, color: "#64748b", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {lines.map((line, i) => (
            <tr
              key={i}
              style={{
                borderBottom: "1px solid #f1f5f9",
                background: line.isCustom ? "#fffbeb" : "#fff",
              }}
            >
              <td style={{ padding: "8px 10px", color: "#374151" }}>{line.rawDescription}</td>
              <td style={{ padding: "8px 10px" }}>
                {line.matchedDescription ? (
                  <span style={{ color: "#0f766e", fontWeight: 500 }}>{line.matchedDescription}</span>
                ) : (
                  <span style={{ color: "#9ca3af", fontStyle: "italic" }}>No match</span>
                )}
              </td>
              <td style={{ padding: "8px 10px", color: "#374151" }}>
                {line.quantity ?? "—"} {line.unit ?? ""}
              </td>
              <td style={{ padding: "8px 10px", color: "#374151" }}>
                {line.unitPrice != null ? `£${line.unitPrice.toFixed(2)}` : "—"}
              </td>
              <td style={{ padding: "8px 10px", fontWeight: 600, color: line.lineTotal ? "#0f172a" : "#9ca3af" }}>
                {line.lineTotal != null ? `£${line.lineTotal.toFixed(2)}` : "Pending"}
              </td>
              <td style={{ padding: "8px 10px" }}>
                {line.isCustom && (
                  <span style={{ background: "#fef3c7", color: "#92400e", padding: "1px 8px", borderRadius: 10, fontSize: 11, fontWeight: 600 }}>
                    Custom pricing
                  </span>
                )}
                {!line.isCustom && line.confidence === "low" && (
                  <span style={{ background: "#f3e8ff", color: "#6b21a8", padding: "1px 8px", borderRadius: 10, fontSize: 11, fontWeight: 600 }}>
                    Low confidence
                  </span>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function QuoteCard({ request, onApprove, onBounce }) {
  const [expanded, setExpanded] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [lineItems, setLineItems] = useState(null);
  const [quote, setQuote] = useState(null);
  const [tier, setTier] = useState("B");
  const [bouncedNotes, setBouncedNotes] = useState("");
  const [showBounceForm, setShowBounceForm] = useState(false);
  const [error, setError] = useState(null);

  const processQuote = useCallback(async () => {
    setProcessing(true);
    setError(null);
    try {
      const items = await normaliseFishDescriptions(request.FishTypes);
      setLineItems(items);
      const calc = calculateQuote(items, Number(request.DeliveriesPerWeek), tier);
      setQuote(calc);
    } catch (e) {
      setError("Failed to process. Check your Cloudflare Worker URL and API key.");
    }
    setProcessing(false);
  }, [request, tier]);

  const handleTierChange = (newTier) => {
    setTier(newTier);
    if (lineItems) {
      setQuote(calculateQuote(lineItems, Number(request.DeliveriesPerWeek), newTier));
    }
  };

  return (
    <div
      style={{
        background: "#fff",
        borderRadius: 14,
        border: "1.5px solid #e2e8f0",
        marginBottom: 16,
        overflow: "hidden",
        boxShadow: expanded ? "0 4px 24px rgba(0,0,0,0.07)" : "0 1px 4px rgba(0,0,0,0.04)",
        transition: "box-shadow 0.2s",
      }}
    >
      {/* Header row */}
      <div
        style={{ display: "flex", alignItems: "center", padding: "14px 20px", gap: 14, cursor: "pointer" }}
        onClick={() => setExpanded((e) => !e)}
      >
        <div style={{ flex: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontWeight: 700, fontSize: 16, color: "#0f172a" }}>{request.CustomerName}</span>
            <StatusBadge status={request.Status} />
            {request.IsRequote && (
              <span style={{ background: "#ede9fe", color: "#4c1d95", padding: "1px 8px", borderRadius: 10, fontSize: 11, fontWeight: 700 }}>
                REQUOTE
              </span>
            )}
          </div>
          <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 3 }}>
            Submitted {format(new Date(request.SubmittedDate), "d MMM yyyy")} · {request.DeliveriesPerWeek}×/wk · Est. £{Number(request.EstMonthlySpend).toLocaleString()}/mo · Currently: {request.CurrentSuppliers}
          </div>
        </div>
        <span style={{ color: "#94a3b8", fontSize: 18 }}>{expanded ? "▲" : "▼"}</span>
      </div>

      {/* Expanded body */}
      {expanded && (
        <div style={{ padding: "0 20px 20px", borderTop: "1px solid #f1f5f9" }}>
          {/* Fish request raw text */}
          <div style={{ marginTop: 14, padding: 12, background: "#f8fafc", borderRadius: 8, fontFamily: "monospace", fontSize: 13, color: "#334155" }}>
            <span style={{ fontWeight: 600, color: "#64748b", fontFamily: "inherit" }}>Rep's request: </span>
            {request.FishTypes}
          </div>

          {request.Notes && (
            <div style={{ marginTop: 8, padding: 10, background: "#f0fdf4", borderRadius: 8, fontSize: 13, color: "#166534" }}>
              <strong>Notes:</strong> {request.Notes}
            </div>
          )}

          {/* Process button */}
          {!quote && (
            <button
              onClick={processQuote}
              disabled={processing}
              style={{
                marginTop: 14,
                padding: "9px 22px",
                background: processing ? "#94a3b8" : "#0ea5e9",
                color: "#fff",
                border: "none",
                borderRadius: 8,
                fontWeight: 700,
                fontSize: 14,
                cursor: processing ? "not-allowed" : "pointer",
                display: "flex",
                alignItems: "center",
                gap: 8,
              }}
            >
              {processing ? "⟳ Processing with Claude..." : "⚡ Process Quote"}
            </button>
          )}

          {error && (
            <div style={{ marginTop: 12, padding: 10, background: "#fee2e2", borderRadius: 8, color: "#991b1b", fontSize: 13 }}>
              {error}
            </div>
          )}

          {/* Quote results */}
          {quote && (
            <div style={{ marginTop: 16 }}>
              <QuoteLineTable lines={quote.lines} tier={tier} onTierChange={handleTierChange} />

              {/* Summary row */}
              <div style={{ display: "flex", gap: 16, marginTop: 14, flexWrap: "wrap" }}>
                <div style={{ padding: "10px 16px", background: "#f8fafc", borderRadius: 10, flex: 1 }}>
                  <div style={{ fontSize: 11, color: "#94a3b8", textTransform: "uppercase", fontWeight: 600, letterSpacing: "0.05em" }}>Quoted Total (known)</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: "#0f172a", marginTop: 2 }}>£{quote.knownTotal.toFixed(2)}</div>
                </div>
                <div
                  style={{
                    padding: "10px 16px",
                    background: quote.belowMinimum ? "#fee2e2" : "#f0fdf4",
                    borderRadius: 10,
                    flex: 1,
                    border: quote.belowMinimum ? "1.5px solid #fca5a5" : "none",
                  }}
                >
                  <div style={{ fontSize: 11, color: "#94a3b8", textTransform: "uppercase", fontWeight: 600, letterSpacing: "0.05em" }}>Revenue / Delivery</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: quote.belowMinimum ? "#dc2626" : "#15803d", marginTop: 2 }}>
                    £{quote.revenuePerDelivery.toFixed(2)}
                    {quote.belowMinimum && <span style={{ fontSize: 12, marginLeft: 8, fontWeight: 600 }}>⚠ Below £{CONFIG.MIN_REVENUE_PER_DELIVERY} minimum</span>}
                  </div>
                </div>
                {quote.hasCustomItems && (
                  <div style={{ padding: "10px 16px", background: "#fffbeb", borderRadius: 10, flex: 1, border: "1.5px solid #fcd34d" }}>
                    <div style={{ fontSize: 11, color: "#92400e", textTransform: "uppercase", fontWeight: 600, letterSpacing: "0.05em" }}>Action Needed</div>
                    <div style={{ fontSize: 13, fontWeight: 600, color: "#92400e", marginTop: 4 }}>Custom items require manual cost lookup (36% margin)</div>
                  </div>
                )}
              </div>

              {/* Approve / Bounce */}
              {request.Status === "Pending" && (
                <div style={{ marginTop: 16, display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <button
                    onClick={() => onApprove(request.id, quote)}
                    style={{ padding: "9px 22px", background: "#16a34a", color: "#fff", border: "none", borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: "pointer" }}
                  >
                    ✓ Approve Quote
                  </button>
                  <button
                    onClick={() => setShowBounceForm((s) => !s)}
                    style={{ padding: "9px 22px", background: "#fff", color: "#dc2626", border: "1.5px solid #dc2626", borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: "pointer" }}
                  >
                    ✕ Bounce to Rep
                  </button>
                </div>
              )}

              {showBounceForm && (
                <div style={{ marginTop: 12 }}>
                  <textarea
                    value={bouncedNotes}
                    onChange={(e) => setBouncedNotes(e.target.value)}
                    placeholder="Reason for bouncing back to rep..."
                    style={{ width: "100%", minHeight: 80, padding: 10, borderRadius: 8, border: "1.5px solid #e2e8f0", fontSize: 13, fontFamily: "inherit", boxSizing: "border-box" }}
                  />
                  <button
                    onClick={() => { onBounce(request.id, bouncedNotes); setShowBounceForm(false); }}
                    style={{ marginTop: 8, padding: "8px 18px", background: "#dc2626", color: "#fff", border: "none", borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: "pointer" }}
                  >
                    Send Back to Rep
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────
export default function App() {
  const [requests, setRequests] = useState(MOCK_REQUESTS);
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");

  // TODO: Replace MOCK_REQUESTS with real SharePoint fetch using Graph API
  // useEffect(() => { fetchFromSharePoint().then(setRequests); }, []);

  const handleApprove = (id, quote) => {
    setRequests((prev) =>
      prev.map((r) => (r.id === id ? { ...r, Status: "Approved", _quote: quote } : r))
    );
    // TODO: PATCH status back to SharePoint
  };

  const handleBounce = (id, notes) => {
    setRequests((prev) =>
      prev.map((r) => (r.id === id ? { ...r, Status: "Bounced", BouncedNotes: notes } : r))
    );
    // TODO: PATCH status + notes back to SharePoint
  };

  const filtered = requests.filter((r) => {
    const matchesFilter = filter === "All" || r.Status === filter;
    const matchesSearch = r.CustomerName.toLowerCase().includes(search.toLowerCase());
    return matchesFilter && matchesSearch;
  });

  const counts = {
    All: requests.length,
    Pending: requests.filter((r) => r.Status === "Pending").length,
    Approved: requests.filter((r) => r.Status === "Approved").length,
    Bounced: requests.filter((r) => r.Status === "Bounced").length,
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'DM Sans', system-ui, sans-serif" }}>
      {/* Header */}
      <div style={{ background: "#0c1a2e", padding: "0 32px" }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <span style={{ fontSize: 22 }}>🐟</span>
            <span style={{ color: "#fff", fontWeight: 800, fontSize: 18, letterSpacing: "-0.02em" }}>
              James Knight <span style={{ color: "#38bdf8", fontWeight: 400 }}>· Quote Manager</span>
            </span>
          </div>
          <div style={{ color: "#64748b", fontSize: 13 }}>
            Week of {format(startOfWeek(new Date(), { weekStartsOn: 1 }), "d MMM yyyy")}
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "28px 32px" }}>
        {/* Stats row */}
        <div style={{ display: "flex", gap: 14, marginBottom: 24, flexWrap: "wrap" }}>
          {Object.entries(counts).map(([status, count]) => (
            <div
              key={status}
              onClick={() => setFilter(status)}
              style={{
                padding: "12px 20px",
                background: filter === status ? "#0ea5e9" : "#fff",
                borderRadius: 10,
                cursor: "pointer",
                border: "1.5px solid",
                borderColor: filter === status ? "#0ea5e9" : "#e2e8f0",
                transition: "all 0.15s",
                minWidth: 110,
              }}
            >
              <div style={{ fontSize: 26, fontWeight: 800, color: filter === status ? "#fff" : "#0f172a" }}>{count}</div>
              <div style={{ fontSize: 12, fontWeight: 600, color: filter === status ? "#bae6fd" : "#94a3b8", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                {status}
              </div>
            </div>
          ))}
        </div>

        {/* Search */}
        <div style={{ marginBottom: 18 }}>
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search by customer name..."
            style={{
              width: "100%",
              maxWidth: 380,
              padding: "9px 14px",
              borderRadius: 8,
              border: "1.5px solid #e2e8f0",
              fontSize: 14,
              outline: "none",
              fontFamily: "inherit",
              boxSizing: "border-box",
            }}
          />
        </div>

        {/* Cards */}
        {filtered.length === 0 ? (
          <div style={{ textAlign: "center", padding: 60, color: "#94a3b8", fontSize: 15 }}>
            No quote requests found.
          </div>
        ) : (
          filtered.map((r) => (
            <QuoteCard key={r.id} request={r} onApprove={handleApprove} onBounce={handleBounce} />
          ))
        )}
      </div>
    </div>
  );
}
