import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════════════
   RSSP Student Budget and Finance Tool

   TABS:
     Budgeting  — Weekly Budget, Annual Summary (Sankey)
     Funding    — Loan Options, Loan Summary (Refugee Student Loan)
     Other      — Forecast Model (multi-year cash flow), Tax, Utilities Guide

   PURPOSE: Normalised weekly income-vs-expenses view for RSSP students,
   extended with an annual Sankey summary, a multi-year cash flow forecast
   (HECS-HELP, SSL, Refugee Student Loan, tax, growth assumptions), and a
   standalone Refugee Student Loan repayment calculator.

   PERSISTENCE: Uses the browser's localStorage. All calls are wrapped in
   try/catch — if storage is unavailable (private browsing etc.) the app
   still functions, it just won't remember inputs between visits.

   LOAN ENGINE: The Refugee Student Loan calculation (calcSparkLoanSchedule)
   is shared identically between the Loan Summary tab and the Forecast
   Model tab, so the two always agree. It replicates the methodology in
   Skill Path's own loan calculation spreadsheet: simple interest accrues
   during the grace (study) period and is capitalised once at the start of
   repayment; the balance then amortises via a level monthly payment over a
   fixed 4-year repayment period. HECS-HELP repayment uses the marginal
   system in effect since 1 July 2025 (2026-27 thresholds).

   DATA SOURCES:
     Youth Allowance max rate: $677.20/fn — Services Australia
     Income test: free area $539, taper1 $646, 50c/60c — Services Australia
     Rent Assistance: indexed 20 Mar 2026 — DVA / Services Australia
     SSL: $1,349/semester — Services Australia
     CSP student contribution bands: education.gov.au (2026 rates)
     HECS-HELP repayment thresholds: ATO (2026-27)
     Refugee Student Loan: Skill Path / Spark Finance
     Utility estimates: AER reference bills, ABS household expenditure survey
   ═══════════════════════════════════════════════════════════════════════ */

const STORAGE_KEY = "rssp-budget-assumptions";

/* Living expense categories */
const LC_LABELS = {
  accom: "Accommodation",
  utilities: "Utilities (Electricity, Gas, Water)",
  transport: "Transport",
  food: "Food",
  personal: "Personal",
  clothing: "Clothing",
  entertainment: "Entertainment",
  other: "Other",
};

/* Hints shown below specific living cost input fields */
const LC_HINTS = {
  personal: "(eg. gym membership, haircut, cosmetics, mobile phone)",
  entertainment: "(eg. cinema, live music/sport, sporting clubs, night out)",
  other: "(eg. study materials, medical expenses)",
};

/* Rent Assistance rates — indexed 20 March 2026 */
const RA_TABLE = [
  { key: "single", situation: "Single",        threshold: 154.80, max: 219.40, get ceiling() { return +(this.threshold + this.max / 0.75).toFixed(2); } },
  { key: "sharer", situation: "Single, sharer", threshold: 154.80, max: 146.27, get ceiling() { return +(this.threshold + this.max / 0.75).toFixed(2); } },
];

/* ═══ UTILITY ESTIMATION DATA ═══ */
const ACCOM_TYPES = [
  { key: "student", label: "Purpose-built student accommodation", bundled: true },
  { key: "sharehouse", label: "Sharehouse (3–4 people)" },
  { key: "solo", label: "Solo rental (1-bed apartment)" },
  { key: "family", label: "Living with family / homestay", bundled: true },
];

const CLIMATE_ZONES = [
  { key: "warm", label: "Warm", cities: "Brisbane, Darwin, Cairns, Townsville" },
  { key: "mild", label: "Mild", cities: "Sydney, Perth, Adelaide, Gold Coast" },
  { key: "cool", label: "Cool", cities: "Melbourne, Canberra, Hobart, Geelong" },
];

const UTIL_ELECTRICITY = { sharehouse: 14, solo: 28 };
const UTIL_GAS = {
  sharehouse: { warm: 0, mild: 5, cool: 8 },
  solo:       { warm: 0, mild: 8, cool: 15 },
};
const UTIL_WATER = { sharehouse: 3, solo: 8 };
const UTIL_INTERNET = { sharehouse: 6, solo: 20 };

function getUtilityEstimate(accomType, climateZone) {
  const accom = ACCOM_TYPES.find(a => a.key === accomType);
  if (!accom || accom.bundled) {
    return { electricity: 0, gas: 0, water: 0, internet: 0, total: 0, bundled: true };
  }
  const electricity = UTIL_ELECTRICITY[accomType] || 0;
  const gas = (UTIL_GAS[accomType] && UTIL_GAS[accomType][climateZone]) || 0;
  const water = UTIL_WATER[accomType] || 0;
  const internet = UTIL_INTERNET[accomType] || 0;
  const total = electricity + gas + water + internet;
  return { electricity, gas, water, internet, total, bundled: false };
}

/* ═══ UTILITY DETAIL CONTENT ═══ */
const UTIL_INFO = {
  electricity: {
    title: "Electricity",
    icon: "⚡",
    description: "Electricity is typically the largest utility expense for students. Your usage depends on whether you have electric heating/cooling, how many appliances you use, and the size of your dwelling.",
    details: [
      "Average retail rate in Australia is approximately 30c per kWh plus a daily supply charge of around $1/day.",
      "A 1-bedroom apartment typically uses 3,500–5,000 kWh per year. In a sharehouse, your share is roughly 1,500–2,500 kWh per year.",
      "Air conditioning in summer (or electric heating in winter) is the biggest single driver of electricity costs.",
      "LED lighting, switching off standby appliances, and using cold-water washing can meaningfully reduce your bill.",
    ],
    tipToSave: "Compare energy plans using the government's free Energy Made Easy tool (energymadeeasy.gov.au). Switching retailers can save $200–400/year.",
  },
  gas: {
    title: "Gas",
    icon: "🔥",
    description: "Gas costs vary significantly depending on where you live in Australia. In warmer climates (Brisbane, Darwin), many homes have no gas connection at all. In cooler climates (Melbourne, Canberra, Hobart), gas heating is common and can add substantially to winter bills.",
    details: [
      "In Melbourne and Canberra, gas heating can add $10–20/week in peak winter months, averaging $8–12/week over the year per person.",
      "In Sydney and Perth, gas usage is moderate — mainly for hot water and cooking — averaging $5–8/week per person.",
      "In Brisbane and Darwin, most student accommodation has no gas connection. Hot water and cooking are electric.",
      "If your accommodation has a gas connection, you will pay a daily supply charge (~60–90c/day) regardless of usage.",
    ],
    tipToSave: "If you have gas heating, set the thermostat to 18–20°C. Each degree above 20°C adds roughly 10% to your heating bill. Draught-proofing doors and windows makes a big difference.",
  },
  water: {
    title: "Water",
    icon: "💧",
    description: "Water billing varies by state and lease type. In many apartment complexes and purpose-built student accommodation, water is included in rent. Where it is separately metered, a single person's usage is typically modest.",
    details: [
      "Typical single-person water usage costs $5–10/week where separately metered.",
      "In Victoria, landlords can only charge tenants for water if the property meets water efficiency standards (low-flow showerheads, dual-flush toilets).",
      "In NSW and Queensland, landlords can generally pass on water usage charges (but not the fixed service charge).",
      "Most purpose-built student accommodation includes water in the rent — check your lease.",
    ],
    tipToSave: "Shorter showers are the single biggest water-saving action. A 4-minute shower uses roughly 36 litres vs 90+ litres for a 10-minute shower.",
  },
  internet: {
    title: "Internet",
    icon: "📶",
    description: "Internet is relatively consistent in price across Australia. The main variable is how many people you split the plan with. Purpose-built student accommodation usually includes WiFi in the rent.",
    details: [
      "A standard NBN 50 plan (suitable for most students) costs $70–90/month.",
      "In a sharehouse of 3–4 people, your share is roughly $5–10/week.",
      "Living alone, you pay the full plan cost: approximately $17–22/week.",
      "Some newer apartment complexes include internet in body corporate fees or strata, so check before signing up for a separate plan.",
    ],
    tipToSave: "If you're in a sharehouse, NBN 50 is usually sufficient for 3–4 people streaming and studying simultaneously. Avoid signing up for NBN 100 unless someone needs it for work.",
  },
};

const C = { navy: "#385592", coral: "#de5240", cyan: "#cdf0f1", teal: "#2dcd9e", grey: "#8a92a6" };

const DEFAULTS = () => ({
  livingCosts: { accom: 375, utilities: 0, transport: 25, food: 235, personal: 50, clothing: 20, entertainment: 30, other: 0 },
  hoursPerWeek: 10, hourlyWage: 25, raType: "sharer", otherNote: "",
});

/* ═══ CALCULATION FUNCTIONS ═══ */
function calcRA(fnRent, threshold, maxRA, taper) {
  if (fnRent <= threshold) return 0;
  return Math.min(Math.round((fnRent - threshold) * taper * 100) / 100, maxRA);
}

function calcIncomeTestReduction(fnWages, freeArea, taper1End, taper1Rate, taper2Rate) {
  if (fnWages <= freeArea) return 0;
  if (fnWages <= taper1End) return Math.round((fnWages - freeArea) * taper1Rate * 100) / 100;
  return Math.round(((taper1End - freeArea) * taper1Rate + (fnWages - taper1End) * taper2Rate) * 100) / 100;
}

/* ═══ TAX CALCULATION (2026-27 resident rates) ═══
   Source: ATO resident tax scale for 2026-27, Medicare levy low-income
   thresholds (2026-27), and the Low Income Tax Offset (LITO).
   ASSUMPTIONS (see Tax tab for full explanation):
   - Student is an Australian resident for tax purposes (permanent/humanitarian
     visa holders generally are; this is NOT checked automatically).
   - Taxable income = annual wages + annual Youth Allowance actually received
     (net of any income-test reduction). Rent Assistance is excluded — it is
     not taxable income.
   - No other income (bank interest, scholarships, etc.) is assumed.
   - The Beneficiary Tax Offset (a separate, more generous offset available
     to Centrelink allowance recipients) is NOT applied — only the standard
     tax-free threshold and LITO are used. This means the estimate is
     conservative (likely to overstate tax for allowance recipients).
   - No work-related deductions are assumed.
   ═══════════════════════════════════════════════════════════════ */
const TAX_BRACKETS_2026_27 = [
  { min: 0, max: 18200, base: 0, rate: 0 },
  { min: 18200, max: 45000, base: 0, rate: 0.15 },
  { min: 45000, max: 135000, base: 4020, rate: 0.30 },
  { min: 135000, max: 190000, base: 31020, rate: 0.37 },
  { min: 190000, max: Infinity, base: 51370, rate: 0.45 },
];
const MEDICARE_LOW = 28011, MEDICARE_HIGH = 35014, MEDICARE_RATE = 0.02;

function calcMarginalTax(taxableIncome) {
  const b = TAX_BRACKETS_2026_27.find(b => taxableIncome > b.min && taxableIncome <= b.max) || TAX_BRACKETS_2026_27[0];
  return taxableIncome <= 18200 ? 0 : b.base + (taxableIncome - b.min) * b.rate;
}
function calcLITO(taxableIncome) {
  if (taxableIncome <= 37500) return 700;
  if (taxableIncome <= 45000) return 700 - (taxableIncome - 37500) * 0.05;
  if (taxableIncome <= 66667) return 325 - (taxableIncome - 45000) * 0.015;
  return 0;
}
function calcMedicareLevy(taxableIncome) {
  if (taxableIncome <= MEDICARE_LOW) return 0;
  if (taxableIncome <= MEDICARE_HIGH) return (taxableIncome - MEDICARE_LOW) * 0.10;
  return taxableIncome * MEDICARE_RATE;
}
function calcAnnualTax(taxableIncome) {
  const gross = calcMarginalTax(taxableIncome);
  const lito = Math.min(calcLITO(taxableIncome), gross);
  const netIncomeTax = Math.max(0, gross - lito);
  const medicare = calcMedicareLevy(taxableIncome);
  return { taxableIncome, gross, lito, netIncomeTax, medicare, total: netIncomeTax + medicare };
}

/* ═══ ANNUAL CASH FLOW: HECS-HELP, CSP bands, SSL, Refugee Student Loan ═══
   CSP student contribution bands confirmed current for 2026 (education.gov.au
   2026 indexed rates; studyassist.gov.au). HECS-HELP compulsory repayment
   uses the marginal system in effect since 1 July 2025, with 2026-27
   thresholds (ATO). SSL and Refugee Student Loan figures match this tool's
   own Funding tab.
   ═══════════════════════════════════════════════════════════════ */
const FIELDS = [
  { name: "Agriculture", band: 1, csp: 4738 }, { name: "Clinical Psychology", band: 1, csp: 4738 },
  { name: "Education", band: 1, csp: 4738 }, { name: "English", band: 1, csp: 4738 },
  { name: "Indigenous & Foreign Languages", band: 1, csp: 4738 }, { name: "Mathematics", band: 1, csp: 4738 },
  { name: "Nursing", band: 1, csp: 4738 }, { name: "Statistics", band: 1, csp: 4738 },
  { name: "Allied Health", band: 2, csp: 9537 }, { name: "Architecture", band: 2, csp: 9537 },
  { name: "Built Environment", band: 2, csp: 9537 }, { name: "Computing", band: 2, csp: 9537 },
  { name: "Engineering", band: 2, csp: 9537 }, { name: "Environmental Studies", band: 2, csp: 9537 },
  { name: "Other Health", band: 2, csp: 9537 }, { name: "Pathology", band: 2, csp: 9537 },
  { name: "Science", band: 2, csp: 9537 }, { name: "Surveying", band: 2, csp: 9537 },
  { name: "Visual & Performing Arts", band: 2, csp: 9537 },
  { name: "Dentistry", band: 3, csp: 13558 }, { name: "Medicine", band: 3, csp: 13558 },
  { name: "Veterinary Science", band: 3, csp: 13558 },
  { name: "Accounting", band: 4, csp: 17399 }, { name: "Administration", band: 4, csp: 17399 },
  { name: "Behavioural Science", band: 4, csp: 17399 }, { name: "Commerce", band: 4, csp: 17399 },
  { name: "Communications", band: 4, csp: 17399 }, { name: "Economics", band: 4, csp: 17399 },
  { name: "Law", band: 4, csp: 17399 }, { name: "Society & Culture", band: 4, csp: 17399 },
];
const HECS_INDEXATION_RATE = 0.028; // 2026 indexation (lower of CPI/WPI) — applied 1 June each year to unpaid balance
const HECS_REPAY_BANDS_2026_27 = [
  { min: 0, max: 69528, base: 0, rate: 0 },
  { min: 69528, max: 129717, base: 0, rate: 0.15 },
  { min: 129717, max: 186051, base: 9028.35, rate: 0.17 },
];
function calcHelpRepayment(repaymentIncome, debtBalance) {
  if (debtBalance <= 0 || repaymentIncome <= 69528) return 0;
  let repay;
  if (repaymentIncome > 186051) {
    repay = repaymentIncome * 0.10; // above top threshold, switches to flat 10% of TOTAL repayment income
  } else {
    const b = HECS_REPAY_BANDS_2026_27.find(b => repaymentIncome > b.min && repaymentIncome <= b.max) || HECS_REPAY_BANDS_2026_27[1];
    repay = b.base + (repaymentIncome - b.min) * b.rate;
  }
  return Math.min(Math.round(repay), Math.round(debtBalance));
}
const SSL_PER_PERIOD = 1349, SSL_PERIODS_PER_YEAR = 2, SSL_ANNUAL = SSL_PER_PERIOD * SSL_PERIODS_PER_YEAR;
const SPARK_MAX_AMT = 5000, SPARK_RATE = 0.07, SPARK_FEE_RATE = 0.05, SPARK_REPAY_YEARS_FIXED = 4;
// Grace period = years the student is drawing down (study period). Repayment period is
// a fixed 4 years, confirmed against Skill Path's own loan calculation engine (not a fixed
// "7 year term" — that only holds when the grace period happens to be 3 years).

/* Shared Refugee Student Loan engine — used identically by the Loan Summary tab and the
   Long Term Cash Flow tab, so both always agree on the same numbers.
   Mechanics (matching the loan engine spreadsheet):
   - Each drawdown (assumed 1 Jan of its year) adds principal + a 5% admin fee, both
     immediately added to the outstanding balance.
   - During the grace period, interest accrues as SIMPLE interest each year on the
     then-outstanding balance (not compounded day to day) — tracked separately.
   - At the start of the repayment period, all accrued grace-period interest is
     capitalised into the balance once.
   - The balance is then repaid via a level monthly payment (standard loan amortisation,
     compounding monthly), fixed for a set number of years (default 4).
   Returns a per-year array covering the grace + repayment years. */
function calcSparkLoanSchedule(drawdownsByYear, graceYears, repayYears = SPARK_REPAY_YEARS_FIXED, rate = SPARK_RATE, feeRate = SPARK_FEE_RATE) {
  const rows = [];
  let balance = 0, accruedGraceInterest = 0;
  for (let y = 1; y <= graceYears; y++) {
    const drawdown = drawdownsByYear[y - 1] || 0;
    const fee = drawdown * feeRate;
    balance += drawdown + fee;
    const interest = balance * rate; // simple interest for the year (balance is constant all year, drawdowns only occur on day 1)
    accruedGraceInterest += interest;
    rows.push({ year: y, phase: "grace", drawdown, fee, interest, repayment: 0, principalPaid: 0, balanceEnd: balance, accruedInterestEnd: accruedGraceInterest });
  }
  const capitalisedStart = balance + accruedGraceInterest;
  const r = rate / 12, n = repayYears * 12;
  const pmt = capitalisedStart <= 0 ? 0 : (r === 0 ? capitalisedStart / n : capitalisedStart * r / (1 - Math.pow(1 + r, -n)));
  let bal = capitalisedStart;
  for (let ry = 1; ry <= repayYears; ry++) {
    let yearInterest = 0, yearPrincipal = 0, yearPaid = 0;
    for (let m = 1; m <= 12; m++) {
      if (bal <= 0.005) break;
      const interest = bal * r;
      const payment = Math.min(pmt, bal + interest);
      const principal = payment - interest;
      bal = Math.max(0, bal - principal);
      yearInterest += interest; yearPrincipal += principal; yearPaid += payment;
    }
    rows.push({ year: graceYears + ry, phase: "repay", drawdown: 0, fee: 0, interest: yearInterest, repayment: yearPaid, principalPaid: yearPrincipal, balanceEnd: bal, accruedInterestEnd: accruedGraceInterest });
  }
  const totalDrawn = drawdownsByYear.reduce((s, v) => s + (v || 0), 0);
  const totalFees = totalDrawn * feeRate;
  const totalInterest = accruedGraceInterest + rows.filter(r => r.phase === "repay").reduce((s, r) => s + r.interest, 0);
  const totalRepaid = rows.filter(r => r.phase === "repay").reduce((s, r) => s + r.repayment, 0);
  return { rows, monthlyPMT: pmt, capitalisedStart, totalDrawn, totalFees, totalInterest, totalRepaid };
}

/* ═══ FORMATTING ═══ */
const fmt = v => {
  if (v == null) return "-";
  return v < 0
    ? `($${Math.abs(Math.round(v)).toLocaleString("en-AU")})`
    : `$${Math.round(v).toLocaleString("en-AU")}`;
};
const fmt2 = v => {
  if (v == null) return "-";
  const abs = Math.abs(v).toFixed(2);
  return v < 0 ? `($${abs})` : `$${abs}`;
};

/* ═══ UI COMPONENTS ═══ */
const Inp = ({ label, value, onChange, min, max, step, note, warn, dollar, disabled, placeholder }) => (
  <div className="flex flex-col gap-0.5">
    <label className="text-xs font-medium" style={{ color: C.navy }}>{label}</label>
    <div className={dollar ? "relative" : ""}>
      {dollar && <span className="absolute left-2 top-1/2 -translate-y-1/2 text-sm text-gray-400">$</span>}
      <input type="number" value={value === 0 && placeholder ? "" : value}
        onChange={e => { const raw = e.target.value; onChange(raw === "" ? 0 : parseFloat(raw) || 0); }}
        placeholder={placeholder} min={min} max={max} step={step || 1} disabled={disabled}
        className={`border rounded py-1.5 text-sm bg-white focus:outline-none focus:ring-2 w-full ${dollar ? "pl-6 pr-2" : "px-2"} ${disabled ? "opacity-50" : ""}`}
        style={{ borderColor: warn ? C.coral : "#d1d5db" }} />
    </div>
    {note && <span className={`text-xs ${warn ? "font-medium" : ""}`} style={{ color: warn ? C.coral : "#9ca3af" }}>{note}</span>}
  </div>
);
const Section = ({ title, children }) => (
  <div className="mb-6">
    <div className="flex items-center gap-2 mb-3 pb-1.5" style={{ borderBottom: `2px solid ${C.teal}` }}>
      <h3 className="text-xs font-bold uppercase tracking-wide" style={{ color: C.navy }}>{title}</h3>
    </div>
    {children}
  </div>
);
const SubHeading = ({ children }) => (
  <p className="text-xs font-bold mb-2" style={{ color: C.navy }}>{children}</p>
);
const GreyNote = ({ children }) => (
  <div className="text-xs mb-3" style={{ color: "#6b7280" }}>{children}</div>
);
const SectionDivider = ({ title }) => (
  <div className="mb-4 -mx-5 px-5 py-3" style={{ backgroundColor: `${C.navy}10`, borderTop: `2px solid ${C.navy}30`, borderBottom: `2px solid ${C.navy}30` }}>
    <h2 className="text-sm font-bold uppercase tracking-wide" style={{ color: C.navy }}>{title}</h2>
  </div>
);
const FN = () => <span style={{ fontWeight: 700, textDecoration: "underline" }}>Fortnightly</span>;

const LOGO = "data:image/png;base64,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";

/* ═══ UTILITY INFO CARD ═══ */
const UtilityCard = ({ info, estimate, isBundled }) => {
  const [open, setOpen] = useState(false);
  return (
    <div className="rounded-lg border overflow-hidden mb-4" style={{ borderColor: "#e5e7eb" }}>
      <button onClick={() => setOpen(!open)} className="w-full flex items-center justify-between p-4 text-left transition hover:bg-gray-50">
        <div className="flex items-center gap-3">
          <span className="text-2xl">{info.icon}</span>
          <div>
            <h4 className="text-sm font-bold" style={{ color: C.navy }}>{info.title}</h4>
            <p className="text-xs" style={{ color: "#6b7280" }}>
              {isBundled ? "Typically included in your rent" : `Estimated: $${estimate}/week`}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-3">
          {!isBundled && (
            <span className="font-mono font-bold text-sm px-3 py-1 rounded" style={{ backgroundColor: `${C.teal}20`, color: C.navy }}>
              ${estimate}/wk
            </span>
          )}
          {isBundled && (
            <span className="text-xs font-semibold px-3 py-1 rounded" style={{ backgroundColor: C.cyan, color: C.navy }}>
              $0/wk
            </span>
          )}
          <svg className={`w-4 h-4 transition-transform ${open ? "rotate-180" : ""}`} style={{ color: C.navy }} fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" /></svg>
        </div>
      </button>
      {open && (
        <div className="px-4 pb-4 border-t" style={{ borderColor: "#f3f4f6" }}>
          <p className="text-xs mt-3 mb-3" style={{ color: "#374151" }}>{info.description}</p>
          <div className="space-y-2 mb-3">
            {info.details.map((d, i) => (
              <div key={i} className="flex gap-2 text-xs" style={{ color: "#6b7280" }}>
                <span style={{ color: C.teal, flexShrink: 0 }}>•</span>
                <span>{d}</span>
              </div>
            ))}
          </div>
          <div className="p-3 rounded-lg text-xs" style={{ backgroundColor: "#f0fdf4", border: "1px solid #bbf7d0" }}>
            <span className="font-semibold" style={{ color: "#065f46" }}>Tip to save: </span>
            <span style={{ color: "#065f46" }}>{info.tipToSave}</span>
          </div>
        </div>
      )}
    </div>
  );
};

/* ═══ SANKEY CHART (Income vs Expenses, annualised) ═══
   Pure display/derivation layer: every $ value it receives is passed in
   from the same `budget` calculation used by the Budget tab. No income-
   test, YA, or RA formulas are re-implemented here.
   Drag-to-adjust: dragging the Accommodation or Part-time Work bar edges
   calls back up to App, which updates the SAME state (livingCosts.accom /
   hoursPerWeek) used everywhere else — so the Budget tab, exports, etc.
   all stay in sync automatically.
   ═══════════════════════════════════════════════════════════════ */
const fmtAnnual = v => `$${Math.round(Math.max(0, v)).toLocaleString("en-AU")}`;

function ribbonPath(x0, y0top, y0bot, x1, y1top, y1bot) {
  const mx = (x0 + x1) / 2;
  return `M${x0},${y0top} C${mx},${y0top} ${mx},${y1top} ${x1},${y1top} L${x1},${y1bot} C${mx},${y1bot} ${mx},${y0bot} ${x0},${y0bot} Z`;
}

function layoutSankeyColumn(nodes, total, top, drawHeight, pad) {
  const scale = total > 0 ? (drawHeight - pad * Math.max(0, nodes.length - 1)) / total : 0;
  let y = top, raw = 0;
  return nodes.map(n => {
    const h = Math.max(n.value * scale, 0);
    const out = { ...n, y0: y, y1: y + h, height: h, scale, raw0: raw, raw1: raw + n.value };
    y += h + pad;
    raw += n.value;
    return out;
  });
}

/* ═══ LOAN WATERFALL CHART ═══
   Renders the same visual pattern as the reference Skill Path loan chart:
   drawdowns/fees/interest step the balance up during the grace period,
   interest/repayments step it back down during the repayment period. */
const LOAN_COLORS = { draw: "#8bc34a", interest: "#e05d44", fee: "#374151", repay: "#5c6bc0" };
function LoanWaterfallChart({ rows, graceYears }) {
  const width = 900, height = 520, top = 40, bottom = 420, left = 34;
  const n = rows.length;
  const plotWidth = width - left - 20;
  const barW = 0.78 * plotWidth / n;
  const gap = n > 1 ? (0.22 * plotWidth) / (n - 1) : 0;

  // Build floating segments by walking a running cumulative balance.
  // Grace years: drawdown/fee/interest stack in sequence (non-overlapping ranges).
  // Repayment years: Repayment and Interest are drawn as two side-by-side bars
  // within the year's column (not stacked at the same x) — otherwise the larger
  // Repayment bar fully paints over the smaller Interest bar and hides it.
  let cum = 0;
  const segs = [];
  const barCenters = [];
  rows.forEach((r, i) => {
    const x = left + i * (barW + gap);
    barCenters.push(x + barW / 2);
    if (r.phase === "grace") {
      if (r.drawdown > 0) { const s = cum; cum += r.drawdown; segs.push({ x, w: barW, from: s, to: cum, color: LOAN_COLORS.draw, big: true }); }
      if (r.fee > 0) { const s = cum; cum += r.fee; segs.push({ x, w: barW, from: s, to: cum, color: LOAN_COLORS.fee }); }
      if (r.interest > 0) { const s = cum; cum += r.interest; segs.push({ x, w: barW, from: s, to: cum, color: LOAN_COLORS.interest }); }
    } else {
      const subGap = barW * 0.08, repayW = barW * 0.60, intW = barW - repayW - subGap;
      if (r.repayment > 0) { const s = cum; cum -= r.repayment; segs.push({ x, w: repayW, from: cum, to: s, color: LOAN_COLORS.repay, big: true }); }
      if (r.interest > 0) { const s = cum; cum += r.interest; segs.push({ x: x + repayW + subGap, w: intW, from: s, to: cum, color: LOAN_COLORS.interest }); }
    }
  });

  const maxVal = Math.max(1, ...segs.map(s => Math.max(s.from, s.to))) * 1.12;
  const scale = (bottom - top) / maxVal;
  const yPix = v => bottom - v * scale;

  const legend = [["draw", "Draw down"], ["interest", "Interest"], ["fee", "Fee"], ["repay", "Repayment"]];
  const graceEndX = left + (graceYears - 1) * (barW + gap) + barW + gap / 2;

  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" style={{ maxWidth: width, display: "block", margin: "0 auto" }}>
      {/* Legend */}
      {legend.map(([key, label], i) => (
        <g key={key} transform={`translate(${left + i * 170}, 12)`}>
          <rect width={18} height={14} y={0} fill={LOAN_COLORS[key]} rx={2} />
          <text x={25} y={12} fontSize="15" fontWeight="700" fill="#374151">{label}</text>
        </g>
      ))}

      {/* Bars */}
      {segs.map((s, i) => (
        <rect key={i} x={s.x} y={yPix(Math.max(s.from, s.to))} width={s.w} height={Math.max(1, Math.abs(s.to - s.from) * scale)} fill={s.color} rx={2} />
      ))}
      {/* Big-segment value labels */}
      {segs.filter(s => s.big).map((s, i) => (
        <text key={i} x={s.x + s.w / 2} y={(yPix(s.from) + yPix(s.to)) / 2 + 5} textAnchor="middle" fontSize="14" fontWeight="700" fill="white">
          {Math.abs(s.to - s.from) * scale > 30 ? `$${Math.round(Math.abs(s.to - s.from)).toLocaleString("en-AU")}` : ""}
        </text>
      ))}

      {/* X-axis year labels */}
      {rows.map((r, i) => (
        <text key={r.year} x={barCenters[i]} y={bottom + 24} textAnchor="middle" fontSize="15" fontWeight="600" fill="#374151">{r.year}</text>
      ))}
      <line x1={left} y1={bottom} x2={width - 20} y2={bottom} stroke="#374151" strokeWidth={1} />

      {/* Phase brackets */}
      <line x1={left} y1={bottom + 52} x2={graceEndX} y2={bottom + 52} stroke="#9ca3af" strokeWidth={1} />
      <text x={(left + graceEndX) / 2} y={bottom + 47} textAnchor="middle" fontSize="14" fill="#6b7280">Course Duration ("Grace Period")</text>
      <line x1={graceEndX} y1={bottom + 52} x2={width - 20} y2={bottom + 52} stroke="#9ca3af" strokeWidth={1} />
      <text x={(graceEndX + width - 20) / 2} y={bottom + 47} textAnchor="middle" fontSize="14" fill="#6b7280">Repayment Period</text>
      <text x={(left + width - 20) / 2} y={bottom + 74} textAnchor="middle" fontSize="15" fontWeight="700" fill="#374151">Years</text>
    </svg>
  );
}

function SankeyChart({ accom, food, other, tax, ra, ya, wages, loan, savings, hourlyWage, onAccomAnnualChange, onWagesAnnualChange }) {
  const width = 940, height = 480, top = 46, drawHeight = 380, pad = 10;
  const leftX = 150, rightX = 620, barW = 130;

  const leftBase = [
    { key: "accom", label: "Accommodation", value: accom, color: C.navy, draggable: true },
    { key: "food", label: "Food", value: food, color: C.navy },
    { key: "other", label: "Other", value: other, color: C.navy },
  ];
  if (tax > 1) leftBase.push({ key: "tax", label: "Income Tax*", value: tax, color: "#6b6f9e" });
  if (savings > 1) leftBase.push({ key: "savings", label: "Net Savings", value: savings, color: C.teal });

  const rightBase = [
    { key: "ra", label: "Rent Assistance", value: ra, color: C.grey, group: "Government" },
    { key: "ya", label: "Youth Allowance", value: ya, color: C.grey, group: "Government" },
    { key: "wages", label: "Part-time Work", value: wages, color: C.coral, group: "Student", draggable: true },
  ];
  if (loan > 1) rightBase.push({ key: "loan", label: "Loan Required", value: loan, color: C.coral, group: "Student", emphasise: true });

  const total = Math.max(1, leftBase.reduce((s, n) => s + n.value, 0));
  const leftNodes = layoutSankeyColumn(leftBase, total, top, drawHeight, pad);
  const rightNodes = layoutSankeyColumn(rightBase, total, top, drawHeight, pad);

  const links = [];
  leftNodes.forEach(l => {
    rightNodes.forEach(r => {
      const lo = Math.max(l.raw0, r.raw0), hi = Math.min(l.raw1, r.raw1);
      if (hi - lo > 0.5) {
        links.push({
          key: `${l.key}-${r.key}`,
          sy0: l.y0 + (lo - l.raw0) * l.scale, sy1: l.y0 + (hi - l.raw0) * l.scale,
          ty0: r.y0 + (lo - r.raw0) * r.scale, ty1: r.y0 + (hi - r.raw0) * r.scale,
        });
      }
    });
  });

  const startDrag = (e, { scale, startAnnual, onChange }) => {
    e.preventDefault(); e.stopPropagation();
    const startY = e.clientY;
    const move = ev => onChange(Math.max(0, startAnnual + (ev.clientY - startY) / scale));
    const up = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", up); };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" style={{ maxWidth: width, display: "block", margin: "0 auto" }}>
      <text x={leftX + barW / 2} y={20} textAnchor="middle" fontSize="13" fontWeight="700" fill={C.navy}>Living Expenses</text>
      <text x={rightX + barW / 2} y={20} textAnchor="middle" fontSize="13" fontWeight="700" fill={C.navy}>Income Source</text>

      {links.map(l => (
        <path key={l.key} d={ribbonPath(leftX + barW, l.sy0, l.sy1, rightX, l.ty0, l.ty1)} fill="#e9ecf1" opacity={0.75} />
      ))}

      {leftNodes.map(n => (
        <g key={n.key}>
          <rect x={leftX} y={n.y0} width={barW} height={Math.max(n.height, 1)} fill={n.color} rx={3} />
          <text x={leftX - 10} y={(n.y0 + n.y1) / 2} textAnchor="end" dominantBaseline="middle" fontSize="12" fontWeight="700" fill={C.navy}>{n.label}</text>
          {n.height > 16 && (
            <text x={leftX + barW / 2} y={(n.y0 + n.y1) / 2 + 4} textAnchor="middle" fontSize="12" fontWeight="700" fill="white">{fmtAnnual(n.value)}</text>
          )}
          {n.draggable && (
            <g style={{ cursor: "ns-resize" }} onPointerDown={e => startDrag(e, { scale: n.scale, startAnnual: n.value, onChange: onAccomAnnualChange })}>
              <rect x={leftX} y={n.y1 - 5} width={barW} height={10} fill="transparent" />
              <rect x={leftX + barW / 2 - 14} y={n.y1 - 2.5} width={28} height={5} rx={2.5} fill="white" opacity={0.9} />
            </g>
          )}
        </g>
      ))}

      {rightNodes.map(n => (
        <g key={n.key}>
          <rect x={rightX} y={n.y0} width={barW} height={Math.max(n.height, 1)} fill={n.color} rx={3} />
          <text x={rightX + barW + 10} y={(n.y0 + n.y1) / 2} textAnchor="start" dominantBaseline="middle"
            fontSize="12" fontWeight={n.emphasise ? 800 : 600} fill={n.emphasise ? "#111827" : "#6b7280"}>{n.label}</text>
          {n.height > 16 && (
            <text x={rightX + barW / 2} y={(n.y0 + n.y1) / 2 + 4} textAnchor="middle" fontSize="12" fontWeight="700" fill="white">{fmtAnnual(n.value)}</text>
          )}
          {n.draggable && (
            <g style={{ cursor: hourlyWage > 0 ? "ns-resize" : "not-allowed" }}
              onPointerDown={e => hourlyWage > 0 && startDrag(e, { scale: n.scale, startAnnual: n.value, onChange: onWagesAnnualChange })}>
              <rect x={rightX} y={n.y1 - 5} width={barW} height={10} fill="transparent" />
              <rect x={rightX + barW / 2 - 14} y={n.y1 - 2.5} width={28} height={5} rx={2.5} fill="white" opacity={0.9} />
            </g>
          )}
        </g>
      ))}
    </svg>
  );
}

/* ═══ MAIN APPLICATION ═══ */
export default function App() {
  const [tab, setTab] = useState("howto");
  const [loaded, setLoaded] = useState(false);

  const [livingCosts, setLivingCosts] = useState(DEFAULTS().livingCosts);
  const [hoursPerWeek, setHoursPerWeek] = useState(DEFAULTS().hoursPerWeek);
  const [hourlyWage, setHourlyWage] = useState(DEFAULTS().hourlyWage);
  const [raType, setRaType] = useState(DEFAULTS().raType);
  const [otherNote, setOtherNote] = useState("");

  /* ═══ Annual Cash Flow tab state ═══ */
  const [studyYears, setStudyYears] = useState(3);
  const [fieldOfStudy, setFieldOfStudy] = useState("Computing");
  const [ssafFee, setSsafFee] = useState(365);
  const [uniContribAmount, setUniContribAmount] = useState(15000);
  const [wageGrowth, setWageGrowth] = useState(0.03);
  const [cfInflation, setCfInflation] = useState(0.03);
  const [graduateSalary, setGraduateSalary] = useState(75000);
  const [preArrivalCost, setPreArrivalCost] = useState(2500);
  const [sslYears, setSslYears] = useState([false, false, false]);
  const [sparkYears, setSparkYears] = useState([false, false, false]);
  const [sparkAmts, setSparkAmts] = useState([5000, 5000, 5000]);
  const toggleSslYear = i => setSslYears(prev => { const n = [...prev]; n[i] = !n[i]; return n; });
  const toggleSparkYear = i => setSparkYears(prev => { const n = [...prev]; n[i] = !n[i]; return n; });
  const updateSparkAmt = (i, v) => setSparkAmts(prev => { const n = [...prev]; n[i] = Math.min(Math.max(0, v), SPARK_MAX_AMT); return n; });

  /* ═══ Loan Summary tab state (standalone — does not affect the Budget or Cash Flow tabs) ═══ */
  const [loanYears, setLoanYears] = useState([true, true, true]);
  const [loanAmts, setLoanAmts] = useState([5000, 5000, 5000]);
  const toggleLoanYear = i => setLoanYears(prev => { const n = [...prev]; n[i] = !n[i]; return n; });
  const updateLoanAmt = (i, v) => setLoanAmts(prev => { const n = [...prev]; n[i] = Math.min(Math.max(0, v), SPARK_MAX_AMT); return n; });
  const [expandedCF, setExpandedCF] = useState(new Set());
  const toggleCF = k => setExpandedCF(prev => { const n = new Set(prev); n.has(k) ? n.delete(k) : n.add(k); return n; });

  const [utilAccomType, setUtilAccomType] = useState("sharehouse");
  const [utilClimate, setUtilClimate] = useState("cool");

  /* Funding summary state (new in v7) */
  const [furnitureCost, setFurnitureCost] = useState(0);
  const [estimatedSavingsInput, setEstimatedSavingsInput] = useState(0);

  const yaMaxRate = 677.20;
  const freeArea = 539;
  const taper1End = 646;
  const taper1Rate = 0.50;
  const taper2Rate = 0.60;

  const raRow = RA_TABLE.find(r => r.key === raType) || RA_TABLE[0];

  /* ── Load ──
     Uses localStorage (a standard browser API) so this works when deployed as a
     real website. Wrapped in try/catch: if storage is unavailable (private
     browsing, disabled cookies, etc.) the app still works, it just won't persist. */
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const s = JSON.parse(raw);
        if (s.livingCosts) setLivingCosts({ ...DEFAULTS().livingCosts, ...s.livingCosts });
        if (s.hoursPerWeek != null) setHoursPerWeek(s.hoursPerWeek);
        if (s.hourlyWage != null) setHourlyWage(s.hourlyWage);
        if (s.raType) setRaType(s.raType);
        if (s.otherNote != null) setOtherNote(s.otherNote);
        if (s.utilAccomType) setUtilAccomType(s.utilAccomType);
        if (s.utilClimate) setUtilClimate(s.utilClimate);
        if (s.furnitureCost != null) setFurnitureCost(s.furnitureCost);
        if (s.estimatedSavingsInput != null) setEstimatedSavingsInput(s.estimatedSavingsInput);
        if (s.studyYears != null) setStudyYears(s.studyYears);
        if (s.fieldOfStudy) setFieldOfStudy(s.fieldOfStudy);
        if (s.ssafFee != null) setSsafFee(s.ssafFee);
        if (s.uniContribAmount != null) setUniContribAmount(s.uniContribAmount);
        if (s.wageGrowth != null) setWageGrowth(s.wageGrowth);
        if (s.cfInflation != null) setCfInflation(s.cfInflation);
        if (s.graduateSalary != null) setGraduateSalary(s.graduateSalary);
        if (s.preArrivalCost != null) setPreArrivalCost(s.preArrivalCost);
        if (Array.isArray(s.sslYears)) setSslYears(s.sslYears);
        if (Array.isArray(s.sparkYears)) setSparkYears(s.sparkYears);
        if (Array.isArray(s.sparkAmts)) setSparkAmts(s.sparkAmts);
        if (Array.isArray(s.loanYears)) setLoanYears(s.loanYears);
        if (Array.isArray(s.loanAmts)) setLoanAmts(s.loanAmts);
      }
    } catch (_) { /* storage unavailable — continue with defaults */ }
    setLoaded(true);
  }, []);

  /* ── Save ── */
  const save = useCallback((snapshot) => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot)); } catch (_) { /* storage unavailable */ }
  }, []);

  useEffect(() => {
    if (!loaded) return;
    save({
      livingCosts, hoursPerWeek, hourlyWage, raType, otherNote,
      utilAccomType, utilClimate, furnitureCost, estimatedSavingsInput,
      studyYears, fieldOfStudy, ssafFee, uniContribAmount, wageGrowth, cfInflation,
      graduateSalary, preArrivalCost, sslYears, sparkYears, sparkAmts, loanYears, loanAmts,
    });
  }, [livingCosts, hoursPerWeek, hourlyWage, raType, otherNote, utilAccomType, utilClimate,
      furnitureCost, estimatedSavingsInput, studyYears, fieldOfStudy, ssafFee, uniContribAmount,
      wageGrowth, cfInflation, graduateSalary, preArrivalCost, sslYears, sparkYears, sparkAmts,
      loanYears, loanAmts, loaded, save]);

  /* ── Reset ── */
  const resetAll = () => {
    const d = DEFAULTS();
    setLivingCosts(d.livingCosts);
    setHoursPerWeek(d.hoursPerWeek);
    setHourlyWage(d.hourlyWage);
    setRaType(d.raType);
    setOtherNote(d.otherNote);
    setFurnitureCost(0);
    setEstimatedSavingsInput(0);
    setStudyYears(3);
    setFieldOfStudy("Computing");
    setSsafFee(365);
    setUniContribAmount(15000);
    setWageGrowth(0.03);
    setCfInflation(0.03);
    setGraduateSalary(75000);
    setPreArrivalCost(2500);
    setSslYears([false, false, false]);
    setSparkYears([false, false, false]);
    setSparkAmts([5000, 5000, 5000]);
    setLoanYears([true, true, true]);
    setLoanAmts([5000, 5000, 5000]);
    try { localStorage.removeItem(STORAGE_KEY); } catch (_) { /* storage unavailable */ }
  };

  /* ── Derived ── */
  const upLC = (k, v) => setLivingCosts(prev => ({ ...prev, [k]: v }));
  const weeklyAccom = livingCosts.accom || 0;
  const weeklyWages = hoursPerWeek * hourlyWage;
  const fnAccomDisplay = weeklyAccom * 2;
  const currentRA = calcRA(fnAccomDisplay, raRow.threshold, raRow.max, 0.75);
  const rentalBond = weeklyAccom * 4;
  const rentInAdvance = weeklyAccom * 2;

  const utilEstimate = useMemo(() => getUtilityEstimate(utilAccomType, utilClimate), [utilAccomType, utilClimate]);

  const applyUtilityEstimate = () => {
    upLC("utilities", utilEstimate.total);
    setTab("budget");
  };

  /* ── Budget calculation ── */
  const budget = useMemo(() => {
    const fnWages = weeklyWages * 2;
    const fnAccom = weeklyAccom * 2;
    const craPerFn = calcRA(fnAccom, raRow.threshold, raRow.max, 0.75);
    const combinedMaxFn = yaMaxRate + craPerFn;
    const incTestRedFn = calcIncomeTestReduction(fnWages, freeArea, taper1End, taper1Rate, taper2Rate);
    const netGovPerFn = Math.max(0, combinedMaxFn - incTestRedFn);
    const weeklyGov = netGovPerFn / 2;
    const weeklyYAMax = yaMaxRate / 2;
    const weeklyRAMax = craPerFn / 2;
    const weeklyReduction = Math.min(incTestRedFn / 2, (yaMaxRate + craPerFn) / 2);
    const expenseItems = Object.entries(LC_LABELS).map(([k, label]) => ({ key: k, label, amount: livingCosts[k] || 0 }));
    const totalExpense = expenseItems.reduce((s, item) => s + item.amount, 0);
    const totalIncome = weeklyWages + weeklyGov;
    const net = totalIncome - totalExpense;
    return { weeklyGov, weeklyYAMax, weeklyRAMax, weeklyReduction, expenseItems, totalExpense, totalIncome, net };
  }, [livingCosts, raRow, weeklyWages, weeklyAccom]);

  /* ── CSV Download ── */
  const downloadXLSX = () => {

    /* Helper: sanitize unicode chars for Excel compatibility */
    const san = s => String(s ?? "").replace(/\u2014/g, "-").replace(/\u2013/g, "-").replace(/\u2018|\u2019/g, "'").replace(/\u201c|\u201d/g, '"');

    /* ── Build rows ── */
    const data = [];
    const push = (...args) => data.push(args.map(a => typeof a === "string" ? san(a) : a));
    const blank = () => data.push([]);

    /* Header */
    push("Student Budget Summary");
    push("All amounts are WEEKLY and are in A$");
    push("Export Date", new Date().toLocaleDateString("en-AU"));
    blank();

    /* Budget Summary */
    push("BUDGET SUMMARY");
    blank();
    push("", "Income", "");
    push("", "  Part-time wages (" + hoursPerWeek + " hrs x $" + hourlyWage + "/hr)", Math.round(weeklyWages));
    push("", "  Government payment (net)", Math.round(budget.weeklyGov));
    push("", "    Youth Allowance (max)", Math.round(budget.weeklyYAMax));
    push("", "    Rent Assistance (max)", Math.round(budget.weeklyRAMax));
    push("", "    Less: Income test reduction", budget.weeklyReduction > 0 ? -Math.round(budget.weeklyReduction) : 0);
    push("", "Total Weekly Income", Math.round(budget.totalIncome));
    blank();
    push("", "Living Expenses", "");
    const accomSuffix = raType === "single" ? " (single)" : " (sharing)";
    budget.expenseItems.forEach(item => {
      const label = item.key === "accom" ? san(item.label) + accomSuffix : san(item.label);
      push("", "  " + label, item.amount);
    });
    push("", "Total Weekly Living Expenses", budget.totalExpense);
    blank();
    push("", "Net Weekly Amount", Math.round(budget.net));
    blank();

    /* ── Create workbook ── */
    const ws = XLSX.utils.aoa_to_sheet(data);

    /* Column widths */
    ws["!cols"] = [
      { wch: 5 },   /* A: narrow — section headings bleed into B */
      { wch: 42 },  /* B: item labels */
      { wch: 14 },  /* C: amounts */
    ];

    /* Number format for currency columns — apply to column C */
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    for (let r = range.s.r; r <= range.e.r; r++) {
      const addr = XLSX.utils.encode_cell({ r, c: 2 });
      const cell = ws[addr];
      if (cell && typeof cell.v === "number") {
        cell.z = '"$"#,##0';
      }
    }

    /* ── Sheet 2: Long Term Cash Flow ── */
    const cf = [];
    const pushCF = (...args) => cf.push(args.map(a => typeof a === "string" ? san(a) : a));
    const blankCF = () => cf.push([]);

    pushCF("RSSP Long Term Cash Flow");
    pushCF("All amounts are ANNUAL and are in A$");
    pushCF("Export Date", new Date().toLocaleDateString("en-AU"));
    blankCF();

    pushCF("ASSUMPTIONS");
    pushCF("Field of Study", fieldOfStudy, `Band ${cfFld.band} — $${cfFld.csp.toLocaleString("en-AU")}/yr (2026 CSP rate)`);
    pushCF("Course Length (years)", studyYears);
    pushCF("SSAF (annual, deferred via SA-HELP)", ssafFee);
    pushCF("University Contribution ($, Year 1 — assumed toward accommodation)", uniContribAmount);
    pushCF("Wage Growth (annual)", wageGrowth);
    pushCF("Inflation (annual)", cfInflation);
    pushCF(`Graduate Salary (Year ${studyYears + 1})`, graduateSalary);
    pushCF("Pre-Arrival Costs (Year 1)", preArrivalCost);
    pushCF("SSL Years Selected", sslYears.map((v, i) => v ? `Year ${i + 1}` : null).filter(Boolean).join(", ") || "None");
    pushCF("Refugee Student Loan Years Selected", sparkYears.map((v, i) => v ? `Year ${i + 1} ($${sparkAmts[i].toLocaleString("en-AU")})` : null).filter(Boolean).join(", ") || "None");
    blankCF();

    pushCF("PROJECTION");
    blankCF();
    const cfHeaderRow = ["Line Item", ...cashflowRows.map(r => `Year ${r.year} (${r.isStudy ? "Study" : "Working"})`)];
    cf.push(cfHeaderRow.map(h => san(h)));
    const cfFirstDataRow = cf.length; // 0-indexed row number of first CF_ROWS line, for number formatting below
    CF_ROWS.forEach(row => {
      if (row.hdr) { cf.push([san(row.label || "")]); return; }
      const label = "  ".repeat((row.indent || 0) + 1) + row.label;
      const vals = cashflowRows.map(r => {
        const raw = r[row.f];
        return Math.round(row.note && raw > 0 ? -raw : raw);
      });
      cf.push([san(label), ...vals]);
    });

    const wsCF = XLSX.utils.aoa_to_sheet(cf);
    wsCF["!cols"] = [{ wch: 44 }, ...cashflowRows.map(() => ({ wch: 16 }))];
    const rangeCF = XLSX.utils.decode_range(wsCF["!ref"] || "A1");
    for (let r = cfFirstDataRow; r <= rangeCF.e.r; r++) {
      for (let c = 1; c <= rangeCF.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = wsCF[addr];
        if (cell && typeof cell.v === "number") cell.z = '"$"#,##0';
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Budget Summary");
    XLSX.utils.book_append_sheet(wb, wsCF, "Long Term Cash Flow");
    XLSX.writeFile(wb, "RSSP_Budget_Summary.xlsx");
  };

  const isDeficit = budget.net < 0;
  const lcTotal = Object.keys(LC_LABELS).reduce((s, k) => s + (livingCosts[k] || 0), 0);

  /* Funding summary calculations */
  const totalUpfrontCosts = rentalBond + rentInAdvance + furnitureCost;
  const fundingGap = totalUpfrontCosts - estimatedSavingsInput;
  const TAB_GROUPS = [
    { group: "Guide", tabs: [["howto", "How to Use"]] },
    { group: "Budgeting", tabs: [["budget", "Weekly Budget"], ["sankey", "Annual Summary"]] },
    { group: "Funding", tabs: [["funding", "Loan Options"], ["loan", "Loan Summary"]] },
    { group: "Other", tabs: [["cashflow", "Forecast Model"], ["tax", "Tax"], ["utilities", "Utilities Guide"]] },
  ];
  const TAB_NUMBERS = { budget: 1, sankey: 2, funding: 3, loan: 4, cashflow: 5, tax: 6, utilities: 7 };
  const TABS = TAB_GROUPS.flatMap(g => g.tabs);

  /* ── Sankey derived values (annualised; reuses `budget` outputs only) ── */
  const netYAWeekly = Math.max(0, budget.weeklyYAMax - budget.weeklyReduction);
  const netRAWeekly = Math.max(0, budget.weeklyGov - netYAWeekly);
  const otherWeekly = Math.max(0, lcTotal - weeklyAccom - (livingCosts.food || 0));
  const annualWagesForTax = weeklyWages * 52;
  const annualYAForTax = netYAWeekly * 52;
  const taxableIncome = annualWagesForTax + annualYAForTax;
  const taxResult = calcAnnualTax(taxableIncome);
  const annualTax = taxResult.total;
  const totalAnnualExpenseInclTax = lcTotal * 52 + annualTax;
  const totalAnnualIncome = budget.totalIncome * 52;
  const netAnnualInclTax = totalAnnualIncome - totalAnnualExpenseInclTax;
  const sankeyAnnual = {
    accom: weeklyAccom * 52,
    food: (livingCosts.food || 0) * 52,
    other: otherWeekly * 52,
    tax: annualTax,
    ra: netRAWeekly * 52,
    ya: netYAWeekly * 52,
    wages: weeklyWages * 52,
    loan: Math.max(0, -netAnnualInclTax),
    savings: Math.max(0, netAnnualInclTax),
  };
  const handleAccomAnnualChange = (annual) => upLC("accom", Math.round((annual / 52) * 100) / 100);
  const handleWagesAnnualChange = (annual) => {
    if (hourlyWage <= 0) return;
    const newWeeklyWage = annual / 52;
    const newHours = Math.min(48, Math.max(0, Math.round((newWeeklyWage / hourlyWage) * 10) / 10));
    setHoursPerWeek(newHours);
  };

  /* ═══ ANNUAL CASH FLOW — multi-year projection ═══
     Reuses THIS tool's own weekly living-cost, RA/YA, and tax logic as the
     study-year baseline (no separate RA/YA/tax formulas are reimplemented
     here). Living costs, and Youth Allowance/Rent Assistance, are inflated
     annually. Wages during study are held flat (fixed part-time job);
     post-study salary grows from Graduate Salary at Wage Growth. */
  const cfFld = FIELDS.find(f => f.name === fieldOfStudy) || FIELDS[0];
  const cfProjectionYears = studyYears + 7;

  // Year 1: the University Contribution is assumed to pay toward accommodation,
  // so the student's own rent (for Rent Assistance purposes) is reduced by it.
  const y1WeeklyRentForRA = Math.max(0, weeklyAccom - uniContribAmount / 52);
  const y1FnRent = y1WeeklyRentForRA * 2;
  const y1FnWages = weeklyWages * 2;
  const y1CraPerFn = calcRA(y1FnRent, raRow.threshold, raRow.max, 0.75);
  const y1IncTestRedFn = calcIncomeTestReduction(y1FnWages, freeArea, taper1End, taper1Rate, taper2Rate);
  const y1NetYAFn = Math.max(0, yaMaxRate - y1IncTestRedFn);
  const y1NetTotalFn = Math.max(0, (yaMaxRate + y1CraPerFn) - y1IncTestRedFn);
  const y1NetRAFn = Math.max(0, y1NetTotalFn - y1NetYAFn);
  const y1NetYAWeekly = y1NetYAFn / 2;
  const y1NetRAWeekly = y1NetRAFn / 2;

  // Refugee Student Loan schedule, computed once via the shared engine (same
  // engine used by the standalone Loan Summary tab) and looked up by year below.
  const cfSparkDrawdowns = useMemo(() => [0, 1, 2].map(i => sparkYears[i] ? (sparkAmts[i] || SPARK_MAX_AMT) : 0),
    [sparkYears[0], sparkYears[1], sparkYears[2], sparkAmts[0], sparkAmts[1], sparkAmts[2]]);
  const cfSparkSchedule = useMemo(() => calcSparkLoanSchedule(cfSparkDrawdowns, studyYears),
    [cfSparkDrawdowns, studyYears]);

  /* Loan Summary tab: standalone use of the same shared engine, fixed 3-year grace
     (matching the Year 1/2/3 drawdown choices exposed on that tab) */
  const loanDrawdowns = useMemo(() => [0, 1, 2].map(i => loanYears[i] ? (loanAmts[i] || 0) : 0),
    [loanYears[0], loanYears[1], loanYears[2], loanAmts[0], loanAmts[1], loanAmts[2]]);
  const loanSchedule = useMemo(() => calcSparkLoanSchedule(loanDrawdowns, 3), [loanDrawdowns]);

  const cashflowRows = useMemo(() => {
    const rows = [];
    let helpDebt = 0;      // combined HECS-HELP + SSL balance

    for (let y = 1; y <= cfProjectionYears; y++) {
      const isStudy = y <= studyYears;
      const infF = Math.pow(1 + cfInflation, y - 1);
      const livingCostsYear = (weeklyAccom + otherWeekly + (livingCosts.food || 0)) * 52 * infF;

      let wages, yaAmount, raAmount, taxableIncomeYear;
      if (isStudy) {
        wages = weeklyWages * 52;
        if (y === 1) {
          yaAmount = y1NetYAWeekly * 52;
          raAmount = y1NetRAWeekly * 52;
        } else {
          yaAmount = netYAWeekly * 52 * infF; // indexed to inflation each year
          raAmount = netRAWeekly * 52 * infF;
        }
        taxableIncomeYear = wages + yaAmount; // RA excluded from taxable income
      } else {
        const wgF = Math.pow(1 + wageGrowth, y - studyYears - 1);
        wages = graduateSalary * wgF;
        yaAmount = 0; raAmount = 0;
        taxableIncomeYear = wages;
      }
      const govPayment = yaAmount + raAmount;
      const taxDetail = calcAnnualTax(taxableIncomeYear);
      const taxYear = taxDetail.total;
      const uniContribYear = y === 1 ? uniContribAmount : 0;
      const preArrivalYear = y === 1 ? preArrivalCost : 0;

      // HECS/HELP debt: accrues CSP+SSAF (+SSL) during study; indexed & repaid after
      let helpRepay = 0;
      const sslThisYear = (isStudy && sslYears[y - 1]) ? SSL_ANNUAL : 0;
      if (isStudy) {
        helpDebt += cfFld.csp + ssafFee + sslThisYear;
      } else {
        helpDebt *= (1 + HECS_INDEXATION_RATE);
        helpRepay = calcHelpRepayment(wages, helpDebt);
        helpDebt = Math.max(0, helpDebt - helpRepay);
      }

      // Refugee Student Loan: look up this year's row from the shared engine schedule
      const sparkRow = cfSparkSchedule.rows.find(r => r.year === y);
      const sparkDrawdownYear = sparkRow ? sparkRow.drawdown : 0;
      const sparkFeeYear = sparkRow ? sparkRow.fee : 0;
      const sparkRepayYear = sparkRow ? sparkRow.repayment : 0;
      const sparkBalEndYear = sparkRow ? (sparkRow.phase === "grace" ? sparkRow.balanceEnd + sparkRow.accruedInterestEnd : sparkRow.balanceEnd) : 0;

      const cashInBeforeFinancing = wages + govPayment + uniContribYear;
      const cashOutBeforeFinancing = taxYear + helpRepay + livingCostsYear + preArrivalYear;
      const netCashflowBeforeFinancing = cashInBeforeFinancing - cashOutBeforeFinancing;
      const netFinancing = sslThisYear + sparkDrawdownYear - sparkRepayYear;
      const netCashflowAfterFinancing = netCashflowBeforeFinancing + netFinancing;

      rows.push({
        year: y, isStudy,
        wages, yaAmount, raAmount, govPayment, uniContribYear, cashInBeforeFinancing,
        taxableIncomeYear,
        taxGross: taxDetail.gross, taxLito: taxDetail.lito, taxMedicare: taxDetail.medicare, tax: taxYear,
        helpRepay, livingCostsYear, preArrivalYear, cashOutBeforeFinancing,
        totalDeductions: taxYear + helpRepay,
        netCashflowBeforeFinancing,
        sslThisYear, sparkDrawdownYear, sparkFeeYear, sparkRepayYear, netFinancing,
        netCashflowAfterFinancing,
        helpDebtEnd: Math.round(Math.max(0, helpDebt)),
        sparkBalEnd: Math.round(Math.max(0, sparkBalEndYear)),
        totalDebtEnd: Math.round(Math.max(0, helpDebt) + Math.max(0, sparkBalEndYear)),
      });
    }
    let cum = 0;
    rows.forEach(r => { cum += r.netCashflowAfterFinancing; r.cumulativeCashflow = cum; });
    return rows;
  }, [studyYears, cfInflation, wageGrowth, graduateSalary, preArrivalCost, ssafFee, uniContribAmount,
      sslYears, cfSparkSchedule, cfFld, weeklyAccom, otherWeekly, livingCosts.food, weeklyWages,
      netYAWeekly, netRAWeekly, cfProjectionYears, y1NetYAWeekly, y1NetRAWeekly]);



  const CF_ROWS = [
    { key: "h_i", label: "INCOME", hdr: true },
    { key: "gov_ya", label: "Youth Allowance", f: "yaAmount", parent: "gov_net", indent: 1, sup: 2 },
    { key: "gov_ra", label: "Rent Assistance", f: "raAmount", parent: "gov_net", indent: 1 },
    { key: "gov_net", label: "Government Payment (YA + Rent Assistance)", f: "govPayment", expandable: true, tip: "Click to see YA/RA detail" },
    { key: "w", label: "Wages", f: "wages" },
    { key: "uc", label: "University Contribution", f: "uniContribYear" },
    { key: "ti", label: "Total Income", f: "cashInBeforeFinancing", sub: true },

    { key: "h_e", label: "EXPENDITURE", hdr: true },
    { key: "pa", label: "Pre-Arrival Costs", f: "preArrivalYear" },
    { key: "lc", label: "Living Costs", f: "livingCostsYear" },

    { key: "h_d", label: "DEDUCTIONS", hdr: true },
    { key: "tx_inc", label: "Taxable Income", f: "taxableIncomeYear", parent: "tx", indent: 1, muted: true },
    { key: "tx_gross", label: "Income Tax (gross)", f: "taxGross", parent: "tx", indent: 1, muted: true },
    { key: "tx_lito", label: "Less: LITO", f: "taxLito", parent: "tx", indent: 1, note: true },
    { key: "tx_medicare", label: "Medicare Levy", f: "taxMedicare", parent: "tx", indent: 1, muted: true },
    { key: "tx", label: "Tax", f: "tax", expandable: true, tip: "Click to see tax detail" },
    { key: "hc", label: "HECS/HELP Repayment", f: "helpRepay", sup: 1 },
    { key: "td", label: "Total Deductions", f: "totalDeductions", sub: true },

    { key: "h_n", label: "", hdr: true },
    { key: "net_op", label: "Net Cash Flow Before Financing", f: "netCashflowBeforeFinancing", total: true },

    { key: "h_f", label: "FINANCING", hdr: true, linkTab: "funding" },
    { key: "h_ssl", label: "Student Start-up Loan", hdr: true, subhdr: true },
    { key: "ssl_dd", label: "SSL Drawdown", f: "sslThisYear", indent: 1, checkbox: "ssl" },
    { key: "h_spark", label: "Refugee Student Loan", hdr: true, subhdr: true },
    { key: "spark_dd", label: "Refugee Loan Drawdown", f: "sparkDrawdownYear", indent: 1, checkbox: "spark" },
    { key: "spark_rp", label: "Refugee Loan Repayment", f: "sparkRepayYear", indent: 1 },
    { key: "nf", label: "Net Financing", f: "netFinancing", sub: true },

    { key: "h_n2", label: "", hdr: true },
    { key: "net", label: "Net Cash Flow (incl. Financing)", f: "netCashflowAfterFinancing", total: true },
    { key: "cum", label: "Cumulative Cash Flow", f: "cumulativeCashflow", cum: true },

    { key: "h_debt", label: "DEBT POSITIONS (YEAR END)", hdr: true },
    { key: "debt", label: "HECS/HELP Debt Outstanding", f: "helpDebtEnd", debt: true, sup: 1 },
    { key: "spark_bal", label: "Refugee Loan Balance", f: "sparkBalEnd", debt: true, sup: 3 },
    { key: "total_debt", label: "Total Debt", f: "totalDebtEnd", debt: true, sub: true },
  ];


  return (
    <div className="min-h-screen" style={{ backgroundColor: "#f8f9fb" }}>
      <div className="max-w-4xl mx-auto p-4">
        {/* Header */}
        <div className="rounded-lg p-4 mb-4 flex items-center justify-between" style={{ background: `linear-gradient(135deg, ${C.navy} 0%, #4a6aaa 100%)` }}>
          <div>
            <h1 className="text-lg font-bold text-white">RSSP Student Budget and Finance Tool</h1>
            <p className="text-xs text-white opacity-70">All currency values are in Australian dollars</p>
          </div>
          <img src={LOGO} alt="Skill Path" style={{ height: 48, objectFit: "contain" }} />
        </div>

        {/* Top line: Guide tab (left) + action buttons (right) */}
        <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
          <button onClick={() => setTab("howto")}
            className="px-3 py-1.5 text-xs rounded transition whitespace-nowrap"
            style={{ backgroundColor: tab === "howto" ? C.grey : "transparent", color: tab === "howto" ? "white" : C.grey, fontWeight: tab === "howto" ? 700 : 600, border: `1px solid ${C.grey}50` }}>
            How to Use
          </button>
          <div className="flex gap-2">
            <button onClick={downloadXLSX} className="px-3 py-1.5 text-xs rounded transition"
              style={{ backgroundColor: "#f3f4f6", color: C.navy, border: "1px solid #e5e7eb" }}>
              Download to Spreadsheet
            </button>
            <button onClick={resetAll} className="px-3 py-1.5 text-xs rounded transition"
              style={{ backgroundColor: "#f3f4f6", color: C.navy, border: "1px solid #e5e7eb" }}>
              Reset to Base-Case Assumptions
            </button>
          </div>
        </div>

        {/* Tab bar: the 3 model-tab groups */}
        <div className="mb-4 border-b overflow-x-auto" style={{ borderColor: "#e5e7eb" }}>
          <div className="flex flex-nowrap items-end gap-2">
            {TAB_GROUPS.filter((g) => g.group !== "Guide").map((g) => {
              const gc = { Budgeting: C.navy, Funding: C.teal, Other: C.coral }[g.group];
              return (
                <div key={g.group} className="flex-1 flex flex-col items-center gap-1 rounded-t-lg px-2 pt-1.5 pb-1" style={{ minWidth: 0, backgroundColor: `${gc}12`, border: `1px solid ${gc}35`, borderBottom: "none" }}>
                  <span className="text-[9px] font-bold uppercase tracking-wider text-center whitespace-nowrap" style={{ color: gc }}>{g.group}</span>
                  <div className="flex gap-0 justify-center flex-nowrap">
                    {g.tabs.map(([k, l]) => (
                      <button key={k} onClick={() => setTab(k)}
                        className="px-1.5 py-1.5 text-xs rounded transition whitespace-nowrap flex-shrink-0"
                        style={{ backgroundColor: tab === k ? gc : "transparent", color: tab === k ? "white" : gc, fontWeight: tab === k ? 700 : 600 }}>
                        {TAB_NUMBERS[k]}. {l}
                      </button>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>


        {/* ═══ TAB: HOW TO USE ═══ */}
        {tab === "howto" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="How to Use This Tool" />
            <GreyNote>
              This tool has been developed for participants in the Refugee Student Settlement Pathway (RSSP). It helps students understand and plan for the financial aspects of their study in Australia. Students can work through the tabs below in order as each one builds on the last. Inputs are saved automatically in your browser, so you can come back and adjust them at any time.
              <br /><br />
              Please note the disclaimer at the bottom of this page.
            </GreyNote>

            <Section title="1. Weekly Budget">
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>What it does:</strong> A simple weekly budget comparing your income against your weekly living expenses.
              </p>
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>Inputs:</strong> Weekly living expenses and income from part-time work. Government payments are pre-populated based on current rates assuming all students receive Youth Allowance.
              </p>
              <p className="text-xs mb-3" style={{ color: C.navy }}>
                <strong>Outputs:</strong> Weekly income and expense breakdown and a net weekly surplus or deficit.
              </p>
              <p className="text-xs italic" style={{ color: "#6b7280" }}>
                Note that this page provides a simple weekly budget model. It doesn't include any pre-arrival (eg. flights) or one-off settlement costs, or any tax obligations. Other pages in the model provide this additional detail. The page is pre-populated with base-case assumptions which need to be modified by the user.
              </p>
            </Section>

            <Section title="2. Annual Summary">
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>What it does:</strong> Converts your Weekly Budget assumptions into an annual view, including income tax.
              </p>
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>Inputs:</strong> No direct inputs. It uses the assumptions already entered on the Weekly Budget tab. You can however adjust your accommodation and part-time work assumptions on this page using the "sliders" on those bars to see how this changes your government payments, tax, and surplus/deficit.
              </p>
              <p className="text-xs mb-3" style={{ color: C.navy }}>
                <strong>Outputs:</strong> An annual income vs. expense summary that includes income tax obligations. A "deficit" is presented as a loan requirement; a "surplus" is presented as savings.
              </p>
              <p className="text-xs italic" style={{ color: "#6b7280" }}>
                Note that this page does not include any pre-arrival costs (e.g. flights) or other one-off costs. These one-off costs may in future periods include costs associated with moving into a rental property once you leave university accommodation. For example, you may be required to pay a refundable rental bond (typically 4 weeks of rent) and 2 weeks of rent in advance.
              </p>
            </Section>

            <Section title="3. Loan Options">
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>What it does:</strong> Provides information about two loan options that you may want to consider to help cover any funding gap, including upfront settlement costs.
              </p>
              <p className="text-xs" style={{ color: C.navy }}>
                <strong>Inputs / Outputs:</strong> No student-specific inputs or outputs are on this page.
              </p>
            </Section>

            <Section title="4. Loan Summary">
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>What it does:</strong> Shows the detailed repayment profile of the Refugee Student Loan for various drawdown assumptions. This tab is standalone and does not affect the Weekly Budget or Forecast Model tabs.
              </p>
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>Inputs:</strong> Which years you draw down the loan, and the amount drawn in each year (up to $5,000/year).
              </p>
              <p className="text-xs mb-3" style={{ color: C.navy }}>
                <strong>Outputs:</strong> Year-by-year balance, interest, and repayment schedule for the loan.
              </p>
            </Section>

            <Section title="5. Forecast Model">
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>What it does:</strong> Presents a multi-year financial summary based on your long term assumptions.
              </p>
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>Inputs:</strong> Automatically takes your weekly budget inputs and allows you to add one-off pre-arrival and settlement expenses (included in Year 1), any university payments you may receive, your tuition fees (deferred via the HECS-HELP loan scheme), drawdowns from the Refugee Student Loan Program, and assumptions about your post-graduation salary.
              </p>
              <p className="text-xs" style={{ color: C.navy }}>
                <strong>Outputs:</strong> A 10 year cash flow model summarising all aspects covered by this budget and finance tool.
              </p>
            </Section>

            <Section title="6. Tax">
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>What it does:</strong> Explains the Australian income tax system as it applies to students, and how the tax figures elsewhere in the tool are calculated.
              </p>
              <p className="text-xs mb-3" style={{ color: C.navy }}>
                <strong>Inputs / Outputs:</strong> No separate inputs — shows the tax calculated from your Weekly Budget assumptions (wages and Youth Allowance).
              </p>
            </Section>

            <Section title="7. Utilities Guide">
              <p className="text-xs mb-2" style={{ color: C.navy }}>
                <strong>What it does:</strong> Provides reference information to help you estimate electricity, gas, and water costs when completing the Weekly Budget tab. These costs will not be incurred when living in on-campus university-provided accommodation only, and are therefore provided as a guide for when you move into your own rental accommodation.
              </p>
              <p className="text-xs" style={{ color: C.navy }}>
                <strong>Inputs / Outputs:</strong> No inputs or outputs of its own.
              </p>
            </Section>

            <Section title="Download to Spreadsheet">
              <p className="text-xs" style={{ color: C.navy }}>
                The <strong>Download to Spreadsheet</strong> button (top right of the page) exports your Weekly Budget and Forecast Model data into an Excel file, so you can review, share, or work with your figures outside the tool.
              </p>
            </Section>

            <div className="mt-2 p-4 rounded text-xs" style={{ backgroundColor: "#fef3c7", color: "#92400e", border: "1px solid #fcd34d" }}>
              <strong>Disclaimer:</strong> This budget and finance tool is provided to students in the RSSP as a guide to assist them in understanding the potential costs associated with studying in Australia. The model is indicative only, is not comprehensive and has not taken into account the student's personal situation. Students should seek professional advice based on their individual circumstances.
            </div>
          </div>
        )}

        {/* ═══ TAB: BUDGET ═══ */}
        {tab === "budget" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Assumptions" />

            <Section title="Living Expenses (Weekly)">
              <GreyNote>
                You can compare your assumed living expenses to those in the Government's cost of living calculator for students which is available{" "}
                <a href="https://costofliving.studyaustralia.gov.au" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>here</a>.
                {" "}For help estimating your utility costs, see the <button onClick={() => setTab("utilities")} className="font-semibold underline" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer" }}>Utilities Guide</button> tab.
              </GreyNote>
              <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-3">
                {Object.entries(LC_LABELS).map(([k, label]) => (
                  <Inp key={k} label={label} value={livingCosts[k]} onChange={v => upLC(k, v)} min={0} step={1} dollar placeholder="0" note={LC_HINTS[k]} />
                ))}
              </div>
              <div className="flex justify-between items-center p-2 rounded text-sm" style={{ backgroundColor: C.cyan }}>
                <span className="font-medium" style={{ color: C.navy }}>Total Weekly</span>
                <span className="font-mono font-semibold" style={{ color: C.navy }}>${lcTotal.toLocaleString()}/wk</span>
              </div>
            </Section>

            <Section title="Income Sources">
              <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-3">
                <Inp label="Part time work — hours/week" value={hoursPerWeek} onChange={v => setHoursPerWeek(v)} min={0} max={48} placeholder="0" />
                <Inp label="Hourly wage" value={hourlyWage} onChange={v => setHourlyWage(v)} step={0.5} dollar placeholder="0" />
              </div>
              <p className="text-xs" style={{ color: "#6b7280" }}>
                If you start working or your income from work changes you will need to notify Centrelink. See{" "}
                <a href="https://www.servicesaustralia.gov.au/when-to-report-your-income-to-centrelink?context=43916" target="_blank" rel="noopener noreferrer"
                  style={{ color: C.navy, textDecoration: "underline", fontWeight: 700 }}>
                  this page
                </a>{" "}for how to report changes in income to Centrelink.
              </p>
            </Section>

            <Section title="Youth Allowance">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <p className="mb-2">Youth Allowance is an Australian Government payment which students receive with the amount paid depending on their situation. This budget model assumes you are between the ages of 15–25, single with no children. If you are older than 25 you will likely be eligible for Austudy payments which are similar.</p>
                <p className="mb-2">The Youth Allowance payments reduce in accordance with a <span style={{ textDecoration: "underline" }}>personal income test</span>. You can earn up to $539 per fortnight before your payment is affected (the "income free area"). For each dollar earned between $539 and $646, your combined payment reduces by 50 cents. For each dollar above $646, your combined payment reduces by 60 cents. The reduction is applied to your total payment from the Government (including any Rent Assistance you may be entitled to — see below).</p>
                <p className="mb-2">An "Income Bank" allows you to accumulate unused income free area credits in low-income fortnights to offset higher-income fortnights.</p>
                <p>More info:{" "}<a href="https://www.servicesaustralia.gov.au/youth-allowance" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>Youth Allowance</a></p>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                <Inp label="Youth Allowance Max (fortnightly)" value={yaMaxRate} onChange={() => {}} step={0.1} dollar disabled note="Single, no children, 18+, away from home" />
              </div>
            </Section>

            <Section title="Rent Assistance">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <SubHeading>How Rent Assistance Works</SubHeading>
                <p className="mb-1">Rent Assistance is an additional government payment for eligible students.</p>
                <p className="mb-1">Rent Assistance is calculated based on the rent you pay. For every $1 of fortnightly rent you pay above the relevant rent threshold, you receive 75 cents in Rent Assistance, up to a maximum amount.</p>
                <p className="mb-1">Rent Assistance is not subject to a separate income test. It is added to your Youth Allowance payment to form a combined maximum rate. The personal income test reduction is then applied to this combined total. This means your Rent Assistance is only affected once the income test reduction exceeds your base Youth Allowance amount.</p>
                <p className="mb-4">More info from the Australian Government{" "}<a href="https://www.servicesaustralia.gov.au/how-much-rent-assistance-you-can-get?context=22206" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>here</a>.</p>

                <SubHeading>Rent Assistance and Purpose Built Student Accommodation (PBSA)</SubHeading>
                <p className="mb-1">
                  Students paying for accommodation in a PBSA property should be eligible for Rent Assistance, which will most likely be considered "Shared" accommodation.
                </p>
                <p className="mb-1">
                  If the property includes catering (regular meals), a portion of the total payment may need to be excluded so that only the rental amount is included. Centrelink uses a <strong>two-thirds rule</strong> to calculate the rent component when a student is paying "Board and Lodging" which includes regular meals.
                </p>
                <p>
                  You will need to speak to Centrelink to confirm the appropriate payments.
                </p>
              </div>

              <div className="mb-4">
                <label className="text-xs font-medium block mb-2" style={{ color: C.navy }}>Your accommodation type</label>
                <div className="flex gap-3">
                  {RA_TABLE.map(row => (
                    <button key={row.key} onClick={() => setRaType(row.key)} className="px-4 py-2 rounded text-sm font-medium transition"
                      style={{ backgroundColor: raType === row.key ? C.navy : "#f3f4f6", color: raType === row.key ? "white" : C.navy, border: `2px solid ${raType === row.key ? C.navy : "#e5e7eb"}` }}>
                      {row.situation}
                    </button>
                  ))}
                </div>
              </div>
              <div className="overflow-x-auto mb-4">
                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ backgroundColor: C.cyan }}>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>If you're</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>Your <FN /> rent is more than</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>To get the maximum payment your <FN /> rent is at least</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>The maximum <FN /> payment is</th>
                    </tr>
                  </thead>
                  <tbody>
                    {RA_TABLE.map(row => (
                      <tr key={row.key} style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: raType === row.key ? `${C.teal}15` : "transparent" }}>
                        <td className="p-2" style={{ fontWeight: raType === row.key ? 600 : 400 }}>{row.situation}</td>
                        <td className="p-2 font-mono">{fmt2(row.threshold)}</td>
                        <td className="p-2 font-mono">{fmt2(row.ceiling)}</td>
                        <td className="p-2 font-mono font-semibold">{fmt2(row.max)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="p-3 rounded-lg mb-4" style={{ backgroundColor: `${C.teal}10`, border: `1px solid ${C.teal}` }}>
                <h4 className="text-xs font-bold mb-1" style={{ color: C.navy }}>Your Estimated Maximum Rent Assistance <span className="font-normal">(before income test)</span></h4>
                <p className="text-xs mb-3" style={{ color: "#6b7280" }}>
                  Your estimated rent assistance payments based on the rent and accommodation type is shown below. This is <span style={{ fontWeight: 700, textDecoration: "underline" }}>BEFORE</span> any reduction due to the income test which is calculated in the budget summary below.
                </p>
                <div className="grid grid-cols-2 gap-x-8 gap-y-2 text-xs max-w-md">
                  <div><span className="text-gray-500">Weekly rent</span><div className="font-mono font-semibold" style={{ color: C.navy }}>{fmt2(weeklyAccom)}</div></div>
                  <div><span className="text-gray-500">Fortnightly rent</span><div className="font-mono font-semibold" style={{ color: C.navy }}>{fmt2(fnAccomDisplay)}</div></div>
                  <div><span className="text-gray-500">Rent Assistance (weekly)</span><div className="font-mono font-semibold" style={{ color: C.teal }}>{fmt2(currentRA / 2)}</div></div>
                  <div><span className="text-gray-500">Rent Assistance (fortnightly)</span><div className="font-mono font-semibold" style={{ color: C.teal }}>{fmt2(currentRA)}</div></div>
                </div>
                {fnAccomDisplay <= raRow.threshold && weeklyAccom > 0 && (
                  <p className="text-xs mt-2" style={{ color: C.coral }}>Your fortnightly rent of {fmt2(fnAccomDisplay)} is below the threshold of {fmt2(raRow.threshold)}. You do not qualify for Rent Assistance.</p>
                )}
                {weeklyAccom === 0 && (
                  <p className="text-xs mt-2" style={{ color: "#9ca3af" }}>Enter your weekly accommodation cost above to calculate your Rent Assistance entitlement.</p>
                )}
              </div>
            </Section>

            {/* ═══ BUDGET SUMMARY ═══ */}
            <SectionDivider title="Budget Summary" />
            <GreyNote>This shows your weekly ongoing budget based on the assumptions above without any university contribution or one-off income sources.</GreyNote>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead><tr style={{ backgroundColor: C.navy }}>
                  <th className="text-left px-4 py-2.5 font-semibold text-white">Item</th>
                  <th className="text-right px-4 py-2.5 font-semibold text-white w-32">Weekly Amount</th>
                </tr></thead>
                <tbody>
                  <tr><td colSpan={2} className="px-4 py-2 font-bold text-xs uppercase tracking-wider" style={{ backgroundColor: `${C.teal}15`, color: C.navy }}>Income</td></tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Part-time wages ({hoursPerWeek} hrs × ${hourlyWage}/hr)</td>
                    <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(Math.round(weeklyWages))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Government payment (net)</td>
                    <td className="px-4 py-2 text-right font-mono font-semibold" style={{ color: C.navy }}>{fmt(Math.round(budget.weeklyGov))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: "#fafafa" }}>
                    <td className="px-4 py-1.5 text-xs" style={{ color: "#6b7280", paddingLeft: 44 }}>Youth Allowance (max)</td>
                    <td className="px-4 py-1.5 text-right font-mono text-xs" style={{ color: "#6b7280" }}>{fmt(Math.round(budget.weeklyYAMax))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: "#fafafa" }}>
                    <td className="px-4 py-1.5 text-xs" style={{ color: "#6b7280", paddingLeft: 44 }}>Rent Assistance (max)</td>
                    <td className="px-4 py-1.5 text-right font-mono text-xs" style={{ color: "#6b7280" }}>{fmt(Math.round(budget.weeklyRAMax))}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6", backgroundColor: "#fafafa" }}>
                    <td className="px-4 py-1.5 text-xs" style={{ color: budget.weeklyReduction > 0 ? C.coral : "#6b7280", paddingLeft: 44 }}>Less: Income test reduction</td>
                    <td className="px-4 py-1.5 text-right font-mono text-xs" style={{ color: budget.weeklyReduction > 0 ? C.coral : "#6b7280" }}>
                      {budget.weeklyReduction > 0 ? `($${Math.round(budget.weeklyReduction).toLocaleString("en-AU")})` : fmt(0)}
                    </td>
                  </tr>
                  <tr style={{ borderTop: "2px solid #e5e7eb", backgroundColor: "#f0fdf4" }}>
                    <td className="px-4 py-2 font-bold" style={{ color: C.navy }}>Total Weekly Income</td>
                    <td className="px-4 py-2 text-right font-mono font-bold" style={{ color: C.navy }}>{fmt(Math.round(budget.totalIncome))}</td>
                  </tr>
                  <tr><td colSpan={2} className="py-1"></td></tr>
                  <tr><td colSpan={2} className="px-4 py-2 font-bold text-xs uppercase tracking-wider" style={{ backgroundColor: `${C.coral}10`, color: C.navy }}>Living Expenses</td></tr>
                  {budget.expenseItems.map((item, i) => (
                    <tr key={`exp-${i}`} style={{ borderBottom: "1px solid #f3f4f6" }}>
                      <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>
                        {item.label}
                      </td>
                      <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(item.amount)}</td>
                    </tr>
                  ))}
                  <tr style={{ borderTop: "2px solid #e5e7eb", backgroundColor: "#fef2f2" }}>
                    <td className="px-4 py-2 font-bold" style={{ color: C.navy }}>Total Weekly Living Expenses</td>
                    <td className="px-4 py-2 text-right font-mono font-bold" style={{ color: C.navy }}>{fmt(budget.totalExpense)}</td>
                  </tr>
                  <tr><td colSpan={2} className="py-1"></td></tr>
                  <tr style={{ backgroundColor: isDeficit ? `${C.coral}15` : `${C.teal}15`, borderTop: `3px solid ${isDeficit ? C.coral : C.teal}` }}>
                    <td className="px-4 py-3 font-bold" style={{ color: isDeficit ? C.coral : "#065f46" }}>Net Weekly Amount</td>
                    <td className="px-4 py-3 text-right font-mono font-bold text-base" style={{ color: isDeficit ? C.coral : "#065f46" }}>{fmt(Math.round(budget.net))}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="mt-5 p-4 rounded-lg" style={{ backgroundColor: isDeficit ? "#fef2f2" : "#f0fdf4", border: `2px solid ${isDeficit ? C.coral : C.teal}` }}>
              <p className="text-sm font-bold" style={{ color: isDeficit ? C.coral : "#065f46" }}>
                Based on your assumptions you are spending {isDeficit ? "more" : "less"} than you earn by ${Math.abs(Math.round(budget.net)).toLocaleString("en-AU")} per week.
              </p>
            </div>
          </div>
        )}

        {/* ═══ TAB: INCOME VS EXPENSES (SANKEY) ═══ */}
        {tab === "sankey" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Income vs Expenses (Annual)" />
            <GreyNote>
              This page converts your weekly assumptions to annual figures.
              <br /><br />
              To investigate different scenarios you can slide the white "handles" at the bottom of the <strong style={{ color: C.navy }}>Accommodation</strong> or <strong style={{ color: C.coral }}>Part-time Work</strong> bars. This will adjust all components dynamically (e.g. increasing income from part-time work will reduce Youth Allowance in accordance with Centrelink's income test, increase tax payable, and reduce the loan required — or increase savings).
              <br /><br />
              For simplicity, the Food and Other Expense categories remain as per the Weekly Budget tab but can be updated there.
              <br /><br />
              The Assumptions on the Weekly Budget tab will update to reflect the chosen scenario if these sliders are used.
            </GreyNote>
            <SankeyChart
              accom={sankeyAnnual.accom}
              food={sankeyAnnual.food}
              other={sankeyAnnual.other}
              tax={sankeyAnnual.tax}
              ra={sankeyAnnual.ra}
              ya={sankeyAnnual.ya}
              wages={sankeyAnnual.wages}
              loan={sankeyAnnual.loan}
              savings={sankeyAnnual.savings}
              hourlyWage={hourlyWage}
              onAccomAnnualChange={handleAccomAnnualChange}
              onWagesAnnualChange={handleWagesAnnualChange}
            />
            <p className="text-xs font-bold text-center mt-4 mb-2" style={{ color: C.navy }}>Weekly Amounts</p>
            <div className="flex flex-wrap gap-4 justify-center text-xs" style={{ color: "#6b7280" }}>
              <span>Accommodation: <strong style={{ color: C.navy }}>${weeklyAccom.toLocaleString("en-AU")}</strong></span>
              <span>Food: <strong style={{ color: C.navy }}>${(livingCosts.food || 0).toLocaleString("en-AU")}</strong></span>
              <span>Other Expenses: <strong style={{ color: C.navy }}>${otherWeekly.toLocaleString("en-AU")}</strong></span>
              <span>Rent Assistance: <strong style={{ color: C.navy }}>${Math.round(netRAWeekly).toLocaleString("en-AU")}</strong></span>
              <span>Youth Allowance: <strong style={{ color: C.navy }}>${Math.round(netYAWeekly).toLocaleString("en-AU")}</strong></span>
              <span>Weekly Wages <strong style={{ color: C.navy }}>${Math.round(weeklyWages).toLocaleString("en-AU")}</strong> ({hoursPerWeek} Hours/week @ ${hourlyWage}/hr)</span>
              {sankeyAnnual.loan > 1 && <span>Annual shortfall (loan): <strong style={{ color: C.coral }}>{fmtAnnual(sankeyAnnual.loan)}</strong></span>}
              {sankeyAnnual.savings > 1 && <span>Annual surplus: <strong style={{ color: C.teal }}>{fmtAnnual(sankeyAnnual.savings)}</strong></span>}
            </div>
            {hourlyWage <= 0 && (
              <p className="text-xs mt-3 text-center" style={{ color: C.coral }}>Set an hourly wage on the Budget tab to enable the Part-time Work slider.</p>
            )}

            <div className="mt-6 pt-4" style={{ borderTop: "1px solid #e5e7eb" }}>
              <p className="text-xs font-bold mb-2" style={{ color: C.navy }}>Notes:</p>
              <ol className="text-xs space-y-2" style={{ color: "#6b7280", listStyle: "decimal", paddingLeft: 18 }}>
                <li>
                  <strong style={{ color: C.navy }}>Tax</strong> — see the{" "}
                  <button onClick={() => setTab("tax")} className="underline font-semibold" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer" }}>Tax</button>{" "}
                  tab for how this is calculated. In practice, students typically pay this progressively through the year as a PAYG deduction from wages, rather than as a single amount.
                </li>
                <li>
                  <strong style={{ color: C.navy }}>University payments</strong> — this annual budget summary does not include any scholarship or one-off payments that you may receive from your university.
                </li>
                <li>
                  <strong style={{ color: C.navy }}>Pre-Arrival and One-off costs</strong> — the summary does not include any pre-arrival costs (e.g. flights) or a range of potential one-off costs. These costs may include costs associated with moving into a rental property once you leave university accommodation. For example, you may be required to pay a refundable rental bond (typically 4 weeks of rent) and 2 weeks of rent in advance. Based on the assumed rental rate this would total approximately <strong style={{ color: C.navy }}>${(rentalBond + rentInAdvance).toLocaleString("en-AU")}</strong>. You may want to fund these items (along with any budget shortfall) via the loan options shown in the{" "}
                  <button onClick={() => setTab("funding")} className="underline font-semibold" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer" }}>Funding Options</button>{" "}
                  tab.
                </li>
              </ol>
            </div>
          </div>
        )}

        {/* ═══ TAB: ANNUAL CASH FLOW ═══ */}
        {tab === "cashflow" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <GreyNote>
              This tab presents a multi-year forecast model from arrival through graduation and into work. It reuses the same weekly living-cost, Rent Assistance/Youth Allowance, and tax logic as the weekly budget tool. The assumptions below are additional course- and career-specific inputs required for this longer-term analysis.
            </GreyNote>

            <Section title="Tuition/Uni Fees">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-2">
                <div className="flex flex-col gap-0.5">
                  <label className="text-xs font-medium" style={{ color: C.navy }}>Field of Study</label>
                  <select value={fieldOfStudy} onChange={e => setFieldOfStudy(e.target.value)}
                    className="border rounded py-1.5 px-2 text-sm bg-white focus:outline-none focus:ring-2 w-full" style={{ borderColor: "#d1d5db" }}>
                    {[1, 2, 3, 4].map(band => (
                      <optgroup key={band} label={`Band ${band} — $${FIELDS.find(f => f.band === band).csp.toLocaleString("en-AU")}/yr`}>
                        {FIELDS.filter(f => f.band === band).map(f => <option key={f.name} value={f.name}>{f.name}</option>)}
                      </optgroup>
                    ))}
                  </select>
                  <span className="text-xs" style={{ color: "#9ca3af" }}>The student's contribution to tuition fees in Australia depends on field of study, grouped in bands. These fees are deferred via the Government's HECS-HELP loan program.</span>
                </div>
                <Inp label="Course Length (years)" value={studyYears} onChange={v => setStudyYears(Math.max(1, Math.min(6, Math.round(v))))} min={1} max={6} step={1}
                  note={`Results in ${studyYears} annual HECS debt entries`} />
              </div>
              <SubHeading>Annual Student Services Amenities Fee (SSAF)</SubHeading>
              <Inp label="Annual SSAF (deferred via SA-HELP)" value={ssafFee} onChange={setSsafFee} step={1} dollar
                note={<>SSAF can be deferred via SA-HELP and is added to your HELP debt. Max $373/yr. Details at{" "}
                  <a href="https://www.studyassist.gov.au" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>studyassist.gov.au (SA-HELP)</a>.</>} />
            </Section>

            <Section title="University Contribution">
              <Inp label="University Contribution ($)" value={uniContribAmount} onChange={setUniContribAmount} step={500} dollar
                note="For simplicity this amount is assumed to go towards the student's accommodation costs in Year 1. Because this reduces the rent the student is actually paying, it also reduces the Rent Assistance the student is assessed as receiving in Year 1 (see Note 2 below)." />
            </Section>

            <Section title="Growth Rates">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <Inp label="Wage Growth (annual)" value={wageGrowth} onChange={setWageGrowth} step={0.01} note="0.03 = 3%. Applied to post-study salary only" />
                <Inp label="Inflation (annual)" value={cfInflation} onChange={setCfInflation} step={0.01} note="Applied to living costs & tuition fees" />
              </div>
            </Section>

            <Section title="Graduate Salary">
              <Inp label={`Starting Salary — Year ${studyYears + 1}`} value={graduateSalary} onChange={setGraduateSalary} step={1000} dollar note="Grows by Wage Growth each year after this" />
            </Section>

            <Section title="Pre-Arrival / Settlement Costs">
              <Inp label="Pre-Arrival Costs (Year 1)" value={preArrivalCost} onChange={setPreArrivalCost} step={100} dollar note="Footnote: flights, exit fees, and similar one-off costs of arriving in Australia" />
            </Section>

            <SectionDivider title={`Projection — ${cfProjectionYears} Years`} />
            <GreyNote>Click row labels with ▸ to expand detail. Loan drawdowns are ticked directly in the Financing section of the table below.</GreyNote>
            <div className="overflow-x-auto mb-4">
              <table className="w-full text-xs">
                <thead>
                  <tr style={{ backgroundColor: C.navy }}>
                    <th className="text-left px-3 py-2 font-semibold text-white sticky left-0 min-w-44" style={{ backgroundColor: C.navy }}>Year</th>
                    {cashflowRows.map(r => (
                      <th key={r.year} className="text-right px-3 py-2 font-semibold text-white whitespace-nowrap">
                        <div className="font-semibold">Year {r.year}</div>
                        <div className="text-xs font-normal" style={{ opacity: 0.7 }}>{r.isStudy ? "Study" : "Working"}</div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {CF_ROWS.map((row, ri) => {
                    if (row.hdr) return (
                      <tr key={row.key}>
                        <td colSpan={cashflowRows.length + 1} className="px-3 py-2 font-bold text-xs uppercase tracking-wider sticky left-0"
                          style={{ backgroundColor: row.subhdr ? "#f0f4ff" : "#f3f4f6", color: C.navy, borderTop: row.subhdr ? "none" : ri > 0 ? `2px solid ${C.navy}30` : "none" }}>
                          {row.label}
                          {row.linkTab && (
                            <button onClick={() => setTab(row.linkTab)} className="ml-2 underline normal-case font-semibold" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer", textTransform: "none", letterSpacing: 0 }}>
                              (See {TABS.find(([k]) => k === row.linkTab)?.[1]} tab)
                            </button>
                          )}
                        </td>
                      </tr>
                    );
                    if (row.parent && !expandedCF.has(row.parent)) return null;
                    const isTotal = row.total, isSub = row.sub, isCum = row.cum, isDebt = row.debt;
                    const isNetHighlight = row.key === "net" || row.key === "net_op";
                    const rowBg = isNetHighlight ? `${C.teal}15` : isTotal ? "#f0fdf4" : "white";
                    return (
                      <tr key={row.key} style={{ backgroundColor: rowBg, borderBottom: "1px solid #f3f4f6" }}>
                        <td className="px-3 py-1.5 sticky left-0" style={{
                          backgroundColor: rowBg, paddingLeft: `${12 + (row.indent || 0) * 16}px`,
                          cursor: row.expandable ? "pointer" : "default",
                          color: row.muted ? "#9ca3af" : row.note ? "#065f46" : C.navy,
                          fontWeight: isSub || isTotal || isCum || isNetHighlight ? 700 : 400,
                          fontStyle: row.note || row.muted ? "italic" : "normal",
                        }} onClick={() => row.expandable && toggleCF(row.key)}>
                          {row.expandable && <span className="mr-1">{expandedCF.has(row.key) ? "▾" : "▸"}</span>}
                          {row.label}{row.sup && <sup style={{ color: C.navy, fontWeight: 700 }}>{row.sup}</sup>}
                        </td>
                        {cashflowRows.map((r, i) => {
                          const raw = r[row.f];
                          const val = row.note && raw > 0 ? -raw : raw; // "Less: LITO" displays as a reduction
                          if (row.checkbox && r.isStudy && i < 3) {
                            const isSsl = row.checkbox === "ssl";
                            return (
                              <td key={r.year} className="px-3 py-1.5 text-right font-mono whitespace-nowrap" style={{ color: C.navy, backgroundColor: rowBg }}>
                                <span className="inline-flex flex-col items-end gap-0.5">
                                  <label className="inline-flex items-center gap-1 cursor-pointer">
                                    <input type="checkbox" checked={isSsl ? sslYears[i] : sparkYears[i]}
                                      onChange={() => isSsl ? toggleSslYear(i) : toggleSparkYear(i)}
                                      style={{ accentColor: isSsl ? C.teal : C.coral }} />
                                    {fmt(val)}
                                  </label>
                                  {!isSsl && sparkYears[i] && (
                                    <input type="number" value={sparkAmts[i]} onChange={e => updateSparkAmt(i, parseFloat(e.target.value) || 0)}
                                      min={0} max={SPARK_MAX_AMT} step={100} onClick={e => e.stopPropagation()}
                                      className="border rounded py-0.5 px-1 text-xs w-20 text-center" style={{ borderColor: "#d1d5db" }} />
                                  )}
                                </span>
                              </td>
                            );
                          }
                          if (val === 0 && !isSub && !isTotal && !isNetHighlight) {
                            return <td key={r.year} className="px-3 py-1.5 text-center whitespace-nowrap" style={{ color: "#d1d5db", backgroundColor: rowBg }}>-</td>;
                          }
                          const neg = val < 0;
                          return (
                            <td key={r.year} className="px-3 py-1.5 text-right font-mono whitespace-nowrap" style={{
                              color: neg ? (row.note ? "#065f46" : C.coral) : row.muted ? "#9ca3af" : isNetHighlight ? "#065f46" : isTotal ? C.navy : isCum ? "#065f46" : isDebt ? C.coral : C.navy,
                              fontWeight: isSub || isTotal || isCum || isNetHighlight ? 700 : 400,
                              fontStyle: row.muted ? "italic" : "normal",
                              backgroundColor: rowBg,
                            }}>
                              {fmt(val)}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="mt-2 pt-4" style={{ borderTop: "1px solid #e5e7eb" }}>
              <p className="text-xs font-bold mb-2" style={{ color: C.navy }}>Notes:</p>
              <ul className="text-xs space-y-2" style={{ color: "#6b7280", listStyle: "none", paddingLeft: 0 }}>
                <li><sup style={{ color: C.navy, fontWeight: 700 }}>1</sup> <strong style={{ color: C.navy }}>HECS/HELP Debt</strong> combines your CSP student contribution, SSAF, and any SSL drawn into a single balance — this matches how the ATO now administers these loans under one combined threshold. It is indexed annually ({(HECS_INDEXATION_RATE * 100).toFixed(1)}%, the 2026 rate) and repaid using the marginal HECS-HELP system in effect since 1 July 2025 (2026–27 thresholds: nil below $69,528; 15% between $69,528–$129,717; 17% above that up to $186,051; a flat 10% of total repayment income above $186,051). See the <button onClick={() => setTab("tax")} className="underline font-semibold" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer" }}>Tax</button> tab for the income tax calculation used each year.</li>
                <li><sup style={{ color: C.navy, fontWeight: 700 }}>2</sup> <strong style={{ color: C.navy }}>Youth Allowance and Rent Assistance</strong> are indexed to inflation each study year (from Year 2 onward) to keep pace with rising living costs. In Year 1, Rent Assistance is reduced because the University Contribution is assumed to pay toward accommodation, lowering the rent the student is assessed as paying out of pocket.</li>
                <li><sup style={{ color: C.navy, fontWeight: 700 }}>3</sup> <strong style={{ color: C.navy }}>The Refugee Student Loan</strong> carries a 5% admin fee, added to the balance on each drawdown (e.g. $250 added if $5,000 is drawn). Interest accrues at 7% p.a. during the study period and is capitalised once repayment begins; the balance is then repaid via a level monthly payment over a fixed 4-year repayment period (see the <button onClick={() => setTab("loan")} className="underline font-semibold" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer" }}>Loan Summary</button> tab for a detailed breakdown).</li>
              </ul>
            </div>
          </div>
        )}

        {/* ═══ TAB 2: UTILITIES GUIDE ═══ */}
        {tab === "utilities" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Utilities Guide" />
            <GreyNote>
              This guide helps you estimate your weekly utility costs based on your accommodation type and where in Australia you are living. Utility expenses can vary significantly and will depend heavily on climate and whether the accommodation uses gas and electricity. You can use this tool to generate an estimate, then apply it to your budget. These costs will likely not be incurred when living in on-campus accommodation. They are provided as a guide for when you move into your own rental accommodation, which may occur after six months or in your second year.
            </GreyNote>

            <Section title="Your Situation">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-4">
                <div>
                  <label className="text-xs font-medium block mb-2" style={{ color: C.navy }}>Accommodation type</label>
                  <div className="flex flex-col gap-2">
                    {ACCOM_TYPES.map(a => (
                      <button key={a.key} onClick={() => setUtilAccomType(a.key)}
                        className="px-4 py-2.5 rounded text-xs font-medium transition text-left"
                        style={{
                          backgroundColor: utilAccomType === a.key ? C.navy : "#f3f4f6",
                          color: utilAccomType === a.key ? "white" : C.navy,
                          border: `2px solid ${utilAccomType === a.key ? C.navy : "#e5e7eb"}`,
                        }}>
                        {a.label}
                        {a.bundled && <span className="ml-2 opacity-70">(utilities usually included)</span>}
                      </button>
                    ))}
                  </div>
                </div>
                <div>
                  <label className="text-xs font-medium block mb-2" style={{ color: C.navy }}>Climate zone</label>
                  <div className="flex flex-col gap-2">
                    {CLIMATE_ZONES.map(cz => (
                      <button key={cz.key} onClick={() => setUtilClimate(cz.key)}
                        className="px-4 py-2.5 rounded text-xs font-medium transition text-left"
                        style={{
                          backgroundColor: utilClimate === cz.key ? C.navy : "#f3f4f6",
                          color: utilClimate === cz.key ? "white" : C.navy,
                          border: `2px solid ${utilClimate === cz.key ? C.navy : "#e5e7eb"}`,
                        }}>
                        <span className="font-semibold">{cz.label}</span>
                        <span className="ml-2 opacity-70">— {cz.cities}</span>
                      </button>
                    ))}
                  </div>
                  <p className="text-xs mt-2" style={{ color: "#9ca3af" }}>Climate zone primarily affects gas/heating costs. Select the zone closest to your university city.</p>
                </div>
              </div>
            </Section>

            <Section title="Your Estimated Weekly Utility Costs">
              {utilEstimate.bundled ? (
                <div className="p-4 rounded-lg mb-4" style={{ backgroundColor: C.cyan, border: `1px solid ${C.navy}20` }}>
                  <p className="text-sm font-semibold" style={{ color: C.navy }}>$0/week — utilities are typically included in your rent</p>
                  <p className="text-xs mt-1" style={{ color: "#6b7280" }}>
                    {utilAccomType === "student"
                      ? "Purpose-built student accommodation almost always includes electricity, gas, water and internet in the weekly rent. Check your lease to confirm."
                      : "When living with family or in a homestay arrangement, utility costs are typically covered by the household. You may want to discuss contributing to household expenses."}
                  </p>
                </div>
              ) : (
                <>
                  <div className="overflow-x-auto mb-4">
                    <table className="w-full text-sm">
                      <thead>
                        <tr style={{ backgroundColor: C.navy }}>
                          <th className="text-left px-4 py-2 font-semibold text-white">Utility</th>
                          <th className="text-right px-4 py-2 font-semibold text-white w-28">Weekly Est.</th>
                          <th className="text-right px-4 py-2 font-semibold text-white w-28">Annual Est.</th>
                        </tr>
                      </thead>
                      <tbody>
                        {[
                          { label: "⚡ Electricity", val: utilEstimate.electricity },
                          { label: "🔥 Gas", val: utilEstimate.gas },
                          { label: "💧 Water", val: utilEstimate.water },
                          { label: "📶 Internet", val: utilEstimate.internet },
                        ].map((row, i) => (
                          <tr key={i} style={{ borderBottom: "1px solid #f3f4f6" }}>
                            <td className="px-4 py-2" style={{ color: C.navy }}>{row.label}</td>
                            <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>${row.val}/wk</td>
                            <td className="px-4 py-2 text-right font-mono text-xs" style={{ color: "#6b7280" }}>${(row.val * 52).toLocaleString("en-AU")}/yr</td>
                          </tr>
                        ))}
                        <tr style={{ borderTop: `2px solid ${C.teal}`, backgroundColor: `${C.teal}15` }}>
                          <td className="px-4 py-2.5 font-bold" style={{ color: C.navy }}>Total Utilities</td>
                          <td className="px-4 py-2.5 text-right font-mono font-bold" style={{ color: C.navy }}>${utilEstimate.total}/wk</td>
                          <td className="px-4 py-2.5 text-right font-mono font-bold text-xs" style={{ color: "#6b7280" }}>${(utilEstimate.total * 52).toLocaleString("en-AU")}/yr</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  <div className="flex items-center gap-4 p-4 rounded-lg mb-4" style={{ backgroundColor: `${C.teal}10`, border: `1px solid ${C.teal}` }}>
                    <div className="flex-1">
                      <p className="text-xs font-semibold" style={{ color: C.navy }}>Apply this estimate to your budget?</p>
                      <p className="text-xs" style={{ color: "#6b7280" }}>
                        This will set your Utilities field on the Budget tab to <strong>${utilEstimate.total}/week</strong>.
                        {livingCosts.utilities > 0 && (
                          <span> It will replace the current value of <strong>${livingCosts.utilities}/week</strong>.</span>
                        )}
                      </p>
                    </div>
                    <button onClick={applyUtilityEstimate}
                      className="px-5 py-2.5 rounded-lg text-sm font-bold text-white transition whitespace-nowrap"
                      style={{ backgroundColor: C.navy }}>
                      Use ${utilEstimate.total}/wk →
                    </button>
                  </div>
                </>
              )}
            </Section>

            <Section title="Understanding Each Utility">
              <GreyNote>Click on each utility below to learn more about what drives costs and how to save money. These estimates are indicative and based on typical usage patterns for students in 2024–25.</GreyNote>
              <UtilityCard info={UTIL_INFO.electricity} estimate={utilEstimate.electricity} isBundled={utilEstimate.bundled} />
              <UtilityCard info={UTIL_INFO.gas} estimate={utilEstimate.gas} isBundled={utilEstimate.bundled} />
              <UtilityCard info={UTIL_INFO.water} estimate={utilEstimate.water} isBundled={utilEstimate.bundled} />
              <UtilityCard info={UTIL_INFO.internet} estimate={utilEstimate.internet} isBundled={utilEstimate.bundled} />
            </Section>

            <Section title="Quick Reference — Estimated Weekly Utility Costs">
              <GreyNote>This table summarises the estimated weekly per-person utility cost by accommodation type and climate zone. All values are in $/week.</GreyNote>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ backgroundColor: C.navy }}>
                      <th className="text-left px-3 py-2 font-semibold text-white">Accommodation Type</th>
                      <th className="text-center px-3 py-2 font-semibold text-white">Warm</th>
                      <th className="text-center px-3 py-2 font-semibold text-white">Mild</th>
                      <th className="text-center px-3 py-2 font-semibold text-white">Cool</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ACCOM_TYPES.map(a => (
                      <tr key={a.key} style={{
                        borderBottom: "1px solid #f3f4f6",
                        backgroundColor: a.key === utilAccomType ? `${C.teal}15` : "transparent",
                      }}>
                        <td className="px-3 py-2" style={{ color: C.navy, fontWeight: a.key === utilAccomType ? 600 : 400 }}>{a.label}</td>
                        {CLIMATE_ZONES.map(cz => {
                          const est = getUtilityEstimate(a.key, cz.key);
                          const isActive = a.key === utilAccomType && cz.key === utilClimate;
                          return (
                            <td key={cz.key} className="px-3 py-2 text-center font-mono" style={{
                              color: est.bundled ? "#9ca3af" : C.navy,
                              fontWeight: isActive ? 700 : 400,
                              backgroundColor: isActive ? `${C.teal}30` : "transparent",
                            }}>
                              {est.bundled ? "Included" : `$${est.total}`}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Section>

            <div className="p-4 rounded text-xs" style={{ backgroundColor: "#fef3c7", color: "#92400e", border: "1px solid #fcd34d" }}>
              <strong>Note:</strong> These estimates are indicative rules of thumb based on average retail energy pricing and typical student usage patterns. Your actual costs will depend on your specific energy retailer, plan, appliances, usage habits, and the energy efficiency of your dwelling. Prices are based on 2024–25 data and are subject to change.
            </div>
          </div>
        )}

        {/* ═══ TAB 3: FUNDING ═══ */}
        {tab === "funding" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Funding Options" />
            <GreyNote>
              If your budget indicates that you may not have enough money to cover your expected costs you may need to take out a loan. Below are two loan options that may help you fund your expenses.
            </GreyNote>
            <Section title="1. Student Start-up Loan (SSL)">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <p className="mb-2">The Student Start-up Loan (SSL) is a voluntary, tax-free loan from the Australian Government designed to help eligible higher education students with the costs of study, including textbooks, equipment, travel and living expenses.</p>
                <p className="mb-2">You can borrow <strong style={{ color: C.navy }}>$1,349 per semester</strong> (up to twice per calendar year, i.e. up to $2,698 per year) for the duration of your course, provided you continue to meet the eligibility requirements and apply each period.</p>
                <p className="mb-2">To be eligible, you must receive at least $1 of Youth Allowance (as a student), Austudy, or ABSTUDY Living Allowance in the relevant fortnight.</p>
                <p className="mb-2">The SSL is added to your HELP debt and is repaid through the tax system once your income exceeds the compulsory repayment threshold (same as HECS-HELP debts). The loan is interest-free but is subject to annual indexation, which means the total amount you repay will be more than you borrow. Indexation is applied on 1 June each year once the debt is at least 11 months old.</p>
                <p>More info:{" "}<a href="https://www.servicesaustralia.gov.au/student-start-up-loan" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>Services Australia — Student Start-up Loan</a></p>
              </div>
              <div className="p-3 rounded text-xs" style={{ backgroundColor: "#f0f4ff", border: `1px solid ${C.navy}20` }}>
                <div className="grid grid-cols-2 gap-2" style={{ maxWidth: 400 }}>
                  <div style={{ color: C.navy }}>Amount per semester</div><div className="font-mono font-semibold" style={{ color: C.navy }}>$1,349</div>
                  <div style={{ color: C.navy }}>Maximum per year</div><div className="font-mono font-semibold" style={{ color: C.navy }}>$2,698</div>
                  <div style={{ color: C.navy }}>Interest rate</div><div className="font-mono font-semibold" style={{ color: C.navy }}>0% (indexed to CPI)</div>
                  <div style={{ color: C.navy }}>Repayment</div><div className="font-semibold" style={{ color: C.navy }}>Via tax system (HELP debt)</div>
                </div>
              </div>
            </Section>
            <Section title="2. Refugee Student Loan Program">
              <div className="text-xs mb-3" style={{ color: "#6b7280" }}>
                <p className="mb-2">The Refugee Student Loan Program has been developed by Skill Path for students in the RSSP. It is administered by Spark Finance. These loans are for up to <strong style={{ color: C.navy }}>$5,000/year</strong> during study. They have a <strong style={{ color: C.navy }}>7% interest rate</strong> and a <strong style={{ color: C.navy }}>4-year repayment term</strong> beginning after you complete your course, with no repayment required during study. See the <button onClick={() => setTab("loan")} className="underline font-semibold" style={{ color: C.navy, background: "none", border: "none", padding: 0, cursor: "pointer" }}>Loan Summary</button> tab for a detailed repayment profile.</p>
                <p>More information about this loan is on the RSSP Student Hub in Notion{" "}<a href="https://www.notion.so/Student-Loan-Scheme-with-Spark-Finance-2c4199deb36c806b9819f6ad2565706c?source=copy_link" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>here</a>.</p>
              </div>
              <div className="p-3 rounded text-xs" style={{ backgroundColor: "#f0f4ff", border: `1px solid ${C.navy}20` }}>
                <div className="grid grid-cols-2 gap-2" style={{ maxWidth: 400 }}>
                  <div style={{ color: C.navy }}>Maximum per year</div><div className="font-mono font-semibold" style={{ color: C.navy }}>$5,000</div>
                  <div style={{ color: C.navy }}>Interest rate</div><div className="font-mono font-semibold" style={{ color: C.navy }}>7%</div>
                  <div style={{ color: C.navy }}>Repayment term</div><div className="font-mono font-semibold" style={{ color: C.navy }}>4 years (fixed)</div>
                  <div style={{ color: C.navy }}>Grace period</div><div className="font-mono font-semibold" style={{ color: C.navy }}>Length of study</div>
                  <div style={{ color: C.navy }}>Repayment during study</div><div className="font-semibold" style={{ color: C.navy }}>None required</div>
                  <div style={{ color: C.navy }}>Administered by</div><div className="font-semibold" style={{ color: C.navy }}>Spark Finance</div>
                </div>
              </div>
            </Section>
          </div>
        )}

        {/* ═══ TAB: LOAN SUMMARY ═══ */}
        {tab === "loan" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="Refugee Student Loan" />
            <GreyNote>
              The summary below shows the repayment profile for the Refugee Student Loan Program for various drawdown assumptions. This tab is standalone from the other tabs so changes here do not affect the Weekly Budget or Long Term Cash Flow tabs.
            </GreyNote>

            <Section title="Drawdown Assumptions">
              <div className="overflow-x-auto mb-2">
                <table className="text-xs">
                  <thead>
                    <tr style={{ backgroundColor: C.cyan }}>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}></th>
                      {[0, 1, 2].map(i => <th key={i} className="text-center p-2 font-semibold" style={{ color: C.navy }}>Year {i + 1}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="p-2" style={{ color: C.navy }}>Drawdown</td>
                      {[0, 1, 2].map(i => (
                        <td key={i} className="p-2 text-center">
                          <div className="flex flex-col items-center gap-1">
                            <input type="checkbox" checked={loanYears[i]} onChange={() => toggleLoanYear(i)} style={{ accentColor: C.coral }} />
                            {loanYears[i] && (
                              <div className="relative" style={{ width: 90 }}>
                                <span className="absolute left-2 top-1/2 -translate-y-1/2 text-xs text-gray-400">$</span>
                                <input type="number" value={loanAmts[i]} onChange={e => updateLoanAmt(i, parseFloat(e.target.value) || 0)}
                                  min={0} max={SPARK_MAX_AMT} step={100}
                                  className="border rounded py-1 pl-5 pr-1 text-xs w-full text-center" style={{ borderColor: "#d1d5db" }} />
                              </div>
                            )}
                          </div>
                        </td>
                      ))}
                    </tr>
                  </tbody>
                </table>
              </div>
              <p className="text-xs" style={{ color: "#9ca3af" }}>Maximum ${SPARK_MAX_AMT.toLocaleString("en-AU")}/year. Default: 3 annual drawdowns of $5,000.</p>
            </Section>

            <SectionDivider title="Repayment Profile" />
            <div className="flex flex-row gap-4 items-start">
              <div className="p-3 rounded-lg flex-shrink-0" style={{ backgroundColor: "#eaf7f7", width: 200 }}>
                <div className="space-y-2 text-xs">
                  <div>
                    <div className="font-bold" style={{ color: C.navy }}>Amount</div>
                    <div style={{ color: C.navy }}>${loanSchedule.totalDrawn.toLocaleString("en-AU")}</div>
                    <div style={{ color: "#6b7280" }}>({loanYears.filter(Boolean).length} annual drawdown{loanYears.filter(Boolean).length === 1 ? "" : "s"}{loanDrawdowns.every(d => d === 0 || d === loanDrawdowns.find(x => x > 0)) ? ` of $${(loanDrawdowns.find(x => x > 0) || 0).toLocaleString("en-AU")}` : ""})</div>
                  </div>
                  <hr style={{ borderColor: "#d1e7e7" }} />
                  <div>
                    <div className="font-bold" style={{ color: C.navy }}>Grace Period</div>
                    <div style={{ color: C.navy }}>3 years</div>
                  </div>
                  <hr style={{ borderColor: "#d1e7e7" }} />
                  <div>
                    <div className="font-bold" style={{ color: C.navy }}>Fee</div>
                    <div style={{ color: C.navy }}>${Math.round(loanSchedule.totalFees).toLocaleString("en-AU")}</div>
                    <div style={{ color: "#6b7280" }}>(5.0% for each drawdown)</div>
                  </div>
                  <hr style={{ borderColor: "#d1e7e7" }} />
                  <div>
                    <div className="font-bold" style={{ color: C.navy }}>Repayments</div>
                    <div style={{ color: C.navy }}>${Math.round(loanSchedule.monthlyPMT).toLocaleString("en-AU")}/mo (minimum)</div>
                  </div>
                  <hr style={{ borderColor: "#d1e7e7" }} />
                  <div>
                    <div className="font-bold" style={{ color: C.navy }}>Total Interest</div>
                    <div style={{ color: C.navy }}>${Math.round(loanSchedule.totalInterest).toLocaleString("en-AU")}</div>
                  </div>
                  <hr style={{ borderColor: "#d1e7e7" }} />
                  <div>
                    <div className="font-bold" style={{ color: C.navy }}>Total Amount Repaid</div>
                    <div style={{ color: C.navy }}>${Math.round(loanSchedule.totalRepaid).toLocaleString("en-AU")}</div>
                  </div>
                </div>
              </div>
              <div style={{ flex: "1 1 0%", minWidth: 0 }}>
                <LoanWaterfallChart rows={loanSchedule.rows} graceYears={3} />
              </div>
            </div>

            <div className="mt-4 pt-4" style={{ borderTop: "1px solid #e5e7eb" }}>
              <p className="text-xs" style={{ color: "#9ca3af" }}>
                Interest accrues as simple interest during the grace period and is capitalised once at the start of repayment; the balance is then repaid via a level monthly payment (standard loan amortisation) over the 4-year repayment period. Figures are estimates only — actual contracts may differ slightly due to exact drawdown dates and day-count conventions.
              </p>
            </div>
          </div>
        )}

        {/* ═══ TAB: TAX ═══ */}
        {tab === "tax" && (
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#e5e7eb" }}>
            <SectionDivider title="A primer on Tax, the Tax-Free Threshold & Youth Allowance" />
            <GreyNote>
              <em>Disclaimer: This is a generic explanation of the tax system and is not tax advice. Consider your own circumstances and seek advice if you need help with tax.</em>
            </GreyNote>

            <Section title="Introduction">
              <p className="text-xs mb-2" style={{ color: "#6b7280" }}>
                In Australia, you need to disclose to the government — via the Australian Taxation Office (ATO) — the income you earn each year. This lets the government work out how much tax you're required to pay (or refund you, if too much has already been paid during the year).
              </p>
            </Section>

            <Section title="Taxable income in Australia">
              <p className="text-xs mb-2" style={{ color: "#6b7280" }}>"Income" is broader than most students expect. It includes:</p>
              <ul className="text-xs mb-3 space-y-1" style={{ color: "#6b7280", listStyle: "disc", paddingLeft: 18 }}>
                <li><strong style={{ color: C.navy }}>Wages and salary</strong> from a job</li>
                <li><strong style={{ color: C.navy }}>Interest</strong> earned on money in a bank account</li>
                <li><strong style={{ color: C.navy }}>Investment earnings</strong> (dividends, capital gains, etc.)</li>
                <li><strong style={{ color: C.navy }}>Government payments</strong>, including Youth Allowance and JobSeeker (but <span style={{ textDecoration: "underline" }}>not</span> Rent Assistance, as this is not considered taxable income)</li>
                <li><strong style={{ color: C.navy }}>Scholarships</strong> from a university, in some circumstances — see the ATO's{" "}
                  <a href="https://www.ato.gov.au/calculators-and-tools/is-my-scholarship-taxable" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>Is my scholarship taxable?</a> tool.
                  Note: initial payments students receive from universities as part of the RSSP are generally <span style={{ textDecoration: "underline" }}>not</span> taxable income.
                </li>
              </ul>
              <p className="text-xs mb-3" style={{ color: "#6b7280" }}>All of it is added together into your <strong style={{ color: C.navy }}>total assessable income</strong> — that total is what your tax is calculated on, not any single source in isolation.</p>
              <div className="p-3 rounded-lg text-xs" style={{ backgroundColor: "#fef3c7", border: "1px solid #fcd34d", color: "#92400e" }}>
                <strong>Important:</strong> the financial year in Australia is 1 July to 30 June (12 months).
              </div>
            </Section>

            <Section title="Tax rates and the “tax-free threshold”">
              <p className="text-xs mb-3" style={{ color: "#6b7280" }}>
                Australia uses a <strong style={{ color: C.navy }}>progressive</strong> tax system: the first slice of income is tax-free, and each additional slice above that is taxed at increasing rates. The tax-free threshold is <strong style={{ color: C.navy }}>$18,200</strong> — the first $18,200 of total assessable income isn't taxed. On top of income tax, most residents also pay a 2% Medicare levy (with a reduction for lower incomes).
              </p>
              <div className="overflow-x-auto mb-3">
                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ backgroundColor: C.cyan }}>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>Taxable income (2026–27)</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>Tax on this income</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}><td className="p-2">$0 – $18,200</td><td className="p-2 font-mono">Nil</td></tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}><td className="p-2">$18,201 – $45,000</td><td className="p-2 font-mono">15c for each $1 over $18,200</td></tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}><td className="p-2">$45,001 – $135,000</td><td className="p-2 font-mono">$4,020 plus 30c for each $1 over $45,000</td></tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}><td className="p-2">$135,001 – $190,000</td><td className="p-2 font-mono">$31,020 plus 37c for each $1 over $135,000</td></tr>
                    <tr><td className="p-2">Over $190,000</td><td className="p-2 font-mono">$51,370 plus 45c for each $1 over $190,000</td></tr>
                  </tbody>
                </table>
              </div>
              <div className="overflow-x-auto mb-3">
                <table className="w-full text-xs">
                  <thead>
                    <tr style={{ backgroundColor: C.cyan }}>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>Medicare levy (2026–27, single, no dependents)</th>
                      <th className="text-left p-2 font-semibold" style={{ color: C.navy }}>Levy</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}><td className="p-2">Up to $28,011</td><td className="p-2 font-mono">Nil</td></tr>
                    <tr style={{ borderBottom: "1px solid #f3f4f6" }}><td className="p-2">$28,012 – $35,014</td><td className="p-2 font-mono">10c for each $1 over $28,011 (phase-in)</td></tr>
                    <tr><td className="p-2">Over $35,014</td><td className="p-2 font-mono">2% of taxable income</td></tr>
                  </tbody>
                </table>
              </div>
              <p className="text-xs mb-1" style={{ color: "#6b7280" }}>Also relevant: the Low Income Tax Offset (LITO) — up to $700 for incomes under $37,500, reducing on a sliding scale to nil at $66,667 — is applied against income tax (not the Medicare levy).</p>
              <div className="p-3 rounded-lg text-xs mb-3" style={{ backgroundColor: "#fef3c7", border: "1px solid #fcd34d", color: "#92400e" }}>
                <strong>Important:</strong> even if your income is below $18,200, you generally still need to lodge a tax return (or a non-lodgment advice) — being under the threshold doesn't excuse reporting.
              </div>
              <div className="p-3 rounded-lg text-xs" style={{ backgroundColor: "#f0f4ff", border: `1px solid ${C.navy}20`, color: "#6b7280" }}>
                This tool covers the basics needed to understand your own estimated bill below. For more detail on <em>electing where you claim the tax-free threshold</em>, how <em>Youth Allowance and tax</em> interact in practice, and a full worked example comparing PAYG strategies, see the{" "}
                <a href="https://app.notion.com/p/3af199deb36c8169b30bca2cf846dd2e" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline", fontWeight: 700 }}>RSSP Student Hub — Tax page</a> on Notion.
              </div>
            </Section>

            {/* ═══ ESTIMATED TAX (from live model assumptions) ═══ */}
            <SectionDivider title="Your Estimated Tax" />
            <GreyNote>
              Calculated from your current assumptions on the Weekly Budget tab: Wages and Youth Allowance are the two taxable income items in this model. Rent Assistance is excluded (not taxable). This uses the 2026–27 resident tax scale, Medicare levy, and LITO shown above.
            </GreyNote>
            <div className="overflow-x-auto mb-4">
              <table className="w-full text-sm">
                <tbody>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy }}>Wages (annual)</td>
                    <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(annualWagesForTax)}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy }}>Youth Allowance received (annual)</td>
                    <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(annualYAForTax)}</td>
                  </tr>
                  <tr style={{ borderTop: "2px solid #e5e7eb" }}>
                    <td className="px-4 py-2 font-bold" style={{ color: C.navy }}>Total assessable income</td>
                    <td className="px-4 py-2 text-right font-mono font-bold" style={{ color: C.navy }}>{fmt(taxResult.taxableIncome)}</td>
                  </tr>
                  <tr><td colSpan={2} className="py-1"></td></tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Income tax on assessable income</td>
                    <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(taxResult.gross)}</td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: "#065f46", paddingLeft: 28 }}>Less: Low Income Tax Offset (LITO)</td>
                    <td className="px-4 py-2 text-right font-mono" style={{ color: taxResult.lito > 0 ? "#065f46" : "#9ca3af" }}>
                      {taxResult.lito > 0 ? `($${Math.round(taxResult.lito).toLocaleString("en-AU")})` : fmt(0)}
                    </td>
                  </tr>
                  <tr style={{ borderBottom: "1px solid #f3f4f6" }}>
                    <td className="px-4 py-2" style={{ color: C.navy, paddingLeft: 28 }}>Plus: Medicare levy</td>
                    <td className="px-4 py-2 text-right font-mono" style={{ color: C.navy }}>{fmt(taxResult.medicare)}</td>
                  </tr>
                  <tr style={{ borderTop: `3px solid ${C.coral}`, backgroundColor: `${C.coral}15` }}>
                    <td className="px-4 py-3 font-bold" style={{ color: C.coral }}>Total estimated tax payable for the year</td>
                    <td className="px-4 py-3 text-right font-mono font-bold text-base" style={{ color: C.coral }}>{fmt(taxResult.total)}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <Section title="Assumptions used in this estimate">
              <ul className="text-xs space-y-1" style={{ color: "#6b7280", listStyle: "disc", paddingLeft: 18 }}>
                <li>Only <strong style={{ color: C.navy }}>Wages</strong> and <strong style={{ color: C.navy }}>Youth Allowance actually received</strong> (after any income-test reduction) are treated as taxable income. Rent Assistance is excluded.</li>
                <li>No other income is assumed (bank interest, scholarships, etc.).</li>
                <li>The <strong style={{ color: C.navy }}>Beneficiary Tax Offset</strong> — a separate, more generous offset available specifically to Centrelink allowance recipients — is <span style={{ textDecoration: "underline" }}>not</span> applied here. Including it may reduce the tax payable estimate above. See the following site for more information on this tax offset:{" "}
                  <a href="https://www.ato.gov.au/individuals-and-families/income-deductions-offsets-and-records/tax-offsets/beneficiary-tax-offset" target="_blank" rel="noopener noreferrer" style={{ color: C.navy, textDecoration: "underline" }}>ATO — Beneficiary Tax Offset</a>.
                </li>
                <li>No work-related deductions are assumed.</li>
                <li>This is the <strong style={{ color: C.navy }}>annual liability</strong>, not what's withheld week to week — see the note under the Annual Summary chart on how this is typically paid via PAYG.</li>
              </ul>
            </Section>
          </div>
        )}
      </div>
    </div>
  );
}
