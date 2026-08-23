import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import _ from "lodash";
import { UploadCloud, FileSpreadsheet, Printer, RotateCcw, ChevronRight, AlertTriangle, Plus, Trash2 } from "lucide-react";

// ---------- Default sender (seller) details — editable in the UI ----------
const DEFAULT_SENDER = {
  name: "Nexxbase Marketing Pvt ltd",
  addressLine: "KHASRA No. 146/25/2/1, Jail Road, Gurgaon, Haryana-122101",
  gstin: "06AADCN0946N1Z8",
};

// ---------- Columns we need, matched by header text ----------
const FIELDS = {
  date: "Date",
  orderNo: "Sale Order Number",
  invoiceNo: "Invoice number",
  portal: "Channel entry",
  productName: "Product Name",
  sku: "Product SKU Code",
  qty: "Qty",
  rate: "Unit Price",
  sales: "Sales",
  tax: "Tax",
  taxLedger: "Tax Ledger",
  total: "Total",
  customerName: "Customer Name",
  shipName: "Shipping Address Name",
  addr1: "Shipping Address Line 1",
  addr2: "Shipping Address Line 2",
  city: "Shipping Address City",
  state: "Shipping Address State",
  stateFull: "State Name",
  country: "Shipping Address Country",
  pincode: "Shipping Address Pincode",
  shippingProvider: "Shipping Provider",
  awb: "AWB num",
  hsn: "HSN Code",
  gstin: "GSTIN Number",
  ewayNo: "E-Way Bill Number",
  ewayDate: "E-Way Bill Date",
  irn: "E-Invoice IRN",
  ackNo: "E-Invoice Ack No",
  ackDate: "E-Invoice Ack Date",
};

function cleanStr(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/^[`']+/g, "").trim();
}

function toNumber(v) {
  const n = typeof v === "number" ? v : parseFloat(cleanStr(v));
  return isFinite(n) ? n : 0;
}

function fmtDate(v) {
  if (!v) return "";
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return cleanStr(v);
  // Excel date-only cells are parsed as UTC midnight by SheetJS. Reading them back
  // with local-time getters can roll the date backward/forward depending on the
  // user's timezone, so we always read the calendar date in UTC.
  return d.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric", timeZone: "UTC" });
}

// Normalize country names to the short code the reference invoices use.
function fmtCountry(country) {
  const c = cleanStr(country);
  if (/^india$/i.test(c)) return "IN";
  return c;
}

const MONTH_ABBR = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jun2: 5, jul: 6, jly: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11,
};

// Some order numbers embed the order date right in the string, e.g.
// "610089862310Jly7041" contains "10Jly" -> 10 July. We don't get a
// separate Order Date column from the sheet, so we recover it here,
// falling back to the invoice date if no such pattern is found.
function extractOrderDate(orderNo, fallbackDate) {
  const cleaned = cleanStr(orderNo);
  const match = cleaned.match(/(\d{1,2})([A-Za-z]{3,4})/);
  if (match) {
    const day = parseInt(match[1], 10);
    const monthKey = match[2].toLowerCase();
    const month = MONTH_ABBR[monthKey];
    if (month !== undefined && day >= 1 && day <= 31) {
      const year = fallbackDate instanceof Date && !isNaN(fallbackDate.getTime())
        ? fallbackDate.getUTCFullYear()
        : new Date().getFullYear();
      const d = new Date(Date.UTC(year, month, day));
      return fmtDate(d);
    }
  }
  return fmtDate(fallbackDate);
}

function fmtMoney(n) {
  return toNumber(n).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ---------- Number to words (Indian numbering system) ----------
const ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
  "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
const TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];

function twoDigits(n) {
  if (n < 20) return ONES[n];
  const t = Math.floor(n / 10), o = n % 10;
  return TENS[t] + (o ? " " + ONES[o] : "");
}
function threeDigits(n) {
  const h = Math.floor(n / 100), r = n % 100;
  return (h ? ONES[h] + " Hundred" + (r ? " " : "") : "") + (r ? twoDigits(r) : "");
}
function numberToWords(num) {
  num = Math.round(num);
  if (num === 0) return "Zero";
  const crore = Math.floor(num / 10000000); num %= 10000000;
  const lakh = Math.floor(num / 100000); num %= 100000;
  const thousand = Math.floor(num / 1000); num %= 1000;
  const hundred = num;
  let parts = [];
  if (crore) parts.push(threeDigits(crore) + " Crore");
  if (lakh) parts.push(threeDigits(lakh) + " Lakh");
  if (thousand) parts.push(threeDigits(thousand) + " Thousand");
  if (hundred) parts.push(threeDigits(hundred));
  return parts.join(" ").trim();
}

// ---------- Tax type detection ----------
function parseTaxType(ledgerText) {
  const t = cleanStr(ledgerText).toUpperCase();
  const rateMatch = t.match(/(\d+(\.\d+)?)/);
  const rate = rateMatch ? parseFloat(rateMatch[1]) : 0;
  if (t.includes("IGST")) return { type: "IGST", rate };
  if (t.includes("CGST") || t.includes("SGST")) return { type: "CGST_SGST", rate };
  return { type: "IGST", rate };
}

// ---------- Excel parsing ----------
function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (!rows.length) throw new Error("The sheet appears to be empty.");
  const header = rows[0].map((h) => (h === null || h === undefined ? "" : String(h).trim()));

  const idx = {};
  Object.entries(FIELDS).forEach(([key, label]) => {
    idx[key] = header.indexOf(label);
  });
  if (idx.invoiceNo === -1) {
    throw new Error('Could not find an "Invoice number" column in this file. Please check the format.');
  }

  const dataRows = rows.slice(1).filter((r) => r && r.length && cleanStr(r[idx.invoiceNo]));

  return dataRows.map((r) => {
    const get = (key) => (idx[key] >= 0 ? r[idx[key]] : "");
    const taxInfo = parseTaxType(get("taxLedger"));
    return {
      date: get("date"),
      orderNo: cleanStr(get("orderNo")),
      invoiceNo: cleanStr(get("invoiceNo")),
      portal: cleanStr(get("portal")),
      productName: cleanStr(get("productName")),
      sku: cleanStr(get("sku")),
      qty: toNumber(get("qty")),
      rate: toNumber(get("rate")),
      sales: toNumber(get("sales")),
      taxAmt: toNumber(get("tax")),
      taxType: taxInfo.type,
      taxRate: taxInfo.rate,
      total: toNumber(get("total")),
      customerName: cleanStr(get("customerName")) || cleanStr(get("shipName")),
      addr1: cleanStr(get("addr1")),
      addr2: cleanStr(get("addr2")),
      city: cleanStr(get("city")),
      state: cleanStr(get("stateFull")) || cleanStr(get("state")),
      country: fmtCountry(get("country")),
      pincode: cleanStr(get("pincode")),
      shippingProvider: cleanStr(get("shippingProvider")),
      awb: cleanStr(get("awb")),
      hsn: cleanStr(get("hsn")),
      buyerGstin: cleanStr(get("gstin")),
      ewayNo: cleanStr(get("ewayNo")),
      ewayDate: get("ewayDate"),
      irn: cleanStr(get("irn")),
      ackNo: cleanStr(get("ackNo")),
      ackDate: get("ackDate"),
    };
  });
}

// ---------- Group rows into editable invoice objects ----------
function buildInvoices(rows) {
  const map = new Map();
  rows.forEach((r) => {
    if (!map.has(r.invoiceNo)) map.set(r.invoiceNo, []);
    map.get(r.invoiceNo).push(r);
  });

  const invoices = [];
  let idx = 0;
  map.forEach((items, invoiceNo) => {
    const first = items[0];
    const poGuess = first.orderNo.replace(/\D/g, "").slice(0, 10);
    const lineItems = items.map((r, i) => ({
      key: `${invoiceNo}-${i}`,
      description: [r.sku, r.productName].filter(Boolean).join(" - "),
      hsn: r.hsn,
      qty: String(r.qty),
      // The reference invoice shows the GST-inclusive per-unit rate (Total / Qty),
      // not the sheet's GST-exclusive "Unit Price" column.
      rate: String(r.qty ? (r.total / r.qty).toFixed(2) : r.rate.toFixed(2)),
      less: "0",
      taxable: String(r.sales.toFixed(2)),
      taxType: r.taxType,
      taxRate: String(r.taxRate),
      taxAmt: String(r.taxAmt.toFixed(2)),
      total: String(r.total.toFixed(2)),
    }));

    invoices.push({
      id: `inv-${idx++}`,
      invoiceNo,
      invoiceDate: fmtDate(first.date),
      orderNo: first.orderNo,
      orderDate: extractOrderDate(first.orderNo, first.date instanceof Date ? first.date : new Date(first.date)),
      poNo: poGuess,
      portal: first.portal,
      paymentMode: "",
      buyer: {
        name: first.customerName,
        addr1: first.addr1,
        addr2: first.addr2,
        city: first.city,
        state: first.state,
        country: first.country,
        pincode: first.pincode,
        gstin: first.buyerGstin,
      },
      awb: first.awb,
      dispatchThrough: first.shippingProvider,
      ewayNo: first.ewayNo,
      ewayDate: fmtDate(first.ewayDate),
      irn: first.irn,
      ackNo: first.ackNo,
      ackDate: fmtDate(first.ackDate),
      declaration: "We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct.",
      signatoryLabel: "Authorised Signatory",
      lineItems,
    });
  });

  invoices.sort((a, b) => a.invoiceNo.localeCompare(b.invoiceNo));
  return invoices;
}

// ---------- Derive totals / HSN summary from (editable) line items ----------
function deriveInvoice(inv) {
  const items = inv.lineItems.map((li) => ({
    ...li,
    qtyN: toNumber(li.qty),
    rateN: toNumber(li.rate),
    lessN: toNumber(li.less),
    taxableN: toNumber(li.taxable),
    taxRateN: toNumber(li.taxRate),
    taxAmtN: toNumber(li.taxAmt),
    totalN: toNumber(li.total),
  }));

  const totals = items.reduce(
    (acc, li) => {
      acc.qty += li.qtyN;
      acc.taxable += li.taxableN;
      acc.total += li.totalN;
      if (li.taxType === "IGST") acc.igst += li.taxAmtN;
      else {
        acc.cgst += li.taxAmtN / 2;
        acc.sgst += li.taxAmtN / 2;
      }
      acc.tax += li.taxAmtN;
      return acc;
    },
    { qty: 0, taxable: 0, igst: 0, cgst: 0, sgst: 0, tax: 0, total: 0 }
  );

  const hasIGST = items.some((li) => li.taxType === "IGST");
  const hasCGST = items.some((li) => li.taxType === "CGST_SGST");

  const hsnMap = new Map();
  items.forEach((li) => {
    const key = li.hsn + "|" + li.taxRateN + "|" + li.taxType;
    if (!hsnMap.has(key)) {
      hsnMap.set(key, { hsn: li.hsn, taxRate: li.taxRateN, taxType: li.taxType, qty: 0, taxable: 0, taxAmt: 0, total: 0 });
    }
    const h = hsnMap.get(key);
    h.qty += li.qtyN;
    h.taxable += li.taxableN;
    h.taxAmt += li.taxAmtN;
    h.total += li.totalN;
  });

  return { items, totals, hasIGST, hasCGST, hsnSummary: Array.from(hsnMap.values()) };
}

// ---------- Small editable primitives ----------
function EditInput({ value, onChange, className = "", align, placeholder, width }) {
  return (
    <input
      className={`ed-input ${className}`}
      style={{ textAlign: align, width: width }}
      value={value}
      placeholder={placeholder}
      onChange={(e) => onChange(e.target.value)}
    />
  );
}

function EditBlock({ value, onChange, className = "" }) {
  return (
    <div
      className={`ed-block ${className}`}
      contentEditable
      suppressContentEditableWarning
      onBlur={(e) => onChange(e.currentTarget.innerText)}
    >
      {value}
    </div>
  );
}

// ---------- Invoice document (editable, matches the reference tax invoice layout) ----------
function InvoiceDoc({ inv, sender, onField, onBuyerField, onLineItem, onAddLine, onRemoveLine, onSenderField }) {
  const d = useMemo(() => deriveInvoice(inv), [inv]);
  const showCGST = d.hasCGST;
  const showIGST = d.hasIGST;

  return (
    <div className="invoice-doc">
      <div className="inv-title">Tax Invoice/Bill of Supply/Cash Memo</div>

      <table className="inv-header-table">
        <tbody>
          <tr>
            <td className="hcell" style={{ width: "38%" }}>
              <div className="label">Sender:</div>
              <EditBlock className="bold" value={sender.name} onChange={(v) => onSenderField("name", v)} />
              <EditBlock value={sender.addressLine} onChange={(v) => onSenderField("addressLine", v)} />
              <div className="bold">GSTIN: <EditInput value={sender.gstin} onChange={(v) => onSenderField("gstin", v)} width="120px" /></div>
            </td>
            <td className="hcell" style={{ width: "32%" }}>
              <div className="label">Invoice No:</div>
              <EditInput className="bold" value={inv.invoiceNo} onChange={(v) => onField("invoiceNo", v)} width="95%" />
            </td>
            <td className="hcell" style={{ width: "30%" }}>
              <div className="label">Invoice Date:</div>
              <EditInput className="bold" value={inv.invoiceDate} onChange={(v) => onField("invoiceDate", v)} width="95%" />
            </td>
          </tr>
          <tr>
            <td className="hcell" rowSpan={2}>
              <div className="label">Buyer:</div>
              <EditBlock className="bold" value={inv.buyer.name} onChange={(v) => onBuyerField("name", v)} />
              <EditBlock
                value={[inv.buyer.addr1, inv.buyer.addr2].filter(Boolean).join(", ") +
                  (inv.buyer.city ? `, ${inv.buyer.city}` : "") +
                  (inv.buyer.state ? `, ${inv.buyer.state}` : "") +
                  (inv.buyer.pincode ? `-${inv.buyer.pincode}` : "") +
                  (inv.buyer.country ? ` ${inv.buyer.country}` : "")}
                onChange={(v) => onBuyerField("addr1", v)}
              />
              <div className="bold">GSTIN: <EditInput value={inv.buyer.gstin} onChange={(v) => onBuyerField("gstin", v)} width="140px" /></div>
            </td>
            <td className="hcell" colSpan={2}>
              <div className="label">Order No:</div>
              <EditInput className="bold" value={inv.orderNo} onChange={(v) => onField("orderNo", v)} width="95%" />
              <div className="label" style={{ marginTop: 4 }}>PO: <EditInput value={inv.poNo} onChange={(v) => onField("poNo", v)} width="140px" /></div>
            </td>
          </tr>
          <tr>
            <td className="hcell">
              <div className="label">Order Date:</div>
              <EditInput className="bold" value={inv.orderDate} onChange={(v) => onField("orderDate", v)} width="95%" />
            </td>
            <td className="hcell">
              <div className="label">Portal:</div>
              <EditInput className="bold" value={inv.portal} onChange={(v) => onField("portal", v)} width="95%" />
            </td>
          </tr>
          <tr>
            <td className="hcell">
              <div className="label">Payment Mode:</div>
              <EditInput className="bold" value={inv.paymentMode} onChange={(v) => onField("paymentMode", v)} width="95%" placeholder="—" />
            </td>
            <td className="hcell">
              <div className="label">AWB No:</div>
              <EditInput className="bold" value={inv.awb} onChange={(v) => onField("awb", v)} width="95%" />
            </td>
            <td className="hcell">
              <div className="label">Dispatch Through:</div>
              <EditInput className="bold" value={inv.dispatchThrough} onChange={(v) => onField("dispatchThrough", v)} width="95%" />
              <div className="label" style={{ marginTop: 4 }}>eWay Bill No: <EditInput className="bold" value={inv.ewayNo} onChange={(v) => onField("ewayNo", v)} width="120px" /></div>
              <div className="label">eWay Bill Date: <EditInput className="bold" value={inv.ewayDate} onChange={(v) => onField("ewayDate", v)} width="120px" /></div>
            </td>
          </tr>
        </tbody>
      </table>

      <table className="inv-items-table">
        <thead>
          <tr>
            <th style={{ width: "4%" }}>Sl No.</th>
            <th style={{ width: "26%" }}>Description of Goods</th>
            <th>Qty</th>
            <th>Rate (Rs)</th>
            <th>Less (Rs)</th>
            <th>Taxable Amt (Rs)</th>
            {showIGST && <th>IGST Amt (Rs)</th>}
            {showCGST && <th>CGST Amt (Rs)</th>}
            {showCGST && <th>SGST Amt (Rs)</th>}
            <th>Total (Rs)</th>
            <th className="no-print" style={{ width: "24px" }}></th>
          </tr>
        </thead>
        <tbody>
          {d.items.map((li, i) => (
            <tr key={li.key}>
              <td className="center">{i + 1}</td>
              <td>
                <EditBlock value={li.description} onChange={(v) => onLineItem(i, "description", v)} />
                <div className="hsn-line">HSN : <EditInput value={li.hsn} onChange={(v) => onLineItem(i, "hsn", v)} width="80px" /></div>
              </td>
              <td className="center"><EditInput align="center" value={li.qty} onChange={(v) => onLineItem(i, "qty", v)} width="42px" /></td>
              <td className="right"><EditInput align="right" value={li.rate} onChange={(v) => onLineItem(i, "rate", v)} width="60px" /></td>
              <td className="right"><EditInput align="right" value={li.less} onChange={(v) => onLineItem(i, "less", v)} width="50px" /></td>
              <td className="right"><EditInput align="right" value={li.taxable} onChange={(v) => onLineItem(i, "taxable", v)} width="70px" /></td>
              {showIGST && (
                <td className="right">
                  {li.taxType === "IGST" ? (
                    <>
                      <EditInput align="right" value={li.taxAmt} onChange={(v) => onLineItem(i, "taxAmt", v)} width="70px" />
                      <div className="rate-line">@ <EditInput align="right" value={li.taxRate} onChange={(v) => onLineItem(i, "taxRate", v)} width="36px" />%</div>
                    </>
                  ) : ("—")}
                </td>
              )}
              {showCGST && (
                <td className="right">{li.taxType === "CGST_SGST" ? fmtMoney(li.taxAmtN / 2) : "—"}</td>
              )}
              {showCGST && (
                <td className="right">{li.taxType === "CGST_SGST" ? fmtMoney(li.taxAmtN / 2) : "—"}</td>
              )}
              <td className="right"><EditInput align="right" value={li.total} onChange={(v) => onLineItem(i, "total", v)} width="70px" /></td>
              <td className="no-print center">
                <button className="icon-btn" title="Remove item" onClick={() => onRemoveLine(i)}><Trash2 size={13} /></button>
              </td>
            </tr>
          ))}
          <tr className="totals-row">
            <td colSpan={2} className="bold">Invoice Total</td>
            <td className="center bold">{d.totals.qty}</td>
            <td></td>
            <td></td>
            <td className="right bold">{fmtMoney(d.totals.taxable)}</td>
            {showIGST && <td className="right bold">{fmtMoney(d.totals.igst)}</td>}
            {showCGST && <td className="right bold">{fmtMoney(d.totals.cgst)}</td>}
            {showCGST && <td className="right bold">{fmtMoney(d.totals.sgst)}</td>}
            <td className="right bold">{fmtMoney(d.totals.total)}</td>
            <td className="no-print"></td>
          </tr>
        </tbody>
      </table>

      <div className="no-print add-line-row">
        <button className="btn btn-secondary btn-tiny" onClick={onAddLine}><Plus size={12} /> Add line item</button>
      </div>

      <table className="inv-footer-table">
        <tbody>
          <tr>
            <td style={{ width: "50%", verticalAlign: "top" }}>
              <div>
                <span className="label">Amount Chargeable (in words): </span>
                <span className="bold">{numberToWords(d.totals.total)} Rupees Only</span>
              </div>
              <div className="declaration">
                <div className="label" style={{ textDecoration: "underline" }}>Declaration</div>
                <EditBlock value={inv.declaration} onChange={(v) => onField("declaration", v)} />
              </div>
              <div className="eoe">E. & O.E</div>
            </td>
            <td style={{ verticalAlign: "top" }}>
              <table className="hsn-summary-table">
                <thead>
                  <tr>
                    <th>HSN Code</th>
                    <th>Tax Rate%</th>
                    <th>Qty</th>
                    <th>Taxable Amt (Rs)</th>
                    {showIGST && <th>IGST Amt (Rs)</th>}
                    {showCGST && <th>CGST Amt (Rs)</th>}
                    {showCGST && <th>SGST Amt (Rs)</th>}
                    <th>Total Tax Amt (Rs)</th>
                    <th>Total (Rs)</th>
                  </tr>
                </thead>
                <tbody>
                  {d.hsnSummary.map((h, i) => (
                    <tr key={i}>
                      <td className="center">{h.hsn}</td>
                      <td className="center">{h.taxRate.toFixed(2)}%</td>
                      <td className="center">{h.qty}</td>
                      <td className="right">{fmtMoney(h.taxable)}</td>
                      {showIGST && <td className="right">{h.taxType === "IGST" ? fmtMoney(h.taxAmt) : "—"}</td>}
                      {showCGST && <td className="right">{h.taxType === "CGST_SGST" ? fmtMoney(h.taxAmt / 2) : "—"}</td>}
                      {showCGST && <td className="right">{h.taxType === "CGST_SGST" ? fmtMoney(h.taxAmt / 2) : "—"}</td>}
                      <td className="right">{fmtMoney(h.taxAmt)}</td>
                      <td className="right">{fmtMoney(h.total)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div className="sign-block">
                <div className="bold">for <EditInput value={sender.name} onChange={(v) => onSenderField("name", v)} align="right" width="220px" /></div>
                <div className="sign-space" />
                <EditInput value={inv.signatoryLabel} onChange={(v) => onField("signatoryLabel", v)} align="right" width="180px" />
              </div>
            </td>
          </tr>
        </tbody>
      </table>

      <div className="inv-post-footer">
        <div>
          <div>This is a Computer Generated Invoice</div>
          <div className="bold">Prepared By : {sender.name}</div>
        </div>
        <div className="irn-block">
          <div>IRN: <EditInput value={inv.irn} onChange={(v) => onField("irn", v)} width="100%" /></div>
          <div>ACK No: <EditInput value={inv.ackNo} onChange={(v) => onField("ackNo", v)} width="100%" /></div>
          <div>ACK Date: <EditInput value={inv.ackDate} onChange={(v) => onField("ackDate", v)} width="100%" /></div>
        </div>
      </div>
    </div>
  );
}

// ---------- Main App ----------
export default function InvoicePortal() {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [invoices, setInvoices] = useState([]);
  const [sender, setSender] = useState(DEFAULT_SENDER);
  const [selectedId, setSelectedId] = useState(null);
  const [printAll, setPrintAll] = useState(false);
  const [dragOver, setDragOver] = useState(false);

  const selectedInvoice = useMemo(
    () => invoices.find((i) => i.id === selectedId) || invoices[0] || null,
    [invoices, selectedId]
  );

  const handleFile = useCallback((file) => {
    setError("");
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = parseWorkbook(e.target.result);
        if (!parsed.length) {
          setError("No usable rows were found. Please check the file matches the expected format.");
          setRows(null);
          return;
        }
        setRows(parsed);
        const built = buildInvoices(parsed);
        setInvoices(built);
        setSelectedId(built[0]?.id || null);
      } catch (err) {
        setError(err.message || "Could not read this file.");
        setRows(null);
      }
    };
    reader.onerror = () => setError("Could not read this file.");
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files?.[0];
    handleFile(file);
  };

  const updateInvoice = useCallback((id, mutator) => {
    setInvoices((prev) => prev.map((inv) => (inv.id === id ? mutator(_.cloneDeep(inv)) : inv)));
  }, []);

  const handleField = (id) => (field, value) =>
    updateInvoice(id, (inv) => { inv[field] = value; return inv; });

  const handleBuyerField = (id) => (field, value) =>
    updateInvoice(id, (inv) => { inv.buyer[field] = value; return inv; });

  const handleLineItem = (id) => (index, field, value) =>
    updateInvoice(id, (inv) => { inv.lineItems[index][field] = value; return inv; });

  const handleAddLine = (id) => () =>
    updateInvoice(id, (inv) => {
      inv.lineItems.push({
        key: `${id}-new-${Date.now()}`,
        description: "New item",
        hsn: "",
        qty: "1",
        rate: "0.00",
        less: "0",
        taxable: "0.00",
        taxType: "IGST",
        taxRate: "18",
        taxAmt: "0.00",
        total: "0.00",
      });
      return inv;
    });

  const handleRemoveLine = (id) => (index) =>
    updateInvoice(id, (inv) => {
      if (inv.lineItems.length > 1) inv.lineItems.splice(index, 1);
      return inv;
    });

  const handleSenderField = (field, value) => setSender((prev) => ({ ...prev, [field]: value }));

  const doPrintOne = () => window.print();

  const doPrintAll = () => {
    setPrintAll(true);
    setTimeout(() => {
      window.print();
      setPrintAll(false);
    }, 50);
  };

  const reset = () => {
    setRows(null);
    setFileName("");
    setError("");
    setInvoices([]);
    setSelectedId(null);
  };

  return (
    <div className="portal-root">
      <style>{`
        :root {
          --ink: #1a1a1a;
          --paper: #ffffff;
          --line: #1a1a1a;
          --muted: #6b6b6b;
          --shell-bg: #f5f4f0;
          --panel: #ffffff;
          --accent: #8a3b12;
          --accent-soft: #f1e4d8;
        }
        .portal-root {
          font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
          background: var(--shell-bg);
          min-height: 100%;
          color: var(--ink);
        }
        .shell-header {
          display: flex; align-items: center; justify-content: space-between;
          padding: 18px 28px; border-bottom: 1px solid #ddd9cf;
          background: var(--panel);
        }
        .brand { display: flex; align-items: baseline; gap: 10px; }
        .brand-mark {
          font-family: Georgia, 'Times New Roman', serif;
          font-weight: 700; font-size: 20px; letter-spacing: 0.2px;
        }
        .brand-sub { color: var(--muted); font-size: 12.5px; }
        .layout {
          display: grid; grid-template-columns: 300px 1fr; min-height: calc(100vh - 62px);
        }
        .sidebar {
          border-right: 1px solid #ddd9cf; background: var(--panel);
          padding: 20px; display: flex; flex-direction: column; gap: 16px;
        }
        .upload-zone {
          border: 1.5px dashed #c9c2b2; border-radius: 10px; padding: 22px 14px;
          text-align: center; cursor: pointer; transition: border-color .15s, background .15s;
          background: #fbfaf7;
        }
        .upload-zone.drag { border-color: var(--accent); background: var(--accent-soft); }
        .upload-zone input { display: none; }
        .upload-icon { color: var(--accent); margin-bottom: 6px; }
        .upload-title { font-weight: 600; font-size: 13.5px; margin-bottom: 2px; }
        .upload-sub { font-size: 11.5px; color: var(--muted); }
        .file-chip {
          display: flex; align-items: center; gap: 8px; font-size: 12.5px;
          background: #fbfaf7; border: 1px solid #e6e1d5; border-radius: 8px; padding: 8px 10px;
        }
        .file-chip svg { flex-shrink: 0; color: var(--accent); }
        .error-box {
          display: flex; gap: 8px; font-size: 12.5px; color: #7a2323;
          background: #fbeaea; border: 1px solid #f0caca; border-radius: 8px; padding: 10px;
        }
        .invoice-list-title {
          font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em;
          color: var(--muted); margin-top: 4px;
        }
        .invoice-list { display: flex; flex-direction: column; gap: 4px; overflow-y: auto; max-height: 42vh; }
        .invoice-list-item {
          display: flex; align-items: center; justify-content: space-between;
          padding: 9px 10px; border-radius: 7px; cursor: pointer; font-size: 12.5px;
          border: 1px solid transparent;
        }
        .invoice-list-item:hover { background: #f3f1ea; }
        .invoice-list-item.active { background: var(--accent-soft); border-color: #e3c9ac; }
        .invoice-list-item .num { font-weight: 600; }
        .invoice-list-item .meta { color: var(--muted); font-size: 11px; }
        .reset-btn {
          margin-top: auto; display: flex; align-items: center; gap: 6px;
          font-size: 12.5px; color: var(--muted); background: none; border: none; cursor: pointer;
          padding: 6px 2px;
        }
        .reset-btn:hover { color: var(--ink); }
        .hint-box { font-size: 11px; color: var(--muted); line-height: 1.5; background: #fbfaf7; border: 1px solid #e6e1d5; border-radius: 8px; padding: 9px 10px; }

        .content { padding: 26px 30px; overflow: auto; }
        .content-toolbar {
          display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px;
        }
        .toolbar-title { font-size: 13px; color: var(--muted); }
        .toolbar-actions { display: flex; gap: 8px; }
        .btn {
          display: inline-flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 600;
          padding: 8px 13px; border-radius: 7px; cursor: pointer; border: 1px solid transparent;
        }
        .btn-primary { background: var(--accent); color: white; }
        .btn-primary:hover { background: #6f2f0d; }
        .btn-secondary { background: white; border-color: #d8d2c2; color: var(--ink); }
        .btn-secondary:hover { background: #f3f1ea; }
        .btn-tiny { padding: 5px 10px; font-size: 11.5px; }

        .empty-state {
          display: flex; flex-direction: column; align-items: center; justify-content: center;
          height: 60vh; color: var(--muted); text-align: center; gap: 6px;
        }
        .empty-state .big { font-size: 15px; color: var(--ink); font-weight: 600; }

        .doc-wrap { display: flex; justify-content: center; }
        .doc-shadow {
          background: white; box-shadow: 0 1px 3px rgba(0,0,0,0.08), 0 8px 24px rgba(0,0,0,0.06);
          padding: 26px 30px; width: 780px; max-width: 100%;
        }

        /* ---- Editable primitives ---- */
        .ed-input {
          font: inherit; color: inherit; font-weight: inherit; border: none; background: transparent;
          border-bottom: 1px dashed transparent; padding: 0 1px; min-width: 24px;
        }
        .ed-input:hover { border-bottom-color: #c9b8a0; }
        .ed-input:focus { outline: none; border-bottom-color: var(--accent); background: #fff8ee; }
        .ed-block {
          border-bottom: 1px dashed transparent; padding: 1px; cursor: text; white-space: pre-wrap;
        }
        .ed-block:hover { border-bottom-color: #c9b8a0; }
        .ed-block:focus { outline: none; border-bottom-color: var(--accent); background: #fff8ee; }
        .icon-btn {
          border: none; background: transparent; cursor: pointer; color: #b3392c; padding: 3px;
          display: inline-flex; border-radius: 4px;
        }
        .icon-btn:hover { background: #fbeaea; }
        .add-line-row { margin: 6px 0 10px; }

        /* ---- Invoice document styling (mirrors the printed tax invoice) ---- */
        .invoice-doc { font-family: Arial, Helvetica, sans-serif; color: #000; font-size: 11.5px; }
        .invoice-doc .inv-title { text-align: center; font-weight: 700; font-size: 13.5px; margin-bottom: 6px; }
        .invoice-doc table { width: 100%; border-collapse: collapse; }
        .invoice-doc .inv-header-table td.hcell { border: 1px solid #000; padding: 5px 7px; vertical-align: top; }
        .invoice-doc .label { color: #333; font-size: 10.5px; }
        .invoice-doc .bold { font-weight: 700; }
        .invoice-doc .inv-items-table { margin-top: -1px; }
        .invoice-doc .inv-items-table th, .invoice-doc .inv-items-table td {
          border: 1px solid #000; padding: 5px 6px; font-size: 11px;
        }
        .invoice-doc .inv-items-table th { background: #f2f2f2; font-weight: 700; }
        .invoice-doc .center { text-align: center; }
        .invoice-doc .right { text-align: right; }
        .invoice-doc .hsn-line { font-size: 10px; color: #333; margin-top: 2px; }
        .invoice-doc .rate-line { font-size: 9.5px; color: #333; }
        .invoice-doc .totals-row td { font-weight: 700; background: #fafafa; }
        .invoice-doc .inv-footer-table { margin-top: -1px; }
        .invoice-doc .inv-footer-table > tbody > tr > td { border: 1px solid #000; padding: 7px; vertical-align: top; }
        .invoice-doc .declaration { margin-top: 10px; font-size: 10.5px; }
        .invoice-doc .eoe { text-align: right; font-size: 10.5px; margin-top: 10px; }
        .invoice-doc .hsn-summary-table th, .invoice-doc .hsn-summary-table td {
          border: 1px solid #999; padding: 3px 5px; font-size: 10px;
        }
        .invoice-doc .hsn-summary-table th { background: #f2f2f2; }
        .invoice-doc .sign-block { text-align: right; margin-top: 14px; font-size: 11px; }
        .invoice-doc .sign-space { height: 46px; }
        .invoice-doc .inv-post-footer {
          display: flex; justify-content: space-between; margin-top: 14px; font-size: 10.5px; color: #222;
        }
        .invoice-doc .irn-block { text-align: right; max-width: 320px; }
        .invoice-doc .irn-block .ed-input { text-align: right; width: 260px; }

        .print-all-container { display: none; }

        @media print {
          .no-print { display: none !important; }
          .shell-header, .sidebar, .content-toolbar { display: none !important; }
          .layout { display: block !important; }
          .content { padding: 0 !important; }
          .doc-wrap { display: block !important; }
          .doc-shadow { box-shadow: none !important; width: 100% !important; padding: 0 !important; }
          .portal-root { background: white !important; }
          .ed-input, .ed-block { border-bottom-color: transparent !important; background: transparent !important; }
          body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
        .printing-all .doc-wrap.single { display: none; }
        .printing-all .print-all-container { display: block; }
        .print-all-container .doc-shadow { box-shadow: none; width: 100%; padding: 0; margin-bottom: 0; }
        .print-all-container .invoice-page { page-break-after: always; padding: 20px 0; }
        .print-all-container .invoice-page:last-child { page-break-after: auto; }
      `}</style>

      <div className={printAll ? "printing-all" : ""}>
        <div className="shell-header no-print">
          <div className="brand">
            <span className="brand-mark">Challan → Invoice</span>
            <span className="brand-sub">generate tax invoices from your RTO upload sheet</span>
          </div>
          {fileName && <div className="file-chip"><FileSpreadsheet size={15} /> {fileName}</div>}
        </div>

        <div className="layout">
          <div className="sidebar no-print">
            <label className={`upload-zone ${dragOver ? "drag" : ""}`}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
            >
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleFile(e.target.files?.[0])} />
              <div className="upload-icon"><UploadCloud size={26} /></div>
              <div className="upload-title">{rows ? "Upload a different file" : "Upload your RTO upload sheet"}</div>
              <div className="upload-sub">.xlsx, .xls or .csv · click or drag & drop</div>
            </label>

            {error && (
              <div className="error-box">
                <AlertTriangle size={15} style={{ flexShrink: 0, marginTop: 1 }} />
                <div>{error}</div>
              </div>
            )}

            {invoices.length > 0 && (
              <>
                <div className="invoice-list-title">{invoices.length} invoice{invoices.length > 1 ? "s" : ""} found</div>
                <div className="invoice-list">
                  {invoices.map((inv) => (
                    <div
                      key={inv.id}
                      className={`invoice-list-item ${selectedInvoice?.id === inv.id ? "active" : ""}`}
                      onClick={() => setSelectedId(inv.id)}
                    >
                      <div>
                        <div className="num">{inv.invoiceNo}</div>
                        <div className="meta">{inv.buyer.name}</div>
                      </div>
                      <ChevronRight size={14} color="#999" />
                    </div>
                  ))}
                </div>
                <div className="hint-box">Every field on the invoice — sender, buyer, line items, totals text, IRN block — is editable. Click any text to edit it, and use "Add line item" for extra rows.</div>
              </>
            )}

            {rows && (
              <button className="reset-btn" onClick={reset}>
                <RotateCcw size={13} /> Start over
              </button>
            )}
          </div>

          <div className="content">
            {!selectedInvoice ? (
              <div className="empty-state no-print">
                <div className="big">No invoice selected yet</div>
                <div>Upload your upload-format sheet on the left to generate invoices.</div>
              </div>
            ) : (
              <>
                <div className="content-toolbar no-print">
                  <div className="toolbar-title">
                    Invoice {selectedInvoice.invoiceNo} · {selectedInvoice.lineItems.length} line item{selectedInvoice.lineItems.length > 1 ? "s" : ""}
                  </div>
                  <div className="toolbar-actions">
                    {invoices.length > 1 && (
                      <button className="btn btn-secondary" onClick={doPrintAll}>
                        <Printer size={14} /> Print all {invoices.length}
                      </button>
                    )}
                    <button className="btn btn-primary" onClick={doPrintOne}>
                      <Printer size={14} /> Print / Save as PDF
                    </button>
                  </div>
                </div>

                <div className="doc-wrap single">
                  <div className="doc-shadow">
                    <InvoiceDoc
                      inv={selectedInvoice}
                      sender={sender}
                      onField={handleField(selectedInvoice.id)}
                      onBuyerField={handleBuyerField(selectedInvoice.id)}
                      onLineItem={handleLineItem(selectedInvoice.id)}
                      onAddLine={handleAddLine(selectedInvoice.id)}
                      onRemoveLine={handleRemoveLine(selectedInvoice.id)}
                      onSenderField={handleSenderField}
                    />
                  </div>
                </div>
              </>
            )}

            <div className="print-all-container">
              {invoices.map((inv) => (
                <div className="invoice-page" key={inv.id}>
                  <div className="doc-shadow">
                    <InvoiceDoc
                      inv={inv}
                      sender={sender}
                      onField={handleField(inv.id)}
                      onBuyerField={handleBuyerField(inv.id)}
                      onLineItem={handleLineItem(inv.id)}
                      onAddLine={handleAddLine(inv.id)}
                      onRemoveLine={handleRemoveLine(inv.id)}
                      onSenderField={handleSenderField}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
