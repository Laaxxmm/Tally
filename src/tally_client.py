"""Robust helpers for connecting to Tally over the HTTP XML interface.

The functions here mirror the working reference script shared by the user:
they clean malformed XML, post requests to 127.0.0.1:9000, and parse Day Book
and ledger responses into friendly Python structures that downstream analytics
can consume.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Dict, Iterable, List
import re
import xml.etree.ElementTree as ET
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import requests


@dataclass
class LedgerEntry:
    ledger_name: str
    amount: float
    is_debit: bool


@dataclass
class Voucher:
    voucher_type: str
    date: date
    ledger_entries: List[LedgerEntry]
    narration: str = ""
    voucher_number: str | None = None


__all__ = [
    "LedgerEntry",
    "Voucher",
    "fetch_companies",
    "fetch_daybook",
    "fetch_ledgers",
    "fetch_ledger_master",
    "fetch_group_master",
    "export_ledger_opening_excel",
    "export_group_master_excel",
]


def _clean_tally_xml(resp: str | None) -> str:
    if not resp:
        return ""
    resp = resp.replace("&", "&amp;")
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", resp)


def _post_xml(xml: str, host: str, port: int) -> str:
    url = f"http://{host}:{port}"
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    data = xml.encode("utf-8")
    req = Request(url, data=data, headers=headers, method="POST")
    try:
        with urlopen(req, timeout=90) as resp:
            return resp.read().decode("utf-8")
    except (HTTPError, URLError) as exc:
        raise ConnectionError(f"Tally connection failed: {exc}")


def fetch_companies(host: str, port: int) -> List[str]:
    xml = """
<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>List of Companies</ID></HEADER>
  <BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES></DESC></BODY>
</ENVELOPE>
"""
    raw = _clean_tally_xml(_post_xml(xml, host, port))
    if not raw:
        return []
    try:
        root = ET.fromstring(raw)
        return sorted({c.get("NAME") for c in root.findall(".//COMPANY") if c.get("NAME")})
    except ET.ParseError:
        return []


def fetch_daybook(company_name: str, start: date, end: date, host: str, port: int) -> List[Voucher]:
    from_str = start.strftime("%Y%m%d")
    to_str = end.strftime("%Y%m%d")
    xml = f"""
<ENVELOPE>
<HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
<BODY>
  <EXPORTDATA>
    <REQUESTDESC>
      <REPORTNAME>Voucher Register</REPORTNAME>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
        <SVFROMDATE>{from_str}</SVFROMDATE>
        <SVTODATE>{to_str}</SVTODATE>
        <EXPLODEFLAG>Yes</EXPLODEFLAG>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
    </REQUESTDESC>
  </EXPORTDATA>
</BODY>
</ENVELOPE>
"""
    raw = _clean_tally_xml(_post_xml(xml, host, port))
    if not raw:
        return []
    try:
        return list(_parse_daybook(raw))
    except ET.ParseError:
        return []


def fetch_ledgers(
    company_name: str,
    host: str,
    port: int,
    start: date | None = None,
    end: date | None = None,
    ) -> List[Dict[str, str | float]]:
    """Return chart-of-accounts ledgers with parent group and opening balance.

    This function is kept for backward compatibility; it now proxies to the
    more explicit ledger master extractor used for Excel exports.
    """

    ledger_rows = fetch_ledger_master(company_name, host, port)
    return [
        {
            "Ledger Name": row["LedgerName"],
            "Under": row["LedgerParent"],
            "Opening Balance": row["OpeningBalanceNormalized"],
        }
        for row in ledger_rows
    ]


def _parse_daybook(raw: str) -> Iterable[Voucher]:
    root = ET.fromstring(raw)
    for voucher in root.findall(".//VOUCHER"):
        vdate_text = voucher.findtext("DATE", "")
        if len(vdate_text) != 8:
            continue
        vdate = date.fromisoformat(f"{vdate_text[:4]}-{vdate_text[4:6]}-{vdate_text[6:]}")
        vtype = _extract_voucher_type(voucher)
        narration = (voucher.findtext("NARRATION", "") or "").strip()
        vnumber = (voucher.findtext("VOUCHERNUMBER", "") or "").strip() or None
        entries: List[LedgerEntry] = []
        entry_tags = [
            "ALLLEDGERENTRIES.LIST",
            "LEDGERENTRIES.LIST",
            "ALLINVENTORYENTRIES.LIST",
            "INVENTORYENTRIES.LIST",
        ]

        for tag in entry_tags:
            for entry in voucher.findall(f".//{tag}"):
                ledger = (
                    entry.findtext("LEDGERNAME", "")
                    or entry.findtext("STOCKITEMNAME", "")
                    or "(Unknown Ledger)"
                )
                amount_raw = _to_float(entry.findtext("AMOUNT", "0"))
                deemed_text = (entry.findtext("ISDEEMEDPOSITIVE", "") or "").strip().lower()
                amount_abs = abs(amount_raw)

                # Prefer the raw amount sign for Dr/Cr; fall back to the deemed flag
                # only when the amount is zero (some exports flip the sign for cr).
                if amount_raw > 0:
                    is_debit = True
                elif amount_raw < 0:
                    is_debit = False
                elif deemed_text in ("no", "n", "false"):
                    is_debit = True
                elif deemed_text in ("yes", "y", "true"):
                    is_debit = False
                else:
                    continue

                if amount_abs == 0:
                    continue

                entries.append(
                    LedgerEntry(
                        ledger_name=ledger,
                        amount=amount_abs,
                        is_debit=is_debit,
                    )
                )
        if entries:
            yield Voucher(
                voucher_type=vtype,
                date=vdate,
                ledger_entries=entries,
                narration=narration,
                voucher_number=vnumber,
        )


def _to_float(value: str | None) -> float:
    cleaned = (value or "0").replace(",", "").strip()
    if not cleaned or cleaned == "-":
        return 0.0

    # Handle strings like "12345 Cr" / "12345 Dr" where Dr is positive and Cr
    # is negative, matching common Tally exports for opening/closing balances.
    match = re.match(r"(-?\d*\.?\d+)(?:\s*(Dr|Cr))?", cleaned, flags=re.IGNORECASE)
    if not match:
        return 0.0

    number_part, drcr = match.groups()
    try:
        value_flt = float(number_part)
    except ValueError:
        return 0.0

    if drcr:
        if drcr.lower() == "cr":
            value_flt = -abs(value_flt)
        else:
            value_flt = abs(value_flt)
    return value_flt


def _extract_parent(node: ET.Element) -> str:
    return _first_non_empty(
        [
            node.findtext("PARENT"),
            node.get("PARENT"),
            node.findtext("PARENTNAME"),
            node.get("PARENTNAME"),
            node.findtext(".//PARENT"),
        ]
    )


def _first_non_empty(candidates):
    for candidate in candidates:
        cleaned = (candidate or "").strip()
        if cleaned:
            return cleaned
    return ""


def _fiscal_year_start(anchor: date) -> date:
    """Return an April 1 fiscal-year start for the given anchor date."""

    year = anchor.year if anchor.month >= 4 else anchor.year - 1
    return date(year, 4, 1)


# ---------------------------------------------------------------------------
# Ledger master extraction with normalized opening balances
# ---------------------------------------------------------------------------

# Sample XML envelope to pull ledger masters from Tally's chart of accounts.
LEDGER_MASTER_REQUEST_TEMPLATE = """
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Ledger Master List</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
        <REMOTECMPNAME>{company_name}</REMOTECMPNAME>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Ledger Master List" ISMODIFY="No">
            <TYPE>Ledger</TYPE>
            <BELONGSTO>Yes</BELONGSTO>
            <FETCH>NAME</FETCH>
            <FETCH>PARENT</FETCH>
            <FETCH>PARENTNAME</FETCH>
            <FETCH>OPENINGBALANCE</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>
"""


LEDGER_MASTER_FALLBACK_REQUEST = """
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>List of Ledgers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
        <REMOTECMPNAME>{company_name}</REMOTECMPNAME>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>
"""


def fetch_ledger_master(company_name: str, host: str, port: int) -> List[Dict[str, str | float]]:
    """Fetch ledger names, parents, and opening balances from Tally.

    Uses the requests library to post XML to Tally's HTTP listener. Opening
    balances are parsed from the Dr/Cr suffixed values Tally returns.
    """

    url = f"http://{host}:{port}"
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    xml_body = LEDGER_MASTER_REQUEST_TEMPLATE.format(company_name=company_name)

    # Attempt the HTTP post; if Tally is not running or HTTP is disabled, raise
    # a clear connection error for the caller to surface.
    try:
        response = requests.post(url, data=xml_body.encode("utf-8"), headers=headers, timeout=60)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise ConnectionError("Tally is not reachable. Ensure it is running with HTTP enabled.") from exc

    if not response.text or not response.text.strip():
        raise RuntimeError("Empty response received from Tally when requesting ledger masters.")

    cleaned = _clean_tally_xml(response.text)
    rows = _parse_ledger_master(cleaned)

    # Some Tally builds omit parents/openings unless you use the built-in
    # "List of Ledgers" collection. If the TDL collection yields nothing, fall
    # back to the built-in export before surfacing an error to the UI.
    if not rows:
        try:
            fallback_body = LEDGER_MASTER_FALLBACK_REQUEST.format(company_name=company_name)
            fb_resp = requests.post(url, data=fallback_body.encode("utf-8"), headers=headers, timeout=60)
            fb_resp.raise_for_status()
        except requests.RequestException as exc:
            raise ConnectionError("Tally is not reachable. Ensure it is running with HTTP enabled.") from exc

        cleaned_fb = _clean_tally_xml(fb_resp.text)
        rows = _parse_ledger_master(cleaned_fb)

    if not rows:
        raise RuntimeError(
            "No ledgers returned from Tally. Open the company, keep it active, and ensure HTTP access is enabled."
        )

    print(f"Extracted {len(rows)} ledgers from Tally")
    return sorted(rows, key=lambda r: r["LedgerName"].lower())


def export_ledger_opening_excel(company_name: str, host: str, port: int) -> bytes:
    """Return an Excel workbook (as bytes) containing ledger master openings."""

    import io
    import pandas as pd

    rows = fetch_ledger_master(company_name, host, port)
    df = pd.DataFrame(rows, columns=[
        "LedgerName",
        "LedgerParent",
        "OpeningBalanceRaw",
        "OpeningBalanceNormalized",
    ])

    output = io.BytesIO()
    # Use openpyxl via pandas to build the Excel file for download.
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return output.read()


# ---------------------------------------------------------------------------
# Group master extraction with BS/P&L classification
# ---------------------------------------------------------------------------

GROUP_MASTER_REQUEST_TEMPLATE = """
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Group Master List</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
        <REMOTECMPNAME>{company_name}</REMOTECMPNAME>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Group Master List" ISMODIFY="No">
            <TYPE>Group</TYPE>
            <BELONGSTO>Yes</BELONGSTO>
            <FETCH>NAME</FETCH>
            <FETCH>PARENT</FETCH>
            <FETCH>PARENTNAME</FETCH>
            <FETCH>NATUREOFGROUP</FETCH>
            <FETCH>ISREVENUE</FETCH>
            <FETCH>AFFECTSGROSSPROFIT</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>
"""


GROUP_MASTER_FALLBACK_REQUEST = """
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>List of Groups</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
        <REMOTECMPNAME>{company_name}</REMOTECMPNAME>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>
"""


def fetch_group_master(company_name: str, host: str, port: int) -> List[Dict[str, str | float | bool]]:
    """Return chart-of-account groups with derived BS/P&L classifications."""

    url = f"http://{host}:{port}"
    headers = {"Content-Type": "text/xml; charset=utf-8"}
    body = GROUP_MASTER_REQUEST_TEMPLATE.format(company_name=company_name)

    try:
        response = requests.post(url, data=body.encode("utf-8"), headers=headers, timeout=60)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise ConnectionError("Tally is not reachable. Ensure it is running with HTTP enabled.") from exc

    cleaned = _clean_tally_xml(response.text)
    rows = _parse_group_master(cleaned)

    if not rows:
        try:
            fallback_body = GROUP_MASTER_FALLBACK_REQUEST.format(company_name=company_name)
            fb_resp = requests.post(url, data=fallback_body.encode("utf-8"), headers=headers, timeout=60)
            fb_resp.raise_for_status()
        except requests.RequestException as exc:
            raise ConnectionError("Tally is not reachable. Ensure it is running with HTTP enabled.") from exc

        rows = _parse_group_master(_clean_tally_xml(fb_resp.text))

    if not rows:
        raise RuntimeError("No groups returned from Tally. Open the company and retry.")

    print(f"Extracted {len(rows)} groups from Tally")
    return sorted(rows, key=lambda r: r["GroupName"].lower())


def export_group_master_excel(company_name: str, host: str, port: int) -> bytes:
    """Return Excel bytes for the group master extract."""

    import io
    import pandas as pd

    rows = fetch_group_master(company_name, host, port)
    df = pd.DataFrame(rows, columns=[
        "GroupName",
        "ParentName",
        "BS_or_PnL",
        "Type",
        "AffectsGrossProfit",
    ])

    output = io.BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return output.read()


def _normalize_drcr(value: str) -> float:
    """Convert Dr/Cr suffixed opening balances to signed floats."""

    cleaned = (value or "").replace(",", "").strip()
    if not cleaned:
        return 0.0

    match = re.match(r"(-?\d*\.?\d+)(?:\s*(Dr|Cr))?", cleaned, flags=re.IGNORECASE)
    if not match:
        return 0.0

    number_part, drcr = match.groups()
    try:
        number_val = float(number_part)
    except ValueError:
        return 0.0

    if drcr:
        # Explicit Dr/Cr suffix wins: Dr should be positive, Cr should be negative
        # regardless of the sign emitted in the raw value.
        if drcr.lower() == "cr":
            return -abs(number_val)
        return abs(number_val)

    # When no Dr/Cr marker is present, invert the original sign so ledgers that
    # surfaced as negative debits / positive credits are normalized to
    # "Debit = positive, Credit = negative" as requested.
    return -number_val


def _parse_ledger_master(raw: str) -> List[Dict[str, str | float]]:
    """Convert a Tally ledger master XML payload into structured rows."""

    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise RuntimeError("Unable to parse Tally ledger master response.") from exc

    rows: List[Dict[str, str | float]] = []
    for ledger in root.findall(".//LEDGER"):
        name = _first_non_empty([ledger.get("NAME"), ledger.findtext("NAME")])
        parent = _extract_parent(ledger) or "(Unknown)"
        opening_raw = (ledger.findtext("OPENINGBALANCE") or ledger.get("OPENINGBALANCE") or "0").strip()
        opening_norm = _normalize_drcr(opening_raw)

        if not name:
            continue

        rows.append(
            {
                "LedgerName": name,
                "LedgerParent": parent,
                "OpeningBalanceRaw": opening_raw,
                "OpeningBalanceNormalized": opening_norm,
            }
        )

    return rows


def get_parent_name(node: ET.Element) -> str:
    """Return the group's parent, defaulting to itself when blank."""

    parent = _extract_parent(node)
    if parent:
        return parent
    name = _first_non_empty([node.get("NAME"), node.findtext("NAME")])
    return name or ""


def classify_bs_or_pnl(is_revenue: str | None, nature: str | None) -> str:
    """Determine whether a group contributes to Balance Sheet or P&L using Tally flags."""

    is_revenue_clean = (is_revenue or "").strip().lower()
    if is_revenue_clean in {"yes", "y", "true"}:
        return "P&L"
    if is_revenue_clean in {"no", "n", "false"}:
        return "Balance Sheet"

    nature_clean = (nature or "").lower()
    if any(token in nature_clean for token in ["income", "expense", "purchase", "sale"]):
        return "P&L"
    if nature_clean:
        return "Balance Sheet"
    return "Balance Sheet"


def classify_type(nature: str | None, group_name: str, parent_name: str) -> str:
    """Classify a group as Asset, Liability, Income, or Expense using Tally's nature."""

    nature_clean = (nature or "").lower()
    if "asset" in nature_clean:
        return "Asset"
    if "liabilit" in nature_clean:
        return "Liability"
    if "income" in nature_clean or "revenue" in nature_clean or "sale" in nature_clean:
        return "Income"
    if "expense" in nature_clean or "purchase" in nature_clean:
        return "Expense"

    # Fallback: infer from parent/name keywords directly from Tally labels
    normalized = " ".join([group_name or "", parent_name or ""]).lower()
    if "income" in normalized or "revenue" in normalized or "sale" in normalized:
        return "Income"
    if "expense" in normalized or "purchase" in normalized:
        return "Expense"
    if "liability" in normalized:
        return "Liability"
    if "asset" in normalized:
        return "Asset"
    return "Asset"


def determine_affects_gross_profit(affects_gp_flag: str | None, nature: str | None, group_name: str, parent_name: str) -> str:
    """Return "Yes" when the group affects gross profit using Tally's own flag."""

    flag_clean = (affects_gp_flag or "").strip().lower()
    if flag_clean in {"yes", "y", "true"}:
        return "Yes"
    if flag_clean in {"no", "n", "false"}:
        return "No"

    nature_clean = (nature or "").lower()
    if any(token in nature_clean for token in ["direct", "purchase", "sales", "sale", "trading"]):
        return "Yes"

    labels = " ".join([group_name or "", parent_name or ""]).lower()
    if any(token in labels for token in ["sales", "sale", "purchase", "direct income", "direct expense"]):
        return "Yes"
    return "No"


def _parse_group_master(raw: str) -> List[Dict[str, str | float | bool]]:
    """Parse a Tally group master payload into structured rows."""

    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise RuntimeError("Unable to parse Tally group master response.") from exc

    rows: List[Dict[str, str | float | bool]] = []
    for group in root.findall(".//GROUP"):
        name = _first_non_empty([group.get("NAME"), group.findtext("NAME")])
        if not name:
            continue

        parent = get_parent_name(group)
        nature = _first_non_empty([group.findtext("NATUREOFGROUP"), group.get("NATUREOFGROUP")])
        is_revenue = _first_non_empty([group.findtext("ISREVENUE"), group.get("ISREVENUE")])
        affects_gp_flag = _first_non_empty([group.findtext("AFFECTSGROSSPROFIT"), group.get("AFFECTSGROSSPROFIT")])

        bs_or_pnl = classify_bs_or_pnl(is_revenue, nature)
        gtype = classify_type(nature, name, parent)
        affects_gp = determine_affects_gross_profit(affects_gp_flag, nature, name, parent)

        rows.append(
            {
                "GroupName": name,
                "ParentName": parent or name,
                "BS_or_PnL": bs_or_pnl,
                "Type": gtype,
                "AffectsGrossProfit": affects_gp,
            }
        )

    return rows


def _extract_voucher_type(voucher: ET.Element) -> str:
    """Return a voucher type using common fields plus attributes."""

    def _clean(text: str | None) -> str:
        return (text or "").strip()

    candidates = [
        voucher.get("VCHTYPE"),
        voucher.get("VOUCHERTYPE"),
        voucher.findtext("VOUCHERTYPENAME"),
        voucher.findtext("VCHTYPE"),
        voucher.findtext("VOUCHERTYPE"),
        voucher.findtext("BASICVCHTYPE"),
    ]

    for candidate in candidates:
        cleaned = _clean(candidate)
        if cleaned:
            return cleaned

    return "(Unknown)"

