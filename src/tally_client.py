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
    "export_ledger_opening_excel",
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
        if drcr.lower() == "cr":
            return -abs(number_val)
        return abs(number_val)
    return number_val


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

