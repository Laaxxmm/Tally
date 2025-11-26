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

    Parent (Under) values come from the ledger master collection (with TDL
    fetches for parent + opening balance). Opening balances are refreshed from
    the Trial Balance report to preserve Dr/Cr directions seen inside Tally.
    """

    # Step 1: pull ledger masters for names, parent groups, and opening values.
    xml_basic = f"""
<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>List of Ledgers</ID></HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="List of Ledgers" ISMODIFY="No">
            <TYPE>Ledger</TYPE>
            <BELONGSTO>Yes</BELONGSTO>
            <FETCH>Name</FETCH>
            <FETCH>Parent</FETCH>
            <FETCH>Parent Name</FETCH>
            <FETCH>ParentName</FETCH>
            <FETCH>OpeningBalance</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>
"""
    raw_basic = _clean_tally_xml(_post_xml(xml_basic, host, port))
    parent_map: Dict[str, str] = {}
    master_opening: Dict[str, float] = {}
    try:
        root_basic = ET.fromstring(raw_basic)
        for ledger in root_basic.findall(".//LEDGER"):
            name = (ledger.get("NAME") or "").strip()
            if not name:
                continue
            parent = _extract_parent(ledger)
            parent_map[name] = parent or "(Unknown)"
            master_opening[name] = _to_float(ledger.findtext("OPENINGBALANCE"))
    except ET.ParseError:
        return []

    # Step 2: pull opening balances from Trial Balance so Dr/Cr values match Tally.
    today = (end or date.today()).strftime("%Y%m%d")
    fy_start = _fiscal_year_start(end or date.today()).strftime("%Y%m%d")
    xml_trial = f"""
<ENVELOPE>
  <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Trial Balance</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
          <SVFROMDATE>{fy_start}</SVFROMDATE>
          <SVTODATE>{today}</SVTODATE>
          <ISDETAILED>Yes</ISDETAILED>
          <EXPLODEFLAG>Yes</EXPLODEFLAG>
          <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        </STATICVARIABLES>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>
"""
    opening_map: Dict[str, float] = {}
    try:
        raw_trial = _clean_tally_xml(_post_xml(xml_trial, host, port))
        root_trial = ET.fromstring(raw_trial)
        for ledger in root_trial.findall(".//LEDGER"):
            name = _first_non_empty([ledger.get("NAME"), ledger.findtext("NAME")])
            if not name:
                continue
            opening_val = _to_float(ledger.findtext("OPENINGBALANCE", "0"))
            opening_map[name.strip()] = round(opening_val, 2)
            if name.strip() not in parent_map:
                trial_parent = _extract_parent(ledger)
                parent_map[name.strip()] = trial_parent or "(Unknown)"
    except ET.ParseError:
        pass

    # Step 3: combine parents + openings and return sorted rows.
    rows: List[Dict[str, str | float]] = []
    for name, parent in parent_map.items():
        rows.append(
            {
                "Ledger Name": name,
                "Under": parent,
                "Opening Balance": opening_map.get(name, master_opening.get(name, 0.0)),
            }
        )

    for name, opening in opening_map.items():
        if name not in parent_map:
            rows.append(
                {
                    "Ledger Name": name,
                    "Under": "(Unknown)",
                    "Opening Balance": opening,
                }
            )

    return sorted(rows, key=lambda r: r["Ledger Name"].lower())


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

