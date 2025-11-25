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
    """Return master ledger groupings and opening balances.

    The current extract focuses on ledger masters only: ledger name, parent
    group, the parent's nature (e.g., Assets, Liabilities, Income, Expense),
    whether that nature flows to the Balance Sheet or Profit & Loss, and the
    opening balance. Date parameters are accepted for API compatibility but are
    not used because this view is master-data oriented.
    """

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
            <FETCH>OpeningBalance</FETCH>
            <FETCH>IsBillWiseOn</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>
"""
    raw_basic = _clean_tally_xml(_post_xml(xml_basic, host, port))
    ledgers: Dict[str, Dict[str, str | float]] = {}
    try:
        root_basic = ET.fromstring(raw_basic)
        for ledger in root_basic.findall(".//LEDGER"):
            name = (ledger.get("NAME") or "").strip()
            if not name:
                continue
            parent = (ledger.findtext("PARENT", default="") or ledger.get("PARENT", "")).strip()
            opening = _to_float(ledger.findtext("OPENINGBALANCE", "0"))
            ledgers[name] = {
                "Parent": parent,
                "Opening Balance": opening,
            }
    except ET.ParseError:
        return []

    group_nature = _fetch_group_nature(company_name, host, port)

    def statement_for(nature: str) -> str:
        lowered = nature.lower()
        if lowered in {"income", "expense"}:
            return "Profit & Loss"
        if lowered in {"assets", "liabilities"}:
            return "Balance Sheet"
        return "(Unknown)"

    rows: List[Dict[str, str | float]] = []
    for name, values in sorted(ledgers.items()):
        parent = values.get("Parent", "")
        nature = group_nature.get(parent, "")
        rows.append(
            {
                "Ledger Name": name,
                "Parent Group": parent,
                "Group Nature": nature or "(Unknown)",
                "Statement": statement_for(nature) if nature else "(Unknown)",
                "Opening Balance": round(values.get("Opening Balance", 0.0), 2),
            }
        )

    return rows


def _fetch_group_nature(company_name: str, host: str, port: int) -> Dict[str, str]:
    """Return a mapping of group name -> nature (Assets/Liabilities/Income/Expense)."""

    xml_groups = f"""
<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>List of Groups</ID></HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{company_name}</SVCURRENTCOMPANY>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="List of Groups" ISMODIFY="No">
            <TYPE>Group</TYPE>
            <BELONGSTO>Yes</BELONGSTO>
            <FETCH>Name</FETCH>
            <FETCH>Parent</FETCH>
            <FETCH>NatureOfGroup</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>
"""
    raw_groups = _clean_tally_xml(_post_xml(xml_groups, host, port))
    mapping: Dict[str, str] = {}
    try:
        root_groups = ET.fromstring(raw_groups)
        for group in root_groups.findall(".//GROUP"):
            name = (group.get("NAME") or group.findtext("NAME", "") or "").strip()
            nature = (group.findtext("NATUREOFGROUP", "") or group.get("NATUREOFGROUP", "")).strip()
            if name:
                mapping[name] = nature
    except ET.ParseError:
        return {}
    return mapping


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

