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
    """Return ledger masters with parent, opening and closing balances.

    When ``start`` and ``end`` are provided, opening balances are as of
    ``start`` and closing balances as of ``end`` using Day Book movements.
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
            billwise = ledger.findtext("ISBILLWISEON", "No")
            ledgers[name] = {
                "Parent": parent,
                "Opening Balance": opening,
                "Billwise": billwise,
                "Closing Balance": 0.0,
            }
    except ET.ParseError:
        return []

    if start is None or end is None:
        _populate_closing_from_trial_balance(company_name, host, port, ledgers)
        return [
            {
                "Ledger Name": name,
                **values,
                "Nett": round(values["Opening Balance"] - values["Closing Balance"], 2),
            }
            for name, values in sorted(ledgers.items())
        ]

    fy_start = _financial_year_start(start)
    vouchers = fetch_daybook(company_name, fy_start, end, host, port)

    pre_start_delta: Dict[str, float] = {}
    in_period_delta: Dict[str, float] = {}
    for voucher in vouchers:
        for entry in voucher.ledger_entries:
            delta = entry.amount if entry.is_debit else -entry.amount
            if voucher.date < start:
                pre_start_delta[entry.ledger_name] = pre_start_delta.get(
                    entry.ledger_name, 0.0
                ) + delta
            else:
                in_period_delta[entry.ledger_name] = in_period_delta.get(
                    entry.ledger_name, 0.0
                ) + delta

    for name, values in ledgers.items():
        opening = values["Opening Balance"] + pre_start_delta.get(name, 0.0)
        closing = opening + in_period_delta.get(name, 0.0)
        values["Opening Balance"] = round(opening, 2)
        values["Closing Balance"] = round(closing, 2)
        values["Nett"] = round(opening - closing, 2)

    return [
        {
            "Ledger Name": name,
            **values,
        }
        for name, values in sorted(ledgers.items())
    ]


def _populate_closing_from_trial_balance(
    company_name: str, host: str, port: int, ledgers: Dict[str, Dict[str, str | float]]
) -> None:
    today = date.today().strftime("%Y%m%d")
    fy_start = _financial_year_start(date.today()).strftime("%Y%m%d")
    xml_tb = f"""
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
          <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        </STATICVARIABLES>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>
"""
    raw_tb = _clean_tally_xml(_post_xml(xml_tb, host, port))
    try:
        root_tb = ET.fromstring(raw_tb)
        for ledger in root_tb.findall(".//LEDGER"):
            name = (ledger.findtext("NAME", "") or "").strip()
            if name in ledgers:
                ledgers[name]["Closing Balance"] = _to_float(
                    ledger.findtext("CLOSINGBALANCE", "0")
                )
    except ET.ParseError:
        return


def _financial_year_start(anchor: date) -> date:
    year = anchor.year
    if anchor.month < 4:
        year -= 1
    return date(year, 4, 1)


def _parse_daybook(raw: str) -> Iterable[Voucher]:
    root = ET.fromstring(raw)
    for voucher in root.findall(".//VOUCHER"):
        vdate_text = voucher.findtext("DATE", "")
        if len(vdate_text) != 8:
            continue
        vdate = date.fromisoformat(f"{vdate_text[:4]}-{vdate_text[4:6]}-{vdate_text[6:]}")
        vtype = voucher.findtext("VOUCHERTYPENAME", "")
        narration = (voucher.findtext("NARRATION", "") or "").strip()
        entries: List[LedgerEntry] = []
        for entry in voucher.findall(".//ALLLEDGERENTRIES.LIST"):
            ledger = entry.findtext("LEDGERNAME", "")
            amount_raw = _to_float(entry.findtext("AMOUNT", "0"))

            # Tally's ISDEEMEDPOSITIVE is the most reliable indicator:
            #   "Yes"  => Credit, "No" => Debit. Only fall back to the sign
            # of the amount when the flag is missing so debits/credits match
            # what Tally shows for Dr/Cr.
            deemed_text = (entry.findtext("ISDEEMEDPOSITIVE", "") or "").strip().lower()
            if deemed_text in ("yes", "y", "true"):
                is_debit = False
            elif deemed_text in ("no", "n", "false"):
                is_debit = True
            elif amount_raw < 0:
                is_debit = False
            else:
                is_debit = True

            entries.append(
                LedgerEntry(
                    ledger_name=ledger,
                    amount=abs(amount_raw),
                    is_debit=is_debit,
                )
            )
        yield Voucher(voucher_type=vtype, date=vdate, ledger_entries=entries, narration=narration)


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

