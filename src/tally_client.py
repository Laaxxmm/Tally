"""Utilities for pulling financial data from Tally over the HTTP XML interface.

This module focuses on three capabilities:
- Establishing a connection to a running Tally instance (127.0.0.1:9000 by default).
- Requesting Day Book entries for a date range.
- Converting the XML responses into structured Python dictionaries that are easy to aggregate.

The client intentionally uses only the Python standard library so it can run anywhere
without extra dependencies.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Iterable, List
import http.client
import xml.etree.ElementTree as ET


@dataclass
class LedgerEntry:
    """Represents a ledger line extracted from a Day Book voucher."""

    ledger_name: str
    amount: float
    is_debit: bool


@dataclass
class Voucher:
    """Represents a voucher extracted from a Day Book response."""

    voucher_type: str
    date: date
    ledger_entries: List[LedgerEntry]

    @property
    def total(self) -> float:
        return sum(entry.amount if entry.is_debit else -entry.amount for entry in self.ledger_entries)


class TallyClient:
    """Minimal XML client for communicating with a running Tally instance."""

    def __init__(self, host: str = "127.0.0.1", port: int = 9000) -> None:
        self.host = host
        self.port = port

    def _post_xml(self, xml_body: str) -> str:
        connection = http.client.HTTPConnection(self.host, self.port, timeout=10)
        headers = {"Content-type": "text/xml"}
        connection.request("POST", "/", body=xml_body.encode("utf-8"), headers=headers)
        response = connection.getresponse()
        payload = response.read().decode("utf-8")
        connection.close()
        if response.status != 200:
            raise RuntimeError(f"Tally responded with HTTP {response.status}: {payload}")
        return payload

    def fetch_daybook(self, start: date, end: date) -> List[Voucher]:
        """Fetch Day Book vouchers between two dates (inclusive)."""

        # Tally uses dd-mm-yyyy format in XML requests.
        start_str = start.strftime("%d-%m-%Y")
        end_str = end.strftime("%d-%m-%Y")

        request_xml = f"""
            <ENVELOPE>
                <HEADER>
                    <VERSION>1</VERSION>
                    <TALLYREQUEST>Export Data</TALLYREQUEST>
                    <TYPE>Data</TYPE>
                    <ID>Day Book</ID>
                </HEADER>
                <BODY>
                    <DESC>
                        <STATICVARIABLES>
                            <SVFROMDATE>{start_str}</SVFROMDATE>
                            <SVTODATE>{end_str}</SVTODATE>
                        </STATICVARIABLES>
                        <TDL>
                            <TDLMESSAGE>
                                <REPORT NAME="Day Book">
                                    <FORMS>Day Book</FORMS>
                                </REPORT>
                                <FORM NAME="Day Book">
                                    <TOPPARTS>DBPart</TOPPARTS>
                                    <XMLTAG>DayBook</XMLTAG>
                                </FORM>
                                <PART NAME="DBPart">
                                    <LINES>DBLine</LINES>
                                </PART>
                                <LINE NAME="DBLine">
                                    <LEFTFIELDS>VoucherTypeName,Date,LedgerEntries</LEFTFIELDS>
                                </LINE>
                            </TDLMESSAGE>
                        </TDL>
                    </DESC>
                </BODY>
            </ENVELOPE>
        """

        response_xml = self._post_xml(request_xml)
        return list(_parse_daybook_response(response_xml))


def _parse_daybook_response(response_xml: str) -> Iterable[Voucher]:
    """Parse the minimal Day Book XML produced by Tally."""

    root = ET.fromstring(response_xml)
    for voucher_element in root.iterfind(".//VOUCHER"):
        voucher_type = voucher_element.get("VCHTYPE", "")
        date_text = voucher_element.findtext("DATE", default="")
        ledger_entries = []

        for ledger_element in voucher_element.findall("ALLLEDGERENTRIES.LIST"):
            ledger_name = ledger_element.findtext("LEDGERNAME", default="")
            amount_text = ledger_element.findtext("AMOUNT", default="0")
            is_debit = ledger_element.findtext("ISDEEMEDPOSITIVE", default="No") == "No"
            ledger_entries.append(
                LedgerEntry(
                    ledger_name=ledger_name,
                    amount=abs(float(amount_text)),
                    is_debit=is_debit,
                )
            )

        voucher_date = date.fromisoformat(
            f"{date_text[:4]}-{date_text[4:6]}-{date_text[6:]}"
        ) if date_text else date.today()
        yield Voucher(voucher_type=voucher_type, date=voucher_date, ledger_entries=ledger_entries)

