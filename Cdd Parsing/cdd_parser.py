"""Parse a Vector CANdela Studio CDD file and export its diagnostic data to Excel.

The workbook has four sheets:
    Service  every diagnostic service of every variant, with request bytes and access rights
    DID      every data identifier, its length, data elements and read/write access
    DTC      every DTC of every fault memory, with the DTC record fields
    RID      every routine identifier with its sub-functions and parameters

Usage:
    python cdd_parser.py <file.cdd> [-o <output.xlsx>]
"""

import argparse
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROUTINE_CONTROL_SID = 0x31
READ_DID_SID = 0x22
WRITE_DID_SID = 0x2E


def text(elem, tag="NAME"):
    """Return the (first language) text of a CANdela text element such as NAME or DESC."""
    if elem is None:
        return ""
    node = elem.find(tag) if tag else elem
    if node is None:
        return ""
    tuv = node.find("TUV")
    value = tuv.text if tuv is not None else node.text
    return (value or "").strip()


def to_hex(value, n_bytes):
    return f"{value:0{n_bytes * 2}X}"


def dtc_code(number):
    """Convert a 3-byte UDS DTC number into its SAE J2012 form, e.g. 0xC41881 -> U0418-81."""
    high = (number >> 16) & 0xFF
    letter = "PCBU"[high >> 6]
    return f"{letter}{(high >> 4) & 0x3}{high & 0xF:X}{(number >> 8) & 0xFF:02X}-{number & 0xFF:02X}"


class CddFile:
    """In-memory view of a CDD file with an index of every element that has an id."""

    def __init__(self, path):
        self.path = Path(path)
        self.tree = ET.parse(self.path)
        self.root = self.tree.getroot()
        self.doc = self.root.find("ECUDOC")
        self.ids = {e.get("id"): e for e in self.root.iter() if e.get("id")}
        self.parent = {child: parent for parent in self.root.iter() for child in parent}
        # mayBeExec="(1,3,4,6)" counts the states of all state groups in document order.
        self.states = []
        for group in self.doc.iter("STATEGROUP"):
            for state in group.findall("STATE"):
                self.states.append((group.get("spec"), text(state)))

    # ---------- generic helpers ----------

    def ref(self, elem, attr):
        return self.ids.get(elem.get(attr)) if elem is not None else None

    def bit_length(self, dtref):
        """Return (min_bits, max_bits) of a data type, or (None, None) if unknown."""
        dtype = self.ids.get(dtref)
        cvt = dtype.find("CVALUETYPE") if dtype is not None else None
        if cvt is None:
            return None, None
        bl = int(cvt.get("bl", 0))
        if cvt.get("qty") == "field":
            return bl * int(cvt.get("minsz", 1)), bl * int(cvt.get("maxsz", 1))
        return bl, bl

    def access(self, service):
        """Return (sessions, security levels) in which a service may be executed."""
        may_be_exec = service.get("mayBeExec")
        template = self.ref(service, "tmplref")
        if may_be_exec is None and template is not None:
            may_be_exec = template.get("mayBeExec")
        if may_be_exec is None:
            return "All", "All"
        indexes = [int(i) for i in may_be_exec.strip("()").split(",") if i]
        picked = [self.states[i - 1] for i in indexes if 0 < i <= len(self.states)]
        sessions = ", ".join(name for spec, name in picked if spec == "session")
        security = ", ".join(name for spec, name in picked if spec == "security")
        return sessions or "-", security or "-"

    def protocol_service(self, service):
        """SERVICE -> DCLSRVTMPL -> PROTOCOLSERVICE."""
        return self.ref(self.ref(service, "tmplref"), "tmplref")

    def request_bytes(self, diaginst, service):
        """Build the constant leading bytes of the request (SID, sub-function, identifier)."""
        proto = self.protocol_service(service)
        req = proto.find("REQ") if proto is not None else None
        if req is None:
            return None, ""
        static_values = {}
        for sv in diaginst.findall("STATICVALUE"):
            shstatic = self.ref(sv, "shstaticref")
            for comp_ref in shstatic.findall("STATICCOMPREF") if shstatic is not None else []:
                static_values[comp_ref.get("idref")] = int(sv.get("v"))
        sid = None
        parts = []
        for comp in req:
            if comp.tag == "CONSTCOMP":
                n_bytes = max(1, int(comp.get("bl", 8)) // 8)
                value = int(comp.get("v"))
                if comp.get("spec") == "sid" and sid is None:
                    sid = value
                parts.append(to_hex(value, n_bytes))
            elif comp.tag == "STATICCOMP":
                bits, _ = self.bit_length(comp.get("dtref"))
                n_bytes = max(1, (bits or 8) // 8)
                value = static_values.get(comp.get("id"))
                parts.append(to_hex(value, n_bytes) if value is not None else "XX" * n_bytes)
            elif comp.tag in ("NAME", "QUAL", "DESC"):
                continue
            else:
                break  # first variable data component ends the constant prefix
        return sid, " ".join(parts)

    # ---------- structure walking ----------

    def variants(self):
        return self.doc.find("ECU").findall("VAR")

    def diag_instances(self, variant):
        """Yield (diag class name, DIAGINST) for instances defined in or referenced by a variant."""
        for child in variant:
            if child.tag == "DIAGCLASS":
                for inst in child.findall("DIAGINST"):
                    yield text(child), inst
            elif child.tag == "DIAGINST":
                yield text(self.ref(child, "tmplref")), child
            elif child.tag == "DIAGINSTREF":
                inst = self.ref(child, "idref")
                if inst is not None:
                    yield text(self.ref(inst, "tmplref")), inst

    def parameters(self, diaginst):
        """Return 'Proxy name: element, element' for each data block of an instance."""
        out = []
        for cont in diaginst.findall("SIMPLECOMPCONT"):
            names = [text(e) for e in cont.iter("DATAOBJ")]
            if names:
                out.append(f"{text(self.ref(cont, 'shproxyref'))}: {', '.join(names)}")
        return "\n".join(out)

    # ---------- sheet builders ----------

    def services(self):
        rows = []
        for variant in self.variants():
            for diag_class, inst in self.diag_instances(variant):
                for service in inst.findall("SERVICE"):
                    proto = self.protocol_service(service)
                    sid, request = self.request_bytes(inst, service)
                    sessions, security = self.access(service)
                    addressing = [a for a, flag in (("Physical", "phys"), ("Functional", "func")) if service.get(flag) == "1"]
                    rows.append({
                        "Variant": text(variant),
                        "Diag Class": diag_class,
                        "Diag Instance": text(inst),
                        "Service": text(service),
                        "Qualifier": service.findtext("SHORTCUTQUAL") or service.findtext("QUAL") or "",
                        "SID": to_hex(sid, 1) if sid is not None else "",
                        "Protocol Service": text(proto),
                        "Request": request,
                        "Addressing": ", ".join(addressing),
                        "Sessions": sessions,
                        "Security": security,
                        "Description": text(inst, "DESC"),
                    })
        return rows

    def dids(self):
        # Which variants list each DID, and which services read or write it.
        did_variants = {}
        for variant in self.variants():
            for ref in variant.iter("DIDREF"):
                did_variants.setdefault(ref.get("didRef"), []).append(text(variant))
        access = {}  # (variant, DID number, SID) -> (sessions, security)
        for variant in self.variants():
            for _, inst in self.diag_instances(variant):
                sv = inst.find("STATICVALUE")
                if sv is None:
                    continue
                for service in inst.findall("SERVICE"):
                    sid, _ = self.request_bytes(inst, service)
                    if sid in (READ_DID_SID, WRITE_DID_SID):
                        access.setdefault((text(variant), int(sv.get("v")), sid), self.access(service))

        rows = []
        for did in self.doc.find("DIDS").findall("DID"):
            number = int(did.get("n"))
            elements, min_bits, max_bits = [], 0, 0
            for obj in did.find("STRUCTURE"):
                lo, hi = self.bit_length(obj.get("dtref")) if obj.get("dtref") else (int(obj.get("bl", 0)),) * 2
                if lo is None:
                    min_bits = max_bits = None
                elif min_bits is not None:
                    min_bits, max_bits = min_bits + lo, max_bits + hi
                elements.append(f"{text(obj)} [{lo if lo == hi else f'{lo}..{hi}'} bit]")
            if min_bits is None:
                length = "?"
            elif min_bits == max_bits:
                length = f"{min_bits // 8}" if min_bits % 8 == 0 else f"{min_bits} bit"
            else:
                length = f"{min_bits // 8}..{max_bits // 8}"
            variants = did_variants.get(did.get("id"), [])
            read = next((access[k] for k in ((v, number, READ_DID_SID) for v in variants) if k in access), None)
            write = next((access[k] for k in ((v, number, WRITE_DID_SID) for v in variants) if k in access), None)
            rows.append({
                "DID": to_hex(number, 2),
                "Name": text(did),
                "Qualifier": did.findtext("QUAL") or "",
                "Variants": ", ".join(variants),
                "Length (bytes)": length,
                "Data Elements": "\n".join(elements),
                "Read (22)": "Yes" if read else "No",
                "Read Sessions": read[0] if read else "",
                "Read Security": read[1] if read else "",
                "Write (2E)": "Yes" if write else "No",
                "Write Sessions": write[0] if write else "",
                "Write Security": write[1] if write else "",
            })
        return rows

    def dtcs(self):
        # Keep the record fields in the order of the DTC record templates.
        fields = [text(item) for tmpl in self.doc.iter("RECORDTMPL") for item in tmpl if item.get("id")]
        rows = []
        for record_dt in self.root.iter("RECORDDT"):
            records = record_dt.findall("RECORD")
            if not records:
                continue
            owner = self.parent.get(record_dt)
            variant, memory = "Record pool (no variant)", ""
            node = owner
            while node is not None:
                if node.tag == "DIAGINST":
                    memory = text(node)
                elif node.tag == "VAR":
                    variant = text(node)
                    break
                node = self.parent.get(node)
            for record in records:
                number = int(record.get("v"))
                row = {
                    "Variant": variant,
                    "Fault Memory": memory,
                    "DTC": to_hex(number, 3),
                    "DTC Code": dtc_code(number),
                    "Description": text(record, "TEXT"),
                    **{field: "" for field in fields},
                }
                for item in record:
                    tmpl = self.ref(item, "idref")
                    if tmpl is None:
                        continue
                    column = text(tmpl)
                    if item.tag == "TRRECORDITEM":
                        row[column] = text(item, "TEXT")
                    elif item.tag == "ENUMRECORDITEM":
                        labels = {e.get("v"): text(e, None) for e in tmpl.findall("ETAG")}
                        row[column] = labels.get(item.get("v"), item.get("v"))
                    else:
                        row[column] = int(item.get("v")) if item.get("v", "").isdigit() else item.get("v")
                rows.append(row)
        return rows

    def rids(self):
        rows = []
        for variant in self.variants():
            for diag_class, inst in self.diag_instances(variant):
                services = inst.findall("SERVICE")
                sv = inst.find("STATICVALUE")
                if not services or sv is None:
                    continue
                requests = [(s, *self.request_bytes(inst, s)) for s in services]
                if not any(sid == ROUTINE_CONTROL_SID for _, sid, _ in requests):
                    continue
                sessions, security = self.access(services[0])
                rows.append({
                    "Variant": text(variant),
                    "RID": to_hex(int(sv.get("v")), 2),
                    "Name": text(inst),
                    "Qualifier": inst.findtext("QUAL") or "",
                    "Diag Class": diag_class,
                    "Sub-functions": ", ".join(text(s) for s in services),
                    "Requests": "\n".join(f"{text(s)}: {req}" for s, _, req in requests),
                    "Sessions": sessions,
                    "Security": security,
                    "Parameters": self.parameters(inst),
                    "Description": text(inst, "DESC"),
                })
        return rows


def write_sheet(workbook, title, rows):
    sheet = workbook.create_sheet(title)
    columns = []
    for row in rows:
        columns.extend(c for c in row if c not in columns)
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(c, "") for c in columns])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for index, column in enumerate(columns, start=1):
        longest = max([len(column)] + [max((len(line) for line in str(r.get(column, "")).split("\n")), default=0) for r in rows])
        sheet.column_dimensions[get_column_letter(index)].width = min(max(10, longest + 2), 60)
    sheet.freeze_panes = "A2"
    if rows:
        sheet.auto_filter.ref = sheet.dimensions


def export(cdd_path, xlsx_path):
    cdd = CddFile(cdd_path)
    sheets = {"Service": cdd.services(), "DID": cdd.dids(), "DTC": cdd.dtcs(), "RID": cdd.rids()}
    workbook = Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets.items():
        write_sheet(workbook, title, rows)
    workbook.save(xlsx_path)
    return {title: len(rows) for title, rows in sheets.items()}


def main(argv=None):
    parser = argparse.ArgumentParser(description="Export DIDs, DTCs, RIDs and services of a CDD file to Excel.")
    parser.add_argument("cdd", help="path to the .cdd file")
    parser.add_argument("-o", "--output", help="output .xlsx path (default: <cdd name>.xlsx next to the CDD)")
    args = parser.parse_args(argv)
    output = args.output or str(Path(args.cdd).with_suffix(".xlsx"))
    counts = export(args.cdd, output)
    print(f"Wrote {output}")
    for title, count in counts.items():
        print(f"  {title}: {count} rows")
    return 0


if __name__ == "__main__":
    sys.exit(main())
