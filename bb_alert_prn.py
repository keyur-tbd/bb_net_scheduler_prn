"""BigBasket's WMS depots mail PRNs in a new format; turn them into the old one.

Since 22 Sep 2026 BigBasket has been moving depots to its WMS. A moved depot stops sending
"Auto PRN Details" from bbnet2@bigbasket.com (GonNo, ShipToCode, ... with the header on row 3)
and instead sends two mails per PRN from alerts@bigbasket.com: "PRN Initiated" when the stock is
set aside and "PRN handovered" when our person collects it. Each carries a
supplier_return_<PRN>_<id>.xlsx with different columns, no ship-to code, and a stylesheet that
openpyxl refuses to load.

Only the handover file is read, so a PRN is never counted twice; it holds the quantities actually
collected. The file is rewritten here in the old layout before it reaches Drive, so the Excel step
in app.py, supabase_sink.py, the bbalertprn tab and the marketing team's Mail tab all keep working
unchanged.

Reads the workbook's XML directly. Standard library plus openpyxl for writing.
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook

# WMS depot -> the (ShipToCode, ShipToName) that depot used on its old-format mails. The marketing
# team's sheet keys every PRN on ShipToCode + PRN and looks the warehouse up from ShipToName, so
# these must stay the old spellings, "Hyderbad" included. Bangalore was matched on the address.
FC_DEPOT = {
    'WMS-Ahmedabad-FC': ('BBAHM001', 'AHMEDABAD SERVER'),
    'WMS-Bangalore-FC': ('BBBLR001', 'BIG BASKET WAREHOUSE'),
    'WMS-Bangalore-FC2': ('BBBLR005', 'Bangalore DC5'),
    'WMS-Chennai-FC': ('BBCHN003', 'Chennai DC3'),
    'WMS-Hyderabad-FC': ('BBHYD002', 'Hyderbad FNV'),
    'WMS-Kolkata-FC': ('BBKOL003', 'Kolkata DC3'),
    'WMS-Kundli-FC': ('BBKUN001', 'Kundli DC'),
    'WMS-Noida-FC': ('BBNOI002', 'Noida DC2'),
    'WMS-Pune-FC': ('BBPUN002', 'Pune DC2'),
}

OLD_HEADER = ['GonNo', 'SupplierCode', 'SupplierName', 'ShipToCode', 'ShipToName', 'ShipTOAddress',
              'InvoiceNo', 'InvoiceDate', 'HSN Code', 'Skucode', 'SkuDesc', 'MRP', 'SGST%', 'SGST Value',
              'CGST%', 'CGST Value', 'IGST%', 'IGST Value', 'GST%', 'GST Amount', 'Cess%', 'Cess Value',
              'Additional Cess Per piece', 'Total additional Cess', 'LandingCost', 'Quantity', 'TotalValue']

_REQUIRED = ['PRN No', 'PRN Timestamp', 'Supplier Code', 'Supplier Name', 'Fc Name', 'Fc Address', 'Sku Code',
             'Sku Description', 'HSN Code', 'Quantity', 'Cost Price (Incl. Tax)', 'MRP', 'SGST %', 'CGST %',
             'IGST %', 'CESS %', 'CESS per Unit', 'SGST Value', 'CGST Value', 'IGST Value', 'Cess Value',
             'CESS per Unit Value', 'Tax Value', 'Total PRN Value (Incl. Tax)']

_NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
       'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
_M = '{%s}' % _NS['m']


class AlertFileError(ValueError):
    """The file is not the layout this module knows; better loud than loaded wrong."""


def _column(ref: str) -> int:
    n = 0
    for ch in re.match(r'[A-Z]+', ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_grid(data: bytes) -> List[List[Any]]:
    """Every row of the first sheet as a list of cell strings (None for an empty cell)."""
    z = zipfile.ZipFile(io.BytesIO(data))
    shared: List[str] = []
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', _NS):
            shared.append(''.join(t.text or '' for t in si.iter(_M + 't')))
    sheet = ET.fromstring(z.read('xl/workbook.xml')).find('m:sheets', _NS)[0]
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    target = rels[sheet.get('{%s}id' % _NS['r'])].lstrip('/')
    target = target if target.startswith('xl/') else 'xl/' + target
    grid = []
    for row in ET.fromstring(z.read(target)).iter(_M + 'row'):
        cells: Dict[int, Any] = {}
        for c in row.findall('m:c', _NS):
            v, inline = c.find('m:v', _NS), c.find('m:is', _NS)
            if inline is not None:
                value = ''.join(t.text or '' for t in inline.iter(_M + 't'))
            elif v is None:
                value = None
            elif c.get('t') == 's':
                value = shared[int(v.text)]
            else:
                value = v.text
            cells[_column(c.get('r'))] = value
        grid.append([cells.get(i) for i in range(max(cells) + 1)] if cells else [])
    return grid


def _number(text: Any) -> float:
    try:
        return float(text) if text not in (None, '') else 0.0
    except (TypeError, ValueError):
        raise AlertFileError(f'not a number: {text!r}')


def _integer(text: Any) -> Any:
    """Codes as whole numbers, as they arrive on the old files; anything else stays as it came."""
    text = str(text or '').strip()
    return int(text) if text.isdigit() else text


def parse_timestamp(text: Any) -> date:
    return datetime.strptime(str(text).strip(), '%d-%m-%Y %I:%M:%S %p').date()


def parse_handover(data: bytes) -> Dict[str, Any]:
    """Read a supplier_return file: the lines by SKU, repeats summed, zero-quantity lines dropped."""
    grid = read_grid(data)
    if not grid:
        raise AlertFileError('empty sheet')
    head = [str(h or '').strip() for h in grid[0]]
    missing = [h for h in _REQUIRED if h not in head]
    if missing:
        raise AlertFileError(f'columns missing: {missing}')
    at = {h: head.index(h) for h in _REQUIRED}

    def get(row, name):
        i = at[name]
        return row[i] if i < len(row) else None

    rows = [r for r in grid[1:] if r and re.fullmatch(r'\d+', str(r[0] or '').strip())]
    if not rows:
        raise AlertFileError('no data rows')
    if len({str(get(r, 'PRN No')) for r in rows}) != 1:
        raise AlertFileError('more than one PRN in the file')

    # The file's own Total row is the check that no line was misread.
    file_total = None
    for r in grid[1:]:
        if r and any(str(x).strip() == 'Total:' for x in r if x is not None):
            file_total = _number(get(r, 'Total PRN Value (Incl. Tax)'))
    value_read = sum(_number(get(r, 'Total PRN Value (Incl. Tax)')) for r in rows)
    if file_total is not None and abs(file_total - value_read) > 1:
        raise AlertFileError(f"lines sum to {value_read:.2f} but the file's Total says {file_total:.2f}")

    first = rows[0]
    lines: Dict[str, Dict[str, Any]] = {}
    for r in rows:
        qty = _number(get(r, 'Quantity'))
        if qty == 0:                          # initiated but not collected
            continue
        sku = str(get(r, 'Sku Code')).strip()
        line = lines.setdefault(sku, dict(
            sku=sku, desc=str(get(r, 'Sku Description') or '').strip(), hsn=get(r, 'HSN Code'),
            mrp=_number(get(r, 'MRP')), sgst_pct=_number(get(r, 'SGST %')), cgst_pct=_number(get(r, 'CGST %')),
            igst_pct=_number(get(r, 'IGST %')), cess_pct=_number(get(r, 'CESS %')),
            cess_unit=_number(get(r, 'CESS per Unit')), qty=0.0, value=0.0, sgst=0.0, cgst=0.0, igst=0.0,
            tax=0.0, cess=0.0, cess_total=0.0))
        line['qty'] += qty
        line['value'] += _number(get(r, 'Total PRN Value (Incl. Tax)'))
        line['sgst'] += _number(get(r, 'SGST Value'))
        line['cgst'] += _number(get(r, 'CGST Value'))
        line['igst'] += _number(get(r, 'IGST Value'))
        line['tax'] += _number(get(r, 'Tax Value'))
        line['cess'] += _number(get(r, 'Cess Value'))
        line['cess_total'] += _number(get(r, 'CESS per Unit Value'))
    return dict(prn=str(get(first, 'PRN No')).strip(), file_date=parse_timestamp(get(first, 'PRN Timestamp')),
                supplier_code=get(first, 'Supplier Code'), supplier_name=str(get(first, 'Supplier Name') or '').strip(),
                fc=str(get(first, 'Fc Name') or '').strip(), fc_address=str(get(first, 'Fc Address') or '').strip(),
                lines=list(lines.values()), value=sum(x['value'] for x in lines.values()),
                qty=sum(x['qty'] for x in lines.values()))


def depot_of(fc: str) -> Tuple[str, str, bool]:
    """(ShipToCode, ShipToName, known). An unknown depot keeps its WMS name, so its rows still load
    and show up as odd rather than vanish: a depot going missing is how this module came to exist."""
    if fc in FC_DEPOT:
        return FC_DEPOT[fc][0], FC_DEPOT[fc][1], True
    return fc, fc, False


def old_layout_rows(parsed: Dict[str, Any], prn_date: Optional[date], handover_date: date) -> List[List[Any]]:
    """The file's lines as rows under OLD_HEADER. prn_date is the Initiated date when known: the
    handover file's own PRN Timestamp moves when it is collected, which can cross a month end."""
    code, name, _ = depot_of(parsed['fc'])
    when = prn_date or parsed['file_date']
    out = []
    for x in parsed['lines']:
        out.append([_integer(parsed['prn']), _integer(parsed['supplier_code']), parsed['supplier_name'], code, name,
                    parsed['fc_address'], handover_date.strftime('%d/%m/%Y'), datetime(when.year, when.month, when.day),
                    _integer(x['hsn']), _integer(x['sku']), x['desc'], x['mrp'], x['sgst_pct'], round(x['sgst'], 4),
                    x['cgst_pct'], round(x['cgst'], 4), x['igst_pct'], round(x['igst'], 4),
                    x['sgst_pct'] + x['cgst_pct'] + x['igst_pct'], round(x['tax'], 4), x['cess_pct'], round(x['cess'], 4),
                    x['cess_unit'], round(x['cess_total'], 4), round(x['value'] / x['qty'], 4), x['qty'],
                    round(x['value'], 4)])
    return out


def to_old_layout(parsed: Dict[str, Any], prn_date: Optional[date], handover_date: date) -> bytes:
    """An .xlsx laid out like an old Auto PRN file: two title rows, then the header on row 3."""
    wb = Workbook()
    ws = wb.active
    ws.title = f"{parsed['prn']}-PRN"
    ws.cell(row=1, column=13, value='AUTO PRN Details (converted from WMS handover mail)')
    ws.cell(row=2, column=13, value='Date: ' + (prn_date or parsed['file_date']).strftime('%d-%b-%Y'))
    for i, h in enumerate(OLD_HEADER, 1):
        ws.cell(row=3, column=i, value=h)
    for r, row in enumerate(old_layout_rows(parsed, prn_date, handover_date), 4):
        for c, value in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=value)
            if isinstance(value, datetime):
                cell.number_format = 'DD-MMM-YYYY'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
