from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from lxml import etree
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

RFC_RECEPTOR_ESPERADO = "MES2301274X9"
TARGET_YEAR = 2026

MESES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

MESES_INV = {v: k.capitalize() for k, v in MESES.items()}

YELLOW = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
RED = PatternFill(start_color="EF9A9A", end_color="EF9A9A", fill_type="solid")

COLUMNS = [
    "Fecha",
    "Proveedor",
    "Proveedor RFC",
    "Folio Factura",
    "UUID",
    "Concepto",
    "Importe",
    "IVA",
    "Otros Impuestos",
    "Total",
    "Comentarios",
    "Empleado",
]

CATEGORY_EXACT = {
    "15101514": "GASOLINA",
    "15101515": "GASOLINA",
    "95111602": "PEAJE",
    "95111603": "PEAJE",
    "78111807": "ESTACIONAMIENTO",
    "90111800": "HOTEL",
    "90111500": "HOTEL / HOSPEDAJE",
    "90101500": "ALIMENTO / BEBIDA",
    "90101501": "ALIMENTO / BEBIDA",
    "90101503": "ALIMENTO / BEBIDA",
    "90101700": "ALIMENTO / BEBIDA",
    "90101800": "ALIMENTO / BEBIDA",
    "78111804": "TAXI",
    "78111800": "TRANSPORTE",
    "78111808": "ALQUILER DE AUTO",
    "78111811": "ALQUILER DE AUTO",
    "83111603": "DATOS MÓVILES",
    "43201415": "DATOS MÓVILES",
    "27113300": "HERRAMIENTAS",
    "27131500": "HERRAMIENTAS",
    "23291900": "HERRAMIENTAS INDUSTRIALES",
    "14111828": "PAPELERÍA",
}

CATEGORY_PREFIX = {
    "831116": "DATOS MÓVILES",
    "2711": "HERRAMIENTAS",
    "2713151": "HERRAMIENTAS",
    "141115": "PAPELERÍA",
    "441217": "PAPELERÍA",
}


@dataclass
class InvoiceRow:
    fecha: datetime
    proveedor: str
    proveedor_rfc: str
    folio_factura: str
    uuid: str
    concepto: str
    importe: Decimal
    iva: Decimal
    otros_impuestos: Decimal
    total: Decimal
    comentarios: str
    empleado: str
    source_month: str
    source_file: Path


class Processor:
    def __init__(self, base_2026: Path, excel_path: Path, logger):
        self.base_2026 = base_2026
        self.excel_path = excel_path
        self.logger = logger

    def run(self, dry_run: bool = False) -> dict:
        wb = load_workbook(self.excel_path)
        existing_uuid_positions = self._collect_uuid_positions(wb)
        new_uuid_positions: dict[str, list[tuple[str, int]]] = {}
        inserted = 0
        warnings = 0
        errors = 0

        for month_name, month_num in MESES.items():
            month_folder = self.base_2026 / month_name.capitalize()
            if not month_folder.exists() or not month_folder.is_dir():
                continue

            for employee_folder in sorted([p for p in month_folder.iterdir() if p.is_dir()], key=lambda p: p.name.lower()):
                for xml_file in sorted(employee_folder.glob("*.xml")):
                    try:
                        row = self._parse_invoice(xml_file, employee_folder.name, month_name)
                    except Exception as exc:
                        errors += 1
                        self.logger(f"ERROR parseando {xml_file}: {exc}")
                        continue

                    if row.fecha.year != TARGET_YEAR:
                        errors += 1
                        self.logger(f"ERROR fecha fuera de 2026 ({row.fecha.date()}) en {xml_file}")
                        continue

                    if not row.uuid:
                        errors += 1
                        self.logger(f"ERROR UUID vacío en {xml_file}")
                        continue

                    target_sheet = f"{MESES_INV[row.fecha.month]} {TARGET_YEAR}"
                    if target_sheet not in wb.sheetnames:
                        ws = wb.create_sheet(target_sheet)
                        self._ensure_headers(ws)
                    ws = wb[target_sheet]
                    self._ensure_headers(ws)

                    if not dry_run:
                        ws.append(
                            [
                                row.fecha,
                                row.proveedor,
                                row.proveedor_rfc,
                                row.folio_factura,
                                row.uuid,
                                row.concepto,
                                float(row.importe),
                                float(row.iva),
                                float(row.otros_impuestos),
                                float(row.total),
                                row.comentarios,
                                row.empleado,
                            ]
                        )
                        inserted_row = ws.max_row
                        inserted += 1

                        if row.source_month.lower() != MESES_INV[row.fecha.month].lower():
                            warnings += 1
                            self._fill_row(ws, inserted_row, YELLOW)
                            self.logger(
                                f"ADVERTENCIA mes distinto: {xml_file.name} en carpeta {row.source_month} pero fecha {row.fecha.date()}"
                            )

                        new_uuid_positions.setdefault(row.uuid, []).append((ws.title, inserted_row))

        if not dry_run:
            self._apply_duplicates(wb, existing_uuid_positions, new_uuid_positions)
            self._sort_all_month_sheets(wb)
            wb.save(self.excel_path)

        duplicate_count = sum(1 for u, pos in new_uuid_positions.items() if len(pos) + len(existing_uuid_positions.get(u, [])) > 1)
        warnings += duplicate_count
        if duplicate_count:
            self.logger(f"ADVERTENCIA UUID duplicados detectados: {duplicate_count}")

        return {
            "inserted": inserted,
            "warnings": warnings,
            "errors": errors,
            "duplicates": duplicate_count,
            "dry_run": dry_run,
        }

    def _collect_uuid_positions(self, wb):
        result: dict[str, list[tuple[str, int]]] = {}
        for name in wb.sheetnames:
            ws = wb[name]
            if ws.max_row < 2:
                continue
            headers = [ws.cell(1, i + 1).value for i in range(len(COLUMNS))]
            if headers[: len(COLUMNS)] != COLUMNS:
                continue
            for r in range(2, ws.max_row + 1):
                uuid_val = ws.cell(r, 5).value
                if uuid_val:
                    result.setdefault(str(uuid_val), []).append((name, r))
        return result

    def _apply_duplicates(self, wb, existing, new):
        for uuid_val, new_positions in new.items():
            positions = list(existing.get(uuid_val, [])) + new_positions
            if len(positions) > 1:
                for sheet, row in positions:
                    self._fill_row(wb[sheet], row, RED)

    def _fill_row(self, ws, row_number: int, fill):
        for c in range(1, len(COLUMNS) + 1):
            ws.cell(row=row_number, column=c).fill = fill

    def _ensure_headers(self, ws):
        for idx, col in enumerate(COLUMNS, start=1):
            ws.cell(row=1, column=idx, value=col)

    def _sort_all_month_sheets(self, wb):
        for name in wb.sheetnames:
            if not name.endswith(f" {TARGET_YEAR}"):
                continue

            ws = wb[name]
            if ws.max_row < 3:
                continue

            rows_to_sort: list[tuple[str, datetime, list, str | None]] = []
            for row_number in range(2, ws.max_row + 1):
                row_values = [ws.cell(row_number, col).value for col in range(1, len(COLUMNS) + 1)]
                row_highlight = self._detect_row_highlight(ws, row_number)
                empleado_key = str(row_values[11] or "").lower()
                fecha_dt = self._coerce_datetime(row_values[0], sheet_name=name, row_number=row_number)
                if fecha_dt is None:
                    # Mantener la fila pero enviarla al final para no romper el procesamiento.
                    fecha_dt = datetime.max

                rows_to_sort.append((empleado_key, fecha_dt, row_values, row_highlight))

            rows_to_sort.sort(key=lambda item: (item[0], item[1]))

            for row_number in range(2, ws.max_row + 1):
                for col in range(1, len(COLUMNS) + 1):
                    ws.cell(row_number, col).value = None
                    ws.cell(row_number, col).fill = PatternFill(fill_type=None)

            for target_row, (_, _, row_values, row_highlight) in enumerate(rows_to_sort, start=2):
                for col in range(1, len(COLUMNS) + 1):
                    ws.cell(target_row, col).value = row_values[col - 1]

                if row_highlight == "YELLOW":
                    self._fill_row(ws, target_row, YELLOW)
                elif row_highlight == "RED":
                    self._fill_row(ws, target_row, RED)

    def _detect_row_highlight(self, ws, row_number: int) -> str | None:
        """Detecta si la fila estaba marcada en amarillo o rojo antes de reordenar."""
        fill = ws.cell(row=row_number, column=1).fill
        color = getattr(fill, "start_color", None)
        rgb = (getattr(color, "rgb", None) or "").upper()
        index = (getattr(color, "index", None) or "").upper()

        if "EF9A9A" in rgb or "EF9A9A" in index:
            return "RED"
        if "FFF59D" in rgb or "FFF59D" in index:
            return "YELLOW"
        return None

    def _coerce_datetime(self, value, sheet_name: str, row_number: int):
        if isinstance(value, datetime):
            return value
        if value in (None, ""):
            self.logger(
                f"ADVERTENCIA fecha vacía en sheet {sheet_name} fila {row_number}; se mantiene la fila y se ordena al final."
            )
            return None

        try:
            return datetime.fromisoformat(str(value))
        except ValueError:
            self.logger(
                f"ADVERTENCIA fecha inválida '{value}' en sheet {sheet_name} fila {row_number}; se mantiene la fila y se ordena al final."
            )
            return None

    def _parse_invoice(self, xml_file: Path, empleado: str, source_month: str) -> InvoiceRow:
        doc = etree.parse(str(xml_file))
        root = doc.getroot()
        ns = {
            "cfdi": "http://www.sat.gob.mx/cfd/4",
            "tfd": "http://www.sat.gob.mx/TimbreFiscalDigital",
        }

        fecha_str = root.get("Fecha", "")
        fecha = datetime.fromisoformat(fecha_str.replace("Z", "+00:00"))

        emisor = root.find("cfdi:Emisor", namespaces=ns)
        receptor = root.find("cfdi:Receptor", namespaces=ns)

        proveedor = emisor.get("Nombre", "") if emisor is not None else ""
        proveedor_rfc = emisor.get("Rfc", "") if emisor is not None else ""

        receptor_rfc = receptor.get("Rfc", "") if receptor is not None else ""
        if receptor_rfc != RFC_RECEPTOR_ESPERADO:
            raise ValueError(f"RFC receptor inválido: {receptor_rfc}")

        complemento = root.find("cfdi:Complemento", namespaces=ns)
        uuid = ""
        if complemento is not None:
            tfd = complemento.find("tfd:TimbreFiscalDigital", namespaces=ns)
            if tfd is not None:
                uuid = tfd.get("UUID", "")

        serie = root.get("Serie", "")
        folio = root.get("Folio", "")
        folio_factura = f"{serie}-{folio}" if serie and folio else (folio or serie)

        subtotal = Decimal(root.get("SubTotal", "0"))
        total = Decimal(root.get("Total", "0"))

        iva = Decimal("0")
        otros = Decimal("0")

        impuestos = root.find("cfdi:Impuestos", namespaces=ns)
        if impuestos is not None:
            traslados = impuestos.findall("cfdi:Traslados/cfdi:Traslado", namespaces=ns)
            retenciones = impuestos.findall("cfdi:Retenciones/cfdi:Retencion", namespaces=ns)
            for tax in list(traslados) + list(retenciones):
                imp_code = tax.get("Impuesto", "")
                amount = Decimal(tax.get("Importe", "0"))
                if imp_code == "002":
                    iva += amount
                else:
                    otros += amount

        concepto_code = self._first_concept_code(root, ns)
        concepto = self._map_category(concepto_code)

        return InvoiceRow(
            fecha=fecha,
            proveedor=proveedor,
            proveedor_rfc=proveedor_rfc,
            folio_factura=folio_factura,
            uuid=uuid,
            concepto=concepto,
            importe=subtotal,
            iva=iva,
            otros_impuestos=otros,
            total=total,
            comentarios="",
            empleado=empleado,
            source_month=source_month,
            source_file=xml_file,
        )

    def _first_concept_code(self, root, ns) -> str:
        conceptos = root.findall("cfdi:Conceptos/cfdi:Concepto", namespaces=ns)
        if not conceptos:
            return ""
        return conceptos[0].get("ClaveProdServ", "")

    def _map_category(self, clave: str) -> str:
        if not clave:
            return ""
        if clave in CATEGORY_EXACT:
            return CATEGORY_EXACT[clave]
        for prefix, category in CATEGORY_PREFIX.items():
            if clave.startswith(prefix):
                return category
        return clave


def discover_month_folders(base_2026: Path) -> Iterable[Path]:
    for p in sorted(base_2026.iterdir(), key=lambda x: x.name.lower()):
        if p.is_dir() and p.name.lower() in MESES:
            yield p
