from __future__ import annotations

from copy import copy
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

    @staticmethod
    def _normalize_uuid(value) -> str:
        return str(value or "").strip().upper().replace("{", "").replace("}", "")

    # ------------------------------------------------------------------
    # Main entry point
    # ------------------------------------------------------------------

    def run(self, dry_run: bool = False) -> dict:
        wb = load_workbook(self.excel_path)

        # Step 1: collect UUIDs already present in the workbook
        existing_uuid_positions = self._collect_uuid_positions(wb)
        existing_uuid_set = set(existing_uuid_positions.keys())

        # Step 2: track UUIDs inserted in this run
        new_uuid_positions: dict[str, list[tuple[str, int]]] = {}

        inserted = 0
        warnings = 0
        errors = 0

        # Step 3: iterate month folders → employee folders → XML files
        for month_name, month_num in MESES.items():
            month_folder = self.base_2026 / month_name.capitalize()
            if not month_folder.exists() or not month_folder.is_dir():
                continue

            employee_folders = sorted(
                [p for p in month_folder.iterdir() if p.is_dir()],
                key=lambda p: p.name.lower(),
            )

            for employee_folder in employee_folders:
                for xml_file in sorted(employee_folder.glob("*.xml")):
                    # --- parse ---
                    try:
                        row = self._parse_invoice(xml_file, employee_folder.name, month_name)
                    except Exception as exc:
                        errors += 1
                        self.logger(f"ERROR parseando {xml_file.name}: {exc}")
                        continue

                    # --- validate year ---
                    if row.fecha.year != TARGET_YEAR:
                        errors += 1
                        self.logger(
                            f"ERROR fecha fuera de {TARGET_YEAR} ({row.fecha.date()}) en {xml_file.name}"
                        )
                        continue

                    # --- validate UUID ---
                    if not row.uuid:
                        errors += 1
                        self.logger(f"ERROR UUID vacío en {xml_file.name}")
                        continue

                    uuid_key = self._normalize_uuid(row.uuid)

                    # --- skip if already in workbook ---
                    if uuid_key in existing_uuid_set:
                        self.logger(
                            f"OMITIDO UUID ya existente: {row.uuid} ({xml_file.name})"
                        )
                        continue

                    # --- determine target sheet ---
                    target_sheet = f"{MESES_INV[row.fecha.month]} {TARGET_YEAR}"
                    if target_sheet not in wb.sheetnames:
                        ws = wb.create_sheet(target_sheet)
                        self._ensure_headers(ws)

                    ws = wb[target_sheet]
                    self._ensure_headers(ws)

                    # --- insert row ---
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

                        # yellow highlight if invoice date doesn't match folder month
                        if row.source_month.lower() != MESES_INV[row.fecha.month].lower():
                            warnings += 1
                            self._fill_row(ws, inserted_row, YELLOW)
                            self.logger(
                                f"ADVERTENCIA mes distinto: {xml_file.name} "
                                f"en carpeta '{row.source_month}' pero fecha {row.fecha.date()}"
                            )

                        # track for duplicate detection
                        new_uuid_positions.setdefault(uuid_key, []).append(
                            (ws.title, inserted_row)
                        )
                        # mark as seen so we don't insert the same UUID twice
                        # even if two XML files in this run share the same UUID
                        existing_uuid_set.add(uuid_key)

        # Step 4: apply red highlight to any duplicated UUIDs
        if not dry_run:
            dup_count = self._apply_duplicates(wb, existing_uuid_positions, new_uuid_positions)
            warnings += dup_count
            if dup_count:
                self.logger(f"ADVERTENCIA UUIDs duplicados detectados y marcados en rojo: {dup_count}")

            # Step 5: sort each month sheet (employee asc, date asc)
            try:
                self._sort_all_month_sheets(wb)
            except Exception as exc:
                errors += 1
                self.logger(f"ERROR ordenando sheets: {exc}")

            wb.save(self.excel_path)

        return {
            "inserted": inserted,
            "warnings": warnings,
            "errors": errors,
            "dry_run": dry_run,
        }

    # ------------------------------------------------------------------
    # UUID helpers
    # ------------------------------------------------------------------

    def _collect_uuid_positions(self, wb) -> dict[str, list[tuple[str, int]]]:
        """Return a mapping of normalised UUID → [(sheet_name, row_number)] for every
        data row already in the workbook."""
        result: dict[str, list[tuple[str, int]]] = {}
        for name in wb.sheetnames:
            ws = wb[name]
            if ws.max_row < 2:
                continue
            headers = [ws.cell(1, i + 1).value for i in range(len(COLUMNS))]
            if headers[: len(COLUMNS)] != COLUMNS:
                continue
            for r in range(2, ws.max_row + 1):
                raw = ws.cell(r, 5).value
                if raw:
                    key = self._normalize_uuid(raw)
                    if key:
                        result.setdefault(key, []).append((name, r))
        return result

    def _apply_duplicates(
        self,
        wb,
        existing: dict[str, list[tuple[str, int]]],
        new: dict[str, list[tuple[str, int]]],
    ) -> int:
        """Highlight in red every row (old + new) that shares a UUID with another row."""
        dup_count = 0
        for uuid_key, new_positions in new.items():
            all_positions = list(existing.get(uuid_key, [])) + new_positions
            if len(all_positions) > 1:
                dup_count += 1
                for sheet_name, row_num in all_positions:
                    if sheet_name in wb.sheetnames:
                        self._fill_row(wb[sheet_name], row_num, RED)
        return dup_count

    # ------------------------------------------------------------------
    # Excel helpers
    # ------------------------------------------------------------------

    def _fill_row(self, ws, row_number: int, fill: PatternFill) -> None:
        for c in range(1, len(COLUMNS) + 1):
            ws.cell(row=row_number, column=c).fill = copy(fill)

    def _ensure_headers(self, ws) -> None:
        if ws.cell(1, 1).value != COLUMNS[0]:
            for idx, col in enumerate(COLUMNS, start=1):
                ws.cell(row=1, column=idx, value=col)
        # Always ensure autofilter covers the header row
        from openpyxl.utils import get_column_letter
        ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    def _detect_row_highlight(self, ws, row_number: int) -> str | None:
        rgb = (
            getattr(getattr(ws.cell(row=row_number, column=1).fill, "start_color", None), "rgb", "")
            or ""
        ).upper()
        if "EF9A9A" in rgb:
            return "RED"
        if "FFF59D" in rgb:
            return "YELLOW"
        return None

    def _sort_all_month_sheets(self, wb) -> None:
        for name in wb.sheetnames:
            if not name.endswith(f" {TARGET_YEAR}"):
                continue
            ws = wb[name]
            if ws.max_row < 3:
                continue

            # Read all data rows with their highlight state
            data: list[tuple[str, datetime, list, str | None]] = []
            for r in range(2, ws.max_row + 1):
                values = [ws.cell(r, c).value for c in range(1, len(COLUMNS) + 1)]
                highlight = self._detect_row_highlight(ws, r)
                empleado = str(values[11] or "").lower()
                fecha_dt = self._coerce_datetime(values[0], name, r)
                if fecha_dt is None:
                    fecha_dt = datetime.max
                data.append((empleado, fecha_dt, values, highlight))

            data.sort(key=lambda x: (x[0], x[1]))

            # Clear and rewrite
            for r in range(2, ws.max_row + 1):
                for c in range(1, len(COLUMNS) + 1):
                    ws.cell(r, c).value = None
                    ws.cell(r, c).fill = PatternFill(fill_type=None)

            for idx, (_, _, values, highlight) in enumerate(data, start=2):
                for c in range(1, len(COLUMNS) + 1):
                    ws.cell(idx, c).value = values[c - 1]
                if highlight == "YELLOW":
                    self._fill_row(ws, idx, YELLOW)
                elif highlight == "RED":
                    self._fill_row(ws, idx, RED)

    def _coerce_datetime(self, value, sheet_name: str, row_number: int) -> datetime | None:
        if isinstance(value, datetime):
            return value
        if value in (None, ""):
            self.logger(
                f"ADVERTENCIA fecha vacía en sheet '{sheet_name}' fila {row_number}; se ordena al final."
            )
            return None
        try:
            return datetime.fromisoformat(str(value))
        except ValueError:
            self.logger(
                f"ADVERTENCIA fecha inválida '{value}' en sheet '{sheet_name}' fila {row_number}; se ordena al final."
            )
            return None

    # ------------------------------------------------------------------
    # XML parsing
    # ------------------------------------------------------------------

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

        uuid = ""
        complemento = root.find("cfdi:Complemento", namespaces=ns)
        if complemento is not None:
            tfd_node = complemento.find("tfd:TimbreFiscalDigital", namespaces=ns)
            if tfd_node is not None:
                uuid = self._normalize_uuid(tfd_node.get("UUID", ""))

        serie = root.get("Serie", "")
        folio = root.get("Folio", "")
        folio_factura = f"{serie}-{folio}" if serie and folio else (folio or serie)

        subtotal = Decimal(root.get("SubTotal", "0"))
        total = Decimal(root.get("Total", "0"))

        iva = Decimal("0")
        otros = Decimal("0")
        impuestos = root.find("cfdi:Impuestos", namespaces=ns)
        if impuestos is not None:
            for tax in impuestos.findall("cfdi:Traslados/cfdi:Traslado", namespaces=ns):
                amount = Decimal(tax.get("Importe", "0"))
                if tax.get("Impuesto", "") == "002":
                    iva += amount
                else:
                    otros += amount
            for tax in impuestos.findall("cfdi:Retenciones/cfdi:Retencion", namespaces=ns):
                amount = Decimal(tax.get("Importe", "0"))
                if tax.get("Impuesto", "") == "002":
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
        return conceptos[0].get("ClaveProdServ", "") if conceptos else ""

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
