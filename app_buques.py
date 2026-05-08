import re
import shutil
import tempfile
import threading
import traceback
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import customtkinter as ctk
import pandas as pd
from openpyxl import load_workbook
from rapidfuzz import fuzz, process
from tksheet import Sheet


DEFAULT_ISAOSA_SHEET = "LINE-UP"
DEFAULT_TWIN_SHEET = "Hoja1"

# ISAOSA es la base maestra acumulada. Los externos sirven para alimentar o
# corregir la base, así que por defecto no reportamos cada registro ISAOSA que
# no apareció en los archivos externos actuales.
INCLUDE_SIN_CAMBIOS = True
INCLUDE_UNMATCHED_ISAOSA = False
MAX_MATCH_DATE_DAYS = 60
BLOCK_DATE_GAP_DAYS = 40

DISPLAY_COLUMNS = [
    "vessel",
    "arrival_date",
    "tonnage",
    "cargo",
    "discharge_port",
    "importer",
    "supplier",
    "loading_port",
    "country_origin",
    "terminal",
    "pier",
    "etb",
    "etd",
    "status",
    "source",
    "source_sheet",
    "source_row",
]

WORK_COLUMNS = [
    "vessel_clean",
    "cargo_clean",
    "discharge_port_clean",
    "importer_clean",
    "supplier_clean",
    "loading_port_clean",
    "arrival_date_clean",
    "tonnage_clean",
    "match_key",
]

RESULT_COLUMNS = [
    "tipo",
    "match",
    "buque_isaosa",
    "buque_externo",
    "campo",
    "valor_isaosa",
    "valor_externo",
    "dato_isaosa",
    "dato_azul",
    "accion_sugerida",
]

WARNING_COLUMNS = ["nivel", "tipo", "archivo", "hoja", "fila", "detalle"]
LOG_COLUMNS = ["momento", "etapa", "detalle"]

FIELD_ALIASES = {
    "vessel": [
        "BARCO",
        "VESSEL",
        "VESSEL NAME",
        "MV",
        "M/V",
    ],
    "laycan": ["LAYCAN"],
    "arrival_date": ["ARRIVAL DATE", "ETA"],
    "tonnage": ["TONNAGE", "MT X VESSEL"],
    "cargo": ["CARGO", "PRODUCT"],
    "discharge_port": [
        "DISCH.PORT",
        "DISCH PORT",
        "DISCH. PORT",
        "DISPORT",
        "DISCHARGE PORT",
    ],
    "importer": ["IMPORTER", "RECEIVER"],
    "cfr_est": ["CFR EST.", "CFR EST"],
    "supplier": [
        "TRADER/SUPPLIER",
        "TRADER/ SUPPLIER",
        "TRADER / SUPPLIER",
        "SHIPPER",
        "SUPPLIER",
    ],
    "loading_port": ["ORIGEN", "LOADING PORT", "LOADPOART", "LOADPORT"],
    "country_origin": ["COUNTRY OF ORIGIN", "ORIGIN COUNTRY"],
    "terminal": ["TERMINAL"],
    "pier": ["PIER"],
    "etb": ["ETB"],
    "etd": ["ETD"],
    "status": ["STATUS"],
    "cost": ["COSTA"],
    "quality_info": ["CALIDAD INFO"],
    "importer_2": ["IMPORTADOR2", "IMPORTADOR 2", "SHORT NAME"],
    "government": ["GOBIERNO"],
    "receiver_tonnage": ["MT X RECEIVER"],
    "port_turnaround_days": ["PORT TURNAROUND (DAYS)"],
    "berthed_time_days": ["BERTHED TIME (DAYS)"],
}

REQUIRED_BY_SOURCE = {
    "ISAOSA": ["vessel", "arrival_date", "tonnage", "cargo", "discharge_port", "importer", "supplier"],
    "buques_azules": ["vessel", "arrival_date", "tonnage", "cargo", "discharge_port", "importer"],
}

ALIAS_TO_FIELD = {}
for standard, aliases in FIELD_ALIASES.items():
    for alias in aliases:
        ALIAS_TO_FIELD[alias] = standard


@dataclass
class ReadResult:
    data: pd.DataFrame
    warnings: list
    meta: dict


def strip_accents(value):
    text = str(value)
    return "".join(
        char for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )


def clean_header_name(value):
    if value is None or pd.isna(value):
        return ""
    text = strip_accents(value)
    text = text.replace("\n", " ").replace("\r", " ")
    text = text.replace('"', "").replace("'", "")
    text = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", text)
    text = re.sub(r"\s+", " ", text).strip().upper()
    return text


def header_key(value):
    text = clean_header_name(value)
    text = text.replace("&", " AND ")
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_text(value):
    if value is None or pd.isna(value):
        return ""
    text = strip_accents(value).upper()
    text = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", text)
    text = re.sub(r"^(M\s*/\s*V|M\.?\s*V\.?|MV)\s+", "", text.strip())
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_date(value):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return pd.NaT
    if isinstance(value, datetime):
        return pd.Timestamp(value).normalize()
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=False)
    if pd.isna(parsed):
        return pd.NaT
    return pd.Timestamp(parsed).normalize()


def parse_user_date(value):
    text = str(value or "").strip()
    if not text:
        return pd.NaT
    return normalize_date(text)


def normalize_number(value):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = strip_accents(value).upper()
    text = text.replace(",", "").replace("$", "")
    text = re.sub(r"\b(MT|MTS|TON|TONS|TM)\b", "", text)
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group())
    except ValueError:
        return None


def safe_display(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value)


def safe_display_field(field, value):
    if field in {"arrival_date", "arrival_date_clean", "fecha", "ETA", "ARRIVAL DATE"}:
        normalized = normalize_date(value)
        return safe_display(normalized) if not pd.isna(normalized) else safe_display(value)
    return safe_display(value)


def build_record_summary(row):
    if row is None:
        return ""
    parts = [
        f"Fecha: {safe_display_field('arrival_date', row.get('arrival_date', ''))}",
        f"Ton: {safe_display(row.get('tonnage', ''))}",
        f"Carga: {safe_display(row.get('cargo', ''))}",
        f"Puerto: {safe_display(row.get('discharge_port', ''))}",
        f"Importador: {safe_display(row.get('importer', ''))}",
        f"Proveedor: {safe_display(row.get('supplier', ''))}",
    ]
    return " | ".join(part for part in parts if not part.endswith(": "))


def filter_by_date_range(df, start_date, end_date):
    if df.empty:
        return df.copy()
    filtered = df.copy()
    if not pd.isna(start_date):
        filtered = filtered[filtered["arrival_date_clean"].isna() | (filtered["arrival_date_clean"] >= start_date)]
    if not pd.isna(end_date):
        filtered = filtered[filtered["arrival_date_clean"].isna() | (filtered["arrival_date_clean"] <= end_date)]
    return filtered.copy()


def describe_range(start_date, end_date):
    start = safe_display(start_date) if not pd.isna(start_date) else "sin inicio"
    end = safe_display(end_date) if not pd.isna(end_date) else "sin fin"
    return f"{start} a {end}"


def warning(nivel, tipo, archivo, hoja="", fila="", detalle=""):
    return {
        "nivel": nivel,
        "tipo": tipo,
        "archivo": Path(str(archivo)).name if archivo else "",
        "hoja": hoja,
        "fila": fila,
        "detalle": detalle,
    }


def friendly_file_error(file_path, exc):
    detail = str(exc)
    if isinstance(exc, PermissionError) or "Permission denied" in detail or "siendo utilizado" in detail:
        return (
            f"No se pudo leer '{Path(str(file_path)).name}' porque Windows lo tiene bloqueado. "
            "Cierra el archivo en Excel, espera a que OneDrive termine de sincronizarlo y vuelve a intentar."
        )
    return detail


def make_temp_excel_copy(file_path):
    source = Path(file_path)
    temp_dir = Path(tempfile.gettempdir()) / "comparador_buques"
    temp_dir.mkdir(parents=True, exist_ok=True)
    target = temp_dir / f"{source.stem}_lectura_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{source.suffix}"
    shutil.copyfile(source, target)
    return target


def now_text():
    return datetime.now().strftime("%H:%M:%S")


def get_workbook_sheet_names(file_path):
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    return workbook.sheetnames


def resolve_sheet_name(file_path, requested_sheet):
    sheet_names = get_workbook_sheet_names(file_path)
    if requested_sheet in sheet_names:
        return requested_sheet, None
    requested_key = header_key(requested_sheet)
    for sheet in sheet_names:
        if header_key(sheet) == requested_key:
            return sheet, warning(
                "ADVERTENCIA",
                "HOJA NORMALIZADA",
                file_path,
                requested_sheet,
                "",
                f"Se usó la hoja real '{sheet}' para la solicitud '{requested_sheet}'.",
            )
    return None, warning(
        "ERROR",
        "HOJA NO ENCONTRADA",
        file_path,
        requested_sheet,
        "",
        f"No existe la hoja '{requested_sheet}'. Hojas disponibles: {', '.join(sheet_names)}",
    )


def detect_header_row(file_path, sheet_name, max_rows=80):
    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=max_rows,
        dtype=object,
        engine="openpyxl",
    )
    alias_keys = {header_key(alias) for alias in ALIAS_TO_FIELD}
    best_row = None
    best_score = 0
    best_headers = []

    for idx, row in preview.iterrows():
        values = [value for value in row.tolist() if clean_header_name(value)]
        keys = [header_key(value) for value in values]
        exact = len(set(keys) & alias_keys)
        fuzzy = 0
        for key in keys:
            if not key:
                continue
            match = process.extractOne(key, alias_keys, scorer=fuzz.ratio)
            if match and match[1] >= 92:
                fuzzy += 1
        density = min(len(values), 20) * 0.05
        score = exact * 2 + fuzzy + density
        if score > best_score:
            best_row = idx
            best_score = score
            best_headers = [clean_header_name(value) for value in values]

    if best_row is None or best_score < 2:
        raise ValueError("No se pudo detectar una fila de encabezados confiable.")
    return int(best_row), best_headers, round(float(best_score), 2)


def map_columns(columns):
    used = {}
    mapped = []
    unmapped = []

    alias_keys = {header_key(alias): field for alias, field in ALIAS_TO_FIELD.items()}
    alias_lookup = list(alias_keys.keys())

    for col in columns:
        clean = clean_header_name(col)
        key = header_key(clean)
        standard = alias_keys.get(key)
        if standard is None and key:
            match = process.extractOne(key, alias_lookup, scorer=fuzz.ratio)
            if match and match[1] >= 94:
                standard = alias_keys[match[0]]

        if standard is None:
            name = clean if clean else "COLUMNA VACIA"
            unmapped.append(name)
            standard = f"extra_{len(unmapped)}_{header_key(name).lower().replace(' ', '_')}"

        if standard in used:
            used[standard] += 1
            standard = f"{standard}_{used[standard]}"
        else:
            used[standard] = 1
        mapped.append(standard)

    return mapped, unmapped


def read_source_sheet(file_path, sheet_name, source_name):
    warnings = []
    read_path = file_path
    temp_path = None
    try:
        real_sheet, sheet_warning = resolve_sheet_name(read_path, sheet_name)
    except Exception as exc:
        try:
            temp_path = make_temp_excel_copy(file_path)
            read_path = temp_path
            real_sheet, sheet_warning = resolve_sheet_name(read_path, sheet_name)
            warnings.append(warning(
                "INFO",
                "LECTURA EN COPIA TEMPORAL",
                file_path,
                sheet_name,
                "",
                "El archivo estaba bloqueado; se leyó una copia temporal para poder comparar.",
            ))
        except Exception:
            warnings.append(warning(
                "ERROR",
                "ARCHIVO NO LEGIBLE",
                file_path,
                sheet_name,
                "",
                friendly_file_error(file_path, exc),
            ))
            return ReadResult(pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS), warnings, {})
    if sheet_warning:
        warnings.append(sheet_warning)
    if not real_sheet:
        return ReadResult(pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS), warnings, {})

    try:
        header_row, headers_found, score = detect_header_row(read_path, real_sheet)
        raw = pd.read_excel(
            read_path,
            sheet_name=real_sheet,
            header=header_row,
            dtype=object,
            engine="openpyxl",
        )
    except Exception as exc:
        if temp_path is None:
            try:
                temp_path = make_temp_excel_copy(file_path)
                read_path = temp_path
                header_row, headers_found, score = detect_header_row(read_path, real_sheet)
                raw = pd.read_excel(
                    read_path,
                    sheet_name=real_sheet,
                    header=header_row,
                    dtype=object,
                    engine="openpyxl",
                )
                warnings.append(warning(
                    "INFO",
                    "LECTURA EN COPIA TEMPORAL",
                    file_path,
                    real_sheet,
                    "",
                    "El archivo estaba bloqueado; se leyó una copia temporal para poder comparar.",
                ))
            except Exception:
                warnings.append(warning("ERROR", "ERROR DE FORMATO", file_path, real_sheet, "", friendly_file_error(file_path, exc)))
                return ReadResult(pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS), warnings, {})
        else:
            warnings.append(warning("ERROR", "ERROR DE FORMATO", file_path, real_sheet, "", friendly_file_error(file_path, exc)))
            return ReadResult(pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS), warnings, {})

    original_columns = [clean_header_name(col) for col in raw.columns]
    mapped_columns, unmapped = map_columns(raw.columns)
    raw.columns = mapped_columns

    for col in DISPLAY_COLUMNS:
        if col not in raw.columns:
            raw[col] = None

    missing = [col for col in REQUIRED_BY_SOURCE.get(source_name, []) if col not in mapped_columns]
    for col in missing:
        warnings.append(warning(
            "ADVERTENCIA",
            "COLUMNA FALTANTE",
            file_path,
            real_sheet,
            header_row + 1,
            f"No se encontró columna estándar '{col}'. Se continuará con valores vacíos.",
        ))

    if unmapped:
        warnings.append(warning(
            "INFO",
            "COLUMNAS NO MAPEADAS",
            file_path,
            real_sheet,
            header_row + 1,
            ", ".join(unmapped[:20]),
        ))

    df = raw.copy()
    df["source"] = source_name
    df["source_sheet"] = real_sheet
    df["source_row"] = df.index + header_row + 2

    df["vessel_clean"] = df["vessel"].apply(normalize_text)
    df["cargo_clean"] = df["cargo"].apply(normalize_text)
    df["discharge_port_clean"] = df["discharge_port"].apply(normalize_text)
    df["importer_clean"] = df["importer"].apply(normalize_text)
    df["supplier_clean"] = df["supplier"].apply(normalize_text)
    df["loading_port_clean"] = df["loading_port"].apply(normalize_text)
    df["arrival_date_clean"] = df["arrival_date"].apply(normalize_date)
    df["tonnage_clean"] = df["tonnage"].apply(normalize_number)
    df["match_key"] = (
        df["vessel_clean"] + "|" + df["discharge_port_clean"] + "|" + df["cargo_clean"]
    )

    empty_before = len(df)
    df = df[df["vessel_clean"] != ""].copy()
    removed = empty_before - len(df)
    if removed:
        warnings.append(warning(
            "INFO",
            "FILAS OMITIDAS",
            file_path,
            real_sheet,
            "",
            f"Se omitieron {removed} filas sin buque después del encabezado.",
        ))

    meta = {
        "archivo": Path(file_path).name,
        "hoja": real_sheet,
        "encabezado_fila": header_row + 1,
        "puntaje_encabezado": score,
        "encabezados_reales": ", ".join(original_columns),
        "encabezados_detectados": ", ".join(headers_found),
        "registros": len(df),
    }
    return ReadResult(df[DISPLAY_COLUMNS + WORK_COLUMNS], warnings, meta)


def load_isaosa(file_path, sheet_name):
    return read_source_sheet(file_path, sheet_name, "ISAOSA")


def load_twin(file_path, sheet_name):
    return read_source_sheet(file_path, sheet_name, "buques_azules")


def date_score(isaosa_date, ext_date):
    if pd.isna(isaosa_date) or pd.isna(ext_date):
        return 50
    days = abs((ext_date - isaosa_date).days)
    if days <= 3:
        return 100
    if days <= 10:
        return 85
    if days <= 30:
        return 65
    if days <= 90:
        return 35
    return 10


def tonnage_score(isaosa_ton, ext_ton):
    if isaosa_ton in (None, 0) or ext_ton in (None, 0):
        return 50
    pct = abs(ext_ton - isaosa_ton) / abs(isaosa_ton) * 100
    if pct <= 3:
        return 100
    if pct <= 8:
        return 80
    if pct <= 20:
        return 55
    return 20


def text_similarity(left, right):
    if not left and not right:
        return 100
    if not left or not right:
        return 50
    return max(
        fuzz.token_sort_ratio(left, right),
        fuzz.token_set_ratio(left, right),
        fuzz.partial_ratio(left, right),
    )


def calculate_match_score(isaosa_row, ext_row):
    vessel = fuzz.token_sort_ratio(ext_row["vessel_clean"], isaosa_row["vessel_clean"])
    port = text_similarity(ext_row["discharge_port_clean"], isaosa_row["discharge_port_clean"]) if ext_row["discharge_port_clean"] and isaosa_row["discharge_port_clean"] else 50
    cargo = text_similarity(ext_row["cargo_clean"], isaosa_row["cargo_clean"]) if ext_row["cargo_clean"] and isaosa_row["cargo_clean"] else 50
    importer = text_similarity(ext_row["importer_clean"], isaosa_row["importer_clean"]) if ext_row["importer_clean"] and isaosa_row["importer_clean"] else 50
    d_score = date_score(isaosa_row["arrival_date_clean"], ext_row["arrival_date_clean"])
    t_score = tonnage_score(isaosa_row["tonnage_clean"], ext_row["tonnage_clean"])
    total = vessel * 0.55 + port * 0.13 + cargo * 0.12 + importer * 0.08 + d_score * 0.06 + t_score * 0.06
    return {
        "total": round(total, 2),
        "vessel": round(vessel, 2),
        "port": round(port, 2),
        "cargo": round(cargo, 2),
        "importer": round(importer, 2),
        "date": round(d_score, 2),
        "tonnage": round(t_score, 2),
    }


def candidate_matches(isaosa_df, ext_row):
    candidates = []
    vessel_choices = isaosa_df["vessel_clean"].fillna("").tolist()
    rough = process.extract(ext_row["vessel_clean"], vessel_choices, scorer=fuzz.token_sort_ratio, limit=80)
    indexes = [item[2] for item in rough if item[1] >= 72]
    for idx in indexes:
        row = isaosa_df.iloc[idx] if isinstance(idx, int) else isaosa_df.loc[idx]
        score = calculate_match_score(row, ext_row)
        if score["vessel"] >= 72:
            candidates.append((row, score))
    if not pd.isna(ext_row["arrival_date_clean"]):
        close_date_candidates = []
        for row, score in candidates:
            if pd.isna(row["arrival_date_clean"]):
                close_date_candidates.append((row, score))
                continue
            days = abs((ext_row["arrival_date_clean"] - row["arrival_date_clean"]).days)
            if days <= MAX_MATCH_DATE_DAYS:
                close_date_candidates.append((row, score))
        candidates = close_date_candidates
    candidates.sort(key=lambda item: item[1]["total"], reverse=True)
    return candidates[:8]


def historical_vessel_matches(isaosa_df, ext_row):
    if pd.isna(ext_row["arrival_date_clean"]):
        return []
    matches = []
    vessel_choices = isaosa_df["vessel_clean"].fillna("").tolist()
    rough = process.extract(ext_row["vessel_clean"], vessel_choices, scorer=fuzz.token_sort_ratio, limit=20)
    for idx, vessel_score, _ in [(item[2], item[1], item[0]) for item in rough if item[1] >= 88]:
        row = isaosa_df.iloc[idx] if isinstance(idx, int) else isaosa_df.loc[idx]
        if pd.isna(row["arrival_date_clean"]):
            continue
        days = abs((ext_row["arrival_date_clean"] - row["arrival_date_clean"]).days)
        if days > MAX_MATCH_DATE_DAYS:
            matches.append((row, days, vessel_score))
    matches.sort(key=lambda item: item[1])
    return matches[:3]


def result_row(tipo, confianza, isaosa_row, ext_row, campo, valor_isaosa, valor_externo, accion, motivo, riesgo):
    if tipo in {"NUEVO REGISTRO", "NUEVO VIAJE", "LINEA AZUL SIN MATCH"}:
        match_status = "NO"
    elif tipo in {"REVISAR POSIBLE COINCIDENCIA", "POSIBLE DUPLICADO", "REVISAR BLOQUE SIMILAR"}:
        match_status = "REVISAR"
    elif tipo == "NO ENCONTRADO EN EXTERNOS":
        match_status = "NO"
    else:
        match_status = "SI"
    dato_isaosa = build_record_summary(isaosa_row)
    dato_azul = build_record_summary(ext_row)
    return {
        "tipo": tipo,
        "match": match_status,
        "buque_isaosa": safe_display(isaosa_row.get("vessel", "")) if isaosa_row is not None else "",
        "buque_externo": safe_display(ext_row.get("vessel", "")) if ext_row is not None else "",
        "campo": campo,
        "valor_isaosa": safe_display_field(campo, valor_isaosa),
        "valor_externo": safe_display_field(campo, valor_externo),
        "dato_isaosa": dato_isaosa,
        "dato_azul": dato_azul,
        "accion_sugerida": accion,
    }


def compare_dates(isaosa_value, external_value):
    if pd.isna(external_value):
        return None
    if pd.isna(isaosa_value):
        return "externo_tiene_dato"
    diff = int((external_value - isaosa_value).days)
    return diff if diff != 0 else None


def compare_tonnage(isaosa_value, external_value):
    if external_value is None:
        return None
    if isaosa_value is None:
        return "externo_tiene_dato"
    if isaosa_value == 0:
        return None
    pct = abs(external_value - isaosa_value) / abs(isaosa_value) * 100
    return round(pct, 2) if pct > 0.5 else None


def detect_changes(isaosa_row, ext_row, score):
    changes = []
    confidence = score["total"]

    date_diff = compare_dates(isaosa_row["arrival_date_clean"], ext_row["arrival_date_clean"])
    if date_diff is not None:
        riesgo = "BAJO" if date_diff == "externo_tiene_dato" or abs(date_diff) <= 10 else "MEDIO"
        changes.append(result_row(
            "CAMBIO FECHA",
            confidence,
            isaosa_row,
            ext_row,
            "arrival_date",
            isaosa_row["arrival_date"],
            ext_row["arrival_date"],
            "Sugerir actualización de fecha",
            f"Diferencia de {date_diff} días" if date_diff != "externo_tiene_dato" else "ISAOSA no tiene fecha y externo sí",
            riesgo,
        ))

    tonnage_diff = compare_tonnage(isaosa_row["tonnage_clean"], ext_row["tonnage_clean"])
    if tonnage_diff is not None and (tonnage_diff == "externo_tiene_dato" or tonnage_diff > 3):
        riesgo = "ALTO"
        tipo = "CAMBIO TONELAJE"
        accion = "Revisar tonelaje; diferencia mayor al 3%"
        changes.append(result_row(
            tipo,
            confidence,
            isaosa_row,
            ext_row,
            "tonnage",
            isaosa_row["tonnage"],
            ext_row["tonnage"],
            accion,
            f"Diferencia aproximada {tonnage_diff}%" if tonnage_diff != "externo_tiene_dato" else "ISAOSA no tiene tonelaje y externo sí",
            riesgo,
        ))

    text_fields = [
        ("cargo", "cargo_clean", "CAMBIO CARGA", "Revisión manual; no actualizar automáticamente"),
        ("discharge_port", "discharge_port_clean", "CAMBIO PUERTO", "Revisión manual; no actualizar automáticamente"),
        ("importer", "importer_clean", "CAMBIO IMPORTADOR", "Revisión manual; no actualizar automáticamente"),
        ("supplier", "supplier_clean", "CAMBIO PROVEEDOR", "Validar proveedor antes de actualizar"),
        ("loading_port", "loading_port_clean", "CAMBIO ORIGEN", "Validar origen antes de actualizar"),
    ]
    for original, clean, tipo, accion in text_fields:
        ext_clean = ext_row[clean]
        isa_clean = isaosa_row[clean]
        if not ext_clean:
            continue
        if not isa_clean:
            changes.append(result_row(
                tipo, confidence, isaosa_row, ext_row, original,
                isaosa_row[original], ext_row[original],
                "Sugerir completar dato faltante en copia/reporte",
                "ISAOSA no tiene dato y externo sí",
                "MEDIO",
            ))
            continue
        similarity = text_similarity(isa_clean, ext_clean)
        threshold = 88 if original in {"cargo", "discharge_port", "importer"} else 82
        if similarity < threshold:
            risk = "ALTO" if original in {"cargo", "discharge_port", "importer"} else "MEDIO"
            changes.append(result_row(
                tipo, confidence, isaosa_row, ext_row, original,
                isaosa_row[original], ext_row[original],
                accion,
                "El texto normalizado no coincide",
                risk,
            ))
    return changes


def row_date(row):
    value = row.get("arrival_date_clean", pd.NaT)
    return value if not pd.isna(value) else pd.NaT


def block_date(block):
    dates = [row_date(row) for row in block["rows"] if not pd.isna(row_date(row))]
    if not dates:
        return pd.NaT
    return min(dates)


def block_date_range_text(block):
    dates = sorted({row_date(row) for row in block["rows"] if not pd.isna(row_date(row))})
    if not dates:
        return "sin fecha"
    if len(dates) == 1:
        return safe_display(dates[0])
    return f"{safe_display(dates[0])} a {safe_display(dates[-1])}"


def block_total_tonnage(block):
    values = [row.get("tonnage_clean") for row in block["rows"] if row.get("tonnage_clean") is not None]
    return sum(values) if values else None


def join_unique(rows, field, limit=8):
    values = []
    seen = set()
    for row in rows:
        value = safe_display(row.get(field, "")).strip()
        key = normalize_text(value)
        if value and key not in seen:
            seen.add(key)
            values.append(value)
    if len(values) > limit:
        return ", ".join(values[:limit]) + f" +{len(values) - limit}"
    return ", ".join(values)


def block_summary_row(block):
    rows = block["rows"]
    return {
        "vessel": block["vessel"],
        "arrival_date": block_date(block),
        "tonnage": block_total_tonnage(block),
        "cargo": join_unique(rows, "cargo"),
        "discharge_port": join_unique(rows, "discharge_port"),
        "importer": join_unique(rows, "importer"),
        "supplier": join_unique(rows, "supplier"),
        "loading_port": join_unique(rows, "loading_port"),
    }


def build_blocks(df, source_label):
    blocks = []
    if df.empty:
        return blocks
    for vessel_clean, group in df[df["vessel_clean"] != ""].groupby("vessel_clean", sort=True):
        rows = []
        for _, row in group.iterrows():
            rows.append(row)
        rows.sort(key=lambda row: (pd.isna(row_date(row)), row_date(row) if not pd.isna(row_date(row)) else pd.Timestamp.max))
        current = []
        last_date = pd.NaT
        for row in rows:
            date = row_date(row)
            split = False
            if current and not pd.isna(date) and not pd.isna(last_date):
                split = abs((date - last_date).days) > BLOCK_DATE_GAP_DAYS
            if split:
                blocks.append(make_block(current, source_label))
                current = []
            current.append(row)
            if not pd.isna(date):
                last_date = date
        if current:
            blocks.append(make_block(current, source_label))
    return blocks


def make_block(rows, source_label):
    first = rows[0]
    return {
        "source": source_label,
        "vessel_clean": first["vessel_clean"],
        "vessel": safe_display(first["vessel"]),
        "rows": rows,
    }


def block_match_score(isaosa_block, external_block):
    vessel = fuzz.token_sort_ratio(external_block["vessel_clean"], isaosa_block["vessel_clean"])
    isa_date = block_date(isaosa_block)
    ext_date = block_date(external_block)
    if pd.isna(isa_date) or pd.isna(ext_date):
        days = None
        date_component = 50
    else:
        days = abs((ext_date - isa_date).days)
        if days <= 3:
            date_component = 100
        elif days <= 10:
            date_component = 90
        elif days <= BLOCK_DATE_GAP_DAYS:
            date_component = 70
        else:
            date_component = 0
    total = vessel * 0.75 + date_component * 0.25
    return total, vessel, days


def block_values_text(block, field):
    return normalize_text(join_unique(block["rows"], field, limit=100))


def line_count_score(isaosa_block, external_block):
    isa_count = len(isaosa_block["rows"])
    ext_count = len(external_block["rows"])
    if not isa_count and not ext_count:
        return 100
    if not isa_count or not ext_count:
        return 0
    diff = abs(isa_count - ext_count)
    return max(0, 100 - (diff / max(isa_count, ext_count)) * 100)


def total_tonnage_similarity(isaosa_block, external_block):
    isa_total = block_total_tonnage(isaosa_block)
    ext_total = block_total_tonnage(external_block)
    return tonnage_score(isa_total, ext_total)


def line_signature(row):
    date_value = row.get("arrival_date_clean", pd.NaT)
    date_key = safe_display(date_value) if not pd.isna(date_value) else ""
    tonnage = row.get("tonnage_clean")
    tonnage_key = round(float(tonnage), 3) if tonnage is not None else None
    return (
        date_key,
        tonnage_key,
        row.get("cargo_clean", ""),
        row.get("discharge_port_clean", ""),
        row.get("importer_clean", ""),
        row.get("supplier_clean", ""),
    )


def exact_line_match_count(isaosa_block, external_block):
    available = {}
    for row in isaosa_block["rows"]:
        key = line_signature(row)
        available[key] = available.get(key, 0) + 1

    matches = 0
    for row in external_block["rows"]:
        key = line_signature(row)
        if available.get(key, 0) > 0:
            available[key] -= 1
            matches += 1
    return matches


def detailed_block_match_score(isaosa_block, external_block):
    vessel = fuzz.token_sort_ratio(external_block["vessel_clean"], isaosa_block["vessel_clean"])
    _, _, days = block_match_score(isaosa_block, external_block)
    if days is None:
        date_component = 45
    elif days <= 3:
        date_component = 100
    elif days <= 10:
        date_component = 92
    elif days <= 20:
        date_component = 78
    elif days <= BLOCK_DATE_GAP_DAYS:
        date_component = 58
    elif days <= 90:
        date_component = 25
    else:
        date_component = 0

    supplier = text_similarity(block_values_text(isaosa_block, "supplier"), block_values_text(external_block, "supplier"))
    importer = text_similarity(block_values_text(isaosa_block, "importer"), block_values_text(external_block, "importer"))
    cargo = text_similarity(block_values_text(isaosa_block, "cargo"), block_values_text(external_block, "cargo"))
    port = text_similarity(block_values_text(isaosa_block, "discharge_port"), block_values_text(external_block, "discharge_port"))
    line_count = line_count_score(isaosa_block, external_block)
    tonnage = total_tonnage_similarity(isaosa_block, external_block)
    exact_lines = exact_line_match_count(isaosa_block, external_block)
    exact_line_score = exact_lines / max(len(external_block["rows"]), 1) * 100

    total = (
        exact_line_score * 0.34
        + vessel * 0.18
        + date_component * 0.18
        + supplier * 0.10
        + importer * 0.08
        + line_count * 0.05
        + port * 0.03
        + cargo * 0.02
        + tonnage * 0.02
    )
    return {
        "total": round(total, 2),
        "vessel": round(vessel, 2),
        "days": days,
        "date": round(date_component, 2),
        "supplier": round(supplier, 2),
        "importer": round(importer, 2),
        "line_count": round(line_count, 2),
        "port": round(port, 2),
        "cargo": round(cargo, 2),
        "tonnage": round(tonnage, 2),
        "exact_lines": exact_lines,
        "exact_line_score": round(exact_line_score, 2),
    }


def find_matching_block(isaosa_blocks, external_block):
    candidates = []
    for block in isaosa_blocks:
        score = detailed_block_match_score(block, external_block)
        if score["vessel"] >= 82:
            candidates.append((block, score))
    candidates.sort(key=lambda item: (item[1]["exact_lines"], item[1]["total"]), reverse=True)
    if not candidates:
        return None
    best_block, best_score = candidates[0]
    if best_score["total"] < 62:
        return None
    if best_score["days"] is not None and best_score["days"] > 90 and best_score["total"] < 78:
        return None
    return best_block, best_score


def explain_block_match(isaosa_block, external_block, score):
    if isaosa_block is None:
        return "No encontró un bloque ISAOSA suficientemente parecido por nombre de buque y fecha."
    pieces = []
    if score["days"] is None:
        pieces.append("no hay fecha suficiente para medir distancia")
    elif score["days"] == 0:
        pieces.append("misma fecha de llegada")
    else:
        pieces.append(f"fechas a {score['days']} días")
    if score["supplier"] >= 85:
        pieces.append("proveedores muy parecidos")
    elif score["supplier"] >= 60:
        pieces.append("proveedores parcialmente parecidos")
    else:
        pieces.append("proveedores diferentes")
    if score["importer"] >= 85:
        pieces.append("importadores muy parecidos")
    elif score["importer"] >= 60:
        pieces.append("importadores parcialmente parecidos")
    else:
        pieces.append("importadores diferentes")
    if score["line_count"] >= 90:
        pieces.append("cantidad de líneas similar")
    else:
        pieces.append(f"líneas distintas: ISAOSA {len(isaosa_block['rows'])}, azul {len(external_block['rows'])}")
    if score.get("exact_lines", 0):
        pieces.append(f"{score['exact_lines']} lineas exactas")
    if score["port"] >= 85:
        pieces.append("puertos parecidos")
    if score["cargo"] >= 85:
        pieces.append("cargas parecidas")
    return "La app eligió este bloque porque " + "; ".join(pieces) + "."


def detail_line_score(isaosa_row, ext_row):
    scores = detail_line_components(isaosa_row, ext_row)
    return (
        scores["importer"] * 0.28
        + scores["cargo"] * 0.22
        + scores["port"] * 0.18
        + scores["tonnage"] * 0.16
        + scores["supplier"] * 0.08
        + scores["origin"] * 0.08
    )


def detail_line_components(isaosa_row, ext_row):
    return {
        "importer": text_similarity(isaosa_row["importer_clean"], ext_row["importer_clean"]),
        "cargo": text_similarity(isaosa_row["cargo_clean"], ext_row["cargo_clean"]),
        "port": text_similarity(isaosa_row["discharge_port_clean"], ext_row["discharge_port_clean"]),
        "tonnage": tonnage_score(isaosa_row["tonnage_clean"], ext_row["tonnage_clean"]),
        "supplier": text_similarity(isaosa_row["supplier_clean"], ext_row["supplier_clean"]),
        "origin": text_similarity(isaosa_row["loading_port_clean"], ext_row["loading_port_clean"]),
    }


def line_match_allowed(components):
    critical_conflicts = 0
    if components["cargo"] < 70:
        critical_conflicts += 1
    if components["port"] < 75:
        critical_conflicts += 1
    if components["importer"] < 75:
        critical_conflicts += 1
    if components["tonnage"] < 55:
        critical_conflicts += 1

    if critical_conflicts >= 2:
        return False
    if components["cargo"] < 60 and components["tonnage"] < 80:
        return False
    if components["port"] < 70 and components["importer"] < 88:
        return False
    return True


def explain_detail_line_match(isaosa_row, ext_row, score):
    parts = []
    labels = [
        ("importer", "importador"),
        ("cargo", "carga"),
        ("port", "puerto"),
        ("tonnage", "tonelaje"),
        ("supplier", "proveedor"),
        ("origin", "origen"),
    ]
    for key, label in labels:
        value = score.get(key, 0)
        if value >= 88:
            parts.append(f"{label} coincide")
        elif value >= 65:
            parts.append(f"{label} parecido")
        else:
            parts.append(f"{label} distinto")
    return "Se ligó con esta línea ISAOSA porque " + "; ".join(parts) + "."


def line_display_title(row):
    importer = safe_display(row.get("importer", "SIN IMPORTADOR")) or "SIN IMPORTADOR"
    cargo = safe_display(row.get("cargo", "SIN CARGA")) or "SIN CARGA"
    tonnage = safe_display(row.get("tonnage", ""))
    port = safe_display(row.get("discharge_port", ""))
    pieces = [importer, cargo]
    if tonnage:
        pieces.append(tonnage)
    if port:
        pieces.append(port)
    return " | ".join(pieces)


def find_best_detail_line(isaosa_rows, ext_row, used_indexes):
    best = None
    for idx, isa_row in enumerate(isaosa_rows):
        if idx in used_indexes:
            continue
        score = detail_line_score(isa_row, ext_row)
        if best is None or score > best[1]:
            best = (idx, score, isa_row)
    return best


def compare_block_details(isaosa_block, external_block, include_sin_cambios, log):
    results = []
    isa_summary = block_summary_row(isaosa_block)
    ext_summary = block_summary_row(external_block)
    isa_total = block_total_tonnage(isaosa_block)
    ext_total = block_total_tonnage(external_block)
    tonnage_diff = compare_tonnage(isa_total, ext_total)
    if tonnage_diff is not None and (tonnage_diff == "externo_tiene_dato" or tonnage_diff > 3):
        results.append(result_row(
            "CAMBIO TONELAJE TOTAL",
            100,
            isa_summary,
            ext_summary,
            "tonnage",
            isa_total,
            ext_total,
            "Revisar tonelaje total del bloque",
            "El tonelaje total del viaje cambió",
            "ALTO",
        ))

    isa_date = block_date(isaosa_block)
    ext_date = block_date(external_block)
    if not pd.isna(isa_date) and not pd.isna(ext_date) and isa_date != ext_date:
        results.append(result_row(
            "CAMBIO FECHA BLOQUE",
            100,
            isa_summary,
            ext_summary,
            "arrival_date",
            isa_date,
            ext_date,
            "Revisar fecha del viaje",
            "La fecha principal del bloque cambió",
            "MEDIO",
        ))

    assignments = {}
    used_isaosa = set()
    used_external = set()

    exact_available = {}
    for isa_idx, isa_row in enumerate(isaosa_block["rows"]):
        exact_available.setdefault(line_signature(isa_row), []).append((isa_idx, isa_row))

    for ext_idx, ext_row in enumerate(external_block["rows"]):
        choices = exact_available.get(line_signature(ext_row), [])
        while choices and choices[0][0] in used_isaosa:
            choices.pop(0)
        if choices:
            isa_idx, isa_row = choices.pop(0)
            assignments[ext_idx] = (isa_idx, 100, isa_row)
            used_isaosa.add(isa_idx)
            used_external.add(ext_idx)

    all_scores = []
    for ext_idx, ext_row in enumerate(external_block["rows"]):
        if ext_idx in used_external:
            continue
        for isa_idx, isa_row in enumerate(isaosa_block["rows"]):
            if isa_idx in used_isaosa:
                continue
            components = detail_line_components(isa_row, ext_row)
            if not line_match_allowed(components):
                continue
            all_scores.append((detail_line_score(isa_row, ext_row), ext_idx, isa_idx, isa_row, ext_row))
    all_scores.sort(reverse=True, key=lambda item: item[0])

    for score, ext_idx, isa_idx, isa_row, ext_row in all_scores:
        if score < 72:
            continue
        if ext_idx in assignments or isa_idx in used_isaosa:
            continue
        assignments[ext_idx] = (isa_idx, score, isa_row)
        used_isaosa.add(isa_idx)

    for ext_idx, ext_row in enumerate(external_block["rows"]):
        best = assignments.get(ext_idx)
        if best is None:
            log("Línea azul sin match", f"{safe_display(ext_row['vessel'])}: no encontré línea ISAOSA equivalente dentro del bloque.")
            results.append(result_row(
                "LINEA AZUL SIN MATCH",
                0,
                isa_summary,
                ext_row,
                "registro",
                "No existe línea equivalente en el bloque ISAOSA",
                safe_display(ext_row.get("vessel", "")),
                "Agregar esta línea al viaje/bloque ISAOSA",
                "Ninguna línea ISAOSA del bloque se parece lo suficiente",
                "MEDIO",
            ))
            continue
        _, _, isa_row = best
        changes = detect_changes(isa_row, ext_row, {"total": 100})
        if changes:
            results.extend(changes)
            for change in changes:
                log("Cambio detectado", f"{external_block['vessel']} bloque {block_date_range_text(external_block)} | {change['tipo']} {change['campo']}")
        elif include_sin_cambios:
            results.append(result_row(
                "SIN CAMBIOS",
                100,
                isa_row,
                ext_row,
                "",
                "",
                "",
                "No hacer cambios",
                "Línea azul coincide con línea ISAOSA del bloque",
                "BAJO",
            ))
    return results


def build_line_pairs_for_blocks(isa_block, ext_block):
    pairs = []
    isa_rows = isa_block["rows"] if isa_block else []
    ext_rows = ext_block["rows"]
    assignments = {}
    used_isa = set()
    used_external = set()

    exact_available = {}
    for isa_idx, isa_row in enumerate(isa_rows):
        exact_available.setdefault(line_signature(isa_row), []).append((isa_idx, isa_row))

    for ext_idx, ext_row in enumerate(ext_rows):
        choices = exact_available.get(line_signature(ext_row), [])
        while choices and choices[0][0] in used_isa:
            choices.pop(0)
        if choices:
            isa_idx, isa_row = choices.pop(0)
            assignments[ext_idx] = (isa_idx, 100, isa_row)
            used_isa.add(isa_idx)
            used_external.add(ext_idx)

    all_scores = []
    for ext_idx, ext_row in enumerate(ext_rows):
        if ext_idx in used_external:
            continue
        for isa_idx, isa_row in enumerate(isa_rows):
            if isa_idx in used_isa:
                continue
            components = detail_line_components(isa_row, ext_row)
            if not line_match_allowed(components):
                continue
            all_scores.append((detail_line_score(isa_row, ext_row), ext_idx, isa_idx, isa_row, ext_row))
    all_scores.sort(reverse=True, key=lambda item: item[0])

    for score, ext_idx, isa_idx, isa_row, ext_row in all_scores:
        if score < 72:
            continue
        if ext_idx in assignments or isa_idx in used_isa:
            continue
        assignments[ext_idx] = (isa_idx, score, isa_row)
        used_isa.add(isa_idx)

    for ext_idx, ext_row in enumerate(ext_rows):
        assignment = assignments.get(ext_idx)
        if assignment is None:
            pairs.append({
                "isaosa": None,
                "azul": ext_row,
                "changes": [],
                "explanation": "No se encontró una línea ISAOSA suficientemente parecida dentro del bloque.",
                "components": {},
            })
            continue
        _, _, isa_row = assignment
        changes = detect_changes(isa_row, ext_row, {"total": 100})
        components = detail_line_components(isa_row, ext_row)
        pairs.append({
            "isaosa": isa_row,
            "azul": ext_row,
            "changes": changes,
            "explanation": explain_detail_line_match(isa_row, ext_row, components),
            "components": components,
        })
    return pairs


def pair_decision_key(pair):
    azul = pair.get("azul")
    isaosa = pair.get("isaosa")
    azul_id = f"{safe_display(azul.get('source_sheet', ''))}:{safe_display(azul.get('source_row', ''))}:{normalize_text(azul.get('vessel', ''))}" if azul is not None else "SIN_AZUL"
    isaosa_id = f"{safe_display(isaosa.get('source_sheet', ''))}:{safe_display(isaosa.get('source_row', ''))}:{normalize_text(isaosa.get('vessel', ''))}" if isaosa is not None else "SIN_ISAOSA"
    return f"{azul_id}=>{isaosa_id}"


def compare_all(
    isaosa_df,
    external_df,
    log_callback=None,
    include_unmatched_isaosa=INCLUDE_UNMATCHED_ISAOSA,
    include_sin_cambios=INCLUDE_SIN_CAMBIOS,
):
    def log(etapa, detalle):
        if log_callback:
            log_callback(etapa, detalle)

    results = []
    external_df = external_df.copy()
    if not external_df.empty and "listas_externas" not in external_df.columns:
        lists_by_vessel = (
            external_df.groupby("vessel_clean")["source"]
            .apply(lambda values: ", ".join(sorted({safe_display(value) for value in values if safe_display(value)})))
            .to_dict()
        )
        external_df["listas_externas"] = external_df["vessel_clean"].map(lists_by_vessel).fillna(external_df["source"])

    if isaosa_df.empty:
        log("Comparación", "ISAOSA está vacía; todo externo se marcará como nuevo registro.")
        for _, ext_row in external_df.iterrows():
            log("Externo -> ISAOSA", f"{safe_display(ext_row['source'])}/{safe_display(ext_row['source_sheet'])}: {safe_display(ext_row['vessel'])} no tiene base ISAOSA comparable.")
            results.append(result_row(
                "NUEVO REGISTRO", 0, None, ext_row, "registro",
                "No existe base ISAOSA cargada", "Existe en fuente externa",
                "Agregar a revisión para alta", "No hay registros ISAOSA comparables", "ALTO",
            ))
        return pd.DataFrame(results, columns=RESULT_COLUMNS)

    isaosa_blocks = build_blocks(isaosa_df, "ISAOSA")
    external_blocks = build_blocks(external_df, "buques_azules")
    log("Bloques azules", f"Se encontraron {len(external_blocks)} bloques en buques_azules.")
    for idx, block in enumerate(external_blocks, start=1):
        log("Bloque azul", f"{idx}. {block['vessel']} | {block_date_range_text(block)} | {len(block['rows'])} líneas.")

    for idx, ext_block in enumerate(external_blocks, start=1):
        match = find_matching_block(isaosa_blocks, ext_block)
        ext_summary = block_summary_row(ext_block)
        if not match:
            log("Bloque sin match", f"{ext_block['vessel']} | {block_date_range_text(ext_block)} no tiene bloque ISAOSA suficientemente parecido.")
            results.append(result_row(
                "NUEVO VIAJE",
                0,
                None,
                ext_summary,
                "bloque",
                "No existe bloque ISAOSA equivalente",
                ext_block["vessel"],
                "Agregar este bloque/viaje de buques_azules a ISAOSA",
                "No hubo bloque ISAOSA con mismo buque y fecha cercana",
                "MEDIO",
            ))
            continue

        isa_block, match_score = match
        log(
            "Bloque con match",
            f"{ext_block['vessel']} | azul {block_date_range_text(ext_block)} coincide con ISAOSA {block_date_range_text(isa_block)} | "
            f"{len(ext_block['rows'])} líneas azul vs {len(isa_block['rows'])} líneas ISAOSA | "
            f"fecha {match_score['date']} proveedor {match_score['supplier']} importador {match_score['importer']} líneas {match_score['line_count']}.",
        )
        block_results = compare_block_details(isa_block, ext_block, include_sin_cambios, log)
        results.extend(block_results)

    return pd.DataFrame(results, columns=RESULT_COLUMNS)


def build_summary(results_df, isaosa_df, external_df, warnings_df):
    rows = [
        ("Total registros ISAOSA", len(isaosa_df)),
        ("Total registros externos", len(external_df)),
        ("Total sin cambios", int((results_df["tipo"] == "SIN CAMBIOS").sum()) if not results_df.empty else 0),
        ("Total cambios fecha", int((results_df["tipo"] == "CAMBIO FECHA").sum()) if not results_df.empty else 0),
        ("Total cambios tonelaje", int((results_df["tipo"] == "CAMBIO TONELAJE").sum()) if not results_df.empty else 0),
        ("Total nuevos registros", int((results_df["tipo"] == "NUEVO REGISTRO").sum()) if not results_df.empty else 0),
        ("Total revisión manual", int(results_df["tipo"].isin(["REVISAR POSIBLE COINCIDENCIA", "POSIBLE DUPLICADO", "CAMBIO CARGA", "CAMBIO PUERTO", "CAMBIO IMPORTADOR"]).sum()) if not results_df.empty else 0),
        ("Total no encontrados", int((results_df["tipo"] == "NO ENCONTRADO EN EXTERNOS").sum()) if not results_df.empty else 0),
        ("Total advertencias", len(warnings_df)),
    ]
    return pd.DataFrame(rows, columns=["Metrica", "Valor"])


def build_block_report(isaosa_df, external_df, decisions=None):
    decisions = decisions or {}
    columns = [
        "buque",
        "decision_usuario",
        "estado",
        "bloque_azul",
        "bloque_isaosa",
        "linea_azul",
        "linea_isaosa",
        "cambios_detectados",
        "criterio_usado",
        "accion_sugerida",
    ]
    if isaosa_df.empty or external_df.empty:
        return pd.DataFrame(columns=columns)

    rows = []
    isaosa_blocks = build_blocks(isaosa_df, "ISAOSA")
    external_blocks = build_blocks(external_df, "buques_azules")
    for ext_block in external_blocks:
        match = find_matching_block(isaosa_blocks, ext_block)
        isa_block = match[0] if match else None
        match_score = match[1] if match else None
        block_explanation = explain_block_match(isa_block, ext_block, match_score) if match else explain_block_match(None, ext_block, {})
        pairs = build_line_pairs_for_blocks(isa_block, ext_block)
        for pair in pairs:
            azul = pair["azul"]
            isaosa = pair["isaosa"]
            if isaosa is None:
                estado = "NUEVA LINEA / SIN MATCH"
                cambios = "No existe una línea ISAOSA suficientemente parecida."
                accion = "Agregar o revisar esta línea en ISAOSA."
            elif pair["changes"]:
                estado = "CON CAMBIOS"
                cambios = "; ".join(
                    f"{safe_display(change['campo'])}: {safe_display(change['valor_isaosa'])} -> {safe_display(change['valor_externo'])}"
                    for change in pair["changes"]
                )
                accion = "; ".join(sorted({safe_display(change["accion_sugerida"]) for change in pair["changes"] if safe_display(change["accion_sugerida"])}))
            else:
                estado = "SIN CAMBIOS"
                cambios = ""
                accion = "No hacer cambios."

            decision = decisions.get(pair_decision_key(pair), "PENDIENTE" if estado != "SIN CAMBIOS" else "")
            rows.append({
                "buque": safe_display(azul.get("vessel", ext_block["vessel"])),
                "decision_usuario": decision,
                "estado": estado,
                "bloque_azul": f"{block_date_range_text(ext_block)} | {len(ext_block['rows'])} líneas | total {safe_display(block_total_tonnage(ext_block))}",
                "bloque_isaosa": f"{block_date_range_text(isa_block)} | {len(isa_block['rows'])} líneas | total {safe_display(block_total_tonnage(isa_block))}" if isa_block else "No encontrado",
                "linea_azul": build_record_summary(azul),
                "linea_isaosa": build_record_summary(isaosa) if isaosa is not None else "Sin línea equivalente",
                "cambios_detectados": cambios,
                "criterio_usado": f"{block_explanation} {pair['explanation']}",
                "accion_sugerida": accion,
            })
    return pd.DataFrame(rows, columns=columns)


def export_report(output_path, results_df, isaosa_df, external_df, warnings_df, logs_df=None, decisions=None):
    summary_df = build_summary(results_df, isaosa_df, external_df, warnings_df)
    logs_df = logs_df if logs_df is not None else pd.DataFrame(columns=LOG_COLUMNS)
    block_report_df = build_block_report(isaosa_df, external_df, decisions)
    accepted_df = block_report_df[block_report_df["decision_usuario"] == "ACEPTAR"].copy() if not block_report_df.empty else block_report_df.copy()
    ignored_df = block_report_df[block_report_df["decision_usuario"] == "IGNORAR"].copy() if not block_report_df.empty else block_report_df.copy()
    pending_df = block_report_df[block_report_df["decision_usuario"] == "PENDIENTE"].copy() if not block_report_df.empty else block_report_df.copy()

    sheets = {
        "Vista por buque": block_report_df,
        "Aceptados": accepted_df,
        "Ignorados": ignored_df,
        "Pendientes": pending_df,
        "Comparacion detalle": results_df,
        "Resumen": summary_df,
        "Log Proceso": logs_df,
    }

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        high_format = workbook.add_format({"bg_color": "#FCE4D6"})
        medium_format = workbook.add_format({"bg_color": "#FFF2CC"})
        low_format = workbook.add_format({"bg_color": "#E2F0D9"})

        for sheet_name, df in sheets.items():
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                sample = [len(str(value))]
                if not df.empty:
                    sample.extend(df.iloc[:200, col_num].map(lambda item: len(safe_display(item))).tolist())
                max_width = 85 if sheet_name == "Vista por buque" else 42
                width = min(max(sample) + 2, max_width)
                worksheet.set_column(col_num, col_num, width)
            if "nivel_riesgo" in df.columns and not df.empty:
                risk_col = df.columns.get_loc("nivel_riesgo")
                worksheet.conditional_format(1, risk_col, len(df), risk_col, {"type": "text", "criteria": "containing", "value": "ALTO", "format": high_format})
                worksheet.conditional_format(1, risk_col, len(df), risk_col, {"type": "text", "criteria": "containing", "value": "MEDIO", "format": medium_format})
                worksheet.conditional_format(1, risk_col, len(df), risk_col, {"type": "text", "criteria": "containing", "value": "BAJO", "format": low_format})


class BuquesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador Local de Buques - ISAOSA")
        self.root.geometry("1500x900")
        self.root.minsize(1180, 720)

        self.isaosa_path = tk.StringVar()
        self.twin_path = tk.StringVar()
        self.isaosa_sheet = tk.StringVar(value=DEFAULT_ISAOSA_SHEET)
        self.twin_sheet = tk.StringVar(value=DEFAULT_TWIN_SHEET)
        self.status_var = tk.StringVar(value="Listo.")
        self.config_visible = tk.BooleanVar(value=False)
        self.include_unmatched_isaosa = tk.BooleanVar(value=False)
        self.include_sin_cambios = tk.BooleanVar(value=True)
        self.isaosa_date_from = tk.StringVar()
        self.isaosa_date_to = tk.StringVar()
        self.azul_date_from = tk.StringVar()
        self.azul_date_to = tk.StringVar()
        self.block_filter_text = tk.StringVar()
        self.block_filter_type = tk.StringVar(value="Todos")

        self.results_df = pd.DataFrame(columns=RESULT_COLUMNS)
        self.isaosa_df = pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS)
        self.external_df = pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS)
        self.warnings_df = pd.DataFrame(columns=WARNING_COLUMNS)
        self.logs_df = pd.DataFrame(columns=LOG_COLUMNS)
        self.meta_rows = []
        self.review_decisions = {}
        self.block_item_keys = {}
        self.worker = None

        self.create_ui()
        self.autodetect_default_files()

    def create_ui(self):
        self.root.configure(fg_color="#F4F7FB")
        main = ctk.CTkFrame(self.root, fg_color="#F4F7FB", corner_radius=0)
        main.pack(fill="both", expand=True)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(header, text="Detectar cambios en buques", font=ctk.CTkFont(size=20, weight="bold"), text_color="#172033").pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Compara la lista azul contra la base ISAOSA. El resultado muestra buques nuevos y campos que cambiaron.",
            font=ctk.CTkFont(size=11),
            text_color="#5B6472",
        ).pack(anchor="w", pady=(1, 0))

        flow = ctk.CTkFrame(main, fg_color="#FFFFFF", corner_radius=10, border_width=1, border_color="#DDE5F0")
        flow.pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkLabel(flow, text="Archivos principales", font=ctk.CTkFont(size=13, weight="bold"), text_color="#172033").grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(7, 3))
        self.simple_file_row(
            flow,
            1,
            "1. Base local ISAOSA",
            "Archivo maestro. La app NO lo modifica; solo lo usa para comparar.",
            self.isaosa_path,
        )
        self.simple_file_row(
            flow,
            2,
            "2. Lista azul a revisar",
            "Archivo externo que se revisa contra ISAOSA para encontrar altas y cambios.",
            self.twin_path,
        )
        quick = ctk.CTkFrame(main, fg_color="transparent")
        quick.pack(fill="x", padx=12, pady=(0, 6))
        self.compare_button = ctk.CTkButton(quick, text="Detectar cambios", command=self.run_comparison, height=30, fg_color="#2457A6", hover_color="#1D4788")
        self.compare_button.pack(side="left", padx=(0, 8))
        ctk.CTkButton(quick, text="Exportar reporte Excel", command=self.export_results, height=30, fg_color="#1F7A55", hover_color="#176242").pack(side="left", padx=(0, 8))
        ctk.CTkButton(quick, text="Limpiar", command=self.clear_results, height=30, fg_color="#8792A2", hover_color="#707B8A").pack(side="left", padx=(0, 8))
        self.config_button = ctk.CTkButton(quick, text="Configuración avanzada", command=self.toggle_config, height=30, fg_color="#FFFFFF", hover_color="#EEF3FA", text_color="#2457A6", border_width=1, border_color="#C9D6E8")
        self.config_button.pack(side="right")

        self.config_frame = ctk.CTkFrame(main, fg_color="#FFFFFF", corner_radius=10, border_width=1, border_color="#DDE5F0")
        ctk.CTkLabel(self.config_frame, text="Configuración avanzada", font=ctk.CTkFont(size=15, weight="bold"), text_color="#172033").grid(row=0, column=0, columnspan=5, sticky="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(
            self.config_frame,
            text="Solo cambia esto si el Excel cambió de hoja o si quieres ver información extra en el reporte.",
            text_color="#5B6472",
        ).grid(row=1, column=0, columnspan=5, sticky="w", padx=14, pady=(0, 8))
        self.sheet_row(self.config_frame, 2, "Hoja de base ISAOSA", self.isaosa_sheet, "Normalmente: LINE-UP")
        self.sheet_row(self.config_frame, 3, "Hoja de lista azul", self.twin_sheet, "Normalmente: Hoja1")
        next_row = 4
        ctk.CTkCheckBox(
            self.config_frame,
            text="Mostrar también buques históricos de ISAOSA que no aparezcan en la lista azul",
            variable=self.include_unmatched_isaosa,
            text_color="#263244",
        ).grid(row=next_row, column=1, columnspan=3, sticky="w", padx=14, pady=(10, 12))
        ctk.CTkCheckBox(
            self.config_frame,
            text="Mostrar también buques que no tuvieron cambios",
            variable=self.include_sin_cambios,
            state="disabled",
            text_color="#263244",
        ).grid(row=next_row, column=4, sticky="w", padx=(8, 14), pady=(10, 12))

        ctk.CTkLabel(main, textvariable=self.status_var, font=ctk.CTkFont(size=11), text_color="#5B6472").pack(anchor="w", padx=12, pady=(0, 4))

        self.notebook = ctk.CTkTabview(main, fg_color="#FFFFFF", segmented_button_fg_color="#E8EEF7", segmented_button_selected_color="#2457A6", segmented_button_selected_hover_color="#1D4788")
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.block_tree = self.create_block_view_tab("Vista por buque")
        self.log_tree = self.create_table_tab("Log")

    def simple_file_row(self, parent, row, title, help_text, path_var):
        ctk.CTkLabel(parent, text=title, font=ctk.CTkFont(size=11, weight="bold"), text_color="#263244").grid(row=row, column=0, sticky="w", padx=(10, 8), pady=3)
        ctk.CTkLabel(parent, text=help_text, font=ctk.CTkFont(size=10), text_color="#5B6472").grid(row=row, column=1, sticky="w", padx=(0, 8), pady=3)
        ctk.CTkEntry(parent, textvariable=path_var, height=26, font=ctk.CTkFont(size=10)).grid(row=row, column=2, sticky="ew", padx=(0, 6), pady=3)
        ctk.CTkButton(parent, text="Cambiar", command=lambda: self.select_file(path_var), height=26, width=82, fg_color="#EEF3FA", hover_color="#E2EAF5", text_color="#2457A6").grid(row=row, column=3, sticky="e", padx=(0, 10), pady=3)
        parent.columnconfigure(2, weight=1)

    def sheet_row(self, parent, row, label, sheet_var, help_text):
        ctk.CTkLabel(parent, text=label, text_color="#263244").grid(row=row, column=0, sticky="w", padx=(14, 8), pady=5)
        ctk.CTkEntry(parent, textvariable=sheet_var, width=190, height=32).grid(row=row, column=1, sticky="w", padx=(0, 8), pady=5)
        ctk.CTkLabel(parent, text=help_text, text_color="#5B6472").grid(row=row, column=2, columnspan=3, sticky="w", pady=5)

    def range_row(self, parent, row, label, help_text, from_var, to_var):
        ctk.CTkLabel(parent, text=label, font=ctk.CTkFont(size=11, weight="bold"), text_color="#263244").grid(row=row, column=0, sticky="w", padx=(10, 8), pady=3)
        ctk.CTkLabel(parent, text=help_text, font=ctk.CTkFont(size=10), text_color="#5B6472").grid(row=row, column=1, sticky="w", padx=(0, 8), pady=3)
        range_frame = ctk.CTkFrame(parent, fg_color="transparent")
        range_frame.grid(row=row, column=2, sticky="w", pady=2)
        ctk.CTkLabel(range_frame, text="Desde", font=ctk.CTkFont(size=10), text_color="#5B6472").pack(side="left", padx=(0, 3))
        ctk.CTkEntry(range_frame, textvariable=from_var, width=92, height=26, font=ctk.CTkFont(size=10), placeholder_text="2026-05-01").pack(side="left", padx=(0, 6))
        ctk.CTkLabel(range_frame, text="Hasta", font=ctk.CTkFont(size=10), text_color="#5B6472").pack(side="left", padx=(0, 3))
        ctk.CTkEntry(range_frame, textvariable=to_var, width=92, height=26, font=ctk.CTkFont(size=10), placeholder_text="2026-05-31").pack(side="left")
        ctk.CTkLabel(parent, text="Vacío = todo", font=ctk.CTkFont(size=10), text_color="#7B8492").grid(row=row, column=3, sticky="w", padx=(6, 10), pady=2)

    def file_row(self, parent, row, label, path_var, sheet_var, sheet_label):
        ctk.CTkLabel(parent, text=label, text_color="#263244").grid(row=row, column=0, sticky="w", padx=(14, 6), pady=5)
        ctk.CTkEntry(parent, textvariable=path_var, height=32).grid(row=row, column=1, sticky="ew", padx=(0, 6), pady=5)
        ctk.CTkButton(parent, text="Buscar", command=lambda: self.select_file(path_var), height=32, width=90).grid(row=row, column=2, sticky="w", padx=(0, 12), pady=5)
        ctk.CTkLabel(parent, text=sheet_label, text_color="#263244").grid(row=row, column=3, sticky="w", padx=(0, 6), pady=5)
        ctk.CTkEntry(parent, textvariable=sheet_var, height=32).grid(row=row, column=4, sticky="ew", padx=(0, 14), pady=5)
        parent.columnconfigure(1, weight=2)
        parent.columnconfigure(4, weight=1)

    def toggle_config(self):
        if self.config_visible.get():
            self.config_frame.pack_forget()
            self.config_visible.set(False)
            self.config_button.configure(text="Configuración avanzada")
        else:
            self.config_frame.pack(fill="x", pady=(0, 10), after=self.config_button.master)
            self.config_visible.set(True)
            self.config_button.configure(text="Ocultar configuración")

    def create_table_tab(self, title):
        self.notebook.add(title)
        frame = self.notebook.tab(title)
        tree = ttk.Treeview(frame, show="headings")
        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def create_block_view_tab(self, title):
        self.notebook.add(title)
        frame = self.notebook.tab(title)

        filters = ctk.CTkFrame(frame, fg_color="#F7FAFE", corner_radius=8)
        filters.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(5, 4))
        ctk.CTkLabel(filters, text="Buscar", font=ctk.CTkFont(size=11), text_color="#263244").pack(side="left", padx=(8, 5), pady=5)
        search_entry = ctk.CTkEntry(filters, textvariable=self.block_filter_text, width=300, height=26, font=ctk.CTkFont(size=10), placeholder_text="Buque, importador, puerto, carga...")
        search_entry.pack(side="left", padx=(0, 8), pady=5)
        search_entry.bind("<Return>", lambda _event: self.refresh_block_view())
        ctk.CTkLabel(filters, text="Tipo", font=ctk.CTkFont(size=11), text_color="#263244").pack(side="left", padx=(0, 5), pady=5)
        ctk.CTkOptionMenu(
            filters,
            variable=self.block_filter_type,
            values=["Todos", "Con cambios", "Nuevos", "Revisar", "Sin cambios"],
            width=132,
            height=26,
            font=ctk.CTkFont(size=10),
            dropdown_font=ctk.CTkFont(size=10),
            command=lambda _: self.refresh_block_view(),
        ).pack(side="left", padx=(0, 8), pady=5)
        ctk.CTkButton(filters, text="Aplicar", height=26, width=76, command=self.refresh_block_view).pack(side="left", padx=(0, 8), pady=5)
        ctk.CTkButton(filters, text="Limpiar", height=26, width=76, fg_color="#8792A2", hover_color="#707B8A", command=self.clear_block_filters).pack(side="left", padx=(0, 8), pady=5)
        ctk.CTkLabel(filters, text="Selecciona una fila y decide:", font=ctk.CTkFont(size=11, weight="bold"), text_color="#172033").pack(side="left", padx=(8, 6), pady=5)
        ctk.CTkButton(filters, text="Aceptar seleccionado", height=26, width=138, fg_color="#1F7A55", hover_color="#176242", command=lambda: self.set_selected_decision("ACEPTAR")).pack(side="left", padx=(0, 8), pady=5)
        ctk.CTkButton(filters, text="Ignorar seleccionado", height=26, width=138, fg_color="#A33A3A", hover_color="#842D2D", command=lambda: self.set_selected_decision("IGNORAR")).pack(side="left", padx=(0, 8), pady=5)
        ctk.CTkButton(filters, text="Dejar pendiente", height=26, width=112, fg_color="#EEF3FA", hover_color="#E2EAF5", text_color="#2457A6", command=lambda: self.set_selected_decision("PENDIENTE")).pack(side="left", padx=(0, 8), pady=5)

        visual_headers = ["Arbol", "Buque", "Fecha", "Tonelaje", "Carga", "Puerto", "Importador", "Proveedor", "Origen", "Resumen / criterio / cambios"]
        visual = Sheet(frame, data=[], headers=visual_headers, show_x_scrollbar=True, show_y_scrollbar=True)
        visual.enable_bindings("single_select", "row_select", "column_width_resize", "arrowkeys", "copy", "right_click_popup_menu")
        visual.set_column_widths([140, 140, 90, 95, 160, 145, 170, 170, 145, 560])
        visual.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=6, pady=(0, 6))
        # Hoja visual disponible internamente; la navegación principal vuelve a ser el árbol por buque.
        self.block_visual_sheet = visual
        self.visual_row_keys = {}

        columns = ["seccion", "fecha", "tonelaje", "carga", "puerto", "importador", "proveedor", "accion"]
        tree = ttk.Treeview(frame, columns=columns, show="tree headings")
        tree.heading("#0", text="Buque / bloque / línea")
        tree.column("#0", width=330, minwidth=240, anchor="w", stretch=True)
        headings = {
            "seccion": "Sección",
            "fecha": "Fecha",
            "tonelaje": "Tonelaje",
            "carga": "Carga",
            "puerto": "Puerto",
            "importador": "Importador",
            "proveedor": "Proveedor",
            "accion": "Cambio / acción",
        }
        widths = {
            "seccion": 105,
            "fecha": 90,
            "tonelaje": 95,
            "carga": 190,
            "puerto": 165,
            "importador": 190,
            "proveedor": 190,
            "accion": 560,
        }
        for col in columns:
            tree.heading(col, text=headings[col])
            min_width = 240 if col == "accion" else 80
            tree.column(col, width=widths[col], minwidth=min_width, anchor="w", stretch=True)

        y_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        # La tabla jerárquica queda oculta: solo se mantiene como estructura interna
        # para compatibilidad con filtros/exportación, pero la vista útil es la hoja visual.
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)

        tree.tag_configure("block", background="#EAF1FB", font=("Segoe UI", 9, "bold"))
        tree.tag_configure("isaosa_header", background="#EEF3FA", font=("Segoe UI", 9, "bold"))
        tree.tag_configure("azul_header", background="#E8F1FF", font=("Segoe UI", 9, "bold"))
        tree.tag_configure("cambio_header", background="#FFF4DA", font=("Segoe UI", 9, "bold"))
        tree.tag_configure("nuevo", background="#FDECEC", foreground="#8A1F1F")
        tree.tag_configure("cambio", background="#FFE1E1", foreground="#8A1F1F", font=("Segoe UI", 9, "bold"))
        tree.tag_configure("revisar", background="#FCE7E7", foreground="#8A1F1F")
        tree.tag_configure("match_ok", background="#F4FBF6", foreground="#1E5A36")
        tree.tag_configure("linked", background="#F7FAFE", foreground="#263244")
        tree.tag_configure("isaosa_line", background="#E4F5E9", foreground="#155B2D")
        tree.tag_configure("azul_line", background="#E5F0FF", foreground="#174A8B")
        tree.tag_configure("criterio", background="#F6F8FB", foreground="#5B6472")
        tree.tag_configure("aceptado", background="#DFF2E6", foreground="#155B2D", font=("Segoe UI", 9, "bold"))
        tree.tag_configure("ignorado", background="#EFEFEF", foreground="#6B7280")
        return tree

    def autodetect_default_files(self):
        folder = Path(__file__).resolve().parent
        for path in folder.glob("*.xlsx"):
            name = path.name.upper()
            if ("BASE_ISAOSA" in name or "REPORTES BUQUES" in name) and not self.isaosa_path.get():
                self.isaosa_path.set(str(path))
            elif ("BUQUES_AZULES" in name or "POSICION DE BUQUES" in name) and not self.twin_path.get():
                self.twin_path.set(str(path))

    def select_file(self, variable):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm *.xls"), ("Todos los archivos", "*.*")],
        )
        if file_path:
            variable.set(file_path)

    def parse_sheets(self, value):
        return [item.strip() for item in value.split(",") if item.strip()]

    def clear_block_filters(self):
        self.block_filter_text.set("")
        self.block_filter_type.set("Todos")
        self.refresh_block_view()

    def set_block_tree_open(self, open_value):
        if not hasattr(self, "block_tree"):
            return

        def apply_open(item):
            self.block_tree.item(item, open=open_value)
            for child in self.block_tree.get_children(item):
                apply_open(child)

        for item in self.block_tree.get_children():
            apply_open(item)

    def reset_visual_text(self):
        self.visual_rows = []
        self.visual_highlights = []
        self.visual_row_keys = {}

    def finish_visual_text(self):
        if not hasattr(self, "block_visual_sheet"):
            return
        self.block_visual_sheet.set_sheet_data(self.visual_rows, reset_highlights=True)
        self.block_visual_sheet.set_column_widths([140, 140, 90, 95, 160, 145, 170, 170, 145, 560])
        grouped = {}
        for row, col, bg, fg in self.visual_highlights:
            grouped.setdefault((bg, fg), []).append((row, col))
        for (bg, fg), cells in grouped.items():
            self.block_visual_sheet.highlight_cells(cells=cells, bg=bg, fg=fg, redraw=False)
        self.block_visual_sheet.redraw()

    def visual_insert(self, text, tag=None):
        return

    def visual_line(self, label, row, changed_fields, base_tag):
        return [
            label,
            safe_display(row.get("vessel", "")),
            safe_display_field("arrival_date", row.get("arrival_date", "")),
            safe_display(row.get("tonnage", "")),
            safe_display(row.get("cargo", "")),
            safe_display(row.get("discharge_port", "")),
            safe_display(row.get("importer", "")),
            safe_display(row.get("supplier", "")),
            safe_display(row.get("loading_port", "")),
            "",
        ]

    def append_visual_block(self, ext_block, isa_block, pairs, explanation, block_index, block_kind):
        if self.visual_rows:
            gap_row = len(self.visual_rows)
            self.visual_rows.append(["", "", "", "", "", "", "", "", "", ""])
            for col in range(10):
                self.visual_highlights.append((gap_row, col, "#EEF3FA", "#EEF3FA"))

        header_row = len(self.visual_rows)
        self.visual_rows.append([
            f"BLOQUE {block_index}",
            ext_block["vessel"],
            block_date_range_text(ext_block),
            safe_display(block_total_tonnage(ext_block)),
            "",
            "",
            "",
            "",
            "",
            f"{block_kind.upper()} | {len(pairs)} líneas azules | CRITERIO DEL BLOQUE: {explanation}",
        ])
        self.visual_rows[header_row][9] = f"{block_kind.upper()} | {len(pairs)} lineas azules | CRITERIO DEL BLOQUE: {explanation}"
        for col in range(10):
            self.visual_highlights.append((header_row, col, "#2457A6", "#FFFFFF"))

        subheader_row = len(self.visual_rows)
        self.visual_rows.append(["", "", "", "", "", "", "", "", "", ""])
        for col in range(10):
            self.visual_highlights.append((subheader_row, col, "#FFFFFF", "#FFFFFF"))

        for pair_no, pair in enumerate(pairs, start=1):
            changed_fields = {safe_display(change.get("campo", "")) for change in pair["changes"]}
            decision_key = pair_decision_key(pair)
            decision = self.review_decisions.get(pair_decision_key(pair), "PENDIENTE" if pair["changes"] or pair["isaosa"] is None else "")
            line_note = pair["explanation"]
            if decision:
                line_note = f"DECISION: {decision} | {line_note}"
            decision_bg = "#CFEEDB" if decision == "ACEPTAR" else "#EFEFEF" if decision == "IGNORAR" else None
            decision_fg = "#155B2D" if decision == "ACEPTAR" else "#6B7280" if decision == "IGNORAR" else None
            pair_status = "SIN MATCH" if pair["isaosa"] is None else "CON CAMBIOS" if pair["changes"] else "SIN CAMBIOS"

            pair_header_row = len(self.visual_rows)
            self.visual_rows.append([
                f"LÍNEA {pair_no}",
                safe_display(pair["azul"].get("vessel", ext_block["vessel"])),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                f"{pair_status} | {line_note}",
            ])
            self.visual_rows[pair_header_row][0] = f"  LINEA {pair_no}"
            self.visual_row_keys[pair_header_row] = decision_key
            pair_header_bg = "#FFF4DA" if pair["changes"] else "#F4FBF6" if pair["isaosa"] is not None else "#FDECEC"
            pair_header_fg = "#8A1F1F" if pair["changes"] or pair["isaosa"] is None else "#155B2D"
            for col in range(10):
                self.visual_highlights.append((pair_header_row, col, pair_header_bg, pair_header_fg))

            if pair["isaosa"] is not None:
                isa_row_idx = len(self.visual_rows)
                isa_row = self.visual_line("    ISAOSA", pair["isaosa"], changed_fields, "isaosa")
                isa_row[9] = line_note
                self.visual_rows.append(isa_row)
                self.visual_row_keys[isa_row_idx] = decision_key
                for col in range(10):
                    self.visual_highlights.append((isa_row_idx, col, "#DFF2E6", "#155B2D"))
            else:
                isa_row_idx = len(self.visual_rows)
                self.visual_rows.append(["ISAOSA", "", "", "", "", "", "", "", "", "SIN LINEA EQUIVALENTE"])
                self.visual_rows[isa_row_idx][0] = "    ISAOSA"
                self.visual_row_keys[isa_row_idx] = decision_key
                for col in range(10):
                    self.visual_highlights.append((isa_row_idx, col, "#FFD7D7", "#8A1F1F"))

            azul_row_idx = len(self.visual_rows)
            azul_row = self.visual_line("    BUQUES AZUL", pair["azul"], changed_fields, "azul")
            if pair["changes"]:
                cambios = "; ".join(
                    f"{safe_display(change['campo'])}: {safe_display(change['valor_isaosa'])} -> {safe_display(change['valor_externo'])}"
                    for change in pair["changes"]
                )
                azul_row[9] = f"CAMBIOS: {cambios}"
            elif pair["isaosa"] is not None:
                azul_row[9] = "Sin diferencias relevantes en esta línea."
            self.visual_rows.append(azul_row)
            self.visual_row_keys[azul_row_idx] = decision_key
            for col in range(10):
                self.visual_highlights.append((azul_row_idx, col, "#DDEBFF", "#174A8B"))
            if decision_bg:
                self.visual_highlights.append((isa_row_idx, 9, decision_bg, decision_fg))
                self.visual_highlights.append((azul_row_idx, 9, decision_bg, decision_fg))

            field_to_col = {
                "arrival_date": 2,
                "tonnage": 3,
                "cargo": 4,
                "discharge_port": 5,
                "importer": 6,
                "supplier": 7,
                "loading_port": 8,
            }
            for field in changed_fields:
                col = field_to_col.get(field)
                if col is not None:
                    self.visual_highlights.append((isa_row_idx, col, "#FFD7D7", "#8A1F1F"))
                    self.visual_highlights.append((azul_row_idx, col, "#FFD7D7", "#8A1F1F"))

            spacer_row = len(self.visual_rows)
            self.visual_rows.append(["", "", "", "", "", "", "", "", "", ""])
            for col in range(10):
                self.visual_highlights.append((spacer_row, col, "#FAFBFD", "#FAFBFD"))

    def refresh_block_view(self):
        if not hasattr(self, "block_tree"):
            return
        tree = self.block_tree
        tree.delete(*tree.get_children())
        self.block_item_keys = {}
        self.reset_visual_text()
        if self.isaosa_df is None or self.external_df is None or self.external_df.empty:
            self.finish_visual_text()
            return

        text_filter = normalize_text(self.block_filter_text.get())
        type_filter = self.block_filter_type.get()
        isaosa_blocks = build_blocks(self.isaosa_df, "ISAOSA")
        external_blocks = build_blocks(self.external_df, "buques_azules")

        for block_index, ext_block in enumerate(external_blocks, start=1):
            related_results = self.results_for_block(ext_block)
            block_kind = self.classify_block_for_filter(ext_block, related_results)
            match = find_matching_block(isaosa_blocks, ext_block)
            isa_block = match[0] if match else None
            pairs = self.build_visual_line_pairs(isa_block, ext_block)
            if match and block_kind == "Nuevos":
                block_kind = "Con cambios"
            if text_filter and text_filter not in self.block_search_blob(ext_block, isa_block, related_results, pairs):
                continue
            if type_filter != "Todos" and block_kind != type_filter:
                continue

            status_text = block_kind
            match_score = match[1] if match else None
            explanation = explain_block_match(isa_block, ext_block, match_score) if match else explain_block_match(None, ext_block, {})
            self.append_visual_block(ext_block, isa_block, pairs, explanation, block_index, block_kind)
            continue
            parent = tree.insert(
                "",
                "end",
                text=f"{ext_block['vessel']} | {block_date_range_text(ext_block)}",
                values=(status_text, block_date_range_text(ext_block), safe_display(block_total_tonnage(ext_block)), "", "", "", "", f"{len(ext_block['rows'])} líneas azul"),
                tags=("block",),
                open=False,
            )
            tree.insert(
                parent,
                "end",
                text="Por qué eligió este bloque",
                values=("EXPLICACIÓN", "", "", "", "", "", "", explanation),
                tags=("linked" if match else "revisar",),
            )

            summary_text = "Bloque ISAOSA encontrado" if isa_block else "No se encontró bloque ISAOSA equivalente"
            tree.insert(
                parent,
                "end",
                text="Resumen del bloque",
                values=("RESUMEN", block_date_range_text(isa_block) if isa_block else "", safe_display(block_total_tonnage(isa_block)) if isa_block else "", "", "", "", "", summary_text),
                tags=("isaosa_header" if isa_block else "revisar",),
            )

            pairs_parent = tree.insert(parent, "end", text="Comparación línea por línea", values=("PARES", "", "", "", "", "", "", "ISAOSA arriba, azul abajo"), tags=("cambio_header",), open=True)
            for pair_no, pair in enumerate(pairs, start=1):
                decision_key = pair_decision_key(pair)
                decision = self.review_decisions.get(decision_key, "PENDIENTE" if pair["changes"] or pair["isaosa"] is None else "")
                pair_tag = "cambio" if pair["changes"] else "nuevo" if pair["isaosa"] is None else "match_ok"
                if decision == "ACEPTAR":
                    pair_tag = "aceptado"
                elif decision == "IGNORAR":
                    pair_tag = "ignorado"
                pair_action = "Cambió" if pair["changes"] else "No existe en ISAOSA" if pair["isaosa"] is None else "Sin cambios"
                if decision:
                    pair_action = f"{decision} | {pair_action}"
                pair_node = tree.insert(
                    pairs_parent,
                    "end",
                    text=f"Azul: {line_display_title(pair['azul'])}",
                    values=("LINK", "", "", "", "", "", "", pair_action),
                    tags=(pair_tag,),
                    open=bool(pair["changes"] or pair["isaosa"] is None),
                )
                self.block_item_keys[pair_node] = decision_key
                if pair["isaosa"] is not None:
                    self.insert_block_line(tree, pair_node, "BASE ISAOSA", "ISAOSA", pair["isaosa"], "isaosa_line")
                    tree.insert(
                        pair_node,
                        "end",
                        text="Criterio de emparejamiento",
                        values=("CRITERIO", "", "", "", "", "", "", pair["explanation"]),
                        tags=("criterio",),
                    )
                else:
                    tree.insert(pair_node, "end", text="BASE ISAOSA", values=("ISAOSA", "", "", "", "", "", "", "Sin línea equivalente"), tags=("revisar",))
                self.insert_block_line(tree, pair_node, "BUQUES AZULES", "AZUL", pair["azul"], "azul_line")
                for change in pair["changes"]:
                    tree.insert(
                        pair_node,
                        "end",
                        text=f"CAMPO DIFERENTE: {safe_display(change['campo'])}",
                        values=(
                            safe_display(change["campo"]),
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            f"{safe_display(change['valor_isaosa'])} -> {safe_display(change['valor_externo'])} | {safe_display(change['accion_sugerida'])}",
                        ),
                        tags=("cambio",),
                    )

        self.finish_visual_text()

    def insert_block_line(self, tree, parent, label, section, row, tag=""):
        tree.insert(
            parent,
            "end",
            text=label,
            values=(
                section,
                safe_display_field("arrival_date", row.get("arrival_date", "")),
                safe_display(row.get("tonnage", "")),
                safe_display(row.get("cargo", "")),
                safe_display(row.get("discharge_port", "")),
                safe_display(row.get("importer", "")),
                safe_display(row.get("supplier", "")),
                "",
            ),
            tags=(tag,) if tag else (),
        )

    def build_visual_line_pairs(self, isa_block, ext_block):
        return build_line_pairs_for_blocks(isa_block, ext_block)
        pairs = []
        isa_rows = isa_block["rows"] if isa_block else []
        ext_rows = ext_block["rows"]
        assignments = {}
        used_isa = set()

        all_scores = []
        for ext_idx, ext_row in enumerate(ext_rows):
            for isa_idx, isa_row in enumerate(isa_rows):
                all_scores.append((detail_line_score(isa_row, ext_row), ext_idx, isa_idx, isa_row, ext_row))
        all_scores.sort(reverse=True, key=lambda item: item[0])

        for score, ext_idx, isa_idx, isa_row, ext_row in all_scores:
            if score < 62:
                continue
            if ext_idx in assignments or isa_idx in used_isa:
                continue
            assignments[ext_idx] = (isa_idx, score, isa_row)
            used_isa.add(isa_idx)

        for ext_idx, ext_row in enumerate(ext_rows):
            assignment = assignments.get(ext_idx)
            if assignment is None:
                pairs.append({"isaosa": None, "azul": ext_row, "changes": [], "explanation": "No se encontró una línea ISAOSA suficientemente parecida dentro del bloque."})
                continue
            _, _, isa_row = assignment
            changes = detect_changes(isa_row, ext_row, {"total": 100})
            components = detail_line_components(isa_row, ext_row)
            pairs.append({
                "isaosa": isa_row,
                "azul": ext_row,
                "changes": changes,
                "explanation": explain_detail_line_match(isa_row, ext_row, components),
                "components": components,
            })
        return pairs

    def block_search_blob(self, ext_block, isa_block, related_results, pairs):
        chunks = [ext_block["vessel"], block_date_range_text(ext_block)]
        if isa_block:
            chunks.extend([isa_block["vessel"], block_date_range_text(isa_block)])
            for row in isa_block["rows"]:
                chunks.append(build_record_summary(row))
        for row in ext_block["rows"]:
            chunks.append(build_record_summary(row))
        for pair in pairs:
            for change in pair["changes"]:
                chunks.extend([safe_display(change.get("tipo", "")), safe_display(change.get("campo", "")), safe_display(change.get("accion_sugerida", ""))])
        if related_results is not None and not related_results.empty:
            chunks.extend(related_results.astype(str).agg(" ".join, axis=1).tolist())
        return normalize_text(" ".join(chunks))

    def results_for_block(self, ext_block):
        if self.results_df is None or self.results_df.empty:
            return pd.DataFrame(columns=RESULT_COLUMNS)
        vessels = {normalize_text(row.get("vessel", "")) for row in ext_block["rows"]}
        mask = self.results_df["buque_externo"].apply(lambda value: normalize_text(value) in vessels)
        return self.results_df[mask].copy()

    def classify_block_for_filter(self, ext_block, related_results):
        if related_results.empty:
            return "Sin cambios"
        tipos = set(related_results["tipo"].fillna("").tolist())
        matches = set(related_results["match"].fillna("").tolist())
        if tipos & {"NUEVO VIAJE", "NUEVO REGISTRO", "LINEA AZUL SIN MATCH"}:
            return "Nuevos"
        if "REVISAR" in matches:
            return "Revisar"
        return "Con cambios"

    def selected_decision_key(self):
        if hasattr(self, "block_visual_sheet"):
            try:
                selected_rows = list(self.block_visual_sheet.get_selected_rows())
                if selected_rows:
                    for row in selected_rows:
                        if row in self.visual_row_keys:
                            return self.visual_row_keys[row]
                selected = self.block_visual_sheet.get_currently_selected()
                if selected is not None:
                    row = selected.row if hasattr(selected, "row") else selected[0]
                    if row in self.visual_row_keys:
                        return self.visual_row_keys[row]
            except Exception:
                pass
        selected = self.block_tree.selection() if hasattr(self, "block_tree") else ()
        if not selected:
            return None
        item = selected[0]
        while item:
            if item in self.block_item_keys:
                return self.block_item_keys[item]
            item = self.block_tree.parent(item)
        return None

    def set_selected_decision(self, decision):
        key = self.selected_decision_key()
        if not key:
            messagebox.showinfo("Selecciona una fila", "Selecciona una fila verde o azul en la hoja visual de Vista por buque.")
            return
        self.review_decisions[key] = decision
        self.refresh_block_view()
        self.status_var.set(f"Cambio marcado como {decision}.")

    def run_comparison(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Proceso en curso", "La comparación todavía está corriendo.")
            return
        isaosa_file = self.isaosa_path.get().strip()
        twin_file = self.twin_path.get().strip()
        if not isaosa_file:
            messagebox.showerror("Falta archivo", "Debes seleccionar la base ISAOSA.")
            return
        if not twin_file:
            messagebox.showerror("Falta archivo externo", "Selecciona buques_azules.")
            return

        self.compare_button.configure(state="disabled")
        self.status_var.set("Comparando en segundo plano...")
        self.logs_df = pd.DataFrame(columns=LOG_COLUMNS)
        self.load_dataframe_to_tree(self.log_tree, self.logs_df)
        self.worker = threading.Thread(target=self._comparison_worker, daemon=True)
        self.worker.start()

    def add_log(self, etapa, detalle):
        row = {"momento": now_text(), "etapa": etapa, "detalle": detalle}
        self.logs_df = pd.concat([self.logs_df, pd.DataFrame([row])], ignore_index=True)
        if len(self.logs_df) <= 300 or len(self.logs_df) % 50 == 0:
            self.root.after(0, lambda: self.load_dataframe_to_tree(self.log_tree, self.logs_df.tail(1000)))

    def _comparison_worker(self):
        try:
            warnings = []
            metas = []
            self.add_log("Inicio", "Arrancó la comparación.")
            self.root.after(0, lambda: self.status_var.set("Leyendo base ISAOSA..."))
            self.add_log("Lectura", f"Leyendo ISAOSA: {self.isaosa_path.get().strip()} | hoja {self.isaosa_sheet.get().strip()}.")
            isaosa_result = load_isaosa(self.isaosa_path.get().strip(), self.isaosa_sheet.get().strip())
            warnings.extend(isaosa_result.warnings)
            if isaosa_result.data.empty:
                detail = "No se pudo leer base_isaosa o no contiene buques válidos."
                if isaosa_result.warnings:
                    detail = isaosa_result.warnings[-1].get("detalle", detail)
                raise RuntimeError(detail)
            if isaosa_result.meta:
                metas.append(isaosa_result.meta)
                self.add_log("Encabezado ISAOSA", f"Hoja {isaosa_result.meta.get('hoja')} | encabezado fila {isaosa_result.meta.get('encabezado_fila')} | registros {isaosa_result.meta.get('registros')}.")
            external_frames = []
            if self.twin_path.get().strip():
                self.root.after(0, lambda: self.status_var.set("Leyendo buques_azules..."))
                self.add_log("Lectura", f"Leyendo buques_azules: {self.twin_path.get().strip()} | hoja {self.twin_sheet.get().strip()}.")
                twin_result = load_twin(self.twin_path.get().strip(), self.twin_sheet.get().strip())
                warnings.extend(twin_result.warnings)
                if twin_result.meta:
                    metas.append(twin_result.meta)
                    self.add_log("Encabezado buques_azules", f"Hoja {twin_result.meta.get('hoja')} | encabezado fila {twin_result.meta.get('encabezado_fila')} | registros {twin_result.meta.get('registros')}.")
                if not twin_result.data.empty:
                    external_frames.append(twin_result.data)

            external_df = pd.concat(external_frames, ignore_index=True) if external_frames else pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS)
            if external_df.empty:
                detail = "No se pudo leer buques_azules o no contiene buques válidos."
                if warnings:
                    detail = warnings[-1].get("detalle", detail)
                raise RuntimeError(detail)
            self.root.after(0, lambda: self.status_var.set("Comparando registros..."))
            self.add_log("Comparación", "Se compara buques_azules contra base_isaosa para detectar nuevos registros y campos cambiados.")
            results_df = compare_all(
                isaosa_result.data,
                external_df,
                log_callback=self.add_log,
                include_unmatched_isaosa=self.include_unmatched_isaosa.get(),
                include_sin_cambios=True,
            )
            self.add_log("Fin", f"Comparación terminada con {len(results_df)} hallazgos y {len(warnings)} advertencias.")
            warnings_df = pd.DataFrame(warnings, columns=WARNING_COLUMNS)
            logs_df = self.logs_df.copy()
            self.root.after(0, lambda: self._comparison_done(results_df, isaosa_result.data, external_df, warnings_df, metas, logs_df))
        except Exception as exc:
            detail = traceback.format_exc()
            self.add_log("Error", detail)
            self.root.after(0, lambda: self._comparison_failed(exc, detail))

    def _comparison_done(self, results_df, isaosa_df, external_df, warnings_df, metas, logs_df):
        self.results_df = results_df
        self.isaosa_df = isaosa_df
        self.external_df = external_df
        self.warnings_df = warnings_df
        self.logs_df = logs_df
        self.meta_rows = metas
        self.load_dataframe_to_tree(self.log_tree, self.logs_df.tail(1000))
        self.refresh_block_view()
        self.compare_button.configure(state="normal")
        self.status_var.set(
            f"Listo. ISAOSA: {len(self.isaosa_df)} | Externos: {len(self.external_df)} | "
            f"Comparaciones: {len(self.results_df)} | Advertencias: {len(self.warnings_df)}"
        )
        messagebox.showinfo("Comparación terminada", "La comparación terminó correctamente.")

    def _comparison_failed(self, exc, detail):
        print(detail)
        self.compare_button.configure(state="normal")
        self.status_var.set("Error en la comparación.")
        messagebox.showerror("Error", f"Ocurrió un error:\n\n{exc}")

    def load_dataframe_to_tree(self, tree, df):
        tree.delete(*tree.get_children())
        if df is None or df.empty:
            tree["columns"] = []
            return
        display_df = df.head(1000).copy()
        columns = list(display_df.columns)
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=160, minwidth=95, anchor="w")
        for _, row in display_df.iterrows():
            tree.insert("", "end", values=[safe_display(row[col]) for col in columns])

    def export_results(self):
        if self.results_df is None or self.results_df.empty:
            messagebox.showwarning("Sin resultados", "Primero compara archivos antes de exportar.")
            return
        output_path = filedialog.asksaveasfilename(
            title="Guardar reporte",
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
            initialfile="reporte_comparacion_buques.xlsx",
        )
        if not output_path:
            return
        try:
            export_report(output_path, self.results_df, self.isaosa_df, self.external_df, self.warnings_df, self.logs_df, self.review_decisions)
            messagebox.showinfo("Reporte guardado", f"El reporte se guardó correctamente en:\n\n{output_path}")
        except Exception as exc:
            messagebox.showerror("Error al exportar", f"No se pudo exportar el reporte:\n\n{exc}")

    def clear_results(self):
        self.results_df = pd.DataFrame(columns=RESULT_COLUMNS)
        self.isaosa_df = pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS)
        self.external_df = pd.DataFrame(columns=DISPLAY_COLUMNS + WORK_COLUMNS)
        self.warnings_df = pd.DataFrame(columns=WARNING_COLUMNS)
        self.logs_df = pd.DataFrame(columns=LOG_COLUMNS)
        self.review_decisions = {}
        self.block_item_keys = {}
        self.load_dataframe_to_tree(self.log_tree, pd.DataFrame())
        self.block_tree.delete(*self.block_tree.get_children())
        self.reset_visual_text()
        self.finish_visual_text()
        self.status_var.set("Resultados limpiados.")


def main():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TFrame", background="#F6F8FB")
    style.configure("TLabel", background="#F6F8FB", foreground="#1F2937", font=("Segoe UI", 10))
    style.configure("TLabelframe", background="#F6F8FB", foreground="#1F2937")
    style.configure("TLabelframe.Label", background="#F6F8FB", foreground="#1F2937", font=("Segoe UI", 10, "bold"))
    style.configure("TButton", font=("Segoe UI", 10), padding=(10, 5))
    style.configure("TCheckbutton", background="#F6F8FB", foreground="#1F2937", font=("Segoe UI", 10))
    style.configure("Treeview", rowheight=24, font=("Segoe UI", 9))
    style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E8EEF7")
    BuquesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
