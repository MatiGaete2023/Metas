#!/usr/bin/env python3
"""Genera Registros.xlsx, META 4 y un reporte de procesamiento desde una carpeta."""

from __future__ import annotations

import argparse
import calendar
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


APP_VERSION = "2.4"

TRIBUNALS = [
    "ARAUCO",
    "CABRERO",
    "CAÑETE",
    "CONCEPCIÓN",
    "CORONEL",
    "CURANILAHUE",
    "FLORIDA",
    "LOS ANGELES",
    "LAJA",
    "LEBU",
    "MULCHÉN",
    "NACIMIENTO",
    "SANTA BÁRBARA",
    "SANTA JUANA",
    "TALCAHUANO",
    "TOMÉ",
    "YUMBEL",
]

TRIBUNAL_PATTERNS = [
    ("SANTA BARBARA", "SANTA BÁRBARA"),
    ("SANTA JUANA", "SANTA JUANA"),
    ("LOS ANGELES", "LOS ANGELES"),
    ("TALCAHUANO", "TALCAHUANO"),
    ("CONCEPCION", "CONCEPCIÓN"),
    ("CONCEPCIN", "CONCEPCIÓN"),
    ("CURANILAHUE", "CURANILAHUE"),
    ("NACIMIENTO", "NACIMIENTO"),
    ("CABRERO", "CABRERO"),
    ("CANETE", "CAÑETE"),
    ("CAETE", "CAÑETE"),
    ("CORONEL", "CORONEL"),
    ("FLORIDA", "FLORIDA"),
    ("ARAUCO", "ARAUCO"),
    ("MULCHEN", "MULCHÉN"),
    ("LAJA", "LAJA"),
    ("LEBU", "LEBU"),
    ("TOME", "TOMÉ"),
    ("YUMBEL", "YUMBEL"),
]

# Alias aceptados. AMBIGUOUS_ALIASES exige validación de contexto porque la
# abreviatura coincide con una palabra común del español.
TRIBUNAL_ALIASES = {
    "THNO": "TALCAHUANO",
    "CCP": "CONCEPCIÓN",
    "LA": "LOS ANGELES",
    "LOSA": "LOS ANGELES",
    "LGS": "LOS ANGELES",
}

AMBIGUOUS_ALIASES = {"LA"}

# Si el alias ambiguo viene precedido por uno de estos términos, es un artículo.
ALIAS_CONNECTORS = {
    "DE", "DEL", "EN", "A", "AL", "CON", "POR", "PARA", "DESDE", "SOBRE",
    "HASTA", "ENTRE", "SEGUN", "Y", "O", "U", "ANTE", "BAJO", "TRAS",
    "TODA", "TODAS", "ESTA", "ESA", "AQUELLA", "SOLO", "SOLAMENTE",
}

# Si el alias ambiguo va seguido por uno de estos términos (o por nada, o por un
# número), es una abreviatura de tribunal.
ALIAS_NEIGHBOURS = {
    "FAE", "FAES", "FAS", "RES", "RESIDENCIA", "RESIDENCIAS", "CREAD",
    "ESPERA", "ESP", "CUMPLIMIENTO", "CUMPLIMEINTO", "CUMPL", "CUMP", "CUM",
    "INFORME", "INFORMES", "JF", "JUZGADO", "FAMILIA",
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
    "AGOSTO", "SEPTIEMBRE", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
    "XLS", "XLSX", "XLSM", "PDF", "PNG", "JPG", "JPEG",
}

ESPERA_HEADERS = [
    "RIT",
    "TRIBUNAL",
    "RUT",
    "NOMBRE",
    "NACIONALIDAD / PAIS",
    "TIPO",
    "SEXO",
    "DERIVACIÓN",
    "CURADOR",
    "T ESPERA",
    "FEC. RESOLUCIÓN",
    "REVISIÓN",
    "DURACIÓN",
    "80 BIS",
    "FEC. NACIMIENTO",
    "EDAD",
    "FEC. ACT. F. INDIVIDUAL",
    "FEC. ACT. F. AMBULATORIA",
    "FEC. ACT. F. FAE/FAS",
    "FEC. OÍDO",
    "PROXS. AUDS.",
    "FECHA OBS.",
    "ARCHIVO ORIGEN",
    "RUTA ORIGEN",
]

CUMPLIMIENTO_HEADERS = [
    "RIT",
    "TRIBUNAL",
    "RUT",
    "NOMBRE",
    "NACIONALIDAD / PAÍS",
    "TIPO",
    "SEXO",
    "DERIVACIÓN",
    "CURADOR",
    "FEC. RESOLUCIÓN",
    "FEC. INGRESO EFECTIVO",
    "FEC. EGRESO PROYECTADO",
    "DÍAS DE CUMPLIMIENTO",
    "DÍAS PARA EGRESAR",
    "80 BIS",
    "FEC. NACIMIENTO",
    "EDAD",
    "FEC. ACT. F. INDIVIDUAL",
    "FEC. ACT. F. AMBULATORIA",
    "FEC. ACT. F. FAE/FAS",
    "FEC. OÍDO",
    "FECHA OBS.",
    "ARCHIVO ORIGEN",
    "RUTA ORIGEN",
]

SOURCE_HEADERS = {
    "Espera": ESPERA_HEADERS,
    "Cumplimiento": CUMPLIMIENTO_HEADERS,
}

SHEET_DEFS = {
    "Residencia Espera": {
        "source": "ESPERA",
        "category": "Residencia",
        "report": "Informe de Ingresos en Lista de Espera",
        "headers": [
            "RIT",
            "TRIBUNAL",
            "TIPO",
            "SEXO",
            "DERIVACIÓN",
            "T ESPERA",
            "FEC. RESOLUCIÓN",
            "FECHA OBS.",
        ],
    },
    "Residencia Cumplimiento": {
        "source": "CUMPLIMIENTO",
        "category": "Residencia",
        "report": "Informe de Ingresos en Lista de Cumplimiento",
        "headers": [
            "RIT",
            "TRIBUNAL",
            "TIPO",
            "SEXO",
            "DERIVACIÓN",
            "FEC. RESOLUCIÓN",
            "FEC. INGRESO EFECTIVO",
            "FECHA OBS.",
        ],
    },
    "FAE Espera": {
        "source": "ESPERA",
        "category": "FAE",
        "report": "Informe de Ingresos en Lista de Espera",
        "headers": [
            "RIT",
            "TRIBUNAL",
            "TIPO",
            "SEXO",
            "DERIVACIÓN",
            "T ESPERA",
            "FEC. RESOLUCIÓN",
            "FECHA OBS.",
        ],
    },
    "FAE Cumplimiento": {
        "source": "CUMPLIMIENTO",
        "category": "FAE",
        "report": "Informe de Ingresos en Lista de Cumplimiento",
        "headers": [
            "RIT",
            "TRIBUNAL",
            "TIPO",
            "SEXO",
            "DERIVACIÓN",
            "FEC. RESOLUCIÓN",
            "FEC. INGRESO EFECTIVO",
            "FECHA OBS.",
        ],
    },
}

EXCEL_SUFFIXES = {".xls", ".xlsx", ".xlsm"}
PDF_SUFFIXES = {".pdf"}
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".gif", ".webp"}

DATE_HEADERS = {
    "FEC RESOLUCION",
    "FEC INGRESO EFECTIVO",
    "FEC EGRESO PROYECTADO",
    "FEC NACIMIENTO",
    "FEC ACT F INDIVIDUAL",
    "FEC ACT F AMBULATORIA",
    "FEC ACT F FAE FAS",
    "FEC OIDO",
    "FECHA OBS",
}


@dataclass
class Issue:
    level: str
    file: str
    combination: str
    message: str


@dataclass
class SourceRow:
    status: str
    category: str | None
    tribunal: str | None
    values: dict[str, Any]
    source_file: str
    source_path: str
    excel_row: int


@dataclass
class Candidate:
    file_result: "FileResult"
    tribunal: str
    category: str
    status: str
    kind: str
    rows: list[SourceRow] = field(default_factory=list)
    emission_date: dt.date | None = None
    modified_at: float = 0.0
    selected: bool = False
    selection_note: str = ""

    @property
    def key(self) -> tuple[str, str, str]:
        return self.tribunal, self.category, self.status


@dataclass
class FileResult:
    path: Path
    relative_path: str
    kind: str
    result: str
    tribunal: str | None = None
    status: str | None = None
    categories: list[str] = field(default_factory=list)
    record_count: int = 0
    emission_date: dt.date | None = None
    sheet_name: str = ""
    header_row: int | None = None
    observations: list[str] = field(default_factory=list)
    candidates: list[Candidate] = field(default_factory=list)
    orphan_rows: list[SourceRow] = field(default_factory=list)

    def selected_label(self) -> str:
        if not self.candidates:
            return "N/A"
        selected = sum(1 for candidate in self.candidates if candidate.selected)
        if selected == len(self.candidates):
            return "SÍ"
        if selected:
            return "PARCIAL"
        return "NO"


def norm(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).replace("\ufffd", "")
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Z0-9]+", " ", text.upper()).strip()


def clean_text(value: Any) -> Any:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, str):
        return re.sub(r"\s+", " ", value).strip()
    return value


def identifier_text(value: Any) -> str | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(clean_text(value))


def excel_value(value: Any) -> Any:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if hasattr(value, "item") and not isinstance(value, (str, bytes, dt.date, dt.datetime)):
        try:
            value = value.item()
        except (TypeError, ValueError):
            pass
    return clean_text(value)


def classify_derivation(value: Any) -> str | None:
    text = norm(value)
    if text.startswith("FAE"):
        return "FAE"
    if text.startswith("R") or text.startswith("CREAD"):
        return "Residencia"
    return None


def segment_tribunals(value: Any) -> tuple[set[str], list[str]]:
    """Devuelve (tribunales hallados, notas) para UN segmento de texto.

    Prioridad: nombre completo > alias. Un alias solo se acepta si el contexto
    descarta que sea una palabra común (ver ALIAS_CONNECTORS/ALIAS_NEIGHBOURS).
    """
    raw = re.sub(r"\bL\.\s?A\.?", "LA", str(value or ""), flags=re.IGNORECASE)
    text = norm(raw)
    if not text:
        return set(), []
    found: set[str] = set()
    for pattern, label in TRIBUNAL_PATTERNS:
        if re.search(rf"(?<![A-Z]){re.escape(pattern)}(?![A-Z])", text):
            found.add(label)
    # L-01: antes se retornaba aquí apenas un nombre completo coincidía, sin
    # revisar alias. Un segmento con un nombre completo Y un alias de OTRO
    # tribunal (p. ej. "Tome CCP") quedaba resuelto solo al nombre completo,
    # ignorando en silencio la mención del alias. Los alias se acumulan en el
    # mismo conjunto `found`; la ambigüedad (len(found) > 1) la resuelve el
    # llamador exactamente igual que con dos nombres completos.

    notes: list[str] = []
    tokens = text.split()
    for index, token in enumerate(tokens):
        label = TRIBUNAL_ALIASES.get(token)
        if not label:
            continue
        if token not in AMBIGUOUS_ALIASES:
            found.add(label)
            continue
        previous = tokens[index - 1] if index else None
        following = tokens[index + 1] if index + 1 < len(tokens) else None
        if previous in ALIAS_CONNECTORS:
            continue
        if following is None or following in ALIAS_NEIGHBOURS or following.isdigit():
            found.add(label)
        else:
            notes.append(
                f"Se ignoró el alias '{token}' en '{str(value)}' porque el contexto no permite "
                f"distinguirlo de una palabra común; renombra el archivo para usarlo como {label}"
            )
    return found, notes


def tribunal_name(value: Any) -> str | None:
    """Clasificación para el contenido de una celda (un solo valor)."""
    found, _ = segment_tribunals(value)
    if len(found) == 1:
        return next(iter(found))
    return None


def path_segments(value: Any) -> list[str]:
    text = str(value or "").replace("\\", "/")
    return [segment for segment in text.split("/") if segment.strip()]


def tribunal_from_path(value: Any) -> tuple[str | None, list[str]]:
    """Clasificación por ruta: segmento más profundo primero, ambigüedad = None."""
    notes: list[str] = []
    for segment in reversed(path_segments(value)):
        found, segment_notes = segment_tribunals(segment)
        notes.extend(segment_notes)
        if len(found) == 1:
            return next(iter(found)), notes
        if len(found) > 1:
            notes.append(
                f"El segmento '{segment}' menciona más de un tribunal ({', '.join(sorted(found))}); "
                "no se clasificó por ruta"
            )
            return None, notes
    return None, notes


def status_from_title(value: Any) -> str | None:
    text = norm(value)
    if "LISTA DE CUMPLIMIENTO" in text:
        return "Cumplimiento"
    if "LISTA DE ESPERA" in text:
        return "Espera"
    return None


def status_from_headers(headers: Iterable[Any]) -> str | None:
    values = {norm(value) for value in headers if norm(value)}
    if "FEC INGRESO EFECTIVO" in values or "DIAS DE CUMPLIMIENTO" in values:
        return "Cumplimiento"
    if "T ESPERA" in values:
        return "Espera"
    return None


def _status_candidates(segment: str) -> set[str]:
    text = norm(segment)
    tokens = set(text.split())
    candidates: set[str] = set()
    if (
        "CUMPLIMIENTO" in text
        or "CUMPLIMEINTO" in text
        or any(token.startswith("CUMPL") for token in tokens)
        or bool(tokens.intersection({"CUMP", "CUM"}))
    ):
        candidates.add("Cumplimiento")
    if "ESPERA" in text or "ESP" in tokens:
        candidates.add("Espera")
    return candidates


def status_from_path(value: Any) -> str | None:
    """Evalúa primero el nombre del archivo y luego las carpetas hacia arriba,
    igual que tribunal_from_path (M-01). Antes se evaluaba toda la ruta como un
    solo texto: una carpeta padre con la señal contraria (ESPERA/cumpl.pdf)
    anulaba un nombre de archivo inequívoco.

    Un segmento con ambas señales a la vez es ambiguo y detiene la búsqueda en
    ese punto (no se sigue subiendo a una carpeta más lejana a adivinar)."""
    for segment in reversed(path_segments(value)):
        candidates = _status_candidates(segment)
        if len(candidates) == 1:
            return next(iter(candidates))
        if len(candidates) > 1:
            return None
    return None


def _category_candidates(segment: str) -> set[str]:
    tokens = set(norm(segment).split())
    candidates: set[str] = set()
    if tokens.intersection({"FAE", "FAES"}):
        candidates.add("FAE")
    if any(token.startswith("RESIDENCIA") for token in tokens) or tokens.intersection({"RES", "CREAD"}):
        candidates.add("Residencia")
    return candidates


def category_from_path(value: Any) -> str | None:
    """Misma prioridad por segmento que status_from_path (M-01)."""
    for segment in reversed(path_segments(value)):
        candidates = _category_candidates(segment)
        if len(candidates) == 1:
            return next(iter(candidates))
        if len(candidates) > 1:
            return None
    return None


def as_date(value: Any) -> dt.date | Any:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().date()
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value > 20000:
            return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))
        return value
    text = str(value).strip()
    if not text or text == "---":
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if pd.notna(parsed):
        return parsed.to_pydatetime().date()
    return clean_text(value)


def parse_emission_value(value: Any) -> dt.date | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (dt.date, dt.datetime, pd.Timestamp)):
        parsed = as_date(value)
        return parsed if isinstance(parsed, dt.date) else None
    text = str(value)
    match = re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{4}|\d{4}-\d{1,2}-\d{1,2})", text)
    if not match:
        return None
    parsed = as_date(match.group(1))
    return parsed if isinstance(parsed, dt.date) else None


def as_wait_days(value: Any) -> Any:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip().replace(",", ".")
    try:
        number = float(text)
        return int(number) if number.is_integer() else number
    except ValueError:
        return clean_text(value)


def easter_sunday(year: int) -> dt.date:
    """Algoritmo gregoriano de Meeus/Jones/Butcher."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    ell = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * ell) // 451
    month = (h + ell - 7 * m + 114) // 31
    day = ((h + ell - 7 * m + 114) % 31) + 1
    return dt.date(year, month, day)


def mondayized(date_value: dt.date) -> dt.date:
    weekday = date_value.weekday()
    if weekday in (1, 2, 3):
        return date_value - dt.timedelta(days=weekday)
    if weekday == 4:
        return date_value + dt.timedelta(days=3)
    return date_value


# Día Nacional de los Pueblos Indígenas (Ley 21.357): solsticio de invierno,
# 20 o 21 de junio según el año. Solo se usa si la librería holidays no está
# instalada. Para años fuera de la tabla se asume el 21 y debe confirmarse en
# assets/feriados_adicionales.txt.
SOLSTICE_DAY = {
    2024: 20, 2025: 20, 2026: 21, 2027: 21, 2028: 20,
    2029: 20, 2030: 21, 2031: 21, 2032: 20, 2033: 20,
}


def fallback_chile_holidays(year: int) -> set[dt.date]:
    easter = easter_sunday(year)
    dates = {
        dt.date(year, 1, 1),
        easter - dt.timedelta(days=2),
        easter - dt.timedelta(days=1),
        dt.date(year, 5, 1),
        dt.date(year, 5, 21),
        dt.date(year, 6, SOLSTICE_DAY.get(year, 21)),
        mondayized(dt.date(year, 6, 29)),
        dt.date(year, 7, 16),
        dt.date(year, 8, 15),
        dt.date(year, 9, 18),
        dt.date(year, 9, 19),
        mondayized(dt.date(year, 10, 12)),
        dt.date(year, 10, 31),
        dt.date(year, 11, 1),
        dt.date(year, 12, 8),
        dt.date(year, 12, 25),
    }
    if dt.date(year, 9, 18).weekday() == 1:
        dates.add(dt.date(year, 9, 17))
    if dt.date(year, 9, 19).weekday() == 3:
        dates.add(dt.date(year, 9, 20))
    return dates


def additional_holidays(base_dir: Path, year: int) -> set[dt.date]:
    path = base_dir / "assets" / "feriados_adicionales.txt"
    result: set[dt.date] = set()
    if not path.exists():
        return result
    try:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    except OSError:
        return result
    for line in lines:
        value = line.split("#", 1)[0].strip()
        if not value:
            continue
        try:
            parsed = dt.datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            continue
        if parsed.year == year:
            result.add(parsed)
    return result


def chile_holidays(year: int, base_dir: Path) -> set[dt.date]:
    try:
        import holidays  # type: ignore

        try:
            generated = holidays.country_holidays("CL", subdiv="BI", years=year)
        except (KeyError, NotImplementedError, ValueError):
            generated = holidays.country_holidays("CL", years=year)
        result = set(generated.keys())
    except (ImportError, AttributeError):
        result = fallback_chile_holidays(year)
    return result | additional_holidays(base_dir, year)


def first_business_day(year: int, month: int, base_dir: Path) -> dt.date:
    day = dt.date(year, month, 1)
    holidays_for_year = chile_holidays(year, base_dir)
    while day.weekday() >= 5 or day in holidays_for_year:
        day += dt.timedelta(days=1)
    return day


def parse_emission_date(args: argparse.Namespace, base_dir: Path) -> dt.date:
    if args.emission_date:
        try:
            return dt.datetime.strptime(args.emission_date.strip(), "%d-%m-%Y").date()
        except ValueError:
            print(
                f"--emission-date inválido: {args.emission_date!r}. Formato esperado DD-MM-AAAA, "
                "por ejemplo 03-08-2026",
                file=sys.stderr,
            )
            raise SystemExit(2)
    if args.month:
        match = re.fullmatch(r"\s*(\d{4})-(\d{1,2})\s*", args.month)
        if not match:
            print(
                f"--month inválido: {args.month!r}. Formato esperado AAAA-MM, por ejemplo 2026-08",
                file=sys.stderr,
            )
            raise SystemExit(2)
        year, month = int(match.group(1)), int(match.group(2))
        if not 1 <= month <= 12:
            print(
                f"--month inválido: el mes debe estar entre 01 y 12, se recibió {month:02d}",
                file=sys.stderr,
            )
            raise SystemExit(2)
    else:
        today = dt.date.today()
        year, month = today.year, today.month
    return first_business_day(year, month, base_dir)


@dataclass
class FolderSelection:
    """Resultado de un intento de mostrar el selector de carpeta.

    NO_DISPONIBLE: el diálogo no pudo mostrarse en este equipo (falta tkinter,
    no es Windows, PowerShell falló). Es el ÚNICO estado que debe activar un
    respaldo.
    CANCELADO: el diálogo se mostró y el usuario lo cerró sin elegir carpeta.
    No debe abrir un segundo diálogo (defecto M-02: antes, cancelar tkinter
    abría igual el respaldo de PowerShell).
    SELECCIONADO: hay una carpeta elegida, en `path`.
    """

    status: str
    path: str = ""


def _folder_dialog_tkinter() -> FolderSelection:
    """Selector nativo de Python. No depende de PowerShell ni de directivas de grupo."""
    try:
        import tkinter
        from tkinter import filedialog
    except ImportError:
        return FolderSelection("NO_DISPONIBLE")
    try:
        root = tkinter.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
        except tkinter.TclError:
            pass
        selection = filedialog.askdirectory(
            title="Seleccione la carpeta que contiene los informes",
            mustexist=True,
        )
        root.destroy()
    except Exception:
        return FolderSelection("NO_DISPONIBLE")
    if selection:
        return FolderSelection("SELECCIONADO", selection)
    return FolderSelection("CANCELADO")


def _folder_dialog_powershell() -> FolderSelection:
    """Respaldo para equipos sin tkinter. Escribe la ruta en UTF-8 a un temporal
    para no depender de la página de códigos de la consola (tildes y ñ)."""
    if os.name != "nt":
        return FolderSelection("NO_DISPONIBLE")
    handle = tempfile.NamedTemporaryFile(prefix="meta4_folder_", suffix=".txt", delete=False)
    handle.close()
    temporary = Path(handle.name)
    marker = temporary.with_suffix(".ok")
    escaped = str(temporary).replace("'", "''")
    escaped_marker = str(marker).replace("'", "''")
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$dialog = New-Object System.Windows.Forms.FolderBrowserDialog;"
        "$dialog.Description = 'Seleccione la carpeta que contiene los informes';"
        "$dialog.ShowNewFolderButton = $false;"
        "if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {"
        f"[System.IO.File]::WriteAllText('{escaped}', $dialog.SelectedPath, "
        "(New-Object System.Text.UTF8Encoding($false)));"
        f"[System.IO.File]::WriteAllText('{escaped_marker}', 'OK') }}"
    )
    try:
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-STA", "-Command", script],
            capture_output=True,
            timeout=600,
            check=False,
        )
        if completed.returncode != 0:
            return FolderSelection("NO_DISPONIBLE")
        if not marker.exists():
            # El diálogo se mostró y se cerró sin marcar OK: el usuario canceló.
            return FolderSelection("CANCELADO")
        selection = temporary.read_text(encoding="utf-8").strip()
    except (OSError, subprocess.SubprocessError, UnicodeDecodeError):
        return FolderSelection("NO_DISPONIBLE")
    finally:
        for candidate in (temporary, marker):
            try:
                candidate.unlink()
            except OSError:
                pass
    if selection:
        return FolderSelection("SELECCIONADO", selection)
    return FolderSelection("CANCELADO")


def select_source_folder() -> FolderSelection:
    """Muestra el primer selector disponible. Solo NO_DISPONIBLE prueba el
    siguiente; CANCELADO se respeta de inmediato y no abre un segundo diálogo."""
    result = _folder_dialog_tkinter()
    if result.status != "NO_DISPONIBLE":
        return result
    return _folder_dialog_powershell()


def expected_keys() -> list[tuple[str, str, str]]:
    return [
        (tribunal, category, status)
        for tribunal in TRIBUNALS
        for category in ("FAE", "Residencia")
        for status in ("Espera", "Cumplimiento")
    ]


def combination_text(key: tuple[str, str, str] | None) -> str:
    return " | ".join(key) if key else ""


def find_header_row(df: pd.DataFrame) -> int | None:
    for row_idx in range(min(len(df.index), 30)):
        values = {norm(value) for value in df.iloc[row_idx] if norm(value)}
        if "RIT" in values and "TRIBUNAL" in values and any(value.startswith("DERIV") for value in values):
            return row_idx
    return None


def header_index(headers: list[Any], wanted: str) -> int | None:
    normalized = [norm(value) for value in headers]
    target = norm(wanted)
    if target in normalized:
        return normalized.index(target)
    if target == "DERIVACION":
        return next((idx for idx, value in enumerate(normalized) if value.startswith("DERIV")), None)
    if target == "FEC RESOLUCION":
        return next((idx for idx, value in enumerate(normalized) if value.startswith("FEC RESOLUC")), None)
    if target == "FECHA OBS":
        return next((idx for idx, value in enumerate(normalized) if value.startswith("FECHA OBS")), None)
    return None


def output_file_reason(path: Path) -> str | None:
    name = norm(path.name)
    if path.name.startswith("~$"):
        return "Archivo temporal de Excel"
    if name == "REGISTROS XLSX":
        return "Salida Registros.xlsx de una ejecución anterior"
    if name.startswith("META 4 ACTUALIZADO"):
        return "Salida META 4 de una ejecución anterior"
    if name.startswith("REPORTE DE PROCESAMIENTO META 4"):
        return "Reporte de una ejecución anterior"
    if name == "META4 TEMPLATE XLSX":
        return "Plantilla interna"
    return None


def relative_string(path: Path, root: Path) -> str:
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def inspect_excel(
    path: Path,
    root: Path,
    readable_path: Path,
    issues: list[Issue],
) -> FileResult:
    relative = relative_string(path, root)
    result = FileResult(path=path, relative_path=relative, kind="EXCEL", result="PROCESADO")
    try:
        with pd.ExcelFile(readable_path) as xls:
            sheets: list[tuple[tuple[int, int], str, pd.DataFrame, int]] = []
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=object)
                header_row = find_header_row(df)
                if header_row is None:
                    continue
                headers = list(df.iloc[header_row])
                rit_idx = header_index(headers, "RIT")
                record_count = 0
                if rit_idx is not None:
                    data = df.iloc[header_row + 1 :]
                    mask = data.iloc[:, rit_idx].notna() & data.iloc[:, rit_idx].astype(str).str.strip().ne("")
                    record_count = int(mask.sum())
                header_score = sum(1 for value in headers if norm(value))
                sheets.append(((record_count, header_score), sheet_name, df, header_row))
    except Exception as exc:
        result.result = "ERROR"
        message = f"No se pudo abrir como Excel: {type(exc).__name__}: {exc}"
        result.observations.append(message)
        issues.append(Issue("ERROR", relative, "", message))
        return result

    if not sheets:
        result.result = "ERROR"
        message = "No se encontró una hoja con encabezados RIT, TRIBUNAL y DERIVACIÓN"
        result.observations.append(message)
        issues.append(Issue("ERROR", relative, "", message))
        return result

    sheets.sort(key=lambda item: item[0], reverse=True)
    _, sheet_name, df, header_row = sheets[0]
    result.sheet_name = sheet_name
    result.header_row = header_row + 1
    if len(sheets) > 1:
        message = f"Se encontraron {len(sheets)} hojas candidatas; se utilizó {sheet_name}"
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))

    headers = list(df.iloc[header_row])
    rit_idx = header_index(headers, "RIT")
    tri_idx = header_index(headers, "TRIBUNAL")
    der_idx = header_index(headers, "DERIVACIÓN")
    if rit_idx is None or tri_idx is None or der_idx is None:
        result.result = "ERROR"
        message = "La hoja seleccionada no contiene las columnas mínimas"
        result.observations.append(message)
        issues.append(Issue("ERROR", relative, "", message))
        return result

    title = df.iat[1, 0] if len(df.index) > 1 else None
    title_status = status_from_title(title)
    column_status = status_from_headers(headers)
    path_status = status_from_path(relative)
    status = title_status or column_status or path_status
    if status is None:
        result.result = "ERROR"
        message = "No se pudo determinar Espera o Cumplimiento"
        result.observations.append(message)
        issues.append(Issue("ERROR", relative, "", message))
        return result
    result.status = status
    if title_status and column_status and title_status != column_status:
        message = f"El título indica {title_status}, pero las columnas indican {column_status}; prevaleció el título"
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))
    elif not title_status:
        message = f"El estado se determinó mediante las columnas: {status}"
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))

    a1 = df.iat[0, 0] if len(df.index) else None
    a1_tribunal = tribunal_name(a1)
    path_tribunal, path_notes = tribunal_from_path(relative)
    for note in path_notes:
        result.observations.append(note)
        issues.append(Issue("ADVERTENCIA", relative, "", note))
    result.emission_date = parse_emission_value(df.iat[2, 0] if len(df.index) > 2 else None)

    data = df.iloc[header_row + 1 :].copy()
    mask = data.iloc[:, rit_idx].notna() & data.iloc[:, rit_idx].astype(str).str.strip().ne("")
    data = data[mask]
    result.record_count = len(data)

    row_tribunals: Counter[str] = Counter()
    category_counts: Counter[str] = Counter()
    rows: list[SourceRow] = []
    canonical_headers = SOURCE_HEADERS[status]
    column_positions = {header: header_index(headers, header) for header in canonical_headers}

    required = ["RIT", "TRIBUNAL", "TIPO", "SEXO", "DERIVACIÓN", "FEC. RESOLUCIÓN"]
    if status == "Espera":
        required.append("T ESPERA")
    else:
        required.append("FEC. INGRESO EFECTIVO")
    missing_columns = [header for header in required if column_positions.get(header) is None]
    if missing_columns:
        message = "Columnas requeridas ausentes: " + ", ".join(missing_columns)
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))

    for df_idx, row in data.iterrows():
        original_tribunal = row.iloc[tri_idx] if tri_idx < len(row) else None
        row_tribunal = tribunal_name(original_tribunal)
        if row_tribunal:
            row_tribunals[row_tribunal] += 1
        original_derivation = row.iloc[der_idx] if der_idx < len(row) else None
        category = classify_derivation(original_derivation)
        if category:
            category_counts[category] += 1

        values: dict[str, Any] = {}
        for header in canonical_headers:
            if header == "ARCHIVO ORIGEN":
                values[header] = path.name
                continue
            if header == "RUTA ORIGEN":
                values[header] = relative
                continue
            col_idx = column_positions.get(header)
            value = row.iloc[col_idx] if col_idx is not None and col_idx < len(row) else None
            if header in {"RIT", "RUT"}:
                value = identifier_text(value)
            elif norm(header) == "T ESPERA":
                value = as_wait_days(value)
            elif norm(header) in DATE_HEADERS:
                value = as_date(value)
            else:
                value = excel_value(value)
            values[header] = value

        rows.append(
            SourceRow(
                status=status,
                category=category,
                tribunal=row_tribunal,
                values=values,
                source_file=path.name,
                source_path=relative,
                excel_row=int(df_idx) + 1,
            )
        )

    if row_tribunals:
        most_common, _ = row_tribunals.most_common(1)[0]
        file_tribunal = most_common
        if len(row_tribunals) > 1:
            message = "La columna TRIBUNAL contiene más de un tribunal; se utilizó el más frecuente"
            result.observations.append(message)
            issues.append(Issue("ADVERTENCIA", relative, "", message))
        if a1_tribunal and a1_tribunal != file_tribunal:
            message = f"A1 indica {a1_tribunal}, pero los registros indican {file_tribunal}; prevalecieron los registros"
            result.observations.append(message)
            issues.append(Issue("ADVERTENCIA", relative, "", message))
    else:
        file_tribunal = a1_tribunal or path_tribunal
    result.tribunal = file_tribunal

    if not file_tribunal:
        result.result = "ERROR"
        message = "No se pudo determinar el tribunal desde A1, la columna TRIBUNAL o la ruta"
        result.observations.append(message)
        issues.append(Issue("ERROR", relative, "", message))
        return result

    categories = list(category_counts)
    if not categories:
        fallback_category = category_from_path(relative)
        if fallback_category:
            categories = [fallback_category]
            message = f"La modalidad se determinó por la ruta: {fallback_category}"
            result.observations.append(message)
            issues.append(Issue("ADVERTENCIA", relative, "", message))
        else:
            result.result = "ERROR"
            message = "No se pudo determinar FAE o Residencia"
            result.observations.append(message)
            issues.append(Issue("ERROR", relative, "", message))
            result.orphan_rows.extend(rows)
            return result
    result.categories = sorted(categories, key=lambda value: (value != "FAE", value))
    if len(result.categories) > 1:
        message = "El Excel mezcla FAE y Residencia; se separaron los registros por derivación"
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))

    modified_at = path.stat().st_mtime
    # Filas cuyo TRIBUNAL propio no coincide con el tribunal del archivo generan
    # un candidato aparte, con SU tribunal. La matriz de cobertura se arma desde
    # candidate.tribunal y el META 4 se reconstruye leyendo la columna TRIBUNAL
    # de cada fila (read_registros); si ambos no usan el mismo criterio, pueden
    # contradecirse: la matriz marca un tribunal FALTANTE mientras el META 4 ya
    # tiene filas de ese tribunal (defecto H-03).
    other_tribunals = {row.tribunal for row in rows if row.tribunal and row.tribunal != file_tribunal}
    if other_tribunals:
        message = (
            f"El archivo está asociado a {file_tribunal}, pero contiene filas de "
            f"{', '.join(sorted(other_tribunals))}; se generó un candidato separado por cada "
            "tribunal detectado en la columna TRIBUNAL"
        )
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))

    for category in result.categories:
        rows_by_tribunal: dict[str, list[SourceRow]] = defaultdict(list)
        for row in rows:
            if row.category != category:
                continue
            rows_by_tribunal[row.tribunal or file_tribunal].append(row)
        if not rows_by_tribunal:
            # El Excel no tiene filas de datos para esta categoría (por ejemplo,
            # cero registros y la modalidad se determinó por la ruta). Debe
            # seguir cubriendo la combinación con 0 registros, igual que antes
            # de dividir candidatos por tribunal.
            rows_by_tribunal[file_tribunal] = []
        for candidate_tribunal, candidate_rows in rows_by_tribunal.items():
            candidate = Candidate(
                file_result=result,
                tribunal=candidate_tribunal,
                category=category,
                status=status,
                kind="EXCEL",
                rows=candidate_rows,
                emission_date=result.emission_date,
                modified_at=modified_at,
            )
            result.candidates.append(candidate)

    unknown_rows = [row for row in rows if row.category is None]
    if unknown_rows:
        result.orphan_rows.extend(unknown_rows)
        message = (
            f"{len(unknown_rows)} registro(s) quedan fuera del alcance FAE/Residencia "
            "y no se incorporan a Registros ni a META 4"
        )
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))
    return result


def find_directory_tribunal(path: Path, root: Path, mapping: dict[Path, str]) -> str | None:
    current = path.parent.resolve()
    root_resolved = root.resolve()
    while True:
        if current in mapping:
            return mapping[current]
        if current == root_resolved or current.parent == current:
            return None
        current = current.parent


def inspect_zero_record_file(
    path: Path,
    root: Path,
    kind: str,
    directory_map: dict[Path, str],
    issues: list[Issue],
) -> FileResult:
    relative = relative_string(path, root)
    tribunal, path_notes = tribunal_from_path(relative)
    if not tribunal:
        tribunal = find_directory_tribunal(path, root, directory_map)
    status = status_from_path(relative)
    category = category_from_path(relative)
    result = FileResult(
        path=path,
        relative_path=relative,
        kind=kind,
        result="SIN REGISTROS",
        tribunal=tribunal,
        status=status,
        categories=[category] if category else [],
        record_count=0,
    )
    for note in path_notes:
        result.observations.append(note)
        issues.append(Issue("ADVERTENCIA", relative, "", note))
    if not tribunal or not status or not category:
        result.result = "NO CLASIFICADO"
        missing = []
        if not tribunal:
            missing.append("tribunal")
        if not category:
            missing.append("modalidad")
        if not status:
            missing.append("estado")
        message = "No se pudo identificar: " + ", ".join(missing)
        result.observations.append(message)
        issues.append(Issue("ADVERTENCIA", relative, "", message))
        return result
    candidate = Candidate(
        file_result=result,
        tribunal=tribunal,
        category=category,
        status=status,
        kind=kind,
        rows=[],
        emission_date=None,
        modified_at=path.stat().st_mtime,
    )
    result.candidates.append(candidate)
    return result


def inspect_ignored(path: Path, root: Path, reason: str) -> FileResult:
    return FileResult(
        path=path,
        relative_path=relative_string(path, root),
        kind="OTRO",
        result="IGNORADO",
        observations=[reason],
    )


def scan_folder(root: Path) -> tuple[list[FileResult], list[Issue]]:
    issues: list[Issue] = []
    file_results: list[FileResult] = []
    discovered: list[Path] = []
    for directory, subdirectories, filenames in os.walk(root, followlinks=False):
        subdirectories[:] = [
            name for name in subdirectories if not Path(directory, name).is_symlink()
        ]
        for filename in filenames:
            path = Path(directory, filename)
            if path.is_file() and not path.is_symlink():
                discovered.append(path)
    files = sorted(discovered, key=lambda path: norm(relative_string(path, root)))

    excel_paths: list[Path] = []
    other_paths: list[Path] = []
    for path in files:
        generated_reason = output_file_reason(path)
        if generated_reason:
            file_results.append(inspect_ignored(path, root, generated_reason))
            continue
        if path.suffix.lower() in EXCEL_SUFFIXES:
            excel_paths.append(path)
        else:
            other_paths.append(path)

    with tempfile.TemporaryDirectory(prefix="meta4_sources_") as temp_dir:
        temp_root = Path(temp_dir)
        for index, path in enumerate(excel_paths, 1):
            readable_path = temp_root / f"{index:04d}{path.suffix.lower()}"
            try:
                shutil.copy2(path, readable_path)
            except OSError as exc:
                relative = relative_string(path, root)
                message = f"No se pudo copiar para lectura: {exc}"
                result = FileResult(path, relative, "EXCEL", "ERROR", observations=[message])
                file_results.append(result)
                issues.append(Issue("ERROR", relative, "", message))
                continue
            file_results.append(inspect_excel(path, root, readable_path, issues))

    directory_votes: dict[Path, Counter[str]] = defaultdict(Counter)
    for result in file_results:
        if result.kind == "EXCEL" and result.tribunal:
            directory_votes[result.path.parent.resolve()][result.tribunal] += 1
    directory_map = {
        directory: votes.most_common(1)[0][0]
        for directory, votes in directory_votes.items()
        if votes
    }

    for path in other_paths:
        suffix = path.suffix.lower()
        if suffix in PDF_SUFFIXES:
            file_results.append(inspect_zero_record_file(path, root, "PDF", directory_map, issues))
        elif suffix in IMAGE_SUFFIXES:
            file_results.append(inspect_zero_record_file(path, root, "IMAGEN", directory_map, issues))
        else:
            file_results.append(inspect_ignored(path, root, "Tipo de archivo no utilizado"))

    file_results.sort(key=lambda result: norm(result.relative_path))
    return file_results, issues


def candidate_sort_key(candidate: Candidate) -> tuple[int, int, float, int, str]:
    """Un Excel siempre vence a un PDF o imagen: su contenido fue leído.

    Entre archivos del mismo rango decide la fecha de emisión y luego la de
    modificación. Sin este primer criterio, un PDF más reciente eliminaba los
    registros de un Excel (defecto F-02).
    """
    effective_date = candidate.emission_date
    if effective_date:
        date_score = effective_date.toordinal()
    else:
        date_score = dt.datetime.fromtimestamp(candidate.modified_at).date().toordinal()
    evidence_score = 1 if candidate.kind == "EXCEL" else 0
    kind_score = {"EXCEL": 3, "PDF": 2, "IMAGEN": 1}.get(candidate.kind, 0)
    return (
        evidence_score,
        date_score,
        candidate.modified_at,
        kind_score,
        norm(candidate.file_result.relative_path),
    )


def select_candidates(
    file_results: list[FileResult],
    issues: list[Issue],
) -> tuple[dict[tuple[str, str, str], Candidate], dict[tuple[str, str, str], list[Candidate]]]:
    groups: dict[tuple[str, str, str], list[Candidate]] = defaultdict(list)
    for result in file_results:
        for candidate in result.candidates:
            groups[candidate.key].append(candidate)

    selected: dict[tuple[str, str, str], Candidate] = {}
    for key in expected_keys():
        candidates = groups.get(key, [])
        if not candidates:
            issues.append(Issue("FALTANTE", "", combination_text(key), "No se encontró un archivo para esta combinación"))
            continue
        chosen = max(candidates, key=candidate_sort_key)
        chosen.selected = True
        chosen.selection_note = "Seleccionado"
        selected[key] = chosen
        if len(candidates) > 1:
            omitted = [candidate for candidate in candidates if candidate is not chosen]
            for candidate in omitted:
                candidate.selection_note = f"Duplicado omitido; se utilizó {chosen.file_result.relative_path}"
                candidate.file_result.observations.append(candidate.selection_note)
            chosen.file_result.observations.append(f"Seleccionado entre {len(candidates)} archivos duplicados")
            omitted_names = "; ".join(candidate.file_result.relative_path for candidate in omitted)
            message = (
                f"Se utilizó {chosen.file_result.relative_path}. Omitidos: {omitted_names}. "
                "Criterio: Excel sobre PDF/imagen, luego fecha de emisión y fecha de modificación"
            )
            issues.append(Issue("DUPLICADO", chosen.file_result.relative_path, combination_text(key), message))

            lost_rows = sum(
                len(candidate.rows) for candidate in omitted if candidate.kind == "EXCEL"
            )
            if not chosen.rows and lost_rows:
                issues.append(
                    Issue(
                        "ERROR",
                        chosen.file_result.relative_path,
                        combination_text(key),
                        f"El archivo seleccionado no aporta registros, pero se omitieron "
                        f"{lost_rows} registro(s) de un Excel. Revisa fechas y nombres antes de usar el META 4",
                    )
                )
            newer_omitted = [
                candidate
                for candidate in omitted
                if candidate.kind != "EXCEL" and candidate.modified_at > chosen.modified_at
            ]
            if chosen.kind == "EXCEL" and newer_omitted:
                newer_names = "; ".join(candidate.file_result.relative_path for candidate in newer_omitted)
                issues.append(
                    Issue(
                        "REQUIERE_REVISION",
                        chosen.file_result.relative_path,
                        combination_text(key),
                        f"Se usó el Excel {chosen.file_result.relative_path} para no perder sus "
                        f"{len(chosen.rows)} registro(s), pero existe un archivo más reciente sin "
                        f"registros para la misma combinación: {newer_names}. Confirma que el Excel "
                        "corresponde al período informado antes de dar el META 4 por definitivo",
                    )
                )
    return selected, groups


def consolidated_rows(
    selected: dict[tuple[str, str, str], Candidate],
) -> dict[str, list[SourceRow]]:
    rows: dict[str, list[SourceRow]] = {"Espera": [], "Cumplimiento": []}
    tribunal_order = {tribunal: index for index, tribunal in enumerate(TRIBUNALS)}
    ordered_candidates = sorted(
        selected.values(),
        key=lambda candidate: (
            tribunal_order.get(candidate.tribunal, 999),
            candidate.category != "FAE",
            candidate.status != "Espera",
            norm(candidate.file_result.relative_path),
        ),
    )
    for candidate in ordered_candidates:
        if candidate.kind != "EXCEL":
            continue
        # Alcance de la herramienta: solo FAE y Residencia. Las filas con otra
        # derivación (OPD, DAM, ambulatorios, etc.) quedan fuera de Registros y
        # solo se informan como advertencia con archivo y conteo.
        rows[candidate.status].extend(candidate.rows)
    return rows


NAVY = "17365D"
BLUE = "2F75B5"
LIGHT_BLUE = "D9EAF7"
PALE_BLUE = "EAF3F8"
GREEN = "E2F0D9"
ORANGE = "FCE4D6"
RED = "F4CCCC"
GRAY = "E7E6E6"
WHITE = "FFFFFF"
DARK_TEXT = "1F2937"
LIGHT_BORDER = "CBD5E1"


def unique_output_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(1, 1000):
        candidate = path.with_name(f"{path.stem} ({index}){path.suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"No se pudo crear un nombre único cerca de {path}")


def save_workbook_atomic(workbook: Workbook, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f"{path.stem}_",
            suffix=path.suffix,
            dir=path.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
        workbook.save(temporary)
        os.replace(temporary, path)
        temporary = None
    finally:
        if temporary and temporary.exists():
            temporary.unlink()


def apply_header_style(ws, row: int, start_col: int, end_col: int) -> None:
    thin = Side(style="thin", color=LIGHT_BORDER)
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    ws.row_dimensions[row].height = 30


def style_data_sheet(ws, headers: list[str], rows: list[SourceRow]) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(rows) + 1)}"
    for col_idx, header in enumerate(headers, 1):
        ws.cell(1, col_idx, header)
    apply_header_style(ws, 1, 1, len(headers))

    for row_idx, source_row in enumerate(rows, 2):
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row_idx, col_idx, source_row.values.get(header))
            cell.font = Font(name="Aptos", size=9, color=DARK_TEXT)
            cell.alignment = Alignment(vertical="top", wrap_text=False)
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F7FAFC")
            normalized = norm(header)
            if normalized in DATE_HEADERS and cell.value is not None:
                cell.number_format = "DD-MM-YYYY"
            elif normalized in {"T ESPERA", "DURACION", "EDAD", "DIAS DE CUMPLIMIENTO", "DIAS PARA EGRESAR"}:
                cell.number_format = "0"
            elif normalized in {"RIT", "RUT"}:
                cell.number_format = "@"

    width_by_header = {
        "RIT": 15,
        "TRIBUNAL": 34,
        "RUT": 15,
        "NOMBRE": 30,
        "NACIONALIDAD PAIS": 20,
        "TIPO": 14,
        "SEXO": 9,
        "DERIVACION": 42,
        "CURADOR": 30,
        "T ESPERA": 12,
        "REVISION": 14,
        "DURACION": 12,
        "80 BIS": 10,
        "EDAD": 9,
        "PROXS AUDS": 22,
        "ARCHIVO ORIGEN": 38,
        "RUTA ORIGEN": 60,
    }
    for col_idx, header in enumerate(headers, 1):
        normalized = norm(header)
        if normalized in width_by_header:
            width = width_by_header[normalized]
        elif normalized in DATE_HEADERS or normalized.startswith("FEC ") or normalized.startswith("FECHA "):
            width = 18
        elif normalized.startswith("DIAS "):
            width = 16
        else:
            width = 16
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def style_simple_table(
    ws,
    header_row: int,
    first_col: int,
    last_col: int,
    last_row: int,
    freeze: str,
) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = freeze
    apply_header_style(ws, header_row, first_col, last_col)
    if last_row >= header_row:
        ws.auto_filter.ref = (
            f"{get_column_letter(first_col)}{header_row}:"
            f"{get_column_letter(last_col)}{max(header_row, last_row)}"
        )
    for row in range(header_row + 1, last_row + 1):
        for col in range(first_col, last_col + 1):
            cell = ws.cell(row, col)
            cell.font = Font(name="Aptos", size=9, color=DARK_TEXT)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F7FAFC")


def write_inventory_sheet(ws, file_results: list[FileResult]) -> None:
    headers = [
        "N°",
        "ARCHIVO",
        "RUTA RELATIVA",
        "TIPO",
        "TRIBUNAL",
        "MODALIDAD",
        "ESTADO",
        "REGISTROS",
        "FECHA EMISIÓN",
        "RESULTADO",
        "SELECCIONADO",
        "OBSERVACIÓN",
    ]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(1, col_idx, header)
    for row_idx, result in enumerate(file_results, 2):
        values = [
            row_idx - 1,
            result.path.name,
            result.relative_path,
            result.kind,
            result.tribunal,
            ", ".join(result.categories),
            result.status,
            result.record_count,
            result.emission_date,
            result.result,
            result.selected_label(),
            " | ".join(dict.fromkeys(result.observations)),
        ]
        for col_idx, value in enumerate(values, 1):
            ws.cell(row_idx, col_idx, value)
        if result.emission_date:
            ws.cell(row_idx, 9).number_format = "DD-MM-YYYY"
        result_cell = ws.cell(row_idx, 10)
        if result.result in {"ERROR", "NO CLASIFICADO"}:
            result_cell.fill = PatternFill("solid", fgColor=RED)
        elif result.result == "SIN REGISTROS":
            result_cell.fill = PatternFill("solid", fgColor=PALE_BLUE)
        elif result.result == "IGNORADO":
            result_cell.fill = PatternFill("solid", fgColor=GRAY)
        else:
            result_cell.fill = PatternFill("solid", fgColor=GREEN)
    style_simple_table(ws, 1, 1, len(headers), len(file_results) + 1, "A2")
    widths = [7, 42, 65, 12, 23, 16, 16, 12, 16, 18, 14, 70]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width


def matrix_cell_text(
    key: tuple[str, str, str],
    selected: dict[tuple[str, str, str], Candidate],
    groups: dict[tuple[str, str, str], list[Candidate]],
) -> str:
    candidate = selected.get(key)
    if not candidate:
        return "FALTANTE"
    if candidate.kind == "EXCEL":
        text = f"EXCEL · {len(candidate.rows)} registro(s)"
        out_of_scope = len(candidate.file_result.orphan_rows)
        if out_of_scope:
            text += f" (+{out_of_scope} fuera de alcance)"
    elif candidate.kind == "PDF":
        text = "PDF · 0 registros"
    else:
        text = "IMAGEN · 0 registros"
    siblings = groups.get(key, [])
    if len(siblings) > 1:
        text += " · DUPLICADO"
    if candidate.kind == "EXCEL" and any(
        sibling.kind != "EXCEL" and sibling.modified_at > candidate.modified_at for sibling in siblings
    ):
        text += " · REQUIERE REVISIÓN"
    return text


def write_matrix_sheet(
    ws,
    selected: dict[tuple[str, str, str], Candidate],
    groups: dict[tuple[str, str, str], list[Candidate]],
) -> None:
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:E1")
    ws["A1"] = "MATRIZ DE COBERTURA DE ARCHIVOS"
    ws["A1"].font = Font(name="Aptos Display", size=16, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30

    found = len(selected)
    duplicates = sum(1 for key in expected_keys() if len(groups.get(key, [])) > 1)
    ws.merge_cells("A2:E2")
    ws["A2"] = f"Cobertura: {found}/68 · Faltantes: {68 - found} · Combinaciones duplicadas: {duplicates}"
    ws["A2"].font = Font(name="Aptos", size=10, bold=True, color=DARK_TEXT)
    ws["A2"].fill = PatternFill("solid", fgColor=LIGHT_BLUE)
    ws["A2"].alignment = Alignment(vertical="center")

    headers = ["TRIBUNAL", "FAE ESPERA", "FAE CUMPLIMIENTO", "RESIDENCIA ESPERA", "RESIDENCIA CUMPLIMIENTO"]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(4, col_idx, header)
    apply_header_style(ws, 4, 1, 5)

    for row_idx, tribunal in enumerate(TRIBUNALS, 5):
        ws.cell(row_idx, 1, tribunal)
        keys = [
            (tribunal, "FAE", "Espera"),
            (tribunal, "FAE", "Cumplimiento"),
            (tribunal, "Residencia", "Espera"),
            (tribunal, "Residencia", "Cumplimiento"),
        ]
        for col_idx, key in enumerate(keys, 2):
            ws.cell(row_idx, col_idx, matrix_cell_text(key, selected, groups))
        for col_idx in range(1, 6):
            cell = ws.cell(row_idx, col_idx)
            cell.font = Font(name="Aptos", size=9, bold=(col_idx == 1), color=DARK_TEXT)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F7FAFC")
        ws.row_dimensions[row_idx].height = 30

    data_range = f"B5:E{4 + len(TRIBUNALS)}"
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["ISNUMBER(SEARCH(\"FALTANTE\",B5))"], fill=PatternFill("solid", fgColor=RED)),
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["ISNUMBER(SEARCH(\"DUPLICADO\",B5))"], fill=PatternFill("solid", fgColor=ORANGE)),
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["ISNUMBER(SEARCH(\"EXCEL\",B5))"], fill=PatternFill("solid", fgColor=GREEN)),
    )
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(formula=["OR(ISNUMBER(SEARCH(\"PDF\",B5)),ISNUMBER(SEARCH(\"IMAGEN\",B5)))"], fill=PatternFill("solid", fgColor=PALE_BLUE)),
    )
    ws.freeze_panes = "B5"
    ws.column_dimensions["A"].width = 24
    for column in "BCDE":
        ws.column_dimensions[column].width = 29


def write_issues_sheet(ws, issues: list[Issue]) -> None:
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:D1")
    ws["A1"] = "ERRORES Y ADVERTENCIAS"
    ws["A1"].font = Font(name="Aptos Display", size=16, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    headers = ["NIVEL", "ARCHIVO", "COMBINACIÓN", "MENSAJE"]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(3, col_idx, header)

    display_issues = issues or [Issue("INFO", "", "", "No se detectaron errores ni advertencias")]
    for row_idx, issue in enumerate(display_issues, 4):
        values = [issue.level, issue.file, issue.combination, issue.message]
        for col_idx, value in enumerate(values, 1):
            ws.cell(row_idx, col_idx, value)
        level_cell = ws.cell(row_idx, 1)
        if issue.level in {"ERROR", "FALTANTE"}:
            level_cell.fill = PatternFill("solid", fgColor=RED)
        elif issue.level in {"ADVERTENCIA", "DUPLICADO", "REQUIERE_REVISION"}:
            level_cell.fill = PatternFill("solid", fgColor=ORANGE)
        else:
            level_cell.fill = PatternFill("solid", fgColor=GREEN)
    style_simple_table(ws, 3, 1, 4, len(display_issues) + 3, "A4")
    widths = [16, 55, 38, 90]
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width


def write_registros_workbook(
    path: Path,
    rows: dict[str, list[SourceRow]],
    file_results: list[FileResult],
    selected: dict[tuple[str, str, str], Candidate],
    groups: dict[tuple[str, str, str], list[Candidate]],
    issues: list[Issue],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    espera = wb.create_sheet("ESPERA")
    style_data_sheet(espera, ESPERA_HEADERS, rows["Espera"])
    cumplimiento = wb.create_sheet("CUMPLIMIENTO")
    style_data_sheet(cumplimiento, CUMPLIMIENTO_HEADERS, rows["Cumplimiento"])
    inventory = wb.create_sheet("INVENTARIO_ARCHIVOS")
    write_inventory_sheet(inventory, file_results)
    matrix = wb.create_sheet("MATRIZ_COBERTURA")
    write_matrix_sheet(matrix, selected, groups)
    issue_sheet = wb.create_sheet("ERRORES_Y_ADVERTENCIAS")
    write_issues_sheet(issue_sheet, issues)
    wb.active = 0
    save_workbook_atomic(wb, path)


def find_sheet(workbook_sheets: Iterable[str], target: str) -> str:
    wanted = norm(target)
    for sheet in workbook_sheets:
        if norm(sheet) == wanted:
            return sheet
    raise KeyError(f"Falta la hoja: {target}")


def find_col(columns: Iterable[Any], kind: str, required: bool = True) -> Any:
    normalized = [(col, norm(col)) for col in columns]
    if kind == "rit":
        matches = [col for col, value in normalized if value == "RIT"]
    elif kind == "tribunal":
        matches = [col for col, value in normalized if value == "TRIBUNAL"]
    elif kind == "tipo":
        matches = [col for col, value in normalized if value == "TIPO"]
    elif kind == "sexo":
        matches = [col for col, value in normalized if value == "SEXO"]
    elif kind == "derivacion":
        matches = [col for col, value in normalized if value.startswith("DERIV")]
    elif kind == "t_espera":
        matches = [col for col, value in normalized if value == "T ESPERA"]
    elif kind == "fec_resolucion":
        matches = [col for col, value in normalized if value.startswith("FEC RESOLUC")]
    elif kind == "fec_ingreso":
        matches = [col for col, value in normalized if value.startswith("FEC INGRESO EFECTIVO")]
    elif kind == "fecha_obs":
        matches = [col for col, value in normalized if value.startswith("FECHA OBS")]
    else:
        matches = []
    if not matches:
        if required:
            raise KeyError(f"Falta una columna para {kind}")
        return None
    return matches[0]


def read_registros(
    path: Path,
) -> tuple[
    dict[str, dict[str, list[list[Any]]]],
    dict[str, int],
    list[dict[str, Any]],
]:
    records = {name: {tribunal: [] for tribunal in TRIBUNALS} for name in SHEET_DEFS}
    source_counts: dict[str, int] = {}
    unmatched: list[dict[str, Any]] = []

    with pd.ExcelFile(path) as xls:
        source_sheet_names = {
            "ESPERA": find_sheet(xls.sheet_names, "ESPERA"),
            "CUMPLIMIENTO": find_sheet(xls.sheet_names, "CUMPLIMIENTO"),
        }
        for source, actual_sheet in source_sheet_names.items():
            df = pd.read_excel(xls, sheet_name=actual_sheet, dtype=object)
            cols = df.columns
            c_rit = find_col(cols, "rit")
            c_tri = find_col(cols, "tribunal")
            c_tipo = find_col(cols, "tipo")
            c_sexo = find_col(cols, "sexo")
            c_der = find_col(cols, "derivacion")
            c_res = find_col(cols, "fec_resolucion")
            c_obs = find_col(cols, "fecha_obs", required=False)
            c_wait = find_col(cols, "t_espera") if source == "ESPERA" else None
            c_ing = find_col(cols, "fec_ingreso") if source == "CUMPLIMIENTO" else None

            real = df[df[c_rit].notna() & df[c_rit].astype(str).str.strip().ne("")].copy()
            source_counts[source] = len(real)
            for idx, row in real.iterrows():
                category = classify_derivation(row[c_der])
                tribunal = tribunal_name(row[c_tri])
                if category is None or tribunal is None:
                    unmatched.append(
                        {
                            "sheet": source,
                            "row": int(idx) + 2,
                            "tribunal": clean_text(row[c_tri]),
                            "derivacion": clean_text(row[c_der]),
                        }
                    )
                    continue

                sheet_name = f"{category} Espera" if source == "ESPERA" else f"{category} Cumplimiento"
                observation = row[c_obs] if c_obs is not None else None
                if source == "ESPERA":
                    rec = [
                        clean_text(row[c_rit]),
                        clean_text(row[c_tri]),
                        clean_text(row[c_tipo]),
                        clean_text(row[c_sexo]),
                        clean_text(row[c_der]),
                        as_wait_days(row[c_wait]),
                        as_date(row[c_res]),
                        as_date(observation),
                    ]
                else:
                    rec = [
                        clean_text(row[c_rit]),
                        clean_text(row[c_tri]),
                        clean_text(row[c_tipo]),
                        clean_text(row[c_sexo]),
                        clean_text(row[c_der]),
                        as_date(row[c_res]),
                        as_date(row[c_ing]),
                        as_date(observation),
                    ]
                records[sheet_name][tribunal].append(rec)
    return records, source_counts, unmatched


def default_template_path(base_dir: Path) -> Path:
    return base_dir / "assets" / "meta4-template.xlsx"


def build_default_template() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    widths = [13, 34, 15, 8, 44, 18, 22, 16]
    thin = Side(style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor=LIGHT_BLUE)

    for sheet_name, spec in SHEET_DEFS.items():
        ws = wb.create_sheet(sheet_name)
        ws.sheet_view.showGridLines = False
        for idx, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        row = 1
        ws.cell(row, 1, "ARAUCO").font = Font(bold=True)
        row += 1
        ws.cell(row, 1, spec["report"])
        row += 1
        ws.cell(row, 1, "Fecha Emisión   : 01-01-2000")
        row += 2
        for col, value in enumerate(spec["headers"], 1):
            cell = ws.cell(row, col, value)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        row += 1
        if "Cumplimiento" in sheet_name:
            dummy = [
                "X-0-0000",
                "Tribunal",
                "Tipo",
                "S",
                "DERIVACIÓN",
                dt.date(2000, 1, 1),
                dt.date(2000, 1, 1),
                dt.date(2000, 1, 1),
            ]
        else:
            dummy = [
                "X-0-0000",
                "Tribunal",
                "Tipo",
                "S",
                "DERIVACIÓN",
                0,
                dt.date(2000, 1, 1),
                dt.date(2000, 1, 1),
            ]
        for col, value in enumerate(dummy, 1):
            cell = ws.cell(row, col, value)
            cell.border = border
            if "Cumplimiento" in sheet_name and col in (6, 7, 8):
                cell.number_format = "DD-MM-YYYY"
            elif "Espera" in sheet_name and col in (7, 8):
                cell.number_format = "DD-MM-YYYY"
            elif "Espera" in sheet_name and col == 6:
                cell.number_format = "0"
        row += 2
        ws.cell(row, 1, "SIN REGISTROS").font = Font(bold=True)
        row += 2
        ws.cell(row, 1, "TOTAL:").font = Font(bold=True)
        ws.cell(row, 2, 0).font = Font(bold=True)
    return wb


def load_template(path: Path | None) -> Workbook:
    if path and path.exists():
        wb = load_workbook(path)
    else:
        wb = build_default_template()
    for ws in list(wb.worksheets):
        clean = ws.title.strip()
        if clean != ws.title:
            ws.title = clean
    return wb


def row_styles(ws, row_num: int | None, width: int = 8) -> list[dict[str, Any]]:
    if row_num is None:
        return [{} for _ in range(width)]
    styles = []
    for col in range(1, width + 1):
        cell = ws.cell(row_num, col)
        styles.append(
            {
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": copy(cell.protection),
            }
        )
    return styles


def find_row(ws, predicate) -> int | None:
    for row in range(1, ws.max_row + 1):
        if predicate(row):
            return row
    return None


def capture_templates(ws) -> dict[str, Any]:
    title = find_row(
        ws,
        lambda row: isinstance(ws.cell(row, 1).value, str)
        and isinstance(ws.cell(row + 1, 1).value, str)
        and ws.cell(row + 1, 1).value.startswith("Informe de Ingresos"),
    )
    header = find_row(ws, lambda row: ws.cell(row, 1).value == "RIT")
    data = header + 1 if header else None
    sin = find_row(ws, lambda row: ws.cell(row, 1).value == "SIN REGISTROS")
    total = find_row(
        ws,
        lambda row: isinstance(ws.cell(row, 1).value, str)
        and ws.cell(row, 1).value.upper().startswith("TOTAL"),
    )
    if not title or not header or not data or not sin or not total:
        raise RuntimeError(f"La hoja de plantilla {ws.title} no contiene las filas de estilo requeridas")
    return {
        "title": row_styles(ws, title),
        "report": row_styles(ws, title + 1),
        "date": row_styles(ws, title + 2),
        "header": row_styles(ws, header),
        "data": row_styles(ws, data),
        "sin": row_styles(ws, sin),
        "total": row_styles(ws, total),
        "heights": {
            "title": ws.row_dimensions[title].height,
            "report": ws.row_dimensions[title + 1].height,
            "date": ws.row_dimensions[title + 2].height,
            "header": ws.row_dimensions[header].height,
            "data": ws.row_dimensions[data].height,
            "sin": ws.row_dimensions[sin].height,
            "total": ws.row_dimensions[total].height,
        },
    }


def apply_style(cell, style: dict[str, Any]) -> None:
    if not style:
        return
    cell.font = copy(style["font"])
    cell.fill = copy(style["fill"])
    cell.border = copy(style["border"])
    cell.alignment = copy(style["alignment"])
    cell.number_format = style["number_format"]
    cell.protection = copy(style["protection"])


def write_row(ws, row_num: int, values: list[Any], styles: list[dict[str, Any]], date_cols: tuple[int, ...] = ()) -> None:
    for col in range(1, 9):
        cell = ws.cell(row_num, col)
        cell.value = values[col - 1] if col <= len(values) else None
        apply_style(cell, styles[col - 1])
        if col in date_cols and cell.value is not None:
            cell.number_format = "DD-MM-YYYY"


def set_height(ws, row_num: int, templates: dict[str, Any], key: str) -> None:
    height = templates["heights"].get(key)
    if height is not None:
        ws.row_dimensions[row_num].height = height


def rebuild(
    wb: Workbook,
    records: dict[str, dict[str, list[list[Any]]]],
    emission_date: dt.date,
) -> dict[str, int]:
    emission_text = emission_date.strftime("%d-%m-%Y")
    summary: dict[str, int] = {}
    for sheet_name, spec in SHEET_DEFS.items():
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            temp = build_default_template()[sheet_name]
            for col in range(1, 9):
                letter = get_column_letter(col)
                ws.column_dimensions[letter].width = temp.column_dimensions[letter].width
            templates = capture_templates(temp)
        else:
            ws = wb[sheet_name]
            templates = capture_templates(ws)
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)

        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        row = 1
        total = 0
        for tribunal in TRIBUNALS:
            tribunal_records = records[sheet_name][tribunal]
            write_row(ws, row, [tribunal], templates["title"])
            set_height(ws, row, templates, "title")
            row += 1
            write_row(ws, row, [spec["report"]], templates["report"])
            set_height(ws, row, templates, "report")
            row += 1
            write_row(ws, row, [f"Fecha Emisión   : {emission_text}"], templates["date"])
            set_height(ws, row, templates, "date")
            row += 2
            if tribunal_records:
                write_row(ws, row, spec["headers"], templates["header"])
                set_height(ws, row, templates, "header")
                row += 1
                date_cols = (7, 8) if "Espera" in sheet_name else (6, 7, 8)
                for rec in tribunal_records:
                    write_row(ws, row, rec, templates["data"], date_cols=date_cols)
                    set_height(ws, row, templates, "data")
                    row += 1
                total += len(tribunal_records)
            else:
                write_row(ws, row, ["SIN REGISTROS"], templates["sin"])
                set_height(ws, row, templates, "sin")
                row += 1
            row += 1
        row += 1
        write_row(ws, row, ["TOTAL:", total], templates["total"])
        set_height(ws, row, templates, "total")
        summary[sheet_name] = total
    wb.active = 0
    return summary


def verify_meta_output(path: Path) -> tuple[dict[str, Any], dict[str, int], list[str]]:
    """Lee el TOTAL escrito y, por separado, recuenta las filas de datos.

    El recuento independiente cierra la verificación circular del programa
    original (§5.4 del informe v2.0).
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        totals: dict[str, Any] = {}
        counted: dict[str, int] = {}
        problems: list[str] = []
        for sheet_name in SHEET_DEFS:
            ws = wb[sheet_name]
            total = None
            data_rows = 0
            in_block = False
            tribunal_labels = {norm(tribunal) for tribunal in TRIBUNALS}
            for row in ws.iter_rows(min_col=1, max_col=2, values_only=True):
                first = row[0]
                text = norm(first) if isinstance(first, str) else ""
                if text == "RIT":
                    in_block = True
                    continue
                if text.startswith("TOTAL"):
                    total = row[1]
                    in_block = False
                    continue
                if text in tribunal_labels or text == "SIN REGISTROS":
                    in_block = False
                    continue
                if in_block and first is not None and str(first).strip():
                    data_rows += 1
            totals[sheet_name] = total
            counted[sheet_name] = data_rows
            if total != data_rows:
                problems.append(
                    f"{sheet_name}: TOTAL declara {total} pero se contaron {data_rows} filas de datos"
                )
        return totals, counted, problems
    finally:
        wb.close()


def _fingerprint_value(value: Any) -> str:
    """Forma canónica de un valor de celda para comparar filas por contenido."""
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (int, float)):
        return str(value)
    return norm(value)


def verify_registros_output(
    path: Path,
    expected_rows: dict[str, list[SourceRow]] | None = None,
) -> dict[str, Any]:
    """Recuenta el archivo escrito y lo contrasta con lo que había en memoria.

    La comparación memoria-vs-archivo es la red de seguridad principal: como
    ambos números provienen de construir el mismo `rows` en distintos momentos,
    una discrepancia solo puede deberse a un error al escribir, no a que el RIT
    se repita.

    El detector de filas repetidas (H-01) usa una huella con TODAS las columnas
    de negocio, no solo RIT: un RIT es la causa y puede legítimamente tener más
    de una fila (hermanos, varias medidas). Usar RIT+archivo como clave produce
    falsos positivos con datos reales. ARCHIVO ORIGEN y RUTA ORIGEN se excluyen
    de la huella (son metadato técnico) y se comparan aparte para distinguir un
    duplicado dentro de un mismo archivo de uno entre archivos distintos.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        result: dict[str, Any] = {"sheets": list(wb.sheetnames), "rows": {}, "problems": []}
        for sheet_name, memory_key in (("ESPERA", "Espera"), ("CUMPLIMIENTO", "Cumplimiento")):
            ws = wb[sheet_name]
            headers = [cell for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())]
            normalized = [norm(header) for header in headers]
            rit_idx = normalized.index("RIT") if "RIT" in normalized else 0
            archivo_idx = normalized.index("ARCHIVO ORIGEN") if "ARCHIVO ORIGEN" in normalized else None
            ruta_idx = normalized.index("RUTA ORIGEN") if "RUTA ORIGEN" in normalized else None

            count = 0
            fingerprints: dict[tuple[str, ...], list[tuple[str, str]]] = defaultdict(list)
            excluded_columns = {index for index in (archivo_idx, ruta_idx) if index is not None}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if rit_idx >= len(row):
                    continue
                rit = row[rit_idx]
                if rit is None or not str(rit).strip():
                    continue
                count += 1
                archivo = str(row[archivo_idx] or "").strip() if archivo_idx is not None and archivo_idx < len(row) else ""
                # Huella de negocio: todas las columnas menos ARCHIVO ORIGEN y RUTA
                # ORIGEN (metadatos técnicos). El archivo se compara aparte para
                # distinguir "misma fila repetida en el mismo archivo" de "misma
                # fila proveniente de archivos distintos".
                business = tuple(
                    _fingerprint_value(value) for index, value in enumerate(row) if index not in excluded_columns
                )
                fingerprints[business].append((str(rit).strip(), archivo))
            result["rows"][sheet_name] = count

            same_file, cross_file = [], []
            for occurrences in fingerprints.values():
                if len(occurrences) < 2:
                    continue
                rit_value = occurrences[0][0]
                archivos = sorted({archivo for _, archivo in occurrences})
                if len(archivos) == 1:
                    same_file.append(f"RIT {rit_value} en {archivos[0]}")
                else:
                    cross_file.append(f"RIT {rit_value} en {', '.join(archivos)}")

            if same_file:
                result["problems"].append(
                    f"{sheet_name}: {len(same_file)} fila(s) con todos los datos idénticos "
                    f"repetidas dentro del mismo archivo de origen. Ejemplos: {'; '.join(same_file[:5])}"
                )
            if cross_file:
                result["problems"].append(
                    f"{sheet_name}: {len(cross_file)} fila(s) con todos los datos idénticos "
                    f"provenientes de archivos distintos. Ejemplos: {'; '.join(cross_file[:5])}"
                )
            if expected_rows is not None and count != len(expected_rows.get(memory_key, [])):
                result["problems"].append(
                    f"{sheet_name}: se escribieron {count} filas pero se consolidaron "
                    f"{len(expected_rows.get(memory_key, []))}"
                )
        result["inventory_rows"] = max(0, wb["INVENTARIO_ARCHIVOS"].max_row - 1)
        result["matrix_rows"] = max(0, wb["MATRIZ_COBERTURA"].max_row - 4)
        return result
    finally:
        wb.close()


def write_text_report(
    path: Path,
    root: Path,
    emission: dt.date,
    file_results: list[FileResult],
    selected: dict[tuple[str, str, str], Candidate],
    groups: dict[tuple[str, str, str], list[Candidate]],
    issues: list[Issue],
    source_counts: dict[str, int],
    meta_summary: dict[str, int],
    registros_path: Path,
    meta_path: Path,
    started_at: dt.datetime,
    finished_at: dt.datetime,
) -> None:
    divider = "=" * 96
    lines: list[str] = [
        divider,
        "REPORTE DE PROCESAMIENTO META 4",
        divider,
        f"Versión                : {APP_VERSION}",
        f"Carpeta examinada      : {root}",
        f"Inicio                  : {started_at.strftime('%d-%m-%Y %H:%M:%S')}",
        f"Término                 : {finished_at.strftime('%d-%m-%Y %H:%M:%S')}",
        f"Fecha emisión META 4    : {emission.strftime('%d-%m-%Y')}",
        "",
        "RESUMEN",
        "-" * 96,
    ]
    kind_counts = Counter(result.kind for result in file_results)
    result_counts = Counter(result.result for result in file_results)
    duplicate_count = sum(1 for key in expected_keys() if len(groups.get(key, [])) > 1)
    lines.extend(
        [
            f"Archivos inventariados  : {len(file_results)}",
            f"Excel                    : {kind_counts.get('EXCEL', 0)}",
            f"PDF                      : {kind_counts.get('PDF', 0)}",
            f"Imágenes                 : {kind_counts.get('IMAGEN', 0)}",
            f"Otros/ignorados          : {kind_counts.get('OTRO', 0)}",
            f"Combinaciones encontradas: {len(selected)}/68",
            f"Combinaciones faltantes  : {68 - len(selected)}",
            f"Combinaciones duplicadas : {duplicate_count}",
            f"Registros Espera         : {source_counts.get('ESPERA', 0)}",
            f"Registros Cumplimiento   : {source_counts.get('CUMPLIMIENTO', 0)}",
            f"Total registros          : {sum(source_counts.values())}",
            "",
            "ARCHIVOS PROCESADOS",
            "-" * 96,
        ]
    )
    for result in file_results:
        classification = " / ".join(
            value
            for value in (
                result.tribunal or "SIN TRIBUNAL",
                ", ".join(result.categories) or "SIN MODALIDAD",
                result.status or "SIN ESTADO",
            )
        )
        lines.append(
            f"[{result.result}] [{result.selected_label()}] {result.kind} | "
            f"{classification} | {result.record_count} registro(s) | {result.relative_path}"
        )
        if result.observations:
            for observation in dict.fromkeys(result.observations):
                lines.append(f"    - {observation}")

    lines.extend(["", "MATRIZ DE COBERTURA", "-" * 96])
    for tribunal in TRIBUNALS:
        lines.append(tribunal)
        for category, status in (
            ("FAE", "Espera"),
            ("FAE", "Cumplimiento"),
            ("Residencia", "Espera"),
            ("Residencia", "Cumplimiento"),
        ):
            key = (tribunal, category, status)
            candidate = selected.get(key)
            if candidate:
                detail = (
                    f"{candidate.kind}; {len(candidate.rows)} registro(s); "
                    f"{candidate.file_result.relative_path}"
                )
                if len(groups.get(key, [])) > 1:
                    detail += f"; {len(groups[key])} archivos encontrados"
            else:
                detail = "FALTANTE"
            lines.append(f"    {category} {status}: {detail}")

    lines.extend(["", "ERRORES Y ADVERTENCIAS", "-" * 96])
    if issues:
        for issue in issues:
            location = issue.file or issue.combination or "GENERAL"
            lines.append(f"[{issue.level}] {location}: {issue.message}")
    else:
        lines.append("No se detectaron errores ni advertencias.")

    lines.extend(
        [
            "",
            "TOTALES META 4",
            "-" * 96,
        ]
    )
    for sheet_name in SHEET_DEFS:
        lines.append(f"{sheet_name}: {meta_summary.get(sheet_name, 0)}")
    lines.extend(
        [
            "",
            "PRODUCTOS GENERADOS",
            "-" * 96,
            f"Registros: {registros_path}",
            f"META 4   : {meta_path}",
            f"Reporte  : {path}",
            "",
            "CONTEO POR RESULTADO DE ARCHIVO",
            "-" * 96,
        ]
    )
    for name, count in sorted(result_counts.items()):
        lines.append(f"{name}: {count}")
    lines.extend(["", divider])
    path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")


def create_template(path: Path) -> None:
    workbook = build_default_template()
    save_workbook_atomic(workbook, path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Examina una carpeta y genera Registros.xlsx, META 4 y un reporte TXT"
    )
    parser.add_argument(
        "--folder",
        help="Carpeta que se examinará recursivamente. Por defecto, la carpeta actual.",
    )
    parser.add_argument(
        "--pick-folder",
        action="store_true",
        help="Abre una ventana para elegir la carpeta cuando no se indica --folder",
    )
    parser.add_argument(
        "--open-output",
        action="store_true",
        help="Abre la carpeta de salida al terminar (solo Windows)",
    )
    parser.add_argument(
        "--output-dir",
        help="Carpeta de salida. Por defecto, la misma carpeta examinada.",
    )
    parser.add_argument("--template", help="Plantilla META 4 opcional en formato .xlsx")
    parser.add_argument("--month", help="Mes de emisión en formato AAAA-MM")
    parser.add_argument("--emission-date", help="Fecha exacta de emisión en formato DD-MM-AAAA")
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Reemplaza el META 4 del mismo mes en lugar de crear un nombre alternativo",
    )
    parser.add_argument(
        "--create-template",
        help="Crea una plantilla META 4 en la ruta indicada y termina",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parents[1]
    if args.create_template:
        template_target = Path(args.create_template).expanduser().resolve()
        create_template(template_target)
        print(json.dumps({"template": str(template_target)}, ensure_ascii=False, indent=2))
        return

    started_at = dt.datetime.now()
    if not args.folder and args.pick_folder:
        selection = select_source_folder()
        if selection.status == "NO_DISPONIBLE":
            print("No se pudo mostrar un selector de carpetas en este equipo.")
            print('Usa --folder "ruta" en su lugar.')
            raise SystemExit(1)
        if selection.status == "CANCELADO":
            print("No se seleccionó ninguna carpeta. Proceso cancelado.")
            raise SystemExit(3)
        args.folder = selection.path
    root = Path(args.folder or ".").expanduser().resolve()
    if not root.exists():
        print(f"No existe la carpeta: {root}", file=sys.stderr)
        raise SystemExit(1)
    if not root.is_dir():
        print(f"La ruta no es una carpeta: {root}", file=sys.stderr)
        raise SystemExit(1)
    output_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else root
    output_dir.mkdir(parents=True, exist_ok=True)
    emission = parse_emission_date(args, base_dir)

    file_results, issues = scan_folder(root)
    selected, groups = select_candidates(file_results, issues)
    rows = consolidated_rows(selected)
    registros_path = output_dir / "Registros.xlsx"
    write_registros_workbook(registros_path, rows, file_results, selected, groups, issues)

    records, source_counts, unmatched = read_registros(registros_path)
    if unmatched:
        for item in unmatched:
            message = (
                f"Registro excluido de META 4: tribunal={item['tribunal']!r}, "
                f"derivación={item['derivacion']!r}"
            )
            issues.append(
                Issue(
                    "ADVERTENCIA",
                    f"{item['sheet']} fila {item['row']}",
                    "",
                    message,
                )
            )
        write_registros_workbook(registros_path, rows, file_results, selected, groups, issues)

    template = (
        Path(args.template).expanduser().resolve()
        if args.template
        else default_template_path(base_dir)
    )
    workbook = load_template(template)
    meta_summary = rebuild(workbook, records, emission)
    meta_path = output_dir / f"META 4 - actualizado {emission.strftime('%Y-%m')}.xlsx"
    if meta_path.exists() and not args.overwrite:
        meta_path = unique_output_path(meta_path)
    save_workbook_atomic(workbook, meta_path)
    verified_meta, counted_meta, meta_problems = verify_meta_output(meta_path)
    verified_registros = verify_registros_output(registros_path, rows)
    verification_issues = list(meta_problems) + list(verified_registros["problems"])
    if verified_meta != meta_summary:
        verification_issues.append(
            f"Los totales verificados {verified_meta} no coinciden con {meta_summary}"
        )
    if counted_meta != meta_summary:
        verification_issues.append(
            f"El recuento de filas por bloque {counted_meta} no coincide con {meta_summary}"
        )
    registros_total = len(rows["Espera"]) + len(rows["Cumplimiento"])
    meta_total = sum(meta_summary.values())
    expected_meta_total = registros_total - len(unmatched)
    if meta_total != expected_meta_total:
        verification_issues.append(
            f"El META 4 contiene {meta_total} fila(s) pero deberían ser {expected_meta_total} "
            f"({registros_total} en Registros menos {len(unmatched)} excluida(s) por tribunal o "
            "derivación no reconocidos). Revisa coincidencia entre matriz de cobertura y META 4"
        )
    if verification_issues:
        for message in verification_issues:
            issues.append(Issue("ERROR", meta_path.name, "", f"Verificación: {message}"))
        write_registros_workbook(registros_path, rows, file_results, selected, groups, issues)
        verified_registros = verify_registros_output(registros_path, rows)

    timestamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    report_path = unique_output_path(
        output_dir / f"Reporte de procesamiento META 4 - {timestamp}.txt"
    )
    finished_at = dt.datetime.now()
    write_text_report(
        report_path,
        root,
        emission,
        file_results,
        selected,
        groups,
        issues,
        source_counts,
        meta_summary,
        registros_path,
        meta_path,
        started_at,
        finished_at,
    )

    print(
        json.dumps(
            {
                "version": APP_VERSION,
                "source_folder": str(root),
                "emission": emission.strftime("%d-%m-%Y"),
                "outputs": {
                    "registros": str(registros_path),
                    "meta4": str(meta_path),
                    "report": str(report_path),
                },
                "files": {
                    "inventoried": len(file_results),
                    "excel": sum(result.kind == "EXCEL" for result in file_results),
                    "pdf": sum(result.kind == "PDF" for result in file_results),
                    "images": sum(result.kind == "IMAGEN" for result in file_results),
                },
                "coverage": {
                    "found": len(selected),
                    "expected": 68,
                    "missing": 68 - len(selected),
                    "duplicate_combinations": sum(
                        1 for key in expected_keys() if len(groups.get(key, [])) > 1
                    ),
                },
                "source_counts": source_counts,
                "meta_summary": meta_summary,
                "verify_meta_totals": verified_meta,
                "verify_meta_counted_rows": counted_meta,
                "verify_registros": verified_registros,
                "verification_problems": verification_issues,
                "out_of_scope_rows": sum(
                    len(result.orphan_rows) for result in file_results
                ),
                "issues": len(issues),
            },
            ensure_ascii=False,
            indent=2,
        )
    )

    if args.open_output:
        opener = getattr(os, "startfile", None)
        if opener is not None:
            try:
                opener(str(output_dir))
            except OSError:
                pass


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, NotADirectoryError, PermissionError, OSError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
