from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, PlainTextResponse
from pydantic import BaseModel
import fitz  # PyMuPDF
import re
import html
import traceback
import base64
from typing import List, Dict, Any, Optional
from fastapi.staticfiles import StaticFiles


app = FastAPI(
    title="API Planes de Transmisión",
    description="Procesa PDFs de planes de transmisión y devuelve tabla limpia de Ventas, Promos y Cortes.",
    version="1.5.0",
)


app.mount("/addin", StaticFiles(directory="addin"), name="addin")

# ============================================================
# MODELOS
# ============================================================

class PDFBase64Request(BaseModel):
    filename: str
    content_base64: str


# ============================================================
# UTILIDADES DE TIEMPO
# ============================================================

def to_seconds(minutes: int, seconds: int) -> int:
    return int(minutes) * 60 + int(seconds)


def seconds_to_hms(total_seconds: int) -> str:
    total_seconds = int(total_seconds or 0)
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h}:{m:02d}:{s:02d}"


def hms_to_seconds(value: str) -> int:
    match = re.fullmatch(r"(\d{2}):(\d{2}):(\d{2})", str(value or "").strip())

    if not match:
        raise ValueError(f"Hora invalida: {value}")

    hours, minutes, seconds = [int(part) for part in match.groups()]
    return hours * 3600 + minutes * 60 + seconds


def duration_between_hms(start_hms: str, end_hms: str) -> int:
    start = hms_to_seconds(start_hms)
    end = hms_to_seconds(end_hms)

    if end < start:
        end += 24 * 3600

    return end - start


# ============================================================
# LECTURA PDF
# ============================================================

def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text_parts = []

        for page in doc:
            text_parts.append(page.get_text("text"))

        text = "\n".join(text_parts)

        if not text.strip():
            raise ValueError("El PDF no devolvió texto. Puede ser escaneado o imagen.")

        return text

    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f"No se pudo leer el PDF: {str(e)}"
        )


# ============================================================
# HEADER
# ============================================================

def extract_header_info(text: str) -> Dict[str, str]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    fecha = ""
    horario = ""
    hora_inicio = ""
    hora_fin = ""
    programa = ""
    estado = ""

    for line in lines:
        if re.search(
            r"Lunes|Martes|Mi[eé]rcoles|Jueves|Viernes|S[aá]bado|Domingo",
            line,
            re.I,
        ):
            fecha = line

        match = re.search(
            r"De:\s*(\d{2}:\d{2}:\d{2})\s*a\s*(\d{2}:\d{2}:\d{2})\s*(.+)",
            line,
            re.I,
        )

        if match:
            hora_inicio = match.group(1)
            hora_fin = match.group(2)
            horario = f"{hora_inicio} a {hora_fin}"
            programa = match.group(3).strip()

    estado_match = re.search(
        r"(DISPONIBLE|SATURADO):\s*(\d+\s*hh\s*\d+\s*min\.?\s*\d+\s*seg)",
        text,
        re.I,
    )

    if estado_match:
        estado = f"{estado_match.group(1).upper()}: {estado_match.group(2)}"

    return {
        "fecha": fecha,
        "horario": horario,
        "hora_inicio": hora_inicio,
        "hora_fin": hora_fin,
        "programa": programa,
        "estado": estado,
    }


# ============================================================
# BLOQUES
# ============================================================

BLOCK_LOOKAHEAD_LINES = 4


def normalize_spaces(value: str) -> str:
    return " ".join(str(value or "").split())


def marker_debug(markers: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [
        {
            "block_number": marker["block_number"],
            "declared_seconds": marker["declared_seconds"],
            "raw_line": marker["raw_line"],
            "index": marker["index"],
        }
        for marker in markers
    ]


def parser_http_error(message: str, markers: Optional[List[Dict[str, Any]]] = None) -> HTTPException:
    return HTTPException(
        status_code=422,
        detail={
            "message": message,
            "debug_bloques": marker_debug(markers or []),
        },
    )


def find_block_markers(lines: List[str]) -> List[Dict[str, Any]]:
    """
    Detecta encabezados de bloque aunque PyMuPDF parta el texto en varias lineas.
    Se mira la linea actual mas cuatro lineas siguientes para capturar duraciones
    separadas como "4" / "min. 29 seg." o "4 min." / "29 seg.".
    """
    markers = []

    for idx, line in enumerate(lines):
        normalized = normalize_spaces(line)

        if "BLOQUE" not in normalized.upper() or "CORTE" not in normalized.upper():
            continue

        window_lines = [line]

        for next_line in lines[idx + 1:idx + BLOCK_LOOKAHEAD_LINES + 1]:
            normalized_next = normalize_spaces(next_line).upper()

            if "BLOQUE" in normalized_next and "CORTE" in normalized_next:
                break

            window_lines.append(next_line)

        raw_window = " | ".join(item.strip() for item in window_lines if item.strip())
        normalized_window = normalize_spaces(" ".join(window_lines))

        block_match = re.search(r"\bBLOQUE\s+(\d+)\b", normalized_window, re.I)
        corte_match = re.search(r"\bCorte\s*:", normalized_window, re.I)

        if not block_match or not corte_match:
            continue

        after_corte = normalized_window[corte_match.end():]
        duration_match = re.search(
            r"\b(\d+)\s*min\.?\s*(\d+)\s*seg\.?",
            after_corte,
            re.I,
        )

        if not duration_match:
            continue

        markers.append({
            "index": idx,
            "block_number": int(block_match.group(1)),
            "declared_seconds": to_seconds(
                int(duration_match.group(1)),
                int(duration_match.group(2)),
            ),
            "raw_line": raw_window,
        })

    return markers


def validate_block_markers(markers: List[Dict[str, Any]]) -> None:
    if not markers:
        raise parser_http_error(
            "No se encontraron bloques con formato reconocible 'BLOQUE X Corte : ... min. ... seg.'.",
            markers,
        )

    first_block = markers[0]["block_number"]

    if first_block != 1:
        raise parser_http_error(
            f"El primer bloque detectado es BLOQUE {first_block}; se esperaba BLOQUE 1.",
            markers,
        )

    detected = [marker["block_number"] for marker in markers]
    expected = list(range(1, len(detected) + 1))

    if detected != expected:
        raise parser_http_error(
            "Hay saltos o desorden en la numeracion de bloques. "
            f"Detectados: {detected}. Esperados: {expected}.",
            markers,
        )


# ============================================================
# DETECCIÓN DE PIEZAS
# ============================================================

def is_integer_line(line: str) -> bool:
    return re.fullmatch(r"\d+", line.strip()) is not None


def is_time_line(line: str) -> bool:
    line = line.strip()
    return re.search(r"\b\d{2}:\d{2}:\d{2}\b", line) is not None


def extract_piece_inline(line: str) -> Optional[Dict[str, Any]]:
    """
    Caso cuando PyMuPDF devuelve la pieza en una sola línea.

    Ejemplo:
    NAC 09:55:13 C P1631... DIGITAL 0 30 -98891
    NAC 09:54:43 P AHORA CAIGO AV HOY DVCAM 0 30
    """

    upper_line = line.upper()

    if "BLOQUE" in upper_line and "CORTE" in upper_line:
        return None

    normalized = normalize_spaces(line)

    type_match = re.search(
        r"\b\d{2}:\d{2}:\d{2}\s+([CP])\b(?P<tail>.*)$",
        normalized,
        re.I,
    )

    if not type_match:
        return None

    piece_type = type_match.group(1).upper()
    tail = type_match.group("tail") or ""

    number_pairs = re.findall(r"\b(\d+)\s+(\d+)\b", tail)

    if not number_pairs:
        return None

    minutes, seconds = number_pairs[-1]
    duration_seconds = to_seconds(int(minutes), int(seconds))

    return {
        "type": piece_type,
        "seconds": duration_seconds,
        "raw_line": normalized,
    }


def summarize_segment(lines: List[str]) -> Dict[str, Any]:
    """
    Suma ventas/promos dentro de un segmento.

    Soporta dos formatos:

    1) Formato en una sola línea:
       NAC 09:55:13 C P1631... DIGITAL 0 30

    2) Formato separado por columnas:
       C
       NAC 09:55:13
       30
       0

    En varios PDFs de estos planes, PyMuPDF devuelve:
       tipo
       hora
       segundos
       minutos

    Por eso para el formato separado:
       seconds = línea i+2
       minutes = línea i+3
    """

    ventas = 0
    promos = 0
    pieces = []

    i = 0

    while i < len(lines):
        line = lines[i].strip()

        # -------------------------------
        # Caso 1: una sola línea
        # -------------------------------
        inline_piece = extract_piece_inline(line)

        if inline_piece:
            if inline_piece["seconds"] > 0:
                if inline_piece["type"] == "C":
                    ventas += inline_piece["seconds"]
                elif inline_piece["type"] == "P":
                    promos += inline_piece["seconds"]

                pieces.append(inline_piece)

            i += 1
            continue

        # -------------------------------
        # Caso 1b: pieza partida en varias lineas.
        # Ejemplo:
        # NAC 23:50:18 C
        # 0 30
        # -------------------------------
        if is_time_line(line) and re.search(r"\b[CP]\b", line):
            combined_lines = lines[i:i + 3]
            combined_piece = extract_piece_inline(" ".join(combined_lines))

            if combined_piece and combined_piece["seconds"] > 0:
                if combined_piece["type"] == "C":
                    ventas += combined_piece["seconds"]
                elif combined_piece["type"] == "P":
                    promos += combined_piece["seconds"]

                pieces.append(combined_piece)
                i += len(combined_lines)
                continue

        # -------------------------------
        # Caso 2: columnas separadas
        # C / P
        # NAC 09:55:13
        # segundos
        # minutos
        # -------------------------------
        if (
            line in ("C", "P")
            and i + 3 < len(lines)
            and is_time_line(lines[i + 1])
            and is_integer_line(lines[i + 2])
            and is_integer_line(lines[i + 3])
        ):
            piece_type = line

            seconds = int(lines[i + 2].strip())
            minutes = int(lines[i + 3].strip())

            duration_seconds = to_seconds(minutes, seconds)

            if duration_seconds > 0:
                if piece_type == "C":
                    ventas += duration_seconds
                elif piece_type == "P":
                    promos += duration_seconds

                pieces.append({
                    "type": piece_type,
                    "seconds": duration_seconds,
                    "raw_line": " | ".join(lines[i:i + 4]),
                })

            i += 4
            continue

        i += 1

    return {
        "ventas_seconds": ventas,
        "promos_seconds": promos,
        "total_seconds": ventas + promos,
        "pieces": pieces,
    }


# ============================================================
# FILAS Y TABLA LIMPIA
# ============================================================

def make_row(
    concepto: str,
    ventas_seconds: int,
    promos_seconds: int,
    declarado_seconds: Optional[int] = None,
) -> Dict[str, Any]:

    total_seconds = ventas_seconds + promos_seconds

    if declarado_seconds is None:
        declarado_text = ""
        cuadra = True
        difference_seconds = 0
    else:
        declarado_text = seconds_to_hms(declarado_seconds)
        difference_seconds = total_seconds - declarado_seconds
        cuadra = difference_seconds == 0

    return {
        "concepto": concepto,
        "ventas": seconds_to_hms(ventas_seconds) if ventas_seconds > 0 else "",
        "promos": seconds_to_hms(promos_seconds) if promos_seconds > 0 else "",
        "corte": seconds_to_hms(total_seconds),
        "ventas_seconds": ventas_seconds,
        "promos_seconds": promos_seconds,
        "total_seconds": total_seconds,
        "declarado_pdf": declarado_text,
        "cuadra": cuadra,
        "difference_seconds": difference_seconds,
    }


def calculate_content_seconds(header: Dict[str, str], total_cuts_seconds: int) -> Dict[str, Any]:
    start_hms = header.get("hora_inicio", "")
    end_hms = header.get("hora_fin", "")

    if not start_hms or not end_hms:
        return {
            "seconds": None,
            "warning": "No se pudo calcular Contenido porque no se encontro hora de inicio y fin del programa.",
        }

    try:
        program_seconds = duration_between_hms(start_hms, end_hms)
    except ValueError as exc:
        return {
            "seconds": None,
            "warning": f"No se pudo calcular Contenido: {str(exc)}.",
        }

    content_seconds = program_seconds - int(total_cuts_seconds or 0)

    if content_seconds < 0:
        return {
            "seconds": None,
            "program_seconds": program_seconds,
            "warning": (
                "No se pudo calcular Contenido: el tiempo de cortes "
                f"({seconds_to_hms(total_cuts_seconds)}) supera la duracion del programa "
                f"({seconds_to_hms(program_seconds)})."
            ),
        }

    return {
        "seconds": content_seconds,
        "program_seconds": program_seconds,
        "warning": "",
    }


def build_email_table_html(result: dict) -> str:
    """
    Tabla HTML pensada para correo Outlook / Power Automate.
    Usa estilos en línea para que Outlook respete bordes, negritas y alineación.
    """

    programa = html.escape(str(result.get("programa", "")))

    table_style = (
        "border-collapse:collapse;"
        "font-family:Arial, sans-serif;"
        "font-size:13px;"
        "color:#000000;"
        "mso-table-lspace:0pt;"
        "mso-table-rspace:0pt;"
    )

    th_style = (
        "border:1px solid #555;"
        "padding:6px 10px;"
        "font-weight:bold;"
        "text-align:center;"
        "vertical-align:middle;"
        "background-color:#ffffff;"
        "color:#000000;"
        "white-space:normal;"
    )

    td_label_style = (
        "border:1px solid #555;"
        "padding:6px 10px;"
        "text-align:left;"
        "vertical-align:middle;"
        "background-color:#ffffff;"
        "color:#000000;"
        "white-space:nowrap;"
    )

    td_time_style = (
        "border:1px solid #555;"
        "padding:6px 10px;"
        "text-align:right;"
        "vertical-align:middle;"
        "background-color:#ffffff;"
        "color:#000000;"
        "white-space:nowrap;"
    )

    td_time_bold_style = td_time_style + "font-weight:bold;"
    td_label_bold_style = td_label_style + "font-weight:bold;"

    html_rows = f"""
    <tr>
        <th style="{th_style}">{programa}</th>
        <th style="{th_style}">Ventas</th>
        <th style="{th_style}">Promos</th>
        <th style="{th_style}">Corte</th>
    </tr>
    """

    for row in result.get("tabla", []):
        concepto = html.escape(str(row.get("concepto", "")))
        ventas = html.escape(str(row.get("ventas", "")))
        promos = html.escape(str(row.get("promos", "")))
        corte = html.escape(str(row.get("corte", "")))

        if concepto.lower() in ("totales", "contenido"):
            label_style = td_label_bold_style
            ventas_style = td_time_bold_style
            promos_style = td_time_bold_style
            corte_style = td_time_bold_style
        else:
            label_style = td_label_style
            ventas_style = td_time_style
            promos_style = td_time_style
            corte_style = td_time_bold_style

        html_rows += f"""
        <tr>
            <td style="{label_style}">{concepto}</td>
            <td style="{ventas_style}">{ventas}</td>
            <td style="{promos_style}">{promos}</td>
            <td style="{corte_style}">{corte}</td>
        </tr>
        """

    return f'<table style="{table_style}" cellpadding="0" cellspacing="0">{html_rows}</table>'


def build_email_result_html(result: dict) -> str:
    programa = html.escape(str(result.get("programa", "")))
    fecha = html.escape(str(result.get("fecha", "")))
    horario = html.escape(str(result.get("horario", "")))
    table_html = result.get("html_table") or build_email_table_html(result)

    title_style = (
        "font-family:Arial, sans-serif;"
        "font-size:18px;"
        "font-weight:bold;"
        "line-height:22px;"
        "color:#000000;"
        "margin:0 0 6px 0;"
    )
    info_style = (
        "font-family:Arial, sans-serif;"
        "font-size:13px;"
        "line-height:18px;"
        "color:#222222;"
        "margin:0 0 14px 0;"
    )

    return f"""
    <div style="font-family:Arial, sans-serif;color:#000000;">
        <div style="{title_style}">{programa}</div>
        <div style="{info_style}">
            <div><strong>Fecha:</strong> {fecha}</div>
            <div><strong>Horario:</strong> {horario}</div>
        </div>
        {table_html}
    </div>
    """


def build_clean_table(text: str) -> Dict[str, Any]:
    raw_lines = [line.strip() for line in text.splitlines() if line.strip()]
    header = extract_header_info(text)
    markers = find_block_markers(raw_lines)
    validate_block_markers(markers)

    rows = []

    # ------------------------------------------------------------
    # PRESENTA = todo lo que está antes del Bloque 1.
    # Solo mostrar si tiene duración positiva.
    # ------------------------------------------------------------

    first_block_index = markers[0]["index"]
    presenta_lines = raw_lines[:first_block_index]
    presenta_summary = summarize_segment(presenta_lines)

    if presenta_summary["total_seconds"] > 0:
        rows.append(make_row(
            concepto="PRESENTA",
            ventas_seconds=presenta_summary["ventas_seconds"],
            promos_seconds=presenta_summary["promos_seconds"],
            declarado_seconds=None,
        ))

    # ------------------------------------------------------------
    # CORTES
    # Piezas después de cada bloque y antes del siguiente bloque.
    #
    # Último bloque:
    # - Si hay piezas reales después, se considera DESPIDE.
    # - Si no hay piezas reales, no se muestra.
    # ------------------------------------------------------------

    for i, marker in enumerate(markers):
        start = marker["index"] + 1
        end = markers[i + 1]["index"] if i + 1 < len(markers) else len(raw_lines)

        segment_lines = raw_lines[start:end]
        summary = summarize_segment(segment_lines)

        actual = summary["total_seconds"]
        declared = marker["declared_seconds"]
        is_last_block = i == len(markers) - 1

        # No mostrar cortes vacíos.
        if actual == 0:
            continue

        if is_last_block:
            concepto = "DESPIDE"
        else:
            concepto = f"Corte {marker['block_number']}"

        rows.append(make_row(
            concepto=concepto,
            ventas_seconds=summary["ventas_seconds"],
            promos_seconds=summary["promos_seconds"],
            declarado_seconds=declared,
        ))

    # ------------------------------------------------------------
    # TOTALES
    # ------------------------------------------------------------

    total_ventas = sum(row["ventas_seconds"] for row in rows)
    total_promos = sum(row["promos_seconds"] for row in rows)
    total_general = sum(row["total_seconds"] for row in rows)

    rows.append({
        "concepto": "Totales",
        "ventas": seconds_to_hms(total_ventas),
        "promos": seconds_to_hms(total_promos),
        "corte": seconds_to_hms(total_general),
        "ventas_seconds": total_ventas,
        "promos_seconds": total_promos,
        "total_seconds": total_general,
        "declarado_pdf": "",
        "cuadra": True,
    })

    content_info = calculate_content_seconds(header, total_general)
    contenido_seconds = content_info.get("seconds")

    if contenido_seconds is not None:
        rows.append({
            "concepto": "Contenido",
            "ventas": "",
            "promos": "",
            "corte": seconds_to_hms(contenido_seconds),
            "ventas_seconds": 0,
            "promos_seconds": 0,
            "total_seconds": int(contenido_seconds),
            "declarado_pdf": "",
            "cuadra": True,
            "difference_seconds": 0,
        })

    # ------------------------------------------------------------
    # ADVERTENCIAS
    # ------------------------------------------------------------

    warnings = []

    if content_info.get("warning"):
        warnings.append(content_info["warning"])

    for row in rows:
        if row["concepto"] in ("Totales", "Contenido"):
            continue

        if row["declarado_pdf"] and not row["cuadra"]:
            difference = int(row.get("difference_seconds", 0))
            abs_difference = abs(difference)
            severity = (
                "Diferencia mayor a 1 segundo"
                if abs_difference > 1
                else "Diferencia de 1 segundo, posible redondeo o extraccion"
            )
            warnings.append(
                f"{row['concepto']} no cuadra: PDF declara {row['declarado_pdf']} "
                f"pero la suma da {row['corte']} ({severity}; diferencia {difference:+d}s)."
            )

    clean_rows = []

    for row in rows:
        clean_rows.append({
            "concepto": row["concepto"],
            "ventas": row["ventas"],
            "promos": row["promos"],
            "corte": row["corte"],
            "declarado_pdf": row["declarado_pdf"],
            "cuadra": row["cuadra"],
        })

    result = {
        "programa": header.get("programa", ""),
        "fecha": header.get("fecha", ""),
        "horario": header.get("horario", ""),
        "estado": header.get("estado", ""),
        "duracion_programa": (
            seconds_to_hms(content_info["program_seconds"])
            if content_info.get("program_seconds") is not None
            else ""
        ),
        "contenido": seconds_to_hms(contenido_seconds) if contenido_seconds is not None else "",
        "contenido_seconds": contenido_seconds,
        "tabla": clean_rows,
        "advertencias": warnings,
        "debug_bloques": marker_debug(markers),
    }

    result["html_table"] = build_email_table_html(result)
    result["html_email"] = build_email_result_html(result)

    return result


# ============================================================
# HTML PARA VISTA WEB
# ============================================================

def result_to_html(result: dict) -> str:
    rows_html = ""

    for row in result["tabla"]:
        concepto = html.escape(str(row.get("concepto", "")))
        ventas = html.escape(str(row.get("ventas", "")))
        promos = html.escape(str(row.get("promos", "")))
        corte = html.escape(str(row.get("corte", "")))
        declarado = html.escape(str(row.get("declarado_pdf", "")))
        cuadra = bool(row.get("cuadra", True))

        if concepto.lower() in ("totales", "contenido"):
            row_class = "total"
        elif not cuadra:
            row_class = "error"
        else:
            row_class = ""

        estado_cuadra = "✅" if cuadra else "❌"

        rows_html += f"""
        <tr class="{row_class}">
            <td>{concepto}</td>
            <td>{ventas}</td>
            <td>{promos}</td>
            <td><strong>{corte}</strong></td>
            <td>{declarado}</td>
            <td>{estado_cuadra}</td>
        </tr>
        """

    warnings_html = ""

    if result.get("advertencias"):
        warnings_html = "<div class='warnings'><h3>Advertencias</h3><ul>"

        for warning in result["advertencias"]:
            warnings_html += f"<li>{html.escape(str(warning))}</li>"

        warnings_html += "</ul></div>"

    programa = html.escape(str(result.get("programa", "")))
    fecha = html.escape(str(result.get("fecha", "")))
    horario = html.escape(str(result.get("horario", "")))
    estado = html.escape(str(result.get("estado", "")))

    return f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Resumen Plan de Transmisión</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background: #f5f6fa;
                padding: 30px;
                color: #222;
            }}

            .container {{
                max-width: 1050px;
                margin: auto;
                background: white;
                padding: 25px;
                border-radius: 12px;
                box-shadow: 0 4px 18px rgba(0,0,0,0.08);
            }}

            h1 {{
                margin-bottom: 5px;
                font-size: 24px;
            }}

            .info {{
                margin-bottom: 20px;
                color: #555;
                line-height: 1.5;
            }}

            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 20px;
                font-size: 15px;
            }}

            th {{
                background: #20242a;
                color: white;
                padding: 10px;
                text-align: left;
            }}

            td {{
                border-bottom: 1px solid #ddd;
                padding: 10px;
            }}

            td:nth-child(2),
            td:nth-child(3),
            td:nth-child(4),
            td:nth-child(5),
            td:nth-child(6) {{
                text-align: right;
            }}

            .total {{
                background: #eef3ff;
                font-weight: bold;
            }}

            .error {{
                background: #ffecec;
            }}

            .warnings {{
                margin-top: 20px;
                padding: 15px;
                background: #fff4d6;
                border-left: 5px solid #f0b400;
            }}

            .back {{
                display: inline-block;
                margin-top: 20px;
                text-decoration: none;
                color: #0066cc;
            }}

            .copy-note {{
                margin-top: 12px;
                font-size: 13px;
                color: #666;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>{programa}</h1>

            <div class="info">
                <div><strong>Fecha:</strong> {fecha}</div>
                <div><strong>Horario:</strong> {horario}</div>
                <div><strong>Estado:</strong> {estado}</div>
            </div>

            <table>
                <thead>
                    <tr>
                        <th>Concepto</th>
                        <th>Ventas</th>
                        <th>Promos</th>
                        <th>Corte</th>
                        <th>Declarado PDF</th>
                        <th>Cuadra</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>

            <div class="copy-note">
                Puedes seleccionar la tabla, copiarla y pegarla en Excel o en un correo.
            </div>

            {warnings_html}

            <a class="back" href="/">Procesar otro PDF</a>
        </div>
    </body>
    </html>
    """


def error_to_html(error_message: str, detail: str = "") -> str:
    return f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Error</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background: #f5f6fa;
                padding: 40px;
            }}
            .container {{
                max-width: 900px;
                margin: auto;
                background: white;
                padding: 30px;
                border-radius: 12px;
                box-shadow: 0 4px 18px rgba(0,0,0,0.08);
            }}
            .error {{
                background: #ffecec;
                padding: 15px;
                border-left: 5px solid #d90000;
                margin-top: 20px;
                white-space: pre-wrap;
                font-family: Consolas, monospace;
                font-size: 13px;
            }}
            a {{
                color: #0066cc;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Ocurrió un error procesando el PDF</h1>
            <p>{html.escape(error_message)}</p>
            <div class="error">{html.escape(detail)}</div>
            <p><a href="/">Volver</a></p>
        </div>
    </body>
    </html>
    """


# ============================================================
# PRUEBAS INTERNAS DEL PARSER
# ============================================================

def run_parser_self_tests() -> None:
    block_cases = [
        (
            ["NAC 09:30:00 ESTA MANANA - BLOQUE 1 Corte : 4 min. 50 seg. 24 43"],
            1,
            290,
        ),
        (
            ["PREVIA MUNDIAL 2026 DIFERIDO - BLOQUE 1 Corte : 4", "min. 29 seg."],
            1,
            269,
        ),
        (
            ["NUESTRO MUNDIAL AL DIA - BLOQUE 1 Corte : 4 min.", "27 seg."],
            1,
            267,
        ),
        (
            ["PARTIDO MUNDIAL 2026 DIFERIDO - BLOQUE 1 Corte : PRESENTACIN", "2 min. 47 seg."],
            1,
            167,
        ),
        (
            ["PARTIDO MUNDIAL 2026 DIFERIDO - BLOQUE 6 Corte : DESPEDIDA", "0 min. 30 seg."],
            6,
            30,
        ),
        (
            ["NUESTRO MUNDIAL AL DIA - BLOQUE 5 Corte : 0 min. 0", "seg."],
            5,
            0,
        ),
    ]

    for lines, expected_block, expected_seconds in block_cases:
        markers = find_block_markers(lines)
        assert len(markers) == 1, lines
        assert markers[0]["block_number"] == expected_block, lines
        assert markers[0]["declared_seconds"] == expected_seconds, lines

    piece_lines = [
        "NAC 22:10:54 C TEST 0 5",
        "_____UIO 22:20:28 C TEST 0 15",
        "***** GYE 11:43:02 P TEST 0 20",
        "DIGITAL adi NAC 22:14:14 C TEST 0 15",
        "DIGITAL ADIC NAC 15:46:20 C TEST 0 31",
        "P1700-PAPA-VISA-MUNDIAL-BPICHINCHA DIGITAL -99076",
        '(30")',
        "NAC 23:50:18 C 0 30",
        "NAC 00:00:00 C 0 0",
    ]
    summary = summarize_segment(piece_lines)
    assert summary["ventas_seconds"] == 96, summary
    assert summary["promos_seconds"] == 20, summary

    despedida_text = "\n".join([
        "NAC 08:00:00 C PRESENTA 0 30",
        "PROGRAMA - BLOQUE 1 Corte : 2 min. 47 seg.",
        "NAC 08:10:00 C VENTA 2 17",
        "NAC 08:12:00 P PROMO 0 30",
        "PROGRAMA - BLOQUE 2 Corte : DESPEDIDA",
        "0 min. 30 seg.",
        "NAC 08:20:00 C DESPIDE 0 30",
    ])
    result = build_clean_table(despedida_text)
    conceptos = [row["concepto"] for row in result["tabla"]]
    assert conceptos == ["PRESENTA", "Corte 1", "DESPIDE", "Totales"], conceptos

    contenido_text = "\n".join([
        "De: 10:00:00 a 11:00:00 PROGRAMA TEST",
        "NAC 10:00:00 C PRESENTA 0 30",
        "PROGRAMA TEST - BLOQUE 1 Corte : 1 min. 0 seg.",
        "NAC 10:10:00 C VENTA 1 0",
        "PROGRAMA TEST - BLOQUE 2 Corte : DESPEDIDA",
        "0 min. 30 seg.",
        "NAC 10:50:00 C DESPIDE 0 30",
    ])
    result = build_clean_table(contenido_text)
    conceptos = [row["concepto"] for row in result["tabla"]]
    assert conceptos == ["PRESENTA", "Corte 1", "DESPIDE", "Totales", "Contenido"], conceptos
    assert result["duracion_programa"] == "1:00:00", result
    assert result["contenido"] == "0:58:00", result


# ============================================================
# ENDPOINTS
# ============================================================

@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Procesar Plan de Transmisión</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background: #f5f6fa;
                padding: 40px;
            }

            .container {
                max-width: 600px;
                margin: auto;
                background: white;
                padding: 30px;
                border-radius: 12px;
                box-shadow: 0 4px 18px rgba(0,0,0,0.08);
            }

            h1 {
                font-size: 24px;
                margin-bottom: 8px;
            }

            p {
                color: #555;
                margin-bottom: 24px;
            }

            input[type="file"] {
                margin-bottom: 20px;
                width: 100%;
            }

            button {
                background: #20242a;
                color: white;
                border: none;
                padding: 12px 20px;
                border-radius: 8px;
                cursor: pointer;
                font-size: 15px;
            }

            button:hover {
                background: #000;
            }

            .links {
                margin-top: 24px;
                font-size: 14px;
                line-height: 1.6;
            }

            .links a {
                color: #0066cc;
                text-decoration: none;
            }

            code {
                background: #f0f0f0;
                padding: 2px 5px;
                border-radius: 4px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Procesar Plan de Transmisión</h1>
            <p>Sube un PDF del plan para generar la tabla limpia de ventas, promos y cortes.</p>

            <form action="/procesar-plan-tabla" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required>
                <br>
                <button type="submit">Generar tabla</button>
            </form>

            <div class="links">
                Endpoint JSON archivo normal: <code>POST /procesar-plan</code><br>
                Endpoint JSON Power Automate: <code>POST /procesar-plan-base64</code><br>
                El JSON ahora incluye: <code>html_table</code><br>
                Ver texto crudo del PDF: <code>POST /debug-texto</code><br>
                Estado API: <a href="/health">/health</a><br>
                Documentación técnica: <a href="/docs">/docs</a>
            </div>
        </div>
    </body>
    </html>
    """


@app.get("/health")
def health():
    return {
        "status": "ok",
        "message": "API activa",
        "version": "1.5.0"
    }


@app.post("/procesar-plan")
async def procesar_plan(file: UploadFile = File(...)):
    """
    Endpoint para pruebas manuales o integraciones con multipart/form-data.
    Devuelve JSON, incluyendo html_table.
    """

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Solo se aceptan archivos PDF.")

    pdf_bytes = await file.read()
    text = extract_text_from_pdf(pdf_bytes)
    result = build_clean_table(text)

    return result


@app.post("/procesar-plan-base64")
async def procesar_plan_base64(payload: PDFBase64Request):
    """
    Endpoint recomendado para Power Automate.
    Recibe el PDF como base64 dentro de JSON.

    Body esperado:
    {
      "filename": "archivo.pdf",
      "content_base64": "JVBERi0x..."
    }

    Devuelve JSON, incluyendo html_table.
    """

    try:
        if not payload.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail="Solo se aceptan archivos PDF.")

        content = payload.content_base64.strip()

        # Por si algún sistema envía formato tipo:
        # data:application/pdf;base64,JVBERi0x...
        if "," in content and "base64" in content[:100].lower():
            content = content.split(",", 1)[1]

        pdf_bytes = base64.b64decode(content)

        text = extract_text_from_pdf(pdf_bytes)
        result = build_clean_table(text)

        return result

    except HTTPException:
        raise

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Error procesando PDF base64: {str(e)}"
        )


@app.post("/procesar-plan-tabla", response_class=HTMLResponse)
async def procesar_plan_tabla(file: UploadFile = File(...)):
    """
    Endpoint visual.
    Devuelve HTML con tabla para revisar en navegador.
    """

    try:
        if not file.filename.lower().endswith(".pdf"):
            return HTMLResponse(
                error_to_html("Solo se aceptan archivos PDF."),
                status_code=400,
            )

        pdf_bytes = await file.read()
        text = extract_text_from_pdf(pdf_bytes)
        result = build_clean_table(text)

        return result_to_html(result)

    except HTTPException as e:
        return HTMLResponse(
            error_to_html(str(e.detail)),
            status_code=e.status_code,
        )

    except Exception:
        detail = traceback.format_exc()
        return HTMLResponse(
            error_to_html("Error interno del servidor.", detail),
            status_code=500,
        )


@app.post("/debug-texto", response_class=PlainTextResponse)
async def debug_texto(file: UploadFile = File(...)):
    """
    Endpoint para ver cómo PyMuPDF está leyendo el PDF.
    Sirve para ajustar el parser.
    """

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Solo se aceptan archivos PDF.")

    pdf_bytes = await file.read()
    text = extract_text_from_pdf(pdf_bytes)
    raw_lines = [line.strip() for line in text.splitlines() if line.strip()]
    markers = find_block_markers(raw_lines)
    debug_lines = [
        "",
        "",
        "================ DEBUG_BLOQUES ================",
    ]

    for marker in marker_debug(markers):
        debug_lines.append(
            f"index={marker['index']} block_number={marker['block_number']} "
            f"declared_seconds={marker['declared_seconds']} raw_line={marker['raw_line']}"
        )

    return text + "\n".join(debug_lines)
