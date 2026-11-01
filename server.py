"""Cotizador de Pisos Blandos (revestimientos de muro MK) — lista de
materiales, sin IA, sin precios.

Replica la lógica de "SELEKTOR REVESTIMIENTO DE MUROS.xlsx" (hojas
"Selector"/"Matriz"/"BD Revestimientos"): el vendedor agrega sectores
(zona, dónde se instala, sobre qué superficie, si hay calor cerca, producto,
m²) y el sistema calcula, por producto, las cajas/piezas necesarias, el
adhesivo MK indicado con su rendimiento, y los complementos de instalación
(perfiles, fijaciones, sello de canterías, limpieza) — igual patrón que el
cotizador de pinturas, pero con un solo tamaño de envase por producto en vez
de combinar tineta/galón/¼.

El catálogo (29 revestimientos: SPC WALL, paneles alveolares, palillaje,
paneles acústicos/3D, palillaje exterior) vive en catalogo_revestimientos.json,
extraído 1 a 1 de las hojas "Matriz" y "BD Revestimientos" de esa planilla —
ver knowledge/ para el Excel original.
"""

from __future__ import annotations

import io
import json
import math
import os
import secrets
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
from fastapi import Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel
from starlette.middleware.sessions import SessionMiddleware

import db

app = FastAPI(title="Cotizador de Pisos Blandos")
db.init_db()

ESTADISTICAS_CLAVE = os.environ.get("ESTADISTICAS_CLAVE", "PisosBlandos2026")

ALLOWED_EMAIL_DOMAINS = {
    d.strip().lower()
    for d in os.environ.get("ALLOWED_EMAIL_DOMAINS", "mk.cl").split(",")
    if d.strip()
}

# Sesión simple (solo correo) — mismo mecanismo que el cotizador de pinturas.
SECRET_KEY = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.add_middleware(SessionMiddleware, secret_key=SECRET_KEY, same_site="lax", https_only=False)


def usuario_actual(request: Request) -> Optional[dict]:
    nombre = request.session.get("nombre")
    return {"nombre": nombre} if nombre else None


def requerir_login(request: Request) -> dict:
    u = usuario_actual(request)
    if not u:
        raise HTTPException(status_code=401, detail="No autenticado.")
    return u

LOGO_MK_PATH = Path(__file__).parent / "web" / "assets" / "logo-mk.png"

# Desperdicio por corte/recorte de paneles. La planilla original no fija un
# % — se asume 10% (más alto que el 5% de pinturas porque acá hay cortes de
# palillaje/plank en esquinas y vanos). Ajustar aquí si Tomás define otro.
MERMA = 0.10

# ---------------------------------------------------------------------------
# Catálogo (extraído de "SELEKTOR REVESTIMIENTO DE MUROS.xlsx" — hojas
# "Matriz" + "BD Revestimientos". Ver knowledge/ para el Excel original.)
# ---------------------------------------------------------------------------

CATALOGO_PATH = Path(__file__).parent / "catalogo_revestimientos.json"
with open(CATALOGO_PATH, encoding="utf-8") as _f:
    CATALOGO: Dict[str, dict] = json.load(_f)

# Lugar de instalación -> clave del dict "lugares" de cada producto en el
# catálogo. "Piso / terraza" existe en la planilla original como pregunta
# pero HOY ningún producto de este Excel la soporta (es "revestimiento de
# MURO") — se deja fuera de las opciones; si en el futuro se agregan pisos
# blandos de piso al catálogo, agregar "Piso / terraza": "piso_terraza" aquí.
LUGAR_CLAVE: Dict[str, str] = {
    "Muro interior": "muro_interior",
    "Muro exterior / fachada": "muro_exterior",
    "Baño con ducha": "bano_ducha",
    "Baño visita (sin ducha)": "bano_visita",
    "Cielo": "cielo",
    "Contenedor": "contenedor",
}
LUGAR_OPCIONES = list(LUGAR_CLAVE.keys())

# Igual a LUGAR_CLAVE pero con "Piso / terraza" incluido — solo para la
# tabla informativa "¿Dónde se puede instalar?" del Selector (hoja
# "Selector", sección 6), que SÍ muestra ese lugar aunque hoy ningún
# producto lo soporte. El formulario de cotizar (LUGAR_CLAVE) lo omite para
# no ofrecer una opción que nunca tiene productos disponibles.
LUGAR_INFO_CLAVE: Dict[str, str] = {**LUGAR_CLAVE, "Piso / terraza": "piso_terraza"}

SUPERFICIE_BASE_OPCIONES = [
    "Muro de cemento / hormigón",
    "Yeso cartón",
    "Placa Internit (fibrocemento)",
    "Placa de madera",
    "Cerámica o porcelanato",
]
SUPERFICIE_CERAMICA = "Cerámica o porcelanato"

# Notas de preparación por superficie (hoja "Selector", sección 5, columna
# "Preparación") — son genéricas, no cambian según el producto elegido.
PREPARACION_POR_SUPERFICIE: Dict[str, str] = {
    "Muro de cemento / hormigón": "Muro firme: revisar que no desprenda polvo ni arena; si desprende, aplicar mano de sellante.",
    "Yeso cartón": "Usar mano de aparejo sobre yeso cartón.",
    "Placa Internit (fibrocemento)": "Usar mano de aparejo sobre Internit.",
    "Placa de madera": "Aparejo (placas de yeso o madera).",
    "Cerámica o porcelanato": "Instalación sobre cerámica: realizar prueba de adherencia antes de instalar.",
}

CONDICION_CALOR_OPCIONES = ["Sin fuente de calor", "Calefacción en muro", "Seca toalla en baño"]
_CAMPO_POR_CONDICION_CALOR = {
    "Calefacción en muro": "calefaccion_muro",
    "Seca toalla en baño": "seca_toalla_bano",
}

# Qué productos se pueden ofrecer para cada lugar (hoja "Matriz", columna
# "RESULTADO PARA EL SECTOR ELEGIDO").
PRODUCTOS_POR_LUGAR: Dict[str, List[str]] = {
    etiqueta: [nombre for nombre, info in CATALOGO.items() if info["lugares"].get(clave)]
    for etiqueta, clave in LUGAR_CLAVE.items()
}

# Plantillas de sector con ícono, para agregar sectores típicos con un clic
# (mismo patrón que "Agregar sector rápido" del cotizador de pinturas). El
# m² siempre lo ingresa el vendedor.
AREAS_PLANTILLA = [
    {"nombre": "Baño visita", "icono": "🚻", "lugar": "Baño visita (sin ducha)",
     "superficie_base": "Cerámica o porcelanato", "condicion_calor": "Sin fuente de calor",
     "producto": "POWER"},
    {"nombre": "Baño con ducha", "icono": "🚿", "lugar": "Baño con ducha",
     "superficie_base": "Cerámica o porcelanato", "condicion_calor": "Sin fuente de calor",
     "producto": "ROCKO"},
    {"nombre": "Living / muro interior", "icono": "🛋️", "lugar": "Muro interior",
     "superficie_base": "Muro de cemento / hormigón", "condicion_calor": "Sin fuente de calor",
     "producto": "RIVIERA"},
    {"nombre": "Cielo", "icono": "☁️", "lugar": "Cielo",
     "superficie_base": "Placa de madera", "condicion_calor": "Sin fuente de calor",
     "producto": "OVIEDO"},
    {"nombre": "Fachada exterior", "icono": "🏠", "lugar": "Muro exterior / fachada",
     "superficie_base": "Muro de cemento / hormigón", "condicion_calor": "Sin fuente de calor",
     "producto": "QUIEBRAVISTA"},
    {"nombre": "Contenedor", "icono": "📦", "lugar": "Contenedor",
     "superficie_base": "Placa de madera", "condicion_calor": "Sin fuente de calor",
     "producto": "ENERGY / SOFT"},
]


# ---------------------------------------------------------------------------
# Modelos
# ---------------------------------------------------------------------------

class Sector(BaseModel):
    sector: str
    lugar: str
    superficie_base: str
    condicion_calor: str = "Sin fuente de calor"
    producto: str
    m2: float
    observaciones: str = ""


class CotizarRequest(BaseModel):
    cliente: str = ""
    vendedor: str = ""
    sectores: List[Sector]
    modificado_de: Optional[int] = None


# ---------------------------------------------------------------------------
# Lógica de cálculo
# ---------------------------------------------------------------------------

def _validar_sector(s: Sector) -> None:
    if s.lugar not in LUGAR_CLAVE:
        raise HTTPException(400, f"'{s.lugar}' no es un lugar de instalación válido.")
    if s.producto not in PRODUCTOS_POR_LUGAR.get(s.lugar, []):
        raise HTTPException(400, f"'{s.producto}' no se puede instalar en '{s.lugar}'.")
    if s.superficie_base not in SUPERFICIE_BASE_OPCIONES:
        raise HTTPException(400, f"'{s.superficie_base}' no es una superficie base válida.")
    if s.condicion_calor not in CONDICION_CALOR_OPCIONES:
        raise HTTPException(400, f"Sector '{s.sector}': condición de calor inválida.")
    if s.m2 <= 0:
        raise HTTPException(400, f"Sector '{s.sector}': la superficie debe ser mayor a 0.")


def _adhesivo_para(info: dict, superficie_base: str) -> Optional[dict]:
    """Adhesivo MK según superficie — "Cerámica o porcelanato" usa un
    adhesivo distinto (Salckh Puk) al resto de las superficies."""
    campo = "adhesivo_ceramica" if superficie_base == SUPERFICIE_CERAMICA else "adhesivo_estandar"
    adh = info.get(campo)
    if not adh or not adh.get("rendimiento_m2"):
        return None
    return adh


def _calcular(req: CotizarRequest) -> dict:
    for s in req.sectores:
        _validar_sector(s)

    m2_por_producto: Dict[str, float] = {}
    m2_por_producto_superficie: Dict[tuple, float] = {}
    detalle_por_area: "Dict[str, list]" = {}
    orden_sectores: List[str] = []
    m2_por_sector_producto: Dict[tuple, float] = {}
    m2_por_sector_producto_superficie: Dict[tuple, float] = {}
    advertencias_calor: List[dict] = []
    vistas_calor: set = set()

    for s in req.sectores:
        info = CATALOGO[s.producto]

        m2_por_producto[s.producto] = m2_por_producto.get(s.producto, 0) + s.m2
        clave_ps = (s.producto, s.superficie_base)
        m2_por_producto_superficie[clave_ps] = m2_por_producto_superficie.get(clave_ps, 0) + s.m2

        detalle_por_area.setdefault(s.sector, []).append(
            f"{s.lugar} ({s.m2} m², {s.producto}, sobre {s.superficie_base})"
        )
        if s.sector not in orden_sectores:
            orden_sectores.append(s.sector)
        clave_sp = (s.sector, s.producto)
        m2_por_sector_producto[clave_sp] = m2_por_sector_producto.get(clave_sp, 0) + s.m2
        clave_sps = (s.sector, s.producto, s.superficie_base)
        m2_por_sector_producto_superficie[clave_sps] = (
            m2_por_sector_producto_superficie.get(clave_sps, 0) + s.m2
        )

        campo_calor = _CAMPO_POR_CONDICION_CALOR.get(s.condicion_calor)
        if campo_calor:
            vista = (s.producto, s.condicion_calor)
            if vista not in vistas_calor:
                vistas_calor.add(vista)
                valor = (info.get(campo_calor) or "NO").strip()
                if valor.upper().startswith("NO"):
                    advertencias_calor.append({
                        "producto": s.producto,
                        "condicion": s.condicion_calor,
                        "detalle": valor,
                    })

    resumen_areas = [
        {"sector": sector, "detalle": "; ".join(items)}
        for sector, items in detalle_por_area.items()
    ]

    def _items_recinto(sector: str) -> list:
        items = []
        for (sec, nombre_producto), m2_sin_merma in m2_por_sector_producto.items():
            if sec != sector:
                continue
            info = CATALOGO[nombre_producto]
            equiv = info["equivalencia_m2"]
            m2_necesarios = m2_sin_merma * (1 + MERMA)
            if equiv:
                items.append({
                    "producto": nombre_producto,
                    "item": info["unidad_venta"],
                    "cantidad": math.ceil(round(m2_necesarios / equiv, 6)),
                })
        for (sec, nombre_producto, superficie), m2_sin_merma in m2_por_sector_producto_superficie.items():
            if sec != sector:
                continue
            info = CATALOGO[nombre_producto]
            adh = _adhesivo_para(info, superficie)
            if not adh:
                continue
            m2_necesarios = m2_sin_merma * (1 + MERMA)
            items.append({
                "producto": nombre_producto,
                "item": f"Adhesivo {adh['sku']}" + (f" ({adh['nombre']})" if adh.get("nombre") else ""),
                "cantidad": math.ceil(round(m2_necesarios / adh["rendimiento_m2"], 6)),
            })
        return items

    resumen_por_recinto = [
        {"sector": sector, "items": _items_recinto(sector)} for sector in orden_sectores
    ]

    productos_resultado = []
    for nombre, m2_total in m2_por_producto.items():
        info = CATALOGO[nombre]
        equiv = info["equivalencia_m2"]
        m2_necesarios = m2_total * (1 + MERMA)
        cantidad = math.ceil(round(m2_necesarios / equiv, 6)) if equiv else None
        productos_resultado.append({
            "producto": nombre,
            "familia": info["familia"],
            "materialidad": info["materialidad"],
            "medidas": info["medidas"],
            "m2_total": round(m2_total, 2),
            "m2_necesarios": round(m2_necesarios, 2),
            "unidad_venta": info["unidad_venta"],
            "cantidad_necesaria": cantidad,
            "cantidad_por_caja_texto": info["cantidad_por_caja_texto"],
            "requisitos": {
                "perfiles": info.get("requiere_perfiles"),
                "perfiles_sku": info.get("perfil_sku_texto"),
                "fijaciones": info.get("requiere_fijaciones"),
                "sello_canterias": info.get("sello_canterias"),
                "sello_sku": info.get("sku_sello"),
                "limpieza": info.get("limpieza_tipo"),
                "limpiador_sku": info.get("limpiador_sku"),
                "pintable": info.get("pintable"),
                "no_pega_motivo": info.get("no_pega_motivo"),
            },
        })

    adhesivos_resultado = []
    for (nombre, superficie), m2_total in m2_por_producto_superficie.items():
        info = CATALOGO[nombre]
        adh = _adhesivo_para(info, superficie)
        if not adh:
            continue
        m2_necesarios = m2_total * (1 + MERMA)
        envases = math.ceil(round(m2_necesarios / adh["rendimiento_m2"], 6))
        adhesivos_resultado.append({
            "producto": nombre,
            "superficie_base": superficie,
            "sku": adh["sku"],
            "nombre_comercial": adh.get("nombre"),
            "envase": adh.get("envase"),
            "rendimiento_m2": adh["rendimiento_m2"],
            "m2_necesarios": round(m2_necesarios, 2),
            "envases_necesarios": envases,
        })
        alt = info.get("alternativa_mercado")
        if alt:
            adhesivos_resultado[-1]["alternativa_mercado"] = alt

    return {
        "cliente": req.cliente,
        "vendedor": req.vendedor,
        "m2_total": round(sum(m2_por_producto.values()), 2),
        "productos": productos_resultado,
        "adhesivos": adhesivos_resultado,
        "resumen_areas": resumen_areas,
        "resumen_por_recinto": resumen_por_recinto,
        "advertencias_calor": advertencias_calor,
    }


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

TITULO_FILL = PatternFill("solid", fgColor="1F2A44")
TITULO_FONT = Font(color="FFFFFF", bold=True, size=14)
ENCABEZADO_FILL = PatternFill("solid", fgColor="1F2A44")
ENCABEZADO_FONT = Font(bold=True, color="FFFFFF")
SUBTITULO_FONT = Font(bold=True, size=12, color="1F2A44")
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

NCOLS = 8


def _escribir_tabla(ws, fila_inicio: int, encabezados: list, filas: list) -> int:
    for col, titulo in enumerate(encabezados, start=1):
        celda = ws.cell(row=fila_inicio, column=col, value=titulo)
        celda.fill = ENCABEZADO_FILL
        celda.font = ENCABEZADO_FONT
        celda.border = _BORDER
    fila = fila_inicio + 1
    for valores in filas:
        for col, valor in enumerate(valores, start=1):
            celda = ws.cell(row=fila, column=col, value=valor)
            celda.border = _BORDER
        fila += 1
    return fila


def _construir_excel(resultado: dict, numero: int, modificado_de: Optional[int] = None) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lista de materiales"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
    ws["A1"] = f"COTIZADOR DE PISOS BLANDOS — LISTA DE MATERIALES N° {numero}"
    ws["A1"].fill = TITULO_FILL
    ws["A1"].font = TITULO_FONT
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40
    if LOGO_MK_PATH.exists():
        try:
            img = XLImage(str(LOGO_MK_PATH))
            img.width = img.height = 36
            ws.add_image(img, "A1")
        except Exception:  # noqa: BLE001 — el logo es decorativo
            pass

    ws["A3"] = "Cliente:"
    ws["B3"] = resultado["cliente"]
    ws["D3"] = "Vendedor:"
    ws["E3"] = resultado["vendedor"]
    ws["A4"] = "Superficie total (m²):"
    ws["B4"] = resultado["m2_total"]
    for celda in ("A3", "D3", "A4"):
        ws[celda].font = Font(bold=True)
    if modificado_de:
        ws["A5"] = f"Modifica la cotización N° {modificado_de}."
        ws["A5"].font = Font(italic=True, color="666666")

    fila = 6
    filas_productos = [
        [p["producto"], p["familia"], p["m2_total"], p["m2_necesarios"],
         p["cantidad_necesaria"], p["unidad_venta"], p["cantidad_por_caja_texto"], p["medidas"]]
        for p in resultado["productos"]
    ]
    fila = _escribir_tabla(ws, fila, [
        "Revestimiento", "Familia", "m² totales", "m² con merma",
        "Cantidad a comprar", "Unidad de venta", "Formato de caja", "Medidas",
    ], filas_productos)

    fila += 1
    ws.cell(row=fila, column=1, value="Adhesivo necesario").font = SUBTITULO_FONT
    fila += 1
    filas_adh = [
        [a["producto"], a["superficie_base"], a["sku"], a.get("nombre_comercial") or "—",
         a["envases_necesarios"], f"Rend. {a['rendimiento_m2']} m²"]
        for a in resultado["adhesivos"]
    ]
    fila = _escribir_tabla(ws, fila, [
        "Producto", "Superficie base", "SKU adhesivo", "Nombre", "Envases necesarios", "Rendimiento",
    ], filas_adh)

    fila += 1
    ws.cell(row=fila, column=1, value="Resumen por área").font = SUBTITULO_FONT
    fila += 1
    filas_areas = [[a["sector"], a["detalle"]] for a in resultado["resumen_areas"]]
    fila = _escribir_tabla(ws, fila, ["Sector", "Detalle"], filas_areas)

    fila += 1
    ws.cell(row=fila, column=1, value="Resumen por recinto").font = SUBTITULO_FONT
    fila += 1
    for r in resultado["resumen_por_recinto"]:
        ws.cell(row=fila, column=1, value=r["sector"]).font = Font(bold=True, size=11)
        fila += 1
        filas_recinto = [[it["producto"], it["item"], it["cantidad"]] for it in r["items"]]
        fila = _escribir_tabla(ws, fila, ["Producto", "Ítem", "Cantidad"], filas_recinto)
        fila += 1

    fila += 1
    ws.cell(row=fila, column=1, value="Complementos de instalación (por producto)").font = SUBTITULO_FONT
    fila += 1
    filas_complementos = []
    for p in resultado["productos"]:
        r = p["requisitos"]
        if r.get("no_pega_motivo"):
            filas_complementos.append([p["producto"], "Fijación (no usa adhesivo)", r["no_pega_motivo"]])
        if r.get("perfiles"):
            filas_complementos.append([p["producto"], "Perfiles de terminación", f"{r['perfiles']} — {r.get('perfiles_sku') or ''}"])
        if r.get("fijaciones"):
            filas_complementos.append([p["producto"], "Fijaciones", r["fijaciones"]])
        if r.get("sello_canterias"):
            filas_complementos.append([p["producto"], "Sello de canterías", f"{r['sello_canterias']} — SKU {r.get('sello_sku') or 'N/A'}"])
        if r.get("pintable"):
            filas_complementos.append([p["producto"], "¿Se puede pintar?", r["pintable"]])
        filas_complementos.append([p["producto"], "Limpieza y mantención", f"{r.get('limpieza') or ''} — {r.get('limpiador_sku') or ''}"])
    fila = _escribir_tabla(ws, fila, ["Producto", "Condición", "Sugerido"], filas_complementos)

    if resultado["advertencias_calor"]:
        fila += 1
        ws.cell(row=fila, column=1, value="⚠ Advertencias de compatibilidad con calor").font = SUBTITULO_FONT
        fila += 1
        filas_adv = [[a["producto"], a["condicion"], a["detalle"]] for a in resultado["advertencias_calor"]]
        fila = _escribir_tabla(ws, fila, ["Producto", "Condición", "Detalle"], filas_adv)

    anchos = [26, 22, 12, 14, 16, 16, 30, 30]
    for i, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = ancho

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------

def _construir_pdf(resultado: dict, numero: int, modificado_de: Optional[int] = None) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import cm, mm
    from reportlab.platypus import Image as RLImage
    from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer,
                                     Table, TableStyle)

    styles = getSampleStyleSheet()
    estilo_normal = styles["Normal"]
    estilo_h2 = ParagraphStyle("h2", parent=styles["Heading2"], textColor=colors.HexColor("#1F2A44"))
    estilo_celda = ParagraphStyle("celda", parent=estilo_normal, fontSize=8.5, leading=10)

    def celda(texto):
        return Paragraph(str(texto), estilo_celda)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm, leftMargin=1.8 * cm, rightMargin=1.8 * cm,
    )
    elementos = []

    if LOGO_MK_PATH.exists():
        logo = RLImage(str(LOGO_MK_PATH), width=14 * mm, height=14 * mm)
        elementos.append(logo)
        elementos.append(Spacer(1, 0.2 * cm))

    titulo = Table([[f"COTIZADOR DE PISOS BLANDOS — N° {numero}"]], colWidths=[17 * cm])
    titulo.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1F2A44")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 14),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    elementos.append(titulo)
    elementos.append(Spacer(1, 0.5 * cm))

    elementos.append(Paragraph(
        f"<b>Cliente:</b> {resultado['cliente'] or '—'} &nbsp;&nbsp; "
        f"<b>Vendedor:</b> {resultado['vendedor'] or '—'} &nbsp;&nbsp; "
        f"<b>Superficie total:</b> {resultado['m2_total']} m²",
        estilo_normal,
    ))
    if modificado_de:
        elementos.append(Paragraph(
            f"Modifica la cotización N° {modificado_de}.",
            ParagraphStyle("mod", parent=estilo_normal, fontName="Helvetica-Oblique", textColor=colors.HexColor("#666666")),
        ))
    elementos.append(Spacer(1, 0.4 * cm))

    elementos.append(Paragraph("Lista de materiales", estilo_h2))
    filas = [["Revestimiento", "Familia", "m² totales", "Cantidad", "Unidad"]]
    for p in resultado["productos"]:
        filas.append([
            celda(p["producto"]), celda(p["familia"]), p["m2_total"],
            p["cantidad_necesaria"], p["unidad_venta"],
        ])
    tabla = Table(filas, colWidths=[4.5 * cm, 4 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm], repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E1F2")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 0.4 * cm))

    elementos.append(Paragraph("Adhesivo necesario", estilo_h2))
    filas = [["Producto", "Superficie base", "SKU", "Envases", "Rendimiento"]]
    for a in resultado["adhesivos"]:
        filas.append([celda(a["producto"]), celda(a["superficie_base"]), a["sku"], a["envases_necesarios"], f"{a['rendimiento_m2']} m²"])
    tabla = Table(filas, colWidths=[4 * cm, 4.5 * cm, 3 * cm, 2 * cm, 2.5 * cm], repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E1F2")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 0.4 * cm))

    elementos.append(Paragraph("Resumen por área", estilo_h2))
    for a in resultado["resumen_areas"]:
        elementos.append(Paragraph(f"<b>{a['sector']}:</b> {a['detalle']}", estilo_normal))
    elementos.append(Spacer(1, 0.4 * cm))

    elementos.append(Paragraph("Complementos de instalación", estilo_h2))
    for p in resultado["productos"]:
        r = p["requisitos"]
        partes = []
        if r.get("no_pega_motivo"):
            partes.append(r["no_pega_motivo"])
        if r.get("perfiles"):
            partes.append(f"Perfiles: {r['perfiles']}")
        if r.get("fijaciones"):
            partes.append(f"Fijaciones: {r['fijaciones']}")
        if r.get("sello_canterias"):
            partes.append(f"Sello de canterías: {r['sello_canterias']}")
        if r.get("pintable"):
            partes.append(f"Pintable: {r['pintable']}")
        partes.append(f"Limpieza: {r.get('limpieza') or ''} — {r.get('limpiador_sku') or ''}")
        elementos.append(Paragraph(f"<b>{p['producto']}:</b> " + " · ".join(partes), ParagraphStyle(
            "req", parent=estilo_normal, fontSize=8.5,
        )))

    if resultado["advertencias_calor"]:
        elementos.append(Spacer(1, 0.4 * cm))
        elementos.append(Paragraph("⚠ Advertencias de compatibilidad con calor", estilo_h2))
        for a in resultado["advertencias_calor"]:
            elementos.append(Paragraph(
                f"<b>{a['producto']} — {a['condicion']}:</b> {a['detalle']}",
                ParagraphStyle("warn", parent=estilo_normal, textColor=colors.HexColor("#8A2020")),
            ))

    doc.build(elementos)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------

@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/catalogo")
def catalogo(_: dict = Depends(requerir_login)):
    return {
        "lugar_opciones": LUGAR_OPCIONES,
        "lugar_info_opciones": list(LUGAR_INFO_CLAVE.keys()),
        "lugar_info_clave": LUGAR_INFO_CLAVE,
        "superficie_base_opciones": SUPERFICIE_BASE_OPCIONES,
        "superficie_ceramica": SUPERFICIE_CERAMICA,
        "preparacion_por_superficie": PREPARACION_POR_SUPERFICIE,
        "condicion_calor_opciones": CONDICION_CALOR_OPCIONES,
        "productos_por_lugar": PRODUCTOS_POR_LUGAR,
        "productos": CATALOGO,
        "areas_plantilla": AREAS_PLANTILLA,
    }


@app.post("/cotizar")
def cotizar(req: CotizarRequest, _: dict = Depends(requerir_login)):
    if not req.sectores:
        raise HTTPException(400, "Agrega al menos un sector.")
    return _calcular(req)


@app.post("/cotizar/generar")
def cotizar_generar(req: CotizarRequest, _: dict = Depends(requerir_login)):
    if not req.sectores:
        raise HTTPException(400, "Agrega al menos un sector.")
    resultado = _calcular(req)
    numero = db.guardar_cotizacion(
        req.cliente, req.vendedor, resultado["m2_total"],
        [s.dict() for s in req.sectores], resultado,
        modificado_de=req.modificado_de,
    )
    return {
        "numero": numero,
        "resultado": resultado,
        "nombre_excel": f"lista_materiales_pisos_blandos_{numero}.xlsx",
        "nombre_pdf": f"cotizacion_pisos_blandos_{numero}.pdf",
    }


@app.get("/cotizacion/{numero}")
def cotizacion_para_modificar(numero: int, _: dict = Depends(requerir_login)):
    registro = db.obtener_cotizacion(numero)
    if registro is None:
        raise HTTPException(404, "N° de cotización no encontrado.")
    return {
        "numero": registro["id"],
        "cliente": registro["cliente"],
        "vendedor": registro["vendedor"],
        "sectores": registro["sectores"],
    }


@app.get("/descarga/{numero}/excel")
def descarga_excel(numero: int, _: dict = Depends(requerir_login)):
    registro = db.obtener_cotizacion(numero)
    if registro is None:
        raise HTTPException(404, "N° de cotización no encontrado.")
    excel_bytes = _construir_excel(registro["resultado"], numero, registro.get("modificado_de"))
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=lista_materiales_pisos_blandos_{numero}.xlsx"},
    )


@app.get("/descarga/{numero}/pdf")
def descarga_pdf(numero: int, _: dict = Depends(requerir_login)):
    registro = db.obtener_cotizacion(numero)
    if registro is None:
        raise HTTPException(404, "N° de cotización no encontrado.")
    pdf_bytes = _construir_pdf(registro["resultado"], numero, registro.get("modificado_de"))
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename=cotizacion_pisos_blandos_{numero}.pdf"},
    )


class EstadisticasRequest(BaseModel):
    clave: str


def _requerir_clave(clave: str) -> None:
    if not secrets.compare_digest(clave, ESTADISTICAS_CLAVE):
        raise HTTPException(401, "Clave incorrecta.")


_NO_CACHE = {"Cache-Control": "no-store"}


@app.get("/")
def index_page(request: Request):
    if usuario_actual(request):
        return RedirectResponse("/cotizador", status_code=303)
    return FileResponse("web/login.html", headers=_NO_CACHE)


@app.post("/login")
def login(request: Request, correo: str = Form(...)):
    correo = correo.strip().lower()
    if "@" not in correo or correo.rsplit("@", 1)[1] not in ALLOWED_EMAIL_DOMAINS:
        raise HTTPException(403, "Correo no registrado.")
    request.session["nombre"] = correo
    return {"redirect": "/cotizador"}


@app.post("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/", status_code=303)


@app.get("/whoami")
def whoami(usuario: dict = Depends(requerir_login)):
    return usuario


@app.get("/cotizador")
def cotizador_page(request: Request):
    if not usuario_actual(request):
        return RedirectResponse("/", status_code=303)
    return FileResponse("web/index.html", headers=_NO_CACHE)


@app.get("/estadisticas")
def estadisticas_page(request: Request):
    if not usuario_actual(request):
        return RedirectResponse("/", status_code=303)
    return FileResponse("web/estadisticas.html", headers=_NO_CACHE)


@app.get("/buscar")
def buscar_page(request: Request):
    if not usuario_actual(request):
        return RedirectResponse("/", status_code=303)
    return FileResponse("web/buscar.html", headers=_NO_CACHE)


@app.get("/selector")
def selector_page(request: Request):
    if not usuario_actual(request):
        return RedirectResponse("/", status_code=303)
    return FileResponse("web/selector.html", headers=_NO_CACHE)


@app.post("/estadisticas/datos")
def estadisticas_datos(req: EstadisticasRequest, _: dict = Depends(requerir_login)):
    _requerir_clave(req.clave)
    return {"cotizaciones": db.listar_cotizaciones()}


app.mount("/", StaticFiles(directory="web", html=True), name="web")
