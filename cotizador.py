"""
Cotizador Internacional — Robertet Ingredients
================================================
App Streamlit de cotización de ingredientes naturales con lógica de margen y forex.

MODELO: se corre LOCALMENTE y guarda en tu carpeta de OneDrive sincronizada.
  - El repositorio NO contiene datos comerciales. Solo código.
  - El histórico vive en un archivo dentro de tu carpeta local de OneDrive
    (p.ej. "C:/Users/Tu/OneDrive - Robertet/cotizaciones/historico.xlsx").
  - La app lee ese archivo al iniciar. Al guardar una sesión, AGREGA las
    cotizaciones nuevas y reescribe el archivo en su lugar, dejando primero
    un respaldo con fecha. OneDrive lo sincroniza solo a la nube y al celular.
  - También existe el modo "subir archivo" (efímero) por si corres sin OneDrive.

Fórmula de precio (verificada):
  USD = (matl + ovrhd) / (1 - GM) / (1 - comm) / (1 - duty) / forex
"""

import io
import os
import re
import shutil
from datetime import date, datetime

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Configuración de página
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Cotizador Internacional — Robertet",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# Paleta e identidad visual (navy tinta + dorado, acentos especiados)
# ----------------------------------------------------------------------------
GOLD = "#C9A84C"
GOLD_HI = "#E4C476"
NAVY = "#0B1524"
PANEL = "#111F35"
PAPRIKA = "#C4553B"
TURMERIC = "#D99A2B"
GREEN = "#4CBF8B"
RED = "#E06C5F"

st.markdown(
    f"""
    <style>
      .stApp {{ background: radial-gradient(1200px 500px at 50% -220px, #16294a 0%, {NAVY} 55%); }}
      section[data-testid="stSidebar"] {{ background: {PANEL}; }}
      .spice-strip {{ height:4px; border-radius:2px;
         background: linear-gradient(90deg, {PAPRIKA}, {TURMERIC} 45%, {GOLD} 80%, {GOLD_HI});
         margin-bottom: 14px; }}
      .brand-title {{ font-family: Georgia, 'Times New Roman', serif; font-size:30px;
         font-weight:700; color:#EDF1F7; line-height:1.05; margin:0; }}
      .brand-sub {{ color:#93A3B8; font-size:13px; font-family: 'SF Mono', Consolas, monospace; }}
      .big-price {{ font-family: Georgia, serif; font-size:44px; font-weight:700;
         background: linear-gradient(100deg, {GOLD_HI}, {GOLD} 55%, {TURMERIC});
         -webkit-background-clip:text; background-clip:text; color:transparent; line-height:1; }}
      div[data-testid="stMetricValue"] {{ font-family: Georgia, serif; }}
      .stButton>button[kind="primary"] {{ background: linear-gradient(180deg, {GOLD_HI}, {GOLD});
         color:#1A1608; font-weight:700; border:none; }}
      .ref-card {{ background:{PANEL}; border:1px solid #1E3050; border-left:3px solid {GOLD};
         border-radius:8px; padding:12px 14px; margin-top:8px; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# i18n
# ----------------------------------------------------------------------------
T = {
    "es": {
        "brand": "Cotizador Internacional",
        "sub": "Robertet Ingredients · sesión efímera",
        "lang_label": "Idioma",
        "load_header": "Cargar histórico de la sesión",
        "load_help": "Sube tu archivo de cotizaciones (CSV o Excel). Se usa solo en esta sesión y se borra al cerrar. No se guarda en el servidor.",
        "load_prompt": "Sube el archivo histórico para comenzar",
        "load_none": "Sin histórico cargado — la búsqueda y las referencias están inactivas. Puedes cotizar igual, pero sin comparación histórica.",
        "loaded_ok": "Histórico cargado: {n} registros ({start} → {end})",
        "load_error": "No se pudo leer el archivo: {err}",
        "detect_ok": "✓ Columnas reconocidas: {cols}",
        "detect_missing_essential": "⚠️ Faltan columnas esenciales: {cols}. El archivo puede no funcionar bien. Usa la plantilla.",
        "detect_missing_important": "ℹ️ No se detectaron: {cols}. Se rellenan por defecto; revisa si tu archivo las tenía con otro nombre.",
        "template_hint": "¿Formato distinto? Descarga la plantilla (abajo) y vacía tus datos ahí.",
        "session_privacy": "🔒 Local: la app corre en tu equipo. El archivo vive en tu OneDrive; nada se envía a servidores externos.",
        "clear_data": "Borrar histórico de la sesión",
        "tab_search": "🔎 Buscar historial",
        "tab_quote": "🧮 Nueva cotización",
        "tab_batch": "📦 Cotización al cliente",
        "tab_insights": "📊 Analítica",
        "search_ph": "Buscar producto, código o cliente… (ej: paprika symrise · 8101 · cardamom)",
        "search_hint": "Los términos se combinan (AND). Escribe producto + cliente, código, o solo producto.",
        "no_results": "Sin resultados. Prueba menos términos o revisa la ortografía.",
        "results_count": "{n} resultados",
        "col_date": "Fecha", "col_customer": "Cliente", "col_product": "Producto",
        "col_code": "Código", "col_cost": "Costo CAD", "col_gm": "GM",
        "col_quote": "Quote USD", "col_uom": "UOM", "col_note": "Nota",
        "use_row": "Usar como base ↓",
        "form_header": "Datos de la cotización",
        "customer": "Cliente", "product": "Producto *", "code": "Código / ENR", "uom": "UOM",
        "product_pick": "Producto (elige o escribe para filtrar) *",
        "code_pick": "Código / ENR (elige o filtra)",
        "write_new": "✏️ Escribir nuevo…",
        "product_new": "Nuevo producto *",
        "code_new": "Nuevo código / ENR",
        "autofill_hint": "💡 Elige un producto del histórico para autocompletar su última base (costo, GM, forex…).",
        "customer_pick": "Cliente (elige o escribe para filtrar)",
        "customer_new": "Nuevo cliente",
        "code_family": "🔗 Códigos ligados a «{p}»",
        "final_price": "PRECIO FINAL A COTIZAR",
        "final_price_label": "Precio final (USD) — decisión del vendedor",
        "final_price_help": "El cotizador sugiere un precio, pero tú decides. Este es el valor que se guarda en la cotización.",
        "suggested": "Sugerido: ${s}",
        "final_vs_calc": "{arrow} {d} vs. sugerido",
        "grp_client": "── Productos de {c} ──",
        "grp_catalog": "── Resto del catálogo ──",
        "client_base": "🎯 Base de {c} (última vez): {info}",
        "client_never": "ℹ️ {c} nunca compró «{p}» — usando base general del producto.",
        "client_last_price": "💰 {c} pagó ${p} la última vez ({date})",
        "customer_pick": "Cliente (elige o escribe para filtrar)",
        "customer_new": "Nuevo cliente",
        "code_filtered_hint": "Códigos de este producto",
        "code_all_hint": "Todos los códigos",
        "from_code_hint": "↑ Producto tomado del código elegido",
        "cost": "Costo (CAD/kg) *", "ovrhd": "Overhead", "gm": "GM (0.30 = 30%)",
        "comm": "Comisión", "duty": "Duty / Arancel", "forex": "Forex CAD→USD",
        "spec": "Especificación", "incoterm": "Incoterm", "moq": "MOQ", "validity": "Validez",
        "comment": "Comentario / nota interna",
        "calc_price": "PRECIO CALCULADO", "per_unit": "USD / {u}",
        "add_quote": "➕ Agregar a la cotización de la sesión",
        "added_ok": "Agregada: {p} @ ${q}",
        "missing": "Falta producto o costo, o los datos rompen la fórmula (GM/comm/duty ≥ 1).",
        "ref_header": "Referencia histórica",
        "ref_none": "Sin coincidencias históricas para este producto.",
        "ref_last": "Último cotizado", "ref_avg_gm": "GM promedio", "ref_range": "Rango histórico",
        "ref_count": "{n} cotizaciones previas",
        "warn_header": "Puntos de atención",
        "warn_gm_low": "GM bajo ({v}) — verifica que cubra costos.",
        "warn_gm_high": "GM alto ({v}) — riesgo de quedar fuera de mercado.",
        "warn_dev": "Precio {d} vs. último histórico (${h}). Justifica la variación.",
        "warn_fx": "Forex = {v}. Revisa el tipo de cambio.",
        "warn_unit": "UOM '{u}' difiere del histórico ('{h}'). Verifica unidades.",
        "batch_header": "Cotización de la sesión",
        "batch_empty": "Aún no agregaste ítems. Ve a «Nueva cotización» y agrega productos.",
        "batch_count": "{n} ítems en la cotización",
        "remove": "Quitar",
        "export_internal": "⬇️ Exportar interno (con costos y márgenes)",
        "export_client": "⬇️ Exportar para cliente (sin costos)",
        "export_client_help": "Versión limpia: solo ítem, especificación, precio, moneda, incoterm, validez y MOQ. Sin costos ni márgenes.",
        "update_header": "Actualizar histórico (modo nube)",
        "update_desc": "Descarga tu histórico completo con las cotizaciones de esta sesión ya agregadas. Reemplaza tu archivo con este y súbelo la próxima vez.",
        "update_btn_csv": "⬇️ Descargar histórico actualizado (CSV)",
        "update_btn_xlsx": "⬇️ Descargar histórico actualizado (Excel)",
        "update_summary": "{base} previas + {new} de esta sesión = {total} registros",
        "update_dups_note": "({d} duplicadas se omitieron)",
        "update_no_hist": "Carga primero un histórico para poder combinarlo con la sesión.",
        "update_no_cart": "No hay cotizaciones nuevas en la sesión. El histórico se descarga tal cual está.",
        "client_meta": "Datos de la cotización al cliente",
        "client_name": "Cliente", "client_ref": "Referencia / RFQ", "client_currency": "Moneda",
        "client_prepared": "Preparado por",
        "insights_header": "Analítica del histórico",
        "insights_none": "Carga un histórico para ver la analítica.",
        "ins_by_customer": "GM promedio por cliente (top 15)",
        "ins_top_products": "Productos más cotizados (top 15)",
        "ins_over_time": "Cotizaciones por mes",
        # --- precarga desde búsqueda ---
        "use_in_quote": "📋 Usar en nueva cotización",
        "preloaded_ok": "✓ «{p}» precargado — ve a la pestaña «Nueva cotización» para ajustar y cotizar.",
        "sel_to_preload": "Selecciona la fila a precargar",
        "row_label": "{date} · {cust} · {prod} · ${q}",
        # --- analytics de producto / ventas ---
        "ins_product_header": "📊 Analítica de «{p}»",
        "ins_source_quotes": "📊 Fuente: cotizaciones históricas",
        "ins_source_sales": "💰 Fuente: ventas reales",
        "ins_no_product": "Selecciona un producto en «Nueva cotización» para ver su analítica específica. Mostrando vista general.",
        "ins_prod_min": "Mínimo", "ins_prod_max": "Máximo", "ins_prod_avg": "Promedio",
        "ins_prod_last": "Última", "ins_prod_count": "N.º registros",
        "ins_price_time": "Precio en el tiempo",
        "ins_by_client_prod": "Por cliente (precio promedio)",
        "ins_no_prod_data": "Sin datos históricos para «{p}» todavía.",
        "sales_header": "Archivo de ventas (opcional)",
        "sales_help": "Sube ventas reales (CSV/Excel). Si está presente, la analítica usa precios de venta reales en vez de cotizados.",
        "sales_prompt": "Sube archivo de ventas",
        "sales_ok": "Ventas cargadas: {n} registros",
        "sales_error": "No se pudo leer ventas: {err}",
        "sales_clear": "Quitar archivo de ventas",
        "footer": "Fórmula: (costo + overhead) / (1−GM) / (1−comisión) / (1−duty) / forex",
        # --- modo OneDrive local ---
        "mode_label": "Fuente del histórico",
        "mode_local": "📁 Archivo en OneDrive (local)",
        "mode_upload": "⬆️ Subir archivo (efímero)",
        "local_path": "Ruta al archivo en tu OneDrive",
        "local_path_help": "Pega la ruta completa a tu archivo dentro de la carpeta OneDrive sincronizada. Ej: C:/Users/Tu/OneDrive - Robertet/cotizaciones/historico.xlsx",
        "local_load": "Cargar desde OneDrive",
        "local_ok": "Leído desde OneDrive: {n} registros",
        "local_missing": "No encontré el archivo en esa ruta. Verifica que exista y que OneDrive lo haya sincronizado.",
        "local_save_header": "Guardar en OneDrive",
        "local_save_btn": "💾 Agregar sesión al histórico y reescribir",
        "local_saved": "Guardado ✓ — {n} nuevas agregadas. Total: {total}. OneDrive sincronizará el cambio.",
        "local_saved_dups": "Guardado ✓ — {n} nuevas agregadas ({d} eran duplicadas y se omitieron). Total: {total}.",
        "local_backup": "Respaldo previo: {b}",
        "local_nothing": "No hay cotizaciones nuevas en la sesión para guardar.",
        "local_no_path": "Primero carga un archivo desde OneDrive (necesito saber dónde escribir).",
        "local_write_err": "Error al escribir: {err}. ¿El archivo está abierto en Excel? Ciérralo e intenta de nuevo.",
        "local_tip": "💡 Cierra el archivo en Excel antes de guardar, o la reescritura fallará.",
    },
    "en": {
        "brand": "International Quoter",
        "sub": "Robertet Ingredients · ephemeral session",
        "lang_label": "Language",
        "load_header": "Load session history",
        "load_help": "Upload your quotes file (CSV or Excel). Used only in this session and cleared on close. Nothing is stored on the server.",
        "load_prompt": "Upload the history file to begin",
        "load_none": "No history loaded — search and references are inactive. You can still quote, without historical comparison.",
        "loaded_ok": "History loaded: {n} records ({start} → {end})",
        "load_error": "Could not read file: {err}",
        "detect_ok": "✓ Recognized columns: {cols}",
        "detect_missing_essential": "⚠️ Missing essential columns: {cols}. The file may not work well. Use the template.",
        "detect_missing_important": "ℹ️ Not detected: {cols}. Filled with defaults; check if your file named them differently.",
        "template_hint": "Different format? Download the template (below) and paste your data there.",
        "session_privacy": "🔒 Local: the app runs on your machine. The file lives in your OneDrive; nothing is sent to external servers.",
        "clear_data": "Clear session history",
        "tab_search": "🔎 Search history",
        "tab_quote": "🧮 New quote",
        "tab_batch": "📦 Client quotation",
        "tab_insights": "📊 Analytics",
        "search_ph": "Search product, code or customer… (e.g. paprika symrise · 8101 · cardamom)",
        "search_hint": "Terms combine (AND). Type product + customer, code, or product alone.",
        "no_results": "No results. Try fewer terms or check spelling.",
        "results_count": "{n} results",
        "col_date": "Date", "col_customer": "Customer", "col_product": "Product",
        "col_code": "Code", "col_cost": "Cost CAD", "col_gm": "GM",
        "col_quote": "Quote USD", "col_uom": "UOM", "col_note": "Note",
        "use_row": "Use as base ↓",
        "form_header": "Quote details",
        "customer": "Customer", "product": "Product *", "code": "Code / ENR", "uom": "UOM",
        "product_pick": "Product (pick or type to filter) *",
        "code_pick": "Code / ENR (pick or filter)",
        "write_new": "✏️ Type new…",
        "product_new": "New product *",
        "code_new": "New code / ENR",
        "autofill_hint": "💡 Pick a product from history to autofill its last base (cost, GM, forex…).",
        "customer_pick": "Customer (pick or type to filter)",
        "customer_new": "New customer",
        "code_family": "🔗 Codes linked to «{p}»",
        "final_price": "FINAL PRICE TO QUOTE",
        "final_price_label": "Final price (USD) — seller's decision",
        "final_price_help": "The tool suggests a price, but you decide. This is the value saved in the quote.",
        "suggested": "Suggested: ${s}",
        "final_vs_calc": "{arrow} {d} vs. suggested",
        "grp_client": "── {c}'s products ──",
        "grp_catalog": "── Rest of catalog ──",
        "client_base": "🎯 {c}'s base (last time): {info}",
        "client_never": "ℹ️ {c} never bought «{p}» — using general product base.",
        "client_last_price": "💰 {c} paid ${p} last time ({date})",
        "customer_pick": "Customer (pick or type to filter)",
        "customer_new": "New customer",
        "code_filtered_hint": "Codes for this product",
        "code_all_hint": "All codes",
        "from_code_hint": "↑ Product taken from the selected code",
        "cost": "Cost (CAD/kg) *", "ovrhd": "Overhead", "gm": "GM (0.30 = 30%)",
        "comm": "Commission", "duty": "Duty / Tariff", "forex": "Forex CAD→USD",
        "spec": "Specification", "incoterm": "Incoterm", "moq": "MOQ", "validity": "Validity",
        "comment": "Comment / internal note",
        "calc_price": "CALCULATED PRICE", "per_unit": "USD / {u}",
        "add_quote": "➕ Add to session quotation",
        "added_ok": "Added: {p} @ ${q}",
        "missing": "Missing product or cost, or data breaks the formula (GM/comm/duty ≥ 1).",
        "ref_header": "Historical reference",
        "ref_none": "No historical matches for this product.",
        "ref_last": "Last quoted", "ref_avg_gm": "Average GM", "ref_range": "Historical range",
        "ref_count": "{n} previous quotes",
        "warn_header": "Points of attention",
        "warn_gm_low": "Low GM ({v}) — verify it covers costs.",
        "warn_gm_high": "High GM ({v}) — risk of being priced out.",
        "warn_dev": "Price {d} vs. last historical (${h}). Justify the variation.",
        "warn_fx": "Forex = {v}. Check the exchange rate.",
        "warn_unit": "UOM '{u}' differs from history ('{h}'). Verify units.",
        "batch_header": "Session quotation",
        "batch_empty": "No items yet. Go to «New quote» and add products.",
        "batch_count": "{n} items in the quotation",
        "remove": "Remove",
        "export_internal": "⬇️ Export internal (with costs & margins)",
        "export_client": "⬇️ Export for client (no costs)",
        "export_client_help": "Clean version: only item, spec, price, currency, incoterm, validity, MOQ. No costs or margins.",
        "update_header": "Update history (cloud mode)",
        "update_desc": "Download your full history with this session's quotes already appended. Replace your file with this one and upload it next time.",
        "update_btn_csv": "⬇️ Download updated history (CSV)",
        "update_btn_xlsx": "⬇️ Download updated history (Excel)",
        "update_summary": "{base} previous + {new} from this session = {total} records",
        "update_dups_note": "({d} duplicates skipped)",
        "update_no_hist": "Load a history first so it can be merged with the session.",
        "update_no_cart": "No new session quotes. History downloads as-is.",
        "client_meta": "Client quotation details",
        "client_name": "Customer", "client_ref": "Reference / RFQ", "client_currency": "Currency",
        "client_prepared": "Prepared by",
        "insights_header": "History analytics",
        "insights_none": "Load a history file to see analytics.",
        "ins_by_customer": "Average GM by customer (top 15)",
        "ins_top_products": "Most quoted products (top 15)",
        "ins_over_time": "Quotes per month",
        "use_in_quote": "📋 Use in new quote",
        "preloaded_ok": "✓ «{p}» preloaded — go to the «New quote» tab to adjust and quote.",
        "sel_to_preload": "Select the row to preload",
        "row_label": "{date} · {cust} · {prod} · ${q}",
        "ins_product_header": "📊 Analytics for «{p}»",
        "ins_source_quotes": "📊 Source: historical quotes",
        "ins_source_sales": "💰 Source: real sales",
        "ins_no_product": "Select a product in «New quote» to see its specific analytics. Showing general view.",
        "ins_prod_min": "Min", "ins_prod_max": "Max", "ins_prod_avg": "Average",
        "ins_prod_last": "Last", "ins_prod_count": "Records",
        "ins_price_time": "Price over time",
        "ins_by_client_prod": "By customer (avg price)",
        "ins_no_prod_data": "No historical data for «{p}» yet.",
        "sales_header": "Sales file (optional)",
        "sales_help": "Upload real sales (CSV/Excel). If present, analytics uses real sale prices instead of quoted ones.",
        "sales_prompt": "Upload sales file",
        "sales_ok": "Sales loaded: {n} records",
        "sales_error": "Could not read sales: {err}",
        "sales_clear": "Remove sales file",
        "footer": "Formula: (cost + overhead) / (1−GM) / (1−commission) / (1−duty) / forex",
        # --- OneDrive local mode ---
        "mode_label": "History source",
        "mode_local": "📁 OneDrive file (local)",
        "mode_upload": "⬆️ Upload file (ephemeral)",
        "local_path": "Path to your OneDrive file",
        "local_path_help": "Paste the full path to your file inside the synced OneDrive folder. E.g. C:/Users/You/OneDrive - Robertet/quotes/history.xlsx",
        "local_load": "Load from OneDrive",
        "local_ok": "Read from OneDrive: {n} records",
        "local_missing": "File not found at that path. Check it exists and OneDrive has synced it.",
        "local_save_header": "Save to OneDrive",
        "local_save_btn": "💾 Append session to history and rewrite",
        "local_saved": "Saved ✓ — {n} new added. Total: {total}. OneDrive will sync the change.",
        "local_saved_dups": "Saved ✓ — {n} new added ({d} duplicates skipped). Total: {total}.",
        "local_backup": "Previous backup: {b}",
        "local_nothing": "No new session quotes to save.",
        "local_no_path": "Load a file from OneDrive first (I need to know where to write).",
        "local_write_err": "Write error: {err}. Is the file open in Excel? Close it and retry.",
        "local_tip": "💡 Close the file in Excel before saving, or the rewrite will fail.",
    },
}

# ----------------------------------------------------------------------------
# Lógica de negocio
# ----------------------------------------------------------------------------
COLS = ["date", "customer", "broker", "product", "code", "matl", "ovrhd",
        "gm", "comm", "duty", "forex", "uom", "quote", "comment"]
NUMERIC = ["matl", "ovrhd", "gm", "comm", "duty", "forex", "quote"]


def normalize_code(code):
    """Normaliza un código a formato NR/ENR.
    - Números puros -> se antepone NR (11064 -> NR11064).
    - Empieza con nr/enr (cualquier caja) -> se pasa a NR/ENR en mayúscula.
    - Formatos raros (X, ????, 08xxx, t1062...) -> se dejan intactos.
    """
    if code is None:
        return ""
    c = str(code).strip()
    if not c or c.lower() in ("nan", "total", "subtotal", "none", "n/a"):
        return ""
    up = c.upper()
    if up.startswith("ENR"):
        return "ENR" + c[3:]
    if up.startswith("NR"):
        return "NR" + c[2:]
    if re.match(r"^\d+$", c):
        return "NR" + c
    return c


def calc_quote(matl, ovrhd, gm, comm, duty, forex):
    """USD = (matl + ovrhd) / (1-GM) / (1-comm) / (1-duty) / forex."""
    fx = forex or 1.0
    if gm >= 1 or comm >= 1 or duty >= 1:
        return float("nan")
    return (matl + ovrhd) / (1 - gm) / (1 - comm) / (1 - duty) / fx


COLUMN_SYNONYMS = {
    # canónico -> lista de posibles nombres (en minúscula) que la app reconocerá
    "date":     ["date", "fecha", "quote date", "fecha cotizacion", "fecha cotización", "día", "dia"],
    "customer": ["customer", "cliente", "client", "account", "cuenta", "empresa", "razon social", "razón social"],
    "broker":   ["broker", "corredor", "intermediario", "agente", "distribuidor"],
    "product":  ["product", "producto", "item", "articulo", "artículo", "descripcion", "descripción",
                 "description", "material", "nombre producto", "product name"],
    "code":     ["code", "codigo", "código", "item code", "sku", "enr", "ref", "referencia",
                 "cod", "clave", "part number", "product code"],
    "matl":     ["matl", "material", "costo", "cost", "material cost", "costo material",
                 "costo materia", "materia prima", "raw cost", "costo cad"],
    "ovrhd":    ["ovrhd", "overhead", "gasto", "gastos", "indirecto", "indirectos", "oh"],
    "gm":       ["gm", "gross margin", "margen", "margen bruto", "margin", "mrg", "% margen", "margen %"],
    "comm":     ["comm", "commission", "comision", "comisión", "comm%", "% comision"],
    "duty":     ["duty", "arancel", "aranceles", "tariff", "impuesto", "impuestos", "tax"],
    "forex":    ["forex", "fx", "tipo de cambio", "tc", "exchange", "exchange rate", "cambio", "tasa cambio"],
    "uom":      ["uom", "unidad", "unit", "unidad medida", "unidad de medida", "u/m", "medida"],
    "quote":    ["quote", "cotizacion", "cotización", "precio", "price", "precio venta",
                 "precio cotizado", "quoted price", "pvp", "precio final", "sell price"],
    "comment":  ["comment", "comentario", "comentarios", "nota", "notas", "observacion",
                 "observación", "observaciones", "remarks", "note"],
}


def remap_columns(df):
    """Traduce los nombres de columna de un archivo ajeno a los canónicos de la app,
    buscando sinónimos en español/inglés. Devuelve (df_renombrado, columnas_detectadas).
    - Coincidencia exacta primero; luego coincidencia parcial (el sinónimo contenido en el encabezado).
    - No pisa una columna canónica ya asignada.
    """
    df = df.copy()
    lower_cols = {c: str(c).strip().lower() for c in df.columns}
    rename = {}
    used_targets = set()

    # 1) coincidencia exacta
    for canon, syns in COLUMN_SYNONYMS.items():
        if canon in used_targets:
            continue
        for orig, low in lower_cols.items():
            if low in syns and orig not in rename:
                rename[orig] = canon
                used_targets.add(canon)
                break
    # 2) coincidencia parcial (sinónimo contenido en el encabezado), sin repetir destino
    for canon, syns in COLUMN_SYNONYMS.items():
        if canon in used_targets:
            continue
        for orig, low in lower_cols.items():
            if orig in rename:
                continue
            if any(s in low for s in syns if len(s) >= 3):
                rename[orig] = canon
                used_targets.add(canon)
                break

    df = df.rename(columns=rename)
    return df, sorted(used_targets)


def normalize_history(df):
    """Ajusta un dataframe cargado al esquema esperado, tolerante a columnas faltantes.
    Primero intenta mapear nombres de columna ajenos (ES/EN) a los canónicos.
    """
    df = df.copy()
    # Paso 1: intento de mapeo flexible ANTES de bajar a minúsculas crudas
    df, _detected = remap_columns(df)
    # Paso 2: normalización estándar
    df.columns = [str(c).strip().lower() for c in df.columns]
    for c in COLS:
        if c not in df.columns:
            df[c] = "" if c not in NUMERIC else 0.0
    df = df[COLS]
    for c in NUMERIC:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["date", "customer", "broker", "product", "code", "uom", "comment"]:
        df[c] = df[c].fillna("").astype(str)
    df["uom"] = df["uom"].replace("", "kg")
    df["code"] = df["code"].apply(normalize_code)  # normaliza códigos a NR/ENR
    return df


def analyze_upload(df):
    """Inspecciona un archivo recién subido y reporta qué columnas se reconocieron
    y cuáles faltan de las esenciales. Para dar feedback claro al usuario."""
    _, detected = remap_columns(df)
    essential = ["product", "quote"]      # mínimo para que el histórico sea útil
    important = ["customer", "code", "matl", "gm", "forex", "date"]
    missing_essential = [c for c in essential if c not in detected]
    missing_important = [c for c in important if c not in detected]
    return detected, missing_essential, missing_important


def load_uploaded(file, return_diagnostics=False):
    name = file.name.lower()
    if name.endswith(".csv"):
        df_raw = pd.read_csv(file)
    elif name.endswith((".xlsx", ".xls")):
        df_raw = pd.read_excel(file)
    else:
        raise ValueError("Formato no soportado (usa CSV o Excel).")
    diag = analyze_upload(df_raw)
    result = normalize_history(df_raw)
    if return_diagnostics:
        return result, diag
    return result


def _clean_sales_product(raw_name):
    """'ANISE OLEORESIN NR0102 [NR0102]' -> 'Anise Oleoresin'."""
    s = str(raw_name).strip()
    s = re.sub(r"\s*\[[^\]]*\]\s*$", "", s)                    # quitar [NR0102]
    s = re.sub(r"\s+(E?N?R)?\d+\s*$", "", s, flags=re.I)       # quitar código final
    s = re.sub(r"\s+", " ", s).strip()
    return s.title()


def _read_sales_r12(raw):
    """Desarma un reporte pivote tipo R12: doble encabezado (año + métrica),
    bloques de métricas repetidos por año, filas de identificación en col 0-7.
    Devuelve filas producto-año con precio real de venta.
    """
    year_row = raw.iloc[0]
    year_starts = {}
    for col in range(8, raw.shape[1]):
        v = year_row[col]
        if pd.notna(v):
            s = str(v).split(".")[0]
            if re.match(r"^20\d\d$", s):
                year_starts[int(s)] = col
    if not year_starts:
        return None

    ID = {"customer": 3, "product": 5, "item_code": 7}
    OFF_PRICE, OFF_QTY = 50, 34  # offsets dentro del bloque de cada año
    data = raw.iloc[2:].reset_index(drop=True)
    recs = []
    for year, start in year_starts.items():
        pc, qc = start + OFF_PRICE, start + OFF_QTY
        for _, r in data.iterrows():
            code = normalize_code(r[ID["item_code"]])
            if not code:                     # sólo filas con código real (evita totales/duplicados)
                continue
            prod = r[ID["product"]]
            if pd.isna(prod) or not str(prod).strip():
                continue
            price = r[pc] if pc < len(r) else None
            if not (pd.notna(price) and isinstance(price, (int, float)) and price > 0):
                continue
            qty = r[qc] if qc < len(r) else None
            cust = r[ID["customer"]]
            cust = re.sub(r"\s*\[[^\]]*\]\s*$", "", str(cust)).strip() if pd.notna(cust) else ""
            recs.append({
                "product": _clean_sales_product(prod), "code": code,
                "price": round(float(price), 2),
                "date": str(year), "customer": cust.title(),
                "qty": round(float(qty), 2) if pd.notna(qty) else pd.NA,
            })
    if not recs:
        return None
    return pd.DataFrame(recs)


def load_sales(file):
    """Lee un archivo de ventas reales. Detecta automáticamente:
    (a) formato pivote R12 (doble encabezado por año), o
    (b) formato simple (una fila por venta, columnas nombradas).
    Normaliza a: product, code, price, date, customer, qty.
    """
    name = file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file)
    elif name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file, header=None)
        # ¿Es el formato pivote R12? La fila 0 suele traer 'R12' o años en col 8+
        first = str(df.iloc[0, 0]).strip().upper() if len(df) else ""
        has_years = any(re.match(r"^20\d\d", str(v).split(".")[0])
                        for v in (df.iloc[0, 8:] if df.shape[1] > 8 else []))
        if first == "R12" or has_years:
            out = _read_sales_r12(df)
            if out is not None and len(out):
                return out
        # Si no era pivote, releer con encabezado normal
        file.seek(0)
        df = pd.read_excel(file)
    else:
        raise ValueError("Formato no soportado (usa CSV o Excel).")

    # --- Formato simple ---
    df.columns = [str(c).strip().lower() for c in df.columns]

    def find(cands):
        for cand in cands:
            for col in df.columns:
                if cand == col or cand in col:
                    return col
        return None

    col_prod = find(["product", "producto", "item", "descripcion", "description", "material"])
    col_code = find(["item code", "code", "codigo", "código", "sku", "enr", "ref"])
    col_price = find(["price", "precio", "sale price", "venta", "unit price",
                      "precio unitario", "importe", "amount", "valor"])
    col_date = find(["date", "fecha", "year", "año", "invoice date"])
    col_cust = find(["customer", "cliente", "client", "account", "cuenta"])
    col_qty = find(["qty", "quantity", "cantidad", "volumen", "volume", "units"])

    if col_price is None or (col_prod is None and col_code is None):
        raise ValueError("El archivo de ventas necesita al menos una columna de precio y una de producto o código.")

    out = pd.DataFrame()
    out["product"] = df[col_prod].fillna("").astype(str).str.strip() if col_prod else ""
    out["code"] = df[col_code].apply(normalize_code) if col_code else ""
    out["price"] = pd.to_numeric(df[col_price], errors="coerce")
    out["date"] = df[col_date].fillna("").astype(str) if col_date else ""
    out["customer"] = df[col_cust].fillna("").astype(str).str.strip() if col_cust else ""
    out["qty"] = pd.to_numeric(df[col_qty], errors="coerce") if col_qty else pd.NA
    out = out[out["price"].notna()]
    return out


def read_local(path):
    """Lee el histórico desde una ruta local (CSV o Excel). Devuelve DataFrame normalizado."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(path)
    else:
        raise ValueError("Formato no soportado (usa .csv o .xlsx).")
    return normalize_history(df)


def cart_to_rows(cart):
    """Convierte las cotizaciones de la sesión al esquema del histórico."""
    rows = []
    for it in cart:
        rows.append({
            "date": it["date"], "customer": it["customer"], "broker": "",
            "product": it["product"], "code": it["code"],
            "matl": it["matl"], "ovrhd": it["ovrhd"], "gm": it["gm"],
            "comm": it["comm"], "duty": it["duty"], "forex": it["forex"],
            "uom": it["uom"], "quote": it["price_usd"], "comment": it["comment"],
        })
    return normalize_history(pd.DataFrame(rows)) if rows else None


def dedup_key(df):
    """Clave de deduplicación por fecha+producto+cliente+precio (redondeado)."""
    return (df["date"].astype(str) + "|" + df["product"].str.lower().str.strip()
            + "|" + df["customer"].str.lower().str.strip()
            + "|" + df["quote"].round(2).astype(str))


def append_dedup(history, new_rows):
    """Une histórico + filas nuevas, elimina duplicados exactos. Devuelve (df, n_agregadas)."""
    if new_rows is None or not len(new_rows):
        return history, 0
    if history is None or not len(history):
        combined = new_rows.copy()
        return combined, len(new_rows)
    existing_keys = set(dedup_key(history))
    mask_new = ~dedup_key(new_rows).isin(existing_keys)
    to_add = new_rows[mask_new]
    combined = pd.concat([history, to_add], ignore_index=True)
    return combined, int(mask_new.sum())


def write_local(path, df):
    """Reescribe el archivo local con respaldo previo. Devuelve la ruta del respaldo."""
    backup = None
    if os.path.exists(path):
        base, ext = os.path.splitext(path)
        backup = f"{base}.backup_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        shutil.copy2(path, backup)
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df.to_csv(path, index=False)
    else:
        df.to_excel(path, index=False, engine="openpyxl")
    return backup


def fmt(v, d=2):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    try:
        return f"{float(v):,.{d}f}"
    except (ValueError, TypeError):
        return "—"


def pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{float(v) * 100:.1f}".rstrip("0").rstrip(".") + "%"


# ----------------------------------------------------------------------------
# Estado de sesión
# ----------------------------------------------------------------------------
if "lang" not in st.session_state:
    st.session_state.lang = "es"
if "history" not in st.session_state:
    st.session_state.history = None          # DataFrame o None
if "cart" not in st.session_state:
    st.session_state.cart = []               # lista de dicts (cotizaciones de la sesión)
if "base_row" not in st.session_state:
    st.session_state.base_row = {}           # fila cargada como base en el formulario
if "active_product" not in st.session_state:
    st.session_state.active_product = ""     # producto en foco (para analytics contextual)
if "sales" not in st.session_state:
    st.session_state.sales = None            # DataFrame de ventas reales (opcional)
if "local_path" not in st.session_state:
    st.session_state.local_path = ""         # ruta al archivo en OneDrive (modo local)

# ----------------------------------------------------------------------------
# Sidebar: idioma + carga efímera + privacidad
# ----------------------------------------------------------------------------
with st.sidebar:
    st.session_state.lang = st.radio(
        T[st.session_state.lang]["lang_label"], ["es", "en"],
        format_func=lambda x: {"es": "🇪🇸 Español", "en": "🇬🇧 English"}[x],
        horizontal=True,
        index=0 if st.session_state.lang == "es" else 1,
    )
    t = T[st.session_state.lang]

    st.markdown("---")
    st.subheader(t["load_header"])

    mode = st.radio(
        t["mode_label"], ["local", "upload"],
        format_func=lambda x: t["mode_local"] if x == "local" else t["mode_upload"],
        index=0,
    )

    if mode == "local":
        # --- Modo OneDrive local: leer y reescribir un archivo en la carpeta sincronizada ---
        path_in = st.text_input(
            t["local_path"], value=st.session_state.local_path,
            help=t["local_path_help"], placeholder="C:/Users/…/OneDrive - Robertet/…/historico.xlsx",
        )
        if st.button(t["local_load"], use_container_width=True):
            p = path_in.strip().strip('"')
            if not os.path.exists(p):
                st.error(t["local_missing"])
            else:
                try:
                    st.session_state.history = read_local(p)
                    st.session_state.local_path = p
                    st.success(t["local_ok"].format(n=len(st.session_state.history)))
                except Exception as exc:  # noqa: BLE001
                    st.error(t["load_error"].format(err=exc))
    else:
        # --- Modo subir (efímero) ---
        uploaded = st.file_uploader(
            t["load_prompt"], type=["csv", "xlsx", "xls"], help=t["load_help"],
            label_visibility="collapsed",
        )
        if uploaded is not None:
            try:
                st.session_state.history, diag = load_uploaded(uploaded, return_diagnostics=True)
                st.session_state.local_path = ""  # subir no permite reescritura
                detected, miss_ess, miss_imp = diag
                if detected:
                    st.caption(t["detect_ok"].format(cols=", ".join(detected)))
                if miss_ess:
                    st.warning(t["detect_missing_essential"].format(cols=", ".join(miss_ess)))
                elif miss_imp:
                    st.caption(t["detect_missing_important"].format(cols=", ".join(miss_imp)))
                st.caption(t["template_hint"])
            except Exception as exc:  # noqa: BLE001
                st.error(t["load_error"].format(err=exc))

    hist = st.session_state.history
    if hist is not None and len(hist):
        dmin = hist["date"].replace("", pd.NA).dropna().min()
        dmax = hist["date"].replace("", pd.NA).dropna().max()
        st.success(t["loaded_ok"].format(n=len(hist), start=dmin, end=dmax))
        if st.button(t["clear_data"]):
            st.session_state.history = None
            st.rerun()
    else:
        st.info(t["load_none"])

    # --- Guardar en OneDrive (solo modo local con ruta cargada) ---
    if st.session_state.local_path:
        st.markdown("---")
        st.subheader(t["local_save_header"])
        st.caption(t["local_tip"])
        if st.button(t["local_save_btn"], type="primary", use_container_width=True):
            if not st.session_state.cart:
                st.warning(t["local_nothing"])
            else:
                new_rows = cart_to_rows(st.session_state.cart)
                combined, n_added = append_dedup(st.session_state.history, new_rows)
                n_dups = len(new_rows) - n_added
                try:
                    backup = write_local(st.session_state.local_path, combined)
                    st.session_state.history = combined
                    st.session_state.cart = []  # sesión guardada, se vacía el carrito
                    if n_dups:
                        st.success(t["local_saved_dups"].format(n=n_added, d=n_dups, total=len(combined)))
                    else:
                        st.success(t["local_saved"].format(n=n_added, total=len(combined)))
                    if backup:
                        st.caption(t["local_backup"].format(b=os.path.basename(backup)))
                except Exception as exc:  # noqa: BLE001
                    st.error(t["local_write_err"].format(err=exc))

    st.caption(t["session_privacy"])

    # --- Archivo de ventas reales (opcional) ---
    st.markdown("---")
    st.subheader(t["sales_header"])
    sales_up = st.file_uploader(t["sales_prompt"], type=["csv", "xlsx", "xls"],
                                help=t["sales_help"], label_visibility="collapsed",
                                key="sales_uploader")
    if sales_up is not None:
        try:
            st.session_state.sales = load_sales(sales_up)
            st.success(t["sales_ok"].format(n=len(st.session_state.sales)))
        except Exception as exc:  # noqa: BLE001
            st.error(t["sales_error"].format(err=exc))
    if st.session_state.sales is not None and len(st.session_state.sales):
        if st.button(t["sales_clear"]):
            st.session_state.sales = None
            st.rerun()

t = T[st.session_state.lang]
hist = st.session_state.history

# ----------------------------------------------------------------------------
# Header
# ----------------------------------------------------------------------------
st.markdown('<div class="spice-strip"></div>', unsafe_allow_html=True)
c1, c2 = st.columns([3, 1])
with c1:
    st.markdown(f'<div class="brand-title">{t["brand"]}</div>', unsafe_allow_html=True)
    n = len(hist) if hist is not None else 0
    st.markdown(
        f'<div class="brand-sub">{t["sub"]}'
        + (f' · {n:,} ' + ("registros" if st.session_state.lang == "es" else "records") if n else "")
        + "</div>",
        unsafe_allow_html=True,
    )

tab_search, tab_quote, tab_batch, tab_insights = st.tabs(
    [t["tab_search"], t["tab_quote"], t["tab_batch"], t["tab_insights"]]
)

# ============================================================================
# TAB 1 — Buscar historial
# ============================================================================
with tab_search:
    if hist is None or not len(hist):
        st.info(t["load_none"])
    else:
        q = st.text_input(t["search_ph"], key="search_q",
                          placeholder=t["search_ph"], label_visibility="collapsed")
        st.caption(t["search_hint"])
        if q.strip():
            terms = q.lower().split()
            hay = (hist["product"].str.lower() + " " + hist["code"].str.lower()
                   + " " + hist["customer"].str.lower() + " " + hist["broker"].str.lower())
            mask = pd.Series(True, index=hist.index)
            for term in terms:
                mask &= hay.str.contains(term, regex=False, na=False)
            res = hist[mask].sort_values("date", ascending=False).head(200)
            if not len(res):
                st.warning(t["no_results"])
            else:
                st.caption(t["results_count"].format(n=len(res)))
                show = res[["date", "customer", "product", "code", "matl",
                            "gm", "quote", "uom", "comment"]].copy()
                show.columns = [t["col_date"], t["col_customer"], t["col_product"],
                                t["col_code"], t["col_cost"], t["col_gm"],
                                t["col_quote"], t["col_uom"], t["col_note"]]
                show[t["col_gm"]] = res["gm"].apply(pct).values
                st.dataframe(show, use_container_width=True, hide_index=True,
                             height=min(560, 60 + 35 * len(res)))

                # --- Precargar una fila en «Nueva cotización» con un clic ---
                res_reset = res.reset_index(drop=True)
                labels = [t["row_label"].format(date=r["date"], cust=(r["customer"] or "—"),
                                                prod=r["product"], q=fmt(r["quote"]))
                          for _, r in res_reset.iterrows()]
                pc1, pc2 = st.columns([3, 1])
                with pc1:
                    sel_idx = st.selectbox(t["sel_to_preload"], range(len(labels)),
                                           format_func=lambda i: labels[i], key="preload_sel")
                with pc2:
                    st.write("")
                    st.write("")
                    if st.button(t["use_in_quote"], type="primary", use_container_width=True):
                        row = res_reset.iloc[sel_idx]
                        st.session_state.base_row = {
                            "customer": row["customer"], "product": row["product"],
                            "code": row["code"], "uom": row["uom"] or "kg",
                            "matl": float(row["matl"]) if pd.notna(row["matl"]) else 0.0,
                            "ovrhd": float(row["ovrhd"]) if pd.notna(row["ovrhd"]) else 2.0,
                            "gm": float(row["gm"]) if pd.notna(row["gm"]) else 0.30,
                            "comm": float(row["comm"]) if pd.notna(row["comm"]) else 0.0,
                            "duty": float(row["duty"]) if pd.notna(row["duty"]) else 0.0,
                            "forex": float(row["forex"]) if pd.notna(row["forex"]) else 1.37,
                        }
                        st.session_state.active_product = row["product"]
                        # Fijar directamente los selectbox del formulario (así los encabezados se llenan)
                        # Se limpian las keys de widget para que el formulario las regenere desde base_row.
                        for k in ("cust_pick", "prod_pick", "code_pick",
                                  "cust_new", "prod_new", "code_new", "final_price_input"):
                            st.session_state.pop(k, None)
                        st.session_state._preload = True
                        st.success(t["preloaded_ok"].format(p=row["product"]))
                        st.rerun()

# ============================================================================
# TAB 2 — Nueva cotización
# ============================================================================
with tab_quote:
    b = st.session_state.base_row
    left, right = st.columns([1.5, 1])

    with left:
        st.subheader(t["form_header"])

        # ============================================================
        # Mapas de relación desde el histórico (case-insensitive)
        # ============================================================
        prod_options, code_options, cust_options = [], [], []
        prod_to_codes = {}       # producto -> [códigos]
        code_to_prods = {}       # código   -> [productos]
        code_to_custs = {}       # código   -> [clientes]
        cust_to_prods = {}       # cliente  -> [productos]
        prod_canonical = {}      # lower(nombre) -> nombre canónico

        if hist is not None and len(hist):
            h2 = hist[hist["product"].astype(str).str.strip() != ""].copy()
            for _, row in h2.sort_values("date").iterrows():
                prod_canonical[str(row["product"]).strip().lower()] = str(row["product"]).strip()
            prod_options = sorted(prod_canonical.values(), key=lambda x: x.lower())
            cust_options = sorted({str(c).strip() for c in hist["customer"] if str(c).strip()},
                                  key=lambda x: x.lower())
            for _, row in h2.iterrows():
                canon = prod_canonical.get(str(row["product"]).strip().lower(), str(row["product"]).strip())
                cd = str(row["code"]).strip()
                cu = str(row["customer"]).strip()
                if cd:
                    prod_to_codes.setdefault(canon, [])
                    if cd not in prod_to_codes[canon]:
                        prod_to_codes[canon].append(cd)
                    code_to_prods.setdefault(cd, [])
                    if canon not in code_to_prods[cd]:
                        code_to_prods[cd].append(canon)
                    if cu:
                        code_to_custs.setdefault(cd, [])
                        if cu not in code_to_custs[cd]:
                            code_to_custs[cd].append(cu)
                if cu:
                    cust_to_prods.setdefault(cu, [])
                    if canon not in cust_to_prods[cu]:
                        cust_to_prods[cu].append(canon)
            code_options = sorted(code_to_prods.keys(), key=lambda x: x.lower())

        NEW = t["write_new"]
        GRP_CAT = t["grp_catalog"]
        b = st.session_state.base_row

        def sep(label):
            """Devuelve un separador visual no seleccionable."""
            return "── " + label + " ──"

        # ============================================================
        # 1) CÓDIGO primero (permite el flujo código -> producto/cliente)
        #    Lo leemos ANTES para poder filtrar cliente y producto por él.
        # ============================================================
        # El código elegido en la corrida previa (si lo hay)
        chosen_code = st.session_state.get("code_pick", None)
        if chosen_code in (NEW, None) or (chosen_code and chosen_code.startswith("──")):
            chosen_code = None

        # ============================================================
        # 2) CLIENTE: híbrido, filtrado por código si hay uno elegido
        # ============================================================
        r1c1, r1c2 = st.columns(2)
        with r1c1:
            if cust_options:
                if chosen_code and chosen_code in code_to_custs:
                    # solo clientes que han cotizado ese código, + resto
                    linked = sorted(code_to_custs[chosen_code], key=lambda x: x.lower())
                    rest = [c for c in cust_options if c not in linked]
                    copts = [NEW] + linked + ([sep(GRP_CAT)] + rest if rest else [])
                else:
                    copts = [NEW] + cust_options
                # valor inicial desde base_row (precarga)
                base_cust = b.get("customer", "")
                if "cust_pick" not in st.session_state and base_cust in cust_options:
                    st.session_state.cust_pick = base_cust
                elif "cust_pick" in st.session_state and st.session_state.cust_pick not in copts:
                    st.session_state.cust_pick = NEW
                cpick = st.selectbox(t["customer_pick"], copts, key="cust_pick")
                if cpick == NEW:
                    customer = st.text_input(t["customer_new"], key="cust_new",
                                             value=base_cust if base_cust and base_cust not in cust_options else "")
                elif cpick.startswith("──"):
                    customer = ""
                else:
                    customer = cpick
            else:
                customer = st.text_input(t["customer"], value=b.get("customer", ""))

        # ============================================================
        # 3) PRODUCTO: híbrido. Prioridad de filtrado:
        #    a) si hay código elegido -> productos de ese código
        #    b) elif hay cliente -> productos del cliente + resto
        #    c) else -> catálogo completo
        # ============================================================
        with r1c2:
            if prod_options:
                if chosen_code and chosen_code in code_to_prods:
                    linked_p = sorted(code_to_prods[chosen_code], key=lambda x: x.lower())
                    rest_p = [p for p in prod_options if p not in linked_p]
                    popts = [NEW] + linked_p + ([sep(GRP_CAT)] + rest_p if rest_p else [])
                elif customer and customer.strip() and customer in cust_to_prods:
                    cprods = sorted(cust_to_prods[customer], key=lambda x: x.lower())
                    rest_p = [p for p in prod_options if p not in cprods]
                    popts = [NEW] + [sep(t["grp_client"].format(c=customer.strip()[:20]))] + cprods \
                            + ([sep(GRP_CAT)] + rest_p if rest_p else [])
                else:
                    popts = [NEW] + prod_options

                base_prod = b.get("product", "")
                if "prod_pick" not in st.session_state and base_prod in prod_options:
                    st.session_state.prod_pick = base_prod
                elif "prod_pick" in st.session_state and st.session_state.prod_pick not in popts:
                    st.session_state.prod_pick = NEW
                pick = st.selectbox(t["product_pick"], popts, key="prod_pick", help=t["autofill_hint"])
                if pick == NEW:
                    product = st.text_input(t["product_new"], key="prod_new",
                                            value=base_prod if base_prod and base_prod not in prod_options else "")
                elif pick.startswith("──"):
                    product = ""
                else:
                    product = pick
            else:
                product = st.text_input(t["product"], value=b.get("product", ""))

        # ============================================================
        # Autocompletar base (prioriza última cotización del cliente)
        # ============================================================
        auto = {}
        client_last = None
        if hist is not None and len(hist) and product:
            fam = hist[hist["product"].str.strip().str.lower() == product.strip().lower()]
            if customer and customer.strip():
                fam_client = fam[fam["customer"].str.strip().str.lower() == customer.strip().lower()]
                if len(fam_client):
                    client_last = fam_client.sort_values("date", ascending=False).iloc[0]
            source_row = client_last if client_last is not None else (
                fam.sort_values("date", ascending=False).iloc[0] if len(fam) else None)
            if source_row is not None:
                last = source_row
                auto = {
                    "code": str(last["code"]) if str(last["code"]).strip() else "",
                    "matl": float(last["matl"]) if pd.notna(last["matl"]) else 0.0,
                    "ovrhd": float(last["ovrhd"]) if pd.notna(last["ovrhd"]) else 2.0,
                    "gm": float(last["gm"]) if pd.notna(last["gm"]) else 0.30,
                    "comm": float(last["comm"]) if pd.notna(last["comm"]) else 0.0,
                    "duty": float(last["duty"]) if pd.notna(last["duty"]) else 0.0,
                    "forex": float(last["forex"]) if pd.notna(last["forex"]) else 1.37,
                    "uom": str(last["uom"]) if str(last["uom"]).strip() else "kg",
                }
        if product and customer and customer.strip():
            if client_last is not None:
                st.caption(t["client_last_price"].format(
                    c=customer.strip()[:20], p=fmt(client_last["quote"]), date=client_last["date"]))
            elif auto:
                st.caption(t["client_never"].format(c=customer.strip()[:20], p=product[:24]))

        # ============================================================
        # 4) CÓDIGO: híbrido, ligado al producto (si hay), + resto
        # ============================================================
        r2c1, r2c2 = st.columns(2)
        with r2c1:
            # familia de códigos del producto elegido
            family_codes = None
            if product:
                key_match = next((p for p in prod_to_codes if p.lower() == product.strip().lower()), None)
                if key_match:
                    family_codes = prod_to_codes[key_match]

            base_code = auto.get("code", b.get("code", ""))
            if code_options:
                if family_codes:
                    rest_c = [c for c in code_options if c not in family_codes]
                    ccopts = [NEW] + family_codes + ([sep(GRP_CAT)] + rest_c if rest_c else [])
                else:
                    ccopts = [NEW] + code_options
                # valor inicial: base_code si aplica; si el producto tiene un solo código, ese
                if "code_pick" not in st.session_state:
                    if base_code and base_code in (family_codes or code_options):
                        st.session_state.code_pick = base_code
                    elif family_codes and len(family_codes) == 1:
                        st.session_state.code_pick = family_codes[0]
                elif st.session_state.code_pick not in ccopts:
                    # el código elegido ya no aplica al nuevo producto -> resetear
                    st.session_state.code_pick = NEW
                cpick_c = st.selectbox(t["code_pick"], ccopts, key="code_pick")
                if cpick_c == NEW:
                    code = st.text_input(t["code_new"], key="code_new",
                                         value=base_code if base_code and base_code not in code_options else "")
                elif cpick_c.startswith("──"):
                    code = ""
                else:
                    code = cpick_c
            else:
                code = st.text_input(t["code"], value=base_code)

        code = normalize_code(code)
        if code and family_codes and product:
            st.caption(t["code_family"].format(n=len(family_codes), p=product[:24]))

        uom = r2c2.text_input(t["uom"], value=auto.get("uom", b.get("uom", "kg")))

        # Limpiar base_row tras consumirla (una sola vez), para no re-forzar valores
        if st.session_state.get("_preload"):
            st.session_state._preload = False

        # Producto en foco para analítica contextual
        if product and product.strip():
            st.session_state.active_product = product.strip()

        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        matl = m1.number_input(t["cost"], min_value=0.0, value=float(auto.get("matl", b.get("matl", 0.0))), step=0.01, format="%.2f")
        ovrhd = m2.number_input(t["ovrhd"], min_value=0.0, value=float(auto.get("ovrhd", b.get("ovrhd", 2.0))), step=0.5, format="%.2f")
        gm = m3.number_input(t["gm"], min_value=0.0, max_value=0.99, value=float(auto.get("gm", b.get("gm", 0.30))), step=0.05, format="%.2f")
        m4, m5, m6 = st.columns(3)
        comm = m4.number_input(t["comm"], min_value=0.0, max_value=0.99, value=float(auto.get("comm", b.get("comm", 0.0))), step=0.01, format="%.2f")
        duty = m5.number_input(t["duty"], min_value=0.0, max_value=0.99, value=float(auto.get("duty", b.get("duty", 0.0))), step=0.01, format="%.2f")
        forex = m6.number_input(t["forex"], min_value=0.0001, value=float(auto.get("forex", b.get("forex", 1.37))), step=0.001, format="%.3f")

        st.markdown("---")
        s1, s2, s3, s4 = st.columns(4)
        spec = s1.text_input(t["spec"], value="")
        incoterm = s2.text_input(t["incoterm"], value="FCA")
        moq = s3.text_input(t["moq"], value="")
        validity = s4.text_input(t["validity"], value="30 " + ("días" if st.session_state.lang == "es" else "days"))
        comment = st.text_input(t["comment"], value="")

    with right:
        price = calc_quote(matl, ovrhd, gm, comm, duty, forex)
        st.markdown(f'<div style="color:#93A3B8;font-size:11px;letter-spacing:.16em;font-weight:700">{t["calc_price"]}</div>', unsafe_allow_html=True)
        if pd.isna(price):
            st.markdown('<div class="big-price">—</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="big-price">${fmt(price)}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="brand-sub">{t["per_unit"].format(u=uom or "kg")}</div>', unsafe_allow_html=True)

        # Cascada de la fórmula
        if not pd.isna(price) and matl > 0:
            v = matl + ovrhd
            steps = [("Base (cost+ovrhd)", v)]
            if 0 < gm < 1: v /= (1 - gm); steps.append((f"÷(1−GM {pct(gm)})", v))
            if 0 < comm < 1: v /= (1 - comm); steps.append((f"÷(1−Comm {pct(comm)})", v))
            if 0 < duty < 1: v /= (1 - duty); steps.append((f"÷(1−Duty {pct(duty)})", v))
            if forex != 1: v /= forex; steps.append((f"÷ FX {fmt(forex,4)}", v))
            st.markdown("###### ⛓️ " + ("Desglose" if st.session_state.lang == "es" else "Breakdown"))
            for label, val in steps:
                st.markdown(f"<div style='display:flex;justify-content:space-between;font-size:12.5px;padding:2px 0;'><span style='color:#93A3B8'>{label}</span><span style='color:#EDF1F7;font-family:monospace'>{fmt(val)}</span></div>", unsafe_allow_html=True)

        # --- PRECIO FINAL A COTIZAR (decisión del vendedor) ---
        st.markdown("---")
        st.markdown(f'<div style="color:{GOLD_HI};font-size:11px;letter-spacing:.14em;font-weight:700">{t["final_price"]}</div>', unsafe_allow_html=True)
        suggested = 0.0 if pd.isna(price) else round(float(price), 2)
        final_price = st.number_input(
            t["final_price_label"], min_value=0.0,
            value=suggested, step=0.01, format="%.2f",
            help=t["final_price_help"], key="final_price_input",
        )
        if suggested > 0:
            st.caption(t["suggested"].format(s=fmt(suggested)))
            if abs(final_price - suggested) > 0.005:
                arrow = "↑" if final_price > suggested else "↓"
                diff_pct = abs(final_price - suggested) / suggested * 100 if suggested else 0
                st.markdown(
                    f"<div style='color:{TURMERIC};font-size:12px'>"
                    + t["final_vs_calc"].format(arrow=arrow, d=f"{diff_pct:.1f}%")
                    + "</div>", unsafe_allow_html=True)

        # --- Referencia histórica ---
        st.markdown("---")
        st.markdown(f"###### 📚 {t['ref_header']}")
        ref = None
        if hist is not None and len(hist) and product.strip():
            ref = hist[hist["product"].str.lower() == product.strip().lower()]
        if ref is not None and len(ref):
            ref_sorted = ref.sort_values("date", ascending=False)
            last_q = ref_sorted.iloc[0]["quote"]
            avg_gm = ref["gm"].mean()
            qmin, qmax = ref["quote"].min(), ref["quote"].max()
            rc1, rc2 = st.columns(2)
            rc1.metric(t["ref_last"], f"${fmt(last_q)}")
            rc2.metric(t["ref_avg_gm"], pct(avg_gm))
            st.caption(f"{t['ref_range']}: ${fmt(qmin)} – ${fmt(qmax)} · " + t["ref_count"].format(n=len(ref)))
        elif product.strip():
            st.caption(t["ref_none"])

        # --- Validaciones / puntos de atención ---
        warns = []
        if 0 < gm < 0.10:
            warns.append(t["warn_gm_low"].format(v=pct(gm)))
        if gm > 0.60:
            warns.append(t["warn_gm_high"].format(v=pct(gm)))
        if forex <= 0.01:
            warns.append(t["warn_fx"].format(v=fmt(forex, 3)))
        if ref is not None and len(ref) and not pd.isna(price):
            last_q = ref.sort_values("date", ascending=False).iloc[0]["quote"]
            if last_q and abs(price - last_q) / last_q > 0.10:
                arrow = "↑" if price > last_q else "↓"
                dv = f"{arrow} {abs(price-last_q)/last_q*100:.0f}%"
                warns.append(t["warn_dev"].format(d=dv, h=fmt(last_q)))
            hist_uom = ref.sort_values("date", ascending=False).iloc[0]["uom"]
            if hist_uom and uom and hist_uom.lower() != uom.lower():
                warns.append(t["warn_unit"].format(u=uom, h=hist_uom))
        if warns:
            st.markdown(f"###### ⚠️ {t['warn_header']}")
            for w in warns:
                st.markdown(f"<div style='color:{TURMERIC};font-size:12.5px;padding:2px 0'>• {w}</div>", unsafe_allow_html=True)

        st.markdown("---")
        if st.button(t["add_quote"], type="primary", use_container_width=True):
            # El precio a guardar es el FINAL (decisión del vendedor), no el calculado.
            price_to_save = final_price if final_price and final_price > 0 else (0.0 if pd.isna(price) else price)
            if not product.strip() or price_to_save <= 0:
                st.error(t["missing"])
            else:
                st.session_state.cart.append({
                    "date": date.today().isoformat(),
                    "customer": customer.strip(), "product": product.strip(),
                    "code": normalize_code(code), "spec": spec.strip(),
                    "matl": round(matl, 4), "ovrhd": round(ovrhd, 4),
                    "gm": round(gm, 4), "comm": round(comm, 4), "duty": round(duty, 4),
                    "forex": round(forex, 4), "uom": uom or "kg",
                    "price_usd": round(price_to_save, 2),
                    "price_suggested": (None if pd.isna(price) else round(float(price), 2)),
                    "incoterm": incoterm.strip(), "moq": moq.strip(),
                    "validity": validity.strip(), "comment": comment.strip(),
                })
                st.session_state.base_row = {}
                st.success(t["added_ok"].format(p=product.strip(), q=fmt(price_to_save)))

# ============================================================================
# TAB 3 — Cotización al cliente (export en lote)
# ============================================================================
with tab_batch:
    st.subheader(t["batch_header"])
    cart = st.session_state.cart
    if not cart:
        st.info(t["batch_empty"])
    else:
        st.caption(t["batch_count"].format(n=len(cart)))
        cart_df = pd.DataFrame(cart)

        # Vista interna (con costos/márgenes)
        internal_view = cart_df[["date", "customer", "product", "code", "spec",
                                 "matl", "ovrhd", "gm", "comm", "duty", "forex",
                                 "price_usd", "uom", "incoterm", "moq", "validity", "comment"]].copy()
        st.dataframe(internal_view, use_container_width=True, hide_index=True)

        # Quitar ítems
        with st.expander("🗑️ " + t["remove"]):
            for i, item in enumerate(cart):
                cols = st.columns([5, 1])
                cols[0].write(f"{item['product']} · {item['customer']} · ${fmt(item['price_usd'])}")
                if cols[1].button(t["remove"], key=f"rm_{i}"):
                    st.session_state.cart.pop(i)
                    st.rerun()

        st.markdown("---")
        st.markdown(f"##### {t['client_meta']}")
        mc1, mc2, mc3, mc4 = st.columns(4)
        cl_name = mc1.text_input(t["client_name"], value=cart[0].get("customer", ""))
        cl_ref = mc2.text_input(t["client_ref"], value="")
        cl_curr = mc3.text_input(t["client_currency"], value="USD")
        cl_by = mc4.text_input(t["client_prepared"], value="Robertet")

        # Vista cliente (SIN costos ni márgenes — regla de la skill)
        client_view = cart_df[["product", "spec", "price_usd", "uom",
                               "incoterm", "moq", "validity"]].copy()
        client_view.columns = ["Product/Producto", "Spec", f"Price ({cl_curr})",
                               "UOM", "Incoterm", "MOQ", "Validity/Validez"]

        st.markdown("---")
        e1, e2 = st.columns(2)

        # Export interno (Excel con costos)
        buf_int = io.BytesIO()
        with pd.ExcelWriter(buf_int, engine="openpyxl") as xw:
            internal_view.to_excel(xw, sheet_name="Cotizacion_Interna", index=False)
        e1.download_button(
            t["export_internal"], data=buf_int.getvalue(),
            file_name=f"cotizacion_interna_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Export cliente (Excel limpio con encabezado)
        buf_cli = io.BytesIO()
        with pd.ExcelWriter(buf_cli, engine="openpyxl") as xw:
            header = pd.DataFrame({
                "": [f"{'Cotización' if st.session_state.lang=='es' else 'Quotation'}: {cl_name}",
                     f"{'Referencia' if st.session_state.lang=='es' else 'Reference'}: {cl_ref}",
                     f"{'Fecha' if st.session_state.lang=='es' else 'Date'}: {date.today().isoformat()}",
                     f"{'Preparado por' if st.session_state.lang=='es' else 'Prepared by'}: {cl_by}",
                     ""]
            })
            header.to_excel(xw, sheet_name="Quotation", index=False, header=False)
            client_view.to_excel(xw, sheet_name="Quotation", index=False, startrow=6)
        e2.download_button(
            t["export_client"], data=buf_cli.getvalue(),
            file_name=f"cotizacion_cliente_{cl_name or 'cliente'}_{datetime.now():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help=t["export_client_help"], use_container_width=True,
        )
        st.caption("🔒 " + t["export_client_help"])

        # ------------------------------------------------------------------
        # Actualizar histórico (modo nube): descargar CSV/Excel completo
        # con las cotizaciones de la sesión ya agregadas y deduplicadas.
        # ------------------------------------------------------------------
        st.markdown("---")
        st.markdown(f"##### 🔄 {t['update_header']}")
        st.caption(t["update_desc"])

        if st.session_state.history is None or not len(st.session_state.history):
            st.info(t["update_no_hist"])
        else:
            new_rows = cart_to_rows(st.session_state.cart)
            combined, n_added = append_dedup(st.session_state.history, new_rows)
            n_base = len(st.session_state.history)
            n_dups = (len(new_rows) - n_added) if new_rows is not None else 0

            summary = t["update_summary"].format(base=n_base, new=n_added, total=len(combined))
            if n_dups:
                summary += " " + t["update_dups_note"].format(d=n_dups)
            st.markdown(f"<div style='color:{GOLD_HI};font-size:13px;margin-bottom:8px'>📊 {summary}</div>",
                        unsafe_allow_html=True)
            if not st.session_state.cart:
                st.caption(t["update_no_cart"])

            u1, u2 = st.columns(2)
            stamp = datetime.now().strftime("%Y%m%d_%H%M")

            # CSV
            csv_bytes = combined.to_csv(index=False).encode("utf-8-sig")
            u1.download_button(
                t["update_btn_csv"], data=csv_bytes,
                file_name=f"historico_actualizado_{stamp}.csv",
                mime="text/csv", use_container_width=True,
            )
            # Excel
            buf_hist = io.BytesIO()
            with pd.ExcelWriter(buf_hist, engine="openpyxl") as xw:
                combined.to_excel(xw, sheet_name="Historico", index=False)
            u2.download_button(
                t["update_btn_xlsx"], data=buf_hist.getvalue(),
                file_name=f"historico_actualizado_{stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
# ============================================================================
with tab_insights:
    sales = st.session_state.sales
    active = st.session_state.get("active_product", "").strip()

    if hist is None or not len(hist):
        st.subheader(t["insights_header"])
        st.info(t["insights_none"])
    elif active:
        # ---------- ANALÍTICA CONTEXTUAL DEL PRODUCTO ----------
        st.subheader(t["ins_product_header"].format(p=active))

        use_sales = sales is not None and len(sales) > 0
        # Filtrar los datos de ese producto en la fuente elegida
        if use_sales:
            # match por producto (case-insensitive) O por cualquier código de esa familia
            prod_codes = set(hist[hist["product"].str.strip().str.lower() == active.lower()]["code"]) - {""}
            sp = sales[
                (sales["product"].str.strip().str.lower() == active.lower())
                | (sales["code"].isin(prod_codes))
            ]
            data = sp.rename(columns={"price": "value"})
            source_label = t["ins_source_sales"]
        else:
            hp = hist[hist["product"].str.strip().str.lower() == active.lower()].copy()
            data = hp.rename(columns={"quote": "value"})
            source_label = t["ins_source_quotes"]

        st.caption(source_label)

        if not len(data) or data["value"].dropna().empty:
            st.info(t["ins_no_prod_data"].format(p=active))
        else:
            vals = data["value"].dropna()
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric(t["ins_prod_min"], f"${fmt(vals.min())}")
            m2.metric(t["ins_prod_avg"], f"${fmt(vals.mean())}")
            m3.metric(t["ins_prod_max"], f"${fmt(vals.max())}")
            data_dated = data[data["date"].astype(str).str.strip() != ""].copy()
            if len(data_dated):
                data_dated = data_dated.sort_values("date")
                m4.metric(t["ins_prod_last"], f"${fmt(data_dated.iloc[-1]['value'])}")
            m5.metric(t["ins_prod_count"], f"{len(vals)}")

            # Precio en el tiempo
            if len(data_dated):
                st.markdown(f"###### {t['ins_price_time']}")
                ts = data_dated.copy()
                ts["_d"] = pd.to_datetime(ts["date"], errors="coerce")
                ts = ts.dropna(subset=["_d"]).set_index("_d")["value"]
                if len(ts):
                    st.line_chart(ts)

            # Por cliente (precio promedio para este producto)
            if "customer" in data.columns and (data["customer"].astype(str).str.strip() != "").any():
                st.markdown(f"###### {t['ins_by_client_prod']}")
                by_c = (data[data["customer"].astype(str).str.strip() != ""]
                        .groupby("customer")["value"].mean().sort_values(ascending=False).head(15))
                st.bar_chart(by_c)

        st.caption(t["ins_no_product"])
    else:
        # ---------- VISTA GENERAL (sin producto activo) ----------
        st.subheader(t["insights_header"])
        st.caption(t["ins_no_product"])
        ic1, ic2 = st.columns(2)
        with ic1:
            st.markdown(f"###### {t['ins_by_customer']}")
            by_cust = (hist[hist["customer"] != ""]
                       .groupby("customer")["gm"].mean().sort_values(ascending=False).head(15))
            st.bar_chart(by_cust)
        with ic2:
            st.markdown(f"###### {t['ins_top_products']}")
            top_prod = hist["product"].value_counts().head(15)
            st.bar_chart(top_prod)

        st.markdown(f"###### {t['ins_over_time']}")
        tmp = hist.copy()
        tmp["month"] = pd.to_datetime(tmp["date"], errors="coerce").dt.to_period("M").astype(str)
        by_month = tmp[tmp["month"] != "NaT"].groupby("month").size()
        st.line_chart(by_month)

st.markdown("---")
st.caption(t["footer"])
