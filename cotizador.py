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
        "client_meta": "Datos de la cotización al cliente",
        "client_name": "Cliente", "client_ref": "Referencia / RFQ", "client_currency": "Moneda",
        "client_prepared": "Preparado por",
        "insights_header": "Analítica del histórico",
        "insights_none": "Carga un histórico para ver la analítica.",
        "ins_by_customer": "GM promedio por cliente (top 15)",
        "ins_top_products": "Productos más cotizados (top 15)",
        "ins_over_time": "Cotizaciones por mes",
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
        "client_meta": "Client quotation details",
        "client_name": "Customer", "client_ref": "Reference / RFQ", "client_currency": "Currency",
        "client_prepared": "Prepared by",
        "insights_header": "History analytics",
        "insights_none": "Load a history file to see analytics.",
        "ins_by_customer": "Average GM by customer (top 15)",
        "ins_top_products": "Most quoted products (top 15)",
        "ins_over_time": "Quotes per month",
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


def calc_quote(matl, ovrhd, gm, comm, duty, forex):
    """USD = (matl + ovrhd) / (1-GM) / (1-comm) / (1-duty) / forex."""
    fx = forex or 1.0
    if gm >= 1 or comm >= 1 or duty >= 1:
        return float("nan")
    return (matl + ovrhd) / (1 - gm) / (1 - comm) / (1 - duty) / fx


def normalize_history(df):
    """Ajusta un dataframe cargado al esquema esperado, tolerante a columnas faltantes."""
    df = df.copy()
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
    return df


def load_uploaded(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file)
    elif name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file)
    else:
        raise ValueError("Formato no soportado (usa CSV o Excel).")
    return normalize_history(df)


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
                st.session_state.history = load_uploaded(uploaded)
                st.session_state.local_path = ""  # subir no permite reescritura
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
                st.caption("💡 " + ("Copia un producto y cárgalo en «Nueva cotización» para calcular con su base."
                                    if st.session_state.lang == "es" else
                                    "Copy a product into «New quote» to calculate from its base."))

# ============================================================================
# TAB 2 — Nueva cotización
# ============================================================================
with tab_quote:
    b = st.session_state.base_row
    left, right = st.columns([1.5, 1])

    with left:
        st.subheader(t["form_header"])
        r1c1, r1c2 = st.columns(2)
        customer = r1c1.text_input(t["customer"], value=b.get("customer", ""))
        product = r1c2.text_input(t["product"], value=b.get("product", ""))
        r2c1, r2c2 = st.columns(2)
        code = r2c1.text_input(t["code"], value=b.get("code", ""))
        uom = r2c2.text_input(t["uom"], value=b.get("uom", "kg"))

        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        matl = m1.number_input(t["cost"], min_value=0.0, value=float(b.get("matl", 0.0)), step=0.01, format="%.2f")
        ovrhd = m2.number_input(t["ovrhd"], min_value=0.0, value=float(b.get("ovrhd", 2.0)), step=0.5, format="%.2f")
        gm = m3.number_input(t["gm"], min_value=0.0, max_value=0.99, value=float(b.get("gm", 0.30)), step=0.05, format="%.2f")
        m4, m5, m6 = st.columns(3)
        comm = m4.number_input(t["comm"], min_value=0.0, max_value=0.99, value=float(b.get("comm", 0.0)), step=0.01, format="%.2f")
        duty = m5.number_input(t["duty"], min_value=0.0, max_value=0.99, value=float(b.get("duty", 0.0)), step=0.01, format="%.2f")
        forex = m6.number_input(t["forex"], min_value=0.0001, value=float(b.get("forex", 1.37)), step=0.001, format="%.3f")

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
            if not product.strip() or matl <= 0 or pd.isna(price):
                st.error(t["missing"])
            else:
                st.session_state.cart.append({
                    "date": date.today().isoformat(),
                    "customer": customer.strip(), "product": product.strip(),
                    "code": code.strip(), "spec": spec.strip(),
                    "matl": round(matl, 4), "ovrhd": round(ovrhd, 4),
                    "gm": round(gm, 4), "comm": round(comm, 4), "duty": round(duty, 4),
                    "forex": round(forex, 4), "uom": uom or "kg",
                    "price_usd": round(price, 2),
                    "incoterm": incoterm.strip(), "moq": moq.strip(),
                    "validity": validity.strip(), "comment": comment.strip(),
                })
                st.session_state.base_row = {}
                st.success(t["added_ok"].format(p=product.strip(), q=fmt(price)))

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

# ============================================================================
# TAB 4 — Analítica
# ============================================================================
with tab_insights:
    st.subheader(t["insights_header"])
    if hist is None or not len(hist):
        st.info(t["insights_none"])
    else:
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
