# Cotizador Internacional — Robertet Ingredients

App de cotización de ingredientes naturales (oleorresinas, especias, colores naturales,
extractos, químicos aromáticos) con lógica de margen y tipo de cambio. Interfaz en
**Español / Inglés**. Construida en Streamlit para desplegarse desde GitHub.

---

## 🔒 Modelo: local + OneDrive (recomendado)

Esta app está pensada para **correr en tu propia computadora** y guardar el histórico en
tu **carpeta de OneDrive sincronizada**. El repositorio **no contiene ningún dato comercial**.

- El repo solo tiene **código** (la app, la lógica de pricing, la interfaz).
- Tu histórico real (costos, márgenes, clientes) vive en un archivo dentro de tu carpeta
  local de OneDrive — por ejemplo `C:/Users/Tu/OneDrive - Robertet/cotizaciones/historico.xlsx`.
- Al iniciar, la app **lee** ese archivo. Al guardar una sesión, **agrega** las cotizaciones
  nuevas y **reescribe** el archivo en su lugar (dejando primero un respaldo con fecha).
- OneDrive sincroniza el cambio **solo** a la nube y a tu **celular**. Sin trámites de Azure,
  sin credenciales, sin exponer nada en internet.

Como la app corre en tu equipo, nada viaja a servidores externos: es más privado que una
URL pública. Y como el archivo vive en OneDrive, está **igual en tu compu y tu celular**,
siempre actualizado.

> ⚠️ **Nunca** subas tu archivo histórico real al repositorio. El `.gitignore` ya bloquea
> `data/*.csv` y `data/*.xlsx` (salvo la plantilla de muestra). Tu archivo real vive solo
> en tu OneDrive.

### Los dos modos de la barra lateral

- **📁 Archivo en OneDrive (local)** — pegas la ruta a tu archivo; la app lo lee y, al guardar,
  lo reescribe ahí mismo. Este es el modo principal.
- **⬆️ Subir archivo (efímero)** — por si corres la app en un equipo sin OneDrive: subes el
  archivo, trabajas, y exportas en lote. No reescribe nada (solo lectura).

---

## 📁 Estructura

```
cotizador-robertet/
├── cotizador.py                     # la app
├── requirements.txt                 # dependencias
├── README.md                        # este archivo
├── .gitignore                       # bloquea datos reales
├── .streamlit/
│   └── config.toml                  # tema visual navy/gold
└── data/
    └── plantilla_historico.csv      # SOLO estructura + 3 filas de ejemplo
```

El archivo `data/plantilla_historico.csv` muestra el **formato esperado**. Tu histórico real
debe tener las mismas columnas:

`date, customer, broker, product, code, matl, ovrhd, gm, comm, duty, forex, uom, quote, comment`

- `matl` = costo de material (CAD/kg) · `ovrhd` = overhead · `gm` = margen bruto (0.30 = 30%)
- `comm` = comisión · `duty` = arancel · `forex` = tipo de cambio CAD→USD
- `quote` = precio calculado (USD) · el resto son texto

**Fórmula:** `precio_USD = (matl + ovrhd) / (1−gm) / (1−comm) / (1−duty) / forex`

Columnas faltantes se rellenan solas; el orden no importa (se normaliza por nombre).

---

## 🚀 Puesta en marcha (local + OneDrive)

### Requisitos previos
- **Python 3.9+** instalado ([python.org](https://www.python.org/downloads/), marca "Add to PATH").
- **OneDrive** instalado y sincronizando tu carpeta de Robertet (ya lo tienes por Microsoft 365).

### Pasos (una sola vez)

1. Descarga este proyecto (o clónalo desde tu repo de GitHub) a una carpeta de tu equipo.
2. Coloca tu histórico real dentro de tu carpeta de OneDrive. Por ejemplo, crea
   `OneDrive - Robertet/cotizaciones/` y guarda ahí tu `historico.xlsx`.
3. Abre una terminal (PowerShell en Windows) en la carpeta del proyecto y ejecuta:

   ```bash
   pip install -r requirements.txt
   streamlit run cotizador.py
   ```

4. Se abre en tu navegador (`http://localhost:8501`).

### Uso diario

1. En la barra lateral, deja el modo en **📁 Archivo en OneDrive (local)**.
2. Pega la ruta completa a tu archivo. Para copiarla en Windows: clic derecho sobre el
   archivo → *Copiar como ruta*. Ejemplo:
   `C:/Users/TuNombre/OneDrive - Robertet/cotizaciones/historico.xlsx`
3. Clic en **Cargar desde OneDrive**. Ya puedes buscar, cotizar y ver analítica.
4. Agrega cotizaciones en «Nueva cotización».
5. Cuando termines, en la barra lateral clic en **💾 Agregar sesión al histórico y reescribir**.
   La app agrega tus cotizaciones nuevas, reescribe el archivo y crea un respaldo con fecha.
   OneDrive sincroniza el cambio a la nube y a tu celular automáticamente.

> 💡 **Cierra el archivo en Excel antes de guardar.** Si está abierto, Windows bloquea la
> reescritura y la app te avisará. Ciérralo y reintenta.

### Sobre los respaldos
Cada vez que guardas, la app deja una copia como
`historico.backup_AAAAMMDD_HHMMSS.xlsx` en la misma carpeta. Si algo sale mal, tienes la
versión anterior intacta. El `.gitignore` también excluye estos respaldos del repo.

---

## ☁️ ¿Y si algún día quiero una URL compartible?

Este proyecto también corre en **Streamlit Community Cloud** (sube el repo, apunta a
`cotizador.py`). Pero en la nube el disco es efímero y **no** puede reescribir tu OneDrive
directamente: ahí funcionaría solo el modo **⬆️ Subir archivo** (lectura + exportar en lote).
Para reescritura automática en la nube haría falta conectar Microsoft Graph, lo que requiere
un registro de app en Azure (normalmente vía IT). Para uso personal, **correr local con
OneDrive es más simple y más privado.**

---

## ✨ Funciones

- **ES / EN** — toda la interfaz cambia de idioma desde la barra lateral.
- **Buscar historial** — búsqueda AND por producto, código y cliente.
- **Nueva cotización** — cálculo en vivo con desglose paso a paso de la fórmula.
- **Referencia histórica automática** — al escribir un producto, muestra último precio,
  GM promedio y rango histórico de ese producto.
- **Puntos de atención** — avisa GM fuera de rango, forex sospechoso, desviación >±10%
  vs. el histórico, y diferencias de unidad (kg vs otra).
- **Cotización al cliente (export en lote)** — acumula ítems durante la sesión y exporta:
  - *versión interna* (Excel con costos y márgenes),
  - *versión cliente* (Excel limpio: solo ítem, spec, precio, incoterm, validez, MOQ —
    **sin costos ni márgenes**, por regla comercial).
- **Analítica** — GM promedio por cliente, productos más cotizados, cotizaciones por mes.

---

## 📝 Notas

- La versión para cliente **nunca** incluye costos ni márgenes internos.
- Verifica siempre las unidades (kg vs lb): el error de unidad es el fallo más costoso.
- Los supuestos (FX, validez, incoterm) quedan documentados en el export interno para que
  cualquier revisor pueda reconstruir cada precio.
