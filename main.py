import os
import time
import re
import tempfile
from playwright.sync_api import sync_playwright
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# =========================================================
# 🔹 CONFIG
# =========================================================
URL_LOGIN = "http://seaap.minsa.gob.pe/web/login"
URL_WEB   = "http://seaap.minsa.gob.pe/web"

USUARIO  = os.getenv("SEAAP_USER")
PASSWORD = os.getenv("SEAAP_PASS")

if not USUARIO or not PASSWORD:
    raise Exception("❌ Faltan credenciales")

# =========================================================
# 🔹 GOOGLE SHEETS
# =========================================================
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDS")

with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as tmp:
    tmp.write(GOOGLE_CREDS_JSON)
    CREDS_PATH = tmp.name

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds  = ServiceAccountCredentials.from_json_keyfile_name(CREDS_PATH, scope)
client = gspread.authorize(creds)

spreadsheet = client.open("DATA COMPROMISO 1 CONSOLIDADO ABRIL ")

# hojas
HOJAS_ACTORES = [
    h.title for h in spreadsheet.worksheets()
    if h.title not in [
        "telefono","Sheet1","RURAL","FIRMAS",
        "HEMOGLOBINA","VACUNAS","SEGUIMIENTO 1",
        "SEGUIMIENTO GESTORA","CONSOLIDADO"
    ]
]

sheets = {}
dni_filas = {}

print("📄 Detectando hojas...")

for nombre in HOJAS_ACTORES:
    sh = spreadsheet.worksheet(nombre)
    sheets[nombre] = sh

    dni_col = sh.col_values(3)

    dni_filas[nombre] = {
        str(dni): i + 1
        for i, dni in enumerate(dni_col)
        if dni
    }

print("🟢 Hojas:", HOJAS_ACTORES)

# =========================================================
# 🔹 ACTORES VALIDOS
# =========================================================
hoja_tel = spreadsheet.worksheet("telefono")

ACTORES_VALIDOS_DNI = {
    str(v).strip()
    for v in hoja_tel.col_values(1)[1:]
    if str(v).strip().isdigit()
}

# =========================================================
# 🔹 HELPERS
# =========================================================
def extraer_dni_actor(texto):
    m = re.match(r"^\[(\d+)\]", str(texto).strip())
    return m.group(1) if m else None

def call_kw(page, payload):
    return page.evaluate("""
        async (payload) => {
            const res = await fetch('/web/dataset/call_kw', {
                method: 'POST',
                credentials: 'same-origin',
                headers: {
                    'Content-Type': 'application/json',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: JSON.stringify(payload)
            });
            return await res.json();
        }
    """, payload)

# =========================================================
# 🔹 LOGIN (EL QUE TE FUNCIONA)
# =========================================================
def esperar_login_real(page):
    for _ in range(25):
        if "login" not in page.url:
            return True
        page.wait_for_timeout(1000)
    return False

def login_seaap(page):

    print("🌐 Abriendo login...")

    page.goto(URL_LOGIN, timeout=60000)
    page.wait_for_selector("input[name='login']", timeout=30000)

    print("🔐 Enviando credenciales...")

    page.fill("input[name='login']", USUARIO)
    page.fill("input[name='password']", PASSWORD)
    page.click("button[type='submit']")

    page.wait_for_load_state("domcontentloaded")

    ok = esperar_login_real(page)

    print("🌐 URL actual:", page.url)

    if not ok or "login" in page.url:
        raise Exception("❌ Login falló")

    print("🟢 Login REAL exitoso")

    page.goto(URL_WEB, timeout=60000)
    page.wait_for_load_state("domcontentloaded")

    if "login" in page.url:
        raise Exception("❌ Sesión inválida")

    print("🟢 Sesión Odoo activa")

# =========================================================
# 🔹 SHEETS
# =========================================================
visitas_para_sheet = []
formatos_para_sheet = []

def registrar_visitas_sheet(dni, registros):

    fila = None
    hoja_destino = None

    for nombre, dic in dni_filas.items():
        if str(dni) in dic:
            fila = dic[str(dni)]
            hoja_destino = nombre
            break

    if not fila:
        return

    registros_validos = [
        r for r in registros if r.get("ficha") in [1,2,4,5]
    ]

    if not registros_validos:
        return

    registros_ordenados = sorted(
        registros_validos,
        key=lambda x: x.get("fecha_visita_1") or ""
    )

    columnas = ["Z","AC","AF"]

    colores = {
        1: {"red":0.75,"green":0.95,"blue":0.75},
        2: {"red":0.75,"green":0.95,"blue":0.75},
        4: {"red":1,"green":0.65,"blue":0.65},
        5: {"red":0.8,"green":0.65,"blue":0.95}
    }

    for i, reg in enumerate(registros_ordenados[:3]):

        fecha = reg.get("fecha_visita_1")
        if not fecha:
            continue

        ficha = int(reg.get("ficha", 0))
        col = columnas[i]

        visitas_para_sheet.append({
            "range": f"{hoja_destino}!{col}{fila}",
            "values": [[fecha]]
        })

        if ficha in colores:
            formatos_para_sheet.append({
                "hoja": hoja_destino,
                "celda": f"{col}{fila}",
                "color": colores[ficha]
            })

def enviar_visitas_a_sheet():

    if not visitas_para_sheet:
        print("📭 Sin datos para enviar")
        return

    print(f"📤 Enviando {len(visitas_para_sheet)} registros...")

    spreadsheet.values_batch_update({
        "valueInputOption": "USER_ENTERED",
        "data": visitas_para_sheet
    })

    if formatos_para_sheet:

        requests = []

        for f in formatos_para_sheet:
            row, col = gspread.utils.a1_to_rowcol(f["celda"])
            sheet_id = sheets[f["hoja"]].id

            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row-1,
                        "endRowIndex": row,
                        "startColumnIndex": col-1,
                        "endColumnIndex": col
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": f["color"]
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })

        spreadsheet.batch_update({"requests": requests})

    print("✅ Sheets actualizado")

# =========================================================
# 🔥 PROCESO OPTIMIZADO
# =========================================================
def procesar(page):

    print("📊 Cargando niños...")

    ninos = call_kw(page, {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "actividades.padron.nominal",
            "method": "search_read",
            "args": [[["parent_id","=",103]]],
            "kwargs": {
                "fields": ["id","name","documento_numero","actor_id","registro_ids"],
                "limit": 3000
            }
        },
        "id": 1
    }).get("result", [])

    print("👶 TOTAL:", len(ninos))

    filtrados = []

    for n in ninos:
        actor = n.get("actor_id")
        if not actor:
            continue

        dni_actor = extraer_dni_actor(actor[1])

        if dni_actor in ACTORES_VALIDOS_DNI:
            filtrados.append(n)

    print("🟢 FILTRADOS:", len(filtrados))

    all_ids = []
    mapa = {}

    for n in filtrados:
        ids = n.get("registro_ids", [])
        if ids:
            mapa[n["id"]] = ids
            all_ids.extend(ids)

    registros = call_kw(page, {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "actividades.registro",
            "method": "read",
            "args": [all_ids],
            "kwargs": {
                "fields": ["id","ficha","fecha_visita_1"]
            }
        },
        "id": 2
    }).get("result", [])

    reg_dict = {r["id"]: r for r in registros}

    for n in filtrados:

        dni = n.get("documento_numero")
        regs = [
            reg_dict[rid]
            for rid in mapa.get(n["id"], [])
            if rid in reg_dict
        ]

        if regs:
            registrar_visitas_sheet(dni, regs)

    enviar_visitas_a_sheet()

# =========================================================
# 🔹 MAIN
# =========================================================
def ejecutar():

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox","--disable-dev-shm-usage"]
        )

        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
            locale="es-PE"
        )

        page = context.new_page()

        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined })
        """)

        login_seaap(page)

        procesar(page)

        browser.close()

# =========================================================
if __name__ == "__main__":
    ejecutar()
