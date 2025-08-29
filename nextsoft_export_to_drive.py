# nextsoft_export_to_drive.py
# Fluxo:
# 1) Login
# 2) Troca de loja (via secret APPNEXT_LOJA_DESTINO)
# 3) Navega: Vendas > Vendedor Analítico
# 4) Abre filtros
# 5) Captura URL do Excel; se não der, fallback via API
# 6) Envia ao Google Drive via rclone

import os, sys, subprocess, traceback, shutil, json, unicodedata, re
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, parse_qs, urlencode

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ================= LOG =================
def log(msg: str):
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {msg}", flush=True)

def step(title: str):
    log("="*70); log(title); log("="*70)

load_dotenv()

# ================ CONFIG ================
REDE  = os.getenv("APPNEXT_REDE", "").strip()
USER  = os.getenv("APPNEXT_USER", "").strip()
PASS  = os.getenv("APPNEXT_PASS", "").strip()

DRIVE_REMOTE    = os.getenv("DRIVE_REMOTE", "GDRIVE:")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID", "").strip()
DRIVE_FILE_NAME = os.getenv("DRIVE_FILE_NAME", "importacaoA.xlsx").strip()

DATA_INICIO = os.getenv("DATA_INICIO", "01/06/2025").strip()
DATA_FIM    = os.getenv("DATA_FIM", "").strip()  # vazio => hoje 23:59:59

TEMPLATE_XLSX = os.getenv("TEMPLATE_XLSX", "importacaoA.xlsx")
TEMPLATE_HEADER_ROW = int(os.getenv("TEMPLATE_HEADER_ROW", "2"))

LOJA_ID_DESTINO = os.getenv("APPNEXT_LOJA_ID_DESTINO", "").strip()

RCLONE_PATH = os.getenv("RCLONE_PATH") or shutil.which("rclone") or r"C:\rclone\rclone.exe"
if not Path(RCLONE_PATH).exists() and shutil.which("rclone") is None:
    log("ERRO: rclone não encontrado.")
    sys.exit(1)

if not (REDE and USER and PASS):
    log("ERRO: Preencha APPNEXT_REDE / APPNEXT_USER / APPNEXT_PASS")
    sys.exit(1)

TS = datetime.now().strftime("%Y%m%d_%H%M%S")
LOCAL_OUT = Path.cwd() / f"export_nextsoft_{TS}.xlsx"
LOGIN_URL = "https://www.appnext.com.br/#/login"

# -------- datas ----------
def to_dt(s: str, end=False):
    try:
        d = datetime.strptime(s, "%d/%m/%Y")
        return d.replace(hour=23, minute=59, second=59) if end else d
    except Exception:
        return None

if not DATA_FIM:
    DATA_FIM = datetime.now().strftime("%d/%m/%Y")

INI_DT = to_dt(DATA_INICIO) or datetime(2025, 6, 1)
FIM_DT = to_dt(DATA_FIM, end=True) or datetime.now().replace(hour=23, minute=59, second=59)

FMT_JSON_INI = INI_DT.strftime("%Y-%m-%dT%H:%M:%S.000Z")
FMT_JSON_FIM = FIM_DT.strftime("%Y-%m-%dT%H:%M:%S.000Z")
FMT_BR_INI   = INI_DT.strftime("%d/%m/%Y, %H:%M:%S")
FMT_BR_FIM   = FIM_DT.strftime("%d/%m/%Y, %H:%M:%S")

# -------- seletores ----------
SEL = {
    "rede":  "input[placeholder='Rede']",
    "email": "input[placeholder='Email']",
    "senha": "input[placeholder='Senha']",
    "entrar": "button:has-text('Entrar'), button:has-text('Login')",
    "menu_vendas": "nav >> text=Vendas, header >> text=Vendas, a[href*='vendas']:has-text('Vendas')",
    "item_vendedor_analitico": "a[href*='vendedor-analitico'], a:has-text('Vendedor Anal')",
    "titulo_rel": "text=Listagem de Vendedor Analítico",
    "btn_filtros": "[title*='Filtro'], button:has(i.mdi-filter)",
    "pane_filtros": "#filtrosForm",
    "btn_atualizar": "button:has-text('Atualizar Filtros')",
    "grid_row": "table tbody tr",
    "excel_btn": "#dataTableButtons button.buttons-excel, button.buttons-excel",
}

# ---------- Helpers ----------
def rclone_copy_latest(local_file: Path):
    cmd = [RCLONE_PATH, "copyto", str(local_file),
           f"{DRIVE_REMOTE}{DRIVE_FILE_NAME}",
           "--drive-root-folder-id", DRIVE_FOLDER_ID, "-v"]
    log(f"rclone -> {' '.join(cmd)}")
    res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if res.stdout.strip():
        print(res.stdout)
    if res.returncode != 0:
        print(res.stderr, file=sys.stderr)
        raise RuntimeError("Falha ao atualizar no Drive.")

def _click_first(page, selectors):
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count():
                loc.scroll_into_view_if_needed(timeout=1500)
                loc.click(timeout=4000, force=True)
                return True
        except Exception:
            pass
    return False

# ---------- Troca de Loja ----------
def trocar_loja(page, loja_nome=None):
    alvo = (loja_nome or os.getenv("APPNEXT_LOJA_DESTINO")
            or "GOIANIA - TEA SHOP FLAMBOYANT")
    step(f"2) Trocando loja → {alvo}")
    try:
        opened = _click_first(page, ["button:has-text('TEA SHOP')",
                                     "[data-bs-toggle='dropdown']",
                                     ".dropdown-toggle"])
        if not opened:
            _click_first(page, ["xpath=(//i[contains(@class,'store')]/ancestor::*[self::button])[1]"])

        try:
            page.get_by_role("menuitem", name=re.compile(alvo, re.I)).first.click(timeout=6000, force=True)
        except Exception:
            _click_first(page, [f"text=^{alvo}$"])
        log(f"Loja selecionada: {alvo}")
    except Exception as e:
        log(f"Falha ao trocar loja: {e}")

# ---------- Goto com retry ----------
def goto_login_with_retries(page, url, tries=3):
    for i in range(1, tries+1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=180_000)
            return
        except Exception as e:
            log(f"[goto-retry] tentativa {i}/{tries} falhou: {e}")
            if i == tries:
                raise
            page.wait_for_timeout(5000)

# ================= MAIN =================
def main():
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(Path.cwd() / "pw_state"),
            headless=True,
            accept_downloads=True,
            viewport={"width": 1600, "height": 950},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(180_000)
        page.set_default_navigation_timeout(240_000)

        page.on("console", lambda m: log(f"[console] {m.type}: {m.text}"))
        page.on("requestfailed", lambda r: log(f"[requestfailed] {r.url} -> {r.failure}"))

        try:
            step("1) Login")
            goto_login_with_retries(page, LOGIN_URL)

            if "/#/login" in page.url:
                page.fill(SEL["rede"], REDE)
                page.fill(SEL["email"], USER)
                page.fill(SEL["senha"], PASS)
                page.click(SEL["entrar"])
            page.wait_for_url("**/#/loja/**", timeout=120_000)

            trocar_loja(page, "GOIANIA - TEA SHOP FLAMBOYANT")
             #trocar_loja(page, os.getenv("APPNEXT_LOJA_DESTINO"))

            # Aqui seguem os passos 3 a 6 (iguais ao seu original)...
            # Ir para Vendas > Vendedor Analítico, abrir filtros, exportar Excel, fallback e enviar pro Drive.

        except Exception:
            log("### ERRO DURANTE O FLUXO ###")
            print(traceback.format_exc())
        finally:
            try: context.close()
            except Exception: pass

    step("FINALIZADO")
    log(f"Local: {LOCAL_OUT.resolve()}")
    log(f"Drive: {DRIVE_FILE_NAME}  (pasta ID {DRIVE_FOLDER_ID})")

if __name__ == "__main__":
    main()
