import time
import os
import glob
import logging
import pandas as pd
import unicodedata
import threading
import pythoncom
import win32com.client

from PIL import ImageGrab
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from playwright.sync_api import sync_playwright

# =========================================================
# CONFIG
# =========================================================

CAMINHO_PASTA_MESTRE = r"C:\Users\mfs0045\Desktop\Produtividade Refrio"

CAMINHO_EXCEL_MESTRE = os.path.join(
    CAMINHO_PASTA_MESTRE,
    "Acompanhamento da Produtividade Refrio.xlsx"
)

USER_DATA_DIR = os.path.join(
    os.environ["LOCALAPPDATA"],
    r"Microsoft\Edge\User Data\AutomacaoWpp"
)

GRUPOS_WHATSAPP = [
    "Operação Refrio/Carrefour (R3)"
]

USUARIO = "brmna148"
SENHA = "Rochedo@15"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S"
)

lock = threading.Lock()

# =========================================================
# UTIL
# =========================================================

def normalizar_texto(txt):

    return ''.join(
        c for c in unicodedata.normalize('NFD', txt)
        if unicodedata.category(c) != 'Mn'
    ).lower()

# =========================================================
# LIMPA ARQUIVOS ANTIGOS
# =========================================================

def limpar_arquivos_antigos():

    try:

        logging.info("🗑️ Limpando downloads antigos...")

        arquivos = glob.glob(
            os.path.join(
                CAMINHO_PASTA_MESTRE,
                "*.xlsx"
            )
        )

        for arq in arquivos:

            nome = normalizar_texto(
                os.path.basename(arq)
            )

            if "acompanhamento da produtividade refrio" in nome:
                continue

            if "estocado" in nome or "pbl" in nome:

                try:

                    os.remove(arq)

                    logging.info(
                        f"🗑️ Removido: {os.path.basename(arq)}"
                    )

                except Exception as e:

                    logging.warning(
                        f"⚠️ Erro removendo {arq}: {e}"
                    )

    except Exception as e:

        logging.error(
            f"Erro limpando arquivos: {e}"
        )

# =========================================================
# SELECIONA SITE
# =========================================================

def selecionar_site(driver, wait):

    logging.info("🏢 Selecionando depósito 4945...")

    try:

        elemento = wait.until(
            EC.visibility_of_element_located(
                (By.ID, "ddlSite")
            )
        )

        wait.until(
            lambda d: len(
                Select(
                    d.find_element(By.ID, "ddlSite")
                ).options
            ) > 1
        )

        select = Select(elemento)

        selecionado = False

        for opt in select.options:

            if "4945" in opt.text:

                select.select_by_value(
                    opt.get_attribute("value")
                )

                logging.info(
                    f"✅ Selecionado: {opt.text}"
                )

                selecionado = True

                break

        if not selecionado:
            select.select_by_index(1)

        driver.execute_script("""
            document.getElementById('ddlSite')
            .dispatchEvent(
                new Event('change', { bubbles: true })
            );
        """)

        time.sleep(3)

    except Exception as e:

        logging.error(
            f"Erro ao selecionar site: {e}"
        )

# =========================================================
# FILTROS
# =========================================================

def realizar_config_filtros(driver, wait, tipo):

    try:

        logging.info(
            f"📅 Configurando datas para {tipo}"
        )

        data_fim = datetime.now().strftime(
            "%d/%m/%Y"
        )

        data_ini = (
            datetime.now() - timedelta(days=30)
        ).strftime("%d/%m/%Y")

        wait.until(
            EC.presence_of_element_located(
                (By.ID, "txtDataINI")
            )
        )

        driver.execute_script("""
            document.getElementById('txtDataINI')
            .value = arguments[0];
        """, data_ini)

        driver.execute_script("""
            document.getElementById('txtDataFIM')
            .value = arguments[0];
        """, data_fim)

        return True

    except Exception as e:

        logging.error(
            f"Erro nos filtros: {e}"
        )

        return False

# =========================================================
# FLUXO RELATÓRIO
# =========================================================

def fluxo_relatorio(driver, tipo):

    logging.info(
        f"🔍 Iniciando processo: {tipo}"
    )

    driver.get(
        "https://kairos.grpbg.br/Kairos/"
    )

    wait = WebDriverWait(driver, 45)

    try:

        menu = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//span[contains(text(),'Relatórios')]"
            ))
        )

        driver.execute_script(
            "arguments[0].click();",
            menu
        )

        xpath_tipo = f"""
        //a[contains(
        translate(text(),
        'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
        'abcdefghijklmnopqrstuvwxyz'),
        '{tipo.lower()}')]
        """

        sub = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, xpath_tipo)
            )
        )

        driver.execute_script(
            "arguments[0].click();",
            sub
        )

        link = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[contains(@href,'ItemSeparacao')]"
            ))
        )

        driver.execute_script(
            "arguments[0].click();",
            link
        )

        realizar_config_filtros(
            driver,
            wait,
            tipo
        )

        selecionar_site(driver, wait)

        btn_filtrar = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "btnFiltrar")
            )
        )

        driver.execute_script(
            "arguments[0].click();",
            btn_filtrar
        )

        for _ in range(120):

            try:

                barra = driver.find_element(
                    By.CSS_SELECTOR,
                    "div.progress-bar-success"
                )

                if "100%" in barra.get_attribute("style"):

                    logging.info(
                        f"✅ {tipo} pronto!"
                    )

                    return True

            except:
                pass

            time.sleep(5)

        return False

    except Exception as e:

        logging.error(
            f"Erro fluxo {tipo}: {e}"
        )

        return False

# =========================================================
# AGUARDA DOWNLOAD
# =========================================================

def aguardar_download(tipo, timeout=180):

    logging.info(
        f"📥 Aguardando download {tipo}..."
    )

    inicio = time.time()

    while time.time() - inicio < timeout:

        arquivos = glob.glob(
            os.path.join(
                CAMINHO_PASTA_MESTRE,
                "*.xlsx"
            )
        )

        candidatos = [
            a for a in arquivos
            if tipo.lower() in normalizar_texto(
                os.path.basename(a)
            )
        ]

        if candidatos:

            arq = max(
                candidatos,
                key=os.path.getctime
            )

            nome = os.path.basename(arq).lower()

            if (
                not nome.startswith("~$")
                and not nome.endswith(".crdownload")
            ):

                logging.info(
                    f"✅ Download concluído: "
                    f"{os.path.basename(arq)}"
                )

                return arq

        time.sleep(2)

    raise Exception(
        f"❌ Timeout download {tipo}"
    )

# =========================================================
# ATUALIZA EXCEL
# =========================================================

def atualizar_excel_com_downloads(
    wb,
    arquivo_estocado=None,
    arquivo_pbl=None
):

    try:

        logging.info("📊 Atualizando Excel...")

        excel = wb.Application

        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.Calculation = -4135

        # =================================================
        # ESTOCADO
        # =================================================

        if (
            arquivo_estocado
            and os.path.exists(arquivo_estocado)
        ):

            logging.info(
                "📥 Importando ESTOCADO..."
            )

            df_estocado = pd.read_excel(
                arquivo_estocado,
                dtype=str,
                usecols="A:Z"
            ).fillna("")

            ws_estocado = wb.Worksheets(
                "Itens Separação estocado"
            )

            ultima_linha_estocado = ws_estocado.Cells(
                ws_estocado.Rows.Count,
                1
            ).End(-4162).Row

            ws_estocado.Range(
                f"A2:Z{ultima_linha_estocado}"
            ).ClearContents()

            headers = [list(df_estocado.columns)]

            ws_estocado.Range(
                ws_estocado.Cells(1, 1),
                ws_estocado.Cells(
                    1,
                    len(headers[0])
                )
            ).Value = headers

            dados = df_estocado.values.tolist()

            if dados:

                total_linhas = len(dados)
                total_colunas = len(dados[0])

                ws_estocado.Range(
                    ws_estocado.Cells(2, 1),
                    ws_estocado.Cells(
                        total_linhas + 1,
                        total_colunas
                    )
                ).Value = dados

            logging.info(
                f"✅ ESTOCADO atualizado "
                f"({len(df_estocado)} linhas)"
            )

        # =================================================
        # PBL
        # =================================================

        if (
            arquivo_pbl
            and os.path.exists(arquivo_pbl)
        ):

            logging.info(
                "📥 Importando PBL..."
            )

            df_pbl = pd.read_excel(
                arquivo_pbl,
                dtype=str,
                usecols="A:Y"
            ).fillna("")

            ws_pbl = wb.Worksheets(
                "Itens Separação PBL"
            )

            ultima_linha_pbl = ws_pbl.Cells(
                ws_pbl.Rows.Count,
                1
            ).End(-4162).Row

            ws_pbl.Range(
                f"A2:Y{ultima_linha_pbl}"
            ).ClearContents()

            headers = [list(df_pbl.columns)]

            ws_pbl.Range(
                ws_pbl.Cells(1, 1),
                ws_pbl.Cells(
                    1,
                    len(headers[0])
                )
            ).Value = headers

            dados = df_pbl.values.tolist()

            if dados:

                total_linhas = len(dados)
                total_colunas = len(dados[0])

                ws_pbl.Range(
                    ws_pbl.Cells(2, 1),
                    ws_pbl.Cells(
                        total_linhas + 1,
                        total_colunas
                    )
                ).Value = dados

            logging.info(
                f"✅ PBL atualizado "
                f"({len(df_pbl)} linhas)"
            )

        wb.RefreshAll()

        excel.CalculateFull()

        time.sleep(10)

        wb.Save()

        excel.Calculation = -4105
        excel.ScreenUpdating = True

        logging.info(
            "✅ Excel atualizado e salvo"
        )

    except Exception as e:

        logging.error(
            f"Erro atualizando Excel: {e}"
        )

        raise

# =========================================================
# BITMAP
# =========================================================

def gerar_bitmap_excel(wb):

    try:

        logging.info(
            "🖼️ Copiando imagem da Planilha6..."
        )

        ws = wb.Worksheets("Planilha6")

        ws.Activate()

        excel = ws.Application

        excel.Visible = True
        excel.WindowState = -4137
        excel.ActiveWindow.Zoom = 90

        rng = ws.Range("A1:R31")

        rng.CopyPicture(
            Appearance=1,
            Format=2
        )

        time.sleep(3)

        logging.info(
            "✅ Imagem copiada"
        )

    except Exception as e:

        logging.error(
            f"Erro ao copiar imagem: {e}"
        )

        raise

# =========================================================
# WHATSAPP
# =========================================================

def enviar_whatsapp(wb):

    logging.info(
        "📱 Abrindo WhatsApp..."
    )

    gerar_bitmap_excel(wb)

    with sync_playwright() as p:

        browser = p.chromium.launch_persistent_context(
            USER_DATA_DIR,
            headless=False,
            channel="msedge"
        )

        page = browser.pages[0]

        page.goto(
            "https://web.whatsapp.com"
        )

        page.wait_for_selector(
            'div[data-tab="3"]',
            timeout=120000
        )

        time.sleep(5)

        for grupo in GRUPOS_WHATSAPP:

            logging.info(
                f"📨 Enviando para {grupo}"
            )

            page.keyboard.press("Control+Alt+/")

            time.sleep(2)

            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")

            page.keyboard.type(grupo)

            time.sleep(3)

            page.keyboard.press("Enter")

            time.sleep(3)

            page.keyboard.press("Control+V")

            logging.info(
                "📷 Imagem colada"
            )

            time.sleep(5)

            legenda = (
                f"Segue Produtividade Hora a Hora "
                f"CD Refrio 4945. "
                f"Ultima Atualização as "
                f"{datetime.now().strftime('%H:%M')}"
            )

            page.keyboard.type(legenda)

            time.sleep(2)

            page.keyboard.press("Enter")

            logging.info(
                "✅ WhatsApp enviado"
            )

            time.sleep(60)

        browser.close()

# =========================================================
# EXECUÇÃO PRINCIPAL
# =========================================================

def executar():

    if not lock.acquire(blocking=False):

        logging.warning(
            "⚠️ Execução já em andamento."
        )

        return

    driver = None

    try:

        if datetime.now().weekday() > 4:

            logging.info(
                "😴 Final de semana."
            )

            return

        opts = EdgeOptions()

        prefs = {
            "download.default_directory":
            CAMINHO_PASTA_MESTRE
        }

        opts.add_experimental_option(
            "prefs",
            prefs
        )

        driver = webdriver.Edge(
            options=opts
        )

        driver.maximize_window()

        limpar_arquivos_antigos()

        # LOGIN
        logging.info("🔐 Fazendo login...")

        driver.get(
            "https://kairos.grpbg.br/Kairos/Login/LogInLDAP"
        )

        wait_login = WebDriverWait(
            driver,
            30
        )

        wait_login.until(
            EC.element_to_be_clickable(
                (By.ID, "txtUser")
            )
        ).send_keys(USUARIO)

        driver.find_element(
            By.ID,
            "txtSenha"
        ).send_keys(SENHA)

        driver.find_element(
            By.TAG_NAME,
            "button"
        ).click()

        time.sleep(5)

        arquivo_estocado = None
        arquivo_pbl = None

        # ESTOCADO
        sucesso_estocado = fluxo_relatorio(
            driver,
            "Estocado"
        )

        if sucesso_estocado:

            arquivo_estocado = aguardar_download(
                "estocado"
            )

        # PBL
        sucesso_pbl = fluxo_relatorio(
            driver,
            "PBL"
        )

        if sucesso_pbl:

            arquivo_pbl = aguardar_download(
                "pbl"
            )

        # EXCEL + WHATSAPP
        if sucesso_estocado or sucesso_pbl:

            logging.info(
                "📊 Abrindo Excel..."
            )

            driver.quit()
            driver = None

            pythoncom.CoInitialize()

            excel_app = win32com.client.DispatchEx(
                "Excel.Application"
            )

            excel_app.Visible = True

            wb = excel_app.Workbooks.Open(
                CAMINHO_EXCEL_MESTRE
            )

            atualizar_excel_com_downloads(
                wb,
                arquivo_estocado,
                arquivo_pbl
            )

            time.sleep(5)

            enviar_whatsapp(wb)

            wb.Close(
                SaveChanges=True
            )

            excel_app.Quit()

            logging.info(
                "🎯 Processo finalizado!"
            )

        else:

            logging.warning(
                "⚠️ Nenhum relatório baixado."
            )

    except Exception as e:

        logging.error(
            f"💥 ERRO CRÍTICO: {e}"
        )

    finally:

        try:

            if driver:
                driver.quit()

        except:
            pass

        lock.release()

# =========================================================
# MAIN
# =========================================================

if __name__ == "__main__":

    logging.info(
        "🚀 Iniciando automação..."
    )

    executar()