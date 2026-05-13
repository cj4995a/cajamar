import os, logging, time, shutil, warnings
import pandas as pd
import win32com.client
import schedule  # Biblioteca para agendamento
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from playwright.sync_api import sync_playwright

# --- CONFIGURAÇÃO ---
URL_SAD = "https://sadsams.grpbg.br/DBSAD/servlet/MTFTR"
PASTA_DESTINO_FINAL = r"C:\Users\mfs0045\Desktop\Brasil\Download_MTFTR"
CAMINHO_EXCEL_MESTRE = r"C:\Users\mfs0045\Desktop\Brasil\Demanda Sams Brasil Atual.xlsx"
DOWNLOAD_TEMP = os.path.join(os.getcwd(), "temp_download")
USER_DATA_DIR = os.path.join(os.environ['LOCALAPPDATA'], r"Microsoft\Edge\User Data\AutomacaoWpp")

NOME_ABA_DADOS = "FTR254"
NOME_ABA_FIFO = "FIFO"
NOME_ABA_REPORT = "Report"
GRUPOS_WHATSAPP = ["Area 51 - PCP"]
FILIAIS = ["4946","4955","6003","6008","4995","4945","6010","4996"]

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(message)s', datefmt='%H:%M:%S')
warnings.filterwarnings("ignore")

# ... (Funções tratar_alerta e processar_e_enviar permanecem iguais) ...

def tratar_alerta(driver, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(EC.alert_is_present())
        driver.switch_to.alert.accept()
        return True
    except: return None

def processar_e_enviar():
    logging.info("📊 Tratando dados com conversão numérica (Data Excel)...")
    arquivos_csv = [os.path.join(PASTA_DESTINO_FINAL, f) for f in os.listdir(PASTA_DESTINO_FINAL) if f.endswith('.csv')]
    if not arquivos_csv: return False

    lista_dfs = []
    for f in arquivos_csv:
        df_temp = pd.read_csv(f, sep=None, engine='python', encoding='latin-1', dtype=str)
        lista_dfs.append(df_temp)
    
    df_novos = pd.concat(lista_dfs, ignore_index=True)
    df_novos = df_novos.iloc[:, :38]

    def data_para_excel_num(valor):
        if pd.isna(valor) or str(valor).strip() == "": return None
        try:
            dt = pd.to_datetime(valor, dayfirst=True, errors='coerce')
            if pd.isna(dt): return None
            if dt.tzinfo is not None:
                dt = dt.tz_localize(None)
            data_base = datetime(1899, 12, 30)
            delta = dt.to_pydatetime() - data_base
            return delta.days + (delta.seconds / 86400.0)
        except: return None

    df_novos.iloc[:, 7] = df_novos.iloc[:, 7].apply(data_para_excel_num)

    def limpar_numero(valor):
        if pd.isna(valor) or str(valor).strip() == "": return 0.0
        s = str(valor).replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.').strip()
        try: return float(s)
        except: return 0.0

    df_novos.iloc[:, 13] = df_novos.iloc[:, 13].apply(limpar_numero)
    df_novos.iloc[:, 32] = df_novos.iloc[:, 32].apply(limpar_numero)

    dados_lista = df_novos.where(pd.notnull(df_novos), None).values.tolist()

    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(CAMINHO_EXCEL_MESTRE, UpdateLinks=0)
        ws_dados = wb.Worksheets(NOME_ABA_DADOS)
        
        u_linha = ws_dados.Cells(ws_dados.Rows.Count, 1).End(-4162).Row
        if u_linha >= 2: ws_dados.Range(f"A2:AL{u_linha}").ClearContents()
        
        ws_dados.Columns(8).NumberFormat = "dd/mm/yyyy"

        if dados_lista:
            ws_dados.Range(ws_dados.Cells(2, 1), ws_dados.Cells(len(dados_lista) + 1, 38)).Value = dados_lista

        logging.info("🔄 Atualizando fórmulas e preparando envios...")
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        time.sleep(5)

        with sync_playwright() as p:
            context = p.chromium.launch_persistent_context(USER_DATA_DIR, headless=False, channel="msedge", no_viewport=True, args=["--start-maximized"])
            while len(context.pages) > 1: context.pages[-1].close()
            page = context.pages[0]
            page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")
            
            try:
                btn = page.wait_for_selector('text="Usar nesta janela"', timeout=8000)
                if btn: btn.click()
            except: pass

            page.wait_for_selector('//div[@data-tab="3"]', timeout=120000)

            config_envios = [
                {"aba": NOME_ABA_REPORT, "range": "A1:J114", "legenda": "Segue Carteira Sam's Brasil"},
                {"aba": NOME_ABA_FIFO, "range": "A1:L45", "legenda": "Segue FIFO BR"}
            ]

            for envio in config_envios:
                logging.info(f"📸 Capturando e enviando: {envio['aba']}")
                ws = wb.Worksheets(envio["aba"])
                ws.Activate()
                ws.Range(envio["range"]).CopyPicture(Appearance=1, Format=2)
                time.sleep(2)

                for grupo in GRUPOS_WHATSAPP:
                    page.keyboard.press("Control+Alt+/")
                    page.wait_for_timeout(1000)
                    page.keyboard.type(grupo, delay=100)
                    page.wait_for_timeout(2000)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(2000)

                    page.locator('//div[@data-tab="10"]').first.click()
                    page.wait_for_timeout(1000)
                    page.keyboard.press("Control+V")
                    page.wait_for_timeout(3000) 
                    page.keyboard.type(envio["legenda"], delay=100)
                    page.wait_for_timeout(1000)
                    page.keyboard.press("Enter")
                    logging.info(f"✅ {envio['aba']} enviado para {grupo}")
                    
            logging.info("⏳ Aguardando conclusão do upload final...")
            page.wait_for_timeout(8000) 
            context.close()

        wb.Save()
        wb.Close()
        return True
    except Exception as e:
        logging.error(f"❌ Erro: {e}")
        return False
    finally:
        if excel: excel.Quit()

def executar_automacao():
    logging.info(f"🚀 Iniciando ciclo de automação: {datetime.now().strftime('%H:%M:%S')}")
    for pasta in [PASTA_DESTINO_FINAL, DOWNLOAD_TEMP]:
        if os.path.exists(pasta): shutil.rmtree(pasta)
        os.makedirs(pasta)

    options = EdgeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", {"download.default_directory": DOWNLOAD_TEMP})
    options.add_experimental_option('excludeSwitches', ['enable-logging']) 
    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 35)

    try:
        logging.info("🚀 Iniciando Extração SAD...")
        driver.get(URL_SAD)
        wait.until(EC.presence_of_element_located((By.ID, "TL_USUARIO"))).send_keys("524537")
        driver.find_element(By.ID, "XTL_SENHA").send_keys("Rochedo0")
        driver.find_element(By.ID, "TL_EMP").send_keys("01")
        driver.find_element(By.ID, "TL_FILIAL_WM").send_keys("4945")
        driver.find_element(By.ID, "BTN_LOGIN_FX").click()
        time.sleep(3)
        tratar_alerta(driver)

        for filial in FILIAIS:
            logging.info(f"🏢 Filial: {filial}")
            wait.until(EC.presence_of_element_located((By.ID, "auxProxTela"))).send_keys("mtftr\n")
            time.sleep(2)
            campo = wait.until(EC.presence_of_element_located((By.ID, "TL_FILIAL")))
            campo.clear()
            campo.send_keys(filial)
            
            dt_ini = (datetime.today() - timedelta(days=90)).strftime("%d/%m/%Y")
            driver.find_element(By.ID, "dtIni").clear()
            driver.find_element(By.ID, "dtIni").send_keys(dt_ini)
            driver.find_element(By.ID, "BT_INQ").click()
            time.sleep(4)
            
            wait.until(EC.element_to_be_clickable((By.ID, "BT_EXP"))).click()
            if tratar_alerta(driver, timeout=100):
                wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'DownLoad Arquivo')]"))).click()
                time.sleep(6)
                for arq in os.listdir(DOWNLOAD_TEMP):
                    if arq.endswith(".csv"):
                        shutil.move(os.path.join(DOWNLOAD_TEMP, arq), os.path.join(PASTA_DESTINO_FINAL, arq))
            
            driver.find_element(By.ID, "imglogo").click()
            time.sleep(2)
            
        driver.quit()
        if processar_e_enviar():
            logging.info("🎉 Automação concluída com sucesso!")

    except Exception as e:
        logging.error(f"❌ Falha crítica: {e}")
        if 'driver' in locals(): driver.quit()

# --- LÓGICA DE AGENDAMENTO ---

if __name__ == "__main__":

    executar_automacao()