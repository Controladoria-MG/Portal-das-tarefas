import time
import os
import sys

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

URL_LOGIN = "https://aplicativo.mgcontecnica.com.br/#/home"
USUARIO = "rjuan"
SENHA = "Palmeiras!"

DATA_INICIO = "01/06/2026"
DATA_FIM = "30/06/2026"

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_DESTINO    = os.path.join(_RAIZ, "dados", "base")
PASTA_COORDENADOR = os.path.join(_RAIZ, "dados", "coordenador")

TIMEOUT_ELEMENTO   = 60
TIMEOUT_DOWNLOAD   = 1500
TIMEOUT_SELENIUM_CMD = 1500

RELATORIOS = [
    {"xpath": "//h5[contains(text(),'São Paulo (Geral)')]",      "nome": "sao_paulo"},
    {"xpath": "//h5[contains(text(),'Rio de Janeiro (Geral)')]", "nome": "rio_de_janeiro"},
    {"xpath": "//h5[contains(text(),'Santos (Geral)')]",         "nome": "santos"},
]


# ── Selenium ──────────────────────────────────────────────────────────────────

def _criar_driver(pasta_destino: str) -> webdriver.Chrome:
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--log-level=3")
    options.add_argument("--disable-logging")
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_experimental_option("prefs", {
        "download.default_directory": pasta_destino,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
    })
    options.add_argument("--safebrowsing-disable-download-protection")

    original_stdout = sys.stdout
    sys.stdout = open(os.devnull, "w")
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    finally:
        sys.stdout = original_stdout

    driver.command_executor.client_config.timeout = TIMEOUT_SELENIUM_CMD
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": pasta_destino},
    )
    return driver


def _navegar_e_exportar(driver, wait, xpath_relatorio: str, data_inicio: str, data_fim: str, pasta_destino: str) -> set:
    xlsx_antes = set(f for f in os.listdir(pasta_destino) if f.endswith(".xlsx"))

    wait.until(EC.element_to_be_clickable((By.XPATH, "//h5[contains(text(),'Relatórios')]"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//h5[contains(text(),'Personalizados')]"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, xpath_relatorio))).click()
    time.sleep(1.5)

    wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_VencimentoRadioButton"))).click()
    time.sleep(2)

    campo_inicio = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_DataInicioTextBox")))
    campo_fim    = wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_DataFimTextBox")))
    campo_inicio.clear(); time.sleep(0.5); campo_inicio.send_keys(data_inicio); time.sleep(0.5)
    campo_fim.clear();    time.sleep(0.5); campo_fim.send_keys(data_fim);       time.sleep(0.5)

    wait.until(EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ExportarRelatorioLinkButton"))).click()

    return xlsx_antes


def _aguardar_e_renomear(pasta_destino: str, xlsx_antes: set, nome_final: str) -> str:
    tempo_inicio = time.time()
    arquivo_final = None

    while (time.time() - tempo_inicio) < TIMEOUT_DOWNLOAD:
        arquivos = os.listdir(pasta_destino)
        if any(a.endswith(".crdownload") for a in arquivos):
            time.sleep(2)
            continue
        novos = [f for f in arquivos if f.endswith(".xlsx") and f not in xlsx_antes]
        if novos:
            arquivo_final = novos[0]
            break
        time.sleep(2)

    if not arquivo_final:
        raise RuntimeError(f"Download de '{nome_final}' não concluído em {TIMEOUT_DOWNLOAD}s.")

    caminho_antigo = os.path.join(pasta_destino, arquivo_final)
    caminho_novo   = os.path.join(pasta_destino, nome_final)
    if os.path.exists(caminho_novo):
        os.remove(caminho_novo)
    os.rename(caminho_antigo, caminho_novo)

    for arq in os.listdir(pasta_destino):
        if arq.lower() == "downloads.htm":
            os.remove(os.path.join(pasta_destino, arq))

    return caminho_novo


def _baixar_um(rel: dict, data_inicio: str, data_fim: str, pasta_destino: str) -> str:
    driver = _criar_driver(pasta_destino)
    wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)

    try:
        driver.get(URL_LOGIN)
        time.sleep(2)

        try:
            campo_usuario = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "usuario"))
            )
            campo_usuario.send_keys(USUARIO)
            driver.find_element(By.ID, "senha").send_keys(SENHA)
            wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Entrar']"))
            ).click()
            time.sleep(2)
        except Exception:
            pass

        wait.until(EC.element_to_be_clickable((By.XPATH, "//h6[@title='MG Controle']"))).click()
        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[-1])

        xlsx_antes = _navegar_e_exportar(
            driver, wait, rel["xpath"], data_inicio, data_fim, pasta_destino
        )
        return _aguardar_e_renomear(pasta_destino, xlsx_antes, f"{rel['nome']}.xlsx")

    finally:
        time.sleep(3)
        driver.quit()


# ── Pós-processamento ─────────────────────────────────────────────────────────

def _carregar_coordenadores() -> dict:
    arquivos = [f for f in os.listdir(PASTA_COORDENADOR) if f.endswith(".xlsx")]
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo xlsx encontrado em {PASTA_COORDENADOR}")
    caminho = os.path.join(PASTA_COORDENADOR, arquivos[0])
    df = pd.read_excel(caminho)
    return dict(zip(df["Nome de Exibição"], df["Coordenador"]))


def _salvar_xlsx(df: pd.DataFrame, caminho: str) -> None:
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Pendencias", index=False)


def _processar(caminho_xlsx: str, nome: str, pasta_destino: str, mapa_coord: dict) -> None:
    df = pd.read_excel(caminho_xlsx, sheet_name="Pendencias")
    df["Coordenador"] = df["UsuarioResponsavel"].map(mapa_coord)

    if nome == "sao_paulo":
        df_sp    = df[df["Unidade"] == "SP"].reset_index(drop=True)
        df_goias = df[df["Unidade"] == "GOIAS"].reset_index(drop=True)
        _salvar_xlsx(df_sp,    os.path.join(pasta_destino, "sao_paulo.xlsx"))
        _salvar_xlsx(df_goias, os.path.join(pasta_destino, "goias.xlsx"))
        print(f"       São Paulo: {len(df_sp)} linhas | Goiás: {len(df_goias)} linhas")
    else:
        _salvar_xlsx(df, caminho_xlsx)


# ── Orquestração ──────────────────────────────────────────────────────────────

def baixar_relatorios(data_inicio: str, data_fim: str, pasta_destino: str) -> None:
    os.makedirs(pasta_destino, exist_ok=True)
    mapa_coord = _carregar_coordenadores()

    for i, rel in enumerate(RELATORIOS, 1):
        print(f"[{i}/{len(RELATORIOS)}] Baixando: {rel['nome']}...")
        caminho = _baixar_um(rel, data_inicio, data_fim, pasta_destino)
        print(f"       Download concluído. Processando...")
        _processar(caminho, rel["nome"], pasta_destino, mapa_coord)
        print(f"       Salvo em: {pasta_destino}")


if __name__ == "__main__":
    print(f"Baixando relatórios de {DATA_INICIO} a {DATA_FIM}...")
    baixar_relatorios(DATA_INICIO, DATA_FIM, PASTA_DESTINO)
