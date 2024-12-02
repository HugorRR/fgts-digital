import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import logging
import pyautogui
from pathlib import Path
import time
import pandas as pd
import os

competencia = '11 2024'

DIRETORIO_LOGS = Path(__file__).parent / 'logs'
logging.basicConfig(filename=f'{DIRETORIO_LOGS}/Log-RPA.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
DOWNLOAD_DIR = Path(__file__).parent / 'Competencias executadas'
PASTA_COMPETENCIA = DOWNLOAD_DIR / f'{competencia}'
IMAGEM_DIR = Path(__file__).parent / "img"
if not DOWNLOAD_DIR.exists():
    DOWNLOAD_DIR.mkdir()
CACHE = Path(__file__).parent / 'perfil-path'
perfil_path = str(CACHE)
options = uc.ChromeOptions()
options.add_argument("--user-data-dir=" + str(CACHE))
options.add_experimental_option("prefs",{"download.default_directory": str(PASTA_COMPETENCIA),"download.prompt_for_download": False,"profile.default_content_settings.popups": 0,})
driver = uc.Chrome(options=options)
driver.get('https://fgtsdigital.sistema.gov.br/portal/login')
driver.maximize_window()
driver.implicitly_wait(10)

def reconhecimento(imagens_referencia, tempo_limite, confidence=1.0):
    """Espera uma das imagens de referência aparecer na tela dentro de um tempo limite."""
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicao = pyautogui.locateCenterOnScreen(imagem_referencia, confidence=confidence)
            if posicao is not None:
                logging.info(f"Imagem encontrada: {imagem_referencia}")
                return True
            logging.info(f"Imagem não encontrada: {imagem_referencia}")
        time.sleep(1)
    logging.info("Nenhuma imagem encontrada dentro do tempo limite.")
    return False

def renomear_arquivo_recente(codigo, competencia):
    try:
        arquivos = list(PASTA_COMPETENCIA.glob("*"))
        arquivo_recente = max(arquivos, key=os.path.getctime)
        novo_nome = PASTA_COMPETENCIA / f"{codigo} - FGTS - {competencia}.pdf"
        arquivo_recente.rename(novo_nome)
        logging.info(f"Arquivo renomeado para: {novo_nome}")
    except Exception as e:
        logging.error(f"Erro ao renomear o arquivo: {e}")

def verificar_arquivo_baixado(codigo, competencia):
    """Verifica se o arquivo foi baixado corretamente com o nome esperado."""
    nome_esperado = PASTA_COMPETENCIA / f"{codigo} - FGTS - {competencia}.pdf"
    if nome_esperado.exists():
        logging.info(f"Arquivo baixado e verificado: {nome_esperado}")
        return True
    else:
        logging.warning(f"Arquivo não encontrado: {nome_esperado}")
        return False

df = pd.read_excel('Dados.xlsx', engine='openpyxl')
codigos = df['COD']
clientes = df['RAZAO SOCIAL']
cnpjs = df['CNPJ']
status = df['STATUS'].astype(str)

def lentidao():
    time.sleep(1.5)

def login():
    try:
        # Botão acesso gov
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//button[@class="br-button is-primary entrar"]')))
        driver.find_element(By.XPATH,'//button[@class="br-button is-primary entrar"]').click()
        lentidao()
        # Botão certificado
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//button[@id="login-certificate"]')))
        driver.find_element(By.XPATH,'//button[@id="login-certificate"]').click()
        lentidao()
        lentidao()
        pyautogui.press('enter')
        # Botão aceitar cookies
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="card0"]/div/div[2]/button[2]')))
        driver.find_element(By.XPATH, '//*[@id="card0"]/div/div[2]/button[2]').click()
        lentidao()
        # Botão definir
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '/html/body/modal-container/div[2]/div/fgtsd-modal-alterar-perfil/div[2]/div/button[2]')))
        driver.find_element(By.XPATH, '/html/body/modal-container/div[2]/div/fgtsd-modal-alterar-perfil/div[2]/div/button[2]').click()
        lentidao()
    except Exception as e:
        logging.error(f"Erro no login: {e}")
login()
for cliente, codigo, cnpj in zip(clientes,  codigos, cnpjs):
    # Verificacao se o arquivo ja fo baixado
    status_series = df.loc[df['CNPJ'] == cnpj, 'STATUS']
    if status_series.astype(str).str.contains('Sem procuracao|Nao ha debitos de interesse|Guia baixada|Erro desconhecido verificar o cliente', case=False, regex=True).any():
        continue
    try:
        time.sleep(2)    
        # Troca de perfil
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//button[@class="br-button secondary botao-barra-perfil"]')))
        driver.find_element(By.XPATH, '//button[@class="br-button secondary botao-barra-perfil"]').click()
        lentidao()
        # Perfil
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//div[@class="ng-input"]')))
        driver.find_element(By.XPATH, '//div[@class="ng-input"]').click()
        lentidao()
        # Selecionar procurador
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//span[text()="Procurador"]')))
        driver.find_element(By.XPATH, '//span[text()="Procurador"]').click()
        lentidao()
        # Campo CNPJ
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//input[@placeholder="Informe CNPJ ou CPF"]')))
        driver.find_element(By.XPATH, '//input[@placeholder="Informe CNPJ ou CPF"]').send_keys(cnpj)
        lentidao()
        # Botao Selecionar
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '/html/body/modal-container/div[2]/div/fgtsd-modal-alterar-perfil/div[2]/div/button[2]')))
        driver.find_element(By.XPATH, '/html/body/modal-container/div[2]/div/fgtsd-modal-alterar-perfil/div[2]/div/button[2]').click()
        lentidao()
        try:
            # Reconhecimento de procuracao
            reconhecimento(IMAGEM_DIR + 'erro_selecao.png', 10, confidence=0.8)
            # Adicionar observacao na planilha
            df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Sem procuracao'
            # Salvar o Excel
            df.to_excel('Dados.xlsx', index=False)
            # Botao cancelar
            WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//button[@class="br-button is-secondary"]')))
            driver.find_element(By.XPATH, '//button[@class="br-button is-secondary"]').click()
            lentidao()
            driver.get('https://fgtsdigital.sistema.gov.br/portal/servicos')
            lentidao()
            continue
        except Exception as e:
            pass
        # Botao gestao de Guias
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="divcard"]/br-card/div/div[2]')))
        driver.find_element(By.XPATH, '//*[@id="divcard"]/br-card/div/div[2]').click()
        lentidao()
        # Botao emissao de guias
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="divcard"]/br-card/div/div[1]')))
        driver.find_element(By.XPATH, '//*[@id="divcard"]/br-card/div/div[1]').click()
        lentidao()
        try:
            # Reconhecimento de erro
            reconhecimento([IMAGEM_DIR + 'rc_erro_1.png'], 10, confidence=0.8)
            # Adicionar informacao na planilha
            df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Nao ha debitos de interesse'
            # Salvar planilha
            df.to_excel('Dados.xlsx', index=False)
            # Botao Home
            lentidao()
            driver.get('https://fgtsdigital.sistema.gov.br/portal/servicos')
            continue
        except Exception as e:
            pass
        # Botao pesquisar
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="pesquisar"]/fieldset/div[2]/div/div[2]/button[2]')))
        driver.find_element(By.XPATH, '//*[@id="pesquisar"]/fieldset/div[2]/div/div[2]/button[2]').click()
        lentidao()
        # Botao Emitir guia
        WebDriverWait(driver, 7).until(EC.visibility_of_element_located((By.XPATH, '/html/body/app-root/fgtsd-main-layout/div/br-main-layout/div/div/div/main/div[2]/app-emissao-guia-rapida/div/app-agrupamento-guia-rapida/div/div/div/div[1]/div/button[1]')))
        lentidao()
        lentidao()
        driver.find_element(By.XPATH, '/html/body/app-root/fgtsd-main-layout/div/br-main-layout/div/div/div/main/div[2]/app-emissao-guia-rapida/div/app-agrupamento-guia-rapida/div/div/div/div[1]/div/button[1]').click()
        time.sleep(15)
        renomear_arquivo_recente(codigo, competencia)
        # Adiciona guia baixada na planilha
        df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'
        # Salva a planilha
        df.to_excel('Dados.xlsx', index=False) 
        # Botao Home
        lentidao()
        driver.get('https://fgtsdigital.sistema.gov.br/portal/servicos')
    except Exception as e:
        # Adicionar informacao na planilha
        df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Nao ha debitos de interesse'
        # Salvar planilha
        df.to_excel('Dados.xlsx', index=False)
        # Botao Home
        lentidao()
        driver.get('https://fgtsdigital.sistema.gov.br/portal/servicos')
        continue