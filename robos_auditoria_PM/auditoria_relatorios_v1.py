import os
import time
import glob
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

BASE_DIR = r"C:"

PASTA_AUDITORIA = os.path.join(BASE_DIR, "auditoria relatorios ponto mais")

if not os.path.exists(PASTA_AUDITORIA):
    os.makedirs(PASTA_AUDITORIA)

URL_PAGINA = "https://" 

nomes_colaboradores = []

PASTA_DOCUMENTOS = os.path.join(PASTA_AUDITORIA, "Documentos_Baixados")
PASTA_DOWNLOADS_TEMP = os.path.join(PASTA_AUDITORIA, "Downloads_Temp")
CAMINHO_EXCEL = os.path.join(PASTA_AUDITORIA, "Relatorio_Auditoria_Britanico.xlsx")

for pasta in [PASTA_DOCUMENTOS, PASTA_DOWNLOADS_TEMP]:
    if not os.path.exists(pasta):
        os.makedirs(pasta)

relatorio = []

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

prefs = {
    "download.default_directory": PASTA_DOWNLOADS_TEMP,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True 
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 15)

def atualizar_excel_tempo_real(nome, status, detalhes="N/A"):
    relatorio.append({"Colaborador": nome, "Status": status, "Dias e Horários Suspeitos": detalhes})
    pd.DataFrame(relatorio).to_excel(CAMINHO_EXCEL, index=False)

def mover_pdf_baixado(nome_colaborador):
    tempo_espera = 0
    while tempo_espera < 30:
        arquivos = glob.glob(os.path.join(PASTA_DOWNLOADS_TEMP, "*"))
        arquivos_validos = [f for f in arquivos if not f.endswith(".crdownload")]
        
        if arquivos_validos:
            arquivo_recente = max(arquivos_validos, key=os.path.getctime)
            nome_arquivo_original = os.path.basename(arquivo_recente)
            
            novo_caminho = os.path.join(PASTA_DOCUMENTOS, nome_arquivo_original)
            
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)
                
            shutil.move(arquivo_recente, novo_caminho)
            print(f"[{nome_colaborador}] PDF salvo como: {nome_arquivo_original}")
            return True
            
        time.sleep(1)
        tempo_espera += 1
        
    print(f"[{nome_colaborador}] Aviso: Falha ou lentidão ao baixar o PDF.")
    return False

print("Abrindo o navegador...")
driver.get(URL_PAGINA) 

print("\n" + "="*70)
print("AÇÃO NECESSÁRIA:")
print("1. Faça o login.")
print("2. Vá até Relatórios > Ajustes de Ponto.")
print("3. Preencha o Período.")
print("4. Selecione Filtrar por 'Colaborador'.")
print("="*70 + "\n")

input(">>> Quando TUDO estiver preenchido, aperte [ENTER] aqui no terminal para começar <<<")

print("\nIniciando as buscas em lote...\n")

for nome in nomes_colaboradores:
    print(f"\n[{nome}] Iniciando...")

    try:
        try:
            btn_limpar = driver.find_element(By.XPATH, "//div[@data-testid='employee_id-input']//span[@title='Limpar']")
            driver.execute_script("arguments[0].click();", btn_limpar)
            time.sleep(1)
        except NoSuchElementException:
            pass 

        campo_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-testid='employee_id-input']//div[@role='combobox']")))
        try:
            campo_combobox.click()
        except:
            driver.execute_script("arguments[0].click();", campo_combobox)
        time.sleep(1.5) 

        inputs_busca = driver.find_elements(By.XPATH, "//input[@placeholder='Digite para buscar']")
        input_busca_visivel = next((inp for inp in inputs_busca if inp.is_displayed()), None)
        
        if not input_busca_visivel:
            raise Exception("O campo de digitar o nome não apareceu.")

        input_busca_visivel.clear()
        input_busca_visivel.send_keys(nome)
        
        print(f"[{nome}] Aguardando a API retornar o nome correto na lista...")
        tempo_limite = time.time() + 15 
        opcao_correta = None
        
        while time.time() < tempo_limite:
            opcoes_lista = driver.find_elements(By.CSS_SELECTOR, "div.ng-option")
            for opcao in opcoes_lista:
                texto_opcao = opcao.text.strip().upper() 
                if nome.upper() in texto_opcao:
                    opcao_correta = opcao
                    break 
            
            if opcao_correta:
                break 
            time.sleep(0.5) 
            
        if not opcao_correta:
            raise Exception(f"A pesquisa não encontrou correspondência exata para '{nome}'.")

        driver.execute_script("arguments[0].click();", opcao_correta) 
        time.sleep(1)

        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)

        print(f"[{nome}] Gerando relatório...")
        botao_gerar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Gerar relatório')]]")))
        driver.execute_script("arguments[0].click();", botao_gerar)
        time.sleep(6) 
        
        print(f"[{nome}] Analisando tabela de horários...")
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody/tr")))
        except:
            pass
            
        linhas_tabela = driver.find_elements(By.XPATH, "//table//tbody/tr")
        horarios_britanicos = []
        
        for linha in linhas_tabela:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            
            if len(colunas) > 5: 
                data_registro = colunas[1].text.strip()
                
                celulas_manuais = linha.find_elements(By.XPATH, ".//td[contains(@title, 'Incluído manualmente')]")
                horas_suspeitas_no_dia = []
                
                for celula in celulas_manuais:
                    titulo = celula.get_attribute("title")
                    if titulo:
                        partes = titulo.split("-")
                        if len(partes) > 1:
                            hora = partes[-1].strip() 
                            if hora.endswith(":00"):
                                horas_suspeitas_no_dia.append(hora)
                
                if horas_suspeitas_no_dia:
                    horarios_britanicos.append(f"{data_registro} -> [{', '.join(horas_suspeitas_no_dia)}]")
                
        if horarios_britanicos:
            texto_log = " | ".join(horarios_britanicos)
            print(f"[{nome}] Horário Britânico ENCONTRADO: {texto_log}")
            
            botao_baixar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Baixar')]]")))
            driver.execute_script("arguments[0].click();", botao_baixar)
            time.sleep(1.5)
            
            botao_pdf = wait.until(EC.element_to_be_clickable((By.ID, "relatorios-baixar-pdf")))
            driver.execute_script("arguments[0].click();", botao_pdf)
            
            movido = mover_pdf_baixado(nome)
            if movido:
                atualizar_excel_tempo_real(nome, "PDF Baixado", texto_log)
            else:
                atualizar_excel_tempo_real(nome, "Erro ao Mover PDF", texto_log)
        else:
            print(f"[{nome}] Nenhum horário britânico manual encontrado.")
            atualizar_excel_tempo_real(nome, "Sem horário britânico")

    except Exception as e:
        print(f"[{nome}] ❌ ERRO: {str(e)}")
        atualizar_excel_tempo_real(nome, "Erro no Script", f"Falha na automação.")

print("\nProcesso concluído!")
driver.quit()
