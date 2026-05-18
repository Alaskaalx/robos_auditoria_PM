import os
import time
import glob
import shutil
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

# --- CONFIGURAÇÕES DE PASTAS ---
BASE_DIR = r""
PASTA_AUDITORIA = os.path.join(BASE_DIR, "auditoria relatorios ponto mais")
PASTA_DOWNLOADS_TEMP = os.path.join(PASTA_AUDITORIA, "Downloads_Temp")

os.makedirs(PASTA_AUDITORIA, exist_ok=True)
os.makedirs(PASTA_DOWNLOADS_TEMP, exist_ok=True)

URL_PAGINA = "https://" 

# --- ESTRUTURA: Separação por Hierarquia e Nomes ---
DADOS_POR_HIERARQUIA = {
    "equipe A": [
    ],
    "equipe B": [
        "",
        "",
        "",
        "",
        "",
        ""    
    ],
    "equipe C": [
        "",
        "",
        "",
        "",
        "",
        "",
        ""    
    ]
}

# --- CONFIGURAÇÃO DO NAVEGADOR ---
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.page_load_strategy = 'eager' # Deixa o robô mais rápido

prefs = {
    "download.default_directory": PASTA_DOWNLOADS_TEMP, 
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True 
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 15)

# --- FUNÇÕES DE CÁLCULO DE HORAS ---
def str_to_mins(time_str):
    """Converte 'HH:MM' para minutos totais."""
    if not time_str or ":" not in time_str:
        return 0
    h, m = map(int, time_str.split(':'))
    return h * 60 + m

def mins_to_str(mins):
    """Converte minutos totais para 'HH:MM'."""
    mins = abs(mins)
    return f"{mins//60:02d}:{mins%60:02d}"

def calcular_duracao_blocos(lista_tempos):
    """Soma o tempo trabalhado: (T2-T1) + (T4-T3) + ..."""
    total = 0
    for i in range(0, len(lista_tempos)-1, 2):
        total += (str_to_mins(lista_tempos[i+1]) - str_to_mins(lista_tempos[i]))
    return total

# --- FUNÇÕES AUXILIARES DE ARQUIVO ---
def atualizar_excel_tempo_real(dados_lista, caminho_arquivo):
    """Salva o Excel isolado da hierarquia em tempo real, com proteção contra arquivo aberto."""
    try:
        df = pd.DataFrame(dados_lista)
        colunas_ordem = ["Hierarquia", "Colaborador", "Data", "Jornada Planejada", "Pontos Realizados", 
                         "Status do Dia", "Tempo Faltante", "Horas Extras", "Pausas Identificadas", "Arquivo PDF"]
        df = df[[c for c in colunas_ordem if c in df.columns]]
        df.to_excel(caminho_arquivo, index=False)
    except PermissionError:
        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f"\n⚠️ AVISO: O arquivo '{nome_arquivo}' está ABERTO! Salvando no BACKUP...")
        try:
            df.to_excel(caminho_arquivo.replace(".xlsx", "_BACKUP.xlsx"), index=False)
        except: pass

def mover_pdf_baixado(nome_colaborador, pasta_destino):
    """Pega o PDF da pasta temporária, renomeia com o nome do colaborador e move para a hierarquia."""
    tempo_espera = 0
    while tempo_espera < 30:
        arquivos = glob.glob(os.path.join(PASTA_DOWNLOADS_TEMP, "*"))
        arquivos_validos = [f for f in arquivos if not f.endswith(".crdownload")]
        if arquivos_validos:
            arquivo_recente = max(arquivos_validos, key=os.path.getctime)
            nome_arquivo_original = os.path.basename(arquivo_recente)
            nome_limpo = nome_colaborador.replace(" ", "_")
            novo_nome = f"{nome_limpo}_{nome_arquivo_original}"
            novo_caminho = os.path.join(pasta_destino, novo_nome)
            if os.path.exists(novo_caminho): os.remove(novo_caminho)
            shutil.move(arquivo_recente, novo_caminho)
            return novo_nome
        time.sleep(1)
        tempo_espera += 1
    return None

# --- FLUXO PRINCIPAL ---
print("Abrindo o navegador...")
driver.get(URL_PAGINA) 

print("\n" + "="*70)
print("AÇÃO NECESSÁRIA:")
print("1. Faça o login e vá até Relatórios > Espelho de Ponto (Jornada).")
print("2. Preencha o Período e selecione Filtrar por 'Colaborador'.")
print("="*70 + "\n")
input(">>> Quando TUDO estiver preenchido, aperte [ENTER] aqui no terminal para começar <<<")

print("\nIniciando extração e cálculo de horas...\n")

for hierarquia, nomes_colaboradores in DADOS_POR_HIERARQUIA.items():
    print(f"\n{'='*40}\n INICIANDO HIERARQUIA: {hierarquia}\n{'='*40}")

    pasta_hierarquia = os.path.join(PASTA_AUDITORIA, hierarquia)
    pasta_documentos = os.path.join(pasta_hierarquia, "Documentos_Baixados")
    os.makedirs(pasta_documentos, exist_ok=True)
    
    caminho_excel = os.path.join(pasta_hierarquia, f"Relatorio_Auditoria_Jornada_{hierarquia}.xlsx")
    relatorio_local = [] 

    for nome in nomes_colaboradores:
        print(f"\n[{nome}] Iniciando...")

        try:
            # --- BUSCA DE COLABORADOR ROBUSTA (ANGULAR/JAVASCRIPT) ---
            try:
                btn_limpar = driver.find_element(By.XPATH, "//div[@data-testid='employee_id-input']//span[@title='Limpar']")
                driver.execute_script("arguments[0].click();", btn_limpar)
                time.sleep(1)
            except NoSuchElementException: pass 

            campo_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-testid='employee_id-input']//div[@role='combobox']")))
            driver.execute_script("arguments[0].click();", campo_combobox)
            time.sleep(1.5) 

            inputs_busca = driver.find_elements(By.XPATH, "//input[@placeholder='Digite para buscar']")
            input_busca_visivel = next((inp for inp in inputs_busca if inp.is_displayed()), None)
            if not input_busca_visivel: 
                raise Exception("O campo de digitar o nome não apareceu.")

            input_busca_visivel.clear()
            input_busca_visivel.send_keys(nome)
            
            print(f"[{nome}] Aguardando a API do Angular retornar a lista...")
            tempo_limite = time.time() + 15 
            opcao_correta = None
            
            while time.time() < tempo_limite:
                try:
                    opcoes_lista = driver.find_elements(By.CSS_SELECTOR, "div.ng-option")
                    for opcao in opcoes_lista:
                        texto_opcao = driver.execute_script("return arguments[0].innerText;", opcao)
                        if texto_opcao and nome.upper() in texto_opcao.strip().upper():
                            opcao_correta = opcao
                            break 
                except StaleElementReferenceException:
                    time.sleep(0.5)
                    continue
                    
                if opcao_correta: break 
                time.sleep(0.5) 
                
            if not opcao_correta: 
                raise Exception(f"Nome '{nome}' não encontrado na lista suspensa.")

            driver.execute_script("arguments[0].click();", opcao_correta) 
            time.sleep(1.5)
            
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
            # --- FIM DA BUSCA ---
            
            # Clica em Gerar Relatório
            botao_gerar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Gerar relatório')]]")))
            driver.execute_script("arguments[0].click();", botao_gerar)
            time.sleep(6)
            
            print(f"[{nome}] Extraindo e calculando blocos de horas...")
            try: wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody/tr"))); time.sleep(2)
            except: pass
                
            dias_processados_colaborador = []
            tentativas_tabela = 0
            
            # --- LEITURA DA TABELA DE ESPELHO DE PONTO ---
            while tentativas_tabela < 3:
                try:
                    linhas_tabela = driver.find_elements(By.XPATH, "//table//tbody/tr")
                    for i in range(len(linhas_tabela)):
                        linha_atual = driver.find_elements(By.XPATH, "//table//tbody/tr")[i]
                        colunas = linha_atual.find_elements(By.TAG_NAME, "td")
                        
                        if len(colunas) >= 20: 
                            data_dia = colunas[0].text.strip()
                            
                            tempos_jornada = [colunas[j].text.strip() for j in range(1, 5) if colunas[j].text.strip()]
                            tempos_pontos = [colunas[j].text.strip() for j in range(5, 9) if colunas[j].text.strip()]
                            
                            inicio_p1, fim_p1 = colunas[9].text.strip(), colunas[10].text.strip()
                            inicio_p2, fim_p2 = colunas[11].text.strip(), colunas[12].text.strip()
                            
                            lista_pausas = []
                            if inicio_p1 and fim_p1:
                                duracao_p1 = str_to_mins(fim_p1) - str_to_mins(inicio_p1)
                                lista_pausas.append(f"Pausa 1: {mins_to_str(duracao_p1)}")
                            if inicio_p2 and fim_p2:
                                duracao_p2 = str_to_mins(fim_p2) - str_to_mins(inicio_p2)
                                lista_pausas.append(f"Pausa 2: {mins_to_str(duracao_p2)}")
                                
                            texto_pausas = " | ".join(lista_pausas) if lista_pausas else "Sem Pausas/Registros"
                            
                            faltante = colunas[14].text.strip() or "00:00"
                            he_50 = colunas[16].text.split('\n')[0].strip() or "00:00"
                            he_100 = colunas[17].text.split('\n')[0].strip() or "00:00"
                            motivo = colunas[20].text.strip()
                            
                            status_dia = "Normal"
                            if motivo:
                                status_dia = f"Afastamento/Motivo: {motivo}"
                            elif len(tempos_pontos) % 2 != 0:
                                status_dia = "Inconsistente (Marcação Ímpar)"
                            elif faltante != "00:00":
                                status_dia = "Déficit de Horas / Atraso"
                            elif he_50 != "00:00" or he_100 != "00:00":
                                status_dia = "Horas Extras Realizadas"
                            
                            dias_processados_colaborador.append({
                                "Hierarquia": hierarquia,
                                "Colaborador": nome,
                                "Data": data_dia,
                                "Jornada Planejada": " - ".join(tempos_jornada) if tempos_jornada else "Sem Jornada",
                                "Pontos Realizados": " - ".join(tempos_pontos) if tempos_pontos else "Sem Pontos",
                                "Status do Dia": status_dia,
                                "Tempo Faltante": faltante,
                                "Horas Extras": f"50%: {he_50} | 100%: {he_100}",
                                "Pausas Identificadas": texto_pausas,
                                "Arquivo PDF": "Pendente"
                            })
                    break 
                    
                except StaleElementReferenceException:
                    time.sleep(2)
                    tentativas_tabela += 1
                    dias_processados_colaborador = [] 
                    
            # --- DOWNLOAD DO PDF ---
            print(f"[{nome}] Baixando PDF...")
            botao_baixar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Baixar')]]")))
            driver.execute_script("arguments[0].click();", botao_baixar)
            time.sleep(1.5)
            
            botao_pdf = wait.until(EC.element_to_be_clickable((By.ID, "relatorios-baixar-pdf")))
            driver.execute_script("arguments[0].click();", botao_pdf)
            
            nome_pdf = mover_pdf_baixado(nome, pasta_documentos)
            
            if nome_pdf:
                print(f"[{nome}] ✅ PDF salvo e {len(dias_processados_colaborador)} dias analisados.")
            else:
                print(f"[{nome}] ⚠️ Dados analisados, mas falha ao baixar PDF.")
                
            for dia in dias_processados_colaborador:
                dia["Arquivo PDF"] = nome_pdf if nome_pdf else "Falha no Download"
                relatorio_local.append(dia)
                
            atualizar_excel_tempo_real(relatorio_local, caminho_excel)

        except Exception as e:
            print(f"[{nome}] ❌ ERRO: {str(e)}")
            relatorio_local.append({"Hierarquia": hierarquia, "Colaborador": nome, "Status do Dia": "Erro no Script", "Tempo Faltante": str(e)})
            atualizar_excel_tempo_real(relatorio_local, caminho_excel)

print("\nAuditoria matemática concluída com sucesso!")
driver.quit()
