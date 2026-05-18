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
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

# --- CONFIGURAÇÕES DE PASTAS ---
BASE_DIR = r"C:\Users\710399\Desktop\Python codes\auditoria ponto mais"
PASTA_AUDITORIA = os.path.join(BASE_DIR, "auditoria relatorios ponto mais")
PASTA_DOWNLOADS_TEMP = os.path.join(PASTA_AUDITORIA, "Downloads_Temp")

# Cria pastas base se não existirem
os.makedirs(PASTA_AUDITORIA, exist_ok=True)
os.makedirs(PASTA_DOWNLOADS_TEMP, exist_ok=True)

URL_PAGINA = "https://atma2.pontomais.com.br/relatorios" 

# --- NOVA ESTRUTURA: Separação por Hierarquia e Nomes ---
DADOS_POR_HIERARQUIA = {
    "Cibely Mazalla Gibulo": [
        "Victor Hugo Barboza Ramos",
        "Gabriela Rodrigues da Silva",
        "Talita Romeika Canete",
        "Mariana Silva Massote Campos"
    ],
    "Giancarlo Tardin Santos": [
        "Maria Altinizia Santos Santana",
        "Claudia Gabriela D Almeida Figueiredo",
        "Edilson de Santana Gomes",
        "Gleide Sales dos Santos",
        "Diego Mota Macario",
        "Carlos Alberto Amaral dos Santos"    
    ],
    "Roberto Costa Bento Bonini": [
        "Lucas Oliveira dos Anjos",
        "Carolina Coelho Ribeiro",
        "Paula Cancherini Sevo",
        "Cibele Alves Siqueira",
        "Samara dos Santos Rodrigues Lemos",
        "Fernanda Talharo Ikeda",
        "Alex Souza de Alcantara Nonato"    
    ]
}

# --- CONFIGURAÇÃO DO NAVEGADOR ---
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

prefs = {
    "download.default_directory": PASTA_DOWNLOADS_TEMP, # Todos os downloads vão pra cá primeiro
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True 
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 15)

# --- FUNÇÕES AUXILIARES ---
def atualizar_excel_tempo_real(dados_lista, caminho_arquivo, hierarquia, nome, status, detalhes="N/A"):
    """Salva o Excel isolado da hierarquia em tempo real, com proteção contra arquivo aberto."""
    dados_lista.append({
        "Hierarquia": hierarquia, 
        "Colaborador": nome, 
        "Status": status, 
        "Dias e Horários Suspeitos": detalhes
    })
    
    try:
        pd.DataFrame(dados_lista).to_excel(caminho_arquivo, index=False)
    except PermissionError:
        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f"\n⚠️ AVISO: O arquivo '{nome_arquivo}' está ABERTO! O Windows bloqueou o salvamento.")
        print("-> Salvando no arquivo de BACKUP temporário...")
        
        caminho_backup = caminho_arquivo.replace(".xlsx", "_BACKUP.xlsx")
        try:
            pd.DataFrame(dados_lista).to_excel(caminho_backup, index=False)
        except:
            print("-> Falha também ao salvar o backup.")
    except Exception as e:
        print(f"\n⚠️ Erro inesperado ao salvar o Excel: {e}")

def mover_pdf_baixado(nome_colaborador, pasta_destino):
    """Pega o PDF da pasta temporária, renomeia com o nome do colaborador e move para a hierarquia."""
    tempo_espera = 0
    while tempo_espera < 30:
        arquivos = glob.glob(os.path.join(PASTA_DOWNLOADS_TEMP, "*"))
        arquivos_validos = [f for f in arquivos if not f.endswith(".crdownload")]
        
        if arquivos_validos:
            # Pega o arquivo mais recente
            arquivo_recente = max(arquivos_validos, key=os.path.getctime)
            nome_arquivo_original = os.path.basename(arquivo_recente)
            
            # Formata o novo nome para incluir o colaborador (evita sobrescrever arquivos)
            nome_limpo = nome_colaborador.replace(" ", "_")
            novo_nome = f"{nome_limpo}_{nome_arquivo_original}"
            novo_caminho = os.path.join(pasta_destino, novo_nome)
            
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)
                
            shutil.move(arquivo_recente, novo_caminho)
            print(f"[{nome_colaborador}] PDF salvo como: {novo_nome}")
            return True
            
        time.sleep(1)
        tempo_espera += 1
        
    print(f"[{nome_colaborador}] Aviso: Falha ou lentidão ao baixar o PDF.")
    return False

# --- FLUXO PRINCIPAL ---
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

# Loop 1: Passa por cada Hierarquia
for hierarquia, nomes_colaboradores in DADOS_POR_HIERARQUIA.items():
    print(f"\n{'='*40}")
    print(f" INICIANDO HIERARQUIA: {hierarquia}")
    print(f"{'='*40}")

    # Cria a estrutura de pastas da hierarquia atual
    pasta_hierarquia = os.path.join(PASTA_AUDITORIA, hierarquia)
    pasta_documentos_hierarquia = os.path.join(pasta_hierarquia, "Documentos_Baixados")
    os.makedirs(pasta_documentos_hierarquia, exist_ok=True)
    
    # Define o caminho do Excel desta hierarquia
    caminho_excel_hierarquia = os.path.join(pasta_hierarquia, f"Relatorio_Auditoria_Britanico_{hierarquia}.xlsx")
    relatorio_local = [] # Zera a lista para esta hierarquia

    # Loop 2: Passa por cada Colaborador
    for nome in nomes_colaboradores:
        print(f"\n[{nome}] Iniciando...")

        try:
            # Limpa a busca anterior
            try:
                btn_limpar = driver.find_element(By.XPATH, "//div[@data-testid='employee_id-input']//span[@title='Limpar']")
                driver.execute_script("arguments[0].click();", btn_limpar)
                time.sleep(1)
            except NoSuchElementException:
                pass 

            # Clica no campo de busca
            campo_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-testid='employee_id-input']//div[@role='combobox']")))
            try:
                campo_combobox.click()
            except:
                driver.execute_script("arguments[0].click();", campo_combobox)
            time.sleep(1.5) 

            # Digita o nome
            inputs_busca = driver.find_elements(By.XPATH, "//input[@placeholder='Digite para buscar']")
            input_busca_visivel = next((inp for inp in inputs_busca if inp.is_displayed()), None)
            
            if not input_busca_visivel:
                raise Exception("O campo de digitar o nome não apareceu.")

            input_busca_visivel.clear()
            input_busca_visivel.send_keys(nome)
            
            print(f"[{nome}] Aguardando a API retornar o nome correto na lista...")
            tempo_limite = time.time() + 15 
            opcao_correta = None
            
            # Encontra a opção correta na lista suspensa
            while time.time() < tempo_limite:
                try:
                    opcoes_lista = driver.find_elements(By.CSS_SELECTOR, "div.ng-option")
                    for opcao in opcoes_lista:
                        texto_opcao = opcao.text.strip().upper() 
                        if nome.upper() in texto_opcao:
                            opcao_correta = opcao
                            break 
                except StaleElementReferenceException:
                    # Se a lista atualizou enquanto o robô lia, ele pausa e tenta de novo
                    time.sleep(0.5)
                    continue
                
                if opcao_correta: break 
                time.sleep(0.5) 
                
            if not opcao_correta:
                raise Exception(f"A pesquisa não encontrou correspondência exata para '{nome}'.")

            driver.execute_script("arguments[0].click();", opcao_correta) 
            time.sleep(1)

            # Fecha a lista suspensa
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
            
            print(f"[{nome}] Gerando relatório...")
            botao_gerar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Gerar relatório')]]")))
            driver.execute_script("arguments[0].click();", botao_gerar)
            time.sleep(6)
            
            print(f"[{nome}] Analisando tabela de horários...")
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, "//table//tbody/tr")))
                time.sleep(2) # Dá um fôlego extra para a tabela renderizar completamente
            except: pass
                
            # Verifica o Horário Britânico (COM PROTEÇÃO STALE ELEMENT E LOOP POR ÍNDICE)
            horarios_britanicos = []
            tentativas_tabela = 0
            
            while tentativas_tabela < 3:
                try:
                    linhas_tabela = driver.find_elements(By.XPATH, "//table//tbody/tr")
                    quantidade_linhas = len(linhas_tabela)
                    
                    # Usar range(len) evita que a lista inteira fique obsoleta (Stale) de uma vez
                    for i in range(quantidade_linhas):
                        # Puxa a linha específica de novo a cada repetição
                        linha_atual = driver.find_elements(By.XPATH, "//table//tbody/tr")[i]
                        colunas = linha_atual.find_elements(By.TAG_NAME, "td")
                        
                        if len(colunas) > 5: 
                            data_registro = colunas[1].text.strip()
                            celulas_manuais = linha_atual.find_elements(By.XPATH, ".//td[contains(@title, 'Incluído manualmente')]")
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
                    
                    # Se leu todas as linhas sem erro, quebra o while
                    break
                    
                except StaleElementReferenceException:
                    print(f"   -> A tabela atualizou durante a leitura. Tentando novamente ({tentativas_tabela+1}/3)...")
                    time.sleep(2)
                    tentativas_tabela += 1
                    horarios_britanicos = [] # Zera os horários para não duplicar na próxima tentativa
                    
            if horarios_britanicos:
                texto_log = " | ".join(horarios_britanicos)
                print(f"[{nome}] Horário Britânico ENCONTRADO: {texto_log}")
                
                # Baixa o PDF
                botao_baixar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Baixar')]]")))
                driver.execute_script("arguments[0].click();", botao_baixar)
                time.sleep(1.5)
                
                botao_pdf = wait.until(EC.element_to_be_clickable((By.ID, "relatorios-baixar-pdf")))
                driver.execute_script("arguments[0].click();", botao_pdf)
                
                # Move para a pasta da hierarquia atual
                movido = mover_pdf_baixado(nome, pasta_documentos_hierarquia)
                
                if movido:
                    atualizar_excel_tempo_real(relatorio_local, caminho_excel_hierarquia, hierarquia, nome, "PDF Baixado", texto_log)
                else:
                    atualizar_excel_tempo_real(relatorio_local, caminho_excel_hierarquia, hierarquia, nome, "Erro ao Mover PDF", texto_log)
            else:
                print(f"[{nome}] Nenhum horário britânico manual encontrado.")
                atualizar_excel_tempo_real(relatorio_local, caminho_excel_hierarquia, hierarquia, nome, "Sem horário britânico")

        except Exception as e:
            print(f"[{nome}] ❌ ERRO: {str(e)}")
            atualizar_excel_tempo_real(relatorio_local, caminho_excel_hierarquia, hierarquia, nome, "Erro no Script", f"Falha na automação.")

print("\nProcesso concluído com sucesso!")
driver.quit()