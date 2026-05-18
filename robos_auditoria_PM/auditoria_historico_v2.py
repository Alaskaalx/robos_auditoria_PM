import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- CONFIGURAÇÕES DE PASTAS ---
BASE_DIR = r""
PASTA_AUDITORIA = os.path.join(BASE_DIR, "auditoria LOG")
os.makedirs(PASTA_AUDITORIA, exist_ok=True)

URL_HISTORICO = ""

# --- NOVA ESTRUTURA: Separação por Hierarquia e Nomes ---
DADOS_POR_HIERARQUIA = {
    "Operacoes_Norte": [
        "", 
        ""
    ],
    "Operacoes_Sul": [
        "", 
        ""
    ]
}

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 15)

def aguardar_e_clicar(by, seletor):
    elemento = wait.until(EC.element_to_be_clickable((by, seletor)))
    elemento.click()
    time.sleep(1.5)

def aguardar_e_preencher(by, seletor, texto):
    elemento = wait.until(EC.presence_of_element_located((by, seletor)))
    elemento.clear()
    elemento.send_keys(texto)
    time.sleep(1)

def capturar_print(caminho_arquivo):
    driver.save_screenshot(caminho_arquivo)
    print(f"Print salvo: {caminho_arquivo}")

# --- LOGIN MANUAL (Feito apenas 1 vez) ---
print("Abrindo o navegador...")
driver.get(URL_HISTORICO) 

print("\n" + "="*50)
print("AÇÃO NECESSÁRIA: Por favor, faça o login no navegador.")
print(f"O script está pausado e aguardando você chegar na página:")
print(f"-> {URL_HISTORICO}")
print("="*50 + "\n")

try:
    wait_login = WebDriverWait(driver, 600)
    wait_login.until(EC.url_to_be(URL_HISTORICO))
    print("Página de histórico detectada! Iniciando a automação em 5 segundos...")
    time.sleep(5) 
except TimeoutException:
    print("Tempo limite de login excedido. Encerrando o script.")
    driver.quit()
    exit()

# --- PROCESSAMENTO POR HIERARQUIA ---
for hierarquia, nomes_colaboradores in DADOS_POR_HIERARQUIA.items():
    print(f"\n{'='*40}")
    print(f" INICIANDO HIERARQUIA: {hierarquia}")
    print(f"{'='*40}")

    # Cria as pastas específicas desta hierarquia
    pasta_hierarquia = os.path.join(PASTA_AUDITORIA, hierarquia)
    pasta_sucesso = os.path.join(pasta_hierarquia, "Sucesso")
    pasta_erros = os.path.join(pasta_hierarquia, "Erros")
    
    for pasta in [pasta_hierarquia, pasta_sucesso, pasta_erros]:
        os.makedirs(pasta, exist_ok=True)

    relatorio_local = [] # Relatório zerado para esta hierarquia

    # --- PROCESSAMENTO POR NOME (Dentro da Hierarquia) ---
    for nome in nomes_colaboradores:
        print(f"\nIniciando busca para: {nome}")
        driver.get(URL_HISTORICO)
        time.sleep(3) 

        try:
            # Pesquisa o nome
            aguardar_e_preencher(By.CSS_SELECTOR, 'input[data-testid="search_text-input"]', nome)
            aguardar_e_clicar(By.CSS_SELECTOR, 'button.search-button')
            time.sleep(3) 
            
            # Cria a pasta com o NOME do colaborador dentro da pasta de Sucesso da Hierarquia
            pasta_colab = os.path.join(pasta_sucesso, nome)
            os.makedirs(pasta_colab, exist_ok=True)
                
            capturar_print(os.path.join(pasta_colab, f"1_Pesquisa_{nome}.png"))

            # Tenta encontrar e clicar no ajuste de ponto
            try:
                xpath_botao_ajuste = "//p[contains(text(), 'Ajuste de ponto')]/ancestor::tr//button[contains(@class, 'pm-btn-icon')]"
                botao_olho = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_ajuste)))
                botao_olho.click()
                time.sleep(3)
                
                capturar_print(os.path.join(pasta_colab, f"2_Detalhes_Ajuste_{nome}.png"))
                
                aguardar_e_clicar(By.XPATH, '//a[@href="/historico" and contains(text(), "Histórico")]')
                time.sleep(3)
                
            except TimeoutException:
                print(f"'Ajuste de ponto' não encontrado para {nome} na primeira página.")

            # Verifica paginação e equivalência
            encontrou_equivalencia = False
            paginas_verificadas = 0
            limite_paginas = 10 
            
            while not encontrou_equivalencia and paginas_verificadas < limite_paginas:
                try:
                    driver.find_element(By.XPATH, "//p[contains(text(), 'Colaboradores (turno)') or contains(text(), 'Gestor de equipe')]")
                    encontrou_equivalencia = True
                    capturar_print(os.path.join(pasta_colab, f"3_Equivalencia_Encontrada_{nome}_Pag_{paginas_verificadas+1}.png"))
                    print(f"Equivalência encontrada para {nome}.")
                    break
                except NoSuchElementException:
                    pass
                    
                try:
                    btn_proxima = driver.find_element(By.CSS_SELECTOR, "div.dx-next-button")
                    
                    if "dx-state-disabled" in btn_proxima.get_attribute("class"):
                        print("Chegou na última página.")
                        break
                    
                    btn_proxima.click()
                    time.sleep(3)
                    paginas_verificadas += 1
                except NoSuchElementException:
                    print("Botão de próxima página não encontrado.")
                    break

            relatorio_local.append({"Hierarquia": hierarquia, "Colaborador": nome, "Status": "Sucesso", "Observação": "Processado com sucesso."})

        except Exception as e:
            print(f"Erro ao processar {nome}: {str(e)}")
            caminho_erro = os.path.join(pasta_erros, f"Erro_{nome}.png")
            driver.save_screenshot(caminho_erro)
            print(f"Print de erro salvo: {caminho_erro}")
            relatorio_local.append({"Hierarquia": hierarquia, "Colaborador": nome, "Status": "Erro", "Observação": "Falha durante o processamento."})

    # --- SALVA O EXCEL DA HIERARQUIA ---
    df_relatorio = pd.DataFrame(relatorio_local)
    caminho_excel = os.path.join(pasta_hierarquia, f"Relatorio_Auditoria_{hierarquia}.xlsx")
    df_relatorio.to_excel(caminho_excel, index=False)
    print(f"\nRelatório da hierarquia {hierarquia} salvo em: {caminho_excel}")

driver.quit()
print("\nTODAS AS HIERARQUIAS FORAM PROCESSADAS COM SUCESSO!")
