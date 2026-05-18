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

# ==============================================================================
# 1. DEFINA AQUI OS SEUS LOTES E NOMES
# ==============================================================================
LOTES_DE_NOME = {
    "Lote_01_Equipe_A": ["", ""],
    "Lote_02_Equipe_B": [""]
}

# --- CONFIGURAÇÕES DE PASTAS ---
BASE_DIR = r"C:"
PASTA_AUDITORIA = os.path.join(BASE_DIR, "auditoria relatorios ponto mais")
PASTA_DOWNLOADS_TEMP = os.path.join(PASTA_AUDITORIA, "Downloads_Temp")
CAMINHO_EXCEL = os.path.join(PASTA_AUDITORIA, "Relatorio_Downloads.xlsx")
URL_PAGINA = ""

relatorio = []

# Cria a pasta temporária se não existir
if not os.path.exists(PASTA_DOWNLOADS_TEMP):
    os.makedirs(PASTA_DOWNLOADS_TEMP)


def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": PASTA_DOWNLOADS_TEMP,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)


def atualizar_excel_tempo_real(lote, nome, status, arquivo=""):
    relatorio.append({"Lote": lote, "Colaborador": nome, "Status": status, "Arquivo Salvo": arquivo})
    pd.DataFrame(relatorio).to_excel(CAMINHO_EXCEL, index=False)


def mover_pdf_baixado(nome_colaborador, pasta_destino_lote):
    """Espera o download terminar, move para a pasta do lote e renomeia o arquivo."""
    tempo_espera = 0
    while tempo_espera < 30:
        arquivos = glob.glob(os.path.join(PASTA_DOWNLOADS_TEMP, "*"))
        arquivos_validos = [f for f in arquivos if not f.endswith(".crdownload") and not f.endswith(".tmp")]

        if arquivos_validos:
            # Pega o arquivo mais recente que caiu na pasta temporária
            arquivo_recente = max(arquivos_validos, key=os.path.getctime)

            # Limpa o nome do colaborador para usar no nome do arquivo
            nome_limpo = "".join([c for c in nome_colaborador if c.isalnum() or c == ' ']).strip().replace(' ', '_')
            novo_nome_arquivo = f"Ponto_{nome_limpo}.pdf"
            novo_caminho = os.path.join(pasta_destino_lote, novo_nome_arquivo)

            # Se já existir um arquivo com esse nome, deleta o antigo para substituir
            if os.path.exists(novo_caminho):
                os.remove(novo_caminho)

            shutil.move(arquivo_recente, novo_caminho)
            print(f"      [OK] PDF salvo como: {novo_nome_arquivo}")
            return True, novo_nome_arquivo

        time.sleep(1)
        tempo_espera += 1

    print(f"      [ERRO] Aviso: Falha ou lentidão ao baixar o PDF de {nome_colaborador}.")
    return False, ""


def processar_lotes():
    driver = iniciar_driver()
    wait = WebDriverWait(driver, 15)

    print("Abrindo o navegador...")
    driver.get(URL_PAGINA)

    print("\n" + "=" * 70)
    print("AÇÃO NECESSÁRIA:")
    print("1. Faça o login no site.")
    print("2. Vá até Relatórios > Ajustes de Ponto.")
    print("3. Preencha o Período.")
    print("4. Selecione Filtrar por 'Colaborador'.")
    print("=" * 70 + "\n")

    input(">>> Quando TUDO estiver preenchido e pronto, aperte [ENTER] aqui no terminal para começar <<<")

    print("\nIniciando as buscas em lote...\n")

    for nome_da_pasta, lista_nomes in LOTES_DE_NOME.items():

        # Cria a pasta destino para esse lote específico
        pasta_destino_lote = os.path.join(PASTA_AUDITORIA, nome_da_pasta)
        if not os.path.exists(pasta_destino_lote):
            os.makedirs(pasta_destino_lote)

        print(f"\n{'-' * 50}")
        print(f"ABRINDO COTA: {nome_da_pasta} ({len(lista_nomes)} Colaboradores)")
        print(f"{'-' * 50}")

        contador = 0
        total = len(lista_nomes)

        for nome in lista_nomes:
            contador += 1
            print(f"\n[{contador}/{total}] Buscando: {nome}...")

            try:
                # 1. Limpa o campo de busca (se houver nome da busca anterior)
                try:
                    btn_limpar = driver.find_element(By.XPATH,
                                                     "//div[@data-testid='employee_id-input']//span[@title='Limpar']")
                    driver.execute_script("arguments[0].click();", btn_limpar)
                    time.sleep(1)
                except NoSuchElementException:
                    pass

                # 2. Clica no combobox para abrir a lista
                campo_combobox = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//div[@data-testid='employee_id-input']//div[@role='combobox']")))
                try:
                    campo_combobox.click()
                except:
                    driver.execute_script("arguments[0].click();", campo_combobox)
                time.sleep(1.5)

                # 3. Digita o nome do colaborador
                inputs_busca = driver.find_elements(By.XPATH, "//input[@placeholder='Digite para buscar']")
                input_busca_visivel = next((inp for inp in inputs_busca if inp.is_displayed()), None)

                if not input_busca_visivel:
                    raise Exception("O campo de digitar o nome não apareceu.")

                input_busca_visivel.clear()
                input_busca_visivel.send_keys(nome)

                # 4. Aguarda a lista filtrar e seleciona o nome
                print(f"   -> Aguardando site localizar '{nome}'...")
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
                    raise Exception(f"Correspondência exata não encontrada na lista para '{nome}'.")

                driver.execute_script("arguments[0].click();", opcao_correta)
                time.sleep(1)

                # Aperta ESC para fechar a caixinha suspensa
                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(1)

                # 5. Clica em Gerar Relatório
                print(f"   -> Gerando relatório no sistema...")
                botao_gerar = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Gerar relatório')]]")))
                driver.execute_script("arguments[0].click();", botao_gerar)

                # --- A MÁGICA ESTÁ AQUI: ESPERA INTELIGENTE ---
                print(f"   -> Aguardando o carregamento da página...")
                wait_longo = WebDriverWait(driver, 60)  # Espera até 1 minuto se o site estiver lento

                try:
                    # O robô olha para a tela e espera até que as linhas da tabela apareçam
                    wait_longo.until(EC.presence_of_element_located((By.XPATH, "//table//tbody/tr")))
                except:
                    print("   [Aviso] A tabela demorou a aparecer, tentando prosseguir...")

                time.sleep(1.5)  # Respira só para dar tempo das animações do site terminarem

                # 6. Baixa o arquivo PDF
                print(f"   -> Iniciando o download...")
                botao_baixar = wait_longo.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(), 'Baixar')]]")))
                driver.execute_script("arguments[0].click();", botao_baixar)
                time.sleep(1.5)

                botao_pdf = wait_longo.until(EC.element_to_be_clickable((By.ID, "relatorios-baixar-pdf")))
                driver.execute_script("arguments[0].click();", botao_pdf)

                # 7. Move da pasta Temporária para a pasta do Lote e renomeia
                movido, nome_final = mover_pdf_baixado(nome, pasta_destino_lote)

                if movido:
                    atualizar_excel_tempo_real(nome_da_pasta, nome, "Sucesso", nome_final)
                else:
                    atualizar_excel_tempo_real(nome_da_pasta, nome, "Erro no Download", "Falha ao mover arquivo")

            except Exception as e:
                print(f"   ❌ ERRO: {str(e)}")
                atualizar_excel_tempo_real(nome_da_pasta, nome, "Erro no Script", f"Falha na automação: {str(e)}")

    print("\n" + "=" * 50)
    print("PROCESSO CONCLUÍDO COM SUCESSO!")
    print(f"Arquivos organizados em: {PASTA_AUDITORIA}")
    print("=" * 50)
    driver.quit()


if __name__ == "__main__":
    processar_lotes()
