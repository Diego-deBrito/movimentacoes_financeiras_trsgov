import os
import pandas as pd
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import uuid
import re

# Configuração de logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def esperar_elemento(driver, xpath, timeout=10):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )


def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        logging.info("Conectado ao navegador existente!")
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        logging.error(f"Erro ao conectar ao navegador: {erro}")
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()


def ler_planilha(arquivo):
    """Lê a planilha de entrada mantendo todos os dados originais da aba Document_CH321"""
    try:
        df = pd.read_excel(arquivo, sheet_name='Document_CH321')
        if 'N° Convênio' not in df.columns:
            raise ValueError("Coluna 'N° Convênio' não encontrada na planilha.")
        logging.info("Planilha lida com sucesso! Mantendo todas as colunas originais.")
        print("✅ Planilha lida com sucesso! Mantendo todas as colunas originais.")
        return df
    except Exception as erro:
        logging.error(f"Erro ao ler planilha: {erro}")
        print(f"❌ Erro ao ler planilha: {erro}")
        exit()


def criar_planilha_saida(arquivo_entrada, df, primeiro_instrumento=False):
    """Cria ou atualiza a planilha de saída com os dados originais mais as novas colunas"""
    try:
        pasta = os.path.dirname(arquivo_entrada)
        nome_base = os.path.basename(arquivo_entrada)

        if '.' in nome_base:
            nome_base = nome_base[:nome_base.rfind('.')]

        arquivo_saida = os.path.join(pasta, f"{nome_base}_COM_DATAS.xlsx")

        # Se for o primeiro instrumento, cria novo arquivo
        if primeiro_instrumento:
            contador = 1
            while os.path.exists(arquivo_saida):
                arquivo_saida = os.path.join(pasta, f"{nome_base}_COM_DATAS_{contador}.xlsx")
                contador += 1

            # Cria nova planilha com a aba Document_CH321
            with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Document_CH321', index=False)
        else:
            # Atualiza o arquivo existente
            with pd.ExcelWriter(arquivo_saida, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='Document_CH321', index=False)

        logging.info(f"Planilha de saída {'criada' if primeiro_instrumento else 'atualizada'}: {arquivo_saida}")
        print(f"✅ Planilha de saída {'criada' if primeiro_instrumento else 'atualizada'}: {arquivo_saida}")
        return arquivo_saida
    except Exception as e:
        logging.error(f"Erro ao {'criar' if primeiro_instrumento else 'atualizar'} planilha de saída: {e}")
        print(f"❌ Erro ao {'criar' if primeiro_instrumento else 'atualizar'} planilha de saída: {e}")
        return None


def navegar_menu_principal(driver, instrumento):
    try:
        esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
        esperar_elemento(driver, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()
        campo_pesquisa = esperar_elemento(driver,
                                          "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        esperar_elemento(driver,
                         "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]").click()
        time.sleep(1)
        esperar_elemento(driver,
                         "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]").click()
        return True
    except:
        print(f"⚠️ Instrumento {instrumento} não encontrado.")
        return False


def acessar_aba_movimentacao_financeira(driver):
    """Acessa a aba de Movimentação Financeira com múltiplos caminhos alternativos"""
    try:
        # Primeiro clique - Tenta dois caminhos diferentes para o menu principal
        try:
            # Tentativa 1: Caminho XPath original
            esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[6]/div/span/span").click()
        except Exception as e:
            # Tentativa 2: Novo JS Path como fallback
            driver.execute_script('document.querySelector("#div_-481524888 > span > span").click()')

        time.sleep(1)

        # Segundo clique - Movimentação Financeira
        try:
            # Tentativa 1: Execução via JavaScript
            driver.execute_script('document.querySelector("#menu_link_-481524888_1304359359 > div > span").click()')
        except Exception as e:
            # Tentativa 2: Caminho XPath alternativo
            esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[26]/div/span/span").click()

        time.sleep(1)
        return True

    except Exception as erro:
        logging.error(f"Erro ao acessar aba Movimentação Financeira: {erro}")
        print(f"❌ Erro ao acessar aba Movimentação Financeira: {erro}")
        return False


def coletar_data_recente(driver):
    """Coleta a data mais recente da tabela de movimentação financeira"""
    try:
        tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[16]/div[2]/form/table"))
        )

        rows = tabela.find_elements(By.TAG_NAME, "tr")
        datas = []

        for row in rows[1:]:  # Pular cabeçalho se existir
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 2:  # Verifica se há pelo menos 2 colunas
                    data_text = cells[1].text.strip()
                    if re.match(r"\d{2}/\d{2}/\d{4}", data_text):
                        data = datetime.strptime(data_text, "%d/%m/%Y")
                        datas.append(data)
            except:
                continue

        if datas:
            data_recente = max(datas).strftime("%d/%m/%Y")
            return data_recente, "SIM"  # Agora retorna apenas a data e "SIM"
        else:
            return None, "NÃO"  # Retorna None e "NÃO" quando não há dados

    except:
        return None, "NÃO"  # Retorna None e "NÃO" quando não encontra a tabela


def navegar_voltar_inicio(driver):
    """Navega de volta para a página inicial entre cada proposta"""
    try:
        driver.get("https://discricionarias.transferegov.sistema.gov.br/voluntarias/Principal/Principal.do")
        time.sleep(2)  # Espera carregar a página
        print("✅ Voltou para a página inicial com sucesso")
        return True
    except Exception as e:
        print(f"❌ Erro ao voltar para página inicial: {e}")
        return False


def main():
    # Configuração inicial
    arquivo_entrada = r"C:\Users\diego.brito\Downloads\robov1\Movimentação Financeira\20250506 - Parcerias em Execução com Desembolso Acumulado.xlsx"

    # Ler planilha mantendo todos os dados originais
    df = ler_planilha(arquivo_entrada)
    total_instrumentos = len(df)

    # Adicionar colunas novas
    df['Data Mais Recente'] = pd.NA
    df['Movimentação'] = pd.NA

    driver = conectar_navegador_existente()
    arquivo_saida = None

    # Variáveis para cálculo de tempo
    tempos = []
    inicio_geral = time.time()

    for index, row in df.iterrows():
        instrumento = row['N° Convênio']
        inicio_instrumento = time.time()

        # Cálculo do progresso
        progresso = (index + 1) / total_instrumentos * 100
        restantes = total_instrumentos - (index + 1)

        print("\n" + "=" * 60)
        print(f"🚀 PROCESSANDO INSTRUMENTO {index + 1}/{total_instrumentos}")
        print(f"📌 Convênio atual: {instrumento}")
        print(f"📊 Progresso: {progresso:.1f}% concluído")
        print(f"🕑 Instrumentos restantes: {restantes}")

        if tempos:
            tempo_medio = sum(tempos) / len(tempos)
            tempo_estimado = tempo_medio * restantes
            horas, rem = divmod(tempo_estimado, 3600)
            minutos, segundos = divmod(rem, 60)
            print(f"⏱ Tempo estimado restante: {int(horas):02d}h {int(minutos):02d}m {int(segundos):02d}s")

        print("=" * 60 + "\n")

        if navegar_menu_principal(driver, instrumento):
            if acessar_aba_movimentacao_financeira(driver):
                data_recente, possui_data = coletar_data_recente(driver)
            else:
                data_recente, possui_data = None, "NÃO"
        else:
            data_recente, possui_data = None, "NÃO"

        # Atualizar DataFrame
        df.at[index, 'Data Mais Recente'] = data_recente
        df.at[index, 'Movimentação'] = possui_data

        # Calcular tempo deste instrumento
        tempo_instrumento = time.time() - inicio_instrumento
        tempos.append(tempo_instrumento)
        print(f"⏳ Tempo deste instrumento: {tempo_instrumento:.1f} segundos")

        # Criar/atualizar planilha de saída após o primeiro instrumento
        if index == 0:
            arquivo_saida = criar_planilha_saida(arquivo_entrada, df, primeiro_instrumento=True)
        elif (index + 1) % 5 == 0:  # Salvar a cada 5 instrumentos
            arquivo_saida = criar_planilha_saida(arquivo_entrada, df)

        # Voltar para página inicial antes do próximo instrumento
        if index < total_instrumentos - 1:
            if not navegar_voltar_inicio(driver):
                print("⚠️ Não conseguiu voltar para página inicial, tentando recarregar...")
                driver.refresh()
                time.sleep(3)

    # Cálculo do tempo total
    tempo_total = time.time() - inicio_geral
    horas, rem = divmod(tempo_total, 3600)
    minutos, segundos = divmod(rem, 60)

    print("\n" + "=" * 60)
    print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
    print(f"⏱ Tempo total: {int(horas):02d}h {int(minutos):02d}m {int(segundos):02d}s")
    print(f"📝 Arquivo gerado: {arquivo_saida}")
    print("=" * 60)

    driver.quit()
if __name__ == "__main__":
    main()