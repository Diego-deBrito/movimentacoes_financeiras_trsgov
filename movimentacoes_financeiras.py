# -*- coding: utf-8 -*-
"""
Este script é um robô de automação (RPA) que enriquece uma planilha Excel
com informações de data coletadas do portal TransfereGov.

O robô lê uma lista de números de convênio de uma planilha, navega até a
página de cada convênio no portal, acessa a aba de "Movimentação Financeira",
coleta a data da transação mais recente e, por fim, adiciona essa data
e um status de movimentação de volta à planilha, salvando o resultado em um
novo arquivo.
"""

import os
import re
import time
import logging
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, WebDriverException

# --- Configurações Globais ---

# Caminho para a planilha de entrada. Altere este valor conforme necessário.
INPUT_EXCEL_PATH = r"C:\Users\diego.brito\Downloads\robov1\Movimentação Financeira\20250506 - Parcerias em Execução com Desembolso Acumulado.xlsx"

# Configuração do sistema de logging para registrar eventos em um arquivo e no console.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("robo_log.txt"), # Salva os logs em um arquivo
        logging.StreamHandler() # Mostra os logs no console
    ]
)


def conectar_navegador_existente():
    """
    Conecta-se a uma instância do Google Chrome já em execução em modo de depuração.

    Pré-requisito: O Chrome deve ser iniciado com a flag de depuração remota, por exemplo:
    `"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222`

    Returns:
        webdriver.Chrome: Instância do driver do Selenium conectada.
        Encerra o script em caso de falha.
    """
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        logging.info("Conexão com o navegador existente bem-sucedida.")
        return driver
    except WebDriverException as e:
        logging.error(f"Não foi possível conectar ao navegador. Verifique se ele está em execução na porta 9222. Erro: {e}")
        exit()
    except Exception as e:
        logging.error(f"Ocorreu um erro inesperado ao conectar ao navegador: {e}")
        exit()


def ler_planilha(caminho_arquivo):
    """
    Lê a aba 'Document_CH321' de uma planilha Excel, mantendo todas as colunas originais.

    Args:
        caminho_arquivo (str): O caminho completo para o arquivo Excel.

    Returns:
        pd.DataFrame: DataFrame com os dados da planilha.
        Encerra o script se o arquivo ou a coluna 'N° Convênio' não forem encontrados.
    """
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name='Document_CH321')
        if 'N° Convênio' not in df.columns:
            raise ValueError("A coluna 'N° Convênio' é obrigatória e não foi encontrada na planilha.")
        logging.info(f"Planilha '{caminho_arquivo}' lida com sucesso.")
        return df
    except FileNotFoundError:
        logging.error(f"Arquivo da planilha de entrada não encontrado em: {caminho_arquivo}")
        exit()
    except Exception as e:
        logging.error(f"Erro ao ler a planilha: {e}")
        exit()


def salvar_planilha_saida(df, arquivo_entrada, primeira_vez=False):
    """
    Cria ou atualiza a planilha de saída com os dados processados.

    Na primeira execução, cria um novo arquivo com sufixo '_COM_DATAS'.
    Nas execuções subsequentes, atualiza o mesmo arquivo.

    Args:
        df (pd.DataFrame): O DataFrame com os dados atualizados.
        arquivo_entrada (str): O caminho do arquivo de entrada original.
        primeira_vez (bool): Flag para indicar se é a primeira vez que o arquivo está sendo salvo.

    Returns:
        str: O caminho do arquivo de saída.
    """
    try:
        pasta = os.path.dirname(arquivo_entrada)
        nome_base = os.path.splitext(os.path.basename(arquivo_entrada))[0]
        arquivo_saida = os.path.join(pasta, f"{nome_base}_COM_DATAS.xlsx")

        # Se for a primeira vez, verifica se o arquivo já existe para criar uma versão.
        if primeira_vez:
            contador = 1
            while os.path.exists(arquivo_saida):
                arquivo_saida = os.path.join(pasta, f"{nome_base}_COM_DATAS_{contador}.xlsx")
                contador += 1

        # Utiliza o ExcelWriter para criar ou substituir a aba, preservando outras abas se existirem.
        mode = 'w' if primeira_vez else 'a'
        if_sheet_exists = None if primeira_vez else 'replace'

        with pd.ExcelWriter(arquivo_saida, engine='openpyxl', mode=mode, if_sheet_exists=if_sheet_exists) as writer:
            df.to_excel(writer, sheet_name='Document_CH321', index=False)
        
        logging.info(f"Planilha de saída salva com sucesso em: {arquivo_saida}")
        return arquivo_saida
    except PermissionError:
        logging.error(f"Erro de permissão. Feche o arquivo '{arquivo_saida}' e tente novamente.")
        return None
    except Exception as e:
        logging.error(f"Erro ao salvar a planilha de saída: {e}")
        return None


def navegar_menu_principal(driver, instrumento):
    """
    Navega pelo menu do sistema e pesquisa por um instrumento específico.

    Args:
        driver (webdriver.Chrome): A instância do driver.
        instrumento (str): O número do convênio a ser pesquisado.

    Returns:
        bool: True se a navegação for bem-sucedida, False caso contrário.
    """
    try:
        # Nota: Os seletores XPath completos são frágeis e quebram facilmente com
        # qualquer mudança no site. É recomendável usar seletores mais robustos (ID, classe, etc.).
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]"))).click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]"))).click()
        
        campo_pesquisa = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")))
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]"))).click()
        time.sleep(1) # Pausa para aguardar a renderização dos resultados.
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]"))).click()
        return True
    except Exception as e:
        logging.warning(f"Instrumento '{instrumento}' não encontrado ou erro na navegação: {e}")
        return False


def acessar_aba_movimentacao_financeira(driver):
    """
    Acessa a aba "Movimentação Financeira" usando seletores primários e de fallback.

    Args:
        driver (webdriver.Chrome): A instância do driver.

    Returns:
        bool: True se o acesso for bem-sucedido, False caso contrário.
    """
    try:
        # Primeiro clique para expandir o menu principal da aba.
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[6]/div/span/span"))).click()
        except TimeoutException:
            logging.info("Seletor principal do menu falhou. Tentando via JavaScript.")
            driver.execute_script('document.querySelector("#div_-481524888 > span > span").click()')
        time.sleep(1)

        # Segundo clique para acessar a sub-aba de movimentação.
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[26]/div/span/span"))).click()
        except TimeoutException:
            logging.info("Seletor principal da sub-aba falhou. Tentando via JavaScript.")
            driver.execute_script('document.querySelector("#menu_link_-481524888_1304359359 > div > span").click()')
        time.sleep(1)
        
        return True
    except Exception as e:
        logging.error(f"Erro crítico ao acessar a aba 'Movimentação Financeira': {e}")
        return False


def coletar_data_recente(driver):
    """
    Coleta a data mais recente da tabela de movimentação financeira.

    Args:
        driver (webdriver.Chrome): A instância do driver.

    Returns:
        tuple: Uma tupla contendo (data_recente_str, status_movimentacao).
               Ex: ("dd/mm/aaaa", "SIM") ou (None, "NÃO").
    """
    try:
        tabela = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[16]/div[2]/form/table")))
        rows = tabela.find_elements(By.TAG_NAME, "tr")
        datas = []

        for row in rows[1:]:  # Pula o cabeçalho.
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 2:
                data_text = cells[1].text.strip()
                if re.match(r"\d{2}/\d{2}/\d{4}", data_text):
                    datas.append(datetime.strptime(data_text, "%d/%m/%Y"))
        
        if datas:
            data_recente = max(datas).strftime("%d/%m/%Y")
            return data_recente, "SIM"
        else:
            return None, "NÃO"
    except TimeoutException:
        logging.info("Tabela de movimentação financeira não encontrada. Considerado como 'NÃO'.")
        return None, "NÃO"
    except Exception as e:
        logging.warning(f"Erro ao coletar data da tabela: {e}")
        return None, "NÃO"


def navegar_voltar_inicio(driver):
    """
    Retorna à página principal do sistema para iniciar o próximo ciclo.

    Args:
        driver (webdriver.Chrome): A instância do driver.
    """
    try:
        url_principal = "https://discricionarias.transferegov.sistema.gov.br/voluntarias/Principal/Principal.do"
        driver.get(url_principal)
        WebDriverWait(driver, 10).until(EC.url_to_be(url_principal))
        logging.info("Navegou de volta para a página inicial.")
    except Exception as e:
        logging.warning(f"Não foi possível retornar à página inicial. Tentando recarregar. Erro: {e}")
        driver.refresh()
        time.sleep(3)


def main():
    """
    Função principal que orquestra todo o processo de automação.
    """
    df = ler_planilha(INPUT_EXCEL_PATH)
    total_instrumentos = len(df)

    # Adiciona as novas colunas ao DataFrame.
    df['Data Mais Recente'] = None
    df['Movimentação'] = None

    driver = conectar_navegador_existente()
    arquivo_saida = None
    tempos_por_instrumento = []
    inicio_geral = time.time()

    for index, row in df.iterrows():
        instrumento = row['N° Convênio']
        inicio_instrumento = time.time()

        # --- Lógica de progresso e estimativa de tempo ---
        progresso = (index + 1) / total_instrumentos * 100
        restantes = total_instrumentos - (index + 1)

        logging.info("=" * 60)
        logging.info(f"Processando instrumento {index + 1}/{total_instrumentos}: {instrumento}")
        logging.info(f"Progresso: {progresso:.1f}% | Restantes: {restantes}")

        if tempos_por_instrumento:
            tempo_medio = sum(tempos_por_instrumento) / len(tempos_por_instrumento)
            tempo_estimado = tempo_medio * restantes
            horas, rem = divmod(tempo_estimado, 3600)
            minutos, seg = divmod(rem, 60)
            logging.info(f"Tempo estimado restante: {int(horas):02d}h {int(minutos):02d}m {int(seg):02d}s")
        
        # --- Processo de Coleta ---
        if navegar_menu_principal(driver, instrumento):
            if acessar_aba_movimentacao_financeira(driver):
                data_recente, possui_movimentacao = coletar_data_recente(driver)
            else:
                data_recente, possui_movimentacao = None, "ERRO_NAVEGACAO"
        else:
            data_recente, possui_movimentacao = None, "NAO_ENCONTRADO"

        # --- Atualização e Salvamento ---
        df.loc[index, 'Data Mais Recente'] = data_recente
        df.loc[index, 'Movimentação'] = possui_movimentacao

        # Salva o arquivo pela primeira vez ou a cada 5 instrumentos.
        if index == 0:
            arquivo_saida = salvar_planilha_saida(df, INPUT_EXCEL_PATH, primeira_vez=True)
        elif (index + 1) % 5 == 0 or (index + 1) == total_instrumentos:
            salvar_planilha_saida(df, arquivo_saida)

        tempo_instrumento = time.time() - inicio_instrumento
        tempos_por_instrumento.append(tempo_instrumento)
        logging.info(f"Instrumento processado em {tempo_instrumento:.1f} segundos.")

        if index < total_instrumentos - 1:
            navegar_voltar_inicio(driver)

    # --- Finalização ---
    tempo_total = time.time() - inicio_geral
    h, rem = divmod(tempo_total, 3600)
    m, s = divmod(rem, 60)

    logging.info("=" * 60)
    logging.info("PROCESSO CONCLUÍDO!")
    logging.info(f"Tempo total de execução: {int(h):02d}h {int(m):02d}m {int(s):02d}s")
    logging.info(f"Planilha de saída final salva em: {arquivo_saida}")
    logging.info("=" * 60)

    driver.quit()


if __name__ == "__main__":
    main()
