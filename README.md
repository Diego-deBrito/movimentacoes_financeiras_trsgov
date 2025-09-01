# RPA para Coleta de Datas de Movimentação Financeira

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![Libraries](https://img.shields.io/badge/Libraries-Selenium%20%7C%20Pandas-orange.svg)
![Status](https://img.shields.io/badge/Status-Funcional-success.svg)

Este projeto é um robô de automação (RPA) construído em Python, projetado para enriquecer uma planilha de controle existente com dados de data extraídos do portal TransfereGov do governo brasileiro.

## Descrição do Projeto

A principal função deste robô é automatizar a consulta da data de movimentação financeira mais recente para uma lista de convênios. O script lê um arquivo Excel, itera sobre cada número de convênio, navega no portal, extrai a data desejada e, por fim, atualiza a planilha original com as novas informações, salvando o resultado em um novo arquivo para manter a integridade dos dados originais.

## Funcionalidades Principais

- **Sistema de Logging:** Utiliza o módulo `logging` do Python para registrar detalhadamente cada passo da execução, salvando os logs em um arquivo `robo_log.txt` e exibindo-os no console.
- **Enriquecimento de Dados:** Lê uma planilha Excel, preserva todas as colunas e dados originais e adiciona duas novas colunas: `Data Mais Recente` e `Movimentação`.
- **Navegação Robusta:** Emprega mecanismos de fallback (tentando seletores alternativos via JavaScript) para navegar em menus que podem ter variações na interface do usuário.
- **Salvamento Incremental:** Salva o progresso na planilha de saída a cada 5 itens processados, garantindo que nenhum dado seja perdido em caso de interrupção durante longas execuções.
- **Controle de Versão de Arquivos:** Cria um novo arquivo de saída com o sufixo `_COM_DATAS`. Se um arquivo com o mesmo nome já existir, ele cria versões numeradas (ex: `_COM_DATAS_1`, `_COM_DATAS_2`) para evitar a sobrescrita de execuções anteriores.
- **Estimativa de Tempo de Execução:** Calcula e exibe o tempo estimado para a conclusão do processo com base no tempo médio gasto por item.
- **Conexão com Navegador Existente:** Permite que o script se conecte a uma sessão do Chrome já autenticada, simplificando o processo de login e evitando problemas com CAPTCHA.

## Pré-requisitos

- [Python 3.7](https://www.python.org/downloads/) ou superior
- [Google Chrome](https://www.google.com/chrome/) (navegador web)

## Instalação e Configuração

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
    cd seu-repositorio
    ```

2.  **Crie e ative um ambiente virtual (recomendado):**
    ```bash
    # Para Windows
    python -m venv venv
    .\venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    Crie um arquivo chamado `requirements.txt` com o seguinte conteúdo:
    ```
    pandas
    selenium
    webdriver-manager
    openpyxl
    ```
    Execute o comando de instalação:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configure o Caminho do Arquivo de Entrada:**
    Abra o script Python e edite a constante `INPUT_EXCEL_PATH` no início do arquivo, inserindo o caminho completo para a sua planilha.
    ```python
    # Caminho para a planilha de entrada. Altere este valor conforme necessário.
    INPUT_EXCEL_PATH = r"C:\caminho\para\sua\planilha.xlsx"
    ```

## Como Executar

1.  **Prepare a Planilha de Entrada:**
    Garanta que sua planilha Excel contenha uma aba chamada `Document_CH321` e, dentro dela, uma coluna chamada `N° Convênio`.

2.  **Inicie o Google Chrome em Modo de Depuração:**
    Feche todas as janelas do Chrome e inicie uma nova usando o terminal com o comando abaixo. Isso é essencial para que o script possa se conectar.
    ```bash
    # Para Windows
    "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
    ```

3.  **Acesse o Sistema Manualmente:**
    Na nova janela do Chrome, navegue até o portal TransfereGov, faça o login e deixe o navegador pronto na página principal.

4.  **Execute o Script:**
    Abra um novo terminal (não o que você usou para iniciar o Chrome), navegue até a pasta do projeto e execute o script:
    ```bash
    python nome_do_script.py
    ```

O robô iniciará a execução, e você poderá acompanhar o progresso através dos logs no console e no arquivo `robo_log.txt`.

## Observações Importantes

> **Fragilidade dos Seletores (XPath):** Este script utiliza seletores XPath absolutos, que são suscetíveis a quebrar com mudanças na estrutura do site. Para maior durabilidade, recomenda-se a substituição por seletores mais estáveis como IDs, classes ou XPaths relativos.

> **Uso Específico:** O robô foi desenhado para o fluxo de navegação e a estrutura de DOM do portal TransfereGov. Ele não funcionará em outros sites sem adaptações significativas.
