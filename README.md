<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">  
</head>
<body>

<header>
  <h1>🤖 Robô Selenium - Consulta de Movimentações Financeiras</h1>
  <p>Automatize a extração de dados de convênios no portal TransfereGov</p>
</header>

<main>

  <section>
    <h2>✨ Funcionalidades</h2>
    <ul>
      <li>✅ Conecta ao navegador já autenticado com Chrome DevTools</li>
      <li>✅ Lê dados da planilha de entrada (.xlsx)</li>
      <li>✅ Navega e pesquisa convênios no sistema TransfereGov</li>
      <li>✅ Extrai datas mais recentes de movimentações financeiras</li>
      <li>✅ Gera planilha de saída com colunas <code>Data Mais Recente</code> e <code>Movimentação</code></li>
    </ul>
  </section>

  <section>
    <h2>📁 Estrutura do Projeto</h2>
    <pre><code>.
├── robo_movimentacao.py
├── requirements.txt
├── README.html
└── planilhas/
    └── entrada.xlsx</code></pre>
  </section>

  <section>
    <h2>🛠️ Requisitos</h2>
    <ul>
      <li>Python 3.8+</li>
      <li>Google Chrome com debug ativado:
        <pre><code>chrome.exe --remote-debugging-port=9222 --user-data-dir="C:/chrome-dev-profile"</code></pre>
      </li>
      <li>Instale as dependências:
        <pre><code>pip install -r requirements.txt</code></pre>
      </li>
    </ul>
    <p><strong>Conteúdo de <code>requirements.txt</code>:</strong></p>
    <pre><code>selenium
pandas
openpyxl
webdriver-manager</code></pre>
  </section>

  <section>
    <h2>📈 Como Usar</h2>
    <ol>
      <li>Abra o Chrome com porta de depuração habilitada</li>
      <li>Altere o caminho da planilha no script</li>
      <li>Execute o script com <code>python robo_movimentacao.py</code></li>
    </ol>
  </section>

  <section>
    <h2>🖼️ Exemplo de Saída</h2>
    <table>
      <thead>
        <tr>
          <th>N° Convênio</th>
          <th>Data Mais Recente</th>
          <th>Movimentação</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>2020XXXX</td>
          <td>12/03/2024</td>
          <td><span class="badge">SIM</span></td>
        </tr>
        <tr>
          <td>2021YYYY</td>
          <td>—</td>
          <td><span class="badge">NÃO</span></td>
        </tr>
      </tbody>
    </table>
  </section>

  <section>
    <h2>⚠️ Possíveis Problemas</h2>
    <ul>
      <li><strong>Erro ao conectar ao navegador:</strong> verifique se o Chrome foi aberto com o parâmetro correto.</li>
      <li><strong>Coluna 'N° Convênio' não encontrada:</strong> confirme o nome da aba na planilha.</li>
      <li><strong>Site lento ou travado:</strong> o script tentará recarregar e seguir em frente.</li>
    </ul>
  </section>

  <section>
    <h2>👨‍💻 Autor</h2>
    <p><strong>Diego Brito</strong><br>
      Desenvolvedor Python e entusiasta de automações com Selenium</p>
  </section>

  <section>
    <h2>📝 Licença</h2>
    <p>Este projeto é de uso livre para fins educacionais e internos. Consulte o autor para fins comerciais.</p>
  </section>

</main>

</body>
</html>
