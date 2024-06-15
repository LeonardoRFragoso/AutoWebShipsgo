# Shipsgo Automation Script

## Descrição
Este projeto consiste em um script Python que automatiza a extração de dados de um banco de dados Firebird, processa os resultados, faz login no site Shipsgo para obter informações adicionais sobre contêineres e atualiza uma planilha Excel com os dados obtidos.

## Funcionalidades
- Verificação do status do serviço Firebird.
- Execução de consultas SQL utilizando `isql`.
- Validação dos dados extraídos do banco de dados.
- Login automatizado no site Shipsgo utilizando Selenium.
- Extração e processamento de informações de contêineres no Shipsgo.
- Atualização de uma planilha Excel com os dados dos contêineres.

## Requisitos
- Python 3.x
- Selenium
- Openpyxl
- Dotenv
- Um navegador Chrome instalado
- Chromedriver compatível com a versão do Chrome instalada

## Instalação

1. Clone o repositório:
    ```sh
    git clone https://github.com/usuario/repositorio.git
    cd repositorio
    ```

2. Crie um ambiente virtual e ative-o:
    ```sh
    python -m venv venv
    source venv/bin/activate # Linux/macOS
    venv\Scripts\activate # Windows
    ```

3. Instale as dependências:
    ```sh
    pip install -r requirements.txt
    ```

4. Configure as variáveis de ambiente:
    Crie um arquivo `.env` na raiz do projeto e defina as seguintes variáveis:
    ```env
    DB_PATH=C:\\robo\\CONTROLE.FDB
    DB_USER=sysdba
    DB_PASSWORD=Q5QIST
    ISQL_PATH=C:\\Program Files\\Firebird\\Firebird_2_5\\bin\\isql.exe
    SCRIPT_SQL_PATH=C:\\robo\\script.sql
    SHIPS_GO_USER=seu_email@exemplo.com
    SHIPS_GO_PASSWORD=sua_senha
    CHROMEDRIVER_PATH=C:\\caminho\\para\\chromedriver.exe
    ```

## Uso

1. Certifique-se de que o serviço Firebird está em execução.

2. Execute o script:
    ```sh
    python script.py
    ```

3. O script irá:
   - Verificar o status do Firebird.
   - Executar a consulta SQL definida.
   - Validar os dados extraídos.
   - Fazer login no Shipsgo e processar os dados dos contêineres.
   - Atualizar a planilha Excel com os dados processados.

## Estrutura do Projeto
```plaintext
.
├── .env.example
├── README.md
├── requirements.txt
├── script.py
└── script.log
