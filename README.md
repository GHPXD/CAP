
# CAP Workflow Automation

Este projeto tem como objetivo automatizar processos do CAP (Customer Approval Process) da Votorantim utilizando automação web com **Selenium** e uma interface gráfica em **Tkinter**. O usuário pode escolher entre duas funções principais: **DownloadCAP** e **AprovaçãoCAP**.

---

## Estrutura do Código

Este projeto é organizado da seguinte forma:

### 1. **Estrutura de Pastas**

```
CAP_Workflow_Automation/
│
├── src/                    # Código-fonte do projeto
│   ├── main.py             # Arquivo principal para rodar a aplicação
│   ├── download_cap.py     # Função para realizar o download
│   ├── aprovacao_cap.py    # Função para aprovar processos
│   └── utils.py            # Funções auxiliares para operações diversas
│
├── assets/                 # Recursos gráficos (imagens, ícones, etc.)
│   └── logo.png            # Logo da aplicação
│
├── dist/                   # Pasta gerada pelo PyInstaller com o executável
│   └── main.exe (Windows)  # Executável para Windows
│   └── main (Linux/macOS)  # Executável para Linux/macOS
│
├── README.md               # Documentação do projeto
└── requirements.txt        # Dependências do projeto
```

### 2. **Arquivo `main.py`**

Este é o arquivo principal que contém a interface gráfica usando o Tkinter, permitindo ao usuário escolher entre as funções **DownloadCAP** e **AprovaçãoCAP**.

- O **Tkinter** é usado para criar a interface gráfica onde o usuário pode escolher qual função rodar.
- O **Selenium** é usado para automatizar o processo de login e download/execução de tarefas na plataforma web da Votorantim.

### 3. **Funções**

- **download_cap.py**: Contém a função `DownloadCAP` que automatiza o processo de login e download de arquivos da plataforma.
- **aprovacao_cap.py**: Contém a função `AprovaçãoCAP` que automatiza o processo de aprovação de tarefas.
- **utils.py**: Contém funções auxiliares para operações diversas, como movimentação de arquivos.

### 4. **Dependências**

- `selenium`: Biblioteca usada para automação do navegador.
- `tkinter`: Biblioteca usada para criar a interface gráfica.
- `PyInstaller`: Ferramenta para empacotar o código em um executável para distribuição.

---

## Instruções para o Usuário

### 1. **Baixar o Arquivo Executável**

O usuário pode baixar o arquivo **executável** para rodar a aplicação sem precisar de instalação de Python ou outras dependências.

1. Baixe o arquivo `CAP_Workflow_Automation.zip` do link fornecido **[Executável](https://example.com/baixar-o-executavel)**.
2. Extraia o conteúdo do arquivo ZIP em uma pasta de sua escolha.

### 2. **Executar o Programa**

1. Navegue até a pasta onde você extraiu o arquivo ZIP.
2. Clique duas vezes no arquivo `main.exe` (Windows) ou `main` (Linux/macOS) para iniciar o programa.

### 3. **Escolher a Função**

Quando a interface gráfica abrir, você verá duas opções:

- **DownloadCAP**: Para automatizar o processo de download de arquivos.
- **AprovaçãoCAP**: Para automatizar o processo de aprovação de tarefas.

Escolha a função desejada e clique no botão correspondente para iniciar a automação. O Selenium irá acessar a plataforma, fazer o login e executar a tarefa selecionada.

### 4. **Saída do Programa**

Após a execução, o programa mostrará uma mensagem informando que a tarefa foi concluída com sucesso ou se ocorreu algum erro.

---

## Como Funciona a Automação

### **1. DownloadCAP**

Quando a opção **DownloadCAP** é selecionada, o programa realiza o seguinte:

1. Abre o navegador e acessa o link da plataforma Votorantim.
2. Realiza o login com as credenciais configuradas.
3. Automação de navegação na plataforma:
   - Seleciona as colunas necessárias.
   - Executa o processo de download dos arquivos.
4. Move os arquivos baixados para uma pasta específica configurada no código.

### **2. AprovaçãoCAP**

Ao selecionar a opção **AprovaçãoCAP**, o programa realiza a aprovação de tarefas na plataforma, automatizando o processo de navegação, filtragem e aprovação dos itens pendentes.

---

## Arquivos Incluídos

- **`main.exe`** (ou `main` para Linux/macOS): O executável do programa.
- **`logo.png`**: Logo da aplicação, incluído na interface gráfica.
- **`requirements.txt`**: Arquivo com as dependências necessárias para rodar o projeto (caso o usuário precise configurar o ambiente de desenvolvimento).

---

## Suporte

Se você encontrar algum problema ao usar o programa ou precisar de mais informações, entre em contato com **[suporte@example.com]** ou consulte o repositório no **[GitHub](https://github.com/usuario/repositorio)**.

---

### Como Gerar o Executável (Para Desenvolvedores)

Caso você precise gerar o executável novamente ou personalizar o processo, siga as instruções abaixo.

1. **Instale as Dependências**:

   Se você não tiver as bibliotecas necessárias, execute:

   ```bash
   pip install -r requirements.txt
   ```

2. **Gerar o Executável com PyInstaller**:

   Navegue até o diretório onde está o código-fonte e execute:

   ```bash
   pyinstaller --onefile --windowed --add-data "assets/logo.png;assets" src/main.py
   ```

   - **`--onefile`**: Cria um único arquivo executável.
   - **`--windowed`**: Impede que a janela de terminal apareça.
   - **`--add-data`**: Inclui os recursos necessários (como `logo.png`).

3. O executável será gerado na pasta `dist/`:

   ```
   dist/
     └── main.exe (Windows) ou main (Linux/macOS)
   ```

---

## Licença

Este projeto é licenciado sob a licença MIT - veja o arquivo **LICENSE** para mais detalhes.

Entendido! A seguir, fiz as correções no README conforme a estrutura final do seu projeto e removi a parte sobre a criação opcional do executável. Além disso, ajustei para refletir a estrutura final do seu projeto com base nos diretórios e arquivos mencionados.

---

# Automação de Aprovação de Tarefas - Sistema CAP

Este repositório contém scripts Python que automatizam o processo de aprovação de tarefas no sistema CAP (Votorantim). O processo envolve a interação com a planilha Excel para filtrar tarefas, acessar o sistema CAP e aprovar as tarefas automaticamente, seguindo a lógica de um código VBA anterior.

## Estrutura do Projeto

```
CAP_Automacao/
│
├── assets/                           # Imagens e arquivos auxiliares
│   └── logo.png                      # Imagem de logo para interface
│
├── build/                            # Arquivos gerados após a criação do executável com PyInstaller
│   └── main/                         # Arquivos temporários criados durante a execução do PyInstaller
│
├── dist/                             # Executáveis gerados com PyInstaller
│   └── main.exe                      # Executável gerado para Windows
│
├── src/                              # Scripts principais
│   ├── aprovacao_cap.py              # Script de aprovação de CAP
│   ├── download_cap.py              # Script de download de CAP
│   ├── main.py                      # Arquivo principal que inicia o processo
│   ├── utils.py                     # Funções utilitárias, como login e validações
│
├── requirements.txt                 # Arquivo de dependências para o Python
├── main.spec                        # Arquivo de configuração do PyInstaller
└── README.md                        # Este arquivo de documentação
```

## Pré-requisitos

Antes de executar os scripts, você precisa ter alguns pré-requisitos instalados:

1. **Python 3.8 ou superior** (certifique-se de que o Python está instalado e configurado corretamente no seu sistema).
2. **Bibliotecas Python necessárias**:
    - `selenium`
    - `pandas`
    - `tkinter` (incluso por padrão no Python para interfaces gráficas)

### Instalando as Bibliotecas

Para instalar as dependências necessárias, execute:

```bash
pip install -r requirements.txt
```

3. **ChromeDriver**: O Selenium precisa de um `ChromeDriver` para interagir com o navegador Chrome. Baixe a versão compatível com a sua versão do Chrome [aqui](https://sites.google.com/a/chromium.org/chromedriver/). Coloque o `chromedriver.exe` na pasta raiz do projeto ou especifique o caminho no código.

## Como Executar

### 1. **Executar via Interface Gráfica**

O script `main.py` oferece uma interface gráfica simples para escolher qual operação executar: **Download de CAP** ou **Aprovação de CAP**.

- Baixe e extraia o projeto em seu computador.
- Execute o arquivo **`main.exe`** (gerado com PyInstaller) para abrir a interface gráfica.
- Na interface, escolha qual processo deseja rodar (Download ou Aprovação).
- Siga as instruções na tela para autenticação e execução do processo.

### 2. **Executar via Linha de Comando**

Caso prefira, você também pode rodar diretamente o código Python via linha de comando, sem a interface gráfica.

#### Executar o **DownloadCAP**:

No terminal, navegue até a pasta `src/` e execute o script:

```bash
python download_cap.py
```

#### Executar o **AprovaçãoCAP**:

Para o processo de aprovação de CAP, execute:

```bash
python aprovacao_cap.py
```

## Como Funciona

### **1. DownloadCAP**

O script `download_cap.py` realiza o seguinte fluxo:

- **Login** no sistema CAP.
- **Navega** até a página de tarefas.
- **Filtra e seleciona** colunas específicas.
- **Baixa arquivos** de tarefas com base em links encontrados na página.

### **2. AprovaçãoCAP**

O script `aprovacao_cap.py` automatiza a aprovação de tarefas com base em dados da planilha Excel. Ele executa o seguinte processo:

1. **Carrega** a planilha de tarefas.
2. **Verifica** se as tarefas estão aprovadas na coluna de status e se as condições de data são atendidas (tarefa deve ser de hoje ou ontem).
3. **Acessa** a interface web do CAP e aprova as tarefas.
4. **Atualiza** a planilha para indicar que a tarefa foi processada.

### **Ajustes de Planilha**

O script assume que a planilha Excel tem uma estrutura específica. Verifique se as seguintes colunas estão presentes:

- **Coluna 1**: Número da tarefa.
- **Coluna 2**: Fluxo de Trabalho.
- **Coluna 14**: Status da tarefa (Deve ser "Aprovado" para processar).
- **Coluna X**: Data relacionada à tarefa (Deve ser igual a "hoje" ou "ontem").

Exemplo de estrutura da planilha:

| Número | Fluxo de Trabalho | ... | Status   | Data       |
|--------|--------------------|-----|----------|------------|
| 12345  | WF001              | ... | Aprovado | 12/01/2025 |
| 67890  | WF002              | ... | Pendente | 13/01/2025 |

### **Configuração do Selenium**

O Selenium usa o `ChromeDriver` para automatizar o navegador. Se o caminho do `ChromeDriver` não estiver configurado corretamente, o código irá falhar. Para configurá-lo:

1. Baixe a versão do `ChromeDriver` compatível com o seu Chrome.
2. Coloque o `chromedriver.exe` na pasta raiz do projeto ou ajuste o caminho no código.

```python
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

service = Service("path/to/chromedriver")  # Substitua o caminho pelo local do seu ChromeDriver
driver = webdriver.Chrome(service=service)
```

### **Variáveis Importantes**

- **`ws.Cells(i, x)`**: Acessa as células da planilha (com base no número da linha `i` e a coluna `x`).
- **Planilha**: O script assume que você está utilizando o Excel para armazenar as tarefas que precisam ser aprovadas ou baixadas.
- **Datas**: O script verifica se a data da tarefa é igual a "hoje" ou "ontem" para processá-la.
