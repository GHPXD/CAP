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

- ## Licença

Este projeto é licenciado sob a licença MIT - veja o arquivo **LICENSE** para mais detalhes.
