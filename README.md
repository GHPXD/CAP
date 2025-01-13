
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