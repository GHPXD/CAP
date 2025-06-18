# ⚙️ Automação de Aprovação de Tarefas - Sistema CAP

[![Python Version](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Automatize tarefas repetitivas no sistema CAP (Votorantim) com este conjunto de scripts Python! A solução interage com a plataforma **ServiceNow** para baixar e aprovar tarefas de forma eficiente, com suporte a interface gráfica e planilhas Excel.

---

## 🚀 Funcionalidades

✅ **Download de Relatórios/Anexos** do CAP  
✅ **Aprovação Automatizada de Tarefas** com base em critérios definidos em planilhas Excel  
✅ **Interface Gráfica Intuitiva** (Tkinter) para facilitar o uso  
✅ **Automação com Selenium** e tratamento robusto de interações com a web  

---

## 📁 Estrutura do Projeto

```
CAP_Automacao/
│
├── assets/                  # Imagens e arquivos auxiliares
│   └── logo.png             # (Opcional) Logo da interface
│
├── src/                     # Scripts principais
│   ├── aprovacao_cap.py     # Aprova tarefas automaticamente
│   ├── download_cap.py      # Baixa relatórios/anexos
│   └── main.py              # Interface gráfica e controle geral
│
├── .gitignore               # Arquivos/pastas ignoradas pelo Git
├── requirements.txt         # Dependências do projeto
├── main.spec                # (Opcional) Configuração para PyInstaller
└── README.md                # Este arquivo
```

---

## 📋 Pré-requisitos

### 🐍 Requisitos de Software

- **Python 3.8 ou superior**
- **Google Chrome instalado**

### 📦 Bibliotecas Python

- `selenium`
- `pandas`
- `tkinter` (normalmente já incluso com Python)
- `webdriver_manager` (para facilitar o uso do ChromeDriver)

> Para instalar todas as dependências:
```bash
pip install -r requirements.txt
```

---

## ⚙️ Executando o Projeto

### 1️⃣ Usando a Interface Gráfica (Recomendado)

```bash
python src/main.py
```

A janela interativa permitirá que você escolha entre:

- ✅ Rodar **Download de CAP**
- ✅ Rodar **Aprovação de CAP**

### 2️⃣ Executando Diretamente via Terminal

- Para **Download de Relatórios**:
  ```bash
  python src/download_cap.py
  ```

- Para **Aprovação de Tarefas**:
  ```bash
  python src/aprovacao_cap.py
  ```

> ⚠️ Execute sempre os scripts a partir da **raiz do projeto** (`CAP_Automacao/`) para garantir o correto funcionamento dos caminhos relativos.

---

## 🔄 Fluxo de Funcionamento

### 🧾 `download_cap.py`

1. Login no sistema CAP (`venergia.capworkflow.com`)
2. Acesso à aba **TAREFAS**
3. Configuração de colunas da tabela
4. Iteração nas tarefas e **download automático** dos arquivos
5. Arquivos movidos para: `C:/Users/SeuUsuario/Documents/teste` (pasta configurável)

---

### ✅ `aprovacao_cap.py`

1. Seleção da planilha Excel (.xlsx) com os dados de controle
2. Leitura e validação das colunas:
   - `Solicitação`
   - `FRS`
   - `Status Aprov.` (deve estar como "Aprovado")
   - `Data Aprovação` (hoje ou ontem)
3. Aprovação automatizada de cada tarefa válida:
   - Navega na aba **TAREFAS**
   - Clica na linha da solicitação
   - Pressiona o botão “Aprovar”
   - Trata possíveis alertas de confirmação

---

## 📊 Estrutura Esperada da Planilha Excel

| Solicitação | FRS        | ... | Status Aprov. | Data Aprovação |
|-------------|------------|-----|----------------|----------------|
| 123456      | 1031234567 | ... | Aprovado       | 2024-06-17     |
| 234567      | 1039876543 | ... | Aprovado       | 2024-06-18     |
| 345678      | 1031928375 | ... | Pendente       | 2024-06-16     |

---

## ⚠️ Considerações Importantes

- 🔐 **Segurança das Credenciais**: Evite manter senhas no código. Use variáveis de ambiente ou arquivos `.env` (ignorados pelo Git).
- 🕵️ **Seletores XPATH**: Mudanças na interface do site podem afetar o funcionamento. Atualize os seletores conforme necessário.
- 🧪 **Tratamento de Erros**: O uso de `try-except` e `logging` facilita a depuração. Verifique o console ou adicione logs se necessário.
- 🌐 **Compatibilidade com Navegador**: Projetado para funcionar com **Google Chrome**.

---

## 📄 Licença

Este projeto está licenciado sob a [Licença MIT](https://opensource.org/licenses/MIT).  
Sinta-se livre para usar, modificar e contribuir!
