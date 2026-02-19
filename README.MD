# FAI-SQL

**Gerenciador de Comandos SQL Desktop** — Software para armazenamento e execução de comandos SQL recorrentes.

O projeto possui **duas versões de interface**:

| Versão | Pasta | Interface | Status |
|--------|-------|-----------|--------|
| **Clássica** | `./` (raiz) | Tkinter | ✅ Estável |
| **Fluent** | `fai-sql-fluent/` | PySide6 + Fluent Widgets | ✅ Nova |

---

## 📋 Recursos

- ✅ **Interface moderna** com temas claro/escuro
- ✅ **Gerenciador de múltiplas conexões** SQL Server
- ✅ **Armazenamento de comandos** SQL para reuso
- ✅ **Edição e exclusão** de comandos salvos
- ✅ **Syntax Highlighting** SQL em tempo real
- ✅ **Exportação** para CSV e Excel
- ✅ **Histórico** de comandos executados
- ✅ **Undo/Redo** (Ctrl+Z / Ctrl+Y)
- ✅ **Criptografia AES-256** com senha mestra
- ✅ **Navegação lateral** estilo Windows 11 (versão Fluent)

---

## 🚀 Instalação

### Pré-requisitos

- Python 3.10+
- Driver ODBC do SQL Server instalado

### Versão Clássica (Tkinter)

```bash
git clone https://github.com/seu-usuario/FAI-SQL.git
cd FAI-SQL
pip install -r requirements.txt
python main.py
```

### Versão Fluent (PySide6)

```bash
cd FAI-SQL
pip install -r fai-sql-fluent/requirements.txt
python fai-sql-fluent/main.py
```

Na primeira execução de qualquer versão, crie sua **senha mestra** para proteger seus dados.

> **💡 Nota**: Ambas as versões compartilham os mesmos arquivos de dados (`config.dat`, `comandos.dat`, `historico.dat`) e o mesmo `crypto_utils.py`.

---

## 📁 Estrutura de Arquivos

```
FAI-SQL/
├── main.py                  # Versão Clássica (Tkinter)
├── crypto_utils.py          # Módulo de criptografia (compartilhado)
├── requirements.txt         # Dependências (versão clássica)
├── icon.ico                 # Ícone da aplicação
│
├── fai-sql-fluent/          # Versão Fluent (PySide6)
│   ├── main.py              # Entry point + login
│   ├── app.py               # Janela principal (FluentWindow)
│   ├── requirements.txt     # Dependências (versão fluent)
│   ├── pages/
│   │   ├── editor.py        # Editor SQL + resultados
│   │   ├── connections.py   # Gerenciador de conexões
│   │   ├── commands.py      # Comandos salvos
│   │   └── history.py       # Histórico de execuções
│   └── utils/
│       ├── database.py      # Funções de conexão e queries
│       └── sql_highlighter.py # Syntax highlighting SQL
│
├── config.dat               # Configurações (criptografado)
├── comandos.dat             # Comandos salvos (criptografado)
├── historico.dat            # Histórico (criptografado)
├── .salt                    # Salt de criptografia
├── .verify                  # Verificador de senha
└── docs/                    # Documentação
```

---

## ⚙️ Configuração

### Conexão com Banco de Dados

1. Acesse a página **Conexões** (versão Fluent) ou **Sistema → Conexão** (versão clássica)
2. Preencha os dados: Nome, Servidor/IP, Usuário, Senha, Banco
3. Clique em **Testar** para verificar
4. Clique em **Salvar** para adicionar

Você pode criar múltiplas conexões e alternar entre elas clicando em **Usar**.

---

## 🎨 Temas

Alterne entre tema claro e escuro:

- **Versão Fluent**: Botão **Alternar Tema** na barra lateral
- **Versão Clássica**: Menu **Sistema → Tema Escuro/Claro**

A preferência é salva automaticamente.

---

## 📤 Exportação

Após executar um SELECT:
- Clique em **CSV** para exportar como CSV
- Clique em **Excel** para exportar como Excel formatado

---

## 🔨 Build (Executável)

### Versão Clássica (Tkinter)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico --name "FAISQL" main.py
```

O executável será gerado em `dist/FAISQL.exe`.

### Versão Fluent (PySide6)

```bash
pip install pyinstaller
python -m PyInstaller --onefile --windowed --icon=icon.ico --name "FAISQL-Fluent" --add-data "crypto_utils.py;." --hidden-import qfluentwidgets --hidden-import PySide6.QtSvg --hidden-import PySide6.QtSvgWidgets fai-sql-fluent/main.py
```

O executável será gerado em `dist/FAISQL-Fluent.exe`.

> **⚠️ Importante**: Execute o comando na raiz do projeto (`FAI-SQL/`), não dentro de `fai-sql-fluent/`.

---

## 🛠️ Tecnologias

| Componente | Versão Clássica | Versão Fluent |
|------------|-----------------|---------------|
| Linguagem | Python 3.x | Python 3.10+ |
| Interface | Tkinter | PySide6 + pyside6-fluent-widgets |
| Banco de Dados | pyodbc (SQL Server) | pyodbc (SQL Server) |
| Exportação Excel | openpyxl | openpyxl |
| Criptografia | cryptography (Fernet/AES-256) | cryptography (Fernet/AES-256) |

---

## 🔐 Segurança

Todos os dados são protegidos por criptografia AES-256:

- **Primeira execução**: Crie uma senha mestra (mínimo 4 caracteres)
- **Próximas execuções**: Digite a senha para acessar
- **Esqueceu a senha?**: Clique no link na tela de login para resetar (apaga todos os dados)

### O que é protegido

- ✅ Conexões de banco de dados (incluindo senhas)
- ✅ Comandos SQL salvos
- ✅ Histórico de execuções

> **⚠️ Importante**: Não há recuperação de senha. Se esquecer, todos os dados serão perdidos.

---

## 📜 Licença

Este projeto está sob licença MIT.

---

## 🤝 Contribuição

Contribuições são bem-vindas! Abra uma issue ou envie um pull request.
