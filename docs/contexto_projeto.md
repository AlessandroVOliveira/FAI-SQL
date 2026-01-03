# Documento de Contexto - FAI-SQL

**Data de Criação:** 02/01/2026  
**Versão:** 1.0

---

## 1. Visão Geral do Projeto

**FAI-SQL** é uma aplicação desktop desenvolvida em Python com interface gráfica Tkinter, projetada para auxiliar desenvolvedores, DBAs e analistas a armazenarem e executarem comandos SQL de uso frequente de forma rápida e eficiente.

O software elimina a necessidade de procurar ou digitar comandos repetitivos, permitindo a execução direta no banco de dados configurado e exibindo os resultados em tempo real.

---

## 2. Stack Tecnológica

| Componente | Tecnologia |
|------------|------------|
| Linguagem | Python 3.x |
| Interface Gráfica | Tkinter (biblioteca padrão Python) |
| Conexão com Banco | pyodbc (SQL Server) |
| Persistência de Dados | Arquivos JSON locais |
| Build/Distribuição | PyInstaller |

---

## 3. Estrutura do Projeto

```
FAI-SQL/
├── main.py              # Arquivo principal da aplicação (356 linhas)
├── config.json          # Configurações de conexão com o banco
├── comandos.json        # Lista de comandos SQL salvos
├── main.spec            # Especificação do PyInstaller
├── icon.ico             # Ícone da aplicação
├── image.png            # Screenshot - Tela inicial
├── image-1.png          # Screenshot - Tela de conexão
├── image-2.png          # Screenshot - Tela de adicionar comandos
├── README.MD            # Documentação do projeto
├── build/               # Arquivos de build do PyInstaller
├── dist/                # Executável distribuível (FAISQL.exe)
└── docs/                # Relatórios de implementação
```

---

## 4. Arquitetura da Aplicação

### 4.1 Componentes Principais

A aplicação é construída como um único arquivo Python (`main.py`) com as seguintes responsabilidades:

1. **Interface Gráfica Principal**
   - Janela principal (900x800 pixels) centralizada na tela
   - Barra de menus (Arquivo, Sistema)
   - Combobox para seleção de comandos salvos
   - Área de texto para visualização/edição de comandos
   - Treeview para exibição de resultados com scrollbars

2. **Tela de Configuração de Conexão** (`abrir_tela_conexao()`)
   - Campos: IP, Usuário, Senha, Banco de Dados
   - Funcionalidade de testar conexão
   - Salva configurações em `config.json`

3. **Tela de Adicionar Comando** (`abrir_tela_adicionar_comando()`)
   - Campos: Nome do Comando, Comando SQL
   - Salva comandos em `comandos.json`

4. **Motor de Execução SQL** (`executar_comando()`)
   - Conecta ao SQL Server usando pyodbc
   - Executa comandos SELECT e exibe resultados na Treeview
   - Executa comandos INSERT/UPDATE/DELETE com commit automático
   - Tratamento de erros com mensagens informativas

### 4.2 Fluxo de Dados

```
┌─────────────────┐     ┌─────────────────┐
│   config.json   │────▶│   Conexão DB    │
└─────────────────┘     └────────┬────────┘
                                 │
┌─────────────────┐              ▼
│ comandos.json   │────▶┌─────────────────┐
└─────────────────┘     │   Execução SQL  │
                        └────────┬────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │  Treeview/Msgs  │
                        └─────────────────┘
```

---

## 5. Arquivos de Configuração

### 5.1 config.json
Armazena as credenciais de conexão com o banco de dados:
```json
{
    "ip": "localhost,porta",
    "usuario": "usuario_db",
    "senha": "senha_db",
    "banco": "nome_banco"
}
```

### 5.2 comandos.json
Lista de comandos SQL salvos para reutilização:
```json
[
    {
        "nome": "Nome descritivo do comando",
        "comando": "SELECT * FROM tabela"
    }
]
```

---

## 6. Funcionalidades Implementadas

- [x] Conexão com SQL Server via ODBC
- [x] Teste de conexão antes de salvar
- [x] Armazenamento persistente de comandos SQL
- [x] Seleção de comandos via dropdown
- [x] Execução de comandos SELECT com exibição em tabela
- [x] Execução de comandos DML (INSERT, UPDATE, DELETE)
- [x] Ajuste automático de largura de colunas
- [x] Scrollbars horizontal e vertical na tabela de resultados
- [x] Tratamento de valores NULL
- [x] Interface responsiva e redimensionável
- [x] Build para executável .exe via PyInstaller

---

## 7. Dependências

```bash
pip install pyodbc      # Conexão com SQL Server
pip install pyinstaller # Build do executável (somente dev)
```

**Requisito adicional:** Driver ODBC do SQL Server instalado no sistema.

---

## 8. Como Executar

### Desenvolvimento
```bash
python main.py
```

### Build para Produção
```bash
pyinstaller --onefile --windowed --name "FAISQL" main.py
```

O executável será gerado em `dist/FAISQL.exe`.

---

## 9. Limitações Conhecidas

1. **Banco de Dados:** Suporta apenas SQL Server (driver fixo "SQL Server")
2. **Segurança:** Credenciais armazenadas em texto plano no `config.json`
3. **Edição de Comandos:** Não é possível editar ou excluir comandos salvos pela interface
4. **Layout Fixo:** Alguns elementos usam posicionamento absoluto (`place()`)
5. **Sem validação:** Não há validação de sintaxe SQL antes da execução

---

## 10. Possíveis Melhorias Futuras


- [ ] Criação de interface moderna com Tkinter
- [ ] Edição e exclusão de comandos salvos
- [ ] Exportação de resultados (CSV, Excel)
- [ ] Histórico de comandos executados
- [ ] Syntax highlighting para SQL
- [ ] Suporte a múltiplas conexões/ambientes
- [ ] Undo/Redo na área de texto

---

*Este documento serve como referência para desenvolvimento e manutenção do projeto FAI-SQL.*
