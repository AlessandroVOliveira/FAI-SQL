import tkinter as tk
from tkinter import ttk, messagebox, font  # Importando o módulo font
import pyodbc  # Biblioteca para conexão com SQL Server
import json    # Biblioteca para manipulação de arquivos JSON
import os      # Biblioteca para verificar se o arquivo existe

def sair():
    root.quit()

def abrir_tela_conexao():
    # Cria uma nova janela para a tela de conexão
    tela_conexao = tk.Toplevel(root)
    tela_conexao.title("Configuração de Conexão")
    tela_conexao.geometry("400x400")  # Tamanho da janela de conexão

    # Função para carregar os dados do arquivo config.json (se existir)
    def carregar_dados():
        if os.path.exists("config.json"):
            with open("config.json", "r") as arquivo:
                dados = json.load(arquivo)
                entry_ip.insert(0, dados.get("ip", ""))
                entry_usuario.insert(0, dados.get("usuario", ""))
                entry_senha.insert(0, dados.get("senha", ""))
                entry_banco.insert(0, dados.get("banco", ""))

    # Função para o botão "Testar"
    def testar_conexao():
        # Obtém os valores dos campos
        ip = entry_ip.get()
        usuario = entry_usuario.get()
        senha = entry_senha.get()
        banco = entry_banco.get()

        # Monta a string de conexão
        string_conexao = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={ip};"
            f"DATABASE={banco};"
            f"UID={usuario};"
            f"PWD={senha};"
        )

        try:
            # Tenta estabelecer a conexão
            conexao = pyodbc.connect(string_conexao)
            conexao.close()  # Fecha a conexão
            messagebox.showinfo("Testar Conexão", "Conexão bem-sucedida!")
        except Exception as e:
            # Exibe mensagem de erro se a conexão falhar
            messagebox.showerror("Testar Conexão", f"Conexão falhou: {str(e)}")

    # Função para o botão "Salvar"
    def salvar_conexao():
        # Obtém os valores dos campos
        ip = entry_ip.get()
        usuario = entry_usuario.get()
        senha = entry_senha.get()
        banco = entry_banco.get()

        # Cria um dicionário com os dados
        dados = {
            "ip": ip,
            "usuario": usuario,
            "senha": senha,
            "banco": banco
        }

        # Salva os dados em um arquivo JSON
        with open("config.json", "w") as arquivo:
            json.dump(dados, arquivo, indent=4)  # indent=4 para formatação legível

        messagebox.showinfo("Salvar Conexão", "Configurações salvas com sucesso!")
        tela_conexao.destroy()  # Fecha a tela de conexão após salvar

    # Label e Entry para IP de Conexão
    label_ip = tk.Label(tela_conexao, text="IP Conexão:", font=("Arial", 12))
    label_ip.pack(pady=5)
    entry_ip = tk.Entry(tela_conexao, width=30)
    entry_ip.pack(pady=5)

    # Label e Entry para Usuário
    label_usuario = tk.Label(tela_conexao, text="Usuário:", font=("Arial", 12))
    label_usuario.pack(pady=5)
    entry_usuario = tk.Entry(tela_conexao, width=30)
    entry_usuario.pack(pady=5)

    # Label e Entry para Senha
    label_senha = tk.Label(tela_conexao, text="Senha:", font=("Arial", 12))
    label_senha.pack(pady=5)
    entry_senha = tk.Entry(tela_conexao, width=30, show="*")  # Mostra asteriscos para a senha
    entry_senha.pack(pady=5)

    # Label e Entry para Banco de Dados
    label_banco = tk.Label(tela_conexao, text="Banco de Dados:", font=("Arial", 12))
    label_banco.pack(pady=5)
    entry_banco = tk.Entry(tela_conexao, width=30)
    entry_banco.pack(pady=5)

    # Botão "Testar"
    botao_testar = tk.Button(tela_conexao, text="Testar", command=testar_conexao, width=15)
    botao_testar.pack(pady=10)

    # Botão "Salvar"
    botao_salvar = tk.Button(tela_conexao, text="Salvar", command=salvar_conexao, width=15)
    botao_salvar.pack(pady=10)

    # Carrega os dados do arquivo config.json (se existir)
    carregar_dados()

def abrir_tela_adicionar_comando():
    # Cria uma nova janela para a tela de adicionar comando
    tela_adicionar_comando = tk.Toplevel(root)
    tela_adicionar_comando.title("Adicionar Comando")
    tela_adicionar_comando.geometry("900x600")  # Tamanho da janela de adicionar comando

    # Função para o botão "Salvar"
    def salvar_comando():
        nome_comando = entry_nome_comando.get()
        comando = caixa_texto_comando.get("1.0", tk.END).strip()  # Obtém o texto da caixa de texto

        # Verifica se os campos estão preenchidos
        if not nome_comando or not comando:
            messagebox.showwarning("Atenção", "Preencha todos os campos!")
            return

        # Cria o dicionário do comando
        novo_comando = {
            "nome": nome_comando,
            "comando": comando
        }

        # Carrega os comandos existentes ou cria uma lista vazia
        if os.path.exists("comandos.json"):
            with open("comandos.json", "r") as arquivo:
                comandos = json.load(arquivo)
        else:
            comandos = []

        # Adiciona o novo comando à lista
        comandos.append(novo_comando)

        # Salva a lista atualizada no arquivo JSON
        with open("comandos.json", "w") as arquivo:
            json.dump(comandos, arquivo, indent=4)  # indent=4 para formatação legível

        messagebox.showinfo("Salvar Comando", "Comando salvo com sucesso!")
        tela_adicionar_comando.destroy()  # Fecha a tela de adicionar comando após salvar

        # Atualiza o Combobox na tela principal
        carregar_comandos_no_combobox()

    # Label e Entry para Nome do Comando
    label_nome_comando = tk.Label(tela_adicionar_comando, text="Nome do Comando:", font=("Arial", 12))
    label_nome_comando.pack(pady=5)
    entry_nome_comando = tk.Entry(tela_adicionar_comando, width=50)
    entry_nome_comando.pack(pady=5)

    # Label e Text para Comando
    label_comando = tk.Label(tela_adicionar_comando, text="Comando:", font=("Arial", 12))
    label_comando.pack(pady=5)
    caixa_texto_comando = tk.Text(tela_adicionar_comando, height=10, wrap=tk.WORD)  # Altura para 10 linhas, quebra de linha por palavra
    caixa_texto_comando.pack(pady=5, fill=tk.BOTH, expand=True)  # Expande para ocupar o espaço disponível

    # Botão "Salvar"
    botao_salvar = tk.Button(tela_adicionar_comando, text="Salvar", command=salvar_comando, width=15)
    botao_salvar.pack(pady=10)

def carregar_comandos_no_combobox():
    # Verifica se o arquivo comandos.json existe
    if os.path.exists("comandos.json"):
        with open("comandos.json", "r") as arquivo:
            comandos = json.load(arquivo)
            # Extrai os nomes dos comandos
            nomes_comandos = [comando["nome"] for comando in comandos]
            # Atualiza os valores do Combobox
            combo_comandos["values"] = nomes_comandos
    else:
        combo_comandos["values"] = []  # Se o arquivo não existir, limpa o Combobox

def ao_selecionar_comando(event):
    # Obtém o nome do comando selecionado
    nome_comando_selecionado = combo_comandos.get()

    # Verifica se o arquivo comandos.json existe
    if os.path.exists("comandos.json"):
        with open("comandos.json", "r") as arquivo:
            comandos = json.load(arquivo)
            # Procura o comando selecionado
            for comando in comandos:
                if comando["nome"] == nome_comando_selecionado:
                    # Exibe o comando na caixa de texto
                    caixa_texto_comando.delete("1.0", tk.END)  # Limpa o conteúdo atual
                    caixa_texto_comando.insert("1.0", comando["comando"])  # Insere o novo conteúdo
                    break

def executar_comando():
    # Obtém o comando da caixa de texto
    comando = caixa_texto_comando.get("1.0", tk.END).strip()

    # Verifica se o comando está vazio
    if not comando:
        messagebox.showwarning("Atenção", "Nenhum comando para executar!")
        return

    # Verifica se o arquivo config.json existe
    if not os.path.exists("config.json"):
        messagebox.showerror("Erro", "Arquivo de configuração não encontrado!")
        return

    # Carrega os dados de conexão do arquivo config.json
    with open("config.json", "r") as arquivo:
        dados_conexao = json.load(arquivo)

    # Monta a string de conexão
    string_conexao = (
        f"DRIVER={{SQL Server}};"
        f"SERVER={dados_conexao['ip']};"
        f"DATABASE={dados_conexao['banco']};"
        f"UID={dados_conexao['usuario']};"
        f"PWD={dados_conexao['senha']};"
    )

    try:
        # Conecta ao banco de dados
        conexao = pyodbc.connect(string_conexao)
        cursor = conexao.cursor()

        # Executa o comando SQL
        cursor.execute(comando)

        # Obtém os resultados (se for um SELECT)
        if comando.strip().upper().startswith("SELECT"):
            resultados = cursor.fetchall()
            # Limpa a Treeview antes de adicionar novos dados
            for row in treeview.get_children():
                treeview.delete(row)
            # Obtém os nomes das colunas
            colunas = [column[0] for column in cursor.description]
            # Configura as colunas da Treeview
            treeview["columns"] = colunas
            treeview.heading("#0", text="ID", anchor="w")  # Coluna de índice
            treeview.column("#0", width=50, stretch=False)  # Tamanho fixo para a coluna de índice
            for coluna in colunas:
                treeview.heading(coluna, text=coluna, anchor="w")
                treeview.column(coluna, width=100, stretch=True)  # Tamanho inicial de 100px
            # Insere os dados na Treeview
            for i, linha in enumerate(resultados):
                # Substitui valores nulos por "NULL"
                valores_formatados = [str(valor) if valor is not None else "NULL" for valor in linha]
                treeview.insert("", "end", text=str(i + 1), values=valores_formatados)
            # Ajusta o tamanho das colunas ao conteúdo
            fonte = font.Font()  # Cria uma instância de fonte para medir o tamanho do texto
            for coluna in colunas:
                coluna_index = colunas.index(coluna)
                largura_cabecalho = fonte.measure(coluna)
                largura_maxima = largura_cabecalho
                for linha in resultados:
                    valor = str(linha[coluna_index]) if linha[coluna_index] is not None else "NULL"
                    largura_maxima = max(largura_maxima, fonte.measure(valor))
                treeview.column(coluna, width=largura_maxima)
        else:
            # Confirma a execução de comandos que não são SELECT (INSERT, UPDATE, DELETE, etc.)
            conexao.commit()
            messagebox.showinfo("Sucesso", "Comando executado com sucesso!")

        # Fecha a conexão
        cursor.close()
        conexao.close()

    except Exception as e:
        # Exibe mensagem de erro se a execução falhar
        messagebox.showerror("Erro", f"Falha ao executar o comando: {str(e)}")

# Cria a janela principal
root = tk.Tk()
root.title("FAI-SQL")



# Define o tamanho inicial da janela (900x800)
largura = 900
altura = 800

# Obtém as dimensões da tela
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

# Calcula a posição x e y para centralizar a janela
pos_x = (largura_tela // 2) - (largura // 2)
pos_y = (altura_tela // 2) - (altura // 2)

# Define a geometria da janela (tamanho + posição)
root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

# Permite redimensionar a janela (horizontal e vertical)
root.resizable(True, True)

# Cria a barra de menus
menubar = tk.Menu(root)

# Cria o menu "Arquivo"
menu_arquivo = tk.Menu(menubar, tearoff=0)
menu_arquivo.add_command(label="Sair", command=sair)
menubar.add_cascade(label="Arquivo", menu=menu_arquivo)

# Cria o menu "Sistema"
menu_sistema = tk.Menu(menubar, tearoff=0)
menu_sistema.add_command(label="Conexão", command=abrir_tela_conexao)
menu_sistema.add_command(label="Adicionar Comando", command=abrir_tela_adicionar_comando)
menubar.add_cascade(label="Sistema", menu=menu_sistema)

# Configura a barra de menus na janela principal
root.config(menu=menubar)

# Adiciona um Label "Comandos:" no canto superior esquerdo
label_comandos = tk.Label(root, text="Comandos:", font=("Arial", 12))
label_comandos.place(x=10, y=10)  # Posiciona o label em (x=10, y=10)

# Cria um Combobox (menu dropdown moderno)
combo_comandos = ttk.Combobox(root, state="readonly")  # state="readonly" impede a digitação
combo_comandos.place(x=10, y=40, width=200)  # Posiciona o Combobox abaixo do label

# Vincula a função ao evento de seleção no Combobox
combo_comandos.bind("<<ComboboxSelected>>", ao_selecionar_comando)

# Carrega os comandos no Combobox ao iniciar a aplicação
carregar_comandos_no_combobox()

# Adiciona a label "Comando Selecionado:"
label_comando_selecionado = tk.Label(root, text="Comando Selecionado:", font=("Arial", 12))
label_comando_selecionado.place(x=10, y=80)  # Posiciona abaixo do Combobox

# Adiciona uma caixa de texto (Text) para o comando selecionado
caixa_texto_comando = tk.Text(root, height=10, wrap=tk.WORD)  # Altura para 10 linhas, quebra de linha por palavra
caixa_texto_comando.place(x=10, y=110, relwidth=0.97)  # Posiciona abaixo da label e ocupa 97% da largura

# Adiciona um botão "Executar"
botao_executar = tk.Button(root, text="Executar", command=executar_comando, width=15)
botao_executar.place(x=10, y=280)  # Posiciona abaixo da caixa de texto do comando

# Adiciona uma Treeview para exibir os resultados
treeview = ttk.Treeview(root)
treeview.place(x=10, y=320, relwidth=0.97, height=400)  # Posiciona abaixo do botão "Executar"

# Adiciona uma barra de rolagem vertical à Treeview
scrollbar_vertical = ttk.Scrollbar(treeview, orient="vertical", command=treeview.yview)
scrollbar_vertical.pack(side = 'right', fill='y')
treeview.configure(yscrollcommand=scrollbar_vertical.set)

# Adiciona uma barra de rolagem horizontal à Treeview
scrollbar_horizontal = ttk.Scrollbar(treeview, orient="horizontal", command=treeview.xview)
scrollbar_horizontal.pack(side = 'bottom', fill='x')
treeview.configure(xscrollcommand=scrollbar_horizontal.set)

# Inicia o loop principal da interface gráfica
root.mainloop()