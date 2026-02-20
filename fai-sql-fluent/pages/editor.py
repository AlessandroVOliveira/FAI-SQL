"""
Página do Editor SQL — FAI-SQL Fluent
Editor principal com syntax highlighting, execução de queries e tabela de resultados.
"""

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPlainTextEdit,
    QSplitter,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    BodyLabel,
    CardWidget,
    ComboBox,
    FluentIcon as FIF,
    InfoBar,
    InfoBarPosition,
    MessageBox,
    PrimaryPushButton,
    PushButton,
    StrongBodyLabel,
    SubtitleLabel,
    TableWidget,
    isDarkTheme,
)

import crypto_utils as crypto
from utils.database import carregar_colunas_tabela, carregar_objetos_banco, executar_query, exportar_csv, exportar_excel
from utils.sql_highlighter import SqlHighlighter
from utils.sql_completer import SqlCompleter


class EditorPage(QWidget):
    """Página do editor SQL com resultados."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("editor_page")
        self._colunas_resultado = []
        self._dados_resultado = []
        self._setup_ui()
        self._carregar_comandos()
        QTimer.singleShot(200, self._atualizar_label_conexao)
        QTimer.singleShot(500, self._carregar_schema_banco)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(12)

        # --- Cabeçalho ---
        header = QHBoxLayout()
        titulo = SubtitleLabel("Editor SQL")
        header.addWidget(titulo)
        header.addStretch()

        self._label_conexao = BodyLabel("⚠️ Nenhuma conexão configurada")
        self._label_conexao.setStyleSheet("color: #888;")
        header.addWidget(self._label_conexao)
        layout.addLayout(header)

        # --- Splitter vertical: editor em cima, resultados em baixo ---
        splitter = QSplitter(Qt.Vertical)
        splitter.setHandleWidth(6)

        # === Parte superior: comandos + editor + botões ===
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(10)

        # Comandos salvos
        cmd_card = CardWidget()
        cmd_layout = QHBoxLayout(cmd_card)
        cmd_layout.setContentsMargins(15, 10, 15, 10)

        cmd_label = BodyLabel("Comando Salvo:")
        cmd_layout.addWidget(cmd_label)

        self._combo_comandos = ComboBox()
        self._combo_comandos.setMinimumWidth(300)
        self._combo_comandos.setPlaceholderText("Selecione um comando...")
        self._combo_comandos.currentIndexChanged.connect(self._ao_selecionar_comando)
        cmd_layout.addWidget(self._combo_comandos, 1)

        top_layout.addWidget(cmd_card)

        # Editor SQL
        editor_card = CardWidget()
        editor_layout = QVBoxLayout(editor_card)
        editor_layout.setContentsMargins(15, 12, 15, 12)

        editor_header = QHBoxLayout()
        editor_title = StrongBodyLabel("SQL")
        editor_header.addWidget(editor_title)
        editor_header.addStretch()
        editor_layout.addLayout(editor_header)

        self._editor = QPlainTextEdit()
        self._editor.setFont(QFont("Consolas", 11))
        self._editor.setPlaceholderText("Digite seu comando SQL aqui...")
        self._editor.setMinimumHeight(120)
        self._editor.setStyleSheet(self._editor_stylesheet())
        editor_layout.addWidget(self._editor)

        # Syntax highlighter
        self._highlighter = SqlHighlighter(
            self._editor.document(), dark=isDarkTheme()
        )

        # Autocomplete
        self._completer = SqlCompleter(self._editor)

        top_layout.addWidget(editor_card)

        # Botões de ação
        btn_layout = QHBoxLayout()

        btn_executar = PrimaryPushButton(FIF.PLAY, "Executar")
        btn_executar.setFixedHeight(36)
        btn_executar.clicked.connect(self._executar_comando)
        btn_layout.addWidget(btn_executar)

        btn_limpar = PushButton(FIF.DELETE, "Limpar")
        btn_limpar.setFixedHeight(36)
        btn_limpar.clicked.connect(self._limpar_tela)
        btn_layout.addWidget(btn_limpar)

        # Separador visual
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setFrameShadow(QFrame.Sunken)
        sep.setFixedWidth(2)
        btn_layout.addWidget(sep)

        btn_csv = PushButton(FIF.DOCUMENT, "CSV")
        btn_csv.setFixedHeight(36)
        btn_csv.clicked.connect(self._exportar_csv)
        btn_layout.addWidget(btn_csv)

        btn_excel = PushButton(FIF.BOOK_SHELF, "Excel")
        btn_excel.setFixedHeight(36)
        btn_excel.clicked.connect(self._exportar_excel)
        btn_layout.addWidget(btn_excel)

        btn_layout.addStretch()

        self._label_status = BodyLabel("")
        btn_layout.addWidget(self._label_status)

        top_layout.addLayout(btn_layout)
        splitter.addWidget(top_widget)

        # === Parte inferior: resultados ===
        results_widget = QWidget()
        results_layout = QVBoxLayout(results_widget)
        results_layout.setContentsMargins(0, 0, 0, 0)

        results_header = StrongBodyLabel("Resultados")
        results_layout.addWidget(results_header)

        self._tabela = TableWidget()
        self._tabela.setWordWrap(False)
        self._tabela.setAlternatingRowColors(True)
        self._tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._tabela.horizontalHeader().setStretchLastSection(True)
        self._tabela.verticalHeader().setVisible(True)
        self._tabela.setEditTriggers(TableWidget.NoEditTriggers)
        self._tabela.setSelectionBehavior(TableWidget.SelectRows)
        results_layout.addWidget(self._tabela)

        splitter.addWidget(results_widget)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)

        layout.addWidget(splitter, 1)

    # ------------------------------------------------------------------
    # Comandos salvos
    # ------------------------------------------------------------------

    def _carregar_comandos(self):
        self._combo_comandos.clear()
        self._combo_comandos.addItem("— Selecione —")
        try:
            comandos = crypto.ler_arquivo_seguro(crypto.ARQUIVO_COMANDOS)
            if comandos:
                for cmd in comandos:
                    self._combo_comandos.addItem(cmd.get("nome", ""))
        except Exception:
            pass

    def _ao_selecionar_comando(self, index: int):
        if index <= 0:
            return
        try:
            comandos = crypto.ler_arquivo_seguro(crypto.ARQUIVO_COMANDOS)
            if comandos and index - 1 < len(comandos):
                sql = comandos[index - 1].get("comando", "")
                self._editor.setPlainText(sql)
        except Exception:
            pass

    def recarregar_comandos(self):
        """Chamado externamente quando comandos são alterados."""
        self._carregar_comandos()

    # ------------------------------------------------------------------
    # Conexão ativa
    # ------------------------------------------------------------------

    def _atualizar_label_conexao(self):
        try:
            conn = self._obter_conexao_ativa()
            if conn:
                self._label_conexao.setText(
                    f"🔗 {conn.get('nome', 'Sem nome')} ({conn.get('banco', '')})"
                )
            else:
                self._label_conexao.setText("⚠️ Nenhuma conexão configurada")
        except Exception:
            pass

    def atualizar_conexao(self):
        """Chamado externo quando a conexão é alterada."""
        self._atualizar_label_conexao()
        self._carregar_schema_banco()

    def _carregar_schema_banco(self):
        """Carrega tabelas e colunas do banco ativo para autocomplete."""
        conn = self._obter_conexao_ativa()
        if conn:
            objetos = carregar_objetos_banco(conn)
            self._completer.set_database_objects(objetos)
            # Callback para carregar colunas de uma tabela sob demanda
            self._completer.set_column_loader(
                lambda tabela: carregar_colunas_tabela(conn, tabela)
            )

    @staticmethod
    def _obter_conexao_ativa():
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados:
                nome_ativa = dados.get("conexao_ativa", "")
                conexoes = dados.get("conexoes", [])
                for conn in conexoes:
                    if conn.get("nome") == nome_ativa:
                        return conn
        except Exception:
            pass
        return None

    # ------------------------------------------------------------------
    # Execução
    # ------------------------------------------------------------------

    # Comandos perigosos e suas descrições
    _COMANDOS_PERIGOSOS = [
        ("DROP",     "🚨 Este comando irá REMOVER um objeto do banco de dados permanentemente."),
        ("TRUNCATE", "🚨 Este comando irá APAGAR TODOS os registros da tabela."),
        ("DELETE",   "⚠️ Este comando irá EXCLUIR dados do banco."),
        ("UPDATE",   "⚠️ Este comando irá ALTERAR dados existentes no banco."),
        ("ALTER",    "⚠️ Este comando irá MODIFICAR a estrutura do banco de dados."),
    ]

    def _confirmar_comando_perigoso(self, comando: str) -> bool:
        """Exibe confirmação para comandos destrutivos. Retorna True se pode prosseguir."""
        cmd_upper = comando.strip().upper()

        for keyword, descricao in self._COMANDOS_PERIGOSOS:
            if cmd_upper.startswith(keyword):
                box = MessageBox(
                    f"Comando {keyword} detectado",
                    f"{descricao}\n\nDeseja realmente executar este comando?",
                    self,
                )
                box.yesButton.setText("Executar")
                box.cancelButton.setText("Cancelar")
                return bool(box.exec())

        return True

    def _executar_comando(self):
        comando = self._editor.toPlainText().strip()
        if not comando:
            InfoBar.warning("Atenção", "Nenhum comando para executar!",
                            parent=self, duration=3000, position=InfoBarPosition.TOP)
            return

        dados_conexao = self._obter_conexao_ativa()
        if not dados_conexao:
            InfoBar.error("Erro", "Configure uma conexão primeiro!",
                          parent=self, duration=3000, position=InfoBarPosition.TOP)
            return

        # Verificar comandos perigosos
        if not self._confirmar_comando_perigoso(comando):
            return

        resultado = executar_query(dados_conexao, comando)

        if resultado["sucesso"]:
            if resultado["tipo"] == "select":
                self._exibir_resultados(resultado["colunas"], resultado["dados"])
                self._label_status.setText(f"✅ {resultado['mensagem']}")
            else:
                self._label_status.setText("✅ " + resultado["mensagem"])
                InfoBar.success("Sucesso", resultado["mensagem"],
                                parent=self, duration=3000, position=InfoBarPosition.TOP)

            self._salvar_historico(
                comando, True,
                resultado["mensagem"],
                resultado["registros"],
                dados_conexao.get("nome", ""),
            )
        else:
            self._label_status.setText("❌ Erro na execução")
            InfoBar.error("Erro", resultado["mensagem"],
                          parent=self, duration=5000, position=InfoBarPosition.TOP)
            self._salvar_historico(comando, False, resultado["mensagem"][:200], 0, dados_conexao.get("nome", ""))

    def _exibir_resultados(self, colunas: list[str], dados: list[list[str]]):
        self._colunas_resultado = colunas
        self._dados_resultado = dados

        self._tabela.setColumnCount(len(colunas))
        self._tabela.setRowCount(len(dados))
        self._tabela.setHorizontalHeaderLabels(colunas)

        for row_idx, linha in enumerate(dados):
            for col_idx, valor in enumerate(linha):
                item = QTableWidgetItem(str(valor))
                self._tabela.setItem(row_idx, col_idx, item)

        # Redimensionar colunas ao conteúdo (com limite)
        self._tabela.resizeColumnsToContents()
        header = self._tabela.horizontalHeader()
        for i in range(len(colunas)):
            if header.sectionSize(i) > 300:
                header.resizeSection(i, 300)

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def _exportar_csv(self):
        if not self._colunas_resultado:
            InfoBar.warning("Atenção", "Não há dados para exportar!",
                            parent=self, duration=3000, position=InfoBarPosition.TOP)
            return

        caminho, _ = QFileDialog.getSaveFileName(
            self, "Salvar como CSV",
            f"exportacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "Arquivo CSV (*.csv);;Todos (*.*)",
        )
        if not caminho:
            return

        try:
            exportar_csv(caminho, self._colunas_resultado, self._dados_resultado)
            self._label_status.setText(f"✅ Exportado: {os.path.basename(caminho)}")
            InfoBar.success("Sucesso", f"CSV salvo em: {caminho}",
                            parent=self, duration=3000, position=InfoBarPosition.TOP)
        except Exception as e:
            InfoBar.error("Erro", str(e), parent=self, duration=4000,
                          position=InfoBarPosition.TOP)

    def _exportar_excel(self):
        if not self._colunas_resultado:
            InfoBar.warning("Atenção", "Não há dados para exportar!",
                            parent=self, duration=3000, position=InfoBarPosition.TOP)
            return

        caminho, _ = QFileDialog.getSaveFileName(
            self, "Salvar como Excel",
            f"exportacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Arquivo Excel (*.xlsx);;Todos (*.*)",
        )
        if not caminho:
            return

        try:
            exportar_excel(caminho, self._colunas_resultado, self._dados_resultado)
            self._label_status.setText(f"✅ Exportado: {os.path.basename(caminho)}")
            InfoBar.success("Sucesso", f"Excel salvo em: {caminho}",
                            parent=self, duration=3000, position=InfoBarPosition.TOP)
        except Exception as e:
            InfoBar.error("Erro", str(e), parent=self, duration=4000,
                          position=InfoBarPosition.TOP)

    # ------------------------------------------------------------------
    # Limpar
    # ------------------------------------------------------------------

    def _limpar_tela(self):
        self._editor.clear()
        self._tabela.clearContents()
        self._tabela.setRowCount(0)
        self._tabela.setColumnCount(0)
        self._combo_comandos.setCurrentIndex(0)
        self._label_status.setText("")
        self._colunas_resultado = []
        self._dados_resultado = []

    # ------------------------------------------------------------------
    # Histórico
    # ------------------------------------------------------------------

    @staticmethod
    def _salvar_historico(comando: str, sucesso: bool, mensagem: str = "", registros: int = 0, conexao: str = ""):
        MAX_HISTORICO = 100
        try:
            historico = crypto.ler_arquivo_seguro(crypto.ARQUIVO_HISTORICO)
            if historico is None:
                historico = []
        except Exception:
            historico = []

        nova_entrada = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "comando": comando[:500],
            "sucesso": sucesso,
            "mensagem": mensagem,
            "registros": registros,
            "conexao": conexao,
        }
        historico.insert(0, nova_entrada)
        historico = historico[:MAX_HISTORICO]
        try:
            crypto.escrever_arquivo_seguro(crypto.ARQUIVO_HISTORICO, historico)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Métodos públicos
    # ------------------------------------------------------------------

    def set_sql_text(self, sql: str):
        """Define o texto do editor (chamado por outras páginas)."""
        self._editor.setPlainText(sql)

    @staticmethod
    def _editor_stylesheet() -> str:
        """Retorna stylesheet do editor adaptado ao tema."""
        color = "#e0e0e0" if isDarkTheme() else "#1a1a1a"
        bg = "#2d2d2d" if isDarkTheme() else "#ffffff"
        border = "#555" if isDarkTheme() else "#ccc"
        return f"""
            QPlainTextEdit {{
                border: 1px solid {border};
                border-radius: 6px;
                padding: 8px;
                color: {color};
                background-color: {bg};
            }}
        """

    def update_theme(self):
        """Atualiza o syntax highlighter e cores do editor com o tema atual."""
        self._highlighter.set_dark(isDarkTheme())
        self._editor.setStyleSheet(self._editor_stylesheet())
