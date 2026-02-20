"""
Página de Gerenciamento de Comandos Salvos — FAI-SQL Fluent
CRUD de comandos SQL + botão para usar no editor.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QHBoxLayout,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    BodyLabel,
    CardWidget,
    FluentIcon as FIF,
    InfoBar,
    InfoBarPosition,
    LineEdit,
    ListWidget,
    MessageBox,
    PrimaryPushButton,
    PushButton,
    StrongBodyLabel,
    SubtitleLabel,
    isDarkTheme,
)

import crypto_utils as crypto
from utils.database import carregar_colunas_tabela, carregar_objetos_banco
from utils.sql_highlighter import SqlHighlighter
from utils.sql_completer import SqlCompleter


class CommandsPage(QWidget):
    """Página para gerenciar comandos SQL salvos."""

    usar_comando_signal = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("commands_page")
        self._setup_ui()
        self._atualizar_lista()
        self._carregar_schema_banco()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(15)

        titulo = SubtitleLabel("Comandos Salvos")
        layout.addWidget(titulo)

        content = QHBoxLayout()
        content.setSpacing(20)

        # === Lista de Comandos (esquerda) ===
        left_card = CardWidget()
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(15, 15, 15, 15)

        left_title = StrongBodyLabel("Comandos")
        left_layout.addWidget(left_title)

        self._lista = ListWidget()
        self._lista.setMinimumWidth(250)
        self._lista.currentRowChanged.connect(self._ao_selecionar)
        left_layout.addWidget(self._lista, 1)

        btn_row = QHBoxLayout()
        btn_usar = PrimaryPushButton(FIF.EDIT, "Editar Comando")
        btn_usar.clicked.connect(self._usar_comando)
        btn_row.addWidget(btn_usar)

        btn_excluir = PushButton(FIF.DELETE, "Excluir")
        btn_excluir.clicked.connect(self._excluir_comando)
        btn_row.addWidget(btn_excluir)
        left_layout.addLayout(btn_row)

        content.addWidget(left_card, 2)

        # === Formulário (direita) ===
        right_card = CardWidget()
        right_layout = QVBoxLayout(right_card)
        right_layout.setContentsMargins(20, 20, 20, 20)
        right_layout.setSpacing(12)

        right_title = StrongBodyLabel("Editar / Novo Comando")
        right_layout.addWidget(right_title)

        lbl_nome = BodyLabel("Nome do Comando")
        right_layout.addWidget(lbl_nome)

        self._entry_nome = LineEdit(self)
        self._entry_nome.setFixedHeight(36)
        self._entry_nome.setClearButtonEnabled(True)
        right_layout.addWidget(self._entry_nome)

        lbl_sql = BodyLabel("SQL")
        right_layout.addWidget(lbl_sql)

        self._editor_sql = QPlainTextEdit()
        self._editor_sql.setFont(QFont("Consolas", 11))
        self._editor_sql.setPlaceholderText("Digite o SQL aqui...")
        self._editor_sql.setStyleSheet(self._editor_stylesheet())
        right_layout.addWidget(self._editor_sql, 1)

        self._highlighter = SqlHighlighter(
            self._editor_sql.document(), dark=isDarkTheme()
        )

        # Autocomplete
        self._completer = SqlCompleter(self._editor_sql)

        form_btns = QHBoxLayout()

        btn_salvar = PrimaryPushButton(FIF.SAVE, "Salvar")
        btn_salvar.setFixedHeight(36)
        btn_salvar.clicked.connect(self._salvar_comando)
        form_btns.addWidget(btn_salvar)

        btn_novo = PushButton(FIF.ADD, "Novo")
        btn_novo.setFixedHeight(36)
        btn_novo.clicked.connect(self._limpar_campos)
        form_btns.addWidget(btn_novo)

        right_layout.addLayout(form_btns)

        content.addWidget(right_card, 3)
        layout.addLayout(content, 1)

    # ------------------------------------------------------------------

    def _carregar_comandos(self) -> list[dict]:
        try:
            comandos = crypto.ler_arquivo_seguro(crypto.ARQUIVO_COMANDOS)
            return comandos if comandos else []
        except Exception:
            return []

    def _atualizar_lista(self):
        self._lista.clear()
        for cmd in self._carregar_comandos():
            self._lista.addItem(cmd.get("nome", ""))

    def _ao_selecionar(self, row: int):
        if row < 0:
            return
        comandos = self._carregar_comandos()
        if row < len(comandos):
            self._entry_nome.setText(comandos[row].get("nome", ""))
            self._editor_sql.setPlainText(comandos[row].get("comando", ""))

    def _usar_comando(self):
        row = self._lista.currentRow()
        if row < 0:
            InfoBar.warning("Atenção", "Selecione um comando!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            return
        comandos = self._carregar_comandos()
        if row < len(comandos):
            self._entry_nome.setText(comandos[row].get("nome", ""))
            self._editor_sql.setPlainText(comandos[row].get("comando", ""))

    def _excluir_comando(self):
        row = self._lista.currentRow()
        if row < 0:
            InfoBar.warning("Atenção", "Selecione um comando!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            return

        box = MessageBox("Confirmar Exclusão", "Excluir este comando?", self)
        if not box.exec():
            return

        try:
            comandos = self._carregar_comandos()
            if row < len(comandos):
                del comandos[row]
                crypto.escrever_arquivo_seguro(crypto.ARQUIVO_COMANDOS, comandos)
                self._atualizar_lista()
                self._limpar_campos()
                InfoBar.success("Sucesso", "Comando excluído!", parent=self,
                                duration=3000, position=InfoBarPosition.TOP)
                self._notificar_editor()
        except Exception as e:
            InfoBar.error("Erro", str(e), parent=self, duration=4000,
                          position=InfoBarPosition.TOP)

    def _salvar_comando(self):
        nome = self._entry_nome.text().strip()
        sql = self._editor_sql.toPlainText().strip()

        if not nome or not sql:
            InfoBar.warning("Atenção", "Preencha nome e SQL!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            return

        try:
            comandos = self._carregar_comandos()

            encontrado = False
            for i, cmd in enumerate(comandos):
                if cmd.get("nome") == nome:
                    comandos[i] = {"nome": nome, "comando": sql}
                    encontrado = True
                    break
            if not encontrado:
                comandos.append({"nome": nome, "comando": sql})

            crypto.escrever_arquivo_seguro(crypto.ARQUIVO_COMANDOS, comandos)
            self._atualizar_lista()
            InfoBar.success("Sucesso", "Comando salvo!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            self._notificar_editor()
        except Exception as e:
            InfoBar.error("Erro", str(e), parent=self, duration=4000,
                          position=InfoBarPosition.TOP)

    def _limpar_campos(self):
        self._entry_nome.clear()
        self._editor_sql.clear()
        self._lista.clearSelection()

    @staticmethod
    def _editor_stylesheet() -> str:
        """Retorna stylesheet do editor SQL adaptado ao tema."""
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

    def _notificar_editor(self):
        """Notifica o editor para recarregar a lista de comandos."""
        main_window = self.window()
        if hasattr(main_window, "_editor_page"):
            main_window._editor_page.recarregar_comandos()

    def _carregar_schema_banco(self):
        """Carrega tabelas e colunas do banco ativo para autocomplete contextual."""
        conn = self._obter_conexao_ativa()
        if conn:
            objetos = carregar_objetos_banco(conn)
            self._completer.set_database_objects(objetos)
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

