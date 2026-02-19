"""
FAI-SQL Fluent — Janela Principal (FluentWindow)
Navegação lateral com 4 páginas: Editor, Conexões, Comandos, Histórico.
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import QApplication

from qfluentwidgets import (
    FluentIcon as FIF,
    FluentWindow,
    NavigationItemPosition,
    setTheme,
    Theme,
    isDarkTheme,
)

import crypto_utils as crypto

from pages.editor import EditorPage
from pages.connections import ConnectionsPage
from pages.commands import CommandsPage
from pages.history import HistoryPage


class MainWindow(FluentWindow):
    """Janela principal com navegação fluent."""

    # Signal para comunicação entre páginas
    comando_solicitado = Signal(str)  # envia SQL do histórico/comandos → editor

    def __init__(self):
        super().__init__()
        self._carregar_tema()
        self._setup_window()
        self._setup_pages()

    def _carregar_tema(self):
        """Carrega preferência de tema salva."""
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados and dados.get("tema") == "escuro":
                setTheme(Theme.DARK)
            else:
                setTheme(Theme.LIGHT)
        except Exception:
            setTheme(Theme.LIGHT)

    def _salvar_tema(self):
        """Salva preferência de tema."""
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados is None:
                dados = {}
            dados["tema"] = "escuro" if isDarkTheme() else "claro"
            crypto.escrever_arquivo_seguro(crypto.ARQUIVO_CONFIG, dados)
        except Exception:
            pass

    def _setup_window(self):
        self.setWindowTitle("FAI-SQL")
        self.resize(1100, 800)
        self.setMinimumSize(900, 650)

        # Centralizar
        screen = QApplication.primaryScreen().geometry()
        self.move(
            (screen.width() - self.width()) // 2,
            (screen.height() - self.height()) // 2,
        )

    def _setup_pages(self):
        # Criar páginas
        self._editor_page = EditorPage(self)
        self._connections_page = ConnectionsPage(self)
        self._commands_page = CommandsPage(self)
        self._history_page = HistoryPage(self)

        # Conectar signal para enviar comando ao editor
        self.comando_solicitado.connect(self._editor_page.set_sql_text)
        self._commands_page.usar_comando_signal.connect(self._usar_comando)
        self._history_page.usar_comando_signal.connect(self._usar_comando)

        # Adicionar páginas à navegação
        self.addSubInterface(self._editor_page, FIF.CODE, "Editor SQL")
        self.addSubInterface(self._connections_page, FIF.CONNECT, "Conexões")
        self.addSubInterface(self._commands_page, FIF.COMMAND_PROMPT, "Comandos")
        self.addSubInterface(self._history_page, FIF.HISTORY, "Histórico")

        # Botão de tema na parte inferior da navegação
        self.navigationInterface.addItem(
            routeKey="toggle_theme",
            icon=FIF.CONSTRACT,
            text="Alternar Tema",
            onClick=self._alternar_tema,
            selectable=False,
            position=NavigationItemPosition.BOTTOM,
        )

    def _usar_comando(self, sql: str):
        """Recebe SQL de outra página e envia para o editor."""
        self._editor_page.set_sql_text(sql)
        self.switchTo(self._editor_page)

    def _alternar_tema(self):
        if isDarkTheme():
            setTheme(Theme.LIGHT)
        else:
            setTheme(Theme.DARK)
        self._salvar_tema()
        # Atualizar syntax highlighter e cores dos editores
        self._editor_page.update_theme()
        self._commands_page._highlighter.set_dark(isDarkTheme())
        self._commands_page._editor_sql.setStyleSheet(
            self._commands_page._editor_stylesheet()
        )
