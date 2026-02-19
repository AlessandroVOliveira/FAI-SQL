"""
Página de Histórico de Execuções — FAI-SQL Fluent
Visualização do histórico com possibilidade de reutilizar comandos.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QHBoxLayout,
    QHeaderView,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    BodyLabel,
    CardWidget,
    FluentIcon as FIF,
    InfoBar,
    InfoBarPosition,
    MessageBox,
    PrimaryPushButton,
    PushButton,
    StrongBodyLabel,
    SubtitleLabel,
    TableWidget,
)

import crypto_utils as crypto


class HistoryPage(QWidget):
    """Página de histórico de execuções SQL."""

    usar_comando_signal = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("history_page")
        self._historico: list[dict] = []
        self._setup_ui()
        self._carregar_historico()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(15)

        # Cabeçalho
        header = QHBoxLayout()
        titulo = SubtitleLabel("Histórico de Execuções")
        header.addWidget(titulo)
        header.addStretch()

        btn_atualizar = PushButton(FIF.SYNC, "Atualizar")
        btn_atualizar.clicked.connect(self._carregar_historico)
        header.addWidget(btn_atualizar)
        layout.addLayout(header)

        # Tabela
        card = CardWidget()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(15, 15, 15, 15)

        self._tabela = TableWidget()
        self._tabela.setColumnCount(4)
        self._tabela.setHorizontalHeaderLabels(["Data/Hora", "Status", "Registros", "Comando"])
        self._tabela.setWordWrap(False)
        self._tabela.setAlternatingRowColors(True)
        self._tabela.setEditTriggers(TableWidget.NoEditTriggers)
        self._tabela.setSelectionBehavior(TableWidget.SelectRows)
        self._tabela.verticalHeader().setVisible(False)

        header_view = self._tabela.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.Fixed)
        header_view.setSectionResizeMode(1, QHeaderView.Fixed)
        header_view.setSectionResizeMode(2, QHeaderView.Fixed)
        header_view.setSectionResizeMode(3, QHeaderView.Stretch)
        self._tabela.setColumnWidth(0, 160)
        self._tabela.setColumnWidth(1, 70)
        self._tabela.setColumnWidth(2, 90)

        self._tabela.doubleClicked.connect(self._usar_comando)

        card_layout.addWidget(self._tabela)
        layout.addWidget(card, 1)

        # Botões
        btn_layout = QHBoxLayout()
        btn_usar = PrimaryPushButton(FIF.PASTE, "Usar Comando Selecionado")
        btn_usar.clicked.connect(self._usar_comando)
        btn_layout.addWidget(btn_usar)

        btn_layout.addStretch()

        btn_limpar = PushButton(FIF.DELETE, "Limpar Histórico")
        btn_limpar.clicked.connect(self._limpar_historico)
        btn_layout.addWidget(btn_limpar)

        layout.addLayout(btn_layout)

    # ------------------------------------------------------------------

    def _carregar_historico(self):
        try:
            hist = crypto.ler_arquivo_seguro(crypto.ARQUIVO_HISTORICO)
            self._historico = hist if hist else []
        except Exception:
            self._historico = []

        self._tabela.setRowCount(len(self._historico))
        for row, item in enumerate(self._historico):
            self._tabela.setItem(row, 0, QTableWidgetItem(item.get("timestamp", "")))

            status = "✅" if item.get("sucesso", False) else "❌"
            status_item = QTableWidgetItem(status)
            status_item.setTextAlignment(Qt.AlignCenter)
            self._tabela.setItem(row, 1, status_item)

            reg_item = QTableWidgetItem(str(item.get("registros", "-")))
            reg_item.setTextAlignment(Qt.AlignCenter)
            self._tabela.setItem(row, 2, reg_item)

            cmd = item.get("comando", "")
            resumo = cmd[:120].replace("\n", " ")
            if len(cmd) > 120:
                resumo += "..."
            self._tabela.setItem(row, 3, QTableWidgetItem(resumo))

    def _usar_comando(self):
        row = self._tabela.currentRow()
        if row < 0:
            InfoBar.warning("Atenção", "Selecione um comando!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            return
        if row < len(self._historico):
            sql = self._historico[row].get("comando", "")
            self.usar_comando_signal.emit(sql)

    def _limpar_historico(self):
        box = MessageBox("Confirmar", "Limpar todo o histórico?", self)
        if not box.exec():
            return
        try:
            crypto.escrever_arquivo_seguro(crypto.ARQUIVO_HISTORICO, [])
            self._historico = []
            self._tabela.setRowCount(0)
            InfoBar.success("Sucesso", "Histórico limpo!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
        except Exception as e:
            InfoBar.error("Erro", str(e), parent=self, duration=4000,
                          position=InfoBarPosition.TOP)

    def showEvent(self, event):
        """Recarrega o histórico sempre que a página é exibida."""
        super().showEvent(event)
        self._carregar_historico()
