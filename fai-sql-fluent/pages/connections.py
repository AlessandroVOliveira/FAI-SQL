"""
Página de Gerenciamento de Conexões — FAI-SQL Fluent
CRUD de conexões SQL Server com teste integrado.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout,
    QListWidgetItem,
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
    PasswordLineEdit,
    PrimaryPushButton,
    PushButton,
    StrongBodyLabel,
    SubtitleLabel,
)

import crypto_utils as crypto
from utils.database import testar_conexao


class ConnectionsPage(QWidget):
    """Página de gerenciamento de conexões."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("connections_page")
        self._setup_ui()
        self._atualizar_lista()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(15)

        titulo = SubtitleLabel("Gerenciador de Conexões")
        layout.addWidget(titulo)

        # Layout horizontal: lista à esquerda, formulário à direita
        content = QHBoxLayout()
        content.setSpacing(20)

        # === Lista de Conexões ===
        left_card = CardWidget()
        left_layout = QVBoxLayout(left_card)
        left_layout.setContentsMargins(15, 15, 15, 15)

        left_title = StrongBodyLabel("Conexões Salvas")
        left_layout.addWidget(left_title)

        self._lista = ListWidget()
        self._lista.setMinimumWidth(280)
        self._lista.currentRowChanged.connect(self._ao_selecionar)
        left_layout.addWidget(self._lista, 1)

        btn_row = QHBoxLayout()
        btn_usar = PrimaryPushButton(FIF.ACCEPT, "Usar")
        btn_usar.clicked.connect(self._usar_conexao)
        btn_row.addWidget(btn_usar)

        btn_excluir = PushButton(FIF.DELETE, "Excluir")
        btn_excluir.clicked.connect(self._excluir_conexao)
        btn_row.addWidget(btn_excluir)
        left_layout.addLayout(btn_row)

        content.addWidget(left_card, 2)

        # === Formulário ===
        right_card = CardWidget()
        right_layout = QVBoxLayout(right_card)
        right_layout.setContentsMargins(20, 20, 20, 20)
        right_layout.setSpacing(12)

        right_title = StrongBodyLabel("Dados da Conexão")
        right_layout.addWidget(right_title)

        # Campos
        self._entries: dict[str, LineEdit | PasswordLineEdit] = {}
        campos = [
            ("Nome", "nome", False),
            ("Servidor / IP", "ip", False),
            ("Usuário", "usuario", False),
            ("Senha", "senha", True),
            ("Banco de Dados", "banco", False),
        ]
        for label_text, key, is_password in campos:
            lbl = BodyLabel(label_text)
            right_layout.addWidget(lbl)

            if is_password:
                entry = PasswordLineEdit(self)
            else:
                entry = LineEdit(self)
            entry.setFixedHeight(36)
            entry.setClearButtonEnabled(True)
            right_layout.addWidget(entry)
            self._entries[key] = entry

        right_layout.addStretch()

        # Botões do formulário
        form_btns = QHBoxLayout()

        btn_testar = PushButton(FIF.WIFI, "Testar")
        btn_testar.setFixedHeight(36)
        btn_testar.clicked.connect(self._testar_conexao)
        form_btns.addWidget(btn_testar)

        btn_salvar = PrimaryPushButton(FIF.SAVE, "Salvar")
        btn_salvar.setFixedHeight(36)
        btn_salvar.clicked.connect(self._salvar_conexao)
        form_btns.addWidget(btn_salvar)

        btn_novo = PushButton(FIF.ADD, "Novo")
        btn_novo.setFixedHeight(36)
        btn_novo.clicked.connect(self._limpar_campos)
        form_btns.addWidget(btn_novo)

        right_layout.addLayout(form_btns)

        content.addWidget(right_card, 3)
        layout.addLayout(content, 1)

    # ------------------------------------------------------------------

    def _carregar_conexoes(self) -> list[dict]:
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados:
                if "conexoes" not in dados:
                    if "ip" in dados:
                        dados["conexoes"] = [{
                            "nome": "Padrão",
                            "ip": dados.get("ip", ""),
                            "usuario": dados.get("usuario", ""),
                            "senha": dados.get("senha", ""),
                            "banco": dados.get("banco", ""),
                        }]
                        dados["conexao_ativa"] = "Padrão"
                        crypto.escrever_arquivo_seguro(crypto.ARQUIVO_CONFIG, dados)
                return dados.get("conexoes", [])
        except Exception:
            pass
        return []

    def _atualizar_lista(self):
        self._lista.clear()
        for conn in self._carregar_conexoes():
            self._lista.addItem(f"{conn.get('nome', 'Sem nome')} — {conn.get('banco', '')}")

    def _ao_selecionar(self, row: int):
        if row < 0:
            return
        conexoes = self._carregar_conexoes()
        if row < len(conexoes):
            conn = conexoes[row]
            self._entries["nome"].setText(conn.get("nome", ""))
            self._entries["ip"].setText(conn.get("ip", ""))
            self._entries["usuario"].setText(conn.get("usuario", ""))
            self._entries["senha"].setText(conn.get("senha", ""))
            self._entries["banco"].setText(conn.get("banco", ""))

    def _usar_conexao(self):
        row = self._lista.currentRow()
        if row < 0:
            InfoBar.warning("Atenção", "Selecione uma conexão!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            return
        conexoes = self._carregar_conexoes()
        if row < len(conexoes):
            nome = conexoes[row].get("nome")
            try:
                dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
                if dados is None:
                    dados = {}
                dados["conexao_ativa"] = nome
                crypto.escrever_arquivo_seguro(crypto.ARQUIVO_CONFIG, dados)
                InfoBar.success("Sucesso", f"Conexão '{nome}' ativada!",
                                parent=self, duration=3000, position=InfoBarPosition.TOP)
                # Atualizar editor page
                main_window = self.window()
                if hasattr(main_window, "_editor_page"):
                    main_window._editor_page.atualizar_conexao()
            except Exception as e:
                InfoBar.error("Erro", str(e), parent=self, duration=4000,
                              position=InfoBarPosition.TOP)

    def _excluir_conexao(self):
        row = self._lista.currentRow()
        if row < 0:
            InfoBar.warning("Atenção", "Selecione uma conexão!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            return

        box = MessageBox("Confirmar Exclusão", "Excluir esta conexão?", self)
        if not box.exec():
            return

        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados is None:
                dados = {}
            conexoes = dados.get("conexoes", [])
            if row < len(conexoes):
                del conexoes[row]
                dados["conexoes"] = conexoes
                crypto.escrever_arquivo_seguro(crypto.ARQUIVO_CONFIG, dados)
                self._atualizar_lista()
                self._limpar_campos()
                InfoBar.success("Sucesso", "Conexão excluída!", parent=self,
                                duration=3000, position=InfoBarPosition.TOP)
        except Exception as e:
            InfoBar.error("Erro", str(e), parent=self, duration=4000,
                          position=InfoBarPosition.TOP)

    def _testar_conexao(self):
        ok, msg = testar_conexao(
            self._entries["ip"].text(),
            self._entries["usuario"].text(),
            self._entries["senha"].text(),
            self._entries["banco"].text(),
        )
        if ok:
            InfoBar.success("Sucesso", msg, parent=self, duration=3000,
                            position=InfoBarPosition.TOP)
        else:
            InfoBar.error("Falha", msg, parent=self, duration=5000,
                          position=InfoBarPosition.TOP)

    def _salvar_conexao(self):
        nome = self._entries["nome"].text().strip()
        if not nome:
            InfoBar.warning("Atenção", "Informe um nome para a conexão!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)
            return

        nova_conn = {
            "nome": nome,
            "ip": self._entries["ip"].text(),
            "usuario": self._entries["usuario"].text(),
            "senha": self._entries["senha"].text(),
            "banco": self._entries["banco"].text(),
        }

        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados is None:
                dados = {}
            conexoes = dados.get("conexoes", [])

            encontrado = False
            for i, conn in enumerate(conexoes):
                if conn.get("nome") == nome:
                    conexoes[i] = nova_conn
                    encontrado = True
                    break
            if not encontrado:
                conexoes.append(nova_conn)

            dados["conexoes"] = conexoes
            if not dados.get("conexao_ativa"):
                dados["conexao_ativa"] = nome
            crypto.escrever_arquivo_seguro(crypto.ARQUIVO_CONFIG, dados)

            self._atualizar_lista()
            InfoBar.success("Sucesso", "Conexão salva!", parent=self,
                            duration=3000, position=InfoBarPosition.TOP)

            main_window = self.window()
            if hasattr(main_window, "_editor_page"):
                main_window._editor_page.atualizar_conexao()
        except Exception as e:
            InfoBar.error("Erro", str(e), parent=self, duration=4000,
                          position=InfoBarPosition.TOP)

    def _limpar_campos(self):
        for entry in self._entries.values():
            entry.clear()
        self._lista.clearSelection()
