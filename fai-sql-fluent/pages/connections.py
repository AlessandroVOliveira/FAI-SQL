"""
Página de Gerenciamento de Conexões — FAI-SQL Fluent
CRUD de conexões SQL Server com teste integrado e proteção por senha.
"""

import hashlib
import os
import sys

sys.path.insert(0, os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.join(os.path.dirname(__file__), "..", ".."))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLineEdit,
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
    MessageBoxBase,
    PasswordLineEdit,
    PrimaryPushButton,
    PushButton,
    StrongBodyLabel,
    SubtitleLabel,
    ToolButton,
    isDarkTheme,
)

import crypto_utils as crypto
from utils.database import testar_conexao


# ======================================================================
# Diálogo de senha da tela de conexões
# ======================================================================

class _PasswordDialog(MessageBoxBase):
    """Diálogo Fluent para criar ou informar a senha da tela de conexões."""

    def __init__(self, is_new: bool, parent=None):
        self._is_new = is_new
        self._result_password: str | None = None
        super().__init__(parent)

        # Montar conteúdo após o super().__init__()
        self.widget.setMinimumWidth(350)

        if self._is_new:
            lbl = BodyLabel("Crie uma senha para proteger a tela de conexões:")
            self.viewLayout.addWidget(lbl)

            self._entry_senha = PasswordLineEdit(self)
            self._entry_senha.setPlaceholderText("Nova senha")
            self._entry_senha.setFixedHeight(36)
            self.viewLayout.addWidget(self._entry_senha)

            self._entry_confirma = PasswordLineEdit(self)
            self._entry_confirma.setPlaceholderText("Confirme a senha")
            self._entry_confirma.setFixedHeight(36)
            self.viewLayout.addWidget(self._entry_confirma)
        else:
            lbl = BodyLabel("Digite a senha para acessar as conexões:")
            self.viewLayout.addWidget(lbl)

            self._entry_senha = PasswordLineEdit(self)
            self._entry_senha.setPlaceholderText("Senha")
            self._entry_senha.setFixedHeight(36)
            self.viewLayout.addWidget(self._entry_senha)

        self._label_erro = BodyLabel("")
        self._label_erro.setStyleSheet("color: #f44336;")
        self.viewLayout.addWidget(self._label_erro)

        self.yesButton.setText("Confirmar")
        self.cancelButton.setText("Cancelar")

        self.yesButton.clicked.disconnect()
        self.yesButton.clicked.connect(self._confirmar)

    def _confirmar(self):
        senha = self._entry_senha.text()

        if self._is_new:
            confirma = self._entry_confirma.text()
            if len(senha) < 4:
                self._label_erro.setText("Senha deve ter pelo menos 4 caracteres!")
                return
            if senha != confirma:
                self._label_erro.setText("As senhas não coincidem!")
                return

        self._result_password = senha
        self.accept()
        self.accepted.emit()

    @property
    def password(self) -> str | None:
        return self._result_password


# ======================================================================
# Página de Conexões
# ======================================================================

def _hash_senha(senha: str) -> str:
    """Gera hash SHA-256 da senha."""
    return hashlib.sha256(senha.encode("utf-8")).hexdigest()


class ConnectionsPage(QWidget):
    """Página de gerenciamento de conexões com proteção por cadeado."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("connections_page")
        self._bloqueado = False       # se a tela está atualmente bloqueada
        self._desbloqueado_sessao = False  # se já desbloqueou nesta visita
        self._setup_ui()
        self._atualizar_lista()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(15)

        # --- Cabeçalho com título e cadeado ---
        header = QHBoxLayout()
        titulo = SubtitleLabel("Gerenciador de Conexões")
        header.addWidget(titulo)
        header.addStretch()

        self._btn_lock = ToolButton(FIF.CERTIFICATE)
        self._btn_lock.setFixedSize(36, 36)
        self._btn_lock.setToolTip("Proteger tela com senha")
        self._btn_lock.clicked.connect(self._on_lock_clicked)
        header.addWidget(self._btn_lock)

        layout.addLayout(header)

        # --- Conteúdo principal (será escondido quando bloqueado) ---
        self._content_widget = QWidget()
        content_layout = QVBoxLayout(self._content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(15)

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
        content_layout.addLayout(content, 1)

        layout.addWidget(self._content_widget, 1)

        # --- Overlay de bloqueio ---
        self._overlay = CardWidget()
        overlay_layout = QVBoxLayout(self._overlay)
        overlay_layout.setAlignment(Qt.AlignCenter)
        overlay_layout.setSpacing(15)

        lock_icon = SubtitleLabel("🔒")
        lock_icon.setAlignment(Qt.AlignCenter)
        lock_icon.setStyleSheet("font-size: 48px;")
        overlay_layout.addWidget(lock_icon)

        lock_label = StrongBodyLabel("Tela protegida por senha")
        lock_label.setAlignment(Qt.AlignCenter)
        overlay_layout.addWidget(lock_label)

        btn_desbloquear = PrimaryPushButton(FIF.FINGERPRINT, "Desbloquear")
        btn_desbloquear.setFixedSize(200, 40)
        btn_desbloquear.clicked.connect(self._solicitar_desbloqueio)
        overlay_layout.addWidget(btn_desbloquear, alignment=Qt.AlignCenter)

        self._overlay.setVisible(False)
        layout.addWidget(self._overlay, 1)

        # Atualizar ícone do cadeado
        self._atualizar_icone_lock()

    # ------------------------------------------------------------------
    # Proteção por senha
    # ------------------------------------------------------------------

    def _tem_senha_tela(self) -> bool:
        """Verifica se há uma senha de proteção configurada."""
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados and dados.get("conexao_senha_hash"):
                return True
        except Exception:
            pass
        return False

    def _verificar_senha_tela(self, senha: str) -> bool:
        """Verifica se a senha informada está correta."""
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados:
                return dados.get("conexao_senha_hash") == _hash_senha(senha)
        except Exception:
            pass
        return False

    def _salvar_senha_tela(self, senha: str):
        """Salva o hash da senha de proteção."""
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados is None:
                dados = {}
            dados["conexao_senha_hash"] = _hash_senha(senha)
            crypto.escrever_arquivo_seguro(crypto.ARQUIVO_CONFIG, dados)
        except Exception:
            pass

    def _remover_senha_tela(self):
        """Remove a senha de proteção."""
        try:
            dados = crypto.ler_arquivo_seguro(crypto.ARQUIVO_CONFIG)
            if dados and "conexao_senha_hash" in dados:
                del dados["conexao_senha_hash"]
                crypto.escrever_arquivo_seguro(crypto.ARQUIVO_CONFIG, dados)
        except Exception:
            pass

    def _atualizar_icone_lock(self):
        """Atualiza ícone e tooltip do cadeado."""
        if self._tem_senha_tela():
            self._btn_lock.setIcon(FIF.CERTIFICATE)
            self._btn_lock.setToolTip("Proteção ativa — clique para remover")
        else:
            self._btn_lock.setIcon(FIF.FINGERPRINT)
            self._btn_lock.setToolTip("Clique para ativar proteção por senha")

    def _bloquear_tela(self):
        """Bloqueia o acesso ao conteúdo."""
        self._bloqueado = True
        self._content_widget.setVisible(False)
        self._overlay.setVisible(True)

    def _desbloquear_tela(self):
        """Desbloqueia o acesso ao conteúdo."""
        self._bloqueado = False
        self._desbloqueado_sessao = True
        self._content_widget.setVisible(True)
        self._overlay.setVisible(False)

    def _on_lock_clicked(self):
        """Ação ao clicar no cadeado."""
        if self._tem_senha_tela():
            # Já tem senha: perguntar se deseja remover
            box = MessageBox(
                "Remover Proteção",
                "Deseja remover a senha de proteção desta tela?",
                self,
            )
            if box.exec():
                # Pedir senha atual para confirmar remoção
                dlg = _PasswordDialog(is_new=False, parent=self)
                if dlg.exec() and dlg.password:
                    if self._verificar_senha_tela(dlg.password):
                        self._remover_senha_tela()
                        self._desbloquear_tela()
                        self._atualizar_icone_lock()
                        InfoBar.success(
                            "Sucesso", "Proteção removida!",
                            parent=self, duration=3000, position=InfoBarPosition.TOP,
                        )
                    else:
                        InfoBar.error(
                            "Erro", "Senha incorreta!",
                            parent=self, duration=3000, position=InfoBarPosition.TOP,
                        )
        else:
            # Não tem senha: criar
            dlg = _PasswordDialog(is_new=True, parent=self)
            if dlg.exec() and dlg.password:
                self._salvar_senha_tela(dlg.password)
                self._atualizar_icone_lock()
                InfoBar.success(
                    "Sucesso",
                    "Proteção ativada! Na próxima visita a senha será solicitada.",
                    parent=self, duration=4000, position=InfoBarPosition.TOP,
                )

    def _solicitar_desbloqueio(self):
        """Abre o diálogo para desbloquear a tela."""
        dlg = _PasswordDialog(is_new=False, parent=self)
        if dlg.exec() and dlg.password:
            if self._verificar_senha_tela(dlg.password):
                self._desbloquear_tela()
                InfoBar.success(
                    "Sucesso", "Tela desbloqueada!",
                    parent=self, duration=2000, position=InfoBarPosition.TOP,
                )
            else:
                InfoBar.error(
                    "Erro", "Senha incorreta!",
                    parent=self, duration=3000, position=InfoBarPosition.TOP,
                )

    def showEvent(self, event):
        """Verifica proteção ao acessar a tela."""
        super().showEvent(event)
        if self._tem_senha_tela() and not self._desbloqueado_sessao:
            self._bloquear_tela()
        else:
            self._desbloquear_tela()

    def hideEvent(self, event):
        """Reseta desbloqueio ao sair da tela."""
        super().hideEvent(event)
        if self._tem_senha_tela():
            self._desbloqueado_sessao = False

    # ------------------------------------------------------------------
    # CRUD de Conexões
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
