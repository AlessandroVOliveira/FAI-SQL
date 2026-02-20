"""
FAI-SQL Fluent — Entry Point + Tela de Login
"""

import sys
import os

# Adiciona a raiz do projeto ao path para importar crypto_utils
if getattr(sys, 'frozen', False):
    sys.path.insert(0, os.path.dirname(sys.executable))
else:
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    InfoBar,
    InfoBarPosition,
    LineEdit,
    PasswordLineEdit,
    PrimaryPushButton,
    PushButton,
    setTheme,
    Theme,
    MessageBox,
)

import crypto_utils as crypto


class LoginDialog(QDialog):
    """Tela de login / criação de senha mestra."""

    def __init__(self):
        super().__init__()
        self._primeira_vez = not crypto.senha_configurada()
        self._autenticado = False
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("🔐 FAI-SQL — Autenticação")
        self.setFixedSize(420, 360 if self._primeira_vez else 300)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        # Centralizar na tela
        screen = QApplication.primaryScreen().geometry()
        self.move(
            (screen.width() - self.width()) // 2,
            (screen.height() - self.height()) // 2,
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 35, 40, 35)
        layout.setSpacing(12)

        # Título
        titulo = QLabel("🔐 FAI-SQL")
        titulo.setFont(QFont("Segoe UI", 22, QFont.Weight.Bold))
        titulo.setAlignment(Qt.AlignCenter)
        titulo.setStyleSheet("color: #4da6ff;")
        layout.addWidget(titulo)

        subtitulo = QLabel("Sistema de Consultas SQL")
        subtitulo.setAlignment(Qt.AlignCenter)
        subtitulo.setStyleSheet("color: #888; font-size: 12px;")
        layout.addWidget(subtitulo)

        layout.addSpacing(10)

        # Label
        if self._primeira_vez:
            label = QLabel("Crie sua senha mestra:")
        else:
            label = QLabel("Digite sua senha:")
        label.setFont(QFont("Segoe UI", 11))
        layout.addWidget(label)

        # Campo de senha
        self._entry_senha = PasswordLineEdit(self)
        self._entry_senha.setPlaceholderText("Senha mestra")
        self._entry_senha.setFixedHeight(38)
        layout.addWidget(self._entry_senha)

        # Campo de confirmação (somente 1ª vez)
        self._entry_confirma = None
        if self._primeira_vez:
            label_conf = QLabel("Confirme a senha:")
            label_conf.setFont(QFont("Segoe UI", 11))
            layout.addWidget(label_conf)

            self._entry_confirma = PasswordLineEdit(self)
            self._entry_confirma.setPlaceholderText("Repita a senha")
            self._entry_confirma.setFixedHeight(38)
            layout.addWidget(self._entry_confirma)

        layout.addSpacing(5)

        # Botão Entrar
        btn_entrar = PrimaryPushButton(
            "Criar Senha" if self._primeira_vez else "Entrar", self
        )
        btn_entrar.setFixedHeight(40)
        btn_entrar.clicked.connect(self._tentar_autenticar)
        layout.addWidget(btn_entrar)

        # Link de reset (somente quando já tem senha)
        if not self._primeira_vez:
            layout.addSpacing(5)
            link_reset = QPushButton("Esqueceu a senha? Clique aqui para resetar")
            link_reset.setFlat(True)
            link_reset.setCursor(Qt.PointingHandCursor)
            link_reset.setStyleSheet(
                "color: #808080; font-size: 10px; text-decoration: underline; border: none;"
            )
            link_reset.clicked.connect(self._resetar_aplicacao)
            layout.addWidget(link_reset, alignment=Qt.AlignCenter)

        # Enter para submeter
        self._entry_senha.returnPressed.connect(self._tentar_autenticar)
        if self._entry_confirma:
            self._entry_confirma.returnPressed.connect(self._tentar_autenticar)

        self._entry_senha.setFocus()

    def _tentar_autenticar(self):
        senha = self._entry_senha.text()

        if self._primeira_vez:
            confirma = self._entry_confirma.text() if self._entry_confirma else ""
            if senha != confirma:
                InfoBar.error("Erro", "As senhas não coincidem!", parent=self,
                              duration=3000, position=InfoBarPosition.TOP)
                return
            if len(senha) < 4:
                InfoBar.error("Erro", "Senha deve ter pelo menos 4 caracteres!", parent=self,
                              duration=3000, position=InfoBarPosition.TOP)
                return

            sucesso, msg = crypto.configurar_senha(senha)
            if sucesso:
                crypto.migrar_dados_antigos()
                self._autenticado = True
                self.accept()
            else:
                InfoBar.error("Erro", msg, parent=self, duration=3000,
                              position=InfoBarPosition.TOP)
        else:
            sucesso, msg = crypto.verificar_senha(senha)
            if sucesso:
                self._autenticado = True
                self.accept()
            else:
                InfoBar.error("Erro", "Senha incorreta!", parent=self,
                              duration=3000, position=InfoBarPosition.TOP)
                self._entry_senha.clear()

    def _resetar_aplicacao(self):
        box = MessageBox(
            "⚠️ Atenção",
            "Isso apagará TODOS os dados!\n\n• Conexões salvas\n• Comandos salvos\n• Histórico\n\nTem certeza?",
            self,
        )
        if box.exec():
            crypto.resetar_dados()
            InfoBar.success("Sucesso", "Dados resetados. Reinicie a aplicação.",
                            parent=self, duration=3000, position=InfoBarPosition.TOP)
            self.reject()

    @property
    def autenticado(self) -> bool:
        return self._autenticado


def main():
    # Setar diretório de trabalho para a pasta do .exe (ou raiz do projeto em dev)
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))
    else:
        os.chdir(os.path.join(os.path.dirname(__file__), ".."))

    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))

    # Tema escuro inicial para login
    setTheme(Theme.DARK)

    # Tela de login
    login = LoginDialog()
    if login.exec() != QDialog.Accepted or not login.autenticado:
        sys.exit(0)

    # Importar aqui para evitar import circular
    from app import MainWindow

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
