"""
SQL Autocomplete Contextual para QPlainTextEdit — FAI-SQL Fluent
Sugere palavras-chave SQL e, após WHERE/ORDER BY/etc., apenas colunas da tabela.
"""

from __future__ import annotations

import re
from typing import Optional

from PySide6.QtCore import Qt, QStringListModel
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCompleter,
    QPlainTextEdit,
)

from utils.sql_highlighter import SQL_KEYWORDS, SQL_FUNCTIONS


# Mínimo de caracteres para ativar o autocomplete
MIN_PREFIX_LEN = 2

# Contextos onde devemos sugerir apenas colunas da tabela
_COLUMN_CONTEXTS = re.compile(
    r'\b(?:WHERE|AND|OR|ORDER\s+BY|GROUP\s+BY|HAVING|SET|ON|SELECT)\b',
    re.IGNORECASE,
)

# Regex para extrair tabela(s) do SQL
_FROM_TABLE = re.compile(
    r'\bFROM\s+([\w.\[\]]+)',
    re.IGNORECASE,
)
_JOIN_TABLE = re.compile(
    r'\bJOIN\s+([\w.\[\]]+)',
    re.IGNORECASE,
)
_UPDATE_TABLE = re.compile(
    r'\bUPDATE\s+([\w.\[\]]+)',
    re.IGNORECASE,
)
_INTO_TABLE = re.compile(
    r'\bINTO\s+([\w.\[\]]+)',
    re.IGNORECASE,
)


def _extrair_tabelas(sql: str) -> list[str]:
    """Extrai nomes de tabelas referenciadas no SQL."""
    tabelas = []
    for pattern in (_FROM_TABLE, _JOIN_TABLE, _UPDATE_TABLE, _INTO_TABLE):
        for m in pattern.finditer(sql):
            nome = m.group(1).strip("[]")
            if nome.upper() not in {k.upper() for k in SQL_KEYWORDS}:
                tabelas.append(nome)
    return tabelas


def _cursor_em_contexto_coluna(sql_ate_cursor: str) -> bool:
    """Verifica se o cursor está num contexto onde colunas são esperadas."""
    # Pegar o último contexto relevante
    matches = list(_COLUMN_CONTEXTS.finditer(sql_ate_cursor))
    if not matches:
        return False

    last_match = matches[-1]
    # Verificar se não há outro FROM/JOIN depois (o que mudaria o contexto)
    texto_apos = sql_ate_cursor[last_match.end():]
    if re.search(r'\bFROM\b', texto_apos, re.IGNORECASE):
        return False
    return True


class SqlCompleter(QCompleter):
    """Completer contextual que sugere colunas da tabela após WHERE."""

    def __init__(self, editor: QPlainTextEdit, parent=None):
        super().__init__(parent or editor)
        self._editor = editor
        self._all_words: list[str] = []       # todas as palavras (keywords + objetos)
        self._table_columns: dict[str, list[str]] = {}  # cache: tabela -> colunas
        self._column_loader = None            # callback para carregar colunas
        self._model = QStringListModel(self)

        # Configuração do popup
        self.setWidget(editor)
        self.setCompletionMode(QCompleter.PopupCompletion)
        self.setCaseSensitivity(Qt.CaseInsensitive)
        self.setFilterMode(Qt.MatchContains)
        self.setMaxVisibleItems(12)

        # Estilo do popup
        popup = self.popup()
        popup.setStyleSheet("""
            QListView {
                font-family: Consolas, monospace;
                font-size: 11pt;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 2px;
            }
            QListView::item {
                padding: 4px 8px;
            }
            QListView::item:selected {
                background-color: #0078d4;
                color: #fff;
            }
        """)
        popup.setSelectionBehavior(QAbstractItemView.SelectRows)

        self.activated.connect(self._inserir_completamento)
        self._editor.textChanged.connect(self._on_text_changed)

        self._build_base_words()

    # ------------------------------------------------------------------
    # Configuração
    # ------------------------------------------------------------------

    def set_database_objects(self, objects: list[str]):
        """Define a lista de objetos do banco (tabelas, colunas)."""
        self._all_words = list(set(
            [kw.upper() for kw in SQL_KEYWORDS]
            + [fn.upper() for fn in SQL_FUNCTIONS]
            + objects
        ))
        self._all_words.sort()

    def set_column_loader(self, loader):
        """Define callback(tabela) -> list[str] para colunas sob demanda."""
        self._column_loader = loader

    def _build_base_words(self):
        """Constrói a lista base de palavras (sem objetos do banco)."""
        self._all_words = sorted(set(
            [kw.upper() for kw in SQL_KEYWORDS]
            + [fn.upper() for fn in SQL_FUNCTIONS]
        ))

    # ------------------------------------------------------------------
    # Lógica contextual
    # ------------------------------------------------------------------

    def _get_word_under_cursor(self) -> str:
        cursor = self._editor.textCursor()
        cursor.movePosition(QTextCursor.StartOfBlock, QTextCursor.KeepAnchor)
        line_text = cursor.selectedText()
        match = re.search(r'[\w.]+$', line_text)
        return match.group() if match else ""

    def _get_sql_before_cursor(self) -> str:
        """Retorna todo o SQL do início até a posição do cursor."""
        cursor = self._editor.textCursor()
        cursor.movePosition(QTextCursor.Start, QTextCursor.KeepAnchor)
        return cursor.selectedText()

    def _get_table_columns(self, tabela: str) -> list[str]:
        """Retorna colunas da tabela, usando cache ou carregando do banco."""
        key = tabela.upper()
        if key in self._table_columns:
            return self._table_columns[key]

        if self._column_loader:
            cols = self._column_loader(tabela)
            if cols:
                self._table_columns[key] = cols
                return cols
        return []

    def _determinar_sugestoes(self) -> list[str]:
        """Decide quais palavras sugerir com base no contexto SQL."""
        sql = self._get_sql_before_cursor()
        tabelas = _extrair_tabelas(self._editor.toPlainText())

        if tabelas and _cursor_em_contexto_coluna(sql):
            # Contexto de coluna: sugerir apenas colunas das tabelas + operadores SQL
            colunas = []
            for t in tabelas:
                colunas.extend(self._get_table_columns(t))

            if colunas:
                # Adicionar também keywords de lógica (AND, OR, IS, NOT, etc.)
                extras = [
                    "AND", "OR", "NOT", "IN", "LIKE", "BETWEEN", "IS",
                    "NULL", "EXISTS", "ASC", "DESC", "AS", "CASE", "WHEN",
                    "THEN", "ELSE", "END", "ORDER", "BY", "GROUP", "HAVING",
                ]
                return sorted(set(colunas + extras))

        # Contexto genérico: todas as palavras
        return self._all_words

    # ------------------------------------------------------------------
    # Eventos
    # ------------------------------------------------------------------

    def _on_text_changed(self):
        prefix = self._get_word_under_cursor()

        if len(prefix) < MIN_PREFIX_LEN:
            self.popup().hide()
            return

        # Determinar sugestões baseado no contexto
        sugestoes = self._determinar_sugestoes()
        self._model.setStringList(sugestoes)
        self.setModel(self._model)

        self.setCompletionPrefix(prefix)

        if self.completionCount() == 0:
            self.popup().hide()
            return

        if (self.completionCount() == 1
                and self.currentCompletion().upper() == prefix.upper()):
            self.popup().hide()
            return

        cr = self._editor.cursorRect()
        cr.setWidth(
            self.popup().sizeHintForColumn(0)
            + self.popup().verticalScrollBar().sizeHint().width()
            + 20
        )
        cr.setWidth(max(cr.width(), 250))
        self.complete(cr)

    def _inserir_completamento(self, text: str):
        prefix = self._get_word_under_cursor()
        cursor = self._editor.textCursor()
        for _ in prefix:
            cursor.deletePreviousChar()
        cursor.insertText(text)
        self._editor.setTextCursor(cursor)
