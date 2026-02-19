"""
SQL Syntax Highlighter para QPlainTextEdit — FAI-SQL Fluent
Usa QSyntaxHighlighter do Qt para highlight em tempo real.
"""

import re

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QSyntaxHighlighter, QTextCharFormat

SQL_KEYWORDS = [
    "SELECT", "FROM", "WHERE", "AND", "OR", "NOT", "IN", "LIKE", "BETWEEN",
    "INSERT", "INTO", "VALUES", "UPDATE", "SET", "DELETE", "CREATE", "DROP",
    "ALTER", "TABLE", "INDEX", "VIEW", "DATABASE", "SCHEMA", "CONSTRAINT",
    "PRIMARY", "KEY", "FOREIGN", "REFERENCES", "UNIQUE", "CHECK", "DEFAULT",
    "NULL", "INNER", "LEFT", "RIGHT", "FULL", "OUTER", "JOIN", "ON", "AS",
    "ORDER", "BY", "ASC", "DESC", "GROUP", "HAVING", "DISTINCT", "TOP",
    "LIMIT", "OFFSET", "UNION", "ALL", "EXCEPT", "INTERSECT", "EXISTS",
    "CASE", "WHEN", "THEN", "ELSE", "END", "IF", "BEGIN", "COMMIT", "ROLLBACK",
    "TRANSACTION", "DECLARE", "EXEC", "EXECUTE", "PROCEDURE", "FUNCTION",
    "TRIGGER", "CURSOR", "FETCH", "OPEN", "CLOSE", "DEALLOCATE", "WITH",
    "CTE", "RECURSIVE", "PARTITION", "OVER", "ROW_NUMBER", "RANK", "DENSE_RANK",
    "IS", "TRUNCATE", "VARCHAR", "INT", "INTEGER", "FLOAT", "DECIMAL", "DATE",
    "DATETIME", "BIT", "TEXT", "NVARCHAR", "CHAR", "BIGINT", "SMALLINT",
]

SQL_FUNCTIONS = [
    "COUNT", "SUM", "AVG", "MIN", "MAX", "ABS", "CEILING", "FLOOR", "ROUND",
    "POWER", "SQRT", "LEN", "LENGTH", "SUBSTRING", "SUBSTR", "LEFT", "RIGHT",
    "LTRIM", "RTRIM", "TRIM", "UPPER", "LOWER", "REPLACE", "CHARINDEX",
    "CONCAT", "COALESCE", "ISNULL", "NULLIF", "CAST", "CONVERT", "GETDATE",
    "DATEADD", "DATEDIFF", "DATEPART", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE",
    "SECOND", "NOW", "CURRENT_DATE", "CURRENT_TIME", "CURRENT_TIMESTAMP",
]

# Cores para tema escuro
DARK_COLORS = {
    "keyword": "#569cd6",
    "function": "#dcdcaa",
    "string": "#ce9178",
    "number": "#b5cea8",
    "comment": "#6a9955",
    "operator": "#d4d4d4",
}

# Cores para tema claro
LIGHT_COLORS = {
    "keyword": "#0000ff",
    "function": "#9932cc",
    "string": "#008000",
    "number": "#ff6600",
    "comment": "#808080",
    "operator": "#cc0066",
}


class SqlHighlighter(QSyntaxHighlighter):
    """Syntax highlighter para SQL com suporte a tema claro/escuro."""

    def __init__(self, parent=None, dark: bool = False):
        super().__init__(parent)
        self._dark = dark
        self._build_rules()

    def set_dark(self, dark: bool):
        self._dark = dark
        self._build_rules()
        self.rehighlight()

    # ------------------------------------------------------------------
    def _make_fmt(self, color: str, bold: bool = False, italic: bool = False) -> QTextCharFormat:
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(color))
        if bold:
            fmt.setFontWeight(QFont.Weight.Bold)
        if italic:
            fmt.setFontItalic(True)
        return fmt

    def _build_rules(self):
        colors = DARK_COLORS if self._dark else LIGHT_COLORS
        self._rules: list[tuple[re.Pattern, QTextCharFormat]] = []

        # Keywords
        kw_pattern = r"\b(?:" + "|".join(SQL_KEYWORDS) + r")\b"
        self._rules.append(
            (re.compile(kw_pattern, re.IGNORECASE), self._make_fmt(colors["keyword"], bold=True))
        )

        # Functions (followed by parenthesis)
        fn_pattern = r"\b(?:" + "|".join(SQL_FUNCTIONS) + r")\s*(?=\()"
        self._rules.append(
            (re.compile(fn_pattern, re.IGNORECASE), self._make_fmt(colors["function"]))
        )

        # Numbers
        self._rules.append(
            (re.compile(r"\b\d+\.?\d*\b"), self._make_fmt(colors["number"]))
        )

        # Operators
        self._rules.append(
            (re.compile(r"[=<>!]+|[+\-*/]"), self._make_fmt(colors["operator"]))
        )

        # Strings (single-quoted)
        self._rules.append(
            (re.compile(r"'[^']*'"), self._make_fmt(colors["string"]))
        )

        # Strings (double-quoted)
        self._rules.append(
            (re.compile(r'"[^"]*"'), self._make_fmt(colors["string"]))
        )

        # Line comments
        self._rules.append(
            (re.compile(r"--[^\n]*"), self._make_fmt(colors["comment"], italic=True))
        )

        # Block comments
        self._rules.append(
            (re.compile(r"/\*[\s\S]*?\*/"), self._make_fmt(colors["comment"], italic=True))
        )

    # ------------------------------------------------------------------
    def highlightBlock(self, text: str):
        for pattern, fmt in self._rules:
            for match in pattern.finditer(text):
                self.setFormat(match.start(), match.end() - match.start(), fmt)
