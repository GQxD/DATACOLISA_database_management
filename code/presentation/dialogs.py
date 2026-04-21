"""Reusable dialog components for the UI."""

from __future__ import annotations

from typing import Any, Dict, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QMessageBox,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class MissingCodesDialog:
    """Dialog to show missing reference codes."""

    @staticmethod
    def show(parent: QWidget, missing_codes: List[str]) -> None:
        """Show missing codes dialog."""
        if not missing_codes:
            QMessageBox.information(
                parent,
                "Codes manquants",
                "Aucun code manquant sur cette plage",
            )
            return

        QMessageBox.information(
            parent,
            "Codes manquants",
            "\n".join(missing_codes),
        )


class ImportResultDialog(QDialog):
    """Dialog to show detailed import results with tabs."""

    def __init__(self, parent: QWidget, result: Dict[str, Any]):
        super().__init__(parent)
        self.result = result
        self.setWindowTitle("Resultats de l'import")
        self.setMinimumSize(800, 600)

        # Inherit parent's stylesheet if available
        if parent and parent.styleSheet():
            self.setStyleSheet(parent.styleSheet())

        self._setup_ui()

    def _setup_ui(self) -> None:
        """Setup the dialog UI with tabs."""
        layout = QVBoxLayout(self)

        # Extract result data
        imported = int(self.result.get("imported", 0) or 0)
        skipped_manual = int(self.result.get("skipped_manual", 0) or 0)
        skipped_validation = int(self.result.get("skipped_validation", 0) or 0)
        duplicates = int(self.result.get("duplicates", 0) or 0)

        # Header with summary
        headline = self._headline(imported, skipped_manual, skipped_validation, duplicates)
        header = QLabel(headline)  # Plain text, no HTML
        # Clear any inline styles to use theme colors
        header.setStyleSheet("")
        # Make it bold and larger using font properties
        from PySide6.QtGui import QFont
        font = header.font()
        font.setPointSize(14)
        font.setBold(True)
        header.setFont(font)
        header.setWordWrap(True)
        header.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(header)

        # Tab widget
        tabs = QTabWidget()
        layout.addWidget(tabs)

        # Tab 1: Summary
        summary_widget = self._create_summary_tab()
        tabs.addTab(summary_widget, "Resume")

        # Tab 2: Imported lines
        imported_refs = self.result.get("imported_refs", [])
        if imported_refs:
            imported_widget = self._create_imported_tab(imported_refs)
            tabs.addTab(imported_widget, f"Importees ({len(imported_refs)})")

        # Tab 3: Validation errors
        skipped_validation_details = self.result.get("skipped_validation_details", [])
        if skipped_validation_details:
            errors_widget = self._create_errors_tab(skipped_validation_details)
            tabs.addTab(errors_widget, f"Erreurs ({len(skipped_validation_details)})")

        # Tab 4: Duplicates
        duplicate_refs = self.result.get("duplicate_refs", [])
        doublons_tab_index = -1
        if duplicate_refs:
            duplicates_widget = self._create_duplicates_tab(duplicate_refs)
            doublons_tab_index = tabs.count()
            tabs.addTab(duplicates_widget, f"Doublons ({len(duplicate_refs)})")

        # Tab 5: Manually excluded
        skipped_manual_refs = self.result.get("skipped_manual_refs", [])
        if skipped_manual_refs:
            excluded_widget = self._create_excluded_tab(skipped_manual_refs)
            tabs.addTab(excluded_widget, f"Exclues ({len(skipped_manual_refs)})")

        # Auto-select Doublons tab if there are duplicates
        if doublons_tab_index >= 0:
            tabs.setCurrentIndex(doublons_tab_index)

        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

    def _create_summary_tab(self) -> QWidget:
        """Create the summary tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        message = self._format_result(self.result)
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(message)
        # Don't override colors - let parent stylesheet handle it
        text_edit.setStyleSheet("")
        layout.addWidget(text_edit)

        return widget

    def _create_imported_tab(self, imported_refs: List[str]) -> QWidget:
        """Create the imported lines tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        label = QLabel(f"<b>{len(imported_refs)} ligne(s) importee(s) avec succes :</b>")
        layout.addWidget(label)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n".join(f"- {ref}" for ref in imported_refs))
        # Don't override colors - let parent stylesheet handle it
        text_edit.setStyleSheet("")
        layout.addWidget(text_edit)

        return widget

    def _create_errors_tab(self, errors: List[Dict[str, Any]]) -> QWidget:
        """Create the validation errors tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        label = QLabel(f"<b>{len(errors)} ligne(s) avec erreurs de validation:</b>")
        layout.addWidget(label)

        lines = []
        for item in errors:
            ref = str(item.get("ref") or "?").strip()
            error_list = [str(err).strip() for err in item.get("errors", []) if str(err).strip()]
            if error_list:
                lines.append(f"- {ref} :")
                for err in error_list:
                    lines.append(f"  * {err}")
                lines.append("")

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n".join(lines))
        # Don't override colors - let parent stylesheet handle it
        text_edit.setStyleSheet("")
        layout.addWidget(text_edit)

        return widget

    def _create_duplicates_tab(self, duplicate_refs: List[str]) -> QWidget:
        """Create the duplicates tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        label = QLabel(f"<b>{len(duplicate_refs)} doublon(s) detecte(s) :</b>")
        layout.addWidget(label)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n".join(f"- {ref} (deja present dans le fichier)" for ref in duplicate_refs))
        # Don't override colors - let parent stylesheet handle it
        text_edit.setStyleSheet("")
        layout.addWidget(text_edit)

        return widget

    def _create_excluded_tab(self, excluded_refs: List[str]) -> QWidget:
        """Create the manually excluded tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        label = QLabel(f"<b>{len(excluded_refs)} ligne(s) exclue(s) manuellement:</b>")
        layout.addWidget(label)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n".join(f"- {ref} (non selectionne)" for ref in excluded_refs))
        # Don't override colors - let parent stylesheet handle it
        text_edit.setStyleSheet("")
        layout.addWidget(text_edit)

        return widget

    @staticmethod
    def show(parent: QWidget, result: Dict[str, Any]) -> None:
        """Show import result dialog."""
        dialog = ImportResultDialog(parent, result)
        dialog.exec()

    @staticmethod
    def _headline(imported: int, skipped_manual: int, skipped_validation: int, duplicates: int) -> str:
        """Return a short headline for the import result."""
        if imported > 0 and duplicates == 0 and skipped_validation == 0 and skipped_manual == 0:
            return "Import reussi sans probleme"
        if duplicates > 0 and imported == 0:
            return f"Attention : {duplicates} doublon(s) detecte(s)"
        if skipped_validation > 0 and imported == 0:
            return f"Attention : {skipped_validation} ligne(s) a corriger"
        if skipped_manual > 0 and imported == 0:
            return f"Attention : {skipped_manual} ligne(s) exclue(s)"
        if duplicates > 0 or skipped_validation > 0 or skipped_manual > 0:
            return "Import termine avec alertes"
        return "Import termine"

    @staticmethod
    def _format_result(result: Dict[str, Any]) -> str:
        """Format import result as user-friendly message."""
        imported = int(result.get("imported", 0) or 0)
        skipped_manual = int(result.get("skipped_manual", 0) or 0)
        skipped_validation = int(result.get("skipped_validation", 0) or 0)
        duplicates = int(result.get("duplicates", 0) or 0)
        not_imported = skipped_manual + skipped_validation + duplicates
        target_out = str(result.get("target_out", ""))
        duplicate_refs = [str(x).strip() for x in result.get("duplicate_refs", []) if str(x).strip()]
        skipped_manual_refs = [str(x).strip() for x in result.get("skipped_manual_refs", []) if str(x).strip()]
        skipped_validation_details = result.get("skipped_validation_details", []) or []

        lines: List[str] = []

        if imported > 0 and not_imported == 0:
            lines = [
                "Les fichiers ont bien ete importes.",
                f"Aucun probleme detecte pendant l'import.",
                f"Lignes importees : {imported}",
            ]
        elif imported > 0:
            lines = [
                "L'import est termine.",
                "Les fichiers ont ete importes, mais avec des problemes a verifier.",
                f"Lignes importees : {imported}",
                f"Lignes non importees : {not_imported}",
                f"Lignes exclues : {skipped_manual}",
                f"Lignes a corriger : {skipped_validation}",
                f"Doublons detectes : {duplicates}",
            ]
        else:
            lines = [
                "L'import est termine, mais aucun nouveau fichier exploitable n'a ete ajoute.",
                f"Lignes importees : {imported}",
                f"Lignes non importees : {not_imported}",
                f"Lignes exclues : {skipped_manual}",
                f"Lignes a corriger : {skipped_validation}",
                f"Doublons detectes : {duplicates}",
            ]

        if target_out:
            lines.extend([
                "",
                "Le fichier COLISA en cours a ete mis a jour.",
                f"Emplacement : {target_out}",
            ])

        if duplicates > 0:
            lines.extend([
                "",
                f"{duplicates} ligne(s) non importee(s) car deja presentes dans le fichier Excel.",
            ])
            if duplicate_refs:
                lines.append(f"Lignes concernees : {', '.join(duplicate_refs)}")
        if skipped_validation > 0:
            lines.extend([
                "",
                f"{skipped_validation} ligne(s) non importee(s) car il manque une information obligatoire.",
            ])
            for item in skipped_validation_details[:5]:
                ref = str(item.get("ref") or "?").strip()
                errors = [str(err).strip() for err in item.get("errors", []) if str(err).strip()]
                if errors:
                    lines.append(f"- {ref} : {', '.join(errors)}")
        if skipped_manual > 0:
            lines.extend([
                "",
                f"{skipped_manual} ligne(s) non importee(s) car elles ne sont pas selectionnees.",
            ])
            if skipped_manual_refs:
                lines.append(f"Lignes concernees : {', '.join(skipped_manual_refs)}")
        if imported == 0 and (duplicates > 0 or skipped_validation > 0):
            lines.extend(
                [
                    "",
                    "Aucune nouvelle ligne n'a ete ajoutee.",
                ]
            )
        if imported > 0:
            lines.extend(
                [
                    "",
                    "Ensuite, lance Collect-Science.",
                    "Le fichier COLISA genere sert a calculer des valeurs comme md_num_individu.",
                ]
            )
        return "\n".join(lines)


class HistoryDialog:
    """Dialog to show import history."""

    @staticmethod
    def show(parent: QWidget, history_data: Dict[str, Any]) -> None:
        """Show history dialog."""
        import json

        rows = history_data.get("rows", [])
        message = json.dumps(
            {
                "updated_at": history_data.get("updated_at"),
                "rows": rows[:50],
                "count": len(rows),
            },
            ensure_ascii=False,
            indent=2,
        )
        QMessageBox.information(parent, "Suivi", message)


class ConfirmationDialog:
    """Generic confirmation dialog."""

    @staticmethod
    def ask(parent: QWidget, title: str, message: str) -> bool:
        """Ask for user confirmation."""
        reply = QMessageBox.question(
            parent,
            title,
            message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        return reply == QMessageBox.Yes


class ErrorDialog:
    """Error dialog with consistent formatting."""

    @staticmethod
    def show(parent: QWidget, title: str, error: Exception | str) -> None:
        """Show error dialog."""
        QMessageBox.critical(parent, title, str(error))


class WarningDialog:
    """Warning dialog with consistent formatting."""

    @staticmethod
    def show(parent: QWidget, title: str, message: str) -> None:
        """Show warning dialog."""
        QMessageBox.warning(parent, title, message)


class InfoDialog:
    """Info dialog with consistent formatting."""

    @staticmethod
    def show(parent: QWidget, title: str, message: str) -> None:
        """Show info dialog."""
        QMessageBox.information(parent, title, message)


