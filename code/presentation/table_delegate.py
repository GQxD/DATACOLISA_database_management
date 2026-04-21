"""Table delegates for custom cell editors."""

from __future__ import annotations

from typing import Any, List, Optional, Sequence, Tuple
from PySide6.QtWidgets import QStyledItemDelegate, QComboBox, QWidget, QStyleOptionViewItem
from PySide6.QtCore import QModelIndex, Qt


class ComboBoxDelegate(QStyledItemDelegate):
    """
    Delegate for columns that use combo box editors.

    Key performance improvement:
    - QComboBox is only created during editing (not one per cell!)
    - In QTableWidget approach, combo boxes are created for ALL cells upfront
    - This delegate creates combo box ONLY when user clicks to edit

    Performance: 100+ rows with combo columns render instantly vs 10s+ freeze
    """

    def __init__(self, options: List[str], parent: Optional[QWidget] = None):
        """
        Initialize delegate with options.

        Args:
            options: List of string options for the combo box
            parent: Parent widget
        """
        super().__init__(parent)
        self.options = options

    def createEditor(
        self,
        parent: QWidget,
        option: QStyleOptionViewItem,
        index: QModelIndex
    ) -> QWidget:
        """
        Create combo box editor when cell is double-clicked.

        This is called ONLY when user enters edit mode, not at render time.
        """
        editor = QComboBox(parent)
        editor.addItems(self.options)
        editor.setEditable(False)  # Dropdown only, no free text
        return editor

    def setEditorData(self, editor: QWidget, index: QModelIndex) -> None:
        """
        Set the current value in the combo box from the model.
        """
        if not isinstance(editor, QComboBox):
            super().setEditorData(editor, index)
            return

        # Get current value from model
        value = index.model().data(index, role=0)  # Qt.DisplayRole
        if value is None:
            value = ""

        # Find and select matching option
        index_in_combo = editor.findText(str(value))
        if index_in_combo >= 0:
            editor.setCurrentIndex(index_in_combo)
        else:
            # If value not in options, add it temporarily
            editor.addItem(str(value))
            editor.setCurrentText(str(value))

    def setModelData(
        self,
        editor: QWidget,
        model: Any,
        index: QModelIndex
    ) -> None:
        """
        Save the selected combo box value back to the model.
        """
        if not isinstance(editor, QComboBox):
            super().setModelData(editor, model, index)
            return

        value = editor.currentText()
        model.setData(index, value, role=Qt.EditRole)

    def updateEditorGeometry(
        self,
        editor: QWidget,
        option: QStyleOptionViewItem,
        index: QModelIndex
    ) -> None:
        """
        Set the editor geometry to match the cell.
        """
        editor.setGeometry(option.rect)


class EditableComboBoxDelegate(QStyledItemDelegate):
    """
    Delegate for columns that use editable combo box (with free text entry).

    Allows user to either select from dropdown OR type custom value.
    """

    def __init__(self, options: List[str], parent: Optional[QWidget] = None):
        """
        Initialize delegate with options.

        Args:
            options: List of string options for the combo box
            parent: Parent widget
        """
        super().__init__(parent)
        self.options = options

    def createEditor(
        self,
        parent: QWidget,
        option: QStyleOptionViewItem,
        index: QModelIndex
    ) -> QWidget:
        """Create editable combo box editor."""
        editor = QComboBox(parent)
        editor.addItems(self.options)
        editor.setEditable(True)  # Allow free text entry
        editor.setInsertPolicy(QComboBox.NoInsert)  # Don't auto-add to list
        return editor

    def setEditorData(self, editor: QWidget, index: QModelIndex) -> None:
        """Set the current value in the combo box."""
        if not isinstance(editor, QComboBox):
            super().setEditorData(editor, index)
            return

        value = index.model().data(index, role=0)
        if value is None:
            value = ""
        editor.setCurrentText(str(value))

    def setModelData(
        self,
        editor: QWidget,
        model: Any,
        index: QModelIndex
    ) -> None:
        """Save the combo box value back to the model."""
        if not isinstance(editor, QComboBox):
            super().setModelData(editor, model, index)
            return

        value = editor.currentText()
        model.setData(index, value, role=Qt.EditRole)

    def updateEditorGeometry(
        self,
        editor: QWidget,
        option: QStyleOptionViewItem,
        index: QModelIndex
    ) -> None:
        """Set the editor geometry to match the cell."""
        editor.setGeometry(option.rect)


class MappedEditableComboBoxDelegate(QStyledItemDelegate):
    """Editable combo with display labels mapped to stored values."""

    def __init__(self, options: Sequence[Tuple[str, str]], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.options = list(options)

    def createEditor(self, parent: QWidget, option: QStyleOptionViewItem, index: QModelIndex) -> QWidget:
        editor = QComboBox(parent)
        for display, value in self.options:
            editor.addItem(display, value)
        editor.setEditable(True)
        editor.setInsertPolicy(QComboBox.NoInsert)
        return editor

    def setEditorData(self, editor: QWidget, index: QModelIndex) -> None:
        if not isinstance(editor, QComboBox):
            super().setEditorData(editor, index)
            return

        value = index.model().data(index, role=0)
        value = "" if value is None else str(value).strip()

        for combo_index in range(editor.count()):
            if str(editor.itemData(combo_index) or "").strip() == value:
                editor.setCurrentIndex(combo_index)
                return
        editor.setEditText(value)

    def setModelData(self, editor: QWidget, model: Any, index: QModelIndex) -> None:
        if not isinstance(editor, QComboBox):
            super().setModelData(editor, model, index)
            return

        value = editor.currentData()
        if value is None or not str(value).strip():
            text = editor.currentText().strip()
            value = text.split(" - ", 1)[0].strip() if " - " in text else text
        model.setData(index, str(value).strip(), role=Qt.EditRole)

    def updateEditorGeometry(self, editor: QWidget, option: QStyleOptionViewItem, index: QModelIndex) -> None:
        editor.setGeometry(option.rect)
