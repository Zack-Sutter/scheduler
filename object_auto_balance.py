import copy
import datetime
import os
import random
import re
import subprocess
import sys
from dataclasses import dataclass

import numpy as np
import pandas as pd
import logging
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QTimer, Signal
from PySide6.QtGui import QBrush, QColor, QCloseEvent, QFontMetrics, QPainter, QPen
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from object_auto_balance_core import (
    NONSTANDARD_SHIFTS,
    RES_FILE_NAME,
    SHIFT_INFO,
    STANDARD_FLOOR_SHIFTS,
    SWAPPABLE_FLOOR_SHIFTS,
    count_column_violations,
    default_balance_rules,
    format_balance_rule_line,
    shift_background_color,
    total_violations,
)

primary_button_color = '#EDD863'
primary_button_hover_color = '#E1D591'
secondary_button_color = '#d15462'
secondary_button_hover_color = '#ea6473'
enable_button_color = '#2e7d32'
enable_button_hover_color = '#388e3c'
disable_button_color = secondary_button_color
disable_button_hover_color = secondary_button_hover_color
background_color = '#ffffff'
frame_color = '#eeeeee'
text_color = '#000000'
subtitle_text_color = '#666666'
text_field_color = '#c6c6c6'
sheet_grid_line_color = '#b0b0b0'
regular_font_size = 11

APP_STYLESHEET = f"""
* {{
    color: {text_color};
    font-size: {regular_font_size}pt;
}}
QMainWindow, QWidget#centralRoot, QDialog {{
    background-color: {background_color};
}}
QFrame {{
    background-color: {frame_color};
}}
QLineEdit, QTextEdit {{
    background-color: {text_field_color};
    color: {text_color};
    font-size: {regular_font_size-1}pt;
}}
QListWidget, QTableView {{
    background-color: {text_field_color};
}}
QHeaderView::section {{
    background-color: {frame_color};
    color: {text_color};
}}
QPushButton#primaryBtn {{
    background-color: {primary_button_color};
    color: {text_color};
}}
QPushButton#primaryBtn:hover {{
    background-color: {primary_button_hover_color};
}}
QPushButton#secondaryBtn {{
    background-color: {secondary_button_color};
    color: {text_color};
}}
QPushButton#secondaryBtn:hover {{
    background-color: {secondary_button_hover_color};
}}
QPushButton#enableBtn {{
    background-color: {enable_button_color};
    color: {text_color};
}}
QPushButton#enableBtn:hover {{
    background-color: {enable_button_hover_color};
}}
QPushButton#disableBtn {{
    background-color: {disable_button_color};
    color: {text_color};
}}
QPushButton#disableBtn:hover {{
    background-color: {disable_button_hover_color};
}}
"""


def label_with_subtitle(title: str, subtitle: str) -> QWidget:
    container = QWidget()
    layout = QVBoxLayout(container)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(0)
    title_label = QLabel(title)
    title_font = title_label.font()
    title_font.setPointSize(regular_font_size)
    title_label.setFont(title_font)
    layout.addWidget(title_label)
    sub = QLabel(subtitle)
    sub_font = sub.font()
    sub_font.setPointSize(max(regular_font_size - 2, 6))
    sub.setFont(sub_font)
    sub.setStyleSheet(f'color: {subtitle_text_color};')
    layout.addWidget(sub)
    return container


@dataclass
class SelectionRegion:
    from_row: int
    upto_row: int
    from_col: int
    upto_col: int

    @property
    def row_count(self) -> int:
        return self.upto_row - self.from_row

    @property
    def col_count(self) -> int:
        return self.upto_col - self.from_col

    def time_range(self, df: pd.DataFrame) -> tuple:
        return df.index[self.from_row], df.index[self.upto_row - 1]

    def column_names(self, df: pd.DataFrame) -> list:
        return list(df.columns[self.from_col : self.upto_col])

    def values_2d(self, df: pd.DataFrame) -> list:
        block = df.iloc[self.from_row : self.upto_row, self.from_col : self.upto_col]
        return block.values.tolist()

    def values_flat(self, df: pd.DataFrame) -> list:
        flat = []
        for row in self.values_2d(df):
            for cell in row:
                flat.append(cell)
        return flat


class ScheduleTableModel(QAbstractTableModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._df = pd.DataFrame()

    def set_dataframe(self, df: pd.DataFrame, full_reset: bool = False) -> None:
        if full_reset or self._df.shape != df.shape or list(self._df.columns) != list(df.columns):
            self.beginResetModel()
            self._df = df.copy()
            self.endResetModel()
        else:
            self._df = df.copy()
            top_left = self.index(0, 0)
            bottom_right = self.index(max(0, len(self._df) - 1), max(0, len(self._df.columns) - 1))
            self.dataChanged.emit(
                top_left, bottom_right, [Qt.DisplayRole, Qt.BackgroundRole, Qt.ForegroundRole]
            )

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._df)

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._df.columns)

    def data(self, index: QModelIndex, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        value = self._df.iat[index.row(), index.column()]
        if role == Qt.DisplayRole:
            return '' if pd.isna(value) else str(value)
        if role == Qt.BackgroundRole:
            color = shift_background_color(value)
            if color:
                return QBrush(QColor(color))
        if role == Qt.ForegroundRole:
            return QBrush(QColor(text_color))
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            if section < len(self._df.columns):
                return str(self._df.columns[section])
        elif section < len(self._df.index):
            return str(self._df.index[section])
        return None

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable


class ScheduleColumnHeader(QHeaderView):
    def __init__(self, parent=None):
        super().__init__(Qt.Horizontal, parent)
        self._grid_color = QColor(sheet_grid_line_color)

    def paintSection(self, painter: QPainter, rect, logicalIndex: int) -> None:
        super().paintSection(painter, rect, logicalIndex)
        painter.save()
        painter.setPen(QPen(self._grid_color, 1))
        right = rect.right()
        bottom = rect.bottom()
        painter.drawLine(right, rect.top(), right, bottom)
        painter.drawLine(rect.left(), bottom, right, bottom)
        painter.restore()


class ScheduleTableView(QTableView):
    regions_changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._regions: list[SelectionRegion] = []
        self._drag_anchor: tuple[int, int] | None = None
        self.setSelectionMode(QTableView.NoSelection)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        column_header = ScheduleColumnHeader(self)
        self.setHorizontalHeader(column_header)
        column_header.sectionClicked.connect(self._on_column_header_clicked)

    def scrollContentsBy(self, dx: int, dy: int) -> None:
        pass

    def wheelEvent(self, event):
        event.ignore()

    def selection_regions(self) -> list[SelectionRegion]:
        return list(self._regions)

    def set_regions(self, regions: list[SelectionRegion]) -> None:
        self._regions = regions
        self.viewport().update()
        self.regions_changed.emit()

    def _on_column_header_clicked(self, column: int) -> None:
        if self.model() is None or self.model().rowCount() == 0:
            return
        region = SelectionRegion(0, self.model().rowCount(), column, column + 1)
        self.set_regions([region])

    def mousePressEvent(self, event):
        index = self.indexAt(event.position().toPoint())
        if index.isValid():
            self._drag_anchor = (index.row(), index.column())
        else:
            self._drag_anchor = None
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if self._drag_anchor is not None:
            end_index = self.indexAt(event.position().toPoint())
            if end_index.isValid():
                r0, c0 = self._drag_anchor
                r1, c1 = end_index.row(), end_index.column()
                region = SelectionRegion(
                    min(r0, r1),
                    max(r0, r1) + 1,
                    min(c0, c1),
                    max(c0, c1) + 1,
                )
                if event.modifiers() & Qt.ControlModifier:
                    regions = self._regions + [region]
                    if len(regions) > 2:
                        regions = regions[-2:]
                    self.set_regions(regions)
                else:
                    self.set_regions([region])
        self._drag_anchor = None
        super().mouseReleaseEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self._regions:
            return

        painter = QPainter(self.viewport())
        pen = QPen(QColor('#2563eb'), 2)
        painter.setPen(pen)
        for region in self._regions:
            top_left = self.visualRect(self.model().index(region.from_row, region.from_col))
            bottom_right = self.visualRect(
                self.model().index(region.upto_row - 1, region.upto_col - 1)
            )
            rect = top_left.united(bottom_right)
            painter.drawRect(rect.adjusted(0, 0, -1, -1))
        painter.end()


class BalanceRulesDialog(QDialog):
    def __init__(self, parent, on_apply):
        super().__init__(parent)
        self.on_apply = on_apply
        self.setWindowTitle('Balance Shifts')
        self.setModal(True)
        self.setMinimumWidth(400)
        self.rules = copy.deepcopy(default_balance_rules())

        layout = QVBoxLayout(self)
        intro = QLabel('Rules run top to bottom. Order matters.')
        intro.setWordWrap(True)
        layout.addWidget(intro)

        content = QHBoxLayout()
        self.rule_list = QListWidget()
        content.addWidget(self.rule_list, stretch=1)

        controls = QVBoxLayout()
        self.toggle_btn = QPushButton('Enable / Disable')
        self.toggle_btn.setObjectName('primaryBtn')
        self.toggle_btn.clicked.connect(self._toggle_selected)
        controls.addWidget(self.toggle_btn)
        up_btn = QPushButton('Move Up')
        up_btn.setObjectName('primaryBtn')
        up_btn.clicked.connect(self._move_up)
        controls.addWidget(up_btn)
        down_btn = QPushButton('Move Down')
        down_btn.setObjectName('primaryBtn')
        down_btn.clicked.connect(self._move_down)
        controls.addWidget(down_btn)
        content.addLayout(controls)
        layout.addLayout(content)

        self.rule_list.currentRowChanged.connect(self._update_toggle_button)
        self.rule_list.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        buttons = QDialogButtonBox()
        cancel_btn = QPushButton('Cancel')
        cancel_btn.setObjectName('secondaryBtn')
        cancel_btn.clicked.connect(self.reject)
        apply_btn = QPushButton('Apply')
        apply_btn.setObjectName('primaryBtn')
        apply_btn.clicked.connect(self._apply)
        buttons.addButton(cancel_btn, QDialogButtonBox.RejectRole)
        buttons.addButton(apply_btn, QDialogButtonBox.AcceptRole)
        layout.addWidget(buttons)

        self.setStyleSheet(APP_STYLESHEET)
        self.refresh_rule_list()
        self.rule_list.setCurrentRow(0)

    def _restyle_button(self, button: QPushButton) -> None:
        style = button.style()
        style.unpolish(button)
        style.polish(button)
        button.update()

    def _update_toggle_button(self) -> None:
        index = self._selected_index()
        if index is None:
            self.toggle_btn.setText('Enable / Disable')
            self.toggle_btn.setObjectName('primaryBtn')
        elif self.rules[index].enabled:
            self.toggle_btn.setText('Disable')
            self.toggle_btn.setObjectName('disableBtn')
        else:
            self.toggle_btn.setText('Enable')
            self.toggle_btn.setObjectName('enableBtn')
        self._restyle_button(self.toggle_btn)

    def refresh_rule_list(self):
        row = self.rule_list.currentRow()
        self.rule_list.clear()
        for index, rule in enumerate(self.rules):
            item = QListWidgetItem(format_balance_rule_line(index, rule))
            item.setForeground(QBrush(QColor(text_color)))
            self.rule_list.addItem(item)
        if self.rules:
            row = min(max(row, 0), len(self.rules) - 1)
            self.rule_list.setCurrentRow(row)
        self._update_toggle_button()

    def _selected_index(self):
        row = self.rule_list.currentRow()
        return None if row < 0 else row

    def _toggle_selected(self):
        index = self._selected_index()
        if index is None:
            return
        self.rules[index].enabled = not self.rules[index].enabled
        self.refresh_rule_list()

    def _move_up(self):
        index = self._selected_index()
        if index is None or index == 0:
            return
        self.rules[index], self.rules[index - 1] = self.rules[index - 1], self.rules[index]
        self.refresh_rule_list()
        self.rule_list.setCurrentRow(index - 1)

    def _move_down(self):
        index = self._selected_index()
        if index is None or index >= len(self.rules) - 1:
            return
        self.rules[index], self.rules[index + 1] = self.rules[index + 1], self.rules[index]
        self.refresh_rule_list()
        self.rule_list.setCurrentRow(index + 1)

    def _apply(self):
        self.on_apply(self.rules)
        self.accept()


class SheetFrame(QWidget):
    FRAME_WIDTH = 975
    FRAME_HEIGHT = 365

    def __init__(self, controller: 'ScheduleApp'):
        super().__init__()
        self.controller = controller
        self._frame_width = self.FRAME_WIDTH
        self._frame_height = self.FRAME_HEIGHT
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.model = ScheduleTableModel(self)
        self.table_view = ScheduleTableView(self)
        self.table_view.setObjectName('scheduleTable')
        self.table_view.setModel(self.model)
        self.table_view.setShowGrid(True)
        self.table_view.setCornerButtonEnabled(False)
        self.table_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table_view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table_view.setStyleSheet(
            f'QTableView#scheduleTable {{ gridline-color: {sheet_grid_line_color}; }}'
        )
        h_header = self.table_view.horizontalHeader()
        v_header = self.table_view.verticalHeader()
        h_header.setSectionResizeMode(QHeaderView.Fixed)
        h_header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        h_header.setMinimumSectionSize(1)
        h_header.setStretchLastSection(False)
        v_header.setSectionResizeMode(QHeaderView.Stretch)
        v_header.setMinimumSectionSize(1)
        v_header.setStretchLastSection(False)
        layout.addWidget(self.table_view)

    def set_frame_size(self, width: int, height: int) -> None:
        self._frame_width = width
        self._frame_height = height
        self.setFixedSize(width, height)

    def showEvent(self, event):
        super().showEvent(event)
        self._fit_table_to_frame()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._fit_table_to_frame()

    def _frame_size(self) -> tuple[int, int]:
        width, height = self.width(), self.height()
        if width <= 1:
            width = self._frame_width
        if height <= 1:
            height = self._frame_height
        return width, height

    def _viewport_size(self) -> tuple[int, int]:
        frame_w, frame_h = self._frame_size()
        view = self.table_view
        h_header = view.horizontalHeader()
        v_header = view.verticalHeader()
        header_h = max(h_header.height(), h_header.sizeHint().height())
        header_w = max(v_header.width(), v_header.sizeHint().width())
        return max(1, frame_w - header_w), max(1, frame_h - header_h)

    def _fit_table_to_frame(self) -> None:
        view = self.table_view
        rows = self.model.rowCount()
        cols = self.model.columnCount()
        if rows == 0 or cols == 0:
            return

        viewport_w = view.viewport().width()
        if viewport_w <= 1:
            viewport_w, _ = self._viewport_size()
        if view.showGrid():
            viewport_w = max(1, viewport_w - max(0, cols - 1))

        base_col_w = viewport_w // cols
        extra_col_w = viewport_w - base_col_w * cols
        for col in range(cols):
            view.setColumnWidth(col, base_col_w + (1 if col < extra_col_w else 0))

        view.verticalHeader().resizeSections()
        view.horizontalScrollBar().setValue(0)
        view.verticalScrollBar().setValue(0)

    def update_sheet(self):
        df = self.controller.df.copy()
        full_reset = self.model.rowCount() == 0 or self.model.columnCount() == 0
        self.model.set_dataframe(df.fillna(''), full_reset=full_reset)
        self.table_view.set_regions([])
        self._fit_table_to_frame()
        QTimer.singleShot(0, self._fit_table_to_frame)

    def column_names(self) -> list:
        return list(self.controller.df.columns)


class InputFrame(QWidget):
    def __init__(self, controller: 'ScheduleApp'):
        super().__init__()
        self.controller = controller
        grid = QGridLayout(self)
        self.standard_frame = StandardShiftFrame(self, controller)
        self.nonstandard_frame = NonStandardShiftFrame(self, controller)
        grid.addWidget(self.standard_frame, 0, 0, 5, 1)
        grid.addWidget(self.nonstandard_frame, 0, 1, 5, 1)

        self.undo_button = QPushButton('Undo')
        self.undo_button.setObjectName('secondaryBtn')
        self.undo_button.clicked.connect(controller.undo)
        grid.addWidget(self.undo_button, 0, 2)

        self.redo_button = QPushButton('Redo')
        self.redo_button.setObjectName('secondaryBtn')
        self.redo_button.clicked.connect(controller.redo)
        grid.addWidget(self.redo_button, 0, 3)

        self.swap_button = QPushButton('Swap')
        self.swap_button.setObjectName('primaryBtn')
        self.swap_button.clicked.connect(
            lambda: controller._perform_with_undo(controller.swap)
        )
        grid.addWidget(self.swap_button, 1, 2, 1, 1)

        self.balance_button = QPushButton('Balance')
        self.balance_button.setObjectName('primaryBtn')
        self.balance_button.clicked.connect(controller.show_balance_rules_dialog)
        grid.addWidget(self.balance_button, 1, 3, 1, 1)


class NonStandardShiftFrame(QFrame):
    def __init__(self, parent: InputFrame, controller: 'ScheduleApp'):
        super().__init__(parent)
        self.controller = controller
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        grid = QGridLayout(self)
        grid.addWidget(
            label_with_subtitle('Add nonstandard shift', 'fill selection'), 0, 0, 1, 1
        )
        self.list_widget = QListWidget()
        self.list_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list_widget.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        fm = QFontMetrics(self.list_widget.font())
        list_width = fm.horizontalAdvance('0' * 20) + 2 * self.list_widget.frameWidth() + 8
        self.list_widget.setFixedWidth(list_width)
        for shift in [*NONSTANDARD_SHIFTS, 'DELETE']:
            item = QListWidgetItem(shift)
            color = shift_background_color(shift) if shift != 'DELETE' else '#D3D3D3'
            if color:
                item.setBackground(QBrush(QColor(color)))
            item.setForeground(QBrush(QColor(text_color)))
            self.list_widget.addItem(item)
        list_height = sum(
            self.list_widget.sizeHintForRow(row) for row in range(self.list_widget.count())
        )
        list_widget_height = list_height + 2 * self.list_widget.frameWidth()
        self.list_widget.setFixedHeight(list_widget_height)
        self.list_widget.itemClicked.connect(self.on_list_item_clicked)
        grid.addWidget(self.list_widget, 1, 0)

        self.entry = QLineEdit()
        self.entry.setPlaceholderText('Custom shift')
        self.entry.setFixedWidth(list_width)
        self.entry.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.entry.returnPressed.connect(self.add_custom_shift_action)
        grid.addWidget(self.entry, 2, 0, 1, 1)

    def _apply_shift_to_selection(self, shift):
        selection = self.controller.get_sheet_selection()
        if not selection:
            return
        selection['shift'] = shift
        self.controller._perform_with_undo(
            self.controller.add_nonstandard_shift, selection
        )

    def on_list_item_clicked(self, item: QListWidgetItem):
        shift = item.text()
        if shift == 'DELETE':
            shift = np.nan
        self._apply_shift_to_selection(shift)

    def add_custom_shift_action(self):
        shift = self.entry.text().strip()
        if not shift:
            return
        self._apply_shift_to_selection(shift)
        self.entry.clear()


class StandardShiftFrame(QFrame):
    def __init__(self, parent: InputFrame, controller: 'ScheduleApp'):
        super().__init__(parent)
        self.controller = controller
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        grid = QGridLayout(self)
        grid.addWidget(label_with_subtitle('Add standard shift', 'one item per row'), 0, 0)
        self.list_widget = QListWidget()
        self.list_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list_widget.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        fm = QFontMetrics(self.list_widget.font())
        list_width = fm.horizontalAdvance('0' * 20) + 2 * self.list_widget.frameWidth() + 8
        self.list_widget.setFixedWidth(list_width)
        for shift in STANDARD_FLOOR_SHIFTS:
            item = QListWidgetItem(shift)
            color = shift_background_color(shift)
            if color:
                item.setBackground(QBrush(QColor(color)))
            item.setForeground(QBrush(QColor(text_color)))
            self.list_widget.addItem(item)
        list_height = sum(
            self.list_widget.sizeHintForRow(row) for row in range(self.list_widget.count())
        )
        list_widget_height = list_height + 2 * self.list_widget.frameWidth()
        self.list_widget.setFixedHeight(list_widget_height)
        self.list_widget.itemClicked.connect(self.on_list_item_clicked)
        grid.addWidget(self.list_widget, 1, 0)

    def on_list_item_clicked(self, item: QListWidgetItem):
        self.controller._perform_with_undo(
            self.controller.add_standard_shift, item.text()
        )


class ScheduleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('MoMath Automatic Scheduler')
        self.resize(1500, 800)
        self.setStyleSheet(APP_STYLESHEET)

        self.nonstandard_shifts = []
        self.action_history_stack = []
        self.action_redo_stack = []
        self.df = pd.DataFrame()
        self.sheet_frame: SheetFrame | None = None
        self.inputs: InputFrame | None = None
        self.paid_workers: list = []
        self.volunteers: list = []

        root = QWidget()
        root.setObjectName('centralRoot')
        self.setCentralWidget(root)
        main_layout = QHBoxLayout(root)

        left = QVBoxLayout()
        left.setContentsMargins(6, 6, 6, 6)
        left.setSpacing(6)
        main_layout.addLayout(left, stretch=1)

        paid_worker_header = label_with_subtitle('Paid workers', 'comma separated')
        paid_worker_header.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        left.addWidget(paid_worker_header)

        self.paid_workers_entry = QTextEdit()
        self.paid_workers_entry.setFixedHeight(60)
        left.addWidget(self.paid_workers_entry)

        volunteer_header = label_with_subtitle('Volunteers', 'comma separated')
        volunteer_header.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        left.addWidget(volunteer_header)
        self.volunteers_entry = QTextEdit()
        self.volunteers_entry.setFixedHeight(60)
        left.addWidget(self.volunteers_entry)

        operating_hours_label = QLabel('Operating hours')
        operating_hours_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        left.addWidget(operating_hours_label)
        self.operating_hours = QLineEdit('10:00 - 5:00')
        left.addWidget(self.operating_hours)

        lunch_label = QLabel('Early or late lunches today?')
        lunch_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        left.addWidget(lunch_label)
        self.lunch_timing_group = QButtonGroup(self)
        early_lunch = QRadioButton('Early')
        late_lunch = QRadioButton('Late')
        late_lunch.setChecked(True)
        self.lunch_timing_group.addButton(early_lunch, 0)
        self.lunch_timing_group.addButton(late_lunch, 1)
        left.addWidget(early_lunch)
        left.addWidget(late_lunch)

        hour_lunch_label = QLabel('Hour Lunches?')
        hour_lunch_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        left.addWidget(hour_lunch_label)
        self.hour_lunch_group = QButtonGroup(self)
        hour_yes = QRadioButton('Yes')
        hour_no = QRadioButton('No')
        hour_yes.setChecked(True)
        self.hour_lunch_group.addButton(hour_yes, 1)
        self.hour_lunch_group.addButton(hour_no, 0)
        left.addWidget(hour_yes)
        left.addWidget(hour_no)

        create_btn = QPushButton('Create Blank')
        create_btn.setObjectName('primaryBtn')
        create_btn.clicked.connect(self.create_schedule)
        left.addWidget(create_btn)

        self.notes_text_box = QTextEdit()
        self.notes_text_box.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding
        )
        left.addWidget(self.notes_text_box, stretch=1)

        right = QVBoxLayout()
        main_layout.addLayout(right, stretch=2)

        self.schedule_area = QWidget()
        schedule_layout = QVBoxLayout(self.schedule_area)
        schedule_layout.setContentsMargins(0, 0, 0, 0)
        self.schedule_placeholder = QWidget()
        schedule_layout.addWidget(self.schedule_placeholder, stretch=1)
        right.addWidget(self.schedule_area, stretch=3)

        self.inputs_host = QWidget()
        self.inputs_layout = QHBoxLayout(self.inputs_host)
        self.inputs_layout.addStretch()
        right.addWidget(self.inputs_host)


        bottom_row = QHBoxLayout()
        bottom_row.addStretch()
        open_btn = QPushButton('Open Schedule in Excel')
        open_btn.setObjectName('primaryBtn')
        open_btn.clicked.connect(self.open_excel)
        bottom_row.addWidget(open_btn)
        close_btn = QPushButton('Close MAS')
        close_btn.setObjectName('secondaryBtn')
        close_btn.clicked.connect(self._request_close)
        bottom_row.addWidget(close_btn)
        right.addLayout(bottom_row)

        self.load_notes()

    def _destroy_schedule_widgets(self) -> None:
        if self.sheet_frame is not None:
            self.sheet_frame.deleteLater()
            self.sheet_frame = None
        if self.inputs is not None:
            self.inputs.deleteLater()
            self.inputs = None

    def update_sheet(self):
        if self.sheet_frame:
            self.sheet_frame.update_sheet()

    def get_sheet_selection(self) -> dict | None:
        if not self.sheet_frame:
            return None
        regions = self.sheet_frame.table_view.selection_regions()
        if not regions:
            print('none selected')
            return None
        region = regions[0]
        time_start, time_end = region.time_range(self.df)
        workers = region.column_names(self.df)
        return {'workers': workers, 'time_start': time_start, 'time_end': time_end}

    def add_nonstandard_shift(self, selection):
        self.df.loc[
            selection['time_start'] : selection['time_end'], selection['workers']
        ] = selection['shift']
        self.nonstandard_shifts.append(selection)
        self.update_sheet()

    def add_standard_shift(self, shift):
        if not self.sheet_frame:
            return
        regions = self.sheet_frame.table_view.selection_regions()
        if not regions:
            workers = self.paid_workers + self.volunteers
            start, end = 0, len(self.df)
        else:
            region = regions[0]
            workers = region.column_names(self.df)
            start, end = region.from_row, region.upto_row

        failed_time_slots = []
        random.shuffle(workers)

        if SHIFT_INFO[shift]['isHour']:
            failed_time_slots = self._standard_full_hour_shift(shift, workers, start, end)
        if not SHIFT_INFO[shift]['isHour']:
            failed_time_slots = self._standard_half_hour_shift(shift, workers, start, end)

        if failed_time_slots and shift:
            msg = ', '.join(failed_time_slots)
            QMessageBox.warning(self, 'Warning', f'Failed to place {shift} at:\n{msg}')
        self.update_sheet()

    def _standard_half_hour_shift(self, shift, workers, start, end):
        failed_time_slots = []
        for curr_row in self.df.index[start:end]:
            if shift in self.df.loc[curr_row].values:
                continue
            nan_set = set(self.df.columns[self.df.loc[curr_row].isna()]) & set(workers)
            workers_with_nan = list(filter(lambda x: x in nan_set, workers))
            if workers_with_nan:
                worker_to_assign = next(w for w in workers if w in workers_with_nan)
                workers.remove(worker_to_assign)
                workers.append(worker_to_assign)
                self.df.at[curr_row, worker_to_assign] = shift
            if shift not in self.df.loc[curr_row].values:
                failed_time_slots.append(curr_row)
        return failed_time_slots

    def _standard_full_hour_shift(self, shift, workers, start, end):
        failed_time_slots = []
        for index, _row in self.df.iloc[start:end:2].iterrows():
            pos = self.df.index.get_loc(index)
            next_index = self.df.index[pos + 1]
            if shift in self.df.loc[index].values and shift in self.df.loc[next_index].values:
                continue
            workers_with_nan = (
                set(self.df.columns[self.df.iloc[pos].isna()])
                & set(self.df.columns[self.df.iloc[pos + 1].isna()])
                & set(workers)
            )
            if workers_with_nan:
                worker_to_assign = next(w for w in workers if w in workers_with_nan)
                workers.remove(worker_to_assign)
                workers.append(worker_to_assign)
                if shift not in self.df.loc[index].values:
                    self.df.at[index, worker_to_assign] = shift
                if shift not in self.df.loc[next_index].values:
                    self.df.at[next_index, worker_to_assign] = shift
            if shift not in self.df.loc[index].values:
                failed_time_slots.append(index)
            if shift not in self.df.loc[index].values:
                failed_time_slots.append(next_index)
        return failed_time_slots

    def _perform_with_undo(self, action_func, *args, **kwargs):
        state_before = self.df.copy()
        result = action_func(*args, **kwargs)
        if not self.df.equals(state_before):
            self.action_history_stack.append(state_before)
            self.action_redo_stack = []
            self.update_labels()
        else:
            print('No changes detected. No state saved')
        return result

    def undo(self):
        if self.action_history_stack:
            last_state = self.action_history_stack.pop()
            current_state = self.df.copy()
            self.action_redo_stack.append(current_state)
            self.df = last_state
            self.update_sheet()
        else:
            print('nothing left to undo.')
        self.update_labels()

    def redo(self):
        if self.action_redo_stack:
            next_state = self.action_redo_stack.pop()
            current_state = self.df.copy()
            self.action_history_stack.append(current_state)
            self.df = next_state
            self.update_sheet()
        else:
            print('nothing left to redo.')
        self.update_labels()

    def swap(self):
        if not self.sheet_frame:
            return
        regions = self.sheet_frame.table_view.selection_regions()
        if len(regions) < 2:
            print('insufficient selections.')
            QMessageBox.warning(self, 'Swap Error', 'Please select 2 equal size segments.')
            return
        sel1, sel2 = regions[0], regions[1]
        if sel1.row_count != sel2.row_count or sel1.col_count != sel2.col_count:
            print('incorrect size match')
            QMessageBox.warning(self, 'Swap Error', 'Incorrect size match.')
            return

        sel1_data = [np.nan if x == '' else x for x in sel1.values_flat(self.df)]
        sel2_data = [np.nan if x == '' else x for x in sel2.values_flat(self.df)]

        sel1_time_start, sel1_time_end = sel1.time_range(self.df)
        sel2_time_start, sel2_time_end = sel2.time_range(self.df)
        sel1_workers = sel1.column_names(self.df)
        sel2_workers = sel2.column_names(self.df)

        sel1_data_loc = self.df.loc[sel1_time_start:sel1_time_end, sel1_workers]
        sel2_data_loc = self.df.loc[sel2_time_start:sel2_time_end, sel2_workers]

        sel1_df = pd.DataFrame(sel1_data)
        sel1_df.index = sel2_data_loc.index
        sel1_df.columns = sel2_data_loc.columns

        sel2_df = pd.DataFrame(sel2_data)
        sel2_df.index = sel1_data_loc.index
        sel2_df.columns = sel1_data_loc.columns

        self.df.loc[sel2_time_start:sel2_time_end, sel2_workers] = sel1_df
        self.df.loc[sel1_time_start:sel1_time_end, sel1_workers] = sel2_df
        self.update_sheet()
        self.update_labels()

    def show_balance_rules_dialog(self):
        if self.df.empty:
            QMessageBox.warning(
                self,
                'Balance Error',
                'No schedule created yet. Create a blank schedule first.',
            )
            return
        balance_rules_dialog = BalanceRulesDialog(self, on_apply=self._apply_balance_rules)
        balance_rules_dialog.show()

    def _apply_balance_rules(self, rules):
        self._perform_with_undo(lambda: self.auto_balance_shifts(rules))

    def auto_balance_shifts(self, balance_rules=None):
        balance_rules = balance_rules or default_balance_rules()
        max_iterations = 100

        for _iteration in range(max_iterations):
            found_violation = False
            for col_name, col_data in self.df.sample(frac=1,axis=1).items():
                col_series = self.df[col_name]
                for i in range(len(col_series)):
                    cell_value = col_series.iloc[i]
                    row_label = col_series.index[i]
                    is_nan = pd.isna(cell_value)
                    if not is_nan and cell_value not in SWAPPABLE_FLOOR_SHIFTS:
                        continue
                    col_violations_before = count_column_violations(col_series, balance_rules)
                    if col_violations_before == tuple(0 for _ in balance_rules):
                        continue
                    temp = col_series.copy()
                    temp.iloc[i] = np.nan
                    if count_column_violations(temp, balance_rules) >= col_violations_before:
                        continue
                    best_swap_col = None
                    best_score = total_violations(self.df, balance_rules)
                    for other_col_name, other_col_data in self.df.sample(frac=1,axis=1).items():
                        if other_col_name == col_name:
                            continue
                        candidate_value = self.df.at[row_label, other_col_name]
                        candidate_is_nan = pd.isna(candidate_value)
                        if not candidate_is_nan and candidate_value not in SWAPPABLE_FLOOR_SHIFTS:
                            continue
                        sim_df = self.df.copy()
                        sim_df.at[row_label, col_name] = candidate_value
                        sim_df.at[row_label, other_col_name] = cell_value
                        score = total_violations(sim_df, balance_rules)
                        if score < best_score:
                            best_score = score
                            best_swap_col = other_col_name
                    if best_swap_col is not None:
                        orig_value = self.df.at[row_label, col_name]
                        self.df.at[row_label, col_name] = self.df.at[row_label, best_swap_col]
                        self.df.at[row_label, best_swap_col] = orig_value
                        found_violation = True
                        break
                if found_violation:
                    break
            if not found_violation:
                break
        else:
            QMessageBox.warning(
                self,
                'Balance Warning',
                'Could not fully resolve all conflicts. '
                'Try running Balance again or adjust the schedule manually.',
            )
        self.update_sheet()

    def update_labels(self):
        if not self.inputs:
            return
        undo_len = len(self.action_history_stack)
        redo_len = len(self.action_redo_stack)
        self.inputs.undo_button.setText(f'Undo ({undo_len})')
        self.inputs.redo_button.setText(f'Redo ({redo_len})')

    def load_notes(self):
        try:
            with open('daily_notes.txt') as file:
                self.notes_text_box.setPlainText(file.read())
        except FileNotFoundError:
            self.notes_text_box.setPlainText(
                'Error: File not found.\nPlease make a file named daily_notes.txt and\n'
                'place it in the same folder as schedule.py'
            )
        except Exception as e:
            self.notes_text_box.setPlainText(f'An error occurred: {e}')

    def save_notes(self):
        with open('daily_notes.txt', 'w') as file:
            file.write(self.notes_text_box.toPlainText())

    def create_schedule(self):
        if self.sheet_frame:
            self.action_history_stack = []
            self.action_redo_stack = []
            self.update_labels()
            self._destroy_schedule_widgets()

        self.paid_workers = self.paid_workers_entry.toPlainText().split(', ')
        self.volunteers = self.volunteers_entry.toPlainText().split(', ')
        is_late_lunch = self.lunch_timing_group.checkedId()
        is_hour_lunch = self.hour_lunch_group.checkedId()

        start = 10
        end = 17
        hours_raw_text = self.operating_hours.text()
        pattern = r'(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})'
        match = re.search(pattern, hours_raw_text)
        if match:
            start = int(match.group(1))
            end = int(match.group(3)) + 12
            end_minutes = int(match.group(4))
            if end_minutes > 0:
                end += 1

        times = pd.to_datetime(
            [datetime.time(h, m).strftime('%H:%M') for h in range(start, end) for m in (0, 30)],
            format='%H:%M',
        ).strftime('%I:%M')
        if self.volunteers[0]:
            self.df = pd.DataFrame(columns=self.paid_workers + self.volunteers, index=times)
        else:
            self.df = pd.DataFrame(columns=self.paid_workers, index=times)

        self.fill_lunch(is_late_lunch, is_hour_lunch)
        if end >= 19:
            self.fill_dinner()

        self.sheet_frame = SheetFrame(self)
        height = 415 if end > 17 else 365
        self.sheet_frame.set_frame_size(975, height)

        schedule_layout = self.schedule_area.layout()
        if self.schedule_placeholder is not None:
            schedule_layout.removeWidget(self.schedule_placeholder)
            self.schedule_placeholder.deleteLater()
            self.schedule_placeholder = None
        schedule_layout.insertWidget(0, self.sheet_frame)

        self.inputs = InputFrame(self)
        while self.inputs_layout.count():
            item = self.inputs_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.inputs_layout.addWidget(self.inputs)

        self.update_sheet()

    def fill_lunch(self, is_late_lunch, is_hour_lunch):
        possible_lunch_times = ['11:00', '11:30', '12:00', '12:30', '01:00', '01:30']
        if is_late_lunch:
            possible_lunch_times = ['01:00', '01:30', '12:00', '12:30', '11:00', '11:30']
        base_lunch_times = [time for time in possible_lunch_times if time in self.df.index]

        worker_groups = [self.paid_workers]
        if self.volunteers[0]:
            worker_groups.append(self.volunteers)

        for workers in worker_groups:
            workers = list(workers)
            random.shuffle(workers)
            lunch_times = list(base_lunch_times)
            for worker in workers:
                pos = self.df.index.get_loc(lunch_times[0])
                if is_hour_lunch:
                    self.df.at[self.df.index[pos], worker] = 'Lunch'
                    self.df.at[self.df.index[pos + 1], worker] = 'Lunch'
                    lunch_times.append(lunch_times.pop(0))
                    lunch_times.append(lunch_times.pop(0))
                else:
                    if lunch_times[0] == '11:00':
                        lunch_times.pop(0)
                        pos = self.df.index.get_loc(lunch_times[0])
                    self.df.at[self.df.index[pos], worker] = 'Lunch'
                    lunch_times.append(lunch_times.pop(0))

    def fill_dinner(self):
        possible_dinner_times = ['05:00', '05:30', '06:00', '06:30']
        base_dinner_times = [time for time in possible_dinner_times if time in self.df.index]

        worker_groups = [self.paid_workers]
        if self.volunteers[0]:
            worker_groups.append(self.volunteers)

        for workers in worker_groups:
            workers = list(workers)
            random.shuffle(workers)
            dinner_times = list(base_dinner_times)
            for worker in workers:
                pos = self.df.index.get_loc(dinner_times[0])
                # if dinner_times[0] == '05:00':
                #     dinner_times.pop(0)
                #     pos = self.df.index.get_loc(dinner_times[0])
                self.df.at[self.df.index[pos], worker] = 'Dinner'
                dinner_times.append(dinner_times.pop(0))

    def make_excel_file(self):
        self.df.index.name = f'{datetime.date.today().month}/{datetime.date.today().day}'
        with pd.ExcelWriter(RES_FILE_NAME, mode='w') as writer:
            self.df.to_excel(writer, sheet_name='Sheet1')
        wb = load_workbook(RES_FILE_NAME)
        ws = wb.active
        for cells_in_row in ws.iter_rows(min_row=2, max_col=len(self.df.columns) + 1):
            for cell in cells_in_row:
                cell_color = SHIFT_INFO.get(cell.internal_value)
                if cell_color:
                    raw = cell_color['color']
                    fg = raw[1:] if raw.startswith('#') else raw
                    cell.fill = PatternFill(patternType='solid', fgColor=fg)
        wb.save(RES_FILE_NAME)

    def open_excel(self):
        self.make_excel_file()
        try:
            self.open_file(RES_FILE_NAME)
        except FileNotFoundError:
            QMessageBox.warning(
                self, 'Excel Error', 'No schedule created yet, cannot open in Excel.'
            )
        except PermissionError:
            QMessageBox.warning(
                self,
                'Excel Error',
                'Please close the current excel sheet before opening the new one.',
            )
        except Exception as e:
            QMessageBox.warning(self, 'Excel Error', str(e))

    def open_file(self, filename: str) -> None:
        if sys.platform == 'win32':
            os.startfile(filename)
        else:
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.call([opener, filename])

    def _request_close(self):
        self.close()

    def closeEvent(self, event: QCloseEvent):
        reply = QMessageBox.question(
            self,
            'Quit',
            'Are you sure you want to quit?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.save_notes()
            event.accept()
        else:
            event.ignore()

def show_ui() -> None:
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_STYLESHEET)
    window = ScheduleApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    show_ui()
