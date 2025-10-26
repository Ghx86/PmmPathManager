# -*- coding: utf-8 -*-
import sys
import os
import subprocess
from pathlib import Path
from PySide6.QtWidgets import QLineEdit, QTableWidget, QTableWidgetItem, QPushButton, QTextEdit
from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon
from source.epm_set import Styles

def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class DragDropLineEdit(QLineEdit):
    def __init__(self, parent=None, accept_folders=False):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.accept_folders = accept_folders
        self.setStyleSheet(Styles.DRAG_DROP_LINE_EDIT)
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                file_path = urls[0].toLocalFile()
                if self.accept_folders:
                    if os.path.isdir(file_path):
                        event.acceptProposedAction()
                        self.setStyleSheet(Styles.DRAG_DROP_LINE_EDIT_HOVER)
                        return
                else:
                    if file_path.lower().endswith('.pmm'):
                        event.acceptProposedAction()
                        self.setStyleSheet(Styles.DRAG_DROP_LINE_EDIT_HOVER)
                        return
        event.ignore()
    def dragLeaveEvent(self, event):
        self.setStyleSheet(Styles.DRAG_DROP_LINE_EDIT)
    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if self.accept_folders and os.path.isdir(file_path):
                self.setText(file_path)
                event.acceptProposedAction()
            elif not self.accept_folders and file_path.lower().endswith('.pmm'):
                self.setText(file_path)
                event.acceptProposedAction()
        self.setStyleSheet(Styles.DRAG_DROP_LINE_EDIT)

class DragDropTextEdit(QTextEdit):
    def __init__(self, parent=None, accept_folders=False):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.accept_folders = accept_folders
        self.setStyleSheet(Styles.DRAG_DROP_TEXT_EDIT)
        self.setFixedHeight(50)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                file_path = urls[0].toLocalFile()
                if self.accept_folders:
                    if os.path.isdir(file_path):
                        event.acceptProposedAction()
                        self.setStyleSheet(Styles.DRAG_DROP_TEXT_EDIT_HOVER)
                        return
                else:
                    if file_path.lower().endswith('.pmm'):
                        event.acceptProposedAction()
                        self.setStyleSheet(Styles.DRAG_DROP_TEXT_EDIT_HOVER)
                        return
        event.ignore()
    
    def dragLeaveEvent(self, event):
        self.setStyleSheet(Styles.DRAG_DROP_TEXT_EDIT)
    
    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if self.accept_folders and os.path.isdir(file_path):
                self.setPlainText(file_path)
                event.acceptProposedAction()
            elif not self.accept_folders and file_path.lower().endswith('.pmm'):
                self.setPlainText(file_path)
                event.acceptProposedAction()
        self.setStyleSheet(Styles.DRAG_DROP_TEXT_EDIT)
    
    def text(self):
        return self.toPlainText()
    
    def setText(self, text):
        self.setPlainText(text)

class EditableDragDropTable(QTableWidget):
    def __init__(self, parent=None, tab_type="model"):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QTableWidget.DropOnly)
        self.tab_type = tab_type
        self.original_paths = []
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectRows)
    def set_original_paths(self, paths):
        self.original_paths = paths
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            item = self.itemAt(event.position().toPoint())
            if item:
                row = item.row()
                self.selectRow(row)
            event.acceptProposedAction()
        else:
            event.ignore()
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            item = self.itemAt(event.position().toPoint())
            if item:
                row = item.row()
                urls = event.mimeData().urls()
                for url in urls:
                    file_path = url.toLocalFile()
                    if os.path.exists(file_path):
                        win_path = file_path.replace('/', '\\')
                        path_column = 2
                        self.setItem(row, path_column, QTableWidgetItem(win_path))
                        break
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

class MediaDragDropTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QTableWidget.DropOnly)
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectRows)
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            item = self.itemAt(event.position().toPoint())
            if item:
                row = item.row()
                self.selectRow(row)
            event.acceptProposedAction()
        else:
            event.ignore()
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            item = self.itemAt(event.position().toPoint())
            if item:
                row = item.row()
                urls = event.mimeData().urls()
                for url in urls:
                    file_path = url.toLocalFile()
                    if os.path.exists(file_path):
                        extension = Path(file_path).suffix.lower()
                        media_type = self.item(row, 1).text()
                        
                        valid = False
                        if media_type == "音声ファイル" and extension == '.wav':
                            valid = True
                        elif media_type == "背景AVI" and extension == '.avi':
                            valid = True
                        elif media_type == "背景画像":
                            valid = True
                        
                        if valid:
                            win_path = file_path.replace('/', '\\')
                            path_column = 2
                            self.setItem(row, path_column, QTableWidgetItem(win_path))
                        
                        break
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

class OpenLocationButton(QPushButton):
    def __init__(self, row, table, parent=None):
        super().__init__(parent)
        self.row = row
        self.table = table
        self.setFixedSize(28, 28)
        self.normal_icon = QIcon(":/img/fol_1.png")
        self.hover_icon = QIcon(":/img/fol_2.png")
        self.pressed_icon = QIcon(":/img/fol_3.png")
        self.setIcon(self.normal_icon)
        self.setStyleSheet("border: none; background: transparent;")
        self.clicked.connect(self.open_file_location)
    
    def enterEvent(self, event):
        self.setIcon(self.hover_icon)
        super().enterEvent(event)
    
    def leaveEvent(self, event):
        self.setIcon(self.normal_icon)
        super().leaveEvent(event)
    
    def mousePressEvent(self, event):
        self.setIcon(self.pressed_icon)
        super().mousePressEvent(event)
    
    def mouseReleaseEvent(self, event):
        self.setIcon(self.normal_icon)
        super().mouseReleaseEvent(event)
    
    def open_file_location(self):
        path_column = 2
        item = self.table.item(self.row, path_column)
        if item:
            file_path = item.text().strip()
            if file_path and os.path.exists(file_path):
                try:
                    subprocess.run(['explorer', '/select,', file_path], check=False)
                except Exception:
                    pass