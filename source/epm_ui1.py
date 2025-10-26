# -*- coding: utf-8 -*-
import os
import shutil
from pathlib import Path
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QFileDialog, 
                               QLabel, QLineEdit, QMessageBox, 
                               QTableWidget, QTableWidgetItem, QHeaderView, QTabWidget,
                               QGroupBox, QTextEdit, QDialog, QDialogButtonBox, QSizePolicy)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QAction

from source.epm_set import *
from source.epm_pmm import PMMProcessor
from source.epm_file import FileHandler
from source.epm_widgets import DragDropLineEdit, EditableDragDropTable, MediaDragDropTable, OpenLocationButton
from source import resources_rc

class OutputNameDialog(QDialog):
    def __init__(self, current_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle("出力ファイル名設定")
        self.setModal(True)
        self.setFixedSize(300, 120)
        
        layout = QVBoxLayout(self)
        
        layout.addWidget(QLabel("出力ファイルの接尾辞を設定:"))
        
        self.name_edit = QLineEdit()
        self.name_edit.setText(current_name)
        self.name_edit.setPlaceholderText("例: _out, _new, _modified")
        layout.addWidget(self.name_edit)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def get_output_name(self):
        return self.name_edit.text().strip()

class PathSettingDialog(QDialog):
    def __init__(self, path_type, current_path, parent=None):
        super().__init__(parent)
        self.path_type = path_type
        
        if path_type == "src":
            self.setWindowTitle("移行元exe設定")
            label_text = "移行元MMD.exeのパスを設定:"
        else:
            self.setWindowTitle("移行先exe設定")
            label_text = "移行先MMD.exeのパスを設定:"
        
        self.setModal(True)
        self.setFixedSize(500, 150)
        
        layout = QVBoxLayout(self)
        
        layout.addWidget(QLabel(label_text))
        
        path_layout = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setText(current_path)
        self.path_edit.setPlaceholderText("パスを入力")
        path_layout.addWidget(self.path_edit)
        
        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self.browse_path)
        path_layout.addWidget(browse_btn)
        
        layout.addLayout(path_layout)
        
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def browse_path(self):
        if self.path_type == "src":
            title = DialogTitles.SELECT_MMD_SRC
        else:
            title = DialogTitles.SELECT_MMD_DEST
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, title, "", FileFilters.EXECUTABLE_FILES)
        if file_path:
            self.path_edit.setText(file_path)
    
    def get_path(self):
        return self.path_edit.text().strip()

class EMMPathExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pmm_processor = PMMProcessor()
        self.file_handler = FileHandler()
        self.current_output_name = DEFAULT_OUTPUT_NAME
        
        self.config_reload_timer = QTimer()
        self.config_reload_timer.setSingleShot(True)
        self.config_reload_timer.timeout.connect(self.reload_config_from_file)
        
        self.updating_from_file = False
        
        self.mmd_src_path_edit = QLineEdit()
        self.mmd_dest_path_edit = QLineEdit()
        
        self.init_ui()
        self.file_handler.setup_file_watcher(self.on_config_file_changed)
        self.load_saved_paths()
    
    def setup_menu_bar(self):
        menubar = self.menuBar()
        try:
            menubar.setNativeMenuBar(False)
        except Exception:
            pass
        
        menubar.setContentsMargins(0, 0, 0, 0)
        
        menubar.setStyleSheet("""
            QMenuBar {
                background-color: rgb(27,27,27);
                color: white;
                min-height: 22px;
            }
            QMenuBar::item {
                padding: 0px 11px;
            }
            QMenuBar::item:selected {
                background: rgb(65,65,68);
            }
        """)
        
        self.menu_style = """
            QMenu {
                background-color: rgb(50,50,50);
                color: white;
                border: 1px solid rgb(96,96,101);
            }
            QMenu::item {
                padding: 5px 18px;
            }
            QMenu::item:selected {
                background-color: rgb(65,65,68);
            }
            QMenu::separator {
                height: 1px;
                background: rgb(96,96,100);
                margin: 2px 0px;
            }
        """
        
        settings_menu = menubar.addMenu('設定')
        settings_menu.setStyleSheet(self.menu_style)
        
        self.mmd_src_action = QAction('移行元exe', self)
        self.mmd_src_action.triggered.connect(self.show_mmd_src_dialog)
        settings_menu.addAction(self.mmd_src_action)
        
        self.mmd_dest_action = QAction('移行先exe', self)
        self.mmd_dest_action.triggered.connect(self.show_mmd_dest_dialog)
        settings_menu.addAction(self.mmd_dest_action)
        
        settings_menu.addSeparator()
        
        self.output_name_action = QAction('出力ファイル名設定', self)
        self.output_name_action.triggered.connect(self.show_output_name_dialog)
        settings_menu.addAction(self.output_name_action)
    
    def show_output_name_dialog(self):
        dialog = OutputNameDialog(self.current_output_name, self)
        if dialog.exec() == QDialog.Accepted:
            new_name = dialog.get_output_name()
            if new_name and new_name != self.current_output_name:
                self.update_output_name(new_name)
    
    def show_mmd_src_dialog(self):
        current_path = self.mmd_src_path_edit.text().strip()
        dialog = PathSettingDialog("src", current_path, self)
        if dialog.exec() == QDialog.Accepted:
            new_path = dialog.get_path()
            if new_path != current_path:
                self.mmd_src_path_edit.setText(new_path)
    
    def show_mmd_dest_dialog(self):
        current_path = self.mmd_dest_path_edit.text().strip()
        dialog = PathSettingDialog("dest", current_path, self)
        if dialog.exec() == QDialog.Accepted:
            new_path = dialog.get_path()
            if new_path != current_path:
                self.mmd_dest_path_edit.setText(new_path)
    
    def update_output_name(self, new_name):
        old_name = self.current_output_name
        self.current_output_name = new_name
        
        success = self.file_handler.save_output_name(new_name)
        
        if success:
            global O_NAME, OUTPUT_PMM_SUFFIX
            O_NAME = new_name
            OUTPUT_PMM_SUFFIX = f"{new_name}.pmm"
            
            self.add_log(f"出力ファイル名を '{old_name}' から '{new_name}' に変更しました")
            QMessageBox.information(self, "設定完了", f"出力ファイル名を '{new_name}' に設定しました")
        else:
            self.current_output_name = old_name
            QMessageBox.warning(self, "エラー", "出力ファイル名の設定に失敗しました")
    
    def load_saved_paths(self):
        try:
            exe_old, exe_new, output_name = self.file_handler.load_mmd_paths()
            
            if exe_old:
                self.mmd_src_path_edit.setText(exe_old)
            if exe_new:
                self.mmd_dest_path_edit.setText(exe_new)
            if output_name and output_name != self.current_output_name:
                self.current_output_name = output_name
                global O_NAME, OUTPUT_PMM_SUFFIX
                O_NAME = output_name
                OUTPUT_PMM_SUFFIX = f"{output_name}.pmm"
                
        except Exception as e:
            print(f"設定読み込みエラー: {e}")
    
    def save_paths(self):
        if self.updating_from_file:
            return
        
        try:
            exe_old = self.mmd_src_path_edit.text().strip().strip('"')
            exe_new = self.mmd_dest_path_edit.text().strip().strip('"')
            
            self.file_handler.save_mmd_paths(exe_old, exe_new, self.current_output_name)
        except Exception as e:
            print(f"設定保存エラー: {e}")
    
    def reload_config_from_file(self):
        if self.updating_from_file:
            return
        
        try:
            self.updating_from_file = True
            
            exe_old, exe_new, output_name = self.file_handler.load_mmd_paths()
            
            self.mmd_src_path_edit.blockSignals(True)
            self.mmd_dest_path_edit.blockSignals(True)
            
            current_old = self.mmd_src_path_edit.text()
            current_new = self.mmd_dest_path_edit.text()
            
            if current_old != exe_old:
                self.mmd_src_path_edit.setText(exe_old)
                self.pmm_processor.set_mmd_src_path(exe_old)
            
            if current_new != exe_new:
                self.mmd_dest_path_edit.setText(exe_new)
                self.pmm_processor.set_mmd_dest_path(exe_new)
            
            if output_name != self.current_output_name:
                self.current_output_name = output_name
                global O_NAME, OUTPUT_PMM_SUFFIX
                O_NAME = output_name
                OUTPUT_PMM_SUFFIX = f"{output_name}.pmm"
            
            self.mmd_src_path_edit.blockSignals(False)
            self.mmd_dest_path_edit.blockSignals(False)
            
            self.update_extract_button_state()
            self.on_selection_changed()
            
        except Exception as e:
            print(f"設定ファイル読み取りエラー: {e}")
        finally:
            self.updating_from_file = False

    def load_pmm_file_and_extract(self, pmm_file_path):
        self.pmm_processor.pmm_file_path = pmm_file_path
        self.pmm_path_edit.setText(pmm_file_path)
        
        mmd_src_path = self.mmd_src_path_edit.text().strip()
        if mmd_src_path and mmd_src_path.lower().endswith('.exe'):
            QTimer.singleShot(100, self.extract_paths)
    
    def init_ui(self):
        self.setWindowTitle(WINDOW_TITLE)
        
        icon_path = os.path.join(os.path.dirname(__file__), "..", "img", ICON_FILE_NAME)
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        self.resize(WINDOW_DEFAULT_WIDTH, WINDOW_DEFAULT_HEIGHT)
        
        self.setup_menu_bar()
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        self.mmd_src_path_edit.textChanged.connect(self.on_mmd_src_path_changed)
        self.mmd_dest_path_edit.textChanged.connect(self.on_mmd_dest_path_changed)
        
        self.setup_pmm_selection_area(layout)
        
        top_button_layout = QHBoxLayout()
        
        self.extract_btn = QPushButton(ButtonTexts.EXTRACT)
        self.extract_btn.clicked.connect(self.extract_paths)
        self.extract_btn.setStyleSheet(Styles.EXTRACT_BUTTON)
        top_button_layout.addWidget(self.extract_btn)
        
        self.write_pmm_btn = QPushButton("パス全置換")
        self.write_pmm_btn.clicked.connect(self.write_pmm_with_replacements)
        self.write_pmm_btn.setEnabled(False)
        self.write_pmm_btn.setStyleSheet(Styles.MAIN_BUTTON)
        top_button_layout.addWidget(self.write_pmm_btn)
        
        layout.addLayout(top_button_layout)
        
        self.setup_log_area(layout)
        self.setup_replace_controls(layout)
        self.setup_path_controls(layout)
        self.setup_tab_widget(layout)
        
        self.update_extract_button_state()
    
    def setup_pmm_selection_area(self, layout):
        pmm_layout = QVBoxLayout()
        pmm_row = QHBoxLayout()
        
        self.pmm_path_edit = DragDropLineEdit(accept_folders=False)
        self.pmm_path_edit.setPlaceholderText(Placeholders.PMM_FILE)
        self.pmm_path_edit.textChanged.connect(self.on_pmm_path_changed)
        
        pmm_browse_btn = QPushButton(ButtonTexts.SELECT)
        pmm_browse_btn.clicked.connect(self.browse_pmm_file)
        
        pmm_row.addWidget(self.pmm_path_edit)
        pmm_row.addWidget(pmm_browse_btn)
        
        pmm_layout.addLayout(pmm_row)
        layout.addLayout(pmm_layout)
    
    
    def setup_log_area(self, layout):
        log_layout = QVBoxLayout()
        log_layout.setContentsMargins(0, 0, 0, 0)
        log_layout.setSpacing(0)
    
        log_label = QLabel("log")
        log_layout.addWidget(log_label)
    
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFixedHeight(25)
        self.log_output.setStyleSheet(Styles.TEXT_AREA)
        self.log_output.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        log_layout.addWidget(self.log_output)
    
        layout.addLayout(log_layout)
    
    def setup_replace_controls(self, layout):
        replace_group = QGroupBox("選択した行の文字置換")
        replace_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        replace_layout = QHBoxLayout(replace_group)
        
        replace_layout.addWidget(QLabel("前:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("置換対象文字を入力")
        replace_layout.addWidget(self.search_edit)
        
        replace_layout.addWidget(QLabel("後:"))
        self.replace_edit = QLineEdit()
        self.replace_edit.setPlaceholderText("置換後の文字を入力")
        replace_layout.addWidget(self.replace_edit)
        
        replace_all_btn = QPushButton("すべて置換")
        replace_all_btn.setStyleSheet(Styles.CONTROL_BUTTON)
        replace_all_btn.clicked.connect(lambda: self.replace_text(self.search_edit.text(), self.replace_edit.text(), all_occurrences=True))
        replace_layout.addWidget(replace_all_btn)
        
        replace_next_btn = QPushButton("置換")
        replace_next_btn.setStyleSheet(Styles.CONTROL_BUTTON)
        replace_next_btn.clicked.connect(lambda: self.replace_text(self.search_edit.text(), self.replace_edit.text(), all_occurrences=False))
        replace_layout.addWidget(replace_next_btn)
        
        layout.addWidget(replace_group)
        
    
    def add_log(self, message):
        self.log_output.clear()
        self.log_output.append(message)
        self.log_output.verticalScrollBar().setValue(self.log_output.verticalScrollBar().maximum())
    
    def setup_path_controls(self, layout):
        path_control_layout = QHBoxLayout()
        
        paths_label = QLabel(LabelTexts.PATH_LIST)
        path_control_layout.addWidget(paths_label)
        
        path_control_layout.addStretch()
        
        self.level_up_btn = QPushButton(ButtonTexts.LEVEL_UP)
        self.level_up_btn.clicked.connect(self.level_up_paths)
        self.level_up_btn.setEnabled(False)
        self.level_up_btn.setStyleSheet(Styles.CONTROL_BUTTON)
        path_control_layout.addWidget(self.level_up_btn)
        
        self.level_down_btn = QPushButton(ButtonTexts.LEVEL_DOWN)
        self.level_down_btn.clicked.connect(self.level_down_paths)
        self.level_down_btn.setEnabled(False)
        self.level_down_btn.setStyleSheet(Styles.CONTROL_BUTTON)
        path_control_layout.addWidget(self.level_down_btn)
        
        self.clear_row_btn = QPushButton(ButtonTexts.CLEAR_ROW)
        self.clear_row_btn.clicked.connect(self.clear_row)
        self.clear_row_btn.setEnabled(False)
        self.clear_row_btn.setStyleSheet(Styles.CONTROL_BUTTON)
        path_control_layout.addWidget(self.clear_row_btn)
        
        self.reset_row_btn = QPushButton(ButtonTexts.RESET_ROW)
        self.reset_row_btn.clicked.connect(self.reset_row)
        self.reset_row_btn.setEnabled(False)
        self.reset_row_btn.setStyleSheet(Styles.CONTROL_BUTTON)
        path_control_layout.addWidget(self.reset_row_btn)
        
        layout.addLayout(path_control_layout)
    
    def setup_tab_widget(self, layout):
        self.tab_widget = QTabWidget()
        self.tab_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.model_tab = QWidget()
        model_layout = QVBoxLayout(self.model_tab)
        model_layout.setContentsMargins(0, 0, 0, 0)
        
        self.model_table = self.create_table(tab_type="model", editable=True, drag_drop=True)
        self.model_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        model_layout.addWidget(self.model_table)
        self.tab_widget.addTab(self.model_tab, LabelTexts.TAB_MODEL)
        
        self.accessory_tab = QWidget()
        accessory_layout = QVBoxLayout(self.accessory_tab)
        accessory_layout.setContentsMargins(0, 0, 0, 0)
        
        self.accessory_table = self.create_table(tab_type="accessory", editable=True, drag_drop=True)
        self.accessory_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        accessory_layout.addWidget(self.accessory_table)
        self.tab_widget.addTab(self.accessory_tab, LabelTexts.TAB_ACCESSORY)
        
        self.media_tab = QWidget()
        media_layout = QVBoxLayout(self.media_tab)
        media_layout.setContentsMargins(0, 0, 0, 0)
        
        self.media_table = self.create_media_table()
        self.media_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        media_layout.addWidget(self.media_table)
        self.tab_widget.addTab(self.media_tab, LabelTexts.TAB_MEDIA)
        
        self.tab_widget.currentChanged.connect(self.on_tab_changed)
        
        layout.addWidget(self.tab_widget)
    
    def write_pmm_with_replacements(self):
        if not self.pmm_processor.pmm_file_path:
            QMessageBox.warning(self, "エラー", "PMMが読み込まれていません")
            return
        
        try:
            pmm_path = Path(self.pmm_processor.pmm_file_path)
            output_pmm_path = pmm_path.parent / f"{pmm_path.stem}{OUTPUT_PMM_SUFFIX}"
            
            shutil.copy2(self.pmm_processor.pmm_file_path, output_pmm_path)
            
            result = self.pmm_processor.initialize_dotnet()
            if not result["success"]:
                QMessageBox.critical(self, "エラー", result["message"])
                return
            
            import clr
            from MikuMikuMethods.Pmm import PolygonMovieMaker # type: ignore
            
            pmm = PolygonMovieMaker(str(output_pmm_path))
            
            for row in range(self.model_table.rowCount()):
                try:
                    idx = int(self.model_table.item(row, 0).text()) - 1
                    name = self.model_table.item(row, 1).text() if self.model_table.item(row, 1) else ""
                    path = self.model_table.item(row, 2).text() if self.model_table.item(row, 2) else ""
                    
                    if idx < pmm.Models.Count:
                        try:
                            pmm.Models[idx].Name = name
                        except Exception:
                            pass
                        try:
                            pmm.Models[idx].Path = path
                        except Exception:
                            pass
                except Exception:
                    continue
            
            for row in range(self.accessory_table.rowCount()):
                try:
                    idx = int(self.accessory_table.item(row, 0).text()) - 1
                    path = self.accessory_table.item(row, 2).text() if self.accessory_table.item(row, 2) else ""
                    
                    if idx < pmm.Accessories.Count:
                        try:
                            pmm.Accessories[idx].Path = path
                        except Exception:
                            pass
                except Exception:
                    continue
            
            bg = getattr(pmm, "BackGroundMedia", None) or getattr(pmm, "BackGround", None) or None
            if bg is not None:
                try:
                    for row in range(self.media_table.rowCount()):
                        media_type = self.media_table.item(row, 1).text() if self.media_table.item(row, 1) else ""
                        path = self.media_table.item(row, 2).text() if self.media_table.item(row, 2) else ""
                        
                        if media_type == "音声ファイル":
                            try:
                                setattr(bg, "Audio", path)
                            except Exception:
                                pass
                        elif media_type == "背景AVI":
                            try:
                                setattr(bg, "Video", path)
                            except Exception:
                                pass
                        elif media_type == "背景画像":
                            try:
                                setattr(bg, "Image", path)
                            except Exception:
                                pass
                except Exception:
                    pass
            
            pmm.Write(str(output_pmm_path))
            self.add_log(f"PMM書き込み完了: {output_pmm_path}")
            QMessageBox.information(self, "完了", "PMMファイルの書き込みが完了しました")
            
        except Exception as e:
            error_msg = f"PMM書き込みエラー: {str(e)}"
            self.add_log(error_msg)
            QMessageBox.critical(self, "エラー", error_msg)
    
    def create_table(self, tab_type="model", editable=False, drag_drop=False):
        if drag_drop:
            table = EditableDragDropTable(tab_type=tab_type)
        else:
            table = QTableWidget()
            table.setAlternatingRowColors(True)
        
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["No.", "名前", "パス", ""])
        table.setColumnWidth(0, 40)
        table.setColumnWidth(1, 150)
        table.setColumnWidth(3, 30)
        
        table.horizontalHeader().setVisible(False)
        
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        
        table.verticalHeader().setVisible(False)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.selectionModel().selectionChanged.connect(self.on_selection_changed)
        
        table.setWordWrap(False)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        table.setStyleSheet(Styles.TABLE_WIDGET)
        
        if editable:
            table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            table.keyPressEvent = lambda event: self.handle_table_key_press(event, table)
        else:
            table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        
        return table
    
    def create_media_table(self):
        table = MediaDragDropTable()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["No.", "メディア種類", "ファイルパス", ""])
        table.setColumnWidth(0, 40)
        table.setColumnWidth(1, 150)
        table.setColumnWidth(3, 30)
        
        table.horizontalHeader().setVisible(False)
        
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        
        table.verticalHeader().setVisible(False)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.selectionModel().selectionChanged.connect(self.on_selection_changed)
        
        table.setWordWrap(False)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        table.setStyleSheet(Styles.TABLE_WIDGET)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.keyPressEvent = lambda event: self.handle_table_key_press(event, table)
        
        return table
    
    def handle_table_key_press(self, event, table):
        if event.key() == Qt.Key.Key_F2:
            current_item = table.currentItem()
            if current_item:
                path_column = 2
                if current_item.column() == path_column:
                    table.editItem(current_item)
        else:
            QTableWidget.keyPressEvent(table, event)
    
    def on_tab_changed(self):
        self.on_selection_changed()
    
    def get_current_table(self):
        current_index = self.tab_widget.currentIndex()
        if current_index == 0:
            return self.model_table
        elif current_index == 1:
            return self.accessory_table
        elif current_index == 2:
            return self.media_table
        else:
            return None
    
    def get_current_paths_and_originals(self):
        current_index = self.tab_widget.currentIndex()
        if current_index == 0:
            return self.pmm_processor.model_paths, self.pmm_processor.original_model_paths
        elif current_index == 1:
            return self.pmm_processor.accessory_paths, self.pmm_processor.original_accessory_paths
        elif current_index == 2:
            return self.pmm_processor.media_paths, self.pmm_processor.media_paths
        else:
            return [], []
    
    def on_selection_changed(self):
        table = self.get_current_table()
        
        if table is None:
            self.clear_row_btn.setEnabled(False)
            self.reset_row_btn.setEnabled(False)
            self.level_up_btn.setEnabled(False)
            self.level_down_btn.setEnabled(False)
            return
            
        has_selection = len(table.selectedItems()) > 0
        
        self.clear_row_btn.setEnabled(has_selection)
        self.reset_row_btn.setEnabled(has_selection)
        self.level_up_btn.setEnabled(has_selection)
        self.level_down_btn.setEnabled(has_selection)
    
    def extract_paths(self):
        if not self.pmm_processor.pmm_file_path:
            QMessageBox.warning(self, DialogTitles.WARNING, Messages.WARNING_NO_PMM_FILE)
            return
        
        try:
            result = self.pmm_processor.extract_paths_from_pmm()
            
            if result["success"]:
                self.add_log(result["message"])
                self.update_table_display()
                self.select_first_row_all_tabs()
                self.on_selection_changed()
                self.write_pmm_btn.setEnabled(True)
            else:
                self.add_log(result["message"])
                QMessageBox.critical(self, DialogTitles.ERROR, result["message"])
            
        except Exception as e:
            error_msg = f"パス抽出中エラー発生: {str(e)}"
            self.add_log(error_msg)
            QMessageBox.critical(self, DialogTitles.ERROR, error_msg)
    
    def select_first_row_all_tabs(self):
        tables = [self.model_table, self.accessory_table, self.media_table]
        for table in tables:
            if table.rowCount() > 0:
                table.selectRow(0)
    
    def update_table_display(self):
        self.update_single_table(
            self.model_table, 
            self.pmm_processor.model_paths, 
            self.pmm_processor.original_model_paths,
            self.pmm_processor.model_names,
            "model"
        )
        self.update_single_table(
            self.accessory_table, 
            self.pmm_processor.accessory_paths, 
            self.pmm_processor.original_accessory_paths,
            self.pmm_processor.accessory_names,
            "accessory"
        )
        self.update_media_table()
    
    def update_single_table(self, table, paths_to_display, original_paths, names, table_type):
        table.setRowCount(0)
        if not paths_to_display:
            return
        
        if hasattr(table, 'set_original_paths'):
            table.set_original_paths(original_paths)
        
        table.setRowCount(len(paths_to_display))
        
        for i, path in enumerate(paths_to_display):
            num_item = QTableWidgetItem(str(i + 1))
            num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            num_item.setFlags(num_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(i, 0, num_item)
            
            name = names[i] if i < len(names) else ""
            name_item = QTableWidgetItem(name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(i, 1, name_item)
            
            path_item = QTableWidgetItem(path)
            path_item.setFlags(path_item.flags() | Qt.ItemFlag.ItemIsEditable)
            table.setItem(i, 2, path_item)
            
            btn = OpenLocationButton(i, table)
            table.setCellWidget(i, 3, btn)
    
    def update_media_table(self):
        table = self.media_table
        table.setRowCount(0)
        
        if not self.pmm_processor.media_paths:
            return
        
        table.setRowCount(len(self.pmm_processor.media_paths))
        
        for i, (media_type, path) in enumerate(zip(self.pmm_processor.media_types, self.pmm_processor.media_paths)):
            num_item = QTableWidgetItem(str(i + 1))
            num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            num_item.setFlags(num_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(i, 0, num_item)
            
            type_item = QTableWidgetItem(media_type)
            type_item.setFlags(type_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(i, 1, type_item)
            
            path_item = QTableWidgetItem(path)
            path_item.setFlags(path_item.flags() | Qt.ItemFlag.ItemIsEditable)
            table.setItem(i, 2, path_item)
            
            btn = OpenLocationButton(i, table)
            table.setCellWidget(i, 3, btn)
    
    def replace_text(self, search_text, replace_text, all_occurrences=True):
        if not search_text:
            return
        
        table = self.get_current_table()
        if not table:
            return
            
        path_column = 2
        replaced_count = 0
        
        for item in table.selectedItems():
            if item.column() == path_column:
                current_text = item.text()
                if all_occurrences:
                    new_text = current_text.replace(search_text, replace_text)
                    if new_text != current_text:
                        item.setText(new_text)
                        replaced_count += 1
                else:
                    new_text = current_text.replace(search_text, replace_text, 1)
                    if new_text != current_text:
                        item.setText(new_text)
                        replaced_count += 1
                        break
    
    def clear_row(self):
        table = self.get_current_table()
        if not table:
            return
            
        path_column = 2
        
        for item in table.selectedItems():
            if item.column() == path_column:
                item.setText("none")
    
    def reset_row(self):
        table = self.get_current_table()
        if not table:
            return
            
        current_paths, original_paths = self.get_current_paths_and_originals()
        path_column = 2
        
        for item in table.selectedItems():
            if item.column() == path_column:
                row = item.row()
                if row < len(current_paths):
                    item.setText(current_paths[row])
    
    def level_up_paths(self):
        table = self.get_current_table()
        if not table:
            return
            
        path_column = 2
        modified = False
        
        for item in table.selectedItems():
            if item.column() == path_column:
                current_path = item.text().strip()
                if current_path and current_path != "none":
                    parent_path = str(Path(current_path).parent)
                    if parent_path != current_path:
                        item.setText(parent_path)
                        modified = True
    
    def level_down_paths(self):
        table = self.get_current_table()
        if not table:
            return
            
        current_paths, original_paths = self.get_current_paths_and_originals()
        path_column = 2
        
        modified = False
        
        for item in table.selectedItems():
            if item.column() == path_column:
                row = item.row()
                if row < len(current_paths):
                    current_path = item.text().strip()
                    reference_path = current_paths[row]
                    
                    if current_path and reference_path:
                        current_parts = Path(current_path).parts
                        reference_parts = Path(reference_path).parts
                        
                        if len(reference_parts) > len(current_parts):
                            next_part = reference_parts[len(current_parts)]
                            down_path = str(Path(current_path) / next_part)
                            item.setText(down_path)
                            modified = True
    
    def update_extract_button_state(self):
        mmd_src_path = self.mmd_src_path_edit.text().strip()
        has_exe_extension = mmd_src_path.lower().endswith('.exe') if mmd_src_path else False
        
        has_pmm_file = bool(self.pmm_processor.pmm_file_path)
        
        self.extract_btn.setEnabled(has_exe_extension and has_pmm_file)
    
    def browse_pmm_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, DialogTitles.SELECT_PMM_FILE, "", FileFilters.PMM_FILES)
        if file_path:
            self.pmm_processor.pmm_file_path = file_path
            self.pmm_path_edit.setText(file_path)
    
    def on_mmd_src_path_changed(self):
        if self.updating_from_file:
            return
        
        path = self.mmd_src_path_edit.text().strip()
        self.pmm_processor.set_mmd_src_path(path)
        self.save_paths()
        self.update_extract_button_state()
    
    def on_mmd_dest_path_changed(self):
        if self.updating_from_file:
            return
        
        path = self.mmd_dest_path_edit.text().strip()
        self.pmm_processor.set_mmd_dest_path(path)
        self.save_paths()
        self.on_selection_changed()
    
    def on_pmm_path_changed(self):
        path = self.pmm_path_edit.text().strip()
        if path and os.path.exists(path) and path.lower().endswith('.pmm'):
            self.pmm_processor.pmm_file_path = path
        else:
            self.pmm_processor.pmm_file_path = ""
        
        self.update_extract_button_state()
    
    def on_config_file_changed(self, path):
        if self.updating_from_file:
            return
        
        self.config_reload_timer.start(CONFIG_RELOAD_DELAY_MS)