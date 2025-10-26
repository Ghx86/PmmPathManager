# -*- coding: utf-8 -*-
SUPPORTED_ENCODINGS = ['shift_jis', 'cp932', 'utf-8', 'utf-8-sig']
SUPPORTED_EXTENSIONS = {
   'pmm': ['.pmm'],
   'executable': ['.exe'],
   'object': ['.x', '.pmx', '.pmd'],
   'effect': ['.fx', '.fxsub'],
}

CONFIG_FILE_NAME = "config.txt"
DEFAULT_OUTPUT_NAME = "_out"


def load_output_name():
    output_name = DEFAULT_OUTPUT_NAME
    try:
        with open(CONFIG_FILE_NAME, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("output_name") and "=" in line:
                    _, value = line.split("=", 1)
                    output_name = value.strip().strip('"')
                    break
    except FileNotFoundError:
        pass
    return output_name

def save_output_name(new_output_name):
    try:
        config_data = {}
        try:
            with open(CONFIG_FILE_NAME, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line:
                        key, value = line.split("=", 1)
                        config_data[key.strip()] = value.strip().strip('"')
        except FileNotFoundError:
            pass
        config_data["output_name"] = new_output_name
        
        with open(CONFIG_FILE_NAME, "w", encoding="utf-8") as f:
            for key, value in config_data.items():
                f.write(f"{key} = {value}\n")
        return True
    except Exception as e:
        print(f"エラー")
        return False


O_NAME = load_output_name()

DEFAULT_MAX_PATH_LENGTH = 50
OUTPUT_PMM_SUFFIX = f"{O_NAME}.pmm"
ICON_FILE_NAME = "../icon/PPM.ico"

MAX_UP_LEVELS = 3

WINDOW_DEFAULT_WIDTH = 850
WINDOW_DEFAULT_HEIGHT = 900
WINDOW_TABLE_HEIGHT = 620
WINDOW_TITLE = "Pmm Path Manager"

CONFIG_RELOAD_DELAY_MS = 100
FILE_WATCHER_RESTART_DELAY_MS = 200

class Messages:
    INITIAL_STATUS = "移行元exe,PMMを設定でインポート有効 / 移行先exeを設定で下ボタン有効(推奨)"
    PMM_FILE_SELECTED = "PMM選択完了"
    PATH_EXTRACT_ERROR = "パス抽出エラー"
    PATH_REWRITE_ERROR = "パス書換エラー"
    
    LEVEL_UP_COMPLETED = "パスを1階層上げました"
    LEVEL_DOWN_COMPLETED = "パスを1階層下げました"
    CLEAR_ROW_COMPLETED = "選択行を空欄にしました"
    RESET_ROW_COMPLETED = "選択行をリセットしました"
    
    WARNING_NO_PMM_FILE = "PMMを選択してください"
    WARNING_NO_SELECTION = "選択されていません"
    WARNING_NO_MMD_DEST_SET = "移行先exeパスが設定されていません"
    WARNING_NO_PMM_SET = "PMMが設定されていません"
    WARNING_NO_REWRITE_SELECTION = "置換対象行が選択されていません"
    
    ERROR_ENCODING_DETECTION = "ファイルのエンコーディングを特定できませんでした"
    ERROR_OUT_PMM_NOT_FOUND = "_out.pmmファイルが見つかりません。先にパス抽出を実行してください"
    ERROR_OUT_PMM_READ_FAILED = "_out.pmmファイルの読み込みに失敗しました"

class Placeholders:
    MMD_SRC_PATH = "パスを入力"
    MMD_DEST_PATH = "パスを入力"
    PMM_FILE = ".pmm D&Dまたは選択"

class DialogTitles:
    SELECT_MMD_SRC = "移行前MMD.exeを選択"
    SELECT_MMD_DEST = "移行後MMD.exeを選択"
    SELECT_PMM_FILE = "PMMファイルを選択"
    WARNING = "警告"
    ERROR = "エラー"

class FileFilters:
    PMM_FILES = "PMM Files (*.pmm);;All Files (*)"
    EXECUTABLE_FILES = "Executable Files (*.exe);;All Files (*)"

class TableSettings:
    COLUMN_LABELS = ["No.", "パス"]
    HEADER_VISIBLE = False
    SELECTION_BEHAVIOR = "SelectRows"

class Styles:
    DRAG_DROP_LINE_EDIT = """
        QLineEdit {
            border: 1px solid #cccccc;
            border-radius: 0px;
            padding: 5px;
        }
        QLineEdit:focus {
            border: 2px solid #2196F3;
        }
    """
    
    DRAG_DROP_LINE_EDIT_HOVER = """
        QLineEdit {
            border: 2px dashed #2196F3;
            border-radius: 0px;
            padding: 5px;
            background-color: #f0f8f0;
        }
    """
    
    DRAG_DROP_TEXT_EDIT = """
        QTextEdit {
            border: 1px solid #cccccc;
            border-radius: 0px;
            padding: 5px;
            font-family: 'Segoe UI', Arial;
            font-size: 11px;
        }
        QTextEdit:focus {
            border: 2px solid #2196F3;
        }
    """
    
    DRAG_DROP_TEXT_EDIT_HOVER = """
        QTextEdit {
            border: 2px dashed #2196F3;
            border-radius: 0px;
            padding: 5px;
            background-color: #f0f8f0;
            font-family: 'Segoe UI', Arial;
            font-size: 11px;
        }
    """
    
    MAIN_BUTTON = """
        QPushButton {
            background-color: #4A4A4A;
            color: white;
            border: none;
            padding: 8px;
            border-radius: 4px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #2196F3;
        }
        QPushButton:disabled {
            background-color: #cccccc;
        }
    """
    
    EXTRACT_BUTTON = """
        QPushButton {
            background-color: #4A4A4A;
            color: white;
            border: none;
            padding: 8px;
            border-radius: 4px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #2196F3;
        }
        QPushButton:disabled {
            background-color: #cccccc;
        }
    """
    
    CONTROL_BUTTON = """
        QPushButton {
            background-color: #4A4A4A;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 3px;
            font-size: 12px;
        }
        QPushButton:hover {
            background-color: #2196F3;
        }
        QPushButton:disabled {
            background-color: #cccccc;
        }
    """
    
    PROGRESS_BAR = """
        QProgressBar {
            border: 2px solid grey;
            border-radius: 5px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #4CAF50;
            width: 20px;
        }
    """
    
    TABLE_WIDGET = """
        QTableWidget {
            gridline-color: transparent;
            font-family: 'Segoe UI', Arial;
            font-size: 11px;
            border: none;
            alternate-background-color: #323232;
            background-color: #242424;
        }
        QTableWidget::item {
            padding: 5px 0px;
            border: none;
            background-color: #242424;
        }
        QTableWidget::item:alternate {
            background-color: #323232;
        }
        QTableWidget::item:selected {
            background-color: #1f5aa7;
            color: white;
        }
        QHeaderView::section {
            background-color: #fafafa;
            padding: 6px;
            border: none;
            font-weight: bold;
            font-size: 12px;
            color: #333;
        }
    """

    TEXT_AREA = """
        QTextEdit {
            font-family: 'Segoe UI', Arial;
            font-size: 11px;
            border: none;
            background-color: #242424;
            color: white;
        }
    """

class ConfigFormat:
    KEY_EXE_OLD = "exe_old"
    KEY_EXE_NEW = "exe_new"
    KEY_OUTPUT_NAME = "output_name"
    ENCODING = 'utf-8'

EXCLUDED_VALUES = ['none', 'true', 'false']

class ButtonTexts:
    BROWSE = SELECT = "選択"
    EXTRACT = "インポート / 再更新"
    UNIMPLEMENTED_1 = "未実装1"
    UNIMPLEMENTED_2 = "未実装2"
    LEVEL_UP = "上"
    LEVEL_DOWN = "下"
    CLEAR_ROW = "none"
    RESET_ROW = "リセット"

class LabelTexts:
    MMD_SRC = "移行元exe:"
    MMD_DEST = "移行先exe:"
    PATH_LIST = "F2キーで編集"
    TAB_MODEL = "モデル"
    TAB_ACCESSORY = "アクセサリ"
    TAB_MEDIA = "メディア"