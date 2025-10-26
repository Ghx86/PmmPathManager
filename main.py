# -*- coding: utf-8 -*-
import sys

REQUIRED_PYTHON_VERSION = (3, 11)

if sys.version_info[:2] != REQUIRED_PYTHON_VERSION:
    print(f"エラー: Pythonバージョンが3.11ではありません。")
    input("Enterキーを押して終了...")
    sys.exit(1)

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
import os
from source.epm_ui1 import EMMPathExtractor
from source.epm_set import ICON_FILE_NAME

def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

def get_dotnet_version():
    try:
        import pythonnet
        pythonnet.load("coreclr")
        
        import clr  # type: ignore
        from System import Environment  # type: ignore
        version = Environment.Version
        
        try:
            from System.Runtime.InteropServices import RuntimeInformation  # type: ignore
            framework_description = str(RuntimeInformation.FrameworkDescription)
            
            if "Framework" in framework_description:
                return f"Framework {version.Major}.{version.Minor}.{version.Build}"
            else:
                return f"{version.Major}.{version.Minor}.{version.Build}"
        except Exception:
            if version.Major == 4:
                return f"Framework {version.Major}.{version.Minor}.{version.Build}"
            else:
                return f"{version.Major}.{version.Minor}.{version.Build}"
    except Exception:
        return "取得失敗"

def main():
    dotnet_version = get_dotnet_version()
    version_log = f"現環境: .NET {dotnet_version}"
    
    app = QApplication(sys.argv)
    
    icon_path = get_resource_path(os.path.join("img", ICON_FILE_NAME))
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    window = EMMPathExtractor()
    window.show()
    
    window.add_log(version_log)
    
    if len(sys.argv) > 1:
        pmm_file_path = sys.argv[1]
        if os.path.exists(pmm_file_path) and pmm_file_path.lower().endswith('.pmm'):
            window.load_pmm_file_and_extract(pmm_file_path)
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()