# -*- coding: utf-8 -*-
import sys
from pathlib import Path
from PySide6.QtCore import QFileSystemWatcher, QTimer
from source.epm_set import *

class FileHandler:
    
    def __init__(self):
        self.file_watcher = QFileSystemWatcher()
        self.config_file_change_callback = None
    
    def setup_file_watcher(self, callback):
        self.config_file_change_callback = callback
        config_path = self.get_config_path()
        
        if config_path.exists():
            self.file_watcher.addPath(str(config_path))
            self.file_watcher.fileChanged.connect(self._on_config_file_changed)
    
    def _on_config_file_changed(self, path):
        if self.config_file_change_callback:
            self.config_file_change_callback(path)
    
    def get_config_path(self):
        if getattr(sys, 'frozen', False):
            executable_path = Path(sys.executable).resolve()
        else:
            executable_path = Path(sys.argv[0]).resolve()
        
        return executable_path.parent / CONFIG_FILE_NAME
    
    def load_mmd_paths(self):
        try:
            config_path = self.get_config_path()
            if not config_path.exists():
                return "", "", DEFAULT_OUTPUT_NAME
            
            with open(config_path, 'r', encoding=ConfigFormat.ENCODING) as f:
                lines = f.readlines()
            
            exe_old = ""
            exe_new = ""
            output_name = DEFAULT_OUTPUT_NAME
            
            for line in lines:
                line = line.strip()
                if '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip().strip('"')
                    
                    if key == ConfigFormat.KEY_EXE_OLD:
                        exe_old = value
                    elif key == ConfigFormat.KEY_EXE_NEW:
                        exe_new = value
                    elif key == ConfigFormat.KEY_OUTPUT_NAME:
                        output_name = value if value else DEFAULT_OUTPUT_NAME
            
            return exe_old, exe_new, output_name
                
        except Exception as e:
            print(f"設定読み込みエラー: {e}")
            return "", "", DEFAULT_OUTPUT_NAME
    
    def save_mmd_paths(self, exe_old, exe_new, output_name=None):
        try:
            config_path = self.get_config_path()
            
            self.file_watcher.removePath(str(config_path))
            
            config_data = {}
            if config_path.exists():
                with open(config_path, 'r', encoding=ConfigFormat.ENCODING) as f:
                    for line in f:
                        line = line.strip()
                        if '=' in line:
                            key, value = line.split('=', 1)
                            config_data[key.strip()] = value.strip().strip('"')
            
            config_data[ConfigFormat.KEY_EXE_OLD] = exe_old
            config_data[ConfigFormat.KEY_EXE_NEW] = exe_new
            if output_name is not None:
                config_data[ConfigFormat.KEY_OUTPUT_NAME] = output_name
            elif ConfigFormat.KEY_OUTPUT_NAME not in config_data:
                config_data[ConfigFormat.KEY_OUTPUT_NAME] = DEFAULT_OUTPUT_NAME
            
            with open(config_path, 'w', encoding=ConfigFormat.ENCODING) as f:
                for key, value in config_data.items():
                    f.write(f"{key} = {value}\n")
            
            QTimer.singleShot(FILE_WATCHER_RESTART_DELAY_MS, lambda: self.file_watcher.addPath(str(config_path)))
            
            print(f"設定保存: exe_old='{exe_old}', exe_new='{exe_new}', output_name='{config_data.get(ConfigFormat.KEY_OUTPUT_NAME, DEFAULT_OUTPUT_NAME)}'")
                
        except Exception as e:
            print(f"設定保存エラー: {e}")
    
    def save_output_name(self, output_name):
        try:
            config_path = self.get_config_path()
            
            config_data = {}
            if config_path.exists():
                with open(config_path, 'r', encoding=ConfigFormat.ENCODING) as f:
                    for line in f:
                        line = line.strip()
                        if '=' in line:
                            key, value = line.split('=', 1)
                            config_data[key.strip()] = value.strip().strip('"')
            
            config_data[ConfigFormat.KEY_OUTPUT_NAME] = output_name
            
            self.file_watcher.removePath(str(config_path))
            
            with open(config_path, 'w', encoding=ConfigFormat.ENCODING) as f:
                for key, value in config_data.items():
                    f.write(f"{key} = {value}\n")
            
            QTimer.singleShot(FILE_WATCHER_RESTART_DELAY_MS, lambda: self.file_watcher.addPath(str(config_path)))
            
            print(f"出力名設定保存: output_name='{output_name}'")
            return True
            
        except Exception as e:
            print(f"出力名設定保存エラー: {e}")
            return False