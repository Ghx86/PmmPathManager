# -*- coding: utf-8 -*-
import os
from pathlib import Path
from source.epm_set import *

def setup_dotnet8_runtime():
    try:
        from pythonnet import load  # type: ignore
        load("coreclr")
        return True
    except Exception as e:
        print(f".NETランタイム設定エラー: {e}")
        return False

class PMMProcessor:
    
    def __init__(self):
        self.pmm_file_path = ""
        self.model_paths = []
        self.accessory_paths = []
        self.original_model_paths = []
        self.original_accessory_paths = []
        self.model_names = []
        self.accessory_names = []
        self.resolved_model_paths = []
        self.resolved_accessory_paths = []
        self.mmd_base_path = ""
        self.mmd_dest_path = ""
        self.pmm_keys = []
        self.pmm_sections = []
        self.model_keys = []
        self.model_sections = []
        self.accessory_keys = []
        self.accessory_sections = []
        self.input_pmm_content = ""
        self.output_pmm_content = ""
        self.dotnet_initialized = False
        self.dll_loaded = False
        self.media_paths = []
        self.media_types = []
    
    def initialize_dotnet(self):
        if self.dotnet_initialized:
            return {"success": True, "message": ".NET既に初期化済み"}
        
        if not setup_dotnet8_runtime():
            return {"success": False, "message": ".NETランタイム設定に失敗しました"}
        
        try:
            import clr  # type: ignore
            dll_path = Path(__file__).parent.parent / "lib" / "MikuMikuMethods.dll"
            if not dll_path.exists():
                return {"success": False, "message": f"DLLが見つかりません: {dll_path}"}
            
            try:
                clr.ClearCache()
            except Exception:
                pass
            
            try:
                clr.AddReference(str(dll_path))
            except Exception:
                try:
                    clr.AddReferenceToFileAndPath(str(dll_path))
                except Exception as e:
                    return {"success": False, "message": f"DLL参照追加失敗: {e}"}
            
            self.dotnet_initialized = True
            self.dll_loaded = True
            return {"success": True, "message": "DLL読み込み成功"}
            
        except Exception as e:
            return {"success": False, "message": f"DLL読み込みエラー: {e}"}
    
    def set_mmd_src_path(self, path):
        if path:
            if path.lower().endswith('.exe'):
                self.mmd_base_path = str(Path(path).parent)
            elif os.path.isdir(path):
                self.mmd_base_path = path
            else:
                self.mmd_base_path = ""
        else:
            self.mmd_base_path = ""
    
    def set_mmd_dest_path(self, path):
        if path:
            if path.lower().endswith('.exe') and os.path.exists(path):
                self.mmd_dest_path = str(Path(path).parent)
            elif os.path.isdir(path):
                self.mmd_dest_path = path
            else:
                self.mmd_dest_path = ""
        else:
            self.mmd_dest_path = ""
    
    def resolve_paths_with_mmd(self, paths):
        if not self.mmd_base_path:
            return paths
        
        resolved_paths = []
        mmd_base = Path(self.mmd_base_path)
        
        for path_str in paths:
            path_str = path_str.strip()
            if not path_str:
                resolved_paths.append("")
                continue
                
            path = Path(path_str)
            
            if path.is_absolute():
                resolved_paths.append(str(path))
            else:
                resolved_path = mmd_base / path
                resolved_paths.append(str(resolved_path))
        
        return resolved_paths
    
    def extract_paths_from_pmm(self):
        try:
            if not self.dll_loaded:
                result = self.initialize_dotnet()
                if not result["success"]:
                    return result
            
            self.model_paths = []
            self.accessory_paths = []
            self.original_model_paths = []
            self.original_accessory_paths = []
            self.model_names = []
            self.accessory_names = []
            self.pmm_keys = []
            self.pmm_sections = []
            self.model_keys = []
            self.model_sections = []
            self.accessory_keys = []
            self.accessory_sections = []
            self.media_paths = []
            self.media_types = []
            
            import clr  # type: ignore
            from MikuMikuMethods.Pmm import PolygonMovieMaker  # type: ignore
            
            pmm = PolygonMovieMaker(self.pmm_file_path)
            
            try:
                model_count = int(pmm.Models.Count)
            except Exception:
                model_count = 0
            
            for i in range(model_count):
                m = pmm.Models[i]
                name = getattr(m, "Name", "") or ""
                path = getattr(m, "Path", "") or ""
                
                self.model_names.append(name)
                self.model_paths.append(path)
                self.original_model_paths.append(path)
            
            try:
                acc_count = int(pmm.Accessories.Count)
            except Exception:
                acc_count = 0
                
            for i in range(acc_count):
                a = pmm.Accessories[i]
                aname = getattr(a, "Name", "") or ""
                apath = getattr(a, "Path", "") or ""
                
                self.accessory_names.append(aname)
                self.accessory_paths.append(apath)
                self.original_accessory_paths.append(apath)
            
            bg = None
            try:
                bg = getattr(pmm, "BackGroundMedia", None) or getattr(pmm, "BackGround", None) or None
            except Exception:
                bg = None

            if bg is not None:
                try:
                    audio_path = str(getattr(bg, "Audio", "") or "")
                    self.media_types.append("音声ファイル")
                    self.media_paths.append(audio_path)
                except Exception:
                    self.media_types.append("音声ファイル")
                    self.media_paths.append("")
                
                try:
                    video_path = str(getattr(bg, "Video", "") or "")
                    self.media_types.append("背景AVI")
                    self.media_paths.append(video_path)
                except Exception:
                    self.media_types.append("背景AVI")
                    self.media_paths.append("")
                
                try:
                    image_path = str(getattr(bg, "Image", "") or "")
                    self.media_types.append("背景画像")
                    self.media_paths.append(image_path)
                except Exception:
                    self.media_types.append("背景画像")
                    self.media_paths.append("")
            
            if self.mmd_base_path:
                self.resolved_model_paths = self.resolve_paths_with_mmd(self.model_paths)
                self.resolved_accessory_paths = self.resolve_paths_with_mmd(self.accessory_paths)
            else:
                self.resolved_model_paths = self.model_paths.copy()
                self.resolved_accessory_paths = self.accessory_paths.copy()
            
            pmm_filename = Path(self.pmm_file_path).name
            
            return {
                "success": True,
                "message": f"{pmm_filename}: Model x{len(self.model_paths)}, Accessory x{len(self.accessory_paths)}"
            }
            
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            return {
                "success": False,
                "message": f"パス抽出中にエラーが発生しました: {str(e)}\n{tb}"
            }
    
    def get_base_data(self, gui_displayed_path, original_path):
        try:
            if not Path(original_path).is_absolute() and self.mmd_base_path:
                bd_p1 = Path(self.mmd_base_path) / original_path
            else:
                bd_p1 = Path(original_path)
        
            bd_p2 = Path(gui_displayed_path)
        
            base_data = bd_p1.relative_to(bd_p2.parent)
            return str(base_data)
        except Exception:
            return ""
    
    def is_relative_reference(self, copy_dest_path, mmd_parent_path):
        try:
            copy_dest = Path(copy_dest_path).resolve()
            mmd_parent = Path(mmd_parent_path).resolve()
            
            try:
                copy_dest.relative_to(mmd_parent)
                return True
            except ValueError:
                return False
        except Exception:
            return False
    
    def generate_new_path(self, base_data, copy_dest_path, mmd_parent_path, is_relative):
        try:
            copy_dest = Path(copy_dest_path)
            mmd_parent = Path(mmd_parent_path)
            
            if is_relative:
                relative_part = copy_dest.relative_to(mmd_parent)
                new_path = relative_part / base_data
                return str(new_path).replace('\\', '/')
            else:
                new_path = copy_dest / base_data
                return str(new_path)
        except Exception:
            return ""
    
    def truncate_path(self, path, max_length=DEFAULT_MAX_PATH_LENGTH):
        if len(path) <= max_length:
            return path
        return path[:max_length] + "..."