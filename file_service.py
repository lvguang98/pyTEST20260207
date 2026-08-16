# file_service.py
import os
import subprocess
import json
import datetime
import win32com.client
import pandas as pd
import logging
from typing import Dict, List, Any
from path_utils import path_utils


class FileService:
    def __init__(self, base_path, logger=None):
        self.BASE_PATH = base_path
        self.logger = logger or self._create_default_logger()
        self.WPS_PATH = self.find_wps_path()
        self.logger.setLevel(logging.WARNING)

    def create_enhanced_case_folder(self, base_path: str, case_info: Dict[str, Any]) -> str:
        """创建案件文件夹 —— 直接用案本号命名"""
        import os

        case_number = case_info['case_number']

        folder_name = case_number
        folder_path = os.path.join(base_path, folder_name)

        # 处理重复
        counter = 1
        while os.path.exists(folder_path):
            folder_name = f"{case_number}-{counter:02d}"
            folder_path = os.path.join(base_path, folder_name)
            counter += 1

        os.makedirs(folder_path, exist_ok=True)
        return folder_path

    def _create_default_logger(self):
        """创建默认日志器"""
        import logging
        logger = logging.getLogger('FileService')
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.INFO)
        return logger

    def open_document(self, file_path):
        """打开文档（统一使用这个方法）"""
        self.logger.info(f"打开文档: {file_path}")
        try:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                self.logger.error(f"文件不存在: {file_path}")
                return False, "文件不存在"

            # 尝试用WPS打开
            if self.WPS_PATH and os.path.exists(self.WPS_PATH):
                try:
                    self.logger.info(f"尝试用WPS打开: {self.WPS_PATH}")
                    subprocess.Popen([self.WPS_PATH, file_path])
                    self.logger.info("WPS打开成功")
                    return True, "用WPS打开成功"
                except Exception as e:
                    self.logger.error(f"WPS打开失败: {e}")

            # 尝试用Word打开
            try:
                self.logger.info("尝试用Word打开")
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = True
                word_app.Documents.Open(file_path)
                self.logger.info("Word打开成功")
                return True, "用Word打开成功"
            except Exception as e:
                self.logger.error(f"Word打开失败: {e}")

            # 尝试用系统默认程序打开
            try:
                self.logger.info("尝试用默认程序打开")
                os.startfile(file_path)
                self.logger.info("默认程序打开成功")
                return True, "用默认程序打开成功"
            except Exception as e:
                self.logger.error(f"默认程序打开失败: {e}")
                return False, "所有打开方式都失败"

        except Exception as e:
            self.logger.error(f"打开文档异常: {e}")
            return False, f"打开文件时发生错误: {str(e)}"

    def ensure_folder_exists(self, folder_path):
        """确保文件夹存在"""
        os.makedirs(folder_path, exist_ok=True)
        return folder_path

    def find_wps_path(self):
        """从原类复制，完全一样"""
        possible_paths = [
            r"C:\Program Files (x86)\Kingsoft\WPS Office\ksolaunch.exe",
            r"C:\Program Files\Kingsoft\WPS Office\ksolaunch.exe",
            os.path.join(os.environ.get('LOCALAPPDATA', ''), "Kingsoft", "WPS Office", "ksolaunch.exe"),
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None

    def get_case_folder_path(self, person_name, case_versions=None):
        """
        获取案件文件夹路径
        person_name: 人员姓名
        case_versions: 版本字典（可选）
        返回：完整的文件夹路径
        """
        if case_versions and person_name in case_versions:
            version = case_versions[person_name]
            folder_name = f"{person_name}{version:02d}"
        else:
            # 扫描现有文件夹
            existing_folders = []
            if os.path.exists(self.BASE_PATH):
                existing_folders = [
                    f for f in os.listdir(self.BASE_PATH)
                    if f.startswith(person_name) and
                       os.path.isdir(os.path.join(self.BASE_PATH, f))
                ]

            if existing_folders:
                max_version = 0
                for folder in existing_folders:
                    version_str = folder.replace(person_name, "").strip()
                    if version_str.isdigit():
                        version = int(version_str)
                        if version > max_version:
                            max_version = version
                version = max_version + 1
            else:
                version = 1

            if case_versions:
                case_versions[person_name] = version

            folder_name = f"{person_name}{version:02d}"

        return os.path.join(self.BASE_PATH, folder_name)

    def create_case_folder(self, person_name, case_versions=None):
        """
        创建案件文件夹
        返回：创建的文件夹路径
        """
        folder_path = self.get_case_folder_path(person_name, case_versions)
        os.makedirs(folder_path, exist_ok=True)
        return folder_path

    def save_to_excel(self, template_path, excel_filename, column_name, new_item, existing_items):
        """保存到Excel - 使用统一的数据目录"""
        # 使用PathUtils的数据目录
        from path_utils import path_utils
        excel_path = path_utils.get_data_path(excel_filename)

        print(f"💾 保存Excel到: {excel_path}")

        try:
            if os.path.exists(excel_path):
                existing_df = pd.read_excel(excel_path)
                existing_list = existing_df[column_name].tolist()
                all_items = list(set(existing_list + [new_item] + existing_items))
                df = pd.DataFrame(all_items, columns=[column_name])
            else:
                all_items = list(set([new_item] + existing_items))
                df = pd.DataFrame(all_items, columns=[column_name])

            df.to_excel(excel_path, index=False)
            return df[column_name].tolist()

        except Exception as e:
            print(f"[FileService ERROR] 保存到Excel失败: {str(e)}")
            return existing_items

    def search_cases_fuzzy(self, base_path: str, keyword: str) -> List[Dict[str, Any]]:
        """根据关键词模糊搜索案件（匹配文件夹名中任意部分）"""
        import os
        import re
        from datetime import datetime

        if not keyword or not keyword.strip():
            return []

        kw = keyword.strip()
        matched_cases = []

        for folder_name in os.listdir(base_path):
            folder_path = os.path.join(base_path, folder_name)
            if not os.path.isdir(folder_path):
                continue

            # 模糊匹配：关键词出现在文件夹名的任意位置
            if kw not in folder_name:
                continue

            # 解析文件夹名：张三-案本202608111234 或 张三-案本202608111234-01
            pattern = r'^(.+?)-(案本|工亡)(\d{8})(\d{4})(-\d{2})?$'
            match = re.match(pattern, folder_name)
            if match:
                name_in_folder = match.group(1)
                case_number = folder_name.rstrip(match.group(5) or '')
                id_last4 = match.group(4)
            else:
                case_number = folder_name
                name_in_folder = ''
                id_last4 = ''

            created_time = os.path.getctime(folder_path)
            created_date = datetime.fromtimestamp(created_time).strftime('%Y-%m-%d')
            file_count = len([f for f in os.listdir(folder_path) if f.endswith('.docx')])

            matched_cases.append({
                'folder_name': folder_name,
                'folder_path': folder_path,
                'case_number': case_number,
                'person_name': name_in_folder,
                'id_last4': id_last4,
                'created_date': created_date,
                'file_count': file_count,
                'has_person': self._check_has_person_record(folder_path, name_in_folder)
            })

        matched_cases.sort(key=lambda x: x['folder_name'], reverse=True)
        return matched_cases

    def search_cases_by_person_name(self, base_path: str, person_name: str) -> List[Dict[str, Any]]:
        """
        根据姓名搜索案件文件夹

        Args:
            base_path: 基础路径
            person_name: 要搜索的姓名

        Returns:
            匹配的案件列表
        """
        import os
        import re
        from datetime import datetime

        if not person_name or not person_name.strip():
            return []

        search_name = person_name.strip()
        matched_cases = []

        for folder_name in os.listdir(base_path):
            folder_path = os.path.join(base_path, folder_name)

            if not os.path.isdir(folder_path):
                continue

            # 解析文件夹名：张三-案本202608111234
            pattern = r'^(.+?)-(案本|工亡)(\d{8})(\d{4})(-\d{2})?$'
            match = re.match(pattern, folder_name)

            if match:
                name_in_folder = match.group(1)

                if name_in_folder == search_name:
                    case_number = folder_name.rstrip(match.group(5) or '')
                    id_last4 = match.group(4)
                    created_time = os.path.getctime(folder_path)
                    created_date = datetime.fromtimestamp(created_time).strftime('%Y-%m-%d')
                    file_count = len([f for f in os.listdir(folder_path)
                                      if f.endswith('.docx')])

                    matched_cases.append({
                        'folder_name': folder_name,
                        'folder_path': folder_path,
                        'case_number': case_number,
                        'person_name': name_in_folder,
                        'id_last4': id_last4,
                        'created_date': created_date,
                        'file_count': file_count,
                        'has_person': self._check_has_person_record(folder_path, name_in_folder)
                    })

        # 按创建时间倒序排列
        matched_cases.sort(key=lambda x: x['folder_name'], reverse=True)
        return matched_cases

    def _check_has_person_record(self, folder_path: str, person_name: str) -> bool:
        """检查文件夹中是否已有本人笔录"""
        import os
        import re

        for filename in os.listdir(folder_path):
            if filename.endswith('.docx'):
                # 匹配：本人谈话笔录（...）（本人）姓名.docx
                if re.search(r'本人谈话笔录.*本人.*' + re.escape(person_name), filename):
                    return True
        return False