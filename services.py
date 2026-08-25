"""services.py - merged backend services.

Public class names: DataService / FileService / TemplateVariableManager.
"""
import os
import subprocess
import datetime
import logging
import win32com.client
import pandas as pd
from typing import Dict, Any, Optional


class DataService:
    """数据处理服务 - 负责所有数据计算、验证、转换操作"""

    def __init__(self):
        """初始化DataService"""
        pass

    @staticmethod
    def calculate_age_from_idcard(idcard: str) -> Optional[int]:
        """
        根据身份证号计算年龄
        :param idcard: 身份证号码（15位或18位）
        :return: 年龄（整数），如果身份证无效返回None
        """
        try:
            if not idcard or len(idcard) not in (15, 18):
                return None

            # 提取出生日期
            if len(idcard) == 18:
                birth_date_str = idcard[6:14]  # YYYYMMDD
            else:  # 15位身份证
                birth_date_str = f"19{idcard[6:12]}"  # 19YYMMDD

            # 转换为日期
            birth_date = datetime.datetime.strptime(birth_date_str, "%Y%m%d")

            # 计算年龄
            today = datetime.datetime.now()
            age = today.year - birth_date.year

            # 如果今年生日还没过，减1岁
            if (today.month, today.day) < (birth_date.month, birth_date.day):
                age -= 1

            return age

        except Exception as e:
            print(f"[DataService] 计算年龄失败: {e}")
            return None

    @staticmethod
    def extract_gender_from_idcard(idcard: str) -> Optional[str]:
        """
        根据身份证号提取性别
        :param idcard: 身份证号码（15位或18位）
        :return: "男" 或 "女"，如果身份证无效返回None
        """
        try:
            if not idcard:
                return None

            if len(idcard) == 18:
                gender_digit = int(idcard[16])  # 第17位
            elif len(idcard) == 15:
                gender_digit = int(idcard[14])  # 第15位
            else:
                return None

            # 奇数男性，偶数女性
            return "男" if gender_digit % 2 == 1 else "女"

        except Exception as e:
            print(f"[DataService] 提取性别失败: {e}")
            return None


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

    def save_to_excel(self, template_path, excel_filename, column_name, new_item, existing_items):
        """保存到Excel - 使用文书模板目录"""
        from path_utils import path_utils
        excel_path = path_utils.get_document_template_path(excel_filename)

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


class TemplateVariableManager:
    """模板变量管理器"""

    def __init__(self, data_model):
        self.data = data_model
        self.variables_cache: Dict[str, Any] = {}
        self.introduction_cache: Dict[str, str] = {}

    def clear_cache(self):
        """清空所有缓存（切换证人/角色等数据变化时调用，避免命中旧数据）"""
        self.variables_cache.clear()
        self.introduction_cache.clear()
