"""
config_service.py
配置管理服务 - 集中管理应用程序的所有配置
修复版本：不依赖 psutil
"""

import os
import sys
import json
import copy
import logging
import shutil
from dataclasses import dataclass, asdict, field
from enum import Enum
from datetime import datetime
import traceback
import platform
from typing import Dict, Any, Optional, List

# ============================================================================
# 配置数据类定义（保持不变）
# ============================================================================
@dataclass
class DatabaseConfig:
    """数据库配置"""
    type: str = "json"  # json, sqlite, mysql
    path: str = "data"
    filename: str = "case_data.json"
    backup_count: int = 10
    auto_backup: bool = True
    backup_interval_hours: int = 24

@dataclass
class TemplateConfig:
    """模板配置"""
    # 使用相对路径，会在 get_full_path 中转换为绝对路径
    base_path: str = ""  # 留空，在运行时设置
    talk_template_subpath: str = "谈话模板"
    document_template_subpath: str = "文书模板"
    backup_enabled: bool = True
    backup_count: int = 5
    auto_update_templates: bool = False

@dataclass
class PathConfig:
    """路径配置"""
    base_storage_path: str = "工伤助手存储案本"
    desktop_path: str = os.path.join(os.path.expanduser("~"), "Desktop")
    temp_folder: str = "_temp"
    backup_folder: str = "_backups"
    log_folder: str = "_logs"
    export_folder: str = "_exports"
    template_folder: str = "resource/模板文件"

@dataclass
class FileConfig:
    """文件配置"""
    default_encoding: str = "utf-8"
    auto_clean_temp_files: bool = True
    temp_file_lifetime_hours: int = 24
    max_file_size_mb: int = 50
    allowed_extensions: List[str] = field(default_factory=lambda: [
        '.docx', '.xlsx', '.txt', '.json', '.pdf', '.jpg', '.png'
    ])
    max_backup_files: int = 100
    compress_backups: bool = True

@dataclass
class CaseConfig:
    """案件配置"""
    case_number_format: str = "案本-{date}-{seq:03d}"
    case_folder_format: str = "{person_name}_{case_number}"
    max_versions_per_case: int = 10
    auto_increment_case_number: bool = True
    default_case_type: str = "工伤案件"
    default_application_type: str = "单位申请"
    date_format: str = "%Y年%m月%d日"
    time_format: str = "%H时%M分"

@dataclass
class UISettings:
    """UI设置"""
    theme: str = "default"
    font_family: str = "Microsoft YaHei"
    font_size: int = 10
    window_width: int = 1200
    window_height: int = 800
    auto_save_interval_minutes: int = 5
    show_confirmation_dialogs: bool = True
    language: str = "zh_CN"
    show_tooltips: bool = True
    animation_enabled: bool = True

@dataclass
class ValidationConfig:
    """验证配置"""
    require_id_card_for_save: bool = True
    min_name_length: int = 2
    max_name_length: int = 20
    validate_age_range: bool = True
    min_age: int = 16
    max_age: int = 100
    require_company_name: bool = False
    validate_phone_number: bool = True
    phone_regex: str = r'^1[3-9]\d{9}$'
    id_card_regex: str = r'^[1-9]\d{5}(18|19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[1-2]\d|3[0-1])\d{3}[\dXx]$'

@dataclass
class SystemConfig:
    """系统配置"""
    debug_mode: bool = False
    log_level: str = "INFO"
    enable_auto_update: bool = False
    check_for_updates_interval_days: int = 7
    max_log_files: int = 10
    max_log_size_mb: int = 10
    enable_error_reporting: bool = True
    enable_usage_statistics: bool = False
    startup_check_disk_space: bool = True
    min_disk_space_gb: int = 1

@dataclass
class ExportConfig:
    """导出配置"""
    default_format: str = "docx"
    compress_exports: bool = True
    include_backup_files: bool = False
    export_metadata: bool = True
    watermark_enabled: bool = False
    watermark_text: str = "工伤案件管理系统"

@dataclass
class BackupConfig:
    """备份配置"""
    auto_backup: bool = True
    backup_interval_hours: int = 6
    max_backup_count: int = 30
    backup_location: str = "auto"  # auto, local, external
    compress_backups: bool = True
    include_logs_in_backup: bool = True

@dataclass
class NetworkConfig:
    """网络配置"""
    enable_network_features: bool = False
    proxy_enabled: bool = False
    proxy_host: str = ""
    proxy_port: int = 8080
    timeout_seconds: int = 30
    retry_count: int = 3

# ============================================================================
# 主配置服务类（修改为不依赖 psutil）
# ============================================================================

class ConfigService:
    """
    配置管理服务
    负责管理应用程序的所有配置设置
    """

    # 单例实例
    _instance = None

    def __new__(cls, config_dir: Optional[str] = None):
        """单例模式"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self, config_dir: Optional[str] = None):
        """初始化配置服务"""
        # 防止重复初始化
        if hasattr(self, '_initialized') and self._initialized:
            return

        # 确定配置文件目录
        self._determine_config_dir(config_dir)

        # 设置默认配置
        self._setup_default_config()

        # 设置文件路径
        self._setup_file_paths()

        # 设置日志
        self._setup_logging()

        # 加载配置
        self.load_config()

        self._initialized = True

    def _determine_config_dir(self, config_dir: Optional[str] = None):
        """确定配置文件目录"""
        if config_dir is None:
            if getattr(sys, 'frozen', False):
                # 打包后的应用
                base_dir = os.path.dirname(sys.executable)
            else:
                # 开发环境：__file__ 所在目录即应用根目录
                base_dir = os.path.dirname(os.path.abspath(__file__))

            self.config_dir = os.path.join(base_dir, "config")
        else:
            self.config_dir = config_dir

        # 确保配置目录存在
        os.makedirs(self.config_dir, exist_ok=True)

    def _setup_default_config(self):
        """设置默认配置"""
        self.default_config = {
            "database": asdict(DatabaseConfig()),
            "template": asdict(TemplateConfig()),
            "paths": asdict(PathConfig()),
            "files": asdict(FileConfig()),
            "case": asdict(CaseConfig()),
            "ui": asdict(UISettings()),
            "validation": asdict(ValidationConfig()),
            "system": asdict(SystemConfig()),
            "export": asdict(ExportConfig()),
            "backup": asdict(BackupConfig()),
            "network": asdict(NetworkConfig()),
            "version": "1.0.0",
            "last_modified": datetime.now().isoformat()
        }

    def _setup_file_paths(self):
        """设置文件路径"""
        self.config_file = os.path.join(self.config_dir, "app_config.json")
        self.user_settings_file = os.path.join(self.config_dir, "user_settings.json")
        self.template_config_file = os.path.join(self.config_dir, "template_config.json")
        self.backup_config_file = os.path.join(self.config_dir, "backup_config.json")

        # 历史版本备份
        self.config_backup_dir = os.path.join(self.config_dir, "backups")
        os.makedirs(self.config_backup_dir, exist_ok=True)

    def _setup_logging(self):
        """设置日志"""
        log_dir = os.path.join(self.config_dir, "logs")
        os.makedirs(log_dir, exist_ok=True)

        log_file = os.path.join(log_dir, "config_service.log")

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )

        self.logger = logging.getLogger(__name__)
        self.logger.info(f"ConfigService 初始化完成，配置目录: {self.config_dir}")

    # ============================================================================
    # 主要公共方法（保持不变）
    # ============================================================================

    def load_config(self) -> Dict[str, Any]:
        """
        加载配置文件

        Returns:
            配置字典
        """
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)

                # 合并默认配置和加载的配置
                self.current_config = self._deep_merge(
                    copy.deepcopy(self.default_config),
                    loaded_config
                )

                # 更新最后修改时间
                self.current_config['last_modified'] = datetime.now().isoformat()

                self.logger.info(f"配置文件加载成功: {self.config_file}")
                self.logger.debug(f"配置内容: {json.dumps(self.current_config, ensure_ascii=False, indent=2)[:500]}...")
            else:
                # 使用默认配置并保存
                self.current_config = copy.deepcopy(self.default_config)
                self.save_config()
                self.logger.info("使用默认配置并创建配置文件")

        except json.JSONDecodeError as e:
            self.logger.error(f"配置文件格式错误: {e}")
            self._backup_corrupted_config()
            self.current_config = copy.deepcopy(self.default_config)
            self.save_config()

        except Exception as e:
            self.logger.error(f"加载配置文件失败: {e}")
            self.logger.debug(traceback.format_exc())
            self.current_config = copy.deepcopy(self.default_config)

        return self.current_config

    def save_config(self, config: Optional[Dict[str, Any]] = None) -> bool:
        """
        保存配置到文件

        Args:
            config: 要保存的配置，如果为None则保存当前配置

        Returns:
            是否成功
        """
        try:
            save_config = config if config is not None else self.current_config

            # 备份旧配置文件
            self._backup_config_file()

            # 确保配置完整性
            complete_config = self._deep_merge(
                copy.deepcopy(self.default_config),
                save_config
            )

            # 更新版本和修改时间
            complete_config['version'] = self.get_version()
            complete_config['last_modified'] = datetime.now().isoformat()

            # 保存配置
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(complete_config, f, ensure_ascii=False, indent=2)

            # 更新当前配置
            self.current_config = complete_config

            self.logger.info(f"配置文件保存成功: {self.config_file}")
            self.logger.debug(f"保存的配置大小: {os.path.getsize(self.config_file)} bytes")

            return True

        except Exception as e:
            self.logger.error(f"保存配置文件失败: {e}")
            self.logger.debug(traceback.format_exc())
            return False

    def get(self, key_path: str, default: Any = None) -> Any:
        """
        获取配置值（支持点分隔路径）

        Args:
            key_path: 配置路径，如 "template.base_path"
            default: 默认值

        Returns:
            配置值
        """
        try:
            keys = key_path.split('.')
            value = self.current_config

            for key in keys:
                if isinstance(value, dict) and key in value:
                    value = value[key]
                else:
                    if self.current_config.get('system', {}).get('debug_mode', False):
                        self.logger.debug(f"配置键不存在: {key_path}，返回默认值: {default}")
                    return default

            return value
        except Exception as e:
            self.logger.warning(f"获取配置失败 {key_path}: {e}")
            return default

    def set(self, key_path: str, value: Any, save_immediately: bool = False) -> bool:
        """
        设置配置值

        Args:
            key_path: 配置路径
            value: 配置值
            save_immediately: 是否立即保存

        Returns:
            是否成功
        """
        try:
            keys = key_path.split('.')
            config = self.current_config

            # 遍历到最后一个父键
            for key in keys[:-1]:
                if key not in config:
                    config[key] = {}
                elif not isinstance(config[key], dict):
                    config[key] = {}
                config = config[key]

            # 设置值
            old_value = config.get(keys[-1])
            config[keys[-1]] = value

            self.logger.debug(f"设置配置: {key_path} = {value} (原值: {old_value})")

            # 立即保存
            if save_immediately:
                return self.save_config()

            return True

        except Exception as e:
            self.logger.error(f"设置配置失败 {key_path}: {e}")
            return False
    def get_case_config(self) -> CaseConfig:
        """获取案件配置对象"""
        return CaseConfig(**self.get('case', {}))

    def get_ui_settings(self) -> UISettings:
        """获取UI设置对象"""
        return UISettings(**self.get('ui', {}))
    def get_version(self) -> str:
        """获取当前版本"""
        return self.get('version', '1.0.0')
    def _backup_config_file(self):
        """备份配置文件"""
        try:
            if os.path.exists(self.config_file):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = os.path.join(self.config_backup_dir, f"config_backup_{timestamp}.json")

                import shutil
                shutil.copy2(self.config_file, backup_file)

                # 清理旧的备份文件
                self._cleanup_old_backups()

                self.logger.debug(f"配置文件备份成功: {backup_file}")

        except Exception as e:
            self.logger.warning(f"备份配置文件失败: {e}")

    def _backup_corrupted_config(self):
        """备份损坏的配置文件"""
        try:
            if os.path.exists(self.config_file):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                corrupted_file = os.path.join(self.config_backup_dir, f"config_corrupted_{timestamp}.json")

                import shutil
                shutil.copy2(self.config_file, corrupted_file)

                self.logger.warning(f"损坏的配置文件已备份: {corrupted_file}")

        except Exception as e:
            self.logger.error(f"备份损坏的配置文件失败: {e}")

    def _cleanup_old_backups(self, max_backups: int = 10):
        """清理旧的备份文件"""
        try:
            if os.path.exists(self.config_backup_dir):
                backup_files = []
                for file in os.listdir(self.config_backup_dir):
                    if file.startswith("config_backup_") and file.endswith(".json"):
                        file_path = os.path.join(self.config_backup_dir, file)
                        backup_files.append((file_path, os.path.getmtime(file_path)))

                # 按修改时间排序
                backup_files.sort(key=lambda x: x[1], reverse=True)

                # 删除多余的备份文件
                for file_path, _ in backup_files[max_backups:]:
                    os.remove(file_path)
                    self.logger.debug(f"删除旧的备份文件: {file_path}")

        except Exception as e:
            self.logger.warning(f"清理旧备份文件失败: {e}")
    def _deep_merge(self, base: Dict, update: Dict) -> Dict:
        """
        深度合并两个字典

        Args:
            base: 基础字典
            update: 更新字典

        Returns:
            合并后的字典
        """
        result = copy.deepcopy(base)

        for key, value in update.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._deep_merge(result[key], value)
            else:
                result[key] = value

        return result
    def __str__(self) -> str:
        """字符串表示"""
        return f"ConfigService(config_dir={self.config_dir}, version={self.get_version()})"

    def __repr__(self) -> str:
        """详细表示"""
        return f"ConfigService(config_dir={self.config_dir}, current_config={self.current_config})"

# ============================================================================
# 测试代码（修改为不依赖 psutil）
# ============================================================================