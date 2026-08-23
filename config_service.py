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
from typing import Dict, Any, Optional, List, Union
from dataclasses import dataclass, asdict, field
from enum import Enum
from datetime import datetime
import traceback
import platform

# ============================================================================
# 配置数据类定义（保持不变）
# ============================================================================

class ConfigFormat(Enum):
    """配置文件格式枚举"""
    JSON = "json"
    YAML = "yaml"

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

    def update_section(self, section: str, values: Dict[str, Any], save_immediately: bool = False) -> bool:
        """
        更新整个配置部分

        Args:
            section: 配置部分名称
            values: 新的值字典
            save_immediately: 是否立即保存

        Returns:
            是否成功
        """
        try:
            if section not in self.current_config:
                self.current_config[section] = {}

            self.current_config[section].update(values)

            self.logger.info(f"更新配置部分: {section}")
            self.logger.debug(f"更新内容: {values}")

            if save_immediately:
                return self.save_config()

            return True

        except Exception as e:
            self.logger.error(f"更新配置部分失败 {section}: {e}")
            return False

    # ============================================================================
    # 配置验证和完整性检查（修改磁盘空间检查）
    # ============================================================================

    def validate_config(self) -> Dict[str, List[str]]:
        """
        验证配置的有效性
        """
        errors = {}

        print("🔍 开始验证配置...")

        # 验证模板（打印详细调试信息）
        template_errors = self._validate_templates()
        if template_errors:
            errors['template'] = template_errors

        print("🔍 模板验证完成")

        # 验证验证规则
        validation_errors = self._validate_validation_rules()
        if validation_errors:
            errors['validation'] = validation_errors

        print("🔍 验证规则验证完成")

        # 验证系统配置
        system_errors = self._validate_system_config()
        if system_errors:
            errors['system'] = system_errors

        print("🔍 系统配置验证完成")

        if errors:
            self.logger.warning(f"配置验证发现错误: {errors}")
            print(f"⚠️ 配置验证发现错误: {errors}")
        else:
            self.logger.info("配置验证通过")
            print("✅ 配置验证通过")

        return errors

    def _validate_paths(self) -> List[str]:
        """验证路径配置"""
        errors = []

        try:
            # 检查模板路径
            template_path = self.get_full_path("templates")
            if not os.path.exists(template_path):
                errors.append(f"模板路径不存在: {template_path}")

            # 检查存储路径
            storage_path = self.get_full_path("storage")
            if not os.path.exists(storage_path):
                # 尝试创建
                try:
                    os.makedirs(storage_path, exist_ok=True)
                except:
                    errors.append(f"无法创建存储路径: {storage_path}")

            # 检查磁盘空间（使用 shutil 替代 psutil）
            if self.get('system.startup_check_disk_space', True):
                min_space_gb = self.get('system.min_disk_space_gb', 1)
                if not self._check_disk_space_safe(storage_path, min_space_gb):
                    errors.append(f"磁盘空间可能不足，需要至少 {min_space_gb}GB")

        except Exception as e:
            errors.append(f"路径验证异常: {str(e)}")

        return errors

    def _validate_templates(self) -> List[str]:
        """验证模板配置"""
        errors = []

        try:
            template_config = self.get_template_config()

            # 不要检查 template.base_path，因为它是空字符串
            # 直接使用 get_full_path 获取模板路径
            template_path = self.get_full_path("templates")

            print(f"🔍 验证模板路径: {template_path}")
            print(f"🔍 路径是否存在: {os.path.exists(template_path)}")

            # 检查谈话模板目录
            talk_template_path = os.path.join(template_path, template_config.talk_template_subpath)
            print(f"🔍 谈话模板路径: {talk_template_path}")
            print(f"🔍 是否存在: {os.path.exists(talk_template_path)}")

            if not os.path.exists(talk_template_path):
                # 尝试创建目录
                try:
                    os.makedirs(talk_template_path, exist_ok=True)
                    print(f"✅ 创建谈话模板目录: {talk_template_path}")
                except Exception as e:
                    errors.append(f"谈话模板目录不存在且无法创建: {talk_template_path} - {e}")

            # 检查文书模板目录
            doc_template_path = os.path.join(template_path, template_config.document_template_subpath)
            print(f"🔍 文书模板路径: {doc_template_path}")
            print(f"🔍 是否存在: {os.path.exists(doc_template_path)}")

            if not os.path.exists(doc_template_path):
                # 尝试创建目录
                try:
                    os.makedirs(doc_template_path, exist_ok=True)
                    print(f"✅ 创建文书模板目录: {doc_template_path}")
                except Exception as e:
                    errors.append(f"文书模板目录不存在且无法创建: {doc_template_path} - {e}")

        except Exception as e:
            errors.append(f"模板验证异常: {str(e)}")
            import traceback
            traceback.print_exc()

        return errors

    def _validate_validation_rules(self) -> List[str]:
        """验证验证规则配置"""
        errors = []

        try:
            validation_config = self.get_validation_config()

            # 检查年龄范围
            if validation_config.validate_age_range:
                if validation_config.min_age >= validation_config.max_age:
                    errors.append(f"年龄范围无效: 最小年龄({validation_config.min_age}) >= 最大年龄({validation_config.max_age})")

            # 检查姓名长度
            if validation_config.min_name_length >= validation_config.max_name_length:
                errors.append(f"姓名长度范围无效")

        except Exception as e:
            errors.append(f"验证规则验证异常: {str(e)}")

        return errors

    def _validate_system_config(self) -> List[str]:
        """验证系统配置"""
        errors = []

        try:
            system_config = self.get_system_config()

            # 检查日志配置
            if system_config.max_log_files <= 0:
                errors.append("最大日志文件数必须大于0")

            if system_config.max_log_size_mb <= 0:
                errors.append("最大日志文件大小必须大于0")

        except Exception as e:
            errors.append(f"系统配置验证异常: {str(e)}")

        return errors

    def _check_disk_space_safe(self, path: str, min_gb: int) -> bool:
        """检查磁盘空间（安全版本，不依赖 psutil）"""
        try:
            # 使用 Python 内置的 shutil
            total, used, free = shutil.disk_usage(path)
            free_gb = free / (1024**3)  # 转换为GB

            self.logger.debug(f"磁盘检查: 路径={path}, 需要={min_gb}GB, 可用={free_gb:.2f}GB")

            return free_gb >= min_gb

        except Exception as e:
            self.logger.warning(f"检查磁盘空间失败，跳过检查: {e}")
            return True  # 如果检查失败，假定空间足够

    # ============================================================================
    # 便捷获取方法（保持不变）
    # ============================================================================

    def get_database_config(self) -> DatabaseConfig:
        """获取数据库配置对象"""
        return DatabaseConfig(**self.get('database', {}))

    def get_template_config(self) -> TemplateConfig:
        """获取模板配置对象"""
        return TemplateConfig(**self.get('template', {}))

    def get_path_config(self) -> PathConfig:
        """获取路径配置对象"""
        return PathConfig(**self.get('paths', {}))

    def get_file_config(self) -> FileConfig:
        """获取文件配置对象"""
        return FileConfig(**self.get('files', {}))

    def get_case_config(self) -> CaseConfig:
        """获取案件配置对象"""
        return CaseConfig(**self.get('case', {}))

    def get_ui_settings(self) -> UISettings:
        """获取UI设置对象"""
        return UISettings(**self.get('ui', {}))

    def get_validation_config(self) -> ValidationConfig:
        """获取验证配置对象"""
        return ValidationConfig(**self.get('validation', {}))

    def get_system_config(self) -> SystemConfig:
        """获取系统配置对象"""
        return SystemConfig(**self.get('system', {}))

    def get_export_config(self) -> ExportConfig:
        """获取导出配置对象"""
        return ExportConfig(**self.get('export', {}))

    def get_backup_config(self) -> BackupConfig:
        """获取备份配置对象"""
        return BackupConfig(**self.get('backup', {}))

    def get_network_config(self) -> NetworkConfig:
        """获取网络配置对象"""
        return NetworkConfig(**self.get('network', {}))

    # ============================================================================
    # 路径管理方法（保持不变）
    # ============================================================================

    def get_full_path(self, path_type: str, *subpaths) -> str:
        """
        获取完整路径 - 委托给 path_utils

        Args:
            path_type: 路径类型 (templates, storage, config, data)
            *subpaths: 子路径

        Returns:
            完整路径字符串
        """
        try:
            # 导入 path_utils
            from path_utils import path_utils

            # 映射到 path_utils 的方法
            if path_type == "storage":
                return str(path_utils.get_storage_path(*subpaths))
            elif path_type == "templates":
                return str(path_utils.get_template_path(*subpaths))
            elif path_type == "config":
                return str(path_utils.get_config_path(*subpaths))
            elif path_type == "data":
                return str(path_utils.get_data_path(*subpaths))
            elif path_type == "temp":
                # temp 路径
                return str(path_utils.get_storage_path("_temp", *subpaths))
            elif path_type == "backup":
                # backup 路径
                return str(path_utils.get_storage_path("_backups", *subpaths))
            elif path_type == "log":
                # log 路径
                return str(path_utils.get_storage_path("_logs", *subpaths))
            elif path_type == "export":
                # export 路径
                return str(path_utils.get_storage_path("_exports", *subpaths))
            else:
                # 未知类型，默认返回存储路径
                print(f"⚠️ ConfigService: 未知路径类型 '{path_type}'，使用存储路径")
                return str(path_utils.get_storage_path(*subpaths))

        except ImportError as e:
            print(f"❌ ConfigService: 无法导入 path_utils: {e}")
            # 直接抛出异常，而不是使用后备逻辑
            raise ImportError(f"无法导入 path_utils: {e}")

    def get_template_file_path(self, filename: str, template_type: str = "talk") -> str:
        """
        获取模板文件完整路径

        Args:
            filename: 模板文件名
            template_type: 模板类型，talk 或 document

        Returns:
            完整模板文件路径
        """
        template_config = self.get_template_config()

        if template_type == "talk":
            subpath = template_config.talk_template_subpath
        elif template_type == "document":
            subpath = template_config.document_template_subpath
        else:
            raise ValueError(f"未知的模板类型: {template_type}")

        return self.get_full_path("templates", subpath, filename)

    # ============================================================================
    # 配置管理方法（保持不变）
    # ============================================================================

    def reset_to_defaults(self, section: Optional[str] = None) -> bool:
        """
        重置配置为默认值

        Args:
            section: 重置特定部分，如果为None则重置全部

        Returns:
            是否成功
        """
        try:
            if section is None:
                self.current_config = copy.deepcopy(self.default_config)
                self.logger.info("重置所有配置为默认值")
            elif section in self.default_config:
                self.current_config[section] = copy.deepcopy(self.default_config[section])
                self.logger.info(f"重置配置部分 {section} 为默认值")
            else:
                self.logger.warning(f"未知的配置部分: {section}")
                return False

            return self.save_config()

        except Exception as e:
            self.logger.error(f"重置配置失败: {e}")
            return False

    def export_config(self, export_path: str, format: ConfigFormat = ConfigFormat.JSON) -> bool:
        """
        导出配置

        Args:
            export_path: 导出路径
            format: 导出格式

        Returns:
            是否成功
        """
        try:
            config_data = self.current_config

            if format == ConfigFormat.JSON:
                with open(export_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, ensure_ascii=False, indent=2)
            elif format == ConfigFormat.YAML:
                try:
                    import yaml
                    with open(export_path, 'w', encoding='utf-8') as f:
                        yaml.dump(config_data, f, allow_unicode=True)
                except ImportError:
                    self.logger.error("PyYAML 库未安装，无法导出 YAML 格式")
                    return False
            else:
                raise ValueError(f"不支持的格式: {format}")

            self.logger.info(f"配置导出成功: {export_path}")
            return True

        except Exception as e:
            self.logger.error(f"导出配置失败: {e}")
            return False

    def import_config(self, import_path: str, merge: bool = True) -> bool:
        """
        导入配置

        Args:
            import_path: 导入文件路径
            merge: 是否与现有配置合并

        Returns:
            是否成功
        """
        try:
            if not os.path.exists(import_path):
                self.logger.error(f"导入文件不存在: {import_path}")
                return False

            # 根据文件扩展名判断格式
            ext = os.path.splitext(import_path)[1].lower()

            if ext == '.json':
                with open(import_path, 'r', encoding='utf-8') as f:
                    imported_config = json.load(f)
            elif ext in ['.yaml', '.yml']:
                try:
                    import yaml
                    with open(import_path, 'r', encoding='utf-8') as f:
                        imported_config = yaml.safe_load(f)
                except ImportError:
                    self.logger.error("PyYAML 库未安装，无法导入 YAML 格式")
                    return False
            else:
                raise ValueError(f"不支持的配置文件格式: {ext}")

            # 验证导入的配置
            validation_errors = self._validate_imported_config(imported_config)
            if validation_errors:
                self.logger.error(f"导入的配置验证失败: {validation_errors}")
                return False

            # 合并或替换配置
            if merge:
                merged_config = self._deep_merge(
                    copy.deepcopy(self.current_config),
                    imported_config
                )
            else:
                merged_config = self._deep_merge(
                    copy.deepcopy(self.default_config),
                    imported_config
                )

            return self.save_config(merged_config)

        except Exception as e:
            self.logger.error(f"导入配置失败: {e}")
            return False

    def _validate_imported_config(self, config: Dict[str, Any]) -> List[str]:
        """验证导入的配置"""
        errors = []

        # 检查必需的部分
        required_sections = ['template', 'paths', 'case']
        for section in required_sections:
            if section not in config:
                errors.append(f"缺少必需的部分: {section}")

        return errors

    def get_version(self) -> str:
        """获取当前版本"""
        return self.get('version', '1.0.0')

    def get_last_modified(self) -> datetime:
        """获取最后修改时间"""
        last_modified_str = self.get('last_modified', '')
        if last_modified_str:
            try:
                return datetime.fromisoformat(last_modified_str)
            except:
                pass
        return datetime.now()

    # ============================================================================
    # 备份和恢复（保持不变）
    # ============================================================================

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

    def restore_backup(self, backup_file: str) -> bool:
        """
        从备份恢复配置

        Args:
            backup_file: 备份文件路径

        Returns:
            是否成功
        """
        try:
            if not os.path.exists(backup_file):
                self.logger.error(f"备份文件不存在: {backup_file}")
                return False

            # 加载备份配置
            with open(backup_file, 'r', encoding='utf-8') as f:
                backup_config = json.load(f)

            # 验证备份配置
            validation_errors = self._validate_imported_config(backup_config)
            if validation_errors:
                self.logger.error(f"备份配置验证失败: {validation_errors}")
                return False

            # 恢复配置
            return self.save_config(backup_config)

        except Exception as e:
            self.logger.error(f"恢复备份失败: {e}")
            return False

    def list_backups(self) -> List[Dict[str, Any]]:
        """
        列出所有备份文件

        Returns:
            备份文件信息列表
        """
        backups = []

        try:
            if os.path.exists(self.config_backup_dir):
                for file in os.listdir(self.config_backup_dir):
                    if file.startswith("config_backup_") and file.endswith(".json"):
                        file_path = os.path.join(self.config_backup_dir, file)
                        stat = os.stat(file_path)

                        backups.append({
                            'filename': file,
                            'path': file_path,
                            'size': stat.st_size,
                            'modified': datetime.fromtimestamp(stat.st_mtime),
                            'created': datetime.fromtimestamp(stat.st_ctime)
                        })

                # 按修改时间排序（最新的在前）
                backups.sort(key=lambda x: x['modified'], reverse=True)

        except Exception as e:
            self.logger.error(f"列出备份文件失败: {e}")

        return backups

    # ============================================================================
    # 工具方法（修改系统信息获取）
    # ============================================================================

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

    def get_system_info(self) -> Dict[str, Any]:
        """获取系统信息（不依赖 psutil）"""
        sys_info = {
            "config_service": {
                "version": self.get_version(),
                "config_dir": self.config_dir,
                "config_file": self.config_file,
                "config_size_kb": round(os.path.getsize(self.config_file) / 1024, 2) if os.path.exists(self.config_file) else 0,
                "last_modified": self.get_last_modified().isoformat(),
                "backup_count": len(self.list_backups())
            },
            "system": {
                "platform": platform.platform(),
                "python_version": sys.version,
                "processor": platform.processor(),
                "machine": platform.machine(),
                "architecture": platform.architecture()[0]
            },
            "disk": {},
            "memory": {"message": "使用 shutil 获取磁盘信息，内存信息需要 psutil 库"}
        }

        try:
            # 使用 shutil 获取磁盘信息
            storage_path = self.get_full_path("storage")
            if os.path.exists(storage_path):
                try:
                    total, used, free = shutil.disk_usage(storage_path)
                    sys_info['disk'][storage_path] = {
                        "total_gb": round(total / (1024**3), 2),
                        "used_gb": round(used / (1024**3), 2),
                        "free_gb": round(free / (1024**3), 2),
                        "percent": round(used / total * 100, 2) if total > 0 else 0
                    }
                except Exception as e:
                    sys_info['disk'][storage_path] = {"error": str(e)}

            # 尝试获取系统根目录
            try:
                root_path = os.path.abspath(os.sep)
                total, used, free = shutil.disk_usage(root_path)
                sys_info['disk'][root_path] = {
                    "total_gb": round(total / (1024**3), 2),
                    "used_gb": round(used / (1024**3), 2),
                    "free_gb": round(free / (1024**3), 2),
                    "percent": round(used / total * 100, 2) if total > 0 else 0
                }
            except Exception as e:
                sys_info['disk'][root_path] = {"error": str(e)}

        except Exception as e:
            sys_info['disk'] = {"error": f"获取磁盘信息失败: {e}"}
            self.logger.warning(f"获取系统信息失败: {e}")

        return sys_info

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return copy.deepcopy(self.current_config)

    def __str__(self) -> str:
        """字符串表示"""
        return f"ConfigService(config_dir={self.config_dir}, version={self.get_version()})"

    def __repr__(self) -> str:
        """详细表示"""
        return f"ConfigService(config_dir={self.config_dir}, current_config={self.current_config})"

# ============================================================================
# 测试代码（修改为不依赖 psutil）
# ============================================================================

def test_config_service():
    """测试配置服务"""
    print("🧪 开始测试 ConfigService（不依赖 psutil）...")

    try:
        # 创建测试目录
        test_dir = os.path.join(os.path.dirname(__file__), "test_config")
        os.makedirs(test_dir, exist_ok=True)

        # 创建配置服务实例
        config = ConfigService(test_dir)

        print(f"✅ 配置服务初始化成功")
        print(f"   配置目录: {config.config_dir}")
        print(f"   配置文件: {config.config_file}")
        print(f"   版本: {config.get_version()}")

        # 测试获取配置
        template_path = config.get('template.base_path')
        print(f"✅ 获取模板路径: {template_path}")

        # 测试设置配置
        config.set('ui.font_size', 12)
        print(f"✅ 设置字体大小: 12")

        # 测试保存配置
        if config.save_config():
            print(f"✅ 配置保存成功")
        else:
            print(f"❌ 配置保存失败")

        # 测试验证配置
        errors = config.validate_config()
        if errors:
            print(f"⚠️ 配置验证发现错误: {errors}")
        else:
            print(f"✅ 配置验证通过")

        # 测试路径获取
        storage_path = config.get_full_path("storage")
        print(f"✅ 存储路径: {storage_path}")

        # 测试系统信息
        sys_info = config.get_system_info()
        print(f"✅ 系统信息获取成功")
        print(f"   平台: {sys_info['system']['platform']}")
        print(f"   Python版本: {sys_info['system']['python_version'].split()[0]}")

        # 测试备份列表
        backups = config.list_backups()
        print(f"✅ 备份文件数量: {len(backups)}")

        print("\n🎉 ConfigService 测试完成！")

        # 清理测试目录
        try:
            import shutil
            shutil.rmtree(test_dir, ignore_errors=True)
        except:
            pass

        return True

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    # 运行测试
    if test_config_service():
        print("\n✅ 所有测试通过！")
        print("💡 提示：此版本不依赖 psutil，使用 Python 内置库实现所有功能")
    else:
        print("\n❌ 测试失败！")