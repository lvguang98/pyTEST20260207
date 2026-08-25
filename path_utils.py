# path_utils.py
"""
统一的路径管理工具
保持向后兼容，不改变现有数据位置
"""
import os
import sys
from pathlib import Path


class PathUtils:
    """路径管理工具 - 单例模式"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if hasattr(self, '_initialized'):
            return

        # 应用根目录（程序所在目录）
        self._determine_app_root()

        # 用户数据目录（保持现有位置）
        self._setup_user_dirs()

        self._initialized = True
        print(f"[OK] PathUtils初始化: app_root={self.app_root}")

    def _determine_app_root(self):
        """确定程序根目录"""
        if getattr(sys, 'frozen', False):
            # 打包后的exe
            self.app_root = Path(sys.executable).parent
        else:
            # 开发环境：main.py所在目录
            self.app_root = Path(__file__).parent

        # 确保目录存在
        self.app_root.mkdir(parents=True, exist_ok=True)

    def _setup_user_dirs(self):
        """设置用户目录（保持现有结构）"""
        # 1. 配置文件目录（保持当前目录）
        self.config_dir = self.app_root / "config"

        # 2. 数据文件目录（保持当前目录）
        self.data_dir = self.app_root

        # 3. 模板资源目录
        self.resource_dir = self.app_root / "resource"
        self.template_dir = self.resource_dir / "模板文件"

        # 4. 用户存储目录（桌面）
        desktop = Path.home() / "Desktop"
        self.storage_dir = desktop / "工伤助手存储案本"

        # 确保所有目录存在
        for dir_path in [self.config_dir, self.resource_dir,
                         self.template_dir, self.storage_dir]:
            dir_path.mkdir(parents=True, exist_ok=True)

    # ============ 主要路径获取方法 ============

    def get_template_path(self, *subpaths) -> Path:
        """获取模板路径"""
        return self.template_dir.joinpath(*subpaths)

    def get_storage_path(self, *subpaths) -> Path:
        """获取存储路径"""
        return self.storage_dir.joinpath(*subpaths)

    def get_config_path(self, *subpaths) -> Path:
        """获取配置路径"""
        return self.config_dir.joinpath(*subpaths)

    def get_data_path(self, *subpaths) -> Path:
        """获取数据路径"""
        return self.data_dir.joinpath(*subpaths)

    def get_talk_template_path(self, *subpaths) -> Path:
        """获取谈话模板路径"""
        return self.template_dir.joinpath("谈话模板", *subpaths)

    def get_document_template_path(self, *subpaths) -> Path:
        """获取文书模板路径"""
        return self.template_dir.joinpath("文书模板", *subpaths)

    def get_users_file(self) -> Path:
        """获取users_api.json路径（保持原位）"""
        return self.data_dir / "users_api.json"


# 全局单例实例
path_utils = PathUtils()