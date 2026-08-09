#!/usr/bin/env python3
"""
工伤助手 - 主程序入口（整合版）
所有功能整合到主窗口，不再弹出独立登录窗口
"""

import sys
import os
import io
import json
from datetime import datetime
from typing import Dict, Any, Optional

# 修复 Windows 控制台 GBK 编码无法打印 emoji 的问题
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from PyQt5.QtWidgets import (
    QApplication, QLineEdit, QToolButton, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QStyle
from path_utils import path_utils


# ============================================================================
# UserManager（用户配置管理）
# ============================================================================

class UserManager:
    """用户管理器（管理API配置的持久化存储）"""

    def __init__(self):
        self.users_file = str(path_utils.get_users_file())
        self.current_user = None
        self.users_data = self._load_users()
        print(f"[INFO] UserManager使用文件: {self.users_file}")

    def _load_users(self) -> Dict[str, Any]:
        """加载用户数据（包含API配置）"""
        if not os.path.exists(self.users_file):
            return {"users": {}, "last_user": None, "version": "2.0"}

        try:
            with open(self.users_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载用户数据失败: {e}")
            return {"users": {}, "last_user": None, "version": "2.0"}

    def _save_users(self):
        """保存用户数据"""
        try:
            with open(self.users_file, 'w', encoding='utf-8') as f:
                json.dump(self.users_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存用户数据失败: {e}")

    def save_user_config(self, username: str, api_url: str = "",
                         api_key: str = "", remember_me: bool = False,
                         service: str = "DeepSeek") -> bool:
        """保存完整的用户配置"""
        try:
            if username not in self.users_data["users"]:
                self.users_data["users"][username] = {}

            user_data = self.users_data["users"][username]
            user_data.update({
                "service": service,
                "api_url": api_url,
                "api_key": api_key,
                "remember_me": remember_me,
                "last_login": datetime.now().isoformat(),
                "configured_at": user_data.get("configured_at", datetime.now().isoformat())
            })

            if api_key:  # 如果提供了新密钥，更新时间
                user_data["configured_at"] = datetime.now().isoformat()

            self.users_data["last_user"] = username
            self._save_users()
            return True

        except Exception as e:
            print(f"保存用户配置失败: {e}")
            return False

    def get_remembered_user(self) -> Optional[Dict[str, Any]]:
        """获取记住的用户信息"""
        last_user = self.users_data.get("last_user")
        if not last_user:
            return None

        user_data = self.users_data["users"].get(last_user, {})
        if user_data.get("remember_me", False):
            return {
                "username": last_user,
                "api_url": user_data.get("api_url", ""),
                "api_key": user_data.get("api_key", "")
            }

        return None

    def get_user_list(self) -> list:
        """获取所有用户名列表"""
        return list(self.users_data.get("users", {}).keys())

    def get_current_user(self) -> Optional[str]:
        """获取当前用户"""
        return self.current_user

    def set_current_user(self, username: str):
        """设置当前用户"""
        self.current_user = username


# ============================================================================
# PasswordLineEdit（带眼睛按钮的密码输入框）
# ============================================================================

class PasswordLineEdit(QLineEdit):
    """带眼睛按钮的密码输入框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEchoMode(QLineEdit.Password)

        # 创建眼睛按钮
        self.toggle_button = QToolButton(self)
        self.toggle_button.setCursor(Qt.PointingHandCursor)
        self.toggle_button.setText("👁")
        self.toggle_button.setStyleSheet("border: none; background: transparent;")
        self.toggle_button.clicked.connect(self.toggle_password_visibility)

        self.update_button_position()

    def toggle_password_visibility(self):
        """切换密码显示/隐藏"""
        if self.echoMode() == QLineEdit.Password:
            self.setEchoMode(QLineEdit.Normal)
            self.toggle_button.setText("🔒")
        else:
            self.setEchoMode(QLineEdit.Password)
            self.toggle_button.setText("👁")

    def update_button_position(self):
        """更新按钮位置"""
        frame_width = self.style().pixelMetric(QStyle.PM_DefaultFrameWidth)
        button_size = self.toggle_button.sizeHint()

        self.setStyleSheet(f"""
            QLineEdit {{
                padding-right: {button_size.width() + frame_width + 4}px;
            }}
        """)

        self.toggle_button.move(
            self.rect().right() - frame_width - button_size.width(),
            (self.rect().bottom() - button_size.height() + 1) // 2
        )

    def resizeEvent(self, event):
        """重写resize事件"""
        super().resizeEvent(event)
        self.update_button_position()


# ============================================================================
# 主程序入口
# ============================================================================

def main():
    """主函数 - 直接启动主窗口"""
    app = QApplication(sys.argv)
    app.setApplicationName("工伤助手")

    try:
        from app_main import MainWindow
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"[ERROR] 启动主程序失败: {e}")
        import traceback
        traceback.print_exc()

        QMessageBox.critical(
            None,
            "启动错误",
            f"启动主程序失败:\n{str(e)}\n\n请检查程序文件是否完整。"
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
