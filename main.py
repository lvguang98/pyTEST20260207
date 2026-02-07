#!/usr/bin/env python3
"""
工伤助手 - 主程序入口（简化版）
将用户登录和API配置合并到一个窗口
"""

import sys
import os
import json
from datetime import datetime
from typing import Dict, Any, Optional

from PyQt5.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QCheckBox, QMessageBox, QComboBox,
    QToolButton, QStyle
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
from path_utils import path_utils


# ============================================================================
# 第一部分：UserManager（从user_manager.py迁移，增加API配置）
# ============================================================================

class UserManager:
    """用户管理器（包含API配置）"""

    def __init__(self):
        # 使用统一的路径
        self.users_file = str(path_utils.get_users_file())  # 转换为字符串保持兼容
        self.current_user = None
        self.users_data = self._load_users()
        print(f"📁 UserManager使用文件: {self.users_file}")

    def _load_users(self) -> Dict[str, Any]:
        """加载用户数据（包含API配置）"""
        if not os.path.exists(self.users_file):
            return {"users": {}, "last_user": None, "version": "2.0"}

        try:
            with open(self.users_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {"users": {}, "last_user": None, "version": "2.0"}

    def _save_users(self):
        """保存用户数据"""
        with open(self.users_file, 'w', encoding='utf-8') as f:
            json.dump(self.users_data, f, ensure_ascii=False, indent=2)

    def save_user_config(self, username: str, api_url: str = "",
                         api_key: str = "", remember_me: bool = False,
                         service: str = "DeepSeek") -> bool:
        """保存完整的用户配置"""
        try:
            if username not in self.users_data["users"]:
                self.users_data["users"][username] = {}

            user_data = self.users_data["users"][username]
            user_data.update({
                "service": service,  # 新增
                "api_url": api_url,
                "api_key": api_key,
                "remember_me": remember_me,
                "last_login": datetime.now().isoformat(),
                "configured_at": user_data.get("configured_at", datetime.now().isoformat())
            })

            if api_key:  # 如果提供了新密钥，更新时间
                user_data["configured_at"] = datetime.now().isoformat()

            # ... 其余代码不变
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

    def get_current_user(self) -> Optional[str]:
        """获取当前用户"""
        return self.current_user

    def set_current_user(self, username: str):
        """设置当前用户"""
        self.current_user = username


# ============================================================================
# 第二部分：带眼睛按钮的密码输入框
# ============================================================================

class PasswordLineEdit(QLineEdit):
    """带眼睛按钮的密码输入框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEchoMode(QLineEdit.Password)

        # 创建眼睛按钮
        self.toggle_button = QToolButton(self)
        self.toggle_button.setCursor(Qt.PointingHandCursor)
        self.toggle_button.setText("👁")  # 使用emoji图标
        self.toggle_button.setStyleSheet("border: none; background: transparent;")
        self.toggle_button.clicked.connect(self.toggle_password_visibility)

        # 更新按钮位置
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

        # 设置按钮位置（右侧居中）
        self.toggle_button.move(
            self.rect().right() - frame_width - button_size.width(),
            (self.rect().bottom() - button_size.height() + 1) // 2
        )

    def resizeEvent(self, event):
        """重写resize事件"""
        super().resizeEvent(event)
        self.update_button_position()


# ============================================================================
# 第三部分：LoginWindow（合并用户登录和API配置）
# ============================================================================

class LoginWindow(QDialog):
    """登录窗口（合并用户登录和API配置）"""

    def __init__(self, user_manager):
        super().__init__()
        self.user_manager = user_manager
        self.username = None
        self.api_config = None

        self.setup_ui()
        self.load_remembered_user()

    def setup_ui(self):
        self.setWindowTitle("工伤助手 - 登录")
        self.setFixedSize(400, 350)

        layout = QVBoxLayout()
        layout.setContentsMargins(30, 30, 30, 30)

        # 标题
        title = QLabel("工伤助手登录")
        title.setFont(QFont("微软雅黑", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        layout.addSpacing(20)

        # 用户名（必填）
        username_layout = QHBoxLayout()
        username_label = QLabel("用户名:")
        username_label.setMinimumWidth(70)
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("请输入用户名（必填）")
        username_layout.addWidget(username_label)
        username_layout.addWidget(self.username_input)
        layout.addLayout(username_layout)

        # API地址（可选） - 原来的代码：
        url_layout = QHBoxLayout()
        url_label = QLabel("API地址:")
        url_label.setMinimumWidth(70)
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("请输入API地址（可选）")
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.url_input)
        layout.addLayout(url_layout)

        # API密钥（可选，带眼睛按钮）
        key_layout = QHBoxLayout()
        key_label = QLabel("API密钥:")
        key_label.setMinimumWidth(70)
        self.key_input = PasswordLineEdit()  # 使用自定义的带眼睛密码框
        self.key_input.setPlaceholderText("请输入API密钥（可选）")
        key_layout.addWidget(key_label)
        key_layout.addWidget(self.key_input)
        layout.addLayout(key_layout)

        layout.addSpacing(10)

        # 记住我
        self.remember_checkbox = QCheckBox("记住我（下次自动填充）")
        self.remember_checkbox.setChecked(True)  # 默认选中
        layout.addWidget(self.remember_checkbox)

        layout.addSpacing(20)

        # 按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.login_btn = QPushButton("登录")
        self.login_btn.setMinimumWidth(80)
        self.login_btn.clicked.connect(self.on_login)

        self.exit_btn = QPushButton("退出")
        self.exit_btn.setMinimumWidth(80)
        self.exit_btn.clicked.connect(self.reject)

        button_layout.addWidget(self.login_btn)
        button_layout.addSpacing(10)
        button_layout.addWidget(self.exit_btn)
        button_layout.addStretch()

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def load_remembered_user(self):
        """加载记住的用户信息"""
        remembered_user = self.user_manager.get_remembered_user()
        if remembered_user:
            self.username_input.setText(remembered_user.get("username", ""))
            self.url_input.setText(remembered_user.get("api_url", ""))
            self.key_input.setText(remembered_user.get("api_key", ""))
            self.remember_checkbox.setChecked(True)

    def on_login(self):
        """登录按钮点击事件"""
        # 获取输入值
        username = self.username_input.text().strip()
        api_url = self.url_input.text().strip()
        api_key = self.key_input.text().strip()
        remember_me = self.remember_checkbox.isChecked()

        # 验证用户名（必填）
        if not username:
            QMessageBox.warning(self, "提示", "用户名不能为空！")
            self.username_input.setFocus()
            return

        # 如果API地址或密钥为空，提示用户
        if not api_url and not api_key:
            reply = QMessageBox.information(
                self,
                "AI功能提示",
                "您没有输入API信息，AI审查功能将不可用。\n\n是否继续登录？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return
        elif not api_url or not api_key:
            # 如果只输入了其中一个
            QMessageBox.information(
                self,
                "API信息不完整",
                "API地址和密钥需要同时填写才能使用AI功能。\n\n如果只填写一个，AI功能仍然不可用。"
            )

        # 保存用户配置
        self.user_manager.save_user_config(
            username=username,
            api_url=api_url,
            api_key=api_key,
            remember_me=remember_me
        )

        # 设置当前用户
        self.user_manager.set_current_user(username)

        # 设置用户名和API配置
        self.username = username
        self.api_config = {
            "service": "DeepSeek",  # 默认服务商
            "api_url": api_url,
            "api_key": api_key
        }

        # 提示AI功能状态
        if api_url and api_key:
            QMessageBox.information(self, "成功", f"登录成功，欢迎 {username}！\nAI功能已启用。")
        else:
            QMessageBox.information(self, "成功", f"登录成功，欢迎 {username}！\nAI功能未启用。")

        # 接受对话框
        self.accept()

    def get_username(self):
        """获取用户名"""
        return self.username

    def get_api_config(self):
        """获取API配置 - 直接从users_api.json获取"""
        if not self.username:
            print("❌ 用户名为空")
            return None

        if "users" not in self.user_manager.users_data:
            print("❌ 用户数据结构错误")
            return None

        user_data = self.user_manager.users_data["users"].get(self.username)

        if not user_data:
            print(f"❌ 用户 {self.username} 不存在")
            return None

        # 检查必要的API配置
        api_url = user_data.get("api_url", "")
        api_key = user_data.get("api_key", "")

        if not api_url or not api_key:
            print(f"⚠️ 用户 {self.username} API配置不完整")
            # 仍返回配置，但标记为不完整
            return {
                "service": user_data.get("service", "DeepSeek"),
                "api_url": api_url,
                "api_key": api_key,
                "configured_at": user_data.get("configured_at", ""),
                "complete": bool(api_url and api_key)
            }

        return {
            "service": user_data.get("service", "DeepSeek"),
            "api_url": api_url,
            "api_key": api_key,
            "configured_at": user_data.get("configured_at", ""),
            "complete": True
        }


# ============================================================================
# 第三部分：主程序启动逻辑
# ============================================================================

def main():
    """主函数 - 程序入口点"""
    # 创建Qt应用
    app = QApplication(sys.argv)
    app.setApplicationName("工伤助手")

    # 1. 初始化管理器
    user_manager = UserManager()

    # 2. 显示登录窗口
    login_window = LoginWindow(user_manager)
    if login_window.exec_() != 1:  # 用户取消或退出
        sys.exit(0)

    username = login_window.get_username()
    api_config = login_window.get_api_config()

    # 3. 启动主程序
    try:
        from app_main import MainWindow  # ✅ 新的导入
        # 创建主窗口实例，传递参数
        window = MainWindow(username=username, api_config=api_config)  # ✅ 使用新类名
        # 显示窗口
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"❌ 启动主程序失败: {e}")
        import traceback
        traceback.print_exc()

        # 显示错误对话框
        QMessageBox.critical(
            None,
            "启动错误",
            f"启动主程序失败:\n{str(e)}\n\n请检查程序文件是否完整。"
        )
        sys.exit(1)


# ============================================================================
# 程序入口
# ============================================================================

if __name__ == "__main__":
    main()