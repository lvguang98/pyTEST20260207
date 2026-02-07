import os
import sys
import datetime
import logging
from pathlib import Path
from typing import Dict, List, Any, Tuple
from ctypes import *
import pandas as pd
from docx import Document
from docxtpl import DocxTemplate
from PyQt5.Qt import *
from PyQt5.QtCore import QThread, pyqtSignal, QTimer, Qt
from PyQt5.QtWidgets import (
    QApplication, QWidget, QMessageBox, QDialog, QVBoxLayout,
    QLabel, QTextEdit, QPushButton, QHBoxLayout, QInputDialog,
    QLineEdit, QComboBox, QCompleter, QCheckBox, QProgressDialog,
    QTableWidget, QTableWidgetItem, QRadioButton, QButtonGroup
)

from ui_main_window import Ui_Form
from file_service import FileService
from data_service import DataService
from template_service import TemplateService, TemplateVariableManager
from ai_service import AIService
from config_service import ConfigService
from path_utils import path_utils
from case_index import CaseIndexManager, CaseIndexEntry

# 设置日志级别
logging.getLogger('config_service').setLevel(logging.WARNING)


class AIWorker(QThread):
    """AI工作线程"""
    finished = pyqtSignal(dict)  # 发送完成信号
    error = pyqtSignal(str)  # 发送错误信号
    progress = pyqtSignal(str, int)  # 发送进度信号 (消息, 进度百分比)

    def __init__(self, ai_service, file_path):
        super().__init__()
        self.ai_service = ai_service
        self.file_path = file_path

    def run(self):
        """线程运行的主函数"""
        try:
            # 第一步：提取文本
            self.progress.emit("正在提取文档文本...", 20)
            document_text = self.ai_service.extract_text_from_docx(self.file_path)

            # 第二步：AI分析
            self.progress.emit("正在调用DeepSeek API进行分析...", 50)
            result = self.ai_service.analyze_legal_document(document_text)

            # 第三步：完成
            self.progress.emit("分析完成，正在生成报告...", 90)
            self.finished.emit(result)

        except Exception as e:
            self.error.emit(str(e))

class CaseDataModel:
    """案件数据模型 - 统一管理所有案件数据"""

    def __init__(self):
        self.basic_info: Dict[str, Any] = {}  # 基础个人信息
        self.company_info: Dict[str, Any] = {}  # 公司相关信息
        self.case_info: Dict[str, Any] = {}  # 案件信息
        self.investigation: Dict[str, Any] = {}  # 调查信息
        self.output_config: Dict[str, Any] = {}  # 输出配置
        self._init_default_values()

    def _init_default_values(self):
        """初始化默认值"""
        self.case_info.update({
            '案件性质': '工伤案件',
            '申请类型': '单位申请'
        })
        self.output_config.update({
            '当前日期': datetime.datetime.now().strftime('%Y年%m月%d日'),
            '当前时间': datetime.datetime.now().strftime('%H时%M分')
        })

    def to_template_dict(self) -> Dict[str, Any]:
        """转换为模板渲染用的字典"""
        template_dict = {}

        # 确保日期时间是最新的
        self.output_config.update({
            '当前日期': datetime.datetime.now().strftime('%Y年%m月%d日'),
            '当前时间': datetime.datetime.now().strftime('%H时%M分')
        })

        # 按优先级合并
        template_dict.update(self.basic_info)
        template_dict.update(self.company_info)
        template_dict.update(self.case_info)
        template_dict.update(self.investigation)
        template_dict.update(self.output_config)

        return template_dict

    def update_basic_info(self, role: str, data: Dict[str, Any]):
        """更新基础信息"""
        prefixed_data = {}
        for key, value in data.items():
            if not key.startswith(role):
                new_key = f"{role}{key}" if key != "姓名" else f"{role}姓名"
            else:
                new_key = key
            prefixed_data[new_key] = value

        self.basic_info.update(prefixed_data)

    def update_company_info(self, company_name: str = "",
                            employer: str = "",
                            site: str = ""):
        """更新公司信息"""
        if company_name:
            self.company_info['公司名称'] = company_name
        if employer:
            self.company_info['用人单位'] = employer
        if site:
            self.company_info['工地名称'] = site

    def clear_role_data(self, role: str):
        """清除特定角色的数据"""
        role_prefix = role if role in ["本人", "证人", "法人"] else ""
        if not role_prefix:
            return

        keys_to_remove = [
            key for key in self.basic_info.keys()
            if key.startswith(role_prefix)
        ]

        for key in keys_to_remove:
            self.basic_info.pop(key, None)


class MainWindow(QWidget, Ui_Form):

    def __init__(self, username=None, api_config=None, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setupUi(self)
        self._setup_radio_connections()

        self.person_transcript_created = False

        # 保存传递过来的参数（不在此处连接按钮）
        self.username = username
        self.api_config = api_config

        # 第一步：统一设置所有路径（必须在所有服务初始化之前）
        print("=" * 50)
        print("🚀 开始初始化 MainWindow")
        print("=" * 50)

        self._setup_paths()  # 统一使用 path_utils 设置路径

        # 第二步：初始化配置服务（但禁用其路径管理功能）
        self.config_service = ConfigService()
        # 禁用config_service的路径检查，避免干扰
        self.config_service.set('system.startup_check_disk_space', False)
        self.config_service.set('template.base_path', self.TEMPLATE_PATH)

        # 第三步：初始化组合框数据（必须在路径设置之后）
        self.init_combobox_data()

        # 工伤告知书按钮连接
        self.pushButton_7.clicked.connect(self.generate_injury_notice)

        # 第四步：初始化其他核心服务（使用正确的路径）
        self.file_service = FileService(self.BASE_PATH)
        self.data_service = DataService()
        self.template_service = TemplateService(self.TEMPLATE_PATH)

        # 第五步：更新所有服务的路径
        self._update_services_paths()

        # 第六步：初始化数据模型
        self.data_model = CaseDataModel()
        self.var_manager = TemplateVariableManager(self.data_model)
        self.data_model.output_config['用户名'] = username or '未登录用户'

        self.case_index_manager = CaseIndexManager()

        # 简化日志系统
        self.setup_logging()

        # 保持向后兼容
        self.dict = self.data_model.to_template_dict()
        # 使用FileService加载计数器
        self._daily_counter = self.file_service.load_counter()
        self.current_case_folder = None
        self.current_person_name = ""
        # self.case_versions = {}  # 注释掉，如果不再使用

        # 获取复选框控件
        self.death_case_checkbox = self.findChild(QCheckBox, "deathCaseCheckbox")
        self.personal_application_checkbox = self.findChild(QCheckBox, "personalApplicationCheckbox")

        # 连接信号
        self.death_case_checkbox.stateChanged.connect(self.on_case_type_changed)
        self.personal_application_checkbox.stateChanged.connect(self.on_case_type_changed)

        # 初始化数据
        case_config = self.config_service.get_case_config()
        self.set_data('案件性质', case_config.default_case_type, 'case')
        self.set_data('申请类型', case_config.default_application_type, 'case')

        # 连接公司相关信号
        self.company_pane.currentTextChanged.connect(self.company)
        self.construction_company.currentTextChanged.connect(self.sync_employer_to_dict)
        self.constuction_plant.currentTextChanged.connect(self.c_plant)

        # 初始化案件类型下拉框
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第一项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第二项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第三项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第四项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第五项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第六项")

        # 初始化组合框（必须在 init_combobox_data 之后）
        self.init_comboboxes()

        # 设置测试数据
        self.setup_test_data()

        # 应用UI设置
        self._apply_ui_settings()

        # 搜索按钮
        self.pushButton_6.clicked.connect(self.smart_search_cases)

        # 初始化AI服务（使用传入的api_config）
        self.ai_service = None  # 先初始化为None
        self.init_ai_service()

        # 连接AI审查按钮
        self.pushButton_ai_review.clicked.connect(self.ai_review_document)
        self._reconnect_approval_button()
        # 谈话通知书按钮连接
        self.pushButton_12.clicked.connect(self.on_pushButton_12_clicked)

        self.pushButton.clicked.disconnect()
        self.pushButton.clicked.connect(self.on_talk_button_clicked)

        # 初始状态设为可用
        self.pushButton.setEnabled(True)

        # 验证模板路径（使用已获取的路径）
        if os.path.exists(self.TEMPLATE_PATH):
            print(f"✅ 模板路径可访问: {self.TEMPLATE_PATH}")
            print(f"✅ 谈话模板路径: {self.TALK_TEMPLATE_PATH}")
            print(f"✅ 文书模板路径: {self.DOCUMENT_TEMPLATE_PATH}")
        else:
            print(f"❌ 模板路径不存在: {self.TEMPLATE_PATH}")

        print("=" * 50)
        print("🎉 MainWindow 初始化完成")
        print("=" * 50)

    def on_talk_button_clicked(self):
        """谈话笔录按钮点击事件处理"""
        try:
            # 记录已创建本人笔录
            current_role = self.get_current_role_type()
            if current_role == "本人":
                self.person_transcript_created = True

            # 调用原有的谈话笔录生成逻辑
            self.work_tlak()

            # 如果是本人笔录，禁用按钮
            if current_role == "本人":
                self.pushButton.setEnabled(False)
                self.pushButton.setStyleSheet("background-color: #cccccc; color: #666666;")
                print("✅ 本人笔录已创建，谈话笔录按钮已禁用")

        except Exception as e:
            print(f"谈话笔录按钮点击异常: {e}")
            import traceback
            traceback.print_exc()

    def on_role_changed(self):
        """当角色切换时调用"""
        current_role = self.get_current_role_type()
        print(f"🔄 角色切换: {current_role}")

        if current_role == "证人" or current_role == "法人":
            # 切换到证人或法人时，恢复按钮可用状态
            self.pushButton.setEnabled(True)
            self.pushButton.setStyleSheet("")  # 恢复默认样式
            print(f"✅ 切换到{current_role}，谈话笔录按钮已启用")
        else:
            # 切换回本人时，检查是否已创建笔录
            if self.person_transcript_created:
                self.pushButton.setEnabled(False)
                self.pushButton.setStyleSheet("background-color: #cccccc; color: #666666;")
                print("🔒 切换回本人，已创建过笔录，按钮保持禁用")
            else:
                self.pushButton.setEnabled(True)
                self.pushButton.setStyleSheet("")  # 恢复默认样式
                print("✅ 切换回本人，未创建笔录，按钮已启用")

    def _setup_paths(self):
        """统一使用PathUtils设置所有路径"""
        print("=" * 50)
        print("🔄 使用统一的PathUtils设置所有路径...")

        # 使用PathUtils获取路径（path_utils.get_xxx() 方法已确保目录存在）
        self.BASE_PATH = str(path_utils.get_storage_path())
        self.TEMPLATE_PATH = str(path_utils.get_template_path())
        self.CONFIG_PATH = str(path_utils.get_config_path(""))
        self.DATA_PATH = str(path_utils.get_data_path(""))

        # 获取模板子目录（path_utils 已确保目录存在）
        self.TALK_TEMPLATE_PATH = str(path_utils.get_talk_template_path())
        self.DOCUMENT_TEMPLATE_PATH = str(path_utils.get_document_template_path())

    def _update_services_paths(self):
        """更新所有服务的路径"""
        print("🔄 更新服务路径...")

        # 更新FileService
        if hasattr(self, 'file_service'):
            self.file_service.BASE_PATH = self.BASE_PATH
            print(f"✅ 更新FileService路径: {self.BASE_PATH}")

        # 更新TemplateService（使用 path_utils 的路径）
        if hasattr(self, 'template_service'):
            self.template_service.template_path = self.TEMPLATE_PATH
            self.template_service.template_base_path = self.TALK_TEMPLATE_PATH  # 使用已获取的路径
            self.template_service.document_template_path = self.DOCUMENT_TEMPLATE_PATH  # 使用已获取的路径
            print(f"✅ 更新TemplateService路径: {self.TEMPLATE_PATH}")

        # 更新数据模型
        if hasattr(self, 'data_model'):
            self.data_model.company_info['存储路径'] = self.BASE_PATH
            self.data_model.company_info['模板路径'] = self.TEMPLATE_PATH

        print("✅ 服务路径更新完成")

    def select_all_questions(self, select_all: bool):
        """全选或全不选问题"""
        if not hasattr(self, 'question_list_widget'):
            return

        for i in range(self.question_list_widget.count()):
            item = self.question_list_widget.item(i)
            item.setCheckState(Qt.Checked if select_all else Qt.Unchecked)

    def insert_selected_questions(self, dialog):
        """将选中的问题插入到笔录文档"""
        try:
            # 获取选中的问题
            selected_questions = []
            for i in range(self.question_list_widget.count()):
                item = self.question_list_widget.item(i)
                if item.checkState() == Qt.Checked:
                    selected_questions.append(item.text())

            if not selected_questions:
                QMessageBox.warning(dialog, "提示", "请至少选择一个要插入的问题")
                return

            print(f"✅ 选择了 {len(selected_questions)} 个问题准备插入")

            # 检查案件文件夹
            if not self.current_case_folder:
                QMessageBox.warning(dialog, "提示", "请先保存案件信息")
                return

            # 查找本人笔录文件
            person_name = self.get_data("本人姓名", "") or self.lineEdit_2.text().strip()
            if not person_name:
                QMessageBox.warning(dialog, "提示", "请先输入受伤职工姓名")
                return

            # 查找笔录文件
            person_files = []
            for file in os.listdir(self.current_case_folder):
                if "本人" in file and file.endswith('.docx') and "审批表" not in file:
                    person_files.append(file)

            if not person_files:
                QMessageBox.warning(dialog, "提示", "未找到本人笔录文件")
                return

            # 使用第一个找到的本人笔录
            file_path = os.path.join(self.current_case_folder, person_files[0])

            # 插入问题到文档
            success, message = self.insert_questions_to_document(file_path, selected_questions)

            if success:
                QMessageBox.information(dialog, "成功",
                                        f"已成功插入 {len(selected_questions)} 个问题到笔录中\n\n"
                                        f"文件：{os.path.basename(file_path)}")
                dialog.close()
            else:
                QMessageBox.critical(dialog, "失败", f"插入失败：{message}")

        except Exception as e:
            print(f"❌ 插入问题失败: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(dialog, "错误", f"插入过程中发生错误：{str(e)}")

    def insert_questions_to_document(self, file_path: str, questions: list) -> tuple:
        """
        将问题插入到Word文档中

        Args:
            file_path: Word文档路径
            questions: 问题列表

        Returns:
            (是否成功, 消息)
        """
        try:
            from docx import Document
            from docx.shared import Pt, RGBColor

            print(f"📄 正在处理文档: {file_path}")

            # 打开文档
            doc = Document(file_path)

            # 查找插入位置（文档末尾）
            # 我们可以找到最后一个段落，然后在其后插入

            # 添加分隔线
            separator = doc.add_paragraph("=" * 50)
            separator.alignment = 1  # 居中

            # 添加标题
            title = doc.add_paragraph("AI建议补充问题：")
            title.runs[0].bold = True
            title.alignment = 0  # 左对齐

            # 插入每个问题
            for i, question_text in enumerate(questions, 1):
                # 添加问题
                question_para = doc.add_paragraph()
                question_para.add_run(f"{i}. {question_text}")

                # 添加答案占位符（带下划线）
                answer_para = doc.add_paragraph()
                answer_run = answer_para.add_run("答：")
                answer_run.font.underline = True
                answer_run.font.color.rgb = RGBColor(0, 0, 0)

                # 添加空行
                doc.add_paragraph()

            # 保存文档
            doc.save(file_path)

            print(f"✅ 成功插入 {len(questions)} 个问题到文档")

            return True, "插入成功"

        except Exception as e:
            print(f"❌ 插入问题到文档失败: {e}")
            import traceback
            traceback.print_exc()
            return False, str(e)

    def parse_ai_result(self, ai_text: str) -> dict:
        """
        解析AI结果，提取审查结果和缺失问题

        Args:
            ai_text: AI返回的完整文本

        Returns:
            包含审查结果和缺失问题的字典
        """
        result = {
            "审查结果": "",
            "缺失问题": [],
            "原始文本": ai_text
        }

        try:
            # 分割审查结果和缺失问题
            if "【审查结果】" in ai_text and "【缺失问题列表】" in ai_text:
                # 提取审查结果部分
                start = ai_text.find("【审查结果】")
                end = ai_text.find("【缺失问题列表】")

                if start != -1 and end != -1:
                    review_text = ai_text[start:end]
                    # 清理标记
                    review_text = review_text.replace("【审查结果】", "").strip()
                    result["审查结果"] = review_text

                    # 提取缺失问题部分
                    questions_text = ai_text[end:]
                    # 按行分割
                    lines = questions_text.split('\n')

                    for line in lines:
                        line = line.strip()
                        # 查找带方框的问题行
                        if "□" in line and "问：" in line:
                            # 提取问题文本（去掉方框和序号）
                            # 示例：□ 1. 问：您与公司是否签订了书面劳动合同？
                            question = line
                            # 去掉方框标记
                            question = question.replace("□", "", 1).strip()
                            # 去掉序号（如"1. "）
                            if "." in question:
                                question = question.split(".", 1)[1].strip()

                            result["缺失问题"].append(question)

            # 如果格式不正确，尝试其他解析方式
            elif "审查结果" in ai_text and "缺失问题" in ai_text:
                # 尝试其他格式解析
                pass

            else:
                # 如果没有找到格式标记，整个文本作为审查结果
                result["审查结果"] = ai_text

        except Exception as e:
            print(f"解析AI结果失败: {e}")
            result["审查结果"] = ai_text

        print(f"✅ 解析结果: 审查结果长度={len(result['审查结果'])}, 问题数量={len(result['缺失问题'])}")
        return result

    def _setup_paths_from_config(self):
        """从配置服务设置所有路径（简化版）"""
        try:
            print("🔍 开始设置路径（简化版）...")

            # 存储路径使用配置服务（path_utils 已确保目录存在）
            self.BASE_PATH = str(path_utils.get_storage_path())

            # 模板路径直接计算（不依赖配置）
            base_dir = os.path.dirname(os.path.abspath(__file__))
            self.TEMPLATE_PATH = str(path_utils.get_template_path())  # 直接使用 path_utils

            # 更新其他服务
            self._update_services_paths()

            # 禁用 ConfigService 的模板路径验证
            self.config_service.set('system.startup_check_disk_space', False)
            self.config_service.set('template.base_path', self.TEMPLATE_PATH)

        except Exception as e:
            print(f"❌ 配置服务异常，使用默认路径: {e}")
            # 使用更简单的后备路径
            self.BASE_PATH = str(path_utils.get_storage_path())  # 即使异常也使用 path_utils
            self.TEMPLATE_PATH = str(path_utils.get_template_path())
            print(f"✅ 默认路径: {self.BASE_PATH}")

    def _reconnect_approval_button(self):
        """重新连接案件审批表按钮 (pushButton_11)"""
        try:
            # 断开所有现有连接
            self.pushButton_11.clicked.disconnect()
            print("🔌 已断开 pushButton_11 的旧连接")
        except:
            pass  # 如果没有连接，忽略

        # 连接到 approve() 方法
        self.pushButton_11.clicked.connect(self.approve)
        print("✅ 案件审批表按钮 (pushButton_11) 已连接到 approve()")

        # 可选：确认按钮文本
        print(f"📝 按钮文本: '{self.pushButton_11.text()}'")

    def on_pushButton_12_clicked(self):
        """谈话通知书按钮点击事件"""
        print("🔄 谈话通知书按钮被点击")

        # 调用谈话通知书生成函数
        self.generate_interview_notice_from_approval()

    def extract_data_from_approval_table(self, file_path):
        """
        从审批表Word文件中提取数据（简化版）
        """
        try:
            from docx import Document

            print(f"📄 开始提取审批表数据: {file_path}")

            # 搜索关键词映射
            search_mapping = {
                '用人单位': '公司名称',
                '职工姓名': '职工姓名',
                '身份证号': '职工身份证号',
                '申请时间': '申请时间',
                '受理时间': '受理时间',
                '受伤经过': '受伤经过',
                '医疗诊断': '医疗证明',  # 搜索"医疗诊断"
            }

            extracted_data = {}
            document = Document(file_path)

            # 遍历所有表格
            for table in document.tables:
                for row in table.rows:
                    # 遍历每个单元格
                    for i, cell in enumerate(row.cells):
                        cell_text = cell.text.strip()

                        # 检查每个搜索词
                        for search_term, data_field in search_mapping.items():
                            # 如果已经找到了，跳过
                            if data_field in extracted_data:
                                continue

                            # 检查是否包含搜索词
                            if search_term in cell_text:
                                print(f"✅ 找到关键词 '{search_term}'")

                                # 尝试获取右边的单元格
                                if i + 1 < len(row.cells):
                                    right_cell = row.cells[i + 1]
                                    right_text = right_cell.text.strip()

                                    # 只有当右边单元格有内容且不是关键词时才使用
                                    if right_text and right_text != search_term:
                                        extracted_data[data_field] = right_text
                                        print(f"  提取 {data_field}: {right_text}")
                                    else:
                                        # 右边单元格是空的，标记为红色
                                        print(f"  ⚠️ {data_field}: 右边单元格为空")
                                else:
                                    print(f"  ⚠️ {data_field}: 没有右侧单元格")

            print("\n📋 提取结果:")
            required_fields = ['公司名称', '职工姓名', '职工身份证号', '申请时间', '受理时间', '受伤经过', '医疗证明']

            for field in required_fields:
                if field in extracted_data:
                    print(f"  ✅ {field}: {extracted_data[field]}")
                else:
                    print(f"  ❌ {field}: 未找到")

            # 填充缺失字段
            current_date = datetime.datetime.now().strftime('%Y年%m月%d日')

            if '申请时间' not in extracted_data:
                extracted_data['申请时间'] = current_date
                print(f"  🟥 使用默认值 申请时间: {current_date}")

            if '受理时间' not in extracted_data:
                extracted_data['受理时间'] = current_date
                print(f"  🟥 使用默认值 受理时间: {current_date}")

            if '医疗证明' not in extracted_data:
                extracted_data['医疗证明'] = '详见医疗诊断证明'
                print(f"  🟥 使用默认值 医疗证明: 详见医疗诊断证明")

            # 其他字段使用界面数据
            if '公司名称' not in extracted_data:
                extracted_data['公司名称'] = self.company_pane.currentText().strip() or '未知公司'

            if '职工姓名' not in extracted_data:
                extracted_data['职工姓名'] = self.lineEdit_2.text().strip() or '未知'

            if '职工身份证号' not in extracted_data:
                extracted_data['职工身份证号'] = self.get_data('本人身份证号', '')

            return extracted_data

        except Exception as e:
            print(f"❌ 提取审批表数据失败: {e}")
            import traceback
            traceback.print_exc()

            # 返回最小可用数据
            current_date = datetime.datetime.now().strftime('%Y年%m月%d日')
            return {
                '公司名称': self.company_pane.currentText().strip() or '未知公司',
                '职工姓名': self.lineEdit_2.text().strip() or '未知',
                '职工身份证号': self.get_data('本人身份证号', ''),
                '申请时间': current_date,
                '受理时间': current_date,
                '受伤经过': '详见谈话笔录',
                '医疗证明': '详见医疗诊断证明'
            }

    def generate_interview_notice_from_approval(self):
        """从审批表生成接受谈话通知书"""
        try:
            # 1. 检查本人姓名
            person_name = self.get_data("本人姓名", "")
            if not person_name:
                person_name = self.lineEdit_2.text().strip()
                if not person_name:
                    self._set_status('请先输入本人姓名', 'red')
                    return

            # 2. 检查案件文件夹
            if not self.current_case_folder or not os.path.exists(self.current_case_folder):
                self._set_status('请先保存案件信息', 'red')
                return

            # 3. 查找本人案件审批表文件
            approval_file_name = f"{person_name}案件审批表.docx"
            approval_file_path = os.path.join(self.current_case_folder, approval_file_name)

            if not os.path.exists(approval_file_path):
                self._set_status(f'未找到审批表: {approval_file_name}', 'red')
                # 尝试查找其他可能的审批表文件
                all_files = os.listdir(self.current_case_folder)
                approval_files = [f for f in all_files if "审批表" in f and f.endswith('.docx')]

                if not approval_files:
                    self._set_status('案件文件夹中没有审批表文件', 'red')
                    return

                # 使用找到的第一个审批表文件
                approval_file_path = os.path.join(self.current_case_folder, approval_files[0])
                print(f"⚠️ 使用替代审批表: {approval_files[0]}")

            print(f"✅ 找到审批表文件: {approval_file_path}")

            # 4. 从审批表提取数据
            self._set_status('正在提取审批表数据...', 'black')
            QApplication.processEvents()

            extracted_data = self.extract_data_from_approval_table(approval_file_path)

            # 处理申请时间
            if '申请时间' in extracted_data:
                apply_time = extracted_data['申请时间']
                if isinstance(apply_time, str) and apply_time.isdigit() and len(apply_time) == 8:
                    extracted_data['申请时间'] = f"{apply_time[0:4]}年{apply_time[4:6]}月{apply_time[6:8]}日"

            # 处理受理时间
            if '受理时间' in extracted_data:
                accept_time = extracted_data['受理时间']
                if isinstance(accept_time, str) and accept_time.isdigit() and len(accept_time) == 8:
                    extracted_data['受理时间'] = f"{accept_time[0:4]}年{accept_time[4:6]}月{accept_time[6:8]}日"

            if not extracted_data:
                self._set_status('提取审批表数据失败', 'red')
                return

            # 检查必要字段
            required_fields = ['公司名称', '职工姓名', '职工身份证号', '申请时间', '受理时间', '受伤经过', '医疗证明']
            missing_required = []

            for field in required_fields:
                if field not in extracted_data or not extracted_data[field]:
                    missing_required.append(field)

            if missing_required:
                self._set_status(f'审批表缺少必要字段: {missing_required}', 'red')
                return

            self._set_status('审批表数据提取成功', 'green')

            # 5. 准备模板数据
            current_date = datetime.datetime.now().strftime('%Y年%m月%d日')

            # 使用docxtpl的RichText来设置红色
            from docxtpl import RichText

            template_data = {
                '公司名称': extracted_data.get('公司名称', self.company_pane.currentText().strip()),
                '本人姓名': extracted_data.get('职工姓名', self.lineEdit_2.text().strip()),
                '职工性别': extracted_data.get('职工性别', self.get_data('本人性别', '')),
                '本人身份证号': extracted_data.get('职工身份证号', self.get_data('本人身份证号', '')),
                '受伤经过': extracted_data.get('受伤经过', '详见谈话笔录'),
                '当前时期': current_date,
                '当前日期': current_date,
                '当前时间': datetime.datetime.now().strftime('%H时%M分'),
                '案本号': self.get_data('案本号', '')
            }

            # 申请时间：如果有就用，没有就用红色的当前日期
            if '申请时间' in extracted_data and extracted_data['申请时间']:
                template_data['申请时间'] = extracted_data['申请时间']
            else:
                rt = RichText()
                rt.add(current_date, color='FF0000')  # 红色
                template_data['申请时间'] = rt

            # 受理时间：如果有就用，没有就用红色的当前日期
            if '受理时间' in extracted_data and extracted_data['受理时间']:
                template_data['受理时间'] = extracted_data['受理时间']
            else:
                rt = RichText()
                rt.add(current_date, color='FF0000')  # 红色
                template_data['受理时间'] = rt

            # 医疗证明：如果有就用，没有就用红色的"详见医疗诊断证明"
            if '医疗证明' in extracted_data and extracted_data['医疗证明']:
                template_data['医疗证明'] = extracted_data['医疗证明']
            else:
                rt = RichText()
                rt.add('详见医疗诊断证明', color='FF0000')  # 红色
                template_data['医疗证明'] = rt

            # ============ 关键修复：在这里检查模板文件并定义 template_path ============
            # 6. 检查模板文件
            template_path = str(path_utils.get_document_template_path('接受谈话通知书（样本）.docx'))

            # 打印哪些字段用了红色
            red_fields = []
            if '申请时间' not in extracted_data or not extracted_data['申请时间']:
                red_fields.append('申请时间')
            if '受理时间' not in extracted_data or not extracted_data['受理时间']:
                red_fields.append('受理时间')
            if '医疗证明' not in extracted_data or not extracted_data['医疗证明']:
                red_fields.append('医疗证明')

            if red_fields:
                print(f"🔴 以下字段使用红色: {red_fields}")

            # 6. 生成文档
            self._set_status('正在生成谈话通知书...', 'black')
            QApplication.processEvents()

            try:
                from docxtpl import DocxTemplate
                word = DocxTemplate(template_path)
                word.render(template_data)

                notice_file_name = f"{person_name}接受谈话通知书.docx"
                target_path = os.path.join(self.current_case_folder, notice_file_name)
                word.save(target_path)

                print(f"✅ 谈话通知书保存到: {target_path}")

                # 打开文件
                success, message = self.file_service.open_document(target_path)
                if success:
                    self._set_status('谈话通知书生成成功', 'green')
                else:
                    self._set_status(f'谈话通知书生成成功，但打开失败: {message}', 'orange')

            except Exception as e:
                print(f"❌ 生成谈话通知书失败: {e}")
                import traceback
                traceback.print_exc()
                self._set_status(f'生成谈话通知书失败: {str(e)}', 'red')

        except Exception as e:
            print(f"❌ 谈话通知书过程异常: {e}")
            import traceback
            traceback.print_exc()
            self._set_status(f'生成谈话通知书异常: {str(e)}', 'red')

    def _cleanup_resources(self):
        """统一清理所有临时资源（对话框、线程等）"""
        resources = ['wait_dialog', 'progress_dialog', 'ai_worker']

        for attr_name in resources:
            if hasattr(self, attr_name):
                try:
                    resource = getattr(self, attr_name)
                    if attr_name == 'ai_worker' and resource.isRunning():
                        resource.terminate()  # 改为terminate
                        resource.wait(5000)  # 等待5秒
                    elif hasattr(resource, 'close'):
                        resource.close()
                except Exception as e:
                    print(f"清理资源 {attr_name} 失败: {e}")

    def _clean_ai_output(self, text):
        """清理AI输出中的重复标点（保留正常句号）"""
        import re

        if not text:
            return text

        # 1. 替换连续的句号（2个或更多）为单个句号
        text = re.sub(r'。{2,}', '。', text)

        # 2. 替换连续的中文逗号为单个逗号
        text = re.sub(r'，{2,}', '，', text)

        # 3. 清理"句号+空格+句号"的情况
        text = re.sub(r'。\s*。', '。', text)

        return text

    def _set_status(self, text, color="black", label="label_14"):
        """设置状态标签"""
        label_widget = getattr(self, label, None)
        if not label_widget:
            return

        label_widget.setText(text)
        if color == "green":
            label_widget.setStyleSheet("QLabel{color:green;}")
        elif color == "red":
            label_widget.setStyleSheet("QLabel{color:red;}")
        elif color == "orange":
            label_widget.setStyleSheet("QLabel{color:orange;}")
        else:
            label_widget.setStyleSheet("QLabel{color:black;}")

    def _handle_ai_result(self, result=None, error=None, canceled=False):
        """
        统一处理AI操作结果
        """
        print(f"🔄 处理AI结果: result={result is not None}, error={error}, canceled={canceled}")

        # 添加简单的防重复
        if hasattr(self, '_is_handling_ai_result') and self._is_handling_ai_result:
            print("⚠️ 已经在处理AI结果，跳过重复调用")
            return

        self._is_handling_ai_result = True

        try:
            # 清理资源
            self._cleanup_resources()

            if canceled:
                print("⏹️ AI操作被用户取消")
                return

            if error:
                print(f"❌ AI操作出错: {error}")
                QMessageBox.critical(self, "AI审查错误", f"操作失败: {error}")
                return

            if result:
                print("✅ AI操作成功，显示结果")
                self.show_ai_review_result(result)
        except Exception as e:
            print(f"❌ 处理AI结果时出错: {e}")
            import traceback
            traceback.print_exc()
        finally:
            # 清理标志
            if hasattr(self, '_is_handling_ai_result'):
                delattr(self, '_is_handling_ai_result')

    def _setup_radio_connections(self):
        """设置单选按钮信号连接（简化版）"""
        # 断开现有连接（如果有）
        try:
            self.radioButton.clicked.disconnect()
            self.radioButton_2.clicked.disconnect()
            self.radioButton_3.clicked.disconnect()
        except:
            pass

        # 重新连接
        self.radioButton.clicked.connect(self.clear_role_fields)
        self.radioButton_2.clicked.connect(self.clear_role_fields)
        self.radioButton_3.clicked.connect(self.clear_role_fields)
        print("✅ 单选按钮信号重新连接")

    def init_ai_service(self):
        """初始化AI服务 - 使用传入的API配置（简化版）"""
        try:
            # 直接检查传入的api_config
            if not self.api_config:
                print("⚠️ 未传入API配置，AI功能将不可用")
                self.ai_service = None
                return

            api_key = self.api_config.get('api_key', '')
            api_url = self.api_config.get('api_url', 'https://api.deepseek.com')

            # 检查配置是否完整
            if not api_key or not api_url:
                print("⚠️ API配置不完整，AI功能将不可用")
                print(f"  API地址: {api_url if api_url else '未设置'}")
                print(f"  API密钥: {'已设置' if api_key else '未设置'}")
                self.ai_service = None
                return

            print(f"✅ 使用传入的API配置初始化AI服务")
            print(f"  API地址: {api_url}")
            print(f"  API密钥前8位: {api_key[:8]}...")

            # 创建AI服务实例
            self.ai_service = AIService(api_key, api_url)
            print("✅ AI服务初始化成功")

        except Exception as e:
            print(f"❌ AI服务初始化失败: {e}")
            import traceback
            traceback.print_exc()
            self.ai_service = None

    def ai_review_document(self):
        """AI审查文档"""
        try:
            print("=" * 50)
            print("🔄 AI审查开始执行...")

            # 检查AI服务
            if not self.ai_service:
                print("❌ AI服务未初始化")
                QMessageBox.warning(self, "AI审查", "请先配置API密钥。")
                self.init_ai_service()
                return

            print("✅ AI服务检查通过")

            # 检查案件文件夹
            if not self.current_case_folder:
                print("❌ 当前案件文件夹为空")
                QMessageBox.warning(self, "AI审查", "请先保存案件信息。")
                return

            print(f"📁 案件文件夹: {self.current_case_folder}")

            # 查找本人笔录文件
            person_files = []
            for file in os.listdir(self.current_case_folder):
                print(f"📄 检查文件: {file}")
                if "本人" in file and file.endswith('.docx') and "审批表" not in file:
                    person_files.append(file)
                    print(f"✅ 找到本人笔录: {file}")

            if not person_files:
                print("❌ 未找到本人笔录文件")
                QMessageBox.warning(self, "AI审查", "未找到本人笔录文件。")
                return

            print(f"✅ 找到{len(person_files)}个本人笔录文件")

            # 使用第一个找到的本人笔录
            file_path = os.path.join(self.current_case_folder, person_files[0])
            print(f"📄 使用文件路径: {file_path}")

            # 检查文件是否存在
            if not os.path.exists(file_path):
                print(f"❌ 文件不存在")
                QMessageBox.warning(self, "AI审查", f"文件不存在: {file_path}")
                return

            print("✅ 文件存在")

            # ======================
            # 在这里创建进度对话框和AI工作线程
            # ======================

            # 创建进度对话框
            self.progress_dialog = QProgressDialog("正在分析文档...", "取消", 0, 100, self)
            self.progress_dialog.setWindowTitle("AI审查")
            self.progress_dialog.setWindowModality(Qt.WindowModal)

            # 创建AI工作线程（现在file_path已经定义）
            self.ai_worker = AIWorker(self.ai_service, file_path)

            # 连接信号到统一处理方法
            self.ai_worker.finished.connect(
                lambda result: self._handle_ai_result(result=result)
            )
            self.ai_worker.error.connect(
                lambda error: self._handle_ai_result(error=error)
            )
            self.ai_worker.progress.connect(
                lambda msg, value: self.progress_dialog.setLabelText(f"{msg} ({value}%)")
            )

            # 进度对话框取消
            self.progress_dialog.canceled.connect(
                lambda: self._handle_ai_result(canceled=True)
            )

            # 显示进度对话框并启动线程
            self.progress_dialog.show()
            self.ai_worker.start()

        except Exception as e:
            print(f"🔥 整体审查过程异常: {str(e)}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "AI审查错误", f"审查失败: {str(e)}")

    def show_ai_review_result(self, review_result):
        """显示AI审查结果 - 带问题选择功能"""
        print(f"🖥️ 显示审查结果，结果类型: {type(review_result)}")

        # 处理不同的结果格式
        if isinstance(review_result, str):
            result_text = review_result
        elif isinstance(review_result, dict):
            if "结果" in review_result:
                result_text = review_result["结果"]
            elif "错误信息" in review_result:
                result_text = f"错误: {review_result['错误信息']}"
            elif "原始回复" in review_result:
                result_text = review_result["原始回复"]
            else:
                result_text = str(review_result)
        else:
            result_text = str(review_result)

        print(f"📝 要解析的文本长度: {len(result_text)}")

        # 解析AI结果
        parsed_result = self.parse_ai_result(result_text)

        # 创建对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("AI法律审查结果")
        dialog.resize(700, 600)

        layout = QVBoxLayout()

        # 标题
        title = QLabel("AI法律审查报告")
        title.setStyleSheet("font-size: 16px; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # 创建标签页
        tab_widget = QTabWidget()

        # 标签1：审查结果
        review_tab = QWidget()
        review_layout = QVBoxLayout()

        review_label = QLabel("审查结果分析：")
        review_label.setStyleSheet("font-weight: bold;")
        review_layout.addWidget(review_label)

        # 审查结果显示区域
        review_text_edit = QTextEdit()
        review_text_edit.setReadOnly(True)
        review_text_edit.setPlainText(parsed_result["审查结果"])
        review_text_edit.setMinimumHeight(300)
        review_layout.addWidget(review_text_edit)

        review_tab.setLayout(review_layout)
        tab_widget.addTab(review_tab, "审查结果")

        # 标签2：缺失问题（如果有）
        if parsed_result["缺失问题"]:
            questions_tab = QWidget()
            questions_layout = QVBoxLayout()

            questions_label = QLabel(f"发现 {len(parsed_result['缺失问题'])} 个缺失问题，请勾选需要添加到笔录的问题：")
            questions_label.setStyleSheet("font-weight: bold; color: #e74c3c;")
            questions_layout.addWidget(questions_label)

            # 创建问题列表（带复选框）
            self.question_list_widget = QListWidget()

            for question in parsed_result["缺失问题"]:
                item = QListWidgetItem(question)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)  # 默认未选中
                self.question_list_widget.addItem(item)

            questions_layout.addWidget(self.question_list_widget)

            # 全选/全不选按钮
            select_buttons_layout = QHBoxLayout()

            btn_select_all = QPushButton("全选")
            btn_select_all.clicked.connect(lambda: self.select_all_questions(True))

            btn_select_none = QPushButton("全不选")
            btn_select_none.clicked.connect(lambda: self.select_all_questions(False))

            select_buttons_layout.addWidget(btn_select_all)
            select_buttons_layout.addWidget(btn_select_none)
            select_buttons_layout.addStretch()

            questions_layout.addLayout(select_buttons_layout)

            questions_tab.setLayout(questions_layout)
            tab_widget.addTab(questions_tab, f"缺失问题 ({len(parsed_result['缺失问题'])})")

        layout.addWidget(tab_widget)

        # 按钮区域
        button_layout = QHBoxLayout()

        # 插入到笔录按钮（只在有问题时显示）
        if parsed_result.get("缺失问题"):
            btn_insert = QPushButton("插入选中问题到笔录")
            btn_insert.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold;")
            btn_insert.clicked.connect(lambda: self.insert_selected_questions(dialog))
            button_layout.addWidget(btn_insert)

        btn_copy = QPushButton("复制结果")
        btn_copy.clicked.connect(lambda: self.copy_to_clipboard(parsed_result["审查结果"]))

        btn_save = QPushButton("保存报告")
        btn_save.clicked.connect(lambda: self.save_ai_report(parsed_result))

        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(dialog.close)

        button_layout.addWidget(btn_copy)
        button_layout.addWidget(btn_save)
        button_layout.addWidget(btn_close)

        layout.addLayout(button_layout)

        dialog.setLayout(layout)

        # 保存解析结果到对话框对象，以便后续使用
        dialog.parsed_result = parsed_result

        dialog.exec_()

    def copy_to_clipboard(self, text, parent=None):
        """复制文本到剪贴板"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        QMessageBox.information(parent or self, "复制成功", "审查结果已复制到剪贴板")

    def save_ai_report(self, parsed_result):
        """保存AI审查报告"""
        if not self.current_case_folder:
            QMessageBox.warning(self, "提示", "请先保存案件信息")
            return

        person_name = self.get_data('本人姓名', '未知')
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{person_name}_AI审查报告_{timestamp}.txt"
        filepath = os.path.join(self.current_case_folder, filename)

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("AI法律审查报告\n")
                f.write("=" * 60 + "\n\n")
                f.write(f"审查时间：{datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n")
                f.write(f"审查对象：{person_name}\n\n")

                f.write("【审查结果】\n")
                f.write(parsed_result.get("审查结果", "无审查结果") + "\n\n")

                if parsed_result.get("缺失问题"):
                    f.write("【缺失问题列表】\n")
                    for i, question in enumerate(parsed_result["缺失问题"], 1):
                        f.write(f"{i}. {question}\n")

            QMessageBox.information(self, "保存成功", f"报告已保存为：\n{filename}")
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"保存失败：{str(e)}")

    def generate_injury_notice(self):
        """生成工伤告知书"""
        try:
            print("🔄 开始生成工伤告知书...")

            # 1. 获取本人姓名
            person_name = self.get_data("本人姓名", "")
            if not person_name:
                person_name = self.lineEdit_2.text().strip()
                if not person_name:
                    self._set_status('请先输入本人姓名', 'red')
                    QMessageBox.warning(self, "提示", "请先输入受伤职工姓名")
                    return

            # 2. 检查案件文件夹
            if not self.current_case_folder or not os.path.exists(self.current_case_folder):
                self._set_status('请先保存案件信息', 'red')
                QMessageBox.warning(self, "提示", "请先保存案件信息")
                return

            # 3. 查找本人案件审批表
            approval_file_name = f"{person_name}案件审批表.docx"
            approval_file_path = os.path.join(self.current_case_folder, approval_file_name)

            if not os.path.exists(approval_file_path):
                # 尝试查找其他可能的审批表文件
                all_files = os.listdir(self.current_case_folder)
                approval_files = [f for f in all_files if "审批表" in f and f.endswith('.docx')]

                if not approval_files:
                    self._set_status('请先生成案件审批表', 'red')
                    QMessageBox.warning(self, "提示", "请先生成案件审批表")
                    return

                approval_file_path = os.path.join(self.current_case_folder, approval_files[0])
                print(f"⚠️ 使用替代审批表: {approval_files[0]}")

            print(f"✅ 找到审批表文件: {approval_file_path}")

            # 4. 从审批表提取数据（复用现有方法）
            self._set_status('正在提取审批表数据...', 'black')
            QApplication.processEvents()

            extracted_data = self.extract_data_from_approval_table(approval_file_path)

            if not extracted_data:
                self._set_status('提取审批表数据失败', 'red')
                return

            # 5. 准备模板数据
            current_date = datetime.datetime.now().strftime('%Y年%m月%d日')
            current_time = datetime.datetime.now().strftime('%H时%M分')

            # 获取公司信息
            company_name = extracted_data.get('公司名称', self.company_pane.currentText().strip())
            if not company_name:
                company_name = self.get_data('公司名称', '未知公司')

            # 获取身份证号
            id_number = extracted_data.get('职工身份证号', self.get_data('本人身份证号', ''))

            # 获取申请时间（如果有的话）
            application_date = extracted_data.get('申请时间', current_date)

            # 准备模板数据
            template_data = {
                '公司名称': company_name,
                '本人姓名': extracted_data.get('职工姓名', person_name),
                '职工性别': extracted_data.get('职工性别', self.get_data('本人性别', '')),
                '本人身份证号': id_number,
                '受伤经过': extracted_data.get('受伤经过', '详见谈话笔录'),
                '当前时期': current_date,
                '当前日期': current_date,
                '当前时间': current_time,
                '申请时间': application_date,
                '案本号': self.get_data('案本号', ''),
                '告知日期': current_date,
                '受理编号': self.generate_case_number()  # 可以生成一个受理编号
            }

            # 6. 检查模板文件
            template_path = str(path_utils.get_document_template_path('工伤告知书（样本）.docx'))

            # 7. 生成文档
            self._set_status('正在生成工伤告知书...', 'black')
            QApplication.processEvents()

            try:
                from docxtpl import DocxTemplate
                word = DocxTemplate(template_path)
                word.render(template_data)

                notice_file_name = f"{person_name}工伤告知书.docx"
                target_path = os.path.join(self.current_case_folder, notice_file_name)
                word.save(target_path)

                print(f"✅ 工伤告知书保存到: {target_path}")

                # 打开文件
                success, message = self.file_service.open_document(target_path)
                if success:
                    self._set_status('工伤告知书生成成功', 'green')
                    QMessageBox.information(self, "成功", f"工伤告知书已生成:\n{notice_file_name}")
                else:
                    self._set_status(f'工伤告知书生成成功，但打开失败: {message}', 'orange')
                    QMessageBox.information(self, "成功",
                                            f"工伤告知书已生成:\n{notice_file_name}\n\n但打开失败: {message}")

            except Exception as e:
                print(f"❌ 生成工伤告知书失败: {e}")
                import traceback
                traceback.print_exc()
                self._set_status(f'生成工伤告知书失败: {str(e)}', 'red')
                QMessageBox.critical(self, "错误", f"生成工伤告知书失败:\n{str(e)}")

        except Exception as e:
            print(f"❌ 工伤告知书过程异常: {e}")
            import traceback
            traceback.print_exc()
            self._set_status(f'生成工伤告知书异常: {str(e)}', 'red')
            QMessageBox.critical(self, "错误", f"生成工伤告知书异常:\n{str(e)}")

    def setup_test_data(self):
        """设置测试数据"""
        self.name_pane.setText("张三")
        self.lineEdit.setText("男")
        self.age_pane.setText("35")
        self.idnumer_pane.setText("110101199001011234")
        self.textEdit.setPlainText("北京市海淀区中关村大街1号")
        self.lineEdit_4.setText("13800138000")
        self.lineEdit_5.setText("建筑工人")

    def _apply_ui_settings(self):
        """应用UI设置"""
        try:
            ui_settings = self.config_service.get_ui_settings()

            # 设置字体
            font = QFont(ui_settings.font_family, ui_settings.font_size)
            self.setFont(font)

            # 设置窗口大小
            self.resize(ui_settings.window_width, ui_settings.window_height)

            # 应用语言设置（如果有国际化支持）
            if hasattr(self, 'retranslateUi'):
                # 重新翻译UI
                pass
        except Exception as e:
            print(f"⚠️ 应用UI设置失败: {e}")

    def clear_role_fields(self):
        """
        清空当前角色的字段 - 注意：这个方法的参数签名需要匹配信号
        """
        # 获取当前角色
        sender = self.sender()
        if sender == self.radioButton:
            role = "本人"
        elif sender == self.radioButton_2:
            role = "证人"
        elif sender == self.radioButton_3:
            role = "法人"
        else:
            role = "本人"

        print(f"🧹 清空{role}字段")

        # 清空输入控件
        self.clear_fields()

        # 当角色切换时，更新按钮状态
        self.on_role_changed()

        if role == "本人":
            self.lineEdit_2.clear()  # 清空受伤职工显示

    def setup_logging(self):
        """简化日志系统"""
        self.log_warning = lambda msg: print(f"警告: {msg}")
        self.log_error = lambda msg: print(f"错误: {msg}")

    def show_validation_errors(self, errors: List[str]) -> bool:
        """显示验证错误"""
        if errors:
            error_msg = "数据验证失败：\n" + "\n".join(errors)
            display_msg = error_msg[:100] + "..." if len(error_msg) > 100 else error_msg
            self._set_status(display_msg, 'red')
            return False
        return True

    def on_case_type_changed(self):
        """当案件类型选择改变时调用"""
        is_death_case = self.death_case_checkbox.isChecked()
        is_personal = self.personal_application_checkbox.isChecked()

        self.set_data('案件性质', "工亡案件" if is_death_case else "工伤案件", 'case')
        self.set_data('申请类型', "个人申请" if is_personal else "单位申请", 'case')
        self.update_case_type_hint()

    def update_case_type_hint(self):
        """更新案件类型提示信息"""
        is_death_case = self.death_case_checkbox.isChecked()
        is_personal = self.personal_application_checkbox.isChecked()

        hint_text = "当前案件类型: "
        if is_death_case and is_personal:
            hint_text += "个人申请的工亡案件"
        elif is_death_case:
            hint_text += "单位申请的工亡案件"
        elif is_personal:
            hint_text += "个人申请的工伤案件"
        else:
            hint_text += "单位申请的工伤案件"

        hint_label = self.findChild(QLabel, "labelCaseHint")
        if hint_label:
            hint_label.setText(hint_text)

    def init_combobox_data(self):
        """初始化组合框数据（简化版）"""
        try:
            # 使用 path_utils 的数据路径
            from path_utils import path_utils

            data_dir = path_utils.get_data_path("")
            print(f"🔍 数据目录: {data_dir}")

            # 公司名称 - 使用数据目录
            company_file = str(data_dir / '公司名称汇总.xlsx')
            print(f"🔍 公司名称文件: {company_file}")

            if os.path.exists(company_file):
                try:
                    file = pd.read_excel(company_file)
                    self.items_list = file['公司名称汇总'].tolist()
                    print(f"✅ 加载公司名称: {len(self.items_list)}个")
                    if self.items_list:
                        print(f"   示例: {self.items_list[:3]}")
                except Exception as e:
                    print(f"❌ 读取公司名称文件失败: {e}")
                    self.items_list = ['公司A', '公司B', '公司C']  # 默认数据
            else:
                print("⚠️ 公司名称文件不存在，创建默认文件")
                self.items_list = ['公司A', '公司B', '公司C']
                # 创建默认文件
                try:
                    df = pd.DataFrame(self.items_list, columns=['公司名称汇总'])
                    df.to_excel(company_file, index=False)
                    print(f"✅ 创建默认公司名称文件")
                except Exception as e:
                    print(f"❌ 创建公司名称文件失败: {e}")

            # 用人单位 - 使用数据目录
            employer_file = str(data_dir / '用人单位汇总.xlsx')
            print(f"🔍 用人单位文件: {employer_file}")

            if os.path.exists(employer_file):
                try:
                    file1 = pd.read_excel(employer_file)
                    self.items_list1 = file1['用人单位汇总'].tolist()
                    print(f"✅ 加载用人单位: {len(self.items_list1)}个")
                except Exception as e:
                    print(f"❌ 读取用人单位文件失败: {e}")
                    self.items_list1 = ['用人单位A', '用人单位B']
            else:
                print("⚠️ 用人单位文件不存在，创建默认文件")
                self.items_list1 = ['用人单位A', '用人单位B']
                try:
                    df = pd.DataFrame(self.items_list1, columns=['用人单位汇总'])
                    df.to_excel(employer_file, index=False)
                    print(f"✅ 创建默认用人单位文件")
                except Exception as e:
                    print(f"❌ 创建用人单位文件失败: {e}")

            # 工地名称 - 使用数据目录
            site_file = str(data_dir / '工地名称汇总.xlsx')
            print(f"🔍 工地名称文件: {site_file}")

            if os.path.exists(site_file):
                try:
                    file2 = pd.read_excel(site_file)
                    self.items_list2 = file2['工地名称汇总'].tolist()
                    print(f"✅ 加载工地名称: {len(self.items_list2)}个")
                except Exception as e:
                    print(f"❌ 读取工地名称文件失败: {e}")
                    self.items_list2 = ['工地A', '工地B']
            else:
                print("⚠️ 工地名称文件不存在，创建默认文件")
                self.items_list2 = ['工地A', '工地B']
                try:
                    df = pd.DataFrame(self.items_list2, columns=['工地名称汇总'])
                    df.to_excel(site_file, index=False)
                    print(f"✅ 创建默认工地名称文件")
                except Exception as e:
                    print(f"❌ 创建工地名称文件失败: {e}")

        except Exception as e:
            print(f"❌ 初始化组合框数据失败: {e}")
            import traceback
            traceback.print_exc()

            # 设置默认数据
            self.items_list = ['公司A', '公司B', '公司C']
            self.items_list1 = ['用人单位A', '用人单位B']
            self.items_list2 = ['工地A', '工地B']
            print("✅ 使用默认数据")

    def on_id_input_finished(self):
        """当身份证输入框完成编辑时自动计算年龄和性别"""
        try:
            role = self.get_current_role_type()
            idcard = self.idnumer_pane.text().strip()

            if not idcard or len(idcard) not in (15, 18):
                return

            self.set_data(f"{role}身份证号", idcard, 'basic')
            self.process_id_info(role)
            self.calculate_age_from_id(role)

        except Exception as e:
            import traceback
            traceback.print_exc()

    def save_company(self):
        """保存公司名称到Excel"""
        new_item = self.company_pane.currentText().strip()
        if new_item and new_item not in self.items_list:
            # 使用FileService保存（不需要传递路径，file_service内部使用path_utils）
            self.items_list = self.file_service.save_to_excel(
                "",  # 第一个参数可以是空字符串
                '公司名称汇总.xlsx',
                '公司名称汇总',
                new_item,
                self.items_list
            )
            # 更新下拉框
            self.init_combobox(self.company_pane, self.items_list)
            print(f"💾 保存公司名称: {new_item}")

    def save_construction_company(self):
        """保存用人单位到Excel"""
        new_item = self.construction_company.currentText().strip()
        if new_item and new_item not in self.items_list1:
            self.items_list1 = self.file_service.save_to_excel(
                "",  # 空字符串
                '用人单位汇总.xlsx',
                '用人单位汇总',
                new_item,
                self.items_list1
            )
            self.init_combobox(self.construction_company, self.items_list1)
            print(f"💾 保存用人单位: {new_item}")

    def save_construction_plant(self):
        """保存工地名称到Excel"""
        new_item = self.constuction_plant.currentText().strip()
        if new_item and new_item not in self.items_list2:
            self.items_list2 = self.file_service.save_to_excel(
                "",  # 空字符串
                '工地名称汇总.xlsx',
                '工地名称汇总',
                new_item,
                self.items_list2
            )
            self.init_combobox(self.constuction_plant, self.items_list2)
            print(f"💾 保存工地名称: {new_item}")

    def init_combobox(self, combobox, items):
        """初始化组合框"""
        print(f"🔍 初始化 {combobox.objectName()}，数据长度: {len(items)}")

        combobox.clear()

        if items:
            for item in items:
                combobox.addItem(str(item))
            print(f"✅ 添加了 {len(items)} 个选项")
        else:
            print("⚠️ 没有数据可添加")
            combobox.addItem("暂无数据")

        combobox.setCurrentIndex(-1)  # 清空选择

        # 设置自动完成
        completer = QCompleter(items)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        combobox.setCompleter(completer)

        print(f"✅ {combobox.objectName()} 初始化完成，当前项数: {combobox.count()}")

    def calculate_age_from_id(self, role):
        """根据身份证号计算年龄（使用DataService）"""
        try:
            idcard = self.get_data(f"{role}身份证号", "")
            if not idcard:
                return

            # 使用DataService计算年龄
            age = self.data_service.calculate_age_from_idcard(idcard)
            if age is None:
                self._set_status('身份证号格式错误', 'red', 'label_12')
                return

            # 设置年龄数据
            self.set_data(f"{role}年龄", age, 'basic')
            self.age_pane.setText(str(age))

            # 超龄检查（仅对本人）
            if role == "本人":
                gender = self.get_data(f"{role}性别", "")
                if (gender == "男" and age > 60) or (gender == "女" and age > 50):
                    self._set_status('此人已经超龄', 'red', 'label_12')
                else:
                    self._set_status('信息提示', 'black', 'label_12')

        except Exception as e:
            print(f"[calculate_age_from_id] 错误: {e}")
            self._set_status('年龄计算错误', 'red', 'label_12')

    def get_data(self, key: str, default: Any = None) -> Any:
        """统一的数据访问方法"""
        model_value = self._get_from_data_model(key)
        if model_value is not None:
            if key not in self.dict or self.dict[key] != model_value:
                self.dict[key] = model_value
            return model_value

        dict_value = self.dict.get(key)
        if dict_value is not None:
            return dict_value

        return default

    def set_data(self, key: str, value: Any, category: str = "auto") -> None:
        """统一的数据设置方法"""
        if value is None:
            value = ""

        if category == "auto":
            category = self._detect_data_category(key)

        self._store_to_data_model(key, value, category)
        self.dict[key] = value
        self._handle_special_keys(key, value)

    def _handle_special_keys(self, key: str, value: Any):
        """处理特殊键的同步逻辑"""
        if key == "本人姓名":
            self.lineEdit_2.setText(str(value))
        elif key == "当前时期":
            if not value:
                current_date = datetime.datetime.now().strftime('%Y年%m月%d日')
                current_time = datetime.datetime.now().strftime('%H时%M分')
                self.set_data(key, f"{current_date}{current_time}", 'output')

    def _detect_data_category(self, key: str) -> str:
        """自动检测数据类别"""
        role_prefixes = ['本人', '证人', '法人']
        for prefix in role_prefixes:
            if key.startswith(prefix):
                base_key = key[len(prefix):]
                return self._detect_base_category(base_key)
        return self._detect_base_category(key)

    def _detect_base_category(self, key: str) -> str:
        """检测基础键名的类别"""
        if key in ['姓名', '年龄', '性别', '身份证号', '身份证地址', '手机号', '岗位', '职务']:
            return 'basic'
        elif key in ['公司名称', '用人单位', '工地名称']:
            return 'company'
        elif key in ['案件性质', '申请类型', '案本号', '案件类型']:
            return 'case'
        elif key in ['当前日期', '当前时间', '当前时期']:
            return 'output'
        else:
            return 'investigation'

    def _get_from_data_model(self, key: str) -> Any:
        """从数据模型的各个部分获取数据"""
        if key in self.data_model.basic_info:
            return self.data_model.basic_info[key]

        role_prefixes = ['本人', '证人', '法人']
        for prefix in role_prefixes:
            if key.startswith(prefix):
                base_key = key[len(prefix):]
                if base_key in self.data_model.basic_info:
                    return self.data_model.basic_info[base_key]

        data_categories = [
            self.data_model.company_info,
            self.data_model.case_info,
            self.data_model.investigation,
            self.data_model.output_config,
        ]

        for category in data_categories:
            if key in category:
                return category[key]

        return None

    def _store_to_data_model(self, key: str, value: Any, category: str) -> None:
        """存储数据到数据模型"""
        role_prefixes = ['本人', '证人', '法人']
        for prefix in role_prefixes:
            if key.startswith(prefix):
                self.data_model.basic_info[key] = value
                return

        category_map = {
            'basic': self.data_model.basic_info,
            'company': self.data_model.company_info,
            'case': self.data_model.case_info,
            'investigation': self.data_model.investigation,
            'output': self.data_model.output_config,
        }

        if category in category_map:
            category_map[category][key] = value
        else:
            self.data_model.basic_info[key] = value

    def clear_fields(self):
        """清空所有输入控件"""
        fields = [
            self.name_pane, self.age_pane, self.lineEdit,
            self.idnumer_pane, self.textEdit, self.lineEdit_4, self.lineEdit_5
        ]
        for field in fields:
            if isinstance(field, QLineEdit):
                field.clear()
            elif isinstance(field, QTextEdit):
                field.clear()

        self._set_status('信息提示', 'black', 'label_14')
        self._set_status('信息提示', 'black', 'label_12')

    def approve(self):
        """打开案件审批表并生成新的审批表（AI优化版）"""
        try:
            # 1. 获取本人姓名
            person_name = self.data_model.basic_info.get('本人姓名', '')
            if not person_name:
                person_name = self.lineEdit_2.text().strip()
                if person_name:
                    self.data_model.basic_info['本人姓名'] = person_name
                else:
                    self._set_status('请先输入本人姓名', 'red')
                    return

            regulation_text = self.comboBox.currentText().strip()
            # ✅ 新增：获取性别（只看lineEdit输入框）
            gender = self.lineEdit.text().strip()  # 直接读取性别输入框
            if gender not in ["男", "女"]:  # 如果不是有效性别
                gender = "待补充"  # 直接设为待补充

            # 2. 检查案件文件夹
            if not self.current_case_folder or not os.path.exists(self.current_case_folder):
                self._set_status('请先保存案件信息', 'red')
                return

            # 3. 查找本人笔录文件
            person_files = [f for f in os.listdir(self.current_case_folder)
                            if "本人" in f and f.endswith('.docx')]

            if not person_files:
                self._set_status('没找到本人笔录，请先保存本人笔录', 'red')
                return

            # 4. 显示进度提示
            self._set_status('正在提取笔录信息...', 'black')
            QApplication.processEvents()

            # 5. 读取笔录文件并提取关键信息
            old_file_path = os.path.join(self.current_case_folder, person_files[0])
            document = Document(old_file_path)

            thispcjh = []  # 存储受伤经过
            basic_info = {}  # 存储基本信息
            medical_info = ""  # 存储送医情况
            injury_location = ""  # 存储受伤部位描述（从医疗结论提取）

            for i, paragraph in enumerate(document.paragraphs):
                text = paragraph.text

                # 提取自我介绍
                if "请介绍一下你的姓名" in text and i + 1 < len(document.paragraphs):
                    next_para = document.paragraphs[i + 1].text
                    if next_para.startswith("答："):
                        basic_info['自我介绍'] = next_para[2:].strip()

                # 提取受伤经过
                elif "详细的描述" in text and i + 1 < len(document.paragraphs):
                    next_para = document.paragraphs[i + 1].text
                    if next_para.startswith("答："):
                        raw_text = next_para[2:].replace('我们', '他们').replace('我', person_name)
                        thispcjh.append(raw_text)

                # 提取送医情况
                elif "你受伤后去的哪个医院" in text or "是谁送你去救治的" in text:
                    if i + 1 < len(document.paragraphs):
                        next_para = document.paragraphs[i + 1].text
                        if next_para.startswith("答："):
                            medical_info = next_para[2:].strip()
                            print(f"✅ 找到送医情况: {medical_info}")

                # 提取医疗结论 - 独立存储，不拼接到描述中
                elif "医疗结论" in text or "诊断结论" in text or "医院诊断" in text:
                    if i + 1 < len(document.paragraphs):
                        next_para = document.paragraphs[i + 1].text
                        if next_para.startswith("答："):
                            raw_diagnosis = next_para[2:].strip()
                            print(f"✅ 找到原始医疗结论: {raw_diagnosis}")

                            # 简单处理：提取受伤部位描述
                            injury_location = raw_diagnosis

                            # 1. 移除常见的开头短语
                            if injury_location.startswith("医院对我的结论是："):
                                injury_location = injury_location[8:]  # 移除"医院对我的结论是："
                            elif injury_location.startswith("医院结论是："):
                                injury_location = injury_location[6:]  # 移除"医院结论是："
                            elif injury_location.startswith("医院诊断是："):
                                injury_location = injury_location[6:]  # 移除"医院诊断是："
                            elif injury_location.startswith("诊断结论是："):
                                injury_location = injury_location[6:]  # 移除"诊断结论是："
                            elif injury_location.startswith("医疗结论是："):
                                injury_location = injury_location[6:]  # 移除"医疗结论是："
                            elif injury_location.startswith("结论是："):
                                injury_location = injury_location[4:]  # 移除"结论是："
                            elif injury_location.startswith("诊断为："):
                                injury_location = injury_location[4:]  # 移除"诊断为："

                            # 2. 如果还有冒号，取冒号后面的部分
                            if "：" in injury_location:
                                parts = injury_location.split("：", 1)
                                if len(parts) > 1:
                                    injury_location = parts[1]

                            # 3. 清理结尾的标点
                            injury_location = injury_location.rstrip('。').rstrip('.').rstrip('，').rstrip(',').strip()

                            print(f"✅ 提取的受伤部位: {injury_location}")

                            # 更新到investigation中 - 作为独立变量
                            self.data_model.investigation['医院诊断'] = injury_location
                            self.data_model.investigation['医疗结论'] = injury_location  # 新增独立变量

                # 提取其他可能的信息
                elif "事故发生时间" in text and i + 1 < len(document.paragraphs):
                    next_para = document.paragraphs[i + 1].text
                    if next_para.startswith("答："):
                        basic_info['事故时间'] = next_para[2:].strip()

                elif "事故发生地点" in text and i + 1 < len(document.paragraphs):
                    next_para = document.paragraphs[i + 1].text
                    if next_para.startswith("答："):
                        basic_info['事故地点'] = next_para[2:].strip()

            # 6. 组合文本：受伤经过 + 送医情况（不含医疗结论）
            combined_text = ""

            # 如果有受伤经过
            if thispcjh:
                combined_text = '。'.join(thispcjh)

            # 如果有送医情况，添加到描述中
            if medical_info:
                if combined_text:
                    combined_text += f"。受伤后送医情况：{medical_info}"
                else:
                    combined_text = f"受伤后送医情况：{medical_info}"

            # 如果什么都没有
            if not combined_text:
                combined_text = "详见谈话笔录"

            print(f"📝 提取的原始描述（不含医疗结论）: {combined_text[:200]}...")

            # 7. 如果有AI服务，用AI优化描述（但不优化医疗结论）
            if hasattr(self, 'ai_service') and self.ai_service and combined_text:
                self._set_status('正在用AI优化描述...', 'black')
                QApplication.processEvents()

                try:
                    # 调用AI优化文本（只优化受伤经过和送医情况部分）
                    optimized_text = self.ai_service.optimize_injury_description(combined_text)

                    if optimized_text:
                        # 清理AI输出中的重复标点
                        optimized_text = self._clean_ai_output(optimized_text)

                        # 只存储优化后的描述，不拼接医疗结论
                        self.data_model.investigation['受伤经过'] = optimized_text
                        self._set_status('AI优化完成', 'green')
                        print(f"✅ AI优化后的描述（不含医疗结论）: {optimized_text[:250]}...")
                    else:
                        # AI优化失败，使用原始文本
                        self.data_model.investigation['受伤经过'] = combined_text
                        self._set_status('AI优化失败，使用原始文本', 'orange')
                        print("⚠️ AI优化失败，使用原始文本")

                except Exception as ai_error:
                    print(f"AI优化出错: {ai_error}")
                    # AI优化失败，使用原始文本
                    self.data_model.investigation['受伤经过'] = combined_text
                    self._set_status('AI优化出错，使用原始文本', 'orange')
            else:
                # 没有AI服务，使用原始文本
                self.data_model.investigation['受伤经过'] = combined_text
                if not self.ai_service:
                    self._set_status('未启用AI优化', 'black')

            # 8. 更新其他调查信息
            if '事故时间' in basic_info:
                self.data_model.investigation['事故发生时间'] = basic_info['事故时间']
            if '事故地点' in basic_info:
                self.data_model.investigation['事故发生地点'] = basic_info['事故地点']

            # 9. 更新日期时间信息
            current_date = datetime.datetime.now().strftime("%Y年%m月%d日")
            current_time = datetime.datetime.now().strftime("%H时%M分")
            self.data_model.output_config['当前日期'] = current_date
            self.data_model.output_config['当前时间'] = current_time

            # 10. 生成案本号（如果不存在）
            if not self.data_model.case_info.get('案本号'):
                case_number = self.generate_case_number()
                self.data_model.case_info['案本号'] = case_number

            # 11. 获取公司名称（如果不存在）
            if not self.data_model.company_info.get('公司名称'):
                company_name = self.company_pane.currentText().strip()
                if company_name:
                    self.data_model.company_info['公司名称'] = company_name

            # 12. 准备模板数据
            all_data = self.data_model.to_template_dict()
            if '当前时期' not in all_data or not all_data['当前时期']:
                all_data['当前时期'] = f"{current_date}{current_time}"

            all_data['引用条例'] = regulation_text

            # 13. 确保所有必填字段都有值（新增医疗结论字段）
            required_template_fields = {
                '受伤职工姓名': person_name,
                '本人性别': gender,
                '用人单位': self.data_model.company_info.get('用人单位', ''),
                '事故时间': self.data_model.investigation.get('事故发生时间', '待补充'),
                '事故地点': self.data_model.investigation.get('事故发生地点', '待补充'),
                '诊断结论': self.data_model.investigation.get('医院诊断', '待补充'),
                '医疗结论': self.data_model.investigation.get('医疗结论', '待补充'),  # 新增独立字段
                '受伤经过': self.data_model.investigation.get('受伤经过', '详见谈话笔录'),
                '引用条例': regulation_text,
                '申请单位意见': '同意申请工伤认定',
                '调查人意见': '情况属实，建议认定工伤',
                '负责人意见': '同意调查人意见'
            }

            for field, default_value in required_template_fields.items():
                if field not in all_data or not all_data[field]:
                    all_data[field] = default_value

            # 14. 渲染并保存审批表
            template_path = str(path_utils.get_document_template_path('2022版工伤认定审批表.docx'))
            if not os.path.exists(template_path):
                self._set_status('审批表模板不存在', 'red')
                return

            self._set_status('正在生成审批表...', 'black')
            QApplication.processEvents()

            word = DocxTemplate(template_path)
            word.render(all_data)

            file_name = f"{person_name}案件审批表.docx"
            file_path = os.path.join(self.current_case_folder, file_name)
            word.save(file_path)

            self.dict.update(all_data)

            # 15. 打开文件
            success, message = self.file_service.open_document(file_path)
            if success:
                self._set_status('审批表生成成功', 'green')
            else:
                self._set_status(f'审批表生成成功，但打开失败: {message}', 'orange')

        except Exception as e:
            print(f"生成审批表异常: {e}")
            import traceback
            traceback.print_exc()
            self._set_status(f'生成审批表失败: {str(e)}', 'red')

    def generate_case_number(self) -> str:
        """自动生成案本号（使用配置中的格式）"""
        try:
            is_death_case = self.death_case_checkbox.isChecked()

            if not hasattr(self, '_daily_counter'):
                self._daily_counter = 1

            # 获取案件配置
            case_config = self.config_service.get_case_config()

            # 使用配置中的格式
            current_date = datetime.datetime.now().strftime("%Y%m%d")

            if is_death_case:
                prefix = "工亡"
            else:
                prefix = "案本"

            # 使用配置的格式
            case_number_format = case_config.case_number_format

            # 替换变量
            case_number = case_number_format.format(
                prefix=prefix,
                date=current_date,
                seq=self._daily_counter
            )

            self._daily_counter += 1

            # 保存计数器
            if hasattr(self, 'file_service'):
                self.file_service.save_counter(self._daily_counter)
                print(f"📊 保存计数器: {self._daily_counter}")

            return case_number

        except Exception as e:
            print(f"[generate_case_number] 错误: {e}")
            # 使用默认格式
            current_date = datetime.datetime.now().strftime("%Y%m%d")
            return f"案本-{current_date}-{self._daily_counter:03d}"

    def work_save(self):
        """保存当前案件信息"""
        try:
            current_role = self.get_current_role_type()

            # 如果是证人/法人，检查是否已关联案件
            if current_role in ["证人", "法人"]:
                person_name = self.lineEdit_2.text().strip()

                if not person_name:
                    self._set_status('请先输入受伤职工姓名', 'red')
                    return

                if not self.current_case_folder or not os.path.exists(self.current_case_folder):
                    self._set_status('请先关联案件文件夹（点击搜索按钮）', 'red')
                    return
            case_number = self.generate_case_number()
            self.set_data('案本号', case_number, 'case')

            is_valid, errors = self.validate_current_data()
            if not self.show_validation_errors(errors):
                return

            # 获取人员信息
            role_type = self.get_current_role_type()

            # 获取本人姓名
            person_name = self.get_data('本人姓名', '')
            if not person_name:
                person_name = self.lineEdit_2.text().strip()
                if person_name:
                    self.set_data('本人姓名', person_name, 'basic')
                else:
                    self._set_status('请输入本人姓名', 'red')
                    return

            # 获取角色姓名
            role_name_key = f"{role_type}姓名"
            role_name = self.get_data(role_name_key, '')
            if not role_name:
                role_name = self.name_pane.text().strip()
                if role_name:
                    self.set_data(role_name_key, role_name, 'basic')
                else:
                    self._set_status(f'请输入{role_type}姓名', 'red')
                    return

            # 获取公司信息
            company_info = self.get_company_info()
            company_name = company_info['公司名称']
            employer = company_info['用人单位']
            site = company_info['工地名称']

            # 更新公司数据
            if company_name:
                self.set_data('公司名称', company_name, 'company')
            if employer:
                self.set_data('用人单位', employer, 'company')
            if site:
                self.set_data('工地名称', site, 'company')

            number = self.comboBox.currentIndex()
            has_employer = bool(employer)
            has_site = bool(site)

            # 获取案件配置
            is_death_case = self.death_case_checkbox.isChecked()
            is_personal = self.personal_application_checkbox.isChecked()

            # 使用TemplateService获取模板路径
            self.open_file_path = self.template_service.get_template_path(
                role=role_type,
                case_type=number,
                is_death_case=is_death_case,
                is_personal=is_personal
            )

            # 检查模板文件是否存在
            if not os.path.exists(self.open_file_path):
                print(f"⚠️ 模板文件不存在: {self.open_file_path}")

                # 尝试查找模板文件
                found_template = self._find_template_file(role_type, number, is_death_case, is_personal)
                if found_template:
                    self.open_file_path = found_template
                    print(f"✅ 找到替代模板: {self.open_file_path}")
                else:
                    self._set_status('模板文件不存在，请检查模板配置', 'red')
                    return

            # 收集所有变量
            all_variables = self.var_manager.collect_variables(
                role_type, number, has_employer, has_site
            )
            all_variables['案本号'] = case_number
            all_variables = self._ensure_required_fields(all_variables, role_type)

            # 添加自我介绍
            if role_type == "法人":
                position = self.get_data('法人职务', '')
            else:
                position = self.get_data(f'{role_type}岗位', '')
            id_address = self.get_data(f'{role_type}身份证地址', '')

            self_intro = self.var_manager.get_self_introduction(
                role_type, company_name, employer, site, role_name, position, id_address
            )
            all_variables['自我介绍内容'] = self_intro

            # 更新字典
            self.dict.clear()
            self.dict.update(all_variables)

            # 创建或获取案件文件夹 - 使用增强版
            if person_name != self.current_person_name or not self.current_case_folder:
                self.current_person_name = person_name

                # 准备案件信息字典
                case_info = {
                    'case_number': case_number,  # 已生成的案本号
                    'person_name': person_name.strip(),  # 姓名（去除首尾空格）
                    'id_card': self.get_data('本人身份证号', '')  # 身份证号
                }

                # 调用新的文件夹创建方法
                self.current_case_folder = self.file_service.create_enhanced_case_folder(
                    self.BASE_PATH,  # 基础存储路径
                    case_info  # 案件信息
                )

                print(f"📁 创建案件文件夹: {os.path.basename(self.current_case_folder)}")

            # 生成文件名
            base_name = os.path.basename(self.open_file_path).replace('.docx', '')
            file_name = f"{base_name}（{role_type}）{role_name}.docx"
            target_path = os.path.join(self.current_case_folder, file_name)

            try:
                # 处理模板
                temp_template_path = self.template_service.process_template(
                    template_path=self.open_file_path,
                    role_type=role_type,
                    case_type=number,
                    variables=all_variables,
                    need_special_questions=self.need_special_questions()
                )

                if not temp_template_path or not os.path.exists(temp_template_path):
                    self._set_status('模板处理失败', 'red')
                    return

                # 渲染并保存文档
                from docxtpl import DocxTemplate
                doc = DocxTemplate(temp_template_path)
                doc.render(all_variables)
                doc.save(target_path)

                self._set_status(f'文件保存成功，案本号: {case_number}', 'green')

                # 保存成功后，更新按钮状态（如果是本人笔录）
                if role_type == "本人":
                    self.person_transcript_created = True
                    # 这里不直接禁用按钮，让on_talk_button_clicked方法处理
                    print(f"✅ 本人笔录保存成功，设置 person_transcript_created = True")

                self._add_case_to_index(
                    case_number=case_number,
                    person_name=person_name,
                    folder_path=self.current_case_folder,
                    transcript_file=file_name
                )

                # 打开文件
                success, message = self.file_service.open_document(target_path)
                if not success:
                    self._set_status(f'文件保存成功，但打开失败: {message}', 'orange')

                # 清理临时文件
                self._cleanup_temp_file(temp_template_path)

            except Exception as e:
                self._set_status(f'文件保存失败: {str(e)}', 'red')
                import traceback
                traceback.print_exc()

        except Exception as e:
            self._set_status(f'保存过程异常: {str(e)}', 'red')
            import traceback
            traceback.print_exc()

    def _add_case_to_index(self, case_number, person_name, folder_path, transcript_file):
        """
        添加案件到索引

        Args:
            case_number: 案本号
            person_name: 受伤职工姓名
            folder_path: 案件文件夹路径
            transcript_file: 本人笔录文件名
        """
        try:
            # 确保导入datetime
            import datetime

            print(f"📝 正在添加案件到索引: {case_number}")

            # 创建索引条目
            from case_index import CaseIndexEntry

            case_entry = CaseIndexEntry(
                case_number=case_number,
                person_name=person_name,
                id_card_last4=self.get_data('本人身份证号', '')[-4:] if self.get_data('本人身份证号', '') else 'xxxx',
                company_name=self.company_pane.currentText().strip(),
                case_type="工亡" if self.death_case_checkbox.isChecked() else "工伤",
                created_date=datetime.datetime.now().strftime("%Y-%m-%d"),
                folder_name=os.path.basename(folder_path),
                transcript_file=transcript_file,
                updated_time=datetime.datetime.now().isoformat()
            )

            # 添加到索引
            success = self.case_index_manager.add_case(case_entry)

            if success:
                print(f"✅ 案件已添加到索引: {case_number}")
                # 保存当前案本号，供后续使用
                self.current_case_number = case_number
                # 更新数据模型
                self.set_data('案本号', case_number, 'case')
            else:
                print("⚠️ 案件索引添加失败，但不影响主要功能")

        except Exception as e:
            print(f"❌ 添加案件到索引失败（非致命错误）: {e}")
            # 这里不抛出异常，避免影响主流程

    def _find_template_file(self, role_type, case_type, is_death_case, is_personal):
        """查找模板文件"""
        # 可能的模板名称
        template_names = []

        if is_death_case:
            if is_personal:
                template_names.append(f"{role_type}谈话笔录（个人申请工亡案件）.docx")
            template_names.append(f"{role_type}谈话笔录（工亡案件）.docx")
        else:
            case_templates = {
                0: f"{role_type}谈话笔录（普通案件）.docx",
                1: f"{role_type}谈话笔录（工作前后案件）.docx",
                2: f"{role_type}谈话笔录（暴力伤害案件）.docx",
                3: f"{role_type}谈话笔录（患职业病案件）.docx",
                4: f"{role_type}谈话笔录（因工外出案件）.docx",
                5: f"{role_type}谈话笔录（上下班时案件）.docx"
            }
            template_names.append(case_templates.get(case_type, f"{role_type}谈话笔录（普通案件）.docx"))

        # 添加默认名称
        template_names.append(f"{role_type}谈话笔录（普通案件）.docx")
        template_names.append(f"{role_type}谈话笔录.docx")
        template_names.append("谈话笔录.docx")

        # 在多个位置查找
        search_paths = [
            self.TEMPLATE_PATH,
            os.path.join(self.TEMPLATE_PATH, "谈话模板"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "resource", "模板文件", "谈话模板"),
        ]

        for search_path in search_paths:
            if os.path.exists(search_path):
                for template_name in template_names:
                    template_path = os.path.join(search_path, template_name)
                    if os.path.exists(template_path):
                        return template_path

        return None

    def _cleanup_temp_file(self, temp_template_path):
        """清理临时模板文件"""
        try:
            if (temp_template_path and
                    os.path.exists(temp_template_path) and
                    temp_template_path != self.open_file_path):
                os.remove(temp_template_path)
        except Exception as e:
            print(f"⚠️ 清理临时文件失败: {e}")

    def get_company_info(self) -> Dict[str, str]:
        """获取公司相关信息"""
        company_name = self.company_pane.currentText().strip()
        employer = self.construction_company.currentText().strip()
        site = self.constuction_plant.currentText().strip()

        if not company_name:
            company_name = self.get_data('公司名称', '')
        if not employer:
            employer = self.get_data('用人单位', '')
        if not site:
            site = self.get_data('工地名称', '')

        return {
            '公司名称': company_name,
            '用人单位': employer,
            '工地名称': site
        }

    def validate_current_data(self) -> Tuple[bool, List[str]]:
        """验证当前数据（使用DataService）"""
        # 收集当前数据
        current_data = {}

        role = self.get_current_role_type()

        # 人员数据
        if role in ["本人", "证人", "法人"]:
            current_data.update({
                '姓名': self.get_data(f"{role}姓名", ''),
                '身份证号': self.get_data(f"{role}身份证号", ''),
                '年龄': self.get_data(f"{role}年龄", '')
            })

        # 公司数据
        current_data['公司名称'] = self.get_data('公司名称', '')

        # 案件数据
        current_data['案本号'] = self.get_data('案本号', '')

        # 使用DataService验证
        errors = self.data_service.validate_case_data(current_data)

        return len(errors) == 0, errors

    def _ensure_required_fields(self, variables: Dict[str, Any], role: str) -> Dict[str, Any]:
        """确保模板所需的所有字段都存在"""
        required_fields = [
            '当前时期', '本人年龄', '本人性别', '本人姓名',
            '本人身份证号', '本人身份证地址', '本人手机号', '公司名称'
        ]

        for field in required_fields:
            if field not in variables or not variables[field]:
                value = None

                if field in ['本人年龄', '本人性别', '本人姓名', '本人身份证号', '本人身份证地址']:
                    data_key = field if field.startswith('本人') else f'本人{field}'
                    value = self.data_model.basic_info.get(data_key, '')

                elif field == '公司名称':
                    value = self.data_model.company_info.get('公司名称', '')

                elif field == '当前时期':
                    current_date = variables.get('当前日期', datetime.datetime.now().strftime('%Y年%m月%d日'))
                    current_time = variables.get('当前时间', datetime.datetime.now().strftime('%H时%M分'))
                    value = f"{current_date}{current_time}"

                if not value:
                    if field == '本人年龄' and self.age_pane.text():
                        value = self.age_pane.text()
                    elif field == '本人性别' and self.lineEdit.text():
                        value = self.lineEdit.text()
                    elif field == '本人姓名' and self.name_pane.text():
                        value = self.name_pane.text()
                    elif field == '本人手机号' and self.lineEdit_4.text():
                        value = self.lineEdit_4.text()

                if value:
                    variables[field] = value

        return variables

    def init_comboboxes(self):
        """初始化所有组合框"""
        self.init_combobox(self.company_pane, self.items_list)
        self.init_combobox(self.construction_company, self.items_list1)
        self.init_combobox(self.constuction_plant, self.items_list2)

        self.company_pane.setCurrentIndex(-1)
        self.construction_company.setCurrentIndex(-1)
        self.constuction_plant.setCurrentIndex(-1)

    def company(self):
        """更新公司信息"""
        company_name = self.company_pane.currentText().strip()
        self.set_data('公司名称', company_name, 'company')

    def sync_employer_to_dict(self):
        """更新用人单位信息"""
        employer_name = self.construction_company.currentText().strip()
        self.set_data('用人单位', employer_name, 'company')

    def c_plant(self):
        """更新工地名称信息"""
        site_name = self.constuction_plant.currentText().strip()
        self.set_data('工地名称', site_name, 'company')

    def id_clicked(self):
        """读取身份证信息"""
        try:
            dll = windll.LoadLibrary("./sdtapi.dll")
            port = c_int32(1001)
            ifopen = c_int32(1)
            pucManaInfo = create_string_buffer(4)
            pucManaMsg = create_string_buffer(8)
            dll.SDT_StartFindIDCard(port, pucManaInfo, ifopen)
            dll.SDT_SelectIDCard(port, pucManaMsg, ifopen)
            pucCHMsg = create_unicode_buffer(256)
            pucPHMsg = create_string_buffer(1024)
            puiCHMsgLen = c_uint(0)
            puiPHMsgLen = c_uint(0)
            ret = dll.SDT_ReadBaseMsg(port, pucCHMsg, byref(puiCHMsgLen), pucPHMsg,
                                      byref(puiPHMsgLen), ifopen)
            if ret == 65:
                return
            dll.SDT_ClosePort(port)
            self.set_data('当前时期', datetime.datetime.now().strftime("%Y年%m月%d日%H时%M分"), 'output')

            role = self.get_current_role_type()
            self.process_id(pucCHMsg, role)

        except Exception as e:
            print(f"读取身份证信息时出错: {str(e)}")

    def process_id(self, pucCHMsg, role):
        """处理身份证信息"""
        try:
            name = pucCHMsg.value[0:15].strip()
            self.data_model.update_basic_info(role, {'姓名': name})
            self.set_data(f"{role}姓名", name, 'basic')
            self.name_pane.setText(name)

            if len(pucCHMsg.value) >= 79:
                id_number = pucCHMsg.value[61:79].strip()
                self.data_model.update_basic_info(role, {'身份证号': id_number})
                self.set_data(f"{role}身份证号", id_number, 'basic')
                self.idnumer_pane.setText(id_number)

            if len(pucCHMsg.value) >= 61:
                address = pucCHMsg.value[26:61].strip()
                self.data_model.update_basic_info(role, {'身份证地址': address})
                self.set_data(f"{role}身份证地址", address, 'basic')
                self.textEdit.setText(address)

            self.process_id_info(role)
            self.calculate_age_from_id(role)

        except Exception as e:
            import traceback
            traceback.print_exc()

    def update_role_info(self, role):
        """更新角色信息"""
        try:
            name = self.name_pane.text().strip()
            idcard = self.idnumer_pane.text().strip()
            address = self.textEdit.toPlainText().strip()
            phone = self.lineEdit_4.text().strip()
            position = self.lineEdit_5.text().strip()

            data_updates = {
                f'{role}姓名': name,
                f'{role}身份证号': idcard,
                f'{role}身份证地址': address,
                f'{role}手机号': phone,
            }

            if role == "法人":
                data_updates[f'{role}职务'] = position
            else:
                data_updates[f'{role}岗位'] = position

            for key, value in data_updates.items():
                if value:
                    self.set_data(key, value, 'basic')

            if role == "本人":
                self.lineEdit_2.setText(name)

            if role == "法人":
                company_name = self.get_data('公司名称', '')
                if not company_name:
                    company_name = self.company_pane.currentText().strip()
                    if company_name:
                        self.set_data('公司名称', company_name, 'company')

        except Exception as e:
            import traceback
            traceback.print_exc()

    def work_tlak(self):
        """根据当前角色更新信息并保存案件"""
        role = self.get_current_role_type()
        self.update_role_info(role)
        self.discriminate()
        self.work_save()

    def discriminate(self):
        """生成模板文件路径"""
        number = self.comboBox.currentIndex()
        role = self.get_current_role_type()
        has_construction_company = bool(self.construction_company.currentText())
        has_constuction_plant = bool(self.constuction_plant.currentText())

        # 获取案件配置
        is_death_case, is_personal = self.get_current_case_config_tuple()

        self.open_file_path = self.template_service.get_template_path(
            role=role,
            case_type=number,
            is_death_case=is_death_case,
            is_personal=is_personal
        )
        if not os.path.exists(self.open_file_path):
            self._set_status('未找到模板文件，请检查配置', 'red')

    def process_id_info(self, role):
        """处理身份证信息并更新性别显示（使用DataService）"""
        try:
            idcard = self.get_data(f"{role}身份证号", "")
            if not idcard:
                return

            # 使用DataService提取性别
            gender = self.data_service.extract_gender_from_idcard(idcard)
            if gender is None:
                print(f"[process_id_info] 无法从身份证提取性别: {idcard}")
                return

            # 设置性别数据
            self.set_data(f"{role}性别", gender, 'basic')
            self.lineEdit.setText(gender)

        except Exception as e:
            print(f"[process_id_info] 错误: {e}")
            import traceback
            traceback.print_exc()

    def get_current_role_type(self) -> str:
        """获取当前选中的角色类型"""
        if self.radioButton.isChecked():
            return "本人"
        elif self.radioButton_2.isChecked():
            return "证人"
        elif self.radioButton_3.isChecked():
            return "法人"
        return "本人"

    def get_current_case_config_tuple(self) -> tuple:
        """获取案件配置（工亡/个人申请）"""
        is_death_case = self.death_case_checkbox.isChecked()
        is_personal = self.personal_application_checkbox.isChecked()
        return is_death_case, is_personal

    def need_special_questions(self) -> bool:
        """判断是否需要插入特殊问题"""
        is_death_case, is_personal = self.get_current_case_config_tuple()
        return is_death_case or is_personal

    def new_case(self):
        """开始新案件"""
        self.current_case_folder = None
        self.current_person_name = ""
        self.clear_fields()
        self.lineEdit_2.clear()

        # 重置谈话笔录按钮状态
        self.person_transcript_created = False
        self.pushButton.setEnabled(True)
        self.pushButton.setStyleSheet("")  # 恢复默认样式
        print("🆕 开始新案件，谈话笔录按钮已重置")

        self._set_status('已开始新案件', 'green')

    def smart_search_cases(self):
        """
        智能搜索案件 - 点击 pushButton_6 时调用
        1. 获取 lineEdit_2 中的姓名
        2. 搜索匹配的文件夹
        3. 根据结果数量自动处理
        """
        try:
            # 1. 获取受伤职工姓名
            person_name = self.lineEdit_2.text().strip()
            if not person_name:
                QMessageBox.warning(self, "提示", "请输入受伤职工姓名")
                return

            # 2. 搜索案件
            matched_cases = self.file_service.search_cases_by_person_name(
                self.BASE_PATH, person_name
            )

            if not matched_cases:
                QMessageBox.information(self, "提示",
                                        f"未找到'{person_name}'的已有案件")
                return

            # 3. 根据结果数量处理
            if len(matched_cases) == 1:
                # 只有一个匹配，自动选择
                self._select_case_automatically(matched_cases[0])
            else:
                # 多个匹配，显示选择对话框
                self._show_case_selection_dialog(matched_cases, person_name)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"搜索失败: {str(e)}")
            import traceback
            traceback.print_exc()

    def _select_case_automatically(self, case_info: Dict[str, Any]):
        """自动选择单个案件"""
        # 设置当前文件夹
        self.current_case_folder = case_info['folder_path']
        self.current_person_name = case_info['person_name']

        # 更新显示
        self.lineEdit_2.setText(case_info['person_name'])

        # 显示提示
        info = (f"已关联到案件: {case_info['case_number']}\n"
                f"创建时间: {case_info['created_date']}\n"
                f"已有文件: {case_info['file_count']}个")

        QMessageBox.information(self, "关联成功", info)

        # 更新状态标签
        self._set_status(f"已关联: {case_info['case_number']}", 'green')

    def _show_case_selection_dialog(self, cases: List[Dict], person_name: str):
        """显示案件选择对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"选择案件 - {person_name} ({len(cases)}个)")
        dialog.resize(600, 400)

        layout = QVBoxLayout()

        # 标题
        title = QLabel(f"找到{len(cases)}个'{person_name}'的案件，请选择:")
        layout.addWidget(title)

        # 案件列表表格
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(['选择', '案本号', '创建日期', '文件数量', '已有本人笔录'])
        table.setRowCount(len(cases))

        # 单选按钮组
        button_group = QButtonGroup()

        for i, case in enumerate(cases):
            # 单选按钮
            radio_btn = QRadioButton()
            if i == 0:  # 默认选择第一个
                radio_btn.setChecked(True)
            button_group.addButton(radio_btn, i)

            # 案本号
            case_number = QTableWidgetItem(case['case_number'])
            case_number.setFlags(case_number.flags() ^ Qt.ItemIsEditable)

            # 创建日期
            created_date = QTableWidgetItem(case['created_date'])

            # 文件数量
            file_count = QTableWidgetItem(str(case['file_count']))

            # 是否有本人笔录
            has_person = QTableWidgetItem("✅ 有" if case['has_person'] else "❌ 无")

            # 设置到表格
            table.setCellWidget(i, 0, radio_btn)
            table.setItem(i, 1, case_number)
            table.setItem(i, 2, created_date)
            table.setItem(i, 3, file_count)
            table.setItem(i, 4, has_person)

        table.resizeColumnsToContents()
        layout.addWidget(table)

        # 按钮区域
        btn_layout = QHBoxLayout()

        btn_select = QPushButton("选择")
        btn_select.clicked.connect(lambda: self._on_case_selected(
            cases, button_group.checkedId(), dialog
        ))

        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(dialog.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_select)
        btn_layout.addWidget(btn_cancel)

        layout.addLayout(btn_layout)
        dialog.setLayout(layout)

        dialog.exec_()

    def _on_case_selected(self, cases: List[Dict], selected_index: int, dialog: QDialog):
        """用户选择案件后的处理"""
        if selected_index < 0:
            QMessageBox.warning(dialog, "提示", "请先选择一个案件")
            return

        selected_case = cases[selected_index]

        # 设置当前文件夹
        self.current_case_folder = selected_case['folder_path']
        self.current_person_name = selected_case['person_name']

        # 更新显示
        self.lineEdit_2.setText(selected_case['person_name'])

        dialog.accept()

        # 显示成功消息
        QMessageBox.information(self, "关联成功",
                                f"已关联到案件: {selected_case['case_number']}")


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())