import os
import sys
import datetime
import logging
from pathlib import Path
from typing import Dict, List, Any, Tuple
from ctypes import windll, byref, create_string_buffer, c_int32, c_uint
import pandas as pd
from docx import Document
from docxtpl import DocxTemplate
from PyQt5.Qt import *
from PyQt5.QtCore import QThread, pyqtSignal, QTimer, Qt
from PyQt5.QtWidgets import (
    QApplication, QWidget, QMessageBox, QDialog, QVBoxLayout,
    QLabel, QTextEdit, QPushButton, QHBoxLayout, QInputDialog,
    QLineEdit, QComboBox, QCompleter, QCheckBox, QProgressDialog,
    QTableWidget, QTableWidgetItem, QRadioButton, QButtonGroup,
    QGroupBox
)
from PyQt5.QtGui import QFont

from ui_main_window import Ui_Form
from file_service import FileService
from data_service import DataService
from template_service import TemplateService, TemplateVariableManager
from ai_service import AIService
from config_service import ConfigService
from path_utils import path_utils
from case_index import CaseIndexManager, CaseIndexEntry
from main import UserManager, PasswordLineEdit

# 设置日志级别
logging.getLogger('config_service').setLevel(logging.WARNING)


# ============================================================================
# 工具函数
# ============================================================================

def _date_now() -> str:
    """当前日期，格式：2025年01月01日"""
    import datetime as _dt
    return _dt.datetime.now().strftime('%Y年%m月%d日')


def _time_now() -> str:
    """当前时间，格式：14时30分"""
    import datetime as _dt
    return _dt.datetime.now().strftime('%H时%M分')


def _timestamp_now() -> str:
    """时间戳，格式：20250101_143000"""
    import datetime as _dt
    return _dt.datetime.now().strftime('%Y%m%d_%H%M%S')


# ============================================================================
# AIWorker
# ============================================================================

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
            '当前日期': _date_now(),
            '当前时间': _time_now()
        })

    def to_template_dict(self) -> Dict[str, Any]:
        """转换为模板渲染用的字典"""
        template_dict = {}

        # 确保日期时间是最新的
        self.output_config.update({
            '当前日期': _date_now(),
            '当前时间': _time_now()
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


# ============================================================================
# F2 轮换测试数据集
# ============================================================================
TEST_DATA_PRESETS = [
    # ═══════════════════════════════════════════════════════════════
    # 案件1: 张三 — 普通工伤(第一项) — 单位申请 — ZZ新城项目
    # 同案三人: 本人(张三) + 证人(李四) + 法人(王五)
    # ═══════════════════════════════════════════════════════════════
    {
        "name": "1/12 张三案-本人谈话",
        "role": "本人",
        "name_pane": "张三",
        "idnumer_pane": "330324199003151234",
        "textEdit": "浙江省永嘉县瓯北街道XX路88号",
        "lineEdit_4": "13888880001",
        "lineEdit_5": "泥水工",
        "lineEdit_2": "张三",
        "comboBox": 0,
        "company_pane": "永嘉县XX建设工程有限公司",
        "construction_company": "温州YY建筑劳务有限公司",
        "construction_plant": "ZZ新城项目一期工地",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "申请人张三，男，1990年3月15日出生，身份证号330324199003151234，住永嘉县瓯北街道XX路88号。2026年7月20日下午16时20分许，在ZZ新城项目一期工地3号楼5层搬运水泥时，被滑落的水泥袋砸伤右脚，经永嘉县人民医院诊断为右足跖骨骨折。该事故属于在工作时间和工作场所内因工作原因受到的事故伤害，应当认定为工伤。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "医院诊断证明书", "provided": True, "notes": "永嘉县人民医院 右足跖骨骨折"},
            {"name": "劳动合同", "provided": False, "notes": ""},
            {"name": "工资银行流水", "provided": False, "notes": ""},
            {"name": "考勤记录", "provided": False, "notes": ""},
        ],
    },
    {
        "name": "2/12 张三案-证人谈话(李四)",
        "role": "证人",
        "name_pane": "李四",
        "idnumer_pane": "330324198508121235",
        "textEdit": "浙江省永嘉县桥头镇YY村123号",
        "lineEdit_4": "13966660002",
        "lineEdit_5": "钢筋工",
        "lineEdit_2": "张三",                         # ← 指向同案受伤职工
        "comboBox": 0,
        "company_pane": "永嘉县XX建设工程有限公司",
        "construction_company": "温州YY建筑劳务有限公司",
        "construction_plant": "ZZ新城项目一期工地",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "目击证人李四，与张三同为ZZ新城项目一期工地的工友。2026年7月20日下午16时20分许，李四在3号楼5层作业时，亲眼看到张三搬运水泥袋时水泥袋滑落砸中右脚，张三当场倒地呼痛。李四立即上前查看并通知了班组长。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "证人证言", "provided": True, "notes": ""},
        ],
    },
    {
        "name": "3/12 张三案-法人谈话(王五)",
        "role": "法人",
        "name_pane": "王五",
        "idnumer_pane": "330324198212011234",
        "textEdit": "浙江省永嘉县上塘镇XX路66号",
        "lineEdit_4": "13777770003",
        "lineEdit_5": "法定代表人",
        "lineEdit_2": "张三",                         # ← 指向同案受伤职工
        "comboBox": 0,
        "company_pane": "永嘉县XX建设工程有限公司",
        "construction_company": "温州YY建筑劳务有限公司",
        "construction_plant": "ZZ新城项目一期工地",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "永嘉县XX建设工程有限公司法定代表人王五确认：张三系我公司正式职工，于2026年3月1日入职，签订有书面劳动合同。公司未为其参加工伤保险。2026年7月20日下午受伤时，张三正在执行公司安排的正常工作任务。公司认可其受伤属于工伤，愿意配合认定并支付相关费用。",
        "materials": [
            {"name": "公司营业执照副本", "provided": True, "notes": ""},
            {"name": "法定代表人身份证", "provided": True, "notes": ""},
            {"name": "劳动合同", "provided": True, "notes": "2026年3月1日签订"},
            {"name": "工资发放记录", "provided": True, "notes": "月薪6000元 银行转账"},
            {"name": "考勤打卡记录", "provided": True, "notes": ""},
        ],
    },

    # ═══════════════════════════════════════════════════════════════
    # 案件2: 孙七 — 工亡 — 个人申请 — ZZ新城项目
    # 同案三人: 本人(孙七,已故) + 证人(李四) + 法人(王五)
    # ═══════════════════════════════════════════════════════════════
    {
        "name": "4/12 孙七案-本人(工亡)",
        "role": "本人",
        "name_pane": "孙七",
        "idnumer_pane": "330324198704201237",
        "textEdit": "浙江省永嘉县岩坦镇XX村12号",
        "lineEdit_4": "13700001111",
        "lineEdit_5": "钢筋工",
        "lineEdit_2": "孙七",
        "comboBox": 0,
        "company_pane": "永嘉县XX建设工程有限公司",
        "construction_company": "温州YY建筑劳务有限公司",
        "construction_plant": "ZZ新城项目一期工地",
        "deathCaseCheckbox": True,
        "personalApplicationCheckbox": True,
        "statement_edit": "申请人系死者孙七的妻子王九，女，1988年6月出生，身份证号330324198806012345，住永嘉县岩坦镇XX村12号。孙七于2026年8月3日上午9时许，在ZZ新城项目一期工地2号楼进行钢筋绑扎作业时，不慎从8米高处坠落，经抢救无效于当日11时20分死亡。死者生前系永嘉县XX建设工程有限公司钢筋工，月工资8000元。申请人认为孙七在工作中因事故死亡，应当认定为工亡。",
        "materials": [
            {"name": "申请人身份证复印件", "provided": True, "notes": "王九 330324198806012345"},
            {"name": "死者身份证复印件", "provided": True, "notes": "孙七 330324198704201237"},
            {"name": "死亡证明", "provided": True, "notes": "永嘉县人民医院出具"},
            {"name": "医院急救记录", "provided": True, "notes": ""},
            {"name": "婚姻证明", "provided": False, "notes": ""},
            {"name": "劳动合同", "provided": False, "notes": ""},
            {"name": "火化证明", "provided": False, "notes": ""},
        ],
    },
    {
        "name": "5/12 孙七案-证人谈话(李四)",
        "role": "证人",
        "name_pane": "李四",
        "idnumer_pane": "330324198508121235",
        "textEdit": "浙江省永嘉县桥头镇YY村123号",
        "lineEdit_4": "13966660002",
        "lineEdit_5": "钢筋工",
        "lineEdit_2": "孙七",                         # ← 指向同案受伤职工
        "comboBox": 0,
        "company_pane": "永嘉县XX建设工程有限公司",
        "construction_company": "温州YY建筑劳务有限公司",
        "construction_plant": "ZZ新城项目一期工地",
        "deathCaseCheckbox": True,
        "personalApplicationCheckbox": True,
        "statement_edit": "目击证人李四，系孙七的同班组工友。2026年8月3日上午9时许，李四与孙七同在ZZ新城项目一期工地2号楼进行钢筋绑扎作业。孙七在8米高的脚手架平台上作业时，因脚下跳板断裂，从高处坠落至地面。李四立即呼救并拨打120，与班组长一起将孙七送往永嘉县人民医院抢救。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "证人证言", "provided": True, "notes": ""},
        ],
    },
    {
        "name": "6/12 孙七案-法人谈话(王五)",
        "role": "法人",
        "name_pane": "王五",
        "idnumer_pane": "330324198212011234",
        "textEdit": "浙江省永嘉县上塘镇XX路66号",
        "lineEdit_4": "13777770003",
        "lineEdit_5": "法定代表人",
        "lineEdit_2": "孙七",                         # ← 指向同案受伤职工
        "comboBox": 0,
        "company_pane": "永嘉县XX建设工程有限公司",
        "construction_company": "温州YY建筑劳务有限公司",
        "construction_plant": "ZZ新城项目一期工地",
        "deathCaseCheckbox": True,
        "personalApplicationCheckbox": True,
        "statement_edit": "永嘉县XX建设工程有限公司法定代表人王五确认：孙七系我公司正式职工，于2025年6月入职，签订有劳动合同。孙七的具体工作由温州YY建筑劳务有限公司安排管理。2026年8月3日事故发生后，公司已垫付了全部抢救费用共计3.2万元。公司认可孙七因工死亡，愿意配合家属办理工亡认定手续。",
        "materials": [
            {"name": "公司营业执照副本", "provided": True, "notes": ""},
            {"name": "法定代表人身份证", "provided": True, "notes": ""},
            {"name": "劳动合同", "provided": True, "notes": "2025年6月签订"},
            {"name": "工资发放记录", "provided": True, "notes": "月薪8000元"},
            {"name": "考勤打卡记录", "provided": True, "notes": ""},
            {"name": "医院抢救费用单据", "provided": True, "notes": "合计3.2万元"},
        ],
    },

    # ═══════════════════════════════════════════════════════════════
    # 案件3: 赵六 — 上下班途中(第六项) — 单位申请
    # 单人案件: 只有本人
    # ═══════════════════════════════════════════════════════════════
    {
        "name": "7/12 赵六案-本人(上下班途中)",
        "role": "本人",
        "name_pane": "赵六",
        "idnumer_pane": "330324199508031238",
        "textEdit": "浙江省永嘉县乌牛街道XX小区5栋301室",
        "lineEdit_4": "13655550004",
        "lineEdit_5": "装配工",
        "lineEdit_2": "赵六",
        "comboBox": 5,
        "company_pane": "温州BB电器有限公司",
        "construction_company": "",
        "construction_plant": "",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "申请人赵六，系温州BB电器有限公司装配工，住永嘉县乌牛街道XX小区5栋301室。2026年8月5日下午17时40分许，申请人下班后骑电动车沿104国道从公司回乌牛街道家中，在104国道乌牛段被一辆小型轿车追尾，经永嘉县交警大队认定对方负全部责任。事故造成申请人左腿胫骨骨折，已送永嘉县人民医院治疗。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "道路交通事故认定书", "provided": True, "notes": "对方全责"},
            {"name": "医院诊断证明", "provided": True, "notes": "左腿胫骨骨折"},
            {"name": "劳动合同", "provided": True, "notes": ""},
            {"name": "路线示意图", "provided": True, "notes": "乌牛街道→104国道→公司"},
            {"name": "考勤记录", "provided": False, "notes": ""},
        ],
    },

    # ═══════════════════════════════════════════════════════════════
    # 案件4: 刘十 — 预备性工作(第二项) — 单位申请
    # 同案两人: 本人(刘十) + 证人(钱七)
    # ═══════════════════════════════════════════════════════════════
    {
        "name": "8/12 刘十案-本人(预备性工作)",
        "role": "本人",
        "name_pane": "刘十",
        "idnumer_pane": "330324199207052345",
        "textEdit": "浙江省永嘉县大若岩镇XX村33号",
        "lineEdit_4": "13511112222",
        "lineEdit_5": "冲压工",
        "lineEdit_2": "刘十",
        "comboBox": 1,
        "company_pane": "温州AA金属制品有限公司",
        "construction_company": "",
        "construction_plant": "",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "申请人刘十，系温州AA金属制品有限公司冲压工。2026年7月25日上午7时40分许（上班时间为8时），申请人按照惯例提前到岗对公司冲压设备进行例行检查和预热，在此过程中右手被机器夹伤，经温州附二医诊断为右手食指和中指挤压伤。该检查和预热工作系申请人职责范围内的预备性工作，应当认定为工伤。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "医院诊断证明", "provided": True, "notes": "右手食指和中指挤压伤"},
            {"name": "劳动合同", "provided": True, "notes": ""},
            {"name": "操作规程手册", "provided": True, "notes": "车间设备预热流程"},
            {"name": "考勤记录", "provided": True, "notes": "7月25日打卡7:35"},
            {"name": "工资发放记录", "provided": False, "notes": ""},
        ],
    },
    {
        "name": "9/12 刘十案-证人谈话(钱七)",
        "role": "证人",
        "name_pane": "钱七",
        "idnumer_pane": "330324199906081239",
        "textEdit": "浙江省永嘉县岩头镇ZZ村88号",
        "lineEdit_4": "13544440005",
        "lineEdit_5": "冲压工",
        "lineEdit_2": "刘十",                         # ← 指向同案受伤职工
        "comboBox": 1,
        "company_pane": "温州AA金属制品有限公司",
        "construction_company": "",
        "construction_plant": "",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "证人钱七，与刘十同为温州AA金属制品有限公司冲压车间工人。2026年7月25日上午约7时40分，钱七刚到车间就看到刘十右手流血被卡在冲压设备中，立即跑过去按下急停按钮并呼叫其他工友帮忙。刘十称是在做设备预热检查时不小心触发了机器。钱七证明该设备每天开工前确实需要预热检查，这是车间一直以来的惯例。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "证人证言", "provided": True, "notes": ""},
        ],
    },

    # ═══════════════════════════════════════════════════════════════
    # 案件5: 周八 — 暴力伤害(第三项) — 个人申请
    # 单人案件: 只有本人
    # ═══════════════════════════════════════════════════════════════
    {
        "name": "10/12 周八案-本人(暴力伤害)",
        "role": "本人",
        "name_pane": "周八",
        "idnumer_pane": "330324199106011240",
        "textEdit": "浙江省永嘉县枫林镇XX村55号",
        "lineEdit_4": "13433330006",
        "lineEdit_5": "保安",
        "lineEdit_2": "周八",
        "comboBox": 2,
        "company_pane": "温州EE物业管理有限公司",
        "construction_company": "",
        "construction_plant": "",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": True,
        "statement_edit": "申请人周八，系温州EE物业管理有限公司保安，派驻永嘉县XX商场担任安保工作。2026年8月1日晚21时许，申请人在商场一楼巡逻时发现一名男子正在盗窃商户商品，上前制止时被对方用随身携带的铁棍击打头部和右臂。商场其他保安闻讯赶来将嫌疑人控制并报警。申请人被送至永嘉县人民医院治疗，诊断为：脑震荡、右臂桡骨骨折。永嘉县公安局已对该盗窃嫌疑人立案侦查。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "医院诊断证明", "provided": True, "notes": "脑震荡、右臂桡骨骨折"},
            {"name": "公安局报案回执", "provided": True, "notes": "永嘉县公安局"},
            {"name": "商场监控录像", "provided": True, "notes": "XX商场一楼"},
            {"name": "保安服务合同", "provided": True, "notes": "温州EE物业→XX商场"},
            {"name": "劳动合同", "provided": False, "notes": ""},
        ],
    },

    # ═══════════════════════════════════════════════════════════════
    # 案件6: 钱七 — 因工外出(第五项) — 单位申请
    # 同案两人: 本人(钱七) + 法人(王五)
    # ═══════════════════════════════════════════════════════════════
    {
        "name": "11/12 钱七案-本人(因工外出)",
        "role": "本人",
        "name_pane": "钱七",
        "idnumer_pane": "330324199906081239",
        "textEdit": "浙江省永嘉县岩头镇ZZ村88号",
        "lineEdit_4": "13544440005",
        "lineEdit_5": "技术员",
        "lineEdit_2": "钱七",
        "comboBox": 4,
        "company_pane": "永嘉县CC环保科技有限公司",
        "construction_company": "",
        "construction_plant": "",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "申请人钱七，系永嘉县CC环保科技有限公司技术员。2026年7月28日，受公司经理王五指派前往乐清市DD化工厂进行设备维护服务。上午9时30分许，乘坐的公司车辆在乐清市柳市镇境内发生侧翻事故。事故造成申请人腰椎压缩性骨折，已送乐清市人民医院治疗。该公司车辆由公司安排，出差事项有公司派工单和出差申请记录。",
        "materials": [
            {"name": "身份证复印件", "provided": True, "notes": ""},
            {"name": "公司派工单", "provided": True, "notes": "2026年7月27日签发"},
            {"name": "出差申请书", "provided": True, "notes": "经理王五审批"},
            {"name": "交通事故认定书", "provided": True, "notes": "单方事故 路面湿滑"},
            {"name": "医院诊断证明", "provided": True, "notes": "腰椎压缩性骨折"},
            {"name": "劳动合同", "provided": False, "notes": ""},
        ],
    },
    {
        "name": "12/12 钱七案-法人谈话(王五)",
        "role": "法人",
        "name_pane": "王五",
        "idnumer_pane": "330324198212011234",
        "textEdit": "浙江省永嘉县上塘镇XX路66号",
        "lineEdit_4": "13777770003",
        "lineEdit_5": "经理",
        "lineEdit_2": "钱七",                         # ← 指向同案受伤职工
        "comboBox": 4,
        "company_pane": "永嘉县CC环保科技有限公司",
        "construction_company": "",
        "construction_plant": "",
        "deathCaseCheckbox": False,
        "personalApplicationCheckbox": False,
        "statement_edit": "永嘉县CC环保科技有限公司经理王五确认：钱七系我公司正式技术员，负责对外设备维护服务。2026年7月28日的出差系我本人亲自指派，有派工单和出差审批记录。当日公司安排司机陈某驾驶公司车辆（浙C·XXXXX）送钱七前往乐清市DD化工厂。途中发生的事故由交警认定系路面湿滑导致车辆失控侧翻，系单方事故。公司认可钱七该事故属于因工外出期间受伤。",
        "materials": [
            {"name": "公司营业执照副本", "provided": True, "notes": ""},
            {"name": "经理身份证", "provided": True, "notes": ""},
            {"name": "派工单", "provided": True, "notes": "2026年7月27日"},
            {"name": "出差审批单", "provided": True, "notes": ""},
            {"name": "公司车辆行驶证", "provided": True, "notes": "浙C·XXXXX"},
            {"name": "劳动合同", "provided": True, "notes": ""},
        ],
    },
]


# ============================================================================
# 关键证据定义（缺失时AI必须在笔录中追问）
# ============================================================================
KEY_EVIDENCE_NAMES = [
    "身份证",           # 身份证明
    "劳动合同",         # 劳动关系证明
    "医院诊断证明",     # 受伤事实证明
    "工资发放记录",     # 工资标准证明
    "考勤记录",         # 工作时间证明
    "证人证言",         # 事故见证
    "事故现场照片",     # 现场证据
    "监控录像",         # 现场证据
    "工伤认定书",       # 前置认定
    "劳动关系裁决书",   # 劳动关系确认
    "道路交通事故认定书",  # 上下班途中/因工外出
    "公安报案回执",     # 暴力伤害
    "死亡证明",         # 工亡
]


# ============================================================================
# MaterialListWidget — 可勾选+可备注的材料列表组件
# ============================================================================

class MaterialListWidget(QWidget):
    """替代原有 QTextEdit 的材料管理组件。

    每行: [☑/☐ 复选框] [材料名称] [备注输入框]
    - 勾选 = 已提供
    - 未勾选 = 缺失，生成笔录时AI会追问
    - 备注 = 对该材料的补充说明
    """

    materials_changed = pyqtSignal()  # 材料变更信号

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: List[Dict[str, Any]] = []  # [{name, provided, notes, widget_refs}]
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(1)

        # 滚动区域
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ccc;
                border-radius: 2px;
                background-color: #fafafa;
            }
        """)

        # 内容容器
        self._container = QWidget()
        self._container.setStyleSheet("background-color: #fafafa;")
        self._row_layout = QVBoxLayout(self._container)
        self._row_layout.setContentsMargins(4, 2, 4, 2)
        self._row_layout.setSpacing(2)
        self._row_layout.addStretch()  # 底部弹簧，把行推到顶部

        self.scroll.setWidget(self._container)
        layout.addWidget(self.scroll)

    def _make_row(self, name: str = "", provided: bool = False, notes: str = ""):
        """创建一行材料条目"""
        row = QWidget()
        row.setFixedHeight(23)
        h = QHBoxLayout(row)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(3)

        # 复选框
        cb = QCheckBox()
        cb.setChecked(provided)
        cb.setFixedWidth(18)
        cb.setToolTip("勾选=已提供  |  不勾选=缺失")
        cb.toggled.connect(self._on_changed)

        # 材料名称输入框（可编辑）
        name_edit = QLineEdit(name if name else "")
        name_edit.setPlaceholderText("材料名称...")
        name_edit.setStyleSheet("""
            QLineEdit {
                font-size: 8pt;
                border: 1px solid #ddd;
                border-radius: 1px;
                padding: 1px 3px;
                background-color: #fff;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
        """)
        name_edit.setMinimumWidth(100)
        name_edit.textChanged.connect(self._on_changed)

        # 备注输入框
        note_edit = QLineEdit()
        note_edit.setText(notes)
        note_edit.setPlaceholderText("备注...")
        note_edit.setStyleSheet(name_edit.styleSheet())
        note_edit.textChanged.connect(self._on_changed)

        h.addWidget(cb)
        h.addWidget(name_edit, 1)   # stretch=1，自动填充剩余空间
        h.addWidget(note_edit, 2)   # stretch=2，备注更宽

        return row, cb, name_edit, note_edit

    def add_row(self, name: str = "", provided: bool = False, notes: str = ""):
        """在末尾添加一行"""
        row, cb, name_edit, note = self._make_row(name, provided, notes)
        # 在 stretch 之前插入
        self._row_layout.insertWidget(self._row_layout.count() - 1, row)

        item = {
            "name": name,
            "provided": provided,
            "notes": notes,
            "_cb": cb,
            "_name_edit": name_edit,
            "_note": note,
            "_row": row,
        }
        self._rows.append(item)

    def set_materials(self, data: List[Dict[str, Any]]):
        """批量设置材料列表"""
        self.clear()
        for item in data:
            self.add_row(
                name=item.get("name", ""),
                provided=item.get("provided", False),
                notes=item.get("notes", "")
            )

    def get_materials(self) -> List[Dict[str, Any]]:
        """获取所有材料数据（同步UI状态）"""
        result = []
        for row in self._rows:
            result.append({
                "name": row["_name_edit"].text(),
                "provided": row["_cb"].isChecked(),
                "notes": row["_note"].text(),
            })
        return result

    def get_provided(self) -> List[str]:
        """获取已提供的材料名称列表"""
        return [r["_name_edit"].text() for r in self._rows if r["_cb"].isChecked()]

    def get_missing(self) -> List[str]:
        """获取缺失的材料名称列表"""
        return [r["_name_edit"].text() for r in self._rows if not r["_cb"].isChecked()]

    def get_missing_key_evidence(self) -> List[str]:
        """获取缺失的关键证据列表"""
        missing = self.get_missing()
        return [m for m in missing if any(
            kw in m for kw in KEY_EVIDENCE_NAMES
        )]

    def clear(self):
        """清空所有行"""
        for row in self._rows:
            row["_row"].setParent(None)
        self._rows.clear()

    def get_summary_text(self) -> str:
        """生成可复制的文本摘要"""
        lines = []
        for i, r in enumerate(self.get_materials(), 1):
            status = "✓" if r["provided"] else "✗"
            line = f"{status} {i}. {r['name']}"
            if r["notes"]:
                line += f"（{r['notes']}）"
            lines.append(line)
        return "\n".join(lines)

    def copy_to_clipboard(self):
        """复制材料摘要到剪贴板"""
        text = self.get_summary_text()
        QApplication.clipboard().setText(text)

    def _on_changed(self):
        """复选框或备注变更时发出信号"""
        self.materials_changed.emit()


class MainWindow(QWidget, Ui_Form):

    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.setupUi(self)
        self._setup_radio_connections()

        self.person_transcript_created = False
        self._test_data_index = -1  # F2 测试数据轮换索引

        # 创建用户管理器并整合API配置UI
        self.user_manager = UserManager()
        self._setup_api_config_ui()
        self._load_saved_user_config()

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
        self.data_model.output_config['用户名'] = self._get_current_username()

        self.case_index_manager = CaseIndexManager()

        # 简化日志系统
        self.setup_logging()

        # 保持向后兼容
        self._template_dict = self.data_model.to_template_dict()
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
        self.construction_plant.currentTextChanged.connect(self.c_plant)

        # 初始化案件类型下拉框
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第一项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第二项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第三项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第四项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第五项")
        self.comboBox.addItem("《工伤保险条例》第十四条第一款第六项")

        # 初始化组合框（必须在 init_combobox_data 之后）
        self.init_comboboxes()

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

        try:
            self.pushButton.clicked.disconnect()
        except TypeError:
            pass
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
            self.work_talk()

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
            current_date = _date_now()

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
            current_date = _date_now()
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
            current_date = _date_now()

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
                '当前时间': _time_now(),
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

    def _setup_api_config_ui(self):
        """在窗口顶部创建折叠配置栏，并下移所有现有控件；右侧增加辅助面板"""
        TOP_BAR_H = 28          # 顶部小横条高度
        RIGHT_PANEL_X = 478     # 右侧面板起始 x
        RIGHT_PANEL_W = 382     # 右侧面板宽度

        # --- 1. 调整窗口大小（左侧 470 + 右侧 400 = 870） ---
        WIN_W, WIN_H = 870, 740
        self.setMinimumSize(WIN_W, WIN_H)
        self.setMaximumSize(WIN_W, WIN_H)
        self.resize(WIN_W, WIN_H)

        # --- 2. 下移所有现有控件（只挪一个薄顶栏的空间） ---
        for child in self.children():
            if isinstance(child, QWidget) and child is not self:
                try:
                    geo = child.geometry()
                    child.setGeometry(geo.x(), geo.y() + TOP_BAR_H,
                                      geo.width(), geo.height())
                except Exception:
                    pass

        # ============================================================
        # 3. 顶部小横条（始终可见）：⚙ 配置 + 状态
        # ============================================================
        self.top_bar = QLabel(self)
        self.top_bar.setGeometry(0, 0, WIN_W, TOP_BAR_H)
        self.top_bar.setStyleSheet(
            "background-color: #e8e8e8; border-bottom: 1px solid #ccc;"
        )

        # 齿轮按钮：展开/收起用户配置
        self.config_toggle_btn = QPushButton("⚙", self)
        self.config_toggle_btn.setGeometry(4, 2, 28, 24)
        self.config_toggle_btn.setToolTip("显示/隐藏用户配置")
        self.config_toggle_btn.setStyleSheet(
            "QPushButton { background: transparent; border: none; font-size: 14px; }"
            "QPushButton:hover { background-color: #d0d0d0; border-radius: 3px; }"
        )
        self.config_toggle_btn.clicked.connect(self._toggle_config_panel)

        # 顶部状态文字
        self.top_status_label = QLabel("", self)
        self.top_status_label.setGeometry(37, 4, 430, 20)
        self.top_status_label.setStyleSheet("color: #888; background: transparent; border: none;")

        # ============================================================
        # 4. 可折叠的用户配置面板（默认隐藏）
        # ============================================================
        self.api_group = QGroupBox("用户配置", self)
        self.api_group.setGeometry(5, TOP_BAR_H + 2, 460, 84)
        self.api_group.setFont(QFont("微软雅黑", 9))
        self.api_group.hide()  # 默认隐藏

        # 第一行：用户名 + API密钥
        lbl_user = QLabel("用户:", self.api_group)
        lbl_user.setGeometry(10, 22, 35, 20)

        self.api_user_combo = QComboBox(self.api_group)
        self.api_user_combo.setEditable(True)
        self.api_user_combo.setGeometry(45, 20, 140, 22)
        self.api_user_combo.setPlaceholderText("输入用户名")
        self.api_user_combo.currentTextChanged.connect(self._on_user_combo_changed)

        lbl_key = QLabel("密钥:", self.api_group)
        lbl_key.setGeometry(195, 22, 35, 20)

        self.api_key_input = PasswordLineEdit(self.api_group)
        self.api_key_input.setGeometry(230, 20, 160, 22)
        self.api_key_input.setPlaceholderText("输入API密钥")

        # 第二行：记住我 + 保存 + 状态
        self.api_remember_cb = QCheckBox("记住我", self.api_group)
        self.api_remember_cb.setGeometry(10, 50, 70, 20)
        self.api_remember_cb.setChecked(True)

        self.api_save_btn = QPushButton("保存配置", self.api_group)
        self.api_save_btn.setGeometry(80, 48, 70, 23)
        self.api_save_btn.clicked.connect(self._on_save_api_config)

        self.api_status_label = QLabel("", self.api_group)
        self.api_status_label.setGeometry(160, 50, 290, 20)
        self.api_status_label.setStyleSheet("color: #888;")

        # ============================================================
        # 5. 右侧面板：案件申请陈述（上）
        # ============================================================
        STATEMENT_H = 350
        self.statement_group = QGroupBox("案件申请陈述", self)
        self.statement_group.setGeometry(RIGHT_PANEL_X, TOP_BAR_H + 3,
                                         RIGHT_PANEL_W, STATEMENT_H)
        self.statement_group.setFont(QFont("微软雅黑", 9))

        self.statement_edit = QTextEdit(self.statement_group)
        self.statement_edit.setGeometry(8, 18, RIGHT_PANEL_W - 16, STATEMENT_H - 50)
        self.statement_edit.setPlaceholderText("在此输入案件申请陈述...")
        self.statement_edit.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ccc;
                border-radius: 2px;
                background-color: #fafafa;
                font-size: 9pt;
            }
            QTextEdit:focus {
                border-color: #3498db;
                background-color: #fff;
            }
        """)

        # 按钮：手动定位，不用 layout（避免和 QGroupBox 冲突）
        btn_y = STATEMENT_H - 28
        stmt_copy_btn = QPushButton("复制", self.statement_group)
        stmt_copy_btn.setGeometry(8, btn_y, 45, 23)
        stmt_copy_btn.clicked.connect(lambda: self._copy_statement())

        stmt_clear_btn = QPushButton("清空", self.statement_group)
        stmt_clear_btn.setGeometry(58, btn_y, 45, 23)
        stmt_clear_btn.clicked.connect(lambda: self.statement_edit.clear())

        # ============================================================
        # 6. 右侧面板：目前提供的材料分类（下）
        # ============================================================
        MATERIAL_H = 235
        material_top = TOP_BAR_H + STATEMENT_H + 10
        self.material_group = QGroupBox("目前提供的材料分类", self)
        self.material_group.setGeometry(RIGHT_PANEL_X, material_top,
                                        RIGHT_PANEL_W, MATERIAL_H)
        self.material_group.setFont(QFont("微软雅黑", 9))

        # 用 MaterialListWidget 替换原来的 QTextEdit
        self.material_list = MaterialListWidget(self.material_group)
        self.material_list.setGeometry(8, 18, RIGHT_PANEL_W - 16, MATERIAL_H - 50)

        # 按钮：手动定位
        mat_btn_y = MATERIAL_H - 28
        mat_copy_btn = QPushButton("复制", self.material_group)
        mat_copy_btn.setGeometry(8, mat_btn_y, 45, 23)
        mat_copy_btn.clicked.connect(lambda: self._copy_material())

        mat_clear_btn = QPushButton("清空", self.material_group)
        mat_clear_btn.setGeometry(58, mat_btn_y, 45, 23)
        mat_clear_btn.clicked.connect(lambda: self.material_list.clear())

        mat_add_btn = QPushButton("新增", self.material_group)
        mat_add_btn.setGeometry(108, mat_btn_y, 45, 23)
        mat_add_btn.clicked.connect(lambda: self.material_list.add_row())

        print("✅ API配置UI已创建")

    def _load_saved_user_config(self):
        """加载已保存的用户配置到UI"""
        remembered = self.user_manager.get_remembered_user()
        if remembered:
            username = remembered.get("username", "")
            api_key = remembered.get("api_key", "")
            # 填充用户下拉列表
            user_list = self.user_manager.get_user_list()
            self.api_user_combo.addItems(user_list)
            if username:
                idx = self.api_user_combo.findText(username)
                if idx >= 0:
                    self.api_user_combo.setCurrentIndex(idx)
                else:
                    self.api_user_combo.setCurrentText(username)
            if api_key:
                self.api_key_input.setText(api_key)
            print(f"✅ 已加载记住的用户: {username}")
        else:
            # 至少填充用户列表
            user_list = self.user_manager.get_user_list()
            self.api_user_combo.addItems(user_list)
            print("ℹ️ 没有记住的用户")

    def _get_current_username(self) -> str:
        """获取当前用户名"""
        if hasattr(self, 'api_user_combo'):
            return self.api_user_combo.currentText().strip()
        return "未登录用户"

    def _toggle_config_panel(self):
        """展开/收起用户配置面板"""
        visible = self.api_group.isVisible()
        self.api_group.setVisible(not visible)
        arrow = "▼" if not visible else "⚙"
        self.config_toggle_btn.setText(arrow)

    def _update_api_status(self):
        """更新API状态标签"""
        if not hasattr(self, 'api_status_label'):
            return
        if self.ai_service:
            self.api_status_label.setText("✅ AI已就绪")
            self.api_status_label.setStyleSheet("color: green; font-weight: bold;")
            self.top_status_label.setText("✅ AI已就绪")
            self.top_status_label.setStyleSheet("color: green; background: transparent; border: none;")
        else:
            username = self._get_current_username()
            if not username or username == "未登录用户":
                msg = "⚠️ 请配置用户名和API密钥"
                self.api_status_label.setText(msg)
                self.api_status_label.setStyleSheet("color: orange;")
                self.top_status_label.setText(msg)
                self.top_status_label.setStyleSheet("color: orange; background: transparent; border: none;")
            else:
                msg = "⚠️ API密钥未配置，AI功能不可用"
                self.api_status_label.setText(msg)
                self.api_status_label.setStyleSheet("color: orange;")
                self.top_status_label.setText(msg)
                self.top_status_label.setStyleSheet("color: orange; background: transparent; border: none;")

    def _on_save_api_config(self):
        """保存API配置"""
        username = self.api_user_combo.currentText().strip()
        api_key = self.api_key_input.text().strip()
        remember = self.api_remember_cb.isChecked()

        if not username:
            QMessageBox.warning(self, "提示", "请输入用户名")
            return

        # 保存到UserManager
        api_url = "https://api.deepseek.com"
        success = self.user_manager.save_user_config(
            username=username,
            api_url=api_url,
            api_key=api_key,
            remember_me=remember,
            service="DeepSeek"
        )

        if success:
            # 更新下拉列表
            if self.api_user_combo.findText(username) < 0:
                self.api_user_combo.addItem(username)
                self.api_user_combo.setCurrentText(username)

            # 更新数据模型
            self.data_model.output_config['用户名'] = username

            # 重新初始化AI服务
            self.init_ai_service()

            self._set_status(f"配置已保存 - 用户: {username}", "green")
            print(f"✅ API配置已保存: 用户={username}, 密钥={'已设置' if api_key else '未设置'}")
        else:
            QMessageBox.critical(self, "错误", "保存配置失败，请重试")

    def _on_user_combo_changed(self, text):
        """用户名下拉框变化时自动加载对应的API密钥"""
        if not text or not text.strip():
            return
        username = text.strip()
        user_data = self.user_manager.users_data.get('users', {}).get(username, {})
        if user_data:
            api_key = user_data.get('api_key', '')
            remember = user_data.get('remember_me', True)
            self.api_key_input.setText(api_key)
            self.api_remember_cb.setChecked(remember)
            if api_key:
                print(f"✅ 已加载用户 '{username}' 的API配置")
            else:
                print(f"ℹ️ 用户 '{username}' 未配置API密钥")

    def init_ai_service(self):
        """初始化AI服务 - 从UserManager读取配置"""
        try:
            # 从UserManager获取当前用户的API配置
            username = self._get_current_username()
            if not username:
                print("⚠️ 未找到用户配置，AI功能将不可用")
                self.ai_service = None
                self._update_api_status()
                return

            user_data = self.user_manager.users_data.get('users', {}).get(username, {})
            api_key = user_data.get('api_key', '')
            api_url = user_data.get('api_url', 'https://api.deepseek.com')

            # 检查配置是否完整
            if not api_key or not api_url:
                print("⚠️ API配置不完整，AI功能将不可用")
                print(f"  API地址: {api_url if api_url else '未设置'}")
                print(f"  API密钥: {'已设置' if api_key else '未设置'}")
                self.ai_service = None
                self._update_api_status()
                return

            print(f"✅ 使用用户 '{username}' 的API配置初始化AI服务")
            print(f"  API地址: {api_url}")
            print(f"  API密钥前8位: {api_key[:8]}...")

            # 创建AI服务实例
            self.ai_service = AIService(api_key, api_url)
            print("✅ AI服务初始化成功")
            self._update_api_status()

        except Exception as e:
            print(f"❌ AI服务初始化失败: {e}")
            import traceback
            traceback.print_exc()
            self.ai_service = None
            self._update_api_status()

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
        QMessageBox.information(parent or self, "复制成功", "已复制到剪贴板")

    def _copy_statement(self):
        """复制案件申请陈述"""
        text = self.statement_edit.toPlainText().strip()
        if text:
            self.copy_to_clipboard(text)
        else:
            QMessageBox.information(self, "提示", "案件申请陈述为空")

    def _copy_material(self):
        """复制材料分类"""
        if hasattr(self, 'material_list'):
            self.material_list.copy_to_clipboard()
        QMessageBox.information(self, "提示", "材料列表已复制到剪贴板")

    def save_ai_report(self, parsed_result):
        """保存AI审查报告"""
        if not self.current_case_folder:
            QMessageBox.warning(self, "提示", "请先保存案件信息")
            return

        person_name = self.get_data('本人姓名', '未知')
        timestamp = _timestamp_now()
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
            current_date = _date_now()
            current_time = _time_now()

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

    def _apply_ui_settings(self):
        """应用UI设置"""
        try:
            ui_settings = self.config_service.get_ui_settings()

            # 设置字体
            font = QFont(ui_settings.font_family, ui_settings.font_size)
            self.setFont(font)

            # 注意：窗口大小由 _setup_api_config_ui 固定为 870x740，此处不再覆盖
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
        new_item = self.construction_plant.currentText().strip()
        if new_item and new_item not in self.items_list2:
            self.items_list2 = self.file_service.save_to_excel(
                "",  # 空字符串
                '工地名称汇总.xlsx',
                '工地名称汇总',
                new_item,
                self.items_list2
            )
            self.init_combobox(self.construction_plant, self.items_list2)
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
            if key not in self._template_dict or self._template_dict[key] != model_value:
                self._template_dict[key] = model_value
            return model_value

        dict_value = self._template_dict.get(key)
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
        self._template_dict[key] = value
        self._handle_special_keys(key, value)

    def _handle_special_keys(self, key: str, value: Any):
        """处理特殊键的同步逻辑"""
        if key == "本人姓名":
            self.lineEdit_2.setText(str(value))
        elif key == "当前时期":
            if not value:
                current_date = _date_now()
                current_time = _time_now()
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
        """从数据模型的各个部分获取数据

        注意：_store_to_data_model 存储时保留完整 key（含角色前缀），
        此处 first check 即可命中，不需要再去前缀查找。
        """
        if key in self.data_model.basic_info:
            return self.data_model.basic_info[key]

        for category in (
            self.data_model.company_info,
            self.data_model.case_info,
            self.data_model.investigation,
            self.data_model.output_config,
        ):
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

        # 清空右侧辅助面板
        if hasattr(self, 'statement_edit'):
            self.statement_edit.clear()
        if hasattr(self, 'material_list'):
            self.material_list.clear()

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
            current_date = _date_now()
            current_time = _time_now()
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

            self._template_dict.update(all_data)

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
            self._template_dict.clear()
            self._template_dict.update(all_variables)

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

                # 如果有 AI 生成的问答内容，插入到文档中
                if hasattr(self, '_ai_transcript_content') and self._ai_transcript_content:
                    self._insert_ai_content_into_doc(target_path)
                    self._ai_transcript_content = ""  # 用完即清

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
        site = self.construction_plant.currentText().strip()

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
                    current_date = variables.get('当前日期', _date_now())
                    current_time = variables.get('当前时间', _time_now())
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
        self.init_combobox(self.construction_plant, self.items_list2)

        self.company_pane.setCurrentIndex(-1)
        self.construction_company.setCurrentIndex(-1)
        self.construction_plant.setCurrentIndex(-1)

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
        site_name = self.construction_plant.currentText().strip()
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
            self.set_data('当前时期', _date_now() + _time_now(), 'output')

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

    def work_talk(self):
        """根据当前角色更新信息并保存案件"""
        role = self.get_current_role_type()
        self.update_role_info(role)

        if self.ai_service:
            # AI 增强模式
            self._transcript_with_ai(role)
        else:
            # 传统模式：直接走模板
            self.discriminate()
            self.work_save()

    def _transcript_with_ai(self, role: str):
        """使用 AI 生成谈话笔录问答内容"""
        try:
            # 1. 案件分类
            from case_classifier import CaseClassifier
            classifier = CaseClassifier()
            c = classifier.classify(self)
            print(classifier.to_summary(c))

            # 2. 构建 Prompt
            from transcript_prompt import (
                build_system_prompt, build_user_prompt, OUTPUT_FORMAT_INSTRUCTION
            )
            system_prompt = build_system_prompt(c.regulation_index, role)

            # 案件陈述和材料清单（所有角色都传入）
            case_statement = ""
            material_summary = ""
            if hasattr(self, 'statement_edit'):
                case_statement = self.statement_edit.toPlainText().strip()
            if hasattr(self, 'material_list'):
                material_summary = self.material_list.get_summary_text()

            user_prompt = build_user_prompt(
                c, case_statement=case_statement, material_summary=material_summary
            ) + "\n" + OUTPUT_FORMAT_INSTRUCTION

            # ============================================================
            # [DEBUG] 临时调试对话框 — 显示将发送给 AI 的内容
            # ============================================================
            if not self._show_debug_dialog(classifier, c, system_prompt, user_prompt):
                self._set_status('已取消（调试预览）', 'black')
                return

            # 3. 状态提示
            self._set_status('正在AI生成谈话笔录...', 'black')
            QApplication.processEvents()

            # 4. 调用 AI
            result = self.ai_service.generate_transcript(system_prompt, user_prompt)

            if result.get("状态") == "成功":
                self._ai_transcript_content = result.get("内容", "")
                usage = result.get("用量", {})
                prompt_tokens = usage.get("prompt_tokens", 0)
                completion_tokens = usage.get("completion_tokens", 0)
                total_tokens = usage.get("total_tokens", 0)
                print(f"✅ AI 笔录生成完成: prompt={prompt_tokens}, completion={completion_tokens}, total={total_tokens}")
                self._set_status(
                    f'AI生成完成 (prompt:{prompt_tokens} completion:{completion_tokens})',
                    'green'
                )
            else:
                error_msg = result.get("错误信息", "未知错误")
                print(f"⚠️ AI 生成失败: {error_msg}")
                self._set_status(f'AI生成失败: {error_msg[:50]}，使用传统模板', 'orange')
                self._ai_transcript_content = ""

        except Exception as e:
            print(f"❌ AI 生成异常: {e}")
            import traceback
            traceback.print_exc()
            self._set_status('AI生成异常，使用传统模板', 'orange')
            self._ai_transcript_content = ""

        # 5. 继续保存流程（work_save 会检测并使用 AI 内容）
        self.discriminate()
        self.work_save()

    # ── 调试对话框（临时） ──────────────────────────────────────

    def _show_debug_dialog(self, classifier, c, system_prompt: str,
                           user_prompt: str) -> bool:
        """
        [临时调试] 在调用 AI 前展示分类结果 + System Prompt + User Prompt。

        Returns:
            True  = 用户点击"继续发送AI"
            False = 用户点击"取消"
        """
        from case_classifier import CaseClassifier

        dlg = QDialog(self)
        dlg.setWindowTitle("🔍 AI Prompt 预览（调试）")
        dlg.resize(950, 750)
        dlg.setMinimumSize(800, 600)

        layout = QVBoxLayout(dlg)

        # ── 标题 ──
        title = QLabel("📋 以下是将发送给 DeepSeek 的内容")
        title.setStyleSheet("font-size: 14px; font-weight: bold; padding: 4px;")
        layout.addWidget(title)

        # ── 页签容器 ──
        tabs = QTabWidget()

        # --- Tab 1: 案件分类摘要 ---
        tab1 = QWidget()
        t1 = QVBoxLayout(tab1)
        summary = QTextEdit()
        summary.setReadOnly(True)
        summary.setFont(QFont("Consolas", 9))
        summary.setPlainText(classifier.to_summary(c))
        t1.addWidget(summary)
        tabs.addTab(tab1, "分类结果")

        # --- Tab 2: System Prompt ---
        tab2 = QWidget()
        t2 = QVBoxLayout(tab2)
        sp = QTextEdit()
        sp.setReadOnly(True)
        sp.setFont(QFont("Consolas", 8))
        sp.setPlainText(system_prompt)
        t2.addWidget(sp)
        stats2 = QLabel(f"字符数: {len(system_prompt)}  |  约 {len(system_prompt)//2} Token（可缓存）")
        stats2.setStyleSheet("color: #888; padding: 2px;")
        t2.addWidget(stats2)
        tabs.addTab(tab2, "System Prompt")

        # --- Tab 3: User Prompt ---
        tab3 = QWidget()
        t3 = QVBoxLayout(tab3)
        up = QTextEdit()
        up.setReadOnly(True)
        up.setFont(QFont("Consolas", 9))
        up.setPlainText(user_prompt)
        t3.addWidget(up)
        stats3 = QLabel(f"字符数: {len(user_prompt)}  |  约 {len(user_prompt)//2} Token")
        stats3.setStyleSheet("color: #888; padding: 2px;")
        t3.addWidget(stats3)
        tabs.addTab(tab3, "User Prompt")

        # --- Tab 4: 合并（发送内容） ---
        tab4 = QWidget()
        t4 = QVBoxLayout(tab4)
        combined = QTextEdit()
        combined.setReadOnly(True)
        combined.setFont(QFont("Consolas", 8))
        combined.setPlainText(
            f"=== SYSTEM PROMPT ({len(system_prompt)} 字符) ===\n\n{system_prompt}\n\n"
            f"=== USER PROMPT ({len(user_prompt)} 字符) ===\n\n{user_prompt}"
        )
        t4.addWidget(combined)
        stats4 = QLabel(
            f"System: {len(system_prompt)} 字符 (~{len(system_prompt)//2} Token，可缓存)  |  "
            f"User: {len(user_prompt)} 字符 (~{len(user_prompt)//2} Token)  |  "
            f"实际消耗约: ~{len(user_prompt)//2} Token"
        )
        stats4.setStyleSheet("color: #888; padding: 2px;")
        t4.addWidget(stats4)
        tabs.addTab(tab4, "合并预览")

        layout.addWidget(tabs)

        # ── 底部按钮 ──
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(dlg.reject)
        btn_layout.addWidget(cancel_btn)

        send_btn = QPushButton("发送给 AI →")
        send_btn.setStyleSheet(
            "QPushButton { background-color: #3498db; color: white; font-weight: bold; "
            "padding: 6px 20px; border-radius: 4px; }"
            "QPushButton:hover { background-color: #2980b9; }"
        )
        send_btn.clicked.connect(dlg.accept)
        btn_layout.addWidget(send_btn)

        layout.addLayout(btn_layout)

        return dlg.exec_() == QDialog.Accepted

    def _insert_ai_content_into_doc(self, file_path: str):
        """
        将 AI 生成的问答内容插入到已渲染的文档中。

        策略：找到文档中第一个"问："段落（Q&A 正文起点的标记），
        删除从该位置到文档末尾的所有内容，然后插入 AI 生成的问答。

        Args:
            file_path: 已渲染的 docx 文件路径
        """
        try:
            from docx import Document
            from docx.oxml.ns import qn
            from docx.shared import Pt

            ai_text = self._ai_transcript_content
            if not ai_text:
                return

            doc = Document(file_path)
            body = doc.element.body

            # ── 找到第一个"问："段落，并收集要删除的元素 ──
            elems_to_remove = []
            found_qa = False
            for child in list(body):
                if child.tag == qn('w:p'):
                    # 提取段落文本
                    texts = []
                    for t in child.iter(qn('w:t')):
                        if t.text:
                            texts.append(t.text)
                    para_text = ''.join(texts).strip()

                    if not found_qa and para_text.startswith('问：'):
                        found_qa = True
                    if found_qa:
                        elems_to_remove.append(child)

            # ── 删除旧的 Q&A 段落 ──
            for elem in elems_to_remove:
                body.remove(elem)

            print(f"🗑️ 已移除 {len(elems_to_remove)} 个旧 Q&A 段落")

            # ── 插入 AI 生成的问答内容 ──
            lines = ai_text.strip().split('\n')
            inserted_count = 0
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                p = doc.add_paragraph()

                if line.startswith('问：'):
                    # 问题是普通格式
                    run = p.add_run(line)
                    run.font.size = Pt(11)
                elif line.startswith('答：'):
                    # 回答需要下划线（笔录规范）
                    run = p.add_run(line)
                    run.font.size = Pt(11)
                    run.underline = True
                else:
                    # 其他行（可能是续行）
                    run = p.add_run(line)
                    run.font.size = Pt(11)

                inserted_count += 1

            doc.save(file_path)
            print(f"✅ 已插入 {inserted_count} 个 AI 问答段落 → {file_path}")

        except Exception as e:
            print(f"❌ 插入 AI 内容失败: {e}")
            import traceback
            traceback.print_exc()

    def discriminate(self):
        """生成模板文件路径"""
        number = self.comboBox.currentIndex()
        role = self.get_current_role_type()
        has_construction_company = bool(self.construction_company.currentText())
        has_construction_plant = bool(self.construction_plant.currentText())

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

    # ========================================================================
    # F2 测试数据轮换
    # ========================================================================

    def keyPressEvent(self, event):
        """F2 键轮换测试数据"""
        if event.key() == Qt.Key_F2:
            self._cycle_test_data()
        else:
            super().keyPressEvent(event)

    def _cycle_test_data(self):
        """按 F2 切换到下一组测试数据"""
        self._test_data_index = (self._test_data_index + 1) % len(TEST_DATA_PRESETS)
        data = TEST_DATA_PRESETS[self._test_data_index]

        print(f"\n{'='*50}")
        print(f"F2 测试数据: {data['name']}")
        print(f"{'='*50}")

        # ── 角色单选按钮 ──
        role_map = {"本人": self.radioButton, "证人": self.radioButton_2, "法人": self.radioButton_3}
        for role_name, btn in role_map.items():
            btn.setChecked(role_name == data["role"])
        # 触发角色切换
        self.clear_role_fields()

        # ── 案例类型复选框 ──
        self.death_case_checkbox.setChecked(data["deathCaseCheckbox"])
        self.personal_application_checkbox.setChecked(data["personalApplicationCheckbox"])
        self.on_case_type_changed()

        # ── 基本信息输入框 ──
        self.name_pane.setText(data["name_pane"])
        self.idnumer_pane.setText(data["idnumer_pane"])
        self.textEdit.setPlainText(data["textEdit"])
        self.lineEdit_4.setText(data["lineEdit_4"])
        self.lineEdit_5.setText(data["lineEdit_5"])
        self.lineEdit_2.setText(data["lineEdit_2"])

        # ── 条例下拉框 ──
        self.comboBox.setCurrentIndex(data["comboBox"])

        # ── 公司下拉框 ──
        self._set_combo_or_type(self.company_pane, data["company_pane"])
        self._set_combo_or_type(self.construction_company, data["construction_company"])
        self._set_combo_or_type(self.construction_plant, data["construction_plant"])

        # ── 自动计算年龄和性别 ──
        role = self.get_current_role_type()
        self.on_id_input_finished()

        # ── 右侧面板（同案沿用） ──
        # 同一受伤职工 → 沿用上一组的案件陈述和材料分类
        # 不同受伤职工 → 用新预设数据覆盖
        current_worker = data.get("lineEdit_2", "")
        prev_worker = getattr(self, '_prev_injured_worker', None)
        is_same_case = (prev_worker is not None and current_worker == prev_worker)

        if hasattr(self, 'statement_edit'):
            if is_same_case:
                # 同案：保持案件陈述不变，只检查是否为空
                if not self.statement_edit.toPlainText().strip() and data.get("statement_edit"):
                    self.statement_edit.setPlainText(data.get("statement_edit", ""))
                print(f"📋 同案沿用案件陈述（{current_worker}）")
            else:
                self.statement_edit.setPlainText(data.get("statement_edit", ""))

        if hasattr(self, 'material_list'):
            if is_same_case:
                # 同案：保持材料分类不变
                if self.material_list.get_materials() == [] and data.get("materials"):
                    self.material_list.set_materials(data.get("materials", []))
                print(f"📋 同案沿用材料分类（{current_worker}）")
            else:
                self.material_list.set_materials(data.get("materials", []))

        self._prev_injured_worker = current_worker

        # ── 重置案件状态，允许重新生成笔录 ──
        self.person_transcript_created = False
        self.pushButton.setEnabled(True)
        self.pushButton.setStyleSheet("")
        self.current_case_folder = None
        self.current_person_name = ""

        # ── 状态提示 ──
        label = (f"[测试 {self._test_data_index + 1}/{len(TEST_DATA_PRESETS)}] "
                 f"{data['name']}  |  F2=下一个")
        self._set_status(label, 'green')
        print(f"OK 测试数据已填充: {data['name']}")

    def _set_combo_or_type(self, combobox, text):
        """设置下拉框的值：如果存在则选中，否则直接输入"""
        if not text:
            combobox.setCurrentIndex(-1)
            return
        idx = combobox.findText(text)
        if idx >= 0:
            combobox.setCurrentIndex(idx)
        else:
            combobox.setCurrentIndex(-1)
            combobox.setEditText(text)


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())