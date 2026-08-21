# -*- coding: utf-8 -*-
"""
案件信息结构化分类模块
在调用 AI 之前，对 UI 收集的数据做本地分类，Zero Token 消耗。

分类维度：
  1. 案件基础分类 — 性质/申请主体/适用条例/角色
  2. 劳动关系要素 — 公司/用工形式/参保状态
  3. 事故要素     — 时间/地点/经过/医疗
  4. 已掌握证据   — 已有/缺失清单
"""

from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field


# ============================================================================
# 数据类定义
# ============================================================================

@dataclass
class PersonInfo:
    """人员信息"""
    name: str = ""
    gender: str = ""
    age: str = ""
    idcard: str = ""
    id_address: str = ""
    phone: str = ""
    position: str = ""  # 岗位（本人/证人）或职务（法人）


@dataclass
class CompanyInfo:
    """公司信息"""
    company_name: str = ""        # 用工单位
    employer: str = ""            # 用人单位
    site: str = ""                # 工地名称
    has_employer: bool = False    # 是否有用人单位
    has_site: bool = False        # 是否有明确工地


@dataclass
class AccidentInfo:
    """事故信息"""
    time: str = ""                # 事故发生时间
    location: str = ""            # 事故发生地点
    description: str = ""         # 受伤经过
    injury_location: str = ""     # 受伤部位
    diagnosis: str = ""           # 诊断结论
    hospital: str = ""            # 送医医院


@dataclass
class EvidenceStatus:
    """证据掌握状态"""
    has_idcard: bool = False
    has_contract: bool = False
    has_insurance: bool = False
    has_medical_record: bool = False
    has_witness: bool = False
    has_salary_record: bool = False
    has_attendance_record: bool = False
    has_work_badge: bool = False
    has_site_evidence: bool = False  # 现场照片/监控等


@dataclass
class CaseClassification:
    """完整的案件分类结果"""
    # 基础分类
    case_nature: str = ""          # 工伤 / 工亡
    applicant_type: str = ""       # 单位申请 / 个人申请
    regulation_index: int = 0      # 条例索引 0-5
    regulation_text: str = ""      # 条例原文摘要
    role: str = ""                 # 本人 / 证人 / 法人

    # 人员
    person: PersonInfo = field(default_factory=PersonInfo)
    injured_worker_name: str = ""  # 受伤职工姓名

    # 公司
    company: CompanyInfo = field(default_factory=CompanyInfo)

    # 事故
    accident: AccidentInfo = field(default_factory=AccidentInfo)

    # 证据
    evidence: EvidenceStatus = field(default_factory=EvidenceStatus)

    # 案件号
    case_number: str = ""


# ============================================================================
# 分类器
# ============================================================================

class CaseClassifier:
    """
    案件信息分类器

    使用方式:
        classifier = CaseClassifier()
        classification = classifier.classify(main_window)
    """

    # ── 条例索引 ──────────────────────────────────────────────
    REGULATIONS = {
        0: {
            "text": "《工伤保险条例》第十四条第一款第一项",
            "short": "第一项",
            "desc": "在工作时间和工作场所内，因工作原因受到事故伤害",
            "elements": ["劳动关系", "工作时间", "工作场所", "工作原因", "受伤事实"],
        },
        1: {
            "text": "《工伤保险条例》第十四条第一款第二项",
            "short": "第二项",
            "desc": "工作时间前后在工作场所内，从事与工作有关的预备性或收尾性工作受到事故伤害",
            "elements": ["工作内容", "预备/收尾行为", "事故时间合理性", "工作场所", "受伤事实"],
        },
        2: {
            "text": "《工伤保险条例》第十四条第一款第三项",
            "short": "第三项",
            "desc": "在工作时间和工作场所内，因履行工作职责受到暴力等意外伤害",
            "elements": ["工作职责", "冲突起因", "对方身份", "伤害行为", "与工作因果"],
        },
        3: {
            "text": "《工伤保险条例》第十四条第一款第四项",
            "short": "第四项",
            "desc": "患职业病",
            "elements": ["工作环境", "接触危害因素", "症状出现时间", "诊断结论", "因果关联"],
        },
        4: {
            "text": "《工伤保险条例》第十四条第一款第五项",
            "short": "第五项",
            "desc": "因工外出期间，由于工作原因受到伤害或发生事故下落不明",
            "elements": ["外出事由", "工作指派", "事故地点", "与工作关联"],
        },
        5: {
            "text": "《工伤保险条例》第十四条第一款第六项",
            "short": "第六项",
            "desc": "在上下班途中，受到非本人主要责任的交通事故或城市轨道交通、客运轮渡、火车事故伤害",
            "elements": ["住址→工作地路线", "合理时间", "非本人主要责任", "事故经过"],
        },
    }

    # ── 证据维度与对应的 data_model key ──────────────────────
    EVIDENCE_CHECKS = [
        ("has_idcard", "身份证号", "身份证信息"),
        ("has_contract", "劳动合同", "劳动合同签订情况"),
        ("has_insurance", "工伤保险", "工伤保险参保情况"),
        ("has_medical_record", "医疗记录", "医院诊断证明/病历"),
        ("has_witness", "证人", "目击证人证言"),
        ("has_salary_record", "工资记录", "工资发放记录/银行流水"),
        ("has_attendance_record", "考勤记录", "考勤记录/打卡记录"),
        ("has_work_badge", "工牌", "工作证/工牌/工作服"),
        ("has_site_evidence", "现场证据", "现场照片/监控录像"),
    ]

    def classify(self, main_window) -> CaseClassification:
        """
        从 MainWindow 实例提取并分类所有案件信息

        Args:
            main_window: MainWindow 实例

        Returns:
            CaseClassification: 结构化分类结果
        """
        c = CaseClassification()

        # 1. 基础分类
        self._classify_basic(c, main_window)

        # 2. 人员信息
        self._classify_person(c, main_window)

        # 3. 公司信息
        self._classify_company(c, main_window)

        # 4. 事故信息
        self._classify_accident(c, main_window)

        # 5. 证据状态
        self._classify_evidence(c, main_window)

        # 6. 案件号
        c.case_number = main_window.get_data("案本号", "")

        return c

    def _classify_basic(self, c: CaseClassification, w):
        """分类基础案件信息"""
        # 案件性质
        c.case_nature = "工亡" if w.death_case_checkbox.isChecked() else "工伤"
        # 申请类型
        c.applicant_type = "个人申请" if w.personal_application_checkbox.isChecked() else "单位申请"
        # 条例索引
        c.regulation_index = w.comboBox.currentIndex()
        # 条例文本
        reg = self.REGULATIONS.get(c.regulation_index, self.REGULATIONS[0])
        c.regulation_text = reg["text"]
        # 角色
        c.role = w.get_current_role_type()

    def _classify_person(self, c: CaseClassification, w):
        """分类人员信息"""
        role = c.role
        c.person.name = w.get_data(f"{role}姓名", "") or w.name_pane.text().strip()
        c.person.gender = w.get_data(f"{role}性别", "") or w.lineEdit.text().strip()
        c.person.age = w.get_data(f"{role}年龄", "") or w.age_pane.text().strip()
        c.person.idcard = w.get_data(f"{role}身份证号", "") or w.idnumer_pane.text().strip()
        c.person.id_address = w.get_data(f"{role}身份证地址", "") or w.textEdit.toPlainText().strip()
        c.person.phone = w.get_data(f"{role}手机号", "") or w.lineEdit_4.text().strip()
        if c.role == "法人":
            c.person.position = w.get_data("法人职务", "") or w.lineEdit_5.text().strip()
        else:
            c.person.position = w.get_data(f"{role}岗位", "") or w.lineEdit_5.text().strip()
        # 受伤职工 — 证人和法人不能用 name_pane 作 fallback（那是证人/法人自己的名字）
        c.injured_worker_name = w.get_data("本人姓名", "")
        if not c.injured_worker_name and c.role != "本人":
            # 尝试从案本号中提取（格式：张三-案本20260811001）
            case_num = w.lineEdit_2.text().strip()
            if case_num and '-' in case_num:
                c.injured_worker_name = case_num.split('-')[0]
        if not c.injured_worker_name:
            c.injured_worker_name = w.name_pane.text().strip()

    def _classify_company(self, c: CaseClassification, w):
        """分类公司信息"""
        info = w.get_company_info()
        c.company.company_name = info.get("用工单位", "")
        c.company.employer = info.get("用人单位", "")
        c.company.site = info.get("工地名称", "")
        c.company.has_employer = bool(c.company.employer)
        c.company.has_site = bool(c.company.site)

    def _classify_accident(self, c: CaseClassification, w):
        """分类事故信息"""
        inv = w.data_model.investigation
        c.accident.time = inv.get("事故发生时间", inv.get("事故时间", ""))
        c.accident.location = inv.get("事故发生地点", inv.get("事故地点", ""))
        c.accident.description = inv.get("受伤经过", "")
        c.accident.diagnosis = inv.get("医院诊断", inv.get("医疗结论", ""))
        c.accident.hospital = inv.get("送医医院", "")

    def _classify_evidence(self, c: CaseClassification, w):
        """分类证据掌握状态"""
        # 身份证：从 data_model 判断
        c.evidence.has_idcard = bool(
            w.get_data(f"{c.role}身份证号", "") and len(w.get_data(f"{c.role}身份证号", "")) >= 15
        )
        # 医疗记录：从 investigation 判断
        c.evidence.has_medical_record = bool(
            c.accident.diagnosis and c.accident.diagnosis not in ("", "待补充", "详见医疗诊断证明")
        )
        # 证人：角色本身是证人，或有证人笔录
        c.evidence.has_witness = (c.role == "证人")

        # 以下从 investigation 中尝试获取
        inv = w.data_model.investigation
        c.evidence.has_contract = bool(inv.get("劳动合同", ""))
        c.evidence.has_insurance = bool(inv.get("工伤保险", ""))
        c.evidence.has_salary_record = bool(inv.get("工资记录", ""))
        c.evidence.has_attendance_record = bool(inv.get("考勤记录", ""))
        c.evidence.has_work_badge = bool(inv.get("工牌", ""))
        c.evidence.has_site_evidence = bool(inv.get("现场证据", ""))

    # ── 辅助方法 ──────────────────────────────────────────────

    def get_evidence_map(self, c: CaseClassification) -> Dict[str, str]:
        """
        获取证据状态摘要

        Returns:
            {"身份证信息": "已有", "劳动合同签订情况": "缺失", ...}
        """
        result = {}
        for attr, _, label in self.EVIDENCE_CHECKS:
            status = "已有" if getattr(c.evidence, attr, False) else "缺失"
            result[label] = status
        return result

    def get_have_list(self, c: CaseClassification) -> List[str]:
        """获取已有证据的名称列表"""
        have = []
        for attr, _, label in self.EVIDENCE_CHECKS:
            if getattr(c.evidence, attr, False):
                have.append(label)
        return have

    def get_miss_list(self, c: CaseClassification) -> List[str]:
        """获取缺失证据的名称列表"""
        miss = []
        for attr, _, label in self.EVIDENCE_CHECKS:
            if not getattr(c.evidence, attr, False):
                miss.append(label)
        return miss

    def get_required_elements(self, c: CaseClassification) -> List[str]:
        """获取当前条例必须证明的法律要件"""
        reg = self.REGULATIONS.get(c.regulation_index, self.REGULATIONS[0])
        return reg["elements"]

    def to_summary(self, c: CaseClassification) -> str:
        """
        生成可读的分类摘要（用于调试/日志，非 AI prompt）

        Returns:
            多行文本摘要
        """
        lines = [
            "=" * 50,
            f"案件分类: {c.case_nature} | {c.applicant_type} | {c.regulation_text}",
            f"角色: {c.role} | 受伤职工: {c.injured_worker_name}",
            "-" * 50,
            f"[人员] {c.person.name}, {c.person.gender}, {c.person.age}岁",
            f"        身份证: {c.person.idcard[:6]}****({c.person.idcard[-4:]})" if c.person.idcard else "        身份证: 未填写",
            f"        地址: {c.person.id_address}",
            f"        电话: {c.person.phone}",
            f"        岗位/职务: {c.person.position}",
            "-" * 50,
            f"[公司] 用工单位: {c.company.company_name or '未填写'}",
            f"        用人单位: {c.company.employer or '无'}",
            f"        工地: {c.company.site or '无'}",
            "-" * 50,
            f"[事故] 时间: {c.accident.time or '未填写'}",
            f"        地点: {c.accident.location or '未填写'}",
            f"        经过: {(c.accident.description or '未填写')[:80]}...",
            f"        诊断: {c.accident.diagnosis or '未填写'}",
            f"        医院: {c.accident.hospital or '未填写'}",
            "-" * 50,
            f"[证据] 已有: {', '.join(self.get_have_list(c)) or '无'}",
            f"        缺失: {', '.join(self.get_miss_list(c))}",
            "-" * 50,
            f"[要件] 必须证明: {' → '.join(self.get_required_elements(c))}",
            f"[案号] {c.case_number or '未生成'}",
            "=" * 50,
        ]
        return "\n".join(lines)
