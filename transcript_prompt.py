# -*- coding: utf-8 -*-
"""
谈话笔录 AI 生成 — 提示词模块

设计原则：
  1. System Prompt 固化法律知识，可被 DeepSeek 缓存
  2. User Prompt 只传个案事实，紧凑结构化格式
  3. AI 根据角色自动生成合适的提问，不做硬编码模板
  4. 证人和法人笔录依赖本人笔录作为基础事实参考
"""

from typing import Dict, List
from case_classifier import CaseClassification


# ============================================================================
# System Prompt
# ============================================================================

def build_system_prompt(role: str) -> str:
    """构建 System Prompt —— 角色专属法律规范，不含硬编码提问模板。"""

    common_rules = """## 谈话笔录法律规范（通用）

### 笔录必需的结构要素
1. 告知程序：调查人员的身份告知、被调查人的权利义务告知
2. 被调查人基本信息：姓名、性别、年龄、身份证号、住址、联系方式、工作单位、岗位
3. 与案件关联的专门问询
4. 结尾确认：被调查人对笔录内容无异议的确认

### 笔录语言规范
- 使用"问："和"答："格式，一问一答
- 提问应具体、明确，避免诱导性提问
- 回答应使用第一人称记录
- 关键事实必须追问到位，不能笼统带过
- 涉及时间、地点、人物、经过等要素必须明确记录

### 证据链完整性要求
- 每个问题的目的都是为了证明《工伤保险条例》规定的法律要件
- 缺失的证明材料必须在提问中明确询问
- 对工伤认定有利的事实要重点展开

### 通用询问要点
- 劳动关系的确认（是否签合同、是否参保）
- 工作时间、工作地点、工作原因的核实
- 事故经过的详细还原（时间、地点、人物、过程、伤情）
- 医疗诊断情况（医院、诊断结论、治疗进展）
- 工资收入情况（月薪、发放方式）
- 各类证明材料的持有情况
"""

    role_templates = {
        "本人": """
## 本人笔录专属规范

你是受伤职工本人。AI 应以调查人员身份向其提问：
- 告知程序后，询问本人基本信息
- 围绕事故发生经过展开详细询问
- 核实劳动关系、工资、参保情况
- 询问医疗诊断和治疗情况
- 最终由本人确认笔录无误

生成要求：全部围绕受伤职工本人的亲身经历展开。
""",
        "证人": """
## 证人笔录专属规范

你是案件的目击证人或知情人。AI 应以调查人员身份向其提问：
- 告知程序后，询问证人基本信息
- 询问证人与受伤职工的关系（同事/工友/上下级/路人等）
- 询问证人对事故的目击情况：何时、何地、看到了什么
- 询问证人对受伤职工工作情况的了解
- 询问证人对事故原因的看法
- 最终由证人确认笔录无误

重要原则：
1. 证人只陈述自己看到、听到、知道的事实，不替受伤职工陈述
2. 证人不回答"我受伤时"之类的问题——受伤的是别人
3. 与本人笔录的一致性：参考本人笔录中的关键事实（时间、地点、经过），如有矛盾应标注
4. 本人笔录中的详情不需要证人逐条复述——证人只需补充自己知晓的部分
""",
        "法人": """
## 法人/负责人笔录专属规范

你是用人单位法定代表人、负责人或授权代表。AI 应以调查人员身份向其提问：
- 告知程序后，询问法人基本信息（姓名、职务）
- 询问公司基本信息（经营范围、用工人数等）
- 确认受伤职工与公司的劳动关系（入职时间、合同、岗位、工资、参保）
- 陈述公司对事故的调查结果和了解的情况
- 表明公司对工伤认定的态度（认可/不认可及理由）
- 最终由法人确认笔录无误

重要原则：
1. 法人代表公司陈述，不是以个人身份
2. 法人不描述"我受伤"——受伤的是职工
3. 与本人笔录的一致性：参考本人笔录中的关键事实，如有矛盾应标注
4. 劳动关系和工资参保信息是法人回答的核心内容
""",
    }

    return common_rules + role_templates.get(role, role_templates["本人"])


# ============================================================================
# User Prompt
# ============================================================================

def build_user_prompt(c: CaseClassification,
                     case_statement: str = "",
                     material_summary: str = "",
                     person_transcript: str = "") -> str:
    """
    构建 User Prompt —— 传入个案事实。

    Args:
        c: 案件分类结果
        case_statement: 案件申请陈述
        material_summary: 材料清单（✓✗来自UI勾选）
        person_transcript: 本人笔录全文（证人/法人角色时传入，作为基础事实参考）

    Returns:
        User Prompt 字符串
    """
    p = c.person
    co = c.company
    a = c.accident

    lines = [
        f"案件：{c.case_nature}，{c.applicant_type}，适用{c.regulation_text}",
        f"当前角色：{c.role}",
        f"受伤职工：{c.injured_worker_name}",
        "",
        f"=== 被调查人信息 ===",
        f"姓名：{p.name}",
        f"性别：{p.gender}" if p.gender else "",
        f"年龄：{p.age}岁" if p.age else "",
        f"身份证号：{p.idcard}" if p.idcard else "",
        f"地址：{p.id_address}" if p.id_address else "",
        f"电话：{p.phone}" if p.phone else "",
        f"岗位：{p.position}" if p.position else "",
        "",
        f"=== 用人单位信息 ===",
        f"公司：{co.company_name or '未填写'}",
        f"用人单位：{co.employer}" if co.employer else "",
        f"工地：{co.site}" if co.site else "",
    ]

    # 事故信息
    has_accident = any([a.time, a.location, a.description, a.diagnosis, a.hospital])
    if has_accident:
        lines.append("")
        lines.append("=== 事故信息 ===")
        if a.time:
            lines.append(f"事故时间：{a.time}")
        if a.location:
            lines.append(f"事故地点：{a.location}")
        if a.description:
            lines.append(f"受伤经过：{a.description[:200]}")
        if a.diagnosis:
            lines.append(f"诊断结论：{a.diagnosis}")
        if a.hospital:
            lines.append(f"送医医院：{a.hospital}")

    # 案件申请陈述
    if case_statement:
        lines.append("")
        lines.append("=== 案件申请陈述 ===")
        lines.append(case_statement[:800])

    # 本人笔录（证人/法人角色）
    if c.role != "本人" and person_transcript:
        lines.append("")
        lines.append("=== 受伤职工本人笔录（基础事实参考） ===")
        lines.append(person_transcript[:2000])
        lines.append("")
        lines.append("【重要】以上是受伤职工本人的谈话笔录。请据此：")
        lines.append("1. 以本人的事实陈述为基础，生成当前角色的提问")
        lines.append("2. 与本人陈述有出入的地方，追问核实并标注【注意：与本人陈述不一致】")
        lines.append("3. 本人已陈述的事实无需重复询问，重点补充当前角色特有的信息")

    # 材料清单
    if material_summary:
        lines.append("")
        lines.append("=== 材料清单（✓=已提供, ✗=缺失） ===")
        lines.append(material_summary)
        lines.append("")
        lines.append("缺失的材料（✗）必须在谈话中逐项追问：是否持有、如不能提供请说明原因。")

    lines.append("")
    lines.append("请根据以上全部信息，生成完整的谈话笔录问答内容。")
    lines.append("已有信息填入回答中。未掌握的用[待核实]标注。")

    return "\n".join(line for line in lines if line)


# ============================================================================
# 输出格式指令（追加到 User Prompt 末尾）
# ============================================================================

OUTPUT_FORMAT_INSTRUCTION = """
## 输出格式要求

只输出问答内容，每个问答以"问："和"答："开头。
不要输出任何标题、说明或解释文字。
不要使用markdown格式。

示例格式：
问：我们是永嘉县人力资源和社会保障局的工作人员（出示执法证件），现向你调查了解相关情况。你有如实提供相关情况的义务，如隐瞒或虚假陈述，将承担相应的法律责任。你听清楚了吗？
答：听清楚了。

问：请介绍一下你的姓名、住址、工作单位以及从事的工作？
答：我叫张三，身份证地址为浙江省永嘉县XX街道XX号，系XX建设工程公司的职工，被指派到XX劳务公司承建的XX工地，从事泥水工工作。
"""
