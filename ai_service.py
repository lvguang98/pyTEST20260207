# ai_service.py
import requests
from docx import Document
import json


REGULATION_ANALYSIS_SYSTEM = """你是工伤认定领域的资深法律专家，精通《工伤保险条例》及工伤认定实务。

《工伤保险条例》认定工伤 / 视同工伤的情形对照表：
1. 第十四条第（一）项：在工作时间和工作场所内，因工作原因受到事故伤害的。
2. 第十四条第（二）项：工作时间前后在工作场所内，从事与工作有关的预备性或收尾性工作受到事故伤害的。
3. 第十四条第（三）项：在工作时间和工作场所内，因履行工作职责受到暴力等意外伤害的。
4. 第十四条第（四）项：患职业病的。
5. 第十四条第（五）项：因工外出期间，由于工作原因受到伤害或者发生事故下落不明的。
6. 第十四条第（六）项：在上下班途中，受到非本人主要责任的交通事故或者城市轨道交通、客运轮渡、火车事故伤害的。
7. 第十四条第（七）项：法律、行政法规规定应当认定为工伤的其他情形。
8. 第十五条第（一）项：在工作时间和工作岗位，突发疾病死亡或者在48小时之内经抢救无效死亡的。
9. 第十五条第（二）项：在抢险救灾等维护国家利益、公共利益活动中受到伤害的。
10. 第十五条第（三）项：职工原在军队服役，因战、因公负伤致残，已取得革命伤残军人证，到用人单位后旧伤复发的。

工伤认定常见关键证据：身份证、劳动合同、医院诊断证明、工资发放记录、考勤记录、证人证言、事故现场照片、监控录像、工伤认定书、劳动关系裁决书、道路交通事故认定书、公安报案回执、死亡证明、职业病诊断证明、病历资料等。

你的任务：根据案件事实判断最可能适用的条例，列出为认定该条例尚缺的关键证据，并对比工作人员填写的"拟用条例"给出一致/分歧判断及理由。只输出 JSON，不要输出任何解释或 markdown。"""


class AIService:
    def __init__(self, api_key: str, api_url: str = "https://api.deepseek.com"):
        self.api_key = api_key
        self.base_url = api_url  # 使用传入的api_url或默认值
        print(f"🔑 AI服务初始化，API密钥前8位: {api_key[:8]}...")
        print(f"🌐 API地址: {self.base_url}")

    def extract_text_from_docx(self, docx_path: str) -> str:
        """从Word文档中提取文本"""
        try:
            print(f"📖 正在读取Word文档: {docx_path}")
            doc = Document(docx_path)
            full_text = []

            for para in doc.paragraphs:
                if para.text.strip():  # 跳过空段落
                    full_text.append(para.text)

            result = "\n".join(full_text)
            print(f"✅ 提取文本成功，共{len(result)}字符")
            return result

        except Exception as e:
            print(f"❌ 提取文本失败: {str(e)}")
            raise Exception(f"读取Word文档失败: {e}")

    def analyze_legal_document(self, document_text: str) -> dict:
        """分析法律文档"""
        try:
            print(f"📤 准备发送AI请求，文本长度: {len(document_text)}")

            # 构建提示词
            prompt = self._build_prompt(document_text)

            # 准备请求
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }

            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "temperature": 0.3,
                "max_tokens": 2000
            }

            print(f"🌐 发送请求到: {self.base_url}/chat/completions")

            # 发送请求
            response = requests.post(
                f"{self.base_url}/chat/completions",  # ← 使用 self.base_url
                headers=headers,
                json=payload,
                timeout=60
            )

            print(f"📥 收到响应，状态码: {response.status_code}")

            if response.status_code == 200:
                result = response.json()
                print(f"✅ API调用成功")

                ai_response = result["choices"][0]["message"]["content"]
                print(f"📝 AI响应长度: {len(ai_response)}")
                print(f"📝 AI响应预览: {ai_response[:200]}...")

                return {
                    "审查状态": "完成",
                    "结果": ai_response,
                    "原始回复": ai_response
                }
            else:
                print(f"❌ API错误: {response.status_code}")
                print(f"❌ 错误详情: {response.text}")
                return {
                    "审查状态": "失败",
                    "错误信息": f"API调用失败: {response.status_code} - {response.text[:200]}"
                }

        except requests.exceptions.Timeout:
            print(f"⏰ 请求超时")
            return {
                "审查状态": "超时",
                "错误信息": "API请求超时，请检查网络连接"
            }
        except Exception as e:
            print(f"❌ 请求异常: {str(e)}")
            return {
                "审查状态": "异常",
                "错误信息": f"请求异常: {str(e)}"
            }

    # 在 ai_service.py 中的 _build_prompt 方法修改
    def _build_prompt(self, document_text: str) -> str:
        """
        构建智能补问提示词（增强版）
        原来的常规审查 + 新增缺失问题识别
        """
        # 限制文本长度
        truncated_text = document_text[:2500] if len(document_text) > 2500 else document_text

        prompt = f"""你是一个工伤案件专家，请审查以下谈话笔录并完成两项任务。

    谈话笔录内容：
    {truncated_text}

    ==================== 任务一：常规审查 ====================
    请按以下要点审查：
    1. 核心信息完整性（姓名、身份证号、用人单位、受伤时间地点等）
    2. 工伤三要素符合性（工作时间、场所、原因）
    3. 证据完整性（医疗证明、证人证言等）
    4. 是否符合工伤认定条件
    5. 需要补充哪些关键证据

    审查时请忽略程序性格式问题：
    - 调查人员签名是否完整
    - 被调查人确认签字是否完整  
    - 笔录格式是否规范
    - 表格填写是否完整

    ==================== 任务二：识别缺失问题（重点！）====================
    请列出笔录中**完全没有问到但非常重要**的问题。
    这些问题应该是工伤认定必需的核心问题。

    【输出格式要求】
    请严格按以下格式输出：

    【审查结果】
    （这里写常规审查结果，按上述要点分析）

    【缺失问题列表】
    □ 1. 问：[问题文本]
    □ 2. 问：[问题文本]
    □ 3. 问：[问题文本]
    ...（最多列出10个最重要的缺失问题）

    【问题筛选标准】
    1. 必须是笔录中完全没涉及的问题
    2. 必须是工伤认定必需的问题
    3. 问题要具体、明确、可回答
    4. 按重要性从高到低排序
    5. 每个问题以"问："开头
    6. 不要写答案，只写问题
    7. 每个问题前面用"□ "（方框加空格）开头，便于程序识别

    示例：
    【缺失问题列表】
    □ 1. 问：您与公司是否签订了书面劳动合同？
    □ 2. 问：公司是否为您缴纳了工伤保险？
    □ 3. 问：事故发生时是否有现场监控录像？

    请用中文回答，条理清晰。"""

        return prompt

    def optimize_injury_description(self, text):
        """
        用AI优化受伤经过描述（不处理医疗结论）
        :param text: 原始受伤经过描述（不含医疗结论）
        :return: 优化后的文本
        """
        try:
            if not text or not text.strip():
                return None

            prompt = f"""请将以下工伤案件受伤经过描述优化为正式的法律文书语言，要求：

    1. 语言正式、客观、准确，符合工伤认定法律文书的规范
    2. 保持事实原貌，不增加或减少任何事实细节
    3. 使用法律专业术语，语句通顺连贯
    4. 将送医情况自然地融入到整个事件描述中
    5. 段落清晰，逻辑严谨，时间顺序合理
    6. 输出格式：一段完整的叙述，不要分点或添加标题
    7. 特别注意：只优化受伤经过和送医情况，不涉及医疗诊断或结论
    8. 避免出现重复的标点符号（如连续两个句号"。。"）

    原始描述：
    {text}

    优化后的法律文书描述（一段完整的叙述）："""

            headers = {
                'Authorization': f'Bearer {self.api_key}',
                'Content-Type': 'application/json'
            }

            data = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.3,
                "max_tokens": 2000
            }

            print(f"📤 发送受伤经过优化请求，文本长度: {len(text)}")

            # 统一使用 self.base_url
            response = requests.post(
                f"{self.base_url}/chat/completions",  # ← 关键修改！
                headers=headers,
                json=data,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                content = result['choices'][0]['message']['content']

                # 清理响应
                optimized_text = content.strip()

                # 移除可能的多余提示
                unwanted_prefixes = [
                    "优化后的法律文书描述：",
                    "优化后的描述：",
                    "根据您提供的描述，",
                    "以下是对受伤经过的优化描述："
                ]

                for prefix in unwanted_prefixes:
                    if optimized_text.startswith(prefix):
                        optimized_text = optimized_text[len(prefix):].strip()

                print(f"✅ 受伤经过优化完成，新长度: {len(optimized_text)}")
                return optimized_text
            else:
                print(f"❌ 受伤经过优化失败，状态码: {response.status_code}")
                return None

        except Exception as e:
            print(f"❌ 优化受伤描述失败: {str(e)}")
            return None

    def generate_injury_and_conclusion(self, transcript_text: str,
                                       regulation_text: str = "",
                                       regulation_desc: str = "",
                                       regulation_elements=None):
        """
        一次性从本人笔录生成两个结果：调查核实情况段落 + 简洁诊断结论

        :param transcript_text: 本人笔录全文
        :param regulation_text: 适用条款原文（如《工伤保险条例》第十四条第一款第一项）
        :param regulation_desc: 适用条款对应情形说明（如"在工作时间和工作场所内，因工作原因受到事故伤害"）
        :param regulation_elements: 该条款须突出的关键证据要素（如三工要素）
        :return: {"受伤经过": str, "诊断结论": str}，失败返回 None
                （"受伤经过"字段即认定工伤决定书的"调查核实情况"段落）
        """
        try:
            if not transcript_text or not transcript_text.strip():
                return None

            truncated = transcript_text[:4000] if len(transcript_text) > 4000 else transcript_text

            # ── 组装条款情形说明（供 AI 突出关键证据要素）──
            clause_lines = []
            if regulation_text:
                clause_lines.append(f"本案拟适用的条款：{regulation_text}")
            if regulation_desc:
                clause_lines.append(f"该条款对应情形：{regulation_desc}")
            if regulation_elements:
                clause_lines.append("该条款要求突出以下关键证据要素：" + "、".join(regulation_elements))
            clause_block = "\n".join(clause_lines)

            if clause_block:
                clause_injection = (
                    "结合本案拟适用的条款情形，如下：\n" + clause_block +
                    "\n请在事实描述中自然嵌入上述关键证据要素，确保证据要素在事实描述中清晰可辨。\n"
                )
            else:
                clause_injection = ""

            prompt = f"""你是一名专业的工伤保险条例审核专家，精通《工伤保险条例》及工伤认定实务，擅长从调查笔录中提取关键事实，并按照法律文书规范撰写认定工伤决定书的事实认定部分。

任务：根据下面受伤职工本人谈话笔录中记载的受伤经过，{clause_injection}以笔录中记载的受伤经过为基础，草拟一份认定工伤决定书的"调查核实情况"段落。

撰写规则：
1. 严格以笔录记载为事实依据，不得添加笔录中没有记载的时间、地点、经过、人物等信息。
2. 如需对笔录中的模糊信息进行合理推断（如"李工"即班组长李龙国），应在草稿中以【】标注，由我最终确认。
3. 只写事实经过本身，不包含医疗诊断结论、法律适用分析、认定结论、救济途径告知等内容。
4. 严禁在段落开头、中间或末尾添加"系在工作时间和工作场所内，因工作原因受到事故伤害"之类的法律判断句、认定结论或条款套话，只客观陈述事实经过。
5. 语言风格使用工伤认定决定书的规范表述，如"系……职工""被指派到……"等，语言客观、准确、简洁。
6. 时间表述格式使用"XXXX年X月X日X时X分许"的格式。
7. 笔录中未明确的信息（如工友全名、具体时间等），以笔录记载为准，不加推测；如确需补充，以【待核实：……】标注。
8. 使用第三人称叙述（用受伤职工姓名）。

参考示例（注意：示例以"送往……医院检查治疗"收尾，没有法律判断句）：
郑秀杰系永嘉奥康鞋业营销有限公司的部门经理。2026年4月21日，郑秀杰在永嘉奥康鞋业营销有限公司的安排下，组织员工在永嘉县瓯北街道楠华广场开展临时特卖活动。当日8时50分许，在活动过程中，郑秀杰因疲劳前往收货台后方椅子处准备坐下休息，坐下时因椅子翻倒，致其摔倒受伤。同事李琼闻声赶至现场，将其扶起。因腰部疼痛加剧，郑秀杰遂联系同事王瑜、缪海龙驾车将其送往温州医科大学附属第二医院检查治疗。

请按以下格式输出，共两部分，不要输出任何其它内容：

第一部分：直接输出"调查核实情况"事实段落。不要加任何标题、不要分点、不要markdown，直接从受伤职工姓名开始写。

第二部分：另起一行，输出【诊断结论】并紧跟一个简短诊断结论词（如【诊断结论】右足跖骨骨折）。若笔录中无明确诊断结论，输出【诊断结论】无。

受伤职工本人谈话笔录：
{truncated}"""

            headers = {
                'Authorization': f'Bearer {self.api_key}',
                'Content-Type': 'application/json'
            }

            data = {
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.3,
                "max_tokens": 2000
            }

            print(f"📤 发送受伤事实+诊断结论生成请求，笔录长度: {len(transcript_text)}")

            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=headers,
                json=data,
                timeout=60
            )

            if response.status_code == 200:
                result = response.json()
                content = result['choices'][0]['message']['content'].strip()

                # 解析：以【诊断结论】为唯一分界，其前为调查核实情况段落，其后为诊断结论
                injury = ""
                conclusion = ""
                m2 = content.find("【诊断结论】")
                if m2 != -1:
                    injury = content[:m2].strip()
                    conclusion = content[m2 + len("【诊断结论】"):].strip()
                else:
                    injury = content.strip()

                # 清理段落开头可能残留的标题
                for header in ("【调查核实情况】", "调查核实情况：", "调查核实情况:", "调查核实情况"):
                    if injury.startswith(header):
                        injury = injury[len(header):].strip()
                        break

                # 清理诊断结论（去括号、去前缀、去多余文字）
                conclusion = conclusion.replace("【", "").replace("】", "").strip()
                for prefix in ["诊断结论：", "医疗诊断结论：", "结论："]:
                    if conclusion.startswith(prefix):
                        conclusion = conclusion[len(prefix):].strip()

                print(f"✅ 生成完成：调查核实情况 {len(injury)}字，诊断结论「{conclusion}」")
                return {"受伤经过": injury, "诊断结论": conclusion}
            else:
                print(f"❌ 生成失败，状态码: {response.status_code}")
                return None

        except Exception as e:
            print(f"❌ 生成失败: {e}")
            return None

    def generate_transcript(self, system_prompt: str, user_prompt: str) -> dict:
        """
        调用 DeepSeek API 生成谈话笔录问答内容。

        使用 System + User 双消息格式：
        - System Prompt 固化法律知识（可被 DeepSeek 缓存，不计入 Token）
        - User Prompt 传入个案事实（约 300 Token）

        Args:
            system_prompt: 法律规范的 System Prompt
            user_prompt: 个案事实的 User Prompt

        Returns:
            {"状态": "成功"/"失败", "内容": str, "错误信息": str}
        """
        try:
            print(f"📤 准备生成谈话笔录")
            print(f"   System Prompt 长度: {len(system_prompt)} 字符")
            print(f"   User Prompt 长度: {len(user_prompt)} 字符")

            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }

            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": 0.3,
                "max_tokens": 4000,
                "top_p": 0.9,
            }

            print(f"🌐 发送请求到: {self.base_url}/chat/completions")

            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=headers,
                json=payload,
                timeout=120  # 生成笔录允许更长的超时
            )

            print(f"📥 收到响应，状态码: {response.status_code}")

            if response.status_code == 200:
                result = response.json()
                content = result["choices"][0]["message"]["content"]

                # 获取 Token 用量信息
                usage = result.get("usage", {})
                prompt_tokens = usage.get("prompt_tokens", 0)
                completion_tokens = usage.get("completion_tokens", 0)
                total_tokens = usage.get("total_tokens", 0)

                print(f"✅ 笔录生成成功")
                print(f"   Prompt Tokens: {prompt_tokens}")
                print(f"   Completion Tokens: {completion_tokens}")
                print(f"   Total Tokens: {total_tokens}")
                print(f"   内容长度: {len(content)} 字符")

                return {
                    "状态": "成功",
                    "内容": content.strip(),
                    "用量": {
                        "prompt_tokens": prompt_tokens,
                        "completion_tokens": completion_tokens,
                        "total_tokens": total_tokens,
                    }
                }
            else:
                error_msg = f"API 返回错误 {response.status_code}: {response.text[:300]}"
                print(f"❌ {error_msg}")
                return {
                    "状态": "失败",
                    "内容": "",
                    "错误信息": error_msg
                }

        except requests.exceptions.Timeout:
            print("⏰ 请求超时（120秒）")
            return {
                "状态": "超时",
                "内容": "",
                "错误信息": "API请求超时，请检查网络连接后重试"
            }
        except Exception as e:
            print(f"❌ 请求异常: {str(e)}")
            return {
                "状态": "异常",
                "内容": "",
                "错误信息": f"请求异常: {str(e)}"
            }

    def generate_transcript_from_text(self, full_text: str) -> dict:
        """把「发送给AI的模板」渲染后的全文发给 AI，生成询问笔录。

        全文已包含角色设定、案件信息与任务指令，故作为 user prompt 一次性发送。
        """
        return self.generate_transcript("", full_text)

    # ========================================================================
    # 条例判断 + 缺失证据 + 拟用条例一致性分析
    # ========================================================================

    def analyze_case_for_regulation(self, case_data: dict) -> dict:
        """根据案件事实判断适用条例 + 缺失关键证据 + 拟用条例一致性。

        Args:
            case_data: 本人案件 JSON（含 name/employer/injury_description/proposed_article/materials 等）

        Returns:
            {"judged_article": str, "judged_article_reason": str,
             "missing_evidence": [str], "consistency": str, "reason": str}
            失败时返回 {"错误": str}
        """
        try:
            user_prompt = self._build_regulation_user_prompt(case_data)

            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            }
            payload = {
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": REGULATION_ANALYSIS_SYSTEM},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.2,
                "max_tokens": 2000,
                "response_format": {"type": "json_object"},
            }

            print("📤 发送条例与证据分析请求...")
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=headers,
                json=payload,
                timeout=120,
            )

            if response.status_code == 200:
                content = response.json()["choices"][0]["message"]["content"]
                print(f"✅ 条例分析完成，返回长度: {len(content)}")
                return self._parse_json_response(content)
            else:
                return {"错误": f"API返回错误 {response.status_code}: {response.text[:200]}"}

        except requests.exceptions.Timeout:
            return {"错误": "AI分析请求超时，请检查网络后重试"}
        except Exception as e:
            return {"错误": f"AI分析异常: {str(e)}"}

    def _build_regulation_user_prompt(self, case_data: dict) -> str:
        """构建条例分析的 User Prompt（个案事实）"""
        provided = [m.get('name', '') for m in case_data.get('materials', [])
                    if isinstance(m, dict) and m.get('name')]
        lines = [
            "请分析以下工伤案件：",
            "",
            f"- 案本号：{case_data.get('case_id', '')}",
            f"- 姓名：{case_data.get('name', '')}",
            f"- 案件性质：{case_data.get('case_nature', '')}",
            f"- 申请类型：{case_data.get('applicant_type', '')}",
            f"- 用人单位：{case_data.get('employer', '')}",
            f"- 用工单位：{case_data.get('labor_unit', '')}",
            f"- 工地名称：{case_data.get('site', '')}",
            f"- 受伤经过：{case_data.get('injury_description', '')}",
            f"- 拟用条例（工作人员填写）：{case_data.get('proposed_article', '')}",
            f"- 已提供证据：{'、'.join(provided) if provided else '无'}",
            "",
            "请以 JSON 格式输出（不要输出 markdown 代码块、不要任何解释），格式如下：",
            "{",
            '  "judged_article": "判断应适用的条例（从对照表中选，格式如：第十四条第（一）项）",',
            '  "judged_article_reason": "判断该条例的依据（结合受伤经过中的事实要点）",',
            '  "missing_evidence": ["尚缺的关键证据1", "尚缺的关键证据2"],',
            '  "consistency": "一致" 或 "分歧",',
            '  "reason": "与拟用条例一致或分歧的理由"',
            "}",
        ]
        return "\n".join(lines)

    @staticmethod
    def _parse_json_response(content: str) -> dict:
        """解析 AI 返回的 JSON（兼容 markdown 代码块包裹）"""
        text = (content or "").strip()
        if text.startswith("```"):
            lines = [l for l in text.split("\n") if not l.strip().startswith("```")]
            text = "\n".join(lines).strip()
        try:
            data = json.loads(text)
            if isinstance(data, dict):
                return data
        except Exception as e:
            print(f"⚠️ JSON 解析失败，尝试正则提取: {e}")
            import re
            m = re.search(r"\{.*\}", text, re.S)
            if m:
                try:
                    return json.loads(m.group(0))
                except Exception:
                    pass
        return {"错误": "无法解析AI返回的JSON", "原始": (content or "")[:500]}