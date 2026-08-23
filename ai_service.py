# ai_service.py
import requests
from docx import Document
import json

from prompt_manager import load_prompt


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

        prompt = load_prompt('review').replace('{{笔录全文}}', truncated_text)

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

            prompt = load_prompt('optimize_injury').replace('{{原始描述}}', text)

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

            prompt = load_prompt('injury_and_conclusion').replace('{{条款信息}}', clause_injection).replace('{{笔录全文}}', truncated)

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
                    {"role": "system", "content": load_prompt('regulation_system')},
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
        prompt = load_prompt('regulation_user')
        prompt = prompt.replace('{{案本号}}', str(case_data.get('case_id', '')))
        prompt = prompt.replace('{{姓名}}', str(case_data.get('name', '')))
        prompt = prompt.replace('{{案件性质}}', str(case_data.get('case_nature', '')))
        prompt = prompt.replace('{{申请类型}}', str(case_data.get('applicant_type', '')))
        prompt = prompt.replace('{{用人单位}}', str(case_data.get('labor_unit', '')))
        prompt = prompt.replace('{{用工单位}}', str(case_data.get('employer', '')))
        prompt = prompt.replace('{{工地名称}}', str(case_data.get('site', '')))
        prompt = prompt.replace('{{受伤经过}}', str(case_data.get('injury_description', '')))
        prompt = prompt.replace('{{拟用条例}}', str(case_data.get('proposed_article', '')))
        prompt = prompt.replace('{{已提供证据}}', '、'.join(provided) if provided else '无')
        return prompt

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