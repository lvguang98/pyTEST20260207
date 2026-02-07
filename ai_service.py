# ai_service.py
import requests
from docx import Document
import json


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