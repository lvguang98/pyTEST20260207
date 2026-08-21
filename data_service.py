# data_service.py
import datetime
from typing import Dict, List, Any, Tuple, Optional


class DataService:
    """数据处理服务 - 负责所有数据计算、验证、转换操作"""

    def __init__(self):
        """初始化DataService"""
        pass

    def validate_case_data(self, case_data: Dict[str, Any]) -> List[str]:
        """
        统一验证案件数据
        :param case_data: 包含所有案件数据的字典
        :return: 错误信息列表
        """
        errors = []

        # 1. 验证人员信息
        person_data = {
            '姓名': case_data.get('姓名', ''),
            '身份证号': case_data.get('身份证号', ''),
            '年龄': case_data.get('年龄', '')
        }
        person_errors = self.validate_person_data(person_data)
        if person_errors:
            errors.extend([f"个人信息: {err}" for err in person_errors])

        # 2. 验证公司信息
        company_name = case_data.get('用工单位', '')
        if not company_name:
            errors.append("用工单位不能为空")

        # 3. 验证案件信息
        case_number = case_data.get('案本号', '')
        if not case_number:
            errors.append("案本号不能为空")

        return errors

    @staticmethod
    def validate_required_fields(data: Dict[str, Any], field_names: List[str]) -> List[str]:
        """
        验证必填字段
        :param data: 数据字典
        :param field_names: 字段名列表
        :return: 错误信息列表
        """
        errors = []
        for field in field_names:
            value = data.get(field)
            if value is None or (isinstance(value, str) and not value.strip()):
                errors.append(f"{field}不能为空")
        return errors

    # ============ 身份证相关方法 ============

    @staticmethod
    def calculate_age_from_idcard(idcard: str) -> Optional[int]:
        """
        根据身份证号计算年龄
        :param idcard: 身份证号码（15位或18位）
        :return: 年龄（整数），如果身份证无效返回None
        """
        try:
            if not idcard or len(idcard) not in (15, 18):
                return None

            # 提取出生日期
            if len(idcard) == 18:
                birth_date_str = idcard[6:14]  # YYYYMMDD
            else:  # 15位身份证
                birth_date_str = f"19{idcard[6:12]}"  # 19YYMMDD

            # 转换为日期
            birth_date = datetime.datetime.strptime(birth_date_str, "%Y%m%d")

            # 计算年龄
            today = datetime.datetime.now()
            age = today.year - birth_date.year

            # 如果今年生日还没过，减1岁
            if (today.month, today.day) < (birth_date.month, birth_date.day):
                age -= 1

            return age

        except Exception as e:
            print(f"[DataService] 计算年龄失败: {e}")
            return None

    @staticmethod
    def extract_gender_from_idcard(idcard: str) -> Optional[str]:
        """
        根据身份证号提取性别
        :param idcard: 身份证号码（15位或18位）
        :return: "男" 或 "女"，如果身份证无效返回None
        """
        try:
            if not idcard:
                return None

            if len(idcard) == 18:
                gender_digit = int(idcard[16])  # 第17位
            elif len(idcard) == 15:
                gender_digit = int(idcard[14])  # 第15位
            else:
                return None

            # 奇数男性，偶数女性
            return "男" if gender_digit % 2 == 1 else "女"

        except Exception as e:
            print(f"[DataService] 提取性别失败: {e}")
            return None

    @staticmethod
    def validate_idcard(idcard: str) -> Tuple[bool, str]:
        """
        验证身份证号格式
        :param idcard: 身份证号码
        :return: (是否有效, 错误信息)
        """
        if not idcard:
            return False, "身份证号不能为空"

        # 移除空格
        idcard = idcard.strip()

        # 检查长度
        if len(idcard) not in (15, 18):
            return False, "身份证号必须是15位或18位"

        # 检查字符
        if len(idcard) == 18:
            # 18位：前17位必须是数字，最后一位可以是数字或X
            if not idcard[:17].isdigit():
                return False, "身份证号前17位必须是数字"
            if not (idcard[17].isdigit() or idcard[17].upper() == 'X'):
                return False, "身份证号最后一位必须是数字或X"
        else:  # 15位
            if not idcard.isdigit():
                return False, "15位身份证号必须全部是数字"

        return True, "身份证号格式正确"

    # ============ 数据验证方法 ============

    @staticmethod
    def validate_person_data(person_data: Dict[str, Any]) -> List[str]:
        """
        验证人员数据
        :param person_data: 人员数据字典
        :return: 错误信息列表
        """
        errors = []

        # 验证姓名
        name = person_data.get('姓名', '')
        if not name:
            errors.append("姓名不能为空")
        elif len(name) > 15:
            errors.append("姓名长度不能超过15个字符")

        # 验证身份证号
        idcard = person_data.get('身份证号', '')
        if idcard:
            is_valid, msg = DataService.validate_idcard(idcard)
            if not is_valid:
                errors.append(f"身份证号: {msg}")

        # 验证年龄
        age = person_data.get('年龄', '')
        if age:
            if isinstance(age, str):
                if not age.isdigit():
                    errors.append("年龄必须是数字")
                else:
                    age_int = int(age)
                    if not (16 <= age_int <= 100):
                        errors.append("年龄必须在16-100岁之间")
            elif isinstance(age, int):
                if not (16 <= age <= 100):
                    errors.append("年龄必须在16-100岁之间")

        return errors

    @staticmethod
    def get_current_datetime(format_str: str = "%Y年%m月%d日%H时%M分") -> str:
        """
        获取当前日期时间并格式化
        :param format_str: 格式字符串
        :return: 格式化后的日期时间字符串
        """
        return datetime.datetime.now().strftime(format_str)

    # ============ 案件相关数据方法 ============

    @staticmethod
    def generate_case_number(base_number: int, is_death_case: bool = False,
                             date_str: Optional[str] = None) -> str:
        """
        生成案本号
        :param base_number: 基础编号
        :param is_death_case: 是否为工亡案件
        :param date_str: 日期字符串（YYYYMMDD），如果不提供则使用当前日期
        :return: 案本号字符串
        """
        if not date_str:
            date_str = datetime.datetime.now().strftime("%Y%m%d")

        case_type = "GW" if is_death_case else "GS"
        return f"{date_str}-{case_type}-{base_number:03d}"

