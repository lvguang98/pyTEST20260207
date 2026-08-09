"""
case_index.py
案件索引管理 - 轻量级JSON索引系统
"""

import os
import json
import datetime
from typing import Dict, Any, List, Optional
from pathlib import Path
from dataclasses import dataclass, asdict
from path_utils import path_utils


@dataclass
class CaseIndexEntry:
    """案件索引条目（最小数据集）"""
    case_number: str = ""  # 案本号，如：案本-20250127-001
    person_name: str = ""  # 受伤职工姓名
    id_card_last4: str = ""  # 身份证后4位
    company_name: str = ""  # 公司名称
    case_type: str = ""  # 案件类型（普通/工亡等）
    created_date: str = ""  # 创建日期
    folder_name: str = ""  # 文件夹名
    transcript_file: str = ""  # 本人笔录文件名
    updated_time: str = ""  # 最后更新时间

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return asdict(self)


class CaseIndexManager:
    """案件索引管理器"""

    def __init__(self, base_path: Optional[str] = None):
        # 使用统一的存储路径
        self.base_path = Path(base_path) if base_path else path_utils.get_storage_path()
        self.index_file = self.base_path / "cases_index.json"
        self.backup_dir = self.base_path / "_index_backups"

        # 确保目录存在
        self.backup_dir.mkdir(parents=True, exist_ok=True)

        # 加载索引
        self.index = self._load_index()
        print(f"[INFO] 案件索引管理器初始化，当前案件数: {len(self.index.get('cases', {}))}")

    def _load_index(self) -> Dict[str, Any]:
        """加载索引文件"""
        if self.index_file.exists():
            try:
                with open(self.index_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    print(f"[OK] 加载案件索引，共 {len(data.get('cases', {}))} 个案件")
                    return data
            except json.JSONDecodeError as e:
                print(f"[ERROR] 索引文件损坏，从备份恢复: {e}")
                return self._restore_from_backup()
            except Exception as e:
                print(f"[ERROR] 加载索引失败: {e}")

        # 创建新索引
        return {
            "version": "1.0",
            "created_date": datetime.datetime.now().isoformat(),
            "last_updated": "",
            "total_cases": 0,
            "cases": {},  # case_number -> CaseIndexEntry数据
            "statistics": {
                "by_month": {},
                "by_company": {},
                "by_case_type": {}
            }
        }

    def _save_index(self):
        """保存索引文件"""
        try:
            # 先备份当前索引
            self._create_backup()

            # 更新元数据
            self.index["last_updated"] = datetime.datetime.now().isoformat()
            self.index["total_cases"] = len(self.index["cases"])

            # 保存
            with open(self.index_file, 'w', encoding='utf-8') as f:
                json.dump(self.index, f, ensure_ascii=False, indent=2)

            print(f"[SAVE] 案件索引已保存，共 {self.index['total_cases']} 个案件")
            return True

        except Exception as e:
            print(f"[ERROR] 保存索引失败: {e}")
            return False

    def _create_backup(self):
        """创建备份"""
        if self.index_file.exists():
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = self.backup_dir / f"index_backup_{timestamp}.json"

            import shutil
            shutil.copy2(self.index_file, backup_file)

            # 清理旧备份（最多保留5个）
            self._cleanup_old_backups()

    def _cleanup_old_backups(self, max_backups: int = 5):
        """清理旧备份"""
        backup_files = list(self.backup_dir.glob("index_backup_*.json"))
        if len(backup_files) > max_backups:
            # 按时间排序，删除最旧的
            backup_files.sort(key=lambda x: x.stat().st_mtime)
            for old_file in backup_files[:-max_backups]:
                old_file.unlink()
                print(f"[CLEAN] 清理旧备份: {old_file.name}")

    def _restore_from_backup(self) -> Dict[str, Any]:
        """从备份恢复"""
        backup_files = list(self.backup_dir.glob("index_backup_*.json"))
        if backup_files:
            # 使用最新的备份
            backup_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            latest_backup = backup_files[0]

            try:
                with open(latest_backup, 'r', encoding='utf-8') as f:
                    print(f"[RESTORE] 从备份恢复索引: {latest_backup.name}")
                    return json.load(f)
            except:
                pass

        # 无法恢复，返回空索引
        return self._load_index()  # 这会递归调用，但会返回新索引

    def add_case(self, case_data: CaseIndexEntry) -> bool:
        """
        添加案件到索引

        Args:
            case_data: 案件索引数据

        Returns:
            是否成功
        """
        try:
            # 转换为字典
            case_dict = case_data.to_dict()
            case_number = case_dict["case_number"]

            if not case_number:
                print("[ERROR] 案本号不能为空")
                return False

            # 添加到索引
            self.index["cases"][case_number] = case_dict

            # 更新统计信息
            self._update_statistics(case_dict)

            # 保存索引
            success = self._save_index()

            if success:
                print(f"[OK] 案件已添加到索引: {case_number}")
            else:
                print(f"[ERROR] 案件添加失败: {case_number}")

            return success

        except Exception as e:
            print(f"[ERROR] 添加案件到索引失败: {e}")
            return False

    def _update_statistics(self, case_data: Dict[str, Any]):
        """更新统计信息"""
        stats = self.index["statistics"]

        # 按月统计
        created_date = case_data.get("created_date", "")
        if created_date and len(created_date) >= 7:
            month_key = created_date[:7]  # YYYY-MM
            stats["by_month"][month_key] = stats["by_month"].get(month_key, 0) + 1

        # 按公司统计
        company_name = case_data.get("company_name", "")
        if company_name:
            stats["by_company"][company_name] = stats["by_company"].get(company_name, 0) + 1

        # 按案件类型统计
        case_type = case_data.get("case_type", "普通案件")
        stats["by_case_type"][case_type] = stats["by_case_type"].get(case_type, 0) + 1

    def find_by_case_number(self, case_number: str) -> Optional[Dict[str, Any]]:
        """按案本号查找案件"""
        return self.index["cases"].get(case_number)

    def find_by_person_name(self, person_name: str) -> List[Dict[str, Any]]:
        """按姓名查找案件"""
        results = []
        person_name_lower = person_name.lower()

        for case_number, case_data in self.index["cases"].items():
            if person_name_lower in case_data.get("person_name", "").lower():
                results.append({
                    "case_number": case_number,
                    **case_data
                })

        # 按创建时间倒序排列
        results.sort(key=lambda x: x.get("created_date", ""), reverse=True)
        return results

    def find_by_company(self, company_name: str) -> List[Dict[str, Any]]:
        """按公司名称查找案件"""
        results = []
        company_lower = company_name.lower()

        for case_number, case_data in self.index["cases"].items():
            if company_lower in case_data.get("company_name", "").lower():
                results.append({
                    "case_number": case_number,
                    **case_data
                })

        return results

    def get_recent_cases(self, limit: int = 10) -> List[Dict[str, Any]]:
        """获取最近添加的案件"""
        all_cases = list(self.index["cases"].items())
        # 按创建时间排序（假设created_date格式为YYYY-MM-DD）
        all_cases.sort(key=lambda x: x[1].get("created_date", ""), reverse=True)

        return [
            {"case_number": case_number, **case_data}
            for case_number, case_data in all_cases[:limit]
        ]

    def get_statistics(self) -> Dict[str, Any]:
        """获取统计信息"""
        return {
            "total_cases": self.index["total_cases"],
            "by_month": self.index["statistics"]["by_month"],
            "by_company": self.index["statistics"]["by_company"],
            "by_case_type": self.index["statistics"]["by_case_type"],
            "last_updated": self.index["last_updated"]
        }

    def remove_case(self, case_number: str) -> bool:
        """从索引中移除案件"""
        if case_number in self.index["cases"]:
            del self.index["cases"][case_number]
            self._save_index()
            print(f"[REMOVE] 已从索引移除案件: {case_number}")
            return True
        return False

    def cleanup_old_cases(self, days: int = 365):
        """清理指定天数前的案件索引"""
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=days)
        cutoff_str = cutoff_date.strftime("%Y-%m-%d")

        cases_to_remove = []
        for case_number, case_data in self.index["cases"].items():
            created_date = case_data.get("created_date", "")
            if created_date and created_date < cutoff_str:
                cases_to_remove.append(case_number)

        for case_number in cases_to_remove:
            self.remove_case(case_number)

        print(f"[CLEAN] 清理了 {len(cases_to_remove)} 个旧案件索引")