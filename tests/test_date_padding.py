"""
回归测试：日期补零
Bug: 配置中单数字月/日（如 8、1）未补零，导致搜索或筛选异常
Fix: email_utils.normalize_date 及 process_expense 内 _norm_d
"""
import sys
from pathlib import Path

# 确保能导入 scripts 下的模块
root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(root / "scripts"))

from utils.email_utils import normalize_date


def test_normalize_date_single_digit_month():
    """2025-8-12 -> 2025-08-12"""
    assert normalize_date("2025-8-12") == "2025-08-12"


def test_normalize_date_single_digit_day():
    """2025-12-1 -> 2025-12-01"""
    assert normalize_date("2025-12-1") == "2025-12-01"


def test_normalize_date_both_padded():
    """2025-8-1 -> 2025-08-01"""
    assert normalize_date("2025-8-1") == "2025-08-01"


def test_normalize_date_already_padded():
    """2025-08-12 保持不变"""
    assert normalize_date("2025-08-12") == "2025-08-12"


def test_normalize_date_empty():
    """空字符串返回空"""
    assert normalize_date("") == ""
    assert normalize_date("   ") == ""
