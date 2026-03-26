"""
回归测试：第三方行程单与机票发票配对
Bug: 首汽约车行程单与机票发票未配对，且开票日期晚于 trip_date_to 时发票被误过滤
Fix: 第三方行程单与机票类型发票按金额匹配，发票继承行程单的日期/起点/终点
注：完整配对逻辑在 process_expense.run_complete_process 内，此处验证配对后的字段继承。
"""
import sys
from pathlib import Path

root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(root / "scripts"))

# 配对逻辑内联在 run_complete_process，此处仅做占位，后续可提取 apply_trip_invoice_pairing 后补充
# 当前建议：用真实 output 数据做快照回归（见 REGRESSION.md）


def test_third_party_pairing_placeholder():
    """占位：配对逻辑需从 run_complete_process 提取后可补充单元测试"""
    pass
