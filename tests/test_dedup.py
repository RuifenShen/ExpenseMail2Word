"""
回归测试：滴滴多行程去重
Bug: 同起点、同金额、不同终点的行程被误去重，导致 Word 表格只显示部分行程
Fix: 去重键包含 end_location
"""
import json
import pytest
from pathlib import Path

from process_expense import dedup_expense_items


def test_didi_multi_trip_same_start_different_end_not_deduped():
    """同起点、不同终点的两条滴滴行程单应保留，不被误去重"""
    fixture_path = Path(__file__).parent / "fixtures" / "dedup_fixture.json"
    items = json.loads(fixture_path.read_text(encoding="utf-8"))
    result = dedup_expense_items(items)
    assert len(result) == 2, "同起点不同终点的两条行程应都保留"
    assert result[0]["end_location"] == "北京南站-西"
    assert result[1]["end_location"] == "北京西站-南"


def test_didi_duplicate_same_route_deduped():
    """同起点、同终点的重复行程应被去重"""
    items = [
        {"type": "didi", "doc_type": "行程单", "date": "2025-11-05", "start_location": "A", "end_location": "B", "amount": 10.0, "original_filename": "a.pdf"},
        {"type": "didi", "doc_type": "行程单", "date": "2025-11-05", "start_location": "A", "end_location": "B", "amount": 10.0, "original_filename": "b.pdf"},
    ]
    result = dedup_expense_items(items)
    assert len(result) == 1
