# 回归测试说明

本文档记录已修复的 bug 及对应的回归测试方式，便于后续改动时验证。

## 已覆盖的回归用例

| Bug | 复现场景 | 测试位置 | 运行方式 |
|-----|----------|----------|----------|
| 滴滴多行程误去重 | 同起点、不同终点的两条行程被合并为一条 | `tests/test_dedup.py` | `pytest tests/test_dedup.py` |
| 日期补零 | 配置中 `2025-8-12` 未补零导致搜索/筛选异常 | `tests/test_date_padding.py` | `pytest tests/test_date_padding.py` |

## 需手动或快照验证的用例

| Bug | 复现场景 | 验证方式 |
|-----|----------|----------|
| 首汽约车行程单识别 | PDF 含「首汽约车」「行程报销单」等关键词 | 需真实 PDF fixture，或保留 `output2/` 快照定期跑全流程 |
| 第三方行程单与机票发票配对 | 开票日期晚于 trip_date_to 时发票被误过滤 | 同上，依赖完整 pipeline + 真实数据 |
| 第三方行程单图片裁剪 | 首汽约车行程单上下空白未裁剪 | 需 PDF + 生成 docx 后人工检查 |

## 运行所有回归测试

```bash
# 安装测试依赖
pip install pytest

# 运行
pytest tests/ -v
```

## 扩展建议

1. **快照回归**：将 `output2/` 中 `intermediate/info_after_filter.json` 作为 golden，每次改动后跑全流程并 diff
2. **CI 集成**：在 GitHub Actions 中 `pytest tests/`，无需邮箱即可跑单元测试
3. **PDF Fixtures**：若可脱敏，可将典型 PDF 放入 `tests/fixtures/pdfs/` 做提取与裁剪的集成测试
