# 运行报告

在项目根目录运行 `python3 scripts/verify_workbook.py` 和 `python3 scripts/reconcile_orders.py`，会在此生成当前核验报告。报告不纳入Git提交；历史43项结果在 `examples/merchant-v0.3/复核结果.json` 中保留。

这些核对比较已保存的工作簿缓存和虚构账单，不运行Excel/WPS实时重算，也不替代真人试用或全部软件兼容测试。
