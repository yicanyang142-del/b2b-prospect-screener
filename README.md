# 客户筛选 Skill

把 To B 销售筛选方法转化为可解释的 Python 原型：先用一票否决排除撞单与财务恶化企业，再用 CAGR、榜单信号、关系裂变和 KP 可触达性打分，输出候选客户清单和 Excel 报告。

## 特性

- 三阶段筛选：一票否决、规则打分、S/A/B 评级
- 每个加分项都保留得分与证据来源
- 支持 JSON / CSV / XLSX 输入
- 输出 JSON 报告和 `filtered_leads_result.xlsx`
- 仅使用本地 Mock 数据，不接真实企查查/天眼查/CRM，也不包含爬虫

## 运行

```bash
pip install -r requirements.txt
python screener.py --input SYNTHETIC_DATASET.json --excel-output filtered_leads_result.xlsx
```

也可以直接使用其他 CSV / XLSX 输入：

```bash
python screener.py --input your_customers.csv
python screener.py --input your_customers.xlsx
```

## 输出

JSON 顶层结构为：

```json
{
  "total_count": 3,
  "prospects": []
}
```

Excel 报告包含公司名、最终评级、得分、一票否决原因、主要加分项和 KP 信息。

## 边界

当前为规则原型，不用于真实企业决策或商业数据抓取。联系人信息只能来自输入数据，不能由模型或脚本编造。
