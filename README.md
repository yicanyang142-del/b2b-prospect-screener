# B2B Prospect Screener｜客户筛选

这是一个可解释的 To B 销售客户筛选原型：先按一票否决排除撞单和财务恶化企业，再用 CAGR、榜单信号、关系裂变和关键人可触达性打分，输出候选客户清单和 Excel 报告。

## 解决什么问题

销售在联系客户前需要判断企业是否值得优先跟进。这个项目把筛选逻辑固化成规则引擎，保留每一条加分原因和证据来源，避免黑盒输出。

## 三阶段流程

1. 一票否决：撞单、严重财务恶化、关键人缺失
2. 规则打分：财务健康度、榜单信号、关系网络、关键人可触达性
3. 评级输出：S级 / A级 / B级 / 淘汰

## 输入

支持 JSON、CSV 和 XLSX。

常用字段包括：

- 公司名称
- 近三年营收
- 关键人姓名、职务和联系方式
- 榜单信号
- 近 6 个月招投标次数
- 与已合作客户的上下游或同行关系
- CRM 跟进状态

## 输出

返回 JSON：

```json
{
  "total_count": 3,
  "prospects": []
}
```

同时生成 `filtered_leads_result.xlsx`，包含公司名、最终评级、得分、一票否决原因、主要加分项和 KP 信息。

## 运行方式

```sh
python -m pip install -r requirements.txt
python screener.py --input SYNTHETIC_DATASET.json --excel-output filtered_leads_result.xlsx
```

也支持其他文件：

```sh
python screener.py --input your_customers.csv
python screener.py --input your_customers.xlsx
```

## 边界与约束

- 当前为规则原型，不接真实企查查 / 天眼查 / CRM API。
- 不包含爬虫。
- 联系人信息只能来自输入数据，不能由模型或脚本编造。
- 不用于真实企业决策或商业数据抓取。
