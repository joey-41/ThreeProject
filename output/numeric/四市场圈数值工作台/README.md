# 四市场圈数值工作台

本目录用于承接“四市场圈数值落地与跑商沙盘首轮计划”，把现有机制文档中的准配置字段整理成可直接进 Excel 的表包，并提供首轮子市场级经济沙盘。

## 数值来源

本工作台只引用以下文档作为首轮口径来源：

1. `/Users/shaojiahao132776/ThreeProject/output/story/武侠放置网游_世界观_节点地图_背景故事线.md`
2. `/Users/shaojiahao132776/ThreeProject/output/design/武侠放置网游_详细机制模块/02_地图与路线.md`
3. `/Users/shaojiahao132776/ThreeProject/output/design/武侠放置网游_详细机制模块/05_制造采集与经济.md`
4. `/Users/shaojiahao132776/ThreeProject/output/design/武侠放置网游_详细机制模块/06_市场交易与运输.md`
5. `/Users/shaojiahao132776/ThreeProject/output/design/武侠放置网游_详细机制模块/09_帮派社交与政治经济.md`
6. `/Users/shaojiahao132776/ThreeProject/output/design/武侠放置网游_详细机制模块/07_PvE玩法.md`

## 目录结构

- `tables/`
  - 9 张 UTF-8 CSV 表，字段与 Excel 页签一一对应。
  - 以 `12` 个子市场作为首轮聚合层，不直接展开到 `196` 节点。
- `reports/`
  - 沙盘脚本生成的场景指标、Markdown 报告和中间结果。
- `四市场圈数值工作台.xlsx`
  - 由脚本根据 `tables/` 与模拟结果自动生成的工作簿。

## 表包说明

工作簿固定包含以下 9 张页签：

1. `01_子市场总表`
2. `02_路线经济表`
3. `03_贸易商品32类表`
4. `04_通用配方24表`
5. `05_产区与产量表`
6. `06_银两收支总账表`
7. `07_玩家分层行为表`
8. `08_沙盘场景表`
9. `09_校验仪表盘`

其中：

- `tables/09_校验仪表盘.csv` 维护的是阈值与验收口径。
- 实际模拟值会由脚本写入 Excel 的 `09_校验仪表盘` 页签，以及 `reports/` 下的结果文件。

## 运行方式

在仓库根目录执行：

```bash
python3 scripts/four_market_sandbox.py
```

脚本会做三件事：

1. 读取 `tables/` 中的 9 张表。
2. 跑 `7` 日基线场景和 `3` 个扰动场景。
3. 生成 Excel 工作簿、指标 CSV、JSON 摘要和首轮报告。

## 首轮默认假设

- 模拟粒度为“子市场日循环”，不是单节点逐跳事件模拟。
- `196` 节点先映射为 `12` 个子市场经济节点，后续可按同字段下钻。
- 默认税档使用平衡档：交易税 `5%`、过路税 `4%`、资源税 `7%-8%`。
- 帮派影响首轮只折算为税率、维护费、军需需求和补给线稳定度。
- 危险区事件态 `2.15x` 收益不作为常态基线，仅在扰动场景中放大局部需求。

## 后续扩展方向

- 把 `sub_market_id` 向下展开为节点级 `NodeRegion` 与 `RouteEdge`。
- 把 `TradeGoodFocusConfig` 与 `RecipeConfig` 迁移为正式配置表。
- 在第二轮中引入黑市补口、命名精英掉落波动和帮派据点实时税权。
