请使用工作区中的上游价格页、Kulinar 配方库、预算说明和 PDF 指南，为 2026 年 3 月 16日至20日的公司食堂制作一份可回放的周菜单交付包。请交付：

- `/app/output/meal_cost_optimizer.py`
- `/app/output/meal_optimization_results.json`
- `/app/output/Weekly_Meal_Budget.xlsx`
- `/app/output/Cafeteria_Meal_Proposal.docx`
- `/app/output/meal_prep_calendar.ics`

这是 Cowork-Bench `terminal-pw-kulinar-excel-word-gcal` 的离线静态改造版。原题通过本地网页、Kulinar MCP 和 Google Calendar 完成；本题保留同一上游价格页与完整 Kulinar 数据，并将日历交付静态化为 ICS，不需要网络、MCP 或外部账号。

输入职责：

- `grocery_prices.html`：上游价格对比页，共 22 种商品和 3 家商店的报价，单位为卢布。
- `all_recipes.json`：上游 Kulinar MCP 的完整 50 条配方库；本题只使用 `category = "горячее"` 的配方。
- `Meal_Planning_Guidelines.pdf` 和 `budget_notes.md`：上游预算、菜单、采购与日历约束。

为了消除原题的计价歧义，本 Benchmark 使用以下明确口径：

- 配方原料名默认与价格表 `Товар` 精确匹配；仅允许三个显式别名：`Говядина (вырезка) -> Говядина`、`Филе рыбы (судак/треска) -> Филе рыбы`、`Фарш мясной (свинина+говядина) -> Фарш мясной`。
- 价格页的报价单位与配方用量单位并不完全可换算，因此不得虚构重量或单价。本题的 `Total_Cost` 定义为“报价篮子成本”：每个能匹配的不重复原料，按三家商店的最低报价计一个报价单位；未报价原料不估值，必须单独披露。
- 仅考虑至少有 3 个可匹配原料的热菜，按 `Total_Cost` 升序、再按 `Recipe_ID` 升序选前 5 个。
- `Cost_Per_Serving = round(Total_Cost / 4, 2)`，四人份口径来自上游 Query/PDF，不使用配方库中不同的原始 `servings`。
- 5 个菜按成本升序对应 Monday-Friday 的 Lunch，按成本降序对应 Dinner，形成 10 行菜单。
- 周预算为 `12000`。“单店基准”指对周菜单中每一个报价篮子都在同一家商店购买，三家中总价最低者为 `Best_Single_Store`。

`meal_cost_optimizer.py` 必须真实读取 HTML 和 JSON，并支持：

```bash
python meal_cost_optimizer.py \
  --prices /app/data/input_files/grocery_prices.html \
  --recipes /app/data/input_files/all_recipes.json \
  --output /app/output/meal_optimization_results.json
```

Judge 会在原始价格页和修改后的价格页上分别回放该脚本，所以不得把商店、菜名、价格或汇总值直接硬编码为固定答案。JSON 必须包含 `price_rows`、`selected_recipes`、`weekly_meal_plan`、`budget_summary` 和 `limitations`。

`Weekly_Meal_Budget.xlsx` 必须包含三张可编辑表：

- `Grocery_Prices`：`Item, Unit, FreshMart_Price, ValueGrocer_Price, BulkBarn_Price, Cheapest_Store, Cheapest_Price`，22 行，按 Item 升序。
- `Recipe_Cost_Analysis`：`Recipe_ID, Recipe, Matched_Ingredient_Count, Unpriced_Ingredient_Count, Total_Cost, Cheapest_Store_Mix, Cost_Per_Serving, Unpriced_Ingredients`，5 行，按 Total_Cost 升序。`Cost_Per_Serving` 必须用 Excel 公式计算。
- `Weekly_Meal_Plan`：`Date, Day, Meal, Recipe_ID, Recipe, Estimated_Cost`，10 行，按日期和 Lunch/Dinner 排序。`Estimated_Cost` 必须通过公式从 `Recipe_Cost_Analysis` 查找。

`Cafeteria_Meal_Proposal.docx` 必须使用上游指定的英文标题 `Weekly Cafeteria Meal Proposal`，并用俄文完整呈现 `Budget Overview`、`Menu Highlights`、`Procurement Strategy` 和 `Limitations`。关键数字必须与 JSON/Excel 一致，并明确说明报价篮子口径。PDF 要求每份至少 20g 蛋白质，但输入没有营养数据；文档必须将其标为“无法验证”，不得声称已满足。

`meal_prep_calendar.ics` 必须包含 5 个 `VEVENT`，分别安排在 2026-03-16 至 2026-03-20 08:00-09:30，时区 `Europe/Moscow`。标题为 `Meal Prep: [当日 Lunch 菜名]`，描述需按商店列出已报价原料，并单列未报价原料。
