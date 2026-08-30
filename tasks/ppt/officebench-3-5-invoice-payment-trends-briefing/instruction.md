请读取 `/app/input_files/Invoice.pdf` 和 `/app/input_files/Invoice_2.pdf`，制作一份 5 页、16:9 的 2023 年 1 月发票支付趋势与财务复核简报。

必须满足：

- 用原生可编辑表格逐字段登记两张发票的来源文件、Invoice number、Date、Customer、Description、Quantity、Unit price、Subtotal、Sales tax、Shipping & handling、Total due 和 Terms。
- 准确汇总：Subtotal `$1,399.99`、Sales tax `$69.99`、Shipping `$24.99`、Total due `$1,494.97`。
- 用原生可编辑图表比较 Cloud service 与 Decorative clay pottery 两张发票的 Subtotal、Sales tax、Shipping & handling 构成。
- 明确指出两份来源都使用 `Invoice #100`，但日期、客户和业务内容不同；不得只按 Invoice number 去重。
- 准确说明两张发票日期相隔 29 天，付款条款均为 due upon receipt / due on receipt。
- 明确区分“付款条款”和“付款状态”：输入没有 Paid / Unpaid、Payment date、Approval、Payment method 或 Clearing status，因此不得把任何发票写成已支付、未支付、逾期或已结清。
- 保留 Cloud service 发票中“供应商名称未打印”的数据质量问题；不得把客户 The Social Strategists 当成供应商。
- 保留 Pottery 发票的 100 件、fragile goods、Express air、`$24.99` shipping 和 Pottery & Co. payee 信息。
- 每页标注两个输入 PDF 的来源页码；不得补充外部供应商、付款或会计系统数据。
- 所有文本、表格和图表必须原生可编辑，无重叠、无裁切、无整页压平。

输出文件必须命名为 `/app/output/January_2023_Invoice_Payment_Review.pptx`。请仅在 `/app/output` 中写入最终交付文件。
