# Board Stock Analysis — Presenter Script

Good morning, board members. Thank you for being present. Today I will walk through a source-bounded review of the stock data from 4 January 2010 through 4 November 2024. We cleaned 3,740 raw rows by removing the five invalid `FakeDate` rows with missing required fields, leaving 3,735 dated observations. The discussion is historical analysis, not investment advice or a forecast.

## Slide 1 — Closing Prices Over Time

The closing price rose from $30.95 on 4 January 2010 to $408.46 on 4 November 2024, a cumulative change of 1,219.74%. The series reached an observed minimum close of $23.01 and an observed maximum close of $467.56. For the board, this establishes strong long-run appreciation but does not imply that the same trajectory will continue.

## Slide 2 — Close Price with 50 and 200-Day Moving Averages

The latest close is $408.46, compared with a 50-day moving average of $420.52 and a 200-day moving average of $420.88. The close is therefore below both reference trends at the end of the sample. This is a monitoring signal for trend context, not a standalone buy or sell recommendation.

## Slide 3 — Correlation Between Columns

The price variables—Adjusted Close, Close, High, Low, and Open—are almost perfectly correlated, at approximately 1.00. Close and Volume have a correlation of -0.37, while Daily Return and Volume have a correlation of -0.07. For planning, these relationships help identify which fields move together, but correlation does not establish causation.

## Slide 4 — Daily Return Distribution

The mean daily return is 0.08%. Annualized volatility, using 252 trading days, is 25.66%. The observed one-day range runs from -14.74% to +14.22%. Most daily outcomes are concentrated near zero, while the tails show the risk of occasional large moves that should be reflected in scenario and liquidity planning.

## Slide 5 — Volume Distribution

Median daily volume is 30.38 million shares, mean daily volume is 36.32 million, and the observed maximum is 319.32 million. The distribution is right-skewed: a small number of sessions carry exceptionally high activity. For the board, these episodes are useful prompts for event review; volume alone does not identify the cause.

That concludes the five-chart review. Thank you for your time and for being present at the meeting. I welcome questions on the cleaning method, trend context, risk distribution, or the evidence limits of this analysis.
