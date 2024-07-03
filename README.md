```
=@BQL("filter(holdings('IQQH GY Equity');grouprank(#chg().value)<=10)";"#chg";"#chg=pct_chg(px_last(dates=range(2024-01-01,2024-07-03)))";"fill=prev";"cols=2;rows=11")
```