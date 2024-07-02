WITH orders AS (SELECT
	trade_date,
    valuta_date,
	asset_isin,
    asset_name,
    asset_class,
	order_action,
	volume_quantity,
    volume_unit,
	price_quantity,
    price_unit,
	xrate_quantity,
    xrate_unit,
    exchange_rate_quantity,
    exchange_rate_unit
FROM
	confirmations
WHERE
	account_id = '17154503'
	AND trade_date > '2024-06-01'),
depot as (SELECT
	reportings.report_date,
	positions.isin,
	positions.name AS position_name,
	positions.average_entry_quote,
	positions.average_entry_xrate
FROM
	reportings
	JOIN accountsegments ON accountsegments.reporting_uuid = reportings.uuid
	JOIN positions ON reportings.uuid = positions.reporting_uuid
WHERE
	positions.account_segment_id = accountsegments.accountsegment_id
	AND reportings.newest = 1
	AND reportings.report = 'positions'
	AND report_date > '2024-06-01'
	AND accountsegments.account_id = '17154503')
SELECT
	report_date as 'Report Date',
    trade_date as 'Trade Date',
    valuta_date as 'Valuta Date',
    asset_name as 'Position Name',
    asset_class as 'Asset Class',
    order_action as 'Order Action',
    average_entry_quote as 'AEQ',
	average_entry_xrate as 'AEX',
    price_quantity as 'Price',
    price_unit as 'Price Unit',
	xrate_quantity as 'XRate',
    xrate_unit as 'XRate Unit'
FROM depot
RIGHT JOIN orders ON depot.isin = orders.asset_isin
    AND CASE
        WHEN orders.order_action IN ('SELL_CLOSE', 'BUY_CLOSE') THEN depot.report_date = orders.trade_date
        WHEN orders.order_action IN ('BUY_OPEN', 'SELL_OPEN') THEN depot.report_date = orders.valuta_date
    END
ORDER BY trade_date;
