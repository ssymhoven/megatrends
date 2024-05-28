from utility import mandate, get_trades, style_trades_with_bars

if __name__ == '__main__':
    "TODO: Fix XRate"

    for name, id in mandate.items():
        trades = get_trades(account_id=id)
        style_trades_with_bars(trades=trades, name=name)