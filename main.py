from utility import get_positions, style_positions_with_bars, get_us_sector_data, get_eu_sector_data, \
    style_index_with_bars, calc_rel_performance, calc_sector_diff, write_mail, mandate, get_trades, \
    style_trades_with_bars

threshold = -8

mail_data = {
    'files': list(),
    'positions': {},
}
mail = False


if __name__ == '__main__':
    positions = get_positions()
    us_sector = get_us_sector_data()
    eu_sector = get_eu_sector_data()
    diff = calc_sector_diff(us=us_sector, eu=eu_sector)

    positions = calc_rel_performance(positions=positions, eu=eu_sector, us=us_sector)

    unique_names = positions.index.get_level_values(0).unique()

    files = mail_data.get('files')
    p = mail_data.get('positions')

    for name in unique_names:
        subset = positions.loc[name]
        details_chart = style_positions_with_bars(positions=subset, name=name)
        files.append(details_chart)

        filtered_subset = subset[
            (subset['1D vs. Sector'] < threshold) |
            (subset['5D vs. Sector'] < threshold) |
            (subset['1MO vs. Sector'] < threshold) |
            (subset['YTD vs. Sector'] < threshold)
        ]

        underperformed_details_chart = style_positions_with_bars(positions=filtered_subset, name=f'{name}_underperformed')
        p.update({name: underperformed_details_chart})

    us_sector_chart = style_index_with_bars(index=us_sector, name='US')
    eu_sector_chart = style_index_with_bars(index=eu_sector, name='EU')
    diff_sector_chart = style_index_with_bars(index=diff, name='EU_vs_US')

    mail_data.update({
        'sector': {
            'SXXP Index': eu_sector_chart,
            'SPX Index': us_sector_chart,
            'SXXP Index vs. SPX Index': diff_sector_chart
        }
    })

    for name, id in mandate.items():
        trades = get_trades(account_id=id)
        style_trades_with_bars(trades=trades, name=name)

    if mail:
        write_mail(data=mail_data)
