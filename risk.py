from utility import get_positions, style_positions_with_bars, get_us_sector_data, get_eu_sector_data, \
    style_index_with_bars, calc_position_rel_performance_vs_sector, calc_sector_diff, write_risk_mail, mail_data, \
    mail, filter_positions, us_sector, eu_sector

if __name__ == '__main__':
    positions = get_positions()

    diff = calc_sector_diff(us=us_sector, eu=eu_sector)

    positions = calc_position_rel_performance_vs_sector(positions=positions, eu=eu_sector, us=us_sector)

    unique_names = positions.index.get_level_values(0).unique()

    files = mail_data.get('files')
    p = mail_data.get('positions')

    for name in unique_names:
        subset = positions.loc[name]
        details_chart = style_positions_with_bars(positions=subset, name=name)
        files.append(details_chart)

        _, negative_positions = filter_positions(positions=subset)

        underperformed_details_chart = style_positions_with_bars(positions=negative_positions, name=f'{name}_underperformed')
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

    if mail:
        write_risk_mail(data=mail_data)
