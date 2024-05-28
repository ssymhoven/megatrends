from utility import style_universe_with_bars, filter_positions, mail_data, write_universe_mail, \
    mail, us, eu

if __name__ == '__main__':

    us_pos_subset, _ = filter_positions(positions=us, sector='US')
    eu_pos_subset, _ = filter_positions(positions=eu, sector='EU')

    us_positive = style_universe_with_bars(positions=us_pos_subset, name=f'us_universe_positive')
    eu_positive = style_universe_with_bars(positions=eu_pos_subset, name=f'eu_universe_positive')

    mail_data.update({
        'eu_positive': eu_positive,
        'us_positive': us_positive
    })

    if mail:
        write_universe_mail(data=mail_data)


