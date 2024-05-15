from utility import get_positions, style_positions_with_bars

if __name__ == '__main__':
    positions = get_positions()
    unique_names = positions.index.get_level_values(0).unique()

    for name in unique_names:
        subset = positions.loc[name]
        details_chart = style_positions_with_bars(positions=subset, name=name)

