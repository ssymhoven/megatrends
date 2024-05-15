import os

import imgkit
import pandas as pd
from matplotlib import pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from source_engine.opus_source import OpusSource
import seaborn as sns
import dataframe_image as dfi
from IPython.display import display

output_dir = "output"
os.makedirs(os.path.join(output_dir, "images"), exist_ok=True)

mandate = {
    'D&R Aktien': '17154631',
    'D&R Aktien Nachhaltigkeit': '79939969',
    'D&R Aktien Strategie': '399443'
}

query = f"""
    SELECT
        accountsegment_id,
        accountsegments.name as Name, 
        reportings.report_date, 
        positions.name as 'Position Name',
        positions.isin as ISIN,
        positions.bloomberg_query as Query,
        positions.average_entry_quote as AEQ,
        positions.currency as Currency,
        positions.volume as Volume,
        positions.gics_industry_sector as Sector,
        positions.dr_class_level_2 as 'Region'
    FROM
        reportings
            JOIN
        accountsegments ON (accountsegments.reporting_uuid = reportings.uuid)
            JOIN
        positions ON (reportings.uuid = positions.reporting_uuid)
    WHERE
            positions.account_segment_id = accountsegments.accountsegment_id
            AND reportings.newest = 1
            AND reportings.report = 'positions'
            AND positions.asset_class = 'STOCK'
            AND accountsegments.accountsegment_id in ({', '.join(mandate.values())})
            AND reportings.report_date = (SELECT
                                            MAX(report_date)
                                          FROM
                                            reportings)
    """

opus = OpusSource()


def get_positions() -> pd.DataFrame:
    df = opus.read_sql(query=query)
    stocks = get_stocks_data()
    positions = pd.merge(df, stocks[['bloomberg_query', 'Last Price', '5D', '1MO', 'YTD']],
                         left_on='Query', right_on='bloomberg_query',
                         how='left')
    positions.set_index(['Name', 'Position Name'], inplace=True)
    positions['% since AEQ'] = pd.to_numeric(((positions['Last Price'] - positions['AEQ']) / positions['AEQ']) * 100, errors='coerce')
    return positions


def get_stocks_data() -> pd.DataFrame:
    df = pd.read_excel('single-stocks.xlsx', sheet_name='Stocks', header=0)
    df = df.rename(columns={'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})
    mask = df['Currency'] == 'GBp'
    df.loc[mask, 'Currency'] = 'GBP'
    df.loc[mask, 'Last Price'] /= 100
    return df


def get_us_sector_data() -> pd.DataFrame:
    return pd.read_excel('single-stocks.xlsx', sheet_name='US Sector', header=0, index_col=0)


def get_eu_sector_data() -> pd.DataFrame:
    df = pd.read_excel('single-stocks.xlsx', sheet_name='EU Sector', header=0)
    df.drop('Query', inplace=True, axis=1)
    df = df.rename(columns={'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})

    def calculate_weighted_trrs(group):
        weight = group['CUR_MKT_CAP'] / group['CUR_MKT_CAP'].sum()
        d = {'5D': (group['5D'] * weight).sum(),
             '1MO': (group['1MO'] * weight).sum(),
             'YTD': (group['YTD'] * weight).sum()}
        return pd.Series(d)

    df = df.groupby('GICS').apply(calculate_weighted_trrs)
    return df


def style_positions_with_bars(positions: pd.DataFrame, name: str) -> str:
    columns_to_show = ['Position Name', 'AEQ', 'Currency', 'Volume', 'Last Price', '% since AEQ', '5D', '1MO', 'YTD']
    positions = positions.copy().reset_index()[columns_to_show]

    aeq_max_abs_value = max(abs(positions['% since AEQ'].min().min()),
                            abs(positions['% since AEQ'].max().max()))
    trr5d_max_abs_value = max(abs(positions['5D'].min().min()),
                              abs(positions['5D'].max().max()))
    trr1mo_max_abs_value = max(abs(positions['1MO'].min().min()),
                              abs(positions['1MO'].max().max()))
    trr_ytd_max_abs_value = max(abs(positions['YTD'].min().min()),
                              abs(positions['YTD'].max().max()))

    positions = positions.sort_values(by='% since AEQ', ascending=False)
    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))

    styled = (positions.style.bar(subset='% since AEQ', cmap=cm, align=0, vmax=aeq_max_abs_value, vmin=-aeq_max_abs_value)
              .bar(subset='5D', cmap=cm, align=0, vmax=trr5d_max_abs_value, vmin=-trr5d_max_abs_value)
              .bar(subset='1MO', cmap=cm, align=0, vmax=trr1mo_max_abs_value, vmin=-trr1mo_max_abs_value)
              .bar(subset='YTD', cmap=cm, align=0, vmax=trr_ytd_max_abs_value, vmin=-trr_ytd_max_abs_value)
              .format({
                    'AEQ': "{:,.2f}",
                    'Last Price': "{:,.2f}",
                    'Volume': "{:,.0f}",
                    '% since AEQ': "{:.2f}%",
                    '5D': "{:.2f}%",
                    '1MO': "{:.2f}%",
                    'YTD': "{:.2f}%",
              }).hide(axis="index"))

    output_path = f"output/images/{name}_Details.png"
    dfi.export(styled, output_path, table_conversion="selenium")
    return output_path
