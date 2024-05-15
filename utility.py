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
        positions.volume as Volume
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
    positions = pd.merge(df, stocks[['bloomberg_query', 'Last Price']],
                         left_on='Query', right_on='bloomberg_query',
                         how='left')
    positions.set_index(['Name', 'Position Name'], inplace=True)
    positions['% since AEQ'] = pd.to_numeric(((positions['Last Price'] - positions['AEQ']) / positions['AEQ']) * 100, errors='coerce')
    return positions


def get_stocks_data() -> pd.DataFrame:
    df = pd.read_excel('single-stocks.xlsx', sheet_name='Stocks', header=0)
    mask = df['Currency'] == 'GBp'
    df.loc[mask, 'Currency'] = 'GBP'
    df.loc[mask, 'Last Price'] /= 100
    return df


def style_positions_with_bars(positions: pd.DataFrame, name: str) -> str:
    columns_to_show = ['Position Name', 'AEQ', 'Currency', 'Volume', 'Last Price', '% since AEQ']
    positions = positions.copy().reset_index()[columns_to_show]
    max_abs_value = max(abs(positions['% since AEQ'].min().min()), abs(positions['% since AEQ'].max().max()))
    positions = positions.sort_values(by='% since AEQ', ascending=False)
    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))
    styled = (positions.style.bar(subset='% since AEQ', cmap=cm, align=0, vmax=max_abs_value, vmin=-max_abs_value)
              .format({
                    'AEQ': "{:,.2f}",
                    'Last Price': "{:,.2f}",
                    'Volume': "{:,.0f}",
                    '% since AEQ': "{:.2f}%"
              }).hide(axis="index"))

    output_path = f"output/images/{name}_Details.png"
    dfi.export(styled, output_path, table_conversion="selenium")
    return output_path
