import os
from datetime import datetime, timedelta
from typing import Dict

import pandas as pd
import win32com.client as win32
from matplotlib.colors import LinearSegmentedColormap
from source_engine.opus_source import OpusSource
import dataframe_image as dfi

output_dir = "output"
os.makedirs(os.path.join(output_dir, "images"), exist_ok=True)

mail_data = {
    'files': list(),
    'positions': {},
}
mail = True

mandate = {
    'D&R Aktien': '17154503',
    'D&R Aktien Nachhaltigkeit': '79939521',
    'D&R Aktien Strategie': '399347',
    'VV-Aktien Aktiv': '93695431'
}

query = f"""
    SELECT
        accountsegment_id,
        account_id,
        accountsegments.name as Name, 
        reportings.report_date, 
        positions.name as 'Position Name',
        positions.isin as ISIN,
        positions.bloomberg_query as Query,
        positions.average_entry_quote as AEQ,
        positions.average_entry_xrate as AEX,
        positions.currency as Crncy,
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
            AND accountsegments.account_id in ({', '.join(mandate.values())})
            AND reportings.report_date = (SELECT
                                            MAX(report_date)
                                          FROM
                                            reportings)
    """

third_party = """
    SELECT
        accountsegments.name as Name, 
        accountsegments.account_id,
        reportings.report_date, 
        positions.bloomberg_query as Query,
        positions.name as 'Position Name',
        positions.average_entry_quote as AEQ,
        positions.volume as Volume,
        positions.last_xrate_quantity as AEX
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
            AND positions.asset_class = 'FUND_CLASS'
            AND positions.dr_class_level_1 = 'EQUITY'
            AND (accountsegments.name LIKE "%VV-ESG%" OR accountsegments.name LIKE "%VV-Flex%")
            AND reportings.report_date = (SELECT
                                            MAX(report_date)
                                          FROM
                                            reportings)
"""

trades = """
    WITH depot AS (SELECT
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
        AND positions.asset_class = 'STOCK'
        AND report_date > '2024-01-01'
        AND accountsegments.account_id = '{account_id}'
    ),
    orders AS (SELECT
            account_id,
            trade_date,
            asset_isin,
            order_action,
            volume_quantity,
            price_quantity,
            xrate_quantity
        FROM 
            confirmations
        WHERE 
            account_id = '{account_id}'
            AND asset_class = 'STOCK'
            AND order_action = 'SELL_CLOSE'
            AND valuta_date > '2024-01-01'
    )
    SELECT
        trade_date as 'Trade Date',
        position_name as 'Position Name',
        asset_isin as 'ISIN',
        average_entry_quote as 'AEQ',
        average_entry_xrate as 'AEX',
        order_action as 'Action',
        volume_quantity as 'Volume',
        price_quantity as 'Price',
        xrate_quantity as 'XRate'
    FROM depot
    RIGHT JOIN orders ON depot.report_date = orders.trade_date AND depot.isin = orders.asset_isin
    ORDER BY depot.report_date;
"""

opus = OpusSource()


def get_positions() -> pd.DataFrame:
    df = opus.read_sql(query=query)
    df['AEQ'] = df['AEQ'] * df['AEX']

    positions = pd.merge(df, stocks[['bloomberg_query', 'Last Price', '1D', '5D', '1MO', 'YTD']],
                         left_on='Query', right_on='bloomberg_query',
                         how='left')

    positions.set_index(['Name', 'Position Name'], inplace=True)
    positions['% since AEQ'] = pd.to_numeric(((positions['Last Price'] - positions['AEQ']) / positions['AEQ']) * 100, errors='coerce')
    return positions


def get_third_party_products() -> pd.DataFrame:
    df = opus.read_sql(query=third_party)
    df['AEQ'] = df['AEQ'] * df['AEX']

    df = pd.merge(df, funds[['bloomberg_query', 'Last Price', '1D', '5D', '1MO', 'YTD']],
                  left_on='Query', right_on='bloomberg_query',
                  how='left')

    df.set_index(['Name', 'Position Name'], inplace=True)
    df['% since AEQ'] = pd.to_numeric(((df['Last Price'] - df['AEQ']) / df['AEQ']) * 100,
                                      errors='coerce')

    return df


def calc_universe_rel_performance_vs_sector(universe: pd.DataFrame, sector: pd.DataFrame) -> pd.DataFrame:
    sector_mapping = sector.index.to_series().str.extract(r'(\d+)\s*(.*)')
    sector_mapping.columns = ['Sector_Number', 'Cleaned_Sector']
    sector_mapping['Full_Sector'] = sector.index
    sector_mapping_dict = sector_mapping.set_index('Cleaned_Sector')['Full_Sector'].to_dict()

    universe['Sector'] = universe['Sector'].map(sector_mapping_dict)

    def calculate_difference(row, sector):
        sector_row = sector.loc[row['Sector']]

        for time_frame in ['1D', '5D', '1MO', 'YTD']:
            row[f'{time_frame} vs. Sector'] = row[time_frame] - sector_row[time_frame]

        return row

    universe = universe.apply(
        lambda row: calculate_difference(row, sector), axis=1
    )

    return universe


def calc_position_rel_performance_vs_sector(positions: pd.DataFrame, us: pd.DataFrame, eu: pd.DataFrame) -> pd.DataFrame:

    def calculate_difference(row, benchmark_df):
        benchmark_row = benchmark_df.loc[row['Sector']]

        for time_frame in ['1D', '5D', '1MO', 'YTD']:
            row[f'{time_frame} vs. Sector'] = row[time_frame] - benchmark_row[time_frame]

        return row

    positions = positions.apply(
        lambda row: calculate_difference(row, eu if row['Region'] == 'EU' else us), axis=1
    )

    return positions


def filter_positions(positions: pd.DataFrame, sector: str = None) -> (pd.DataFrame, pd.DataFrame):
    def get_quantiles(row):
        if sector:
            return us_quantiles if sector == 'US' else eu_quantiles
        else:
            return us_quantiles if row['Region'] == 'US' else eu_quantiles

    positives = []
    negatives = []

    for _, row in positions.iterrows():
        quantiles = get_quantiles(row)
        if sector:
            pos_condition = (
                    ((row['1D vs. Sector'] > quantiles.loc['1D vs. Sector', '99th Quantile']) |
                    (row['5D vs. Sector'] > quantiles.loc['5D vs. Sector', '99th Quantile']) |
                    (row['1MO vs. Sector'] > quantiles.loc['1MO vs. Sector', '99th Quantile']) |
                    (row['YTD vs. Sector'] > quantiles.loc['YTD vs. Sector', '99th Quantile'])) &
                    ((row['1D'] > quantiles.loc['1D', '99th Quantile']) |
                    (row['5D'] > quantiles.loc['5D', '99th Quantile']) |
                    (row['1MO'] > quantiles.loc['1MO', '99th Quantile']) |
                    (row['YTD'] > quantiles.loc['YTD', '99th Quantile']))
            )
            neg_condition = (
                    ((row['1D vs. Sector'] < quantiles.loc['1D vs. Sector', '1th Quantile']) |
                    (row['5D vs. Sector'] < quantiles.loc['5D vs. Sector', '1th Quantile']) |
                    (row['1MO vs. Sector'] < quantiles.loc['1MO vs. Sector', '1th Quantile']) |
                    (row['YTD vs. Sector'] < quantiles.loc['YTD vs. Sector', '1th Quantile'])) &
                    ((row['1D'] < quantiles.loc['1D', '1th Quantile']) |
                    (row['5D'] < quantiles.loc['5D', '1th Quantile']) |
                    (row['1MO'] < quantiles.loc['1MO', '1th Quantile']) |
                    (row['YTD'] < quantiles.loc['YTD', '1th Quantile']))
            )
        else:
            pos_condition = (
                    (row['1D vs. Sector'] > quantiles.loc['1D vs. Sector', '99th Quantile']) |
                    (row['5D vs. Sector'] > quantiles.loc['5D vs. Sector', '99th Quantile']) |
                    (row['1MO vs. Sector'] > quantiles.loc['1MO vs. Sector', '99th Quantile']) |
                    (row['YTD vs. Sector'] > quantiles.loc['YTD vs. Sector', '99th Quantile'])
            )
            neg_condition = (
                    ((row['1D vs. Sector'] < quantiles.loc['1D vs. Sector', '1th Quantile']) |
                    (row['5D vs. Sector'] < quantiles.loc['5D vs. Sector', '1th Quantile']) |
                    (row['1MO vs. Sector'] < quantiles.loc['1MO vs. Sector', '1th Quantile']) |
                    (row['YTD vs. Sector'] < quantiles.loc['YTD vs. Sector', '1th Quantile'])) | (row['% since AEQ'] < -10)
            )

        if pos_condition:
            positives.append(row)
        if neg_condition:
            negatives.append(row)

    positive_positions = pd.DataFrame(positives)
    negative_positions = pd.DataFrame(negatives)

    return positive_positions, negative_positions


def get_trades(account_id: str) -> pd.DataFrame:
    df = opus.read_sql(query=trades.format(account_id=account_id))
    df['AEQ'] = df['AEQ'] * df['AEX']
    df['Price'] = df['Price'] * df['XRate']
    df['% since AEQ'] = ((df['Price'] - df['AEQ']) / df['AEQ']) * 100
    return df


def calculate_quantiles(df: pd.DataFrame, columns: list) -> pd.DataFrame:
    quantiles = {}
    for column in columns:
        quantiles[column] = {
            '1th Quantile': df[column].quantile(0.01),
            '99th Quantile': df[column].quantile(0.99)
        }
    df = pd.DataFrame(quantiles).transpose()
    df = df.apply(lambda x: round(x * 2) / 2)
    return df


def get_universe_data(universe: str) -> pd.DataFrame:
    universe = pd.read_excel('stocks.xlsx', sheet_name=universe, header=0)
    universe.fillna(0, inplace=True)
    universe = universe.rename(
        columns={'name': 'Name', 'gics_sector_name': 'Sector', 'CURRENT_TRR_1D': '1D',
                 'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})
    return universe


def get_stocks_data() -> pd.DataFrame:
    df = pd.read_excel('stocks.xlsx', sheet_name='Stocks', header=0)
    df.fillna(0, inplace=True)
    df = df.rename(columns={'CURRENT_TRR_1D': '1D', 'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})
    return df


def get_funds_data() -> pd.DataFrame:
    df = pd.read_excel('stocks.xlsx', sheet_name='Funds', header=0)
    df.fillna(0, inplace=True)
    df = df.rename(
        columns={'CURRENT_TRR_1D': '1D', 'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})
    return df


def get_us_sector_data() -> pd.DataFrame:
    df = pd.read_excel('stocks.xlsx', sheet_name='US Sector', header=0, index_col=0)
    df.drop('Query', inplace=True, axis=1)
    df.fillna(0, inplace=True)
    df = df.rename(columns={'CURRENT_TRR_1D': '1D', 'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})
    return df


def get_eu_sector_data() -> pd.DataFrame:
    df = pd.read_excel('stocks.xlsx', sheet_name='EU Sector', header=0)
    df.drop('Query', inplace=True, axis=1)
    df.fillna(0, inplace=True)
    df = df.rename(columns={'CURRENT_TRR_1D': '1D', 'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})

    def calculate_weighted_trrs(group):
        weight = group['CUR_MKT_CAP'] / group['CUR_MKT_CAP'].sum()
        d = {'1D': (group['1D'] * weight).sum(),
             '5D': (group['5D'] * weight).sum(),
             '1MO': (group['1MO'] * weight).sum(),
             'YTD': (group['YTD'] * weight).sum()}
        return pd.Series(d)

    df = df.groupby('GICS').apply(calculate_weighted_trrs)
    return df


def calc_sector_diff(us: pd.DataFrame, eu: pd.DataFrame) -> pd.DataFrame:
    common_columns = us.columns.intersection(eu.columns)
    common_index = us.index.intersection(eu.index)

    diff = pd.DataFrame(index=common_index, columns=common_columns)

    for col in common_columns:
        for idx in common_index:
            diff.at[idx, col] = eu.at[idx, col] - us.at[idx, col]

    last_row_diff = eu.loc[eu.index[-1]] - us.loc[us.index[-1]]
    name = f"{eu.index[-1]} - {us.index[-1]}"
    diff = pd.concat([diff, pd.DataFrame([last_row_diff.values],
                                         columns=common_columns,
                                         index=[name])])
    diff.index.name = 'GICS'
    return diff


def style_universe_with_bars(positions: pd.DataFrame, name: str) -> str:

    trr1d_max_abs_value = max(abs(positions['1D'].min().min()),
                              abs(positions['1D'].max().max()))
    trr5d_max_abs_value = max(abs(positions['5D'].min().min()),
                              abs(positions['5D'].max().max()))
    trr1mo_max_abs_value = max(abs(positions['1MO'].min().min()),
                               abs(positions['1MO'].max().max()))
    trr_ytd_max_abs_value = max(abs(positions['YTD'].min().min()),
                                abs(positions['YTD'].max().max()))

    rel_trr1d_max_abs_value = max(abs(positions['1D vs. Sector'].min().min()),
                                  abs(positions['1D vs. Sector'].max().max()))
    rel_trr5d_max_abs_value = max(abs(positions['5D vs. Sector'].min().min()),
                                abs(positions['5D vs. Sector'].max().max()))
    rel_trr_trr1mo_max_abs_value = max(abs(positions['1MO vs. Sector'].min().min()),
                                abs(positions['1MO vs. Sector'].max().max()))
    rel_trr_ytd_max_abs_value = max(abs(positions['YTD vs. Sector'].min().min()),
                                abs(positions['YTD vs. Sector'].max().max()))

    positions['Highlight'] = positions['ID'].isin(stocks['bloomberg_query'])

    positions.set_index('ID', inplace=True)
    positions.sort_values('YTD vs. Sector', ascending=False, inplace=True)

    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))

    def highlight_index(s):
        return ['background-color: #D1FFBD' if val else '' for val in s]

    styled = (positions.style
              .apply(highlight_index, subset=pd.IndexSlice[positions.index[positions['Highlight']], :], axis=0)
              .bar(subset='1D', cmap=cm, align=0, vmax=trr1d_max_abs_value, vmin=-trr1d_max_abs_value)
              .bar(subset='5D', cmap=cm, align=0, vmax=trr5d_max_abs_value, vmin=-trr5d_max_abs_value)
              .bar(subset='1MO', cmap=cm, align=0, vmax=trr1mo_max_abs_value, vmin=-trr1mo_max_abs_value)
              .bar(subset='YTD', cmap=cm, align=0, vmax=trr_ytd_max_abs_value, vmin=-trr_ytd_max_abs_value)
              .bar(subset='1D vs. Sector', cmap=cm, align=0, vmax=rel_trr1d_max_abs_value,
                   vmin=-rel_trr1d_max_abs_value)
              .bar(subset='5D vs. Sector', cmap=cm, align=0, vmax=rel_trr5d_max_abs_value, vmin=-rel_trr5d_max_abs_value)
              .bar(subset='1MO vs. Sector', cmap=cm, align=0, vmax=rel_trr_trr1mo_max_abs_value, vmin=-rel_trr_trr1mo_max_abs_value)
              .bar(subset='YTD vs. Sector', cmap=cm, align=0, vmax=rel_trr_ytd_max_abs_value, vmin=-rel_trr_ytd_max_abs_value)
              .set_table_styles([
                    {'selector': 'th.col0',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'td.col0',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'th.col2',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'td.col2',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'th.col6',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'td.col6',
                     'props': [('border-left', '1px solid black')]},
                    {
                        'selector': 'th.index_name',
                        'props': [('min-width', '150px'), ('white-space', 'nowrap')]
                    },
                    {
                        'selector': 'td.col0',
                        'props': [('min-width', '200px'), ('white-space', 'nowrap')]
                    },
                    {
                        'selector': 'td.col1',
                        'props': [('min-width', '150px'), ('white-space', 'nowrap')]
                    }
              ])
              .format({
                    '1D': "{:.2f}%",
                    '5D': "{:.2f}%",
                    '1MO': "{:.2f}%",
                    'YTD': "{:.2f}%",
                    '1D vs. Sector': "{:.2f}%",
                    '5D vs. Sector': "{:.2f}%",
                    '1MO vs. Sector': "{:.2f}%",
                    'YTD vs. Sector': "{:.2f}%",
              })).hide(['Highlight'], axis='columns')

    output_path = f"output/images/{name}.png"
    dfi.export(styled, output_path, table_conversion="selenium", max_rows=-1)
    return output_path


def style_third_party(positions: pd.DataFrame, name: str) -> str:
    aeq_max_abs_value = max(abs(positions['% since AEQ'].min().min()),
                            abs(positions['% since AEQ'].max().max()))
    trr1d_max_abs_value = max(abs(positions['1D'].min().min()),
                              abs(positions['1D'].max().max()))
    trr5d_max_abs_value = max(abs(positions['5D'].min().min()),
                              abs(positions['5D'].max().max()))
    trr1mo_max_abs_value = max(abs(positions['1MO'].min().min()),
                               abs(positions['1MO'].max().max()))
    trr_ytd_max_abs_value = max(abs(positions['YTD'].min().min()),
                                abs(positions['YTD'].max().max()))

    output_path = f"output/images/{name}.png"
    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))
    positions.sort_values('% since AEQ', ascending=False, inplace=True)

    styled = ((((positions.style.bar(subset='% since AEQ', cmap=cm, align=0, vmax=aeq_max_abs_value, vmin=-aeq_max_abs_value)
          .bar(subset='1D', cmap=cm, align=0, vmax=trr1d_max_abs_value, vmin=-trr1d_max_abs_value))
          .bar(subset='5D', cmap=cm, align=0, vmax=trr5d_max_abs_value, vmin=-trr5d_max_abs_value))
          .bar(subset='1MO', cmap=cm, align=0, vmax=trr1mo_max_abs_value, vmin=-trr1mo_max_abs_value))
          .bar(subset='YTD', cmap=cm, align=0, vmax=trr_ytd_max_abs_value, vmin=-trr_ytd_max_abs_value)
          .set_table_styles([
                {'selector': 'th.col0',
                 'props': [('border-left', '1px solid black')]},
                {'selector': 'td.col0',
                 'props': [('border-left', '1px solid black')]},
                {'selector': 'th.col4',
                 'props': [('border-left', '1px solid black')]},
                {'selector': 'td.col4',
                 'props': [('border-left', '1px solid black')]},
                {
                    'selector': 'th.index_name',
                    'props': [('min-width', '250px'), ('white-space', 'nowrap')]
                }
          ]).format({
                'AEQ': "{:,.2f}",
                'Last Price': "{:,.2f}",
                '% since AEQ': "{:.2f}%",
                '1D': "{:.2f}%",
                '5D': "{:.2f}%",
                '1MO': "{:.2f}%",
                'YTD': "{:.2f}%"
          }))
    dfi.export(styled, output_path, table_conversion="selenium")
    return output_path


def style_positions_with_bars(positions: pd.DataFrame, name: str) -> str:
    columns_to_show = ['Sector', 'AEQ', 'Volume', 'Last Price', '% since AEQ', '1D', '5D', '1MO', 'YTD',
                       '1D vs. Sector', '5D vs. Sector', '1MO vs. Sector', 'YTD vs. Sector']
    positions = positions.copy()[columns_to_show]

    aeq_max_abs_value = max(abs(positions['% since AEQ'].min().min()),
                            abs(positions['% since AEQ'].max().max()))
    trr1d_max_abs_value = max(abs(positions['1D'].min().min()),
                              abs(positions['1D'].max().max()))
    trr5d_max_abs_value = max(abs(positions['5D'].min().min()),
                              abs(positions['5D'].max().max()))
    trr1mo_max_abs_value = max(abs(positions['1MO'].min().min()),
                               abs(positions['1MO'].max().max()))
    trr_ytd_max_abs_value = max(abs(positions['YTD'].min().min()),
                                abs(positions['YTD'].max().max()))

    rel_trr1d_max_abs_value = max(abs(positions['1D vs. Sector'].min().min()),
                                  abs(positions['1D vs. Sector'].max().max()))
    rel_trr5d_max_abs_value = max(abs(positions['5D vs. Sector'].min().min()),
                                abs(positions['5D vs. Sector'].max().max()))
    rel_trr_trr1mo_max_abs_value = max(abs(positions['1MO vs. Sector'].min().min()),
                                abs(positions['1MO vs. Sector'].max().max()))
    rel_trr_ytd_max_abs_value = max(abs(positions['YTD vs. Sector'].min().min()),
                                abs(positions['YTD vs. Sector'].max().max()))

    positions.sort_values('% since AEQ', ascending=False, inplace=True)
    positions.index.name = 'Position Name'

    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions) if len(positions) != 0 else 3)

    styled = (positions.style.bar(subset='% since AEQ', cmap=cm, align=0, vmax=aeq_max_abs_value, vmin=-aeq_max_abs_value)
              .bar(subset='1D', cmap=cm, align=0, vmax=trr1d_max_abs_value, vmin=-trr1d_max_abs_value)
              .bar(subset='5D', cmap=cm, align=0, vmax=trr5d_max_abs_value, vmin=-trr5d_max_abs_value)
              .bar(subset='1MO', cmap=cm, align=0, vmax=trr1mo_max_abs_value, vmin=-trr1mo_max_abs_value)
              .bar(subset='YTD', cmap=cm, align=0, vmax=trr_ytd_max_abs_value, vmin=-trr_ytd_max_abs_value)
              .bar(subset='1D vs. Sector', cmap=cm, align=0, vmax=rel_trr1d_max_abs_value,
                   vmin=-rel_trr1d_max_abs_value)
              .bar(subset='5D vs. Sector', cmap=cm, align=0, vmax=rel_trr5d_max_abs_value, vmin=-rel_trr5d_max_abs_value)
              .bar(subset='1MO vs. Sector', cmap=cm, align=0, vmax=rel_trr_trr1mo_max_abs_value, vmin=-rel_trr_trr1mo_max_abs_value)
              .bar(subset='YTD vs. Sector', cmap=cm, align=0, vmax=rel_trr_ytd_max_abs_value, vmin=-rel_trr_ytd_max_abs_value)
              .set_table_styles([
                    {'selector': 'th.col0',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'td.col0',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'th.col5',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'td.col5',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'th.col9',
                     'props': [('border-left', '1px solid black')]},
                    {'selector': 'td.col9',
                     'props': [('border-left', '1px solid black')]},
                    {
                        'selector': 'th.index_name',
                        'props': [('min-width', '250px'), ('white-space', 'nowrap')]
                    },
                    {
                        'selector': 'td.col0',
                        'props': [('min-width', '200px'), ('white-space', 'nowrap')]
                    }
              ])
              .format({
                    'AEQ': "{:,.2f}",
                    'Last Price': "{:,.2f}",
                    'Volume': "{:,.0f}",
                    '% since AEQ': "{:.2f}%",
                    '1D': "{:.2f}%",
                    '5D': "{:.2f}%",
                    '1MO': "{:.2f}%",
                    'YTD': "{:.2f}%",
                    '1D vs. Sector': "{:.2f}%",
                    '5D vs. Sector': "{:.2f}%",
                    '1MO vs. Sector': "{:.2f}%",
                    'YTD vs. Sector': "{:.2f}%",
              }))

    output_path = f"output/images/{name.replace(' ', '_')}_Details.png"
    dfi.export(styled, output_path, table_conversion="selenium")
    return output_path


def style_trades_with_bars(trades: pd.DataFrame, name: str) -> str:
    max_abs_value = trades['% since AEQ'].abs().max()
    trades.sort_values('Trade Date', ascending=False, inplace=True)
    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(trades))
    styled = (
        trades.style.bar(subset='% since AEQ', cmap=cm, align=0, vmax=max_abs_value, vmin=-max_abs_value)
        .format({
            'AEQ': "{:,.2f}",
            'Price': "{:,.2f}",
            'Volume': "{:,.0f}",
            '% since AEQ': "{:.2f}%"
        })).hide(axis='index')
    output_path = f"output/images/{name.replace(' ', '_')}_Trades.png"
    dfi.export(styled, output_path, table_conversion="selenium")
    return output_path


def style_index_with_bars(index: pd.DataFrame, name: str) -> str:
    for col in ['1D', '5D', '1MO', 'YTD']:
        index[col] = pd.to_numeric(index[col], errors='coerce')

    trr1d_max_abs_value = index['1D'].abs().max()
    trr5d_max_abs_value = index['5D'].abs().max()
    trr1mo_max_abs_value = index['1MO'].abs().max()
    trr_ytd_max_abs_value = index['YTD'].abs().max()

    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(index))
    styled = (
        index.style.bar(subset='1D', cmap=cm, align=0, vmax=trr1d_max_abs_value, vmin=-trr1d_max_abs_value)
        .bar(subset='5D', cmap=cm, align=0, vmax=trr5d_max_abs_value, vmin=-trr5d_max_abs_value)
        .bar(subset='1MO', cmap=cm, align=0, vmax=trr1mo_max_abs_value, vmin=-trr1mo_max_abs_value)
        .bar(subset='YTD', cmap=cm, align=0, vmax=trr_ytd_max_abs_value, vmin=-trr_ytd_max_abs_value)
        .set_table_styles([
            {'selector': 'th.col0',
             'props': [('border-left', '1px solid black')]},
            {'selector': 'td.col0',
             'props': [('border-left', '1px solid black')]},
            {'selector': 'tr:last-child th, tr:last-child td',
             'props': [('border-top', '1px solid black')]}
        ])
        .format({
            '1D': "{:,.2f}",
            '5D': "{:,.2f}",
            '1MO': "{:,.2f}",
            'YTD': "{:,.2f}"
        }))

    output_path = f"output/images/{name.replace(' ', '_')}_Details.png"
    dfi.export(styled, output_path, table_conversion="selenium")
    return output_path


def get_last_business_day():
    today = datetime.now()

    if today.weekday() == 0:
        last_business_day = today - timedelta(days=3)
    elif today.weekday() == 6:
        last_business_day = today - timedelta(days=2)
    else:
        last_business_day = today - timedelta(days=1)

    last_business_day_str = last_business_day.strftime('%d.%m.%Y')

    return last_business_day_str


def group_funds(positions: pd.DataFrame) -> pd.DataFrame:
    positions.reset_index(inplace=True)

    positions = positions.groupby('Position Name').agg({
        '1D': 'first',
        '5D': 'first',
        '1MO': 'first',
        'YTD': 'first',
        'Last Price': 'first',
        'AEQ': 'mean',
        '% since AEQ': 'mean'
    })

    return positions


def write_risk_mail(data: Dict):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Subject = "Daily Reporting - Risikomanagement"
    mail.Recipients.Add("pm-aktien")
    mail.Recipients.Add("amstatuser@donner-reuschel.lu")
    mail.Recipients.Add("jan.sandermann@donner-reuschel.de")
    mail.Recipients.Add("sadettin.yildiz@donner-reuschel.de").Type = 2

    def inplace_chart(image_path: str):
        image_path = os.path.abspath(image_path)
        attachment = mail.Attachments.Add(Source=image_path)
        cid = os.path.basename(image_path)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
        return cid

    sector_images_html = ''
    for key, image_path in data.get('sector', {}).items():
        cid = inplace_chart(image_path)
        sector_images_html += f'<p><b>{key}</b><br><img src="cid:{cid}"><br></p>'

    position_images_html = ''
    for key, image_path in data.get('positions', {}).items():
        cid = inplace_chart(image_path)
        position_images_html += f'<p><b>{key}</b><br><img src="cid:{cid}"><br></p>'

    mail.HTMLBody = f"""
        <html>
          <head></head>
          <body>
            <p>Hi zusammen, <br><br>
                hier sind absoluten Entwicklungen der <b>Sektoren</b> für den <b>Euro Stoxx 600</b> und den <b>S&P 500</b>, 
                sowie die Out/Underperformance des <b>Euro Stoxx 600</b> gegenüber dem <b>S&P 500</b>.<br><br> 
                Alle Kurse in EUR, Kursreferenz: Letzter Preis am {get_last_business_day()}:<br><br>
                {sector_images_html}
                <br><br>
                
                Hier sind die <b>Positionen</b>, die sich schlechter als der jeweilige Sektor entwickelt haben. 
                Folgende Schwellenwerte werden dabei berücksichtigt, basierend auf dem <b>1. Perzentil</b> der entsprechenden Daten:<br><br>
                <b>US</b><br>
                {us_metrics_positions}<br><br>
                <b>EU</b><br>
                {eu_metrics_positions},<br><br>
                oder seit Kauf mehr als <b>10%</b> verloren haben.<br><br>
                
                {position_images_html}
                <br><br>
                Im Anhang findet Ihr für alle Fonds eine detaillierte Übersicht aller Positionen.
                <br><br>
                Liebe Grüße
            </p>
          </body>
        </html>
    """

    for file_path in data.get('files', []):
        if os.path.exists(file_path):
            mail.Attachments.Add(Source=os.path.abspath(file_path))

    mail.Recipients.ResolveAll()
    mail.Display(True)


def write_universe_mail(data: Dict):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Subject = "Daily Reporting - Investment Universe Momentum"
    mail.Recipients.Add("pm-aktien")
    mail.Recipients.Add("amstatuser@donner-reuschel.lu")
    mail.Recipients.Add("jan.sandermann@donner-reuschel.de")
    mail.Recipients.Add("sadettin.yildiz@donner-reuschel.de").Type = 2

    def inplace_chart(image_path: str):
        image_path = os.path.abspath(image_path)
        attachment = mail.Attachments.Add(Source=image_path)
        cid = os.path.basename(image_path)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
        return cid

    mail.HTMLBody = f"""
        <html>
          <head></head>
          <body>
            <p>Hi zusammen, <br><br>
                hier ist eine Momentumanalyse aller Titel des <b>Euro Stoxx 600</b> und des <b>S&P 500</b>.
                Alle Kurse in EUR, Kursreferenz: Letzter Preis am {get_last_business_day()}.<br><br>
                
                Folgende Schwellenwerte werden dabei berücksichtigt, basierend auf dem <b>99. Perzentil</b> der entsprechenden Daten: <br><br>
                <b>US</b><br>
                {us_metrics_sector}<br><br>
                <b>EU</b><br>
                {eu_metrics_sector}<br><br>
                
                Grün hinterlegte Positionen befinden sich bereits in einem der Fonds.<br><br>
                
                <b>S&P 500 Universum</b><br><br>
                 <p><img src="cid:{inplace_chart(data.get('us_positive'))}"><br></p><br><br>
                
                <b>STOXX Europe 600 Universum</b><br><br>
                <p><img src="cid:{inplace_chart(data.get('eu_positive'))}"><br></p>
                <br><br>
                Liebe Grüße
            </p>
          </body>
        </html>
    """

    mail.Recipients.ResolveAll()
    mail.Display(True)


def write_third_party_mail(data: Dict):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Subject = "Daily Reporting - VV Drittprodukte"
    mail.Recipients.Add("pm-aktien")
    mail.Recipients.Add("amstatuser@donner-reuschel.lu")
    mail.Recipients.Add("jan.sandermann@donner-reuschel.de")
    mail.Recipients.Add("sadettin.yildiz@donner-reuschel.de").Type = 2

    def inplace_chart(image_path: str):
        image_path = os.path.abspath(image_path)
        attachment = mail.Attachments.Add(Source=image_path)
        cid = os.path.basename(image_path)
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
        return cid

    mail.HTMLBody = f"""
        <html>
          <head></head>
          <body>
            <p>Hi zusammen, <br><br>
                
                folgende Drittprodukte aus der VV-Flex und VV-ESG haben sich wie folgt entwickelt. Der durchschnittliche 
                Einstiegskurs ist über alle Varianten gemittelt.<br><br>
                
                Alle Kurse in EUR, Kursreferenz: Letzter Preis am {get_last_business_day()}.<br><br>

                <b>VV-Flex</b><br><br>
                 <p><img src="cid:{inplace_chart(data.get('flex'))}"><br></p><br><br>

                <b>VV-ESG</b><br><br>
                <p><img src="cid:{inplace_chart(data.get('esg'))}"><br></p>
                <br><br>
                Liebe Grüße
            </p>
          </body>
        </html>
    """

    mail.Recipients.ResolveAll()
    mail.Display(True)


stocks = get_stocks_data()
funds = get_funds_data()

us_universe = get_universe_data(universe="S&P 500")
us_sector = get_us_sector_data()

eu_universe = get_universe_data(universe="STOXX Europe 600")
eu_sector = get_eu_sector_data()

us = calc_universe_rel_performance_vs_sector(universe=us_universe, sector=us_sector)
eu = calc_universe_rel_performance_vs_sector(universe=eu_universe, sector=eu_sector)

columns_to_analyze = ['1D', '5D', '1MO', 'YTD', '1D vs. Sector', '5D vs. Sector', '1MO vs. Sector', 'YTD vs. Sector']
us_quantiles = calculate_quantiles(us, columns_to_analyze)
eu_quantiles = calculate_quantiles(eu, columns_to_analyze)

us_metrics_positions = f"""
   1D vs. Sector < <b>{us_quantiles.loc['1D vs. Sector', '1th Quantile']}%</b>, oder<br>
   5D vs. Sector < <b>{us_quantiles.loc['5D vs. Sector', '1th Quantile']}%</b>, oder<br>
   1MO vs. Sector < <b>{us_quantiles.loc['1MO vs. Sector', '1th Quantile']}%</b>, oder<br>
   YTD vs. Sector < <b>{us_quantiles.loc['YTD vs. Sector', '1th Quantile']}%</b>
"""

eu_metrics_positions = f"""
   1D vs. Sector < <b>{eu_quantiles.loc['1D vs. Sector', '1th Quantile']}%</b>, oder<br>
   5D vs. Sector < <b>{eu_quantiles.loc['5D vs. Sector', '1th Quantile']}%</b>, oder<br>
   1MO vs. Sector < <b>{eu_quantiles.loc['1MO vs. Sector', '1th Quantile']}%</b>, oder<br>
   YTD vs. Sector < <b>{eu_quantiles.loc['YTD vs. Sector', '1th Quantile']}%</b>
"""

us_metrics_sector = f"""
   1D > <b>{us_quantiles.loc['1D', '99th Quantile']}%</b>, oder <br>
   5D > <b>{us_quantiles.loc['5D', '99th Quantile']}%</b>, oder <br>
   1MO > <b>{us_quantiles.loc['1MO', '99th Quantile']}%</b>, oder  <br>
   YTD > <b>{us_quantiles.loc['YTD', '99th Quantile']}%</b><br><br>
   und<br><br>
   1D vs. Sector > <b>{us_quantiles.loc['1D vs. Sector', '99th Quantile']}%</b>, oder<br>
   5D vs. Sector > <b>{us_quantiles.loc['5D vs. Sector', '99th Quantile']}%</b>, oder<br>
   1MO vs. Sector > <b>{us_quantiles.loc['1MO vs. Sector', '99th Quantile']}%</b>, oder<br>
   YTD vs. Sector > <b>{us_quantiles.loc['YTD vs. Sector', '99th Quantile']}%</b>
"""

eu_metrics_sector = f"""
   1D > <b>{eu_quantiles.loc['1D', '99th Quantile']}%</b>, oder<br>
   5D > <b>{eu_quantiles.loc['5D', '99th Quantile']}%</b>, oder<br>
   1MO > <b>{eu_quantiles.loc['1MO', '99th Quantile']}%</b>, oder<br>
   YTD > <b>{eu_quantiles.loc['YTD', '99th Quantile']}%</b><br><br>
   und<br><br>
   1D vs. Sector > <b>{eu_quantiles.loc['1D vs. Sector', '99th Quantile']}%</b>, oder<br>
   5D vs. Sector > <b>{eu_quantiles.loc['5D vs. Sector', '99th Quantile']}%</b>, oder<br>
   1MO vs. Sector > <b>{eu_quantiles.loc['1MO vs. Sector', '99th Quantile']}%</b>, oder<br>
   YTD vs. Sector > <b>{eu_quantiles.loc['YTD vs. Sector', '99th Quantile']}%</b>
"""
