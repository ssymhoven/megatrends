import os
from datetime import datetime
from typing import Dict

import pandas as pd
import win32com.client as win32
from matplotlib.colors import LinearSegmentedColormap
from source_engine.opus_source import OpusSource
import dataframe_image as dfi

output_dir = "output"
os.makedirs(os.path.join(output_dir, "images"), exist_ok=True)

mandate = {
    'D&R Aktien': '17154503',
    'D&R Aktien Nachhaltigkeit': '79939521',
    'D&R Aktien Strategie': '399347'
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


trades = """
    WITH depot AS (SELECT
        reportings.report_date, 
        positions.isin,
        positions.name AS position_name,
        positions.average_entry_quote
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
            price_quantity
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
        order_action as 'Action',
        volume_quantity as 'Volume',
        price_quantity as 'Price'
    FROM depot
    RIGHT JOIN orders ON depot.report_date = orders.trade_date AND depot.isin = orders.asset_isin
    ORDER BY depot.report_date;
"""

opus = OpusSource()


def get_positions() -> pd.DataFrame:
    df = opus.read_sql(query=query)
    stocks = get_stocks_data()

    def combine_currencies(row):
        if row['Crncy'] == row['Currency']:
            return row['Crncy']
        else:
            return f"{row['Crncy']}/{row['Currency']}"

    positions = pd.merge(df, stocks[['bloomberg_query', 'Currency', 'Last Price', '1D', '5D', '1MO', 'YTD']],
                         left_on='Query', right_on='bloomberg_query',
                         how='left')
    positions['Currency'] = positions.apply(combine_currencies, axis=1)
    positions.drop(columns=['Crncy'], inplace=True)

    positions.set_index(['Name', 'Position Name'], inplace=True)
    positions['% since AEQ'] = pd.to_numeric(((positions['Last Price'] - positions['AEQ']) / positions['AEQ']) * 100, errors='coerce')
    return positions


def calc_rel_performance(positions: pd.DataFrame, us: pd.DataFrame, eu: pd.DataFrame) -> pd.DataFrame:

    def calculate_difference(row, benchmark_df):
        benchmark_row = benchmark_df.loc[row['Sector']]

        for time_frame in ['1D', '5D', '1MO', 'YTD']:
            row[f'{time_frame} vs. Sector'] = row[time_frame] - benchmark_row[time_frame]

        return row

    positions = positions.apply(
        lambda row: calculate_difference(row, eu if row['Region'] == 'EU' else us), axis=1
    )

    return positions


def get_trades(account_id: str) -> pd.DataFrame:
    df = opus.read_sql(query=trades.format(account_id=account_id))
    df['P&L'] = ((df['Price'] - df['AEQ']) / df['AEQ']) * 100
    return df


def get_stocks_data() -> pd.DataFrame:
    df = pd.read_excel('single-stocks.xlsx', sheet_name='Stocks', header=0)
    df = df.rename(columns={'CURRENT_TRR_1D': '1D', 'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})
    mask = df['Currency'] == 'GBp'
    df.loc[mask, 'Currency'] = 'GBP'
    df.loc[mask, 'Last Price'] /= 100
    return df


def get_us_sector_data() -> pd.DataFrame:
    df = pd.read_excel('single-stocks.xlsx', sheet_name='US Sector', header=0, index_col=0)
    df.drop('Query', inplace=True, axis=1)
    df = df.rename(columns={'CURRENT_TRR_1D': '1D', 'CURRENT_TRR_5D': '5D', 'CURRENT_TRR_1MO': '1MO', 'CURRENT_TRR_YTD': 'YTD'})
    return df


def get_eu_sector_data() -> pd.DataFrame:
    df = pd.read_excel('single-stocks.xlsx', sheet_name='EU Sector', header=0)
    df.drop('Query', inplace=True, axis=1)
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


def style_positions_with_bars(positions: pd.DataFrame, name: str) -> str:
    columns_to_show = ['Position Name', 'AEQ', 'Currency', 'Volume', 'Last Price', '% since AEQ', '1D', '5D', '1MO', 'YTD',
                       '1D vs. Sector', '5D vs. Sector', '1MO vs. Sector', 'YTD vs. Sector']
    positions = positions.copy().reset_index()[columns_to_show]

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
    positions.set_index('Position Name', inplace=True)

    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))

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
    max_abs_value = trades['P&L'].abs().max()
    trades.sort_values('Trade Date', ascending=False, inplace=True)
    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(trades))
    styled = (
        trades.style.bar(subset='P&L', cmap=cm, align=0, vmax=max_abs_value, vmin=-max_abs_value)
        .format({
            'AEQ': "{:,.2f}",
            'Price': "{:,.2f}",
            'Volume': "{:,.0f}",
            'P&L': "{:.2f}%"
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


def write_mail(data: Dict):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Subject = "Daily Reporting - Risikomanagement"
    mail.Recipients.Add("pm-aktien")
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
                die Kurse der anhängenden Charts sind vom <b>{datetime.now().strftime('%d.%m.%Y %H:%M')}</b>.<br><br>
                Hier sind absoluten Entwicklungen der <b>Sektoren</b> für den <b>Euro Stoxx 600</b> und den <b>S&P 500</b>, 
                sowie die Out/Underperformance des <b>Euro Stoxx 600</b> gegenüber dem <b>S&P 500</b>:<br><br>
                {sector_images_html}
                <br><br>
                Hier sind die <b>Positionen</b> die sich in einem der jeweiligen Betrachtungszeiträume 
                <b>mindestens 8% schlechter</b> gegebenüber dem jeweiligen Sektor entwickelt haben:<br><br>
                {position_images_html}
                <br><br>
                Im Anhang findet Ihr für alle Fonds eine Übersicht über alle enthaltenen Positionen.
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
