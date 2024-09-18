import os
from typing import Dict

import pandas as pd
import win32com.client as win32
from matplotlib import pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
import dataframe_image as dfi
import seaborn as sns


output_dir = "output"
os.makedirs(os.path.join(output_dir, "images"), exist_ok=True)

mail = True


def get_megatrends() -> pd.DataFrame:
    trends = pd.read_excel('megatrends.xlsm', sheet_name='Themes', header=0, index_col=0)
    trends = trends.rename(
        columns={'SECURITY_NAME': 'Name', 'CHG_PCT_YTD': 'YTD', 'CHG_PCT_5D': '5D',
                 'CHG_PCT_1M': '1MO', 'CHG_PCT_3M': '3MO', 'CHG_PCT_6M': '6MO'}
    )
    return trends


def get_top_performers() -> pd.DataFrame:
    top_performer = pd.read_excel('megatrends.xlsm', sheet_name='TopPerformers', header=0)
    top_performer = top_performer.dropna()
    return top_performer


def style_mean_with_bars(positions: pd.DataFrame, name: str) -> str:

    max_val = positions['% Change'].abs().max()
    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))

    styled = (positions.style
    .bar(subset='% Change', cmap=cm, align=0, vmax=max_val, vmin=-max_val)
    .format({
        '% Change': "{:.2f}%",
    })).hide(axis="index")

    output_path = f"output/images/{name}.png"
    dfi.export(styled, output_path, table_conversion="selenium", max_rows=-1)
    return output_path


def style_trends_with_bars(positions: pd.DataFrame, name: str) -> str:
    d5_max_abs_value = max(abs(positions['5D'].min().min()),
                           abs(positions['5D'].max().max()))
    mo1_max_abs_value = max(abs(positions['1MO'].min().min()),
                            abs(positions['1MO'].max().max()))
    mo3_max_abs_value = max(abs(positions['3MO'].min().min()),
                            abs(positions['3MO'].max().max()))
    mo6_max_abs_value = max(abs(positions['6MO'].min().min()),
                            abs(positions['6MO'].max().max()))
    ytd_max_abs_value = max(abs(positions['YTD'].min().min()),
                            abs(positions['YTD'].max().max()))

    positions = positions.sort_values('6MO', ascending=False)

    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))

    styled = (positions.style
    .bar(subset='5D', cmap=cm, align=0, vmax=d5_max_abs_value, vmin=-d5_max_abs_value)
    .bar(subset='1MO', cmap=cm, align=0, vmax=mo1_max_abs_value, vmin=-mo1_max_abs_value)
    .bar(subset='3MO', cmap=cm, align=0, vmax=mo3_max_abs_value, vmin=-mo3_max_abs_value)
    .bar(subset='6MO', cmap=cm, align=0, vmax=mo6_max_abs_value, vmin=-mo6_max_abs_value)
    .bar(subset='YTD', cmap=cm, align=0, vmax=ytd_max_abs_value, vmin=-ytd_max_abs_value)
    .set_table_styles([
        {'selector': 'th.col0',
         'props': [('border-left', '1px solid black')]},
        {'selector': 'td.col0',
         'props': [('border-left', '1px solid black')]},
        {
            'selector': 'th.index_name',
            'props': [('min-width', '170px'), ('white-space', 'nowrap')]
        },
        {
            'selector': 'td.col0',
            'props': [('min-width', '300px'), ('white-space', 'nowrap')]
        },
        {
            'selector': 'td.col1',
            'props': [('min-width', '100px'), ('white-space', 'nowrap')]
        },
        {
            'selector': 'td.col2',
            'props': [('min-width', '200px'), ('white-space', 'nowrap')]
        }
    ])
    .format({
        '5D': "{:.2f}%",
        '1MO': "{:.2f}%",
        '3MO': "{:.2f}%",
        '6MO': "{:.2f}%",
        'YTD': "{:.2f}%"
    }))

    output_path = f"output/images/{name}.png"
    dfi.export(styled, output_path, table_conversion="selenium", max_rows=-1)
    return output_path


def plot_histogram(data: pd.DataFrame):
    plt.figure(figsize=(8, 6))
    sns.histplot(data['% Change'], bins=100, kde=True)
    plt.title('Distribution of % Change')
    plt.xlabel('% Change')
    plt.ylabel('Frequency')
    output_path = f"output/images/top10-theme-positions-histogram.png"

    plt.savefig(output_path)
    return output_path


def write_mail(data: Dict):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.Subject = "Daily Reporting - strukturelle Themen & Zukunftsthemen"
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

    theme_images_html = f'<p><br><img src="cid:{inplace_chart(data.get("theme"))}"><br></p>'

    mail.HTMLBody = f"""
        <html>
          <head></head>
          <body>
            <p>Hi zusammen, <br><br>
                
                anbei eine Übersicht über die Performance verschiedener Themen-Investments, sortiert nach 6MO-Performance:
                <br><br>
                {theme_images_html}
                <br><br>
    
                Liebe Grüße
            </p>
          </body>
        </html>
    """

    mail.Recipients.ResolveAll()
    mail.Display(True)
