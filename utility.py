import os
from typing import Dict

import pandas as pd
import win32com.client as win32
from matplotlib.colors import LinearSegmentedColormap
import dataframe_image as dfi

output_dir = "output"
os.makedirs(os.path.join(output_dir, "images"), exist_ok=True)

mail = True


def get_megatrends() -> pd.DataFrame:
    trends = pd.read_excel('megatrends.xlsx', sheet_name='Themes', header=0, index_col=0)
    trends = trends.rename(
        columns={'SECURITY_NAME': 'Name', 'CHG_PCT_YTD': 'YTD', 'CHG_PCT_5D': '5D',
                 'CHG_PCT_1M': '1MO', 'CHG_PCT_3M': '3MO', 'CHG_PCT_6M': '6MO', 'CHG_PCT_HIGH_52WEEK': 'Δ 52W High'})
    return trends


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
    delta_max_abs_value = max(abs(positions['Δ 52W High'].min().min()),
                              abs(positions['Δ 52W High'].max().max()))

    #positions.drop('Description', inplace=True, axis=1)
    positions.sort_values('YTD', ascending=False, inplace=True)

    cm = LinearSegmentedColormap.from_list("custom_red_green", ["red", "white", "green"], N=len(positions))

    styled = (positions.style
    .bar(subset='5D', cmap=cm, align=0, vmax=d5_max_abs_value, vmin=-d5_max_abs_value)
    .bar(subset='1MO', cmap=cm, align=0, vmax=mo1_max_abs_value, vmin=-mo1_max_abs_value)
    .bar(subset='3MO', cmap=cm, align=0, vmax=mo3_max_abs_value, vmin=-mo3_max_abs_value)
    .bar(subset='6MO', cmap=cm, align=0, vmax=mo6_max_abs_value, vmin=-mo6_max_abs_value)
    .bar(subset='YTD', cmap=cm, align=0, vmax=ytd_max_abs_value, vmin=-ytd_max_abs_value)
    .bar(subset='Δ 52W High', cmap=cm, align=0, vmax=delta_max_abs_value, vmin=-delta_max_abs_value)
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
        'YTD': "{:.2f}%",
        'Δ 52W High': "{:.2f}%",
    }))

    output_path = f"output/images/{name}.png"
    dfi.export(styled, output_path, table_conversion="selenium", max_rows=-1)
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
                
                anbei eine Übersicht über die Performance aktueller Zukunftsthemen und struktureller Themen:
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
