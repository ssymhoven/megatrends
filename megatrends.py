from matplotlib import pyplot as plt

from utility import get_megatrends, style_trends_with_bars, write_mail, get_top_performers, plot_histogram, \
    style_mean_with_bars
import seaborn as sns

if __name__ == '__main__':
    trends = get_megatrends()
    theme_image = style_trends_with_bars(positions=trends, name='Trends')

    top_performers = get_top_performers()
    top_10_theme_positions_histogram = plot_histogram(top_performers)

    theme_mean_performance = top_performers.groupby('Theme')['% Change'].mean().sort_values(ascending=False).reset_index()
    theme_mean = style_mean_with_bars(positions=theme_mean_performance, name="theme_mean")

    sector_mean_performance = top_performers.groupby('Sector')['% Change'].mean().sort_values(ascending=False).reset_index()
    sector_mean = style_mean_with_bars(positions=sector_mean_performance, name="sector_mean")

    sectors = top_performers['Sector'].unique()
    
    for sector in sectors:
        sector_df = top_performers[top_performers['Sector'] == sector]
        sector_df = sector_df.drop_duplicates(subset='Query').sort_values('% Change', ascending=False).head(5)
        image_path = style_mean_with_bars(sector_df[['Name', 'Query', '% Change']], f"{sector}_mean_performance")

    write_mail(data={
        'theme': theme_image
    })

