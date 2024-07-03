from utility import get_megatrends, style_trends_with_bars, write_mail

if __name__ == '__main__':
    trends = get_megatrends()
    theme_image = style_trends_with_bars(positions=trends, name='Trends')

    write_mail(data={
        'theme': theme_image
    })
