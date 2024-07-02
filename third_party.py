from utility import get_third_party_products, style_third_party, group_funds, write_third_party_mail

third_party = get_third_party_products()


if __name__ == '__main__':
    esg = third_party[third_party.index.get_level_values('Name').str.contains("ESG")]
    flex = third_party[third_party.index.get_level_values('Name').str.contains("Flex")]
    strategie_select = third_party[third_party.index.get_level_values('Name').str.contains("Strategie - Select")]
    premium_select = third_party[third_party.index.get_level_values('Name').str.contains("Premium Select")]

    esg_chart = style_third_party(positions=group_funds(esg), name="VV-ESG")
    flex_chart = style_third_party(positions=group_funds(flex), name="VV-Flex")
    strategie_select_chart = style_third_party(positions=group_funds(strategie_select), name="Strategie-Select")
    premium_select_chart = style_third_party(positions=group_funds(premium_select), name="Premium-Select")

    write_third_party_mail(data={
        'flex': flex_chart,
        'esg': esg_chart,
        'strategie-select': strategie_select_chart,
        'premium-select': premium_select_chart
    })


