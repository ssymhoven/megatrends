from utility import get_third_party_products, style_third_party, group_funds, write_third_party_mail

third_party = get_third_party_products()


if __name__ == '__main__':
    esg = third_party[third_party.index.get_level_values('Name').str.contains("ESG")]
    flex = third_party[third_party.index.get_level_values('Name').str.contains("Flex")]
    select = third_party[third_party.index.get_level_values('Name').str.contains("Strategie - Select")]

    esg_chart = style_third_party(positions=group_funds(esg), name="VV-ESG")
    flex_chart = style_third_party(positions=group_funds(flex), name="VV-Flex")
    strategie_chart = style_third_party(positions=group_funds(select), name="Strategie - Select")

    write_third_party_mail(data={
        'flex': flex_chart,
        'esg': esg_chart,
        'strategie': strategie_chart
    })


