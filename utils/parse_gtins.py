from contextlib import suppress

from tools.web import Web


def parse_all_gtins_to_out(web: Web, url: str):

    all_id_code = []
    all_gtin_code = []
    all_goods_name = []

    all_goods = dict()

    web.get(url)

    web.wait_element('//*[@id="1"]/span', timeout=120)
    web.find_element('//*[@id="1"]/span').click()

    id_code = []

    gtin_code = []

    goods_name = []

    while True:

        for el in web.find_elements("//a[contains(@class, 'jcKonU')]/div"):
            id_code.append(el.get_attr('text'))
            # break

        for el in web.find_elements("//a[contains(@class, 'jcKonU')]/../preceding-sibling::div[3]/div"):
            gtin_code.append(el.get_attr('text'))
            # break

        for el in web.find_elements("//a[contains(@class, 'jcKonU')]/../preceding-sibling::div[2]"):
            goods_name.append(el.get_attr('text'))
            # break

        try:
            current_page = int(web.find_element("//li[@class='sc-giadOv ekuBbr']", timeout=5).get_attr('text'))
        except:
            break

        available_page = None
        for pages in web.find_elements("//li[@class='sc-giadOv hALQxM']"):
            with suppress(Exception):
                if int(pages.get_attr('text')) > current_page:
                    available_page = int(pages.get_attr('text'))
                    break
        # break
        if available_page is None:
            break

        web.find_element(f"//li[text()='{available_page}']").click()

    print(id_code)
    print(gtin_code)
    print(goods_name)

    all_goods.update({url: [id_code, gtin_code, goods_name]})

    # if ind == 2:
    #     break

    # all_id_code.append(id_code)
    # all_gtin_code.append(gtin_code)
    # all_goods_name.append(goods_name)

    # print(all_id_code)
    # print(all_gtin_code)
    # print(all_goods_name)

    # print()

    return all_goods # all_id_code, all_gtin_code, all_goods_name
