import requests
import xlwt,datetime
from lxml import html, etree

checko_url = 'https://checko.ru/search?query='
headers = [
    'НАЗВАНИЕ',
    'ОГРН',
    'ИНН',
    'ГЕН.ДИР',
    'АДРЕС',
    'Учредители',
    'Выручка',
    'Себестоимость',
    'Прибыль',
    'КОДЫ'
]


def parse(inn_str):
    info = dict.fromkeys(headers, '')
    try:
        url_str = checko_url+inn_str
        resp = requests.get(url_str)
        if 'По вашему запросу не найдено ни одного совпадения' in resp.text:
            raise ValueError('неправильный ИНН или ОГРН')
        elif 'Слишком короткий запрос' in resp.text:
            raise ValueError('неправильный ИНН или ОГРН')

        tree = html.fromstring(resp.text)
        tree = tree.xpath('//main')[0]
        tree = tree.xpath(
            '//div[@class="uk-width-expand@m uk-margin-medium-top"]')[0]

        name = tree.xpath(
            '//div[@class="uk-grid uk-grid-small"]')[0].xpath('.//h1')[0].text
        info['НАЗВАНИЕ'] = name
        table = tree.xpath('//table[@id="shortcut:information"]')[0]
        ogrn = table.xpath('.//tr')[0].xpath('.//td/span')[0].text
        #print('ОГРН: '+ogrn)
        info['ОГРН'] = ogrn

        inn = table.xpath('.//tr')[1].xpath('.//td/span')[0].text
        #print('ИНН: '+inn)
        info['ИНН'] = inn

        gen_fio = table.xpath('//tr')[7].xpath('.//td/a')[0].text
        #print('ГЕН.ДИР: '+gen_fio)
        info['ГЕН.ДИР'] = gen_fio

        addr = table.xpath('.//tr')[5].xpath('.//td')[0].text
        #print('АДРЕС: '+addr)
        info['АДРЕС'] = addr

        # print("Учредители:")
        founders = []
        founders_block = tree.xpath(
            '//section[@id="shortcut:founders"]/table/tbody/tr[@class="data-line"]')
        for i in founders_block:
            founder = i.xpath('.//td')[0].text
            if founder == None:
                founder = i.xpath('.//td/a')[0].text
            # print(founder)
            founders.append(founder)
        info['Учредители'] = ', '.join(founders)
        if len(tree.xpath('//section[@id="shortcut:accounting"]')) > 0:
            extra_link = tree.xpath(
                '//section[@id="shortcut:accounting"]/p/a/@href')[0]
            extra_link = 'https://checko.ru'+extra_link
            (revenue, cost_price, clear_revenue) = get_account(extra_link)
            #print('Выручка '+revenue)
            #print('себестоимость '+cost_price)
            #print('прибыль '+clear_revenue)
            info['Выручка'] = revenue
            info['Себестоимость'] = cost_price
            info['Прибыль'] = clear_revenue

        activity_block = tree.xpath('//section[@id="shortcut:activity"]')[0]
        activity_list = []
        if len(activity_block.xpath('.//tr[@class="td-padding-top"]')) > 0:
            activity_link = activity_block.xpath(
                './/tr[@class="td-padding-top"]/td/a/@href')[0]
            activity_link = 'https://checko.ru'+activity_link
            activity_list = get_ativity(activity_link)
        else:
            activity_tr_list = activity_block.xpath('.//tr')
            activity_list = [
                i.xpath('.//td')[0].text for i in activity_tr_list]
        #print('Виды деятельности:')
        # for i in activity_list:
        #     print(i)
        info['КОДЫ'] = ', '.join(activity_list)
        return info
        # with open('test.html', 'w') as f:
        # f.write(etree.tostring(tree).decode("utf-8"))
    except Exception as err:
        print(err)
        print('ИНН/ОГРН ошибки:'+inn_str)
        return None


def get_account(url_str):
    resp = requests.get(url_str)
    tree = html.fromstring(resp.text).xpath(
        './/div[@class="uk-switcher uk-margin"]')[0]
    tree = tree.xpath(
        './/table[@class="uk-table basic-financial-data full-financial-data"]')[1]
    tr_list = tree.xpath('.//tbody/tr')
    revenue = tr_list[0].xpath('.//td')[-1].text.replace(',', '').split()[0]
    cost_price = tr_list[1].xpath('.//td')[-1].text.replace(',', '').split()[0]
    clear_revenue = tr_list[17].xpath(
        './/td')[-1].text.replace(',', '').split()[0]
    return (revenue, cost_price, clear_revenue)


def get_ativity(url_str):
    resp = requests.get(url_str)
    tr_list = html.fromstring(resp.text).xpath('//tbody/tr')
    activity = [i.xpath('.//td')[0].text for i in tr_list]
    return activity


def make_xls(check_list):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('list')
    i = 0
    for elem in headers:
        ws.write(0, i, elem)
        i += 1
    i = 1
    for elem in check_list:
        print('parsing '+elem)
        info = parse(elem)
        if info == None:
            continue
        for j in range(len(headers)):
            ws.write(i, j, info[headers[j]])
        i += 1
    today = datetime.datetime.today()
    time_str= today.strftime("%Y-%m-%d-%H:%M:%S")
    file_name='checko_list_'+time_str+'.xls'
    wb.save(file_name)

