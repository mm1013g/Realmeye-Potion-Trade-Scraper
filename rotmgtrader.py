from bs4 import BeautifulSoup
import requests
import xlsxwriter

base_trade_url = 'https://www.realmeye.com/offers-by/'

pots = {
    2793: "Life Potion",
    2592: "Defense Potion",
    2591: "Attack Potion",
    2593: "Speed Potion",
    2636: "Dexterity Potion",
    2613: "Wisdom Potion",
    2612: "Vitality Potion",
    2794: "Mana Potion"
}
rev_pots = {v: k for k, v in pots.items()}

# In terms of Speed Pots ETA
pot_values = {
    "Life Potion": 8,
    "Defense Potion": 3,
    "Attack Potion": 3,
    "Speed Potion": 1,
    "Dexterity Potion": 1,
    "Wisdom Potion": 1.2,
    "Vitality Potion": 2,
    "Mana Potion": 5,
}


class Item:
    def __init__(self, item_id, qty):
        self.item_id = item_id
        self.qty = int(qty[1:])
        if item_id in pots.keys():
            self.name = pots[item_id]
            self.worth = self.qty * pot_values[self.name]
        else:
            self.name = 'OTHER_ITEM'
            self.worth = 0

    def __str__(self):
        return f'{self.qty} {self.name}'


class Trade:
    def __init__(self, selling_items, buying_items, qty, added, author, seen, server):
        self.selling_items = selling_items
        self.buying_items = buying_items
        self.qty = qty
        self.added = added
        self.author = author
        self.seen = seen
        self.server = server
        self.selling_worth = sum([i.worth for i in selling_items])
        self.buying_worth = sum(i.worth for i in buying_items)


def getBuyingList(id):
    buying_spd_pots_url = f'https://www.realmeye.com/offers-to/buy/{id}/pots'
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'}
    r = requests.get(buying_spd_pots_url, headers=headers)
    soup = BeautifulSoup(r.content, 'html.parser')

    tbl = soup.find("table", id="g").find("tbody")
    raw_trades = tbl.findAll("tr")

    trades_list = []

    for trade in raw_trades:
        trade_attributes = trade.findAll('td')

        selling_list = trade_attributes[0].findAll('span', {'class': 'item-static'})
        buying_list = trade_attributes[1].findAll('span', {'class': 'item-static'})

        selling_itms = []
        for itm in selling_list:
            data_id = int(itm.find('span', {'class': 'item'})['data-item'])
            data_qt = itm.find('span', {'class': 'item-quantity-static'}).text
            temp_item = Item(data_id, data_qt)
            selling_itms.append(temp_item)

        buying_itms = []
        for itm in buying_list:
            data_id = int(itm.find('span', {'class': 'item'})['data-item'])
            data_qt = itm.find('span', {'class': 'item-quantity-static'}).text
            buying_itms.append(Item(data_id, data_qt))

        trade_qty = trade_attributes[2].find('span')
        trade_added = trade_attributes[3].find('span').text
        trade_author = trade_attributes[5].find('a').text
        trade_seen = trade_attributes[6].find('span')
        trade_server = trade_attributes[7].find('abbr')
        if trade_server is not None:
            trade_server = trade_server.text

        trades_list.append(Trade(selling_itms, buying_itms, trade_qty, trade_added, trade_author, trade_seen, trade_server))
    return trades_list


col_headers = [
    {'header': 'USER'},
    {'header': 'SERVER'},
    {'header': 'SELLING'},
    {'header': 'BUYING'},
    {'header': 'SELLING WORTH'},
    {'header': 'BUYING WORTH'},
    {'header': 'DIFFERENCE'},
]

trades_file = open('trades.txt', 'w')
workbook = xlsxwriter.Workbook('Trades.xlsx')
worksheet = workbook.add_worksheet()

data = []
for pot_type in pots.values():
    trades_list = getBuyingList(rev_pots[pot_type])

    for trade in trades_list:
        buying_items_string = ''
        for item in trade.buying_items:
            if item.name == 'OTHER_ITEM':
                continue
            buying_items_string += f'{item}, '

        selling_items_string = ''
        for item in trade.selling_items:
            if item.name == 'OTHER_ITEM':
                continue
            selling_items_string += f'{item}, '

        user_url = f'{base_trade_url}{trade.author}'
        excel_hyperlink = f'=HYPERLINK(\"{user_url}\",\"{trade.author}\")'

        data.append([excel_hyperlink, trade.server, selling_items_string, buying_items_string, trade.selling_worth,
                     trade.buying_worth, trade.selling_worth - trade.buying_worth])

        printing_string = f'{trade.author}({trade.server}) has {selling_items_string}and wants {buying_items_string}'
        print(printing_string)
        trades_file.write(printing_string + '\n')

worksheet.add_table(0, 0, len(data), len(col_headers) - 1, {'data': data, 'columns': col_headers})
trades_file.close()
workbook.close()
