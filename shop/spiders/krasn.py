import scrapy
import openpyxl


class QuotesSpider(scrapy.Spider):
    name = 'krasn'
    start_urls = [
        'https://krasn.russcvet.ru/catalog/enamels/',
    ]

    @staticmethod
    def write_to_file(items):
        """функция записи данных парсинга в файл"""

        wb = openpyxl.Workbook()

        # добавляем новый лист
        wb.create_sheet(title='Первый лист', index=0)

        # получаем лист, с которым будем работать
        sheet = wb['Первый лист']

        # добавляем заголовки
        sheet['A1'] = 'Name'
        sheet['B1'] = 'Price'

        # добавляем данные в таблицу
        for i in range(2, len(items) + 2):
            sheet[f'A{i}'] = (items[i - 2]['name'])
            sheet[f'B{i}'] = (items[i - 2]['price'])

        # сохнанение в файле
        wb.save('result.xlsx')

    def parse(self, response):
        """процесс парсинга"""

        count = 0  #переменная для подсчета колл-ва элементов каталога
        items = []

        for item in response.css('div.catalog-item'):
            items.append({
                'name': item.css('a::text').getall()[2],
                'price': item.css('span::text').get(),
            })
            count += 1
            if count >= 200:
                break

        # в случае изменения колличества элементов каталога в большую сторону,
        # переход на след. страницу каталога.

        next_page = response.css('a[class="modern-page-next"]::attr("href")').get()
        if next_page is not None:

            print(f'https://spb.russcvet.ru{next_page}')
            next_page = response.urljoin(next_page)
            req = scrapy.Request(url=next_page, callback=self.parse)
            req = response.follow(next_page, callback=self.parse)
            yield req

        self.write_to_file(items)


