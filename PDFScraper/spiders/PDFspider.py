import pandas
from scrapy.spiders import XMLFeedSpider
from openpyxl import load_workbook
import os


# how to run: scrapy crawl pdfspider

article = 92 # +\- number of article to start
stop_article = 40 # how much articles to scrap
# path_to_xml = 'test.xml'
path_to_xls = r'Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format (2).xlsx' # XLS
dir_path = os.path.dirname(os.path.realpath(__file__))

class Spider(XMLFeedSpider):
    name = "pdfspider"
    start_urls = [f"file:{dir_path}\\converted.xml"] # problem with directory in windows and linux

    def parse(self, response):
        i = 5
        for x in range(stop_article):
            index = response.xpath(f'//Sect[{x + article}]/H5/text()').get()
            title = response.xpath(f'//Sect[{x + article}]/P[1]/text()').get()
            author = response.xpath(f'//Sect[{x + article}]/P[2]/text()').get()
            author = ''.join([c for c in author if not c.isdigit()])
            text = response.xpath(f'//Sect[{x + article}]/P/text()').getall()
            text = text[2:]
            text = " ".join(text)

            index_book = pandas.DataFrame({'Index': [index]})
            title_book = pandas.DataFrame({'Title': [title]})
            author_book = pandas.DataFrame({'Author': [author]})
            text_book = pandas.DataFrame({'Text': [text]})

            book = load_workbook(path_to_xls)
            writer = pandas.ExcelWriter(path_to_xls, engine='openpyxl', mode='a')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            i = i + 1 # Add next row (until use max_row)
            index_book.to_excel(writer, "Sheet1", startrow=i, startcol=3, header=False, index=False)
            title_book.to_excel(writer, "Sheet1", startrow=i, startcol=4, header=False, index=False)
            author_book.to_excel(writer, "Sheet1", startrow=i, startcol=0, header=False, index=False)
            text_book.to_excel(writer, "Sheet1", startrow=i, startcol=5, header=False, index=False)
            writer.save()