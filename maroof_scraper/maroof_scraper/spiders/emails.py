# -*- coding: utf-8 -*-
import csv
import glob
import os.path
from openpyxl import Workbook

from scrapy.spiders import Spider
from scrapy.http import Request


class EmailsSpider(Spider):
    name = 'emails'
    allowed_domains = ['maroof.sa']
    start_urls = [
        'https://maroof.sa/BusinessType/MoreBusinessList?businessTypeId=14&pageNumber=0&sortProperty=BestRating&desc=True'
    ]

    def parse(self, response):
        data = response.json()
        businesses = data["Businesses"]

        for business in businesses:
            yield Request(f"https://maroof.sa/" + str(business["Id"]), self.parse_item)

    def parse_item(self, response):
        email = response.xpath(
            '//p[contains(text(),"البريد الإلكتروني ")]/following-sibling::p//text()').extract_first()

        yield {
            'email': email.strip(),
        }

    def close(spider, reason):
        csv_file = max(glob.iglob('*.csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r', encoding="utf8") as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
