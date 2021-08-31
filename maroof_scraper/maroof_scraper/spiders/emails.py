# -*- coding: utf-8 -*-
import csv
import glob
import os.path
from openpyxl import Workbook
import names

from scrapy.spiders import Spider
from scrapy.http import Request

PAGE = 0
SIZE = 10
BASE_URL = "https://maroof.sa/"


class EmailsSpider(Spider):
    name = 'emails'
    allowed_domains = ['maroof.sa']
    start_urls = [
        BASE_URL
    ]

    def parse(self, response):
        yield Request(
            f'{BASE_URL}/BusinessType/MoreBusinessList?businessTypeId=23&pageNumber={PAGE}&sortProperty=BestRating&desc=True',
            callback=self.parse_items)

    def parse_items(self, response):
        if response.text != "":
            data = response.json()
            businesses = data["Businesses"]

            for business in businesses:
                yield Request(f"{BASE_URL}" + str(business["Id"]), self.parse_item)

            page_number = data["PageNumber"]
            size = data["Size"]
            count = data["Count"]

            if (page_number * size) < count:
                yield Request(response.urljoin(
                    f'?businessTypeId=23&pageNumber={page_number}&sortProperty=BestRating&desc=True'),
                    self.parse_items)

    def parse_item(self, response):
        email = response.xpath(
            '//p[contains(text(),"البريد الإلكتروني ")]/following-sibling::p//text()').extract_first()

        yield {
            'First name': names.get_first_name(),
            'Last name': names.get_last_name(),
            'Email': email,
        }

    def close(spider, reason):
        csv_file = max(glob.iglob('*.csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r', encoding="utf8") as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
