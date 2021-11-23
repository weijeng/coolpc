# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "Graphic_" + now.strftime("%Y-%m-%d") + ".xlsx"

class coolpcSpider(scrapy.Spider):
	name = "Graphic"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
			'https://benchmarks.ul.com/compare/best-gpus'
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)
		yield scrapy.Request(url=urls[1], callback=self.benchmarks)

	def parse(self, response):
		wb = openpyxl.Workbook()
		ws = wb.active
		for row in response.xpath('//select/optgroup')[11].xpath('.//option'):
			name=type=price=None
			result=row.extract()
			try: print(result)
			except: continue
			if "cm" not in result: continue
			name = re.sub(r'[\*\s]*參考價[\$\s]*(\d+)',r'$\1',result)
			try: price = re.findall('\$\d+', name)[1]
			except: price = re.findall('\$\s?\d+', name)[0]
			name = re.sub('<.+?>','',name)
			name = re.search('^.+?\$', name).group()
			try: type = re.findall('\(.+?[\)\*\$]', name)[1]
			except: type = re.findall('\(.+?[\)\*\$]', name)[0]
			name = name.replace(type, "").replace(",","")
			type = re.sub('[\(\)\$]','',type)
			name = re.sub('(,\s)?\$|❤ ','',name)
			ws.append([name, type, price])
			print("  => " + name)
			print("  => " + type)
			print("  => " + price)
		wb.save(filename)
		time.sleep(3)
		
	def benchmarks(self, response):
		wb = openpyxl.load_workbook(filename)
		ws = wb.create_sheet("score")
		gpu = response.xpath('//td/a/text()').extract()
		score = response.xpath('//td[@class="small-pr1"]/div/div/span/text()').extract()
		price = response.xpath('//td[@class="list-tiny-none"]/span/div/text()').extract()
		for x, y, z in zip(gpu, score, price):
			ws.append([x.strip(), y.strip(), z.strip()])
		wb.save(filename)
	