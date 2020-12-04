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
			try: print result
			except: continue
			name = re.search('^.+?\$', result).group()
			name = re.sub('<.+?>','',name)
			name = re.sub(',\s\$','',name)
			if "cm" not in name: continue
			type = re.findall('\(.+?\)', name)
			for list in type:
				if "cm" not in list: continue
				type = list
			name = name.replace(type, "").replace(",","")
			type = type.replace("(","").replace(")","")
			try: price = re.findall('\$\d+', result)[1]
			except: price = re.findall('\$\d+', result)[0]
			ws.append([name, type, price])
			print "  ==> " + name
			print "  ==> " + type + "\t" + price
			#time.sleep(2)
		wb.save(filename)
		
	def benchmarks(self, response):
		wb = openpyxl.load_workbook(filename)
		ws = wb.create_sheet("score")
		gpu = response.xpath('//td/a/text()').extract()
		score = response.xpath('//td[@class="small-pr1"]/div/div/span[@class="bar-score"]/text()').extract()
		for x, y in zip(gpu, score):
			ws.append([x.strip(), y.strip()])
		wb.save(filename)
	