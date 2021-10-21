# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "HDD_" + now.strftime("%Y-%m-%d") + ".xlsx"

class coolpcSpider(scrapy.Spider):
	name = "hdd"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)

	def parse(self, response):
		wb = openpyxl.Workbook()
		ws = wb.active
		ws.append(["Name", "Size", "Cache", "Speed", "Model", "Price"])
		for row in response.xpath('//select/optgroup')[7].xpath('.//option'):
			name=size=cache=speed=model=price=None
			result=row.extract()
			try: print(result)
			except: continue
			try: name = re.search('\w+\s\d+(TB|GB?)', result).group()
			except: continue
			if "TB" in name:
				size = re.search('\d+',name).group()
				size = int(size) * 1000
			if "G" in name:
				size = re.search('\d+',name).group()
			cache = re.search('(16|32|64|128|256|512)M?B?/', result).group()
			cache = cache.strip("/")
			print("  ==> " + name)
			try: speed = re.search('(54|57|59|72)00', result).group()
			except: speed = "N/A"
			try: model = re.findall('\(.+?\)', result)[1].replace("(","").replace(")","")
			except: model = re.findall('\(.+?\)', result)[0].replace("(","").replace(")","")
			try: price = re.findall('\$\d+', result)[1]
			except: price = re.findall('\$\d+', result)[0]
			ws.append([name, int(size), cache, speed, model, price])
			print("  ==> " + str(size) + "\t" + cache + "\t" + speed + "\t" + model + "\t" + price)
		wb.save(filename)