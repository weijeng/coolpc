# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "RAM_" + now.strftime("%Y-%m-%d") + ".xlsx"

class coolpcSpider(scrapy.Spider):
	name = "ram"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)

	def parse(self, response):
		wb = openpyxl.Workbook()
		ws = wb.active
		for row in response.xpath('//select/optgroup')[5].xpath('.//option'):
			name=size=type=price=None
			result=row.extract()
			try: print result
			except: continue
			name = re.search('^.+?,', result).group()
			name = re.sub('<.+?>','',name)
			if "D4" in name: name = name.replace("D4", "DDR4")
			if "DDR4 -" in name: name = name.replace("DDR4 -","DDR4-")
			if "DDR" not in name: continue
			print "  ==> " + name
			try: size = re.search('[\-\s]?\d+(GB|BG)(\*\d)?', name).group()
			except: size = re.search('\d+G\(\d+G\*\d\)', name).group()
			name = name.replace(size, "")
			size = size.replace("BG","GB").replace(" ","")
			type = re.search('DDR\dL?[\-\s](\d+|ECC)', name).group()
			name = name.replace(type, "").replace(",","")
			if "/" in name: name = name.replace("/","")
			try: price = re.findall('\$\d+', result)[1]
			except: price = re.findall('\$\d+', result)[0]
			ws.append([name, size, type, price])
			print "  ==> " + size + "\t" + type + "\t" + price
			#time.sleep(2)
		wb.save(filename)
	