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
		sheet_type = set()
		for row in response.xpath('//select/optgroup')[5].xpath('.//option'):
			name=size=type=latency=price=None
			result=row.extract()
			try: print(result)
			except: continue
			try: price = re.findall('\$\d+', result)[1]
			except: price = re.findall('\$\d+', result)[0]
			result = re.search('^.+?,', result).group()
			result = re.sub('<.+?>','',result)
			result = re.sub(r'(D4-|DD[DR]4 -?)(\d+)',r'DDR4-\2',result)
			result = re.sub(r'(\d+)MHz D(DR)?4',r'DDR4-\1',result)
			if "DDR" not in result: continue
			try: size = re.search('\(?\d+G\*\d\)?', result).group()
			except: size = re.search('\d+GB?', result).group()
			type = re.search('DDR\dL?-\d+', result).group()
			name = re.sub(type,'',result)
			try:
				latency = re.search('CL ?\d+(-\d+-\d+)?', result).group()
				name = re.sub(latency,'',name)
			except: pass
			name = name.replace(size,'')
			size = re.sub('[\(\)]','',size)
			if "B" not in size: size = re.sub('G','GB',size)
			name = re.sub('[,/]','',name)
			name = re.sub('\s+',' ',name)
			print("  ==> " + name)
			if type not in sheet_type:
				ws = wb.create_sheet(type)
				sheet_type.add(type)
				ws.append(["Name", "Size", "Type", "Latency", "Price"])
				ws.append([name, size, type, latency, price])
			else:
				ws = wb[type]
				ws.append([name, size, type, latency, price])
		ws = wb.worksheets[0]
		wb.remove_sheet(ws)
		wb.save(filename)
	