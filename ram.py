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
			name=size=type=latency=price=price_list=None
			result=row.extract()
			if "disabled" in result: continue
			if "ECC" in result: continue
			try: print(result)
			except: continue
			price = re.findall('\$\d+', result)
			price = min(price).strip('$')
			result = re.search('^.+?,', result).group()
			result = re.sub('<.+?>', '', result)
			result = re.sub(r'\(\d+\*\d\)', '', result)				# modify: (2048*8)
			result = re.sub(r'(D4-|DD[DR]4 -?)(\d+)', r'DDR4-\2', result)
			result = re.sub(r'D5[\s-](\d+)', r'DDR5-\1', result)	# modify: D5-6000, D5 6000
			result = re.sub(r'(\d+)MHz D(DR)?4', r'DDR4-\1', result)# modify: 3200MHz D4
			result = re.sub(r'DDR5 (\d+)', r'DDR5-\1', result)		# modify: DDR5 5600
			result = re.sub(r'(\d+)G?B?[*x](\d)', r'\1GBx\2', result)
			result = re.sub(r'(\d)G DDR', r'\1GB DDR', result)		# modify: 4G DDR3-1600
			name = re.search('^.+?GB', result).group()
			name = re.sub(r'\d+GB$', '', name).strip(' ')
			try: size = re.search('\d+GBx\d', result).group()
			except: size = re.search('\d+GB', result).group()
			type = re.search('DDR\dL?-\d+', result).group()
			try: latency = re.search('CL ?\d+(-\d+-\d+)?', result).group()
			except: latency = 'n/a'
			print("  => " + name + '|' + size + '|' + type + '|' + latency + '|' + price)
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
	