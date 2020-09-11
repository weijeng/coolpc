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
		ws.append(["Name", "Size", "Type", "Latency", "Price"])
		for row in response.xpath('//select/optgroup')[5].xpath('.//option'):
			name=size=type=latency=price=None
			result=row.extract()
			try: print result
			except: continue
			try: price = re.findall('\$\d+', result)[1]
			except: price = re.findall('\$\d+', result)[0]
			result = re.search('^.+?,', result).group()
			result = re.sub('<.+?>','',result)
			result = re.sub('(D4-|DDR4 -)','DDR4-',result)
			if "DDR" not in result: continue
			try: size = re.search('[\-\s]?\d+(GB|BG|G)(\*\d)?', result).group()
			except: size = re.search('\d+G\(\d+G\*\d\)', result).group()
			size = size.replace("BG","GB").replace(" ","")
			type = re.search('DDR\dL?[\-\s](\d+|ECC)', result).group()
			try: latency = re.search('CL ?\d+(-\d+-\d+)?', result).group()
			except: pass
			name = re.sub(size,'',result)
			name = re.sub(type,'',name)
			try: name = re.sub(latency,'',name)
			except: pass
			name = re.sub('[,/]','',name)
			type = re.sub(' ','-',type)
			print "  ==> " + name
			ws.append([name, size, type, latency, price])
		wb.save(filename)
	