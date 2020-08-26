# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "SSD_" + now.strftime("%Y-%m-%d") + ".xlsx"

class coolpcSpider(scrapy.Spider):
	name = "ssd"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)

	def parse(self, response):
		wb = openpyxl.Workbook()
		ws = wb.active
		for row in response.xpath('//select/optgroup')[6].xpath('.//option'):
			name=size=read=write=type=price=None
			result=row.extract()
			try: print result
			except: continue
			name = re.search('^.+?\/', result).group()
			name = re.sub('<.+?>','',name)
			name = name.replace("/", "")
			name = re.sub('[^w]:\d+M','',name)
			try: size = re.search('\d{3,4}GB?|\dTB', result).group()
			except: pass
			if size: 
				name = name.replace(size, "")
				size = size.replace(" ", "").replace("-", "").replace("G", "").replace("B", "")
				if "T" in size:
					size = size.replace("T", "")
					size = int(size)*1024
			try: read = re.findall('\:\d{3,4}', result)[0].replace(":", "")
			except: continue
			write = re.findall('\:\d{3,4}', result)[1].replace(":", "")
			try: type = re.search('[TQM]LC', result).group()
			except: pass
			price = re.search('\$\d+', result).group()
			ws.append([name, size, read, write, type, price])
			#print "  ==> " + name + '\t' + size + '\t' + read + '\t' + write + '\t' + price
			#time.sleep(3)
		wb.save(filename)
	