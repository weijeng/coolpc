# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "Notebook_" + now.strftime("%Y-%m-%d") + ".xlsx"

class coolpcSpider(scrapy.Spider):
	name = "Notebook"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)

	def parse(self, response):
		wb = openpyxl.Workbook()
		ws = wb.active
		notebook=response.xpath('//select/optgroup')[1]
		i=0
		ws.append(['Series', 'Name', 'Size', 'CPU', 'Ram', 'Storage', 'GPU', 'Price', 'Difference'])
		for row in notebook.xpath('.//optgroup/@label'):
			result=row.extract()
			print("=> " + result)
			size=series=None
			try: size=re.search('\d+(\.\d)?吋', result).group()
			except AttributeError: i+=1; continue
			series=result.strip(size).strip(' 系列')
			for row2 in notebook.xpath('.//optgroup')[i].xpath('./option'):
				result2=row2.extract()
				if "disabled" in result2: continue
				if "延長保固" in result2: continue
				if "加贈" in result2: continue
				if "外接顯示卡" in result2: continue
				if "序號" in result2: continue
				if "任一款" in result2: continue
				try: print("   " + result2)
				except: continue
				name=cpu=ram=storage=gpu=price=refresh=difference=None
				result2 = re.sub('<.+?>', '', result2)
				spec_list = result2.split('/')
				try: cpu = re.search('[iR][3579]( PRO)?\-\w+', spec_list[0]).group()
				except AttributeError: pass
				try: cpu = re.search('(Pentium|Celeron) \w+', spec_list[0]).group()
				except AttributeError: pass
				name = re.sub(cpu, '', spec_list[0]).strip('(')
				ram = spec_list[1].strip('/')
				storage = spec_list[2].replace(' SSD', '')
				try:
					refresh = re.search('d+Hz', result2).group()
					result2 = re.sub(refresh, '', result2)					
				except AttributeError: refresh = 'n/a'
				try: gpu = re.search('\w+([\s-]\w+)?', spec_list[3]).group()
				except AttributeError: print("GPU AttributeError")
				except IndexError: print("GPU IndexError")
				#try: gpu = re.search('(RTX)?30\d\d([\w\-]+)?', spec_list[3]).group()
				#except AttributeError: pass
				#try: gpu = re.search('Radeon', spec_list[3]).group()
				#except AttributeError: pass
				if len(spec_list) > 3:
					if "UHD" in spec_list[3]: gpu = "Intel UHD"
					if "Xe" in spec_list[3]: gpu = "Intel Iris Xe"
				price = re.findall('\$\d+', result2)
				if len(price) > 1:
					difference = int(price[1].strip('$')) - int(price[0].strip('$'))
					price = int(price[1].strip('$'))
				else: price = int(price[0].strip('$'))
				uhd_list = ["X515EA", "Pentium N6000"]
				for u in uhd_list:
					if u in spec_list[0]: gpu = "Intel UHD"
				if "J0052" in spec_list[0]: gpu = "Intel Iris Xe"
				if "Z90Q" in spec_list[0]: gpu = "Intel Iris Xe"
				parsing=[size,cpu,ram,storage,gpu,refresh]
				print(' => ' + ' | '.join(parsing))
				ws.append([series, name, size, cpu, ram, storage, gpu, price, difference])
			i+=1
		wb.save(filename)
	