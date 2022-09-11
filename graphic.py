# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "Graphic_" + now.strftime("%Y-%m-%d") + ".xlsx"
wb = openpyxl.Workbook()

class coolpcSpider(scrapy.Spider):
	name = "graphic"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
			'https://benchmarks.ul.com/compare/best-gpus'
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)
		yield scrapy.Request(url=urls[1], callback=self.benchmarks)

	def parse(self, response):
		ws = wb.active
		graphic=response.xpath('//select/optgroup')[11]
		i=0
		ws.append(['Name', 'Model', 'Core clock', 'Memory', 'Length', 'Warranty', 'Price', 'Others'])
		for line in graphic.xpath('.//optgroup/@label'):
			model=None
			model=line.extract()
			if "Quadro" in model: i+=1; continue
			if "外接顯卡" in model: i+=1; continue
			print(model + ', i = ' + str(i))
			for row in graphic.xpath('.//optgroup')[i].xpath('./option'):
				name=core=memory=length=warranty=price=None
				result=row.extract()
				if "$" not in result: continue
				try: print('  ' + result)
				except: continue
				spec = re.findall('\(.+?\)', result)
				spec = max(spec, key=len)
				name = re.sub('<.+?>', '', result)
				name = re.sub(',.+$', '', name)
				name = re.sub('\$.+$', '', name)
				name = name.replace(spec, ' ')
				spec = spec.strip('(').strip(')')
				spec_list = spec.split('/')
				for j in spec_list:
					if re.match(r'(OC上)?\d+M[Hh][zZ]', j): core = j; spec_list.remove(j); break
					if re.match(r'\d+Hz', j):
						core = re.sub(r'(\d+)Hz', r'\1MHz', j)
						spec_list.remove(j); break
				if core == None: core = 'n/a'
				if 'Z' in core: core = core.replace('Z', 'z')
				if 'OC上' in core: core = core.replace('OC上', '')
				if 'h' in core: core = core.replace('h', 'H')
				for j in spec_list: 
					if re.match(r'\w+\s?DDR\d', j): memory = j; spec_list.remove(j); break
				if memory == None:
					match = re.search(r'\dGB\sG?DDR\d', name)
					try:
						memory = name[match.start():match.end()]
						name = name.replace(memory, '')
					except: memory = 'n/a'
				for j in spec_list: 
					if re.match(r'[\d\.]+cm', j): length = j; spec_list.remove(j); break
				if length == None: length = 'n/a'
				for j in spec_list: 
					if re.match(r'註?三年保?', j): warranty = '3 years'; spec_list.remove(j)
					if re.match(r'註?冊?四年保?', j): warranty = '4 years'; spec_list.remove(j); break
					if re.match(r'註?冊?五年保?', j): warranty = '5 years'; spec_list.remove(j); break
				price = re.findall('\$\d+', result)
				price = min(price)
				print('  => ' + model + ' | ' + core + ' | ' + memory + ' | ' + length)
				spec_others = ' '.join(spec_list)
				ws.append([name.rstrip(' '), model, core, memory, length, warranty, int(price.strip('$')), spec_others])
			i+=1
		
	def benchmarks(self, response):
		ws = wb.create_sheet("score")
		gpu = response.xpath('//td/a/text()').extract()
		score = response.xpath('//td[@class="small-pr1"]/div/div/span/text()').extract()
		price = response.xpath('//td[@class="list-tiny-none"]/span/div/text()').extract()
		for x, y, z in zip(gpu, score, price):
			ws.append([x.strip(), y.strip(), z.strip()])
		wb.save(filename)
	