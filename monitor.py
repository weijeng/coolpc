# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "Monitor_" + now.strftime("%Y-%m-%d") + ".xlsx"
currentTime = time.strftime("%Y%m%d_%H%M%S")
rawData = open("Monitor_%s.log" %currentTime, "a+")
wb = openpyxl.Workbook()

class coolpcSpider(scrapy.Spider):
	name = "monitor"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)

	def parse(self, response):
		ws = wb.active
		monitor=response.xpath('//select/optgroup')[12]
		i=0
		panel_except = ['Ark', 'ROG', 'G5', 'MPG ARTYMIS', 'MAG ARTYMIS', 'TUF', 'Modern', 'PRO', 'QHD', 'Summit', '白', 'Odyssey']
		r1920 = ['ED320QR', '22B2HN', 'C27T550FDC', 'C32T550FDC', 'DA430', 'MP271']
		r2560 = ['PD2500Q']
		ws.append(['Name', 'Size', 'Resolution', 'Input', 'Response', 'Panel', 'Refresh', 'HDR', 'Sync', 'Speaker', 'Price', 'Difference', 'Spec_others'])
		for row in monitor.xpath('.//optgroup/@label'):
			result=row.extract()
			if "外接式" in result: break
			print(result)
			size=resolution=None
			try:
				size=re.search('\d+吋', result).group()
			except AttributeError: size = 'n/a'
			try:
				resolution=re.search('\d+\*\d+', result).group()
			except AttributeError: resolution = 'n/a'
			for row2 in monitor.xpath('.//optgroup')[i].xpath('./option'):
				result2=row2.extract()
				if "掛架" in result2: continue
				if "disabled" in result2: continue
				try: print(result2)
				except: print(" debug:fail to print result2"); continue
				name=spec=input=response=panel=refresh=hdr=sync=speaker=price=difference=spec_others=resolution2=""
				name = re.sub('<.+?>', '', result2)
				name = re.sub(',.+$', '', name)
				name = re.sub('\$.+$', '', name)
				spec = re.findall('\(.+?\)', result2)
				spec = max(spec, key=len)
				name = name.replace(spec, ' ')
				if 'VA含喇叭' in spec: spec = spec.replace('VA含喇叭', 'VA/含喇叭')
				if 'VA無喇叭' in spec: spec = spec.replace('VA無喇叭', 'VA/無喇叭')
				if '含喇叭FreeSync' in spec: spec = spec.replace('含喇叭FreeSync', '含喇叭/FreeSync')
				spec = re.sub("VA165Hz", "VA/165Hz", spec)
				spec = spec.strip('(').strip(')')
				spec_list = spec.split('/')
				for j in spec_list: 
					if re.match(r'(\d[A-P])+', j): input = j
					if re.match(r'\d(\.\d)?ms', j): response = j
					if re.match(r'\d+Hz', j): refresh = j
					if re.match(r'HDR\d+', j): hdr = j
					if re.match(r'\s?(Ad|Fr|G-).+$', j): sync = j
					if re.match(r'TN|(Rapid|Nano|Fast)?\s?IPS|VA(曲面)?|OLED(曲面)?', j): panel = j
				if sync == "":
					match = re.search(r'G-.+?兼容', name)
					try:
						sync = name[match.start():match.end()]
						name = name.replace(sync, '')
						print(" => debug:G")
					except: sync = ''
				if '無喇叭' in spec_list: speaker = 'No'; spec_list.remove('無喇叭')
				if '含喇叭' in spec_list: speaker = 'Yes'; spec_list.remove('含喇叭')
				for j in r1920: 
					if j in name: resolution2 = '1920*1080'
				for j in r2560: 
					if j in name: resolution2 = '2560*1440'
				price = re.findall('\$\d+', result2)
				if len(price) > 1:
					difference = int(price[1].strip('$')) - int(price[0].strip('$'))
					price = int(price[1].strip('$'))
				else: price = int(price[0].strip('$'))
				parsing=[input,response,panel,refresh,hdr,sync.lstrip(' ')]
				print(' => ' + ' | '.join(parsing))
				spec_others = '/'.join(spec_list)
				if resolution2 == "":
					ws.append([name.rstrip(' '), size, resolution, input, response, panel, refresh, hdr, sync.lstrip(' '), speaker, price, difference, spec_others])
				else:
					ws.append([name.rstrip(' '), size, resolution2, input, response, panel, refresh, hdr, sync.lstrip(' '), speaker, price, difference, spec_others])
				rawData.write(result2 + '\n')
			i+=1
		rawData.flush()
		wb.save(filename)
	