# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "Monitor_" + now.strftime("%Y-%m-%d") + ".xlsx"

class coolpcSpider(scrapy.Spider):
	name = "monitor"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)

	def parse(self, response):
		wb = openpyxl.Workbook()
		ws = wb.active
		monitor=response.xpath('//select/optgroup')[12]
		i=0
		panel_except = ['Ark', 'ROG', 'G5', 'MPG ARTYMIS', 'MAG ARTYMIS', 'TUF', 'Modern', 'PRO', 'QHD', 'Summit', '白']
		ws.append(['Name', 'Size', 'Resolution', 'Input', 'Response', 'Panel', 'Refresh', 'HDR', 'Sync', 'Speaker', 'Price', 'Spec_others'])
		for row in monitor.xpath('.//optgroup/@label'):
			result=row.extract()
			if "投影機" in result: break
			print(result)
			size=resolution=None
			try:
				size=re.search('\d+吋', result).group()
				resolution=re.search('\d+\*\d+', result).group()
			except: pass
			for row2 in monitor.xpath('.//optgroup')[i].xpath('./option'):
				result2=row2.extract()
				if "掛架" in result2: continue
				if "$" not in result2: continue
				if "市價" in result2: continue
				try: print('  ' + result2)
				except: continue
				name=spec=input=response=panel=refresh=hdr=sync=speaker=price=spec_others=None
				name = re.sub('<.+?>', '', result2)
				name = re.sub(',.+$', '', name)
				name = re.sub('\$.+$', '', name)
				spec = re.findall('\(.+?\)', result2)
				spec = max(spec, key=len)
				name = name.replace(spec, ' ')
				if 'VA含喇叭' in spec: spec = spec.replace('VA含喇叭', 'VA/含喇叭')
				spec = spec.strip('(').strip(')')
				spec_list = spec.split('/')
				for j in spec_list: 
					if re.match(r'(\d[A-P])+', j): input = j; spec_list.remove(j); break
				for j in spec_list: 
					if re.match(r'\d(\.\d)?ms', j): response = j; spec_list.remove(j); break
				for j in spec_list: 
					if re.match(r'\d+Hz', j): refresh = j; spec_list.remove(j); break
				for j in spec_list: 
					if re.match(r'HDR\d+', j): hdr = j; spec_list.remove(j); break
				for j in spec_list: 
					if re.match(r'\s?(Ad|Fr|G-).+$', j): sync = j; spec_list.remove(j); break
				if sync == None:
					match = re.search(r'G-.+?兼容', name)
					try:
						sync = name[match.start():match.end()]
						name = name.replace(sync, '')
					except: sync = ''
				if '無喇叭' in spec_list: speaker = 'No'; spec_list.remove('無喇叭')
				if '含喇叭' in spec_list: speaker = 'Yes'; spec_list.remove('含喇叭')
				if spec_list[0] in panel_except: panel = spec_list[1]; spec_list.pop(1)
				else: panel = spec_list[0]; spec_list.pop(0)
				if panel == 'VA165Hz': panel = 'VA'; refresh = '165Hz'
				if panel == '4m': panel = spec_list[0]; response = '4ms'; spec_list.pop(0)
				if panel == '': panel = spec_list[0]; spec_list.pop(0)
				print(' => ' + input, end = '')
				try: print(' | ' + response, end = '')
				except TypeError: pass
				print(' | ' + panel, end = '')
				try: print(' | ' + refresh); print(' => Others: ', end = '') 
				except TypeError: print('\n => Others: ', end = '')
				print(spec_list)
				price = re.findall('\$\d+', result2)
				price = min(price)
				spec_others = ' '.join(spec_list)
				ws.append([name.rstrip(' '), size, resolution, input, response, panel, refresh, hdr, sync.lstrip(' '), speaker, int(price.strip('$')), spec_others])
			i+=1
		wb.save(filename)
	