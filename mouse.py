# -*- coding: utf-8 -*-

import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "mouse_" + now.strftime("%Y-%m-%d") + ".xlsx"
currentTime = time.strftime("%Y%m%d_%H%M%S")
rawData = open("Mouse_%s.log" %currentTime, "a+")
wb = openpyxl.Workbook()

class coolpcSpider(scrapy.Spider):
	name = "mouse"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)

	def parse(self, response):
		ws = wb.active
		mouse=response.xpath('//select/optgroup')[17]
		i=0
		size={"M171":[98,62,35], "M190":[115,66,40], "M221":[99,60,39], "M235n":[95,55,39], "M280":[105,68,38],
		"G304":[117,62,38], "M310T":[111,62,39], "M325":[95,57,39], "M331 SilentPlus":[105,68,38], "M350":[107,59,27],
		"Ms350r":[95,58,36], "M550r":[98,59,37], "M830r":[110,80,41], "M850r":[126,82,45], "M900r":[101.5,70.5,38], "M985r":[122,66,66],
		"M585":[103,64,40], "M650":[108,61,39], "Ms650r":[100,75,40], "Ms930r":[116,82.5,53], "Ms950r":[125,84,45],
		"irocks M23":[106,57,33.5], "S1000 Plus":[104,61,34], "B700r":[110,55,25], "Tuf Gaming M4":[126,63.5,40],
		"Razer Atheris":[100,63,34], "Orochi V2":[108,60,38], "Basilisk":[130,60,42], "DeathAdder V2":[127,62,43],
		"Katar Pro Wireless":[116,64,38], "Pop Mouse":[105,59,35], "SteelSeries Rival 3":[121,58,21.5],
		"Cherry Mw5180":[107,62,31]}
		ws.append(['Name', '有線', '無線', '藍芽', 'DPI', 'Weight', 'RGB', 'Price', 'Difference', '長', '寬', '高', 'Spec'])
		for line in mouse.xpath('.//optgroup/@label'):
			vendor=None
			vendor=line.extract()
			if "促銷" in vendor: i+=1; continue
			if "簡報" in vendor: i+=1; continue
			if "鼠墊" in vendor: i+=1; continue
			if "繪圖板" in vendor: i+=1; continue
			for row in mouse.xpath('.//optgroup')[i].xpath('./option'):
				name=price=parsing=size2=difference=None
				dpi=rgb=weight=""
				w1=w2=w3="N"
				result=row.extract()
				if "disabled" in result: continue
				#if "登錄" in result: continue
				#if "購買" in result: continue
				#if "加購" in result: continue
				if "鼠線夾" in result: continue
				try: print(result)
				except: continue
				result = result.replace('無線/藍芽', '無線-藍芽').replace('藍芽/無線', '無線-藍芽')
				name = re.sub('<.+?>', '', result)
				name = re.sub('dpi送啟動卡,', 'dpi,', name)
				name = re.sub('\s?,.+$', '', name)
				name = re.sub(' 原價.+$', '', name)
				name = re.sub('\+ Speed', '/Speed', name)
				try: spec = re.search('/.+', name).group().strip(',')
				except AttributeError: continue
				name = name.replace(spec, '')
				spec_list = spec.split('/')
				spec_list.remove('')
				print(spec_list)
				for j in spec_list:
					if re.match(r'有線-無線-藍芽', j): w1='Y'; w2='Y'; w3='Y'; break
					if re.match(r'有線[\-\+]無線', j): w1='Y'; w2='Y'; break
					if "有線-藍芽-2.4g" in j: w1='Y'; w2='Y'; w3='Y'; break
					if re.match(r'無線[\-\+]藍芽(5.1)?$', j): w2='Y'; w3='Y'; break
					if re.search(r'2\.4[gG][\-\+]藍芽', j): w2='Y'; w3='Y'; break
					if re.match(r'有線', j): w1='Y'; break
					if re.match(r'無線$', j): w2='Y'; break
					if "藍芽" in j: w3='Y'; break
				if "2.4G" in name: w2='Y'
				if "無線" in name: w2='Y'
				if "Pop Mouse" in name: w3='Y'
				if "有線" in name: w1='Y'
				if "威剛 XPG Alpha 電競滑鼠" in name: w1='Y'
				if "M425G" in name: w1='Y'
				if "M585" in name: w3='Y'
				if "Harpoon Rgb Wireless" in name: w1='Y'; w3='Y'
				if "Katar Pro Wireless" in name: w3='Y'
				for j in spec_list:
					if re.match(r'\d+[Dd](pi|PI)', j): dpi = j.strip('dpi').strip('D').strip('DPI'); break
				for j in spec_list:
					if re.match(r'R(gb|GB)', j): rgb = 'Yes'; break
				for j in spec_list:
					if re.match(r'^\d+(g|克)', j): weight = j.strip('輕量化'); break
				price = re.findall('\$\d+', result)
				if len(price) > 1:
					difference = int(price[1].strip('$')) - int(price[0].strip('$'))
					price = int(price[1].strip('$'))
				else: price = int(price[0].strip('$'))
				parsing=[w1,w2,w3,dpi,weight,rgb]
				print(' => ' + ' | '.join(parsing))
				spec_others = '/'.join(spec_list)
				for x, y in size.items():
					if x in name: size2=y; break
					else: size2=['','','']
				ws.append([name.rstrip(' '), w1, w2, w3, dpi, weight, rgb, price, difference, size2[0], size2[1], size2[2], spec_others])
				rawData.write(result + '\n')
				rawData.flush()
			i+=1		
		wb.save(filename)
	
