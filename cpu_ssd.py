import scrapy
import re
import openpyxl
import datetime
import time
now=datetime.datetime.now()
filename = "CPU_SSD_" + now.strftime("%Y-%m-%d") + ".xlsx"

class coolpcSpider(scrapy.Spider):
	name = "cpu"
	def start_requests(self):
		urls = [
			'http://www.coolpc.com.tw/evaluate.php',
			'https://browser.geekbench.com/processor-benchmarks/'
		]
		yield scrapy.Request(url=urls[0], callback=self.parse)
		yield scrapy.Request(url=urls[1], callback=self.parse_geek)

	def parse(self, response):
		wb = openpyxl.Workbook()
		ws = wb.active
		ws.title = "CPU"
		ws.append(["Name", "Cores", "Watt", "Price", "Singel Score", "Multiple Score", "Score/Price"])
		dict_cpu = {}
		for row in response.xpath('//select/optgroup')[3].xpath('.//option'):
			cpu=cores=price=watt=None
			line=row.extract()
			try: print(line)
			except: print(" ==> Can not print"); continue
			if "ROG STRIX" in line: continue
			if "+AMD" in line: continue
			if "NT.$" in line: continue
			price = re.findall('\$\d+', line)
			if len(price) > 1: price = price[1]
			elif len(price) == 1: price = price[0]
			else: print(" ==> Does not find price"); continue
			line = re.sub(r'(AMD )?R(\d)', r'AMD Ryzen \2', line)
			line = re.sub(r'(AMD )?Athlon', r'AMD Athlon', line)
			line = re.sub('TR', 'Threadripper', line)
			line = re.sub('Intel i', 'Intel Core i', line)
			if "G5400" in line: line=line.replace("G5400", "Gold G5400")
			cpu = re.search('(Intel|AMD)[\s\-\w\+]+', line).group()
			try: watt = re.search('[\/\)]\d+W\/?', line).group().strip("W/").replace("/","").replace(")","")
			except: pass
			if 'Intel Xeon E5-2620 V4' in cpu: watt = "85"
			try: cores = re.search('\d+核\/\d+緒', line).group()
			except AttributeError: cores = re.search('\d+([^\w\s\"\-\.\/]{1,2}|C\d+T)', line).group()
			if dict_cpu.get(cpu) == None: dict_cpu[cpu] = cores, watt, int(price.strip("$"))
			elif dict_cpu.get(cpu)[2] > int(price.strip("$")) and dict_cpu.get(cpu)[2] - int(price.strip("$")) < 5000:
				watt = dict_cpu.get(cpu)[1]
				dict_cpu[cpu] = cores, watt, int(price.strip("$"))
			else: continue
		i=2
		for x, y in dict_cpu.items():
			j=str(i)
			s1 = "=VLOOKUP(A" + j + ",Geek_single!A:B,2,0)"
			s2 = "=VLOOKUP(A" + j + ",Geek_multiple!A:B,2,0)"
			sp_ratio = "=(E"+j+"*2+F"+j+")/D"+j
			if "G6400" in x: s1 = 1017; s2 = 2405
			if "G6405" in x: s1 = 1074; s2 = 2417
			if "5600G" in x: s1 = 1508; s2 = 7610
			if "5700G" in x: s1 = 1581; s2 = 9407
			if "10105F" in x: s1 = 1161; s2 = 4558
			ws.append([ x, y[0], y[1], y[2], s1, s2, sp_ratio ])
			i+=1
		wb.save(filename)
		time.sleep(3)
		
		# SSD
		ws = wb.create_sheet("SSD")
		ws.append(["Name", "Size", "Read", "Write", "Type", "Price"])	
		for row in response.xpath('//select/optgroup')[6].xpath('.//option'):
			name=size=read=write=type=price=None
			result=row.extract()
			try: print(result)
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
		wb.save(filename)
		time.sleep(3)
		
	def parse_geek(self, response):
		wb = openpyxl.load_workbook(filename)
		ws = wb.create_sheet("Geek_single")
		latest = 0
		score = 0
		for row in response.selector.xpath("//tr/td[@class]"):
			line = row.extract()
			if "name" in line:
				cpu = re.search('(Intel|AMD|HP|A4)( [\-\w\+]+)+', line).group()
			elif "score" in line: 
				score = re.search('\d+', line).group()
			if int(score) - int(latest) > 10000:
				ws = wb.create_sheet("Geek_multiple")
			if int(score) > 0:
				ws.append([cpu, score])
				latest = score
				score = 0
		wb.save(filename)

