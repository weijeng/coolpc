import scrapy
import re
import openpyxl
import datetime
now=datetime.datetime.now()
filename = "CPU_" + now.strftime("%Y-%m-%d") + ".xlsx"

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
		ws.append(["Name", "Cores", "Watt", "Price", "Singel Score", "Multiple Score", "Score/Price"])
		dict_cpu = {}
		for row in response.xpath('//select/optgroup')[3].xpath('.//option'):
			cpu=cores=price=watt=None
			line=row.extract()
			try: print line
			except: continue
                        try: price = re.search('\$\d+', line).group()
                        except: continue
			if "R3" in line: line=line.replace("R3", "AMD Ryzen 3")
			if "R5" in line: line=line.replace("R5", "AMD Ryzen 5")
			if "R7" in line: line=line.replace("R7", "AMD Ryzen 7")
			if "R9" in line: line=line.replace("R9", "AMD Ryzen 9")
			if "TR" in line: line=line.replace("TR", "Threadripper")
			if "TR2" in line: line=line.replace("TR2", "Threadripper")
			if "G5400" in line: line=line.replace("G5400", "Gold G5400")
			if "Intel i" in line: line=line.replace("Intel i","Intel Core i")
			if "AMD AMD" in line: line=line.replace("AMD AMD", "AMD")
			try: watt = re.search('[\/\)]\d+W\/?', line).group().strip("W/").replace("/","").replace(")","")
			except: pass
			cpu = re.search('(Intel|AMD)[\s\-\w\+]+', line).group()
			if 'Intel Xeon E5-2620 V4' in cpu: watt = "85"
                        if watt == None: continue
			try: cores = re.search('\d+[^\w\s,]\/\d+[^\w,](GPU)?', line).group()
			except AttributeError: cores = re.search('\d[^\w\s\"]{1,2}(\d[^\w]{1})?', line).group()
			if dict_cpu.get(cpu) == None: dict_cpu[cpu] = cores, watt, int(price.strip("$"))
			elif dict_cpu.get(cpu)[2] > int(price.strip("$")) and dict_cpu.get(cpu)[2] - int(price.strip("$")) < 5000:
				watt = dict_cpu.get(cpu)[1]
				dict_cpu[cpu] = cores, watt, int(price.strip("$"))
			else: continue
		i=2
		for x, y in dict_cpu.items():
			j=str(i)
			ws.append([ x, y[0], y[1], y[2], "=VLOOKUP(A" + j + ",Geek_single!A:B,2,0)", "=VLOOKUP(A" + j + ",Geek_multiple!A:B,2,0)", "=(E"+j+"*2+F"+j+")/D"+j ])
			i+=1
		wb.save(filename)
		
	def parse_geek(self, response):
		wb = openpyxl.load_workbook(filename)
		ws = wb.create_sheet(title="Geek_single")
		latest = 0
		score = 0
		for row in response.selector.xpath("//tr/td[@class]"):
			line = row.extract()
			#try: print line
			#except: pass
			if "name" in line:
				cpu = re.search('(Intel|AMD|HP|A4)( [\-\w\+]+)+', line).group()
			elif "score" in line: 
				score = re.search('\d+', line).group()
			if int(score) - int(latest) > 10000:
				print "==> change to multiple sheet"
				ws = wb.create_sheet(title="Geek_multiple")
			if score > 0:
				ws.append([cpu, score])
				#print " => cpu: " + cpu
				#print " => score: " + score + "\n"
				latest = score
				score = 0
		wb.save(filename)
