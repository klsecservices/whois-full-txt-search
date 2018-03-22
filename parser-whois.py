# -*- coding: utf-8 -*-
import urllib
import urllib2
import json
import xlwt
import xlrd
import sys, getopt
#import pprint
import os

def PrintHelp():
	msg = "Usage: python parser-whois.py <file> \r\n"
	msg += "This script performs full text search in WHOIS data of RIPE database.\r\n"
	msg += "Please, install xlwt and xlrdr packages (pip install <name>).\r\n"
	msg += "HTTPS protocol is used. Check your Internet connection before start.\r\n"
	msg += "Enter search patterns for your organisations line by line into input file.\r\n"
	msg += "Your result will be printed into result.xls file. Don't forget to validate it!\r\n"
	print msg
	sys.exit()

def RemoveDuplicates():

	tmp_file = xlrd.open_workbook('tmp.xls')
	write_file = xlwt.Workbook()

	for sheet in tmp_file.sheets():
		no_rows = sheet.nrows
		no_cols = sheet.ncols
		name = sheet.name
		gen_sheets = write_file.add_sheet(name)
		line_list = []
		r = 0
		for row in range(0, no_rows):
			line_sublist = [sheet.cell(row, col).value for col in range(0, no_cols)]
			if line_sublist not in line_list:
				line_list.append(line_sublist)
				for col in range(0, no_cols):
					gen_sheets.write(r,col,line_sublist[col])
				r = r + 1
	write_file.save('result.xls')

def main(argv):
	inputfile = ''
	try:
		opts, args = getopt.getopt(argv,"hi:",["ifile="])
	except getopt.GetoptError:
		print 'parser-whois.py <file>'
		sys.exit(2)
	for opt, arg in opts:
		if (opt == '-h') or (opt == '--help'):
			PrintHelp()
			sys.exit()
		else:
			inputfile = arg

#Handling arguments
if len(sys.argv) == 1: sys.argv[1:] = ["-h"]
main(sys.argv[1:])

opener = urllib2.build_opener()
headers = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 5.1; rv:10.0.1) Gecko/20100101 Firefox/10.0.1',
}
opener.addheaders = headers.items()

#Creating temporary workbook
file = open (sys.argv[1])
orgs = file.read().splitlines()

book = xlwt.Workbook()
sh = book.add_sheet('whois')

sh.row(0).write(0,'Network range')
sh.row(0).write(1,'Description')
row = 1
inetnum=''
info = ''

#Kostyl time. Seriosly, I don't know how to solve it in other way.
for j in range(1,10):
	for org in orgs:
		#Quering WHOIS for every organisation pattern
		url = 'https://apps.db.ripe.net/db-web-ui/api/rest/fulltextsearch/select?facet=true&format=xml&hl=true&q=('+urllib2.quote(org, safe='')+')&start=0&wt=json'
		print 'Proceeding '+org+' request'
		try:
			opener.open(url)
		except urllib2.HTTPError as e:
			print e.code
		else:
			#Loading results for every organisation
			resp = urllib2.urlopen(url)
			string = resp.read().decode('utf-8')
			json_obj = json.loads(string)
			if json_obj['result'] is not None:
				#Counting pages and exact amount of inetnums
				page_count = json_obj['result']['numFound'] / 10
				for data in json_obj['lsts']:
					if data['lst']['lsts'] is not None:
						for data2 in data['lst']['lsts']:
							if data2['lst']['lsts'] is not None:
								for data3 in data2['lst']['lsts']:
									if data3['lst']['ints'] is not None:
										for data4 in data3['lst']['ints']:
											if data4['int']['name'] is not None:
												if data4['int']['name'] == 'inetnum':
													print str(data4['int']['value'])+' inetnums were found'
				for i in range(0,page_count+1):
					#Quering WHOIS for every page. Yes, it sucks.
					url = 'https://apps.db.ripe.net/db-web-ui/api/rest/fulltextsearch/select?facet=true&format=xml&hl=true&q=('+urllib2.quote(org, safe='')+')&start='+str(i*10)+'&wt=json'
					resp = urllib2.urlopen(url)
					string = resp.read().decode('utf-8')
					json_obj = json.loads(string)
					#Parsing results. If you want to add or remove fields from WHOIS, do it here.
					if json_obj['result']['docs'] is not None:
						for data in json_obj['result']['docs']:
							#pprint.pprint (data)
							if data['doc']['strs'] is not None:
								for data2 in data['doc']['strs']:
									if data2['str']['name'] is not None:
										if data2['str']['name'] == 'inetnum':
											inetnum = data2['str']['value']
										if data2['str']['name'] == 'netname':
											info += data2['str']['name']+": "+data2['str']['value']+"\r\n"
										if data2['str']['name'] == 'descr':
											info += data2['str']['name']+": "+data2['str']['value']+"\r\n"
										if data2['str']['name'] == 'country':
											info += data2['str']['name']+": "+data2['str']['value']
								if inetnum != '':
									sh.row(row).write(0,inetnum)
									sh.row(row).write(1,info)
									row = row + 1
									netrange_only=0
								info = ''
								inetnum = ''

book.save('tmp.xls')
file.close()
RemoveDuplicates()
os.remove('tmp.xls')