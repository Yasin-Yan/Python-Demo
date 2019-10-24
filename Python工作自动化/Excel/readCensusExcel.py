import openpyxl, pprint

print('打开工作簿...')
wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb['Population by Census Tract']
countyData = {}

print('读取行...')
for row in range(2, sheet.max_row + 1):
	state = sheet['B' + str(row)].value
	county = sheet['C' + str(row)].value
	pop = sheet['D' + str(row)].value

	# 确保state键存在
	countyData.setdefault(state, {})
	# 确保county键存在
	countyData[state].setdefault(county, {'tracts': 0, 'pop': 0})
	# 每行代表一个普查区域，每次+1
	countyData[state][county]['tracts'] += 1
	# 将普查区域的人口加到郡的人口中来
	countyData[state][county]['pop'] += int(pop)

# 打开一个新的text文件写入countyData的内容
print('正在写入结果...')
resultFile = open('Census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print('完成.')