## 利用openpyxl操作Excel
### 关键操作代码
```python3
# 打开一个工作簿
wb = openpyxl.load_workbook('censuspopdata.xlsx')

# 得到一个工作表
sheet = wb['Population by Census Tract']

# 得到全部工作表名
wb.sheetnames

# 得到活动表
wb.active

# 得到工作表名
sheet.title

# 得到一个单元格
c = sheet[B1]

# 得到行标/最大行标
c.row/c.max_row

# 得到列标/最大列标
c.column/c.max_column

# 单元格取值
c.value

# 单元格标志
c.coordinate

# 打开一个新的工作簿
wb = openpyxl.Workbook()

# 创建一个新的工作表
wb.create_sheet()

# 以指定名称创建一个工作表并设置序号
wb.create_sheet(index=0, title="First Sheet")

# 删除一个指定工作表
wb.remove(wb['Sheet1'])
```
