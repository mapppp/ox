from openpyxl import load_workbook, Workbook


class Divider:
	def __init__(self):
		self.rb = Workbook()
	
	def join(self, xname):
		print("从"+xname+"中读取数据")
		wb = load_workbook("%s.xlsx" % xname)
		sheets = wb.sheetnames
		for sheet in sheets:
			for row in wb[sheet]["A"]:
				self.rb.active.append([row.value, ])
		return self.rb

	def divide(self, d_num, f_name):
		sheet = self.rb.active
		sheet.title = "总表"
		a = list(sheet["A"])
		end_num = sheet.max_row % d_num
		print("判断行数是否是"+str(d_num)+"的整数")
		if end_num == 0:
			e = 1
		else:
			e = 2
		print("创建表并将数据保存")
		for xname in range(1, int(sheet.max_row / d_num) + e):
			print("判断剩余的数据量是否够取"+str(d_num)+"个")
			if len(a) > end_num:
				print("提取数据")
				group = a[0:d_num]
				print("创建新表"+str(xname))
				self.rb.create_sheet("分表" + str(xname))
				f_s = self.rb["分表" + str(xname)]
				print("删除总表中提取出来的数据")
				for i in group:
					a.remove(i)
				print("将数据添加到分表中")
				c_n = 1
				for i in group:
					f_s["A" + str(c_n)] = i.value
					c_n += 1
			else:
				if len(a) != 0:
					print("创建新表"+str(xname))
					self.rb.create_sheet("分表" + str(xname))
					f_s = self.rb["分表" + str(xname)]
					print("将数据添加到分表中")
					c_n = 1
					for i in a:
						f_s["A" + str(c_n)] = i.value
						c_n += 1
				else:
					pass
				# print(a)
		self.rb.save(f_name+".xlsx")


if __name__ == "__main__":
	d = Divider()
	d.join("数据5")
	d.divide(2000, "数据6")
