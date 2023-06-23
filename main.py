# A program that creates a Virtual Workbook
# Adds data from haveibeenpwned API 
# Uses OpenPyXL to write data that involves the company name, date, and number of breaches that has occured into excel sheets
# Create Excel chart on sheets which covers following :
# BarChart() - We covered this in our class
# LineChart()
# ScatterChart()
# PieChart()
# Stores the xlxs file with the name breaches1.xlsx


import openpyxl  # imports openpyxl
import random  # imports random
from openpyxl import Workbook  # imports excel workbook
from openpyxl.chart import BarChart, LineChart, ScatterChart, PieChart, Reference  # imports chart
import requests  # imports request
  
wb = Workbook()  # opens workbook
sheet1 = wb.active   # makes cells active

url = 'https://haveibeenpwned.com/api/v2/breaches'  # api that identifies companies that had breaches
response = requests.get(url)  # gets url
data = response.json()  


random_company = random.sample(data, 10)  #randomly selects sample data
sort_company = sorted(random_company, key=lambda x: x["PwnCount"], reverse=True)  # sorts data from highest to lowest

# cloumn headers
sheet1['A1'] = "Breach Date"  
sheet1['B1'] = "Comapany Name"  
sheet1['C1'] = "Domain"  
sheet1['D1'] = "PwnCount"

#lops through sorted data
for i, breach in enumerate(sort_company, start=2):
  Breach_Date = breach["BreachDate"]
  Company_Name = breach["Title"]
  Domain = breach["Domain"]
  PwnCount = breach["PwnCount"]
  sheet1['A' + str(i)] = Breach_Date
  sheet1['B' + str(i)] = Company_Name
  sheet1['C' + str(i)] = Domain
  sheet1['D' + str(i)] = PwnCount

# define the data range for the chart
data_range = Reference(sheet1, min_col=4, min_row=1, max_col=4, max_row=11)
categories = Reference(sheet1, min_col=2, min_row=2, max_row=11)

# plots bar chart
bar_chart = BarChart()
bar_chart.add_data(data_range, titles_from_data=True)
bar_chart.set_categories(categories)
bar_chart.title = "PwnCount by Company"
sheet1.add_chart(bar_chart, "A14")
bar_chart.x_axis.majorGridlines = None
bar_chart.y_axis.majorGridlines = None

# plots line chart
line_chart = LineChart()
line_chart.add_data(data_range, titles_from_data=True)
line_chart.set_categories(categories)
line_chart.title = "PwnCount by Company"
sheet1.add_chart(line_chart, "H14")
line_chart.x_axis.majorGridlines = None
line_chart.y_axis.majorGridlines = None

#plots scatter chart
scatter_chart = ScatterChart()
scatter_chart.add_data(data_range, titles_from_data=True)
scatter_chart.set_categories(categories)
scatter_chart.title = "PwnCount by Company"
sheet1.add_chart(scatter_chart, "A29")
scatter_chart.x_axis.majorGridlines = None
scatter_chart.y_axis.majorGridlines = None

#plots pie chart
# pie_chart = PieChart()
# pie_chart.add_data(data_range, titles_from_data=True)
# pie_chart.set_categories(categories)
# pie_chart.title = "PwnCount by Company"
# pie_chart.add_data(pie_chart, "H29")
# pie_chart.x_axis.majorGridlines = None
# pie_chart.y_axis.majorGridlines = None

# saves file
file_name = "breaches1.xlsx"
wb.save(file_name)
