from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker, DataPoint

wb = Workbook()
ws = wb.active
ws.title = "Business Data"

data = [
    ['Department', 'Employees', 'Profit (k$)'],
    ['Sales', 50, 120],
    ['Marketing', 40, 100],
    ['R&D', 70, 150],
    ['HR', 30, 60],
    ['Support', 45, 80],
    ['IT', 60, 130],
]

for row in data:
    ws.append(row)

chart = ScatterChart()
chart.title = "Employees vs Monthly Profit"
chart.x_axis.title = "Number of Employees"
chart.y_axis.title = "Monthly Profit (in $1000s)"
chart.scatterStyle = 'marker'

colors = ["FF0000", "00FF00", "0000FF", "FFA500", "800080", "008080"]
for i, color in enumerate(colors, start=2):
    x = Reference(ws, min_col=2, min_row=i, max_row=i)
    y = Reference(ws, min_col=3, min_row=i, max_row=i)
    series = Series(y, x, title=data[i-1][0])
    series.marker = Marker('diamond')
    series.marker.size = 8
    series.graphicalProperties.line.width = 0
    series.graphicalProperties.solidFill = color
    chart.series.append(series)

ws.add_chart(chart, "E2")

wb.save("scatter_chart.xlsx")
