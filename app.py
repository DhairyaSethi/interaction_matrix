import xlsxwriter
import os
from flask import Flask, request, send_from_directory, render_template, current_app

app = Flask(__name__)


@app.route('/')
def index():
	return render_template('index.html')

@app.route('/result', methods = ['POST', 'GET'])
def result():
	if request.method=='POST':
		data = request.form['components'].split(',')
		script(data)
#		return render_template('result.html', result = data)
		dir = os.path.join(current_app.root_path, 'downloads')
		return send_from_directory(dir, "hello.xlsx", as_attachment = True)


if __name__ == '__main__':
	app.run(debug=False, port=8000)



def script(data):
	workbook = xlsxwriter.Workbook('./downloads/hello.xlsx')
	worksheet = workbook.add_worksheet()
	k = 1
	i = 0

	length_list = [len(x) for x in data]
	max_width = max(length_list)

	cell_bold = workbook.add_format({'bold': True, 'center_across': True, 'bg_color': '#f0f0c7'})
	cell_grey = workbook.add_format()
	cell_grey.set_bg_color('#D3D3D3')

	for item in data:
		worksheet.write(k, 0, item, cell_bold)
		worksheet.write(0, k, item, cell_bold)
		worksheet.write(k, k,'' , cell_grey)
		worksheet.set_column(i, i, max_width)
		k+=1
		i+=1
	worksheet.set_column(len(data), len(data), max_width)


	workbook.close()