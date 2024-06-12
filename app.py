# app.py
from flask import Flask, render_template, request, send_file
from main import get_cell_value, write_to_excel

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        file1 = request.files['file1']
        file2 = request.files['file2']

        # Save uploaded files
        file1_path = f"uploads/{file1.filename}"
        file2_path = f"uploads/{file2.filename}"
        file1.save(file1_path)
        file2.save(file2_path)

        # Define sheet names and cells to compare for each file
        sheet_name_file1 = 'NOC Checklist'
        sheet_name_file2 = 'Sheet1'
        cells_to_compare_file1 = ['D3', 'C3', 'D7', 'C7']
        cells_to_compare_file2 = ['B9', 'G9', 'C9', 'H9']
        cells_to_compare_file1V1 = ['B14','C14','B15','D14']
        cells_to_compare_file1V2 = ['A14','C15','A15','D15']
        cells_to_compare_file2V1 = ['D9','Z8','AA8','AJ8','AK8','AQ8','AR8','BA8','BB8']
        cells_to_compare_file2V2 = ['I9','Z9','AA9','AJ9','AK9','AQ9','AR9','BA9','BB9']

        # Retrieve cell values from files
        cell_value_file1 = get_cell_value(file1_path, sheet_name_file1, cells_to_compare_file1)
        cell_value_file2 = get_cell_value(file2_path, sheet_name_file2, cells_to_compare_file2)
        cell_value_file2V1 = get_cell_value(file2_path, sheet_name_file2, cells_to_compare_file2V1)
        cell_value_file2V2 = get_cell_value(file2_path, sheet_name_file2, cells_to_compare_file2V2)
        cell_value_file1V1 = get_cell_value(file1_path, sheet_name_file1, cells_to_compare_file1V1)
        cell_value_file1V2 = get_cell_value(file1_path, sheet_name_file1, cells_to_compare_file1V2)

        # Process files and generate output
        output_filename = 'output.xlsx'
        write_to_excel(cell_value_file1, cell_value_file2, cell_value_file2V1, cell_value_file2V2, cell_value_file1V1, cell_value_file1V2, output_filename)

        return send_file(output_filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
