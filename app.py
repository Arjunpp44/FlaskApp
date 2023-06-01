import base64
import os

import plotly
from flask import Flask, request, render_template
from plotly.graph_objs import Figure

from enc import aesEncrypt, aesDecrypt, get_data_from_excel, create_quick_exel

app = Flask(__name__)


@app.route('/')
def hello_world():  # put application's code here
    return render_template('form.html')


@app.route('/plotly', methods=['GET'])
def plotlyl():
    try:
        str = Figure({
    'data': [{'marker': {'color': 'green'},
              'mode': 'lines+markers',
              'name': 'Desired',
              'type': 'scatter',
              'x': ['STAKEHOLDER MANAGEMENT', 'PROCESS EXCELLENCE', 'COLLABORATION',
                    'RESULT ORIENTATION', 'DEVELOPING TALENT', 'STRATEGIC AGILITY',
                    'CUSTOMER CENTRICITY', 'INNOVATION AND CREATIVITY'],
              'y': [4.0, 4.0, 4.0, 4.0, 4.0, 4.0, 4.0, 4.0]},
             {'marker': {'color': '#00008B'},
              'mode': 'lines+markers',
              'name': 'Actual',
              'type': 'scatter',
              'x': ['STAKEHOLDER MANAGEMENT', 'PROCESS EXCELLENCE', 'COLLABORATION',
                    'RESULT ORIENTATION', 'DEVELOPING TALENT', 'STRATEGIC AGILITY',
                    'CUSTOMER CENTRICITY', 'INNOVATION AND CREATIVITY'],
              'y': [2.46, 3.07, 2.71, 2.54, 2.45, 1.48, 2.42, 2.38]}],
    'layout': {'paper_bgcolor': 'rgba(0,0,0,0)',
               'plot_bgcolor': 'rgba(0,0,0,0)',
               'xaxis': {'title': {'text': 'Competencies'}},
               'yaxis': {'title': {'text': 'Score'}}}
})
        png = plotly.io.to_image(str, scale=3,width=1200, height=800, format='png', engine="kaleido" )
        png_base64 = base64.b64encode(png).decode('ascii')
        print(png_base64)
        return png_base64
    except Exception as e:
        print(e)
    return "ff"


@app.route('/encryption', methods=['GET', 'POST'])
def encryption():
    try:
        if request.method == 'POST':
            data = request.form.get('data')
            key = request.form.get('key')
            iv = request.form.get('iv')
            text = aesEncrypt(data, key, iv)
            return text
        else:
            return ''
    except Exception as e:
        return str(e)


@app.route('/decryption', methods=['GET', 'POST'])
def decryption():
    try:
        if request.method == 'POST':
            data = request.form.get('data')
            key = request.form.get('key')
            iv = request.form.get('iv')
            text = aesDecrypt(data, key, iv)
            return text
        else:
            return ''
    except Exception as e:
        return str(e)


@app.route('/excel', methods=['GET', 'POST'])
def excel_decryption():
    try:
        if request.method == 'POST':
            key = request.form.get('key')
            iv = request.form.get('iv')

            xlsx = request.files['input_file']
            curr_time = 'inpnjik'
            save_path = os.path.join('C:\\Users\\remshad\\PycharmProjects\\flaskPdfKit', curr_time)
            if not os.path.exists(save_path):
                os.makedirs(save_path)

            file_path = os.path.join(save_path, f'manager_list_{curr_time}.xlsx')
            xlsx.save(file_path)
            data = list()
            [excel_status, excel_data] = get_data_from_excel(file_path, 'Sheet1')
            if excel_status:
                for i_data in excel_data:
                    print(i_data)
                    # print(key,iv)
                    # print('data: ',(i_data['middle_name']))
                    # print(aesDecrypt(str(i_data['middle_name']), key, iv))
                    item = {
                        'first_name': aesDecrypt(str(i_data['first_name']), key, iv) if (
                                i_data['first_name'] and i_data['first_name'] != 'NULL' and len(
                            i_data['first_name']) > 5) else '',
                        # 'middle_name': aesDecrypt(str(i_data['middle_name']), key, iv) if (
                        #         i_data['middle_name'] and i_data['middle_name'] != 'NULL' and len(
                        #     i_data['middle_name']) > 5) else '',
                        'last_name': aesDecrypt(str(i_data['last_name']), key, iv) if (
                                i_data['last_name'] and i_data['last_name'] != 'NULL' and len(
                            i_data['last_name']) > 5) else '',
                        # 'display_name': aesDecrypt(str(i_data['display_name']), key, iv) if (
                        #         i_data['display_name'] and i_data['display_name'] != 'NULL' and len(
                        #     i_data['display_name']) > 5) else '',
                        'employee_id': aesDecrypt(str(i_data['employee_id']), key, iv) if (
                                i_data['employee_id'] and i_data['employee_id'] != 'NULL' and len(
                            i_data['employee_id']) > 5) else '',
                        # 'id': i_data['id'],
                        # 'username': aesDecrypt(str(i_data['username']), key, iv) if (
                        #         i_data['username'] and i_data['username'] != 'NULL' and len(
                        #     i_data['username']) > 5) else '',
                        # 'email': aesDecrypt(str(i_data['email']), key, iv) if (
                        #         i_data['email'] and i_data['email'] != 'NULL' and len(
                        #     i_data['email']) > 5) else ''

                    }
                    print(item)
                    data.append(item)
                    # print(data)

            return create_quick_exel(data, 'result')
        else:
            return ''
    except Exception as e:
        return str(e)


@app.route('/getverify', methods=['GET', 'POST'])
def excel_getverify():
    try:
        if request.method == 'POST':
            # key = request.form.get('key')
            # iv = request.form.get('iv')

            xlsx = request.files['input_file']
            curr_time = 'inpnjik'
            save_path = os.path.join('C:\\Users\\remshad\\PycharmProjects\\flaskPdfKit', curr_time)
            if not os.path.exists(save_path):
                os.makedirs(save_path)

            file_path = os.path.join(save_path, f'manager_list_{curr_time}.xlsx')
            xlsx.save(file_path)
            data = list()
            missing_data = list()
            [missing_excel_status, missing_excel_data] = get_data_from_excel(
                'C:\\Users\\remshad\\PycharmProjects\\flaskPdfKit\\missing_data.xlsx', 'Sheet1')
            if missing_excel_status:
                for i_data in missing_excel_data:
                    missing_data.append(i_data['employee_id'])

            [excel_status, excel_data] = get_data_from_excel(file_path, 'Sheet1')
            if excel_status:
                for i_data in excel_data:
                    # print(i_data)
                    # print(key,iv)
                    # print('data: ',(i_data['middle_name']))
                    # print(aesDecrypt(str(i_data['middle_name']), key, iv))
                    item = {
                        'Employee ID': i_data['Employee ID'],
                        'Reporting Manager': i_data['Reporting Manager']
                    }
                    data.append(item)
                    print(item)
            list_no_data = list()
            list_unmapped = list()
            for emp_id in missing_data:
                flag = 0
                for item in data:
                    if (item['Employee ID'] == emp_id):
                        flag = 1
                        for items in data:
                            if items['Employee ID'] == item['Reporting Manager']:
                                flag = 2
                                list_unmapped.append({
                                    'Employee ID': item['Employee ID'],
                                    'Reporting Manager': item['Reporting Manager'],
                                    'flag': 'reupload'
                                })
                        if flag == 1:
                            list_unmapped.append({
                                'Employee ID': item['Employee ID'],
                                'Reporting Manager': item['Reporting Manager'],
                                'flag': 'missing_manager'
                            })
                if flag == 0:
                    list_no_data.append({'Employee ID': emp_id, 'flag': 'Not in current excel'})
            create_quick_exel(list_unmapped, 'list_unmapped')
            create_quick_exel(list_no_data, 'list_no_data')
            return 'done'
        else:
            return ''
    except Exception as e:
        return str(e)


if __name__ == '__main__':
    app.run(port=5001)
