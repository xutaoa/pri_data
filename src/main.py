from flask import Flask, render_template, request, jsonify, send_from_directory
import os
import pandas as pd
from werkzeug.utils import secure_filename
import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB限制
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# 确保上传目录存在
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # 读取Excel文件的所有sheet名[1][5]
        try:
            excel_file = pd.ExcelFile(filepath)
            sheets = excel_file.sheet_names
            return jsonify({
                'status': 'success',
                'filename': filename,
                'sheets': sheets
            })
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/analyze', methods=['POST'])
def analyze_data():
    data = request.get_json()
    filename = data.get('filename')
    sheet_name = data.get('sheet_name')
    file_num = data.get('file_num')

    text_result = ""

    # 获取当前日期
    current_date = datetime.date.today()

    # 计算明天日期
    tomorrow = current_date + datetime.timedelta(days=1)
    weekday = tomorrow.weekday()
    weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    w = weekdays[weekday]
    # 格式化输出明天日期
    tomorrow_date = tomorrow.strftime("%m月%d日") + "   " + w

    print(tomorrow_date)

    if not filename or not sheet_name:
        return jsonify({'error': 'Missing parameters'}), 400
    
    try:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        # df = pd.read_excel(filepath, sheet_name=sheet_name)
        
        if file_num == 1:
            text_result = fun_1(filepath, sheet_name, tomorrow_date)
        elif file_num == 2:
            text_result = fun_2(filepath, sheet_name, tomorrow_date)
        else:
            text_result = fun_3(filepath, sheet_name, tomorrow_date)
        
        return jsonify({
            'status': 'success',
            'result': text_result
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def get_text(df, start, end, row):
    subset = df.iloc[start:end, row]
    text = "\n"
    for index, value in subset.items():
        if str(value) != 'nan':
            text += str(value).replace('借', '') + "\n"
    
    if len(text) < 10:
        return ""
    return text 

def fun_1(filepath, sheet_name, tomorrow_date):
    text = ""
    text += tomorrow_date + '\n'
    text += '【德众堂】宝山区新沪路1073号\n'
    df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=[0])
    
    text += '上午'
    #上午数据
    for i in range(1, 4):  
        text += get_text(df, 0, 17, i)

    for i in range(5, 8):  
        text += get_text(df, 0, 17, i)

    for i in range(9, 12):  
        text += get_text(df, 0, 17, i)
    
    for i in range(1, 4):  
        text += get_text(df, 17, 34, i)

    text += '\n下午\n'
    #下午数据
    for i in range(5, 8):  
        text += get_text(df, 17, 34, i)

    for i in range(9, 12):  
        text += get_text(df, 17, 34, i)

    for i in range(1, 4):  
        text += get_text(df, 34, 52, i)        

    return text

def fun_2(filepath, sheet_name, tomorrow_date):
    text = ""
    text += tomorrow_date + '\n'
    text += '【轩济堂】顾村镇菊泉街675号4幢2楼\n' 

    df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=[0])
    
    text += '上午'
    #上午数据
    for i in range(0, 3):  
        text += get_text(df, 0, 7, i)

    for i in range(3, 6):  
        text += get_text(df, 0, 7, i)

    for i in range(0, 3):  
        text += get_text(df, 7, 14, i)
    
    for i in range(3, 6):  
        text += get_text(df, 7, 14, i)

    text += '\n下午\n'
    #下午数据
    for i in range(0, 3):  
        text += get_text(df, 14, 21, i)

    for i in range(3, 6):  
        text += get_text(df, 14, 21, i)

    for i in range(0, 3):  
        text += get_text(df, 21, 28, i)        

    return text

def fun_3(filepath, sheet_name, tomorrow_date):
    text = ""
    text += tomorrow_date + '\n'
    text += '【宁合中医】灵石路健康智谷7号楼2楼\n' 

    df = pd.read_excel(filepath, sheet_name=sheet_name)

    text += '上午'
    #上午数据
    for i in range(1, 4):  
        text += get_text(df, 0, 12, i)

    for i in range(5, 8):  
        text += get_text(df, 0, 12, i)

    for i in range(9, 12):  
        text += get_text(df, 0, 12, i)
    
    for i in range(1, 4):  
        text += get_text(df, 12, 24, i)

    text += '\n下午\n'
    #下午数据
    for i in range(5, 8):  
        text += get_text(df, 12, 24, i)

    for i in range(9, 12):  
        text += get_text(df, 12, 24, i)

    for i in range(1, 4):  
        text += get_text(df, 24, 48, i)        

    return text

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
