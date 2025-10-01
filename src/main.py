from flask import Flask, render_template, request, jsonify, send_from_directory
import os
import pandas as pd
from werkzeug.utils import secure_filename
import datetime
import traceback
import json

# 初始化Flask应用
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB限制
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# 配置文件路径
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'dezhongtang_config.json')

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def get_default_dezhongtang_config():
    """获取德众堂默认配置"""
    return {
        'startColumn': 1,
        'columnStep': 4,
        'startRow': 0,
        'rowStep': 14
    }

def get_dezhongtang_config():
    """获取德众堂配置，如果配置文件不存在则返回默认配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"读取配置文件出错: {e}")
            return get_default_dezhongtang_config()
    else:
        return get_default_dezhongtang_config()

def save_dezhongtang_config(config):
    """保存德众堂配置到文件"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"保存配置文件出错: {e}")
        return False

def allowed_file(filename):
    """检查文件类型是否允许"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # 读取Excel文件的所有sheet名
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

@app.route('/get-dezhongtang-config', methods=['GET'])
def get_dezhongtang_config_api():
    """获取德众堂配置的API"""
    try:
        config = get_dezhongtang_config()
        return jsonify({
            'status': 'success',
            'config': config
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/save-dezhongtang-config', methods=['POST'])
def save_dezhongtang_config_api():
    """保存德众堂配置的API"""
    try:
        config = request.get_json()
        if save_dezhongtang_config(config):
            return jsonify({'status': 'success', 'message': '配置保存成功'})
        else:
            return jsonify({'error': '配置保存失败'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/analyze', methods=['POST'])
def analyze_data():
    """分析上传的Excel文件"""
    data = request.get_json()
    filename = data.get('filename')
    sheet_name = data.get('sheet_name')
    file_num = data.get('file_num')
    # 获取德众堂配置参数（如果有的话）
    dezhongtang_config = data.get('dezhongtang_config')

    if not filename or not sheet_name:
        return jsonify({'error': 'Missing parameters'}), 400
    
    try:
        # 计算明天日期和星期
        tomorrow = datetime.date.today() + datetime.timedelta(days=1)
        weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
        weekday_str = weekdays[tomorrow.weekday()]
        tomorrow_date = tomorrow.strftime("%m月%d日") + "   " + weekday_str
        print(tomorrow_date)

        # 根据文件编号调用相应的处理函数
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        processors = {
            1: process_dezhongtang,
            2: process_xuanjitang,
            3: process_ninghezhongyi
        }
        
        processor = processors.get(file_num)
        if processor:
            # 如果是德众堂且有配置参数，则传递配置参数
            if file_num == 1 and dezhongtang_config:
                text_result = processor(filepath, sheet_name, tomorrow_date, weekday_str, dezhongtang_config)
            else:
                text_result = processor(filepath, sheet_name, tomorrow_date, weekday_str)
        else:
            return jsonify({'error': 'Invalid file number'}), 400
        
        return jsonify({
            'status': 'success',
            'result': text_result
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def get_text(df, start, end, row):
    """从DataFrame中提取文本数据"""
    text = "\n"
    try:
        subset = df.iloc[start:end, row]
        for index, value in subset.items():
            if str(value) != 'nan':
                text += str(value).replace('借', '') + "\n"
        
        return text if len(text) >= 10 else ""
    except IndexError as e:
        print(e)
        return text

def process_dezhongtang(filepath, sheet_name, tomorrow_date, weekday_str, config=None):
    """处理德众堂数据"""
    # 如果没有提供配置，则使用默认配置
    if config is None:
        config = get_dezhongtang_config()
    
    try:
        text = f"{tomorrow_date}\n【德众堂】宝山区新沪路1073号\n"
        df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=[0])

        start_colum = config['startColumn']
        colum_step = config['columnStep']

        start = config['startRow']
        step = config['rowStep']

        text += '上午'
        # 上午数据
        for i in range(start_colum, start_colum+colum_step):  
            text += get_text(df, start, start+step, i)
        for i in range(start_colum+colum_step, start_colum+colum_step*2):  
            text += get_text(df, start, start+step, i)
        for i in range(start_colum+colum_step*2, start_colum+colum_step*3):  
            text += get_text(df, start, start+step, i)
        for i in range(start_colum, start_colum+colum_step):  
            text += get_text(df, start+step, start+step*2, i)

        text += '\n下午'
        # 下午数据
        for i in range(start_colum, start_colum+colum_step):  
            text += get_text(df, start+step, start+step*2, i)
        for i in range(start_colum+colum_step, start_colum+colum_step*2):  
            text += get_text(df, start+step, start+step*2, i)
        for i in range(start_colum+colum_step*2, start_colum+colum_step*3):  
            text += get_text(df, start+step, start+step*2, i)

        # 星期二夜班数据
        if weekday_str == "星期二":
            text += '\n晚'
            for i in range(start_colum, start_colum+colum_step):  
                text += get_text(df, start+step*2, start+step*3, i)
            for i in range(start_colum+colum_step, start_colum+colum_step*2):  
                text += get_text(df, start+step*2, start+step*3, i)

        return text
    except Exception as e:
        print(f"德众堂处理错误: {e}")
        traceback.print_exc()
        return ""

def process_xuanjitang(filepath, sheet_name, tomorrow_date, weekday_str):
    """处理轩济堂数据"""
    try:
        text = f"{tomorrow_date}\n【轩济堂】顾村镇菊泉街675号4幢2楼\n"
        df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=[0])
        
        text += '上午'
        # 上午数据
        for i in range(0, 3):  
            text += get_text(df, 0, 7, i)
        for i in range(3, 6):  
            text += get_text(df, 0, 7, i)
        for i in range(0, 3):  
            text += get_text(df, 7, 14, i)
        for i in range(3, 6):  
            text += get_text(df, 7, 14, i)

        text += '\n下午'
        # 下午数据
        for i in range(0, 3):  
            text += get_text(df, 14, 21, i)
        for i in range(3, 6):  
            text += get_text(df, 14, 21, i)
        for i in range(0, 3):  
            text += get_text(df, 21, 28, i)

        return text
    except Exception as e:
        print(f"轩济堂处理错误: {e}")
        traceback.print_exc()
        return ""

def process_ninghezhongyi(filepath, sheet_name, tomorrow_date, weekday_str):
    """处理宁合中医数据"""
    try:
        text = f"{tomorrow_date}\n【宁合中医】灵石路健康智谷7号楼2楼\n"
        df = pd.read_excel(filepath, sheet_name=sheet_name)

        text += '上午'
        # 上午数据
        for i in range(1, 4):  
            text += get_text(df, 0, 14, i)
        for i in range(5, 8):  
            text += get_text(df, 0, 14, i)
        for i in range(9, 12):  
            text += get_text(df, 0, 14, i)
        for i in range(1, 4):  
            text += get_text(df, 14, 28, i)

        text += '\n下午'
        # 下午数据
        for i in range(5, 8):  
            text += get_text(df, 14, 28, i)
        for i in range(9, 12):  
            text += get_text(df, 14, 28, i)
        for i in range(1, 4):  
            text += get_text(df, 28, 42, i)

        return text
    except Exception as e:
        print(f"宁合中医处理错误: {e}")
        traceback.print_exc()
        return ""

@app.route('/')
def index():
    """主页路由"""
    return render_template('index.html')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    """文件下载路由"""
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
