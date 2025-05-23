# import sys
# import os
# sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))  # DON'T CHANGE THIS !!!

from flask import Flask, render_template, request, jsonify
# import pandas as pd
import tempfile
import os
# import re

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件大小为16MB

# 存储处理结果
processed_results = {
    '1': None,
    '2': None,
    '3': None
}

# 文件1处理方法：提取所有数值并计算总和
def process_file1(file_path):
    return "1"

# 文件2处理方法：提取文本并统计词频
def process_file2(file_path):
    return "1"

# 文件3处理方法：提取日期时间信息并分析
def process_file3(file_path):
    return "1"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有文件部分'}), 400
    
    file = request.files['file']
    file_num = request.form.get('fileNum', '1')
    
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '只支持Excel文件(.xlsx, .xls)'}), 400
    
    try:
        # 创建临时文件
        temp_dir = tempfile.mkdtemp()
        temp_path = os.path.join(temp_dir, file.filename)
        file.save(temp_path)
        
        # 根据文件编号选择处理方法
        if file_num == '1':
            result = process_file1(temp_path)
        elif file_num == '2':
            result = process_file2(temp_path)
        elif file_num == '3':
            result = process_file3(temp_path)
        else:
            result = "未知的文件编号"
        
        # 存储处理结果
        processed_results[file_num] = result
        
        # 清理临时文件
        os.remove(temp_path)
        os.rmdir(temp_dir)
        
        return jsonify({'result': result})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
