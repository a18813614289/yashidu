from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import os
from werkzeug.utils import secure_filename
import tempfile
import shutil

# 导入Excel到Word转换的核心功能
from excel_to_word import run_excel_to_word_automation

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# 配置临时文件目录
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 确保上传目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xlsx', 'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    # 检查是否有文件上传
    if 'excel_file' not in request.files or 'word_file' not in request.files:
        flash('请上传Excel和Word文件')
        return redirect(url_for('index'))
    
    excel_file = request.files['excel_file']
    word_file = request.files['word_file']
    
    # 检查文件是否有名称
    if excel_file.filename == '' or word_file.filename == '':
        flash('请选择文件')
        return redirect(url_for('index'))
    
    # 检查文件类型
    if not (allowed_file(excel_file.filename) and allowed_file(word_file.filename)):
        flash('只允许上传.xlsx和.docx文件')
        return redirect(url_for('index'))
    
    # 获取复制次数
    try:
        copy_count = int(request.form.get('copy_count', 20))
        if copy_count < 1:
            copy_count = 1
    except ValueError:
        copy_count = 20
    
    try:
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 保存上传的文件
            excel_filename = secure_filename(excel_file.filename)
            word_filename = secure_filename(word_file.filename)
            
            excel_path = os.path.join(temp_dir, excel_filename)
            word_path = os.path.join(temp_dir, word_filename)
            
            excel_file.save(excel_path)
            word_file.save(word_path)
            
            # 创建输出文件路径
            output_filename = f'output_{os.path.splitext(word_filename)[0]}.docx'
            output_path = os.path.join(temp_dir, output_filename)
            
            # 执行转换
            def status_callback(message):
                print(message)
                # Flask不支持直接从线程中更新UI，这里简单打印日志
            
            run_excel_to_word_automation(excel_path, word_path, copy_count, output_path, status_callback)
            
            # 复制到上传目录以便下载
            download_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            shutil.copy2(output_path, download_path)
            
            # 提供下载
            return send_file(download_path, as_attachment=True, attachment_filename=output_filename)
    except Exception as e:
        flash(f'处理文件时出错: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # 生产环境请设置debug=False
    app.run(debug=True)