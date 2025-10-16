from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import os
import sys
import logging
from werkzeug.utils import secure_filename
import tempfile
import shutil
import uuid

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

# 导入Excel到Word转换的核心功能
from excel_to_word import run_excel_to_word_automation

# 创建Flask应用
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'supersecretkey')

# 配置临时文件目录 - 在Vercel中使用/tmp目录
UPLOAD_FOLDER = os.environ.get('TMPDIR', '/tmp')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
logger.info(f"上传目录: {UPLOAD_FOLDER}")

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xlsx', 'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        # 检查是否有文件上传
        if 'excel_file' not in request.files or 'word_file' not in request.files:
            logger.error('缺少文件上传')
            flash('请上传Excel和Word文件')
            return redirect(url_for('index'))
        
        excel_file = request.files['excel_file']
        word_file = request.files['word_file']
        
        # 检查文件是否有名称
        if excel_file.filename == '' or word_file.filename == '':
            logger.error('文件名称为空')
            flash('请选择文件')
            return redirect(url_for('index'))
        
        # 检查文件类型
        if not (allowed_file(excel_file.filename) and allowed_file(word_file.filename)):
            logger.error(f'文件类型不支持: {excel_file.filename}, {word_file.filename}')
            flash('只允许上传.xlsx和.docx文件')
            return redirect(url_for('index'))
        
        # 获取复制次数
        try:
            copy_count = int(request.form.get('copy_count', 20))
            if copy_count < 1:
                copy_count = 1
            logger.info(f'复制次数: {copy_count}')
        except ValueError:
            copy_count = 20
            logger.warning('复制次数参数无效，使用默认值20')
        
        # 创建唯一的临时目录名
        temp_dir_name = f"temp_{uuid.uuid4().hex}"
        temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], temp_dir_name)
        os.makedirs(temp_dir, exist_ok=True)
        logger.info(f'创建临时目录: {temp_dir}')
        
        try:
            # 保存上传的文件
            excel_filename = secure_filename(excel_file.filename)
            word_filename = secure_filename(word_file.filename)
            
            excel_path = os.path.join(temp_dir, excel_filename)
            word_path = os.path.join(temp_dir, word_filename)
            
            logger.info(f'保存Excel文件到: {excel_path}')
            excel_file.save(excel_path)
            
            logger.info(f'保存Word文件到: {word_path}')
            word_file.save(word_path)
            
            # 创建输出文件路径
            output_filename = f'output_{os.path.splitext(word_filename)[0]}.docx'
            output_path = os.path.join(temp_dir, output_filename)
            
            # 执行转换
            def status_callback(message):
                logger.info(f'转换状态: {message}')
            
            logger.info('开始执行Excel到Word转换')
            run_excel_to_word_automation(excel_path, word_path, copy_count, output_path, status_callback)
            logger.info(f'转换完成，输出文件: {output_path}')
            
            # 提供下载
            return send_file(output_path, as_attachment=True, attachment_filename=output_filename)
        except Exception as e:
            logger.error(f'处理文件时出错: {str(e)}', exc_info=True)
            flash(f'处理文件时出错: {str(e)}')
            return redirect(url_for('index'))
        finally:
            # 清理临时文件（在生产环境中，Vercel会自动清理/tmp目录）
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    logger.info(f'已清理临时目录: {temp_dir}')
            except Exception as cleanup_error:
                logger.error(f'清理临时文件时出错: {cleanup_error}')
    except Exception as e:
        logger.error(f'请求处理异常: {str(e)}', exc_info=True)
        flash('服务器内部错误')
        return redirect(url_for('index'))

# 用于Vercel的gunicorn入口点
# Vercel会查找名为'application'的变量作为WSGI应用
def application(environ, start_response):
    # 更新Flask应用的环境变量
    for key, value in environ.items():
        os.environ[key] = value
    # 返回Flask的WSGI应用
    return app.wsgi_app(environ, start_response)

# 开发环境入口
if __name__ == '__main__':
    # 使用环境变量中的端口，默认为5000
    port = int(os.environ.get("PORT", 5000))
    # 生产环境设置debug=False，开发环境可以设为True
    debug_mode = os.environ.get("FLASK_DEBUG", "False").lower() == "true"
    logger.info(f'启动Flask应用，端口: {port}, 调试模式: {debug_mode}')
    app.run(host="0.0.0.0", port=port, debug=debug_mode)