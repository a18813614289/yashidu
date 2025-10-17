from flask import Flask, request, render_template_string, send_file, jsonify
import os
import sys
import tempfile
from werkzeug.utils import secure_filename
import traceback

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入原始脚本的功能
import _9 as压实度_module

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB限制

# 临时文件目录
TEMP_DIR = tempfile.gettempdir()

@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>压实度计算工具</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #f5f5f5;
                max-width: 800px;
                margin: 0 auto;
            }
            h1 {
                color: #333;
                text-align: center;
            }
            .container {
                background-color: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            .form-group {
                margin-bottom: 15px;
            }
            label {
                display: block;
                margin-bottom: 5px;
                font-weight: bold;
            }
            input[type="file"] {
                width: 100%;
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            input[type="number"] {
                width: 100px;
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            button {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 15px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 16px;
            }
            button:hover {
                background-color: #45a049;
            }
            .progress {
                width: 100%;
                height: 20px;
                background-color: #f1f1f1;
                border-radius: 10px;
                margin-top: 20px;
                display: none;
            }
            .progress-bar {
                height: 100%;
                background-color: #4CAF50;
                border-radius: 10px;
                text-align: center;
                line-height: 20px;
                color: white;
                width: 0%;
            }
            .status {
                margin-top: 15px;
                color: #666;
                min-height: 20px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>压实度计算工具</h1>
            <form id="processForm">
                <div class="form-group">
                    <label for="excelFile">Excel文件上传：</label>
                    <input type="file" id="excelFile" name="excelFile" accept=".xlsx" required>
                </div>
                <div class="form-group">
                    <label for="wordFile">Word模板文件上传：</label>
                    <input type="file" id="wordFile" name="wordFile" accept=".docx" required>
                </div>
                <div class="form-group">
                    <label for="copyCount">复制次数：</label>
                    <input type="number" id="copyCount" name="copyCount" min="1" value="1" required>
                </div>
                <button type="submit">处理文件</button>
            </form>
            
            <div class="progress" id="progressBar">
                <div class="progress-bar" id="progressBarInner">0%</div>
            </div>
            
            <div class="status" id="statusMessage"></div>
            
            <div id="downloadLink" style="margin-top: 20px; display: none;">
                <a href="" id="downloadBtn" download>下载生成的Word文件</a>
            </div>
        </div>
        
        <script>
            const form = document.getElementById('processForm');
            const progressBar = document.getElementById('progressBar');
            const progressBarInner = document.getElementById('progressBarInner');
            const statusMessage = document.getElementById('statusMessage');
            const downloadLink = document.getElementById('downloadLink');
            const downloadBtn = document.getElementById('downloadBtn');
            
            form.addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const formData = new FormData(form);
                
                // 显示进度条和状态
                progressBar.style.display = 'block';
                statusMessage.textContent = '开始处理文件...';
                downloadLink.style.display = 'none';
                
                try {
                    // 首先发送文件处理请求
                    const response = await fetch('/process', {
                        method: 'POST',
                        body: formData
                    });
                    
                    if (!response.ok) {
                        throw new Error('处理失败');
                    }
                    
                    const data = await response.json();
                    
                    if (data.success) {
                        statusMessage.textContent = '处理完成！';
                        // 设置下载链接
                        downloadBtn.href = `/download?filename=${encodeURIComponent(data.filename)}`;
                        downloadLink.style.display = 'block';
                    } else {
                        statusMessage.textContent = '处理失败: ' + data.error;
                    }
                } catch (error) {
                    statusMessage.textContent = '发生错误: ' + error.message;
                } finally {
                    progressBar.style.display = 'none';
                }
            });
        </script>
    </body>
    </html>
    ''')

@app.route('/process', methods=['POST'])
def process_files():
    try:
        # 获取上传的文件
        excel_file = request.files['excelFile']
        word_file = request.files['wordFile']
        copy_count = int(request.form['copyCount'])
        
        # 保存临时文件
        excel_path = os.path.join(TEMP_DIR, secure_filename(excel_file.filename))
        word_path = os.path.join(TEMP_DIR, secure_filename(word_file.filename))
        excel_file.save(excel_path)
        word_file.save(word_path)
        
        # 生成输出文件名
        output_filename = f"output_{os.path.splitext(secure_filename(word_file.filename))[0]}.docx"
        output_path = os.path.join(TEMP_DIR, output_filename)
        
        # 调用原始脚本的功能
        # 这里需要根据原始脚本的实际函数名和参数进行调整
        # 假设原始脚本有一个名为run_excel_to_word_automation的主函数
        def status_callback(message):
            print(f"处理状态: {message}")
        
        # 尝试调用原始脚本的主要功能
       压实度_module.run_excel_to_word_automation(excel_path, word_path, copy_count, output_path, status_callback)
        
        # 返回成功信息
        return jsonify({
            'success': True,
            'filename': output_filename
        })
    
    except Exception as e:
        error_trace = traceback.format_exc()
        print(f"错误详情: {error_trace}")
        return jsonify({
            'success': False,
            'error': str(e)
        })
    
    finally:
        # 清理临时文件（可选，取决于需求）
        pass

@app.route('/download')
def download_file():
    filename = request.args.get('filename')
    filepath = os.path.join(TEMP_DIR, filename)
    
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        return jsonify({'error': '文件不存在'}), 404

@app.route('/status')
def get_status():
    # 这里可以实现一个简单的状态查询接口
    return jsonify({'status': 'running'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)