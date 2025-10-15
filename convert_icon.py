from PIL import Image
import sys
import os

def convert_svg_to_ico(svg_path, ico_path):
    try:
        # 首先需要手动将SVG转换为PNG（使用Inkscape或其他工具）
        if not os.path.exists(svg_path.replace('.svg', '.png')):
            print(f"转换失败: 请先手动将 {svg_path} 转换为PNG格式")
            print("请先使用在线工具将SVG转换为PNG，然后重命名为{svg_path.replace('.svg', '.png')}")
            return False
            
        # 读取PNG文件
        img = Image.open(svg_path.replace('.svg', '.png'))
        
        # 转换为ICO格式
        img.save(ico_path, format='ICO', sizes=[(32,32), (48,48), (64,64)])
        print(f"图标已成功转换为 {ico_path}")
        return True
    except Exception as e:
        print(f"转换失败: {str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python convert_icon.py <输入svg路径> <输出ico路径>")
        sys.exit(1)
        
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    
    if not input_path.lower().endswith('.svg'):
        print("错误: 输入文件必须是SVG格式")
        sys.exit(1)
        
    if not output_path.lower().endswith('.ico'):
        print("错误: 输出文件必须是ICO格式")
        sys.exit(1)
        
    if convert_svg_to_ico(input_path, output_path):
        sys.exit(0)
    else:
        sys.exit(1)