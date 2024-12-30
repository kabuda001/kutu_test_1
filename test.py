import os
import barcode
from barcode.writer import ImageWriter

CUSTOM_OPTIONS = {
    "module_width": 0.2,       # 单个条纹的最小宽度, mm
    "module_height": 15.0,     # 条纹带的高度, mm
    "quiet_zone": 6.5,         # 图片两边与首尾两条纹之间的距离, mm
    "font_size": 10,           # 条纹底部文本的大小,pt
    "text_distance": 5.0,      # 条纹底部与条纹之间的距离, mm
}

# 生成条形码并保存
def generate_barcode(file_name):
    # 获取条形码格式
    barcode_format = barcode.get_barcode_class('code128')
    if barcode_format is None:
        print("条形码格式 'code128' 未找到")
        return None
    # barcode.generate()
    # 使用 ImageWriter 生成条形码图像
    barcode_image = barcode_format(file_name, writer=ImageWriter())
    barcode_path = f"{file_name}_barcode.png"

    try:
        barcode_image.save(barcode_path,options=CUSTOM_OPTIONS)
        print(f"条形码已保存为: {barcode_path}")
        return barcode_path
    except Exception as e:
        print(f"生成条形码时发生错误: {e}")
        return None


# 测试条形码生成
cdr_file = r"C:\Users\zhaoc\Desktop\测试条形码\241213-206202498590293.cdr"
file_name = os.path.splitext(os.path.basename(cdr_file))[0]  # 获取文件名并去掉 .cdr 后缀
barcode_path = generate_barcode(file_name)

if barcode_path:
    print(f"条形码路径：{barcode_path}")
else:
    print("条形码生成失败")