import win32com.client as win32
import os

# 设置 CorelDRAW 应用
coreldraw = win32.Dispatch("CorelDRAW.Application")
coreldraw.Visible = True  # 确保 CorelDRAW 可见

# 指定要打开的 CDR 文件和要插入的 PNG 图像文件路径
cdr_file = r"C:\Users\zhaoc\Desktop\测试条形码\Backup_of_1-款号的文字.cdr"  # 请替换为实际路径
png_file = r"D:\python_workspace\kutu_test_1\241213-206202498590293_barcode.png"  # 请替换为实际路径
test_file = r"C:\Users\zhaoc\Desktop\测试条形码\241213-206202498590293.cdr"
# 打开 CDR 文件
document = coreldraw.OpenDocument(cdr_file)

# 获取当前页面
page = document.ActivePage

# 导入 PNG 图像
try:
    # 使用 Import 方法导入 PNG 图像
    picture = document.Import(test_file)
    print("PNG 图像已成功导入！")
except Exception as e:
    print(f"导入图片时发生错误: {e}")

# 保存并关闭文件
document.Save()
document.Close()
