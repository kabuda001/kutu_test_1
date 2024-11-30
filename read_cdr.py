import win32com.client

# 连接 CorelDRAW 应用
coreldraw = win32com.client.Dispatch("CorelDRAW.Application")
coreldraw.OpenDocument("\\\\xinguo\\贴纸.生产资料\\美工做的新款，待检查\\11月\\11.25\\DK182-F.cdr")


def get_shape_size_in_units(shape):
    """ 获取形状的宽度和高度，并根据单位转换为合适的单位 """
    # 获取当前文档的单位
    unit = coreldraw.ActiveDocument.Unit  # 单位可以是 mm, cm, inches, pixels 等
    width = shape.SizeWidth
    height = shape.SizeHeight

    # 英寸（inches）转换为厘米（cm）
    if unit == 1:  # 1: inches
        width *= 25.4  # 将英寸转换为厘米
        height *= 25.4
    elif unit == 2:  # 2: mm
        # width /= 10  # 将毫米转换为厘米
        # height /= 10
        pass
    elif unit == 3:  # 3: cm
        width *= 10  # 厘米转毫米
        height *= 10
    elif unit == 7:  # 7: pixels
        # 无需转换，单位已经是像素
        pass
    else:
        print(f"不支持的单位: {unit}")

    return width, height

# 获取页面长宽
doc = coreldraw.ActiveDocument
width = doc.ActivePage.SizeWidth
height = doc.ActivePage.SizeHeight
unit = coreldraw.ActiveDocument.Unit
# print(unit)
# 获取所有对象并打印每个对象的宽度和高度
for shape in doc.ActivePage.Shapes:
    # print(type(shape))
    if shape.Type == 1:  # 检查是否是图形对象 (1 是形状类型)
        width,height = get_shape_size_in_units(shape)
        # width = shape.SizeWidth  # 获取宽度
        # height = shape.SizeHeight  # 获取高度
        print(f"形状 ID: {shape.ID}, 宽度: {width}, 高度: {height}")

    elif shape.Type == 5:  # 如果是图片类型
        # width = shape.SizeWidth
        # height = shape.SizeHeight
        width, height = get_shape_size_in_units(shape)
        print(f"形状NAME: {shape.Name},宽度: {width}, 高度: {height}")
# print(f"CDR 文件宽度: {width}, 高度: {height}")

# 关闭文档
doc.Close()