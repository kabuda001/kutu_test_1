import win32com.client


def get_shape_size_in_units(shape):
    """ 获取形状的宽度和高度，并根据单位转换为合适的单位 """
    # 获取当前文档的单位
    unit = corel.ActiveDocument.Unit  # 单位可以是 mm, cm, inches, pixels 等
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

def change_length(original_width,original_height,new_height):

    # 计算缩放比例（目标高度 / 原始高度）
    scale_factor = new_height / original_height

    # 计算新的宽度
    new_width = original_width * scale_factor


    # 获取解散后的所有形状（shapes）
    shapes = doc.ActivePage.Shapes
    for shape in shapes:
        shape.SetSize(shape.SizeWidth * scale_factor, shape.SizeHeight * scale_factor)

    # 输出调整后的宽度和高度（单位：mm）
    print(f"原始宽度：{original_width} mm, 原始高度：{original_height} mm")
    print(f"缩放后的宽度：{new_width} mm, 缩放后的高度：{new_height} mm")

try:
    # 连接 CorelDRAW 应用
    corel = win32com.client.Dispatch("CorelDRAW.Application")
    corel.Visible = True  # 不显示CorelDRAW界面，可以设置为True查看
    doc = corel.OpenDocument("C:\\Users\\zhaoc\\Desktop\\测试\\ZC085-B.cdr")

    # 打开 CDR 文件
    # doc = corel.OpenDocument("C:\\path\\to\\your\\file.cdr")

    # 获取页面上的所有对象
    page = doc.Pages(1)  # 获取第一页
    shapes = page.Shapes

    # 清空所有选择（通过取消选中所有图形）
    for shape in shapes:
        shape.Selected = False



    # 选择所有对象
    for shape in shapes:
        shape.Selected = True

    # 获取当前选中的对象
    selection = corel.ActiveSelection

    # 创建一个组
    group = selection.Group()
    # 获取组合的组的宽度和高度
    group_width, group_height = get_shape_size_in_units(group)

    print(f"组合的组宽度: {group_width}, 高度: {group_height}")
    change_length(group_width, group_height, 200)
    # 保存为新的路径
    doc.Save()
    doc.Close()
except Exception as e:
    # print(f"发生错误: {e}")
    print("发生错误:", e)
finally:
    # if 'corel' in locals():
    #     corel.Quit()
    pass
