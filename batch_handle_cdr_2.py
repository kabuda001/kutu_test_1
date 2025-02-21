import sys
import os
import time
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLineEdit, QLabel, QComboBox, \
    QHBoxLayout, QProgressBar, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import win32com.client
import openpyxl


class LoadingThread(QThread):
    """后台线程模拟处理任务"""
    progress = pyqtSignal(int)  # 进度条信号
    finished = pyqtSignal()  # 完成信号

    def __init__(self, folder_path, static_file, parent=None):
        super().__init__(parent)
        self.folder_path = folder_path
        self.static_file = static_file

    def run(self):
        # 模拟耗时操作
        # for i in range(101):
        #     time.sleep(0.05)  # 模拟操作的时间
        #     self.progress.emit(i)  # 更新进度条
        # cdr_files = self.get_cdr_files(self.folder_path)
        orderMap = self.get_multiple_size()
        self.progress.emit(30)
        # 遍历cdr文件，并更改大小，如果尺寸符合，就不调整
        # for cdr_file in cdr_files:
        #     self.handle_cdr(cdr_file)
        self.magnify_mulit_cdr(orderMap)
        self.progress.emit(90)
        # self.delete_backup_cdr_files(self.folder_path)
        self.finished.emit()  # 完成任务

    def magnify_mulit_cdr(self,orderMap):
        for root, dirs, files in os.walk(self.folder_path):
            for file in files:
                if self.is_cdr_file(file):
                    file_name = os.path.splitext(os.path.basename(file))[0] # 获取文件名（不含路径）
                    parts = file_name.split("_")
                    self.handle_cdr(os.path.join(root, file), int(orderMap.get(parts[0])))
    def get_multiple_size(self):
        excel_data = self.read_excel(order_file=self.static_file)
        orderMap = {}
        for row_data in excel_data.values():
            order_num = row_data.get('订单编号')
            longest_side = row_data.get('最长边')
            orderMap[order_num] = longest_side
        return orderMap

    def read_excel(self,order_file):
        workbook = openpyxl.load_workbook(order_file)
        # 选择活动工作表
        sheet = workbook.active
        # 读取第一行作为键（keys）
        header = [cell.value for cell in sheet[1]]  # 第一行作为键
        # 读取数据（从第二行开始）
        data = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_data = {header[i]: row[i] for i in range(len(header))}
            data[row_data[header[0]]] = row_data  # 使用第一列的值作为外层字典的 key
        return data
    def handle_cdr(self, cdr_file,selected_size):
        try:
            # 连接 CorelDRAW 应用
            corel = win32com.client.Dispatch("CorelDRAW.Application")
            corel.Visible = False  # 不显示CorelDRAW界面，可以设置为True查看
            doc = corel.OpenDocument(cdr_file)
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
            group_width, group_height = self.get_shape_size_in_units(corel,group)
            print(f"组合的组宽度: {group_width}, 高度: {group_height}")
            width_or_height = self.get_value_based_on_threshold(group_width,group_height)
            if width_or_height:
                self.change_width(doc,group_width, group_height, selected_size*10)
            else:
                self.change_length(doc,group_width, group_height, selected_size*10)
            # 保存为新的路径
            doc.Save()
            doc.Close()
        except Exception as e:
            # print(f"发生错误: {e}")
            print("发生错误:", e)

    # true 表示变更宽，false 表示变更长
    def get_value_based_on_threshold(self,group_width, group_height, threshold_ratio=0.05):
        # 计算两个值的差值
        difference = abs(group_width - group_height)

        # 计算差值与最大值的比例
        max_value = max(group_width, group_height)
        ratio = difference / max_value

        # 根据比例判断返回较小值还是较大值
        if ratio < threshold_ratio:
            # 比例小于阈值，返回较小值，判断是长还是宽
            if group_width < group_height:
                return True
            else:
                return False
        else:
            # 比例大于或等于阈值，返回较大值，判断是长还是宽
            if group_width > group_height:
                return True
            else:
                return False

    def change_length(self,doc,original_width, original_height, new_height):

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

    def change_width(self,doc,original_width, original_height, new_width):

        # 计算缩放比例（目标高度 / 原始高度）
        scale_factor = new_width / original_width

        # 计算新的宽度
        new_height = original_height * scale_factor

        # 获取解散后的所有形状（shapes）
        shapes = doc.ActivePage.Shapes
        for shape in shapes:
            shape.SetSize(shape.SizeWidth * scale_factor, shape.SizeHeight * scale_factor)

        # 输出调整后的宽度和高度（单位：mm）
        print(f"原始宽度：{original_width} mm, 原始高度：{original_height} mm")
        print(f"缩放后的宽度：{new_width} mm, 缩放后的高度：{new_height} mm")

    def get_shape_size_in_units(self,corel,shape):
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

    def delete_backup_cdr_files(folder_path):
        # 遍历文件夹中的所有文件
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                # 检查文件名是否以 "Backup_of_" 开头且以 ".cdr" 结尾
                if file.startswith("Backup_of_") and file.endswith(".cdr"):
                    file_path = os.path.join(root, file)
                    try:
                        # 删除文件
                        os.remove(file_path)
                        print(f"Deleted: {file_path}")
                    except Exception as e:
                        print(f"Failed to delete {file_path}: {e}")

    def get_cdr_files(self,directory):
        cdr_files = []  # 用来存储符合条件的文件路径
        for root, dirs, files in os.walk(directory):  # 遍历目录及其子目录
            for file in files:
                if file.endswith('.cdr'):  # 判断文件后缀是否为 .cdr
                    cdr_files.append(os.path.join(root, file))  # 添加完整路径到列表中
        return cdr_files

    def is_cdr_file(self,file_path):
        """检查文件是否为 .cdr 格式"""
        return file_path.lower().endswith('.cdr')
class FolderApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CDR批量调整大小(合单版)-zc")

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 输入框，用于显示或输入文件夹路径
        self.folder_input = QLineEdit(self)
        self.folder_input.setPlaceholderText("输入或选择文件夹路径")

        # 选择文件夹按钮
        self.select_button = QPushButton("选择CDR所在的文件夹", self)
        self.select_button.clicked.connect(self.select_folder)


        # 创建按钮：选择文件
        self.open_button = QPushButton('统计列表，请选择 .xlsx 文件')
        self.open_button.clicked.connect(self.open_file_dialog)
        # 创建标签，用于显示选择的文件路径
        self.static_file = QLabel('未选择文件', self)
        # self.order_file = QLabel('E:/627个.xlsx', self)


        # 确定按钮
        self.confirm_button = QPushButton("确定", self)
        self.confirm_button.clicked.connect(self.on_confirm)

        # 取消按钮
        self.cancel_button = QPushButton("取消", self)
        self.cancel_button.clicked.connect(self.close)

        self.delete_button = QPushButton("删除中间文件")
        self.delete_button.clicked.connect(self.delete_intermediate_files)

        # 进度条
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setHidden(True)

        # 布局设置
        layout.addWidget(self.folder_input)
        layout.addWidget(self.select_button)
        # 将按钮和标签加入布局
        layout.addWidget(self.open_button)
        layout.addWidget(self.static_file)
        # 创建水平布局，将按钮放在同一排
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.confirm_button)
        button_layout.addWidget(self.cancel_button)

        # 将按钮布局添加到主布局
        layout.addLayout(button_layout)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)
        self.setGeometry(200, 200, 600, 150)

    def select_folder(self):
        """打开文件夹选择对话框"""
        folder_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder_path:
            self.folder_input.setText(folder_path)

    def open_file_dialog(self):
        # 打开文件选择框，限制只选择 .xlsx 文件
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "选择 .xlsx 文件", "", "Excel 文件 (*.xlsx);;所有文件 (*)",
                                                   options=options)

        if file_name:
            # 显示选择的文件路径
            self.static_file.setText(f"{file_name}")
        else:
            self.static_file.setText("未选择文件")

    def on_confirm(self):
        """点击确认按钮后的操作"""
        folder_path = self.folder_input.text()
        static_file = self.static_file.text()
        if not folder_path or not os.path.exists(folder_path):
            QMessageBox.warning(self, "错误", "请提供有效的文件夹路径")
            return

        if not static_file or not os.path.isfile(static_file):
            QMessageBox.warning(self, '警告', '统计excel选择有误！')
            return

        # 显示进度条
        self.progress_bar.setHidden(False)

        # 创建并启动后台线程执行任务
        self.thread = LoadingThread(folder_path,static_file)
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.on_loading_finished)
        self.thread.start()

        # 禁用按钮，防止重复点击
        self.confirm_button.setDisabled(True)
        self.cancel_button.setDisabled(True)

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    def on_loading_finished(self):
        """任务完成后的操作"""
        self.progress_bar.setHidden(True)
        QMessageBox.information(self, "完成", "操作完成！")

        # 启用按钮，允许再次操作
        self.confirm_button.setEnabled(True)
        self.cancel_button.setEnabled(True)

    def delete_intermediate_files(self):
        """点击确认按钮后的操作"""
        folder_path = self.folder_input.text()
        if not folder_path or not os.path.exists(folder_path):
            QMessageBox.warning(self, "错误", "请提供有效的文件夹路径")
            return
        # 遍历文件夹中的所有文件
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                # 检查文件名是否以 "Backup_of_" 开头且以 ".cdr" 结尾
                if file.startswith("Backup_of_") and file.endswith(".cdr"):
                    file_path = os.path.join(root, file)
                    try:
                        # 删除文件
                        os.remove(file_path)
                        print(f"Deleted: {file_path}")
                    except Exception as e:
                        print(f"Failed to delete {file_path}: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FolderApp()
    window.show()
    sys.exit(app.exec_())