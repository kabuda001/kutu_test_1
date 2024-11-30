import os
import sys


from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QHBoxLayout, QLabel,QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal


class LoadThread(QThread):
    # 定义信号，当任务完成时发送消息
    finished_signal = pyqtSignal(str)

    def __init__(self, folder1):
        super().__init__()
        self.folder1 = folder1
        self.out_file = os.path.join(folder1, '输出.txt')
        if os.path.exists(self.out_file):
            os.remove(self.out_file)
        else:
            pass

    def run(self):
        # 在这个方法中执行耗时的任务
        try:
            print(f"开始加载文件夹1: {self.folder1}")
            self.write_to_out_file(self.folder1)
            # 完成后，发送信号到主线程
            self.finished_signal.emit("处理完毕!")
        except Exception as e:
            self.finished_signal.emit(f"处理失败: {str(e)}")

    def write_to_out_file(self,folder1):
        # 获取文件夹中的所有文件名
        file_names = []
        for root, dirs, files in os.walk(folder1):
            for file in files:
                if self.is_cdr_file(file):
                    file_names.append(os.path.splitext(os.path.basename(file))[0]) # 获取文件名（不含路径）
        # 将文件名以逗号分隔并写入文本文件
        with open(self.out_file, 'a') as txt_file:  # 使用 'a' 模式打开文件，确保追加内容
            # 将文件名列表转换为逗号分隔的字符串
            file_names_str = ','.join(file_names)
            # 写入文件并添加换行符
            txt_file.write(file_names_str + '\n')

    def is_cdr_file(self,file_path):
        """检查文件是否为 .cdr 格式"""
        return file_path.lower().endswith('.cdr')

class FolderSelector(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        # 创建控件
        self.label1 = QLabel('请选择文件夹:')

        self.folder1_path = QLabel('未选择文件夹')
        # self.folder1_path = QLabel('//xinguo/贴纸.生产资料/美工做的新款，待检查')


        self.select_folder1_btn = QPushButton('文件夹')
        self.ok_btn = QPushButton('确定')
        self.cancel_btn = QPushButton('取消')

        # 按钮点击事件
        self.select_folder1_btn.clicked.connect(self.select_folder1)
        self.ok_btn.clicked.connect(self.ok_clicked)
        self.cancel_btn.clicked.connect(self.cancel_clicked)

        # 布局设置
        layout = QVBoxLayout()

        # 第一个文件夹选择行
        folder1_layout = QHBoxLayout()
        folder1_layout.addWidget(self.label1)
        folder1_layout.addWidget(self.folder1_path)
        folder1_layout.addWidget(self.select_folder1_btn)
        layout.addLayout(folder1_layout)


        # 确定和取消按钮
        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(self.ok_btn)
        buttons_layout.addWidget(self.cancel_btn)
        layout.addLayout(buttons_layout)

        self.setLayout(layout)

        # 窗口设置
        self.setWindowTitle('文件名组合-zc')
        self.setGeometry(200, 200, 600, 150)


    def select_folder1(self):
        folder_path = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if folder_path:
            self.folder1_path.setText(folder_path)


    def ok_clicked(self):
        # 获取选择的文件夹路径
        folder1 = self.folder1_path.text()
        if not folder1  or not os.path.exists(folder1):
            QMessageBox.warning(self, '警告', '请确保文件夹已经选择且正确！')
            return
        # 启动后台线程执行任务
        self.thread = LoadThread(folder1)
        self.thread.finished_signal.connect(self.on_load_finished)
        self.thread.start()

        # 禁用按钮以防止重复点击
        self.ok_btn.setDisabled(True)
        self.cancel_btn.setDisabled(True)

    def cancel_clicked(self):
        # 创建确认对话框
        reply = QMessageBox.question(self, '确认取消', '你确定要取消操作吗?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        print("操作已取消")
        self.close()

    def on_load_finished(self, message):
        # 加载完成后弹出提示框
        QMessageBox.information(self, '提示', message)

        # 启用按钮
        self.ok_btn.setEnabled(True)
        self.cancel_btn.setEnabled(True)

        # 如果操作完成，关闭窗口
        self.close()


def main():
    app = QApplication(sys.argv)
    window = FolderSelector()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()