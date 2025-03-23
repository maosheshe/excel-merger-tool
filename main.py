import sys
import os
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                           QWidget, QFileDialog, QMessageBox, QListWidget, QLabel,
                           QHBoxLayout, QTableWidget, QTableWidgetItem, QProgressBar)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QIcon, QColor
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill
from excel_processor import ExcelProcessor
from file_preview import FilePreviewWindow

class MergeWorker(QThread):
    progress_updated = Signal(int)  # 进度信号
    finished = Signal(bool, str)    # 完成信号
    error = Signal(str)             # 错误信号

    def __init__(self, processor, files, template_file):
        super().__init__()
        self.processor = processor
        self.files = files
        self.template_file = template_file

    def run(self):
        try:
            # 更新进度：开始处理
            self.progress_updated.emit(10)

            # 合并文件
            merged_data, message = self.processor.merge_files(self.files)
            if merged_data is None:
                self.error.emit(message)
                return

            # 更新进度：文件合并完成
            self.progress_updated.emit(50)

            # 生成输出文件名
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            current_date = datetime.now().strftime('%Y%m%d')
            output_file = os.path.join(desktop_path, f'附录2：营销现场作业计划审批表_{current_date}.xlsx')

            # 保存文件
            success, message = self.processor.save_output(self.template_file, merged_data, output_file)

            # 更新进度：完成
            self.progress_updated.emit(100)

            # 发送完成信号
            self.finished.emit(success, message if success else f"保存失败：{message}")

        except Exception as e:
            self.error.emit(str(e))

class ExcelMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel文件合并工具")
        self.setMinimumSize(600, 400)
        
        # 获取程序运行目录
        self.exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        if getattr(sys, 'frozen', False):
            # 如果是打包后的exe
            self.default_template = os.path.join(self.exe_dir, "输出模版.xlsx")
        else:
            # 如果是开发环境
            self.default_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "输出模版.xlsx")
        
        # 初始化变量
        self.selected_files = []
        self.template_file = None
        self.preview_window = None
        self.processor = ExcelProcessor()
        
        # 检查默认模板是否存在
        if os.path.exists(self.default_template):
            self.template_file = self.default_template
        
        self.init_ui()
        self.center_window()

    def center_window(self):
        """将窗口居中显示"""
        screen = QApplication.primaryScreen().geometry()
        size = self.geometry()
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        self.move(x, y)

    def init_ui(self):
        """初始化UI"""
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建主布局
        layout = QVBoxLayout()
        
        # 创建按钮布局
        button_layout = QHBoxLayout()
        
        # 创建选择文件按钮
        self.select_button = QPushButton("选择Excel文件")
        self.select_button.clicked.connect(self.select_files)
        button_layout.addWidget(self.select_button)
        
        # 创建清除选择按钮
        self.clear_button = QPushButton("清除选择")
        self.clear_button.clicked.connect(self.clear_selection)
        button_layout.addWidget(self.clear_button)
        
        # 创建规范检查按钮
        self.preview_button = QPushButton("规范检查")
        self.preview_button.clicked.connect(self.preview_file)
        button_layout.addWidget(self.preview_button)
        
        layout.addLayout(button_layout)
        
        # 创建文件列表
        self.file_list = QListWidget()
        layout.addWidget(self.file_list)
        
        # 创建选择模板按钮
        self.template_button = QPushButton("选择输出模板")
        self.template_button.clicked.connect(self.select_template)
        layout.addWidget(self.template_button)
        
        # 创建模板标签
        self.template_label = QLabel("未选择模板文件" if not self.template_file else f"已选择模板：\n{os.path.basename(self.template_file)}")
        layout.addWidget(self.template_label)
        
        # 创建合并按钮
        self.merge_button = QPushButton("合并文件")
        self.merge_button.clicked.connect(self.merge_files)
        layout.addWidget(self.merge_button)
        
        # 创建进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # 创建状态标签
        self.status_label = QLabel("")
        layout.addWidget(self.status_label)
        
        main_widget.setLayout(layout)

    def clear_selection(self):
        """清除已选择的文件"""
        self.selected_files = []
        self.file_list.clear()
        self.status_label.setText('已清除选择的文件')

    def select_files(self):
        """选择Excel文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择Excel文件",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if files:
            if len(files) > 5:
                QMessageBox.warning(self, '警告', '最多只能选择5个文件！')
                files = files[:5]
            
            self.selected_files = files
            self.file_list.addItems([os.path.basename(f) for f in files])
            self.status_label.setText(f'已选择 {len(files)} 个文件')

    def select_template(self):
        """选择模板文件"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "选择输出模板",
            os.path.dirname(self.template_file) if self.template_file else self.exe_dir,
            "Excel Files (*.xlsx)"
        )
        
        if file:
            self.template_file = file
            self.template_label.setText(f"已选择模板：\n{os.path.basename(file)}")
            self.status_label.setText("已选择模板文件")

    def merge_files(self):
        """合并文件"""
        if not self.selected_files:
            QMessageBox.warning(self, '警告', '请先选择要合并的Excel文件！')
            return
        
        if not self.template_file:
            QMessageBox.warning(self, '警告', '请先选择输出模板文件！')
            return

        # 禁用按钮，显示进度条
        self.select_button.setEnabled(False)
        self.clear_button.setEnabled(False)
        self.template_button.setEnabled(False)
        self.merge_button.setEnabled(False)
        self.preview_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # 创建并启动工作线程
        self.merge_thread = MergeWorker(self.processor, self.selected_files, self.template_file)
        self.merge_thread.progress_updated.connect(self.update_progress)
        self.merge_thread.finished.connect(self.merge_completed)
        self.merge_thread.error.connect(self.merge_error)
        self.merge_thread.start()

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    def merge_completed(self, success, message):
        """合并完成的处理"""
        # 恢复按钮状态
        self.select_button.setEnabled(True)
        self.clear_button.setEnabled(True)
        self.template_button.setEnabled(True)
        self.merge_button.setEnabled(True)
        self.preview_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # 显示成功消息
        if success:
            QMessageBox.information(self, "成功", message)
            self.status_label.setText(message)
        else:
            QMessageBox.critical(self, '错误', message)
            self.status_label.setText('合并失败：' + message)

    def merge_error(self, error_message):
        """合并错误的处理"""
        QMessageBox.critical(self, '错误', error_message)
        self.status_label.setText('合并失败：' + error_message)
        
        # 恢复按钮状态
        self.select_button.setEnabled(True)
        self.clear_button.setEnabled(True)
        self.template_button.setEnabled(True)
        self.merge_button.setEnabled(True)
        self.preview_button.setEnabled(True)
        self.progress_bar.setVisible(False)

    def preview_file(self):
        """打开文件预览窗口"""
        if not self.preview_window:
            self.preview_window = FilePreviewWindow()
        self.preview_window.show()

def main():
    app = QApplication(sys.argv)
    window = ExcelMergerApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 