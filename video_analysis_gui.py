import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QGridLayout, QLabel, QLineEdit, QPushButton, 
                             QRadioButton, QButtonGroup, QTextEdit, QFileDialog,
                             QMessageBox, QFrame, QSizePolicy, QScrollArea)
from PyQt6.QtCore import Qt, pyqtSignal, QSettings
from PyQt6.QtGui import QFont, QPixmap, QIcon
from video_analysis_engine import VideoAnalysisEngine
import json
import os

class VideoAnalysisGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("VideoAnalysis", "VideoAnalysisGUI")
        self.loading_settings = False  # 标记是否正在加载设置
        self.init_ui()
        self.load_settings()
        
    def init_ui(self):
        self.setWindowTitle("视频分析助手")
        self.setFixedSize(950, 1000)  # 再次增加窗口大小以容纳所有组件
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #333333;
                font-size: 14px;
            }
            QLineEdit, QTextEdit {
                padding: 8px;
                border: 2px solid #ddd;
                border-radius: 6px;
                font-size: 14px;
                background-color: #ffffff;
                color: #333333;
            }
            QLineEdit:focus, QTextEdit:focus {
                border-color: #4CAF50;
            }
            QPushButton {
                padding: 10px 20px;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: bold;
                background-color: #4CAF50;
                color: white;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton#browseButton {
                background-color: #2196F3;
                padding: 8px 16px;
            }
            QPushButton#browseButton:hover {
                background-color: #1976D2;
            }
            QPushButton#cancelButton {
                background-color: #f44336;
            }
            QPushButton#cancelButton:hover {
                background-color: #d32f2f;
            }
            QRadioButton {
                font-size: 14px;
                color: #333333;
                spacing: 8px;
            }
            QRadioButton::indicator {
                width: 18px;
                height: 18px;
            }
            QRadioButton::indicator:unchecked {
                border: 2px solid #ddd;
                border-radius: 9px;
                background-color: white;
            }
            QRadioButton::indicator:checked {
                border: 2px solid #4CAF50;
                border-radius: 9px;
                background-color: #4CAF50;
            }
            QTextEdit {
                border: 2px solid #ddd;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
                background-color: white;
            }
            QTextEdit:focus {
                border-color: #4CAF50;
            }
            QFrame#separator {
                background-color: #ddd;
                max-height: 1px;
            }
        """)
        
        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(12)  # 减少间距
        main_layout.setContentsMargins(25, 15, 25, 15)  # 减少边距
        
        # 分析类型选择
        type_label = QLabel("本程序支持两种使用方式:")
        type_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        main_layout.addWidget(type_label)
        
        # 单选按钮组
        self.button_group = QButtonGroup()
        
        self.youtube_radio = QRadioButton("YouTube分析")
        self.youtube_radio.setChecked(True)
        self.youtube_radio.toggled.connect(self.on_type_changed)
        self.button_group.addButton(self.youtube_radio)
        main_layout.addWidget(self.youtube_radio)
        
        self.local_radio = QRadioButton("本地视频分析")
        self.local_radio.toggled.connect(self.on_type_changed)
        self.button_group.addButton(self.local_radio)
        main_layout.addWidget(self.local_radio)
        
        # YouTube分析配置
        self.youtube_frame = QWidget()
        youtube_layout = QVBoxLayout(self.youtube_frame)
        youtube_layout.setContentsMargins(0, 10, 0, 10)
        
        youtube_url_label = QLabel("选择对应视频链接Excel文件:")
        youtube_layout.addWidget(youtube_url_label)
        
        youtube_input_layout = QHBoxLayout()
        self.youtube_path_input = QLineEdit()
        self.youtube_path_input.setPlaceholderText("请选择包含YouTube链接的Excel文件")
        self.youtube_path_input.setMinimumHeight(40)  # 统一高度
        self.youtube_path_input.setMinimumWidth(500)  # 设置最小宽度
        youtube_input_layout.addWidget(self.youtube_path_input, 1)  # 给输入框更多空间
        
        self.youtube_browse_btn = QPushButton("浏览")
        self.youtube_browse_btn.setObjectName("browseButton")
        self.youtube_browse_btn.setMinimumHeight(40)  # 与输入框同高
        self.youtube_browse_btn.setMinimumWidth(80)   # 设置最小宽度
        self.youtube_browse_btn.clicked.connect(self.browse_youtube_excel)
        youtube_input_layout.addWidget(self.youtube_browse_btn, 0)  # 按钮固定大小
        
        youtube_layout.addLayout(youtube_input_layout)
        main_layout.addWidget(self.youtube_frame)
        
        # 本地视频分析配置
        self.local_frame = QWidget()
        local_layout = QVBoxLayout(self.local_frame)
        local_layout.setContentsMargins(0, 10, 0, 10)
        
        local_path_label = QLabel("选择视频保存路径:")
        local_layout.addWidget(local_path_label)
        
        local_input_layout = QHBoxLayout()
        self.local_path_input = QLineEdit()
        self.local_path_input.setPlaceholderText("请选择视频保存的文件夹路径")
        self.local_path_input.setMinimumHeight(40)  # 统一高度
        self.local_path_input.setMinimumWidth(500)  # 设置最小宽度
        local_input_layout.addWidget(self.local_path_input, 1)  # 给输入框更多空间
        
        self.local_browse_btn = QPushButton("浏览")
        self.local_browse_btn.setObjectName("browseButton")
        self.local_browse_btn.setMinimumHeight(40)  # 与输入框同高
        self.local_browse_btn.setMinimumWidth(80)   # 设置最小宽度
        self.local_browse_btn.clicked.connect(self.browse_local_folder)
        local_input_layout.addWidget(self.local_browse_btn, 0)  # 按钮固定大小
        
        local_layout.addLayout(local_input_layout)
        main_layout.addWidget(self.local_frame)
        
        # 输出内容保存路径
        output_label = QLabel("输出内容保存路径:")
        output_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        main_layout.addWidget(output_label)
        
        output_input_layout = QHBoxLayout()
        self.output_path_input = QLineEdit()
        self.output_path_input.setPlaceholderText("请选择分析结果保存的文件夹路径")
        self.output_path_input.setMinimumHeight(40)  # 统一高度
        self.output_path_input.setMinimumWidth(500)  # 设置最小宽度
        output_input_layout.addWidget(self.output_path_input, 1)  # 给输入框更多空间
        
        self.output_browse_btn = QPushButton("浏览")
        self.output_browse_btn.setObjectName("browseButton")
        self.output_browse_btn.setMinimumHeight(40)  # 与输入框同高
        self.output_browse_btn.setMinimumWidth(80)   # 设置最小宽度
        self.output_browse_btn.clicked.connect(self.browse_output_folder)
        output_input_layout.addWidget(self.output_browse_btn, 0)  # 按钮固定大小
        
        main_layout.addLayout(output_input_layout)
        
        # 新增：比特浏览器窗口ID输入
        self.bit_window_id_label = QLabel("比特浏览器窗口ID:")
        self.bit_window_id_input = QLineEdit()
        self.bit_window_id_input.setPlaceholderText("请从比特浏览器客户端复制窗口ID")
        main_layout.addWidget(self.bit_window_id_label)
        main_layout.addWidget(self.bit_window_id_input)
        
        # 操作延时配置 - 超简化版本
        delay_layout = QHBoxLayout()
        delay_layout.setContentsMargins(0, 10, 0, 10)
        
        delay_label = QLabel("操作延时范围(秒):")
        delay_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        delay_layout.addWidget(delay_label)
        
        self.min_delay_input = QLineEdit("1")
        self.min_delay_input.setFixedWidth(60)
        self.min_delay_input.setFixedHeight(35)
        self.min_delay_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.min_delay_input.setFont(QFont("Arial", 14))
        delay_layout.addWidget(self.min_delay_input)
        
        dash_label = QLabel(" - ")
        dash_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        delay_layout.addWidget(dash_label)
        
        self.max_delay_input = QLineEdit("3")
        self.max_delay_input.setFixedWidth(60)
        self.max_delay_input.setFixedHeight(35)
        self.max_delay_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.max_delay_input.setFont(QFont("Arial", 14))
        delay_layout.addWidget(self.max_delay_input)
        
        help_label = QLabel("  (每操作后随机等待，避免机器人检测)")
        help_label.setStyleSheet("color: #666; font-size: 12px;")
        delay_layout.addWidget(help_label)
        
        delay_layout.addStretch()
        main_layout.addLayout(delay_layout)
        
        # 分析提示输入
        prompt_label = QLabel("输入您的分析提示词:")
        prompt_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        main_layout.addWidget(prompt_label)
        
        self.prompt_text = QTextEdit()
        self.prompt_text.setMinimumHeight(120)  # 减少高度
        self.prompt_text.setMaximumHeight(120)
        self.prompt_text.setPlainText("## 【视频分析任务重要提示生成 v4.0 - 全面关关系统化版】\n\n### 1. 角色定义与核心目标")
        main_layout.addWidget(self.prompt_text)
        
        # 添加一些间距
        main_layout.addSpacing(10)  # 减少间距
        
        # 日志显示区域 - 与其他输入框风格一致
        log_label = QLabel("运行日志:")
        log_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        main_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(100)
        self.log_text.setMinimumHeight(100)
        self.log_text.setReadOnly(True)
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.setPlainText(f"[{timestamp}] 准备就绪")
        self.log_text.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)  # 启用自动换行
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 2px solid #ddd;
                border-radius: 6px;
                padding: 8px;
                font-size: 12px;
                background-color: white;
            }
        """)
        main_layout.addWidget(self.log_text)
        main_layout.addSpacing(5)  # 减少间距
        
        # 按钮区域
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.start_btn = QPushButton("开始分析")
        self.start_btn.clicked.connect(self.start_analysis)
        button_layout.addWidget(self.start_btn)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setObjectName("cancelButton")
        self.cancel_btn.clicked.connect(self.close)
        button_layout.addWidget(self.cancel_btn)
        
        main_layout.addLayout(button_layout)
        
        # 应用初始状态
        self.on_type_changed(True)
    
    def load_settings(self):
        """加载用户设置"""
        try:
            self.loading_settings = True  # 设置加载标记
            
            # 加载分析类型
            analysis_type = self.settings.value("analysis_type", "youtube")
            if analysis_type == "youtube":
                self.youtube_radio.setChecked(True)
            else:
                self.local_radio.setChecked(True)
            
            # 加载文件路径
            youtube_path = self.settings.value("youtube_path", "")
            if youtube_path:
                self.youtube_path_input.setText(youtube_path)
            
            local_path = self.settings.value("local_path", "")
            if local_path:
                self.local_path_input.setText(local_path)
            
            output_path = self.settings.value("output_path", "")
            if output_path:
                self.output_path_input.setText(output_path)
            
            # 加载提示词
            prompt = self.settings.value("prompt", "")
            if prompt:
                self.prompt_text.setPlainText(prompt)
            
            # 加载比特浏览器窗口ID
            bit_window_id = self.settings.value("bit_window_id", "")
            if bit_window_id:
                self.bit_window_id_input.setText(bit_window_id)
            
            # 加载延时配置
            self.min_delay_input.setText(str(self.settings.value("min_delay", "1")))
            self.max_delay_input.setText(str(self.settings.value("max_delay", "3")))
            
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.append(f"[{timestamp}] 已加载上次的设置")
            
        except Exception as e:
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.append(f"[{timestamp}] 加载设置失败: {str(e)}")
        finally:
            self.loading_settings = False  # 确保总是清除加载标记
    
    def save_settings(self):
        """保存用户设置"""
        try:
            # 保存分析类型
            analysis_type = "youtube" if self.youtube_radio.isChecked() else "local"
            self.settings.setValue("analysis_type", analysis_type)
            
            # 保存文件路径
            self.settings.setValue("youtube_path", self.youtube_path_input.text())
            self.settings.setValue("local_path", self.local_path_input.text())
            self.settings.setValue("output_path", self.output_path_input.text())
            
            # 保存提示词
            self.settings.setValue("prompt", self.prompt_text.toPlainText())
            
            # 保存比特浏览器窗口ID
            self.settings.setValue("bit_window_id", self.bit_window_id_input.text())
            
            # 保存延时配置
            self.settings.setValue("min_delay", self.min_delay_input.text())
            self.settings.setValue("max_delay", self.max_delay_input.text())

        except Exception as e:
            # 在这种情况下，我们不希望有任何弹窗，只是静默失败
            print(f"Error saving settings: {e}")
        
    def on_type_changed(self, checked):
        """当分析类型改变时更新界面"""
        if self.youtube_radio.isChecked():
            self.youtube_frame.setVisible(True)
            self.local_frame.setVisible(False)
        else:
            self.youtube_frame.setVisible(False)
            self.local_frame.setVisible(True)
    
    def browse_youtube_excel(self):
        """浏览选择YouTube链接Excel文件"""
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel 文件 (*.xlsx *.xls)")
        if path:
            self.youtube_path_input.setText(path)
            
    def browse_local_folder(self):
        """浏览选择本地视频文件夹"""
        path = QFileDialog.getExistingDirectory(self, "选择视频文件夹", "")
        if path:
            self.local_path_input.setText(path)

    def browse_output_folder(self):
        """浏览选择输出内容保存文件夹"""
        path = QFileDialog.getExistingDirectory(self, "选择输出文件夹", "")
        if path:
            self.output_path_input.setText(path)

    def validate_inputs(self):
        """验证输入是否有效"""
        if self.youtube_radio.isChecked():
            if not self.youtube_path_input.text().strip():
                QMessageBox.warning(self, "错误", "请选择包含YouTube链接的Excel文件")
                return False
            if not os.path.exists(self.youtube_path_input.text()):
                QMessageBox.warning(self, "错误", "选择的Excel文件不存在")
                return False
        else:
            if not self.local_path_input.text().strip():
                QMessageBox.warning(self, "错误", "请选择视频保存文件夹")
                return False
            if not os.path.exists(self.local_path_input.text()):
                QMessageBox.warning(self, "错误", "选择的文件夹不存在")
                return False
        
        if not self.output_path_input.text().strip():
            QMessageBox.warning(self, "错误", "请选择输出内容保存文件夹")
            return False
        
        if not os.path.exists(self.output_path_input.text()):
            QMessageBox.warning(self, "错误", "选择的输出文件夹不存在")
            return False
        
        if not self.prompt_text.toPlainText().strip():
            QMessageBox.warning(self, "错误", "请输入分析提示词")
            return False
        
        if not self.bit_window_id_input.text().strip():
            QMessageBox.warning(self, "错误", "请输入比特浏览器窗口ID")
            return False
        
        return True
    
    def start_analysis(self):
        """开始分析"""
        if not self.validate_inputs():
            return
        
        # 获取配置信息
        is_youtube = self.youtube_radio.isChecked()
        analysis_type = "YouTube分析" if is_youtube else "本地视频分析"
        file_path = self.youtube_path_input.text() if is_youtube else self.local_path_input.text()
        output_path = self.output_path_input.text()
        prompt = self.prompt_text.toPlainText().strip()
        bit_window_id = self.bit_window_id_input.text().strip()
        
        # 获取延时配置
        try:
            min_delay = float(self.min_delay_input.text() or "1")
            max_delay = float(self.max_delay_input.text() or "3")
        except ValueError:
            min_delay = 1
            max_delay = 3
            
        # 显示配置确认
        config_text = f"""分析配置:
类型: {analysis_type}
{'Excel文件' if self.youtube_radio.isChecked() else '视频文件夹'}: {file_path}
输出路径: {output_path}
提示词: {prompt[:100]}{'...' if len(prompt) > 100 else ''}
比特浏览器窗口ID: {bit_window_id}
操作延时范围: {min_delay}-{max_delay}秒

确认开始分析吗？"""
        
        reply = QMessageBox.question(
            self, 
            "确认分析", 
            config_text,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # 开始分析
            self.start_analysis_process(file_path, output_path, prompt, bit_window_id)
    
    def start_analysis_process(self, file_path, output_path, prompt, bit_window_id):
        """开始分析处理"""
        try:
            # 保存设置
            self.save_settings()
            
            # 禁用开始按钮，防止重复点击
            self.start_btn.setEnabled(False)
            self.start_btn.setText("分析中...")
            
            # 清空日志并显示开始信息
            self.log_text.clear()
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.append(f"[{timestamp}] 开始分析...")
            
            # 获取延时配置
            try:
                min_delay = float(self.min_delay_input.text() or "1")
                max_delay = float(self.max_delay_input.text() or "3")
                if min_delay >= max_delay:
                    self.analysis_error("最小延时必须小于最大延时")
                    return
            except ValueError:
                self.analysis_error("延时配置格式错误，请输入有效的数字")
                return
            
            # 创建分析配置
            config = {
                'analysis_type': 'youtube' if self.youtube_radio.isChecked() else 'local',
                'file_path': file_path,
                'output_path': output_path,
                'prompt': prompt,
                'bit_window_id': bit_window_id,
                # 延时配置
                'min_delay': min_delay,
                'max_delay': max_delay
            }
            
            # 创建并启动分析引擎
            self.analysis_engine = VideoAnalysisEngine(config)
            self.analysis_engine.progress_update.connect(self.update_log)
            self.analysis_engine.analysis_complete.connect(self.analysis_finished)
            self.analysis_engine.error_occurred.connect(self.analysis_error)
            self.analysis_engine.start()
            
        except Exception as e:
            self.analysis_error(f"启动分析失败: {str(e)}")
    
    def update_log(self, message):
        """更新日志显示"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text.append(formatted_message)
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def analysis_finished(self, result):
        """分析完成处理"""
        self.start_btn.setEnabled(True)
        self.start_btn.setText("开始分析")
        
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if result['success']:
            self.log_text.append(f"[{timestamp}] ✅ 分析完成！{result['message']}")
            QMessageBox.information(
                self, 
                "分析完成", 
                f"{result['message']}\n\n结果已保存到指定路径。"
            )
        else:
            self.log_text.append(f"[{timestamp}] ❌ 分析失败: {result.get('message', '未知错误')}")
            QMessageBox.warning(self, "分析失败", result.get('message', '未知错误'))
    
    def analysis_error(self, error_message):
        """分析错误处理"""
        self.start_btn.setEnabled(True)
        self.start_btn.setText("开始分析")
        
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] ❌ 分析错误: {error_message}")
        
        QMessageBox.critical(self, "分析错误", error_message)

    def closeEvent(self, event):
        """在窗口关闭时保存设置"""
        self.save_settings()
        self.settings.sync() # 确保设置被立即写入
        super().closeEvent(event)

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion样式获得更好的外观
    
    window = VideoAnalysisGUI()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main() 