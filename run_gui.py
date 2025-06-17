#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
视频分析助手启动脚本
运行前请确保已安装PyQt6: pip install PyQt6
"""

import os
import sys

# 禁止显示macOS的输入法相关警告
os.environ['QT_MAC_WANTS_LAYER'] = '1'  # 解决macOS上的部分渲染问题
os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.qpa.*=false'  # 禁用Qt调试日志

# 重定向stderr以过滤掉特定警告
if sys.platform == 'darwin':  # 仅在macOS上执行
    import io
    
    class WarningFilter(io.StringIO):
        def __init__(self, real_stderr):
            super().__init__()
            self.real_stderr = real_stderr
            
        def write(self, text):
            # 过滤掉特定的警告信息
            if any(pattern in text for pattern in [
                '_TIPropertyValueIsValid',
                'imkxpc_setApplicationProperty',
                'TSM AdjustCapsLockLEDForKeyTransitionHandling',
                'Layer-backing is always enabled',
                'QT_MAC_WANTS_LAYER',
                'qt.qpa.drawing',
                'qt.qpa'
            ]):
                return  # 忽略这些警告
            self.real_stderr.write(text)
            
        def flush(self):
            self.real_stderr.flush()
    
    sys.stderr = WarningFilter(sys.stderr)

try:
    from video_analysis_gui import main
    main()
except ImportError as e:
    print("缺少必要的依赖包，请运行以下命令安装：")
    print("pip install PyQt6")
    print(f"错误详情: {e}")
except Exception as e:
    print(f"程序运行出错: {e}")
    input("按任意键退出...") 