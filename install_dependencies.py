#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
视频分析助手依赖安装脚本
"""

import subprocess
import sys
import os

def run_command(command):
    """运行命令并显示输出"""
    print(f"执行命令: {command}")
    try:
        result = subprocess.run(command, shell=True, check=True, 
                              capture_output=True, text=True)
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"命令执行失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False

def install_dependencies():
    """安装依赖包"""
    print("正在安装视频分析助手的依赖包...")
    print("=" * 50)
    
    # 更新pip
    print("1. 更新pip...")
    run_command(f"{sys.executable} -m pip install --upgrade pip")
    
    # 安装基础依赖
    print("\n2. 安装基础依赖...")
    dependencies = [
        "PyQt6>=6.5.0",
        "playwright>=1.40.0", 
        "pandas>=2.0.0",
        "openpyxl>=3.1.0"
    ]
    
    for dep in dependencies:
        print(f"\n安装 {dep}...")
        if not run_command(f"{sys.executable} -m pip install {dep}"):
            print(f"安装 {dep} 失败，请检查网络连接或手动安装")
            return False
    
    # 安装Playwright浏览器
    print("\n3. 安装Playwright浏览器...")
    if not run_command(f"{sys.executable} -m playwright install chromium"):
        print("安装Playwright浏览器失败")
        return False
    
    print("\n" + "=" * 50)
    print("✅ 所有依赖安装完成！")
    print("\n现在您可以运行以下命令启动程序：")
    print("python run_gui.py")
    
    return True

def check_python_version():
    """检查Python版本"""
    if sys.version_info < (3, 8):
        print("❌ 错误：需要Python 3.8或更高版本")
        print(f"当前版本：{sys.version}")
        return False
    
    print(f"✅ Python版本检查通过：{sys.version}")
    return True

def main():
    """主函数"""
    print("视频分析助手依赖安装程序")
    print("=" * 50)
    
    # 检查Python版本
    if not check_python_version():
        input("按任意键退出...")
        return
    
    # 安装依赖
    if install_dependencies():
        print("\n🎉 安装成功！程序已准备就绪。")
    else:
        print("\n❌ 安装过程中出现错误，请检查网络连接或手动安装依赖。")
    
    input("\n按任意键退出...")

if __name__ == "__main__":
    main() 