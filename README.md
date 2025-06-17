# 视频分析助手

一个基于PyQt6和Playwright的视频分析工具，支持YouTube视频和本地视频的自动化分析，使用Gemini AI Studio进行智能分析并生成Excel格式的分镜表。

## 功能特点

- 🎬 **YouTube视频分析**：从Excel文件读取YouTube链接，批量分析视频内容
- 📁 **本地视频分析**：处理本地视频文件（开发中）
- 🤖 **AI智能分析**：使用Gemini AI Studio进行视频内容分析
- 📊 **Excel输出**：自动生成包含分镜信息的Excel表格
- 🎨 **现代化界面**：基于PyQt6的美观用户界面
- ⚡ **自动化操作**：使用Playwright自动化浏览器操作

## 系统要求

- Python 3.8 或更高版本
- Chrome浏览器
- 网络连接（用于访问Gemini AI Studio）

## 安装步骤

### 方法一：自动安装（推荐）

1. 下载项目文件到本地
2. 运行安装脚本：
   ```bash
   python install_dependencies.py
   ```
3. 等待安装完成

### 方法二：手动安装

1. 安装Python依赖：
   ```bash
   pip install PyQt6 playwright pandas openpyxl
   ```

2. 安装Playwright浏览器：
   ```bash
   playwright install chromium
   ```

## 使用方法

### 启动程序

```bash
python run_gui.py
```

程序会自动启动Chrome浏览器并保留您的登录信息，无需额外设置。

### YouTube视频分析

1. **选择分析类型**：选择"YouTube分析"
2. **选择Excel文件**：点击"浏览"选择包含YouTube链接的Excel文件
   - Excel文件格式：第一列为视频标题，第二列为YouTube链接
   - 程序会自动识别包含youtube.com或youtu.be的链接
3. **设置输出路径**：选择分析结果保存的文件夹
4. **输入分析提示词**：在文本框中输入或修改AI分析的提示词
5. **开始分析**：点击"开始分析"按钮

### 本地视频分析

1. **选择分析类型**：选择"本地视频分析"
2. **选择保存路径**：选择视频保存的文件夹路径
3. **设置输出路径**：选择分析结果保存的文件夹
4. **输入分析提示词**：在文本框中输入分析提示词
5. **开始分析**：点击"开始分析"按钮

## 分析流程

程序会自动执行以下步骤：

1. 📖 读取Excel文件中的YouTube链接
2. 🌐 打开Gemini AI Studio网页
3. ✍️ 输入分析提示词
4. 🎥 添加YouTube视频
5. ▶️ 运行AI分析
6. ⏳ 等待分析完成
7. 🔄 检测并处理分析错误（自动重试）
8. 📄 提取分析结果
9. 📊 保存为Excel格式

## 输出格式

分析结果会保存为Excel文件，包含以下列：

- **分镜**：分镜编号（分镜1、分镜2...）
- **关键帧图片生成提示词**：用于图片生成的提示词
- **图生视频提示词**：用于视频生成的提示词

## 注意事项

- 🔐 首次使用需要登录Google账号访问Gemini AI Studio
- 🌐 确保网络连接稳定
- ⏱️ 分析过程可能需要较长时间，请耐心等待
- 🔄 程序会自动处理分析错误并重试
- 💾 结果文件会按时间戳命名，避免覆盖

## 故障排除

### 常见问题

1. **浏览器启动失败**
   - 确保已安装Chrome浏览器
   - 运行 `playwright install chromium` 重新安装浏览器

2. **Excel文件读取失败**
   - 检查Excel文件格式是否正确
   - 确保文件中包含有效的YouTube链接

3. **网页元素找不到**
   - Gemini AI Studio界面可能已更新
   - 请联系开发者更新程序

4. **分析超时**
   - 检查网络连接
   - 尝试使用更简洁的提示词

## 文件结构

```
video_tools/
├── video_analysis_gui.py      # 主GUI界面
├── video_analysis_engine.py   # 分析引擎
├── run_gui.py                 # 启动脚本
├── start_chrome_debug.py      # Chrome调试模式启动脚本
├── install_dependencies.py   # 依赖安装脚本
├── requirements.txt           # 依赖列表
└── README.md                  # 说明文档
```

## 技术栈

- **GUI框架**：PyQt6
- **浏览器自动化**：Playwright
- **数据处理**：Pandas
- **Excel处理**：OpenPyXL
- **AI分析**：Gemini AI Studio

## 开发者信息

本项目使用现代化的Python技术栈开发，具有良好的扩展性和维护性。

## 许可证

本项目仅供学习和个人使用。 