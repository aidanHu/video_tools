import os
import time
import pandas as pd
import re
from datetime import datetime
import unicodedata
from playwright.sync_api import sync_playwright
from PyQt6.QtCore import QThread, pyqtSignal
import requests
import json
import random
import math
import shutil

class VideoAnalysisEngine(QThread):
    """视频分析引擎，使用Playwright和比特浏览器API进行自动化操作"""
    
    # 信号定义
    progress_update = pyqtSignal(str)  # 进度更新信号
    analysis_complete = pyqtSignal(dict)  # 分析完成信号
    error_occurred = pyqtSignal(str)  # 错误信号
    
    def __init__(self, config):
        super().__init__()
        self.config = config
        self.browser = None
        self.page = None
        # 添加延时配置，使用新的简化参数
        self.delay_config = {
            'min_delay': config.get('min_delay', 1),  # 最小延时时间（秒）
            'max_delay': config.get('max_delay', 3),  # 最大延时时间（秒）
        }
        
        # 记录延时配置到日志
        self.progress_update.emit(f"延时配置: 最小={self.delay_config['min_delay']}秒, 最大={self.delay_config['max_delay']}秒")
    
    def run(self):
        """主执行方法"""
        try:
            if self.config['analysis_type'] == 'youtube':
                self.analyze_youtube_videos()
            else:
                self.analyze_local_videos()
        except Exception as e:
            self.error_occurred.emit(f"分析过程中发生错误: {str(e)}")
    
    def analyze_youtube_videos(self):
        """分析YouTube视频，并标记已完成的任务"""
        try:
            self.progress_update.emit("正在读取并检查Excel文件...")
            excel_path = self.config['file_path']
            
            try:
                df = pd.read_excel(excel_path, engine='openpyxl')
            except FileNotFoundError:
                self.error_occurred.emit(f"Excel文件不存在: {excel_path}")
                return

            # 确保状态列存在
            status_col_name = "状态"
            if status_col_name not in df.columns:
                # 插入到第四列位置
                df.insert(3, status_col_name, "")

            # 找出需要分析的视频
            youtube_data = []
            for index, row in df.iterrows():
                # 检查是否已分析
                if row.get(status_col_name) == "已分析分镜提示词":
                    self.progress_update.emit(f"➡️ 跳过已完成: {row.iloc[0]}")
                    continue

                if len(row) >= 2:
                    title = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else f"视频_{index+1}"
                    url = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                    if 'youtube.com' in url or 'youtu.be' in url:
                        youtube_data.append({
                            'title': title,
                            'url': url,
                            'index': index # 存储原始DataFrame索引
                        })

            if not youtube_data:
                self.progress_update.emit("✅ 所有任务均已完成，无需分析。")
                self.analysis_complete.emit({'success': True, 'message': '所有任务均已完成', 'results_count': 0})
                return
                
            self.progress_update.emit(f"找到 {len(youtube_data)} 个新任务，开始分析...")
            self.start_browser()
            
            saved_count = 0
            total_videos = len(youtube_data)
            for i, video_data in enumerate(youtube_data):
                self.progress_update.emit(f"\n--- [ {i+1}/{total_videos} ] 开始处理: {video_data['title']} ---")
                
                try:
                    result = self.analyze_single_youtube_video(video_data['url'], video_data['title'])
                    
                    if result and result.get('content'):
                        self.progress_update.emit(f"✅ 分析完成，正在保存...")
                        if self.save_single_result(result):
                            saved_count += 1
                            self.progress_update.emit(f"--- ✅ [ {i+1}/{total_videos} ] 保存成功 ---")
                            
                            # 关键步骤：更新Excel状态并保存
                            try:
                                df.loc[video_data['index'], status_col_name] = "已分析分镜提示词"
                                df.to_excel(excel_path, index=False, engine='openpyxl')
                                self.progress_update.emit(f"✏️ 已在Excel中标记 '{video_data['title']}' 为完成。")
                            except Exception as e:
                                self.progress_update.emit(f"⚠️ 更新Excel文件失败: {e}")

                        else:
                            self.progress_update.emit(f"--- ❌ [ {i+1}/{total_videos} ] 保存失败 ---\n")
                    else:
                        self.progress_update.emit(f"⚠️ 分析未返回有效结果，跳过。")

                except Exception as e:
                    self.error_occurred.emit(f"处理 '{video_data['title']}' 时出错: {e}")
                    self.progress_update.emit("将尝试继续处理下一个视频...")

            self.progress_update.emit("--- ✅ 所有视频处理流程完毕 ---")
            self.analysis_complete.emit({'success': True, 'message': f'成功保存 {saved_count}/{total_videos} 个视频', 'results_count': saved_count})
            
        except Exception as e:
            self.error_occurred.emit(f"YouTube分析流程失败: {str(e)}")
        finally:
            self.cleanup_browser()
    
    def analyze_local_videos(self):
        """分析文件夹内视频，并将已完成的移入子文件夹"""
        try:
            self.progress_update.emit("开始本地视频批量分析...")
            folder_path = self.config['file_path']
            if not os.path.isdir(folder_path):
                self.error_occurred.emit(f"指定的路径不是一个有效的文件夹: {folder_path}")
                return

            # 创建"已分析"子文件夹
            completed_folder = os.path.join(folder_path, "已分析分镜提示词")
            os.makedirs(completed_folder, exist_ok=True)

            supported_formats = ['.mp4', '.mov', '.avi', '.mkv', '.webm', '.flv']
            video_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
                           if os.path.isfile(os.path.join(folder_path, f)) and 
                           os.path.splitext(f)[1].lower() in supported_formats]

            if not video_files:
                self.error_occurred.emit(f"在文件夹 {folder_path} 中未找到支持的视频文件。")
                return

            self.progress_update.emit(f"在文件夹中找到 {len(video_files)} 个视频文件，准备开始处理...")
            self.start_browser()

            saved_count = 0
            total_videos = len(video_files)
            for i, file_path in enumerate(video_files):
                video_name = os.path.basename(file_path)
                self.progress_update.emit(f"\n--- [ {i+1}/{total_videos} ] 开始处理: {video_name} ---")
                
                try:
                    result = self.analyze_single_local_video(file_path)
                    
                    if result and result.get('content'):
                        self.progress_update.emit(f"✅ 分析完成，正在保存...")
                        if self.save_single_result(result):
                            saved_count += 1
                            self.progress_update.emit(f"--- ✅ [ {i+1}/{total_videos} ] 保存成功 ---")

                            # 关键步骤：移动已处理的视频文件
                            try:
                                dest_path = os.path.join(completed_folder, video_name)
                                shutil.move(file_path, dest_path)
                                self.progress_update.emit(f"🚚 已将 '{video_name}' 移动到 '已分析分镜提示词' 文件夹。")
                            except Exception as e:
                                self.progress_update.emit(f"⚠️ 移动视频文件失败: {e}")
                        else:
                            self.progress_update.emit(f"--- ❌ [ {i+1}/{total_videos} ] 保存失败 ---\n")
                    else:
                        self.progress_update.emit(f"⚠️ 分析未返回有效结果。")

                except Exception as e:
                    self.error_occurred.emit(f"处理 '{video_name}' 时出错: {e}")
                    self.progress_update.emit("将尝试继续处理下一个视频...")

            self.progress_update.emit("--- ✅ 所有视频处理流程完毕 ---")
            self.analysis_complete.emit({'success': True, 'message': f'成功保存 {saved_count}/{total_videos} 个视频', 'results_count': saved_count})

        except Exception as e:
            self.error_occurred.emit(f"本地视频分析失败: {str(e)}")
        finally:
            self.cleanup_browser()

    def analyze_single_local_video(self, file_path):
        """在单个页面上分析本地视频"""
        try:
            self.progress_update.emit("正在导航到Gemini AI Studio...")
            self.page.goto("https://aistudio.google.com/prompts/new_chat", timeout=60000)
            self.page.wait_for_load_state("networkidle", timeout=60000)

            video_title = os.path.basename(file_path)
            self.progress_update.emit(f"正在分析: {video_title}")
            
            self.smart_delay()

            prompt_element = self.page.locator("//ms-chunk-input//textarea").first
            self.human_like_input(prompt_element, self.config['prompt'], "提示词")
            
            self.progress_update.emit("准备上传文件...")
            select_button = self.page.locator("//ms-add-chunk-menu//button/span[@class='mat-mdc-button-persistent-ripple mdc-icon-button__ripple']")
            self.human_like_click(select_button, "选择按钮")
            self.smart_delay()
            
            with self.page.expect_file_chooser() as fc_info:
                upload_button = self.page.locator("button:has-text('Upload')")
                self.human_like_click(upload_button, "Upload按钮")
            
            file_chooser = fc_info.value
            file_chooser.set_files(file_path)
            self.progress_update.emit("正在上传文件，请稍候...")

            # 4. 等待文件块出现在UI中，确认文件已添加
            self.progress_update.emit("确认文件添加中...")
            try:
                self.page.locator("//ms-video-chunk").first.wait_for(state="visible", timeout=30000)
                self.progress_update.emit("✅ 文件已在输入区显示。")
            except Exception:
                self.progress_update.emit("⚠️ 未检测到文件在输入区显示，但继续尝试...")

            # 5. 等待Run按钮变为可点击状态
            self.progress_update.emit("等待Run按钮激活...")
            run_button_selector = "//button[contains(@class, 'run-button') and @aria-disabled='false' and not(@disabled)]"
            run_button = self.page.locator(run_button_selector).first
            try:
                run_button.wait_for(state="visible", timeout=120000)
                self.progress_update.emit("✅ Run按钮已激活。")
            except Exception:
                self.progress_update.emit("⚠️ 等待Run按钮激活超时，但仍会尝试继续...")

            # 6. 点击run按钮
            self.human_like_click(run_button, "Run按钮")
            
            self.wait_for_analysis_completion()
            
            max_retries = 3
            retry_count = 0
            while retry_count < max_retries and self.check_generation_error():
                self.progress_update.emit(f"检测到生成错误，重试 ({retry_count + 1}/{max_retries})...")
                self.retry_generation()
                self.wait_for_analysis_completion()
                retry_count += 1
            
            if retry_count >= max_retries:
                self.progress_update.emit("达到最大重试次数，跳过。")
                return None

            time.sleep(2)
            result_content = self.get_analysis_result()
            
            if result_content:
                return {
                    'url': file_path,
                    'title': video_title,
                    'content': result_content,
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
            return None
        except Exception as e:
            self.progress_update.emit(f"分析本地视频时出错: {str(e)}")
            return None

    def start_browser(self):
        """通过比特浏览器API打开窗口并连接，只保留一个页面"""
        self.progress_update.emit("正在通过比特浏览器API启动窗口...")
        
        bit_window_id = self.config.get('bit_window_id')
        if not bit_window_id:
            raise ValueError("未提供比特浏览器窗口ID")

        # 比特浏览器本地API地址
        bit_api_url = "http://127.0.0.1:54345"
        headers = {'Content-Type': 'application/json'}

        # 1. 调用API打开浏览器窗口
        try:
            self.progress_update.emit(f"正在打开窗口ID: {bit_window_id}")
            open_url = f"{bit_api_url}/browser/open"
            open_data = {'id': bit_window_id}
            res = requests.post(open_url, data=json.dumps(open_data), headers=headers)
            res.raise_for_status() # 如果请求失败则抛出异常
            res_json = res.json()
            
            # 检查API调用是否成功
            if not res_json.get('success'):
                error_message = f"API打开窗口失败: {res_json.get('msg', '无详细错误信息')}. "
                error_message += f"完整API响应: {json.dumps(res_json, ensure_ascii=False)}"
                raise Exception(error_message)

            # 2. 从返回结果中获取CDP地址，键名是 'ws'
            cdp_address = res_json.get('data', {}).get('ws')
            if not cdp_address:
                raise Exception(f"API返回结果中未找到CDP地址 (ws). 完整响应: {json.dumps(res_json, ensure_ascii=False)}")
            
            self.progress_update.emit(f"成功获取CDP地址")

        except requests.exceptions.RequestException as e:
            self.error_occurred.emit(f"无法连接到比特浏览器API，请确认比特浏览器已启动并且API服务在运行中。错误: {e}")
            raise
        except Exception as e:
            self.error_occurred.emit(f"打开比特浏览器窗口时出错: {e}")
            raise

        # 3. 使用Playwright连接到获取的CDP地址
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.connect_over_cdp(cdp_address)
            self.context = self.browser.contexts[0]
            
            # 确保我们有一个干净的页面
            pages = self.context.pages
            if pages:
                self.page = pages[0] # 使用第一个已存在的页面
                # 关闭所有其他页面
                for i, p in enumerate(pages):
                    if i > 0:
                        p.close()
            else:
                self.page = self.context.new_page() # 如果没有页面则创建一个

            # 添加反机器人检测设置
            self.progress_update.emit("正在设置反机器人检测...")
            
            # 隐藏webdriver相关属性
            await_js_code = """
            // 删除webdriver属性
            delete navigator.webdriver;
            
            // 重写navigator.plugins属性
            Object.defineProperty(navigator, 'plugins', {
                get: () => [1, 2, 3, 4, 5].map(i => ({
                    name: `Plugin ${i}`,
                    description: `Plugin Description ${i}`,
                    filename: `plugin${i}.dll`,
                    length: 3
                }))
            });
            
            // 重写navigator.languages属性
            Object.defineProperty(navigator, 'languages', {
                get: () => ['zh-CN', 'zh', 'en-US', 'en']
            });
            
            // 重写navigator.permissions查询
            const originalQuery = window.navigator.permissions.query;
            window.navigator.permissions.query = (parameters) => (
                parameters.name === 'notifications' ?
                    Promise.resolve({ state: Notification.permission }) :
                    originalQuery(parameters)
            );
            
            // 重写chrome属性
            window.chrome = {
                runtime: {},
                loadTimes: function() { return {}; },
                csi: function() { return {}; },
                app: { isInstalled: false }
            };
            
            // 添加WebGL指纹伪装
            const getParameterProto = WebGLRenderingContext.prototype.getParameter;
            WebGLRenderingContext.prototype.getParameter = function(parameter) {
                // UNMASKED_VENDOR_WEBGL
                if (parameter === 37445) {
                    return 'Intel Inc.';
                }
                // UNMASKED_RENDERER_WEBGL
                if (parameter === 37446) {
                    return 'Intel Iris OpenGL Engine';
                }
                return getParameterProto.call(this, parameter);
            };
            
            // 添加Canvas指纹伪装
            const originalGetContext = HTMLCanvasElement.prototype.getContext;
            HTMLCanvasElement.prototype.getContext = function(contextType, contextAttributes) {
                const context = originalGetContext.call(this, contextType, contextAttributes);
                if (contextType === '2d') {
                    const originalFillText = context.fillText;
                    context.fillText = function() {
                        arguments[0] = arguments[0].toString();
                        return originalFillText.apply(this, arguments);
                    };
                    
                    const originalToDataURL = HTMLCanvasElement.prototype.toDataURL;
                    HTMLCanvasElement.prototype.toDataURL = function() {
                        // 添加微小噪点以改变指纹
                        const ctx = originalGetContext.call(this, '2d');
                        ctx.fillStyle = '#FFFFFF01';
                        ctx.fillRect(0, 0, 1, 1);
                        return originalToDataURL.apply(this, arguments);
                    };
                }
                return context;
            };
            
            // 模拟真实用户行为 - 鼠标移动跟踪
            window.mouseX = 0;
            window.mouseY = 0;
            document.addEventListener('mousemove', function(e) {
                window.mouseX = e.clientX;
                window.mouseY = e.clientY;
            });
            
            // 模拟真实用户行为 - 随机滚动
            let lastScrollTime = Date.now();
            document.addEventListener('scroll', function() {
                lastScrollTime = Date.now();
            });
            
            // 模拟真实用户行为 - 键盘事件
            document.addEventListener('keydown', function() {
                // 记录键盘活动
            });
            
            // 修改屏幕分辨率和颜色深度信息
            Object.defineProperty(screen, 'colorDepth', { value: 24 });
            Object.defineProperty(screen, 'pixelDepth', { value: 24 });
            
            // 修改硬件并发数
            Object.defineProperty(navigator, 'hardwareConcurrency', { value: 8 });
            
            // 修改设备内存
            Object.defineProperty(navigator, 'deviceMemory', { value: 8 });
            
            // 模拟电池API
            if (navigator.getBattery) {
                navigator.getBattery = function() {
                    return Promise.resolve({
                        charging: true,
                        chargingTime: 0,
                        dischargingTime: Infinity,
                        level: 1.0,
                        addEventListener: function() {}
                    });
                };
            }
            
            // 修改User-Agent客户端提示
            if (navigator.userAgentData) {
                Object.defineProperty(navigator, 'userAgentData', {
                    value: {
                        brands: [
                            {brand: 'Google Chrome', version: '119'},
                            {brand: 'Chromium', version: '119'},
                            {brand: 'Not=A?Brand', version: '24'}
                        ],
                        mobile: false,
                        platform: 'Windows'
                    }
                });
            }
            
            // 伪装已安装的扩展程序
            if (typeof chrome !== 'undefined' && chrome.runtime) {
                chrome.runtime.sendMessage = function() {
                    return Promise.resolve({success: false});
                };
            }
            
            // 修改AudioContext指纹
            const originalGetChannelData = AudioBuffer.prototype.getChannelData;
            if (originalGetChannelData) {
                AudioBuffer.prototype.getChannelData = function(channel) {
                    const array = originalGetChannelData.call(this, channel);
                    // 添加微小噪声
                    if (array.length > 0) {
                        array[0] = array[0] + 0.0000001;
                    }
                    return array;
                };
            }
            """
            
            try:
                # 添加初始化脚本，在每个页面加载时执行
                self.context.add_init_script(await_js_code)
                self.progress_update.emit("✅ 反机器人检测设置完成")
            except Exception as js_error:
                self.progress_update.emit(f"⚠️ 反机器人设置出现问题，但继续执行: {js_error}")

            self.progress_update.emit("✅ 成功连接到比特浏览器，并已清理无关页面")

        except Exception as e:
            self.error_occurred.emit(f"Playwright连接浏览器失败: {e}")
            raise
            
    def analyze_single_youtube_video(self, youtube_url, video_title=""):
        """在单个页面上分析YouTube视频，复用此页面"""
        try:
            # 1. 导航到目标网址
            self.progress_update.emit("正在导航到Gemini AI Studio...")
            self.page.goto("https://aistudio.google.com/prompts/new_chat", timeout=60000)
            self.page.wait_for_load_state("networkidle", timeout=60000)
            self.progress_update.emit("✅ 页面加载完成。")

            display_title = video_title if video_title else youtube_url
            self.progress_update.emit(f"正在分析: {display_title}")
            
            self.smart_delay()
            
            # 模拟用户行为 - 随机移动鼠标和轻微滚动
            try:
                viewport_size = self.page.viewport_size
                if viewport_size:
                    center_x = viewport_size['width'] // 2 + random.randint(-100, 100)
                    center_y = viewport_size['height'] // 2 + random.randint(-100, 100)
                    self.page.mouse.move(center_x, center_y)
                    time.sleep(0.5)
                    self.page.mouse.wheel(0, random.randint(-200, 200))
                    time.sleep(0.3)
            except Exception:
                pass # 用户行为模拟失败不影响主流程

            # 1. 在输入框中输入提示词
            try:
                prompt_element = self.page.locator("//ms-chunk-input//textarea").first
                prompt_element.wait_for(timeout=10000)
                self.human_like_input(prompt_element, self.config['prompt'], "提示词")
            except Exception as e:
                self.progress_update.emit(f"输入提示词失败: {str(e)}")
                raise e
            
            # 2. 点击选择按钮
            select_button = self.page.locator("//ms-add-chunk-menu//button/span[@class='mat-mdc-button-persistent-ripple mdc-icon-button__ripple']")
            self.human_like_click(select_button, "选择按钮")
            self.smart_delay()
            
            # 3. 点击YouTube按钮
            youtube_button = self.page.locator("//button[.//span[text()='YouTube Video']]")
            self.human_like_click(youtube_button, "YouTube按钮")
            self.smart_delay()
            
            # 4. 在弹出的输入框中填写网址
            url_input = self.page.locator("//input[@aria-label='YouTube URL']")
            self.human_like_input(url_input, youtube_url, "YouTube URL")
            
            # 5. 点击save按钮
            save_button = self.page.locator("//button[.//span[text()='Save']]")
            self.human_like_click(save_button, "Save按钮")
            time.sleep(self.delay_config['max_delay'])
            
            # 6. 等待Run按钮变为可点击状态
            self.progress_update.emit("等待Run按钮激活...")
            run_button_selector = "//button[contains(@class, 'run-button') and @aria-disabled='false' and not(@disabled)]"
            run_button = self.page.locator(run_button_selector).first
            try:
                run_button.wait_for(state="visible", timeout=60000)
                self.progress_update.emit("✅ Run按钮已激活。")
            except Exception as e:
                self.progress_update.emit(f"⚠️ 等待Run按钮激活超时: {e}，但仍会尝试继续...")

            # 7. 点击run按钮
            self.human_like_click(run_button, "Run按钮")
            
            self.smart_delay()
            
            # 8. 等待AI分析完成
            self.wait_for_analysis_completion()
            
            # 9. 检查是否生成成功，如果失败则重试
            max_retries = 3
            retry_count = 0
            
            while retry_count < max_retries:
                if self.check_generation_error():
                    self.progress_update.emit(f"检测到生成错误，重试 ({retry_count + 1}/{max_retries})...")
                    time.sleep(random.uniform(2, 4))
                    self.retry_generation()
                    self.wait_for_analysis_completion()
                    retry_count += 1
                else:
                    break
            
            if retry_count >= max_retries:
                self.progress_update.emit("达到最大重试次数，跳过。")
                return None
            
            # 10. 获取分析结果
            time.sleep(2)
            result_content = self.get_analysis_result()
            
            if result_content:
                return {
                    'url': youtube_url,
                    'title': video_title,
                    'content': result_content,
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
            
            return None
            
        except Exception as e:
            self.progress_update.emit(f"分析视频时出错: {str(e)}")
            return None
    
    def wait_for_analysis_completion(self):
        """等待AI分析完成"""
        max_wait_time = 300  # 最多等待5分钟
        start_time = time.time()
        
        while time.time() - start_time < max_wait_time:
            try:
                # 检测stop按钮是否消失
                stop_button = self.page.locator("//run-button/button/div[.//text()[contains(., 'Stop')]]")
                if not stop_button.is_visible():
                    self.progress_update.emit("AI分析完成")
                    return True
                time.sleep(2)
            except:
                # 如果找不到stop按钮，可能已经完成
                return True
        
        self.progress_update.emit("等待超时")
        return False
    
    def check_generation_error(self):
        """检查是否生成失败"""
        try:
            error_element = self.page.locator("(//ms-chat-turn)[last()]//ms-prompt-feedback/button/span[1]")
            return error_element.is_visible()
        except:
            return False
    
    def retry_generation(self):
        """重新生成"""
        try:
            # 在输入框输入重试提示词
            prompt_textarea = self.page.locator("//ms-chunk-input//textarea")
            prompt_textarea.fill("按照要求输出完整分镜提示词")
            time.sleep(random.uniform(1, 2)) # 增加延时
            
            # 点击run按钮 - 使用新的选择器
            run_button_selector = "//button[contains(@class, 'run-button') and @aria-disabled='false' and not(@disabled)]"
            run_button = self.page.locator(run_button_selector).first
            self.human_like_click(run_button, "Run按钮(重试)")
            time.sleep(random.uniform(1, 2)) # 增加延时
            
        except Exception as e:
            self.progress_update.emit(f"重试时出错: {str(e)}")
    
    def get_analysis_result(self):
        """获取分析结果，专门提取表格内容"""
        try:
            self.progress_update.emit("正在获取分析结果...")
            
            # 等待内容完全加载
            time.sleep(2)
            
            # 首先尝试直接获取HTML表格内容并解析
            table_content = None
            table_data = []
            
            # 方法1: 直接解析HTML表格结构
            try:
                self.progress_update.emit("尝试解析HTML表格结构...")
                
                # 检查页面中是否存在表格元素
                table_exists = self.page.evaluate("""() => {
                    const tables = document.querySelectorAll('table');
                    return tables.length > 0;
                }""")
                
                if table_exists:
                    self.progress_update.emit("✅ 检测到HTML表格元素")
                    
                    # 直接从HTML中提取表格数据
                    table_data = self.page.evaluate("""() => {
                        const tables = document.querySelectorAll('table');
                        const table = tables[tables.length - 1]; // 使用最后一个表格
                        
                        const data = [];
                        const rows = table.querySelectorAll('tr');
                        
                        // 获取表头
                        const headers = [];
                        const headerCells = rows[0].querySelectorAll('th, td');
                        for (let cell of headerCells) {
                            headers.push(cell.innerText.trim());
                        }
                        
                        // 确定列索引
                        let shotColIdx = -1;
                        let keyframeColIdx = -1;
                        let videoColIdx = -1;
                        
                        headers.forEach((header, idx) => {
                            const headerLower = header.toLowerCase();
                            if (headerLower.includes('分镜')) {
                                shotColIdx = idx;
                            } else if (headerLower.includes('关键帧') || headerLower.includes('图片生成')) {
                                keyframeColIdx = idx;
                            } else if (headerLower.includes('视频') || headerLower.includes('图生视频')) {
                                videoColIdx = idx;
                            }
                        });
                        
                        // 使用默认索引如果无法识别
                        if (shotColIdx === -1 && headers.length > 0) shotColIdx = 0;
                        if (keyframeColIdx === -1 && headers.length > 1) keyframeColIdx = 1;
                        if (videoColIdx === -1 && headers.length > 2) videoColIdx = 2;
                        
                        // 获取数据行
                        for (let i = 1; i < rows.length; i++) {
                            const cells = rows[i].querySelectorAll('td');
                            if (cells.length >= Math.max(shotColIdx, keyframeColIdx, videoColIdx) + 1) {
                                // 提取分镜号
                                let shotNumber = i;
                                const shotText = cells[shotColIdx].innerText.trim();
                                const shotMatch = shotText.match(/\d+/);
                                if (shotMatch) {
                                    shotNumber = parseInt(shotMatch[0]);
                                }
                                
                                const keyframeText = cells[keyframeColIdx].innerText.trim();
                                const videoText = videoColIdx >= 0 && videoColIdx < cells.length ? 
                                                  cells[videoColIdx].innerText.trim() : "";
                                
                                data.push([shotNumber, keyframeText, videoText]);
                            }
                        }
                        
                        return data;
                    }""")
                    
                    if table_data and len(table_data) > 0:
                        self.progress_update.emit(f"✅ 成功通过HTML表格解析获取 {len(table_data)} 行数据")
                        
                        # 构建表格文本内容用于备份
                        headers = ["分镜", "关键帧图片生成提示词", "图生视频提示词"]
                        table_content = "\t".join(headers) + "\n"
                        
                        for row in table_data:
                            shot_num, keyframe, video = row
                            table_content += f"分镜{shot_num}\t{keyframe}\t{video}\n"
                        
                        self.progress_update.emit(f"✅ 成功通过HTML表格解析获取 {len(table_data)} 行数据")
                        return table_content
                    else:
                        self.progress_update.emit("⚠️ HTML表格解析未获取到有效数据")
                else:
                    self.progress_update.emit("⚠️ 页面中未检测到HTML表格元素")
            except Exception as e:
                self.progress_update.emit(f"⚠️ HTML表格解析失败: {str(e)}")
            
            # 如果HTML表格解析失败，回退到原来的方法
            
            # 方法2: 获取表格的文本内容
            try:
                table_container = self.page.locator(".table-container table")
                if table_container.is_visible():
                    table_content = table_container.inner_text()
                    self.progress_update.emit("✅ 成功通过table-container获取表格内容")
                else:
                    self.progress_update.emit("⚠️ table-container不可见")
            except Exception as e:
                self.progress_update.emit(f"⚠️ 方法2失败: {str(e)}")
            
            # 方法3: 如果方法2失败，尝试获取整个聊天回复内容
            if not table_content:
                try:
                    # 获取最后一个模型回复的完整内容
                    chat_turn = self.page.locator("div.chat-turn-container.model.render").last
                    if chat_turn.is_visible():
                        table_content = chat_turn.inner_text()
                        self.progress_update.emit("✅ 成功通过chat-turn-container获取内容")
                    else:
                        self.progress_update.emit("⚠️ chat-turn-container不可见")
                except Exception as e:
                    self.progress_update.emit(f"⚠️ 方法3失败: {str(e)}")
            
            # 方法4: 如果前面都失败，尝试获取所有表格相关内容
            if not table_content:
                try:
                    # 查找包含"分镜"的元素
                    elements_with_fengjing = self.page.locator("text=分镜")
                    if elements_with_fengjing.count() > 0:
                        # 获取包含分镜内容的父容器
                        parent_container = elements_with_fengjing.first.locator("xpath=ancestor::div[contains(@class,'turn-content') or contains(@class,'model-prompt-container')]")
                        if parent_container.is_visible():
                            table_content = parent_container.inner_text()
                            self.progress_update.emit("✅ 成功通过分镜元素定位获取内容")
                except Exception as e:
                    self.progress_update.emit(f"⚠️ 方法4失败: {str(e)}")
            
            # 如果所有方法都失败，记录页面状态
            if not table_content:
                try:
                    page_content = self.page.content()
                    self.progress_update.emit(f"⚠️ 所有方法都失败，页面标题: {self.page.title()}")
                    
                    # 检查是否有表格元素
                    table_count = self.page.locator("table").count()
                    self.progress_update.emit(f"页面中表格元素数量: {table_count}")
                    
                    # 检查是否有分镜相关文本
                    fengjing_count = self.page.locator("text=分镜").count()
                    self.progress_update.emit(f"页面中'分镜'文本数量: {fengjing_count}")
                    
                    return None
                except Exception as e:
                    self.progress_update.emit(f"获取页面状态时出错: {str(e)}")
                    return None
            
            # 验证获取的内容
            if table_content:
                self.progress_update.emit(f"获取到内容长度: {len(table_content)} 字符")
                self.progress_update.emit(f"内容预览 (前300字符): {table_content[:300]}...")
                
                # 检查内容是否包含预期的表格结构
                if "分镜" in table_content and ("关键帧" in table_content or "提示词" in table_content):
                    self.progress_update.emit("✅ 内容验证通过，包含预期的表格结构")
                    return table_content
                else:
                    self.progress_update.emit("⚠️ 内容验证失败，未找到预期的表格结构")
                    self.progress_update.emit(f"完整内容: {table_content}")
                    return table_content  # 仍然返回内容，让后续处理判断
            
            return None
            
        except Exception as e:
            self.progress_update.emit(f"获取结果时出错: {str(e)}")
            return None
    
    def save_single_result(self, result):
        """保存单个分析结果"""
        if not result:
            return False
        
        output_path = self.config['output_path']

        try:
            file_name = result.get('title', f"YouTube_Analysis_{result.get('timestamp', '')}")
            
            content = result.get('content')
            if not content:
                return False
                
            processed_result = self.process_text(
                output_path, 
                content, 
                file_name
            )
            
            return processed_result and processed_result.get('success')
                
        except Exception as e:
            self.progress_update.emit(f"❌ 保存结果时发生严重错误: {str(e)}")
            return False

    def cleanup_browser(self):
        """关闭比特浏览器窗口并清理资源"""
        try:
            # 1. 清理Playwright资源
            if hasattr(self, 'playwright') and self.playwright:
                self.playwright.stop()
            self.progress_update.emit("Playwright会话已断开")

            # # 2. 通过API关闭浏览器窗口 (根据用户要求，暂时注释掉)
            # bit_window_id = self.config.get('bit_window_id')
            # if bit_window_id:
            #     try:
            #         bit_api_url = "http://127.0.0.1:54345"
            #         headers = {'Content-Type': 'application/json'}
            #         close_url = f"{bit_api_url}/browser/close"
            #         close_data = {'id': bit_window_id}
            #         requests.post(close_url, data=json.dumps(close_data), headers=headers)
            #         self.progress_update.emit(f"已通过API关闭窗口ID: {bit_window_id}")
            #     except Exception as e:
            #         self.progress_update.emit(f"通过API关闭窗口时出错: {e}")

        except Exception as e:
            self.progress_update.emit(f"清理资源时出错: {str(e)}")
    
    # 以下是您提供的文本处理函数
    def sanitize_filename(self, filename_str, max_length=100):
        """清理并规范化文件名"""
        if not filename_str or not isinstance(filename_str, str):
            return f"invalid_filename_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        # 删除#号及其后面的内容
        name = filename_str.split('#')[0].strip()
        
        # 删除常见的YouTube视频标题后缀
        patterns_to_remove = [
            r'\s*\|\s*.*$',                # 删除 | 及其后面的内容
            r'\s*-\s*YouTube\s*$',         # 删除 - YouTube 后缀
            r'\s*\(\d{4}\)\s*$',           # 删除年份 (2023) 等
            r'\s*\[[^\]]+\]\s*$',          # 删除方括号内容 [HD] 等
            r'\s*\{[^}]+\}\s*$',           # 删除花括号内容
            r'\s*【[^】]+】\s*$',           # 删除中文方括号内容
            r'\s*「[^」]+」\s*$',           # 删除中文引号内容
        ]
        
        for pattern in patterns_to_remove:
            name = re.sub(pattern, '', name)

        try:
            normalized_name = unicodedata.normalize('NFKD', name)
            processed_name = "".join([c for c in normalized_name if not unicodedata.combining(c)])
        except TypeError:
            processed_name = name

        # 替换文件系统不允许的字符
        illegal_chars_pattern = r'[?%*:|"<>\x00-\x1f]'
        name = re.sub(illegal_chars_pattern, '', processed_name)
        
        # 替换斜杠为空格
        name = re.sub(r'[/\\]+', ' ', name)
        
        # 合并多个空格为单个空格
        name = re.sub(r'\s+', ' ', name).strip()

        # 删除结尾的点号
        if name.endswith('.'):
            name = name[:-1].strip()
        
        # 删除开头的点号（避免隐藏文件）
        if name.startswith('.'):
            name = name[1:].strip()

        # 截断过长的文件名
        name = name[:max_length].strip()

        # 如果处理后文件名为空，使用默认名称
        if not name or name.isspace():
            return f"video_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
        return name

    def clean_text_content(self, text_content):
        """清理文本内容，移除不必要的标记和行"""
        patterns_to_remove = [
            r'^edit\s*$', 
            r'^more_vert\s*$',
            r'^thumb_up\s*$', 
            r'^thumb_down\s*$',
            r'^content_copy\s*$',
            r'^download\s*$',
            r'Use code with caution\.\s*$',
            r'\d+\.\d+s\s*$'
        ]
        
        lines = text_content.split('\n')
        processed_lines = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            should_remove = False
            for pattern in patterns_to_remove:
                if re.match(pattern, line, re.IGNORECASE):
                    should_remove = True
                    break
            
            if not should_remove:
                processed_lines.append(line)
        
        return "\n".join(processed_lines)

    def parse_tab_separated_table(self, text_content):
        """解析制表符或多空格分隔的表格，兼容HTML表格提取的文本"""
        try:
            cleaned_text = self.clean_text_content(text_content)
            lines = cleaned_text.split('\n')
            
            header_line = None
            header_index = -1
            
            for i, line in enumerate(lines):
                line_lower = line.lower()
                if ('分镜' in line and ('关键帧' in line or '提示词' in line or '图片生成' in line or '视频' in line)) or \
                   (line.count('分镜') > 0 and line.count('提示词') > 0):
                    header_line = line
                    header_index = i
                    break
            
            if not header_line:
                for i, line in enumerate(lines):
                    if re.search(r'分镜\s*\d+', line):
                        header_line = "分镜\t关键帧图片生成提示词\t图生视频提示词"
                        header_index = i - 1
                        break
                
                if not header_line:
                    return []
            
            if '\t' in header_line:
                headers = header_line.split('\t')
                separator_type = "tab"
            else:
                headers = re.split(r'\s{2,}', header_line)
                if len(headers) < 3:
                    headers = header_line.split()
                separator_type = "space"
            
            headers = [h.strip() for h in headers if h.strip()]
            
            shot_col_idx = keyframe_col_idx = video_col_idx = -1
            
            for i, header in enumerate(headers):
                header_lower = header.lower().replace(' ', '')
                if '分镜' in header_lower:
                    shot_col_idx = i
                elif '关键帧' in header_lower or '图片生成' in header_lower:
                    keyframe_col_idx = i
                elif '图生视频' in header_lower or ('视频' in header_lower and '图生' in header_lower):
                    video_col_idx = i
            
            if shot_col_idx == -1 and len(headers) > 0: shot_col_idx = 0
            if keyframe_col_idx == -1 and len(headers) > 1: keyframe_col_idx = 1
            if video_col_idx == -1 and len(headers) > 2: video_col_idx = 2
            
            results = []
            data_start_index = max(0, header_index + 1)
            
            for i in range(data_start_index, len(lines)):
                line = lines[i].strip()
                if not line:
                    continue
                
                if line.lower() in ['edit', 'more_vert', 'thumb_up', 'thumb_down'] or \
                   re.match(r'^\d+\.\d+s$', line):
                    continue
                
                if separator_type == "tab":
                    cells = line.split('\t')
                else:
                    cells = re.split(r'\s{2,}', line)
                    if len(cells) < 3:
                        match = re.match(r'(分镜\d+)\s+(.+)', line)
                        if match:
                            shot_part = match.group(1)
                            rest_content = match.group(2)
                            
                            cells = [shot_part]
                            
                            split_patterns = [r'。\s*(?=[电影感镜头|成年|白色|男人])', r'\.\s+', r'；\s*']
                            split_found = False
                            
                            for pattern in split_patterns:
                                parts = re.split(pattern, rest_content, 1)
                                if len(parts) == 2:
                                    cells.extend([parts[0].strip(), parts[1].strip()])
                                    split_found = True
                                    break
                            
                            if not split_found:
                                if len(rest_content) > 100:
                                    mid_point = len(rest_content) // 2
                                    cells.extend([rest_content[:mid_point].strip(), rest_content[mid_point:].strip()])
                                else:
                                    cells.extend([rest_content, ""])
                        else:
                            cells = [line]
                
                cells = [c.strip() for c in cells]
                
                while len(cells) <= max(shot_col_idx, keyframe_col_idx, video_col_idx):
                    cells.append("")
                
                shot_number = i - data_start_index + 1
                if shot_col_idx >= 0 and shot_col_idx < len(cells):
                    shot_text = cells[shot_col_idx]
                    shot_match = re.search(r'(\d+)', shot_text)
                    if shot_match:
                        shot_number = int(shot_match.group(1))
                
                keyframe_prompt = cells[keyframe_col_idx] if keyframe_col_idx >= 0 and keyframe_col_idx < len(cells) else ""
                video_prompt = cells[video_col_idx] if video_col_idx >= 0 and video_col_idx < len(cells) else ""
                
                if keyframe_prompt.strip() or video_prompt.strip():
                    results.append((shot_number, keyframe_prompt.strip(), video_prompt.strip()))
            
            return results
            
        except Exception as e:
            self.progress_update.emit(f"❌ 表格解析错误: {e}")
            return []

    def process_text(self, folder_path, text_content, file_name=None):
        """处理文本并保存到Excel"""
        try:
            os.makedirs(folder_path, exist_ok=True)

            table_data = self.parse_tab_separated_table(text_content)
            
            if not table_data:
                self.progress_update.emit(f"警告: 未能从 '{file_name}' 的分析结果中解析出有效数据。")
                return {"success": False, "message": "No valid storyboard data found."}
            
            df_data = []
            for shot_num, keyframe, video in table_data:
                df_data.append({
                    '分镜': f'分镜{shot_num}',
                    '关键帧图片生成提示词': keyframe,
                    '图生视频提示词': video
                })
            
            df = pd.DataFrame(df_data)
            
            if file_name:
                sanitized_base_name = self.sanitize_filename(file_name)
            else:
                sanitized_base_name = f"分析结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            subfolder_path = os.path.join(folder_path, sanitized_base_name)
            os.makedirs(subfolder_path, exist_ok=True)

            excel_file_path = os.path.join(subfolder_path, f"{sanitized_base_name}.xlsx")
            
            try:
                with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='分镜表', index=False)
                    worksheet = writer.sheets['分镜表']
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max(max_length + 2, 20), 100)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

                if os.path.exists(excel_file_path):
                    self.progress_update.emit(f"文件已保存到: {subfolder_path}")
                    return {"success": True, "output_file": excel_file_path}
                else:
                    return {"success": False, "message": "Excel file not found after save."}

            except Exception as e:
                self.progress_update.emit(f"❌ 保存Excel文件时发生错误: {str(e)}")
                return {"success": False, "message": f"Failed to save Excel: {e}"}

        except Exception as e:
            self.progress_update.emit(f"❌ 处理和保存文本时发生严重错误: {e}")
            return {"success": False, "message": f"Processing failed: {str(e)}"}

    def smart_delay(self, delay_type='base'):
        """智能延时功能"""
        min_delay = self.delay_config.get('min_delay', 1)
        max_delay = self.delay_config.get('max_delay', 3)
        
        # 在最小和最大延时之间随机选择
        delay = random.uniform(min_delay, max_delay)
        # 记录实际延时值（调试用）
        self.progress_update.emit(f"智能延时 {delay:.1f}s (范围: {min_delay}-{max_delay}秒)")
        time.sleep(delay)
    
    def human_like_input(self, element, text, description=""):
        """人性化输入函数"""
        try:
            if description:
                self.progress_update.emit(f"正在输入{description}...")
            
            # 点击元素获得焦点
            element.click()
            self.smart_delay()
            
            # 清空并一次性输入
            element.fill("")
            time.sleep(0.1)
            element.fill(text)
            
            self.smart_delay()
            
        except Exception as e:
            self.progress_update.emit(f"输入时出错: {str(e)}")
            raise 

    def human_like_click(self, selector_or_element, description="元素"):
        """人性化点击函数，模拟真实用户行为"""
        try:
            self.progress_update.emit(f"正在点击{description}...")
            
            # 确定元素
            element = selector_or_element
            if isinstance(selector_or_element, str):
                element = self.page.locator(selector_or_element).first
            
            # 确保元素可见
            if not element.is_visible():
                self.progress_update.emit(f"⚠️ {description}不可见")
                return False
                
            # 获取元素位置和大小
            bbox = element.bounding_box()
            if not bbox:
                self.progress_update.emit(f"⚠️ 无法获取{description}的位置")
                return False
                
            # 计算目标点，添加随机偏移使点击更自然
            target_x = bbox['x'] + bbox['width'] * random.uniform(0.3, 0.7)
            target_y = bbox['y'] + bbox['height'] * random.uniform(0.3, 0.7)
            
            # 获取当前鼠标位置
            current_position = self.page.evaluate("""() => {
                return {x: window.mouseX || 0, y: window.mouseY || 0}
            }""")
            
            current_x = current_position.get('x', 0)
            current_y = current_position.get('y', 0)
            
            # 如果没有当前位置，使用视口中的随机位置
            if current_x == 0 and current_y == 0:
                viewport = self.page.viewport_size
                if viewport:
                    current_x = random.randint(0, viewport['width'])
                    current_y = random.randint(0, viewport['height'])
            
            # 生成自然的移动轨迹点
            points = self.generate_natural_curve(current_x, current_y, target_x, target_y)
            
            # 执行鼠标移动
            for point in points:
                self.page.mouse.move(point[0], point[1])
                # 随机微小停顿
                if random.random() < 0.2:  # 20%概率停顿
                    time.sleep(random.uniform(0.01, 0.05))
            
            # 鼠标悬停在元素上
            time.sleep(random.uniform(0.1, 0.3))
            
            # 偶尔添加微小抖动
            if random.random() < 0.3:  # 30%概率抖动
                for _ in range(2):
                    jitter_x = target_x + random.uniform(-2, 2)
                    jitter_y = target_y + random.uniform(-2, 2)
                    self.page.mouse.move(jitter_x, jitter_y)
                    time.sleep(random.uniform(0.01, 0.03))
                # 移回目标位置
                self.page.mouse.move(target_x, target_y)
                time.sleep(random.uniform(0.05, 0.1))
            
            # 执行点击，模拟按下和释放的时间差
            self.page.mouse.down()
            time.sleep(random.uniform(0.05, 0.15))  # 人类点击通常按住0.05-0.15秒
            self.page.mouse.up()
            
            # 点击后的随机短暂停顿
            time.sleep(random.uniform(0.1, 0.2))
            
            self.progress_update.emit(f"✅ 成功点击{description}")
            return True
            
        except Exception as e:
            self.progress_update.emit(f"❌ 点击{description}时出错: {str(e)}")
            return False
    
    def generate_natural_curve(self, start_x, start_y, end_x, end_y, control_points=3):
        """生成自然的鼠标移动轨迹曲线"""
        points = []
        # 添加起点
        points.append((start_x, start_y))
        
        # 计算直线距离
        distance = math.sqrt((end_x - start_x) ** 2 + (end_y - start_y) ** 2)
        
        # 根据距离确定控制点数量和总点数
        if distance < 100:
            control_points = 1
            total_points = 5
        elif distance < 300:
            control_points = 2
            total_points = 10
        else:
            control_points = 3
            total_points = 15
            
        # 生成控制点
        control_xs = []
        control_ys = []
        
        for i in range(control_points):
            # 控制点在起点和终点连线附近随机偏移
            ratio = (i + 1) / (control_points + 1)
            control_x = start_x + (end_x - start_x) * ratio
            control_y = start_y + (end_y - start_y) * ratio
            
            # 添加随机偏移，距离越远偏移越大
            max_offset = min(100, distance * 0.15)
            offset_x = random.uniform(-max_offset, max_offset)
            offset_y = random.uniform(-max_offset, max_offset)
            
            control_xs.append(control_x + offset_x)
            control_ys.append(control_y + offset_y)
        
        # 生成贝塞尔曲线点
        for i in range(1, total_points):
            t = i / total_points
            
            # 简化的贝塞尔曲线计算
            x = start_x
            y = start_y
            
            # 线性插值所有控制点
            for j in range(control_points):
                x += (control_xs[j] - x) * t
                y += (control_ys[j] - y) * t
            
            # 最后插值到终点
            x += (end_x - x) * t
            y += (end_y - y) * t
            
            points.append((x, y))
        
        # 添加终点
        points.append((end_x, end_y))
        return points 