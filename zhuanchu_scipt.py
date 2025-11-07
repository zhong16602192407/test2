import pandas as pd
import os
import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import base64
import json
from urllib.parse import urlparse


def sanitize_filename(filename):
    """清理文件名，移除不合法字符"""
    # 移除Windows文件名不允许的字符
    invalid_chars = r'[<>:"/\\|?*]'
    filename = re.sub(invalid_chars, '_', filename)
    # 限制文件名长度
    if len(filename) > 200:
        filename = filename[:200]
    return filename


def get_domain_type(url):
    """
    识别网站类型
    :param url: 网页URL
    :return: 网站类型字符串
    """
    try:
        parsed = urlparse(url)
        domain = parsed.netloc.lower()

        if 'mdpi.com' in domain:
            return 'mdpi'
        elif 'globalbiodefense.com' in domain:
            return 'globalbiodefense'
        elif 'imec' in domain:  # imec.be 或 imec-int.com
            return 'imec'
        elif 'plos.org' in domain:
            return 'plos'
        else:
            return 'other'
    except:
        return 'other'


def get_content_cleanup_script(domain_type):
    """
    根据网站类型生成内容清理脚本
    :param domain_type: 网站类型
    :return: JavaScript清理脚本
    """
    scripts = {
        'mdpi': """
            // MDPI: 只保留 middle-column 内容
            var mainContent = document.querySelector('#middle-column');
            if (mainContent) {
                // 移除 middle-column__help 导航栏
                var helpSection = mainContent.querySelector('.middle-column__help');
                if (helpSection) {
                    helpSection.remove();
                }

                // 保存主内容
                var contentHTML = mainContent.innerHTML;

                // 清空body并只添加主内容
                document.body.innerHTML = '<div id="middle-column" class="content__column">' + contentHTML + '</div>';

                // 添加基本样式
                var style = document.createElement('style');
                style.textContent = 'body { padding: 20px; font-family: Arial, sans-serif; } #middle-column { max-width: 900px; margin: 0 auto; }';
                document.head.appendChild(style);
            }
        """,

        'globalbiodefense': """
            // Global Biodefense: 只保留 main-content
            var mainContent = document.querySelector('.col-8.main-content.s-post-contain');
            if (mainContent) {
                var contentHTML = mainContent.innerHTML;
                document.body.innerHTML = '<div class="main-content">' + contentHTML + '</div>';

                var style = document.createElement('style');
                style.textContent = 'body { padding: 20px; font-family: Arial, sans-serif; } .main-content { max-width: 900px; margin: 0 auto; }';
                document.head.appendChild(style);
            }
        """,

        'imec': """
            // IMEC: 只保留 layout-content
            var mainContent = document.querySelector('.layout_layout-content___o3j2');
            if (mainContent) {
                var contentHTML = mainContent.innerHTML;
                document.body.innerHTML = '<div class="layout-content">' + contentHTML + '</div>';

                var style = document.createElement('style');
                style.textContent = 'body { padding: 20px; font-family: Arial, sans-serif; } .layout-content { max-width: 1200px; margin: 0 auto; }';
                document.head.appendChild(style);
            }
        """,

        'plos': """
            // PLOS: 只保留 main-content
            var mainContent = document.querySelector('#main-content');
            if (mainContent) {
                var contentHTML = mainContent.innerHTML;
                document.body.innerHTML = '<main id="main-content">' + contentHTML + '</main>';

                var style = document.createElement('style');
                style.textContent = 'body { padding: 20px; font-family: Arial, sans-serif; } #main-content { max-width: 900px; margin: 0 auto; }';
                document.head.appendChild(style);
            }
        """,

        'other': """
            // 其他网站：移除常见的导航栏、侧边栏、页脚等
            var elementsToRemove = [
                'header', 'nav', 'aside', 'footer',
                '.header', '.navigation', '.sidebar', '.footer',
                '.advertisement', '.ad', '.banner'
            ];

            elementsToRemove.forEach(function(selector) {
                var elements = document.querySelectorAll(selector);
                elements.forEach(function(el) {
                    // 只移除明显的导航和广告元素，保留文章内的这些标签
                    if (el.offsetHeight < 200 || selector.includes('ad') || selector.includes('banner')) {
                        el.remove();
                    }
                });
            });
        """
    }

    return scripts.get(domain_type, scripts['other'])


def setup_driver(download_dir, proxy_settings=None):
    """
    配置Chrome浏览器
    :param download_dir: 下载目录
    :param proxy_settings: 代理设置，格式: {'proxy_type': 'http', 'host': '127.0.0.1', 'port': '1080'}
    """
    chrome_options = Options()

    # 无头模式（不显示浏览器窗口）
    # chrome_options.add_argument('--headless')  # 如果需要看到浏览器运行过程，可以注释这行

    # 其他配置
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')

    # 设置打印参数
    chrome_options.add_argument('--kiosk-printing')

    # 添加代理设置
    if proxy_settings:
        proxy_str = f"{proxy_settings['host']}:{proxy_settings['port']}"

        if proxy_settings.get('proxy_type') == 'socks5':
            chrome_options.add_argument(f'--proxy-server=socks5://{proxy_str}')
        elif proxy_settings.get('proxy_type') == 'socks4':
            chrome_options.add_argument(f'--proxy-server=socks4://{proxy_str}')
        else:  # 默认HTTP代理
            chrome_options.add_argument(f'--proxy-server=http://{proxy_str}')

        print(f"已设置代理: {proxy_settings['proxy_type']}://{proxy_str}")

    # 可选：忽略证书错误（如果代理有证书问题）
    chrome_options.add_argument('--ignore-certificate-errors')

    # 创建下载目录
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    driver = webdriver.Chrome(options=chrome_options)
    return driver


def save_page_as_pdf(driver, url, output_path, wait_time=5, max_retries=2):
    """
    将网页保存为PDF
    :param driver: WebDriver实例
    :param url: 网页URL
    :param output_path: 输出PDF路径
    :param wait_time: 页面加载等待时间（秒）
    :param max_retries: 最大重试次数
    :return: 是否成功
    """
    for attempt in range(max_retries):
        try:
            print(f"  正在访问: {url}")
            if attempt > 0:
                print(f"  第 {attempt + 1} 次重试...")

            driver.get(url)

            # 等待页面加载
            time.sleep(wait_time)

            # 等待body元素加载
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # 额外等待图片和其他资源加载
            driver.execute_script("return document.readyState") == "complete"
            time.sleep(3)  # 增加等待时间确保国外网站加载完成

            # 清理页面内容 - 只保留正文
            domain_type = get_domain_type(url)
            cleanup_script = get_content_cleanup_script(domain_type)

            if domain_type != 'other':
                print(f"  检测到网站类型: {domain_type}，正在清理内容...")

            try:
                driver.execute_script(cleanup_script)
                time.sleep(1)  # 等待DOM更新
                print(f"  内容清理完成")
            except Exception as e:
                print(f"  内容清理警告: {str(e)}")
                # 继续执行，即使清理失败

            # 使用Chrome的打印功能生成PDF
            print_options = {
                'landscape': False,
                'displayHeaderFooter': False,
                'printBackground': True,
                'preferCSSPageSize': True,
                'paperWidth': 8.27,  # A4宽度（英寸）
                'paperHeight': 11.69,  # A4高度（英寸）
                'marginTop': 0.4,
                'marginBottom': 0.4,
                'marginLeft': 0.4,
                'marginRight': 0.4,
            }

            result = driver.execute_cdp_cmd("Page.printToPDF", print_options)

            # 保存PDF
            with open(output_path, 'wb') as f:
                f.write(base64.b64decode(result['data']))

            print(f"  ✓ 转换成功")
            return True

        except Exception as e:
            print(f"  ✗ 转换失败 (尝试 {attempt + 1}/{max_retries})")
            print(f"  错误信息: {str(e)}")

            if attempt < max_retries - 1:
                print("  等待5秒后重试...")
                time.sleep(5)
            else:
                return False


def main(excel_file, output_dir='html_pdfs', url_column_name='来源网址', wait_time=8, proxy_settings=None):
    """
    主函数
    :param excel_file: Excel文件路径
    :param output_dir: 输出目录
    :param url_column_name: URL列的列名
    :param wait_time: 每个页面的等待加载时间（秒），访问国外网站建议增加
    :param proxy_settings: 代理设置
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)

        print(f"读取Excel文件: {excel_file}")
        print(f"共 {len(df)} 行数据\n")

        # 获取列名
        columns = df.columns.tolist()
        print(f"Excel列名: {columns}\n")

        # 假设第一列是序号，第三列是标题
        col_index = columns[0]  # 第一列
        col_title = columns[2]  # 第三列（正文标题）

        # 查找URL列
        if url_column_name not in columns:
            print(f"错误: 找不到列 '{url_column_name}'")
            print(f"可用的列名: {columns}")
            return

        col_url = url_column_name

        # 初始化浏览器（带代理）
        print("正在启动浏览器...\n")
        driver = setup_driver(output_dir, proxy_settings)

        success_count = 0
        fail_count = 0

        try:
            # 遍历每一行
            for idx, row in df.iterrows():
                序号 = str(row[col_index])
                标题 = str(row[col_title])
                网址 = row[col_url]

                # 检查URL是否有效
                if pd.isna(网址) or not str(网址).startswith('http'):
                    print(f"行 {idx + 1}: 序号[{序号}] - 无效的URL，跳过\n")
                    continue

                print(f"行 {idx + 1}: 序号[{序号}] - 标题[{标题}]")

                # 构建文件名
                filename = f"{序号}-{标题}.pdf"
                filename = sanitize_filename(filename)
                output_path = os.path.join(output_dir, filename)

                # 转换为PDF
                if save_page_as_pdf(driver, str(网址), output_path, wait_time):
                    success_count += 1
                else:
                    fail_count += 1

                print()  # 空行分隔

                # 添加延迟，避免请求过快
                time.sleep(2)

        finally:
            # 关闭浏览器
            driver.quit()
            print("浏览器已关闭\n")

        # 打印统计信息
        print(f"{'=' * 50}")
        print(f"转换完成!")
        print(f"成功: {success_count} 个")
        print(f"失败: {fail_count} 个")
        print(f"文件保存在: {output_dir}")

    except Exception as e:
        print(f"错误: {str(e)}")
        import traceback
        traceback.print_exc()


# 代理配置示例
PROXY_CONFIGS = {
    # Clash 默认配置
    'clash_http': {
        'proxy_type': 'http',
        'host': '127.0.0.1',
        'port': '7890'
    },
    'clash_socks5': {
        'proxy_type': 'socks5',
        'host': '127.0.0.1',
        'port': '7890'
    },
    # V2Ray 默认配置
    'v2ray_http': {
        'proxy_type': 'http',
        'host': '127.0.0.1',
        'port': '10809'
    },
    'v2ray_socks5': {
        'proxy_type': 'socks5',
        'host': '127.0.0.1',
        'port': '10808'
    },
    # Shadowsocks 默认配置
    'ss_http': {
        'proxy_type': 'http',
        'host': '127.0.0.1',
        'port': '1080'
    },
    # 自定义代理
    'custom': {
        'proxy_type': 'http',  # 或 'socks5', 'socks4'
        'host': '127.0.0.1',
        'port': '8080'
    }
}

if __name__ == "__main__":
    # 使用示例
    excel_file = "总数据.xlsx"  # 替换为你的Excel文件路径
    output_dir = "html_pdfs"  # 输出目录
    url_column_name = "来源网址"  # URL所在列的列名

    # 代理设置 - 根据你的本地代理选择对应的配置
    proxy_settings = PROXY_CONFIGS['clash_http']  # 例如使用Clash的HTTP代理

    # 访问国外网站建议增加等待时间
    wait_time = 8

    print("代理配置信息:")
    print(f"类型: {proxy_settings['proxy_type']}")
    print(f"地址: {proxy_settings['host']}:{proxy_settings['port']}")
    print(f"等待时间: {wait_time}秒\n")

    main(excel_file, output_dir, url_column_name, wait_time, proxy_settings)