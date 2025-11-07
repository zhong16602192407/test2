import pandas as pd
import requests
import re
import os
from urllib.parse import urlparse
import time
import json
from datetime import datetime
import random


def extract_urls(text):
    """从文本中提取所有URL链接"""
    if pd.isna(text):
        return []

    # 匹配http或https开头的URL
    url_pattern = r'https?://[^\s\u4e00-\u9fa5，。！；）】」》\)\]\}]+'
    urls = re.findall(url_pattern, str(text))
    return urls


def sanitize_filename(filename):
    """清理文件名，移除不合法字符"""
    # 移除Windows文件名不允许的字符
    invalid_chars = r'[<>:"/\\|?*]'
    filename = re.sub(invalid_chars, '_', filename)
    # 限制文件名长度
    if len(filename) > 200:
        filename = filename[:200]
    return filename


def download_pdf(url, filename, output_dir='downloads', max_retries=3):
    """
    下载PDF文件，支持重试机制
    :param url: 下载链接
    :param filename: 文件名
    :param output_dir: 输出目录
    :param max_retries: 最大重试次数
    :return: (成功与否, 错误信息)
    """
    # 创建输出目录
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    filepath = os.path.join(output_dir, filename)

    # 检查文件是否已存在
    if os.path.exists(filepath):
        file_size = os.path.getsize(filepath)
        if file_size > 0:  # 文件存在且不为空
            print(f"⊙ 文件已存在，跳过: {filename} ({file_size} bytes)")
            return True, None

    # 准备多个User-Agent，把成功率高的放前面
    user_agents = [
        # Firefox系列（测试发现这个最好用）
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:122.0) Gecko/20100101 Firefox/122.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:121.0) Gecko/20100101 Firefox/121.0',
        # Edge系列
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
        # Safari系列
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15',
        # Chrome系列
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    ]

    # 重试循环
    for attempt in range(max_retries):
        try:
            # 第一次尝试用最好的User-Agent，之后随机选择避免被识别
            if attempt == 0:
                selected_ua = user_agents[0]  # 第一次用最好的（Firefox）
            else:
                selected_ua = random.choice(user_agents)  # 重试时随机选择

            headers = {
                'User-Agent': selected_ua,
                'Accept': 'application/pdf,application/octet-stream,*/*',
                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
                'Referer': url,
            }

            # 发送请求
            response = requests.get(url, headers=headers, timeout=30, stream=True)
            response.raise_for_status()

            # 保存文件
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

            # 验证文件是否下载成功
            if os.path.getsize(filepath) > 0:
                print(f"✓ 下载成功: {filename}")
                return True, None
            else:
                # 文件为空，删除并重试
                os.remove(filepath)
                raise Exception("下载的文件为空")

        except requests.exceptions.HTTPError as e:
            error_msg = f"HTTP错误 {e.response.status_code}"
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt  # 指数退避: 1s, 2s, 4s
                print(f"  ⚠ {error_msg}，{wait_time}秒后重试 ({attempt + 1}/{max_retries})...")
                time.sleep(wait_time)
            else:
                print(f"✗ 下载失败: {filename}")
                print(f"  错误: {error_msg}")
                return False, error_msg

        except requests.exceptions.Timeout as e:
            error_msg = "请求超时"
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                print(f"  ⚠ {error_msg}，{wait_time}秒后重试 ({attempt + 1}/{max_retries})...")
                time.sleep(wait_time)
            else:
                print(f"✗ 下载失败: {filename}")
                print(f"  错误: {error_msg}")
                return False, error_msg

        except requests.exceptions.ConnectionError as e:
            error_msg = "连接错误"
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                print(f"  ⚠ {error_msg}，{wait_time}秒后重试 ({attempt + 1}/{max_retries})...")
                time.sleep(wait_time)
            else:
                print(f"✗ 下载失败: {filename}")
                print(f"  错误: {error_msg}")
                return False, error_msg

        except Exception as e:
            error_msg = str(e)
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                print(f"  ⚠ {error_msg}，{wait_time}秒后重试 ({attempt + 1}/{max_retries})...")
                time.sleep(wait_time)
            else:
                print(f"✗ 下载失败: {filename}")
                print(f"  错误: {error_msg}")
                return False, error_msg

    # 如果所有重试都失败
    return False, "所有重试都失败"


def main(excel_file, output_dir='downloads'):
    """
    主函数
    :param excel_file: Excel文件路径
    :param output_dir: 输出目录
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)

        print(f"读取Excel文件: {excel_file}")
        print(f"共 {len(df)} 行数据\n")

        # 获取列名
        columns = df.columns.tolist()
        print(f"Excel列名: {columns}\n")

        # 假设第一列是序号，第三列是标题，最后一列是备注（含链接）
        col_index = columns[0]  # 第一列
        col_title = columns[2]  # 第三列
        col_remark = columns[-1]  # 最后一列

        success_count = 0
        fail_count = 0
        skip_count = 0
        failed_items = []  # 记录失败的项

        # 遍历每一行
        for idx, row in df.iterrows():
            序号 = str(row[col_index])
            标题 = str(row[col_title])
            备注 = row[col_remark]

            # 提取备注中的所有链接
            urls = extract_urls(备注)

            if not urls:
                print(f"行 {idx + 1}: 序号[{序号}] - 未找到链接，跳过")
                continue

            print(f"\n行 {idx + 1}: 序号[{序号}] - 标题[{标题}]")
            print(f"  找到 {len(urls)} 个链接")

            # 下载每个链接
            for url_idx, url in enumerate(urls, 1):
                # 构建文件名
                if len(urls) > 1:
                    # 多个链接时，添加序号后缀
                    filename = f"{序号}-{标题}-{url_idx}.pdf"
                else:
                    filename = f"{序号}-{标题}.pdf"

                # 清理文件名
                filename = sanitize_filename(filename)

                print(f"  [{url_idx}/{len(urls)}] 下载: {url}")

                # 下载文件
                success, error_msg = download_pdf(url, filename, output_dir)
                if success:
                    # 判断是否是跳过的文件
                    if error_msg is None and os.path.exists(os.path.join(output_dir, filename)):
                        file_existed_before = True  # 这个逻辑可以优化，但当前简化处理
                    success_count += 1
                else:
                    fail_count += 1
                    # 记录失败信息
                    failed_items.append({
                        '行号': idx + 1,
                        '序号': 序号,
                        '标题': 标题,
                        '链接': url,
                        '文件名': filename,
                        '错误': error_msg
                    })

                # 添加延迟，避免请求过快
                time.sleep(1)

        # 保存失败记录
        if failed_items:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            failed_log_file = f'failed_downloads_{timestamp}.json'
            failed_txt_file = f'failed_downloads_{timestamp}.txt'

            # 保存为JSON格式（详细信息）
            with open(failed_log_file, 'w', encoding='utf-8') as f:
                json.dump(failed_items, f, ensure_ascii=False, indent=2)

            # 保存为TXT格式（便于阅读）
            with open(failed_txt_file, 'w', encoding='utf-8') as f:
                f.write(f"下载失败记录 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")
                for item in failed_items:
                    f.write(f"行号: {item['行号']}\n")
                    f.write(f"序号: {item['序号']}\n")
                    f.write(f"标题: {item['标题']}\n")
                    f.write(f"链接: {item['链接']}\n")
                    f.write(f"文件名: {item['文件名']}\n")
                    f.write(f"错误: {item['错误']}\n")
                    f.write("-" * 80 + "\n\n")

            print(f"\n失败记录已保存:")
            print(f"  - {failed_log_file} (JSON格式)")
            print(f"  - {failed_txt_file} (文本格式)")

        # 打印统计信息
        print(f"\n{'=' * 50}")
        print(f"下载完成!")
        print(f"成功: {success_count} 个")
        print(f"失败: {fail_count} 个")
        if fail_count > 0:
            print(f"  提示: 失败的文件已记录在 failed_downloads_*.txt 中，可以后续手动处理")
        print(f"文件保存在: {output_dir}")

    except Exception as e:
        print(f"错误: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    # 使用示例
    excel_file = "总数据.xlsx"  # 替换为你的Excel文件路径
    output_dir = "downloads"  # 输出目录

    main(excel_file, output_dir)