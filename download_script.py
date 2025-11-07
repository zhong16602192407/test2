import pandas as pd
import requests
import re
import os
from urllib.parse import urlparse
import time


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


def download_pdf(url, filename, output_dir='downloads'):
    """下载PDF文件"""
    try:
        # 创建输出目录
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 发送请求
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=30, stream=True)
        response.raise_for_status()

        # 保存文件
        filepath = os.path.join(output_dir, filename)
        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        print(f"✓ 下载成功: {filename}")
        return True

    except Exception as e:
        print(f"✗ 下载失败: {filename}")
        print(f"  错误信息: {str(e)}")
        return False


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
                if download_pdf(url, filename, output_dir):
                    success_count += 1
                else:
                    fail_count += 1

                # 添加延迟，避免请求过快
                time.sleep(1)

        # 打印统计信息
        print(f"\n{'=' * 50}")
        print(f"下载完成!")
        print(f"成功: {success_count} 个")
        print(f"失败: {fail_count} 个")
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