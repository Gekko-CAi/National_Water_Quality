# -*- coding: utf-8 -*-
"""
国家水质自动综合监管平台 - 数据抓取脚本
从 https://szzdjc.cnemc.cn:8070/GJZ/Business/Publish/Main.html 抓取实时水质监测数据
GitHub Actions 每2小时运行一次，数据按天合并保存为 CSV（北京时间）
"""

import os
import sys
import re
import csv
import time
from datetime import datetime, timedelta, timezone
from collections import defaultdict, OrderedDict

import requests
import urllib3

# 北京时区 UTC+8
BEIJING_TZ = timezone(timedelta(hours=8))

# 禁用 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ======================== 配置 ========================
API_URL = "https://szzdjc.cnemc.cn:8070/GJZ/Ajax/Publish.ashx"
REFERER_URL = "https://szzdjc.cnemc.cn:8070/GJZ/Business/Publish/RealDatas.html"
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "National Water Quality")
PAGE_SIZE = 2000  # 每页记录数
REQUEST_TIMEOUT = 60  # 请求超时秒数
PAGE_DELAY = 2  # 每页请求间隔秒数

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Content-Type": "application/x-www-form-urlencoded",
    "Referer": REFERER_URL,
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://szzdjc.cnemc.cn:8070",
}

# 水质类别映射
WATER_QUALITY_MAP = {
    "1": "I类", "2": "II类", "3": "III类",
    "4": "IV类", "5": "V类", "6": "劣V类", "7": "未监测"
}

# 去重键：以 (监测断面, 监测时间) 作为唯一标识
DEDUP_KEYS = ["监测断面", "监测时间"]


# ======================== 数据抓取 ========================
def parse_cell_value(html_str):
    """解析单元格 HTML 值，提取原始值和显示值"""
    if not html_str or html_str == "--":
        return ("--", "--")

    original_match = re.search(r"原始值[：:]\s*([\d.]+)", html_str)
    display_match = re.search(r">([^<]+)<", html_str)

    original_val = original_match.group(1) if original_match else html_str
    display_val = display_match.group(1) if display_match else html_str

    if "<" in original_val:
        original_val = re.sub(r"<[^>]+>", "", original_val)
    if "<" in display_val:
        display_val = re.sub(r"<[^>]+>", "", display_val)

    return (original_val.strip(), display_val.strip())


def parse_thead(thead_list):
    """解析表头，提取干净的列名"""
    clean_headers = []
    for h in thead_list:
        clean = re.sub(r"<[^>]+>", "", h)
        clean = re.sub(r"\s+", " ", clean).strip()
        clean_headers.append(clean)
    return clean_headers


def fetch_page(page_index, page_size=PAGE_SIZE, max_retries=3):
    """抓取单页数据"""
    data = {
        "action": "getRealDatas",
        "AreaID": "",
        "RiverID": "",
        "MNName": "",
        "PageIndex": str(page_index),
        "PageSize": str(page_size),
    }

    for attempt in range(max_retries):
        try:
            response = requests.post(
                API_URL, headers=HEADERS, data=data,
                verify=False, timeout=REQUEST_TIMEOUT
            )
            result = response.json()
            if result.get("result") and result["result"] != 0:
                return result
            else:
                print(f"  [警告] 第{page_index}页返回空数据")
                return None
        except requests.exceptions.Timeout:
            print(f"  [重试 {attempt+1}/{max_retries}] 第{page_index}页请求超时")
            time.sleep(5)
        except Exception as e:
            print(f"  [重试 {attempt+1}/{max_retries}] 第{page_index}页请求异常: {e}")
            time.sleep(5)

    print(f"  [错误] 第{page_index}页抓取失败（已重试{max_retries}次）")
    return None


def fetch_all_data():
    """抓取所有页的数据"""
    now_bj = datetime.now(BEIJING_TZ)
    print(f"[{now_bj.strftime('%Y-%m-%d %H:%M:%S')}] 开始抓取水质监测数据（北京时间）...")

    first_page = fetch_page(1)
    if not first_page:
        print("[错误] 第一页数据抓取失败，无法继续")
        return None

    total_pages = first_page.get("total", 1)
    total_records = first_page.get("records", 0)
    thead = first_page.get("thead", [])
    all_tbody = first_page.get("tbody", [])

    clean_headers = parse_thead(thead)
    print(f"  总记录数: {total_records}, 总页数: {total_pages}")
    print(f"  表头列: {clean_headers}")

    for page_idx in range(2, total_pages + 1):
        print(f"  正在抓取第 {page_idx}/{total_pages} 页...")
        time.sleep(PAGE_DELAY)
        page_data = fetch_page(page_idx)
        if page_data and page_data.get("tbody"):
            all_tbody.extend(page_data["tbody"])
        else:
            print(f"  [警告] 第{page_idx}页无数据，跳过")

    print(f"  抓取完成，共获取 {len(all_tbody)} 条记录")
    return {"headers": clean_headers, "data": all_tbody}


# ======================== 数据处理 ========================
def process_data(raw_data):
    """将原始数据解析为结构化记录列表"""
    headers = raw_data["headers"]
    rows = raw_data["data"]
    records = []

    # 使用北京时间的年份
    current_year = datetime.now(BEIJING_TZ).year

    for row in rows:
        record = {}
        for i, val in enumerate(row):
            if i < len(headers):
                col_name = headers[i]
            else:
                col_name = f"列{i+1}"

            if i == 4:  # 水质类别
                quality_level = WATER_QUALITY_MAP.get(str(val), val)
                record[col_name] = quality_level
            elif i >= 5:  # 监测指标值（含 HTML）
                original, display = parse_cell_value(val)
                record[col_name] = original
            elif i == 3:  # 监测时间
                time_str = str(val).strip()
                if time_str and time_str != "--":
                    record[col_name] = f"{current_year}-{time_str}"
                else:
                    record[col_name] = time_str
            else:
                record[col_name] = str(val).strip() if val else ""

        # 抓取时间用北京时间
        record["抓取时间"] = datetime.now(BEIJING_TZ).strftime("%Y-%m-%d %H:%M:%S")
        records.append(record)

    return records, headers


# ======================== 读取已有 CSV ========================
def read_existing_csv(filepath, all_headers):
    """读取已有 CSV 文件并返回按去重键索引的记录字典"""
    existing = OrderedDict()
    if not os.path.exists(filepath):
        return existing

    try:
        with open(filepath, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                key_parts = []
                for k in DEDUP_KEYS:
                    key_parts.append(row.get(k, "").strip())
                key = "||".join(key_parts)
                existing[key] = row
        print(f"  读取已有文件: {os.path.basename(filepath)} ({len(existing)} 条)")
    except Exception as e:
        print(f"  [警告] 读取已有文件失败: {e}")

    return existing


# ======================== CSV 按天合并保存 ========================
def save_daily_csv(records, headers):
    """
    按监测时间的日期分组，每天合并保存为一个 CSV 文件。
    文件名格式: National_Water_YYYYMMDD.csv
    合并策略：以 (监测断面, 监测时间) 为唯一键去重，新数据覆盖旧数据。
    """
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    all_headers = list(headers) + ["抓取时间"]

    # 按监测时间日期分组
    daily_data = defaultdict(list)
    for record in records:
        time_val = str(record.get("监测时间", ""))
        if time_val and time_val != "--" and len(time_val) >= 10:
            date_str = time_val[:10]  # "2026-07-09"
        else:
            date_str = "未知日期"
        daily_data[date_str].append(record)

    saved_files = []
    total_new = 0  # 本次新增记录数
    total_existing = 0  # 已有记录数

    for date_str, new_records in sorted(daily_data.items()):
        date_compact = date_str.replace("-", "")
        filename = f"National_Water_{date_compact}.csv"
        filepath = os.path.join(OUTPUT_DIR, filename)

        # 1. 读取已有数据
        existing = read_existing_csv(filepath, all_headers)
        total_existing += len(existing)

        # 2. 合并新数据（以去重键覆盖已有记录）
        added = 0
        updated = 0
        for record in new_records:
            key_parts = []
            for k in DEDUP_KEYS:
                key_parts.append(record.get(k, "").strip())
            key = "||".join(key_parts)

            if key in existing:
                # 已有记录：用最新抓取数据覆盖
                existing[key] = record
                updated += 1
            else:
                existing[key] = record
                added += 1

        # 3. 写入合并后的全部数据
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=all_headers)
            writer.writeheader()
            for rec in existing.values():
                writer.writerow(rec)

        saved_files.append(filepath)
        total_new += len(existing)
        print(f"  合并保存: {filename} "
              f"(已有 {len(existing) - added - updated}, "
              f"新增 {added}, 更新 {updated}, "
              f"合计 {len(existing)} 条)")

    return saved_files, total_new, total_existing


# ======================== 主函数 ========================
def main():
    """主函数：抓取数据并按天合并保存为 CSV（北京时间）"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # 抓取数据
    raw_data = fetch_all_data()
    if not raw_data or not raw_data.get("data"):
        print("[错误] 未获取到数据，程序退出")
        return False

    # 处理数据
    records, headers = process_data(raw_data)
    if not records:
        print("[错误] 数据处理后为空，程序退出")
        return False

    # 按天合并保存为 CSV
    print(f"\n按天合并保存数据为 CSV（北京时间）...")
    saved_files, total_new, total_existing = save_daily_csv(records, headers)

    now_bj = datetime.now(BEIJING_TZ)
    print(f"\n{'='*60}")
    print(f"抓取完成！")
    print(f"  保存文件数: {len(saved_files)}")
    print(f"  本次抓取: {len(records)} 条")
    print(f"  合并后总计: {total_new} 条")
    print(f"  抓取时间: {now_bj.strftime('%Y-%m-%d %H:%M:%S')} (北京时间)")
    print(f"{'='*60}")
    return True


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
