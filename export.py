#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Stars Exporter v2.1
导出 GitHub Stars 为 JSON / CSV / XLSX / HTML（含主题切换、搜索、排序、文件下载）
支持 GitHub Actions 自动部署到 GitHub Pages
"""

import requests
import json
import csv
import base64
import os
import sys
import time
import argparse
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("[提示] 未安装 openpyxl，XLSX 导出将不可用。安装命令: pip install openpyxl")

API_BASE = "https://api.github.com"

def get_headers(token):
    headers = {"Accept": "application/vnd.github.v3+json"}
    if token:
        headers["Authorization"] = f"token {token}"
    return headers

def fetch_all_starred(username, token):
    repos = []
    page = 1
    per_page = 100
    headers = get_headers(token)
    print(f"[1/4] 正在获取用户 {username} 的星标仓库...")
    while True:
        url = f"{API_BASE}/users/{username}/starred?per_page={per_page}&page={page}"
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 404:
            print(f"  错误：用户 '{username}' 不存在")
            return []
        if resp.status_code == 403:
            print(f"  错误：API 速率限制。建议添加 GITHUB_TOKEN")
            return []
        resp.raise_for_status()
        data = resp.json()
        if not data:
            break
        repos.extend(data)
        print(f"  第 {page} 页: {len(data)} 个，累计 {len(repos)} 个")
        if len(data) < per_page:
            break
        page += 1
        time.sleep(0.2)
    print(f"  共获取 {len(repos)} 个星标仓库\n")
    return repos

def fetch_readme(owner, repo_name, token):
    url = f"{API_BASE}/repos/{owner}/{repo_name}/readme"
    try:
        resp = requests.get(url, headers=get_headers(token), timeout=15)
        if resp.status_code == 404:
            return ""
        resp.raise_for_status()
        data = resp.json()
        return base64.b64decode(data.get("content", "")).decode("utf-8", errors="ignore")
    except Exception:
        return ""

def build_repo_data(repos, token, fetch_readme_flag=True):
    results = []
    total = len(repos)
    print(f"[2/4] 正在处理 {total} 个仓库...")
    for idx, repo in enumerate(repos, 1):
        owner = repo["owner"]["login"]
        name = repo["name"]
        full_name = repo["full_name"]
        readme = fetch_readme(owner, name, token) if fetch_readme_flag else ""
        if fetch_readme_flag:
            time.sleep(0.08)
        item = {
            "序号": idx, "项目名称": name, "项目全名": full_name,
            "项目链接": repo["html_url"], "仓库描述": repo.get("description") or "",
            "主页网址": repo.get("homepage") or "", "README": readme,
            "项目语言": repo.get("language") or "", "星标数": repo.get("stargazers_count", 0),
            "Fork数": repo.get("forks_count", 0), "Watch数": repo.get("watchers_count", 0),
            "Open_Issues": repo.get("open_issues_count", 0), "默认分支": repo.get("default_branch", ""),
            "创建时间": repo.get("created_at", ""), "更新时间": repo.get("updated_at", ""),
            "推送时间": repo.get("pushed_at", ""), "仓库大小_KB": repo.get("size", 0),
            "是否为Fork": repo.get("fork", False), "Topics": ", ".join(repo.get("topics", [])),
            "License": (repo.get("license") or {}).get("name", ""), "Owner": owner,
            "Owner类型": repo["owner"].get("type", ""), "Owner头像": repo["owner"].get("avatar_url", ""),
        }
        results.append(item)
        if idx % 200 == 0 or idx == total:
            print(f"  [{idx}/{total}] 已处理")
    print(f"  全部处理完成\n")
    return results

def export_json(data, filepath):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  JSON: {filepath}")

def export_csv(data, filepath):
    if not data: return
    fieldnames = list(data[0].keys())
    csv_fieldnames = [f for f in fieldnames if f not in ["README", "Owner头像"]]
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=csv_fieldnames)
        writer.writeheader()
        for row in data:
            writer.writerow({k: v for k, v in row.items() if k in csv_fieldnames})
    print(f"  CSV: {filepath}")

def export_xlsx(data, filepath):
    if not HAS_OPENPYXL:
        print("  未安装 openpyxl，跳过 XLSX")
        return
    if not data: return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "GitHub Stars"
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="24292F", end_color="24292F", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style="thin", color="D0D7DE"), right=Side(style="thin", color="D0D7DE"), top=Side(style="thin", color="D0D7DE"), bottom=Side(style="thin", color="D0D7DE"))
    fieldnames = list(data[0].keys())
    for col_idx, field in enumerate(fieldnames, 1):
        cell = ws.cell(row=1, column=col_idx, value=field)
        cell.font = header_font; cell.fill = header_fill; cell.alignment = header_align; cell.border = thin_border
    for row_idx, row in enumerate(data, 2):
        for col_idx, field in enumerate(fieldnames, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row.get(field, ""))
            cell.alignment = cell_align; cell.border = thin_border
    col_widths = {"序号": 6, "项目名称": 25, "项目全名": 35, "项目链接": 45, "仓库描述": 50, "主页网址": 40, "README": 60, "项目语言": 12, "星标数": 10, "Fork数": 10, "Watch数": 10, "Open_Issues": 12, "默认分支": 12, "创建时间": 20, "更新时间": 20, "推送时间": 20, "仓库大小_KB": 14, "是否为Fork": 12, "Topics": 40, "License": 20, "Owner": 18, "Owner类型": 12, "Owner头像": 45}
    for col_idx, field in enumerate(fieldnames, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(field, 20)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(filepath)
    print(f"  XLSX: {filepath}")
