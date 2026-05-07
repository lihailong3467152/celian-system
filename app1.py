#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
策链系统管理平台 - 策链系统(ULD-CeLink)
面向咨询生产与项目协调的系统
管理端 + 第三方机构端(主账号 + 子账号)
支持多电脑部署、数据实时同步、账号互通、权限隔离
"""

import streamlit as st
import streamlit.components.v1 as components
import sqlite3
import pandas as pd
import bcrypt
import os
import json
import base64
import hashlib
import socket
import tempfile
import mimetypes
import html
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import shutil
import re
import textwrap

# ==================== 配置 ====================
DATABASE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'performance.db')
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
ADMIN_DEFAULT_PASSWORD = "Admin@123456"
ORG_DEFAULT_PASSWORD = "Org@123456"
PASSWORD_RULE_TEXT = "密码至少10位，且必须包含大写字母、小写字母、数字和特殊字符"
APP_SCHEMA_VERSION = "20260506_unique_user_contact_project_steps_v1"

# 创建上传目录
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# 项目分类配置
PROJECT_CATEGORIES = {
    '1': {
        'name': '统计调查研究',
        'subcategories': {
            '1': '多源数据采集',
            '2': '专项调查研究'
        }
    },
    '2': {
        'name': '政府绩效评估',
        'subcategories': {
            '1': '财政绩效评估',
            '2': '行政绩效评估'
        }
    },
    '3': {
        'name': '社会经济咨询',
        'subcategories': {
            '1': '企业管理咨询',
            '2': '公共决策咨询'
        }
    },
    '0': {
        'name': '其他项目',
        'subcategories': {}
    }
}

ROLE_NAMES = {
    'super_admin': '超级管理员',
    'org_admin': '机构主账号',
    'org_user': '机构子账号',
}

CLIENT_ALLOWED_ROLES = {
    'admin': {'super_admin'},
    'org': {'org_admin', 'org_user'},
}

PAGE_MENU = {
    'super_admin': [
        ("📊", "数据大盘", "dashboard"),
        ("🏢", "机构管理", "organizations"),
        ("👥", "账号管理", "users"),
        ("📋", "项目审核", "projects"),
        ("📝", "日志查看", "logs"),
        ("📥", "数据导出", "export"),
        ("✅", "审批待办", "approval"),
        ("📨", "消息通知", "messages"),
        ("📚", "项目智库", "indicators"),
        ("📈", "可视化大屏", "visualization"),
    ],
    'org_admin': [
        ("🏠", "工作台", "dashboard"),
        ("🏢", "信息维护", "info"),
        ("👥", "子账号管理", "sub_accounts"),
        ("📋", "项目管理", "projects"),
        ("📚", "项目智库", "knowledge"),
        ("✅", "待办事项", "todos"),
        ("📨", "消息通知", "messages"),
    ],
    'org_user': [
        ("🏠", "工作台", "dashboard"),
        ("🏢", "信息维护", "info"),
        ("📋", "项目管理", "projects"),
        ("📚", "项目智库", "knowledge"),
        ("✅", "待办事项", "todos"),
        ("📨", "消息通知", "messages"),
    ],
}

PAGE_ACCESS = {
    page: {role for role, items in PAGE_MENU.items() for _, _, item_page in items if item_page == page}
    for page in {item_page for items in PAGE_MENU.values() for _, _, item_page in items}
}

# 阶段名称
STAGE_NAMES = {
    1: 'G0｜议题生成（U）',
    2: 'G1｜立项论证（U + D）',
    3: 'G2｜数据构建（D）',
    4: 'G3｜结构分析（I）',
    5: 'G4｜机理建模（K）',
    6: 'G5｜决策判断（W）',
    7: 'G6｜方案设计（S）',
    8: 'G7｜交付实施（S）',
    9: 'G8｜风险治理（X）'
}

# 阶段工程目的
STAGE_PURPOSES = {
    1: '把“模糊关切/直觉判断/外部诉求”转化为可进入研究与决策体系的问题对象。',
    2: '把“问题”转化为可研究、可决策、可投入资源的项目对象。',
    3: '构建可审计、可复核的数据证据系统。',
    4: '把数据转化为结构化信息系统。',
    5: '从结构中提炼规律、机制与因果关系。',
    6: '形成可负责的价值判断与取舍方案。',
    7: '把判断转化为工程化方案体系。',
    8: '将方案转化为执行系统。',
    9: '构建不确定性治理系统。'
}

TOTAL_STAGES = len(STAGE_NAMES)

def format_gate(stage):
    """统一Gate显示，兼容历史异常阶段值"""
    if stage is None:
        return "-"
    if stage in STAGE_NAMES:
        return STAGE_NAMES[stage]
    try:
        stage_num = int(stage)
        gate_num = stage_num - 1 if stage_num > 0 else 0
        return f"G{gate_num}"
    except Exception:
        return str(stage)

def format_datetime_display(value):
    """统一时间显示为 YYYY-MM-DD HH:MM:SS，自动去除微秒"""
    if not value:
        return "-"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    try:
        text = str(value).strip()
        if not text:
            return "-"
        # 兼容 SQLite/ISO 字符串，如 2026-04-03 11:16:21.972716 / 2026-04-03T11:16:21
        normalized = text.replace("T", " ")
        if normalized.endswith("Z"):
            normalized = normalized[:-1]
        dt = datetime.fromisoformat(normalized)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        # 无法解析时，保底去掉小数秒部分
        text = str(value)
        return text.split(".")[0] if "." in text else text

def get_default_password_for_role(role):
    """按账号角色返回初始化/重置后的默认密码。"""
    return ADMIN_DEFAULT_PASSWORD if role == 'super_admin' else ORG_DEFAULT_PASSWORD

def is_default_password(password):
    return password in {ADMIN_DEFAULT_PASSWORD, ORG_DEFAULT_PASSWORD}

def validate_password_policy(password, allow_default=False):
    """校验密码复杂度；改密时不允许继续使用默认密码。"""
    if not password:
        return "请输入密码"
    if len(password) < 10:
        return "密码长度不能少于10位"
    if not re.search(r"[A-Z]", password):
        return "密码必须包含大写字母"
    if not re.search(r"[a-z]", password):
        return "密码必须包含小写字母"
    if not re.search(r"\d", password):
        return "密码必须包含数字"
    if not re.search(r"[^A-Za-z0-9]", password):
        return "密码必须包含特殊字符"
    if not allow_default and is_default_password(password):
        return "新密码不能与系统默认密码相同"
    return None

def normalize_password_hash(password_hash):
    """兼容 SQLite 返回 bytes 或 str 的 bcrypt hash。"""
    if isinstance(password_hash, bytes):
        return password_hash
    if isinstance(password_hash, str):
        return password_hash.encode('utf-8')
    return bytes(password_hash)

# 页面配置
st.set_page_config(
    page_title="策链系统(ULD-CeLink)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义样式 ====================
def apply_custom_styles():
    """应用自定义CSS样式"""
    st.markdown('''
    <style>
    /* 主题色：主背景 (蓝绿色)、深蓝色边栏/卡片和中间色 */
    :root {
        --main-bg: #008B8B; /* 蓝绿色主背景 */
        --deep-blue: #00264D; /* 深蓝色（侧边栏、强调色） */
        --mid-blue: #004D73; /* 中间色，深蓝与蓝绿色之间 */
        --card-bg: linear-gradient(135deg, var(--deep-blue) 0%, var(--mid-blue) 50%, var(--main-bg) 100%);
        --surface-bg: rgba(255,255,255,0.06); /* 浅表面，用于卡片边框/内衬 */
        --text-on-dark: #FFFFFF; /* 深色背景上的文字 */
        --text-on-light: #0b1a2b; /* 浅背景上的文字，保证对比 */
        --accent: #00b3a6;
    }
    
    /* 隐藏Streamlit默认元素 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {display: none !important;}
    .anchor-link, a.anchor-link {
        display: none !important;
        visibility: hidden !important;
    }

    /* 确保按钮在列内平行对齐，消除 Streamlit 默认的外边距差异 */
    div[data-testid="stButton"], div[data-testid="stDownloadButton"] {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
    }

    /* 页面背景——主区域使用蓝绿色，增加轻微纹理感以降低单色视觉疲劳 */
    .reportview-container, .main, .block-container {
        background: var(--main-bg) !important;
        color: var(--text-on-dark) !important;
        background-image: linear-gradient(180deg, rgba(255,255,255,0.02), transparent);
    }

    /* 去除主内容区顶部大块留白 */
    .block-container {
        padding-top: 0.6rem !important;
    }

    /* 页面切换短遮罩：只隐藏 Streamlit 前端应用 delta 时短暂混合的新旧主内容 */
    html.celink-page-switching section.main [data-testid="stAppViewBlockContainer"] {
        opacity: 0 !important;
        pointer-events: none !important;
        transition: none !important;
    }
    
    /* 登录页面样式 */
    .login-container {
        max-width: 520px;
        margin: 50px auto;
        padding: 36px;
        background: var(--card-bg);
        border-radius: 16px;
        box-shadow: 0 14px 40px rgba(0,0,0,0.35);
        color: var(--text-on-dark);
    }
    
    .login-title {
        text-align: center;
        color: #cfefff;
        font-size: 42px;
        font-weight: 800;
        margin-bottom: 12px;
        letter-spacing: 0.6px;
        line-height: 1.2;
        text-shadow: 0 4px 14px rgba(0,0,0,0.35);
    }
    
    .login-subtitle {
        text-align: center;
        color: rgba(230, 250, 255, 0.92);
        font-size: 18px;
        font-weight: 500;
        margin-bottom: 24px;
        opacity: 1;
        letter-spacing: 0.2px;
    }

    /* 管理端首页顶部标题（红框区域） */
    .dashboard-main-title {
        margin: 0;
        color: #9fdfff; /* 浅蓝 */
        font-size: 46px;
        font-weight: 800;
        letter-spacing: 0.6px;
        line-height: 1.15;
    }

    .dashboard-sub-title {
        margin: 6px 0 18px 0;
        color: rgba(224, 245, 255, 0.96); /* 蓝白色 */
        font-size: 22px; /* 比主标题小 */
        font-weight: 550;
        letter-spacing: 0.2px;
        line-height: 1.3;
    }

    /* 机构端工作台 */
    .org-workbench {
        color: #eefcff;
        animation: workbenchFadeIn 0.45s ease both;
    }

    .org-hero {
        position: relative;
        overflow: hidden;
        padding: 30px 34px;
        border-radius: 26px;
        border: 1px solid rgba(255,255,255,0.18);
        background:
            radial-gradient(circle at 8% 0%, rgba(112, 224, 255, 0.28), transparent 34%),
            radial-gradient(circle at 88% 12%, rgba(255, 213, 117, 0.20), transparent 26%),
            linear-gradient(135deg, rgba(0,38,77,0.96) 0%, rgba(0,77,115,0.88) 48%, rgba(0,139,139,0.82) 100%);
        box-shadow: 0 22px 60px rgba(0, 20, 45, 0.30);
    }

    .org-hero::after {
        content: "";
        position: absolute;
        inset: auto -120px -180px auto;
        width: 360px;
        height: 360px;
        border-radius: 50%;
        background: rgba(255,255,255,0.10);
    }

    .org-hero-grid {
        position: relative;
        z-index: 1;
        display: grid;
        grid-template-columns: minmax(0, 1.25fr) minmax(280px, 0.75fr);
        gap: 22px;
        align-items: end;
    }

    .org-kicker {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 7px 12px;
        border-radius: 999px;
        background: rgba(255,255,255,0.12);
        color: rgba(234,251,255,0.88);
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 0.4px;
    }

    .org-hero h1 {
        margin: 18px 0 10px 0;
        color: #f6fdff;
        font-size: clamp(34px, 4vw, 58px);
        line-height: 1.05;
        font-weight: 900;
        letter-spacing: -1px;
    }

    .org-hero-tags {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-top: 22px;
    }

    .org-hero-tag {
        display: inline-flex;
        align-items: center;
        padding: 8px 12px;
        border-radius: 999px;
        background: rgba(255,255,255,0.10);
        border: 1px solid rgba(255,255,255,0.14);
        color: rgba(234,251,255,0.84);
        font-size: 13px;
        font-weight: 800;
        letter-spacing: 0.2px;
    }

    .org-hero-panel {
        padding: 18px;
        border-radius: 20px;
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.16);
        backdrop-filter: blur(10px);
    }

    .org-hero-panel-title {
        color: rgba(234,251,255,0.72);
        font-size: 13px;
        font-weight: 700;
        margin-bottom: 10px;
    }

    .org-hero-panel-value {
        color: #ffffff;
        font-size: 24px;
        font-weight: 850;
        line-height: 1.35;
        word-break: break-word;
    }

    .org-dashboard-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 16px;
        margin: 18px 0;
    }

    .org-workbench a {
        text-decoration: none !important;
    }

    .org-click-card {
        display: block;
        color: inherit !important;
        border-radius: 22px;
        transition: transform 0.18s ease, box-shadow 0.18s ease, filter 0.18s ease;
    }

    .org-click-card:hover {
        transform: translateY(-4px);
        filter: brightness(1.02);
    }

    .org-click-card,
    .org-panel-badge,
    .org-list-item,
    .org-project-row:not(.org-project-head) {
        cursor: pointer;
    }

    .org-hidden-click-layer {
        display: none;
    }

    .element-container:has(.org-metric-click-layer) + div[data-testid="stHorizontalBlock"] {
        position: relative;
        z-index: 10;
        margin-top: -169px;
        margin-bottom: 18px;
    }

    .element-container:has(.org-metric-click-layer) + div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] {
        height: 150px;
    }

    .element-container:has(.org-metric-click-layer) + div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] > button {
        height: 150px;
        min-height: 150px;
        opacity: 0;
        cursor: pointer;
        border-radius: 22px;
        box-shadow: none !important;
    }


    .org-metric-card {
        position: relative;
        overflow: hidden;
        min-height: 132px;
        padding: 20px;
        border-radius: 22px;
        background: linear-gradient(160deg, rgba(255,255,255,0.94), rgba(225,250,249,0.86));
        color: #08313f;
        border: 1px solid rgba(255,255,255,0.70);
        box-shadow: 0 18px 42px rgba(0, 31, 63, 0.18);
    }

    .org-metric-card::before {
        content: "";
        position: absolute;
        width: 84px;
        height: 84px;
        right: -26px;
        top: -24px;
        border-radius: 26px;
        transform: rotate(22deg);
        background: rgba(0,139,139,0.12);
    }

    .org-metric-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 12px;
    }

    .org-metric-icon {
        width: 42px;
        height: 42px;
        display: grid;
        place-items: center;
        border-radius: 15px;
        background: linear-gradient(135deg, #00264d, #008b8b);
        color: #fff;
        font-size: 22px;
        box-shadow: 0 12px 24px rgba(0, 52, 86, 0.22);
    }

    .org-metric-number {
        margin-top: 16px;
        color: #00264d;
        font-size: 38px;
        font-weight: 900;
        line-height: 1;
    }

    .org-metric-label {
        margin-top: 8px;
        color: rgba(8,49,63,0.70);
        font-size: 14px;
        font-weight: 800;
    }

    .org-content-grid {
        display: grid;
        grid-template-columns: minmax(0, 0.9fr) minmax(0, 1.1fr);
        gap: 18px;
        margin-top: 18px;
    }

    .org-panel {
        padding: 22px;
        border-radius: 24px;
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.18);
        box-shadow: 0 18px 46px rgba(0, 31, 63, 0.18);
        backdrop-filter: blur(8px);
    }

    .org-panel-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 16px;
        margin-bottom: 16px;
    }

    .org-panel-title {
        color: #f6fdff;
        font-size: 24px;
        font-weight: 900;
        letter-spacing: -0.3px;
    }

    .org-panel-badge {
        padding: 5px 10px;
        border-radius: 999px;
        background: rgba(255,255,255,0.14);
        color: rgba(238,252,255,0.86);
        font-size: 12px;
        font-weight: 800;
        white-space: nowrap;
        transition: background 0.18s ease, transform 0.18s ease;
        cursor: pointer;
    }

    .org-panel-badge:hover {
        background: rgba(255,255,255,0.24);
        transform: translateY(-1px);
    }

    .org-list {
        display: grid;
        gap: 10px;
    }

    .org-list-item {
        display: grid;
        grid-template-columns: 34px minmax(0,1fr);
        gap: 12px;
        align-items: start;
        padding: 13px 14px;
        border-radius: 16px;
        background: rgba(255,255,255,0.86);
        color: #092f3c;
        border: 1px solid rgba(255,255,255,0.7);
        transition: transform 0.18s ease, background 0.18s ease;
    }

    .org-list-item:hover {
        transform: translateX(3px);
        background: rgba(255,255,255,0.95);
    }

    .org-list-icon {
        width: 34px;
        height: 34px;
        display: grid;
        place-items: center;
        border-radius: 12px;
        background: rgba(0,139,139,0.12);
        color: #005b67;
        font-size: 17px;
    }

    .org-list-title {
        font-size: 15px;
        font-weight: 850;
        color: #092f3c;
        line-height: 1.45;
    }

    .org-list-meta {
        margin-top: 4px;
        color: rgba(9,47,60,0.62);
        font-size: 12px;
        font-weight: 700;
    }

    .org-empty {
        padding: 28px;
        border-radius: 18px;
        text-align: center;
        color: rgba(238,252,255,0.78);
        background: rgba(255,255,255,0.08);
        border: 1px dashed rgba(255,255,255,0.26);
    }

    .org-project-panel {
        margin-top: 18px;
    }

    .org-project-table {
        overflow: hidden;
        border-radius: 18px;
        background: rgba(255,255,255,0.94);
        color: #092f3c;
        border: 1px solid rgba(255,255,255,0.70);
    }

    .org-project-row {
        display: grid;
        grid-template-columns: minmax(180px, 1.05fr) minmax(210px, 1.2fr) 110px 160px;
        gap: 0;
        align-items: center;
        min-height: 46px;
        border-bottom: 1px solid rgba(0,38,77,0.08);
    }

    .org-project-row:last-child {
        border-bottom: 0;
    }

    .org-project-row > div {
        padding: 12px 14px;
        min-width: 0;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .org-project-head {
        background: rgba(0,38,77,0.07);
        color: rgba(9,47,60,0.70);
        font-size: 13px;
        font-weight: 900;
    }

    .org-status-pill {
        display: inline-flex;
        align-items: center;
        padding: 5px 10px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 900;
        background: rgba(0,139,139,0.12);
        color: #005b67;
    }

    .org-status-rejected {
        background: rgba(255, 82, 82, 0.14);
        color: #a12626;
    }

    .org-status-completed {
        background: rgba(28, 146, 96, 0.15);
        color: #166a49;
    }

    .admin-status-list {
        display: grid;
        gap: 12px;
    }

    .admin-status-item {
        padding: 14px;
        border-radius: 16px;
        background: rgba(255,255,255,0.88);
        color: #092f3c;
        border: 1px solid rgba(255,255,255,0.70);
    }

    .admin-status-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin-bottom: 10px;
        font-weight: 900;
    }

    .admin-status-count {
        color: rgba(9,47,60,0.68);
        font-size: 13px;
        font-weight: 900;
    }

    .admin-status-track {
        height: 10px;
        overflow: hidden;
        border-radius: 999px;
        background: rgba(0,38,77,0.08);
    }

    .admin-status-fill {
        height: 100%;
        border-radius: 999px;
        background: linear-gradient(90deg, #00264d, #00a99d);
    }

    .admin-activity-table {
        overflow: hidden;
        border-radius: 18px;
        background: rgba(255,255,255,0.94);
        color: #092f3c;
        border: 1px solid rgba(255,255,255,0.70);
    }

    .admin-activity-row {
        display: grid;
        grid-template-columns: 150px minmax(150px, 1fr) 130px 150px;
        align-items: center;
        min-height: 46px;
        border-bottom: 1px solid rgba(0,38,77,0.08);
    }

    .admin-activity-row:last-child {
        border-bottom: 0;
    }

    .admin-activity-row > div {
        padding: 12px 14px;
        min-width: 0;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .admin-activity-head {
        background: rgba(0,38,77,0.07);
        color: rgba(9,47,60,0.70);
        font-size: 13px;
        font-weight: 900;
    }

    .admin-todo-panel {
        min-height: 300px;
        padding: 22px;
        border-radius: 24px;
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.18);
        box-shadow: 0 18px 46px rgba(0, 31, 63, 0.18);
        backdrop-filter: blur(8px);
    }

    .admin-todo-items {
        display: grid;
        gap: 12px;
        margin-top: 16px;
        margin-bottom: 16px;
    }

    .admin-todo-card {
        display: grid;
        grid-template-columns: 42px minmax(0, 1fr) auto;
        gap: 12px;
        align-items: center;
        padding: 14px 16px;
        border-radius: 17px;
        background: rgba(255,255,255,0.88);
        color: #092f3c;
        border: 1px solid rgba(255,255,255,0.70);
        cursor: pointer;
        transition: transform 0.18s ease, background 0.18s ease, box-shadow 0.18s ease;
    }

    .admin-todo-card:hover {
        transform: translateX(3px);
        background: rgba(255,255,255,0.96);
        box-shadow: 0 12px 24px rgba(0, 31, 63, 0.12);
    }

    .admin-click-panel {
        cursor: pointer;
        transition: transform 0.18s ease, box-shadow 0.18s ease, background 0.18s ease;
    }

    .admin-click-panel:hover {
        transform: translateY(-2px);
        box-shadow: 0 20px 50px rgba(0, 31, 63, 0.22);
        background: rgba(255,255,255,0.15);
    }

    .admin-todo-icon {
        width: 42px;
        height: 42px;
        display: grid;
        place-items: center;
        border-radius: 14px;
        background: rgba(0,139,139,0.13);
        color: #005b67;
        font-weight: 900;
    }

    .admin-todo-title {
        color: #092f3c;
        font-size: 15px;
        font-weight: 900;
        line-height: 1.4;
    }

    .admin-todo-meta {
        margin-top: 4px;
        color: rgba(9,47,60,0.62);
        font-size: 12px;
        font-weight: 800;
    }

    .admin-todo-count {
        min-width: 48px;
        text-align: center;
        color: #00264d;
        font-size: 28px;
        font-weight: 950;
    }

    @keyframes workbenchFadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }

    @media (max-width: 768px) {
        .login-title {
            font-size: 32px;
        }

        .login-subtitle {
            font-size: 16px;
        }

        .dashboard-main-title {
            font-size: 34px;
        }

        .dashboard-sub-title {
            font-size: 17px;
        }

        .org-hero {
            padding: 24px;
        }

        .org-hero-grid,
        .org-content-grid {
            grid-template-columns: 1fr;
        }

        .org-dashboard-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }

        .org-project-row {
            grid-template-columns: minmax(130px, 1fr) minmax(150px, 1fr);
        }

        .org-project-row > div:nth-child(3),
        .org-project-row > div:nth-child(4) {
            display: none;
        }
    }
    
    /* 卡片样式 */
    .stat-card {
        background: var(--card-bg);
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 8px 30px rgba(0,0,0,0.25);
        text-align: center;
        transition: transform 0.25s ease, box-shadow 0.25s ease;
        cursor: pointer;
        color: var(--text-on-dark) !important;
        border: 1px solid rgba(255,255,255,0.06);
        overflow: hidden;
        backdrop-filter: blur(4px);
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    }
    
    .stat-number {
        font-size: 34px;
        font-weight: 800;
        color: rgba(255,255,255,0.95);
    }
    
    .stat-label {
        font-size: 13px;
        color: rgba(255,255,255,0.85);
        margin-top: 6px;
    }
    
    .stat-icon {
        font-size: 36px;
        margin-bottom: 8px;
        opacity: 0.95;
    }
    
    /* 按钮样式 */
    .stButton>button, .stDownloadButton>button, [data-testid="stPopover"]>button {
        width: 100%;
        border-radius: 10px;
        padding: 10px 18px;
        font-weight: 700;
        transition: all 0.18s ease;
        color: var(--text-on-dark) !important;
        background: linear-gradient(90deg, var(--deep-blue), var(--mid-blue)) !important;
        border: none !important;
        box-shadow: 0 6px 18px rgba(0,0,0,0.25);
    }
    
    /* 强制按钮内文字颜色为白色，适配部分情况下的嵌套元素 */
    .stButton>button * , .stDownloadButton>button *, [data-testid="stPopover"]>button * {
        color: var(--text-on-dark) !important;
    }
    
    .stButton>button:hover, .stDownloadButton>button:hover, [data-testid="stPopover"]>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 30px rgba(0,0,0,0.28);
        color: var(--text-on-dark) !important;
        filter: brightness(1.03);
    }

    /* 侧边栏内按钮使用蓝绿色背景以区分（按钮本身样式保持一致） */
    [data-testid="stSidebar"] .stButton>button, [data-testid="stSidebar"] .stDownloadButton>button {
        background: linear-gradient(90deg, var(--main-bg), var(--mid-blue)) !important;
        color: var(--text-on-dark) !important;
        font-weight: 700;
        box-shadow: none !important;
        border: 1px solid rgba(255,255,255,0.06) !important;
    }

    [data-testid="stSidebar"] .stButton>button:hover, [data-testid="stSidebar"] .stDownloadButton>button:hover {
        filter: brightness(1.06);
        box-shadow: 0 10px 24px rgba(0,0,0,0.28) !important;
    }
    
    /* 数据大盘/模块卡片字体，深色背景上使用白色文字以保证可读性 */
    .stat-card, .stat-card * {
        color: var(--text-on-dark) !important;
    }
    
    /* 侧边栏样式 */
    /* 侧边栏：深蓝色背景，侧边栏上文字保持白色高对比 */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, var(--deep-blue) 0%, var(--mid-blue) 100%) !important;
        color: var(--text-on-dark) !important;
        border-right: 1px solid rgba(255,255,255,0.04);
    }

    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] .css-1d391kg {
        color: var(--text-on-dark) !important;
    }
    
    /* 表格样式 */
    .dataframe {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 8px 24px rgba(0,0,0,0.18);
        border: 1px solid rgba(255,255,255,0.04);
        background: rgba(255,255,255,0.02);
    }
    
    /* 成功/错误消息 */
    .element-container .stSuccess, .element-container .stError, 
    .element-container .stWarning, .element-container .stInfo {
        border-radius: 10px;
        padding: 12px;
        color: var(--text-on-dark) !important;
    }
    
    /* 输入框样式 */
    .stTextInput>div>div>input, .stSelectbox>div>div>select, .stTextArea>div>div>textarea {
        border-radius: 8px;
        background: rgba(255,255,255,0.96) !important;
        color: #000000 !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
        padding: 8px;
    }

    /* 数字输入样式（stNumberInput） */
    .stNumberInput>div>div>input, input[type="number"] {
        border-radius: 8px !important;
        background: rgba(255,255,255,0.96) !important;
        color: #000000 !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
        padding: 6px 8px !important;
    }

    /* 下拉选择、按钮组等保持深色主题 */
    .stSelectbox>div>div>select, .stMultiselect>div>div>select {
        background: rgba(255,255,255,0.96) !important;
        color: #000000 !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
    }

    /* 指标管理局部覆盖：使当前指标下的输入/占位文字为黑色以提升可读性 */
    .indicator-row input, .indicator-row textarea, .indicator-row select {
        color: #000 !important;
        background: rgba(255,255,255,0.96) !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
    }

    .indicator-row input::placeholder, .indicator-row textarea::placeholder {
        color: rgba(0,0,0,0.5) !important;
    }

    /* 搜索框空内容提示：仅“请输入内容”占位符显示红色 */
    input[placeholder="请输入内容"]::placeholder {
        color: #ff3b30 !important;
        opacity: 1 !important;
    }
    input[placeholder="请输入内容"]::-webkit-input-placeholder {
        color: #ff3b30 !important;
    }
    input[placeholder="请输入内容"]:-ms-input-placeholder {
        color: #ff3b30 !important;
    }

    /* 登录容器内输入框颜色优化，确保密码点/占位符可读 */
    .login-container .stTextInput>div>div>input,
    .login-container .stTextInput>div>div>input::placeholder,
    .login-container .stTextInput>div>div>input:-ms-input-placeholder {
        color: rgba(255,255,255,0.95) !important;
        background: rgba(0,0,0,0.18) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
    }

    .login-container .stTextInput>div>div>input::placeholder {
        color: rgba(255,255,255,0.75) !important;
    }

    /* 保证审批行内布局：按钮与输入垂直居中并保持行内 */
    .stButton>button, .stTextInput, .stTextInput>div, .stTextInput>div>div {
        vertical-align: middle !important;
        display: inline-block !important;
    }

    /* 当审批卡片使用浅色背景时，保证其内部元素也采用主题色文本 */
    div[style*="background: rgba(255,255,255,0.03)"] h4 { color: rgba(255,255,255,0.98) !important; }
    div[style*="background: rgba(255,255,255,0.03)"] p { color: rgba(255,255,255,0.9) !important; }
    
    /* 文件夹样式 */
    .folder-item {
        padding: 10px 15px;
        margin: 6px 0;
        background: rgba(255,255,255,0.03);
        border-radius: 8px;
        cursor: pointer;
        transition: all 0.18s ease;
        color: var(--text-on-dark) !important;
    }

    .folder-item:hover {
        background: rgba(255,255,255,0.06);
        transform: translateY(-2px);
    }
    
    /* 进度条样式 */
    .stProgress > div > div > div {
        border-radius: 10px;
        background: linear-gradient(90deg, var(--accent), rgba(255,255,255,0.6));
    }
    
    /* 标签页样式 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
        background: rgba(255,255,255,0.03);
        color: var(--text-on-dark) !important;
    }
    
    /* 徽章样式 */
    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 700;
        color: var(--text-on-dark) !important;
    }
    
    .badge-success { background: #d4edda; color: #155724; }
    .badge-warning { background: #fff3cd; color: #856404; }
    .badge-danger { background: #f8d7da; color: #721c24; }
    .badge-info { background: #d1ecf1; color: #0c5460; }
    
    /* 滚动条样式 */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 4px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: rgba(0,0,0,0.25);
        border-radius: 4px;
    }

    ::-webkit-scrollbar-thumb:hover {
        background: rgba(0,0,0,0.35);
    }
    
    /* 右上角悬浮菜单样式 */
    div[data-testid="stPopover"] {
        width: 100% !important;
        display: flex !important;
        justify-content: flex-end !important;
        margin: 6px 0 18px 0 !important;
    }
    
    
    /* 悬浮菜单（个人账户）按钮样式调整 - 确保文字颜色清晰 */
    div[data-testid="stPopover"] button[data-testid="baseButton-secondary"] {
        width: min(280px, 100%) !important;
        min-height: 44px !important;
        padding: 8px 14px !important;
        background: linear-gradient(135deg, #ffffff 0%, #eef8f6 100%) !important;
        border: 2px solid #0e5f70 !important;
        border-radius: 12px !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        box-shadow: 0 10px 24px rgba(0, 34, 58, 0.22) !important;
    }
    div[data-testid="stPopover"] button[data-testid="baseButton-secondary"],
    div[data-testid="stPopover"] button[data-testid="baseButton-secondary"] * {
        color: #153746 !important;
        font-weight: 700 !important;
        -webkit-text-fill-color: #153746 !important; /* 强制覆盖 Webkit 的文字颜色 */
        opacity: 1 !important;
    }
    div[data-testid="stPopover"] button[data-testid="baseButton-secondary"]:hover {
        transform: translateY(-1px) !important;
        filter: none !important;
        box-shadow: 0 12px 28px rgba(0, 34, 58, 0.28) !important;
        border-color: #0b4f5d !important;
    }
    
    /* 优化悬浮菜单内部（如退出登录）的按钮大小，并固定其宽度与外部一致 */
    div[data-testid="stPopoverBody"] {
        min-width: 260px !important;
        max-width: min(260px, calc(100vw - 32px)) !important;
        width: min(260px, calc(100vw - 32px)) !important;
        padding: 8px !important;
        border-radius: 14px !important;
    }
    div[data-testid="stPopoverBody"] button {
        padding: 8px 12px !important;
        min-height: 40px !important;
        font-size: 14px !important;
        line-height: 1.5 !important;
        width: 100% !important;
        white-space: nowrap !important;
    }

    /* 可视化大屏样式 */
    .viz-hero {
        background: linear-gradient(120deg, rgba(3, 44, 74, 0.92), rgba(0, 109, 121, 0.82));
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 16px;
        padding: 18px 22px;
        margin: 4px 0 14px 0;
        box-shadow: 0 14px 30px rgba(0, 20, 35, 0.26);
    }

    .viz-hero-title {
        color: #dff6ff;
        font-size: 28px;
        font-weight: 800;
        margin: 0;
        letter-spacing: 0.4px;
    }

    .viz-hero-subtitle {
        color: rgba(224, 247, 255, 0.9);
        margin: 8px 0 0 0;
        font-size: 14px;
        line-height: 1.55;
    }

    .viz-kpi-card {
        background: linear-gradient(150deg, rgba(5, 38, 66, 0.9), rgba(4, 88, 103, 0.82));
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 14px;
        padding: 14px 16px;
        box-shadow: 0 10px 20px rgba(0, 18, 34, 0.24);
        min-height: 112px;
    }

    .viz-kpi-label {
        color: rgba(214, 243, 255, 0.9);
        font-size: 13px;
        font-weight: 650;
        margin: 0 0 8px 0;
    }

    .viz-kpi-value {
        color: #ffffff;
        font-size: 34px;
        font-weight: 800;
        line-height: 1.1;
        margin: 0;
    }

    .viz-kpi-extra {
        color: rgba(191, 238, 255, 0.88);
        font-size: 12px;
        margin-top: 8px;
    }

    .viz-section-title {
        color: #d8f5ff;
        font-size: 24px;
        font-weight: 800;
        margin: 12px 0 8px 0;
        letter-spacing: 0.3px;
    }

    .viz-panel {
        background: linear-gradient(155deg, rgba(4, 56, 82, 0.62), rgba(2, 80, 90, 0.44));
        border: 1px solid rgba(166, 230, 255, 0.18);
        border-radius: 14px;
        padding: 12px 14px 10px 14px;
        margin-bottom: 12px;
        min-height: 70px;
        box-shadow: 0 10px 24px rgba(0, 27, 39, 0.2);
        backdrop-filter: blur(2px);
    }

    .viz-panel-title {
        margin: 0;
        color: #f4fcff;
        font-size: 16px;
        font-weight: 760;
        letter-spacing: 0.25px;
    }

    .viz-panel-subtitle {
        margin: 4px 0 0 0;
        color: rgba(224, 248, 255, 0.85);
        font-size: 12px;
        line-height: 1.45;
    }

    /* 强制改密页 */
    .password-change-header {
        margin: 24px 0 18px 0;
    }
    .password-change-title {
        color: #082f49;
        font-size: 38px;
        font-weight: 850;
        line-height: 1.15;
        margin: 0 0 10px 0;
    }
    .password-change-subtitle {
        color: rgba(8, 47, 73, 0.78);
        font-size: 15px;
        line-height: 1.6;
        margin: 0;
    }
    .password-change-notice {
        border: 1px solid rgba(8, 145, 178, 0.22);
        background: rgba(236, 253, 245, 0.66);
        border-radius: 8px;
        padding: 12px 14px;
        color: #064e3b;
        font-size: 15px;
        line-height: 1.55;
        margin: 6px 0 18px 0;
    }
    .password-change-card-title {
        color: #083344;
        font-size: 24px;
        font-weight: 800;
        margin: 6px 0 6px 0;
    }
    .password-change-rule {
        color: rgba(8, 47, 73, 0.72);
        font-size: 13px;
        line-height: 1.5;
        margin-bottom: 12px;
    }

    /* 全局反馈提示：红色高对比 + 图标 + 加粗 */
    .flash-feedback {
        margin: 6px 0 14px 0;
        padding: 12px 14px;
        border-radius: 10px;
        border: 2px solid rgba(255, 59, 48, 0.62);
        background: rgba(255, 59, 48, 0.10);
        color: #ff3b30 !important;
        display: flex;
        align-items: center;
        gap: 8px;
        overflow: hidden;
        animation: flash-feedback-autohide 5.35s ease forwards;
    }
    .flash-feedback,
    .flash-feedback * {
        color: #ff3b30 !important;
    }
    .flash-feedback-icon {
        font-size: 18px;
        line-height: 1;
    }
    .flash-feedback-text {
        font-size: 18px;
        font-weight: 800;
        line-height: 1.45;
    }
    @keyframes flash-feedback-autohide {
        0%, 92% {
            opacity: 1;
            max-height: 80px;
            margin-top: 6px;
            margin-bottom: 14px;
            padding-top: 12px;
            padding-bottom: 12px;
            border-width: 2px;
        }
        100% {
            opacity: 0;
            max-height: 0;
            margin-top: 0;
            margin-bottom: 0;
            padding-top: 0;
            padding-bottom: 0;
            border-width: 0;
            pointer-events: none;
            visibility: hidden;
        }
    }

    @media (max-width: 768px) {
        div[data-testid="stPopover"] {
            margin: 4px 0 14px 0 !important;
        }
        div[data-testid="stPopover"] button[data-testid="baseButton-secondary"] {
            width: min(240px, 100%) !important;
        }
        div[data-testid="stPopoverBody"] {
            min-width: min(240px, calc(100vw - 24px)) !important;
            max-width: min(240px, calc(100vw - 24px)) !important;
            width: min(240px, calc(100vw - 24px)) !important;
        }
    }
    </style>
    ''', unsafe_allow_html=True)


def install_page_switch_mask():
    """点击导航时短暂遮住主内容，避开新旧 DOM 混合的一瞬间。"""
    navigation_labels = [
        "数据大盘", "机构管理", "账号管理", "项目审核", "日志查看", "数据导出",
        "审批待办", "消息通知", "项目智库", "可视化大屏",
        "工作台", "信息维护", "子账号管理", "项目管理", "待办事项",
        "查看机构", "查看用户", "查看项目", "去处理审批",
    ]
    labels_json = json.dumps(navigation_labels, ensure_ascii=False)
    components.html(f"""
    <script>
    (function() {{
        const doc = window.parent.document;
        if (!doc || doc.__celinkPageSwitchMaskInstalled) {{
            return;
        }}

        doc.__celinkPageSwitchMaskInstalled = true;
        const labels = {labels_json};
        const root = doc.documentElement;

        function clearSwitchTimer() {{
            if (doc.__celinkPageSwitchTimer) {{
                window.parent.clearTimeout(doc.__celinkPageSwitchTimer);
                doc.__celinkPageSwitchTimer = null;
            }}
        }}

        function shouldMask(button) {{
            const text = (button.innerText || button.textContent || "").replace(/\\s+/g, " ").trim();
            if (!text) {{
                return false;
            }}
            return labels.some(label => text.includes(label));
        }}

        doc.addEventListener("click", function(event) {{
            const button = event.target && event.target.closest ? event.target.closest("button") : null;
            if (!button || !shouldMask(button)) {{
                return;
            }}

            root.classList.add("celink-page-switching");
            clearSwitchTimer();
            doc.__celinkPageSwitchTimer = window.parent.setTimeout(function() {{
                root.classList.remove("celink-page-switching");
                doc.__celinkPageSwitchTimer = null;
            }}, 180);
        }}, true);
    }})();
    </script>
    """, height=0, width=0)


def render_page_ready_marker():
    """页面主体渲染到末尾后兜底解除切换短遮罩。"""
    components.html("""
    <script>
    (function() {
        const doc = window.parent.document;
        if (!doc) {
            return;
        }

        window.parent.requestAnimationFrame(function() {
            window.parent.requestAnimationFrame(function() {
                doc.documentElement.classList.remove("celink-page-switching");
                if (doc.__celinkPageSwitchTimer) {
                    window.parent.clearTimeout(doc.__celinkPageSwitchTimer);
                    doc.__celinkPageSwitchTimer = null;
                }
            });
        });
    })();
    </script>
    """, height=0, width=0)

# ==================== 数据库初始化 ====================
def set_app_meta_with_cursor(cursor, key, value):
    """在数据库初始化事务内写入 app_meta，避免迁移中途另开连接。"""
    cursor.execute(
        '''
        INSERT INTO app_meta (key, value, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(key) DO UPDATE SET
            value = excluded.value,
            updated_at = excluded.updated_at
        ''',
        (key, str(value), datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    )


def normalize_user_unique_contact_fields(cursor):
    """清理历史空值，便于手机号/邮箱使用部分唯一索引。"""
    cursor.execute("UPDATE users SET phone = TRIM(phone) WHERE phone IS NOT NULL")
    cursor.execute("UPDATE users SET email = TRIM(email) WHERE email IS NOT NULL")
    cursor.execute("UPDATE users SET phone = NULL WHERE phone IS NOT NULL AND phone = ''")
    cursor.execute("UPDATE users SET email = NULL WHERE email IS NOT NULL AND email = ''")


def create_user_contact_unique_index(cursor, column, index_name):
    """仅在历史数据无重复时创建唯一索引；有重复则记录阻塞原因，不中断启动。"""
    cursor.execute(
        f'''
        SELECT {column}, COUNT(*) AS cnt
        FROM users
        WHERE {column} IS NOT NULL
        GROUP BY {column}
        HAVING COUNT(*) > 1
        LIMIT 5
        '''
    )
    duplicates = cursor.fetchall()
    meta_key = f"{index_name}_blocked"
    if duplicates:
        blocked_values = ", ".join(f"{row[0]}({row[1]})" for row in duplicates)
        set_app_meta_with_cursor(cursor, meta_key, blocked_values)
        return False

    cursor.execute(
        f"CREATE UNIQUE INDEX IF NOT EXISTS {index_name} ON users({column}) WHERE {column} IS NOT NULL"
    )
    set_app_meta_with_cursor(cursor, meta_key, "")
    return True


def deduplicate_project_steps(cursor):
    """清理历史重复 Gate，保留业务进度最高的一条，并迁移文件引用。"""
    cursor.execute(
        '''
        SELECT project_id, stage, COUNT(*) AS cnt
        FROM project_steps
        GROUP BY project_id, stage
        HAVING COUNT(*) > 1
        '''
    )
    groups = cursor.fetchall()
    removed_count = 0

    for project_id, stage, _ in groups:
        cursor.execute(
            '''
            SELECT id
            FROM project_steps
            WHERE project_id = ? AND stage = ?
            ORDER BY
                CASE status
                    WHEN 'approved' THEN 0
                    WHEN 'submitted' THEN 1
                    WHEN 'rejected' THEN 2
                    WHEN 'pending' THEN 3
                    ELSE 4
                END,
                COALESCE(reviewed_at, submitted_at, created_at) DESC,
                id ASC
            ''',
            (project_id, stage)
        )
        step_ids = [row[0] for row in cursor.fetchall()]
        if len(step_ids) <= 1:
            continue

        keep_id = step_ids[0]
        duplicate_ids = step_ids[1:]
        placeholders = ",".join("?" for _ in duplicate_ids)
        cursor.execute(
            f"UPDATE project_files SET step_id = ? WHERE step_id IN ({placeholders})",
            (keep_id, *duplicate_ids)
        )
        cursor.execute(
            f"DELETE FROM project_steps WHERE id IN ({placeholders})",
            duplicate_ids
        )
        removed_count += len(duplicate_ids)

    if removed_count:
        set_app_meta_with_cursor(cursor, "deduplicate_project_steps_v1_removed", removed_count)


def init_database():
    """初始化数据库表结构"""
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    
    # 用户表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'org_user',
            org_id INTEGER,
            phone TEXT,
            email TEXT,
            real_name TEXT,
            status TEXT DEFAULT 'active',
            must_change_password INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (org_id) REFERENCES organizations(id)
        )
    ''')
    
    # 机构表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS organizations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            org_type TEXT,
            credit_code TEXT UNIQUE,
            legal_person TEXT,
            contact_person TEXT,
            contact_phone TEXT,
            contact_email TEXT,
            address TEXT,
            description TEXT,
            status TEXT DEFAULT 'active',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 项目表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            org_id INTEGER NOT NULL,
            category TEXT,
            subcategory TEXT,
            description TEXT,
            current_stage INTEGER DEFAULT 1,
            status TEXT DEFAULT 'pending',
            created_by INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (org_id) REFERENCES organizations(id),
            FOREIGN KEY (created_by) REFERENCES users(id)
        )
    ''')
    
    # 项目阶段表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS project_steps (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            stage INTEGER NOT NULL,
            status TEXT DEFAULT 'pending',
            submitted_by INTEGER,
            submitted_at TIMESTAMP,
            reviewed_by INTEGER,
            reviewed_at TIMESTAMP,
            review_comment TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (project_id) REFERENCES projects(id),
            FOREIGN KEY (submitted_by) REFERENCES users(id),
            FOREIGN KEY (reviewed_by) REFERENCES users(id)
        )
    ''')
    
    # 项目文件表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS project_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            step_id INTEGER,
            title TEXT NOT NULL,
            file_type TEXT,
            category TEXT,
            subcategory TEXT,
            file_path TEXT,
            file_name TEXT,
            file_size INTEGER,
            publish_org TEXT,
            description TEXT,
            upload_by INTEGER,
            upload_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            approval_status TEXT DEFAULT 'pending',
            approved_by INTEGER,
            approved_at TIMESTAMP,
            approval_comment TEXT,
            FOREIGN KEY (project_id) REFERENCES projects(id),
            FOREIGN KEY (step_id) REFERENCES project_steps(id),
            FOREIGN KEY (upload_by) REFERENCES users(id)
        )
    ''')
    
    # 主评人表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS evaluators (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id INTEGER NOT NULL,
            account_user_id INTEGER,
            name TEXT NOT NULL,
            title TEXT,
            specialty TEXT,
            phone TEXT,
            email TEXT,
            status TEXT DEFAULT 'active',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (org_id) REFERENCES organizations(id),
            FOREIGN KEY (account_user_id) REFERENCES users(id)
        )
    ''')
    
    # 业绩记录表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS achievements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content TEXT,
            achievement_date DATE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (org_id) REFERENCES organizations(id)
        )
    ''')
    
    # 培训记录表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS trainings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            trainer TEXT,
            training_date DATE,
            duration INTEGER,
            participants INTEGER,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (org_id) REFERENCES organizations(id)
        )
    ''')
    
    # 指标库表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS indicator_library (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            subcategory TEXT,
            indicator_name TEXT NOT NULL,
            weight REAL DEFAULT 10,
            description TEXT,
            max_score INTEGER DEFAULT 100,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 政策文件表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS policy_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            file_path TEXT,
            file_name TEXT,
            file_size INTEGER,
            upload_by INTEGER,
            upload_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (upload_by) REFERENCES users(id)
        )
    ''')
    
    # 待办事项表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS todos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content TEXT,
            status TEXT DEFAULT 'pending',
            priority TEXT DEFAULT 'medium',
            related_type TEXT,
            related_id INTEGER,
            due_date DATE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            completed_at TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    ''')

    # 机构信息变更申请表：子账号提交，机构主账号在待办中审批
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS org_info_update_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            org_id INTEGER NOT NULL,
            submitted_by INTEGER NOT NULL,
            approver_id INTEGER,
            status TEXT DEFAULT 'pending',
            old_data TEXT,
            new_data TEXT,
            review_comment TEXT,
            submitted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            reviewed_at TIMESTAMP,
            FOREIGN KEY (org_id) REFERENCES organizations(id),
            FOREIGN KEY (submitted_by) REFERENCES users(id),
            FOREIGN KEY (approver_id) REFERENCES users(id)
        )
    ''')
    
    # 消息通知表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content TEXT,
            msg_type TEXT DEFAULT 'system',
            is_read INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    ''')
    
    # 操作日志表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            org_name TEXT,
            action TEXT,
            module TEXT,
            ip_address TEXT,
            details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
    ''')
    
    # 文件评估表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS file_evaluations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id INTEGER NOT NULL,
            indicator_id INTEGER NOT NULL,
            score REAL,
            comment TEXT,
            evaluated_by INTEGER,
            evaluated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (file_id) REFERENCES project_files(id),
            FOREIGN KEY (indicator_id) REFERENCES indicator_library(id),
            FOREIGN KEY (evaluated_by) REFERENCES users(id)
        )
    ''')

    # 应用元数据表：记录一次性迁移/初始化状态，避免每次页面进入都重复扫描
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS app_meta (
            key TEXT PRIMARY KEY,
            value TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute("PRAGMA table_info(users)")
    user_columns = {row[1] for row in cursor.fetchall()}
    if 'must_change_password' not in user_columns:
        cursor.execute("ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0")

    normalize_user_unique_contact_fields(cursor)

    # 历史库中如果仍有人使用系统默认密码，强制其下次登录后修改。
    cursor.execute("SELECT id, role, password_hash FROM users WHERE status = 'active'")
    for user_id, role, password_hash in cursor.fetchall():
        default_password = get_default_password_for_role(role)
        try:
            if check_password(default_password, password_hash):
                cursor.execute(
                    "UPDATE users SET must_change_password = 1 WHERE id = ?",
                    (user_id,)
                )
        except Exception:
            pass

    # 旧版本创建的机构子账号没有 must_change_password 标记；补一次，确保子账号首次进入也要改密。
    cursor.execute("SELECT value FROM app_meta WHERE key = 'force_org_user_password_change_v1'")
    if not cursor.fetchone():
        cursor.execute(
            "UPDATE users SET must_change_password = 1 WHERE role = 'org_user' AND status = 'active'"
        )
        cursor.execute(
            '''
            INSERT INTO app_meta (key, value, updated_at)
            VALUES ('force_org_user_password_change_v1', '1', ?)
            ON CONFLICT(key) DO UPDATE SET
                value = excluded.value,
                updated_at = excluded.updated_at
            ''',
            (datetime.now().strftime('%Y-%m-%d %H:%M:%S'),)
        )

    # 旧库补列：主评人可绑定本机构主账号或子账号，多个主评人可共用一个账号
    cursor.execute("PRAGMA table_info(evaluators)")
    evaluator_columns = {row[1] for row in cursor.fetchall()}
    if 'account_user_id' not in evaluator_columns:
        cursor.execute("ALTER TABLE evaluators ADD COLUMN account_user_id INTEGER")

    cursor.execute("PRAGMA table_info(todos)")
    todo_columns = {row[1] for row in cursor.fetchall()}
    if 'related_type' not in todo_columns:
        cursor.execute("ALTER TABLE todos ADD COLUMN related_type TEXT")
    if 'related_id' not in todo_columns:
        cursor.execute("ALTER TABLE todos ADD COLUMN related_id INTEGER")

    cursor.execute('''
        UPDATE todos
        SET priority = 'low',
            content = REPLACE(COALESCE(content, ''), '请审批', '已审批')
        WHERE related_type = 'org_info_update' AND status = 'completed'
    ''')

    deduplicate_project_steps(cursor)
    create_user_contact_unique_index(cursor, "phone", "idx_users_phone_unique")
    create_user_contact_unique_index(cursor, "email", "idx_users_email_unique")

    # 高频查询索引：减少侧栏计数、待办、消息、审批、文件列表等页面的扫描开销
    cursor.executescript('''
        CREATE INDEX IF NOT EXISTS idx_users_org_status ON users(org_id, status);
        CREATE INDEX IF NOT EXISTS idx_users_role_status ON users(role, status);
        CREATE INDEX IF NOT EXISTS idx_projects_org_status ON projects(org_id, status);
        CREATE INDEX IF NOT EXISTS idx_projects_status_created_at ON projects(status, created_at DESC);
        CREATE UNIQUE INDEX IF NOT EXISTS idx_project_steps_project_stage_unique ON project_steps(project_id, stage);
        CREATE INDEX IF NOT EXISTS idx_project_steps_project_status ON project_steps(project_id, status);
        CREATE INDEX IF NOT EXISTS idx_project_steps_status ON project_steps(status);
        CREATE INDEX IF NOT EXISTS idx_project_files_step_id ON project_files(step_id);
        CREATE INDEX IF NOT EXISTS idx_project_files_project_approval ON project_files(project_id, approval_status);
        CREATE INDEX IF NOT EXISTS idx_project_files_approval_upload_at ON project_files(approval_status, upload_at DESC);
        CREATE INDEX IF NOT EXISTS idx_messages_user_read_created_at ON messages(user_id, is_read, created_at DESC);
        CREATE INDEX IF NOT EXISTS idx_logs_created_at ON logs(created_at DESC);
        CREATE INDEX IF NOT EXISTS idx_todos_user_status_created_at ON todos(user_id, status, created_at DESC);
        CREATE INDEX IF NOT EXISTS idx_indicator_library_category_subcategory ON indicator_library(category, subcategory);
        CREATE INDEX IF NOT EXISTS idx_evaluators_org_account ON evaluators(org_id, account_user_id);
        CREATE INDEX IF NOT EXISTS idx_todos_related ON todos(related_type, related_id);
        CREATE INDEX IF NOT EXISTS idx_org_info_update_requests_org_status ON org_info_update_requests(org_id, status);
    ''')

    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA synchronous=NORMAL")

    # 迁移历史项目：补齐 G0~G8 Gate，避免旧项目在 G4 后断档
    cursor.execute("SELECT id, current_stage, status FROM projects")
    projects = cursor.fetchall()
    for project_id, current_stage, project_status in projects:
        cursor.execute("SELECT stage, status FROM project_steps WHERE project_id = ?", (project_id,))
        step_rows = cursor.fetchall()
        existing_stages = {row[0] for row in step_rows}

        for stage in range(1, TOTAL_STAGES + 1):
            if stage not in existing_stages:
                cursor.execute(
                    "INSERT INTO project_steps (project_id, stage, status) VALUES (?, ?, 'pending')",
                    (project_id, stage)
                )

        # 纠正异常 current_stage
        safe_stage = current_stage if isinstance(current_stage, int) else 1
        if safe_stage < 1:
            safe_stage = 1
        if safe_stage > TOTAL_STAGES:
            safe_stage = TOTAL_STAGES
        if safe_stage != current_stage:
            cursor.execute(
                "UPDATE projects SET current_stage = ?, updated_at = ? WHERE id = ?",
                (safe_stage, datetime.now(), project_id)
            )

        # 旧项目若被提前标记 completed（但未走完 9 个 Gate），回调为进行中
        cursor.execute("SELECT status FROM project_steps WHERE project_id = ? ORDER BY stage", (project_id,))
        statuses = [row[0] for row in cursor.fetchall()]
        all_approved = len(statuses) == TOTAL_STAGES and all(s == 'approved' for s in statuses)
        if project_status == 'completed' and not all_approved:
            cursor.execute(
                "UPDATE projects SET status = ?, updated_at = ? WHERE id = ?",
                ('in_progress', datetime.now(), project_id)
            )
    
    conn.commit()
    
    # 初始化默认数据
    init_default_data(conn)
    
    conn.close()

def init_default_data(conn):
    """初始化默认数据"""
    cursor = conn.cursor()
    
    # 检查是否已有超级管理员
    cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'super_admin'")
    if cursor.fetchone()[0] == 0:
        # 创建超级管理员
        password_hash = hash_password(ADMIN_DEFAULT_PASSWORD)
        cursor.execute('''
            INSERT INTO users (username, password_hash, role, real_name, status, must_change_password)
            VALUES (?, ?, ?, ?, ?, 1)
        ''', ('admin', password_hash, 'super_admin', '超级管理员', 'active'))
    
    # 初始化指标库（每类10个默认指标）
    cursor.execute("SELECT COUNT(*) FROM indicator_library")
    if cursor.fetchone()[0] == 0:
        default_indicators = []
        for cat_key, cat_val in PROJECT_CATEGORIES.items():
            if cat_key == '0':
                continue
            for sub_key, sub_name in cat_val['subcategories'].items():
                for i in range(1, 11):
                    indicator_name = f"{sub_name}指标{i}"
                    default_indicators.append((cat_key, sub_key, indicator_name, 10, f"{indicator_name}描述", 100))
        
        cursor.executemany('''
            INSERT INTO indicator_library (category, subcategory, indicator_name, weight, description, max_score)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', default_indicators)
    
    conn.commit()

# ==================== 数据库操作 ====================
def get_connection():
    """获取数据库连接"""
    conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA temp_store=MEMORY")
    conn.execute("PRAGMA cache_size=-20000")
    return conn

@st.cache_resource(show_spinner=False)
def get_query_cache_state():
    """全局查询缓存版本号；写操作后递增以触发只读缓存失效"""
    return {"version": 0}

def get_query_cache_version():
    return get_query_cache_state()["version"]

def invalidate_query_cache():
    get_query_cache_state()["version"] += 1

def normalize_query_params(params):
    if params is None:
        return ()
    if isinstance(params, tuple):
        return params
    if isinstance(params, list):
        return tuple(params)
    return params

@st.cache_data(show_spinner=False)
def cached_fetch_query(query, params, version):
    """缓存只读查询结果；version 变化时自动失效"""
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(query, params)
        columns = [description[0] for description in cursor.description]
        results = cursor.fetchall()
        return [dict(zip(columns, row)) for row in results]
    finally:
        conn.close()

def execute_query(query, params=(), fetch=False, commit=False):
    """执行数据库查询"""
    params = normalize_query_params(params)

    if fetch and not commit:
        return cached_fetch_query(query, params, get_query_cache_version())

    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(query, params)
        if commit:
            conn.commit()
            invalidate_query_cache()
            return cursor.lastrowid
        if fetch:
            columns = [description[0] for description in cursor.description]
            results = cursor.fetchall()
            return [dict(zip(columns, row)) for row in results]
        return cursor
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        conn.close()

def can_access_client(role, client_type):
    """校验账号角色是否允许登录当前客户端。"""
    if client_type == 'auto':
        return role in {allowed_role for roles in CLIENT_ALLOWED_ROLES.values() for allowed_role in roles}
    return role in CLIENT_ALLOWED_ROLES.get(client_type, set())

def resolve_client_type(role, client_type):
    """自动识别登录端类型，避免 auto 参与后续权限判断。"""
    if client_type != 'auto':
        return client_type
    return 'admin' if role == 'super_admin' else 'org'

def can_access_page(role, page):
    """集中页面权限判断，所有菜单和路由入口必须走这里。"""
    return role in PAGE_ACCESS.get(page, set())

def can_manage_org_evaluators(role):
    """只有机构主账号可维护主评人及其账号关联。"""
    return role == 'org_admin'

def get_menu_items_for_role(role):
    """返回当前角色可见菜单项。"""
    return [
        item for item in PAGE_MENU.get(role, [])
        if can_access_page(role, item[2])
    ]

def get_default_page_for_role(role):
    """当前角色默认落地页。"""
    items = get_menu_items_for_role(role)
    return items[0][2] if items else None

def ensure_current_page_access(role):
    """防止通过 session_state/current_page 绕过侧边栏权限。"""
    current_page = st.session_state.get('current_page') or get_default_page_for_role(role)
    if not can_access_page(role, current_page):
        default_page = get_default_page_for_role(role)
        st.session_state['current_page'] = default_page
        if default_page and can_access_page(role, default_page):
            set_flash_message("已切换到当前账号可访问的默认页面", level='warning')
            return default_page, True
        return default_page, False
    st.session_state['current_page'] = current_page
    return current_page, True

def render_access_denied():
    st.error("当前账号没有权限访问该页面")

def set_flash_message(message, level='success'):
    """设置跨 rerun 的全局提示消息。"""
    st.session_state['_flash_message'] = {
        'level': level,
        'message': str(message)
    }

def notify_and_rerun(message, level='success'):
    """先记录提示，再重跑页面，避免提示被瞬间吞掉。"""
    set_flash_message(message, level=level)
    st.rerun()

def render_flash_message():
    """渲染并消费一次性提示消息。"""
    flash = st.session_state.pop('_flash_message', None)
    if not flash:
        return
    level = str(flash.get('level', 'info')).strip().lower()
    message = str(flash.get('message', '')).strip()
    if not message:
        return

    icon = "✅" if level == 'success' else "❌"
    safe_message = html.escape(message)

    st.markdown(
        f"""
        <div class="flash-feedback">
            <span class="flash-feedback-icon">{icon}</span>
            <span class="flash-feedback-text"><strong>{safe_message}</strong></span>
        </div>
        """,
        unsafe_allow_html=True
    )

@st.cache_resource(show_spinner=False)
def ensure_app_initialized(schema_version=APP_SCHEMA_VERSION):
    """仅在进程生命周期内初始化一次数据库结构与索引"""
    init_database()
    invalidate_query_cache()
    return True

def get_app_meta(key, default=None):
    rows = execute_query("SELECT value FROM app_meta WHERE key = ?", (key,), fetch=True)
    return rows[0]['value'] if rows else default

def set_app_meta(key, value):
    execute_query(
        '''
        INSERT INTO app_meta (key, value, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(key) DO UPDATE SET
            value = excluded.value,
            updated_at = excluded.updated_at
        ''',
        (key, str(value), datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        commit=True
    )

@st.cache_data(show_spinner=False)
def get_base64_file(path):
    """缓存静态文件的 base64 内容，避免每次重跑都重复读盘"""
    if not path or not os.path.exists(path):
        return None
    with open(path, "rb") as file_obj:
        return base64.b64encode(file_obj.read()).decode()

@st.cache_data(show_spinner=False, max_entries=128)
def get_file_bytes(path, modified_ts, file_size):
    """缓存下载文件内容，避免列表页反复读盘"""
    with open(path, "rb") as file_obj:
        return file_obj.read()

def render_cached_download_button(file_path, file_name, key, label="📥 下载", mime=None):
    """渲染带缓存的数据下载按钮"""
    if file_path and os.path.exists(file_path):
        try:
            file_bytes = get_file_bytes(file_path, os.path.getmtime(file_path), os.path.getsize(file_path))
            st.download_button(label=label, data=file_bytes, file_name=file_name, mime=mime, key=key)
        except Exception:
            st.write("读取失败")
    else:
        st.write("文件丢失")

def get_sidebar_badges(user_id):
    rows = execute_query(
        '''
        SELECT
            (SELECT COUNT(*) FROM project_steps WHERE status = 'pending') as pending_approvals,
            (SELECT COUNT(*) FROM project_files WHERE approval_status = 'pending') as pending_files,
            (SELECT COUNT(*) FROM todos WHERE user_id = ? AND status = 'pending') as pending_todos,
            (SELECT COUNT(*) FROM messages WHERE user_id = ? AND is_read = 0) as unread_messages
        ''',
        (user_id, user_id),
        fetch=True
    )
    return rows[0] if rows else {
        'pending_approvals': 0,
        'pending_files': 0,
        'pending_todos': 0,
        'unread_messages': 0
    }

def get_admin_dashboard_snapshot():
    rows = execute_query(
        '''
        SELECT
            (SELECT COUNT(*) FROM organizations) as org_count,
            (SELECT COUNT(*) FROM users WHERE status = 'active') as user_count,
            (
                SELECT COUNT(*)
                FROM projects p
                JOIN organizations o ON p.org_id = o.id
                WHERE o.status = 'active' AND p.status IN ('pending', 'in_progress')
            ) as active_projects,
            (
                SELECT COUNT(*)
                FROM projects p
                JOIN organizations o ON p.org_id = o.id
                WHERE o.status = 'inactive' AND p.status = 'completed'
            ) as completed_projects,
            (SELECT COUNT(*) FROM project_steps WHERE status = 'pending') as pending_approvals,
            (SELECT COUNT(*) FROM project_files WHERE approval_status = 'pending') as pending_files
        ''',
        fetch=True
    )
    return rows[0] if rows else {
        'org_count': 0,
        'user_count': 0,
        'active_projects': 0,
        'completed_projects': 0,
        'pending_approvals': 0,
        'pending_files': 0
    }

def get_org_dashboard_snapshot(org_id, user_id):
    rows = execute_query(
        '''
        SELECT
            (SELECT COUNT(*) FROM projects WHERE org_id = ?) as project_count,
            (SELECT COUNT(*) FROM projects WHERE org_id = ? AND status IN ('pending', 'in_progress')) as active_count,
            (SELECT COUNT(*) FROM projects WHERE org_id = ? AND status = 'completed') as completed_count,
            (SELECT COUNT(*) FROM project_files WHERE upload_by = ?) as file_count
        ''',
        (org_id, org_id, org_id, user_id),
        fetch=True
    )
    return rows[0] if rows else {
        'project_count': 0,
        'active_count': 0,
        'completed_count': 0,
        'file_count': 0
    }

def add_log(user_id, username, org_name, action, module, details, ip_address=""):
    """添加操作日志"""
    # 使用系统本地时间作为日志时间，避免 SQLite 的 CURRENT_TIMESTAMP 返回 UTC
    created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    execute_query('''
        INSERT INTO logs (user_id, username, org_name, action, module, ip_address, details, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (user_id, username, org_name, action, module, ip_address, details, created_at), commit=True)

def add_message(user_id, title, content, msg_type='system', created_at=None):
    """添加消息通知"""
    if created_at is None:
        created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    execute_query('''
        INSERT INTO messages (user_id, title, content, msg_type, created_at)
        VALUES (?, ?, ?, ?, ?)
    ''', (user_id, title, content, msg_type, created_at), commit=True)

def add_todo(user_id, title, content, priority='medium', due_date=None, related_type=None, related_id=None):
    """添加待办事项"""
    execute_query('''
        INSERT INTO todos (user_id, title, content, priority, due_date, related_type, related_id)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (user_id, title, content, priority, due_date, related_type, related_id), commit=True)

def ensure_project_steps(project_id):
    """确保项目拥有完整的 G0~G8 Gate 记录"""
    steps = execute_query("SELECT stage FROM project_steps WHERE project_id = ?", (project_id,), fetch=True)
    existing_stages = {s['stage'] for s in steps}
    for stage in range(1, TOTAL_STAGES + 1):
        if stage not in existing_stages:
            execute_query(
                "INSERT INTO project_steps (project_id, stage, status) VALUES (?, ?, 'pending')",
                (project_id, stage),
                commit=True
            )

def refresh_project_status(project_id):
    """按完整 Gate 流程刷新项目状态，只有 G0~G8 全部通过才完成"""
    ensure_project_steps(project_id)
    steps = execute_query(
        "SELECT status FROM project_steps WHERE project_id = ? ORDER BY stage",
        (project_id,),
        fetch=True
    )
    if not steps:
        return

    statuses = [s['status'] for s in steps]
    if len(statuses) == TOTAL_STAGES and all(s == 'approved' for s in statuses):
        new_status = 'completed'
    elif any(s == 'rejected' for s in statuses):
        new_status = 'rejected'
    elif any(s in ('submitted', 'approved') for s in statuses):
        new_status = 'in_progress'
    else:
        new_status = 'pending'

    execute_query(
        "UPDATE projects SET status = ?, updated_at = ? WHERE id = ?",
        (new_status, datetime.now(), project_id),
        commit=True
    )

# ==================== 认证系统 ====================
def authenticate_user(login_id, password):
    """用户认证（支持用户名/手机号/邮箱登录）"""
    user = execute_query('''
        SELECT u.*, o.status as org_status
        FROM users u
        LEFT JOIN organizations o ON u.org_id = o.id
        WHERE (u.username = ? OR u.phone = ? OR u.email = ?) AND u.status = 'active'
    ''', (login_id, login_id, login_id), fetch=True)
    
    if user and user[0]['role'] in ('org_admin', 'org_user') and user[0].get('org_status') != 'active':
        return None

    if user and check_password(password, user[0]['password_hash']):
        return user[0]
    return None

def get_current_user_access_state(user_id):
    rows = execute_query('''
        SELECT u.id, u.username, u.role, u.status, u.org_id, o.status as org_status
        FROM users u
        LEFT JOIN organizations o ON u.org_id = o.id
        WHERE u.id = ?
    ''', (user_id,), fetch=True)
    return rows[0] if rows else None

def ensure_current_account_active():
    """已登录会话也要校验账号/机构状态，避免停用后仍可继续操作。"""
    user = st.session_state.get('user')
    if not user:
        return True

    state = get_current_user_access_state(user.get('id'))
    if not state or state.get('status') != 'active':
        logout_current_user("当前账号已被冻结或不存在，请联系管理员", level='error')

    if state.get('role') in ('org_admin', 'org_user') and state.get('org_status') != 'active':
        logout_current_user("所属机构已停用，当前账号无法继续登录", level='error')

    return True

def get_client_ip():
    """获取客户端IP地址"""
    try:
        return socket.gethostbyname(socket.gethostname())
    except:
        return "unknown"

def logout_current_user(message=None, level='info'):
    """统一退出登录，避免清空 session 后当前脚本继续访问 user。"""
    user = dict(st.session_state.get('user') or {})
    try:
        org_name = ""
        if user.get('org_id'):
            org = execute_query("SELECT name FROM organizations WHERE id = ?", (user['org_id'],), fetch=True)
            if org:
                org_name = org[0]['name']
        if user:
            add_log(user.get('id'), user.get('username'), org_name, '退出登录', 'auth', '用户退出登录', get_client_ip())
    except Exception:
        pass

    st.session_state.clear()
    if message:
        set_flash_message(message, level=level)
    st.rerun()
    st.stop()

def hash_password(password):
    """密码加密"""
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())

def check_password(password, password_hash):
    """验证密码"""
    return bcrypt.checkpw(password.encode('utf-8'), normalize_password_hash(password_hash))

# ==================== 文件操作 ====================
def save_uploaded_file(uploaded_file, subfolder="", project_id=None):
    """保存上传的文件"""
    if uploaded_file is None:
        return None, None, 0
    
    # 创建子文件夹
    save_dir = os.path.join(UPLOAD_DIR, subfolder) if subfolder else UPLOAD_DIR
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    
    # 生成唯一文件名，带项目ID前缀（如提供）
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_ext = os.path.splitext(uploaded_file.name)[1]
    base_name = uploaded_file.name
    if project_id:
        file_name = f"{project_id}_{timestamp}_{base_name}"
    else:
        file_name = f"{timestamp}_{base_name}"
    file_path = os.path.join(save_dir, file_name)
    
    # 保存文件
    with open(file_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    
    return file_path, file_name, uploaded_file.size

def get_file_content(file_path, file_name):
    """获取文件内容用于预览"""
    if not os.path.exists(file_path):
        return None, None
    
    file_ext = os.path.splitext(file_name)[1].lower()
    
    with open(file_path, 'rb') as f:
        file_content = f.read()
    
    mime_type, _ = mimetypes.guess_type(file_name)
    if mime_type is None:
        mime_type = 'application/octet-stream'
    
    return file_content, mime_type


def safe_fname(name: str) -> str:
    """生成文件/文件夹安全名称，保留中文、字母和数字，其他替换为下划线"""
    if not name:
        return "unnamed"
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fa5_-]", "_", name)


def ensure_project_export_dirs_and_copy(project_id):
    """为项目创建 1_<项目名>,2_<项目名>,3_<项目名> 目录并复制已通过的文件到对应目录。
    返回字典：{'base': base_dir, 'copied': n, 'failed': m, 'paths': [...]}
    """
    try:
        proj_rows = execute_query("SELECT * FROM projects WHERE id = ?", (project_id,), fetch=True)
        if not proj_rows:
            return None
        proj = proj_rows[0]
        safe_name = safe_fname(proj.get('name') or f"proj_{project_id}")

        base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'exports')
        if not os.path.exists(base_dir):
            os.makedirs(base_dir, exist_ok=True)

        dest_dirs = {}
        for i in (1,2,3):
            d = os.path.join(base_dir, f"{i}_{safe_name}")
            os.makedirs(d, exist_ok=True)
            dest_dirs[i] = d

        other_dir = os.path.join(base_dir, f"other_{safe_name}")
        os.makedirs(other_dir, exist_ok=True)

        # 查询已通过的文件并按所属阶段复制
        rows = execute_query(
            "SELECT pf.*, ps.stage as stage FROM project_files pf LEFT JOIN project_steps ps ON pf.step_id = ps.id WHERE pf.project_id = ? AND pf.approval_status = 'approved'",
            (project_id,), fetch=True
        )

        copied = 0
        failed = 0
        paths = []
        for r in rows or []:
            src = r.get('file_path') or os.path.join(UPLOAD_DIR, r.get('file_name') or '')
            if not src or not os.path.exists(src):
                failed += 1
                continue
            stage = r.get('stage')
            if stage in (1,2,3):
                dest = os.path.join(dest_dirs[stage], r.get('file_name'))
            else:
                dest = os.path.join(other_dir, r.get('file_name'))
            try:
                shutil.copy2(src, dest)
                copied += 1
                paths.append(dest)
            except Exception:
                failed += 1

        return {'base': base_dir, 'copied': copied, 'failed': failed, 'paths': paths, 'proj_name': proj.get('name')}
    except Exception:
        return None

def display_file_preview(file_path, file_name):
    """显示文件预览"""
    if not os.path.exists(file_path):
        st.warning("文件不存在")
        return

    # safe_fname & ensure_project_export_dirs_and_copy 已移动到模块作用域

    file_ext = os.path.splitext(file_name)[1].lower()

    try:
        file_size = os.path.getsize(file_path)
    except Exception:
        file_size = 0

    # 严格限制内联预览阈值以避免冻结：PDF <=1MB, 文本 <=512KB, Excel <=1MB
    if file_ext == '.pdf':
        if file_size > 1 * 1024 * 1024:
            st.warning("PDF 文件较大，已切换为下载以避免卡顿。")
            with open(file_path, 'rb') as f:
                st.download_button(label="📥 下载 PDF", data=f, file_name=file_name, mime='application/pdf')
            return
        # 文件较小，尝试内联预览
        try:
            with open(file_path, 'rb') as f:
                base64_pdf = base64.b64encode(f.read()).decode('utf-8')
            pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
            st.markdown(pdf_display, unsafe_allow_html=True)
        except Exception:
            st.warning("无法在浏览器中预览该PDF，提供下载。")
            with open(file_path, 'rb') as f:
                st.download_button(label="📥 下载 PDF", data=f, file_name=file_name, mime='application/pdf')
        return
    
    elif file_ext in ['.txt', '.md']:
        # 对于较大的文本文件，避免一次性读取到内存中导致卡顿
        if file_size > 512 * 1024:
            st.warning("文本文件较大，已切换为下载以避免卡顿。")
            with open(file_path, 'rb') as f:
                st.download_button(label="📥 下载文件", data=f, file_name=file_name)
        else:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            st.text_area("文件内容", content, height=400)
        return
    
    elif file_ext in ['.xlsx', '.xls']:
        # 对于较大的 Excel 文件，避免读取到内存导致卡顿；提供下载
        if file_size > 1 * 1024 * 1024:
            st.warning("Excel 文件较大，已切换为下载以避免卡顿。")
            with open(file_path, 'rb') as f:
                st.download_button(label="📥 下载 Excel", data=f, file_name=file_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            return
        try:
            df = pd.read_excel(file_path)
            st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(f"无法读取Excel文件: {e}")
        return
    
    elif file_ext in ['.docx', '.doc']:
        st.info("Word文件预览功能需要安装python-docx库，请下载后查看")
        with open(file_path, 'rb') as f:
            st.download_button(
                label="📥 下载文件",
                data=f,
                file_name=file_name,
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
    else:
        with open(file_path, 'rb') as f:
            st.download_button(
                label="📥 下载文件",
                data=f,
                file_name=file_name
            )

# ==================== 登录页面 ====================
def render_login_page():
    """渲染登录页面"""
    apply_custom_styles()
    
    # 登录容器
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    
    # 标题
    st.markdown('''
        <div class="login-title">策链系统(ULD-CeLink)</div>
        <div class="login-subtitle">面向咨询生产与项目协调的系统</div>
    ''', unsafe_allow_html=True)
    
    # 使用中间列承载登录控件，确保真正居中且宽度更克制
    _, login_col, _ = st.columns([1, 1.8, 1])

    with login_col:
        # 客户端类型选择
        client_type = st.selectbox(
            "🖥️ 客户端类型",
            options=["auto", "admin", "org"],
            format_func=lambda x: {
                "auto": "🔄 自动识别",
                "admin": "🏢 管理端",
                "org": "🏛️ 机构端"
            }[x],
            key="client_type_select"
        )
        
        st.markdown("---")
        
        # 登录表单
        with st.form("login_form"):
            login_id = st.text_input("👤 用户名/手机号/邮箱", placeholder="请输入用户名、手机号或邮箱")
            password = st.text_input("🔑 密码", type="password", placeholder="请输入密码")
            
            submit = st.form_submit_button("🔐 登录", use_container_width=True)
        
            if submit:
                if not login_id or not password:
                    st.error("请填写完整的登录信息")
                else:
                    user = authenticate_user(login_id, password)
                    
                    if user:
                        # 根据客户端类型验证
                        if not can_access_client(user['role'], client_type):
                            target_name = "管理端" if client_type == "admin" else "机构端"
                            st.error(f"该账号无权访问{target_name}")
                        else:
                            resolved_client_type = resolve_client_type(user['role'], client_type)
                            # 登录成功
                            st.session_state['logged_in'] = True
                            st.session_state['user'] = dict(user)
                            st.session_state['client_type'] = resolved_client_type
                            st.session_state['current_page'] = get_default_page_for_role(user['role']) or 'dashboard'
                            
                            # 获取机构名称
                            org_name = ""
                            if user['org_id']:
                                org = execute_query("SELECT name FROM organizations WHERE id = ?", (user['org_id'],), fetch=True)
                                if org:
                                    org_name = org[0]['name']
                            
                            # 记录登录日志
                            add_log(user['id'], user['username'], org_name, '登录', 'auth', '用户登录成功', get_client_ip())

                            st.rerun()
                    else:
                        st.error("用户名或密码错误，或账号已被冻结")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # 底部信息
    st.markdown('''
        <div style="text-align: center; margin-top: 30px; color: rgba(230, 250, 255, 0.82);">
            <p>© 2026 策链系统(ULD-CeLink) | 技术支持</p>
        </div>
    ''', unsafe_allow_html=True)

# ==================== 侧边栏 ====================
def navigate_to(page):
    """切换页面，并清理跨页面的临时展示状态。"""
    user = st.session_state.get('user') or {}
    role = user.get('role')
    if not can_access_page(role, page):
        st.session_state['current_page'] = get_default_page_for_role(role)
        set_flash_message("当前账号没有权限访问该页面", level='error')
        return
    st.session_state['current_page'] = page
    st.session_state['show_change_pwd'] = False
    st.session_state.pop('view_file_id', None)


def handle_query_navigation():
    """处理工作台卡片链接跳转，统一复用侧边栏权限与清理逻辑。"""
    try:
        target_page = st.query_params.get("goto")
    except Exception:
        target_page = None

    if isinstance(target_page, list):
        target_page = target_page[0] if target_page else None

    if not target_page:
        return

    navigate_to(str(target_page))
    try:
        del st.query_params["goto"]
    except Exception:
        try:
            st.query_params.clear()
        except Exception:
            pass


def render_sidebar():
    """渲染侧边栏"""
    user = st.session_state.get('user')
    if not user:
        logout_current_user()
        return
    role = user['role']
    
    with st.sidebar:
        # 机构Logo
        logo_path = r"C:\Users\Administrator\Desktop\总结\优兰德模板汇总\logo.jpg"
        encoded_string = get_base64_file(logo_path)
        if encoded_string:
            st.markdown(f"""
            <div style="text-align: center; padding: 10px; background: rgba(255,255,255,0.9); border-radius: 12px; margin-bottom: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.2);">
                <img src="data:image/jpeg;base64,{encoded_string}" style="width: 100%; border-radius: 8px;">
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="text-align: center; padding: 20px; background: rgba(255,255,255,0.1); border-radius: 10px; margin-bottom: 20px;">
                <div style="font-size: 24px; font-weight: bold; color: #fff;">第三方评估系统</div>
            </div>
            """, unsafe_allow_html=True)
        
        # 导航菜单
        menu_items = get_menu_items_for_role(role)

        # 实时计数：合并查询，减少侧边栏重跑时的数据库访问次数
        try:
            badge_snapshot = get_sidebar_badges(user['id'])
            pending_total = (badge_snapshot.get('pending_approvals') or 0) + (badge_snapshot.get('pending_files') or 0)
            pending_todos = badge_snapshot.get('pending_todos') or 0
            msg_unread = badge_snapshot.get('unread_messages') or 0
        except Exception:
            pending_total = 0
            pending_todos = 0
            msg_unread = 0

        # 渲染菜单，使用两列：按钮 + 气泡计数
        for icon, name, page in menu_items:
            col_a, col_b = st.columns([8, 1])
            with col_a:
                st.button(
                    f"{icon} {name}",
                    key=f"nav_{page}",
                    use_container_width=True,
                    on_click=navigate_to,
                    args=(page,),
                )
            with col_b:
                badge_count = 0
                if page == 'approval':
                    badge_count = pending_total
                elif page == 'todos':
                    badge_count = pending_todos
                elif page == 'messages':
                    badge_count = msg_unread

                if badge_count and badge_count > 0:
                    # 缩小间距并拉近到按钮侧边
                    st.markdown(f"<div style=\"background:#ff4d4f;color:#fff;border-radius:999px;padding:4px 8px;text-align:center;font-weight:700;margin-left:-18px;line-height:20px;\">{badge_count}</div>", unsafe_allow_html=True)
                else:
                    st.markdown("", unsafe_allow_html=True)
        
        st.markdown("---")
        
        # 修改密码 & 退出登录（调整为与菜单按钮一致的宽度）
        col_a, col_b = st.columns([8,1])
        with col_a:
            if st.button("🔑 修改密码", key="change_pwd_btn"):
                st.session_state['show_change_pwd'] = True
        with col_b:
            st.markdown("", unsafe_allow_html=True)

        col_c, col_d = st.columns([8,1])
        with col_c:
            if st.button("🚪 退出登录", key="logout_btn"):
                logout_current_user()
        with col_d:
            st.markdown("", unsafe_allow_html=True)
        
        # 修改密码弹窗
        if st.session_state.get('show_change_pwd'):
            render_change_password_modal()

def refresh_session_user(user_id):
    """从数据库刷新当前会话用户，避免权限/改密标记使用旧值。"""
    rows = execute_query("SELECT * FROM users WHERE id = ?", (user_id,), fetch=True)
    if rows:
        st.session_state['user'] = dict(rows[0])
        return st.session_state['user']
    return None

def is_password_change_required():
    user = st.session_state.get('user')
    if not user:
        return False
    current_user = refresh_session_user(user['id'])
    if not current_user:
        logout_current_user()
    return int(current_user.get('must_change_password') or 0) == 1

def render_required_password_change():
    """首次登录或重置密码后，强制用户先修改密码。"""
    _, content_col, _ = st.columns([0.08, 0.84, 0.08])
    with content_col:
        st.markdown(
            f"""
            <div class="password-change-header">
                <div class="password-change-title">🔑 修改初始密码</div>
                <p class="password-change-subtitle">当前账号使用初始或重置密码，完成修改后才能进入系统。</p>
            </div>
            <div class="password-change-notice">
                {PASSWORD_RULE_TEXT}，且不能与系统默认密码相同。
            </div>
            """,
            unsafe_allow_html=True
        )
        if st.session_state.get('_password_change_completed'):
            st.success("密码修改成功，请点击下方按钮进入系统。")
            if st.button("进入系统", use_container_width=True, key="enter_after_password_change"):
                st.session_state.pop('_password_change_completed', None)
                st.rerun()
            return
        render_change_password_modal(force=True)

def render_change_password_modal(force=False):
    """修改密码弹窗"""
    st.markdown('<div class="password-change-card-title">修改密码</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="password-change-rule">{PASSWORD_RULE_TEXT}，且不能与系统默认密码相同。</div>',
        unsafe_allow_html=True
    )
    
    user = st.session_state['user']
    feedback_placeholder = st.empty()

    old_pwd = st.text_input("原密码", type="password", key=f"change_pwd_old_{'force' if force else 'normal'}")
    new_pwd = st.text_input("新密码", type="password", key=f"change_pwd_new_{'force' if force else 'normal'}")
    confirm_pwd = st.text_input("确认新密码", type="password", key=f"change_pwd_confirm_{'force' if force else 'normal'}")
    
    col1, col2 = st.columns(2)
    with col1:
        submit = st.button("确认修改", use_container_width=True, key=f"change_pwd_submit_{'force' if force else 'normal'}")
    with col2:
        cancel_label = "取消并退出登录" if force else "取消"
        cancel = st.button(cancel_label, use_container_width=True, key=f"change_pwd_cancel_{'force' if force else 'normal'}")
    
    if cancel:
        if force:
            logout_current_user()
            return
        else:
            st.session_state['show_change_pwd'] = False
        st.rerun()
    
    if submit:
        try:
            if not old_pwd or not new_pwd or not confirm_pwd:
                feedback_placeholder.error("请填写所有字段")
            else:
                # 验证原密码
                current_user = execute_query("SELECT * FROM users WHERE id = ?", (user['id'],), fetch=True)
                if not current_user or not check_password(old_pwd, current_user[0]['password_hash']):
                    feedback_placeholder.error("原密码错误")
                    return

                if new_pwd != confirm_pwd:
                    feedback_placeholder.error("两次输入的新密码不一致")
                    return

                policy_error = validate_password_policy(new_pwd, allow_default=False)
                if policy_error:
                    feedback_placeholder.error(policy_error)
                    return

                if check_password(new_pwd, current_user[0]['password_hash']):
                    feedback_placeholder.error("新密码不能与当前密码相同")
                    return

                # 更新密码
                new_hash = hash_password(new_pwd)
                execute_query(
                    "UPDATE users SET password_hash = ?, must_change_password = 0, updated_at = ? WHERE id = ?",
                    (new_hash, datetime.now(), user['id']),
                    commit=True
                )
                refresh_session_user(user['id'])
                
                org_name = ""
                if user.get('org_id'):
                    org = execute_query("SELECT name FROM organizations WHERE id = ?", (user['org_id'],), fetch=True)
                    if org:
                        org_name = org[0]['name']
                
                add_log(user['id'], user['username'], org_name, '修改密码', 'auth', '用户修改密码成功', get_client_ip())
                add_message(user['id'], '密码修改成功', f'您于{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}成功修改密码')
                
                st.session_state['show_change_pwd'] = False
                if force:
                    st.session_state['_password_change_completed'] = True
                    feedback_placeholder.success("密码修改成功，请点击“进入系统”继续。")
                else:
                    notify_and_rerun("密码修改成功", level='success')
        except Exception as e:
            feedback_placeholder.error(f"密码修改失败: {e}")

    if force and st.session_state.get('_password_change_completed'):
        if st.button("进入系统", use_container_width=True, key="enter_after_password_change_inline"):
            st.session_state.pop('_password_change_completed', None)
            st.rerun()

# ==================== 管理端页面 ====================
def render_admin_dashboard():
    """管理端数据大盘"""
    dashboard_snapshot = get_admin_dashboard_snapshot()
    org_count = dashboard_snapshot['org_count']
    user_count = dashboard_snapshot['user_count']
    active_projects = dashboard_snapshot['active_projects']
    completed_projects = dashboard_snapshot['completed_projects']
    pending_approvals = dashboard_snapshot['pending_approvals']
    pending_files = dashboard_snapshot['pending_files']
    pending_project_rows = execute_query(
        "SELECT COUNT(*) as cnt FROM projects WHERE status = 'pending'",
        fetch=True
    )
    pending_projects = pending_project_rows[0]['cnt'] if pending_project_rows else 0
    pending_todo_total = pending_approvals + pending_files

    st.markdown(textwrap.dedent(f"""
    <div class="org-workbench">
        <section class="org-hero">
            <div class="org-hero-grid">
                <div>
                    <div class="org-kicker">超级管理员数据大盘</div>
                    <h1>策链系统(ULD-CeLink)</h1>
                    <div class="org-hero-tags">
                        <span class="org-hero-tag">机构统筹</span>
                        <span class="org-hero-tag">项目审核</span>
                        <span class="org-hero-tag">项目管理</span>
                        <span class="org-hero-tag">交付监管</span>
                    </div>
                </div>
                <div class="org-hero-panel">
                    <div class="org-hero-panel-title">运行概览</div>
                    <div class="org-hero-panel-value">管理驾驶舱</div>
                    <div class="org-list-meta" style="color: rgba(234,251,255,0.72); margin-top: 10px;">
                        待处理 {pending_projects + pending_todo_total} 项 · 进行中项目 {active_projects} 个 · 机构 {org_count} 家
                    </div>
                </div>
            </div>
        </section>
    </div>
    """).strip(), unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
            <div class="org-click-card" data-admin-nav="机构管理" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">🏢</div></div>
                    <div class="org-metric-number">{org_count}</div>
                    <div class="org-metric-label">机构总数</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        st.button(
            "查看机构",
            key="goto_orgs",
            use_container_width=True,
            on_click=navigate_to,
            args=('organizations',),
        )
    
    with col2:
        st.markdown(f"""
            <div class="org-click-card" data-admin-nav="账号管理" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">👥</div></div>
                    <div class="org-metric-number">{user_count}</div>
                    <div class="org-metric-label">用户总数</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        st.button(
            "查看用户",
            key="goto_users",
            use_container_width=True,
            on_click=navigate_to,
            args=('users',),
        )
    
    with col3:
        st.markdown(f"""
            <div class="org-click-card" data-admin-nav="项目审核" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">🚀</div></div>
                    <div class="org-metric-number">{active_projects}</div>
                    <div class="org-metric-label">进行中项目</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        st.button(
            "查看项目",
            key="goto_projects",
            use_container_width=True,
            on_click=navigate_to,
            args=('projects',),
        )
    
    with col4:
        st.markdown(f"""
            <div class="org-click-card" data-admin-nav="项目审核" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">✅</div></div>
                    <div class="org-metric-number">{completed_projects}</div>
                    <div class="org-metric-label">已完成项目</div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        st.button(
            "查看项目",
            key="goto_projects_completed",
            use_container_width=True,
            on_click=navigate_to,
            args=('projects',),
        )

    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"""
        <div class="admin-todo-panel">
            <div class="org-panel-header">
                <div class="org-panel-title">待处理事项</div>
                <div class="org-panel-badge">{pending_projects + pending_todo_total} 项待处理</div>
            </div>
            <div class="admin-todo-items">
                <div class="admin-todo-card" data-admin-nav="项目审核" role="button" tabindex="0">
                    <div class="admin-todo-icon">G</div>
                    <div>
                        <div class="admin-todo-title">待审核项目</div>
                        <div class="admin-todo-meta">项目阶段流转审核事项</div>
                    </div>
                    <div class="admin-todo-count">{pending_projects}</div>
                </div>
                <div class="admin-todo-card" data-admin-nav="审批待办" role="button" tabindex="0">
                    <div class="admin-todo-icon">F</div>
                    <div>
                        <div class="admin-todo-title">待审批待办</div>
                        <div class="admin-todo-meta">Gate 审批待办 + 文件审批待办</div>
                    </div>
                    <div class="admin-todo-count">{pending_todo_total}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if pending_approvals + pending_files > 0:
            st.button(
                "去处理审批",
                key="goto_approval",
                use_container_width=True,
                on_click=navigate_to,
                args=('approval',),
            )
        else:
            st.success("暂无待处理事项")
        components.html("""
        <script>
        (function() {
            const doc = window.parent.document;
            if (!doc) {
                return;
            }

            function findSidebarButton(label) {
                const sidebar = doc.querySelector('[data-testid="stSidebar"]') || doc;
                const buttons = Array.from(sidebar.querySelectorAll('button'));
                return buttons.find(function(button) {
                    return (button.innerText || button.textContent || '').trim().includes(label);
                });
            }

            function triggerAdminTodoNav(event) {
                if (event.type === 'keydown' && event.key !== 'Enter' && event.key !== ' ') {
                    return;
                }
                const target = event.target && event.target.closest ? event.target.closest('[data-admin-nav]') : null;
                if (!target) {
                    return;
                }
                const button = findSidebarButton(target.getAttribute('data-admin-nav'));
                if (!button) {
                    return;
                }
                event.preventDefault();
                event.stopPropagation();
                button.click();
            }

            if (doc.__adminTodoNavHandler) {
                doc.removeEventListener('click', doc.__adminTodoNavHandler, true);
                doc.removeEventListener('keydown', doc.__adminTodoNavHandler, true);
            }
            doc.__adminTodoNavHandler = triggerAdminTodoNav;
            doc.addEventListener('click', triggerAdminTodoNav, true);
            doc.addEventListener('keydown', triggerAdminTodoNav, true);
        })();
        </script>
        """, height=0)
    
    with col2:
        st.markdown("""
        <div class="org-panel admin-click-panel" data-admin-nav="项目审核" role="button" tabindex="0">
            <div class="org-panel-header">
                <div class="org-panel-title">项目状态分布</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        status_data = execute_query('''
            SELECT status, COUNT(*) as cnt FROM projects GROUP BY status
        ''', fetch=True)
        
        if status_data:
            df = pd.DataFrame(status_data)
            status_map = {'pending': '待审核', 'in_progress': '进行中', 'completed': '已完成', 'rejected': '已驳回'}
            df['status_name'] = df['status'].map(status_map)
            
            fig = px.pie(df, values='cnt', names='status_name',
                        color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(height=300, margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("暂无项目数据")
    
    st.markdown("""
    <div class="org-panel org-project-panel admin-click-panel" data-admin-nav="日志查看" role="button" tabindex="0">
        <div class="org-panel-header">
            <div class="org-panel-title">最近活动</div>
            <div class="org-panel-badge" data-admin-nav="日志查看" role="button" tabindex="0">最近 10 条</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    recent_logs = execute_query('''
        SELECT * FROM logs ORDER BY created_at DESC LIMIT 10
    ''', fetch=True)
    
    if recent_logs:
        df = pd.DataFrame(recent_logs)
        df = df[['id', 'username', 'org_name', 'action', 'module', 'ip_address', 'created_at']]
        df.columns = ['日志ID', '用户名', '机构名', '操作', '模块', 'IP地址', '操作时间']
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("暂无活动记录")

def render_admin_organizations():
    """管理端机构管理"""
    st.title("🏢 机构管理")
    
    tab1, tab2 = st.tabs(["机构列表", "新增机构"])
    
    with tab1:
        # 搜索和筛选
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            search = st.text_input("搜索机构", placeholder="输入机构ID（纯数字）或机构名称/信用代码")
        with col2:
            status_filter = st.selectbox("状态筛选", ["全部", "active", "inactive"], format_func=lambda x: {"全部": "全部", "active": "启用", "inactive": "停用"}.get(x, x))
        with col3:
            st.write("")  # 占位
        
        # 查询机构列表
        query = '''
            SELECT o.*, 
                   (SELECT COUNT(*) FROM users WHERE org_id = o.id) as user_count,
                   (SELECT COUNT(*) FROM projects WHERE org_id = o.id) as project_count
            FROM organizations o WHERE 1=1
        '''
        params = []
        
        if search:
            s = str(search).strip()
            if s.isdigit():
                # 纯数字：按机构ID精确匹配，避免“1”命中大量信用代码
                query += " AND o.id = ?"
                params.append(int(s))
            else:
                query += " AND (o.name LIKE ? OR o.credit_code LIKE ?)"
                params.extend([f"%{s}%", f"%{s}%"])
        
        if status_filter != "全部":
            query += " AND o.status = ?"
            params.append(status_filter)
        
        query += " ORDER BY o.created_at DESC"
        
        orgs = execute_query(query, params, fetch=True)
        
        if orgs:
            for org in orgs:
                with st.expander(f"**{org['name']}** ({'启用' if org['status'] == 'active' else '停用'})"):
                    col1, col2, col3 = st.columns([2, 2, 1])
                    
                    with col1:
                        st.write(f"**机构类型:** {org['org_type'] or '-'}")
                        st.write(f"**统一社会信用代码:** {org['credit_code'] or '-'}")
                        st.write(f"**法定代表人:** {org['legal_person'] or '-'}")
                        st.write(f"**联系人:** {org['contact_person'] or '-'}")
                    
                    with col2:
                        st.write(f"**联系电话:** {org['contact_phone'] or '-'}")
                        st.write(f"**联系邮箱:** {org['contact_email'] or '-'}")
                        st.write(f"**机构地址:** {org['address'] or '-'}")
                        st.write(f"**创建时间:** {format_datetime_display(org['created_at'])}")
                    
                    with col3:
                        st.write(f"**用户数:** {org['user_count']}")
                        st.write(f"**项目数:** {org['project_count']}")
                    
                    if org['description']:
                        st.write(f"**机构简介:** {org['description']}")
                    
                    # 操作按钮
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        if org['status'] == 'active':
                            if st.button("⏸️ 停用", key=f"deactivate_{org['id']}"):
                                execute_query("UPDATE organizations SET status = 'inactive', updated_at = ? WHERE id = ?", 
                                            (datetime.now(), org['id']), commit=True)
                                notify_and_rerun("机构已停用", level='success')
                        else:
                            if st.button("▶️ 启用", key=f"activate_{org['id']}"):
                                execute_query("UPDATE organizations SET status = 'active', updated_at = ? WHERE id = ?", 
                                            (datetime.now(), org['id']), commit=True)
                                notify_and_rerun("机构已启用", level='success')
                    with col2:
                        if st.button("✏️ 编辑", key=f"edit_{org['id']}"):
                            st.session_state['edit_org_id'] = org['id']
                            st.rerun()
                    
                    with col3:
                        if st.button("🔑 重置密码", key=f"reset_pwd_{org['id']}"):
                            # 获取机构主账号
                            main_user = execute_query("SELECT * FROM users WHERE org_id = ? AND role = 'org_admin'", 
                                                     (org['id'],), fetch=True)
                            if main_user:
                                default_password = get_default_password_for_role(main_user[0]['role'])
                                new_hash = hash_password(default_password)
                                execute_query(
                                    "UPDATE users SET password_hash = ?, must_change_password = 1, updated_at = ? WHERE id = ?",
                                    (new_hash, datetime.now(), main_user[0]['id']),
                                    commit=True
                                )
                                notify_and_rerun(f"密码已重置为: {default_password}，下次登录必须修改密码", level='success')
                            else:
                                st.warning("未找到机构主账号")
                    
                    with col4:
                        if st.button("🗑️ 删除", key=f"delete_{org['id']}"):
                            # 检查是否有关联数据
                            if org['user_count'] > 0 or org['project_count'] > 0:
                                st.error("该机构下有用户或项目，无法删除")
                            else:
                                execute_query("DELETE FROM organizations WHERE id = ?", (org['id'],), commit=True)
                                execute_query("DELETE FROM users WHERE org_id = ?", (org['id'],), commit=True)
                                notify_and_rerun("机构已删除", level='success')
        else:
            st.info("暂无机构数据")
    
    # 如果从列表点击编辑机构，展示编辑表单（优先于新增机构）
    if st.session_state.get('edit_org_id'):
        edit_id = st.session_state.get('edit_org_id')
        org_row = execute_query("SELECT * FROM organizations WHERE id = ?", (edit_id,), fetch=True)
        if org_row:
            org_obj = org_row[0]
            st.subheader(f"编辑机构: {org_obj.get('name')}")
            with st.form("edit_org_form"):
                col1, col2 = st.columns(2)
                with col1:
                    name = st.text_input("机构名称 *", value=org_obj.get('name') or "")
                    org_type = st.selectbox("机构类型 *", ["企业", "事业单位", "社会团体", "民办非企业", "其他"], index=0)
                    credit_code = st.text_input("统一社会信用代码 *", value=org_obj.get('credit_code') or "")
                    legal_person = st.text_input("法定代表人 *", value=org_obj.get('legal_person') or "")
                    contact_person = st.text_input("联系人 *", value=org_obj.get('contact_person') or "")
                with col2:
                    contact_phone = st.text_input("联系电话 *", value=org_obj.get('contact_phone') or "")
                    contact_email = st.text_input("联系邮箱 *", value=org_obj.get('contact_email') or "")
                    address = st.text_input("机构地址 *", value=org_obj.get('address') or "")
                    description = st.text_area("机构简介", value=org_obj.get('description') or "")
                submit = st.form_submit_button("✅ 保存修改", use_container_width=True)
                cancel = st.form_submit_button("取消", use_container_width=True)
                if cancel:
                    del st.session_state['edit_org_id']
                    st.experimental_rerun()
                if submit:
                    # 基本验证
                    required_fields = [name, org_type, credit_code, legal_person, contact_person, contact_phone, contact_email, address]
                    if not all(required_fields):
                        st.error("请填写所有必填项")
                    else:
                        # 检查信用代码冲突（排除当前机构）
                        existing = execute_query("SELECT id FROM organizations WHERE credit_code = ? AND id != ?", (credit_code, edit_id), fetch=True)
                        if existing:
                            st.error("该统一社会信用代码已被其他机构使用")
                        else:
                            execute_query('''
                                UPDATE organizations SET name = ?, org_type = ?, credit_code = ?, legal_person = ?, contact_person = ?,
                                                   contact_phone = ?, contact_email = ?, address = ?, description = ?, updated_at = ?
                                WHERE id = ?
                            ''', (name, org_type, credit_code, legal_person, contact_person, contact_phone, contact_email, address, description, datetime.now(), edit_id), commit=True)
                            add_log(st.session_state['user']['id'], st.session_state['user']['username'], '', '编辑机构', 'organizations', f'编辑机构: {name}', get_client_ip())
                            del st.session_state['edit_org_id']
                            notify_and_rerun("机构信息已更新", level='success')
    
    with tab2:
        st.subheader("新增机构")
        
        with st.form("add_org_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                name = st.text_input("机构名称 *", placeholder="请输入机构名称")
                org_type = st.selectbox("机构类型 *", ["企业", "事业单位", "社会团体", "民办非企业", "其他"])
                credit_code = st.text_input("统一社会信用代码 *", placeholder="请输入18位信用代码")
                legal_person = st.text_input("法定代表人 *", placeholder="请输入法定代表人姓名")
                contact_person = st.text_input("联系人 *", placeholder="请输入联系人姓名")
            
            with col2:
                contact_phone = st.text_input("联系电话 *", placeholder="请输入联系电话")
                contact_email = st.text_input("联系邮箱 *", placeholder="请输入联系邮箱")
                address = st.text_input("机构地址 *", placeholder="请输入机构地址")
                description = st.text_area("机构简介", placeholder="请输入机构简介（选填）")
            
            submit = st.form_submit_button("✅ 创建机构", use_container_width=True)
            
            if submit:
                # 验证必填项
                required_fields = [name, org_type, credit_code, legal_person, contact_person, contact_phone, contact_email, address]
                if not all(required_fields):
                    st.error("请填写所有必填项")
                else:
                    # 检查信用代码是否重复
                    existing = execute_query("SELECT id FROM organizations WHERE credit_code = ?", (credit_code,), fetch=True)
                    if existing:
                        st.error("该统一社会信用代码已存在")
                    else:
                        # 创建机构
                        org_id = execute_query('''
                            INSERT INTO organizations (name, org_type, credit_code, legal_person, contact_person, 
                                                      contact_phone, contact_email, address, description)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (name, org_type, credit_code, legal_person, contact_person, 
                              contact_phone, contact_email, address, description), commit=True)
                        
                        # 创建机构主账号
                        username = f"org_{org_id}"
                        default_password = get_default_password_for_role('org_admin')
                        password_hash = hash_password(default_password)
                        
                        execute_query('''
                            INSERT INTO users (username, password_hash, role, org_id, real_name, phone, email, status, must_change_password)
                            VALUES (?, ?, 'org_admin', ?, ?, ?, ?, 'active', 1)
                        ''', (username, password_hash, org_id, contact_person, contact_phone, contact_email), commit=True)
                        
                        # 记录日志
                        user = st.session_state['user']
                        add_log(user['id'], user['username'], '', '新增机构', 'organizations', 
                               f'创建机构: {name}', get_client_ip())
                        
                        notify_and_rerun(f"机构创建成功！主账号: {username}，默认密码: {default_password}，首次登录必须修改密码", level='success')
def render_admin_users():
    """管理端账号管理"""
    st.title("👥 账号管理")
    
    tab1, tab2 = st.tabs(["账号列表", "新增账号"])
    
    with tab1:
        # 筛选
        col1, col2, col3 = st.columns(3)
        with col1:
            search = st.text_input("搜索账号", placeholder="用户名/手机号/邮箱/机构名")
        with col2:
            role_filter = st.selectbox("角色筛选", ["全部", "super_admin", "org_admin", "org_user"],
                                       format_func=lambda x: {"全部": "全部", "super_admin": "超级管理员", 
                                                             "org_admin": "机构主账号", "org_user": "机构子账号"}.get(x, x))
        with col3:
            status_filter = st.selectbox("状态筛选", ["全部", "active", "inactive"],
                                        format_func=lambda x: {"全部": "全部", "active": "正常", "inactive": "冻结"}.get(x, x))
        
        # 查询用户列表
        query = '''
            SELECT u.*, o.name as org_name,
                   admin_u.username as org_admin_username
            FROM users u 
            LEFT JOIN organizations o ON u.org_id = o.id 
            LEFT JOIN users admin_u
                ON admin_u.org_id = u.org_id
               AND admin_u.role = 'org_admin'
               AND admin_u.status = 'active'
            WHERE 1=1
        '''
        params = []
        
        if search:
            query += " AND (u.username LIKE ? OR u.phone LIKE ? OR u.email LIKE ? OR u.real_name LIKE ? OR o.name LIKE ?)"
            params.extend([f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%"])
        
        if role_filter != "全部":
            query += " AND u.role = ?"
            params.append(role_filter)
        
        if status_filter != "全部":
            query += " AND u.status = ?"
            params.append(status_filter)
        
        query += " ORDER BY u.created_at DESC"
        
        users = execute_query(query, params, fetch=True)
        
        if users:
            for user in users:
                role_names = {'super_admin': '超级管理员', 'org_admin': '机构主账号', 'org_user': '机构子账号'}
                status_names = {'active': '正常', 'inactive': '冻结'}
                display_username = user['username']
                if user['role'] == 'org_user' and user['org_admin_username']:
                    display_username = f"{user['org_admin_username']}-{user['username']}"
                
                with st.expander(f"**{display_username}** - {role_names.get(user['role'], user['role'])} ({status_names.get(user['status'], user['status'])})"):
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.write(f"**姓名:** {user['real_name'] or '-'}")
                        st.write(f"**手机号:** {user['phone'] or '-'}")
                        st.write(f"**邮箱:** {user['email'] or '-'}")
                    
                    with col2:
                        st.write(f"**所属机构:** {user['org_name'] or '-'}")
                        st.write(f"**创建时间:** {user['created_at']}")
                    
                    with col3:
                        st.write(f"**角色:** {role_names.get(user['role'], user['role'])}")
                        st.write(f"**状态:** {status_names.get(user['status'], user['status'])}")
                    
                    # 操作按钮
                    if user['role'] != 'super_admin':
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            if user['status'] == 'active':
                                if st.button("🔒 冻结", key=f"freeze_{user['id']}"):
                                    execute_query("UPDATE users SET status = 'inactive' WHERE id = ?", (user['id'],), commit=True)
                                    notify_and_rerun("账号已冻结", level='success')
                            else:
                                if st.button("🔓 解冻", key=f"unfreeze_{user['id']}"):
                                    execute_query("UPDATE users SET status = 'active' WHERE id = ?", (user['id'],), commit=True)
                                    notify_and_rerun("账号已解冻", level='success')
                        with col2:
                            if st.button("🔑 重置密码", key=f"resetpwd_{user['id']}"):
                                default_password = get_default_password_for_role(user['role'])
                                new_hash = hash_password(default_password)
                                execute_query(
                                    "UPDATE users SET password_hash = ?, must_change_password = 1, updated_at = ? WHERE id = ?",
                                    (new_hash, datetime.now(), user['id']),
                                    commit=True
                                )
                                notify_and_rerun(f"密码已重置为: {default_password}，下次登录必须修改密码", level='success')
                        
                        with col3:
                            if st.button("✏️ 编辑", key=f"edituser_{user['id']}"):
                                st.session_state['edit_user_id'] = user['id']
                                st.rerun()
                        
                        with col4:
                            if st.button("🗑️ 删除", key=f"deleteuser_{user['id']}"):
                                execute_query("UPDATE evaluators SET account_user_id = NULL WHERE account_user_id = ?",
                                              (user['id'],), commit=True)
                                execute_query("DELETE FROM users WHERE id = ?", (user['id'],), commit=True)
                                notify_and_rerun("账号已删除", level='success')
        else:
            st.info("暂无账号数据")
    
    with tab2:
        st.subheader("新增账号")
        
        with st.form("add_user_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                new_username = st.text_input("用户名 *", placeholder="请输入用户名")
                new_real_name = st.text_input("姓名 *", placeholder="请输入姓名")
                new_phone = st.text_input("手机号 *", placeholder="请输入手机号")
                new_email = st.text_input("邮箱 *", placeholder="请输入邮箱")
            
            with col2:
                new_role = st.selectbox("角色 *", ["org_admin", "org_user"],
                                       format_func=lambda x: {"org_admin": "机构主账号", "org_user": "机构子账号"}[x])
                
                # 获取机构列表
                orgs = execute_query("SELECT id, name FROM organizations WHERE status = 'active' ORDER BY name", fetch=True)
                org_options = {f"{o['name']}": o['id'] for o in orgs} if orgs else {}
                
                new_org = st.selectbox("所属机构 *", list(org_options.keys()) if org_options else ["暂无可用机构"])
                new_password = st.text_input("初始密码 *", type="password", placeholder="请输入初始密码")
            
            submit = st.form_submit_button("✅ 创建账号", use_container_width=True)
            
            if submit:
                required_fields = [new_username, new_real_name, new_phone, new_email, new_password]
                if not all(required_fields):
                    st.error("请填写所有必填项")
                elif not org_options:
                    st.error("请先创建机构")
                elif validate_password_policy(new_password, allow_default=True):
                    st.error(validate_password_policy(new_password, allow_default=True))
                else:
                    # 检查用户名是否重复
                    existing = execute_query("SELECT id FROM users WHERE username = ? OR phone = ? OR email = ?", 
                                           (new_username, new_phone, new_email), fetch=True)
                    if existing:
                        st.error("用户名、手机号或邮箱已存在")
                    else:
                        password_hash = hash_password(new_password)
                        org_id = org_options.get(new_org)
                        
                        execute_query('''
                            INSERT INTO users (username, password_hash, role, org_id, real_name, phone, email, status, must_change_password)
                            VALUES (?, ?, ?, ?, ?, ?, ?, 'active', 1)
                        ''', (new_username, password_hash, new_role, org_id, new_real_name, new_phone, new_email), commit=True)
                        
                        # 记录日志
                        user = st.session_state['user']
                        add_log(user['id'], user['username'], '', '新增账号', 'users', 
                               f'创建账号: {new_username}', get_client_ip())
                        
                        notify_and_rerun("账号创建成功，首次登录必须修改密码", level='success')
def render_admin_projects():
    """管理端项目审核"""
    st.title("📋 项目审核")
    
    # 筛选
    col1, col2, col3 = st.columns(3)
    with col1:
        org_filter = st.text_input("机构名称", placeholder="输入机构名称搜索")
    with col2:
        status_filter = st.selectbox("状态筛选", ["全部", "pending", "in_progress", "completed", "rejected"],
                                    format_func=lambda x: {"全部": "全部", "pending": "待审核", "in_progress": "进行中", 
                                                          "completed": "已完成", "rejected": "已驳回"}.get(x, x))
    with col3:
        stage_filter = st.selectbox(
            "Gate筛选",
            ["全部"] + list(STAGE_NAMES.keys()),
            format_func=lambda x: "全部" if x == "全部" else STAGE_NAMES.get(x, str(x))
        )
    
    # 查询项目
    query = '''
        SELECT p.*, o.name as org_name, u.username as creator_name
        FROM projects p
        JOIN organizations o ON p.org_id = o.id
        LEFT JOIN users u ON p.created_by = u.id
        WHERE 1=1
    '''
    params = []
    
    if org_filter:
        query += " AND o.name LIKE ?"
        params.append(f"%{org_filter}%")
    
    if status_filter != "全部":
        query += " AND p.status = ?"
        params.append(status_filter)
    
    if stage_filter != "全部":
        # Gate筛选逻辑：仅显示该Gate尚未审核通过的项目（已通过则不显示）
        query += '''
            AND EXISTS (
                SELECT 1 FROM project_steps psf
                WHERE psf.project_id = p.id
                  AND psf.stage = ?
                  AND psf.status != 'approved'
            )
        '''
        params.append(stage_filter)
    
    query += " ORDER BY p.created_at DESC"
    
    projects = execute_query(query, params, fetch=True)
    
    if projects:
        for proj in projects:
            status_map = {'pending': '待审核', 'in_progress': '进行中', 'completed': '已完成', 'rejected': '已驳回'}
            
            with st.expander(f"**{proj['name']}** - {proj['org_name']} ({status_map.get(proj['status'], proj['status'])})"):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write(f"**项目ID:** {proj['id']}")
                    st.write(f"**所属机构:** {proj['org_name']}")
                    st.write(f"**项目分类:** {PROJECT_CATEGORIES.get(proj['category'], {}).get('name', '-') if proj['category'] else '-'}")
                    st.write(f"**创建人:** {proj['creator_name'] or '-'}")
                
                with col2:
                    st.write(f"**状态:** {status_map.get(proj['status'], proj['status'])}")
                    current_stage = proj['current_stage']
                    st.write(f"**当前Gate:** {format_gate(current_stage)}")
                    st.write(f"**创建时间:** {format_datetime_display(proj['created_at'])}")
                
                if proj['description']:
                    st.write(f"**项目描述:** {proj['description']}")
                
                # Gate详情
                st.subheader("Gate详情")
                
                steps = execute_query('''
                    SELECT ps.*, u.username as submitter_name, r.username as reviewer_name
                    FROM project_steps ps
                    LEFT JOIN users u ON ps.submitted_by = u.id
                    LEFT JOIN users r ON ps.reviewed_by = r.id
                    WHERE ps.project_id = ?
                    ORDER BY ps.stage
                ''', (proj['id'],), fetch=True)
                
                if steps:
                    for step in steps:
                        step_status_map = {'pending': '⏳ 待审核', 'approved': '✅ 已通过', 'rejected': '❌ 已驳回'}
                        
                        step_stage = step['stage']
                        st.markdown(f"**{format_gate(step_stage)}** - {step_status_map.get(step['status'], step['status'])}")
                        st.caption(f"工程目的：{STAGE_PURPOSES.get(step_stage, '-')}")
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.write(f"提交人: {step['submitter_name'] or '-'}")
                        with col2:
                            st.write(f"提交时间: {format_datetime_display(step['submitted_at'])}")
                        with col3:
                            st.write(f"审核人: {step['reviewer_name'] or '-'}")
                        
                        if step['review_comment']:
                            st.write(f"审核意见: {step['review_comment']}")
                        
                        # 查看该阶段的文件
                        step_files = execute_query('''
                            SELECT * FROM project_files WHERE step_id = ?
                        ''', (step['id'],), fetch=True)
                        
                        if step_files:
                            st.write("📁 Gate文件:")
                            for f in step_files:
                                col1, col2, col3 = st.columns([4, 1, 1])
                                with col1:
                                    st.write(f"  - {f['file_name']}")
                                with col2:
                                    if st.button("👁️ 查看", key=f"view_step_file_{f['id']}"):
                                        st.session_state['view_file_id'] = f['id']
                                        st.rerun()
                                with col3:
                                    # 通过 & 驳回按钮（仅在待审批状态显示）
                                    if f.get('approval_status') == 'pending':
                                        if st.button("✅ 通过", key=f"approve_file_{f['id']}"):
                                            execute_query(
                                                "UPDATE project_files SET approval_status = 'approved', approved_by = ?, approved_at = ? WHERE id = ?",
                                                (st.session_state['user']['id'], datetime.now(), f['id']),
                                                commit=True
                                            )
                                            # 发送消息给上传者
                                            add_message(f.get('upload_by'), '文件审核通过', f'您上传的文件"{f.get("title") or f.get("file_name")}"已审核通过')
                                            notify_and_rerun("文件已通过", level='success')
                                    else:
                                        # 显示当前状态简短标签
                                        st.markdown(f"<div style='font-size:12px;color:#666'>{f.get('approval_status')}</div>", unsafe_allow_html=True)
                                # 驳回原因输入和按钮单独一行以避免布局挤压
                                if f.get('approval_status') == 'pending':
                                    reject_col1, reject_col2 = st.columns([5,1])
                                    with reject_col1:
                                        reason = st.text_input(
                                            "",
                                            key=f"reject_reason_file_{f['id']}",
                                            placeholder="驳回原因（可选）",
                                            label_visibility="collapsed"
                                        )
                                    with reject_col2:
                                        if st.button("❌ 驳回", key=f"reject_file_{f['id']}"):
                                            if not reason:
                                                st.error("请填写驳回原因")
                                            else:
                                                execute_query(
                                                    "UPDATE project_files SET approval_status = 'rejected', approved_by = ?, approved_at = ?, approval_comment = ? WHERE id = ?",
                                                    (st.session_state['user']['id'], datetime.now(), reason, f['id']),
                                                    commit=True
                                                )
                                                add_message(f.get('upload_by'), '文件审核驳回', f'您上传的文件"{f.get("title") or f.get("file_name")}"被驳回，原因: {reason}')
                                                notify_and_rerun("已驳回", level='success')
                        # 审核操作
                        if step['status'] == 'pending':
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.button("✅ 通过", key=f"approve_step_{step['id']}"):
                                    execute_query('''
                                        UPDATE project_steps SET status = 'approved', reviewed_by = ?, reviewed_at = ?, review_comment = '审核通过'
                                        WHERE id = ?
                                    ''', (st.session_state['user']['id'], datetime.now(), step['id']), commit=True)

                                    # 更新 current_stage（如果不是最后阶段）
                                    if step['stage'] < TOTAL_STAGES:
                                        execute_query('''
                                            UPDATE projects SET current_stage = ?, updated_at = ?
                                            WHERE id = ?
                                        ''', (step['stage'] + 1, datetime.now(), proj['id']), commit=True)

                                    # 刷新项目总体状态（completed/in_progress/rejected）
                                    try:
                                        refresh_project_status(proj['id'])
                                    except Exception:
                                        pass

                                    # 发送消息
                                    add_message(proj['created_by'], '项目Gate审核通过', 
                                              f'您的项目"{proj["name"]}"{STAGE_NAMES.get(step["stage"], "")}已审核通过')

                                    notify_and_rerun("审核通过", level='success')
                            with col2:
                                comment = st.text_input(
                                    "",
                                    key=f"reject_reason_step_{proj['id']}_{step['id']}",
                                    placeholder="驳回原因",
                                    label_visibility="collapsed"
                                )
                                if st.button("❌ 驳回", key=f"reject_step_{step['id']}"):
                                    if not comment:
                                        st.error("请填写驳回原因")
                                    else:
                                        execute_query('''
                                            UPDATE project_steps SET status = 'rejected', reviewed_by = ?, reviewed_at = ?, review_comment = ?
                                            WHERE id = ?
                                        ''', (st.session_state['user']['id'], datetime.now(), comment, step['id']), commit=True)

                                        # 刷新项目总体状态
                                        try:
                                            refresh_project_status(proj['id'])
                                        except Exception:
                                            pass

                                        add_message(proj['created_by'], '项目Gate审核驳回', 
                                                  f'您的项目"{proj["name"]}"{STAGE_NAMES.get(step["stage"], "")}已被驳回，原因: {comment}')

                                        notify_and_rerun("已驳回", level='success')
                        st.markdown("---")
                else:
                    st.info("暂无Gate数据")
    else:
        st.info("暂无项目数据")

def render_admin_logs():
    """管理端日志查看"""
    st.title("📝 日志查看")
    
    # 筛选条件
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        username_search = st.text_input("用户名", placeholder="输入用户名")
    
    with col2:
        module_filter = st.selectbox("功能模块", ["全部", "auth", "organizations", "users", "projects", "files", "indicators"],
                                    format_func=lambda x: {"全部": "全部", "auth": "认证", "organizations": "机构管理", 
                                                          "users": "账号管理", "projects": "项目管理", 
                                                          "files": "文件管理", "indicators": "指标管理"}.get(x, x))
    
    with col3:
        time_range = st.selectbox("时间范围", ["全部", "1天", "7天", "30天", "自定义"], index=3)
    
    with col4:
        if time_range == "自定义":
            start_date = st.date_input("开始日期")
            end_date = st.date_input("结束日期")
        else:
            page_size = st.selectbox("每页条数", [20, 50, 100], index=1, key="admin_logs_page_size")
    
    if time_range == "自定义":
        page_size = st.selectbox("每页条数", [20, 50, 100], index=1, key="admin_logs_page_size_custom")

    # 构建筛选条件，先计数再分页读取，避免日志量大时一次性加载全表。
    where_sql = " WHERE 1=1"
    params = []
    
    if username_search:
        where_sql += " AND username LIKE ?"
        params.append(f"%{username_search}%")
    
    if module_filter != "全部":
        where_sql += " AND module = ?"
        params.append(module_filter)
    
    if time_range == "1天":
        where_sql += " AND created_at >= datetime('now', '-1 day')"
    elif time_range == "7天":
        where_sql += " AND created_at >= datetime('now', '-7 days')"
    elif time_range == "30天":
        where_sql += " AND created_at >= datetime('now', '-30 days')"
    elif time_range == "自定义":
        where_sql += " AND created_at >= ? AND created_at <= ?"
        params.extend([f"{start_date} 00:00:00", f"{end_date} 23:59:59"])

    count_rows = execute_query(f"SELECT COUNT(*) as cnt FROM logs{where_sql}", params, fetch=True)
    total_count = int(count_rows[0]['cnt']) if count_rows else 0
    total_pages = max(1, (total_count + page_size - 1) // page_size)
    page_key = "admin_logs_page_no"
    current_page = int(st.session_state.get(page_key, 1) or 1)
    if current_page > total_pages:
        st.session_state[page_key] = total_pages
    elif current_page < 1:
        st.session_state[page_key] = 1
    page_no = st.number_input(
        "页码",
        min_value=1,
        max_value=total_pages,
        value=1,
        step=1,
        key=page_key
    )
    offset = (page_no - 1) * page_size

    logs = execute_query(
        f"SELECT * FROM logs{where_sql} ORDER BY created_at DESC LIMIT ? OFFSET ?",
        tuple(params + [page_size, offset]),
        fetch=True
    )

    if logs:
        df = pd.DataFrame(logs)
        df = df[['id', 'username', 'org_name', 'action', 'module', 'ip_address', 'created_at']]
        df.columns = ['日志编号', '用户名', '机构名', '操作', '功能模块', 'IP地址', '操作时间']
        st.caption(f"共 {total_count} 条日志，当前第 {page_no}/{total_pages} 页")
        st.dataframe(df, use_container_width=True, hide_index=True)
        
        export_limit = 5000
        if total_count > export_limit:
            st.warning(f"当前筛选结果 {total_count} 条，页面导出最多生成前 {export_limit} 条，超大数据请分时间段导出。")

        if st.button("📥 导出当前筛选结果", key="export_filtered_logs"):
            export_rows = execute_query(
                f"SELECT * FROM logs{where_sql} ORDER BY created_at DESC LIMIT ?",
                tuple(params + [export_limit]),
                fetch=True
            )
            export_df = pd.DataFrame(export_rows)
            export_df = export_df[['id', 'username', 'org_name', 'action', 'module', 'ip_address', 'created_at']]
            export_df.columns = ['日志编号', '用户名', '机构名', '操作', '功能模块', 'IP地址', '操作时间']
            export_to_excel(export_df, "操作日志")
    else:
        st.info("暂无日志数据")

def render_admin_export():
    """管理端数据导出"""
    st.title("📥 数据导出")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📤 导出数据")
        
        export_type = st.selectbox("选择导出类型", [
            "用户数据", "机构数据", "项目数据", "项目文件数据", "用户日志数据"
        ])
        
        if st.button("📥 导出Excel", use_container_width=True):
            if export_type == "用户数据":
                data = execute_query('''
                    SELECT u.id, u.username, u.real_name, u.phone, u.email, u.role, o.name as org_name, u.status, u.created_at
                    FROM users u LEFT JOIN organizations o ON u.org_id = o.id
                ''', fetch=True)
                df = pd.DataFrame(data) if data else pd.DataFrame()
                df.columns = ['用户ID', '用户名', '姓名', '手机号', '邮箱', '角色', '所属机构', '状态', '创建时间']
            
            elif export_type == "机构数据":
                data = execute_query('''
                    SELECT id, name, org_type, credit_code, legal_person, contact_person, 
                           contact_phone, contact_email, address, status, created_at
                    FROM organizations
                ''', fetch=True)
                df = pd.DataFrame(data) if data else pd.DataFrame()
                df.columns = ['机构ID', '机构名称', '机构类型', '信用代码', '法定代表人', '联系人', 
                             '联系电话', '联系邮箱', '机构地址', '状态', '创建时间']
            
            elif export_type == "项目数据":
                data = execute_query('''
                    SELECT p.id, p.name, o.name as org_name, p.category, p.subcategory, 
                           p.current_stage, p.status, p.created_at
                    FROM projects p JOIN organizations o ON p.org_id = o.id
                ''', fetch=True)
                df = pd.DataFrame(data) if data else pd.DataFrame()
                df.columns = ['项目ID', '项目名称', '所属机构', '项目分类', '子分类', '当前Gate', '状态', '创建时间']
            
            elif export_type == "项目文件数据":
                data = execute_query('''
                    SELECT pf.id, pf.title, pf.file_name, pf.category, pf.subcategory, 
                           pf.approval_status, u.username as uploader, pf.upload_at
                    FROM project_files pf LEFT JOIN users u ON pf.upload_by = u.id
                ''', fetch=True)
                df = pd.DataFrame(data) if data else pd.DataFrame()
                df.columns = ['文件ID', '文件标题', '文件名', '分类', '子分类', '审批状态', '上传者', '上传时间']
            
            elif export_type == "用户日志数据":
                export_limit = 5000
                total_rows = execute_query(
                    "SELECT COUNT(*) as cnt FROM logs WHERE created_at >= datetime('now', '-30 days')",
                    fetch=True
                )
                total_count = int(total_rows[0]['cnt']) if total_rows else 0
                if total_count > export_limit:
                    st.warning(f"最近30天日志共 {total_count} 条，本次导出前 {export_limit} 条；如需全量请按时间分段导出。")
                data = execute_query('''
                    SELECT id, user_id, username, org_name, action, module, ip_address, details, created_at
                    FROM logs
                    WHERE created_at >= datetime('now', '-30 days')
                    ORDER BY created_at DESC
                    LIMIT ?
                ''', (export_limit,), fetch=True)
                df = pd.DataFrame(data) if data else pd.DataFrame()
                if not df.empty:
                    df.columns = ['日志ID', '用户ID', '用户名', '机构名', '操作', '模块', 'IP地址', '详情', '操作时间']
            
            if not df.empty:
                export_to_excel(df, export_type)
            else:
                st.warning("暂无数据可导出")
    
    with col2:
        st.subheader("📥 导入数据")
        
        import_type = st.selectbox("选择导入类型", [
            "用户数据", "机构数据"
        ])
        
        uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'])
        
        import_policy = st.selectbox(
            "导入策略",
            ["仅新增，跳过已有数据", "按唯一字段更新已有数据"],
            help="默认仅新增，不会覆盖系统已有记录；确需更新时再选择第二项。"
        )
        
        def clean_import_value(value, default=""):
            if pd.isna(value):
                return default
            text = str(value).strip()
            return text if text else default
        
        def get_import_value(row, aliases, default=""):
            for alias in aliases:
                if alias in row:
                    value = clean_import_value(row.get(alias), default=None)
                    if value is not None:
                        return value
            return default
        
        if uploaded_file:
            try:
                uploaded_file.seek(0)
                df = pd.read_excel(uploaded_file)
                df = df.where(pd.notna(df), None)
                
                preview_rows = []
                if import_type == "用户数据":
                    for index, row in df.iterrows():
                        username = get_import_value(row, ["用户名", "username"])
                        phone = get_import_value(row, ["手机号", "phone"])
                        email = get_import_value(row, ["邮箱", "email"])
                        if not username:
                            preview_rows.append({"行号": index + 2, "唯一字段": "-", "处理方式": "无效", "原因": "缺少用户名"})
                            continue
                        
                        conflict = execute_query('''
                            SELECT id, username, phone, email FROM users
                            WHERE username = ? OR (? != '' AND phone = ?) OR (? != '' AND email = ?)
                            LIMIT 1
                        ''', (username, phone, phone, email, email), fetch=True)
                        if conflict:
                            reason = "用户名/手机号/邮箱已存在"
                            action = "更新" if import_policy == "按唯一字段更新已有数据" and conflict[0]['username'] == username else "跳过"
                            if action == "更新" and (
                                (phone and conflict[0]['phone'] not in ("", None) and conflict[0]['phone'] != phone) or
                                (email and conflict[0]['email'] not in ("", None) and conflict[0]['email'] != email)
                            ):
                                action = "跳过"
                                reason = "同用户名记录的手机号或邮箱不一致，需手工处理"
                            preview_rows.append({"行号": index + 2, "唯一字段": username, "处理方式": action, "原因": reason})
                        else:
                            preview_rows.append({"行号": index + 2, "唯一字段": username, "处理方式": "新增", "原因": ""})
                
                elif import_type == "机构数据":
                    for index, row in df.iterrows():
                        name = get_import_value(row, ["机构名称", "name"])
                        credit_code = get_import_value(row, ["信用代码", "credit_code"])
                        if not name:
                            preview_rows.append({"行号": index + 2, "唯一字段": "-", "处理方式": "无效", "原因": "缺少机构名称"})
                            continue
                        
                        conflict = execute_query('''
                            SELECT id FROM organizations
                            WHERE (? != '' AND credit_code = ?) OR name = ?
                            LIMIT 1
                        ''', (credit_code, credit_code, name), fetch=True)
                        if conflict:
                            action = "更新" if import_policy == "按唯一字段更新已有数据" else "跳过"
                            preview_rows.append({"行号": index + 2, "唯一字段": credit_code or name, "处理方式": action, "原因": "机构名称或信用代码已存在"})
                        else:
                            preview_rows.append({"行号": index + 2, "唯一字段": credit_code or name, "处理方式": "新增", "原因": ""})
                
                preview_df = pd.DataFrame(preview_rows)
                if not preview_df.empty:
                    st.markdown("#### 导入预览")
                    summary = preview_df["处理方式"].value_counts().to_dict()
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("总行数", len(preview_df))
                    c2.metric("新增", summary.get("新增", 0))
                    c3.metric("跳过", summary.get("跳过", 0))
                    c4.metric("无效", summary.get("无效", 0))
                    if summary.get("更新", 0):
                        st.warning(f"当前策略将更新 {summary.get('更新', 0)} 条已有数据，请确认后再导入。")
                    st.dataframe(preview_df, use_container_width=True, hide_index=True)
                
                if st.button("📤 确认导入", use_container_width=True):
                    imported_count = 0
                    updated_count = 0
                    skipped_count = 0
                    invalid_count = 0
                    
                    if import_type == "用户数据":
                        org_rows = execute_query("SELECT id, name FROM organizations", fetch=True)
                        org_name_to_id = {org['name']: org['id'] for org in (org_rows or [])}
                        
                        for _, row in df.iterrows():
                            username = get_import_value(row, ["用户名", "username"])
                            real_name = get_import_value(row, ["姓名", "real_name"])
                            phone = get_import_value(row, ["手机号", "phone"])
                            email = get_import_value(row, ["邮箱", "email"])
                            role = get_import_value(row, ["角色", "role"], "org_user")
                            status = get_import_value(row, ["状态", "status"], "active")
                            org_name = get_import_value(row, ["所属机构", "org_name"])
                            org_id_value = org_name_to_id.get(org_name)
                            
                            if not username:
                                invalid_count += 1
                                continue
                            
                            conflict = execute_query('''
                                SELECT id, username, phone, email FROM users
                                WHERE username = ? OR (? != '' AND phone = ?) OR (? != '' AND email = ?)
                                LIMIT 1
                            ''', (username, phone, phone, email, email), fetch=True)
                            
                            if conflict:
                                existing = conflict[0]
                                can_update = (
                                    import_policy == "按唯一字段更新已有数据" and
                                    existing['username'] == username and
                                    not (phone and existing['phone'] not in ("", None) and existing['phone'] != phone) and
                                    not (email and existing['email'] not in ("", None) and existing['email'] != email)
                                )
                                if not can_update:
                                    skipped_count += 1
                                    continue
                                execute_query('''
                                    UPDATE users
                                    SET role = ?, org_id = COALESCE(?, org_id), real_name = ?, phone = ?, email = ?, status = ?, updated_at = ?
                                    WHERE id = ?
                                ''', (role, org_id_value, real_name, phone, email, status, datetime.now(), existing['id']), commit=True)
                                updated_count += 1
                            else:
                                default_password = get_default_password_for_role(role)
                                password_hash = hash_password(default_password)
                                execute_query('''
                                    INSERT INTO users (username, password_hash, role, org_id, real_name, phone, email, status, must_change_password)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
                                ''', (username, password_hash, role, org_id_value, real_name, phone, email, status), commit=True)
                                imported_count += 1
                    
                    elif import_type == "机构数据":
                        for _, row in df.iterrows():
                            name = get_import_value(row, ["机构名称", "name"])
                            org_type = get_import_value(row, ["机构类型", "org_type"])
                            credit_code = get_import_value(row, ["信用代码", "credit_code"])
                            legal_person = get_import_value(row, ["法定代表人", "legal_person"])
                            contact_person = get_import_value(row, ["联系人", "contact_person"])
                            contact_phone = get_import_value(row, ["联系电话", "contact_phone"])
                            contact_email = get_import_value(row, ["联系邮箱", "contact_email"])
                            address = get_import_value(row, ["机构地址", "address"])
                            status = get_import_value(row, ["状态", "status"], "active")
                            
                            if not name:
                                invalid_count += 1
                                continue
                            
                            conflict = execute_query('''
                                SELECT id FROM organizations
                                WHERE (? != '' AND credit_code = ?) OR name = ?
                                LIMIT 1
                            ''', (credit_code, credit_code, name), fetch=True)
                            
                            if conflict:
                                if import_policy != "按唯一字段更新已有数据":
                                    skipped_count += 1
                                    continue
                                execute_query('''
                                    UPDATE organizations
                                    SET name = ?, org_type = ?, credit_code = ?, legal_person = ?, contact_person = ?,
                                        contact_phone = ?, contact_email = ?, address = ?, status = ?, updated_at = ?
                                    WHERE id = ?
                                ''', (
                                    name, org_type, credit_code, legal_person, contact_person,
                                    contact_phone, contact_email, address, status, datetime.now(), conflict[0]['id']
                                ), commit=True)
                                updated_count += 1
                            else:
                                execute_query('''
                                    INSERT INTO organizations (name, org_type, credit_code, legal_person,
                                                               contact_person, contact_phone, contact_email, address, status)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (
                                    name, org_type, credit_code, legal_person,
                                    contact_person, contact_phone, contact_email, address, status
                                ), commit=True)
                                imported_count += 1
                 
                    add_log(st.session_state['user']['id'], st.session_state['user']['username'], '',
                            '导入数据', 'export',
                            f'{import_type}: 新增{imported_count}条，更新{updated_count}条，跳过{skipped_count}条，无效{invalid_count}条',
                            get_client_ip())
                    notify_and_rerun(
                        f"导入完成：新增 {imported_count} 条，更新 {updated_count} 条，跳过 {skipped_count} 条，无效 {invalid_count} 条",
                        level='success'
                    )
            except Exception as e:
                st.error(f"导入失败: {e}")

def export_to_excel(df, filename):
    """导出DataFrame到Excel"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='数据')
        
        # 获取工作表并设置样式
        workbook = writer.book
        worksheet = writer.sheets['数据']
        
        # 设置列宽
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    
    st.download_button(
        label=f"📥 下载 {filename}.xlsx",
        data=output,
        file_name=f"{filename}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def render_admin_approval():
    """管理端审批待办"""
    st.title("✅ 审批待办")
    
    tab1, tab2 = st.tabs(["项目Gate审批", "项目文件审批"])
    
    with tab1:
        st.subheader("项目Gate审批")
        
        pending_steps = execute_query('''
            SELECT ps.*, p.name as project_name, o.name as org_name, u.username as submitter
            FROM project_steps ps
            JOIN projects p ON ps.project_id = p.id
            JOIN organizations o ON p.org_id = o.id
            LEFT JOIN users u ON ps.submitted_by = u.id
            WHERE ps.status = 'pending'
            ORDER BY ps.submitted_at ASC
        ''', fetch=True)
        
        if pending_steps:
            for step in pending_steps:
                with st.container():
                    st.markdown(f"""
                    <div style="background: rgba(255,255,255,0.03); padding: 15px; border-radius: 10px; margin: 10px 0; color: var(--text-on-dark);">
                        <h4 style="margin:0; padding:0;">{step['project_name']} - {format_gate(step['stage'])}</h4>
                        <p style="margin:4px 0 0 0; font-size:13px; color: rgba(255,255,255,0.9);">所属机构: {step['org_name']} | 提交人: {step['submitter'] or '-'} | 提交时间: {format_datetime_display(step['submitted_at'])}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    col1, col2, col3 = st.columns([2, 2, 1])
                    
                    with col1:
                        if st.button("✅ 通过", key=f"approve_{step['id']}", use_container_width=True):
                            execute_query('''
                                UPDATE project_steps SET status = 'approved', reviewed_by = ?, reviewed_at = ?, review_comment = '审核通过'
                                WHERE id = ?
                            ''', (st.session_state['user']['id'], datetime.now(), step['id']), commit=True)
                            
                            # 更新项目
                            if step['stage'] < TOTAL_STAGES:
                                execute_query('''
                                    UPDATE projects SET current_stage = ?, updated_at = ?
                                    WHERE id = ?
                                ''', (step['stage'] + 1, datetime.now(), step['project_id']), commit=True)

                            # 刷新项目总体状态
                            try:
                                refresh_project_status(step['project_id'])
                            except Exception:
                                pass
                            
                            # 获取项目创建者
                            proj = execute_query("SELECT created_by, name FROM projects WHERE id = ?", (step['project_id'],), fetch=True)
                            if proj:
                                add_message(proj[0]['created_by'], '项目Gate审核通过', 
                                          f'您的项目"{proj[0]["name"]}"{STAGE_NAMES.get(step["stage"], "")}已审核通过')
                            
                            notify_and_rerun("审核通过", level='success')
                    with col2:
                        reject_reason = st.text_input("", key=f"reason_{step['id']}", placeholder="驳回原因", label_visibility="collapsed")
                    
                    with col3:
                        if st.button("❌ 驳回", key=f"reject_{step['id']}", use_container_width=True):
                            if not reject_reason:
                                st.error("请填写驳回原因")
                            else:
                                execute_query('''
                                    UPDATE project_steps SET status = 'rejected', reviewed_by = ?, reviewed_at = ?, review_comment = ?
                                    WHERE id = ?
                                ''', (st.session_state['user']['id'], datetime.now(), reject_reason, step['id']), commit=True)
                                
                                execute_query('''
                                    UPDATE projects SET status = 'rejected', updated_at = ?
                                    WHERE id = ?
                                ''', (datetime.now(), step['project_id']), commit=True)
                                
                                proj = execute_query("SELECT created_by, name FROM projects WHERE id = ?", (step['project_id'],), fetch=True)
                                if proj:
                                    add_message(proj[0]['created_by'], '项目Gate审核驳回', 
                                              f'您的项目"{proj[0]["name"]}"{STAGE_NAMES.get(step["stage"], "")}已被驳回，原因: {reject_reason}')
                                
                                notify_and_rerun("已驳回", level='success')
        else:
            st.info("暂无待审批的项目阶段")
    
    with tab2:
        st.subheader("项目文件审批")
        
        pending_files = execute_query('''
            SELECT pf.*, p.name as project_name, o.name as org_name, u.username as uploader
            FROM project_files pf
            LEFT JOIN projects p ON pf.project_id = p.id
            LEFT JOIN organizations o ON p.org_id = o.id
            LEFT JOIN users u ON pf.upload_by = u.id
            WHERE pf.approval_status = 'pending'
            ORDER BY pf.upload_at ASC
        ''', fetch=True)
        
        if pending_files:
            for file in pending_files:
                with st.container():
                    st.markdown(f"""
                    <div style="background: rgba(255,255,255,0.03); padding: 15px; border-radius: 10px; margin: 10px 0; color: var(--text-on-dark);">
                        <h4 style="margin:0; padding:0;">{file['title']}</h4>
                        <p style="margin:4px 0 0 0; font-size:13px; color: rgba(255,255,255,0.9);">文件名: {file['file_name']} | 所属机构: {file['org_name'] or '-'} | 上传者: {file['uploader'] or '-'} | 上传时间: {format_datetime_display(file['upload_at'])}</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # 查看文件
                    if st.button("👁️ 查看文件", key=f"view_file_{file['id']}"):
                        st.session_state['view_file_id'] = file['id']
                    
                    col1, col2, col3 = st.columns([2, 2, 1])
                    
                    with col1:
                        if st.button("✅ 通过", key=f"approve_file_{file['id']}", use_container_width=True):
                            execute_query('''
                                UPDATE project_files SET approval_status = 'approved', approved_by = ?, approved_at = ?
                                WHERE id = ?
                            ''', (st.session_state['user']['id'], datetime.now(), file['id']), commit=True)
                            
                            if file['upload_by']:
                                add_message(file['upload_by'], '文件审核通过', f'您上传的文件"{file["title"]}"已审核通过')
                            
                            notify_and_rerun("审核通过", level='success')
                    with col2:
                        file_reject_reason = st.text_input("", key=f"file_reason_{file['id']}", placeholder="驳回原因", label_visibility="collapsed")
                    
                    with col3:
                        if st.button("❌ 驳回", key=f"reject_file_{file['id']}", use_container_width=True):
                            if not file_reject_reason:
                                st.error("请填写驳回原因")
                            else:
                                execute_query('''
                                    UPDATE project_files SET approval_status = 'rejected', approved_by = ?, approved_at = ?, approval_comment = ?
                                    WHERE id = ?
                                ''', (st.session_state['user']['id'], datetime.now(), file_reject_reason, file['id']), commit=True)
                                
                                if file['upload_by']:
                                    add_message(file['upload_by'], '文件审核驳回', 
                                              f'您上传的文件"{file["title"]}"已被驳回，原因: {file_reject_reason}')
                                
                                notify_and_rerun("已驳回", level='success')
        else:
            st.info("暂无待审批的项目文件")

def delete_user_messages(user_id, message_ids):
    """删除当前用户指定消息，防止跨账号删除。"""
    clean_ids = [int(message_id) for message_id in message_ids if message_id]
    if not clean_ids:
        return 0
    placeholders = ",".join(["?"] * len(clean_ids))
    execute_query(
        f"DELETE FROM messages WHERE user_id = ? AND id IN ({placeholders})",
        tuple([user_id] + clean_ids),
        commit=True
    )
    return len(clean_ids)

def render_messages_page(title="📨 消息通知", key_prefix="messages"):
    """统一渲染管理端/机构端消息通知。"""
    st.title(title)
    user = st.session_state['user']

    filter_col, size_col = st.columns([2, 1])
    with filter_col:
        read_filter = st.selectbox(
            "消息状态",
            ["全部", "未读", "已读"],
            key=f"{key_prefix}_read_filter"
        )
    with size_col:
        page_size = st.selectbox(
            "每页条数",
            [10, 20, 50],
            index=1,
            key=f"{key_prefix}_page_size"
        )

    where_sql = " WHERE user_id = ?"
    params = [user['id']]
    if read_filter == "未读":
        where_sql += " AND is_read = 0"
    elif read_filter == "已读":
        where_sql += " AND is_read = 1"

    count_rows = execute_query(f"SELECT COUNT(*) as cnt FROM messages{where_sql}", params, fetch=True)
    total_count = int(count_rows[0]['cnt']) if count_rows else 0
    total_pages = max(1, (total_count + page_size - 1) // page_size)
    page_key = f"{key_prefix}_page_no"
    current_page = int(st.session_state.get(page_key, 1) or 1)
    if current_page > total_pages:
        st.session_state[page_key] = total_pages
    elif current_page < 1:
        st.session_state[page_key] = 1
    page_no = st.number_input(
        "页码",
        min_value=1,
        max_value=total_pages,
        value=1,
        step=1,
        key=page_key
    )
    offset = (page_no - 1) * page_size

    messages = execute_query(
        f"SELECT * FROM messages{where_sql} ORDER BY created_at DESC LIMIT ? OFFSET ?",
        tuple(params + [page_size, offset]),
        fetch=True
    )

    message_ids = [msg['id'] for msg in (messages or [])]
    selected_ids = [
        message_id for message_id in message_ids
        if st.session_state.get(f"{key_prefix}_select_{message_id}", False)
    ]
    all_selected = bool(message_ids) and len(selected_ids) == len(message_ids)

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        if st.button("✅ 一键已读", use_container_width=True, key=f"{key_prefix}_read_all"):
            execute_query("UPDATE messages SET is_read = 1 WHERE user_id = ?", (user['id'],), commit=True)
            notify_and_rerun("所有消息已标记为已读", level='success')
    with col2:
        if st.button("🧹 清空所有通知", use_container_width=True, key=f"{key_prefix}_clear_all"):
            execute_query("DELETE FROM messages WHERE user_id = ?", (user['id'],), commit=True)
            notify_and_rerun("所有通知已清空", level='success')
    with col3:
        if st.button("🧽 删除已读", use_container_width=True, key=f"{key_prefix}_delete_read"):
            read_rows = execute_query(
                "SELECT COUNT(*) as cnt FROM messages WHERE user_id = ? AND is_read = 1",
                (user['id'],),
                fetch=True
            )
            read_count = read_rows[0]['cnt'] if read_rows else 0
            if read_count:
                execute_query("DELETE FROM messages WHERE user_id = ? AND is_read = 1", (user['id'],), commit=True)
                notify_and_rerun(f"已删除 {read_count} 条已读通知", level='success')
            else:
                st.warning("暂无已读通知可删除")
    with col4:
        select_all_label = "⬜ 取消全选" if all_selected else "☑️ 全选当前页"
        if st.button(select_all_label, use_container_width=True, key=f"{key_prefix}_toggle_select_all"):
            target_value = not all_selected
            for message_id in message_ids:
                st.session_state[f"{key_prefix}_select_{message_id}"] = target_value
            st.rerun()
    with col5:
        if st.button("🗑️ 删除已选", use_container_width=True, key=f"{key_prefix}_delete_selected"):
            if selected_ids:
                deleted_count = delete_user_messages(user['id'], selected_ids)
                for message_id in selected_ids:
                    st.session_state.pop(f"{key_prefix}_select_{message_id}", None)
                notify_and_rerun(f"已删除 {deleted_count} 条通知", level='success')
            else:
                st.warning("请先勾选要删除的通知")

    if selected_ids:
        st.caption(f"已选择 {len(selected_ids)} 条通知")
    if total_count:
        st.caption(f"共 {total_count} 条通知，当前第 {page_no}/{total_pages} 页")

    if messages:
        for msg in messages:
            msg_id = msg['id']
            select_col, content_col, action_col = st.columns([0.06, 0.76, 0.18])
            with select_col:
                st.checkbox(
                    "选择",
                    key=f"{key_prefix}_select_{msg_id}",
                    label_visibility="collapsed"
                )

            with content_col:
                read_status = "✅" if msg['is_read'] else "🔔"
                left_color = '#4caf50' if msg['is_read'] else 'var(--accent)'
                safe_title = html.escape(str(msg['title'] or ''))
                safe_content = html.escape(str(msg['content'] or ''))
                safe_created_at = html.escape(format_datetime_display(msg['created_at']))
                st.markdown(f"""
                <div style="background: var(--surface-bg); padding: 15px; border-radius: 8px; margin: 10px 0; border-left: 4px solid {left_color}; color: rgba(255,255,255,0.96);">
                    <h4 style="margin:0 0 6px 0; color: rgba(255,255,255,0.98);">{read_status} {safe_title}</h4>
                    <p style="margin:0 0 6px 0; color: rgba(255,255,255,0.92);">{safe_content}</p>
                    <small style="color: rgba(255,255,255,0.65);">{safe_created_at}</small>
                </div>
                """, unsafe_allow_html=True)

            with action_col:
                st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
                if not msg['is_read']:
                    if st.button("标记已读", key=f"{key_prefix}_read_{msg_id}", use_container_width=True):
                        execute_query("UPDATE messages SET is_read = 1 WHERE user_id = ? AND id = ?", (user['id'], msg_id), commit=True)
                        notify_and_rerun("已标记为已读", level='success')
                if st.button("删除", key=f"{key_prefix}_delete_{msg_id}", use_container_width=True):
                    delete_user_messages(user['id'], [msg_id])
                    st.session_state.pop(f"{key_prefix}_select_{msg_id}", None)
                    notify_and_rerun("通知已删除", level='success')
    else:
        st.info("暂无消息")

def render_admin_messages():
    """管理端消息通知"""
    render_messages_page("📨 消息通知", "admin_messages")

def render_admin_indicators():
    """管理端项目智库管理"""
    st.title("📚 项目智库管理")

    # 一次性迁移：将历史 project_files 的文件名与磁盘文件重命名为 项目ID_原名 格式
    def migrate_files_to_project_prefix():
        migrated = 0
        failed = 0
        rows = execute_query("SELECT id, project_id, file_name, file_path FROM project_files", fetch=True)
        for r in rows:
            try:
                pid = r['project_id']
                if not pid:
                    continue
                old_name = r['file_name'] or ''
                # 如果已经以 projectID_ 开头，跳过
                if old_name.startswith(f"{pid}_"):
                    continue
                old_path = r['file_path']
                if not old_path or not os.path.exists(old_path):
                    failed += 1
                    continue
                dirpath = os.path.dirname(old_path)
                new_name = f"{pid}_{old_name}"
                new_path = os.path.join(dirpath, new_name)
                # 如果目标已存在，跳过或使用唯一后缀
                if os.path.exists(new_path):
                    # 改用时间戳后缀
                    ts = datetime.now().strftime("%Y%m%d%H%M%S")
                    new_name = f"{pid}_{ts}_{old_name}"
                    new_path = os.path.join(dirpath, new_name)
                os.rename(old_path, new_path)
                # 更新数据库字段：file_name 与 file_path
                execute_query("UPDATE project_files SET file_name = ?, file_path = ? WHERE id = ?", (new_name, new_path, r['id']), commit=True)
                migrated += 1
            except Exception:
                failed += 1
        return migrated, failed

    # 使用数据库持久化迁移标记，避免服务重启后再次全表扫描
    files_migrated_flag = get_app_meta("files_prefixed_v1", "0")
    if files_migrated_flag != "1":
        try:
            mig, fail = migrate_files_to_project_prefix()
            set_app_meta("files_prefixed_v1", "1")
            st.success(f"已完成文件迁移: {mig} 个，失败: {fail} 个（如有失败请检查文件权限或路径）")
        except Exception as e:
            set_app_meta("files_prefixed_v1", "1")
            st.error(f"文件迁移时发生错误: {e}")
    
    tab1, tab2, tab3, tab4 = st.tabs(["当前项目", "项目文件管理", "指标管理", "项目文件评估"])

    # 点击“打开项目文件夹”后，自动切到“项目文件管理”页签
    if st.session_state.pop('indicators_jump_to_file_tab', False):
        components.html("""
        <script>
        (function () {
            const switchTab = () => {
                const tabs = window.parent.document.querySelectorAll('button[role="tab"]');
                for (const tab of tabs) {
                    if ((tab.innerText || '').trim() === '项目文件管理') {
                        tab.click();
                        return true;
                    }
                }
                return false;
            };
            switchTab() || setTimeout(switchTab, 120);
            setTimeout(switchTab, 360);
        })();
        </script>
        """, height=0)
    
    with tab1:
        st.subheader("当前项目")
        st.info("显示机构创建且处于审核流程中的项目；可直接打开项目文件夹查看已通过文件。")
        
        # 查询未完成的项目
        current_projects = execute_query('''
            SELECT p.*, o.name as org_name, 
                   (SELECT COUNT(*) FROM project_steps WHERE project_id = p.id AND status != 'approved') as pending_stages
            FROM projects p
            JOIN organizations o ON p.org_id = o.id
            WHERE p.status IN ('pending', 'in_progress', 'rejected')
            ORDER BY p.created_at DESC
        ''', fetch=True)
        
        if current_projects:
            project_ids = [proj['id'] for proj in current_projects]
            project_steps = []
            steps_by_project = {}
            files_by_step = {}
            project_file_stats = {}

            if project_ids:
                project_placeholders = ",".join(["?"] * len(project_ids))
                project_steps = execute_query(f'''
                    SELECT ps.*,
                           COALESCE(pf.file_count, 0) as file_count
                    FROM project_steps ps
                    LEFT JOIN (
                        SELECT step_id, COUNT(*) as file_count
                        FROM project_files
                        GROUP BY step_id
                    ) pf ON pf.step_id = ps.id
                    WHERE ps.project_id IN ({project_placeholders})
                    ORDER BY ps.project_id, ps.stage
                ''', tuple(project_ids), fetch=True)

                project_file_rows = execute_query(f'''
                    SELECT
                        project_id,
                        COUNT(*) AS total_cnt,
                        SUM(CASE WHEN approval_status = 'approved' THEN 1 ELSE 0 END) AS approved_cnt
                    FROM project_files
                    WHERE project_id IN ({project_placeholders})
                    GROUP BY project_id
                ''', tuple(project_ids), fetch=True)
                project_file_stats = {
                    row['project_id']: {
                        'total': int(row.get('total_cnt') or 0),
                        'approved': int(row.get('approved_cnt') or 0)
                    }
                    for row in project_file_rows
                }

                for step in project_steps:
                    steps_by_project.setdefault(step['project_id'], []).append(step)

                step_ids_with_files = [step['id'] for step in project_steps if step['file_count'] > 0]
                if step_ids_with_files:
                    step_placeholders = ",".join(["?"] * len(step_ids_with_files))
                    all_step_files = execute_query(
                        f"SELECT * FROM project_files WHERE step_id IN ({step_placeholders}) ORDER BY step_id, upload_at DESC",
                        tuple(step_ids_with_files),
                        fetch=True
                    )
                    for file_row in all_step_files:
                        files_by_step.setdefault(file_row['step_id'], []).append(file_row)

            for proj in current_projects:
                with st.expander(f"**{proj['name']}** - {proj['org_name']} (待处理Gate: {proj['pending_stages']})"):
                    st.write(f"**项目分类:** {PROJECT_CATEGORIES.get(proj['category'], {}).get('name', '-') if proj['category'] else '-'}")
                    st.write(f"**当前Gate:** {format_gate(proj['current_stage'])}")
                    st.write(f"**状态:** {proj['status']}")
                    st.write(f"**创建时间:** {format_datetime_display(proj['created_at'])}")

                    # 查看各Gate文件统计
                    gate_label_col, gate_button_col, _ = st.columns([1.0, 1.8, 4.2])
                    with gate_label_col:
                        st.markdown("**Gate文件:**")
                    with gate_button_col:
                        if st.button("📂 打开项目文件夹", key=f"open_proj_{proj['id']}", use_container_width=True):
                            proj_stats = project_file_stats.get(proj['id'], {'total': 0, 'approved': 0})
                            if proj_stats['total'] <= 0:
                                notify_and_rerun(f'项目“{proj["name"]}”还未上传文件', level='warning')
                            elif proj_stats['approved'] <= 0:
                                notify_and_rerun(f'项目“{proj["name"]}”暂无已通过文件', level='warning')
                            # 导出已通过文件到本地目录 1_<项目名>/2_<项目名>/3_<项目名>
                            export_result = ensure_project_export_dirs_and_copy(proj['id'])
                            st.session_state['export_info'] = export_result
                            if export_result:
                                st.session_state['export_base'] = export_result.get('base')
                            st.session_state['indicators_open_project'] = proj['id']
                            st.session_state['file_view_mode'] = '项目文件夹'
                            st.session_state['file_search_query'] = ''
                            st.session_state['file_search_input'] = ''
                            st.session_state['file_search_empty'] = False
                            st.session_state['indicators_jump_to_file_tab'] = True
                            st.experimental_rerun()
                    steps = steps_by_project.get(proj['id'], [])

                    if steps:
                        for step in steps:
                            step_stage = step['stage']
                            st.write(f"- {format_gate(step_stage)}: {step['file_count']} 个文件")
                            if step['file_count'] > 0:
                                step_files = files_by_step.get(step['id'], [])
                                for f in step_files:
                                    col1, col2, col3 = st.columns([4, 1, 1])
                                    with col1:
                                        st.write(f"  📄 {f['file_name']}")
                                    with col2:
                                        if st.button("👁️ 查看", key=f"admin_know_view_{f['id']}"):
                                            st.session_state['view_file_id'] = f['id']
                                            st.rerun()
                                    with col3:
                                        file_path = f.get('file_path') or os.path.join(UPLOAD_DIR, f.get('file_name') or '')
                                        if file_path and os.path.exists(file_path):
                                            try:
                                                render_cached_download_button(
                                                    file_path,
                                                    f.get('file_name', 'download'),
                                                    key=f"admin_know_dl_{f['id']}"
                                                )
                                            except Exception:
                                                st.write("读取失败")
                                        else:
                                            st.write("文件丢失")
        else:
            st.info("暂无当前项目")
    
    with tab2:
        st.subheader("项目文件管理")
        
        # 搜索（支持点击按钮或按 Enter）
        if 'file_search_query' not in st.session_state:
            st.session_state['file_search_query'] = ''
        if 'file_search_input' not in st.session_state:
            st.session_state['file_search_input'] = st.session_state.get('file_search_query', '')
        if 'file_search_empty' not in st.session_state:
            st.session_state['file_search_empty'] = False

        placeholder_text = "请输入内容" if (
            st.session_state.get('file_search_empty')
            and not str(st.session_state.get('file_search_input', '')).strip()
        ) else "输入项目或文件关键词查询"

        with st.form("file_search_form", clear_on_submit=False):
            col1, col2 = st.columns([5, 1])
            with col1:
                search_input = st.text_input(
                    "搜索项目或文件",
                    key="file_search_input",
                    placeholder=placeholder_text,
                    label_visibility="collapsed"
                )
            with col2:
                search_submitted = st.form_submit_button("🔍 搜索", use_container_width=True)

        # 用户重新输入后清除空提示
        if str(st.session_state.get('file_search_input', '')).strip():
            st.session_state['file_search_empty'] = False

        if search_submitted:
            query = (search_input or '').strip()
            if query:
                st.session_state['file_search_query'] = query
                st.session_state['file_search_empty'] = False
            else:
                # 空搜索视为清空筛选：显示全部文件
                st.session_state['file_search_query'] = ''
                st.session_state['file_search_empty'] = True

        search = (st.session_state.get('file_search_query') or '').strip()

        # 如果搜索框有内容，直接显示搜索结果
        if search:
            st.markdown(f"### 🔍 搜索结果: {search}")
            files = execute_query('''
                SELECT pf.*, p.name as project_name, u.username as uploader,
                       COALESCE(pf.category, p.category) as eff_category,
                       COALESCE(pf.subcategory, p.subcategory) as eff_subcategory
                FROM project_files pf
                LEFT JOIN projects p ON pf.project_id = p.id
                LEFT JOIN users u ON pf.upload_by = u.id
                WHERE pf.approval_status = 'approved' 
                AND (pf.file_name LIKE ? OR pf.title LIKE ? OR p.name LIKE ?)
                ORDER BY pf.upload_at DESC
            ''', (f'%{search}%', f'%{search}%', f'%{search}%'), fetch=True)
            
            if files:
                projects_in_search = {}
                for f in files:
                    p_id = f.get('project_id')
                    if not p_id:
                        p_id = f"unlinked_{f['id']}"
                        
                    if p_id not in projects_in_search:
                        p_name = f.get('project_name') or '未关联项目'
                        cat = f.get('eff_category') or '0'
                        sub = f.get('eff_subcategory')
                        prefix = f"{cat}-{sub}" if sub else f"{cat}"
                        
                        projects_in_search[p_id] = {
                            'name': p_name,
                            'display_name': f"{prefix} {p_name}",
                            'files': []
                        }
                    projects_in_search[p_id]['files'].append(f)

                for p_id, p_info in projects_in_search.items():
                    with st.expander(f"📁 {p_info['display_name']}", expanded=True):
                        for f in p_info['files']:
                            cat = f.get('eff_category') or '0'
                            sub = f.get('eff_subcategory')
                            file_prefix = f"{cat}-{sub}" if sub else f"{cat}"
                            fname = f.get('file_name') or ''
                            if not fname.startswith(f"{file_prefix}_"):
                                display_name = f"{file_prefix}_{fname}"
                            else:
                                display_name = fname

                            col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
                            with col1:
                                st.write(f"📄 {display_name}")
                            with col2:
                                st.write(f"{f['upload_at'][:10]}")
                            with col3:
                                if st.button("👁️ 查看", key=f"view_search_pf_{f['id']}"):
                                    st.session_state['view_file_id'] = f['id']
                                    st.rerun()
                            with col4:
                                file_path = f.get('file_path') or os.path.join(UPLOAD_DIR, fname)
                                if file_path and os.path.exists(file_path):
                                    try:
                                        render_cached_download_button(
                                            file_path,
                                            f.get('file_name', 'download'),
                                            key=f"dl_search_pf_{f['id']}"
                                        )
                                    except Exception:
                                        st.write("读取失败")
                                else:
                                    st.write("文件丢失")
                            with col5:
                                if st.button("📝 评估", key=f"eval_search_pf_{f['id']}"):
                                    st.session_state['indicators_eval_file'] = f['id']
                                    st.rerun()
            else:
                st.info("未找到匹配的文件")
        else:
            view_mode = st.radio("视图模式", ["项目分类文件夹", "项目文件夹"], horizontal=True, key="file_view_mode")
            
            if view_mode == "项目文件夹":
                st.markdown("### 📁 项目文件夹")
                all_approved_files = execute_query('''
                    SELECT pf.*, p.name as project_name, u.username as uploader,
                           COALESCE(pf.category, p.category) as eff_category,
                           COALESCE(pf.subcategory, p.subcategory) as eff_subcategory
                    FROM project_files pf
                    LEFT JOIN projects p ON pf.project_id = p.id
                    LEFT JOIN users u ON pf.upload_by = u.id
                    WHERE pf.approval_status = 'approved'
                    ORDER BY pf.upload_at DESC
                ''', fetch=True)
                
                if not all_approved_files:
                    st.info("暂无已通过文件的项目")
                else:
                    projects_dict = {}
                    for f in all_approved_files:
                        p_id = f.get('project_id')
                        if not p_id:
                            p_id = f"unlinked_{f['id']}"
                            
                        if p_id not in projects_dict:
                            p_name = f.get('project_name') or '未关联项目'
                            cat = f.get('eff_category') or '0'
                            sub = f.get('eff_subcategory')
                            prefix = f"{cat}-{sub}" if sub else f"{cat}"
                            projects_dict[p_id] = {
                                'name': p_name,
                                'display_name': f"{prefix} {p_name}",
                                'files': []
                            }
                        projects_dict[p_id]['files'].append(f)
                        
                    for p_id, p_info in projects_dict.items():
                        open_flag = (st.session_state.get('indicators_open_project') == p_id)
                        with st.expander(f"📁 {p_info['display_name']}", expanded=open_flag):
                            for f in p_info['files']:
                                cat = f.get('eff_category') or '0'
                                sub = f.get('eff_subcategory')
                                file_prefix = f"{cat}-{sub}" if sub else f"{cat}"
                                fname = f.get('file_name') or ''
                                if not fname.startswith(f"{file_prefix}_"):
                                    display_name = f"{file_prefix}_{fname}"
                                else:
                                    display_name = fname
                                    
                                col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
                                with col1:
                                    st.write(f"📄 {display_name}")
                                with col2:
                                    st.write(f"{f['upload_at'][:10]}")
                                with col3:
                                    if st.button("👁️ 查看", key=f"view_pf_proj_{f['id']}"):
                                        st.session_state['view_file_id'] = f['id']
                                        st.rerun()
                                with col4:
                                    file_path = f.get('file_path') or os.path.join(UPLOAD_DIR, fname)
                                    if file_path and os.path.exists(file_path):
                                        try:
                                            render_cached_download_button(
                                                file_path,
                                                f.get('file_name', 'download'),
                                                key=f"dl_pf_proj_{f['id']}"
                                            )
                                        except Exception:
                                            st.write("读取失败")
                                    else:
                                        st.write("文件丢失")
                                with col5:
                                    if st.button("📝 评估", key=f"eval_pf_proj_{f['id']}"):
                                        st.session_state['indicators_eval_file'] = f['id']
                                        st.rerun()
            else:
                # 分类文件夹视图
                st.markdown("### 📁 项目分类文件夹")
                
                # 定义固定的7个顶级文件夹
                folder_definitions = [
                    {'cat': '1', 'sub': '1', 'label': '1-1 多源数据采集'},
                    {'cat': '1', 'sub': '2', 'label': '1-2 专项调查研究'},
                    {'cat': '2', 'sub': '1', 'label': '2-1 财政绩效评估'},
                    {'cat': '2', 'sub': '2', 'label': '2-2 行政绩效评估'},
                    {'cat': '3', 'sub': '1', 'label': '3-1 企业管理咨询'},
                    {'cat': '3', 'sub': '2', 'label': '3-2 公共决策咨询'},
                    {'cat': '0', 'sub': None, 'label': '0 其他项目'}
                ]
                
                # 获取所有已通过文件及其关联的项目信息
                all_approved_files = execute_query('''
                    SELECT pf.*, p.name as project_name, u.username as uploader,
                           COALESCE(pf.category, p.category) as eff_category,
                           COALESCE(pf.subcategory, p.subcategory) as eff_subcategory
                    FROM project_files pf
                    LEFT JOIN projects p ON pf.project_id = p.id
                    LEFT JOIN users u ON pf.upload_by = u.id
                    WHERE pf.approval_status = 'approved'
                    ORDER BY pf.upload_at DESC
                ''', fetch=True)
    
                if not all_approved_files:
                    st.info("暂无已通过分类的文件")
                else:
                    for f_def in folder_definitions:
                        # 过滤属于该顶级文件夹的文件
                        cat = f_def['cat']
                        sub = f_def['sub']
                        
                        if sub is not None:
                            folder_files = [f for f in all_approved_files if str(f.get('eff_category')) == str(cat) and str(f.get('eff_subcategory')) == str(sub)]
                        else:
                            # 对于分类0，或者没有子分类的情况
                            folder_files = [f for f in all_approved_files if str(f.get('eff_category')) == str(cat) and not f.get('eff_subcategory')]
                            
                        if not folder_files:
                            continue
                        
                        # 检查是否有需要自动展开的项目属于该分类
                        target_proj_id = st.session_state.get('indicators_open_project')
                        has_open_proj = any((f.get('project_id') == target_proj_id) for f in folder_files)
                        
                        cat_state_key = f"cat_open_{cat}_{sub}"
                        if has_open_proj and target_proj_id:
                            st.session_state[cat_state_key] = True
                            
                        is_open = st.session_state.get(cat_state_key, False)
                        
                        # 该文件夹下有文件，显示可点击的顶级文件夹
                        if st.button(f"{'📂' if is_open else '📁'} {f_def['label']}", key=f"btn_cat_{cat}_{sub}", use_container_width=True):
                            st.session_state[cat_state_key] = not is_open
                            st.rerun()
                            
                        if st.session_state.get(cat_state_key, False):
                            # 按项目分组
                            projects_in_folder = {}
                            for f in folder_files:
                                p_id = f.get('project_id')
                                if not p_id:
                                    p_id = f"unlinked_{f['id']}"  # 无关联项目的文件独立分组
                                
                                if p_id not in projects_in_folder:
                                    p_name = f.get('project_name') or '未关联项目'
                                    prefix = f"{cat}-{sub}" if sub else f"{cat}"
                                    projects_in_folder[p_id] = {
                                        'name': p_name,
                                        'display_name': f"{prefix} {p_name}",
                                        'files': []
                                    }
                                projects_in_folder[p_id]['files'].append(f)
                                
                            # 遍历该顶级文件夹下的各个项目
                            for p_id, p_info in projects_in_folder.items():
                                open_flag = (st.session_state.get('indicators_open_project') == p_id)
                                with st.expander(f"📁 {p_info['display_name']}", expanded=open_flag):
                                    for f in p_info['files']:
                                        file_prefix = f"{cat}-{sub}" if sub else f"{cat}"
                                        fname = f.get('file_name') or ''
                                        # 如果文件名还没有此前缀，则在显示时动态补齐
                                        if not fname.startswith(f"{file_prefix}_"):
                                            display_name = f"{file_prefix}_{fname}"
                                        else:
                                            display_name = fname
                                            
                                        col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
                                        with col1:
                                            st.write(f"📄 {display_name}")
                                        with col2:
                                            st.write(f"{f['upload_at'][:10]}")
                                        with col3:
                                            if st.button("👁️ 查看", key=f"view_pf_{f['id']}"):
                                                st.session_state['view_file_id'] = f['id']
                                                st.rerun()
                                        with col4:
                                            file_path = f.get('file_path') or os.path.join(UPLOAD_DIR, fname)
                                            if file_path and os.path.exists(file_path):
                                                try:
                                                    render_cached_download_button(
                                                        file_path,
                                                        f.get('file_name', 'download'),
                                                        key=f"dl_pf_{f['id']}"
                                                    )
                                                except Exception:
                                                    st.write("读取失败")
                                            else:
                                                st.write("文件丢失")
                                        with col5:
                                            if st.button("📝 评估", key=f"eval_pf_{f['id']}"):
                                                st.session_state['indicators_eval_file'] = f['id']
                                                st.rerun()
    
    with tab3:
        st.subheader("指标管理")
        
        # 选择项目分类
        cat_select = st.selectbox("选择项目分类", list(PROJECT_CATEGORIES.keys()),
                                 format_func=lambda x: f"{x} - {PROJECT_CATEGORIES[x]['name']}")
        
        # 获取子分类
        subcats = PROJECT_CATEGORIES[cat_select].get('subcategories', {})
        if subcats:
            sub_select = st.selectbox("选择二级分类", list(subcats.keys()),
                                     format_func=lambda x: f"{x} - {subcats[x]}")
        else:
            sub_select = None
        
        # 获取指标列表
        if sub_select:
            indicators = execute_query('''
                SELECT * FROM indicator_library WHERE category = ? AND subcategory = ?
            ''', (cat_select, sub_select), fetch=True)
        else:
            indicators = execute_query('''
                SELECT * FROM indicator_library WHERE category = ?
            ''', (cat_select,), fetch=True)
        
        if indicators:
            st.markdown("#### 当前指标")
            
            for ind in indicators:
                col1, col2, col3, col4 = st.columns([3, 1, 2, 1])
                
                with col1:
                    st.write(f"**{ind['indicator_name']}**")
                
                with col2:
                    new_weight = st.number_input("权重", value=ind['weight'], min_value=0.0, max_value=100.0, 
                                                key=f"weight_{ind['id']}", label_visibility="collapsed")
                
                with col3:
                    new_desc = st.text_input("描述", value=ind['description'] or "", 
                                            key=f"desc_{ind['id']}", label_visibility="collapsed")
                
                with col4:
                    if st.button("💾", key=f"save_ind_{ind['id']}"):
                        execute_query('''
                            UPDATE indicator_library SET weight = ?, description = ?, updated_at = ?
                            WHERE id = ?
                        ''', (new_weight, new_desc, datetime.now(), ind['id']), commit=True)
                        st.success("已保存")
                    
                    if st.button("🗑️", key=f"del_ind_{ind['id']}"):
                        execute_query("DELETE FROM indicator_library WHERE id = ?", (ind['id'],), commit=True)
                        notify_and_rerun("已删除", level='success')
        # 添加新指标
        st.markdown("#### 添加新指标")
        
        with st.form("add_indicator_form"):
            new_ind_name = st.text_input("指标名称")
            new_ind_weight = st.number_input("权重", value=10.0, min_value=0.0, max_value=100.0)
            new_ind_desc = st.text_area("描述")
            
            if st.form_submit_button("添加指标"):
                if new_ind_name:
                    execute_query('''
                        INSERT INTO indicator_library (category, subcategory, indicator_name, weight, description)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (cat_select, sub_select, new_ind_name, new_ind_weight, new_ind_desc), commit=True)
                    notify_and_rerun("指标添加成功", level='success')
    with tab4:
        st.subheader("项目文件评估")
        # 兼容历史入口：若其他位置仍设置 indicators_eval_file，则自动转入“评估结果”
        if st.session_state.get('indicators_eval_file') and not st.session_state.get('tab4_eval_target_file'):
            st.session_state['tab4_eval_target_file'] = st.session_state.get('indicators_eval_file')
            st.session_state.pop('indicators_eval_file', None)

        def _load_file_detail(file_id):
            rows = execute_query('''
                SELECT pf.*, p.name as project_name,
                       COALESCE(pf.category, p.category) as eff_category,
                       COALESCE(pf.subcategory, p.subcategory) as eff_subcategory
                FROM project_files pf
                LEFT JOIN projects p ON pf.project_id = p.id
                WHERE pf.id = ?
            ''', (file_id,), fetch=True)
            return rows[0] if rows else None

        def _render_eval_row_actions(file_id, key_suffix):
            b1, b2 = st.columns([1, 1])
            with b1:
                if st.button("📝 开始评估", key=f"tab4_eval_{key_suffix}_{file_id}", use_container_width=True):
                    st.session_state['tab4_eval_target_file'] = file_id
                    st.session_state.pop('tab4_view_result_file', None)
                    st.rerun()
            with b2:
                if st.button("👁️ 查看结果", key=f"tab4_view_{key_suffix}_{file_id}", use_container_width=True):
                    st.session_state['tab4_view_result_file'] = file_id
                    st.session_state.pop('tab4_eval_target_file', None)
                    st.rerun()

        # 评估结果录入区（按钮“评估结果”触发）
        eval_target = st.session_state.get('tab4_eval_target_file')
        if eval_target:
            file_info = _load_file_detail(eval_target)
            if file_info:
                head1, head2 = st.columns([6, 1])
                with head1:
                    st.markdown(f"#### 📝 评估结果 - {file_info.get('file_name') or file_info.get('title')}")
                    st.caption(f"项目: {file_info.get('project_name') or '-'}")
                with head2:
                    if st.button("关闭", key=f"tab4_close_eval_{eval_target}"):
                        st.session_state.pop('tab4_eval_target_file', None)
                        st.rerun()

                file_cat = file_info.get('eff_category') or file_info.get('category')
                subcats = PROJECT_CATEGORIES.get(file_cat, {}).get('subcategories', {}) if file_cat else {}
                if file_cat and subcats:
                    sub_select_eval = st.selectbox(
                        "选择二级分类",
                        list(subcats.keys()),
                        format_func=lambda x: f"{x} - {subcats[x]}",
                        key=f"tab4_sub_eval_{eval_target}"
                    )
                    indicators = execute_query(
                        'SELECT * FROM indicator_library WHERE category = ? AND subcategory = ?',
                        (file_cat, sub_select_eval),
                        fetch=True
                    )
                elif file_cat:
                    indicators = execute_query(
                        'SELECT * FROM indicator_library WHERE category = ?',
                        (file_cat,),
                        fetch=True
                    )
                else:
                    indicators = []

                if indicators:
                    st.markdown("##### 评估打分")
                    total_score = 0.0
                    total_weight = sum(float(ind.get('weight') or 0) for ind in indicators)
                    for ind in indicators:
                        c1, c2, c3 = st.columns([3, 1, 1])
                        with c1:
                            st.write(f"**{ind['indicator_name']}** (权重: {ind['weight']})")
                        with c2:
                            score_key = f"tab4_score_{eval_target}_{ind['id']}"
                            score = st.slider("得分", 0, 100, 80, key=score_key)
                        with c3:
                            weighted_score = score * float(ind.get('weight') or 0) / 100 if total_weight > 0 else 0
                            total_score += weighted_score
                            st.write(f"加权: {weighted_score:.1f}")

                    total_weight_display = int(total_weight) if float(total_weight).is_integer() else round(total_weight, 1)
                    st.markdown(f"**总分: {total_score:.1f} / {total_weight_display}**")
                    if st.button("提交评估结果", key=f"tab4_submit_eval_{eval_target}"):
                        for ind in indicators:
                            score = st.session_state.get(f"tab4_score_{eval_target}_{ind['id']}", 80)
                            execute_query('''
                                INSERT INTO file_evaluations (file_id, indicator_id, score, evaluated_by)
                                VALUES (?, ?, ?, ?)
                            ''', (eval_target, ind['id'], score, st.session_state['user']['id']), commit=True)
                        st.session_state.pop('tab4_eval_target_file', None)
                        notify_and_rerun("评估结果已提交", level='success')
                else:
                    st.info("该文件所属分类暂无可用指标")
            else:
                st.warning("目标文件不存在或已删除")

        # 查看结果区（按钮“查看结果”触发）
        view_target = st.session_state.get('tab4_view_result_file')
        if view_target:
            file_info = _load_file_detail(view_target)
            if file_info:
                head1, head2 = st.columns([6, 1])
                with head1:
                    st.markdown(f"#### 👁️ 查看结果 - {file_info.get('file_name') or file_info.get('title')}")
                    st.caption(f"项目: {file_info.get('project_name') or '-'}")
                with head2:
                    if st.button("关闭", key=f"tab4_close_view_{view_target}"):
                        st.session_state.pop('tab4_view_result_file', None)
                        st.rerun()

                eval_rows = execute_query('''
                    SELECT fe.id, fe.file_id, fe.indicator_id, fe.score, fe.comment, fe.evaluated_at,
                           datetime(fe.evaluated_at, 'localtime') as evaluated_at_local,
                           il.indicator_name, il.weight,
                           COALESCE(u.real_name, u.username) as evaluator
                    FROM file_evaluations fe
                    JOIN indicator_library il ON fe.indicator_id = il.id
                    LEFT JOIN users u ON fe.evaluated_by = u.id
                    WHERE fe.file_id = ?
                    ORDER BY fe.evaluated_at DESC, fe.id DESC
                ''', (view_target,), fetch=True)

                if eval_rows:
                    latest_by_indicator = {}
                    for row in eval_rows:
                        ind_id = row.get('indicator_id')
                        if ind_id not in latest_by_indicator:
                            latest_by_indicator[ind_id] = row
                    latest_rows = list(latest_by_indicator.values())
                    total_weight = sum(float(r.get('weight') or 0) for r in latest_rows)
                    total_score = (
                        sum(float(r.get('score') or 0) * float(r.get('weight') or 0) / 100 for r in latest_rows)
                        if total_weight > 0 else 0
                    )
                    total_weight_display = int(total_weight) if float(total_weight).is_integer() else round(total_weight, 1)
                    st.metric("最新总分", f"{total_score:.1f} / {total_weight_display}")

                    rows_for_table = [{
                        '指标': r.get('indicator_name') or '-',
                        '权重': r.get('weight') or 0,
                        '得分': round(float(r.get('score') or 0) * float(r.get('weight') or 0) / 100, 1),
                        '评估人': r.get('evaluator') or '-',
                        '评估时间': format_datetime_display(r.get('evaluated_at_local') or r.get('evaluated_at'))
                    } for r in latest_rows]
                    summary_evaluator = latest_rows[0].get('evaluator') if latest_rows else '-'
                    summary_time = format_datetime_display(
                        latest_rows[0].get('evaluated_at_local') or latest_rows[0].get('evaluated_at')
                    ) if latest_rows else '-'
                    # 末行汇总
                    rows_for_table.append({
                        '指标': '汇总',
                        '权重': total_weight_display,
                        '得分': round(total_score, 1),
                        '评估人': summary_evaluator or '-',
                        '评估时间': summary_time or '-'
                    })
                    df = pd.DataFrame(rows_for_table)
                    st.dataframe(df, use_container_width=True, hide_index=True)
                else:
                    st.info("该文件暂无有效评估结果（若指标已删除，历史记录会自动过滤）")
            else:
                st.warning("目标文件不存在或已删除")

        st.markdown("---")

        # 搜索（支持点击按钮或按 Enter）
        if 'eval_file_search_query' not in st.session_state:
            st.session_state['eval_file_search_query'] = ''
        if 'eval_file_search_input' not in st.session_state:
            st.session_state['eval_file_search_input'] = st.session_state.get('eval_file_search_query', '')
        if 'eval_file_search_empty' not in st.session_state:
            st.session_state['eval_file_search_empty'] = False

        eval_placeholder = "请输入内容" if (
            st.session_state.get('eval_file_search_empty')
            and not str(st.session_state.get('eval_file_search_input', '')).strip()
        ) else "输入项目或文件关键词查询"

        with st.form("tab4_eval_search_form", clear_on_submit=False):
            s1, s2 = st.columns([5, 1])
            with s1:
                eval_search_input = st.text_input(
                    "搜索项目或文件",
                    key="eval_file_search_input",
                    placeholder=eval_placeholder,
                    label_visibility="collapsed"
                )
            with s2:
                eval_search_submit = st.form_submit_button("🔍 搜索", use_container_width=True)

        if str(st.session_state.get('eval_file_search_input', '')).strip():
            st.session_state['eval_file_search_empty'] = False

        if eval_search_submit:
            query = (eval_search_input or '').strip()
            if query:
                st.session_state['eval_file_search_query'] = query
                st.session_state['eval_file_search_empty'] = False
            else:
                st.session_state['eval_file_search_query'] = ''
                st.session_state['eval_file_search_empty'] = True

        search = (st.session_state.get('eval_file_search_query') or '').strip()

        if search:
            st.markdown(f"### 🔍 搜索结果: {search}")
            files = execute_query('''
                SELECT pf.*, p.name as project_name, u.username as uploader,
                       COALESCE(pf.category, p.category) as eff_category,
                       COALESCE(pf.subcategory, p.subcategory) as eff_subcategory
                FROM project_files pf
                LEFT JOIN projects p ON pf.project_id = p.id
                LEFT JOIN users u ON pf.upload_by = u.id
                WHERE pf.approval_status = 'approved'
                  AND (pf.file_name LIKE ? OR pf.title LIKE ? OR p.name LIKE ?)
                ORDER BY pf.upload_at DESC
            ''', (f'%{search}%', f'%{search}%', f'%{search}%'), fetch=True)

            if files:
                projects_in_search = {}
                for f in files:
                    p_id = f.get('project_id') or f"unlinked_{f['id']}"
                    if p_id not in projects_in_search:
                        p_name = f.get('project_name') or '未关联项目'
                        cat = f.get('eff_category') or '0'
                        sub = f.get('eff_subcategory')
                        prefix = f"{cat}-{sub}" if sub else f"{cat}"
                        projects_in_search[p_id] = {'display_name': f"{prefix} {p_name}", 'files': []}
                    projects_in_search[p_id]['files'].append(f)

                for p_id, p_info in projects_in_search.items():
                    with st.expander(f"📁 {p_info['display_name']}", expanded=True):
                        for f in p_info['files']:
                            cat = f.get('eff_category') or '0'
                            sub = f.get('eff_subcategory')
                            file_prefix = f"{cat}-{sub}" if sub else f"{cat}"
                            fname = f.get('file_name') or ''
                            display_name = fname if fname.startswith(f"{file_prefix}_") else f"{file_prefix}_{fname}"
                            c1, c2, c3 = st.columns([4, 1, 2])
                            with c1:
                                st.write(f"📄 {display_name}")
                            with c2:
                                st.write(f"{f['upload_at'][:10]}")
                            with c3:
                                _render_eval_row_actions(f['id'], "search")
            else:
                st.info("未找到匹配的文件")
        else:
            view_mode = st.radio("视图模式", ["项目分类文件夹", "项目文件夹"], horizontal=True, key="eval_file_view_mode")

            if view_mode == "项目文件夹":
                st.markdown("### 📁 项目文件夹")
                all_approved_files = execute_query('''
                    SELECT pf.*, p.name as project_name, u.username as uploader,
                           COALESCE(pf.category, p.category) as eff_category,
                           COALESCE(pf.subcategory, p.subcategory) as eff_subcategory
                    FROM project_files pf
                    LEFT JOIN projects p ON pf.project_id = p.id
                    LEFT JOIN users u ON pf.upload_by = u.id
                    WHERE pf.approval_status = 'approved'
                    ORDER BY pf.upload_at DESC
                ''', fetch=True)

                if not all_approved_files:
                    st.info("暂无已通过文件的项目")
                else:
                    projects_dict = {}
                    for f in all_approved_files:
                        p_id = f.get('project_id') or f"unlinked_{f['id']}"
                        if p_id not in projects_dict:
                            p_name = f.get('project_name') or '未关联项目'
                            cat = f.get('eff_category') or '0'
                            sub = f.get('eff_subcategory')
                            prefix = f"{cat}-{sub}" if sub else f"{cat}"
                            projects_dict[p_id] = {'display_name': f"{prefix} {p_name}", 'files': []}
                        projects_dict[p_id]['files'].append(f)

                    for p_id, p_info in projects_dict.items():
                        open_flag = (st.session_state.get('indicators_open_project') == p_id)
                        with st.expander(f"📁 {p_info['display_name']}", expanded=open_flag):
                            for f in p_info['files']:
                                cat = f.get('eff_category') or '0'
                                sub = f.get('eff_subcategory')
                                file_prefix = f"{cat}-{sub}" if sub else f"{cat}"
                                fname = f.get('file_name') or ''
                                display_name = fname if fname.startswith(f"{file_prefix}_") else f"{file_prefix}_{fname}"
                                c1, c2, c3 = st.columns([4, 1, 2])
                                with c1:
                                    st.write(f"📄 {display_name}")
                                with c2:
                                    st.write(f"{f['upload_at'][:10]}")
                                with c3:
                                    _render_eval_row_actions(f['id'], "proj")
            else:
                st.markdown("### 📁 项目分类文件夹")
                folder_definitions = [
                    {'cat': '1', 'sub': '1', 'label': '1-1 多源数据采集'},
                    {'cat': '1', 'sub': '2', 'label': '1-2 专项调查研究'},
                    {'cat': '2', 'sub': '1', 'label': '2-1 财政绩效评估'},
                    {'cat': '2', 'sub': '2', 'label': '2-2 行政绩效评估'},
                    {'cat': '3', 'sub': '1', 'label': '3-1 企业管理咨询'},
                    {'cat': '3', 'sub': '2', 'label': '3-2 公共决策咨询'},
                    {'cat': '0', 'sub': None, 'label': '0 其他项目'}
                ]

                all_approved_files = execute_query('''
                    SELECT pf.*, p.name as project_name, u.username as uploader,
                           COALESCE(pf.category, p.category) as eff_category,
                           COALESCE(pf.subcategory, p.subcategory) as eff_subcategory
                    FROM project_files pf
                    LEFT JOIN projects p ON pf.project_id = p.id
                    LEFT JOIN users u ON pf.upload_by = u.id
                    WHERE pf.approval_status = 'approved'
                    ORDER BY pf.upload_at DESC
                ''', fetch=True)

                if not all_approved_files:
                    st.info("暂无已通过分类的文件")
                else:
                    for f_def in folder_definitions:
                        cat = f_def['cat']
                        sub = f_def['sub']
                        if sub is not None:
                            folder_files = [f for f in all_approved_files if str(f.get('eff_category')) == str(cat) and str(f.get('eff_subcategory')) == str(sub)]
                        else:
                            folder_files = [f for f in all_approved_files if str(f.get('eff_category')) == str(cat) and not f.get('eff_subcategory')]
                        if not folder_files:
                            continue

                        target_proj_id = st.session_state.get('indicators_open_project')
                        has_open_proj = any((f.get('project_id') == target_proj_id) for f in folder_files)
                        cat_state_key = f"tab4_cat_open_{cat}_{sub}"
                        if has_open_proj and target_proj_id:
                            st.session_state[cat_state_key] = True

                        is_open = st.session_state.get(cat_state_key, False)
                        if st.button(f"{'📂' if is_open else '📁'} {f_def['label']}", key=f"tab4_btn_cat_{cat}_{sub}", use_container_width=True):
                            st.session_state[cat_state_key] = not is_open
                            st.rerun()

                        if st.session_state.get(cat_state_key, False):
                            projects_in_folder = {}
                            for f in folder_files:
                                p_id = f.get('project_id') or f"unlinked_{f['id']}"
                                if p_id not in projects_in_folder:
                                    p_name = f.get('project_name') or '未关联项目'
                                    prefix = f"{cat}-{sub}" if sub else f"{cat}"
                                    projects_in_folder[p_id] = {'display_name': f"{prefix} {p_name}", 'files': []}
                                projects_in_folder[p_id]['files'].append(f)

                            for p_id, p_info in projects_in_folder.items():
                                open_flag = (st.session_state.get('indicators_open_project') == p_id)
                                with st.expander(f"📁 {p_info['display_name']}", expanded=open_flag):
                                    for f in p_info['files']:
                                        file_prefix = f"{cat}-{sub}" if sub else f"{cat}"
                                        fname = f.get('file_name') or ''
                                        display_name = fname if fname.startswith(f"{file_prefix}_") else f"{file_prefix}_{fname}"
                                        c1, c2, c3 = st.columns([4, 1, 2])
                                        with c1:
                                            st.write(f"📄 {display_name}")
                                        with c2:
                                            st.write(f"{f['upload_at'][:10]}")
                                        with c3:
                                            _render_eval_row_actions(f['id'], "cat")

def render_admin_visualization():
    """管理端可视化大屏"""
    st.markdown(
        """
        <div class="viz-hero">
            <p class="viz-hero-title">📈 可视化大屏</p>
            <p class="viz-hero-subtitle">
                聚焦“已审批文件”全量数据，展示规模、评估进度、质量分布与近30天趋势，辅助管理端快速发现风险与亮点。
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    def style_viz_figure(fig, height=360):
        fig.update_layout(
            template="plotly_dark",
            height=height,
            margin=dict(l=24, r=20, t=56, b=38),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(4, 53, 70, 0.58)",
            font=dict(color="#EAFBFF", size=13, family="Microsoft YaHei, Segoe UI, sans-serif"),
            legend=dict(
                bgcolor="rgba(0,0,0,0)",
                orientation="h",
                x=0,
                y=1.12,
                xanchor="left",
                yanchor="bottom",
                font=dict(color="#EAFBFF", size=12)
            ),
            hoverlabel=dict(
                bgcolor="rgba(4, 36, 57, 0.95)",
                font=dict(color="#ffffff")
            ),
            title=dict(font=dict(size=17, color="#F1FDFF"))
        )
        fig.update_xaxes(
            showgrid=True,
            gridcolor="rgba(255,255,255,0.12)",
            zeroline=False,
            title_font=dict(color="#BFEFFF"),
            tickfont=dict(color="#D9F3FF"),
            automargin=True
        )
        fig.update_yaxes(
            showgrid=True,
            gridcolor="rgba(255,255,255,0.12)",
            zeroline=False,
            title_font=dict(color="#BFEFFF"),
            tickfont=dict(color="#D9F3FF"),
            automargin=True
        )
        return fig

    def render_viz_panel(title, subtitle):
        st.markdown(
            f"""
            <div class="viz-panel">
                <p class="viz-panel-title">{title}</p>
                <p class="viz-panel-subtitle">{subtitle}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    def plot_viz_chart(fig, height=360, move_legend_right=False):
        styled_fig = style_viz_figure(fig, height=height)
        if move_legend_right:
            styled_fig.update_layout(
                legend=dict(
                    orientation='v',
                    x=1.08,
                    y=1.0,
                    xanchor='left',
                    yanchor='top',
                    bgcolor='rgba(0,0,0,0)',
                    font=dict(color='#EAFBFF', size=12)
                ),
                margin=dict(l=24, r=150, t=56, b=38)
            )
        st.plotly_chart(
            styled_fig,
            use_container_width=True,
            config={
                "displayModeBar": False,
                "displaylogo": False,
                "responsive": True
            }
        )

    CHART_HEIGHT = 360

    total_files = execute_query(
        "SELECT COUNT(*) as cnt FROM project_files WHERE approval_status = 'approved'",
        fetch=True
    )[0]['cnt']
    evaluated_files = execute_query(
        '''
        SELECT COUNT(DISTINCT fe.file_id) as cnt
        FROM file_evaluations fe
        INNER JOIN project_files pf ON pf.id = fe.file_id
        WHERE pf.approval_status = 'approved'
        ''',
        fetch=True
    )[0]['cnt']
    avg_score = execute_query(
        '''
        SELECT AVG(fe.score) as avg
        FROM file_evaluations fe
        INNER JOIN project_files pf ON pf.id = fe.file_id
        WHERE pf.approval_status = 'approved'
        ''',
        fetch=True
    )[0]['avg'] or 0

    unevaluated = max(total_files - evaluated_files, 0)
    coverage = (evaluated_files / total_files * 100) if total_files else 0

    kpi_items = [
        {"label": "总文件数", "value": total_files, "extra": "已审批文件总量"},
        {"label": "已评估文件", "value": evaluated_files, "extra": f"覆盖率 {coverage:.1f}%"},
        {"label": "未评估文件", "value": unevaluated, "extra": "建议优先推进评估"},
        {"label": "平均得分", "value": f"{avg_score:.1f}", "extra": "基于全部评估明细"}
    ]

    kpi_cols = st.columns(4)
    for col, item in zip(kpi_cols, kpi_items):
        with col:
            st.markdown(
                f"""
                <div class="viz-kpi-card">
                    <p class="viz-kpi-label">{item['label']}</p>
                    <p class="viz-kpi-value">{item['value']}</p>
                    <div class="viz-kpi-extra">{item['extra']}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown('<div class="viz-section-title">文件规模与评估状态</div>', unsafe_allow_html=True)
    row1_col1, row1_col2 = st.columns(2)

    with row1_col1:
        project_file_rows = execute_query(
            '''
            SELECT
                COALESCE(NULLIF(TRIM(p.name), ''), '未命名项目') AS project_name,
                COUNT(pf.id) AS file_cnt
            FROM project_files pf
            JOIN projects p ON p.id = pf.project_id
            GROUP BY p.id, p.name
            ORDER BY file_cnt DESC, p.id DESC
            ''',
            fetch=True
        )

        project_file_df = pd.DataFrame(project_file_rows) if project_file_rows else pd.DataFrame(
            columns=['project_name', 'file_cnt']
        )
        submitted_file_total = int(project_file_df['file_cnt'].sum()) if not project_file_df.empty else 0

        render_viz_panel(
            "项目文件提交分布",
            f"按项目统计已提交文件量，总计 {submitted_file_total} 份"
        )

        if not project_file_df.empty:
            project_file_df['file_cnt'] = project_file_df['file_cnt'].astype(int)
            project_file_df['share_pct'] = (
                project_file_df['file_cnt'] / max(submitted_file_total, 1) * 100
            )
            project_file_df['label_text'] = project_file_df.apply(
                lambda row: f"{int(row['file_cnt'])} 份 · {row['share_pct']:.1f}%",
                axis=1
            )

            top_n = min(8, len(project_file_df))
            top_df = project_file_df.head(top_n).sort_values('file_cnt', ascending=True)
            x_max = int(top_df['file_cnt'].max()) if not top_df.empty else 1
            x_max = x_max + max(1, int(x_max * 0.25))

            fig = go.Figure(
                go.Bar(
                    y=top_df['project_name'],
                    x=top_df['file_cnt'],
                    orientation='h',
                    name='文件数量',
                    marker=dict(
                        color=top_df['file_cnt'],
                        colorscale=[[0, '#23d1bd'], [1, '#2c9fff']],
                        line=dict(color='rgba(255,255,255,0.12)', width=1)
                    ),
                    text=top_df['label_text'],
                    textposition='outside',
                    customdata=top_df[['share_pct']].values,
                    hovertemplate=(
                        "项目: %{y}<br>"
                        "文件数量: %{x}<br>"
                        "占比: %{customdata[0]:.1f}%<extra></extra>"
                    )
                )
            )
            fig.update_layout(
                title=f"各项目文件量 TOP{top_n}",
                showlegend=False,
                xaxis_title="文件数量",
                yaxis_title=""
            )
            fig.update_xaxes(range=[0, x_max], tickmode='linear', dtick=max(1, x_max // 6))
            plot_viz_chart(fig, height=CHART_HEIGHT)

        else:
            st.info("暂无项目文件提交数据")

    with row1_col2:
        render_viz_panel("评估状态分布", "已评估与未评估文件占比，一眼识别评估覆盖短板")
        if total_files > 0:
            status_df = pd.DataFrame(
                {'状态': ['已评估', '未评估'], '数量': [evaluated_files, unevaluated]}
            )
            fig = px.pie(
                status_df,
                values='数量',
                names='状态',
                hole=0.55,
                color='状态',
                color_discrete_map={'已评估': '#15c39a', '未评估': '#ff8c3a'}
            )
            fig.update_traces(textinfo='label+percent', textfont_size=13, pull=[0.02, 0])
            fig.update_layout(title="评估状态占比")
            plot_viz_chart(fig, height=CHART_HEIGHT, move_legend_right=True)
        else:
            st.info("暂无文件，暂无法生成状态分布")

    st.markdown('<div class="viz-section-title">质量与趋势分析</div>', unsafe_allow_html=True)
    row2_col1, row2_col2 = st.columns(2)

    with row2_col1:
        render_viz_panel("近30天评估趋势", "柱状显示每日评估次数，折线展示日均分与7日均线，观察波动更直观")
        trend_rows = execute_query(
            '''
            SELECT DATE(fe.evaluated_at, 'localtime') as eval_date, COUNT(*) as cnt, AVG(fe.score) as avg_score
            FROM file_evaluations fe
            INNER JOIN project_files pf ON pf.id = fe.file_id
            WHERE pf.approval_status = 'approved'
              AND DATE(fe.evaluated_at, 'localtime') >= DATE('now', 'localtime', '-29 day')
            GROUP BY DATE(fe.evaluated_at, 'localtime')
            ORDER BY eval_date
            ''',
            fetch=True
        )

        full_dates = pd.date_range(end=datetime.now().date(), periods=30, freq='D')
        full_df = pd.DataFrame({'eval_date': pd.to_datetime(full_dates)})

        if trend_rows:
            raw_trend_df = pd.DataFrame(trend_rows)
            raw_trend_df['eval_date'] = pd.to_datetime(raw_trend_df['eval_date'], errors='coerce')
            raw_trend_df['cnt'] = pd.to_numeric(raw_trend_df['cnt'], errors='coerce')
            raw_trend_df['avg_score'] = pd.to_numeric(raw_trend_df['avg_score'], errors='coerce')
            raw_trend_df = raw_trend_df.dropna(subset=['eval_date'])
            trend_df = full_df.merge(raw_trend_df[['eval_date', 'cnt', 'avg_score']], on='eval_date', how='left')
        else:
            trend_df = full_df.copy()
            trend_df['cnt'] = 0
            trend_df['avg_score'] = None

        trend_df['cnt'] = trend_df['cnt'].fillna(0).astype(int)
        trend_df['avg_score'] = pd.to_numeric(trend_df['avg_score'], errors='coerce')

        # 仅基于“有评估记录”的日期计算均线，避免无数据日期干扰趋势
        score_active_mask = trend_df['avg_score'].notna()
        active_scores = trend_df.loc[score_active_mask, 'avg_score']
        trend_df['avg_score_7d'] = None
        active_score_days = int(active_scores.shape[0])
        if not active_scores.empty:
            # 小样本阶段也展示滚动均值，避免图例存在但曲线不可见
            rolling_mean = active_scores.rolling(window=7, min_periods=1).mean()
            trend_df.loc[score_active_mask, 'avg_score_7d'] = rolling_mean.values

        score_point_df = trend_df.dropna(subset=['avg_score']).copy()
        score_trend_df = trend_df.dropna(subset=['avg_score_7d']).copy()
        trend_mode = 'lines+markers' if len(score_trend_df) <= 1 else 'lines'

        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=trend_df['eval_date'],
                y=trend_df['cnt'],
                name='日评估次数',
                marker_color='rgba(65, 192, 255, 0.58)',
                hovertemplate="日期: %{x|%m-%d}<br>评估次数: %{y}<extra></extra>"
            )
        )
        fig.add_trace(
            go.Scatter(
                x=score_point_df['eval_date'],
                y=score_point_df['avg_score'],
                name='日均分',
                mode='lines+markers',
                line=dict(color='#ffd166', width=3),
                marker=dict(size=6, color='#ffd166'),
                yaxis='y2',
                connectgaps=False,
                hovertemplate="日期: %{x|%m-%d}<br>日均分: %{y:.1f}<extra></extra>"
            )
        )

        if not score_trend_df.empty:
            fig.add_trace(
                go.Scatter(
                    x=score_trend_df['eval_date'],
                    y=score_trend_df['avg_score_7d'],
                    name='7日均线',
                    mode=trend_mode,
                    line=dict(color='#ff6b6b', width=2.4, dash='dot'),
                    marker=dict(size=7, color='#ff6b6b'),
                    yaxis='y2',
                    connectgaps=False,
                    hovertemplate="日期: %{x|%m-%d}<br>7日均分: %{y:.1f}<extra></extra>"
                )
            )
        fig.update_layout(
            title="近30天评估趋势（次数 + 得分）",
            xaxis_title="日期",
            yaxis=dict(title="评估次数", rangemode='tozero'),
            yaxis2=dict(
                title="平均得分",
                overlaying='y',
                side='right',
                range=[0, 100],
                showgrid=False
            ),
            hovermode='x unified'
        )
        fig.update_xaxes(tickformat="%m-%d", nticks=10)
        plot_viz_chart(fig, height=CHART_HEIGHT, move_legend_right=True)

        if int(trend_df['cnt'].sum()) == 0:
            st.info("近30天暂无评估记录")
        elif active_score_days < 7:
            st.caption(f"提示：当前仅有 {active_score_days} 个评估日，7日均线按现有评估日滚动计算。")

    with row2_col2:
        render_viz_panel("文件得分区间分布", "按分数段统计文件数量，快速识别质量集中区间")
        score_band_rows = execute_query(
            '''
            WITH file_score AS (
                SELECT pf.id AS file_id, AVG(fe.score) AS file_avg_score
                FROM project_files pf
                LEFT JOIN file_evaluations fe ON pf.id = fe.file_id
                WHERE pf.approval_status = 'approved'
                GROUP BY pf.id
            )
            SELECT file_avg_score
            FROM file_score
            WHERE file_avg_score IS NOT NULL
            ''',
            fetch=True
        )

        if score_band_rows:
            score_df = pd.DataFrame(score_band_rows)
            score_df['score_band'] = pd.cut(
                score_df['file_avg_score'],
                bins=[0, 60, 70, 80, 90, 101],
                labels=['0-59', '60-69', '70-79', '80-89', '90-100'],
                right=False,
                include_lowest=True
            )
            dist_df = score_df.groupby('score_band').size().reset_index(name='cnt')
            dist_df['cnt'] = pd.to_numeric(dist_df['cnt'], errors='coerce').fillna(0).astype(int)

            max_cnt = int(dist_df['cnt'].max()) if not dist_df.empty else 0
            target_ticks = 6
            raw_dtick = max(1, (max_cnt + target_ticks - 1) // target_ticks)  # 整数步长，避免小数刻度
            if raw_dtick <= 5:
                y_dtick = raw_dtick
            else:
                magnitude = 10 ** (len(str(raw_dtick)) - 1)
                y_dtick = ((raw_dtick + magnitude - 1) // magnitude) * magnitude
            y_max = max(1, ((max_cnt + y_dtick - 1) // y_dtick) * y_dtick + y_dtick)

            fig = px.bar(
                dist_df,
                x='score_band',
                y='cnt',
                text='cnt',
                color='score_band',
                color_discrete_sequence=['#ef476f', '#ff7f50', '#ffd166', '#06d6a0', '#118ab2']
            )
            fig.update_layout(
                title="文件得分区间分布",
                xaxis_title="得分区间",
                yaxis=dict(
                    title="文件数量",
                    range=[0, y_max],
                    tickmode='linear',
                    tick0=0,
                    dtick=y_dtick
                ),
                showlegend=False
            )
            fig.update_traces(textposition='outside')
            plot_viz_chart(fig, height=CHART_HEIGHT)
        else:
            st.info("暂无得分数据")

# ==================== 机构端页面 ====================
def render_org_dashboard():
    """机构端工作台"""
    user = st.session_state['user']
    org_id = user['org_id']

    def safe_text(value, default="-"):
        if value is None:
            return default
        text = str(value).strip()
        return html.escape(text if text else default)

    org = execute_query("SELECT * FROM organizations WHERE id = ?", (org_id,), fetch=True)
    org_name = org[0]['name'] if org else "未知机构"
    display_name = user.get('real_name') or user['username']
    role_label = ROLE_NAMES.get(user['role'], '机构端')

    org_dashboard_snapshot = get_org_dashboard_snapshot(org_id, user['id'])
    project_count = org_dashboard_snapshot['project_count']
    active_count = org_dashboard_snapshot['active_count']
    completed_count = org_dashboard_snapshot['completed_count']
    file_count = org_dashboard_snapshot['file_count']

    todos = execute_query('''
        SELECT * FROM todos WHERE user_id = ? AND status = 'pending' ORDER BY created_at DESC LIMIT 5
    ''', (user['id'],), fetch=True)

    messages = execute_query('''
        SELECT * FROM messages WHERE user_id = ? ORDER BY created_at DESC LIMIT 5
    ''', (user['id'],), fetch=True)

    recent_projects = execute_query('''
        SELECT * FROM projects WHERE org_id = ? ORDER BY created_at DESC LIMIT 5
    ''', (org_id,), fetch=True)

    priority_map = {'high': '高优先级', 'medium': '中优先级', 'low': '低优先级'}
    todo_items = []
    for todo in todos or []:
        due_date = todo['due_date'] if 'due_date' in todo.keys() else None
        meta_parts = [priority_map.get(todo['priority'], '普通事项')]
        if due_date:
            meta_parts.append(f"截止 {format_datetime_display(due_date).split(' ')[0]}")
        content = safe_text(todo['content'], "") if todo['content'] else ""
        content_html = f"<div class='org-list-meta'>{content}</div>" if content else ""
        todo_items.append(textwrap.dedent(f"""
            <div class="org-list-item" data-org-nav="todos" role="button" tabindex="0">
                <div class="org-list-icon">✓</div>
                <div>
                    <div class="org-list-title">{safe_text(todo['title'])}</div>
                    {content_html}
                    <div class="org-list-meta">{safe_text(' / '.join(meta_parts))}</div>
                </div>
            </div>
        """).strip())
    todo_html = "".join(todo_items) if todo_items else "<div class='org-empty'>当前没有待办事项</div>"

    message_items = []
    unread_count = 0
    for msg in messages or []:
        is_unread = not bool(msg['is_read'])
        if is_unread:
            unread_count += 1
        icon = "🔔" if is_unread else "✓"
        created_date = format_datetime_display(msg['created_at']).split(' ')[0]
        message_items.append(textwrap.dedent(f"""
            <div class="org-list-item" data-org-nav="messages" role="button" tabindex="0">
                <div class="org-list-icon">{icon}</div>
                <div>
                    <div class="org-list-title">{safe_text(msg['title'])}</div>
                    <div class="org-list-meta">{safe_text(created_date)}</div>
                </div>
            </div>
        """).strip())
    message_html = "".join(message_items) if message_items else "<div class='org-empty'>当前没有消息通知</div>"

    status_map = {'pending': '待审核', 'in_progress': '进行中', 'completed': '已完成', 'rejected': '已驳回'}
    status_class_map = {
        'completed': 'org-status-completed',
        'rejected': 'org-status-rejected',
    }
    project_rows = []
    for project in recent_projects or []:
        status = project['status']
        status_class = status_class_map.get(status, '')
        created_at = format_datetime_display(project['created_at'])
        project_rows.append(textwrap.dedent(f"""
            <div class="org-project-row" data-org-nav="projects" role="button" tabindex="0">
                <div title="{safe_text(project['name'])}">{safe_text(project['name'])}</div>
                <div title="{safe_text(format_gate(project['current_stage']))}">{safe_text(format_gate(project['current_stage']))}</div>
                <div><span class="org-status-pill {status_class}">{safe_text(status_map.get(status, status))}</span></div>
                <div title="{safe_text(created_at)}">{safe_text(created_at)}</div>
            </div>
        """).strip())

    if project_rows:
        project_html = textwrap.dedent(f"""
            <div class="org-project-table">
                <div class="org-project-row org-project-head">
                    <div>项目名称</div>
                    <div>当前Gate</div>
                    <div>状态</div>
                    <div>创建时间</div>
                </div>
                {''.join(project_rows)}
            </div>
        """).strip()
    else:
        project_html = "<div class='org-empty'>暂无项目，创建项目后会在这里展示最新进展</div>"

    hero_html = textwrap.dedent(f"""
    <div class="org-workbench">
        <section class="org-hero">
            <div class="org-hero-grid">
                <div>
                    <div class="org-kicker">{safe_text(role_label)}工作台</div>
                    <h1>欢迎，{safe_text(display_name)}</h1>
                    <div class="org-hero-tags">
                        <span class="org-hero-tag">策链协同</span>
                        <span class="org-hero-tag">评估交付</span>
                        <span class="org-hero-tag">Gate 管控</span>
                        <span class="org-hero-tag">资料归集</span>
                    </div>
                </div>
                <div class="org-hero-panel">
                    <div class="org-hero-panel-title">当前机构</div>
                    <div class="org-hero-panel-value">{safe_text(org_name)}</div>
                    <div class="org-list-meta" style="color: rgba(234,251,255,0.72); margin-top: 10px;">
                        账号：{safe_text(user['username'])} · 待办 {len(todos or [])} 项 · 未读消息 {unread_count} 条
                    </div>
                </div>
            </div>
        </section>

        <section class="org-dashboard-grid">
            <div class="org-click-card" data-org-nav="projects" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">📁</div></div>
                    <div class="org-metric-number">{project_count}</div>
                    <div class="org-metric-label">项目总数</div>
                </div>
            </div>
            <div class="org-click-card" data-org-nav="projects" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">🚀</div></div>
                    <div class="org-metric-number">{active_count}</div>
                    <div class="org-metric-label">进行中项目</div>
                </div>
            </div>
            <div class="org-click-card" data-org-nav="projects" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">✅</div></div>
                    <div class="org-metric-number">{completed_count}</div>
                    <div class="org-metric-label">已完成项目</div>
                </div>
            </div>
            <div class="org-click-card" data-org-nav="knowledge" role="button" tabindex="0">
                <div class="org-metric-card">
                    <div class="org-metric-top"><div class="org-metric-icon">📎</div></div>
                    <div class="org-metric-number">{file_count}</div>
                    <div class="org-metric-label">上传文件数</div>
                </div>
            </div>
        </section>
    </div>
    """).strip()
    hero_html = "\n".join(line.strip() for line in hero_html.splitlines() if line.strip())
    st.markdown(hero_html, unsafe_allow_html=True)

    dashboard_html = textwrap.dedent(f"""
    <div class="org-workbench">
        <section class="org-content-grid">
            <div class="org-panel">
                <div class="org-panel-header">
                    <div class="org-panel-title">待办事项</div>
                    <div class="org-panel-badge" data-org-nav="todos" role="button" tabindex="0">{len(todos or [])} 项待处理</div>
                </div>
                <div class="org-list">{todo_html}</div>
            </div>
            <div class="org-panel">
                <div class="org-panel-header">
                    <div class="org-panel-title">最新消息</div>
                    <div class="org-panel-badge" data-org-nav="messages" role="button" tabindex="0">{unread_count} 条未读</div>
                </div>
                <div class="org-list">{message_html}</div>
            </div>
        </section>

        <section class="org-panel org-project-panel">
            <div class="org-panel-header">
                <div class="org-panel-title">最近项目</div>
                <div class="org-panel-badge" data-org-nav="projects" role="button" tabindex="0">最近 5 个项目</div>
            </div>
            {project_html}
        </section>
    </div>
    """).strip()
    dashboard_html = "\n".join(line.strip() for line in dashboard_html.splitlines() if line.strip())
    st.markdown(dashboard_html, unsafe_allow_html=True)

    nav_bridge_map = json.dumps({
        "projects": "项目管理",
        "knowledge": "项目智库",
        "todos": "待办事项",
        "messages": "消息通知",
    }, ensure_ascii=False)
    components.html(f"""
    <script>
    (function() {{
        const doc = window.parent.document;
        if (!doc) {{
            return;
        }}
        const navMap = {nav_bridge_map};

        function findSidebarButton(label) {{
            const sidebar = doc.querySelector('[data-testid="stSidebar"]') || doc;
            const buttons = Array.from(sidebar.querySelectorAll('button'));
            return buttons.find(function(button) {{
                return (button.innerText || button.textContent || '').trim().includes(label);
            }});
        }}

        function triggerDashboardNav(event) {{
            if (event.type === 'keydown' && event.key !== 'Enter' && event.key !== ' ') {{
                return;
            }}
            const target = event.target && event.target.closest ? event.target.closest('[data-org-nav]') : null;
            if (!target) {{
                return;
            }}
            const page = target.getAttribute('data-org-nav');
            const buttonLabel = navMap[page];
            const button = buttonLabel ? findSidebarButton(buttonLabel) : null;
            if (!button) {{
                return;
            }}
            event.preventDefault();
            event.stopPropagation();
            button.click();
        }}

        if (doc.__orgDashboardNavHandler) {{
            doc.removeEventListener('click', doc.__orgDashboardNavHandler, true);
            doc.removeEventListener('keydown', doc.__orgDashboardNavHandler, true);
        }}
        doc.__orgDashboardNavHandler = triggerDashboardNav;
        doc.addEventListener('click', triggerDashboardNav, true);
        doc.addEventListener('keydown', triggerDashboardNav, true);
    }})();
    </script>
    """, height=0)

def render_org_info():
    """机构端信息维护"""
    st.title("🏢 信息维护")
    
    user = st.session_state['user']
    org_id = user['org_id']
    
    tab1, tab2, tab3, tab4 = st.tabs(["机构信息", "主评人管理", "业绩记录", "培训记录"])
    
    with tab1:
        st.subheader("机构信息")
        
        org = execute_query("SELECT * FROM organizations WHERE id = ?", (org_id,), fetch=True)
        
        if org:
            org_data = org[0]
            
            with st.form("update_org_form"):
                col1, col2 = st.columns(2)
                
                with col1:
                    name = st.text_input("机构名称", value=org_data['name'])
                    org_type = st.selectbox("机构类型", ["企业", "事业单位", "社会团体", "民办非企业", "其他"],
                                           index=["企业", "事业单位", "社会团体", "民办非企业", "其他"].index(org_data['org_type']) if org_data['org_type'] else 0)
                    credit_code = st.text_input("统一社会信用代码", value=org_data['credit_code'] or "")
                    legal_person = st.text_input("法定代表人", value=org_data['legal_person'] or "")
                    contact_person = st.text_input("联系人", value=org_data['contact_person'] or "")
                
                with col2:
                    contact_phone = st.text_input("联系电话", value=org_data['contact_phone'] or "")
                    contact_email = st.text_input("联系邮箱", value=org_data['contact_email'] or "")
                    address = st.text_input("机构地址", value=org_data['address'] or "")
                    description = st.text_area("机构简介", value=org_data['description'] or "")
                
                if st.form_submit_button("更新信息", use_container_width=True):
                    old_data = {
                        'name': org_data['name'] or '',
                        'org_type': org_data['org_type'] or '',
                        'credit_code': org_data['credit_code'] or '',
                        'legal_person': org_data['legal_person'] or '',
                        'contact_person': org_data['contact_person'] or '',
                        'contact_phone': org_data['contact_phone'] or '',
                        'contact_email': org_data['contact_email'] or '',
                        'address': org_data['address'] or '',
                        'description': org_data['description'] or '',
                    }
                    new_data = {
                        'name': name.strip(),
                        'org_type': org_type,
                        'credit_code': credit_code.strip(),
                        'legal_person': legal_person.strip(),
                        'contact_person': contact_person.strip(),
                        'contact_phone': contact_phone.strip(),
                        'contact_email': contact_email.strip(),
                        'address': address.strip(),
                        'description': description.strip(),
                    }
                    
                    if old_data == new_data:
                        st.info("机构信息未发生变化")
                    elif user['role'] == 'org_user':
                        main_users = execute_query(
                            "SELECT id FROM users WHERE org_id = ? AND role = 'org_admin' AND status = 'active' ORDER BY id LIMIT 1",
                            (org_id,),
                            fetch=True
                        )
                        if not main_users:
                            st.error("当前机构未找到可审批的机构主账号")
                        else:
                            approver_id = main_users[0]['id']
                            pending_request = execute_query('''
                                SELECT id FROM org_info_update_requests
                                WHERE org_id = ? AND submitted_by = ? AND status = 'pending'
                                LIMIT 1
                            ''', (org_id, user['id']), fetch=True)
                            if pending_request:
                                st.warning("已有待审批的机构信息变更申请，请等待机构主账号处理后再提交")
                                return
                            request_id = execute_query('''
                                INSERT INTO org_info_update_requests (
                                    org_id, submitted_by, approver_id, status, old_data, new_data, submitted_at
                                )
                                VALUES (?, ?, ?, 'pending', ?, ?, ?)
                            ''', (
                                org_id,
                                user['id'],
                                approver_id,
                                json.dumps(old_data, ensure_ascii=False),
                                json.dumps(new_data, ensure_ascii=False),
                                datetime.now()
                            ), commit=True)
                            add_todo(
                                approver_id,
                                "机构信息变更审批",
                                f"{user.get('real_name') or user['username']} 提交了机构信息变更申请，请审批。",
                                priority='high',
                                related_type='org_info_update',
                                related_id=request_id
                            )
                            add_message(approver_id, '机构信息变更待审批', f'{user.get("real_name") or user["username"]} 提交了机构信息变更申请')
                            add_log(user['id'], user['username'], name, '提交机构信息变更申请', 'organizations',
                                    f'申请ID: {request_id}', get_client_ip())
                            notify_and_rerun("已提交机构主账号审核", level='success')
                    else:
                        execute_query('''
                            UPDATE organizations SET name = ?, org_type = ?, credit_code = ?, legal_person = ?,
                            contact_person = ?, contact_phone = ?, contact_email = ?, address = ?, description = ?, updated_at = ?
                            WHERE id = ?
                        ''', (
                            new_data['name'], new_data['org_type'], new_data['credit_code'], new_data['legal_person'],
                            new_data['contact_person'], new_data['contact_phone'], new_data['contact_email'],
                            new_data['address'], new_data['description'], datetime.now(), org_id
                        ), commit=True)
                        
                        add_log(user['id'], user['username'], name, '更新机构信息', 'organizations', '更新机构信息', get_client_ip())
                        notify_and_rerun("信息更新成功", level='success')
    
    with tab2:
        st.subheader("主评人管理")
        can_manage_evaluators = can_manage_org_evaluators(user['role'])
        
        if can_manage_evaluators:
            account_users = execute_query('''
                SELECT id, username, real_name, role, status
                FROM users
                WHERE org_id = ? AND role IN ('org_admin', 'org_user')
                ORDER BY CASE role WHEN 'org_admin' THEN 0 ELSE 1 END, created_at DESC, id DESC
            ''', (org_id,), fetch=True)
        else:
            account_users = execute_query('''
                SELECT id, username, real_name, role, status
                FROM users
                WHERE id = ? AND org_id = ?
            ''', (user['id'], org_id), fetch=True)
        account_by_id = {account['id']: account for account in (account_users or [])}
        account_options = [None] + [account['id'] for account in (account_users or [])]
        account_role_names = {"org_admin": "机构主账号", "org_user": "机构子账号"}
        account_status_names = {"active": "正常", "inactive": "冻结"}
        
        def format_account_option(account_id):
            if account_id is None:
                return "未关联账号"
            account = account_by_id.get(account_id)
            if not account:
                return "账号已删除"
            real_name = f" - {account['real_name']}" if account['real_name'] else ""
            role_name = account_role_names.get(account['role'], account['role'])
            status_name = account_status_names.get(account['status'], account['status'])
            return f"{account['username']}{real_name}（{role_name}/{status_name}）"
        
        if can_manage_evaluators:
            evaluators = execute_query('''
                SELECT e.*, u.username AS account_username, u.real_name AS account_real_name,
                       u.role AS account_role, u.status AS account_status
                FROM evaluators e
                LEFT JOIN users u ON e.account_user_id = u.id AND u.org_id = e.org_id
                WHERE e.org_id = ?
                ORDER BY CASE e.status WHEN 'active' THEN 0 ELSE 1 END, e.created_at DESC, e.id DESC
            ''', (org_id,), fetch=True)
        else:
            st.info("机构子账号仅可查看与当前账号关联的主评人，不能维护主评人或调整账号关联。")
            evaluators = execute_query('''
                SELECT e.*, u.username AS account_username, u.real_name AS account_real_name,
                       u.role AS account_role, u.status AS account_status
                FROM evaluators e
                LEFT JOIN users u ON e.account_user_id = u.id AND u.org_id = e.org_id
                WHERE e.org_id = ? AND e.account_user_id = ?
                ORDER BY CASE e.status WHEN 'active' THEN 0 ELSE 1 END, e.created_at DESC, e.id DESC
            ''', (org_id, user['id']), fetch=True)
        status_options = ["active", "inactive"]
        status_names = {"active": "在职", "inactive": "停用"}
        
        total_count = len(evaluators or [])
        active_count = sum(1 for eva in (evaluators or []) if eva['status'] == 'active')
        inactive_count = total_count - active_count
        col1, col2, col3 = st.columns(3)
        col1.metric("主评人总数", total_count)
        col2.metric("在职", active_count)
        col3.metric("停用", inactive_count)
        
        if evaluators:
            for eva in evaluators:
                status_name = status_names.get(eva['status'], eva['status'] or '-')
                title_text = eva['title'] or '未填写职称'
                if eva.get('account_user_id') and eva.get('account_username'):
                    account_text = format_account_option(eva['account_user_id'])
                elif eva.get('account_user_id'):
                    account_text = "账号已删除"
                else:
                    account_text = "未关联账号"
                
                with st.expander(f"**{eva['name']}** - {title_text} ({status_name}，{account_text})"):
                    if can_manage_evaluators:
                        with st.form(f"edit_evaluator_form_{eva['id']}"):
                            edit_col1, edit_col2 = st.columns(2)
                            
                            with edit_col1:
                                edit_name = st.text_input("姓名 *", value=eva['name'], key=f"eva_name_{eva['id']}")
                                edit_title = st.text_input("职称/职务", value=eva['title'] or "", key=f"eva_title_{eva['id']}")
                                edit_specialty = st.text_input("专业方向", value=eva['specialty'] or "", key=f"eva_specialty_{eva['id']}")
                            
                            with edit_col2:
                                edit_phone = st.text_input("联系电话", value=eva['phone'] or "", key=f"eva_phone_{eva['id']}")
                                edit_email = st.text_input("邮箱", value=eva['email'] or "", key=f"eva_email_{eva['id']}")
                                edit_status = st.selectbox(
                                    "状态",
                                    status_options,
                                    index=status_options.index(eva['status']) if eva['status'] in status_options else 0,
                                    format_func=lambda x: status_names.get(x, x),
                                    key=f"eva_status_{eva['id']}"
                                )
                                current_account_id = eva.get('account_user_id')
                                edit_account_user_id = st.selectbox(
                                    "关联账号",
                                    account_options,
                                    index=account_options.index(current_account_id) if current_account_id in account_options else 0,
                                    format_func=format_account_option,
                                    key=f"eva_account_{eva['id']}"
                                )
                            
                            if st.form_submit_button("保存修改", use_container_width=True):
                                edit_name = edit_name.strip()
                                if not edit_name:
                                    st.error("请填写主评人姓名")
                                else:
                                    execute_query('''
                                        UPDATE evaluators
                                        SET name = ?, title = ?, specialty = ?, phone = ?, email = ?, status = ?, account_user_id = ?
                                        WHERE id = ? AND org_id = ?
                                    ''', (
                                        edit_name,
                                        edit_title.strip(),
                                        edit_specialty.strip(),
                                        edit_phone.strip(),
                                        edit_email.strip(),
                                        edit_status,
                                        edit_account_user_id,
                                        eva['id'],
                                        org_id
                                    ), commit=True)
                                    add_log(user['id'], user['username'], '', '编辑主评人', 'evaluators',
                                            f'编辑主评人: {edit_name}', get_client_ip())
                                    notify_and_rerun("主评人信息已保存", level='success')
                        
                        op_col1, op_col2, op_col3 = st.columns([1, 1, 4])
                        with op_col1:
                            if eva['status'] == 'active':
                                if st.button("停用", key=f"deactivate_eva_{eva['id']}"):
                                    execute_query("UPDATE evaluators SET status = 'inactive' WHERE id = ? AND org_id = ?",
                                                  (eva['id'], org_id), commit=True)
                                    notify_and_rerun("主评人已停用", level='success')
                            else:
                                if st.button("启用", key=f"activate_eva_{eva['id']}"):
                                    execute_query("UPDATE evaluators SET status = 'active' WHERE id = ? AND org_id = ?",
                                                  (eva['id'], org_id), commit=True)
                                    notify_and_rerun("主评人已启用", level='success')
                        with op_col2:
                            if st.button("删除", key=f"delete_eva_{eva['id']}"):
                                execute_query("DELETE FROM evaluators WHERE id = ? AND org_id = ?",
                                              (eva['id'], org_id), commit=True)
                                add_log(user['id'], user['username'], '', '删除主评人', 'evaluators',
                                        f'删除主评人: {eva["name"]}', get_client_ip())
                                notify_and_rerun("主评人已删除", level='success')
                    else:
                        view_col1, view_col2 = st.columns(2)
                        with view_col1:
                            st.write(f"**姓名:** {eva['name']}")
                            st.write(f"**职称/职务:** {eva['title'] or '-'}")
                            st.write(f"**专业方向:** {eva['specialty'] or '-'}")
                        with view_col2:
                            st.write(f"**联系电话:** {eva['phone'] or '-'}")
                            st.write(f"**邮箱:** {eva['email'] or '-'}")
                            st.write(f"**状态:** {status_name}")
        else:
            st.info("暂无主评人")
        
        if can_manage_evaluators:
            st.markdown("---")
            st.markdown("#### 添加主评人")
            
            with st.form("add_evaluator_form"):
                add_col1, add_col2 = st.columns(2)
                
                with add_col1:
                    eva_name = st.text_input("姓名 *")
                    eva_title = st.text_input("职称/职务")
                    eva_specialty = st.text_input("专业方向")
                
                with add_col2:
                    eva_phone = st.text_input("联系电话")
                    eva_email = st.text_input("邮箱")
                    eva_status = st.selectbox("状态", status_options, format_func=lambda x: status_names.get(x, x))
                    eva_account_user_id = st.selectbox("关联账号", account_options, format_func=format_account_option)
                
                if st.form_submit_button("添加主评人", use_container_width=True):
                    eva_name = eva_name.strip()
                    if not eva_name:
                        st.error("请填写主评人姓名")
                    else:
                        execute_query('''
                            INSERT INTO evaluators (org_id, account_user_id, name, title, specialty, phone, email, status)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            org_id,
                            eva_account_user_id,
                            eva_name,
                            eva_title.strip(),
                            eva_specialty.strip(),
                            eva_phone.strip(),
                            eva_email.strip(),
                            eva_status
                        ), commit=True)
                        add_log(user['id'], user['username'], '', '新增主评人', 'evaluators',
                                f'新增主评人: {eva_name}', get_client_ip())
                        notify_and_rerun("主评人添加成功", level='success')
    with tab3:
        st.subheader("业绩记录")
        
        achievements = execute_query("SELECT * FROM achievements WHERE org_id = ? ORDER BY achievement_date DESC", (org_id,), fetch=True)
        
        if achievements:
            for ach in achievements:
                with st.expander(f"**{ach['title']}** - {ach['achievement_date'] or '-'}"):
                    st.write(ach['content'] or '')
                    
                    if st.button("删除", key=f"del_ach_{ach['id']}"):
                        execute_query("DELETE FROM achievements WHERE id = ?", (ach['id'],), commit=True)
                        notify_and_rerun("已删除", level='success')
        st.markdown("---")
        st.markdown("#### 添加业绩记录")
        
        with st.form("add_achievement_form"):
            ach_title = st.text_input("业绩标题 *")
            ach_content = st.text_area("业绩内容")
            ach_date = st.date_input("业绩日期")
            
            if st.form_submit_button("添加"):
                if ach_title:
                    execute_query('''
                        INSERT INTO achievements (org_id, title, content, achievement_date)
                        VALUES (?, ?, ?, ?)
                    ''', (org_id, ach_title, ach_content, ach_date), commit=True)
                    notify_and_rerun("添加成功", level='success')
    with tab4:
        st.subheader("培训记录")
        
        trainings = execute_query("SELECT * FROM trainings WHERE org_id = ? ORDER BY training_date DESC", (org_id,), fetch=True)
        
        if trainings:
            for train in trainings:
                with st.expander(f"**{train['title']}** - {train['training_date'] or '-'}"):
                    st.write(f"**培训讲师:** {train['trainer'] or '-'}")
                    st.write(f"**培训时长:** {train['duration'] or '-'} 小时")
                    st.write(f"**参与人数:** {train['participants'] or '-'} 人")
                    st.write(f"**培训内容:** {train['content'] or '-'}")
                    
                    if st.button("删除", key=f"del_train_{train['id']}"):
                        execute_query("DELETE FROM trainings WHERE id = ?", (train['id'],), commit=True)
                        notify_and_rerun("已删除", level='success')
        st.markdown("---")
        st.markdown("#### 添加培训记录")
        
        with st.form("add_training_form"):
            train_title = st.text_input("培训标题 *")
            train_trainer = st.text_input("培训讲师")
            train_date = st.date_input("培训日期")
            train_duration = st.number_input("培训时长(小时)", min_value=0)
            train_participants = st.number_input("参与人数", min_value=0)
            train_content = st.text_area("培训内容")
            
            if st.form_submit_button("添加"):
                if train_title:
                    execute_query('''
                        INSERT INTO trainings (org_id, title, trainer, training_date, duration, participants, content)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (org_id, train_title, train_trainer, train_date, train_duration, train_participants, train_content), commit=True)
                    notify_and_rerun("添加成功", level='success')
def render_org_sub_accounts():
    """机构端子账号管理"""
    st.title("👥 子账号管理")
    
    user = st.session_state['user']
    org_id = user['org_id']
    
    # 页面入口由集中权限表兜底，这里保留二次校验防止函数被误调用。
    if not can_access_page(user['role'], 'sub_accounts'):
        st.warning("您没有权限管理子账号")
        return
    
    tab1, tab2 = st.tabs(["子账号列表", "新增子账号"])
    
    with tab1:
        sub_users = execute_query('''
            SELECT * FROM users WHERE org_id = ? AND role = 'org_user' ORDER BY created_at DESC
        ''', (org_id,), fetch=True)
        
        if sub_users:
            for sub in sub_users:
                status_name = "正常" if sub['status'] == 'active' else "冻结"
                is_editing = st.session_state.get('edit_sub_id') == sub['id']
                
                with st.expander(f"**{sub['username']}** - {sub['real_name'] or '-'} ({status_name})", expanded=is_editing):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write(f"**姓名:** {sub['real_name'] or '-'}")
                        st.write(f"**手机号:** {sub['phone'] or '-'}")
                        st.write(f"**邮箱:** {sub['email'] or '-'}")
                    
                    with col2:
                        st.write(f"**状态:** {status_name}")
                        st.write(f"**创建时间:** {sub['created_at']}")
                        linked_evaluators = execute_query('''
                            SELECT name, title, status
                            FROM evaluators
                            WHERE org_id = ? AND account_user_id = ?
                            ORDER BY CASE status WHEN 'active' THEN 0 ELSE 1 END, created_at DESC, id DESC
                        ''', (org_id, sub['id']), fetch=True)
                        if linked_evaluators:
                            evaluator_names = [
                                f"{eva['name']}（{eva['title'] or '未填写职称'}）"
                                for eva in linked_evaluators
                            ]
                            st.write(f"**关联主评人:** {'、'.join(evaluator_names)}")
                        else:
                            st.write("**关联主评人:** -")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        if sub['status'] == 'active':
                            if st.button("🔒 冻结", key=f"freeze_sub_{sub['id']}"):
                                execute_query("UPDATE users SET status = 'inactive' WHERE id = ?", (sub['id'],), commit=True)
                                notify_and_rerun("已冻结", level='success')
                        else:
                            if st.button("🔓 解冻", key=f"unfreeze_sub_{sub['id']}"):
                                execute_query("UPDATE users SET status = 'active' WHERE id = ?", (sub['id'],), commit=True)
                                notify_and_rerun("已解冻", level='success')
                    with col2:
                        if st.button("🔑 重置密码", key=f"reset_sub_{sub['id']}"):
                            default_password = get_default_password_for_role(sub['role'])
                            new_hash = hash_password(default_password)
                            execute_query(
                                "UPDATE users SET password_hash = ?, must_change_password = 1, updated_at = ? WHERE id = ?",
                                (new_hash, datetime.now(), sub['id']),
                                commit=True
                            )
                            notify_and_rerun(f"密码已重置为: {default_password}，下次登录必须修改密码", level='success')
                    
                    with col3:
                        if st.button("✏️ 编辑", key=f"edit_sub_{sub['id']}"):
                            st.session_state['edit_sub_id'] = sub['id']
                            st.rerun()
                    
                    with col4:
                        if st.button("🗑️ 删除", key=f"del_sub_{sub['id']}"):
                            execute_query("UPDATE evaluators SET account_user_id = NULL WHERE account_user_id = ? AND org_id = ?",
                                          (sub['id'], org_id), commit=True)
                            execute_query("DELETE FROM users WHERE id = ?", (sub['id'],), commit=True)
                            notify_and_rerun("已删除", level='success')
                    
                    if is_editing:
                        st.markdown("---")
                        st.markdown("#### 编辑子账号")
                        
                        with st.form(f"edit_sub_form_{sub['id']}"):
                            edit_col1, edit_col2 = st.columns(2)
                            
                            with edit_col1:
                                edit_username = st.text_input("用户名 *", value=sub['username'], key=f"edit_sub_username_{sub['id']}")
                                edit_real_name = st.text_input("姓名 *", value=sub['real_name'] or "", key=f"edit_sub_real_name_{sub['id']}")
                                edit_phone = st.text_input("手机号 *", value=sub['phone'] or "", key=f"edit_sub_phone_{sub['id']}")
                            
                            with edit_col2:
                                edit_email = st.text_input("邮箱 *", value=sub['email'] or "", key=f"edit_sub_email_{sub['id']}")
                                edit_status = st.selectbox(
                                    "账号状态",
                                    ["active", "inactive"],
                                    index=0 if sub['status'] == 'active' else 1,
                                    format_func=lambda x: {"active": "正常", "inactive": "冻结"}[x],
                                    key=f"edit_sub_status_{sub['id']}"
                                )
                            
                            save_col, cancel_col = st.columns(2)
                            with save_col:
                                save_edit = st.form_submit_button("保存修改", use_container_width=True)
                            with cancel_col:
                                cancel_edit = st.form_submit_button("取消编辑", use_container_width=True)
                            
                            if cancel_edit:
                                st.session_state.pop('edit_sub_id', None)
                                st.rerun()
                            
                            if save_edit:
                                edit_username = edit_username.strip()
                                edit_real_name = edit_real_name.strip()
                                edit_phone = edit_phone.strip()
                                edit_email = edit_email.strip()
                                
                                if not all([edit_username, edit_real_name, edit_phone, edit_email]):
                                    st.error("请填写所有必填项")
                                else:
                                    duplicate = execute_query('''
                                        SELECT id FROM users
                                        WHERE id != ? AND (username = ? OR phone = ? OR email = ?)
                                    ''', (sub['id'], edit_username, edit_phone, edit_email), fetch=True)
                                    
                                    if duplicate:
                                        st.error("用户名、手机号或邮箱已存在")
                                    else:
                                        execute_query('''
                                            UPDATE users
                                            SET username = ?, real_name = ?, phone = ?, email = ?, status = ?, updated_at = ?
                                            WHERE id = ? AND org_id = ? AND role = 'org_user'
                                        ''', (
                                            edit_username,
                                            edit_real_name,
                                            edit_phone,
                                            edit_email,
                                            edit_status,
                                            datetime.now(),
                                            sub['id'],
                                            org_id
                                        ), commit=True)
                                        st.session_state.pop('edit_sub_id', None)
                                        add_log(user['id'], user['username'], '', '编辑子账号', 'users',
                                                f'编辑子账号: {edit_username}', get_client_ip())
                                        notify_and_rerun("子账号信息已保存", level='success')
        else:
            st.info("暂无子账号")
    
    with tab2:
        st.subheader("新增子账号")
        
        with st.form("add_sub_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                sub_username = st.text_input("用户名 *")
                sub_real_name = st.text_input("姓名 *")
                sub_phone = st.text_input("手机号 *")
            
            with col2:
                sub_email = st.text_input("邮箱 *")
                sub_password = st.text_input("初始密码 *", type="password")
            
            if st.form_submit_button("创建子账号", use_container_width=True):
                required = [sub_username, sub_real_name, sub_phone, sub_email, sub_password]
                if not all(required):
                    st.error("请填写所有必填项")
                elif validate_password_policy(sub_password, allow_default=True):
                    st.error(validate_password_policy(sub_password, allow_default=True))
                else:
                    existing = execute_query("SELECT id FROM users WHERE username = ? OR phone = ? OR email = ?",
                                           (sub_username, sub_phone, sub_email), fetch=True)
                    if existing:
                        st.error("用户名、手机号或邮箱已存在")
                    else:
                        password_hash = hash_password(sub_password)
                        execute_query('''
                            INSERT INTO users (username, password_hash, role, org_id, real_name, phone, email, status, must_change_password)
                            VALUES (?, ?, 'org_user', ?, ?, ?, ?, 'active', 1)
                        ''', (sub_username, password_hash, org_id, sub_real_name, sub_phone, sub_email), commit=True)
                        
                        add_log(user['id'], user['username'], '', '新增子账号', 'users', f'创建子账号: {sub_username}', get_client_ip())
                        notify_and_rerun("子账号创建成功，首次登录必须修改密码", level='success')
def render_org_projects():
    """机构端项目管理"""
    st.title("📋 项目管理")
    
    user = st.session_state['user']
    org_id = user['org_id']
    
    tab1, tab2 = st.tabs(["项目列表", "新建项目"])
    
    with tab1:
        projects = execute_query('''
            SELECT * FROM projects WHERE org_id = ? ORDER BY created_at DESC
        ''', (org_id,), fetch=True)
        
        if projects:
            for proj in projects:
                status_map = {'pending': '待审核', 'in_progress': '进行中', 'completed': '已完成', 'rejected': '已驳回'}
                
                with st.expander(f"**{proj['name']}** - {status_map.get(proj['status'], proj['status'])}"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write(f"**项目分类:** {PROJECT_CATEGORIES.get(proj['category'], {}).get('name', '-') if proj['category'] else '-'}")
                        st.write(f"**当前Gate:** {format_gate(proj['current_stage'])}")
                    
                    with col2:
                        st.write(f"**状态:** {status_map.get(proj['status'], proj['status'])}")
                        st.write(f"**创建时间:** {format_datetime_display(proj['created_at'])}")
                    
                    if proj['description']:
                        st.write(f"**描述:** {proj['description']}")
                    
                    # Gate操作
                    st.markdown("#### Gate操作")
                    
                    steps = execute_query('''
                        SELECT ps.*, (SELECT COUNT(*) FROM project_files WHERE step_id = ps.id) as file_count
                        FROM project_steps ps WHERE ps.project_id = ? ORDER BY ps.stage
                    ''', (proj['id'],), fetch=True)
                    
                    if not steps:
                        # 初始化阶段
                        for stage in range(1, TOTAL_STAGES + 1):
                            execute_query('''
                                INSERT INTO project_steps (project_id, stage, status)
                                VALUES (?, ?, 'pending')
                            ''', (proj['id'], stage), commit=True)
                        st.rerun()
                    else:
                        # 兼容旧项目：自动补齐缺失阶段，确保每个阶段都需要审批
                        existing_stages = {s['stage'] for s in steps}
                        missing_stages = [s for s in range(1, TOTAL_STAGES + 1) if s not in existing_stages]
                        if missing_stages:
                            for stage in missing_stages:
                                execute_query('''
                                    INSERT INTO project_steps (project_id, stage, status)
                                    VALUES (?, ?, 'pending')
                                ''', (proj['id'], stage), commit=True)
                            st.rerun()
                    
                    for step in steps:
                        step_status = {'pending': '⏳ 待提交', 'submitted': '📤 已提交', 'approved': '✅ 已通过', 'rejected': '❌ 已驳回'}
                        
                        step_stage = step['stage']
                        st.markdown(f"**{format_gate(step_stage)}** - {step_status.get(step['status'], step['status'])}")
                        st.caption(f"工程目的：{STAGE_PURPOSES.get(step_stage, '-')}")
                        
                        # 显示已上传文件
                        if step['file_count'] > 0:
                            st.write(f"📁 已上传 {step['file_count']} 个文件")
                        
                        # 上传文件（允许重新提交未通过的阶段）
                        if step['status'] in ('pending', 'rejected'):
                            uploaded_files = st.file_uploader(
                                f"上传{STAGE_NAMES.get(step['stage'], '')}文件 (可多选)",
                                type=['pdf', 'docx', 'doc', 'xlsx', 'xls', 'txt'],
                                accept_multiple_files=True,
                                key=f"upload_{step['id']}"
                            )
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                if st.button(f"📤 批量上传并保存", key=f"upload_btn_{step['id']}"):
                                    if uploaded_files:
                                        for uploaded_file in uploaded_files:
                                            file_path, file_name, file_size = save_uploaded_file(uploaded_file, f"projects/{proj['id']}/{step['stage']}", project_id=proj['id'])
                                            
                                            execute_query('''
                                                INSERT INTO project_files (project_id, step_id, title, file_name, file_path, file_size, upload_by)
                                                VALUES (?, ?, ?, ?, ?, ?, ?)
                                            ''', (proj['id'], step['id'], uploaded_file.name, file_name, file_path, file_size, user['id']), commit=True)
                                            
                                        notify_and_rerun("文件批量上传成功", level='success')
                            with col2:
                                if st.button(f"✅ 提交审核", key=f"submit_btn_{step['id']}"):
                                    # 检查是否上传了文件
                                    files = execute_query("SELECT id FROM project_files WHERE step_id = ?", (step['id'],), fetch=True)
                                    if files:
                                        execute_query('''
                                            UPDATE project_steps SET status = 'submitted', submitted_by = ?, submitted_at = ?
                                            WHERE id = ?
                                        ''', (user['id'], datetime.now(), step['id']), commit=True)
                                        
                                        execute_query("UPDATE projects SET status = 'pending' WHERE id = ?", (proj['id'],), commit=True)
                                        
                                        # 通知超级管理员
                                        admins = execute_query("SELECT id FROM users WHERE role = 'super_admin'", fetch=True)
                                        for admin in admins:
                                            add_message(admin['id'], '新项目待审批', 
                                                      f'机构有新的项目Gate待审批: {proj["name"]} - {format_gate(step["stage"])}')
                                        
                                        notify_and_rerun("已提交审核", level='success')
                                    else:
                                        st.error("请先上传文件")
                        
                        # 查看文件
                        if step['file_count'] > 0:
                            step_files = execute_query("SELECT * FROM project_files WHERE step_id = ?", (step['id'],), fetch=True)
                            for f in step_files:
                                col1, col2 = st.columns([4, 1])
                                with col1:
                                    st.write(f"  📄 {f['file_name']}")
                                with col2:
                                    if st.button("👁️ 查看", key=f"view_org_file_{f['id']}"):
                                        st.session_state['view_file_id'] = f['id']
                                        st.rerun()
        else:
            st.info("暂无项目，请新建项目")
    
    with tab2:
        st.subheader("新建项目")

        category_options = list(PROJECT_CATEGORIES.keys())

        if st.session_state.pop("reset_org_new_project_form", False):
            st.session_state["org_new_proj_name"] = ""
            st.session_state["org_new_proj_desc"] = ""
            if category_options:
                st.session_state["org_new_proj_category"] = category_options[0]
            st.session_state["org_new_proj_subcategory"] = None

        def _sync_org_project_subcategory():
            selected_category = st.session_state.get(
                "org_new_proj_category",
                category_options[0] if category_options else None
            )
            subcats_map = PROJECT_CATEGORIES.get(selected_category, {}).get('subcategories', {})
            current_subcategory = st.session_state.get("org_new_proj_subcategory")

            # 一级分类变化时，二级分类必须落在当前分类的合法集合里；“其他项目”固定为无
            if not subcats_map:
                st.session_state["org_new_proj_subcategory"] = None
            elif current_subcategory not in subcats_map:
                st.session_state["org_new_proj_subcategory"] = next(iter(subcats_map.keys()))

        if category_options and "org_new_proj_category" not in st.session_state:
            st.session_state["org_new_proj_category"] = category_options[0]
        _sync_org_project_subcategory()

        proj_name = st.text_input("项目名称 *", key="org_new_proj_name")

        col1, col2 = st.columns(2)
        with col1:
            proj_category = st.selectbox(
                "项目分类 *",
                category_options,
                format_func=lambda x: f"{x} - {PROJECT_CATEGORIES[x]['name']}",
                key="org_new_proj_category",
                on_change=_sync_org_project_subcategory
            )

        with col2:
            subcats = PROJECT_CATEGORIES.get(proj_category, {}).get('subcategories', {})
            if subcats:
                proj_subcategory = st.selectbox(
                    "二级分类 *",
                    list(subcats.keys()),
                    format_func=lambda x: f"{x} - {subcats[x]}",
                    key="org_new_proj_subcategory"
                )
            else:
                st.selectbox(
                    "二级分类",
                    ["无"],
                    index=0,
                    disabled=True,
                    key="org_new_proj_subcategory_display"
                )
                proj_subcategory = None
                st.session_state["org_new_proj_subcategory"] = None

        proj_desc = st.text_area("项目描述", key="org_new_proj_desc")

        if st.button("创建项目", use_container_width=True, key="create_org_project_btn"):
            if proj_name and proj_name.strip():
                clean_name = proj_name.strip()
                project_id = execute_query('''
                    INSERT INTO projects (name, org_id, category, subcategory, description, created_by, status)
                    VALUES (?, ?, ?, ?, ?, ?, 'pending')
                ''', (clean_name, org_id, proj_category, proj_subcategory, proj_desc, user['id']), commit=True)

                # 创建G0~G8共9个阶段（每个阶段均需提交并审批）
                for stage in range(1, TOTAL_STAGES + 1):
                    execute_query('''
                        INSERT INTO project_steps (project_id, stage, status)
                        VALUES (?, ?, 'pending')
                    ''', (project_id, stage), commit=True)

                add_log(user['id'], user['username'], '', '新建项目', 'projects', f'创建项目: {clean_name}', get_client_ip())
                st.session_state["reset_org_new_project_form"] = True
                notify_and_rerun("项目创建成功", level='success')
            else:
                st.error("请填写项目名称")

def render_org_knowledge():
    """机构端项目智库管理"""
    st.title("📚 项目智库管理")
    
    user = st.session_state['user']
    org_id = user['org_id']
    
    tab1, tab2 = st.tabs(["我的项目", "上传文件"])
    
    with tab1:
        st.subheader("我的项目")
        
        projects = execute_query('''
            SELECT * FROM projects WHERE org_id = ? ORDER BY created_at DESC
        ''', (org_id,), fetch=True)
        
        if projects:
            for proj in projects:
                with st.expander(f"**{proj['name']}**"):
                    st.write(f"**分类:** {PROJECT_CATEGORIES.get(proj['category'], {}).get('name', '-') if proj['category'] else '-'}")
                    st.write(f"**状态:** {proj['status']}")
                    
                    # 查看项目文件
                    st.markdown("**Gate文件:**")
                    steps = execute_query('''
                        SELECT ps.*, (SELECT COUNT(*) FROM project_files WHERE step_id = ps.id AND approval_status = 'approved') as file_count
                        FROM project_steps ps WHERE ps.project_id = ? ORDER BY ps.stage
                    ''', (proj['id'],), fetch=True)
                    
                    if steps:
                        for step in steps:
                            step_stage = step['stage']
                            st.write(f"- {format_gate(step_stage)}: {step['file_count']} 个文件")
                            if step['file_count'] > 0:
                                step_files = execute_query("SELECT * FROM project_files WHERE step_id = ? AND approval_status = 'approved'", (step['id'],), fetch=True)
                                for f in step_files:
                                    col1, col2, col3 = st.columns([4, 1, 1])
                                    with col1:
                                        st.write(f"  📄 {f['file_name']}")
                                    with col2:
                                        if st.button("👁️ 查看", key=f"view_know_file_{f['id']}"):
                                            st.session_state['view_file_id'] = f['id']
                                            st.rerun()
                                    with col3:
                                        file_path = f.get('file_path') or os.path.join(UPLOAD_DIR, f.get('file_name') or '')
                                        if file_path and os.path.exists(file_path):
                                            try:
                                                render_cached_download_button(
                                                    file_path,
                                                    f.get('file_name', 'download'),
                                                    key=f"dl_know_file_{f['id']}"
                                                )
                                            except Exception:
                                                st.write("读取失败")
                                        else:
                                            st.write("文件丢失")
        else:
            st.info("暂无项目")
    
    with tab2:
        st.subheader("上传文件")

        if st.session_state.pop("reset_org_upload_form", False):
            st.session_state["org_upload_file_title"] = ""
            st.session_state["org_upload_file_desc"] = ""
            st.session_state["org_upload_uploader_version"] = st.session_state.get("org_upload_uploader_version", 0) + 1

        upload_category_options = list(PROJECT_CATEGORIES.keys())

        def _sync_org_upload_subcategory():
            selected_category = st.session_state.get(
                "org_upload_file_category",
                upload_category_options[0] if upload_category_options else None
            )
            subcats_map = PROJECT_CATEGORIES.get(selected_category, {}).get('subcategories', {})
            current_subcategory = st.session_state.get("org_upload_file_subcategory")

            # 一级分类变化时，自动校准二级分类；“其他项目”固定显示为无
            if not subcats_map:
                st.session_state["org_upload_file_subcategory"] = None
            elif current_subcategory not in subcats_map:
                st.session_state["org_upload_file_subcategory"] = next(iter(subcats_map.keys()))

        if upload_category_options and "org_upload_file_category" not in st.session_state:
            st.session_state["org_upload_file_category"] = upload_category_options[0]
        _sync_org_upload_subcategory()

        file_title = st.text_input("文件标题 *", key="org_upload_file_title")

        col1, col2 = st.columns(2)
        with col1:
            file_category = st.selectbox(
                "文件类型 *",
                upload_category_options,
                format_func=lambda x: f"{x} - {PROJECT_CATEGORIES[x]['name']}",
                key="org_upload_file_category",
                on_change=_sync_org_upload_subcategory
            )

        with col2:
            subcats = PROJECT_CATEGORIES.get(file_category, {}).get('subcategories', {})
            if subcats:
                file_subcategory = st.selectbox(
                    "二级分类",
                    list(subcats.keys()),
                    format_func=lambda x: f"{x} - {subcats[x]}",
                    key="org_upload_file_subcategory"
                )
            else:
                st.selectbox(
                    "二级分类",
                    ["无"],
                    index=0,
                    disabled=True,
                    key="org_upload_file_subcategory_display"
                )
                file_subcategory = None
                st.session_state["org_upload_file_subcategory"] = None

        org_projects = execute_query(
            "SELECT id, name FROM projects WHERE org_id = ? ORDER BY created_at DESC",
            (org_id,),
            fetch=True
        )
        project_name_map = {p['id']: (p.get('name') or f"项目{p['id']}") for p in (org_projects or [])}
        project_options = list(project_name_map.keys())

        if project_options and st.session_state.get("org_upload_target_project_id") not in project_name_map:
            st.session_state["org_upload_target_project_id"] = project_options[0]

        if project_options:
            target_project_id = st.selectbox(
                "发布项目 *",
                project_options,
                format_func=lambda pid: f"{project_name_map[pid]}（ID:{pid}）",
                key="org_upload_target_project_id"
            )
            target_project_name = project_name_map[target_project_id]
        else:
            st.selectbox(
                "发布项目",
                ["无（请先创建项目）"],
                index=0,
                disabled=True,
                key="org_upload_target_project_display"
            )
            target_project_id = None
            target_project_name = None
            st.info("请先在“项目管理 -> 新建项目”创建项目，再上传文件。")

        file_desc = st.text_area("文件描述", key="org_upload_file_desc")

        uploader_version = st.session_state.get("org_upload_uploader_version", 0)
        uploaded_file = st.file_uploader(
            "上传文件",
            type=['pdf', 'docx', 'doc', 'xlsx', 'xls', 'txt'],
            key=f"org_upload_file_uploader_{uploader_version}"
        )

        if st.button("上传文件", use_container_width=True, key="org_upload_file_btn"):
            if not target_project_id:
                st.error("请先选择发布项目")
            elif file_title and file_title.strip() and uploaded_file:
                clean_title = file_title.strip()
                file_path, file_name, file_size = save_uploaded_file(
                    uploaded_file,
                    f"projects/{target_project_id}/knowledge",
                    project_id=target_project_id
                )

                # 子账号上传需要审批
                approval_status = 'pending' if user['role'] == 'org_user' else 'pending'

                execute_query('''
                    INSERT INTO project_files (
                        project_id, title, file_type, category, subcategory, file_path, file_name,
                        file_size, publish_org, description, upload_by, approval_status
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    target_project_id, clean_title, uploaded_file.type, file_category, file_subcategory,
                    file_path, file_name, file_size, target_project_name, file_desc, user['id'], approval_status
                ), commit=True)

                # 通知审批
                if user['role'] == 'org_user':
                    # 通知机构主账号
                    main_user = execute_query("SELECT id FROM users WHERE org_id = ? AND role = 'org_admin'", (org_id,), fetch=True)
                    if main_user:
                        add_message(main_user[0]['id'], '新文件待审批', f'有新的项目文件待审批: {clean_title}（项目：{target_project_name}）')
                else:
                    # 通知超级管理员
                    admins = execute_query("SELECT id FROM users WHERE role = 'super_admin'", fetch=True)
                    for admin in admins:
                        add_message(admin['id'], '新文件待审批', f'机构上传了新的项目文件待审批: {clean_title}（项目：{target_project_name}）')

                add_log(user['id'], user['username'], '', '上传文件', 'files', f'上传文件: {clean_title}', get_client_ip())
                st.session_state["reset_org_upload_form"] = True
                notify_and_rerun("文件上传成功，等待审批", level='success')
            elif not (file_title and file_title.strip()):
                st.error("请填写文件标题")
            else:
                st.error("请先选择要上传的文件")

def render_org_todos():
    """机构端待办事项"""
    st.title("✅ 待办事项")
    
    user = st.session_state['user']
    org_info_labels = {
        'name': '机构名称',
        'org_type': '机构类型',
        'credit_code': '统一社会信用代码',
        'legal_person': '法定代表人',
        'contact_person': '联系人',
        'contact_phone': '联系电话',
        'contact_email': '联系邮箱',
        'address': '机构地址',
        'description': '机构简介',
    }
    
    tab1, tab2 = st.tabs(["待办列表", "新增待办"])
    
    with tab1:
        todos = execute_query('''
            SELECT * FROM todos WHERE user_id = ? ORDER BY 
            CASE WHEN status = 'pending' THEN 0 ELSE 1 END,
            CASE priority WHEN 'high' THEN 0 WHEN 'medium' THEN 1 ELSE 2 END,
            created_at DESC
        ''', (user['id'],), fetch=True)
        
        if todos:
            for todo in todos:
                status_icon = "⏳" if todo['status'] == 'pending' else "✅"
                priority_colors = {'high': '#f8d7da', 'medium': '#fff3cd', 'low': '#d4edda'}
                display_priority = todo['priority'] if todo['status'] == 'pending' else 'low'
                display_content = todo['content'] or ''
                if todo['status'] != 'pending':
                    display_content = display_content.replace('请审批', '已审批')
                
                st.markdown(f"""
                <div style="background: {priority_colors.get(display_priority, '#f8f9fa')}; 
                            padding: 15px; border-radius: 10px; margin: 10px 0;
                            border-left: 4px solid {'#dc3545' if display_priority == 'high' else '#ffc107' if display_priority == 'medium' else '#28a745'};">
                    <h4>{status_icon} {todo['title']}</h4>
                    <p>{display_content}</p>
                    <small>优先级: {display_priority} | 截止日期: {todo['due_date'] or '未设置'}</small>
                </div>
                """, unsafe_allow_html=True)
                
                if todo.get('related_type') == 'org_info_update' and todo.get('related_id') and todo['status'] == 'pending':
                    request_rows = execute_query('''
                        SELECT r.*, u.username AS submitter_username, u.real_name AS submitter_name
                        FROM org_info_update_requests r
                        LEFT JOIN users u ON r.submitted_by = u.id
                        WHERE r.id = ? AND r.approver_id = ? AND r.status = 'pending'
                    ''', (todo['related_id'], user['id']), fetch=True)
                    
                    if request_rows:
                        request = request_rows[0]
                        try:
                            old_data = json.loads(request['old_data'] or '{}')
                            new_data = json.loads(request['new_data'] or '{}')
                        except json.JSONDecodeError:
                            old_data = {}
                            new_data = {}
                        
                        st.markdown("**机构信息变更内容**")
                        changed_rows = []
                        for field, label in org_info_labels.items():
                            old_value = old_data.get(field, '')
                            new_value = new_data.get(field, '')
                            if old_value != new_value:
                                changed_rows.append({
                                    "字段": label,
                                    "原值": old_value or "-",
                                    "新值": new_value or "-"
                                })
                        if changed_rows:
                            st.dataframe(pd.DataFrame(changed_rows), use_container_width=True, hide_index=True)
                        else:
                            st.info("未检测到字段变化")
                        
                        review_comment = st.text_area("审批意见", key=f"org_info_review_comment_{todo['id']}")
                        approve_col, reject_col = st.columns(2)
                        
                        with approve_col:
                            if st.button("通过并更新机构信息", key=f"approve_org_info_{todo['id']}", use_container_width=True):
                                execute_query('''
                                    UPDATE organizations SET name = ?, org_type = ?, credit_code = ?, legal_person = ?,
                                    contact_person = ?, contact_phone = ?, contact_email = ?, address = ?, description = ?, updated_at = ?
                                    WHERE id = ?
                                ''', (
                                    new_data.get('name', ''),
                                    new_data.get('org_type', ''),
                                    new_data.get('credit_code', ''),
                                    new_data.get('legal_person', ''),
                                    new_data.get('contact_person', ''),
                                    new_data.get('contact_phone', ''),
                                    new_data.get('contact_email', ''),
                                    new_data.get('address', ''),
                                    new_data.get('description', ''),
                                    datetime.now(),
                                    request['org_id']
                                ), commit=True)
                                execute_query('''
                                    UPDATE org_info_update_requests
                                    SET status = 'approved', review_comment = ?, reviewed_at = ?
                                    WHERE id = ? AND approver_id = ?
                                ''', (review_comment.strip(), datetime.now(), request['id'], user['id']), commit=True)
                                execute_query(
                                    '''
                                    UPDATE todos
                                    SET status = 'completed',
                                        priority = 'low',
                                        content = REPLACE(COALESCE(content, ''), '请审批', '已审批'),
                                        completed_at = ?
                                    WHERE id = ?
                                    ''',
                                    (datetime.now(), todo['id']),
                                    commit=True
                                )
                                add_message(request['submitted_by'], '机构信息变更已通过', '您提交的机构信息变更申请已通过')
                                add_log(user['id'], user['username'], '', '审批机构信息变更', 'organizations',
                                        f'通过申请ID: {request["id"]}', get_client_ip())
                                notify_and_rerun("已通过并更新机构信息", level='success')
                        
                        with reject_col:
                            if st.button("驳回申请", key=f"reject_org_info_{todo['id']}", use_container_width=True):
                                execute_query('''
                                    UPDATE org_info_update_requests
                                    SET status = 'rejected', review_comment = ?, reviewed_at = ?
                                    WHERE id = ? AND approver_id = ?
                                ''', (review_comment.strip(), datetime.now(), request['id'], user['id']), commit=True)
                                execute_query(
                                    '''
                                    UPDATE todos
                                    SET status = 'completed',
                                        priority = 'low',
                                        content = REPLACE(COALESCE(content, ''), '请审批', '已审批'),
                                        completed_at = ?
                                    WHERE id = ?
                                    ''',
                                    (datetime.now(), todo['id']),
                                    commit=True
                                )
                                add_message(request['submitted_by'], '机构信息变更已驳回',
                                            f'您提交的机构信息变更申请已驳回。意见: {review_comment.strip() or "无"}')
                                add_log(user['id'], user['username'], '', '审批机构信息变更', 'organizations',
                                        f'驳回申请ID: {request["id"]}', get_client_ip())
                                notify_and_rerun("已驳回申请", level='success')
                    else:
                        st.warning("关联审批申请不存在或已处理")
                else:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        if todo['status'] == 'pending' and st.button("完成", key=f"complete_{todo['id']}"):
                            execute_query("UPDATE todos SET status = 'completed', completed_at = ? WHERE id = ?",
                                        (datetime.now(), todo['id']), commit=True)
                            notify_and_rerun("已完成", level='success')
                    with col2:
                        if st.button("删除", key=f"del_todo_{todo['id']}"):
                            execute_query("DELETE FROM todos WHERE id = ?", (todo['id'],), commit=True)
                            notify_and_rerun("已删除", level='success')
        else:
            st.info("暂无待办事项")
    
    with tab2:
        with st.form("add_todo_form"):
            todo_title = st.text_input("待办标题 *")
            todo_content = st.text_area("待办内容")
            
            col1, col2 = st.columns(2)
            with col1:
                todo_priority = st.selectbox("优先级", ["high", "medium", "low"],
                                            format_func=lambda x: {"high": "高", "medium": "中", "low": "低"}[x])
            with col2:
                todo_due = st.date_input("截止日期")
            
            if st.form_submit_button("添加待办", use_container_width=True):
                if todo_title:
                    execute_query('''
                        INSERT INTO todos (user_id, title, content, priority, due_date)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (user['id'], todo_title, todo_content, todo_priority, todo_due), commit=True)
                    notify_and_rerun("添加成功", level='success')
def render_org_messages():
    """机构端消息通知"""
    render_messages_page("📨 消息通知", "org_messages")

# ==================== 文件预览弹窗 ====================
def render_file_preview():
    """渲染文件预览弹窗"""
    if 'view_file_id' not in st.session_state:
        return
    
    file_id = st.session_state['view_file_id']
    file_info = execute_query("SELECT * FROM project_files WHERE id = ?", (file_id,), fetch=True)
    
    if file_info:
        file_data = file_info[0]
        
        # 使用 Streamlit 的 modal（如果可用），否则回退到 container
        title = f"📄 {file_data['title']}"
        if hasattr(st, 'modal'):
            with st.modal(title=title):
                st.write(f"**文件名:** {file_data['file_name']}")
                st.write(f"**上传时间:** {format_datetime_display(file_data['upload_at'])}")
                if file_data['file_path'] and os.path.exists(file_data['file_path']):
                    display_file_preview(file_data['file_path'], file_data['file_name'])
                else:
                    st.error("文件不存在")

                if st.button("❌ 关闭", key="close_preview"):
                    del st.session_state['view_file_id']
                    st.rerun()
        else:
            # 兼容旧版本 Streamlit：最小化自定义 overlay 的 DOM 操作
            st.markdown("### " + title)
            st.write(f"**文件名:** {file_data['file_name']}")
            st.write(f"**上传时间:** {format_datetime_display(file_data['upload_at'])}")
            if file_data['file_path'] and os.path.exists(file_data['file_path']):
                display_file_preview(file_data['file_path'], file_data['file_name'])
            else:
                st.error("文件不存在")

            if st.button("❌ 关闭", key="close_preview"):
                del st.session_state['view_file_id']
                st.rerun()

# ==================== 主函数 ====================
def main():
    """主函数"""
    # 初始化数据库（仅首次执行）
    ensure_app_initialized(APP_SCHEMA_VERSION)
    
    # 应用自定义样式
    apply_custom_styles()
    install_page_switch_mask()
    
    # 检查登录状态
    if (
        'logged_in' not in st.session_state
        or not st.session_state['logged_in']
        or 'user' not in st.session_state
        or not st.session_state.get('user')
    ):
        st.session_state['logged_in'] = False
        render_flash_message()
        render_login_page()
        return

    ensure_current_account_active()

    if is_password_change_required():
        render_flash_message()
        render_required_password_change()
        return

    handle_query_navigation()
    
    # 渲染侧边栏
    render_sidebar()
    
    # 根据角色和页面渲染内容
    user = st.session_state['user']
    role = user['role']
    page, allowed = ensure_current_page_access(role)
    
    # 在主内容区域的右上角增加用户信息与退出折叠面板
    role_name = ROLE_NAMES.get(role, role)
    display_name = user.get('real_name') or user['username']
    
    # 注入锚点，CSS会抓取此锚点来将其下一个相邻的容器悬浮定位
    with st.popover(f"👤 {display_name} ({role_name})"):
        if st.button("🚪 退出登录", key="top_right_logout"):
            logout_current_user()

    # 文件预览内容（放在右上角账号按钮之后渲染，避免挤压按钮位置）
    render_file_preview()
    render_flash_message()

    page_container = st.empty()
    with page_container.container():
        admin_routes = {
            'dashboard': render_admin_dashboard,
            'organizations': render_admin_organizations,
            'users': render_admin_users,
            'projects': render_admin_projects,
            'logs': render_admin_logs,
            'export': render_admin_export,
            'approval': render_admin_approval,
            'messages': render_admin_messages,
            'indicators': render_admin_indicators,
            'visualization': render_admin_visualization,
        }
        org_routes = {
            'dashboard': render_org_dashboard,
            'info': render_org_info,
            'sub_accounts': render_org_sub_accounts,
            'projects': render_org_projects,
            'knowledge': render_org_knowledge,
            'todos': render_org_todos,
            'messages': render_org_messages,
        }
        routes = admin_routes if role == 'super_admin' else org_routes
        renderer = routes.get(page)
        if not allowed or renderer is None:
            render_access_denied()
        else:
            renderer()

        render_page_ready_marker()

if __name__ == "__main__":
    main()
