import streamlit as st
import pandas as pd
import asyncio
import os
import subprocess
import glob
import time
import logging
import warnings
import streamlit.components.v1 as components

# 静默三方库中英文冗余 Debug/Info 日志输出，保持终端纯净清爽
warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.ERROR)
for _logger_name in ["telethon", "playwright", "asyncio", "urllib3", "httpx", "httpcore", "PIL", "tornado", "streamlit"]:
    logging.getLogger(_logger_name).setLevel(logging.ERROR)

from config import (
    MODE1_EXCEL, MODE2_EXCEL, MODE3_EXCEL, MODE4_EXCEL,
    TELEGRAM_SESSION_DIR, DATA_DIR, DSERS_SESSION_DIR, MABANG_SESSION_DIR,
    DIANXIAOMI_SESSION_DIR, DIANXIAOMI_SHOPS_CACHE,
    DSERS_TEMPLATE, DSERS_IMPORT_XLSX, DSERS_IMPORT_CSV,
    SCRIPT_TEMPLATE, ORDER_TEMPLATE, SESSIONS_DIR
)
from automators.telegram_cpf_bot import run_cpf_query
from automators.dsers_update_bot import run_dsers_rename
from automators.order_template_utils import clean_order_template_to_script, sync_cpf_results_to_order_template
from automators.dianxiaomi_stock_bot import (
    load_cached_shops, fetch_dianxiaomi_shops, run_dianxiaomi_stock_update
)
from automators.mabang_export_bot import run_mabang_export as run_mabang_cpf_export
from automators.mabang_dsers_export import run_mabang_export as run_mabang_dsers_export
from automators.mabang_update_bot import run_mabang_batch_update
from automators.dsers_clean_and_map import run_dsers_clean_and_map
from automators.dsers_cpf_bridge import export_to_cpf_template, merge_cpf_results
from automators.dsers_import_bot import run_dsers_import
from automators.task_manager import global_task_manager

st.set_page_config(page_title="Global Pipeline Studio", layout="wide", initial_sidebar_state="collapsed")

def _safe_param(key):
    val = st.query_params.get(key, "")
    if isinstance(val, list):
        return val[0] if val else ""
    return str(val) if val is not None else ""

_relay_action = _safe_param("relay_action")
_relay_task_id = _safe_param("relay_task_id")
_url_route = _safe_param("route")

if 'route' not in st.session_state:
    st.session_state.route = _url_route if _url_route in ("A", "B") else None

if _relay_action and _relay_task_id:
    if _relay_action == "kill":
        global_task_manager.cancel_task(_relay_task_id)
    elif _relay_action == "pause":
        global_task_manager.pause_task(_relay_task_id)
    elif _relay_action == "resume":
        global_task_manager.resume_task(_relay_task_id)
    
    st.query_params.clear()
    if st.session_state.route:
        st.query_params["route"] = st.session_state.route
    st.rerun()

if st.session_state.route and st.query_params.get("route") != st.session_state.route:
    st.query_params["route"] = st.session_state.route

@st.dialog("选择表格类型")
def select_uploaded_template_dialog(file_buffer, file_id):
    st.markdown("请指定您刚上传的表格属于以下哪种类型：")
    template_choice = st.radio(
        "请选择对应模板：",
        ["dsers模板.xlsx", "import_orders.xlsx", "脚本模板.xlsx", "下单模板.xlsx"],
        index=0,
        label_visibility="collapsed"
    )
    if st.button("确认并载入", type="primary", use_container_width=True):
        st.session_state["handled_file_id"] = file_id
        st.session_state["active_template_type"] = template_choice
        
        if template_choice == "dsers模板.xlsx":
            dest_path = DSERS_TEMPLATE
        elif template_choice == "import_orders.xlsx":
            dest_path = DSERS_IMPORT_XLSX
        elif template_choice == "脚本模板.xlsx":
            dest_path = SCRIPT_TEMPLATE
        else:
            dest_path = ORDER_TEMPLATE
            
        with open(dest_path, "wb") as f:
            f.write(file_buffer)
            
        if template_choice == "下单模板.xlsx":
            is_rename = (st.session_state.get("route") == "B" and st.session_state.get("sw_dsers_rename_key", False))
            clean_order_template_to_script(dest_path, SCRIPT_TEMPLATE, st.session_state.route, sw_dsers_rename=is_rename)
            
        st.rerun()

def on_dsers_rename_change():
    if st.session_state.get("sw_dsers_rename_key", False):
        st.session_state["sw_dsers_clean_key"] = False
        st.session_state["sw_dsers_cpf_check_key"] = False
        st.session_state["sw_dsers_cpf_merge_key"] = False
        st.session_state["sw_dsers_mabang_key"] = False
        st.session_state["sw_dsers_import_key"] = False
        if os.path.exists(ORDER_TEMPLATE):
            clean_order_template_to_script(ORDER_TEMPLATE, SCRIPT_TEMPLATE, "B", sw_dsers_rename=True)

def on_dsers_normal_change():
    if any([
        st.session_state.get("sw_dsers_clean_key", False),
        st.session_state.get("sw_dsers_cpf_check_key", False),
        st.session_state.get("sw_dsers_cpf_merge_key", False),
        st.session_state.get("sw_dsers_mabang_key", False),
        st.session_state.get("sw_dsers_import_key", False)
    ]):
        st.session_state["sw_dsers_rename_key"] = False
        if os.path.exists(ORDER_TEMPLATE):
            clean_order_template_to_script(ORDER_TEMPLATE, SCRIPT_TEMPLATE, "B", sw_dsers_rename=False)

@st.fragment(run_every="2s")
def render_live_task_hub():
    tasks = global_task_manager.get_all_tasks()
    active_tasks = global_task_manager.get_active_tasks()
    active_count = len(active_tasks)
    total_count = len(tasks)

    # 1. 构造抽屉内任务列表的 HTML
    task_items_html = ""
    if not tasks:
        task_items_html = """
        <div style="text-align: center; color: #888888; padding: 3rem 1rem; font-size: 0.9rem;">
            <div style="font-size: 2rem; margin-bottom: 0.5rem;">📋</div>
            暂无正在运行或历史任务记录
        </div>
        """
    else:
        for task in reversed(tasks):
            action_req = getattr(task, 'action_required', None)
            if not action_req and task.error:
                import re
                m = re.search(r'(?:需配置选项|需选择选项|需处理选项|需要选项|阻断拦截|店小秘拦截)[：:\s]+([^\n，。]+)', task.error)
                if m:
                    action_req = m.group(1).strip()
                else:
                    m2 = re.search(r'(请选择[^\n，。】\]]+)', task.error)
                    if m2:
                        action_req = m2.group(1).strip()

            status_badge_style = {
                "RUNNING": "background: rgba(52, 199, 89, 0.12); color: #248a3d; border: 1px solid rgba(52, 199, 89, 0.25);",
                "PAUSED": "background: rgba(255, 149, 0, 0.12); color: #c97500; border: 1px solid rgba(255, 149, 0, 0.25);",
                "SUCCESS": "background: rgba(0, 122, 255, 0.12); color: #007aff; border: 1px solid rgba(0, 122, 255, 0.25);",
                "FAILED": "background: rgba(255, 59, 48, 0.12); color: #ff3b30; border: 1px solid rgba(255, 59, 48, 0.25);",
                "CANCELLED": "background: rgba(142, 142, 147, 0.15); color: #8e8e93; border: 1px solid rgba(142, 142, 147, 0.3);"
            }.get(task.status, "background: #eee; color: #666;")
            
            status_text = {
                "RUNNING": "运行中",
                "PAUSED": "已暂停",
                "SUCCESS": "已完成",
                "FAILED": "失败",
                "CANCELLED": "已停止"
            }.get(task.status, task.status)

            if action_req:
                if task.status == "FAILED":
                    status_badge_style = "background: rgba(255, 59, 48, 0.15); color: #e11d48; border: 1px solid rgba(255, 59, 48, 0.35);"
                    status_text = "待选选项"
                elif task.status == "RUNNING":
                    status_badge_style = "background: rgba(255, 149, 0, 0.15); color: #c97500; border: 1px solid rgba(255, 149, 0, 0.35);"
                    status_text = "需选选项"
            
            logs = task.get_logs()
            log_text = "\n".join(logs[-25:]) if logs else "正在初始化并启动独立运行环境..."
            log_html_safe = log_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

            # 彩色控制胶囊按钮 —— 使用 top-level window._vanguardHandleTaskAction 保证永远响应
            actions_btn_html = ""
            btn_pause_style = "background:rgba(255,149,0,0.12);color:#b45309;border:1px solid rgba(255,149,0,0.3);border-radius:6px;padding:4px 11px;font-size:0.74rem;font-weight:600;cursor:pointer;"
            btn_resume_style = "background:rgba(52,199,89,0.12);color:#15803d;border:1px solid rgba(52,199,89,0.3);border-radius:6px;padding:4px 11px;font-size:0.74rem;font-weight:600;cursor:pointer;"
            btn_stop_style = "background:rgba(255,59,48,0.10);color:#e11d48;border:1px solid rgba(255,59,48,0.28);border-radius:6px;padding:4px 11px;font-size:0.74rem;font-weight:600;cursor:pointer;"

            if task.status == "RUNNING":
                actions_btn_html += f"""<button type="button" onclick="window.parent._vanguardHandleTaskAction('pause', '{task.task_id}')" style="{btn_pause_style}">⏸ 暂停</button>"""
                actions_btn_html += f"""<button type="button" onclick="window.parent._vanguardHandleTaskAction('kill', '{task.task_id}')" style="{btn_stop_style}">✕ 停止</button>"""
            elif task.status == "PAUSED":
                actions_btn_html += f"""<button type="button" onclick="window.parent._vanguardHandleTaskAction('resume', '{task.task_id}')" style="{btn_resume_style}">▶ 继续</button>"""
                actions_btn_html += f"""<button type="button" onclick="window.parent._vanguardHandleTaskAction('kill', '{task.task_id}')" style="{btn_stop_style}">✕ 停止</button>"""

            action_banner_html = ""
            if action_req:
                action_banner_html = f"""
                <div style="background: rgba(255, 59, 48, 0.08); border: 1px solid rgba(255, 59, 48, 0.25); border-left: 4px solid #ff3b30; border-radius: 8px; padding: 7px 11px; margin: 0.45rem 0 0.4rem 0; font-size: 0.78rem; line-height: 1.45;">
                    <div style="display: flex; align-items: center; gap: 0.35rem; font-weight: 700; color: #d70015; margin-bottom: 2px;">
                        <span>⚠️</span>
                        <span>任务阻断提示（需人工选择选项）</span>
                    </div>
                    <div style="color: #1d1d1f; font-weight: 600; padding-left: 1.25rem;">
                        👉 <strong>所需选项：</strong><span style="color: #b91c1c; background: rgba(255, 59, 48, 0.12); padding: 1px 6px; border-radius: 4px;">{action_req}</span>
                    </div>
                </div>
                """

            err_html = f"<div style='color:#ff3b30;font-size:0.78rem;margin-top:0.3rem;'>{task.error}</div>" if (task.error and not action_req) else ""
            now_ms = time.time() * 1000
            accumulated_sec = task.get_duration_seconds()

            # 简明数字进度 (如: 15/120 条 或 30/50 件)
            progress_text = getattr(task, 'progress_text', '')
            progress_html = f"""<span style="font-size:0.75rem;font-weight:700;color:#0071e3;background:rgba(0,113,227,0.08);border:1px solid rgba(0,113,227,0.18);padding:1px 7px;border-radius:6px;">⚡ {progress_text}</span>""" if progress_text else ""

            # 动态进度条
            progress_pct = getattr(task, 'progress_percent', 0.0)
            if task.total_progress > 0 and task.current_progress > 0:
                progress_pct = min(100.0, max(0.0, round(task.current_progress / task.total_progress * 100, 1)))
            progress_bar_html = f"""
            <div style="background: rgba(0,0,0,0.06); border-radius: 9999px; height: 5px; width: 100%; overflow: hidden; margin: 0.35rem 0 0.15rem 0;">
                <div style="background: linear-gradient(90deg, #007aff, #34c759); height: 100%; width: {progress_pct}%; border-radius: 9999px; transition: width 0.4s ease;"></div>
            </div>
            """ if (task.total_progress > 0) else ""

            task_items_html += f"""
            <div style="background:rgba(255,255,255,0.88);backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.95);border-radius:12px;padding:0.9rem;margin-bottom:0.8rem;box-shadow:0 2px 12px rgba(0,0,0,0.03);">
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:0.4rem;">
                    <span style="font-size:0.72rem;font-weight:700;color:#555;background:#eef1f6;padding:2px 6px;border-radius:4px;">{task.category}</span>
                    <div style="display:flex;align-items:center;gap:0.4rem;">
                        <span style="font-size:0.72rem;font-weight:600;padding:2px 8px;border-radius:9999px;{status_badge_style}">{status_text}</span>
                        {actions_btn_html}
                    </div>
                </div>
                <div style="font-size:0.88rem;font-weight:600;color:#1d1d1f;margin-bottom:0.35rem;word-break:break-all;">{task.name}</div>
                {action_banner_html}
                <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:0.2rem;">
                    <div class="vanguard-task-timer" data-render-time="{now_ms}" data-accumulated-sec="{accumulated_sec}" data-status="{task.status}" style="font-size:0.75rem;color:#888;">耗时: {task.get_duration_str()}</div>
                    {progress_html}
                </div>
                {progress_bar_html}
                {err_html}
                <details style="margin-top:0.35rem;font-size:0.75rem;">
                    <summary style="cursor:pointer;color:#007aff;font-weight:500;">查看实时日志 ({len(logs)}行)</summary>
                    <pre style="background:#1e1e1e;color:#f0f0f0;padding:0.6rem;border-radius:6px;font-size:0.7rem;max-height:160px;overflow-y:auto;margin-top:0.4rem;white-space:pre-wrap;word-break:break-all;">{log_html_safe}</pre>
                </details>
            </div>
            """

    has_action_needed = any(getattr(t, 'action_required', None) for t in tasks if t.status in ("RUNNING", "FAILED", "PAUSED"))
    pulse_anim = """<span style="width: 8px; height: 8px; border-radius: 50%; background: #ff3b30; box-shadow: 0 0 8px #ff3b30; display: inline-block; animation: vPulse 1.2s infinite;"></span>""" if has_action_needed else ("""<span style="width: 8px; height: 8px; border-radius: 50%; background: #34c759; box-shadow: 0 0 8px #34c759; display: inline-block; animation: vPulse 1.5s infinite;"></span>""" if active_count > 0 else """<span style="width: 8px; height: 8px; border-radius: 50%; background: #8e8e93; display: inline-block;"></span>""")
    action_warn_tag = """<span style="font-size: 0.72rem; font-weight: 700; color: #ffffff; background: #ff3b30; padding: 2px 7px; border-radius: 9999px;">需处理</span>""" if has_action_needed else ""

    drawer_script = f"""
    <script>
    (function() {{
        const parentDoc = window.parent.document;
        
        // 1. 创建或更新右上角常驻浮动触发胶囊
        let trigger = parentDoc.getElementById("vanguard-task-trigger");
        if (!trigger) {{
            trigger = parentDoc.createElement("div");
            trigger.id = "vanguard-task-trigger";
            trigger.style.cssText = "position: fixed; top: 1.2rem; right: 1.5rem; z-index: 999999; background: rgba(255, 255, 255, 0.92); backdrop-filter: blur(24px) saturate(140%); -webkit-backdrop-filter: blur(24px) saturate(140%); border: 1px solid rgba(255, 255, 255, 0.95); border-radius: 9999px; padding: 0.45rem 1.1rem; box-shadow: 0 4px 18px rgba(0,0,0,0.08), inset 0 1px 1px #fff; cursor: pointer; display: flex; align-items: center; gap: 0.55rem; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);";
            parentDoc.body.appendChild(trigger);
        }}
        trigger.innerHTML = `
            {pulse_anim}
            <span style="font-size: 0.86rem; font-weight: 600; color: #1d1d1f; letter-spacing: 0.01em;">任务看板</span>
            {action_warn_tag}
            <span style="font-size: 0.74rem; font-weight: 700; color: #ffffff; background: #1d1d1f; padding: 2px 7px; border-radius: 9999px;">{active_count}</span>
        `;
        trigger.onmouseover = () => {{ trigger.style.transform = "translateY(-2px) scale(1.03)"; trigger.style.boxShadow = "0 8px 24px rgba(0,0,0,0.12), inset 0 1px 1px #fff"; }};
        trigger.onmouseout = () => {{ trigger.style.transform = "translateY(0) scale(1)"; trigger.style.boxShadow = "0 4px 18px rgba(0,0,0,0.08), inset 0 1px 1px #fff"; }};

        // 2. 创建遮罩层与抽屉面板
        let overlay = parentDoc.getElementById("vanguard-task-overlay");
        if (!overlay) {{
            overlay = parentDoc.createElement("div");
            overlay.id = "vanguard-task-overlay";
            // 首次创建时初始化为隐藏；fragment 刷新时 overlay 已存在则保留其当前显示状态，
            // 防止 run_every=2s 重渲染时将已显示的遮罩层错误重置为 pointer-events:none
            overlay.style.cssText = "position: fixed; inset: 0; background: rgba(0, 0, 0, 0.2); backdrop-filter: blur(4px); -webkit-backdrop-filter: blur(4px); z-index: 9999998; opacity: 0; pointer-events: none; transition: opacity 0.35s cubic-bezier(0.16, 1, 0.3, 1);";
            parentDoc.body.appendChild(overlay);
        }}

        let drawer = parentDoc.getElementById("vanguard-task-drawer");
        const savedDrawerOpen = window.parent.sessionStorage.getItem('vanguard_task_drawer_open') === 'true';
        const wasOpen = (drawer ? (drawer.getAttribute("data-is-open") === "true") : false) || savedDrawerOpen;

        if (savedDrawerOpen) {{
            window.parent.sessionStorage.removeItem('vanguard_task_drawer_open');
        }}

        let scrollPos = 0;
        let openLogs = [];
        if (drawer) {{
            const contentDiv = drawer.querySelector('.vanguard-drawer-content');
            if (contentDiv) {{
                scrollPos = contentDiv.scrollTop;
                drawer.querySelectorAll('details').forEach((d, idx) => {{
                    if (d.open) openLogs.push(idx);
                }});
            }}
        }}

        if (!drawer) {{
            drawer = parentDoc.createElement("div");
            drawer.id = "vanguard-task-drawer";
            drawer.setAttribute("data-is-open", "false");
            drawer.style.cssText = "position: fixed; top: 0; right: 0; bottom: 0; width: 420px; max-width: 90vw; background: rgba(255, 255, 255, 0.90); backdrop-filter: blur(40px) saturate(160%); -webkit-backdrop-filter: blur(40px) saturate(160%); border-left: 1px solid rgba(255, 255, 255, 0.95); box-shadow: -10px 0 40px rgba(0, 0, 0, 0.1); z-index: 9999999; transform: translateX(100%); transition: transform 0.35s cubic-bezier(0.16, 1, 0.3, 1); display: flex; flex-direction: column; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;";
            parentDoc.body.appendChild(drawer);
        }}

        drawer.innerHTML = `
            <div style="padding: 1.2rem 1.4rem; border-bottom: 1px solid rgba(0,0,0,0.06); display: flex; align-items: center; justify-content: space-between;">
                <div>
                    <div style="font-size: 1.1rem; font-weight: 700; color: #1d1d1f;">任务运行看板</div>
                    <div style="font-size: 0.78rem; color: #888; margin-top: 2px;">共 {total_count} 条记录 · {active_count} 个运行中</div>
                </div>
                <div style="display: flex; align-items: center; gap: 0.5rem;">
                    <button id="vanguard-btn-drawer-refresh" style="background: rgba(0,0,0,0.05); border: none; border-radius: 6px; padding: 5px 10px; font-size: 0.78rem; font-weight: 600; color: #1d1d1f; cursor: pointer;">刷新</button>
                    <button id="vanguard-btn-drawer-close" style="width: 32px; height: 32px; border-radius: 50%; background: rgba(0,0,0,0.06); border: none; font-size: 1.1rem; color: #555; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">✕</button>
                </div>
            </div>
             <div class="vanguard-drawer-content" style="padding:1.2rem;flex:1;overflow-y:auto;">
                {task_items_html}
            </div>
        `;

        if (scrollPos > 0) {{
            const newContentDiv = drawer.querySelector('.vanguard-drawer-content');
            if (newContentDiv) newContentDiv.scrollTop = scrollPos;
        }}
        if (openLogs.length > 0) {{
            drawer.querySelectorAll('details').forEach((d, idx) => {{
                if (openLogs.includes(idx)) d.open = true;
            }});
        }}

        // ── 任务控制全局函数（直接触发 Streamlit 原生 Relay 按钮，突破 iframe sandbox 限制）──
        window.parent._vanguardHandleTaskAction = function(action, taskId) {{
            if (!action || !taskId) return;
            try {{
                window.parent.sessionStorage.setItem('vanguard_task_drawer_open', 'true');
            }} catch(e) {{}}
            
            const targetTag = action.toUpperCase() + '_TASK_' + taskId;
            const buttons = Array.from(parentDoc.querySelectorAll('button'));
            const match = buttons.find(b => (b.textContent || '').trim() === targetTag);
            if (match) {{
                match.click();
            }}
        }};

        // 绑定展开与关闭动作
        const openDrawer = () => {{
            drawer.setAttribute("data-is-open", "true");
            drawer.style.transform = "translateX(0)";
            overlay.style.opacity = "1";
            overlay.style.pointerEvents = "auto";
        }};
        const closeDrawer = () => {{
            try {{
                window.parent.sessionStorage.removeItem('vanguard_task_drawer_open');
            }} catch(e) {{}}
            drawer.setAttribute("data-is-open", "false");
            drawer.style.transform = "translateX(100%)";
            overlay.style.opacity = "0";
            overlay.style.pointerEvents = "none";
        }};

        if (wasOpen) {{
            drawer.setAttribute("data-is-open", "true");
            drawer.style.transform = "translateX(0)";
            overlay.style.opacity = "1";
            overlay.style.pointerEvents = "auto";
        }}

        trigger.onclick = (e) => {{ e.preventDefault(); openDrawer(); }};
        overlay.onclick = () => {{ closeDrawer(); }};
        const closeBtn = drawer.querySelector("#vanguard-btn-drawer-close");
        if (closeBtn) closeBtn.onclick = () => {{ closeDrawer(); }};

        // 刷新看板：直接刷新页面
        const refreshBtn = drawer.querySelector("#vanguard-btn-drawer-refresh");
        if (refreshBtn) {{
            refreshBtn.onclick = (e) => {{
                e.preventDefault();
                e.stopPropagation();
                parentDoc.location.reload();
            }};
        }}

        // 客户端高精度秒级计时器 (仅对 RUNNING 状态动态计时，PAUSED 状态严格冻结不增加)
        if (window.parent._vanguardTimerTicker) {{
            clearInterval(window.parent._vanguardTimerTicker);
        }}
        window.parent._vanguardTimerTicker = setInterval(() => {{
            const timerEls = parentDoc.querySelectorAll('.vanguard-task-timer[data-status="RUNNING"]');
            const now = Date.now();
            timerEls.forEach(el => {{
                const renderTs = parseFloat(el.getAttribute('data-render-time') || '0');
                const baseSec = parseFloat(el.getAttribute('data-accumulated-sec') || '0');
                if (renderTs > 0) {{
                    const diffSec = baseSec + Math.max(0, Math.floor((now - renderTs) / 1000));
                    const m = Math.floor(diffSec / 60).toString().padStart(2, '0');
                    const s = (diffSec % 60).toString().padStart(2, '0');
                    el.textContent = `耗时: ${{m}}:${{s}}`;
                }}
            }});
        }}, 1000);

        // 隐藏任务中继控制按钮
        parentDoc.querySelectorAll('button').forEach(btn => {{
            const txt = (btn.textContent || '').trim();
            if (txt.startsWith('KILL_TASK_') || txt.startsWith('PAUSE_TASK_') || txt.startsWith('RESUME_TASK_')) {{
                const el = btn.closest('[data-testid="stElementContainer"]') || btn;
                el.style.position = 'fixed';
                el.style.top = '-9999px';
                el.style.left = '-9999px';
                el.style.opacity = '0';
                el.style.height = '0';
                el.style.pointerEvents = 'none';
            }}
        }});
    }})();
    </script>
    """
    components.html(drawer_script, height=0, width=0)

    # 任务控制原生中继按钮（与抽屉中彩色胶囊按钮一一绑定，通过 WebSocket 即时处理）
    for t in tasks:
        if t.status in ("RUNNING", "PAUSED"):
            if st.button(f"PAUSE_TASK_{t.task_id}", key=f"btn_pause_relay_{t.task_id}"):
                global_task_manager.pause_task(t.task_id)
                st.rerun()
            if st.button(f"RESUME_TASK_{t.task_id}", key=f"btn_resume_relay_{t.task_id}"):
                global_task_manager.resume_task(t.task_id)
                st.rerun()
            if st.button(f"KILL_TASK_{t.task_id}", key=f"btn_kill_relay_{t.task_id}"):
                global_task_manager.cancel_task(t.task_id)
                st.rerun()



def render_dianxiaomi_stock_module(key_prefix: str = "main"):
    with st.container(border=True):
        st.markdown("### 店小秘在线商品库存修改")

        shops = load_cached_shops()
        shop_options = [f"[{s['shortName']}] {s['text']}" for s in shops]
        shop_codes = [s['code'] for s in shops]

        # 默认自动匹配并选中 k店 与 o店
        default_indices = []
        for i, s in enumerate(shops):
            if s.get('firstLetter') in ['k', 'o'] or 'k店' in s.get('shortName', '') or 'o店' in s.get('shortName', ''):
                default_indices.append(i)
        if not default_indices and len(shop_options) > 0:
            default_indices = [0, 1] if len(shop_options) > 1 else [0]

        selected_indices = st.multiselect(
            "目标店铺 (支持多选，系统将同时开启多个网页并发修改)",
            range(len(shop_options)),
            default=default_indices,
            format_func=lambda i: shop_options[i],
            key=f"{key_prefix}_dxm_shop_idx"
        )

        col_num, col_btn = st.columns([1, 1])
        with col_num:
            target_stock_val = st.number_input(
                "修改库存为",
                min_value=0,
                max_value=999999,
                value=0,
                step=1,
                key=f"{key_prefix}_dxm_stock_val"
            )

        with col_btn:
            st.markdown("<div style='height: 1.72rem;'></div>", unsafe_allow_html=True)
            st.markdown(f"""
            <div class="compact-action-outer compact-primary" id="visual-btn-dxm-{key_prefix}" style="margin-top: 0 !important; margin-bottom: 0 !important; min-height: 60px !important;">
                <div class="compact-action-inner" style="padding: 0.4rem 1.2rem !important;">
                    <div class="compact-btn-left">
                        <div class="compact-btn-icon">⚡</div>
                        <div class="compact-btn-title" style="font-size: 0.95rem !important;">批量修改店铺库存</div>
                    </div>
                    <div class="compact-island">
                        <span class="compact-island-text">BATCH UPDATE</span>
                        <span class="compact-island-circle">↗</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            btn_dxm_run = st.button(f"RELAY_BTN_DXM_{key_prefix.upper()}", key=f"btn_dxm_relay_{key_prefix}")

        if btn_dxm_run:
            if not selected_indices:
                st.warning("请至少选择一个目标店铺。")
            else:
                for idx, s_idx in enumerate(selected_indices):
                    shop_code = shop_codes[s_idx]
                    shop_name = shop_options[s_idx]
                    task_id = f"dxm_{int(time.time())}_{idx}"
                    global_task_manager.submit_task(
                        task_id=task_id,
                        name=f"店小秘改库存 - {shop_name[:18]}",
                        target_fn=run_dianxiaomi_stock_update,
                        kwargs={
                            "shop_code": shop_code,
                            "target_stock": int(target_stock_val),
                            "user_data_dir": DIANXIAOMI_SESSION_DIR,
                            "worker_id": str(idx),
                            "headless": False
                        },
                        category="店小秘"
                    )
                st.toast(f"已同时在后台启动 {len(selected_indices)} 个店铺的库存修改任务", icon=None)
                st.rerun()

def execute_pipeline_task(
    route: str,
    sw_auto_export: bool,
    days: int,
    hours: int,
    search_val: str,
    current_excel_path: str,
    active_template_type: str,
    sw_cpf_rename: bool,
    sw_cpf_merge: bool,
    sw_mabang_update: bool,
    sw_dsers_clean: bool,
    sw_dsers_cpf_check: bool,
    sw_dsers_cpf_merge: bool,
    sw_dsers_mabang: bool,
    sw_dsers_import: bool,
    sw_dsers_rename: bool,
    use_vault: bool,
    vault_file_choice: str,
    sku_filter: str = "code",
    mabang_update_mode: str = "both",
    dsers_update_fields: list = None,
    progress_callback = None,
    task_info = None
):
    def log(msg):
        print(msg, flush=True)
        if task_info:
            task_info.check_pause()
        if progress_callback:
            progress_callback(msg)

    pipeline_name = "CPF 订单核对" if route == "A" else "DSERS 批量下单"
    log(f"🚀 开始执行流水线：{pipeline_name}")

    # 清洗残余单例锁
    for lock_file in glob.glob(os.path.join(SESSIONS_DIR, "*", "Singleton*")):
        try: os.remove(lock_file)
        except: pass

    if task_info: task_info.check_pause()

    # 下单模板清洗与映射
    if active_template_type == "下单模板.xlsx":
        is_rename_active = (route == "B" and sw_dsers_rename)
        clean_order_template_to_script(ORDER_TEMPLATE, SCRIPT_TEMPLATE, route, sw_dsers_rename=is_rename_active)
        current_excel_path = SCRIPT_TEMPLATE
        try:
            import pandas as pd
            if os.path.exists(SCRIPT_TEMPLATE):
                df_init = pd.read_excel(SCRIPT_TEMPLATE)
                if len(df_init) > 0 and task_info:
                    task_info.total_progress = len(df_init)
                    task_info.progress_text = f"共 {len(df_init)} 条订单"
        except Exception:
            pass

    # 阶段 1：马帮 ERP 导出 (原生运行，实时支持暂停与继续)
    if sw_auto_export:
        if task_info: task_info.check_pause()
        log("[阶段 1/4] 正在调取马帮 ERP 订单数据..." if route == "A" else "[阶段 1/5] 正在调取马帮 ERP 订单数据...")
        if route == "A":
            task_info.run_async(run_mabang_cpf_export(
                MABANG_SESSION_DIR,
                days=days,
                hours=hours,
                customer_id=search_val,
                sku_filter=sku_filter,
                headless=False,
                progress_callback=log,
                task_info=task_info
            ))
        else:
            task_info.run_async(run_mabang_dsers_export(
                MABANG_SESSION_DIR,
                days=days,
                hours=hours,
                sku_val=search_val,
                headless=False,
                progress_callback=log,
                task_info=task_info
            ))
        log("[阶段 1] 马帮订单数据提取完成。")
        current_excel_path = SCRIPT_TEMPLATE

    # 阶段 2：管线专属执行
    if route == "A":
        if sw_cpf_rename:
            if task_info: task_info.check_pause()
            log("[阶段 2/4] 正在进行 Telegram CPF 姓名查询...")
            task_info.run_async(run_cpf_query(current_excel_path, TELEGRAM_SESSION_DIR, False, lambda m: log(f"[TG 实时] {m}"), task_info=task_info))
            log("[阶段 2] Telegram 查名与校验完成。")
            
            if active_template_type == "下单模板.xlsx":
                log("[同步] 正在同步最新姓名至下单模板 E 列...")
                sync_cpf_results_to_order_template(SCRIPT_TEMPLATE, ORDER_TEMPLATE)
                log("[同步] 下单模板 E 列已更新完成。")
                
        if sw_cpf_merge:
            if task_info: task_info.check_pause()
            log("[阶段 3/4] 独立同步：回填最新姓名至 DSers 导入模板...")
            merge_cpf_results(progress_callback=log)
            log("[阶段 3] 独立同步完成，DSers 订单姓名已更新。")

        if sw_mabang_update:
            if task_info: task_info.check_pause()
            log(f"[阶段 4/4] 正在同步数据至马帮 ERP (模式: {mabang_update_mode})...")
            task_info.run_async(run_mabang_batch_update(current_excel_path, MABANG_SESSION_DIR, update_mode=mabang_update_mode, headless=False, progress_callback=log, task_info=task_info))
            log("[阶段 4] 马帮 ERP 数据同步完成。")

    else: # Route B
        if sw_dsers_clean:
            if task_info: task_info.check_pause()
            log("[阶段 2/5] 正在清理数据字段并映射格式...")
            run_dsers_clean_and_map(progress_callback=log)
            log("[阶段 2] 数据清理与格式映射完成。")
                
        if sw_dsers_cpf_check:
            if task_info: task_info.check_pause()
            log("[阶段 3/5] 姓名核对 1/2: 提取 DSers 订单至 CPF 模板...")
            if use_vault and vault_file_choice == "脚本模板.xlsx":
                log("[阶段 3/5] 姓名核对 1/2: (已从脚本模板继续，跳过桥接提取)")
            else:
                export_to_cpf_template(progress_callback=log)
                
            log("[阶段 3/5] 姓名核对 2/2: 正在通过 Telegram 进行姓名校对...")
            task_info.run_async(run_cpf_query(SCRIPT_TEMPLATE, TELEGRAM_SESSION_DIR, False, lambda m: log(f"[TG 实时] {m}"), task_info=task_info))
            
            if active_template_type == "下单模板.xlsx":
                log("[同步] 正在同步最新姓名至下单模板 E 列...")
                sync_cpf_results_to_order_template(SCRIPT_TEMPLATE, ORDER_TEMPLATE)
                log("[同步] 下单模板 E 列已更新完成。")
                
        if sw_dsers_cpf_merge:
            if task_info: task_info.check_pause()
            log("[阶段 3.5/5] 正在同步真实姓名至 DSers 导入模板...")
            merge_cpf_results(progress_callback=log)
            log("[阶段 3.5] 姓名同步完成，所有 DSers 订单姓名已更新。")
            
        if sw_dsers_mabang:
            if task_info: task_info.check_pause()
            log(f"[阶段 4/5] 正在同步真实姓名至马帮 ERP (模式: {mabang_update_mode})...")
            task_info.run_async(run_mabang_batch_update(SCRIPT_TEMPLATE, MABANG_SESSION_DIR, update_mode=mabang_update_mode, headless=False, progress_callback=log, task_info=task_info))
            log("[阶段 4] 马帮 ERP 数据同步完成。")
                
        if sw_dsers_import:
            if task_info: task_info.check_pause()
            log("[阶段 5/5] 正在向 DSers 批量创建与推送订单...")
            task_info.run_async(run_dsers_import(DSERS_IMPORT_CSV, DSERS_SESSION_DIR, headless=False, progress_callback=log, task_info=task_info))
            log("[阶段 5] 批量上传并创建 DSers 订单成功。")

        if sw_dsers_rename:
            if task_info: task_info.check_pause()
            active_fields_str = "全部" if not dsers_update_fields else ", ".join(dsers_update_fields)
            log(f"[阶段 独立] 正在处理 DSers 网页端订单自动修改 (目标字段: {active_fields_str})...")
            task_info.run_async(run_dsers_rename(current_excel_path, DSERS_SESSION_DIR, False, lambda m: log(f"[DSers 实时] {m}"), task_info=task_info, update_fields=dsers_update_fields))
            log("[阶段 独立] DSers 订单网页修改处理完成。")

    log(f"🎉 {pipeline_name} 流水线全部步骤执行完毕！")


# COMMON CSS (High-end Minimalist, High Transparency Glassmorphism)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=SF+Pro+Display:wght@300;400;500;600&display=swap');

    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif, "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol" !important;
    }

    /* =========================================
       苹果原生级：彻底统一规范所有输入框与下拉选择框（消除任何内胆断层与多重重影）
       ========================================= */
    /* 1. 最外层完整容器：单一完整圆角白底胶囊（涵盖文本、数字加减按钮全区、下拉列表） */
    [data-testid="stTextInput"] div[data-baseweb="input"],
    [data-testid="stNumberInput"] > div:has(input),
    [data-testid="stNumberInput"] > div[class*="eaba2yi0"],
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    [data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
        background: rgba(255, 255, 255, 0.95) !important;
        background-color: rgba(255, 255, 255, 0.95) !important;
        border: 1px solid rgba(0, 0, 0, 0.14) !important;
        border-radius: 10px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03) !important;
        overflow: hidden !important;
    }

    /* 2. 强力清除所有内部子元素（数字输入框内胆、按钮容器、select 外壳等）的自带边框、圆角、背景与阴影 */
    [data-testid="stNumberInput"] div[data-baseweb="input"],
    [data-testid="stNumberInput"] div[data-baseweb="base-input"],
    [data-testid="stNumberInput"] div[class*="eaba2yi1"],
    [data-testid="stTextInput"] div[data-baseweb="base-input"],
    [data-testid="stSelectbox"] div[data-baseweb="select"],
    [data-testid="stMultiSelect"] div[data-baseweb="select"],
    div[data-baseweb="base-input"] {
        border: none !important;
        box-shadow: none !important;
        background: transparent !important;
        background-color: transparent !important;
        border-radius: 0 !important;
        outline: none !important;
    }

    /* 3. 内部原生 input 文本：全透明无边框无圆角 */
    [data-testid="stNumberInput"] input,
    [data-testid="stTextInput"] input,
    div[data-baseweb="input"] input,
    div[data-baseweb="base-input"] input {
        background: transparent !important;
        background-color: transparent !important;
        border: none !important;
        border-radius: 0 !important;
        box-shadow: none !important;
        outline: none !important;
        color: #1d1d1f !important;
        -webkit-text-fill-color: #1d1d1f !important;
        font-size: 0.95rem !important;
        padding-left: 0.8rem !important;
    }

    /* 4. 数字输入框的加减按钮：平整通透，无独立背景与多余边框 */
    [data-testid="stNumberInput"] button {
        background: transparent !important;
        background-color: transparent !important;
        border: none !important;
        border-radius: 0 !important;
        box-shadow: none !important;
        color: #1d1d1f !important;
        width: 32px !important;
        height: 100% !important;
    }
    [data-testid="stNumberInput"] button:hover {
        background: rgba(0, 0, 0, 0.05) !important;
        background-color: rgba(0, 0, 0, 0.05) !important;
    }

    /* 4. 修复开关（Toggle）颜色为苹果官方极简石墨黑 / 经典优雅微晶质感，彻底移除廉价刺眼蓝 */
    div[data-testid="stToggle"] label span[role="checkbox"][aria-checked="true"],
    div[data-testid="stCheckbox"] label span[role="checkbox"][aria-checked="true"],
    div[data-baseweb="toggle"] div[aria-checked="true"],
    [data-testid="stToggle"] div[role="checkbox"][aria-checked="true"] {
        background-color: #1d1d1f !important;
        background: #1d1d1f !important;
    }

    /* 下拉选择弹出层背景与文字颜色 */
    div[data-baseweb="popover"],
    ul[data-baseweb="menu"],
    li[data-baseweb="menu-item"] {
        background-color: #ffffff !important;
        background: #ffffff !important;
        color: #1d1d1f !important;
    }
    li[data-baseweb="menu-item"]:hover {
        background-color: #f2f2f7 !important;
    }

    /* 所有标签文字颜色强制为经典深深灰 (#1d1d1f) */
    label, [data-testid="stWidgetLabel"] p, [data-testid="stMarkdownContainer"] p {
        color: #1d1d1f !important;
    }

    /* 彻底隐藏 Streamlit 原生侧边栏与顶栏，杜绝白线与残留交互 */
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapsedControl"],
    section[data-testid="stSidebar"],
    [data-testid="stSidebar"],
    header {
        display: none !important;
        visibility: hidden !important;
        width: 0 !important;
        height: 0 !important;
        border: none !important;
        box-shadow: none !important;
    }

    [data-testid="stAppViewContainer"] {
        flex-direction: row !important;
    }

    /* 多选框选中的标签：具有明显对比度的精致灰白毛玻璃芯片（Apple Chip 风格） */
    div[data-baseweb="tag"],
    span[data-baseweb="tag"],
    [data-testid="stMultiSelect"] div[data-baseweb="tag"],
    [data-testid="stMultiSelect"] span[data-baseweb="tag"] {
        background: #eef1f6 !important;
        border: 1px solid #d4d9e2 !important;
        border-radius: 6px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06) !important;
        padding: 3px 8px !important;
        margin: 2px !important;
    }
    div[data-baseweb="tag"] span,
    div[data-baseweb="tag"] div,
    [data-testid="stMultiSelect"] span,
    [data-testid="stMultiSelect"] div {
        color: #1d1d1f !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
    }
    div[data-baseweb="tag"] svg {
        fill: #666666 !important;
        color: #666666 !important;
    }
    div[data-baseweb="tag"] svg:hover {
        fill: #1d1d1f !important;
    }

    /* =========================================
       VISION PRO 级：空间计算 3D 极致动态引擎 (Spatial Computing Engine)
       ========================================= */

    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: transparent !important;
    }
    
    /* 沸腾的极光能量场背景 */
    .stApp::before {
        content: "";
        position: fixed;
        top: -50%; left: -50%;
        width: 200%; height: 200%;
        background: 
            radial-gradient(circle at 20% 30%, rgba(255, 107, 107, 0.4), transparent 60%),
            radial-gradient(circle at 80% 80%, rgba(72, 219, 251, 0.4), transparent 60%),
            radial-gradient(circle at 20% 80%, rgba(255, 159, 67, 0.4), transparent 60%),
            radial-gradient(circle at 80% 20%, rgba(84, 160, 255, 0.4), transparent 60%);
        animation: energyField 12s cubic-bezier(0.4, 0, 0.2, 1) infinite alternate;
        z-index: -2;
        filter: blur(80px);
    }
    
    .stApp::after {
        content: "";
        position: fixed;
        inset: 0;
        background: #f4f5f7;
        z-index: -3;
    }

    @keyframes energyField {
        0% { transform: scale(1) rotate(0deg); }
        100% { transform: scale(1.3) rotate(45deg); }
    }

    /* 恢复 liquidReveal 给按钮使用 */
    @keyframes liquidReveal {
        0% { opacity: 0; transform: translateY(40px) scale(0.98); filter: blur(12px); }
        100% { opacity: 1; transform: translateY(0) scale(1); filter: blur(0); }
    }

    /* 夸张的 3D 弹簧入场 (Hyper-Spring Reveal) */
    @keyframes hyperSpringReveal {
        0% { 
            opacity: 0; 
            transform: perspective(2000px) translateY(150px) translateZ(-300px) rotateX(20deg) scale(0.8); 
            filter: blur(30px); 
        }
        100% { 
            opacity: 1; 
            transform: perspective(2000px) translateY(0) translateZ(0) rotateX(0deg) scale(1); 
            filter: blur(0); 
        }
    }

    /* 极致高透毛玻璃容器 - 3D 悬浮态 */
    [data-testid="stVerticalBlockBorderWrapper"],
    div[data-testid="stVerticalBlock"] > div[style*="border"],
    div[class*="st-emotion-cache"][style*="border"] {
        background: rgba(255, 255, 255, 0.2) !important;
        backdrop-filter: blur(60px) saturate(250%) !important;
        -webkit-backdrop-filter: blur(60px) saturate(250%) !important;
        border: 1px solid rgba(255, 255, 255, 0.7) !important;
        border-radius: 2.5rem !important; 
        box-shadow: 
            0 10px 30px rgba(0,0,0,0.05),
            inset 0 2px 0 0 rgba(255,255,255,0.8),
            inset 0 0 20px rgba(255,255,255,0.5) !important;
        padding: 3.5rem !important;
        
        /* 核心 3D 交互配置 */
        transform-style: preserve-3d;
        transition: transform 0.6s cubic-bezier(0.34, 1.56, 0.64, 1), 
                    box-shadow 0.6s cubic-bezier(0.34, 1.56, 0.64, 1),
                    background 0.4s !important;
        
        /* 夸张弹簧入场 */
        animation: hyperSpringReveal 1.2s cubic-bezier(0.2, 0.8, 0.2, 1.2) forwards;
        position: relative;
        overflow: hidden;
    }
    
    /* 容器上的无限反射光效 (Infinite Shimmer Sweep) */
    [data-testid="stVerticalBlockBorderWrapper"]::after,
    div[data-testid="stVerticalBlock"] > div[style*="border"]::after {
        content: "";
        position: absolute;
        top: 0; left: -150%;
        width: 50%; height: 100%;
        background: linear-gradient(to right, rgba(255,255,255,0) 0%, rgba(255,255,255,0.6) 50%, rgba(255,255,255,0) 100%);
        transform: skewX(-25deg);
        animation: glassSweep 4s ease-in-out infinite;
        pointer-events: none;
    }

    @keyframes glassSweep {
        0% { left: -150%; }
        50% { left: 200%; }
        100% { left: 200%; }
    }
    
    /* 暴力 3D 翻转破雪交互 (Violent 3D Hover) */
    [data-testid="stVerticalBlockBorderWrapper"]:hover,
    div[data-testid="stVerticalBlock"] > div[style*="border"]:hover {
        background: rgba(255, 255, 255, 0.4) !important;
        /* 直接让卡片飞出来并发生 3D 偏转 */
        transform: perspective(2000px) translateZ(80px) translateY(-10px) rotateX(4deg) rotateY(-2deg) scale(1.02);
        box-shadow: 
            -20px 40px 80px -10px rgba(0,0,0,0.15), 
            inset 0 2px 0 0 rgba(255,255,255,1), 
            inset 0 0 60px rgba(255,255,255,0.8) !important;
        border: 1px solid rgba(255, 255, 255, 1) !important;
    }

    h1, h2, h3, h4, h5 {
        color: #1d1d1f !important; /* Apple typography color */
        font-weight: 600 !important;
        letter-spacing: -0.02em !important;
    }
    
    p, span, div {
        color: #424245 !important;
    }

    /* 调整顶部空白 */
    .block-container {
        padding-top: 4rem !important;
        max-width: 1200px !important;
    }

    /* ====================
       GLOBAL ULTRA-CLEAR ETHEREAL GLASS CARDS (全站通用近乎全透玻璃材质与 0.8s 丝滑联动)
       ==================== */
    .premium-card-outer {
        padding: 0.5rem;
        border-radius: 2.5rem;
        background: linear-gradient(135deg, rgba(255,255,255,0.15) 0%, rgba(255,255,255,0.02) 100%);
        border: 1px solid rgba(255, 255, 255, 0.4);
        border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        border-right: 1px solid rgba(255, 255, 255, 0.1);
        box-shadow: 0 30px 60px rgba(0,0,0,0.05);
        transition: transform 0.8s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.8s cubic-bezier(0.16, 1, 0.3, 1), background 0.8s cubic-bezier(0.16, 1, 0.3, 1), background-color 0.8s cubic-bezier(0.16, 1, 0.3, 1), border 0.8s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.8s cubic-bezier(0.16, 1, 0.3, 1), opacity 0.8s cubic-bezier(0.16, 1, 0.3, 1) !important;
        position: relative;
        overflow: hidden;
        height: 460px;
        will-change: transform;
    }
    
    .premium-card-inner {
        background: rgba(255, 255, 255, 0.05);
        border-radius: 2.125rem;
        height: 100%;
        padding: 2.6rem 2.8rem 2.4rem 2.8rem;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(50px) saturate(120%);
        -webkit-backdrop-filter: blur(50px) saturate(120%);
        transition: transform 0.8s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.8s cubic-bezier(0.16, 1, 0.3, 1), background 0.8s cubic-bezier(0.16, 1, 0.3, 1), background-color 0.8s cubic-bezier(0.16, 1, 0.3, 1), border 0.8s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.8s cubic-bezier(0.16, 1, 0.3, 1), opacity 0.8s cubic-bezier(0.16, 1, 0.3, 1) !important;
        position: relative;
        z-index: 10;
    }
    
    /* 径向网格渐变背景 */
    .premium-card-outer::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle at 50% 0%, rgba(255,255,255,0.8), transparent 60%);
        opacity: 0.2;
        transition: opacity 0.8s ease;
        z-index: 1;
        pointer-events: none;
    }
    
    /* Kinetic Hover Physics */
    .premium-card-outer:hover {
        transform: translateY(-12px) scale(1.02);
        border: 1px solid rgba(255, 255, 255, 1);
        background: linear-gradient(135deg, rgba(255,255,255,0.6) 0%, rgba(255,255,255,0.1) 100%);
        box-shadow: 0 40px 80px rgba(0,0,0,0.1), 0 0 0 1px rgba(255,255,255,0.5);
        cursor: pointer;
    }
    
    .premium-card-outer:hover .premium-card-inner {
        background: rgba(255, 255, 255, 0.4);
        box-shadow: inset 0 1px 1px rgba(255, 255, 255, 1);
    }
    
    .premium-card-outer:hover::before {
        opacity: 1;
    }
    
    .premium-card-outer:active {
        transform: scale(0.97);
        transition: all 0.15s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }
    
    /* Typography inside cards */
    .card-icon-wrapper {
        width: 72px;
        height: 72px;
        border-radius: 1.4rem;
        background: linear-gradient(135deg, rgba(255,255,255,0.8) 0%, rgba(255,255,255,0.1) 100%);
        border: 1px solid rgba(255, 255, 255, 1);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 2.3rem;
        margin-bottom: 1.2rem;
        box-shadow: inset 0 2px 4px rgba(255,255,255,0.5), 0 10px 20px rgba(0,0,0,0.1);
        transition: all 0.8s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    
    .card-title {
        font-size: 3.4rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 0.8rem;
        color: #1d1d1f !important;
        background: linear-gradient(135deg, #000000 0%, #434345 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        line-height: 1.1;
    }
    
    .card-desc {
        font-size: 1.3rem;
        color: #333333 !important;
        line-height: 1.65;
        font-weight: 600;
        letter-spacing: 0.01em;
    }
    
    /* Button-in-Button Architecture (Island CTA) */
    .island-btn {
        display: inline-flex !important;
        align-items: center !important;
        background: rgba(255, 255, 255, 0.5) !important;
        border: 1px solid rgba(255, 255, 255, 0.8) !important;
        border-radius: 9999px !important;
        padding: 0.6rem 0.6rem 0.6rem 1.6rem !important;
        margin-top: 1.4rem;
        width: fit-content !important;
        transition: all 0.5s cubic-bezier(0.16, 1, 0.3, 1) !important;
        box-shadow: 0 4px 15px rgba(0,0,0,0.03) !important;
    }
    
    .island-btn-text {
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        color: #222222 !important;
        margin-right: 1.8rem !important;
        letter-spacing: 0.06em !important;
        text-transform: uppercase !important;
        transition: color 0.5s ease !important;
    }
    
    .island-btn-icon {
        width: 44px !important;
        height: 44px !important;
        border-radius: 50% !important;
        background: #1d1d1f !important;
        color: #ffffff !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 1.4rem !important;
        font-weight: 600 !important;
        transition: all 0.5s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    
    /* 联动 Hover Physics */
    .premium-card-outer:hover .card-icon-wrapper {
        transform: scale(1.05) translateY(-4px) !important;
        background: linear-gradient(135deg, rgba(255,255,255,1) 0%, rgba(255,255,255,0.4) 100%) !important;
    }
    .premium-card-outer:hover .island-btn {
        background: rgba(255, 255, 255, 0.85) !important;
        border-color: rgba(255, 255, 255, 1) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 20px rgba(0,0,0,0.06) !important;
    }
    .premium-card-outer:hover .island-btn-text {
        color: #000000 !important;
    }
    .premium-card-outer:hover .island-btn-icon {
        transform: translateX(4px) scale(1.05) !important;
        background: #000000 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2) !important;
    }
</style>
""", unsafe_allow_html=True)


import streamlit.components.v1 as components
import os
import urllib.parse

try:
    bg_html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "3d_background", "dist", "index.html")
    if os.path.exists(bg_html_path):
        with open(bg_html_path, "r", encoding="utf-8") as f:
            html_data = f.read()
        
        html_data_encoded = urllib.parse.quote(html_data)
        
        inject_script = f'''
        <script>
            const parentDoc = window.parent.document;
            if (!parentDoc.getElementById("vanguard-3d-bg")) {{
                const iframe = parentDoc.createElement("iframe");
                iframe.id = "vanguard-3d-bg";
                iframe.style.position = "fixed";
                iframe.style.top = "0";
                iframe.style.left = "0";
                iframe.style.width = "100vw";
                iframe.style.height = "100vh";
                iframe.style.zIndex = "-1";
                iframe.style.border = "none";
                iframe.style.pointerEvents = "none";
                iframe.srcdoc = decodeURIComponent("{html_data_encoded}");
                parentDoc.body.appendChild(iframe);
            }}
        </script>
        '''
        components.html(inject_script, height=0, width=0)
except Exception as e:
    pass

# 侧边栏全局后台任务中枢
render_live_task_hub()

if st.session_state.route is None:
    # ==========================
    # 首页大屏视觉效果 (Home Screen) - Agency Tier
    # ==========================
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');
        
        /* 强制全屏容器使用特定字体，并去除强制深黑背景 */
        [data-testid="stAppViewContainer"], .block-container {
            font-family: 'Plus Jakarta Sans', sans-serif !important;
        }
        
        .block-container {
            padding-top: 8rem !important;
            max-width: 1400px !important;
        }

        /* 完美隐藏原生按钮，不影响 DOM 布局 */
        [data-testid="stButton"] {
            position: absolute !important;
            opacity: 0 !important;
            height: 0 !important;
            width: 0 !important;
            overflow: hidden !important;
            pointer-events: none !important;
        }
    </style>
    
    <div style="text-align: center; margin-bottom: 3.5rem; animation: liquidReveal 1s cubic-bezier(0.16, 1, 0.3, 1) forwards; opacity: 0; transform: translateY(20px);">
        <div style="display:inline-block; padding: 0.6rem 1.4rem; border-radius: 9999px; border: 1px solid rgba(0,0,0,0.1); background: rgba(255,255,255,0.4); font-size: 0.9rem; text-transform: uppercase; letter-spacing: 0.2em; margin-bottom: 1.8rem; color: #444444; backdrop-filter: blur(10px); font-weight: 600;">
            Vanguard Systems
        </div>
        <h1 style="font-size: 5.2rem; font-weight: 800; letter-spacing: -0.04em; background: linear-gradient(180deg, #1d1d1f 0%, #555555 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 1.2rem; line-height: 1;">Pipeline Studio</h1>
        <p style="font-size: 1.6rem; font-weight: 500; color: #555555; letter-spacing: 0.02em;">Select your intelligence distribution route.</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")
    
    with col1:
        st.markdown("""
        <div class="premium-card-outer" id="home-card-cpf" style="animation: liquidReveal 1s cubic-bezier(0.16, 1, 0.3, 1) 0.1s forwards; opacity: 0; cursor: pointer;">
            <div class="premium-card-inner">
                <div>
                    <div class="card-icon-wrapper">A</div>
                    <div class="card-title">CPF</div>
                    <div class="card-desc">
                        自动调取 Telegram CPF 接口<br>
                        核对并纠正错误的买家姓名<br>
                        同步更新至马帮和下单表格
                    </div>
                </div>
                <div class="island-btn">
                    <span class="island-btn-text">进入操作</span>
                    <span class="island-btn-icon">↗</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("CPF_ROUTE_HIDDEN_123", key="btn_home_cpf_relay"):
            st.session_state.route = "A"
            st.rerun()
            
    with col2:
        st.markdown("""
        <div class="premium-card-outer" id="home-card-dsers" style="animation: liquidReveal 1s cubic-bezier(0.16, 1, 0.3, 1) 0.2s forwards; opacity: 0; cursor: pointer;">
            <div class="premium-card-inner">
                <div>
                    <div class="card-icon-wrapper">B</div>
                    <div class="card-title">DSERS</div>
                    <div class="card-desc">
                        自动转换马帮订单为 DSERS 格式<br>
                        支持自动校验买家 CPF 并改名<br>
                        批量上传至 DSERS 后台创建订单
                    </div>
                </div>
                <div class="island-btn">
                    <span class="island-btn-text">进入操作</span>
                    <span class="island-btn-icon">↗</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("DSERS_ROUTE_HIDDEN_123", key="btn_home_dsers_relay"):
            st.session_state.route = "B"
            st.rerun()

    # ==========================
    # 彻底解决点击问题：JS 精准中继穿透引擎
    # ==========================
    import streamlit.components.v1 as components
    components.html("""
    <script>
        const parentDoc = window.parent.document;
        let attempts = 0;
        const bindClicks = setInterval(() => {
            attempts++;
            
            const cardCpf = parentDoc.getElementById('home-card-cpf');
            const cardDsers = parentDoc.getElementById('home-card-dsers');
            
            const findRelayBtn = (textTag) => {
                const allBtns = parentDoc.querySelectorAll('button');
                for (let btn of allBtns) {
                    if (btn.textContent && btn.textContent.includes(textTag)) {
                        let container = btn.closest('.element-container') || btn.closest('[data-testid="stButton"]') || btn;
                        container.style.setProperty('display', 'none', 'important');
                        container.style.setProperty('position', 'absolute', 'important');
                        container.style.setProperty('opacity', '0', 'important');
                        container.style.setProperty('height', '0', 'important');
                        container.style.setProperty('pointer-events', 'none', 'important');
                        return btn;
                    }
                }
                return null;
            };
            
            const bindHomeAction = (cardId, textTag) => {
                const card = parentDoc.getElementById(cardId);
                if (card) {
                    card.style.cursor = 'pointer';
                    card.onclick = (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        const liveBtn = findRelayBtn(textTag);
                        if (liveBtn) {
                            liveBtn.click();
                        }
                    };
                }
            };
            
            bindHomeAction('home-card-cpf', 'CPF_ROUTE_HIDDEN_123');
            bindHomeAction('home-card-dsers', 'DSERS_ROUTE_HIDDEN_123');
            
            const btnCpf = findRelayBtn('CPF_ROUTE_HIDDEN_123');
            const btnDsers = findRelayBtn('DSERS_ROUTE_HIDDEN_123');
            if (btnCpf && btnDsers) {
                clearInterval(bindClicks);
            }
            
            if (attempts > 30) {
                clearInterval(bindClicks);
            }
        }, 100);
    </script>
    """, height=0, width=0)

else:
    # ==========================
    # 内部操作流视图 (Inner Route View)
    # ==========================
    st.markdown("""
    <style>
        /* ====================
           主页同款美学形态之精炼横向胶囊按钮 (极其丝滑的 GPU 硬件加速光层与精准属性过渡)
           ==================== */
        .compact-action-outer {
            height: auto !important;
            min-height: 64px !important;
            padding: 0.3rem !important;
            border-radius: 9999px !important;
            margin-bottom: 0.5rem !important;
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.48) 0%, rgba(255, 255, 255, 0.15) 100%) !important;
            border: 1px solid rgba(255, 255, 255, 0.8) !important;
            box-shadow: 0 4px 18px rgba(0,0,0,0.04), inset 0 1px 1px rgba(255,255,255,0.8) !important;
            backdrop-filter: blur(40px) saturate(140%) !important;
            -webkit-backdrop-filter: blur(40px) saturate(140%) !important;
            transition: transform 0.7s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.7s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
            position: relative;
            overflow: hidden;
            will-change: transform;
        }

        /* 硬件加速的 ::before 高透光变层，实现 100% 丝滑渐亮，杜绝渐变突变 */
        .compact-action-outer::before {
            content: '';
            position: absolute;
            inset: 0;
            border-radius: 9999px !important;
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.92) 0%, rgba(255, 255, 255, 0.40) 100%) !important;
            opacity: 0;
            transition: opacity 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
            z-index: 1;
            pointer-events: none;
        }

        .compact-action-inner {
            height: auto !important;
            border-radius: 9999px !important;
            padding: 0.5rem 1.4rem !important;
            display: flex !important;
            align-items: center !important;
            justify-content: space-between !important;
            background: transparent !important;
            border: none !important;
            box-shadow: none !important;
            position: relative;
            z-index: 2;
        }

        .compact-btn-left {
            display: flex !important;
            align-items: center !important;
            gap: 0.8rem !important;
        }

        .compact-btn-icon {
            width: 38px !important;
            height: 38px !important;
            border-radius: 50% !important;
            background: linear-gradient(135deg, rgba(255,255,255,0.9) 0%, rgba(255,255,255,0.2) 100%) !important;
            border: 1px solid rgba(255, 255, 255, 0.9) !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-size: 1.15rem !important;
            flex-shrink: 0 !important;
            color: #1d1d1f !important;
            box-shadow: 0 3px 8px rgba(0,0,0,0.05) !important;
            transition: transform 0.7s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.7s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
        }

        .compact-btn-title {
            font-size: 1.12rem !important;
            font-weight: 700 !important;
            color: #1d1d1f !important;
            letter-spacing: 0.01em !important;
            margin: 0 !important;
            white-space: nowrap !important;
        }

        .compact-island {
            display: inline-flex !important;
            align-items: center !important;
            background: rgba(255, 255, 255, 0.6) !important;
            border: 1px solid rgba(255, 255, 255, 0.85) !important;
            border-radius: 9999px !important;
            padding: 0.35rem 0.35rem 0.35rem 1rem !important;
            gap: 0.6rem !important;
            flex-shrink: 0 !important;
            transition: transform 0.7s cubic-bezier(0.16, 1, 0.3, 1), background-color 0.7s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.7s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.03) !important;
        }

        .compact-island-text {
            font-size: 0.85rem !important;
            font-weight: 800 !important;
            color: #222 !important;
            letter-spacing: 0.05em !important;
            text-transform: uppercase !important;
        }

        .compact-island-circle {
            width: 30px !important;
            height: 30px !important;
            border-radius: 50% !important;
            background: #1d1d1f !important;
            color: #ffffff !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-size: 0.95rem !important;
            font-weight: 700 !important;
            transition: transform 0.7s cubic-bezier(0.16, 1, 0.3, 1), background-color 0.7s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
        }

        /* 强制停止按钮次级底色配置 */
        .compact-secondary {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.32) 0%, rgba(255, 255, 255, 0.08) 100%) !important;
            border: 1px solid rgba(255, 255, 255, 0.6) !important;
        }
        .compact-secondary .compact-btn-icon {
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.5) 0%, rgba(255, 255, 255, 0.1) 100%) !important;
            color: #444 !important;
        }

        /* 统一极致丝滑 0.7s 联动态（靠 GPU 层 opacity 渐显，绝无跳色） */
        .compact-action-outer:hover {
            transform: translateY(-4px) scale(1.015) !important;
            border-color: rgba(255, 255, 255, 1) !important;
            box-shadow: 0 16px 32px rgba(0,0,0,0.08), inset 0 1px 2px rgba(255,255,255,1) !important;
            cursor: pointer !important;
        }

        .compact-action-outer:hover::before {
            opacity: 1 !important;
        }

        .compact-action-outer:hover .compact-btn-icon {
            transform: scale(1.08) rotate(8deg) !important;
            border-color: rgba(255, 255, 255, 1) !important;
            box-shadow: 0 6px 14px rgba(0,0,0,0.1) !important;
        }

        .compact-action-outer:hover .compact-island {
            background-color: rgba(255, 255, 255, 0.95) !important;
            border-color: rgba(255, 255, 255, 1) !important;
            transform: translateX(-2px) !important;
            box-shadow: 0 6px 16px rgba(0,0,0,0.06) !important;
        }

        .compact-action-outer:hover .compact-island-circle {
            transform: translateX(3px) scale(1.05) !important;
            background-color: #000 !important;
            box-shadow: 0 4px 10px rgba(0,0,0,0.2) !important;
        }

        .compact-action-outer:active {
            transform: scale(0.98) !important;
            transition: transform 0.15s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }

        /* 返回主页精美胶囊卡片 (同样挂载 GPU 丝滑光感层) */
        .back-card-outer {
            height: auto !important;
            padding: 0.35rem !important;
            border-radius: 9999px !important;
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.55) 0%, rgba(255, 255, 255, 0.20) 100%) !important;
            border: 1px solid rgba(255, 255, 255, 0.85) !important;
            box-shadow: 0 4px 15px rgba(0,0,0,0.04), inset 0 1px 1px rgba(255,255,255,0.8) !important;
            backdrop-filter: blur(40px) saturate(140%) !important;
            -webkit-backdrop-filter: blur(40px) saturate(140%) !important;
            transition: transform 0.7s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.7s cubic-bezier(0.16, 1, 0.3, 1), border-color 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
            position: relative;
            overflow: hidden;
            will-change: transform;
        }

        .back-card-outer::before {
            content: '';
            position: absolute;
            inset: 0;
            border-radius: 9999px !important;
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.92) 0%, rgba(255, 255, 255, 0.45) 100%) !important;
            opacity: 0;
            transition: opacity 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
            z-index: 1;
            pointer-events: none;
        }

        .back-card-inner {
            border-radius: 9999px !important;
            padding: 0.45rem 1.4rem !important;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.6rem;
            height: auto !important;
            background: transparent !important;
            border: none !important;
            box-shadow: none !important;
            position: relative;
            z-index: 2;
        }

        .back-icon {
            font-size: 1.2rem !important;
            font-weight: 800;
            color: #1d1d1f;
            transition: transform 0.7s cubic-bezier(0.16, 1, 0.3, 1) !important;
        }

        .back-text {
            font-size: 1.05rem;
            font-weight: 700;
            color: #1d1d1f;
            letter-spacing: 0.02em;
        }

        .back-card-outer:hover {
            transform: translateY(-3px) scale(1.02) !important;
            border-color: rgba(255, 255, 255, 1) !important;
            box-shadow: 0 15px 30px rgba(0,0,0,0.08), inset 0 1px 2px rgba(255,255,255,1) !important;
            cursor: pointer !important;
        }

        .back-card-outer:hover::before {
            opacity: 1 !important;
        }

        .back-card-outer:hover .back-icon {
            transform: translateX(-4px) !important;
        }

        /* 统一完全隐藏作为事件中继的 Streamlit 原生 Relay 按钮容器（不使用 display:none 以确保 JS 可程序化点击触发） */
        .element-container:has(#visual-btn-back) + .element-container,
        .element-container:has(#visual-btn-launch) + .element-container,
        .element-container:has(#visual-btn-kill) + .element-container,
        .element-container:has([id*="visual-btn"]) + .element-container,
        .element-container:has([id*="home-card"]) + .element-container,
        .back-card-outer + div,
        .compact-action-outer + div,
        div:has(> #visual-btn-back) ~ div[data-testid="stButton"],
        div:has(> #visual-btn-launch) ~ div[data-testid="stButton"],
        div:has(> #visual-btn-kill) ~ div[data-testid="stButton"],
        div:has(> #visual-btn-dxm-cpf) ~ div[data-testid="stButton"],
        div:has(> #visual-btn-dxm-dsers) ~ div[data-testid="stButton"],
        div:has(> #home-card-cpf) ~ div[data-testid="stButton"],
        div:has(> #home-card-dsers) ~ div[data-testid="stButton"] {
            position: fixed !important;
            top: -9999px !important;
            left: -9999px !important;
            opacity: 0 !important;
            height: 1px !important;
            width: 1px !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow: hidden !important;
            z-index: -9999 !important;
        }
    </style>
    """, unsafe_allow_html=True)

    route_name = "CPF" if st.session_state.route == "A" else "DSERS"
    
    col1, col2 = st.columns([8, 2])
    with col1:
        st.markdown(f'<h1 style="margin-top:0;">{route_name}</h1>', unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="back-card-outer" id="visual-btn-back">
            <div class="back-card-inner">
                <span class="back-icon">←</span>
                <span class="back-text">返回主页</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("RELAY_BTN_BACK", key="btn_back_relay"):
            st.session_state.route = None
            if "route" in st.query_params:
                del st.query_params["route"]
            st.rerun()

    st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

    current_excel_path = SCRIPT_TEMPLATE if st.session_state.route == "A" else DSERS_TEMPLATE
    days = 1
    hours = 0
    search_val = ""
    sku_filter = "code"
    mabang_update_mode = "both"

    with st.container(border=True):
        st.markdown("### 第一步：获取要处理的订单数据")
        data_source = st.radio("选择获取方式", ["从马帮 ERP 提取", "上传本地表格", "使用已产生的表格"], horizontal=True, label_visibility="collapsed")
        sw_auto_export = (data_source == "从马帮 ERP 提取")
        use_vault = (data_source == "使用已产生的表格")

        if sw_auto_export:
            st.session_state["active_template_type"] = "脚本模板.xlsx"
            st.markdown("<br><b>马帮提取参数：</b>", unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                col_d, col_h = st.columns(2)
                with col_d:
                    days = st.number_input("提取最近几天", min_value=0, max_value=30, value=1)
                with col_h:
                    hours = st.number_input("精确到小时", min_value=0, max_value=23, value=0)
            with c2:
                if st.session_state.route == "A":
                    col_c1, col_c2 = st.columns(2)
                    with col_c1:
                        search_val = st.text_input("条件1: 买家客户ID (选填)", value="1000000257")
                    with col_c2:
                        sku_filter = st.text_input("条件2: 排除包含SKU (选填)", value="code")
                else:
                    search_val = st.text_input("要过滤排除的 SKU", value="code")
                    sku_filter = search_val
        elif data_source == "上传本地表格":
            st.markdown("<br>", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("请选择或拖入 .xlsx 表格文件", type=["xlsx", "xls"])
            if uploaded_file:
                file_id = f"{uploaded_file.name}_{uploaded_file.size}"
                if st.session_state.get("handled_file_id") != file_id:
                    select_uploaded_template_dialog(uploaded_file.getbuffer(), file_id)
                else:
                    t_type = st.session_state.get("active_template_type", "脚本模板.xlsx")
                    if t_type == "dsers模板.xlsx":
                        current_excel_path = DSERS_TEMPLATE
                    elif t_type == "import_orders.xlsx":
                        current_excel_path = DSERS_IMPORT_XLSX
                    elif t_type == "下单模板.xlsx":
                        current_excel_path = ORDER_TEMPLATE
                    else:
                        current_excel_path = SCRIPT_TEMPLATE
                    
                    c_msg, c_btn = st.columns([8, 2])
                    with c_msg:
                        st.success(f"已选定表格（识别类型：{t_type}）")
                    with c_btn:
                        if st.button("更改类型"):
                            st.session_state["handled_file_id"] = None
                            st.rerun()
            else:
                st.session_state["handled_file_id"] = None
        else:
            st.markdown("<br>", unsafe_allow_html=True)
            vault_file_choice = st.selectbox("请选择要继续操作的表格文件", ["dsers模板.xlsx", "import_orders.xlsx", "脚本模板.xlsx", "下单模板.xlsx"])
            st.session_state["active_template_type"] = vault_file_choice
            if vault_file_choice == "dsers模板.xlsx":
                current_excel_path = DSERS_TEMPLATE
            elif vault_file_choice == "import_orders.xlsx":
                current_excel_path = DSERS_IMPORT_XLSX
            elif vault_file_choice == "下单模板.xlsx":
                current_excel_path = ORDER_TEMPLATE
            else:
                current_excel_path = SCRIPT_TEMPLATE
            
            if os.path.exists(current_excel_path):
                st.success(f"已选择 {vault_file_choice}，可以继续执行后续处理。")
            else:
                st.warning(f"系统还未生成 {vault_file_choice}，请先在第一步生成或上传该表格。")

    st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

    dsers_update_fields = ["name", "cpf", "birthday"]
    if st.session_state.route == "A":
        with st.container(border=True):
            st.markdown("### 第二步：选择要执行的步骤")
            sw_cpf_rename = st.toggle("自动连接 Telegram 执行 CPF 查询与核对", value=True)
            sw_cpf_merge = st.toggle("将 CPF 查询结果回填至 DSERS 导入表格中", value=False)
            sw_mabang_update = st.toggle("将 CPF 查询核准结果同步更新回马帮 ERP", value=True)
            if sw_mabang_update:
                mode_choice = st.radio(
                    "马帮回传更新内容",
                    ["全部更新（姓名 + 出生日期）", "仅更新客户姓名", "仅更新出生日期（公司名称）"],
                    index=0,
                    horizontal=True
                )
                if mode_choice == "仅更新客户姓名":
                    mabang_update_mode = "name_only"
                elif mode_choice == "仅更新出生日期（公司名称）":
                    mabang_update_mode = "birth_only"
                else:
                    mabang_update_mode = "both"
        
        st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)
        render_dianxiaomi_stock_module(key_prefix="cpf")
    else:
        with st.container(border=True):
            st.markdown("### 第二步：选择要执行的步骤")
            if "sw_dsers_clean_key" not in st.session_state: st.session_state["sw_dsers_clean_key"] = True
            if "sw_dsers_cpf_check_key" not in st.session_state: st.session_state["sw_dsers_cpf_check_key"] = True
            if "sw_dsers_cpf_merge_key" not in st.session_state: st.session_state["sw_dsers_cpf_merge_key"] = True
            if "sw_dsers_mabang_key" not in st.session_state: st.session_state["sw_dsers_mabang_key"] = True
            if "sw_dsers_import_key" not in st.session_state: st.session_state["sw_dsers_import_key"] = True
            if "sw_dsers_rename_key" not in st.session_state: st.session_state["sw_dsers_rename_key"] = False

            sw_dsers_clean = st.toggle("把马帮表格清理并转换成 DSERS 要求的数据格式", key="sw_dsers_clean_key", on_change=on_dsers_normal_change)
            sw_dsers_cpf_check = st.toggle("自动连接 Telegram 查询并核对买家 CPF 真实姓名", key="sw_dsers_cpf_check_key", on_change=on_dsers_normal_change)
            sw_dsers_cpf_merge = st.toggle("把 CPF 查询核准后的正确姓名填入 DSERS 表格中", key="sw_dsers_cpf_merge_key", on_change=on_dsers_normal_change)
            sw_dsers_mabang = st.toggle("把 CPF 查询核准完成的姓名同步更新回马帮 ERP", key="sw_dsers_mabang_key", on_change=on_dsers_normal_change)
            sw_dsers_import = st.toggle("一键把整理好的表格上传到 DSERS 后台并批量建单", key="sw_dsers_import_key", on_change=on_dsers_normal_change)
            st.markdown("---")
            sw_dsers_rename = st.toggle("直接打开 DSERS 网页端，针对后台已有订单自动修改买家信息", key="sw_dsers_rename_key", on_change=on_dsers_rename_change)
            if sw_dsers_rename:
                st.markdown("<span style='font-size: 0.9rem; color: #475569; font-weight: 500;'>选择需要修改的字段（支持勾选组合）：</span>", unsafe_allow_html=True)
                col_f1, col_f2, col_f3 = st.columns(3)
                with col_f1:
                    dsers_mod_name = st.checkbox("客户姓名", value=True, key="dsers_mod_name_key")
                with col_f2:
                    dsers_mod_cpf = st.checkbox("CPF 税号", value=True, key="dsers_mod_cpf_key")
                with col_f3:
                    dsers_mod_birth = st.checkbox("出生日期 (Birthday)", value=True, key="dsers_mod_birth_key")

                dsers_update_fields = []
                if dsers_mod_name: dsers_update_fields.append("name")
                if dsers_mod_cpf: dsers_update_fields.append("cpf")
                if dsers_mod_birth: dsers_update_fields.append("birthday")
                if not dsers_update_fields:
                    st.caption("⚠️ 未勾选任何修改项，请至少勾选一个字段。")


        st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)
        render_dianxiaomi_stock_module(key_prefix="dsers")


    st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)

    # ==========================
    # 数据保险库 Data Vault (实时预览)
    # ==========================
    with st.expander("查看与检查数据表格 (点击展开或修改)", expanded=False):
        st.markdown("在这里可以直接查看、修改和保存系统中流转的数据表格，不用去文件夹里反复翻找。")
        vault_tabs = st.tabs(["dsers模板.xlsx", "import_orders.xlsx", "脚本模板.xlsx", "下单模板.xlsx"])
        
        with vault_tabs[0]:
            if os.path.exists(DSERS_TEMPLATE):
                try:
                    df1 = pd.read_excel(DSERS_TEMPLATE, dtype=str)
                    st.caption(f"当前规格：**{df1.shape[0]}** 行 × **{df1.shape[1]}** 列")
                    edited_df1 = st.data_editor(df1, num_rows="dynamic", use_container_width=True, key="edit_dsers")
                    
                    colA, colB = st.columns(2)
                    with colA:
                        if st.button("保存修改", key="save_dsers"):
                            if not df1.equals(edited_df1):
                                edited_df1.to_excel(DSERS_TEMPLATE, index=False)
                                st.success("修改已保存至 dsers模板.xlsx")
                            else:
                                st.info("内容无变化，无需保存。")
                    with colB:
                        with open(DSERS_TEMPLATE, "rb") as f:
                            st.download_button("下载 dsers模板", f, "dsers模板.xlsx", key="dl_dsers")
                except Exception as e:
                    st.warning(f"无法读取: {e}")
            else:
                st.info("暂无文件数据。")
                
        with vault_tabs[1]:
            if os.path.exists(DSERS_IMPORT_XLSX):
                try:
                    df2 = pd.read_excel(DSERS_IMPORT_XLSX, dtype=str)
                    st.caption(f"当前规格：**{df2.shape[0]}** 行 × **{df2.shape[1]}** 列")
                    edited_df2 = st.data_editor(df2, num_rows="dynamic", use_container_width=True, key="edit_import")
                    
                    colA, colB = st.columns(2)
                    with colA:
                        if st.button("保存修改", key="save_import"):
                            if not df2.equals(edited_df2):
                                edited_df2.to_excel(DSERS_IMPORT_XLSX, index=False)
                                st.success("修改已保存至 import_orders.xlsx")
                            else:
                                st.info("内容无变化，无需保存。")
                    with colB:
                        with open(DSERS_IMPORT_XLSX, "rb") as f:
                            st.download_button("下载 import_orders", f, "import_orders.xlsx", key="dl_import")
                except Exception as e:
                    st.warning(f"无法读取: {e}")
            else:
                st.info("暂无文件数据。")
                
        with vault_tabs[2]:
            if os.path.exists(SCRIPT_TEMPLATE):
                try:
                    df3 = pd.read_excel(SCRIPT_TEMPLATE, dtype=str)
                    st.caption(f"当前规格：**{df3.shape[0]}** 行 × **{df3.shape[1]}** 列")
                    edited_df3 = st.data_editor(df3, num_rows="dynamic", use_container_width=True, key="edit_script")
                    
                    colA, colB = st.columns(2)
                    with colA:
                        if st.button("保存修改", key="save_script"):
                            if not df3.equals(edited_df3):
                                edited_df3.to_excel(SCRIPT_TEMPLATE, index=False)
                                st.success("修改已保存至 脚本模板.xlsx")
                            else:
                                st.info("内容无变化，无需保存。")
                    with colB:
                        with open(SCRIPT_TEMPLATE, "rb") as f:
                            st.download_button("下载 脚本模板", f, "脚本模板.xlsx", key="dl_script")
                except Exception as e:
                    st.warning(f"无法读取: {e}")
            else:
                st.info("暂无文件数据。")
                
        with vault_tabs[3]:
            if os.path.exists(ORDER_TEMPLATE):
                try:
                    df4 = pd.read_excel(ORDER_TEMPLATE, dtype=str)
                    st.caption(f"当前规格：**{df4.shape[0]}** 行 × **{df4.shape[1]}** 列")
                    edited_df4 = st.data_editor(df4, num_rows="dynamic", use_container_width=True, key="edit_order")
                    
                    colA, colB = st.columns(2)
                    with colA:
                        if st.button("保存修改", key="save_order"):
                            if not df4.equals(edited_df4):
                                edited_df4.to_excel(ORDER_TEMPLATE, index=False)
                                st.success("修改已保存至 下单模板.xlsx")
                            else:
                                st.info("内容无变化，无需保存。")
                    with colB:
                        with open(ORDER_TEMPLATE, "rb") as f:
                            st.download_button("下载 下单模板", f, "下单模板.xlsx", key="dl_order")
                except Exception as e:
                    st.warning(f"无法读取: {e}")
            else:
                st.info("暂无文件数据。")

    st.markdown("<div style='height: 2rem;'></div>", unsafe_allow_html=True)

    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        st.markdown("""
        <div class="compact-action-outer" id="visual-btn-launch">
            <div class="compact-action-inner">
                <div class="compact-btn-left">
                    <div class="compact-btn-icon">⚡</div>
                    <div class="compact-btn-title">立即开始处理</div>
                </div>
                <div class="compact-island">
                    <span class="compact-island-text">RUN PIPELINE</span>
                    <span class="compact-island-circle">↗</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        btn_launch = st.button("RELAY_BTN_LAUNCH", key="btn_launch_relay")
    with c_btn2:
        st.markdown("""
        <div class="compact-action-outer compact-secondary" id="visual-btn-kill">
            <div class="compact-action-inner">
                <div class="compact-btn-left">
                    <div class="compact-btn-icon">✕</div>
                    <div class="compact-btn-title">强制停止并清理卡死进程</div>
                </div>
                <div class="compact-island">
                    <span class="compact-island-text">TERMINATE</span>
                    <span class="compact-island-circle">↺</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        btn_kill = st.button("RELAY_BTN_KILL", key="btn_kill_relay")

    # ==========================
    # 操作按钮 JS 物理透传引擎 (跨 iframe 精准中继点击，并深度隐藏原生 relay 按钮)
    # ==========================
    components.html("""
    <script>
        const parentDoc = window.parent.document;
        let attempts = 0;
        const bindRelays = setInterval(() => {
            attempts++;
            
            // 全匹配隐藏并定位真正承载逻辑的原生 Relay Button
            const findRelayBtn = (textTag) => {
                const allBtns = parentDoc.querySelectorAll('button');
                for (let btn of allBtns) {
                    if (btn.textContent && btn.textContent.includes(textTag)) {
                        let container = btn.closest('.element-container') || btn.closest('[data-testid="stButton"]') || btn;
                        container.style.setProperty('display', 'none', 'important');
                        container.style.setProperty('position', 'absolute', 'important');
                        container.style.setProperty('opacity', '0', 'important');
                        container.style.setProperty('height', '0', 'important');
                        container.style.setProperty('pointer-events', 'none', 'important');
                        return btn;
                    }
                }
                return null;
            };
            
            const bindPair = (cardId, textTag) => {
                const card = parentDoc.getElementById(cardId);
                if (card) {
                    card.style.cursor = 'pointer';
                    card.onclick = (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        const liveBtn = findRelayBtn(textTag);
                        if (liveBtn) {
                            liveBtn.click();
                        }
                    };
                }
            };
            
            bindPair('visual-btn-back', 'RELAY_BTN_BACK');
            bindPair('visual-btn-launch', 'RELAY_BTN_LAUNCH');
            bindPair('visual-btn-kill', 'RELAY_BTN_KILL');
            bindPair('visual-btn-dxm-cpf', 'RELAY_BTN_DXM_CPF');
            bindPair('visual-btn-dxm-dsers', 'RELAY_BTN_DXM_DSERS');
            
            if (attempts > 30) clearInterval(bindRelays);
        }, 100);
    </script>
    """, height=0, width=0)
        
    if btn_kill:
        st.warning("正在停止所有后台程序并清理缓存锁...")
        for t in global_task_manager.get_active_tasks():
            global_task_manager.cancel_task(t.task_id)
        os.system("pkill -i -f playwright")
        os.system("pkill -i -f 'remote-debugging-pipe'")
        os.system("pkill -i -f 'user-data-dir.*sessions'")
        for lock_file in glob.glob(os.path.join(SESSIONS_DIR, "*", "Singleton*")):
            try: os.remove(lock_file)
            except: pass
        st.success("后台程序与任务池已被完全停止清理，您可以重新开始执行了。")
        st.stop()

    if btn_launch:
        if not sw_auto_export and not os.path.exists(current_excel_path):
            if use_vault:
                st.error(f"无法开始：'{vault_file_choice}' 还不存在，请先执行第一步把它生成出来。")
            else:
                st.error("无法开始：请先在第一步中准备好表格数据。")
        else:
            task_id = f"pipeline_{st.session_state.route.lower()}_{int(time.time())}"
            pipeline_title = "CPF 订单核对流水线" if st.session_state.route == "A" else "DSERS 批量下单流水线"
            global_task_manager.submit_task(
                task_id=task_id,
                name=pipeline_title,
                target_fn=execute_pipeline_task,
                kwargs={
                    "route": st.session_state.route,
                    "sw_auto_export": sw_auto_export,
                    "days": days,
                    "hours": hours,
                    "search_val": search_val,
                    "current_excel_path": current_excel_path,
                    "active_template_type": st.session_state.get("active_template_type", ""),
                    "sw_cpf_rename": sw_cpf_rename if st.session_state.route == "A" else False,
                    "sw_cpf_merge": sw_cpf_merge if st.session_state.route == "A" else False,
                    "sw_mabang_update": sw_mabang_update if st.session_state.route == "A" else False,
                    "sw_dsers_clean": sw_dsers_clean if st.session_state.route == "B" else False,
                    "sw_dsers_cpf_check": sw_dsers_cpf_check if st.session_state.route == "B" else False,
                    "sw_dsers_cpf_merge": sw_dsers_cpf_merge if st.session_state.route == "B" else False,
                    "sw_dsers_mabang": sw_dsers_mabang if st.session_state.route == "B" else False,
                    "sw_dsers_import": sw_dsers_import if st.session_state.route == "B" else False,
                    "sw_dsers_rename": sw_dsers_rename if st.session_state.route == "B" else False,
                    "dsers_update_fields": dsers_update_fields if (st.session_state.route == "B" and sw_dsers_rename) else ["name", "cpf", "birthday"],
                    "use_vault": use_vault,
                    "vault_file_choice": vault_file_choice if use_vault else "",
                    "sku_filter": sku_filter,
                    "mabang_update_mode": mabang_update_mode,
                },
                category="流水线"
            )
            st.toast(f"{pipeline_title} 已在后台启动！可随时在右上角【任务看板】查看实时耗时与日志。")
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

with st.expander("页面背景光效调节", expanded=False):
    thickness = st.slider("核心厚度 (Thickness)", 0.05, 0.30, 0.131, 0.001)
    blur = st.slider("上下晕染 (Blur)", 0.01, 0.30, 0.12, 0.001)
    brightness = st.slider("整体曝光 (Brightness)", 0.1, 2.0, 0.8, 0.01)

st.components.v1.html(f'''
<script>
    const parentDoc = window.parent.document;
    const iframe = parentDoc.getElementById("vanguard-3d-bg");
    if (iframe && iframe.contentWindow) {{
        iframe.contentWindow.postMessage({{
            type: "UPDATE_STRANDS",
            payload: {{ thickness: {thickness}, blur: {blur}, brightness: {brightness} }}
        }}, "*");
    }}
</script>
''', height=0)
