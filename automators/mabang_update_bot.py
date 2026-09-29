import asyncio
import os
import yaml
from openpyxl import load_workbook
from playwright.async_api import async_playwright

async def run_mabang_batch_update(excel_path: str, user_data_dir: str, update_mode: str = "both", headless: bool = False, progress_callback=None, task_info=None):
    def log(msg):
        print(msg)
        if task_info:
            task_info.check_pause()
        if progress_callback:
            progress_callback(msg)

    # 1. 读取配置文件获取账号密码
    config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.yaml")
    with open(config_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    mb_user = str(config['mabang']['username'])
    mb_pwd = str(config['mabang']['password'])

    import pandas as pd
    log(f"[*] 读取 Excel 数据: {excel_path} (更新模式: {update_mode})")
    df = pd.read_excel(excel_path, dtype=str)
    
    total_count = len(df)
    processed_count = 0
    update_data = []
    birth_update_data = []
    for idx, row in df.iterrows():
        if task_info:
            task_info.check_pause()
        processed_count += 1
        order_id = str(row.iloc[0]).strip()
        # 严格过滤无订单号的空白行
        if not order_id or order_id.lower() in ["", "nan", "none"]:
            continue

        # 优先通过表头名称定位“查询结果”或“TG_Result”，避免索引错位
        result_col = next((c for c in df.columns if str(c).strip() in ['查询结果', 'TG_Result', '姓名查询结果']), None)
        if result_col:
            new_name = str(row[result_col]).strip()
        elif len(row) > 4:
            new_name = str(row.iloc[4]).strip()
        else:
            new_name = ""

        # 防御校验：绝不把指令前缀 /cpf1 误当成真实客户姓名提交回传
        if new_name.startswith('/cpf1'):
            new_name = ""

        birth_date = ""
        if '出生日期' in df.columns:
            birth_date = str(row['出生日期']).strip()
        elif len(row) > 5:
            birth_date = str(row.iloc[5]).strip()
            
        # 防止读取出科学计数法（如果真的被Excel弄坏了的话，尝试去掉 .0 等）
        if 'e+' in order_id.lower() or 'E+' in order_id:
            try:
                order_id = str(int(float(order_id)))
            except:
                pass
        elif order_id.endswith('.0'):
            order_id = order_id[:-2]

        # 过滤无效的名字（兼容各种可能的空或错误状态）
        if order_id and new_name and new_name not in ["", "nan", "None", "无", "遇到验证码且未能通过", "查询超时", "提取失败"]:
            update_data.append(f"{order_id}\t{new_name}")

        # 过滤无效的出生日期
        if order_id and birth_date and birth_date not in ["", "nan", "None", "无", "遇到验证码且未能通过", "查询超时", "提取失败"]:
            birth_update_data.append(f"{order_id}\t{birth_date}")

        log(f"[*] 进度提示：现在是 {processed_count}/{total_count} (共需读取 {total_count} 条，准备好姓名: {len(update_data)} 条, 出生日期: {len(birth_update_data)} 条)")
        if task_info:
            task_info.set_progress(processed_count, total_count, prefix="马帮同步", unit="条")
            
    should_update_name = (update_mode in ["both", "name_only"]) and bool(update_data)
    should_update_birth = (update_mode in ["both", "birth_only"]) and bool(birth_update_data)

    if not should_update_name and not should_update_birth:
        log(f"[!] 根据当前模式 '{update_mode}'，没有需要更新的数据（有效姓名: {len(update_data)} 条, 有效出生日期: {len(birth_update_data)} 条），流程结束。")
        return True
        
    tsv_text = "\n".join(update_data)
    birth_tsv_text = "\n".join(birth_update_data)
    log(f"[*] 数据提取就绪：姓名待更新 {len(update_data)} 条 (执行: {should_update_name})，出生日期待更新 {len(birth_update_data)} 条 (执行: {should_update_birth})。")

    for item in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        p_lock = os.path.join(user_data_dir, item)
        if os.path.exists(p_lock) or os.path.islink(p_lock):
            try: os.remove(p_lock)
            except Exception: pass

    log("[*] 启动浏览器...")
    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=headless,
            viewport={'width': 1280, 'height': 800}
        )
        if task_info:
            task_info.register_context(context)
        page = context.pages[0] if len(context.pages) > 0 else await context.new_page()
        
        try:
            # === 第一步：从首页进入并处理自动登录 ===
            log("[*] 正在打开马帮首页...")
            await page.goto("https://901067.private.mabangerp.com/index.htm", wait_until='domcontentloaded', timeout=60000)
            
            # 判断是否需要登录
            await asyncio.sleep(2)
            if await page.locator('#login-but').is_visible():
                log("[*] 检测到未登录状态，开始自动登录...")
                await page.locator('input[name="username"]').first.fill(mb_user)
                await page.locator('input[name="password"]').first.fill(mb_pwd)
                await page.locator('#login-but').click()
                log("[*] 登录已提交，等待页面加载...")
                await page.wait_for_load_state('domcontentloaded')
                await asyncio.sleep(3)
            
            # === 第二步：处理首页的“店铺授权提醒”弹窗 ===
            log("[*] 等待并检查是否有授权提醒弹窗...")
            try:
                # 弹窗加载可能较慢，等待最多8秒
                await page.wait_for_selector('text="店铺授权提醒"', timeout=8000)
                log("[*] 发现授权提醒弹窗，正在关闭...")
                
                # 尝试勾选 7天内不再重复提醒
                checkbox = page.locator('text="7天内不再重复提醒"')
                if await checkbox.count() > 0:
                    await checkbox.click()
                    
                # layui的确认按钮通常是一个 <a> 标签带有 layui-layer-btn0 class
                await page.locator('.layui-layer-btn0, a:has-text("确认")').first.click()
                log("[*] 已点击确认关闭弹窗。")
                await asyncio.sleep(2)
            except Exception:
                log("[*] 没有发现弹窗，继续操作。")
            
            # === 第三步：直接导航到订单列表（比模拟菜单点击更稳定） ===
            log("[*] 导航至【订单列表】...")
            await page.goto("https://901067.private.mabangerp.com/index.php?mod=order.list&Order_orderStatus=2", wait_until='domcontentloaded')
            
            # 等待订单列表页面完全加载（等待批处理功能按钮出现）
            await page.wait_for_selector('span.text.mr5.ml5:has-text("批处理功能")', timeout=30000)
            await asyncio.sleep(2) # 缓冲一下，防止 DOM 渲染中
            
            # === 第四步：定义打开弹窗函数 ===
            async def open_batch_modal():
                log("[*] 正在打开【更新订单基本信息】弹窗...")
                # 直接调用马帮内置函数，避免因 UI 重叠、分辨率导致 Hover 和 Click 失效
                await page.evaluate("""
                    if(typeof updateEmailShow === 'function') {
                        updateEmailShow();
                    } else {
                        let a = document.querySelector('a[onclick*="updateEmailShow"]');
                        if (a) a.click();
                    }
                """)
                
                # 额外做个安全校验，如果弹窗没出来，尝试备用点击
                try:
                    await page.wait_for_selector('#updateEmail', state='visible', timeout=4000)
                except Exception:
                    log("[*] 警告：快速弹窗失败，尝试使用原生点击...")
                    batch_btn = page.locator('span.text.mr5.ml5:has-text("批处理功能")').first
                    if await batch_btn.count() > 0:
                        await batch_btn.click(force=True)
                        await asyncio.sleep(1)
                    
                    menu_item = page.locator('li[data-customlink="批量更新订单信息"]').first
                    if await menu_item.count() > 0:
                        await menu_item.hover()
                        await asyncio.sleep(1)
                    
                    sub_item = page.locator('a[onclick*="updateEmailShow"]').first
                    if await sub_item.count() > 0:
                        await sub_item.click(force=True)
                        
                await asyncio.sleep(2)

            # === 定义等待更新完成并自动关闭弹窗函数 ===
            async def wait_and_handle_update_modals(field_desc: str):
                log(f"[*] {field_desc}注入成功！请手动检查数据无误后，自行点击【确定】保存。")
                log("[*] 脚本将监控更新完成状态（若出现返回结果将自动完成关闭）...")

                # 监听页面是否有“返回结果”弹窗出现的标志
                await page.evaluate("""() => {
                    window.__mabang_result_shown = false;
                    if (!window.__mabang_result_observer) {
                        window.__mabang_result_observer = new MutationObserver(() => {
                            const dialogs = document.querySelectorAll('.modal, [role="dialog"], .bootbox, .layui-layer, div[class*="dialog"], div[class*="modal"]');
                            for (const d of dialogs) {
                                if (d.innerText && d.innerText.includes('返回结果')) {
                                    window.__mabang_result_shown = true;
                                }
                            }
                        });
                        window.__mabang_result_observer.observe(document.body, { childList: true, subtree: true });
                    }
                }""")

                while True:
                    if task_info:
                        task_info.check_pause()

                    # 检查批量更新弹窗 #updateEmail 是否仍然可见
                    is_update_open = False
                    try:
                        is_update_open = await page.locator('#updateEmail').is_visible()
                    except Exception:
                        # 页面可能刷新或导航完成
                        break

                    # 检查当前是否有图2【返回结果】弹窗（或曾经弹出过）
                    has_result_modal = await page.evaluate("""() => {
                        if (window.__mabang_result_shown) return true;
                        const dialogs = Array.from(document.querySelectorAll('.modal, [role="dialog"], .bootbox, .layui-layer, div[class*="dialog"], div[class*="modal"]'));
                        for (const d of dialogs) {
                            const s = window.getComputedStyle(d);
                            if (s.display !== 'none' && s.visibility !== 'hidden' && d.offsetHeight > 0) {
                                if (d.innerText && d.innerText.includes('返回结果')) {
                                    return true;
                                }
                            }
                        }
                        return false;
                    }""")

                    # 如果图1弹窗已经自然关闭且未出现图2返回结果弹窗，说明正常无误结束
                    if not is_update_open and not has_result_modal:
                        break

                    if has_result_modal:
                        log(f"[*] 检测到图2【返回结果】弹窗，等待更新加载完全...")
                        # 缓冲等待，确保数据与请求完全处理完毕
                        await asyncio.sleep(2)

                        # 1. 自动关闭图2【返回结果】弹窗
                        log("[*] 正在自动关闭图2【返回结果】弹窗...")
                        try:
                            close_btn = page.locator('.modal:has-text("返回结果") button:has-text("关闭"), [role="dialog"]:has-text("返回结果") button:has-text("关闭"), .bootbox:has-text("返回结果") button:has-text("关闭"), div:has-text("返回结果") button:has-text("关闭"), button:has-text("关闭")').first
                            if await close_btn.count() > 0:
                                await close_btn.click()
                            else:
                                x_btn = page.locator('.modal:has-text("返回结果") .close, [role="dialog"]:has-text("返回结果") .close').first
                                if await x_btn.count() > 0:
                                    await x_btn.click()
                                else:
                                    await page.evaluate("""() => {
                                        const dialogs = Array.from(document.querySelectorAll('.modal, [role="dialog"], .bootbox, .layui-layer, div[class*="dialog"], div[class*="modal"]'));
                                        for (const d of dialogs) {
                                            if (d.innerText && d.innerText.includes('返回结果')) {
                                                const b = d.querySelector('button, .close');
                                                if (b) b.click();
                                                else if (typeof $ !== 'undefined' && $(d).modal) $(d).modal('hide');
                                                else d.style.display = 'none';
                                            }
                                        }
                                    }""")
                        except Exception as e:
                            log(f"[*] 关闭返回结果弹窗提示: {e}")

                        await asyncio.sleep(1)

                        # 2. 自动关闭图1【批量更新】弹窗
                        try:
                            if await page.locator('#updateEmail').is_visible():
                                log("[*] 正在自动关闭图1【批量更新】弹窗...")
                                cancel_btn = page.locator('#updateEmail button:has-text("取消"), #updateEmail a:has-text("取消"), #updateEmail input[value="取消"]').first
                                if await cancel_btn.count() > 0:
                                    await cancel_btn.click()
                                else:
                                    x_btn = page.locator('#updateEmail .close, #updateEmail [data-dismiss="modal"]').first
                                    if await x_btn.count() > 0:
                                        await x_btn.click()
                                    else:
                                        await page.evaluate("""() => {
                                            if (typeof $ !== 'undefined' && $('#updateEmail').modal) {
                                                $('#updateEmail').modal('hide');
                                            }
                                            const el = document.querySelector('#updateEmail');
                                            if (el) el.style.display = 'none';
                                            document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
                                        }""")
                        except Exception as e:
                            log(f"[*] 关闭批量更新弹窗提示: {e}")
                        
                        await page.evaluate("""() => {
                            window.__mabang_result_shown = false;
                            document.querySelectorAll('.modal-backdrop').forEach(b => b.remove());
                            document.body.classList.remove('modal-open');
                        }""")

                        await asyncio.sleep(1)
                        break

                    await asyncio.sleep(0.5)

                log(f"[*] 检测到{field_desc}更新完成并关闭弹窗。")
                await asyncio.sleep(2)

            # === 第五步：第一次更新 - 客户姓名 ===
            if should_update_name:
                await open_batch_modal()
                log("[*] 已打开更新弹窗。正在切换更新字段为【客户姓名】...")
                
                try:
                    await page.locator('select[name="select1"]').select_option(label="按 订单编号")
                except Exception:
                    pass

                # 使用 Playwright 专门处理原生 select 的 select_option 方法，绝对稳定
                await page.locator('select[name="select2"]').select_option(label="客户姓名")
                await asyncio.sleep(1)
                
                log("[*] 正在注入【客户姓名】更新数据...")
                # 使用原生唯一的 name 属性定位输入框，不会和其他输入框冲突
                textarea = page.locator('textarea[name="updateData"]')
                await textarea.fill("")
                await textarea.fill(tsv_text)
                await wait_and_handle_update_modals("客户姓名")
            else:
                log(f"[*] 跳过客户姓名更新 (模式: {update_mode}, 有效姓名: {len(update_data)} 条)。")

            # === 第六步：第二次更新 - 公司/门店名称 (出生日期) ===
            if should_update_birth:
                log("[*] 正在进行批量更新：准备将出生日期更新至【公司/门店名称】...")
                await open_batch_modal()
                
                try:
                    await page.locator('select[name="select1"]').select_option(label="按 订单编号")
                except Exception:
                    pass

                log("[*] 已重新打开更新弹窗。正在切换更新字段为【公司/门店名称】...")
                try:
                    await page.locator('select[name="select2"]').select_option(label="公司/门店名称")
                except Exception:
                    options = await page.locator('select[name="select2"] option').all_inner_texts()
                    target_opt = next((opt.strip() for opt in options if "公司" in opt), "公司/门店名称")
                    await page.locator('select[name="select2"]').select_option(label=target_opt)
                await asyncio.sleep(1)
                
                log("[*] 正在注入【公司/门店名称】(出生日期) 更新数据...")
                textarea = page.locator('textarea[name="updateData"]')
                await textarea.fill("")
                await textarea.fill(birth_tsv_text)
                await wait_and_handle_update_modals("公司/门店名称(出生日期)")
            else:
                log(f"[*] 跳过公司/门店名称(出生日期)更新 (模式: {update_mode}, 有效出生日期: {len(birth_update_data)} 条)。")

            log("[*] 马帮批量更新流程全部执行完毕！")
            
        except Exception as e:
            log(f"[!] 发生错误: {e}")
            try:
                await page.screenshot(path="debug_mabang_error.png")
            except:
                pass
            import sys
            sys.exit(1)
            
        finally:
            await context.close()
            log("[*] 测试结束。")

if __name__ == "__main__":
    import sys
    sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from config import SCRIPT_TEMPLATE

    EXCEL_PATH = SCRIPT_TEMPLATE
    SESSION_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sessions", "mabang_session")
    asyncio.run(run_mabang_batch_update(EXCEL_PATH, SESSION_DIR, headless=False))
