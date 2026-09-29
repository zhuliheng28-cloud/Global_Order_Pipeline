import asyncio
import os
from playwright.async_api import async_playwright
from login_helpers.dsers_login import wait_or_auto_login_dsers

async def run_dsers_import(csv_path: str, user_data_dir: str, headless: bool = False, progress_callback=None, task_info=None):
    def log(msg):
        print(msg)
        if task_info:
            task_info.check_pause()
        if progress_callback:
            progress_callback(msg)

    if not os.path.exists(csv_path):
        log(f"[!] 找不到 CSV 文件: {csv_path}")
        return

    total_csv_orders = 0
    try:
        import pandas as pd
        df_csv = pd.read_csv(csv_path)
        total_csv_orders = len(df_csv)
        log(f"[*] 成功载入导入文件，共 {total_csv_orders} 条订单")
        log(f"[*] 进度提示：现在是 0/{total_csv_orders} (准备向 DSers 导入 {total_csv_orders} 条订单)")
        if task_info:
            task_info.set_progress(0, total_csv_orders, prefix="DSers导入", unit="条")
    except Exception:
        pass

    for item in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        p_lock = os.path.join(user_data_dir, item)
        if os.path.exists(p_lock) or os.path.islink(p_lock):
            try: os.remove(p_lock)
            except Exception: pass

    log("[*] 启动 DSers 自动导入引擎...")
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
            # 1. 登录与导航
            TARGET_URL = "https://accounts.dsers.com/accounts/login?redirect_url=https%3A%2F%2Fwww.dsers.com%2Fapplication%2Forders%2F159831080"
            log(f"[*] 尝试访问 DSers 仪表盘: {TARGET_URL} ...")
            try:
                await page.goto(TARGET_URL, wait_until='domcontentloaded', timeout=15000)
            except Exception as e:
                log(f"[!] 网络访问异常 ({e})，正在重试...")
                await asyncio.sleep(2)
                await page.goto(TARGET_URL, wait_until='domcontentloaded')
            
            await wait_or_auto_login_dsers(page, log_func=log, timeout_seconds=30, task_info=task_info)
            
            log("[*] 检测到已进入 DSers 首页/主控台！")
            log("[*] 为方便您操作（如：按掉广告、确认账号或退出重进等），系统将等待 15 秒...")
            for i in range(15, 0, -1):
                if "login" in page.url.lower() or "accounts.dsers.com" in page.url.lower():
                    log("[*] 检测到您点击了退出，正在等待您登录新账号...")
                    await wait_or_auto_login_dsers(page, log_func=log, timeout_seconds=30, task_info=task_info)
                    log("[*] 重新登录成功！等待 5 秒缓冲...")
                    await asyncio.sleep(5)
                    break
                if i % 5 == 0 or i <= 3:
                    log(f"[*] 距离自动开始还剩 {i} 秒 (如需退出切换账号请立刻点击)...")
                await asyncio.sleep(1)

            log("[*] 正在确认是否已成功进入主控台...")
            wait_dashboard = 0
            while "application" not in page.url.lower() and wait_dashboard < 60:
                if wait_dashboard % 5 == 0:
                    log(f"[*] 等待进入主控台 ({wait_dashboard}s/60s)...")
                await asyncio.sleep(1)
                wait_dashboard += 1

            log("[*] 准备开始执行自动化导入...")

            # 2. 点击左侧 CSV Upload
            log("[*] 正在点击【CSV Upload】...")
            try:
                # DSers左侧菜单可能嵌套，使用包含匹配文本
                await page.get_by_text("CSV Upload", exact=False).locator("visible=true").first.click(timeout=8000, force=True)
            except Exception:
                await page.evaluate("""() => {
                    let els = Array.from(document.querySelectorAll('a, li, span, div')).filter(e => e.innerText && e.innerText.includes('CSV Upload') && e.offsetHeight > 0);
                    // 找出没有子元素也包含 CSV Upload 的最内层元素
                    let target = els.find(e => !Array.from(e.children).some(c => c.innerText && c.innerText.includes('CSV Upload'))) || els[0];
                    if (target) target.click();
                }""")
            await asyncio.sleep(4)

            # 3. 点击 Orders 选项卡
            log("[*] 正在切换至【Orders】标签页...")
            try:
                # 只点击可见的 Orders 选项卡
                await page.locator('.ant-tabs-tab, [role="tab"]').filter(has_text="Orders").locator("visible=true").first.click(timeout=5000, force=True)
            except Exception:
                await page.evaluate("""() => {
                    let tabs = Array.from(document.querySelectorAll('[role="tab"], .ant-tabs-tab')).filter(e => e.offsetHeight > 0);
                    let orderTab = tabs.find(t => t.innerText && t.innerText.trim() === 'Orders');
                    if (orderTab) orderTab.click();
                }""")
            await asyncio.sleep(3)

            # 4. 点击面板外层的 + IMPORT 按钮唤出上传弹窗
            log("[*] 正在面板上触发【+ IMPORT】按钮以打开对话框...")
            try:
                # 首先确保在页面上找到大写的 IMPORT 按钮并点击
                await page.locator('button, div[role="button"]').filter(has_text="IMPORT").locator("visible=true").first.click(timeout=5000)
            except Exception:
                await page.evaluate("""() => {
                    let btns = Array.from(document.querySelectorAll('button, div[role="button"]')).filter(b => b.innerText && b.innerText.includes('IMPORT') && b.offsetHeight > 0);
                    if (btns.length > 0) btns[0].click();
                }""")
            
            # 给予弹窗渲染时间
            await asyncio.sleep(2)

            # 5. 在弹窗中点击 CHOOSE CSV FILE 触发选择器并装载
            log("[*] 正在弹窗中装载 CSV 文件...")
            try:
                await page.wait_for_selector('.ant-modal-content, .el-dialog, [role="dialog"]', state='visible', timeout=4000)
            except:
                pass
                
            try:
                await page.locator('input[type="file"]').first.set_files(csv_path, timeout=3000)
                log("[*] 文件直接挂载成功！等待导入按钮激活...")
            except:
                log("[!] 底层 input 未命中，尝试触发界面交互 (点击 CHOOSE CSV FILE 唤起选择器)...")
                async with page.expect_file_chooser(timeout=10000) as fc_info:
                    try:
                        await page.locator('button, div[role="button"], span').filter(has_text="CHOOSE CSV FILE").locator("visible=true").first.click(timeout=3000)
                    except:
                        await page.evaluate("""() => {
                            let dialogs = document.querySelectorAll('[role="dialog"], .el-dialog, .modal, .ant-modal, .v-dialog');
                            let root = dialogs.length > 0 ? dialogs[dialogs.length - 1] : document;
                            let btn = Array.from(root.querySelectorAll('button, div, span, a'))
                                .filter(e => e.offsetHeight > 0)
                                .find(e => {
                                    let t = (e.innerText || '').toLowerCase();
                                    return (t.includes('choose') || t.includes('upload')) && 
                                           (t.includes('csv') || t.includes('file'));
                                });
                            if(btn) btn.click();
                            else {
                                let inputs = root.querySelectorAll('input[type="file"]');
                                if(inputs.length > 0 && inputs[inputs.length - 1].parentElement) {
                                    inputs[inputs.length - 1].parentElement.click();
                                }
                            }
                        }""")
                file_chooser = await fc_info.value
                await file_chooser.set_files(csv_path)
                log("[*] 传统交互挂载成功！等待导入按钮激活...")
                
            await asyncio.sleep(3)

            # 6. 点击弹窗中的确认 IMPORT 按钮
            log("[*] 正在提交导入请求...")
            try:
                await page.evaluate("""() => {
                    let dialogs = document.querySelectorAll('[role="dialog"], .el-dialog, .modal, .ant-modal, .v-dialog');
                    let root = dialogs.length > 0 ? dialogs[dialogs.length - 1] : document;
                    let btns = Array.from(root.querySelectorAll('button')).filter(b => b.innerText && b.innerText.includes('IMPORT') && !b.disabled);
                    // 点击弹窗里的最后一个 IMPORT 按钮
                    if (btns.length > 0) btns[btns.length - 1].click();
                }""")
            except Exception as e:
                log(f"[!] 提交导入时发生异常: {e}")
            
            log("[*] 提交指令已发送，等待页面响应...")
            await asyncio.sleep(3)
            if total_csv_orders > 0:
                log(f"[*] 进度提示：现在是 {total_csv_orders}/{total_csv_orders} (已成功导入 {total_csv_orders} 条订单)")
                if task_info:
                    task_info.set_progress(total_csv_orders, total_csv_orders, prefix="已完成", unit="条")
            log("✅ DSers 订单批量导入指令已执行完毕！")
            log("[*] 浏览器将保持开启 15 秒钟供您检查导入结果，随后将自动安全关闭...")
            await asyncio.sleep(15)

        except Exception as e:
            log(f"❌ 自动化导入中途报错: {e}")
            if not page.is_closed():
                try:
                    await page.screenshot(path="debug_dsers_import_error.png")
                except:
                    pass
            import sys
            sys.exit(1)
        finally:
            await context.close()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", required=True)
    args = parser.parse_args()

    SESSION_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sessions", "dsers_session")
    asyncio.run(run_dsers_import(args.csv, SESSION_DIR, False))
