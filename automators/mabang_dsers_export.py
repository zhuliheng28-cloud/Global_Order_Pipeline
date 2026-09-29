import asyncio
import os
import yaml
import pandas as pd
from playwright.async_api import async_playwright
import sys

# 导入中心化配置
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from automators.excel_utils import save_df_to_excel
from config import DSERS_TEMPLATE

async def run_mabang_export(user_data_dir: str, days: int = 1, hours: int = 0, sku_val: str = 'code', headless: bool = False, progress_callback=None, task_info=None):
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
            viewport={'width': 1280, 'height': 800},
            accept_downloads=True
        )
        if task_info:
            task_info.register_context(context)
        page = context.pages[0] if len(context.pages) > 0 else await context.new_page()
        
        try:
            # === 第一步：从首页进入并处理自动登录 ===
            log("[*] [步骤 1/5] 正在打开马帮并处理登录...")
            await page.goto("https://901067.private.mabangerp.com/index.htm", wait_until='domcontentloaded', timeout=60000)
            
            # 判断是否需要登录
            await asyncio.sleep(2)
            if await page.locator('#login-but').is_visible():
                log("[*] [步骤 1/5] 检测到未登录状态，开始自动登录...")
                await page.locator('input[name="username"]').first.fill(mb_user)
                await page.locator('input[name="password"]').first.fill(mb_pwd)
                await page.locator('#login-but').click()
                log("[*] [步骤 1/5] 登录已提交，等待页面加载...")
                await page.wait_for_load_state('domcontentloaded')
                await asyncio.sleep(3)
            
            # === 第二步：处理首页的“店铺授权提醒”弹窗 ===
            log("[*] [步骤 2/5] 等待并检查是否有授权提醒弹窗...")
            try:
                await page.wait_for_selector('text="店铺授权提醒"', timeout=5000)
                log("[*] [步骤 2/5] 发现授权提醒弹窗，正在关闭...")
                checkbox = page.locator('text="7天内不再重复提醒"')
                if await checkbox.count() > 0:
                    await checkbox.click()
                await page.locator('.layui-layer-btn0, a:has-text("确认")').first.click()
                log("[*] [步骤 2/5] 已点击确认关闭弹窗。")
                await asyncio.sleep(2)
            except Exception:
                log("[*] [步骤 2/5] 没有发现弹窗，继续操作。")
            
            # === 第三步：导航到订单列表 ===
            log("[*] [步骤 3/5] 正在导航至【订单列表】...")
            await page.goto("https://901067.private.mabangerp.com/index.php?mod=order.list&Order_orderStatus=2", wait_until='domcontentloaded')
            await page.wait_for_selector('span.text.mr5.ml5:has-text("批处理功能")', timeout=30000)
            await asyncio.sleep(3)
            
            # === 第四步：高级搜索 ===
            log("[*] [步骤 4/5] 正在打开高级搜索并配置筛选...")
            try:
                await page.locator('text="高级搜索"').last.click(timeout=5000)
            except Exception:
                await page.evaluate("""() => {
                    let els = Array.from(document.querySelectorAll('a, button, span, div.btn')).filter(e => e.innerText && e.innerText.trim() === '高级搜索' && e.offsetParent !== null);
                    if (els.length > 0) {
                        els[0].click();
                    } else {
                        let btn = document.querySelector('#AdvanceSearchBtn button') || document.querySelector('#AdvanceSearchBtn');
                        if(btn) btn.click();
                    }
                }""")
            await asyncio.sleep(3)
            
            log("[*] 配置高级搜索条件：创建时间、起始时间、客户ID...")
            try:
                from datetime import datetime, timedelta
                now_dt = datetime.now()
                now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
                start_dt_str = (now_dt - timedelta(days=days, hours=hours)).strftime("%Y-%m-%d %H:%M:%S")
                
                await page.evaluate(f"""() => {{
                    // 1. 设置创建时间
                    let qTime = document.querySelector('#searchDetail select[name="queryTime"]');
                    if(qTime) {{
                        qTime.value = 'createDate';
                        qTime.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    
                    // 2. 设置起始和截止时间
                    let startTime = document.querySelector('#searchDetail input[name="startTime1"]');
                    if(startTime) {{
                        startTime.value = '{start_dt_str}';
                        startTime.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    
                    let endTime = document.querySelector('#searchDetail input[name="endTime1"]');
                    if(endTime) {{
                        endTime.value = '{now_str}';
                        endTime.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    
                    // 3. 设置筛选条件: 按库存SKU模糊搜索
                    let fKey = document.querySelector('#searchDetail #fuzzySearchKey');
                    if(fKey && fKey.selectize) {{
                        let opts = fKey.selectize.options;
                        let targetVal = null;
                        for(let k in opts) {{
                            if(opts[k].text.includes('库存SKU模糊搜索')) {{
                                targetVal = opts[k].value;
                                break;
                            }}
                        }}
                        if(targetVal) {{
                            fKey.selectize.setValue(targetVal);
                        }} else {{
                            fKey.selectize.setValue('b.sku');
                        }}
                    }}
                    
                    // 3.5 设置匹配逻辑为：不包含
                    let selects = Array.from(document.querySelectorAll('#searchDetail select'));
                    let conditionSelect = null;
                    let notContainVal = null;
                    for (let sel of selects) {{
                        for (let opt of Array.from(sel.options)) {{
                            if (opt.text && opt.text.includes('不包含')) {{
                                conditionSelect = sel;
                                notContainVal = opt.value;
                                break;
                            }}
                        }}
                        if (conditionSelect) break;
                    }}
                    if (conditionSelect && notContainVal) {{
                        conditionSelect.value = notContainVal;
                        conditionSelect.dispatchEvent(new Event('change', {{bubbles: true}}));
                        if (conditionSelect.selectize) {{
                            conditionSelect.selectize.setValue(notContainVal);
                        }}
                    }}
                    
                    // 4. 填入搜索值
                    let fVal = document.querySelector('#searchDetail #fuzzySearchValue');
                    if(fVal) {{
                        fVal.disabled = false;
                        fVal.value = '{sku_val}';
                        fVal.dispatchEvent(new Event('input', {{bubbles: true}}));
                        fVal.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                }}""")
            except Exception as e:
                log(f"[!] 配置筛选条件发生错误: {e}")
            await asyncio.sleep(2)
            
            try:
                if await page.locator('input[placeholder="-查询条件1-"]').count() > 0:
                    input_box = page.locator('input[placeholder="-查询条件1-"]').last
                    if await input_box.is_visible():
                        log("[*] JS可能未生效，使用原生点击选择 '按客户ID'")
                        await input_box.click(force=True)
                        await asyncio.sleep(1)
                        await page.locator('.selectize-dropdown-content div.option:has-text("按客户ID")').last.click(force=True)
                        await asyncio.sleep(1)
            except Exception as e:
                pass
            
            log("[*] 提交高级搜索...")
            try:
                async with page.expect_navigation(timeout=8000):
                    await page.evaluate('''() => {
                        let btn = document.getElementById('searchMore');
                        if(btn) {
                            btn.click();
                        } else {
                            let btns = Array.from(document.querySelectorAll('button'));
                            let b = btns.find(x => x.innerText && x.innerText.trim() === '搜索');
                            if(b) b.click();
                        }
                    }''')
            except Exception:
                pass

            await asyncio.sleep(4)
            
            try:
                await page.evaluate("""() => {
                    let modal = document.querySelector('#AdvanceSearch') || document.querySelector('#searchDetail');
                    if (modal) modal.style.display = 'none';
                    document.querySelectorAll('.modal-backdrop').forEach(el => el.style.display = 'none');
                }""")
            except:
                pass
            await asyncio.sleep(2)
            
            log("[*] 设置每页显示 500 条...")
            try:
                dropdown_btn = page.locator('button:has-text("每页")').last
                if await dropdown_btn.count() > 0:
                    await dropdown_btn.click(force=True)
                    await asyncio.sleep(1)
                    await page.locator('a[onclick*="getPaginationData"]').filter(has_text="500").last.click(force=True)
                    log("[*] 已点击 500 条/页，等待数据重新加载...")
                    await asyncio.sleep(6)
                else:
                    await page.evaluate("if(typeof getPaginationData === 'function') getPaginationData(1,500);")
                    await asyncio.sleep(6)
            except Exception as e:
                log(f"[!] 设置每页 500 条失败: {e}")
            
            # ==========================
            #   多页循环导出逻辑开始
            # ==========================
            all_dfs = []
            page_index = 1
            
            while True:
                log(f"[*] ================= 正在处理第 {page_index} 页订单数据 =================")
                
                # 勾选全选当前页订单
                await page.evaluate("""() => {
                    let checkallBtn = document.getElementById('checkall');
                    if(checkallBtn) {
                        if(!checkallBtn.checked) {
                            checkallBtn.click();
                        }
                    } else {
                        document.querySelectorAll('input.orderCheck').forEach(cb => {
                            if(!cb.checked) cb.click();
                        });
                    }
                }""")
                await asyncio.sleep(1)
                
                log(f"[*] 准备触发第 {page_index} 页【订单导出】...")
                async def trigger_export():
                    await page.evaluate("""() => {
                        if (typeof gotoExportOrderTemplate === 'function') {
                            gotoExportOrderTemplate(1);
                        } else {
                            let link = Array.from(document.querySelectorAll('#upLoadMenu a')).find(a => a.innerText.includes('订单导出'));
                            if(link) link.click();
                        }
                    }""")

                async with context.expect_page() as new_page_info:
                    await trigger_export()
                    
                export_page = await new_page_info.value
                await export_page.wait_for_load_state('domcontentloaded')
                await asyncio.sleep(4)
                
                log("[*] 选择【dsers order】模板...")
                for f in export_page.frames:
                    try:
                        await f.evaluate("""() => {
                            let dropBtn = Array.from(document.querySelectorAll('button, a, div, span')).find(e => e.innerText && e.innerText.includes('加载导出模板'));
                            if(dropBtn) dropBtn.click();
                        }""")
                    except: pass
                await asyncio.sleep(1)
                
                for f in export_page.frames:
                    try:
                        await f.evaluate("""() => {
                            let dsersBtn = Array.from(document.querySelectorAll('li, a, dd, div')).find(e => e.innerText && e.innerText.trim() === 'dsers order');
                            if(dsersBtn) dsersBtn.click();
                        }""")
                    except: pass
                await asyncio.sleep(2)
                
                download_future = asyncio.get_event_loop().create_future()
                def handle_download(d):
                    if not download_future.done():
                        download_future.set_result(d)
                export_page.on("download", handle_download)
                
                log("[*] 正在提交【实时导出】...")
                try:
                    for f in export_page.frames:
                        try:
                            await f.evaluate("""() => {
                                let btns = Array.from(document.querySelectorAll('button, a, input[type="button"], span'));
                                let exportBtn = btns.find(b => b.innerText && b.innerText.includes("实时导出"));
                                if(exportBtn) exportBtn.click();
                            }""")
                        except: pass
                    
                    download = await asyncio.wait_for(download_future, timeout=120.0)
                except Exception as e:
                    log(f"[!] 第 {page_index} 页导出下载超时或失败: {e}")
                    raise e
                    
                download_dir = os.path.dirname(os.path.abspath(__file__))
                temp_path = os.path.join(download_dir, f"temp_dsers_page_{page_index}.xls")
                await download.save_as(temp_path)
                
                log(f"[*] 第 {page_index} 页下载成功，正在读取合并...")
                try:
                    try:
                        df_part = pd.read_excel(temp_path, dtype=str)
                    except:
                        try:
                            df_part = pd.read_csv(temp_path, dtype=str)
                        except:
                            from bs4 import BeautifulSoup
                            with open(temp_path, 'r', encoding='utf-8', errors='replace') as f:
                                soup = BeautifulSoup(f.read(), 'html.parser')
                            tables = soup.find_all('table')
                            rows = []
                            for tr in tables[0].find_all('tr'):
                                rows.append([td.get_text(strip=True) for td in tr.find_all(['th', 'td'])])
                            df_part = pd.DataFrame(rows[1:], columns=rows[0])
                    all_dfs.append(df_part)
                except Exception as e:
                    log(f"[!] 读取第 {page_index} 页表格失败: {e}")
                
                # 关闭导出的标签页
                await export_page.close()
                await asyncio.sleep(1)
                
                # === 检查并点击下一页 ===
                page_status = await page.evaluate("""() => {
                    let activeBtn = document.querySelector('.btn-group button.text-danger') || document.querySelector('.btn-group .active');
                    let currentPage = activeBtn ? parseInt(activeBtn.innerText.trim()) : 1;
                    if (isNaN(currentPage)) currentPage = 1;
                    
                    let nextBtn = document.querySelector('button[data-original-title="下一页"]') || document.querySelector('button[title="下一页"]') || document.querySelector('button:has(i.ico-arrow-right22)');
                    let nextOnclick = nextBtn ? (nextBtn.getAttribute('onclick') || '') : '';
                    let hasNextByBtn = nextBtn && nextOnclick && !nextOnclick.includes('javascript:void(0)') && !nextBtn.className.includes('disabled');
                    
                    let totalMatch = document.body.innerText.match(/共\\s*(\\d+)\\s*条/);
                    let totalOrders = totalMatch ? parseInt(totalMatch[1]) : 0;
                    let totalPages = totalOrders > 0 ? Math.ceil(totalOrders / 500) : 1;
                    let hasNextByTotal = (totalPages > currentPage);

                    let firstOrder = document.querySelector('input.orderCheck') ? document.querySelector('input.orderCheck').value : '';

                    return {
                        currentPage,
                        totalPages,
                        totalOrders,
                        hasNext: Boolean(hasNextByBtn || hasNextByTotal),
                        firstOrder
                    };
                }""")
                
                if page_status['hasNext']:
                    target_page = page_index + 1
                    total_p_str = page_status.get('totalPages', '?')
                    total_o_str = page_status.get('totalOrders', '?')
                    log(f"[*] 检测到存在下一页数据（总计约 {total_o_str} 条，约 {total_p_str} 页），准备翻页至第 {target_page} 页...")
                    old_order = page_status.get('firstOrder', '')
                    
                    await page.evaluate(f"""() => {{
                        let nextBtn = document.querySelector('button[data-original-title="下一页"]') || document.querySelector('button[title="下一页"]') || document.querySelector('button:has(i.ico-arrow-right22)');
                        let numBtn = document.querySelector('button[onclick*="getPaginationData({target_page},"]');
                        if (nextBtn && nextBtn.getAttribute('onclick') && !nextBtn.getAttribute('onclick').includes('javascript:void(0)')) {{
                            nextBtn.click();
                        }} else if (numBtn) {{
                            numBtn.click();
                        }} else if (typeof getPaginationData === 'function') {{
                            getPaginationData({target_page}, 500);
                        }}
                    }}""")
                    
                    # 轮询等待新页面渲染完成
                    page_changed = False
                    for _ in range(25):
                        await asyncio.sleep(1)
                        cur_status = await page.evaluate("""() => {
                            let activeBtn = document.querySelector('.btn-group button.text-danger') || document.querySelector('.btn-group .active');
                            let curPage = activeBtn ? parseInt(activeBtn.innerText.trim()) : 0;
                            let firstOrder = document.querySelector('input.orderCheck') ? document.querySelector('input.orderCheck').value : '';
                            return { curPage, firstOrder };
                        }""")
                        if cur_status['curPage'] == target_page and cur_status['firstOrder'] != old_order:
                            page_changed = True
                            log(f"[*] 页面已成功刷新并定位到第 {target_page} 页！")
                            break
                            
                    if not page_changed:
                        log(f"[!] 警告：翻页等待超时，当前未检测到第 {target_page} 页完全渲染，尝试继续处理。")
                        
                    await asyncio.sleep(3) # 给予表格 DOM 充分的渲染缓冲
                    page_index += 1
                else:
                    log(f"[*] 所有页数据已全部导出完毕（共处理 {page_index} 页）！")
                    break

            # === 保存合并后的大表 ===
            if all_dfs:
                final_df = pd.concat(all_dfs, ignore_index=True)
                final_df.columns = final_df.columns.str.strip()
                log(f"[*] 完美合并完毕，总行数: {len(final_df)}，读取到的列: {list(final_df.columns)}")
                
                # 强制转换为字符串类型，防止超出 15 位的订单号被 Pandas 和 Excel 科学计数法吞噬尾数
                if 'Order_number' in final_df.columns:
                    final_df['Order_number'] = final_df['Order_number'].astype(str)
                if '交易编号' in final_df.columns:
                    final_df['交易编号'] = final_df['交易编号'].astype(str)
                
                output_excel = DSERS_TEMPLATE
                try:
                    save_df_to_excel(final_df, output_excel)
                    log(f"[*] 进度提示：已成功生成 DSers 格式最终模板，共 {len(final_df)} 条订单数据")
                    log(f"[*] 成功生成 DSers 格式最终模板: {output_excel}")
                    # 清理临时分页文件
                    for f_tmp in os.listdir(download_dir):
                        if f_tmp.startswith("temp_dsers_page_") and f_tmp.endswith(".xls"):
                            try: os.remove(os.path.join(download_dir, f_tmp))
                            except Exception: pass
                except Exception as e:
                    log(f"[!] 无法保存表格 (可能表格正被 Excel 占用打开): {e}")
                    raise e
            else:
                log("[!] 未读取到任何有效数据，模板生成失败。")
                raise Exception("未读取到任何有效数据，模板生成失败。请检查是否没有订单数据，或重试。")
                
        except Exception as e:
            log(f"[!] 发生错误: {e}"); sys.exit(1)
            if not page.is_closed():
                try:
                    await page.screenshot(path="debug_export_error.png")
                except:
                    pass
            
        finally:
            await context.close()
            log("[*] 测试结束。")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--days", type=int, default=1)
    parser.add_argument("--hours", type=int, default=0)
    parser.add_argument("--sku", default="code")
    args = parser.parse_args()

    SESSION_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sessions", "mabang_session")
    asyncio.run(run_mabang_export(
        user_data_dir=SESSION_DIR, 
        days=args.days, 
        hours=args.hours, 
        sku_val=args.sku, 
        headless=False
    ))
