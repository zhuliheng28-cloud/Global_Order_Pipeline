import asyncio
import os
import yaml
import pandas as pd
from playwright.async_api import async_playwright
import sys

# 导入中心化配置
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from automators.excel_utils import save_df_to_excel
from config import SCRIPT_TEMPLATE

async def run_mabang_export(user_data_dir: str, days: int = 1, hours: int = 0, customer_id: str = '', sku_filter: str = 'code', headless: bool = False, progress_callback=None, task_info=None):
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
            log("[*] 正在打开马帮首页并处理登录...")
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
                await page.wait_for_selector('text="店铺授权提醒"', timeout=8000)
                log("[*] 发现授权提醒弹窗，正在关闭...")
                checkbox = page.locator('text="7天内不再重复提醒"')
                if await checkbox.count() > 0:
                    await checkbox.click()
                await page.locator('.layui-layer-btn0, a:has-text("确认")').first.click()
                log("[*] 已点击确认关闭弹窗。")
                await asyncio.sleep(2)
            except Exception:
                log("[*] 没有发现弹窗，继续操作。")
            
            # === 第三步：直接导航到订单列表 ===
            log("[*] 导航至【订单列表】...")
            await page.goto("https://901067.private.mabangerp.com/index.php?mod=order.list&Order_orderStatus=2", wait_until='domcontentloaded')
            await asyncio.sleep(3)
            
            # === 第四步：高级搜索 ===
            log("[*] [步骤 4/5] 正在打开高级搜索面板...")
            try:
                # 1. 优先使用 Playwright 原生定位并点击【高级搜索】按钮
                btn = page.locator('#AdvanceSearchBtn button, button:has-text("高级搜索"), a:has-text("高级搜索")').first
                if await btn.count() > 0 and await btn.is_visible():
                    await btn.click(force=True)
                    log("[*] 已点击【高级搜索】按钮。")
                else:
                    await page.evaluate("""() => {
                        if (typeof advanceSearchShow === 'function') {
                            advanceSearchShow();
                        } else {
                            let b = document.querySelector('#AdvanceSearchBtn button') || 
                                    Array.from(document.querySelectorAll('button, a')).find(el => el.innerText && el.innerText.trim().includes('高级搜索'));
                            if (b) b.click();
                        }
                    }""")
                    log("[*] 已触发高级搜索展开函数。")
            except Exception as e:
                log(f"[*] 尝试打开高级搜索面板触发告警: {e}")

            # 2. 严格等待高级搜索弹窗真实可见
            try:
                await page.wait_for_selector('#AdvanceSearch', state='visible', timeout=8000)
                log("[*] 高级搜索面板已成功打开且可见！")
            except Exception:
                log("[!] 等待 #AdvanceSearch 可见超时，尝试调用 advanceSearchShow() 补救...")
                await page.evaluate("if (typeof advanceSearchShow === 'function') advanceSearchShow();")
                await asyncio.sleep(2)
                if await page.locator('#AdvanceSearch').is_visible():
                    log("[*] 补救成功，高级搜索面板已展开。")
                else:
                    log("[!] 警告：未能确认高级搜索面板完全展开，继续尝试执行注入...")
            await asyncio.sleep(2)
            
            from datetime import datetime, timedelta
            now_dt = datetime.now()
            now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
            start_dt_str = (now_dt - timedelta(days=days, hours=hours)).strftime("%Y-%m-%d %H:%M:%S")
            
            log(f"[*] 设置筛选时间范围 (创建时间): {start_dt_str} 至 {now_str}")
            log(f"[*] 筛选条件配置: 客户ID='{customer_id}', 排除SKU='{sku_filter}'")
            
            clean_cust_id = str(customer_id or '').strip()
            clean_sku_f = str(sku_filter or '').strip()

            try:
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
                        startTime.disabled = false;
                        startTime.value = '{start_dt_str}';
                        startTime.dispatchEvent(new Event('input', {{bubbles: true}}));
                        startTime.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    
                    let endTime = document.querySelector('#searchDetail input[name="endTime1"]');
                    if(endTime) {{
                        endTime.disabled = false;
                        endTime.value = '{now_str}';
                        endTime.dispatchEvent(new Event('input', {{bubbles: true}}));
                        endTime.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}
                    
                    // 辅助函数：根据关键词在 selectize options 中查找最佳匹配并设值
                    function setSelectize(selectId, keywords, fallback) {{
                        let el = document.querySelector('#searchDetail ' + selectId);
                        if (!el || !el.selectize) return false;
                        let opts = el.selectize.options;
                        let targetVal = null;
                        for (let k in opts) {{
                            let txt = opts[k].text || '';
                            for (let kw of keywords) {{
                                if (txt.includes(kw)) {{
                                    targetVal = opts[k].value;
                                    break;
                                }}
                            }}
                            if (targetVal) break;
                        }}
                        if (!targetVal && fallback) targetVal = fallback;
                        if (targetVal) {{
                            el.selectize.setValue(targetVal);
                            return true;
                        }}
                        return false;
                    }}

                    // 3. 多条件联合配置
                    let custId = '{clean_cust_id}';
                    let skuF = '{clean_sku_f}';
                    
                    if (custId) {{
                        // 条件1: 买家客户ID 等于 custId
                        setSelectize('#fuzzySearchKey', ['客户ID', '买家ID'], 'a.buyerUserId');
                        let c1 = document.querySelector('select[name="OrderSearch.fuzzySearchConditions1"]');
                        if(c1) {{
                            c1.disabled = false;
                            c1.value = '1'; // 等于
                            c1.dispatchEvent(new Event('change', {{bubbles: true}}));
                        }}
                        let fVal = document.querySelector('#searchDetail #fuzzySearchValue');
                        if(fVal) {{
                            fVal.disabled = false;
                            fVal.value = custId;
                            fVal.dispatchEvent(new Event('input', {{bubbles: true}}));
                            fVal.dispatchEvent(new Event('change', {{bubbles: true}}));
                        }}
                        
                        // 条件2: SKU 不包含 skuF
                        if (skuF) {{
                            setSelectize('#fuzzySearchKey1', ['库存SKU', '商品SKU'], 'b.sku');
                            let c2 = document.querySelector('select[name="OrderSearch.fuzzySearchConditions2"]');
                            if(c2) {{
                                c2.disabled = false;
                                c2.value = '2'; // 不包含
                                c2.dispatchEvent(new Event('change', {{bubbles: true}}));
                            }}
                            let fVal1 = document.querySelector('#searchDetail #fuzzySearchValue1');
                            if(fVal1) {{
                                fVal1.disabled = false;
                                fVal1.value = skuF;
                                fVal1.dispatchEvent(new Event('input', {{bubbles: true}}));
                                fVal1.dispatchEvent(new Event('change', {{bubbles: true}}));
                            }}
                        }}
                    }} else if (skuF) {{
                        // 无客户ID限制，条件1直接配置为 SKU 不包含 skuF
                        setSelectize('#fuzzySearchKey', ['库存SKU', '商品SKU'], 'b.sku');
                        let c1 = document.querySelector('select[name="OrderSearch.fuzzySearchConditions1"]');
                        if(c1) {{
                            c1.disabled = false;
                            c1.value = '2'; // 不包含
                            c1.dispatchEvent(new Event('change', {{bubbles: true}}));
                        }}
                        let fVal = document.querySelector('#searchDetail #fuzzySearchValue');
                        if(fVal) {{
                            fVal.disabled = false;
                            fVal.value = skuF;
                            fVal.dispatchEvent(new Event('input', {{bubbles: true}}));
                            fVal.dispatchEvent(new Event('change', {{bubbles: true}}));
                        }}
                    }}
                }}""")
            except Exception as e:
                log(f"[!] 配置筛选条件发生错误: {e}")
            await asyncio.sleep(2)
            
            # 4. 点击高级搜索面板内的【搜索】(#searchMore)按钮
            log("[*] 提交高级搜索...")
            try:
                search_btn = page.locator('#AdvanceSearch #searchMore, #searchDetail #searchMore').first
                if await search_btn.count() > 0:
                    await search_btn.click(force=True)
                    log("[*] 已点击高级搜索内部【搜索】(#searchMore)。")
                else:
                    await page.evaluate('''() => {
                        let btn = document.querySelector('#AdvanceSearch #searchMore') || document.getElementById('searchMore');
                        if(btn) {
                            btn.click();
                        }
                    }''')
                    log("[*] 已通过 JS 点击 #searchMore。")
            except Exception as e:
                log(f"[!] 点击高级搜索提交按钮异常: {e}")

            await asyncio.sleep(5)
            
            # 隐藏高级搜索面板与遮罩（防止遮挡后续列表操作）
            try:
                await page.evaluate("""() => {
                    let modal = document.querySelector('#AdvanceSearch') || document.querySelector('#searchDetail');
                    if (modal) modal.style.display = 'none';
                    document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
                }""")
            except Exception as e:
                log(f"[*] 隐藏弹窗面板: {e}")
            await asyncio.sleep(2)
            
            # === 第4.5步：设置每页显示500条 ===
            log("[*] 设置每页显示 500 条...")
            try:
                # 尝试点击每页显示数量下拉框
                dropdown_btn = page.locator('button:has-text("每页")').last
                if await dropdown_btn.count() > 0:
                    await dropdown_btn.click(force=True)
                    await asyncio.sleep(1)
                    # 点击 500
                    await page.locator('a[onclick*="getPaginationData"]').filter(has_text="500").last.click(force=True)
                    log("[*] 已点击 500 条/页，等待数据重新加载...")
                    await asyncio.sleep(6)  # 给足够的时间让大量数据加载出来
                else:
                    # Fallback
                    await page.evaluate("if(typeof getPaginationData === 'function') getPaginationData(1,500);")
                    await asyncio.sleep(6)
            except Exception as e:
                log(f"[!] 设置每页 500 条失败: {e}")
            
            # 检查高级搜索后是否有订单数据
            total_orders = await page.evaluate("""() => {
                let totalMatch = document.body.innerText.match(/共\\s*(\\d+)\\s*条/);
                return totalMatch ? parseInt(totalMatch[1]) : 0;
            }""")
            log(f"[*] 高级搜索执行完成，列表当前共检索到 {total_orders} 条订单。")
            if total_orders == 0:
                log("[!] 提示：未检索到符合条件的订单（共 0 条），请确认搜索时间范围或过滤条件！")
                raise Exception("未检索到符合条件的订单（共 0 条），请检查马帮时间范围与筛选配置。")

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
                        checkallBtn.click();
                    }
                    document.querySelectorAll('input[name="item"], input.orderCheck').forEach(cb => {
                        if(!cb.checked) cb.checked = true;
                    });
                }""")
                await asyncio.sleep(1)
                
                # 点击导出菜单
                log(f"[*] 点击【导入/出相关】菜单...")
                await page.locator('#upLoadMenu button').click()
                await asyncio.sleep(1)
                
                log(f"[*] 准备触发第 {page_index} 页【订单导出】...")
                export_link = page.locator('#upLoadMenu a:text-is("订单导出")')
                
                async with context.expect_page() as new_page_info:
                    await export_link.click(force=True)
                    
                export_page = await new_page_info.value
                await export_page.wait_for_load_state('domcontentloaded')
                log(f"[*] 成功进入新标签页，URL: {export_page.url}")
                await asyncio.sleep(2)
                
                # 配置CPF导出弹窗字段
                log("[*] 清空默认选中字段...")
                for f in export_page.frames:
                    try:
                        await f.evaluate("""() => {
                            document.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                                if(cb.parentElement && cb.parentElement.innerText && cb.parentElement.innerText.includes("全选/清空")) return;
                                if(cb.checked) cb.click();
                            });
                        }""")
                    except Exception:
                        pass
                await asyncio.sleep(1)
                
                log("[*] 勾选：订单编号、客户姓名、abnnumber、SKU")
                for f in export_page.frames:
                    try:
                        await f.evaluate("""() => {
                            let fields = ["订单编号", "客户姓名", "abnnumber", "SKU"];
                            let checkboxes = document.querySelectorAll('input[type="checkbox"]');
                            checkboxes.forEach(cb => {
                                let parent = cb.parentElement;
                                if (parent && parent.innerText) {
                                    let txt = parent.innerText.trim();
                                    if (fields.some(f => txt === f)) {
                                        if (!cb.checked) {
                                            cb.click();
                                        }
                                    }
                                }
                            });
                        }""")
                    except Exception:
                        pass
                await asyncio.sleep(1)
                
                download_future = asyncio.get_event_loop().create_future()
                def handle_download(d):
                    if not download_future.done():
                        download_future.set_result(d)
                        
                export_page.on("download", handle_download)
                page.on("download", handle_download)
                context.on("page", lambda p: p.on("download", handle_download))
                
                log("[*] 正在提交【实时导出】...")
                try:
                    for f in export_page.frames:
                        try:
                            await f.evaluate("""() => {
                                let btns = Array.from(document.querySelectorAll('button, a, input[type="button"], span'));
                                let exportBtn = btns.find(b => b.innerText && b.innerText.includes("实时导出"));
                                if(exportBtn) exportBtn.click();
                            }""")
                        except Exception:
                            pass
                    
                    download = await asyncio.wait_for(download_future, timeout=120.0)
                except Exception as e:
                    log(f"[!] 第 {page_index} 页导出下载失败: {e}")
                    raise e
                    
                download_dir = os.path.dirname(os.path.abspath(__file__))
                temp_path = os.path.join(download_dir, f"temp_cpf_page_{page_index}.xls")
                log(f"[*] 正在下载文件至: {temp_path}")
                await download.save_as(temp_path)
                
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
                    log(f"[*] 第 {page_index} 页导出成功，已获取 {len(df_part)} 行数据")
                except Exception as e:
                    log(f"[!] 读取第 {page_index} 页表格失败: {e}")
                
                await export_page.close()
                await asyncio.sleep(2)
                
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
                        
                    await asyncio.sleep(3)
                    page_index += 1
                else:
                    log(f"[*] 所有页数据已全部导出完毕（共处理 {page_index} 页）！")
                    break

            # === 合并并生成模板 ===
            if all_dfs:
                df = pd.concat(all_dfs, ignore_index=True)
                df.columns = df.columns.str.strip()
                log(f"[*] 完美合并完毕，总行数: {len(df)}，读取到的列: {list(df.columns)}")
                
                # 1. 过滤掉无订单号的空白行
                order_col = next((c for c in df.columns if str(c).strip() in ['订单编号', '交易编号', 'Shiopify order number', 'Shopify order number', 'Order_number']), None)
                if not order_col and len(df.columns) > 0:
                    order_col = df.columns[0]
                if order_col:
                    raw_len = len(df)
                    valid_mask = df[order_col].fillna('').astype(str).str.strip().ne('') & ~df[order_col].astype(str).str.strip().str.lower().isin(['nan', 'none'])
                    df = df[valid_mask].copy()
                    if len(df) < raw_len:
                        log(f"[*] 已清理无订单号的空白行: 清理前 {raw_len} 行 -> 清理后 {len(df)} 行")

                # 2. 本地过滤 SKU 包含 code 的行 (双重保障)
                sku_col = next((c for c in df.columns if 'sku' in str(c).lower()), None)
                if sku_col and sku_filter and str(sku_filter).strip():
                    f_val = str(sku_filter).strip().lower()
                    before_sku = len(df)
                    df = df[~df[sku_col].fillna('').astype(str).str.lower().str.contains(f_val)].copy()
                    log(f"[*] 已本地二次过滤排除 SKU 包含 '{sku_filter}' 的订单 (过滤前 {before_sku} 行 -> 过滤后 {len(df)} 行)")

                abn_col = [c for c in df.columns if 'abnnumber' in c.lower()]
                if not abn_col:
                    log("[!] 找不到 abnnumber 列！")
                    raise Exception("导出的表格中缺失关键列 abnnumber，无法生成 CPF 模板！请检查马帮导出字段或检查导出是否为空。")
                else:
                    abn_col = abn_col[0]
                    # 严格按照标准 6 列脚本模板规范组装输出 DataFrame：
                    # 第1列: 订单编号
                    # 第2列: 客户姓名
                    # 第3列: abnnumber
                    # 第4列: cpf1_abn (用于发给 TG 查名，格式: /cpf1 12345678900)
                    # 第5列: 查询结果 (留空，供 TG 机器人回填真实姓名)
                    # 第6列: 出生日期 (留空，供 TG 机器人回填出生日期)
                    name_col = next((c for c in df.columns if str(c).strip() in ['客户姓名', '姓名', 'Contact Name', 'Customer Name']), None)
                    
                    df_out = pd.DataFrame()
                    df_out['订单编号'] = df[order_col].astype(str) if order_col else df.iloc[:, 0].astype(str)
                    df_out['客户姓名'] = df[name_col].astype(str) if name_col else ''
                    df_out['abnnumber'] = df[abn_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    df_out['cpf1_abn'] = '/cpf1 ' + df_out['abnnumber']
                    df_out['查询结果'] = ''
                    df_out['出生日期'] = ''

                    output_excel = SCRIPT_TEMPLATE
                    save_df_to_excel(df_out, output_excel)
                    log(f"[*] 进度提示：已成功生成更新模板，共 {len(df_out)} 条订单数据")
                    log(f"[*] 成功生成更新模板: {output_excel}")
                    log(f"[*] 预览前几行数据:\n{df_out.head(3)}")
                    
                    # 清理临时文件
                    for f_tmp in os.listdir(download_dir):
                        if f_tmp.startswith("temp_cpf_page_") and f_tmp.endswith(".xls"):
                            try: os.remove(os.path.join(download_dir, f_tmp))
                            except Exception: pass
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
    parser.add_argument("--customer_id", default="1000000257")
    args = parser.parse_args()

    SESSION_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sessions", "mabang_session")
    asyncio.run(run_mabang_export(
        user_data_dir=SESSION_DIR, 
        days=args.days, 
        hours=args.hours, 
        customer_id=args.customer_id, 
        headless=False
    ))
