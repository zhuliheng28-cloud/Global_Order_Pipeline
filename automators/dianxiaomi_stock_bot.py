import asyncio
import os
import sys
import json
import yaml
import time
from typing import List, Dict, Any, Optional, Callable
from playwright.async_api import async_playwright

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(BASE_DIR)

from config import DIANXIAOMI_SESSION_DIR, DIANXIAOMI_SHOPS_CACHE

CONFIG_YAML_PATH = os.path.join(BASE_DIR, "config.yaml")

def get_dxm_credentials():
    username = "396107072"
    password = "Yy19841027"
    url = "https://www.dianxiaomi.com/web/smt/smtProductList/online"
    if os.path.exists(CONFIG_YAML_PATH):
        try:
            with open(CONFIG_YAML_PATH, "r", encoding="utf-8") as f:
                cfg = yaml.safe_load(f)
                if cfg and "dianxiaomi" in cfg:
                    dxm = cfg["dianxiaomi"]
                    username = str(dxm.get("username", username))
                    password = str(dxm.get("password", password))
                    url = str(dxm.get("url", url))
        except Exception:
            pass
    return username, password, url

def cleanup_singleton_locks(user_data_dir: str):
    for item in ["SingletonLock", "SingletonCookie", "SingletonSocket"]:
        p = os.path.join(user_data_dir, item)
        if os.path.exists(p) or os.path.islink(p):
            try:
                os.remove(p)
            except Exception:
                pass

def get_worker_user_data_dir(base_dir: str = DIANXIAOMI_SESSION_DIR, worker_id: Optional[str] = None) -> str:
    if not worker_id or worker_id == "0":
        return base_dir
    worker_dir = f"{base_dir}_worker_{worker_id}"
    if not os.path.exists(worker_dir) and os.path.exists(base_dir):
        try:
            import shutil
            shutil.copytree(base_dir, worker_dir, dirs_exist_ok=True, ignore=shutil.ignore_patterns("Singleton*"))
        except Exception:
            pass
    return worker_dir

# 默认预置店铺列表（优先读取缓存）
DEFAULT_SHOPS = [
    {"shortName": "k店", "text": "cn1099827039knrae(母婴玩具-备用号-小伙堆)", "code": "cn1099827039knrae", "firstLetter": "k"},
    {"shortName": "o店", "text": "cn1089891100oplae (257玩具店)-钰晨", "code": "cn1089891100oplae", "firstLetter": "o"},
    {"shortName": "t店", "text": "cn1072320371teoae（消费电子,殊禾商贸）", "code": "cn1072320371teoae", "firstLetter": "t"},
    {"shortName": "a店", "text": "cn1081563145avgae 美容个护(瑾昕贸易)", "code": "cn1081563145avgae", "firstLetter": "a"},
    {"shortName": "u店", "text": "cn1091787220usdae（珠宝饰品手表，小水滴跨境）", "code": "cn1091787220usdae", "firstLetter": "u"},
    {"shortName": "p店", "text": "cn1092337914psoae(家装家居工具,钰晨跨境)", "code": "cn1092337914psoae", "firstLetter": "p"},
    {"shortName": "f店", "text": "cn1094139016fpcae(美容个护  特货，小水滴）", "code": "cn1094139016fpcae", "firstLetter": "f"},
    {"shortName": "r店", "text": "cn1095721177ryhae(箱包鞋类，小水滴）", "code": "cn1095721177ryhae", "firstLetter": "r"},
    {"shortName": "v店", "text": "cn1098506021vtqae(服装服饰，鞋类--小伙堆)", "code": "cn1098506021vtqae", "firstLetter": "v"},
    {"shortName": "r店", "text": "cn1099851036riqae-个护美容特货-备用号-小伙堆", "code": "cn1099851036riqae", "firstLetter": "r"},
    {"shortName": "s店", "text": "cn1099875036sbnae-服装服饰,鞋类备用号--小伙堆", "code": "cn1099875036sbnae", "firstLetter": "s"},
    {"shortName": "q店", "text": "cn1101073048qmjae-3C-消费电子-备用号已封店-小伙堆", "code": "cn1101073048qmjae", "firstLetter": "q"},
    {"shortName": "n店", "text": "cn1101145038nmmae-家居家具家装灯具工具-备用号--小伙堆", "code": "cn1101145038nmmae", "firstLetter": "n"},
    {"shortName": "a店", "text": "cn1517993896atoq(已退店)", "code": "cn1517993896atoq", "firstLetter": "a"},
    {"shortName": "t店", "text": "cn1521843766tosa(已退店，有剩余资金,冻结至26年3月26）", "code": "cn1521843766tosa", "firstLetter": "t"},
    {"shortName": "p店", "text": "cn1535788279pgob(翊涵科技)", "code": "cn1535788279pgob", "firstLetter": "p"}
]

def load_cached_shops() -> List[Dict[str, str]]:
    if os.path.exists(DIANXIAOMI_SHOPS_CACHE):
        try:
            with open(DIANXIAOMI_SHOPS_CACHE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
        except Exception:
            pass
    return DEFAULT_SHOPS

def save_cached_shops(shops: List[Dict[str, str]]):
    try:
        os.makedirs(os.path.dirname(DIANXIAOMI_SHOPS_CACHE), exist_ok=True)
        with open(DIANXIAOMI_SHOPS_CACHE, "w", encoding="utf-8") as f:
            json.dump(shops, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

async def _ensure_logged_in(page, context, username, password, target_url, log, task_info=None):
    # 先等待几秒检测是否已在在线页面或跳转到登录页
    for _ in range(6):
        if task_info:
            await task_info.async_check_pause()
        curr_url = page.url.lower()
        if "index.htm" in curr_url or "login" in curr_url:
            break
        if await page.query_selector("#exampleInputName, #loginBtn, input[name='username']"):
            break
        if await page.query_selector(".d-tag-group-item, .d-tag-group-item__inner, .vxe-table"):
            break
        await asyncio.sleep(1)

    curr_url = page.url.lower()
    login_input = await page.query_selector("#exampleInputName, input[name='username']")
    if "index.htm" in curr_url or "login" in curr_url or login_input:
        log("🔑 检测到店小秘登录页面，自动填充账号与密码...")
        try:
            name_input = await page.query_selector("#exampleInputName, input[name='username']")
            if name_input:
                await name_input.fill(username)
                log(f"[*] 已填入账号: {username}")

            pwd_input = await page.query_selector("#exampleInputPassword, input[name='password']")
            if pwd_input:
                await pwd_input.fill(password)
                log(f"[*] 已填入密码: {'*'*len(password)}")

            remember = await page.query_selector("input[name='remeber']")
            if remember and not await remember.is_checked():
                await remember.check()

            try:
                import ddddocr
                verify_img = await page.query_selector("#verifyImgCode")
                if verify_img:
                    img_bytes = await verify_img.screenshot()
                    ocr = ddddocr.DdddOcr(show_ad=False)
                    code = ocr.classification(img_bytes)
                    log(f"[*] OCR 识别验证码: {code}")
                    code_input = await page.query_selector("#verifyCode")
                    if code_input:
                        await code_input.fill(code)
            except Exception:
                pass

            login_btn = await page.query_selector("#loginBtn")
            if login_btn:
                await login_btn.click()

        except Exception as e:
            log(f"[!] 自动填写表单提示: {e}")

        log("👉 如果出现滑块或短信验证码，请在浏览器中手动完成，系统将自动检测登录...")
        while True:
            if task_info:
                await task_info.async_check_pause()
            curr = page.url.lower()
            if "smtproductlist/online" in curr:
                log("🎉 登录成功，已进入速卖通在线商品列表！")
                break
            elif "index.htm" not in curr and "login" not in curr and ("smt" in curr or "web" in curr or "home" in curr):
                log("🎉 登录成功，正在导航至速卖通在线商品列表...")
                await page.goto(target_url, wait_until='domcontentloaded', timeout=60000)
                break
            await asyncio.sleep(1)

    if "smtproductlist/online" not in page.url.lower():
        await page.goto(target_url, wait_until='domcontentloaded', timeout=60000)
    await asyncio.sleep(3)


async def fetch_dianxiaomi_shops(user_data_dir: str = DIANXIAOMI_SESSION_DIR, headless: bool = False, progress_callback: Optional[Callable[[str], None]] = None) -> List[Dict[str, str]]:
    def log(msg):
        print(msg, flush=True)
        if progress_callback:
            progress_callback(msg)

    username, password, target_url = get_dxm_credentials()
    cleanup_singleton_locks(user_data_dir)

    log("🌐 正在连接店小秘提取最新店铺列表...")
    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=headless,
            viewport={'width': 1440, 'height': 900},
            args=['--disable-blink-features=AutomationControlled', '--no-sandbox']
        )
        page = context.pages[0] if len(context.pages) > 0 else await context.new_page()

        await page.goto(target_url, wait_until='domcontentloaded', timeout=60000)
        await _ensure_logged_in(page, context, username, password, target_url, log)

        log("[*] 正在解析店铺标签...")
        await asyncio.sleep(2)

        raw_shops = await page.evaluate('''() => {
            const list = [];
            const elements = Array.from(document.querySelectorAll('.d-tag-group-item, .d-tag-group-item__inner, span, a'));
            for (const el of elements) {
                const text = el.textContent.trim();
                const match = text.match(/(cn\\d+([a-zA-Z]+))(.*)/);
                if (match) {
                    const fullCode = match[1];
                    const letters = match[2];
                    const firstLetter = letters ? letters[0].toLowerCase() : '';
                    const extra = match[3] || '';
                    list.push({
                        text: text,
                        code: fullCode,
                        firstLetter: firstLetter,
                        shortName: firstLetter ? `${firstLetter}店` : '其他',
                        extra: extra
                    });
                }
            }
            return list;
        }''')

        seen = set()
        unique_shops = []
        for s in raw_shops:
            if s['code'] not in seen:
                seen.add(s['code'])
                unique_shops.append(s)

        # 排序：k店和o店置顶
        def sort_key(s):
            fl = s['firstLetter']
            if fl == 'k':
                return (0, s['text'])
            elif fl == 'o':
                return (1, s['text'])
            else:
                return (2, s['text'])

        sorted_shops = sorted(unique_shops, key=sort_key)
        if sorted_shops:
            save_cached_shops(sorted_shops)
            log(f"✅ 成功刷新并保存 {len(sorted_shops)} 个店铺。")
            return sorted_shops
        else:
            log("⚠️ 未能提取到店铺，返回缓存列表。")
            return load_cached_shops()


async def dismiss_announcement_popup(page, log=None):
    """
    自动检测并关闭店小秘首页及列表页弹出的营销/活动公告弹窗（如“线下活动”、“小秘公告”等）
    """
    try:
        # 1. 优先使用 Playwright 原生 Locator 定位包含活动公告特征的弹窗
        modal_loc = page.locator('.modal, .modal-dialog, .ant-modal, .ant-modal-content, .d-modal, .el-dialog, [role="dialog"], div').filter(has_text="线下活动").filter(has_text="小秘公告")
        if await modal_loc.count() > 0 and await modal_loc.first.is_visible():
            close_btn = modal_loc.first.locator('button:has-text("关闭"), a:has-text("关闭"), .close, .ant-modal-close, .ant-modal-close-x').filter(state="visible").first
            if await close_btn.count() > 0:
                await close_btn.click(force=True)
                if log: log("🔔 检测到小秘公告/线下活动弹窗，已自动点击【关闭】！")
                await asyncio.sleep(1)
                return
    except Exception:
        pass

    try:
        # 2. 深度 DOM 探测与安全关闭
        closed = await page.evaluate('''() => {
            let closedAny = false;
            const dialogs = Array.from(document.querySelectorAll('.modal, .modal-dialog, .ant-modal, .ant-modal-content, .d-modal, .el-dialog, [role="dialog"], div[class*="dialog"], div[class*="modal"]'));
            for (const d of dialogs) {
                const t = (d.innerText || '');
                if (t.includes('线下活动') || t.includes('小秘公告') || t.includes('美客多拉美招商') || t.includes('拉美蓝海') || t.includes('活动倒计时')) {
                    const closeBtn = Array.from(d.querySelectorAll('button, a, span')).find(b => (b.innerText || '').trim() === '关闭');
                    if (closeBtn) {
                        closeBtn.click();
                        closedAny = true;
                        continue;
                    }
                    const closeX = d.querySelector('.close, .ant-modal-close, .ant-modal-close-x, [aria-label*="close" i], [class*="close"]');
                    if (closeX) {
                        closeX.click();
                        closedAny = true;
                        continue;
                    }
                    d.style.display = 'none';
                    closedAny = true;
                }
            }
            return closedAny;
        }''')
        if closed:
            if log: log("🔔 检测到小秘公告/活动推广弹窗，已自动关闭！")
            await asyncio.sleep(1)
    except Exception:
        pass


async def run_dianxiaomi_stock_update(
    shop_code: str,
    target_stock: int,
    user_data_dir: str = DIANXIAOMI_SESSION_DIR,
    worker_id: Optional[str] = None,
    headless: bool = False,
    progress_callback: Optional[Callable[[str], None]] = None,
    task_info: Optional[Any] = None
) -> Dict[str, Any]:
    """
    独立执行店小秘速卖通库存批量修改
    """
    def log(msg):
        print(msg, flush=True)
        if task_info:
            task_info.check_pause()
        if progress_callback:
            progress_callback(msg)

    actual_user_data_dir = get_worker_user_data_dir(user_data_dir, worker_id)
    username, password, target_url = get_dxm_credentials()
    cleanup_singleton_locks(actual_user_data_dir)

    log("==================================================")
    log(f"启动店小秘速卖通库存修改: 店铺[{shop_code}] -> 目标库存[{target_stock}]")
    log("==================================================")

    async with async_playwright() as p:
        log("[*] 启动 Chromium 独立执行环境...")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=actual_user_data_dir,
            headless=headless,
            viewport={'width': 1440, 'height': 900},
            args=['--disable-blink-features=AutomationControlled', '--no-sandbox']
        )
        if task_info:
            task_info.register_context(context)
        page = context.pages[0] if len(context.pages) > 0 else await context.new_page()

        log(f"[*] 导航至速卖通在线列表: {target_url}")
        await page.goto(target_url, wait_until='domcontentloaded', timeout=60000)
        await _ensure_logged_in(page, context, username, password, target_url, log, task_info=task_info)
        await dismiss_announcement_popup(page, log)

        # 1. 点击目标店铺（最多等待 15 秒供 Vue 动态店铺列表加载完成）
        log(f"[*] [步骤 1/7] 定位并点击店铺 [{shop_code}]...")
        clicked_shop = False
        click_text = ""
        for retry_idx in range(15):
            if task_info:
                await task_info.async_check_pause()
            await dismiss_announcement_popup(page, log)
            click_res = await page.evaluate('''(code) => {
                const items = Array.from(document.querySelectorAll('.d-tag-group-item, .d-tag-group-item__inner, span')).filter(el => {
                    return el.textContent.includes(code) && (el.classList.contains('d-tag-group-item') || el.classList.contains('d-tag-group-item__inner') || el.tagName === 'SPAN');
                });
                if (items.length > 0) {
                    items[0].click();
                    return { success: true, text: items[0].textContent.trim() };
                }
                return { success: false };
            }''', shop_code)

            if click_res.get('success'):
                clicked_shop = True
                click_text = click_res.get('text', shop_code)
                break
            await asyncio.sleep(1)

        if not clicked_shop:
            raise RuntimeError(f"未在页面找到店铺代号【{shop_code}】对应的选项卡，请检查店铺代号是否正确！")
        log(f"✅ 已选中店铺: {click_text}")
        await asyncio.sleep(3)
        await dismiss_announcement_popup(page, log)

        # 2. 切换分页为 300条/页
        log("[*] [步骤 2/7] 切换分页条数为 300条/页...")
        pager_input = await page.query_selector(".vxe-pager--sizes input.vxe-input--inner, .vxe-pager--sizes")
        if pager_input:
            await pager_input.click()
            await asyncio.sleep(1)
            await page.evaluate('''() => {
                const opt = Array.from(document.querySelectorAll('.vxe-select-option, .vxe-select--panel-option, div, li, span')).find(e => e.textContent.trim() === '300条/页');
                if (opt) opt.click();
            }''')
            log("✅ 已选择 300条/页，等待列表数据刷新...")
            await asyncio.sleep(5)
        else:
            log("⚠️ 未找到分页选择器，使用当前默认条数继续...")
        await dismiss_announcement_popup(page, log)

        # 3. 全选在线商品
        log("[*] [步骤 3/7] 点击表头复选框进行全选...")
        th_chk = await page.query_selector(".vxe-header--column.col--checkbox .ant-checkbox-wrapper, .vxe-header--column.col--checkbox input")
        if th_chk:
            await th_chk.click()
            log("✅ 已勾选表头全选")
            await asyncio.sleep(1)
        else:
            await page.evaluate('''() => {
                const chk = document.querySelector('th .ant-checkbox-input, th input[type="checkbox"]');
                if (chk) chk.click();
            }''')
            log("✅ 已通过备用选择器勾选全选")
            await asyncio.sleep(1)

        # 4. 批量操作 -> 批量编辑
        log("[*] [步骤 4/7] 鼠标悬停【批量操作】并触发【批量编辑】...")
        batch_btn = await page.query_selector("button:has-text('批量操作')")
        if batch_btn:
            await batch_btn.hover()
            await asyncio.sleep(1)

        async with context.expect_page(timeout=20000) as new_page_info:
            log("[*] 点击【批量编辑】菜单项...")
            await page.evaluate('''() => {
                const el = Array.from(document.querySelectorAll('*')).find(e => e.textContent.trim() === '批量编辑');
                if (el) el.click();
            }''')

        edit_page = await new_page_info.value
        log(f"🎉 成功打开批量编辑页面: {edit_page.url}")
        await edit_page.wait_for_load_state('domcontentloaded')
        
        log("[*] 等待批量编辑表格渲染...")
        await asyncio.sleep(8)

        # 获取待修改商品总数
        product_count_res = await edit_page.evaluate('''() => {
            const rows = document.querySelectorAll('table tbody tr');
            return rows ? rows.length : 0;
        }''')
        if product_count_res and product_count_res > 0:
            log(f"[*] 进度提示：已载入 {product_count_res} 件待修改商品 (共需修改 {product_count_res} 件)")
            if task_info:
                task_info.set_progress(0, product_count_res, prefix="修改库存", unit="件")

        # 5. 点击库存数量下方的修改按钮 (通过两行表头精确对齐【库存数量】列)
        log("[*] [步骤 5/7] 定位【库存数量】列下方的【修改】图标...")
        clicked_pencil_success = False
        for p_retry in range(25):
            if task_info:
                await task_info.async_check_pause()
            click_pencil = await edit_page.evaluate('''() => {
                const theadRows = Array.from(document.querySelectorAll('table thead tr, .vxe-table--header-wrapper tr, .el-table__header-wrapper tr, table tr'));
                
                // 1. 寻找包含“库存数量”或“库存”的标题行及列索引
                let titleRow = null;
                let stockColIndex = -1;
                let stockHeaderText = '';
                
                for (let r = 0; r < theadRows.length; r++) {
                    const cells = Array.from(theadRows[r].children);
                    for (let c = 0; c < cells.length; c++) {
                        const txt = (cells[c].innerText || cells[c].textContent || '').trim();
                        const isStock = (txt === '库存数量' || txt === '库存' || txt.includes('库存') || txt.includes('可售')) && !txt.includes('SKU分类');
                        const isOther = txt.includes('标题') || txt.includes('名称') || txt.includes('价格') || txt.includes('运费') || txt.includes('重量') || txt.includes('零售价');
                        if (isStock && !isOther) {
                            titleRow = theadRows[r];
                            stockColIndex = c;
                            stockHeaderText = txt;
                            break;
                        }
                    }
                    if (stockColIndex !== -1) break;
                }

                // 2. 若找到库存列索引，查找对应位置的编辑图标
                if (titleRow && stockColIndex !== -1) {
                    const parentTable = titleRow.closest('table, .vxe-table, .el-table') || document;
                    const allRows = Array.from(parentTable.querySelectorAll('tr'));
                    const titleRowIndex = allRows.indexOf(titleRow);
                    
                    // 2.1 先在下一行（即操作图标行）同列查找
                    if (titleRowIndex !== -1 && titleRowIndex + 1 < allRows.length) {
                        const actionRow = allRows[titleRowIndex + 1];
                        const actionCells = Array.from(actionRow.children);
                        if (stockColIndex < actionCells.length) {
                            const targetActionCell = actionCells[stockColIndex];
                            // 在该操作单元格中点击第一个编辑图标（排除还原和删除图标）
                            const editIcon = targetActionCell.querySelector('.action-icon-edit, i.anticon-edit, i[class*="edit"], a[class*="edit"], svg, span[class*="edit"], i, a, span');
                            if (editIcon && editIcon.offsetParent !== null) {
                                editIcon.click();
                                return { success: true, method: 'action-row-aligned', col: stockColIndex, headerText: stockHeaderText };
                            }
                        }
                    }

                    // 2.2 检查标题行单元格内部是否直接包含编辑图标
                    const titleCell = titleRow.children[stockColIndex];
                    const directIcon = titleCell.querySelector('.action-icon-edit, i.anticon-edit, i[class*="edit"], a[class*="edit"], svg, .batch-edit-icon, button');
                    if (directIcon && directIcon.offsetParent !== null) {
                        directIcon.click();
                        return { success: true, method: 'title-cell-direct', col: stockColIndex, headerText: stockHeaderText };
                    }
                }

                // 3. 兜底策略：全表格遍历，按列文本精确匹配并寻找同列图标
                for (const table of Array.from(document.querySelectorAll('table'))) {
                    const rows = Array.from(table.querySelectorAll('tr'));
                    let foundCol = -1;
                    let foundHeader = '';
                    for (const row of rows) {
                        const cells = Array.from(row.children);
                        for (let i = 0; i < cells.length; i++) {
                            const t = (cells[i].innerText || '').trim();
                            if ((t === '库存数量' || t === '库存' || t === '可售数量') && !t.includes('标题') && !t.includes('价格')) {
                                foundCol = i;
                                foundHeader = t;
                                break;
                            }
                        }
                        if (foundCol !== -1) break;
                    }
                    if (foundCol !== -1) {
                        for (const row of rows) {
                            const cells = Array.from(row.children);
                            if (cells.length > foundCol) {
                                const c = cells[foundCol];
                                const icon = c.querySelector('.action-icon-edit, i, svg, a, span, button');
                                if (icon && icon.offsetParent !== null && (c.querySelectorAll('i, a, svg, span').length >= 1 || icon.className.includes('edit'))) {
                                    icon.click();
                                    return { success: true, method: 'table-column-scan', col: foundCol, headerText: foundHeader };
                                }
                            }
                        }
                    }
                }

                return { success: false };
            }''')

            if click_pencil.get('success'):
                clicked_pencil_success = True
                log(f"✅ 已成功定位并点击【{click_pencil.get('headerText', '库存数量')}】第 {click_pencil.get('col', 0) + 1} 列修改图标 (方式: {click_pencil.get('method')})")
                break
            await asyncio.sleep(1)

        if not clicked_pencil_success:
            raise RuntimeError("未在批量编辑页面中找到【库存数量】修改按钮，请检查页面是否加载完成！")
        await asyncio.sleep(2)

        # 6. 弹窗选择“直接修改为”，输入 target_stock，点击“确定”
        log(f"[*] [步骤 6/7] 弹窗中选择【直接修改为】，填入数值【{target_stock}】并点击【确定】...")
        modal_confirmed = False
        for m_retry in range(15):
            if task_info:
                await task_info.async_check_pause()
            modal_res = await edit_page.evaluate('''(stockVal) => {
                const modals = Array.from(document.querySelectorAll('.ant-modal, .d-modal, .el-dialog, [role="dialog"], .modal-dialog, .ui-dialog, div[class*="modal"], div[class*="dialog"]'));
                const activeModal = modals.find(m => m.offsetParent !== null && m.clientHeight > 50) || document;

                // 安全防线：若误打开了标题弹窗，立即关闭并拒绝修改
                const modalTitle = (activeModal.querySelector('.ant-modal-title, .el-dialog__title, .modal-title, .ui-dialog-title') || activeModal).innerText || '';
                if (modalTitle.includes('标题') || modalTitle.includes('商品名')) {
                    const closeBtn = activeModal.querySelector('.ant-modal-close, .el-dialog__close, button[aria-label="Close"], .ui-dialog-titlebar-close');
                    if (closeBtn) closeBtn.click();
                    return { success: false, is_wrong_modal: true };
                }

                const radioLabels = Array.from(activeModal.querySelectorAll('label, .ant-radio-wrapper, .el-radio, span, input[type="radio"]')).filter(el => {
                    const txt = el.innerText || el.textContent || '';
                    return txt.includes('直接修改为') || txt.includes('修改为') || (el.value && el.value.includes('modify'));
                });
                if (radioLabels.length > 0) {
                    radioLabels[0].click();
                }

                const inputs = Array.from(activeModal.querySelectorAll('input[type="text"], input[type="number"], input.ant-input-number-input, input.ant-input, input'));
                for (const inp of inputs) {
                    if (inp.type !== 'radio' && inp.type !== 'checkbox' && inp.offsetParent !== null && inp.readOnly !== true) {
                        inp.focus();
                        inp.value = stockVal;
                        inp.dispatchEvent(new Event('input', { bubbles: true }));
                        inp.dispatchEvent(new Event('change', { bubbles: true }));
                        break;
                    }
                }

                const confirmBtns = Array.from(activeModal.querySelectorAll('button, .ant-btn-primary, .el-button--primary, a, input[type="button"]')).filter(el => {
                    const t = (el.innerText || el.textContent || el.value || '').trim();
                    return t === '确定' || t === '确认' || t === '保存' || t === 'OK';
                });
                if (confirmBtns.length > 0) {
                    confirmBtns[0].click();
                    return { success: true };
                }
                return { success: false };
            }''', target_stock)

            if modal_res.get('success'):
                modal_confirmed = True
                break
            if modal_res.get('is_wrong_modal'):
                log("[!] 检测到非库存弹窗，已自动关闭防误触，正在重新定位库存列...")
                await asyncio.sleep(1)
            await asyncio.sleep(1)

        if not modal_confirmed:
            raise RuntimeError("未能成功确认库存修改弹窗！")
        log(f"✅ 弹窗已确认，库存修改为 {target_stock}")
        await asyncio.sleep(3)

        # 7. 点击右上角的【保存】/【更新】按钮
        log("[*] [步骤 7/7] 点击右上角【保存】/【更新】按钮...")
        update_clicked = False
        for u_retry in range(15):
            if task_info:
                await task_info.async_check_pause()
            update_btn_res = await edit_page.evaluate('''() => {
                const btns = Array.from(document.querySelectorAll('button, a, .ant-btn, .btn, span, input[type="button"]')).filter(el => {
                    const t = (el.innerText || el.textContent || el.value || '').trim();
                    return (t === '保存' || t === '更新' || t === '批量更新' || t === '保存更新' || t === '保存并更新') && el.offsetParent !== null;
                });
                if (btns.length > 0) {
                    btns[0].click();
                    return { success: true, btnText: (btns[0].innerText || btns[0].textContent || '').trim() };
                }
                return { success: false };
            }''')
            if update_btn_res.get('success'):
                update_clicked = True
                log(f"✅ 已成功点击【{update_btn_res.get('btnText', '保存')}】按钮")
                break
            await asyncio.sleep(1)

        if not update_clicked:
            raise RuntimeError("未能找到批量编辑页右上角【保存】或【更新】按钮！")
        
        # 7.1 处理可能弹出的二次保存确认弹窗
        await asyncio.sleep(2)
        await edit_page.evaluate('''() => {
            const confirmBtns = Array.from(document.querySelectorAll('.ant-modal button, .el-dialog button, .d-modal button, .modal button, .ui-dialog button')).filter(btn => {
                const t = (btn.innerText || btn.textContent || '').trim();
                return (t === '确定' || t === '确认' || t === '是') && btn.offsetParent !== null;
            });
            if (confirmBtns.length > 0) {
                confirmBtns[0].click();
            }
        }''')

        log("✅ 已点击【保存/更新】按钮，正在监控系统响应...")

        # 8. 持续监控并等待全部商品库存更新彻底完成（最长等待 10 分钟）
        log("[*] 正在等待全部商品批量更新完成...")
        start_time = time.time()
        max_wait_seconds = 600
        is_completed = False
        captured_errors = []

        while time.time() - start_time < max_wait_seconds:
            if task_info:
                await task_info.async_check_pause()
            await asyncio.sleep(2)
            elapsed = int(time.time() - start_time)

            if edit_page.is_closed():
                log("✅ 批量编辑页面已自动完成并关闭。")
                is_completed = True
                break

            curr_url = edit_page.url.lower()
            if "smtproductlist/online" in curr_url:
                log("✅ 页面已自动返回速卖通商品列表，批量更新全部完成！")
                is_completed = True
                break

            status_info = await edit_page.evaluate('''() => {
                let isFinished = false;
                let progressSummary = "";
                let currCount = 0;
                let totalCount = 0;
                const topErrors = [];

                // 1. 专门识别【批量修改产品】主进度弹窗（正常流程：进行中... -> 已完成!）
                const modals = Array.from(document.querySelectorAll('.ant-modal, .d-modal, .el-dialog, [role="dialog"], .modal-dialog'));
                for (const m of modals) {
                    if (m.offsetParent !== null || m.clientHeight > 0) {
                        const txt = m.textContent.trim();
                        if (txt.includes('批量修改产品') || txt.includes('状态:') || txt.includes('状态：') || txt.includes('详情:') || txt.includes('详情：') || txt.includes('修改中') || txt.includes('进行中') || txt.includes('更新中')) {
                            progressSummary = txt.replace(/\\s+/g, ' ');

                            // 格式 1: 匹配店小秘详情 "已成功修改 7 个产品， 失败了 0 个"
                            const mSucc = txt.match(/(?:已成功修改|成功修改|成功)\\s*(\\d+)\\s*个/);
                            const mFail = txt.match(/(?:失败了|失败)\\s*(\\d+)\\s*个/);
                            if (mSucc) {
                                const sCount = parseInt(mSucc[1]);
                                const fCount = mFail ? parseInt(mFail[1]) : 0;
                                currCount = sCount + fCount;
                            }

                            // 格式 2: 匹配常规比例 "7/50" 或 "7 / 50"
                            const mRatio = txt.match(/(\\d+)\\s*[/／]\\s*(\\d+)/);
                            if (mRatio) {
                                currCount = parseInt(mRatio[1]);
                                totalCount = parseInt(mRatio[2]);
                            }

                            // 格式 3: 进度条百分比
                            const pBar = m.querySelector('.ant-progress-bg, .el-progress-bar__inner');
                            if (pBar && pBar.style.width) {
                                const pVal = parseFloat(pBar.style.width);
                                if (!isNaN(pVal) && pVal > 0) {
                                    progressSummary += " (" + Math.round(pVal) + "%)";
                                }
                            }

                            if (txt.includes('已完成') || txt.includes('全部完成')) {
                                isFinished = true;
                                const closeBtn = Array.from(m.querySelectorAll('button, a')).find(b => {
                                    const bt = b.textContent.trim();
                                    return bt === '关闭' || bt === '确定' || bt === '我知道了';
                                });
                                if (closeBtn) {
                                    closeBtn.click();
                                }
                            }
                        }
                    }
                }

                // 2. 仅检测真正的“顶部提示条/Toast浮窗”
                const toastEls = Array.from(document.querySelectorAll('.ant-message-warning, .ant-message-error, .el-message--warning, .el-message--error, .ant-notification-warning, .ant-notification-error'));
                for (const a of toastEls) {
                    if (a.offsetParent !== null || a.clientHeight > 0) {
                        const alertText = a.textContent.trim();
                        if (alertText && !alertText.includes('批量修改') && !alertText.includes('进行中') && !topErrors.includes(alertText)) {
                            topErrors.push(alertText);
                        }
                    }
                }

                return {
                    isFinished,
                    progressSummary,
                    currCount,
                    totalCount,
                    topErrors
                };
            }''')

            if status_info.get("topErrors"):
                for err in status_info["topErrors"]:
                    if err not in captured_errors:
                        captured_errors.append(err)
                        log(f"⚠️ 检测到顶部提示: {err}")

            has_active_progress = bool(status_info.get("progressSummary")) or (status_info.get("currCount", 0) > 0)

            if not has_active_progress and elapsed >= 8:
                blocking_keywords = ['请选择', '请填写', '未选择', '必填', '不能为空', '缺少']
                blocking_prompts = [e for e in captured_errors if any(k in e for k in blocking_keywords)]
                if blocking_prompts:
                    blocking_text = "；".join(blocking_prompts)
                    log(f"❌ [阻断拦截] 未能启动更新进度，顶部提示需选择选项：{blocking_text}")
                    raise RuntimeError(f"需配置选项: {blocking_text}")

            total_target = status_info.get("totalCount") or product_count_res or 0
            curr_cnt = status_info.get("currCount", 0)

            if total_target > 0:
                if task_info:
                    task_info.set_progress(curr_cnt, total_target, prefix="修改库存", unit="件")
                log(f"[*] 进度提示：现在是 {curr_cnt}/{total_target} (修改库存: {curr_cnt}/{total_target} 件)")
            elif elapsed % 4 == 0:
                summary = status_info.get("progressSummary", "")
                if summary:
                    clean_s = summary
                    if "状态" in summary:
                        clean_s = summary[summary.find("状态"):].strip()
                        if "关闭" in clean_s:
                            clean_s = clean_s[:clean_s.find("关闭")].strip()
                    log(f"[*] 批量更新进度: {clean_s} (已耗时 {elapsed}s)")
                else:
                    log(f"[*] 正在等待后台保存完成... (已耗时 {elapsed}s)")

            if status_info.get("isFinished"):
                if product_count_res and product_count_res > 0:
                    if task_info:
                        task_info.set_progress(product_count_res, product_count_res, prefix="修改库存完成", unit="件")
                    log(f"[*] 进度提示：现在是 {product_count_res}/{product_count_res} (修改库存完成: {product_count_res}/{product_count_res} 件)")
                log(f"✅ 批量修改已成功全部完成，已自动关闭弹窗！(共耗时 {elapsed} 秒)")
                is_completed = True
                break
                await asyncio.sleep(2)
                break

        if not is_completed:
            if captured_errors:
                err_summary = "；".join(captured_errors)
                raise RuntimeError(f"库存修改未完成，阻断拦截: {err_summary}")
            else:
                raise RuntimeError(f"批量修改库存超时（已等待 {max_wait_seconds} 秒），未能完成！")

        log("🎉 店小秘速卖通库存批量修改执行完毕！")
        return {
            "success": True,
            "shop_code": shop_code,
            "target_stock": target_stock,
            "message": "店小秘速卖通库存批量修改成功完成！"
        }

if __name__ == '__main__':
    # 命令行直接运行支持
    import argparse
    parser = argparse.ArgumentParser(description="店小秘速卖通库存批量修改")
    parser.add_argument("--shop", type=str, default="cn1089891100oplae", help="店铺代号")
    parser.add_argument("--stock", type=int, default=100, help="修改库存数值")
    parser.add_argument("--refresh_shops", action="store_true", help="仅刷新店铺列表")
    args = parser.parse_args()

    if args.refresh_shops:
        shops = asyncio.run(fetch_dianxiaomi_shops())
        print(json.dumps(shops, ensure_ascii=False, indent=2))
    else:
        asyncio.run(run_dianxiaomi_stock_update(args.shop, args.stock))
