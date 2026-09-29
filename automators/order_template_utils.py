import os
import sys
import re
import pandas as pd

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import SCRIPT_TEMPLATE, ORDER_TEMPLATE
from automators.excel_utils import save_df_to_excel

def extract_shopify_orders(merged_val):
    """
    从被合并订单单元格中提取所有合法 Shopify 短单号：
    - 排除长单号（如 18 位 ERP 订单编号、16 位 AliExpress 单号）
    - 兼容以中英文逗号、分号、斜杠、换行符或空格分隔的混杂单号
    - 纯数字且长度为 3~11 位的短单号（通常 5~8 位，排除 0 等无效单号）
    """
    if merged_val is None:
        return []
    s = str(merged_val).strip()
    if not s or s.lower() in ['nan', 'none', 'null', '-', '无']:
        return []
    
    tokens = re.split(r'[,，;；/\s\n\r]+', s)
    shopify_orders = []
    seen = set()
    for tok in tokens:
        tok = tok.strip()
        if not tok:
            continue
        tok = re.sub(r'\.0$', '', tok)
        tok = tok.lstrip('#').strip()
        if tok.isdigit() and 3 <= len(tok) <= 11 and int(tok) > 0:
            if tok not in seen:
                seen.add(tok)
                shopify_orders.append(tok)
    return shopify_orders

def find_col_by_aliases(columns, aliases, exclude=None):
    """
    通过别名列表动态查找列名（忽略大小写与首尾空格）
    可指定 exclude 排除词，严格防止混淆不同类型单号（如 shopify order 与 ak order/订单编号）
    """
    exclude = [e.lower() for e in exclude] if exclude else []
    col_map = {str(c).strip().lower(): c for c in columns if c is not None}
    
    # 1. 精确全词匹配 (最高优先级)
    for alias in aliases:
        a_norm = alias.strip().lower()
        if a_norm in col_map:
            if not any(ex in a_norm for ex in exclude):
                return col_map[a_norm]
                
    # 2. 包含匹配 (严格校验排除词)
    for alias in aliases:
        a_norm = alias.strip().lower()
        for c_norm, orig_col in col_map.items():
            if any(ex in c_norm for ex in exclude):
                continue
            if a_norm in c_norm:
                return orig_col
    return None

def clean_order_template_to_script(order_path, script_path, route, sw_dsers_rename=False):
    """
    将下单模板清洗并倒模映射到脚本模板（全动态表头识别，不依赖固定列顺序）
    """
    if not os.path.exists(order_path):
        return
    try:
        df_order = pd.read_excel(order_path, dtype=str)
    except Exception as e:
        print(f"[!] 读取下单模板失败: {e}")
        return

    df_out = pd.DataFrame()
    
    # 查找订单编号列 (严格只认 Ak order 或 订单编号，二者为同一种单号；绝不与 Shopify order 或 Aliexpress order 混淆)
    col_order_name = find_col_by_aliases(
        df_order.columns, 
        ['ak order', 'ak_order', 'ak单号', '订单编号'],
        exclude=['shopify', 'shiopify', 'aliexpress']
    )
    if col_order_name:
        col_order = df_order[col_order_name].fillna("").astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    else:
        col_order = pd.Series([""] * len(df_order))
    df_out['订单编号'] = col_order

    # 查找 Shopify 单号列 (动态识别: 兼容 Shopify order, Shiopify order number 等，方便 DSers 等网页定位)
    col_shopify_name = find_col_by_aliases(
        df_order.columns, 
        ['shopify order', 'shopify order number', 'shiopify order number', 'shopify单号', 'shopify'],
        exclude=['aliexpress', 'ak']
    )
    if col_shopify_name:
        df_out['Shopify order'] = df_order[col_shopify_name].fillna("").astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    else:
        df_out['Shopify order'] = pd.Series([""] * len(df_order))


    # 查找客户姓名列 (动态按表头名称识别: 兼容 Customer name, 客户姓名, 姓名 等)
    col_name_header = find_col_by_aliases(
        df_order.columns, 
        ['customer name', 'customer_name', '客户姓名', '姓名', 'contact name', 'contact_name', 'contact_person']
    )
    if col_name_header:
        col_name = df_order[col_name_header].fillna("").astype(str).str.strip()
    else:
        col_name = pd.Series([""] * len(df_order))
    df_out['客户姓名'] = col_name

    # 查找 CPF / abnnumber 列 (动态按表头名称识别: 兼容 CPF, abnnumber, 税号 等)
    col_cpf_name = find_col_by_aliases(
        df_order.columns, 
        ['cpf', 'abnnumber', 'abn', '税号', 'cpf(brazil; optional)', 'cpf/abn']
    )
    if col_cpf_name:
        cpf_raw = df_order[col_cpf_name].fillna("").astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    else:
        cpf_raw = pd.Series([""] * len(df_order))
    df_out['abnnumber'] = cpf_raw

    # cpf_abn则是加了前缀“/cpf1 “的f列
    df_out['cpf1_abn'] = '/cpf1 ' + cpf_raw

    # 关键业务逻辑判断：
    # 若是在cpf管线里面 则客户姓名不用清洗到脚本模板的e列（也就是更改查询后的姓名）保持脚本模板里面的更改后姓名列空着
    # 若是在dsers里面 且打开独立分支，点击了运行后 需要把下单模板里的姓名清洗到脚本模板里的e列（查询更改后的姓名）
    if route == "B" and sw_dsers_rename:
        df_out['查询结果'] = col_name
    else:
        df_out['查询结果'] = ""

    # 查找出生日期列 (动态按表头名称识别: 兼容 Date of birth, 公司/门店名称, 出生日期, birthday 等)
    col_birth_name = find_col_by_aliases(
        df_order.columns, 
        ['date of birth', 'date_of_birth', 'dob', 'birthday', '出生日期', '公司/门店名称', '公司名称', '门店名称']
    )
    if col_birth_name:
        col_birth = df_order[col_birth_name].fillna("").astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    else:
        col_birth = pd.Series([""] * len(df_order))
    df_out['出生日期'] = col_birth

    # 查找被合并订单列 (动态识别: 兼容 被合并订单(订单编号), 被合并订单, 被合并 等)
    col_merged_name = find_col_by_aliases(
        df_order.columns,
        ['被合并订单(订单编号)', '被合并订单', '被合并', '被合并单号', 'merged order', 'merged orders', 'merged']
    )

    # DSers 改名模式：如果存在被合并订单列，提取其中的 Shopify 短单号并单独拆分为行（共享同一行客户信息）
    if sw_dsers_rename and col_merged_name:
        extra_rows = []
        for idx in range(len(df_order)):
            merged_val = df_order[col_merged_name].iloc[idx]
            shopify_ids = extract_shopify_orders(merged_val)
            if not shopify_ids:
                continue

            current_main_shopify = str(df_out['Shopify order'].iloc[idx]).strip()
            assigned_to_main = False
            # 若主行原本无 Shopify 单号，则将提取到的第一个 Shopify 单号回填给主行
            if not current_main_shopify:
                df_out.at[idx, 'Shopify order'] = shopify_ids[0]
                current_main_shopify = shopify_ids[0]
                assigned_to_main = True

            # 其余提取到的 Shopify 单号，均构造独立新行（完全共享同一行客户信息）
            pending_ids = shopify_ids[1:] if assigned_to_main else shopify_ids
            for sh_id in pending_ids:
                if sh_id == current_main_shopify:
                    continue
                extra_rows.append({
                    '订单编号': df_out['订单编号'].iloc[idx],
                    'Shopify order': sh_id,
                    '客户姓名': df_out['客户姓名'].iloc[idx],
                    'abnnumber': df_out['abnnumber'].iloc[idx],
                    'cpf1_abn': df_out['cpf1_abn'].iloc[idx],
                    '查询结果': df_out['查询结果'].iloc[idx],
                    '出生日期': df_out['出生日期'].iloc[idx]
                })

        if extra_rows:
            df_extra = pd.DataFrame(extra_rows)
            df_out = pd.concat([df_out, df_extra], ignore_index=True)
            print(f"[*] [DSers改名] 从被合并订单列中成功解析并新增了 {len(extra_rows)} 条独立的 Shopify 订单处理行")

    # 剔除无订单号的空白行及已经提示修改的行 (只要“订单编号”或“Shopify order”存在其一即为有效行)
    has_order = (
        (df_out['订单编号'].fillna('').astype(str).str.strip().ne('') & ~df_out['订单编号'].astype(str).str.strip().str.lower().isin(['nan', 'none'])) |
        (df_out['Shopify order'].fillna('').astype(str).str.strip().ne('') & ~df_out['Shopify order'].astype(str).str.strip().str.lower().isin(['nan', 'none']))
    )
    valid_mask = has_order
    # 过滤 customer name 为 already changed 的记录
    valid_mask = valid_mask & ~df_out['客户姓名'].astype(str).str.lower().str.contains('already.*change', regex=True, na=False)
    df_out = df_out[valid_mask].copy()

    try:
        save_df_to_excel(df_out, script_path)
        print(f"✅ 已成功将下单模板清洗映射并生成至脚本模板: {script_path}")
    except Exception as e:
        print(f"❌ 写入脚本模板失败: {e}")

def sync_cpf_results_to_order_template(script_path, order_path):
    """
    如果在cpf管线里面，导入的是下单模板后，运行后获取的新姓名要新增返回到下单模板的客户姓名列（按表头动态识别）
    """
    if not os.path.exists(script_path) or not os.path.exists(order_path):
        return
    try:
        df_script = pd.read_excel(script_path, dtype=str)
        df_order = pd.read_excel(order_path, dtype=str)
    except Exception as e:
        print(f"[!] 读取待同步表格失败: {e}")
        return

    # 动态匹配下单模板中的订单号列 (严格只认 Ak order 或 订单编号) 和客户姓名列
    order_col_name = find_col_by_aliases(
        df_order.columns, 
        ['ak order', 'ak_order', 'ak单号', '订单编号'],
        exclude=['shopify', 'shiopify', 'aliexpress']
    )
    name_col_name = find_col_by_aliases(
        df_order.columns, 
        ['customer name', 'customer_name', '客户姓名', '姓名', 'contact name', 'contact_person']
    )

    if not order_col_name or not name_col_name:
        print("[!] 下单模板中未找到订单编号列或客户姓名列，无法回填。")
        return

    # 动态从脚本模板中查找“订单编号”和“查询结果”列
    script_order_col = find_col_by_aliases(
        df_script.columns, 
        ['订单编号', 'ak order', 'ak_order'], 
        exclude=['shopify', 'aliexpress']
    ) or df_script.columns[0]
    script_result_col = find_col_by_aliases(df_script.columns, ['查询结果', 'tg_result', '姓名查询结果']) or (df_script.columns[4] if len(df_script.columns) > 4 else None)

    if not script_result_col:
        print("[!] 脚本模板中未找到查询结果列，无法回填。")
        return

    valid_names = {}
    for idx, row in df_script.iterrows():
        order_no = str(row[script_order_col]).strip()
        result = str(row[script_result_col]).strip()
        if order_no and result and result not in ["", "nan", "无", "遇到验证码且未能通过", "查询超时", "提取失败"]:
            valid_names[order_no] = result

    update_count = 0
    total_orders = len(df_order)
    processed_orders = 0

    # 更新下单模板对应的客户姓名列
    for idx, row in df_order.iterrows():
        processed_orders += 1
        order_no = str(row[order_col_name]).strip()
        if order_no in valid_names:
            df_order.at[idx, name_col_name] = valid_names[order_no]
            update_count += 1
        elif idx < len(df_script):
            script_order = str(df_script.at[idx, script_order_col]).strip()
            if not order_no or order_no == script_order:
                res = str(df_script.at[idx, script_result_col]).strip()
                if res and res not in ["", "nan", "无", "遇到验证码且未能通过", "查询超时", "提取失败"]:
                    df_order.at[idx, name_col_name] = res
                    update_count += 1
        print(f"[*] 进度提示：现在是 {processed_orders}/{total_orders} (共 {total_orders} 条，当前成功同步 {update_count} 条)")

    try:
        save_df_to_excel(df_order, order_path)
        print(f"✅ 成功回填更新 {update_count} 个最新查询姓名至下单模板客户姓名列: {order_path}")
    except Exception as e:
        print(f"❌ 同步保存下单模板失败: {e}")
