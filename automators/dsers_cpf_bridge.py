import pandas as pd
import argparse
import sys
import os

# 导入中心化配置
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from automators.excel_utils import save_df_to_excel
from config import DSERS_IMPORT_XLSX, DSERS_IMPORT_CSV, SCRIPT_TEMPLATE

def export_to_cpf_template(progress_callback=None):
    def log(msg):
        print(msg)
        if progress_callback:
            progress_callback(msg)

    import_path = DSERS_IMPORT_XLSX
    script_template_path = SCRIPT_TEMPLATE
    
    log("[*] 正在读取 DSers 映射文件...")
    try:
        df_dsers = pd.read_excel(import_path, dtype=str)
    except Exception as e:
        log(f"❌ 读取 import_orders.xlsx 失败: {e}")
        raise e
        
    log(f"[*] 成功读取 {len(df_dsers)} 条订单数据。")
    
    df_cpf = pd.DataFrame()
    df_cpf['订单编号'] = df_dsers.get('Order_number', '').fillna('').astype(str).str.strip()
    df_cpf['客户姓名'] = df_dsers.get('Contact_person', '')
    
    cpf_raw = df_dsers.get('CPF(Brazil; Optional)', '').fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df_cpf['abnnumber'] = cpf_raw
    df_cpf['cpf1_abn'] = '/cpf1 ' + cpf_raw
    df_cpf['TG_Result'] = '' # 空列用于接收结果
    df_cpf['出生日期'] = '' # 空列用于接收出生日期
    
    # 剔除无订单号的空白行
    valid_mask = df_cpf['订单编号'].ne('') & ~df_cpf['订单编号'].str.lower().isin(['nan', 'none'])
    df_cpf = df_cpf[valid_mask].copy()
    
    try:
        save_df_to_excel(df_cpf, script_template_path)
        log(f"✅ 成功将 DSers 订单桥接到 CPF 查名模板: {script_template_path}")
        return True
    except Exception as e:
        log(f"❌ 写入 脚本模板.xlsx 失败: {e}")
        raise e

def merge_cpf_results(progress_callback=None):
    def log(msg):
        print(msg)
        if progress_callback:
            progress_callback(msg)

    import_path = DSERS_IMPORT_XLSX
    import_csv_path = DSERS_IMPORT_CSV
    script_template_path = SCRIPT_TEMPLATE
    
    log("[*] 正在读取查名完成的 脚本模板.xlsx ...")
    try:
        df_cpf = pd.read_excel(script_template_path, dtype=str)
        df_dsers = pd.read_excel(import_path, dtype=str)
    except Exception as e:
        log(f"❌ 读取文件失败: {e}")
        raise e
        
    log(f"[*] 成功读取 DSers 模板 ({len(df_dsers)}条) 和 CPF 结果表 ({len(df_cpf)}条)。")
    
    cols = df_cpf.columns.tolist()
    if len(cols) >= 5:
        result_col = cols[4] # 索引4即第五列(E列)
    else:
        log("❌ 脚本模板.xlsx 格式错误，缺少结果列(E列)！")
        raise RuntimeError("缺少结果列(E列)")
        
    update_count = 0
    df_cpf['订单编号'] = df_cpf['订单编号'].fillna('').astype(str).str.strip()
    df_dsers['Order_number'] = df_dsers['Order_number'].fillna('').astype(str).str.strip()
    
    valid_names = {}
    for idx, row in df_cpf.iterrows():
        order_no = row['订单编号']
        result = str(row[result_col]).strip()
        if order_no and result and result not in ["", "nan", "无", "遇到验证码且未能通过", "查询超时", "提取失败"]:
            valid_names[order_no] = result

    total_dsers = len(df_dsers)
    dsers_processed = 0
    for idx, row in df_dsers.iterrows():
        dsers_processed += 1
        order_no = row['Order_number']
        if order_no in valid_names:
            df_dsers.at[idx, 'Contact_person'] = valid_names[order_no]
            update_count += 1
        if dsers_processed % 50 == 0 or dsers_processed == total_dsers:
            log(f"[*] 桥接进度提示：现在是 {dsers_processed}/{total_dsers} (共 {total_dsers} 条，当前成功回填 {update_count} 条)")
            
    from config import DSERS_TEMPLATE
    try:
        df_raw = pd.read_excel(DSERS_TEMPLATE, dtype=str)
        if '交易编号' in df_raw.columns and '客户姓名' in df_raw.columns:
            for idx, row in df_raw.iterrows():
                order_no = str(row['交易编号']).strip()
                if order_no in valid_names:
                    df_raw.at[idx, '客户姓名'] = valid_names[order_no]
            df_raw.to_excel(DSERS_TEMPLATE, index=False)
            log(f"[*] 原始 dsers模板.xlsx 的 '客户姓名' 也已同步更新！")
    except Exception as e:
        log(f"[!] 同步更新 dsers模板.xlsx 失败: {e}")
            
    log(f"[*] 匹配完毕！共成功修正并回填了 {update_count} 个真实的客户姓名。")
    
    try:
        save_df_to_excel(df_dsers, import_path)
        df_dsers.to_csv(import_csv_path, index=False, encoding='utf-8-sig')
        log(f"✅ 修正后的 DSers 订单已重新保存并覆盖: {import_csv_path}")
        return True
    except Exception as e:
        log(f"❌ 保存覆写文件失败: {e}")
        raise e

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="DSers 和 Telegram CPF 的桥接器")
    parser.add_argument("--mode", choices=["export", "merge"], required=True, help="export: 将DSers提取成脚本模板 | merge: 将查名结果合回DSers模板")
    args = parser.parse_args()
    
    if args.mode == "export":
        export_to_cpf_template()
    elif args.mode == "merge":
        merge_cpf_results()
