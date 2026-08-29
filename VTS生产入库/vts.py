import pandas as pd
import requests
import random
import time
import re
import json

# -------------------- 配置 --------------------
USERID = "N33274"
USERPWD = "123456"
IP = "10.224.151.59"
SCAN_TYPE = "1"

BASE_URL_1 = "http://10.224.245.101:8082/Wip/Inspection_wip.aspx"
BASE_URL_2 = "http://10.224.245.101:8082/Wip/Inspection_wip.aspx"
BASE_URL_3 = "http://10.224.245.101:8083/Api.ashx"

HEADERS = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Connection": "keep-alive",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
}

session = requests.Session()
session.headers.update(HEADERS)

# -------------------- 辅助函数 --------------------
def generate_sn_list(start_sn, end_sn):
    match_start = re.search(r'(\d+)$', start_sn)
    match_end = re.search(r'(\d+)$', end_sn)
    if not match_start or not match_end:
        raise ValueError(f"无法从SN中提取数字: {start_sn}, {end_sn}")
    num_start = int(match_start.group(1))
    num_end = int(match_end.group(1))
    prefix = start_sn[:-len(match_start.group(1))]
    width = len(match_start.group(1))
    sn_list = []
    for num in range(num_start, num_end + 1):
        sn = prefix + str(num).zfill(width)
        sn_list.append(sn)
    return sn_list

def extract_step_name(station):
    parts = station.split('-')
    return parts[-1] if parts else station

def parse_callback_response(text):
    text = text.strip()
    if text.startswith('t(') and text.endswith(')'):
        json_str = text[2:-1]
    else:
        json_str = text
    try:
        data = json.loads(json_str)
        if isinstance(data, list) and len(data) > 0:
            return data[0].get('error'), data[0].get('msg', '')
        else:
            return None, "响应格式异常"
    except json.JSONDecodeError:
        return None, f"JSON解析失败: {text}"

# -------------------- 主流程 --------------------
def main():
    # 读取Excel（优先 data.xlsx，其次 date.xlsx）
    try:
        df = pd.read_excel('date.xlsx')
    except FileNotFoundError:
        try:
            df = pd.read_excel('data.xlsx')
        except FileNotFoundError:
            print("未找到 data.xlsx 或 date.xlsx，退出。")
            return

    result_rows = []

    for idx, row in df.iterrows():
        so = str(row['SO']).strip()
        station = str(row['STATION']).strip()
        start_sn = str(row['START_SN']).strip()
        end_sn = str(row['END_SN']).strip()

        print(f"\n处理行 {idx+1}: SO={so}, STATION={station}")

        # ---------- 第一步：获取PN ----------
        key1 = random.random()
        url1 = f"{BASE_URL_1}?Oper=load_pn_info&value={so}&key={key1}"
        try:
            resp1 = session.post(url1, data={})
            resp1.raise_for_status()
            raw_text = resp1.text
            print(f"第一步响应: {raw_text[:200]}")   # 只打印前200字符

            # 解析JSON数组，提取 PRODUCT_NAME
            data_list = json.loads(raw_text)
            if isinstance(data_list, list) and len(data_list) > 0:
                pn = data_list[0].get('PRODUCT_NAME')
                if not pn:
                    # 尝试其他可能的字段名
                    pn = data_list[0].get('PN') or data_list[0].get('pn')
            else:
                pn = None

            if not pn:
                raise ValueError("未从响应中提取到PN")

            print(f"提取到的PN: {pn}")
        except Exception as e:
            print(f"第一步请求或解析失败: {e}")
            pn = "ERROR_PN"

        # ---------- 第二步：查询通过数量（可选） ----------
        step_name = extract_step_name(station)
        key2 = random.random()
        url2 = f"{BASE_URL_2}?Oper=load_pass_qty&lot_name={so}&step_name={step_name}&key={key2}"
        try:
            resp2 = session.post(url2, data={})
            resp2.raise_for_status()
            pass_qty = resp2.text.strip()
            print(f"第二步响应: {pass_qty}")
        except Exception as e:
            print(f"第二步请求失败: {e}")
            pass_qty = "ERROR"

        # ---------- 生成SN列表 ----------
        try:
            sn_list = generate_sn_list(start_sn, end_sn)
        except Exception as e:
            print(f"生成SN列表失败: {e}")
            continue
        print(f"共生成 {len(sn_list)} 个SN")

        # ---------- 第三步：逐个提交 ----------
        for sn in sn_list:
            key3 = random.random()
            timestamp = int(time.time() * 1000)
            url3 = (f"{BASE_URL_3}?type=1&action=complete&sn={sn}&pn={pn}&lot_name={so}"
                    f"&station={station}&scan_type={SCAN_TYPE}&userid={USERID}&userpwd={USERPWD}"
                    f"&defect_code=&defect_location=&ip={IP}&key={key3}&callback=t&_={timestamp}")

            try:
                resp3 = session.get(url3)
                resp3.raise_for_status()
                text3 = resp3.text
                print(f"SN {sn} 响应: {text3}")
                error, msg = parse_callback_response(text3)
                if error is None and msg is None:
                    error = -1
                    msg = text3
            except Exception as e:
                error = -2
                msg = str(e)

            result_rows.append({
                'SO': so,
                'STATION': station,
                'SN': sn,
                'PN': pn,
                'StepName': step_name,
                'PassQty': pass_qty,
                'Error': error,
                'Message': msg
            })

    # ---------- 输出结果 ----------
    if result_rows:
        result_df = pd.DataFrame(result_rows)
        result_df.to_excel('log.xlsx', index=False)
        print(f"\n完成！结果已保存到 log.xlsx，共 {len(result_rows)} 条记录。")
    else:
        print("没有生成任何结果。")

if __name__ == "__main__":
    main()