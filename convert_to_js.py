import pandas as pd
import json

# 定義來源與目標檔案之名
excel_path = '台灣寺廟網香客房一覽表.xlsx'
js_path = 'data.js'

try:
    # 讀取 Excel 檔案
    data_frame = pd.read_excel(excel_path)

    # 將資料轉換為 JSON 字串
    json_string = data_frame.to_json(orient='records', force_ascii=False, indent=4)

    # 構建 JavaScript 檔案內容
    js_content = f"const pilgrimData = {json_string};"

    # 將內容寫入 data.js
    with open(js_path, 'w', encoding='utf-8') as f:
        f.write(js_content)

    print(f"資料已成功寫入 '{js_path}'")

except FileNotFoundError:
    print(f"錯誤：尋無此檔 '{excel_path}'")
except Exception as e:
    print(f"處理之際，發生未料之誤：{e}")
