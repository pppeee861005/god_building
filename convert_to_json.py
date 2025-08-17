import pandas as pd

# 定義來源與目標檔案之名
excel_path = '台灣寺廟網香客房一覽表.xlsx'
json_path = '香客大樓.json'

try:
    # 讀取 Excel 檔案
    data_frame = pd.read_excel(excel_path)

    # 將資料轉換為 JSON 格式並寫入檔案
    # orient='records' 使其成為物件陣列，易於取用
    # force_ascii=False 確保中文字符正常顯示
    # indent=4 使其格式優美，易於閱讀
    data_frame.to_json(json_path, orient='records', force_ascii=False, indent=4)

    print(f"檔案已成功轉換，存於 '{json_path}'")

except FileNotFoundError:
    print(f"錯誤：尋無此檔 '{excel_path}'")
except Exception as e:
    print(f"轉換之際，發生未料之誤：{e}")
