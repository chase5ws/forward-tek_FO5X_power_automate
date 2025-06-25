import pandas as pd
import re
from datetime import datetime, timedelta

# 使用者輸入變數 x，定義篩選天數
x = 7

def parse_data_from_excel(file_path):
    """
    從 Excel 檔案中讀取數據並進行解析。
    """
    try:
        df = pd.read_excel(file_path, header=None, engine="openpyxl")
        raw_data = "".join(df.iloc[:, 0].astype(str))  # 假設數據在第一列
    except Exception as e:
        print(f"讀取 Excel 檔案失敗: {e}")
        return None

    # 定義正則表達式來匹配數據格式
    pattern = r"(FO5X\d{12})(\d{4}/\d{2}/\d{2})(\d{4}-[\u4e00-\u9fa5]+)(\d{4}-[\u4e00-\u9fa5]+)(\d\.[\u4e00-\u9fa5A-Za-z0-9\(\)]+)(\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2})(已簽核)(同意|不同意)"
    
    # 匹配所有行
    matches = re.findall(pattern, raw_data)
    
    if not matches:
        print("未找到符合的數據格式！")
        return None

    # 將匹配的數據轉換為 DataFrame
    columns = ["表單代號", "申請日期", "填表人", "填表部門", "需求類別", "結案日期", "簽核狀態", "簽核結果"]
    processed_df = pd.DataFrame(matches, columns=columns)
    
    return processed_df

def filter_and_sort_data(df, days):
    """
    根據申請日期進行篩選和排序。
    - days: 篩選距離最近日期的天數。
    """
    # 將申請日期轉換為 datetime 格式
    df["申請日期"] = pd.to_datetime(df["申請日期"], format="%Y/%m/%d")
    
    # 計算篩選的日期範圍
    latest_date = df["申請日期"].max()
    cutoff_date = latest_date - timedelta(days=days)
    
    # 篩選出符合條件的數據
    filtered_df = df[df["申請日期"] >= cutoff_date]
    
    # 根據申請日期進行降序排序
    sorted_df = filtered_df.sort_values(by="申請日期", ascending=False)
    
    # 將日期轉回字串格式
    sorted_df["申請日期"] = sorted_df["申請日期"].dt.strftime("%Y/%m/%d")
    
    return sorted_df

def generate_statistics(df):
    """
    根據填表部門和需求類別進行統計。
    """
    # 統計每個部門的填寫次數
    department_counts = df["填表部門"].value_counts().reset_index()
    department_counts.columns = ["填表部門", "填寫次數"]

    # 統計每個需求類別的次數
    category_counts = df["需求類別"].value_counts().reset_index()
    category_counts.columns = ["需求類別", "需求次數"]

    # 合併統計結果
    stats_df = pd.concat([department_counts, category_counts], axis=1)
    
    return stats_df

def save_to_excel_with_statistics(df, stats_df, output_file):
    """
    將處理後的數據和統計結果保存為新的 Excel 檔案。
    - stats_df: 統計數據
    - df: 篩選後的數據
    """
    try:
        # 創建一個新的 Excel writer
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # 將統計結果寫入 Excel 的第一個工作表
            stats_df.to_excel(writer, sheet_name="統計結果", index=False)
            
            # 將篩選後的數據寫入同一個 Excel 的第二個工作表
            df.to_excel(writer, sheet_name="篩選後數據", index=False)
        
        print(f"數據已保存至 {output_file}")
    except Exception as e:
        print(f"保存 Excel 檔案失敗: {e}")

def main():
    # 讓使用者輸入檔案路徑
    file_path = input("請輸入要處理的 Excel 檔案路徑（例如 C:\\Users\\user\\Downloads\\weekreport.xlsx）：\n")
    output_file = input("請輸入處理後的 Excel 檔案儲存路徑（例如 C:\\Users\\user\\Downloads\\processed_weekreport.xlsx）：\n")
    
    # 解析數據
    processed_df = parse_data_from_excel(file_path)
    
    if processed_df is not None:
        # 篩選和排序數據
        filtered_sorted_df = filter_and_sort_data(processed_df, x)
        
        # 生成統計數據
        stats_df = generate_statistics(filtered_sorted_df)
        
        # 保存結果至 Excel
        save_to_excel_with_statistics(filtered_sorted_df, stats_df, output_file)
        print(f"處理完成！結果已儲存至：{output_file}")
    else:
        print("處理失敗，請檢查數據格式或檔案內容。")

if __name__ == "__main__":
    main()

