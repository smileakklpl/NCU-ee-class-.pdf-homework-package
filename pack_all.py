import os
import shutil
import pandas as pd

def process_student_submissions(source_dir, dest_dir, student_list_df):
    """
    處理所有學生的作業提交，PDF保留原名，其餘檔案重新命名，並產生報告。
    """
    # 建立目標資料夾
    os.makedirs(dest_dir, exist_ok=True)
    print(f"目標資料夾 '{os.path.basename(dest_dir)}' 已準備就緒。")

    # 取得磁碟上實際存在的學生資料夾名稱 (忽略隱藏資料夾)
    try:
        existing_folders = {f for f in os.listdir(source_dir) if os.path.isdir(os.path.join(source_dir, f)) and not f.startswith('.')}
    except FileNotFoundError:
        print(f"錯誤：找不到來源資料夾 '{source_dir}'")
        return [], {}, 0 # 回傳0個檔案

    # 準備記錄名單和檔案計數器
    missing_list = []
    wrong_format_list = {}
    copied_files_count = 0 # <--- 新增檔案計數器

    # 遍歷官方學生名單
    for index, row in student_list_df.iterrows():
        try:
            expected_folder_name = f"{row['學號']} ({row['姓名']})"
        except KeyError:
            print("錯誤：CSV檔案中必須包含 '學號' 和 '姓名' 兩個欄位。")
            return [], {}, 0

        # 狀況 A: 學生資料夾不存在
        if expected_folder_name not in existing_folders:
            missing_list.append(expected_folder_name)
            continue

        # 狀況 B: 學生資料夾存在
        student_folder_path = os.path.join(source_dir, expected_folder_name)
        files_in_folder = [f for f in os.listdir(student_folder_path) if not f.startswith('.')]

        # 狀況 B.1: 資料夾是空的
        if not files_in_folder:
            missing_list.append(expected_folder_name)
            continue

        # 狀況 B.2: 資料夾內有檔案
        has_pdf = False
        other_files_extensions = []
        for filename in files_in_folder:
            
            if filename.lower().endswith('.html'):
                continue

            source_file_path = os.path.join(student_folder_path, filename)

            if os.path.isfile(source_file_path):
                
                if filename.lower().endswith('.pdf'):
                    new_filename = filename
                else:
                    new_filename = f"{expected_folder_name}_{filename}"

                dest_file_path = os.path.join(dest_dir, new_filename)
                
                if filename.lower().endswith('.pdf') and os.path.exists(dest_file_path):
                    print(f"警告：檔案 '{filename}' 已存在，將被覆蓋。")
                
                try:
                    shutil.copy2(source_file_path, dest_file_path)
                    copied_files_count += 1 # <--- 每複製一個檔案就+1
                except Exception as e:
                    print(f"錯誤：複製檔案 {source_file_path} 失敗: {e}")

                # 判斷檔案類型
                if filename.lower().endswith('.pdf'):
                    has_pdf = True
                else:
                    ext = os.path.splitext(filename)[1]
                    if ext:
                        other_files_extensions.append(ext)

        # 判斷最終狀態
        if not has_pdf:
            if other_files_extensions:
                wrong_format_list[expected_folder_name] = sorted(list(set(other_files_extensions)))
            else: 
                missing_list.append(expected_folder_name)
    
    print("\n所有檔案複製完成，開始產生報告...")
    # 回傳時，多回傳一個檔案總數
    return sorted(missing_list), wrong_format_list, copied_files_count

if __name__ == "__main__":
    current_directory = os.getcwd()
    
    # --- 1. 指定學生名單和目標資料夾名稱 ---
    STUDENT_LIST_FILENAME = '學生名單.xlsx' # 根據您的程式碼，這裡是讀取 .xlsx
    DESTINATION_FOLDER_NAME = ""
    DESTINATION_FOLDER_NAME = input("請輸入欲輸出的資料夾名稱:")
    source_folder_name = None
        
    # --- 2. 讀取學生名單 ---
    try:
        df_students = pd.read_excel(STUDENT_LIST_FILENAME)
        if '學號' not in df_students.columns or '姓名' not in df_students.columns:
            raise KeyError
        df_students['學號'] = df_students['學號'].astype(str)
        print(f"成功讀取學生名單 '{STUDENT_LIST_FILENAME}'。")
    except FileNotFoundError:
        print(f"錯誤：找不到學生名單 '{STUDENT_LIST_FILENAME}'。")
        print("請確認此程式、學生名單Excel檔、作業資料夾都在同一個目錄下。")
        exit()
    except KeyError:
        print(f"錯誤：學生名單 '{STUDENT_LIST_FILENAME}' 中必須同時包含 '學號' 和 '姓名' 兩個欄位標題。")
        exit()
        
    # --- 3. 自動偵測作業來源資料夾 ---
    for item in os.listdir(current_directory):
        item_path = os.path.join(current_directory, item)
        if os.path.isdir(item_path) and item != DESTINATION_FOLDER_NAME and not item.startswith('.'):
            source_folder_name = item
            print(f"自動偵測到作業來源資料夾為: '{source_folder_name}'")
            break

    if source_folder_name is None:
        print(f"錯誤：在目前目錄下找不到可用的作業來源資料夾。")
    else:
        source_path = os.path.join(current_directory, source_folder_name)
        dest_path = os.path.join(current_directory, DESTINATION_FOLDER_NAME)

        # --- 4. 執行主要處理程序 ---
        # 接收函式回傳的第三個值：檔案總數
        missing, wrong_format, total_files = process_student_submissions(source_path, dest_path, df_students)

        # 在終端機印出總數
        print(f"報告 '繳交狀況報告.txt' 已產生於 '{DESTINATION_FOLDER_NAME}' 資料夾中。")
        print(f"總共打包了 {total_files} 個檔案。") # <--- 在終端機顯示總數

        # --- 5. 產生最終的 txt 報告 ---
        report_path = os.path.join(dest_path, '繳交狀況報告.txt')
        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write("作業繳交狀況報告\n")
                f.write("====================================\n\n")
                
                # 將檔案總數寫入報告
                f.write(f"總共打包檔案數量：{total_files}\n\n") # <--- 在報告中加入總數

                if not missing and not wrong_format:
                    f.write("恭喜！所有學生都已成功繳交 PDF 檔案。\n")
                
                if missing:
                    f.write("缺交人員:\n")
                    for i, name in enumerate(missing, 1):
                        f.write(f"{i}. {name}\n")
                    f.write("\n")
                
                if wrong_format:
                    f.write("格式出錯 (繳交了非 PDF 檔案):\n")
                    i = 1
                    for name, exts in sorted(wrong_format.items()):
                        f.write(f"{i}. {name} (上傳了 {', '.join(exts)} 檔案)\n")
                        i += 1

        except Exception as e:
            print(f"錯誤：寫入報告 '{report_path}' 失敗: {e}")