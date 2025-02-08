import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def clear_excel_files():
    folder_path = filedialog.askdirectory(title="選擇Excel資料夾")
    if not folder_path:
        return
    
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xls'))]
    if not excel_files:
        messagebox.showwarning("警告", "選擇的資料夾中沒有Excel檔案！")
        return
    
    confirm = messagebox.askyesno("確認清空", f"確定要清空 {len(excel_files)} 個Excel檔案的資料嗎？（僅保留標題列）")
    if not confirm:
        return
    
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                continue
            
            # 保留標題但清空資料
            df.iloc[0:0].to_excel(file_path, index=False)
            print(f"清空完成: {file}")
        except Exception as e:
            print(f"處理 {file} 時發生錯誤: {e}")
    
    messagebox.showinfo("完成", "所有Excel檔案的資料已清空，只保留標題列！")

# 建立主視窗
root = tk.Tk()
root.title("Excel 資料清空工具")
root.geometry("300x150")

tk.Label(root, text="Excel 資料清空工具", font=("Arial", 14)).pack(pady=10)

clear_button = tk.Button(root, text="選擇資料夾並清空Excel", command=clear_excel_files)
clear_button.pack(pady=10)

exit_button = tk.Button(root, text="退出", command=root.quit)
exit_button.pack(pady=5)

root.mainloop()
