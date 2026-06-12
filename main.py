import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


class ExcelViewerApp:

    def __init__(self, root):
        self.root = root
        self.root.title("工單資料欄位擷取器")

        # 根據作業系統微調視窗大小
        self.root.geometry("800x500" if os.name == "nt" else "840x520")

        # --- 介面配置 ---
        # 1. 上方操作區
        self.top_frame = tk.Frame(root)
        self.top_frame.pack(pady=15, padx=20, fill=tk.X)

        self.select_btn = tk.Button(
            self.top_frame,
            text=" 選擇 Excel 檔案 (.xlsx)",
            command=self.open_excel,
            bg="#0078D7",
            fg="white",
            font=("Microsoft JhengHei", 11, "bold"),
            padx=10,
            pady=3,
        )
        self.select_btn.pack(side=tk.LEFT)

        self.file_label = tk.Label(
            self.top_frame,
            text="尚未載入檔案",
            fg="gray",
            font=("Microsoft JhengHei", 10),
        )
        self.file_label.pack(side=tk.LEFT, padx=15)

        # 2. 中間資料顯示區 (表格與滾動條)
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        # 垂直與水平滾動條
        self.v_scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.VERTICAL)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.h_scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # 建立表格元件
        self.tree = ttk.Treeview(
            self.table_frame,
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set,
        )
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.v_scrollbar.config(command=self.tree.yview)
        self.h_scrollbar.config(command=self.tree.xview)

        # 設定表格樣式 (讓字體大一點、增加行高)
        style = ttk.Style()
        style.configure("Treeview", font=("Microsoft JhengHei", 10), rowheight=25)
        style.configure(
            "Treeview.Heading", font=("Microsoft JhengHei", 11, "bold")
        )

    def open_excel(self):
        # 彈出檔案選擇視窗
        file_path = filedialog.askopenfilename(
            title="請選擇工單 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )

        if not file_path:
            return

        self.file_label.config(text=f"已載入：{os.path.basename(file_path)}", fg="green")

        try:
            # 讀取 Excel (預設讀取第一個 Sheet)
            df = pd.read_excel(file_path)

            # 💡 【核心自訂區】在這裡設定您想要擷取的特定欄位名稱
            # 這些欄位名稱必須完全對應您 Excel 的第一行表頭
            target_columns = ["工單編號", "產品名稱", "工單數量", "狀態", "建立日"]

            # 檢查 Excel 裡面是否缺少這些欄位
            missing_cols = [col for col in target_columns if col not in df.columns]
            if missing_cols:
                messagebox.showerror(
                    "欄位不匹配",
                    f"在此 Excel 中找不到以下預設欄位：\n{', '.join(missing_cols)}\n\n請檢查 Excel 或修改程式中的 target_columns。",
                )
                return

            # 擷取特定欄位資料
            filtered_df = df[target_columns]

            # 清除舊的表格資料與欄位設定
            self.tree.delete(*self.tree.get_children())

            # 重新設定表格欄位標題
            self.tree["columns"] = target_columns
            self.tree["show"] = "headings"

            for col in target_columns:
                self.tree.heading(col, text=col)
                # 設定欄位寬度與置中
                self.tree.column(col, width=140, anchor="center")

            # 將資料逐行寫入畫面
            for _, row in filtered_df.iterrows():
                # 將 NaN (空值) 轉為空字串，避免畫面顯示 "nan"
                row_values = [
                    "" if pd.isna(val) else str(val) for val in row.values
                ]
                self.tree.insert("", tk.END, values=row_values)

            messagebox.showinfo("成功", f"已成功擷取並呈現 {len(filtered_df)} 筆資料！")

        except Exception as e:
            messagebox.showerror("讀取失敗", f"讀取 Excel 檔案時發生錯誤：\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
