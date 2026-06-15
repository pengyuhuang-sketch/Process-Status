import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl


class ExcelViewerApp:

    def __init__(self, root):
        self.root = root
        self.root.title("工單詳細資料擷取器 (過濾折疊版)")

        # 設定視窗初始大小
        self.root.geometry("950x550" if os.name == "nt" else "1000x580")

        # --- 介面配置 ---
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

        # 雙向滾動條設計
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        self.v_scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.VERTICAL)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.h_scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            self.table_frame,
            yscrollcommand=self.v_scrollbar.set,
            xscrollcommand=self.h_scrollbar.set,
        )
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.v_scrollbar.config(command=self.tree.yview)
        self.h_scrollbar.config(command=self.tree.xview)

        style = ttk.Style()
        style.configure("Treeview", font=("Microsoft JhengHei", 10), rowheight=25)
        style.configure(
            "Treeview.Heading", font=("Microsoft JhengHei", 11, "bold")
        )

    def open_excel(self):
        file_path = filedialog.askopenfilename(
            title="請選擇工單 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )

        if not file_path:
            return

        self.file_label.config(text=f"已載入：{os.path.basename(file_path)}", fg="green")

        try:
            # 1. 核心：使用 openpyxl 先行載入，用來辨識哪些列被「折疊隱藏」了
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active # 讀取當前活動的工作表
            
            # 找出哪些列是隱藏的 (Row index 是從 1 開始)
            hidden_rows = []
            for row_idx, row_dim in ws.row_dimensions.items():
                if row_dim.hidden:
                    hidden_rows.append(row_idx)

            # 2. 用 pandas 正常讀取整張表
            df = pd.read_excel(file_path)

            # 💡 關鍵：過濾掉在 Excel 中被折疊隱藏的資料列
            # pandas 的 index 是從 0 開始，對應 Excel 的 row_idx - 2 (扣除表頭列)
            if hidden_rows:
                # 轉換為 pandas 的索引位置 (Excel第2行對應pandas第0行)
                pandas_hidden_indices = [r - 2 for r in hidden_rows if r >= 2]
                df = df.drop(index=pandas_hidden_indices, errors='ignore')

            # 3. 處理欄位對齊與重新命名邏輯
            col_mapping = {}
            for real_col in df.columns:
                cleaned = str(real_col).strip().replace(" ", "")
                
                if cleaned == "工單編號":
                    col_mapping[real_col] = "工單編號"
                elif cleaned in ["工單產品編號", "產品編號"]:
                    col_mapping[real_col] = "產品編號"  # 統一對齊使用者要求的名稱
                elif cleaned == "產品名稱":
                    col_mapping[real_col] = "產品名稱"
                elif cleaned == "工單數量":
                    col_mapping[real_col] = "工單數量"
                elif cleaned == "批號":
                    col_mapping[real_col] = "批號"
                elif cleaned in ["目前數量", "目 前 數 量"]:
                    col_mapping[real_col] = "目前數量"
                elif cleaned == "作業站編號":
                    col_mapping[real_col] = "作業站編號"
                elif cleaned == "狀態":
                    col_mapping[real_col] = "狀態"
                elif cleaned == "到達時間":
                    col_mapping[real_col] = "到達時間"
                elif cleaned == "區段編號":
                    col_mapping[real_col] = "區段編號"
                elif cleaned == "訂單編號":
                    col_mapping[real_col] = "訂單編號"

            df = df.rename(columns=col_mapping)

            # 4. 您嚴格指定的由左至右 11 個顯示欄位
            display_headers = [
                "工單編號",
                "產品編號",
                "產品名稱",
                "工單數量",
                "批號",
                "目前數量",
                "作業站編號",
                "狀態",
                "到達時間",
                "區段編號",
                "訂單編號"
            ]

            # 檢查重命名後的欄位是否完整
            missing_cols = [col for col in display_headers if col not in df.columns]
            if missing_cols:
                messagebox.showerror(
                    "欄位不匹配",
                    f"在此 Excel 中找不到以下必要欄位：\n{', '.join(missing_cols)}\n\n請確認 Excel 欄位名稱是否正確。",
                )
                return

            # 5. 擷取並排序
            filtered_df = df[display_headers]

            # 6. 清除畫面舊資料
            self.tree.delete(*self.tree.get_children())

            # 7. 設定表格欄位
            self.tree["columns"] = display_headers
            self.tree["show"] = "headings"

            for col in display_headers:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor="center")

            # 8. 將未折疊的資料寫入畫面
            for _, row in filtered_df.iterrows():
                row_values = [
                    "" if pd.isna(val) else str(val) for val in row.values
                ]
                self.tree.insert("", tk.END, values=row_values)

            messagebox.showinfo("成功", f"已成功過濾折疊資料，共呈現 {len(filtered_df)} 筆可視資料！")

        except Exception as e:
            messagebox.showerror("讀取失敗", f"讀取 Excel 檔案時發生錯誤：\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()
