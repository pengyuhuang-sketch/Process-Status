import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


class ExcelViewerApp:

    def __init__(self, root):
        self.root = root
        self.root.title("工單詳細資料擷取器")

        # 設定視窗初始大小
        self.root.geometry("900x550" if os.name == "nt" else "950x580")

        # --- 介面配置 ---
        # 1. 上方操作按鈕區
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

        # 2. 中間資料顯示區 (表格與雙向滾動條)
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        # 建立垂直 (Y) 與水平 (X) 滾動條
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

        # 設定滾動條連動
        self.v_scrollbar.config(command=self.tree.yview)
        self.h_scrollbar.config(command=self.tree.xview)

        # 設定表格美化樣式 (字體與行高)
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
            # 讀取 Excel (預設讀取第一個工作表 Sheet1)
            df = pd.read_excel(file_path)

            # 💡 依照您的需求，嚴格定義從左到右的 11 個目標欄位
            target_columns = [
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
                "訂單編號",
            ]

            # 檢查 Excel 裡面是否缺少這些欄位
            missing_cols = [col for col in target_columns if col not in df.columns]
            if missing_cols:
                messagebox.showerror(
                    "欄位不匹配",
                    f"在此 Excel 中找不到以下必要欄位：\n{', '.join(missing_cols)}\n\n請檢查 Excel 的第一行表頭文字是否正確。",
                )
                return

            # 擷取並自動依 target_columns 的順序重新排列欄位
            filtered_df = df[target_columns]

            # 清除舊的表格資料與舊的欄位設定
            self.tree.delete(*self.tree.get_children())

            # 重新設定表格的欄位與標題
            self.tree["columns"] = target_columns
            self.tree["show"] = "headings"

            for col in target_columns:
                self.tree.heading(col, text=col)
                # 設定每個欄位的基本寬度為 120 像素，並靠中對齊
                self.tree.column(col, width=120, anchor="center")

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
