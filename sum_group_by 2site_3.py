import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import subprocess

out_path_global = None  # để lưu đường dẫn output cho nút "Open Folder"

def calculate_sumifs():
    global out_path_global
    file_path = filedialog.askopenfilename(
        title="Chọn file Excel",
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsb")]
    )
    if not file_path:
        return

    try:
        # Đọc file Excel
        df = pd.read_excel(file_path, sheet_name=0)
        df.columns = df.columns.str.strip()

        # ==== B1. Cột Past due → Total_Demand ====
        main_cols = df.loc[:, 'Past due':'Total_Demand'].columns.tolist()

        # ==== B2. Group Firm+Forecast ====
        filtered = df[df['Type'].isin(['Firm', 'Forecast'])]
        grouped = (
            filtered.groupby(['Part_No', 'Vendor_Code'])[main_cols]
            .sum()
            .reset_index()
        )
        grouped['Type'] = "Firm+Forecast"

        # ==== B3. Store_Qty & IQC_QTY theo site ====
        firm_rows = df[df['Type'] == 'Firm'].copy()

        s1 = (
            firm_rows[firm_rows['Site'] == 'TH3-SHTP']
            .groupby(['Part_No','Vendor_Code'])[['Store_Qty','IQC_QTY']]
            .sum()
            .reset_index()
            .rename(columns={'Store_Qty':'Store_TH3','IQC_QTY':'IQC_TH3'})
        )

        s2 = (
            firm_rows[firm_rows['Site'] == 'TD3-DDK']
            .groupby(['Part_No','Vendor_Code'])[['Store_Qty','IQC_QTY']]
            .sum()
            .reset_index()
            .rename(columns={'Store_Qty':'Store_TD3','IQC_QTY':'IQC_TD3'})
        )

        store_qty = pd.merge(s1, s2, on=['Part_No','Vendor_Code'], how='outer').fillna(0)
        store_qty['Store_Qty'] = store_qty['Store_TH3'] + store_qty['Store_TD3']
        store_qty['IQC_QTY'] = store_qty['IQC_TH3'] + store_qty['IQC_TD3']
        store_qty = store_qty[['Part_No','Vendor_Code','Store_Qty','IQC_QTY']]

        # ==== B4. Metadata ====
        meta_cols = [c for c in ['Part_No','Vendor_Code','Buyer','Planner','Vendor','Org','Site'] if c in df.columns]
        meta_info = df[meta_cols].drop_duplicates(subset=['Part_No','Vendor_Code'])

        # ==== B5. Merge ====
        result = pd.merge(grouped, store_qty, on=['Part_No','Vendor_Code'], how='left')
        result = pd.merge(result, meta_info, on=['Part_No','Vendor_Code'], how='left')

        # ==== B5.1. Gộp cột theo Month (YYYY-MM) ====
        date_cols = []
        for c in result.columns:
            try:
                parsed = pd.to_datetime(c, errors='coerce')
                if not pd.isna(parsed):
                    date_cols.append(c)
            except:
                pass

        if date_cols:
            melted = result.melt(
                id_vars=[c for c in result.columns if c not in date_cols],
                value_vars=date_cols,
                var_name="Date",
                value_name="Value"
            )
            melted['Month'] = pd.to_datetime(melted['Date']).dt.to_period("M").astype(str)

            # group lại theo month
            pivoted = (
                melted.pivot_table(
                    index=[c for c in result.columns if c not in date_cols],
                    columns="Month",
                    values="Value",
                    aggfunc="sum"
                )
                .reset_index()
            )
            pivoted.columns.name = None
            result = pivoted

        # ==== B6. Reorder columns ====
        cols = result.columns.tolist()
        new_order = ['Part_No','Vendor_Code','Type']
        if 'Store_Qty' in cols: new_order.append('Store_Qty')
        if 'IQC_QTY' in cols: new_order.append('IQC_QTY')

        # Thêm cột tháng YYYY-MM
        new_order += [c for c in cols if c not in new_order and c not in meta_cols]
        new_order += [c for c in meta_cols if c not in new_order]

        result = result[new_order]

        # ==== B6.1 Xóa cột Total_Demand và Site ====
        drop_cols = [c for c in ['Total_Demand','Site'] if c in result.columns]
        result = result.drop(columns=drop_cols, errors='ignore')

        # ==== B7. Xuất Excel ====
        base, ext = os.path.splitext(file_path)
        out_path_global = f"{base}_FirmForecast_Sum.xlsx"
        result.to_excel(out_path_global, index=False)

        messagebox.showinfo("Thành công", f"Đã xuất kết quả: {out_path_global}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))


def open_output_folder():
    """Mở thư mục chứa file kết quả"""
    global out_path_global
    if out_path_global and os.path.exists(out_path_global):
        folder = os.path.dirname(out_path_global)
        try:
            if os.name == 'nt':  # Windows
                subprocess.Popen(f'explorer /select,"{out_path_global}"')
            else:  # Mac/Linux
                subprocess.Popen(["open", folder])
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
    else:
        messagebox.showwarning("Chưa có file", "Chưa có file kết quả để mở.")


# ==== GUI ====
root = tk.Tk()
root.title("📊 Firm+Forecast 2Site-Supplier Capacity Monthly")
root.geometry("450x200")
root.configure(bg="#f0f4f7")

style = ttk.Style()
style.configure("TButton", font=("Segoe UI", 11), padding=8)
style.configure("TLabel", font=("Segoe UI", 12), background="#f0f4f7")

label = ttk.Label(root, text="Chọn file Excel để gộp Firm + Forecast và tính Store_Qty, IQC_QTY")
label.pack(pady=15)

btn_process = ttk.Button(root, text="📂 Chọn file Excel & Xử lý", command=calculate_sumifs)
btn_process.pack(pady=10)

btn_open = ttk.Button(root, text="📁 Mở thư mục kết quả", command=open_output_folder)
btn_open.pack(pady=10)

root.mainloop()
