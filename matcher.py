import pandas as pd
from pandas import ExcelWriter
import tkinter as tk
from tkinter import filedialog, messagebox


def load_excel_files(file1_path, file2_path):
    try:
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        return df1, df2
    except Exception as e:
        messagebox.showerror("Помилка", f"Помилка при завантаженні файлів: {e}")
        return None, None


def compare_dataframes(df1, df2, key_column):
    merged = df1.merge(df2, on=key_column, how='outer', indicator=True)
    both = merged[merged['_merge'] == 'both']
    only_df1 = merged[merged['_merge'] == 'left_only']
    only_df2 = merged[merged['_merge'] == 'right_only']
    return both, only_df1, only_df2


def browse_file(entry_widget):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, filename)


def run_comparison():
    file1_path = entry_file1.get()
    file2_path = entry_file2.get()

    if not file1_path or not file2_path:
        messagebox.showwarning("Увага", "Будь ласка, виберіть обидва файли.")
        return

    df1, df2 = load_excel_files(file1_path, file2_path)
    if df1 is not None and df2 is not None:
        both, only_df1, only_df2 = compare_dataframes(df1, df2, "Invoice")

        output_path = "sample_data/matched_report.xlsx"
        with ExcelWriter(output_path, engine="openpyxl") as writer:
            both.to_excel(writer, sheet_name="Matches", index=False)
            only_df1.to_excel(writer, sheet_name="Only in File 1", index=False)
            only_df2.to_excel(writer, sheet_name="Only in File 2", index=False)

        messagebox.showinfo("Успіх", f"Результат збережено у {output_path}")


# === GUI ===
root = tk.Tk()
root.title("Excel Data Matcher")
root.geometry("500x250")

label1 = tk.Label(root, text="Файл 1:")
label1.pack(pady=(10, 0))
entry_file1 = tk.Entry(root, width=60)
entry_file1.pack()
btn_browse1 = tk.Button(root, text="Вибрати файл 1", command=lambda: browse_file(entry_file1))
btn_browse1.pack(pady=(0, 10))

label2 = tk.Label(root, text="Файл 2:")
label2.pack()
entry_file2 = tk.Entry(root, width=60)
entry_file2.pack()
btn_browse2 = tk.Button(root, text="Вибрати файл 2", command=lambda: browse_file(entry_file2))
btn_browse2.pack(pady=(0, 10))

btn_run = tk.Button(root, text="Порівняти та зберегти звіт", bg="#4CAF50", fg="white", command=run_comparison)
btn_run.pack(pady=10)

root.mainloop()
