import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename
import csv
import datetime
import pandas as pd
root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("600x400")
def save_to_csv():
    data = {
        "Mã": entry_ma.get(),
        "Tên": entry_ten.get(),
        "Đơn vị": combobox_donvi.get(),
        "Chức danh": entry_chucdanh.get(),
        "Ngày sinh": entry_ngaysinh.get(),
        "Giới tính": gender_var.get(),
        "Số CMND": entry_cmnd.get(),
        "Ngày cấp": entry_ngaycap.get(),}
    if not all(data.values()):
        messagebox.showwarning("Cảnh báo", "Vui lòng điền đầy đủ thông tin!")
        return
    with open("nhanvien.csv", mode="a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=data.keys())
        if file.tell() == 0:
            writer.writeheader()
        writer.writerow(data)
    messagebox.showinfo("Thành công", "Lưu thông tin thành công!")
    clear_fields()
def clear_fields():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    combobox_donvi.set("")
    entry_chucdanh.delete(0, tk.END)
    entry_ngaysinh.delete(0, tk.END)
    gender_var.set("Nam")
    entry_cmnd.delete(0, tk.END)
    entry_ngaycap.delete(0, tk.END)
def show_today_birthdays():
    try:
        today = datetime.datetime.now().strftime("%d/%m/%Y")
        with open("nhanvien.csv", mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            birthdays = [row for row in reader if row["Ngày sinh"] == today]
        if not birthdays:
            messagebox.showinfo("Kết quả", "Không có nhân viên nào sinh nhật hôm nay.")
            return
            result = "\n".join([f"{b['Tên']} - {b['Đơn vị']}" for b in birthdays])
        messagebox.showinfo("Sinh nhật hôm nay", result)
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên!")
def export_to_excel():
    try:
        with open("nhanvien.csv", mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            data = sorted(reader, key=lambda x: datetime.datetime.strptime(x["Ngày sinh"], "%d/%m/%Y"), reverse=True)

        filepath = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            df = pd.DataFrame(data)
            df.to_excel(filepath, index=False, encoding="utf-8")
            messagebox.showinfo("Thành công", "Xuất file Excel thành công!")
    except FileNotFoundError:
        messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên!")
frame = tk.Frame(root)
frame.pack(pady=10, padx=10, fill="x")
labels = ["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Ngày cấp"]
for i, label in enumerate(labels):
    tk.Label(frame, text=label).grid(row=i, column=0, pady=5, sticky="e")
entry_ma = tk.Entry(frame)
entry_ma.grid(row=0, column=1, pady=5, sticky="w")
entry_ten = tk.Entry(frame)
entry_ten.grid(row=1, column=1, pady=5, sticky="w")
combobox_donvi = ttk.Combobox(frame, values=["Phân xưởng cơ khí", "Phân xưởng cơ điện", "Phân xưởng vận hành"])
combobox_donvi.grid(row=2, column=1, pady=5, sticky="w")
entry_chucdanh = tk.Entry(frame)
entry_chucdanh.grid(row=3, column=1, pady=5, sticky="w")
entry_ngaysinh = tk.Entry(frame)
entry_ngaysinh.grid(row=4, column=1, pady=5, sticky="w")
gender_var = tk.StringVar(value="Nam")
tk.Radiobutton(frame, text="Nam", variable=gender_var, value="Nam").grid(row=5, column=1, sticky="w")
tk.Radiobutton(frame, text="Nữ", variable=gender_var, value="Nữ").grid(row=5, column=2, sticky="w")
entry_cmnd = tk.Entry(frame)
entry_cmnd.grid(row=6, column=1, pady=5, sticky="w")
entry_ngaycap = tk.Entry(frame)
entry_ngaycap.grid(row=7, column=1, pady=5, sticky="w")
btn_save = tk.Button(root, text="Lưu thông tin", command=save_to_csv)
btn_save.pack(pady=5)
btn_birthday = tk.Button(root, text="Sinh nhật ngày hôm nay", command=show_today_birthdays)
btn_birthday.pack(pady=5)
btn_export = tk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel)
btn_export.pack(pady=5)
root.mainloop()
