import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename
import csv
from datetime import datetime
import os
import pandas as pd

# File lưu trữ dữ liệu CSV
data_file = "employees.csv"

# Kiểm tra và tạo file CSV nếu chưa tồn tại
if not os.path.exists(data_file):
    with open(data_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Nơi cấp", "Ngày cấp"])

def save_employee():
    employee_data = [
        entry_id.get(),
        entry_name.get(),
        combo_unit.get(),
        combo_position.get(),
        entry_dob.get(),
        gender_var.get(),
        entry_id_number.get(),
        entry_issuing_place.get(),
        entry_issue_date.get()
    ]

    if any(not field for field in employee_data):
        messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin.")
        return

    with open(data_file, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(employee_data)

    messagebox.showinfo("Thành công", "Đã lưu thông tin nhân viên.")
    clear_fields()

def clear_fields():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    combo_unit.set("")
    combo_position.set("")
    entry_dob.delete(0, tk.END)
    gender_var.set("Nam")
    entry_id_number.delete(0, tk.END)
    entry_issuing_place.delete(0, tk.END)
    entry_issue_date.delete(0, tk.END)

def show_birthdays():
    today = datetime.today().strftime("%d/%m")
    birthdays = []

    with open(data_file, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if today == row["Ngày sinh"][:5]:
                birthdays.append(row)

    if not birthdays:
        messagebox.showinfo("Thông báo", "Hôm nay không có nhân viên nào sinh nhật.")
    else:
        result = "\n".join([f"{emp['Tên']} ({emp['Mã']}) - {emp['Đơn vị']}" for emp in birthdays])
        messagebox.showinfo("Danh sách sinh nhật hôm nay", result)

def export_employees():
    if not os.path.exists(data_file):
        messagebox.showerror("Lỗi", "Không tìm thấy dữ liệu nhân viên.")
        return

    employees = []
    with open(data_file, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            row["Tuổi"] = calculate_age(row["Ngày sinh"])
            employees.append(row)

    employees.sort(key=lambda x: x["Tuổi"], reverse=True)
    output_file = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

    if output_file:
        df = pd.DataFrame(employees)
        df.drop(columns=["Tuổi"], inplace=True)
        df.to_excel(output_file, index=False, encoding='utf-8')
        messagebox.showinfo("Thành công", f"Đã xuất danh sách ra file {output_file}")

def calculate_age(dob):
    try:
        birth_date = datetime.strptime(dob, "%d/%m/%Y")
        today = datetime.today()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return age
    except ValueError:
        return 0

# Tạo giao diện chính
root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("600x400")

# Các trường nhập liệu
frame = ttk.LabelFrame(root, text="Thông tin nhân viên")
frame.pack(padx=10, pady=10, fill="x")

labels = ["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Nơi cấp", "Ngày cấp"]
for i, text in enumerate(labels):
    ttk.Label(frame, text=text + " *").grid(row=i, column=0, sticky="w", padx=5, pady=5)

entry_id = ttk.Entry(frame)
entry_name = ttk.Entry(frame)
combo_unit = ttk.Combobox(frame, values=["Phân xưởng que hàn", "Phân xưởng khác"])
combo_position = ttk.Combobox(frame, values=["Nhân viên", "Quản lý"])
entry_dob = ttk.Entry(frame)
gender_var = tk.StringVar(value="Nam")
radio_male = ttk.Radiobutton(frame, text="Nam", variable=gender_var, value="Nam")
radio_female = ttk.Radiobutton(frame, text="Nữ", variable=gender_var, value="Nữ")
entry_id_number = ttk.Entry(frame)
entry_issuing_place = ttk.Entry(frame)
entry_issue_date = ttk.Entry(frame)

entries = [entry_id, entry_name, combo_unit, combo_position, entry_dob, entry_id_number, entry_issuing_place, entry_issue_date]

for i, widget in enumerate(entries):
    widget.grid(row=i, column=1, sticky="ew", padx=5, pady=5)
radio_male.grid(row=5, column=1, sticky="w", padx=5)
radio_female.grid(row=5, column=1, sticky="e", padx=5)

# Các nút chức năng
btn_frame = ttk.Frame(root)
btn_frame.pack(pady=10)

btn_save = ttk.Button(btn_frame, text="Lưu thông tin", command=save_employee)
btn_birthdays = ttk.Button(btn_frame, text="Sinh nhật hôm nay", command=show_birthdays)
btn_export = ttk.Button(btn_frame, text="Xuất toàn bộ danh sách", command=export_employees)

btn_save.grid(row=0, column=0, padx=10)
btn_birthdays.grid(row=0, column=1, padx=10)
btn_export.grid(row=0, column=2, padx=10)

root.mainloop()
