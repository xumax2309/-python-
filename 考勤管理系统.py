import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import pandas as pd
import pyodbc
import requests
from io import BytesIO

class ImageDownloader:
    @staticmethod
    def download_image(url):
        response = requests.get(url)
        image_data = response.content
        return Image.open(BytesIO(image_data))

class DatabaseConnection:
    def __init__(self):
        self.conn = None
        self.cursor = None

    def connect(self, dsn):
        self.conn = pyodbc.connect(dsn)
        self.cursor = self.conn.cursor()

    def execute(self, query, params=None):
        if params is None:
            params = []
        self.cursor.execute(query, params)
        self.conn.commit()

    def fetchall(self, query, params=None):
        if params is None:
            params = []
        self.cursor.execute(query, params)
        return self.cursor.fetchall()

    def close(self):
        if self.conn:
            self.conn.close()

class AttendanceManagementApp:
    def __init__(self, master):
        self.master = master
        self.master.title("考勤管理系统")
        self.master.withdraw()  # 隐藏主窗口，直到登录成功

        self.font = ("楷体", 16)  # 设置字体为楷体，大小为16
        self.db = DatabaseConnection()
        self.login_image_url = "https://n.sinaimg.cn/ent/transform/775/w630h945/20191021/13b4-ihfpfwa1327234.jpg"
        self.main_image_url = "https://wx4.sinaimg.cn/mw690/005Yovzjgy1hqi44ui07mj32c03404qr.jpg"
        self.create_login_window()

    def create_login_window(self):
        self.login_window = tk.Toplevel(self.master)
        self.login_window.title("登录")

        # 下载并加载背景图片
        self.login_background_image = ImageDownloader.download_image(self.login_image_url)
        self.login_bg = ImageTk.PhotoImage(self.login_background_image)

        # 获取图片尺寸
        image_width, image_height = self.login_background_image.size

        # 设置窗口大小与图片相同
        self.login_window.geometry(f"{image_width}x{image_height}")

        self.canvas = tk.Canvas(self.login_window, width=image_width, height=image_height)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_image(0, 0, image=self.login_bg, anchor="nw")

        # 添加用户名和密码输入框
        self._create_entry(self.canvas, image_width // 2, image_height // 2 - 40, "用户名:", self.font)
        self.username_entry = self._create_entry_window(self.canvas, image_width // 2, image_height // 2 - 10, self.font)

        self._create_entry(self.canvas, image_width // 2, image_height // 2 + 20, "密码:", self.font)
        self.password_entry = self._create_entry_window(self.canvas, image_width // 2, image_height // 2 + 50, self.font, show="*")

        # 添加登录按钮
        self.login_button = tk.Button(self.login_window, text="登录", command=self.login, font=self.font, bg="lightblue", bd=2)
        self.canvas.create_window(image_width // 2, image_height // 2 + 90, window=self.login_button)

        # 绑定回车键触发登录功能
        self.login_window.bind('<Return>', lambda event: self.login())

    def _create_entry(self, canvas, x, y, text, font):
        canvas.create_text(x, y, text=text, fill="white", font=font)

    def _create_entry_window(self, canvas, x, y, font, show=None):
        entry = tk.Entry(canvas, font=font, bd=2, relief="sunken", show=show)
        canvas.create_window(x, y, window=entry)
        return entry

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if self.validate_login(username, password):
            self.login_window.withdraw()  # 隐藏登录窗口
            self.master.deiconify()  # 显示主窗口
            self.create_main_window()
            # 延迟进行数据库连接
            self.master.after(100, self.connect_to_database)
        else:
            messagebox.showerror("错误", "用户名或密码错误")

    def connect_to_database(self):
        self.db.connect('Driver={SQL Server};Server=毛毛的ROG幻16;Database=AttendanceManagement;Trusted_Connection=yes;')

    def validate_login(self, username, password):
        return username == "毛毛" and password == "778899"

    def create_main_window(self):
        # 下载并加载背景图片
        self.main_background_image = ImageDownloader.download_image(self.main_image_url)
        self.main_bg = ImageTk.PhotoImage(self.main_background_image)

        # 获取图片尺寸
        image_width, image_height = self.main_background_image.size

        # 设置窗口大小与图片相同
        self.master.geometry(f"{image_width}x{image_height}")

        # 创建Canvas并显示背景图片
        self.canvas = tk.Canvas(self.master, width=image_width, height=image_height)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_image(0, 0, image=self.main_bg, anchor="nw")

        # 创建菜单
        self.create_menu()

    def create_menu(self):
        menu = tk.Menu(self.master, font=self.font)
        self.master.config(menu=menu)

        attendance_menu = tk.Menu(menu, font=self.font, tearoff=0)
        menu.add_cascade(label="考勤管理", menu=attendance_menu)
        attendance_menu.add_command(label="删除/查询员工考勤记录", command=self.delete_query_attendance_records, font=self.font)
        attendance_menu.add_command(label="公司策略设定", command=self.set_company_policy, font=self.font)
        attendance_menu.add_command(label="显示当日迟到、缺勤明细", command=self.show_absent_details, font=self.font)

        data_menu = tk.Menu(menu, font=self.font, tearoff=0)
        menu.add_cascade(label="数据", menu=data_menu)
        data_menu.add_command(label="Excel历史数据导入", command=self.import_excel_data, font=self.font)

    def delete_query_attendance_records(self):
        self._create_window("删除/查询员工考勤记录", self._create_delete_query_content)

    def _create_window(self, title, content_creator):
        window = tk.Toplevel(self.master)
        window.title(title)
        window.geometry("400x200")  # 设置窗口大小和位置
        content_creator(window)

    def _create_delete_query_content(self, window):
        self._create_label_entry(window, "员工ID:", 0, self._set_employee_id_entry)
        self._create_buttons(window, [("删除", self.delete_attendance_record), ("查询", self.query_attendance_record)], 1)

    def _create_label_entry(self, window, text, row, setter):
        label = tk.Label(window, text=text, font=self.font)
        label.grid(row=row, column=0, pady=10, padx=10)
        entry = tk.Entry(window, font=self.font, bd=2, relief="sunken")
        entry.grid(row=row, column=1, pady=10, padx=10)
        setter(entry)

    def _set_employee_id_entry(self, entry):
        self.employee_id_entry = entry

    def _create_buttons(self, window, buttons, row):
        for idx, (text, command) in enumerate(buttons):
            button = tk.Button(window, text=text, command=command, font=self.font, bg="lightblue", bd=2)
            button.grid(row=row, column=idx, pady=10, padx=10)

    def delete_attendance_record(self):
        employee_id = self.employee_id_entry.get()
        self.db.execute("DELETE FROM AttendanceRecords WHERE EmployeeID=?", (employee_id,))
        messagebox.showinfo("提示", f"员工ID为 {employee_id} 的考勤记录已删除")

    def query_attendance_record(self):
        employee_id = self.employee_id_entry.get()
        records = self.db.fetchall("SELECT * FROM AttendanceRecords WHERE EmployeeID=?", (employee_id,))

        if records:
            messagebox.showinfo("查询结果", f"员工ID为 {employee_id} 的考勤记录:\n{records}")
        else:
            messagebox.showinfo("查询结果", f"未找到员工ID为 {employee_id} 的考勤记录")

    def set_company_policy(self):
        self._create_window("公司策略设定", self._create_policy_content)

    def _create_policy_content(self, window):
        self._create_label_entry(window, "上班时间:", 0, self._set_start_time_entry)
        self._create_label_entry(window, "下班时间:", 1, self._set_end_time_entry)
        self._create_buttons(window, [("保存", self.save_company_policy)], 2)

    def _set_start_time_entry(self, entry):
        self.start_time_entry = entry

    def _set_end_time_entry(self, entry):
        self.end_time_entry = entry

    def save_company_policy(self):
        start_time = self.start_time_entry.get()
        end_time = self.end_time_entry.get()
        self.db.execute("UPDATE CompanyPolicy SET StartTime=?, EndTime=?", (start_time, end_time))
        messagebox.showinfo("提示", "公司策略已更新")

    def show_absent_details(self):
        self._create_window("当日迟到、缺勤明细", self._create_absent_details_content)

    def _create_absent_details_content(self, window):
        self._create_label_entry(window, "日期(YYYY-MM-DD):", 0, self._set_date_entry)
        self._create_buttons(window, [("显示明细", self.show_details)], 1)

    def _set_date_entry(self, entry):
        self.date_entry = entry

    def show_details(self):
        date_str = self.date_entry.get()

        try:
            late_records = self.db.fetchall("SELECT EmployeeID FROM AttendanceRecords WHERE AttendanceDate=? AND AttendanceTime > '09:00:00'", (date_str,))
            absent_records = self.db.fetchall("SELECT DISTINCT EmployeeID FROM Employees WHERE EmployeeID NOT IN (SELECT EmployeeID FROM AttendanceRecords WHERE AttendanceDate=?)", (date_str,))

            late_employee_ids = [record[0] for record in late_records]
            absent_employee_ids = [record[0] for record in absent_records]

            messagebox.showinfo("当日迟到员工", f"迟到员工ID: {late_employee_ids}")
            messagebox.showinfo("当日缺勤员工", f"缺勤员工ID: {absent_employee_ids}")

        except Exception as e:
            messagebox.showerror("错误", f"查询出错: {str(e)}")

    def import_excel_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            for _, row in df.iterrows():
                employee_id = row['EmployeeID']
                attendance_date = row['AttendanceDate']
                attendance_time = row['AttendanceTime']
                clock_out = row['ClockOut']
                self.db.execute("INSERT INTO AttendanceRecords (EmployeeID, AttendanceDate, AttendanceTime, ClockOut) VALUES (?, ?, ?, ?)", (employee_id, attendance_date, attendance_time, clock_out))
            
            messagebox.showinfo("成功", "Excel数据导入成功")
        except Exception as e:
            messagebox.showerror("错误", f"导入出错: {str(e)}")

    def __del__(self):
        self.db.close()

def main():
    root = tk.Tk()
    app = AttendanceManagementApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
