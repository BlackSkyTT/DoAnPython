import tkinter as tk
from tkinter import Menu, ttk, messagebox, filedialog
import mysql.connector
import pandas as pd

class TableApp():
    def __init__(self, root):
        self.root = root
        self.root.title("Kết nối database")
        # File menu
        self.menu_bar = Menu(self.root)
        file_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Xuất", command=self.export_to_excel)
        file_menu.add_command(label="Exit", command=self.quit_app)
        # Help menu
        help_menu = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.msg_box_info)
        self.root.config(menu=self.menu_bar)
        self.root.title("Quản lý sinh viên")
        self.db_name = tk.StringVar(value='DanhSachSV')
        self.user = tk.StringVar(value='root')
        self.password = tk.StringVar(value='123456')
        self.host = tk.StringVar(value='localhost')
        self.port = tk.StringVar(value='3306')
        self.table_name = tk.StringVar(value='sinhvien')

        width = 700
        height = 700
        self.root.geometry(f"{height}x{width}")
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f'{width}x{height}+{x}+{y}')

        self.widgets_connect()

    def widgets_connect(self):
        self.connection_frame = tk.Frame(self.root)
        self.connection_frame.pack(pady=10)

        tk.Label(self.connection_frame, text="DB Name:").grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(self.connection_frame, textvariable=self.db_name).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.connection_frame, text="User:").grid(row=2, column=0, padx=5, pady=5)
        tk.Entry(self.connection_frame, textvariable=self.user).grid(row=2, column=1, padx=5, pady=5)

        tk.Label(self.connection_frame, text="Password:").grid(row=3, column=0, padx=5, pady=5)
        tk.Entry(self.connection_frame, textvariable=self.password, show="*").grid(row=3, column=1, padx=5, pady=5)

        tk.Label(self.connection_frame, text="Host:").grid(row=4, column=0, padx=5, pady=5)
        tk.Entry(self.connection_frame, textvariable=self.host).grid(row=4, column=1, padx=5, pady=5)

        tk.Label(self.connection_frame, text="Port:").grid(row=5, column=0, padx=5, pady=5)
        tk.Entry(self.connection_frame, textvariable=self.port).grid(row=5, column=1, padx=5, pady=5)

        tk.Button(self.connection_frame, text="Connect", command=self.connect_to_manage).grid(row=6, columnspan=2, pady=10)

    def widgets_manage(self):
        # Khung chứa bảng
        self.table_frame = tk.Frame(self.root)
        self.table_frame.pack(pady=10)
        tk.Label(self.table_frame, text="Quản lý sinh viên", border=2, font=(
            "Helvetica", 16, "bold")).grid(column=0, row=0, columnspan=3, pady=20)
        # Tạo bảng Treeview
        self.data_table = ttk.Treeview(self.table_frame, columns=(
            "MSSV", "Họ", "Tên"), show="headings", height=10)
        self.data_table.heading("MSSV", text="MSSV")
        self.data_table.heading("Họ", text="Họ")
        self.data_table.heading("Tên", text="Tên")

        self.data_table.column("MSSV", width=100)
        self.data_table.column("Họ", width=200)
        self.data_table.column("Tên", width=200)
        self.data_table.grid(column=0, row=2)
        self.data_table.bind("<ButtonRelease-1>", self.on_tree_select)
        self.load_data()

        # Khung chứa các ô nhập liệu
        form_frame = ttk.LabelFrame(self.root, text="")
        form_frame.pack(pady=10)

        # Nhãn và ô nhập MSSV
        input_frame = ttk.LabelFrame(form_frame, text="Entry Frame")
        input_frame.grid(padx=10, row=0, column=0)
        ttk.Label(input_frame, text="MSSV : ").grid(
            row=0, column=0, padx=5, pady=5)
        self.mssv = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.mssv).grid(
            row=0, column=1, padx=5, pady=5)

        # Nhãn và ô nhập Họ Tên
        ttk.Label(input_frame, text="Họ : ").grid(
            row=1, column=0, padx=5, pady=5)
        self.ho = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.ho).grid(
            row=1, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Tên : ").grid(
            row=2, column=0, padx=5, pady=5)
        self.ten = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.ten).grid(
            row=2, column=1, padx=5, pady=5)

        # Frame button
        button_frame = ttk.LabelFrame(form_frame, text="Button Frame")
        button_frame.grid(padx=10, row=0, column=1)
        # Nút để thêm hàng mới
        add_button = ttk.Button(
            button_frame, text="Thêm hàng mới", command=self.add_data_button)
        add_button.grid(column=0, row=0, padx=5, pady=10)

        # Nút để load data
        load_button = ttk.Button(
            button_frame, text="Load data", command=self.load_data_button)
        load_button.grid(column=1, row=0, padx=5, pady=10)
        # Nút để xóa data đang chọn
        delete_button = ttk.Button(
            button_frame, text="Xóa các hàng đang chọn", command=self.delete_selected_row)
        delete_button.grid(column=0, row=1, padx=5, pady=10)

        # Nút để xóa data trong ô
        delete_input_button = ttk.Button(
            button_frame, text="Xóa dữ liệu trong ô MSSV", command=self.delete_data_button)
        delete_input_button.grid(column=1, row=1, padx=5, pady=10)

        # Nút clear input
        clear_button = ttk.Button(
            button_frame, text="Xóa nội dung trong ô", command=self.clear_inputs)
        clear_button.grid(column=0, row=2, padx=5, pady=10)
        # Nút update data
        update_button = ttk.Button(
            button_frame, text="Cập nhật nội dung trong ô", command=self.update_data_button)
        update_button.grid(column=1, row=2, padx=5, pady=10)

    def on_tree_select(self, event):
        """Xử lý khi người dùng chọn hàng trong Treeview"""
        selected_items = self.data_table.selection()
        if selected_items:
            item = selected_items[0]
            values = self.data_table.item(item, 'values')
            mssv = values[0]
            ho = values[1]
            ten = values[2]
            self.mssv.set(mssv)
            self.ho.set(ho)
            self.ten.set(ten)

    def add_data_button(self):
        """Thêm sinh viên mới vào cơ sở dữ liệu và bảng Treeview"""
        mssv = self.mssv.get().strip()
        ho = self.ho.get().strip()
        ten = self.ten.get().strip()
        if not mssv or not ho or not ten:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
            return
        self.add_data(mssv, ho, ten)
        self.clear_inputs()

    def clear_inputs(self):
        """Xóa nội dung các ô nhập liệu"""
        self.mssv.set("")
        self.ho.set("")
        self.ten.set("")

    def delete_selected_row(self):
        """Xóa hàng được chọn trong Treeview và cơ sở dữ liệu"""
        selected_items = self.data_table.selection()
        if selected_items:
            item = selected_items[0]
            values = self.data_table.item(item, 'values')
            mssv = values[0]
            self.delete_data(mssv)
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn hàng để xóa.")

    def update_data_button(self):
        """Cập nhật dữ liệu của sinh viên trong cơ sở dữ liệu và Treeview"""
        mssv = self.mssv.get().strip()
        ho = self.ho.get().strip()
        ten = self.ten.get().strip()
        if not mssv or not ho or not ten:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
            return
        self.update_data(mssv, ho, ten)

    def load_data_button(self):
        """Tải lại toàn bộ dữ liệu sinh viên từ cơ sở dữ liệu"""
        self.load_data()

    def delete_data_button(self):
        """Xóa sinh viên trong cơ sở dữ liệu dựa trên MSSV"""
        mssv = self.mssv.get().strip()
        if not mssv:
            messagebox.showwarning(
                "Cảnh báo", "Vui lòng nhập MSSV để xóa sinh viên.")
            return
        self.delete_data(mssv)
        self.clear_inputs()

    def export_to_excel(self):
        """Xuất dữ liệu sinh viên ra file Excel"""
        data = self.fetch_all_data()
        if not data:
            messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất.")
            return
        df = pd.DataFrame(data, columns=["MSSV", "Họ", "Tên"])
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Thông báo", "Xuất file thành công")

    def msg_box_info(self):
        """Hộp thoại thông tin về chương trình"""
        messagebox.showinfo("Thông tin", "Phần mềm quản lý sinh viên phiên bản 1.0")

    def quit_app(self):
        """Thoát ứng dụng"""
        self.root.quit()

    def connect_to_manage(self):
        try:
            self.conn = mysql.connector.connect(
                host=self.host.get(),
                port=self.port.get(),
                user=self.user.get(),
                password=self.password.get(),
            )
            self.cursor = self.conn.cursor()
        
        # Kiểm tra hoặc tạo database nếu chưa tồn tại
            self.cursor.execute(f"CREATE DATABASE IF NOT EXISTS {self.db_name.get()}")
            self.conn.database = self.db_name.get()  # Chọn database vừa tạo hoặc đã tồn tại
        
        # Kiểm tra hoặc tạo bảng sinhvien nếu chưa tồn tại
            create_table_query = f"""
            CREATE TABLE IF NOT EXISTS {self.table_name.get()} (
                mssv VARCHAR(20) PRIMARY KEY,
                ho VARCHAR(50),
                ten VARCHAR(50)
            );
            """
            self.cursor.execute(create_table_query)
            
            self.connection_frame.pack_forget()  # Ẩn form kết nối
            self.widgets_manage()  # Chuyển đến giao diện quản lý
        except mysql.connector.Error as e:
            messagebox.showerror("Lỗi kết nối", f"Không thể kết nối cơ sở dữ liệu: {e}")



    def add_data(self, mssv, ho, ten):
        """Thêm sinh viên mới vào cơ sở dữ liệu"""
        query = "INSERT INTO sinhvien (MSSV, Ho, Ten) VALUES (%s, %s, %s)"
        try:
            self.cursor.execute(query, (mssv, ho, ten))
            self.conn.commit()
            messagebox.showinfo("Thành công", "Thêm sinh viên thành công")
            self.load_data()
        except mysql.connector.Error as e:
            messagebox.showerror("Lỗi", f"Không thể thêm sinh viên: {e}")
            self.conn.rollback()

    def update_data(self, mssv, ho, ten):
        """Cập nhật thông tin sinh viên trong cơ sở dữ liệu"""
        query = "UPDATE sinhvien SET Ho=%s, Ten=%s WHERE MSSV=%s"
        try:
            self.cursor.execute(query, (ho, ten, mssv))
            self.conn.commit()
            messagebox.showinfo("Thành công", "Cập nhật sinh viên thành công")
            self.load_data()
        except mysql.connector.Error as e:
            messagebox.showerror("Lỗi", f"Không thể cập nhật sinh viên: {e}")
            self.conn.rollback()

    def delete_data(self, mssv):
        """Xóa sinh viên trong cơ sở dữ liệu"""
        query = "DELETE FROM sinhvien WHERE MSSV=%s"
        try:
            self.cursor.execute(query, (mssv,))
            self.conn.commit()
            messagebox.showinfo("Thành công", "Xóa sinh viên thành công")
            self.load_data()
        except mysql.connector.Error as e:
            messagebox.showerror("Lỗi", f"Không thể xóa sinh viên: {e}")
            self.conn.rollback()

    def load_data(self):
        """Tải dữ liệu sinh viên từ cơ sở dữ liệu vào bảng Treeview"""
        for i in self.data_table.get_children():
            self.data_table.delete(i)
        query = "SELECT MSSV, Ho, Ten FROM sinhvien"
        self.cursor.execute(query)
        rows = self.cursor.fetchall()
        for row in rows:
            self.data_table.insert("", tk.END, values=row)

    def fetch_all_data(self):
        """Lấy toàn bộ dữ liệu sinh viên từ cơ sở dữ liệu"""
        query = "SELECT MSSV, Ho, Ten FROM sinhvien"
        self.cursor.execute(query)
        return self.cursor.fetchall()

if __name__ == "__main__":
    root = tk.Tk()
    app = TableApp(root)
    root.mainloop()
