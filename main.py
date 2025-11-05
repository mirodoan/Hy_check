import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *

    BOOTSTRAP_AVAILABLE = True
except ImportError:
    BOOTSTRAP_AVAILABLE = False
    # Fallback constants
    SUCCESS = "success"
    WARNING = "warning"
    DANGER = "danger"
    PRIMARY = "primary"
    INFO = "info"
    BOTH = tk.BOTH
    X = tk.X
    Y = tk.Y
    LEFT = tk.LEFT
    RIGHT = tk.RIGHT
    VERTICAL = tk.VERTICAL
    W = tk.W

from database import Database
from excel_handler import ExcelHandler


class KiotCheckApp:
    def __init__(self):
        # Khởi tạo database và excel handler
        self.db = Database()
        self.excel_handler = ExcelHandler(self.db)

        # Tạo main window với ttkbootstrap theme
        if BOOTSTRAP_AVAILABLE:
            self.root = tb.Window(themename="cosmo")
        else:
            self.root = tk.Tk()
        self.root.title("KiotCheck - Quản lý sản phẩm")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        self.setup_ui()

    def setup_ui(self):
        """Thiết lập giao diện chính"""
        # Tạo notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Tab quản lý sản phẩm
        self.product_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.product_frame, text="📦 Quản lý sản phẩm")
        self.setup_product_tab()

        # Tab quét mã vạch
        self.scanner_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.scanner_frame, text="🔍 Quét mã vạch")
        self.setup_scanner_tab()

        # Tab import/export
        self.excel_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.excel_frame, text="📊 Import/Export")
        self.setup_excel_tab()

    def setup_product_tab(self):
        """Tab quản lý sản phẩm"""
        # Frame cho buttons
        btn_frame = ttk.Frame(self.product_frame)
        btn_frame.pack(fill=X, padx=10, pady=5)

        # Buttons
        ttk.Button(
            btn_frame,
            text="➕ Thêm sản phẩm",
            bootstyle=SUCCESS,
            command=self.add_product,
        ).pack(side=LEFT, padx=5)
        ttk.Button(
            btn_frame, text="✏️ Sửa", bootstyle=WARNING, command=self.edit_product
        ).pack(side=LEFT, padx=5)
        ttk.Button(
            btn_frame, text="🗑️ Xóa", bootstyle=DANGER, command=self.delete_product
        ).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 Refresh", command=self.refresh_products).pack(
            side=LEFT, padx=5
        )

        # Search frame
        search_frame = ttk.Frame(self.product_frame)
        search_frame.pack(fill=X, padx=10, pady=5)

        ttk.Label(search_frame, text="🔍 Tìm kiếm:").pack(side=LEFT, padx=5)
        self.search_var = tk.StringVar()
        # Sử dụng trace_add thay cho trace (Python >=3.7, bắt buộc ở 3.14)
        self.search_var.trace_add("write", lambda *args: self.on_search_change())
        search_entry = ttk.Entry(
            search_frame, textvariable=self.search_var, font=("Arial", 11)
        )
        search_entry.pack(side=LEFT, fill=X, expand=True, padx=5)

        # Treeview cho danh sách sản phẩm
        tree_frame = ttk.Frame(self.product_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        # Tạo Treeview chỉ với 2 cột: mã vạch, tên sản phẩm
        columns = ("barcode", "name")
        self.product_tree = ttk.Treeview(
            tree_frame, columns=columns, show="tree headings"
        )
        # Thiết lập headers
        self.product_tree.heading("#0", text="ID")
        self.product_tree.heading("barcode", text="Mã vạch")
        self.product_tree.heading("name", text="Tên sản phẩm")
        # Thiết lập column widths
        self.product_tree.column("#0", width=50)
        self.product_tree.column("barcode", width=120)
        self.product_tree.column("name", width=300)
        # Scrollbar
        scrollbar = ttk.Scrollbar(
            tree_frame, orient=VERTICAL, command=self.product_tree.yview
        )
        self.product_tree.configure(yscrollcommand=scrollbar.set)
        # Pack treeview và scrollbar
        self.product_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)
        # Load dữ liệu ban đầu
        self.refresh_products()

    def setup_scanner_tab(self):
        """Tab quét mã vạch"""
        # Title
        title_frame = ttk.Frame(self.scanner_frame)
        title_frame.pack(fill=X, padx=20, pady=20)

        title_label = ttk.Label(
            title_frame, text="🔍 QUÉT MÃ VẠCH SẢN PHẨM", font=("Arial", 16, "bold")
        )
        title_label.pack()

        # Scanner input
        scanner_frame = ttk.Frame(self.scanner_frame)
        scanner_frame.pack(fill=X, padx=20, pady=10)

        ttk.Label(
            scanner_frame, text="Nhập hoặc quét mã vạch:", font=("Arial", 12)
        ).pack(anchor=W)

        self.barcode_var = tk.StringVar()
        barcode_entry = ttk.Entry(
            scanner_frame, textvariable=self.barcode_var, font=("Arial", 14), width=30
        )
        barcode_entry.pack(pady=5)
        barcode_entry.bind("<Return>", self.scan_barcode)
        barcode_entry.focus()

        ttk.Button(
            scanner_frame,
            text="🔍 Tìm kiếm",
            bootstyle=PRIMARY,
            command=self.scan_barcode,
        ).pack(pady=10)

        # Kết quả
        self.result_frame = ttk.LabelFrame(self.scanner_frame, text="Kết quả tìm kiếm")
        self.result_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)

        # Treeview cho kết quả
        result_columns = ("name", "unit", "price")
        self.result_tree = ttk.Treeview(
            self.result_frame, columns=result_columns, show="headings"
        )

        self.result_tree.heading("name", text="Tên sản phẩm")
        self.result_tree.heading("unit", text="Đơn vị")
        self.result_tree.heading("price", text="Giá")

        self.result_tree.column("name", width=400)
        self.result_tree.column("unit", width=100)
        self.result_tree.column("price", width=100)

        self.result_tree.pack(fill=BOTH, expand=True, padx=10, pady=10)

    def setup_excel_tab(self):
        """Tab import/export Excel"""
        # Import section
        import_frame = ttk.LabelFrame(self.excel_frame, text="📥 Import từ Excel/CSV")
        import_frame.pack(fill=X, padx=20, pady=10)

        ttk.Label(
            import_frame, text="Chọn file Excel hoặc CSV để import sản phẩm:"
        ).pack(anchor=W, padx=10, pady=5)

        import_btn_frame = ttk.Frame(import_frame)
        import_btn_frame.pack(fill=X, padx=10, pady=10)

        ttk.Button(
            import_btn_frame,
            text="📁 Chọn file Import",
            bootstyle=SUCCESS,
            command=self.import_excel,
        ).pack(side=LEFT, padx=5)
        ttk.Button(
            import_btn_frame, text="📋 Tải template mẫu", command=self.download_template
        ).pack(side=LEFT, padx=5)

        # Export section
        export_frame = ttk.LabelFrame(self.excel_frame, text="📤 Export ra Excel")
        export_frame.pack(fill=X, padx=20, pady=10)

        ttk.Label(
            export_frame, text="Xuất toàn bộ danh sách sản phẩm ra file Excel:"
        ).pack(anchor=W, padx=10, pady=5)

        ttk.Button(
            export_frame,
            text="💾 Export Excel",
            bootstyle=INFO,
            command=self.export_excel,
        ).pack(anchor=W, padx=10, pady=10)

        # Log area
        log_frame = ttk.LabelFrame(self.excel_frame, text="📝 Log hoạt động")
        log_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)

        self.log_text = tk.Text(log_frame, height=10, font=("Consolas", 10))
        log_scrollbar = ttk.Scrollbar(
            log_frame, orient=VERTICAL, command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        self.log_text.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)
        log_scrollbar.pack(side=RIGHT, fill=Y, pady=10)

    # Product management methods
    def refresh_products(self):
        """Refresh danh sách sản phẩm"""
        # Xóa tất cả items cũ
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)
        # Load lại từ database
        products = self.db.get_all_products()
        for product in products:
            barcode = product[1]
            self.product_tree.insert(
                "", "end", text=product[0], values=(barcode, product[2])
            )

    def on_search_change(self, *args):
        """Xử lý khi search text thay đổi"""
        search_term = self.search_var.get()
        # Xóa tất cả items cũ
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)
        # Tìm kiếm và hiển thị
        if search_term.strip():
            products = self.db.search_products(search_term)
        else:
            products = self.db.get_all_products()
        for product in products:
            barcode = product[1]
            self.product_tree.insert(
                "", "end", text=product[0], values=(barcode, product[2])
            )

    def add_product(self):
        """Thêm sản phẩm mới"""
        self.product_dialog()

    def delete_product(self):
        """Xóa sản phẩm"""
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn sản phẩm cần xóa!")
            return
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa sản phẩm này?"):
            item = self.product_tree.item(selected[0])
            product_id = item["text"]
            self.db.delete_product(product_id)
            self.refresh_products()
            messagebox.showinfo("Thành công", "Đã xóa sản phẩm!")

    def edit_product(self):
        """Sửa sản phẩm"""
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn sản phẩm cần sửa!")
            return
        item = self.product_tree.item(selected[0])
        product_id = item["text"]
        values = item["values"]
        self.product_dialog(product_id, values)

    def product_dialog(self, barcode=None, values=None):
        """Dialog thêm/sửa sản phẩm với quản lý nhiều đơn vị tính"""

        dialog = tk.Toplevel(self.root)
        dialog.title("Thêm sản phẩm" if barcode is None else "Sửa sản phẩm")
        dialog.geometry("700x520")
        dialog.transient(self.root)
        dialog.grab_set()
        # Căn giữa màn hình
        dialog.update_idletasks()
        w = 700
        h = 520
        x = (dialog.winfo_screenwidth() // 2) - (w // 2)
        y = (dialog.winfo_screenheight() // 2) - (h // 2)
        dialog.geometry(f"{w}x{h}+{x}+{y}")

        barcode_var = tk.StringVar(value=values[0] if values else (barcode or ""))
        name_var = tk.StringVar(value=values[1] if values else "")

        # Form fields - đồng bộ tiêu đề size/bold
        ttk.Label(dialog, text="Mã vạch:", font=("Arial", 17, "bold")).pack(anchor=W, padx=20, pady=8)
        barcode_entry = ttk.Entry(dialog, textvariable=barcode_var, width=40, font=("Arial", 15))
        barcode_entry.pack(padx=20, pady=8)
        ttk.Label(dialog, text="Tên sản phẩm:", font=("Arial", 17, "bold")).pack(anchor=W, padx=20, pady=8)
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=40, font=("Arial", 15))
        name_entry.pack(padx=20, pady=8)


        # Đơn vị tính + giá
        unit_frame = ttk.LabelFrame(dialog)
        unit_frame.pack(fill=X, padx=20, pady=12)
        ttk.Label(unit_frame, text="Đơn vị tính & Giá bán", font=("Arial", 17, "bold")).pack(anchor=W, padx=8, pady=4)

        # Đơn vị tính + giá
        unit_frame = ttk.LabelFrame(dialog)
        unit_frame.pack(fill=X, padx=20, pady=12)
        ttk.Label(unit_frame, text="Đơn vị tính & Giá bán", font=("Arial", 17, "bold")).pack(anchor=W, padx=8, pady=4)

        add_unit_frame = ttk.Frame(unit_frame)
        add_unit_frame.pack(fill=X, padx=5, pady=5)
        add_unit_var = tk.StringVar()
        add_price_var = tk.StringVar()

        unit_entry = ttk.Entry(add_unit_frame, textvariable=add_unit_var, width=16, font=("Arial", 13))
        unit_entry.pack(side=LEFT, padx=4)
        unit_entry.config(foreground="#888")
        def clear_unit_placeholder(event):
            if unit_entry.get() == "Đơn vị":
                unit_entry.delete(0, tk.END)
                unit_entry.config(foreground="#000")
        def restore_unit_placeholder(event):
            if not unit_entry.get():
                unit_entry.insert(0, "Đơn vị")
                unit_entry.config(foreground="#888")
        unit_entry.insert(0, "Đơn vị")
        unit_entry.bind("<FocusIn>", clear_unit_placeholder)
        unit_entry.bind("<FocusOut>", restore_unit_placeholder)

        price_entry = ttk.Entry(add_unit_frame, textvariable=add_price_var, width=16, font=("Arial", 13))
        price_entry.pack(side=LEFT, padx=4)
        price_entry.config(foreground="#888")
        def clear_price_placeholder(event):
            if price_entry.get() == "Giá bán":
                price_entry.delete(0, tk.END)
                price_entry.config(foreground="#000")
        def restore_price_placeholder(event):
            if not price_entry.get():
                price_entry.insert(0, "Giá bán")
                price_entry.config(foreground="#888")
        price_entry.insert(0, "Giá bán")
        price_entry.bind("<FocusIn>", clear_price_placeholder)
        price_entry.bind("<FocusOut>", restore_price_placeholder)

        def add_unit():
            barcode = barcode_var.get().strip()
            unit = add_unit_var.get().strip()
            price_str = add_price_var.get().replace(".", "").replace(",", ".")
            try:
                price = float(price_str)
            except:
                messagebox.showerror("Lỗi", "Giá phải là số!")
                return
            if not barcode or not unit or price <= 0:
                messagebox.showerror("Lỗi", "Điền đủ barcode, đơn vị, giá > 0!")
                return
            if self.db.add_unit(barcode, unit, price):
                load_units()
                add_unit_var.set("")
                add_price_var.set("")
                restore_unit_placeholder(None)
                restore_price_placeholder(None)
            else:
                messagebox.showerror("Lỗi", "Trùng barcode + đơn vị!")

        ttk.Button(add_unit_frame, text="➕", width=3, command=add_unit, style="success.TButton").pack(side=LEFT, padx=4)

        unit_tree = ttk.Treeview(unit_frame, columns=("unit", "price"), show="headings", height=12)
        unit_tree.heading("unit", text="Đơn vị tính")
        unit_tree.heading("price", text="Giá bán")
        unit_tree.column("unit", width=220, anchor="center")
        unit_tree.column("price", width=220, anchor="center")
        unit_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=8, pady=8)

        unit_scroll = ttk.Scrollbar(unit_frame, orient=VERTICAL, command=unit_tree.yview)
        unit_tree.configure(yscrollcommand=unit_scroll.set)
        unit_scroll.pack(side=RIGHT, fill=Y)

        # Nút sửa/xóa cho từng đơn vị
        def show_action_buttons(event=None):
            if hasattr(unit_tree, "action_btns") and unit_tree.action_btns:
                for btn in unit_tree.action_btns:
                    btn.destroy()
                unit_tree.action_btns = []
            selected = unit_tree.selection()
            if selected:
                bbox = unit_tree.bbox(selected[0])
                if bbox:
                    x, y, w, h = bbox
                    def edit_unit():
                        item = unit_tree.item(selected[0])
                        old_unit = item['values'][0]
                        old_price = item['values'][1].replace(",", "").replace(".0", "")
                        add_unit_var.set(old_unit)
                        add_price_var.set(old_price)
                        delete_unit(selected[0])
                        restore_unit_placeholder(None)
                        restore_price_placeholder(None)
                    edit_btn = ttk.Button(unit_tree, text="✏️ Sửa", width=7, style="warning.TButton", command=edit_unit)
                    edit_btn.place(x=w+30, y=y)
                    del_btn = ttk.Button(unit_tree, text="🗑️ Xóa", width=7, style="danger.TButton", command=lambda: delete_unit(selected[0]))
                    del_btn.place(x=w+100, y=y)
                    unit_tree.action_btns = [edit_btn, del_btn]

        def hide_action_buttons(event=None):
            if hasattr(unit_tree, "action_btns") and unit_tree.action_btns:
                for btn in unit_tree.action_btns:
                    btn.destroy()
                unit_tree.action_btns = []

        unit_tree.action_btns = []
        unit_tree.bind("<<TreeviewSelect>>", show_action_buttons)
        unit_tree.bind("<Button-1>", hide_action_buttons)

        def load_units():
            unit_tree.delete(*unit_tree.get_children())
            barcode = barcode_var.get().strip()
            if barcode:
                units = self.db.get_units_by_barcode(barcode)
                for u in units:
                    unit_tree.insert("", "end", iid=u[0], values=(u[2], f"{u[3]:,.0f}"))
            hide_action_buttons()

        def delete_unit(unit_id):
            if messagebox.askyesno("Xác nhận", "Xóa đơn vị tính này?"):
                self.db.delete_unit(unit_id)
                load_units()

        load_units()
        unit_tree.bind("<Button-1>", hide_action_buttons)

        def load_units():
            unit_tree.delete(*unit_tree.get_children())
            barcode = barcode_var.get().strip()
            if barcode:
                units = self.db.get_units_by_barcode(barcode)
                for u in units:
                    unit_tree.insert("", "end", iid=u[0], values=(u[2], f"{u[3]:,.0f}"))
            hide_action_buttons()
        load_units()

        def delete_unit(unit_id):
            if messagebox.askyesno("Xác nhận", "Xóa đơn vị tính này?"):
                self.db.delete_unit(unit_id)
                load_units()

        # Gợi ý Google Images (chuyển xuống cuối)
        def open_google_images():
            import webbrowser

            q = barcode_var.get().strip() or name_var.get().strip()
            if not q:
                messagebox.showinfo("Gợi ý", "Nhập mã vạch hoặc tên sản phẩm trước!")
                return
            url = f"https://www.google.com/search?tbm=isch&q={q}"
            webbrowser.open(url)

        google_frame = ttk.Frame(dialog)
        google_frame.pack(fill=X, padx=20, pady=8, side=tk.BOTTOM)
        ttk.Label(
            google_frame, text="Gợi ý ảnh Google Images:", font=("Arial", 15)
        ).pack(side=LEFT)
        ttk.Button(
            google_frame,
            text="🔎 Mở Google Images",
            command=open_google_images,
            style="info.TButton",
        ).pack(side=LEFT, padx=10)

        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=24)

        def save_product():
            barcode = barcode_var.get().strip()
            name = name_var.get().strip()
            existed = self.db.get_product_by_barcode(barcode)
            if not barcode or not name:
                messagebox.showerror("Lỗi", "Điền đủ mã vạch và tên sản phẩm!")
                return
            existed = self.db.get_product_by_barcode(barcode)
            # Nếu đang sửa (barcode đã tồn tại và là dialog sửa), chỉ cập nhật
            if barcode and existed and (barcode == (values[0] if values else barcode)):
                self.db.update_product(barcode, name)
                messagebox.showinfo("Thành công", "Đã cập nhật sản phẩm!")
                dialog.destroy()
                self.refresh_products()
                return
            # Nếu đang thêm mới (barcode chưa tồn tại)
            if not existed:
                if self.db.add_product(barcode, name):
                    messagebox.showinfo("Thành công", "Đã thêm sản phẩm!")
                    dialog.destroy()
                    self.refresh_products()
                else:
                    messagebox.showerror("Lỗi", "Trùng mã vạch!")
                return
            # Nếu barcode đã tồn tại và không phải sửa đúng sản phẩm đó
            messagebox.showerror("Lỗi", "Trùng mã vạch!")

        ttk.Button(
            btn_frame,
            text="� Lưu",
            bootstyle=SUCCESS,
            command=save_product,
            style="success.TButton",
            width=12,
        ).pack(side=LEFT, padx=8)
        ttk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            style="danger.TButton",
            width=12,
        ).pack(side=LEFT, padx=8)

        # Gợi ý Google Images (chuyển xuống cuối)
        def open_google_images():
            import webbrowser

            q = barcode_var.get().strip() or name_var.get().strip()
            if not q:
                messagebox.showinfo("Gợi ý", "Nhập mã vạch hoặc tên sản phẩm trước!")
                return
            url = f"https://www.google.com/search?tbm=isch&q={q}"
            webbrowser.open(url)

        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=24)

        def save_product():
            barcode = barcode_var.get().strip()
            name = name_var.get().strip()
            existed = self.db.get_product_by_barcode(barcode)
            if not barcode or not name:
                messagebox.showerror("Lỗi", "Điền đủ mã vạch và tên sản phẩm!")
                return
            existed = self.db.get_product_by_barcode(barcode)
            # Nếu đang sửa (barcode đã tồn tại và là dialog sửa), chỉ cập nhật
            if barcode and existed and (barcode == (values[0] if values else barcode)):
                self.db.update_product(barcode, name)
                messagebox.showinfo("Thành công", "Đã cập nhật sản phẩm!")
                dialog.destroy()
                self.refresh_products()
                return
            # Nếu đang thêm mới (barcode chưa tồn tại)
            if not existed:
                if self.db.add_product(barcode, name):
                    messagebox.showinfo("Thành công", "Đã thêm sản phẩm!")
                    dialog.destroy()
                    self.refresh_products()
                else:
                    messagebox.showerror("Lỗi", "Trùng mã vạch!")
                return
            # Nếu barcode đã tồn tại và không phải sửa đúng sản phẩm đó
            messagebox.showerror("Lỗi", "Trùng mã vạch!")

        ttk.Button(btn_frame, text="💾 Lưu", bootstyle=SUCCESS, command=save_product, style="success.TButton", width=12).pack(side=LEFT, padx=8)
        ttk.Button(btn_frame, text="❌ Hủy", command=dialog.destroy, style="danger.TButton", width=12).pack(side=LEFT, padx=8)

        # Google Images chỉ 1 lần
        google_frame = ttk.Frame(dialog)
        google_frame.pack(fill=X, padx=20, pady=8, side=tk.BOTTOM)
        ttk.Label(google_frame, text="Gợi ý ảnh Google Images:", font=("Arial", 15)).pack(side=LEFT)
        ttk.Button(google_frame, text="🔎 Mở Google Images", command=open_google_images, style="info.TButton").pack(side=LEFT, padx=10)

    # Scanner methods
    def scan_barcode(self, event=None):
        """Quét mã vạch và hiện popup thông tin sản phẩm nếu có"""
        barcode = self.barcode_var.get().strip()
        if not barcode:
            return

        # Xóa kết quả cũ
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # Xóa/thêm label thêm mới nếu có
        if hasattr(self, "add_new_label") and self.add_new_label:
            self.add_new_label.destroy()
            self.add_new_label = None

        product = self.db.get_product_by_barcode(barcode)
        units = self.db.get_units_by_barcode(barcode)

        if product:
            # Popup thông tin sản phẩm
            popup = tk.Toplevel(self.root)
            popup.title("Thông tin sản phẩm")
            popup.geometry("700x520")
            popup.transient(self.root)
            popup.grab_set()
            popup.update_idletasks()
            w = 700
            h = 520
            x = (popup.winfo_screenwidth() // 2) - (w // 2)
            y = (popup.winfo_screenheight() // 2) - (h // 2)
            popup.geometry(f"{w}x{h}+{x}+{y}")
            # Barcode & tên
            ttk.Label(
                popup, text=f"Mã vạch: {product[1]}", font=("Arial", 20, "bold")
            ).pack(anchor=W, padx=30, pady=16)
            ttk.Label(
                popup, text=f"Tên sản phẩm: {product[2]}", font=("Arial", 18)
            ).pack(anchor=W, padx=30, pady=8)
            # Đơn vị & giá
            unit_frame = ttk.LabelFrame(
                popup, text="Đơn vị tính & Giá bán", font=("Arial", 16, "bold")
            )
            unit_frame.pack(fill=X, padx=30, pady=16)
            unit_tree = ttk.Treeview(
                unit_frame, columns=("unit", "price"), show="headings", height=7
            )
            unit_tree.heading("unit", text="Đơn vị tính")
            unit_tree.heading("price", text="Giá bán")
            unit_tree.column("unit", width=180, anchor="center")
            unit_tree.column("price", width=180, anchor="center")
            unit_tree.pack(fill=X, padx=8, pady=8)
            for u in units:
                unit_tree.insert("", "end", values=(u[2], f"{u[3]:,.0f} VNĐ"))

            # Nút Sửa
            def open_edit():
                popup.destroy()
                self.product_dialog(barcode=barcode, values=[product[1], product[2]])

            btn_frame = ttk.Frame(popup)
            btn_frame.pack(pady=18)
            ttk.Button(
                btn_frame,
                text="✏️ Sửa sản phẩm",
                command=open_edit,
                style="success.TButton",
                width=14,
            ).pack(side=LEFT, padx=10)
            ttk.Button(
                btn_frame,
                text="Đóng",
                command=popup.destroy,
                style="danger.TButton",
                width=14,
            ).pack(side=LEFT, padx=10)
        else:
            # Custom popup với nút Thêm sản phẩm mới
            popup = tk.Toplevel(self.root)
            popup.title("Không tìm thấy sản phẩm")
            popup.geometry("700x320")
            popup.transient(self.root)
            popup.grab_set()
            popup.update_idletasks()
            w = 700
            h = 320
            x = (popup.winfo_screenwidth() // 2) - (w // 2)
            y = (popup.winfo_screenheight() // 2) - (h // 2)
            popup.geometry(f"{w}x{h}+{x}+{y}")
            label = ttk.Label(
                popup,
                text=f"Không tìm thấy sản phẩm với mã vạch: {barcode}",
                font=("Arial", 18),
            )
            label.pack(pady=40)

            def open_add():
                popup.destroy()
                self.product_dialog(barcode=barcode, values=[barcode, "", "", "", ""])

            add_btn = ttk.Button(
                popup,
                text="➕ Thêm sản phẩm mới",
                command=open_add,
                style="success.TButton",
                width=18,
            )
            add_btn.pack(pady=18)
            close_btn = ttk.Button(
                popup,
                text="Đóng",
                command=popup.destroy,
                style="danger.TButton",
                width=18,
            )
            close_btn.pack(pady=8)

        # Clear input
        self.barcode_var.set("")

    # Excel methods
    def import_excel(self):
        """Import từ Excel/CSV"""
        file_path = filedialog.askopenfilename(
            title="Chọn file import",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )

        if file_path:
            success, message = self.excel_handler.import_from_excel(file_path)

            if success:
                self.log_message(f"✅ {message}")
                self.refresh_products()
                messagebox.showinfo("Thành công", message)
            else:
                self.log_message(f"❌ {message}")
                messagebox.showerror("Lỗi", message)

    def export_excel(self):
        """Export ra Excel"""
        file_path = filedialog.asksaveasfilename(
            title="Lưu file export",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )

        if file_path:
            success, message = self.excel_handler.export_to_excel(file_path)

            if success:
                self.log_message(f"✅ {message}")
                messagebox.showinfo("Thành công", message)
            else:
                self.log_message(f"❌ {message}")
                messagebox.showerror("Lỗi", message)

    def download_template(self):
        """Tải template mẫu"""
        file_path = filedialog.asksaveasfilename(
            title="Lưu template mẫu",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )

        if file_path:
            try:
                template_df = self.excel_handler.get_template_data()
                template_df.to_excel(file_path, index=False)
                self.log_message(f"✅ Đã tải template mẫu: {file_path}")
                messagebox.showinfo("Thành công", "Đã tải template mẫu!")
            except Exception as e:
                self.log_message(f"❌ Lỗi tải template: {str(e)}")
                messagebox.showerror("Lỗi", f"Lỗi tải template: {str(e)}")

    def log_message(self, message):
        """Ghi log message"""
        import datetime

        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    def run(self):
        """Chạy ứng dụng"""
        self.root.mainloop()


if __name__ == "__main__":
    app = KiotCheckApp()
    app.run()
