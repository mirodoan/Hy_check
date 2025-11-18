import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import platform
import sys, os
FONT_FAMILY = "Segoe UI" if platform.system() == "Windows" else "Arial"
from PIL import Image, ImageTk

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

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


class HyCheckApp:
    def __init__(self):
        # Khởi tạo database và excel handler
        self.db = Database()
        self.excel_handler = ExcelHandler(self.db)

        # Tạo main window với ttkbootstrap theme
        if BOOTSTRAP_AVAILABLE:
            self.root = tb.Window(themename="cosmo")
        else:
            self.root = tk.Tk()
        self.root.title("HyCheck - Quản lý sản phẩm")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        self.setup_ui()

    def setup_ui(self):
        """Thiết lập giao diện chính"""
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=(FONT_FAMILY, 16, "bold"), padding=[24, 12])
        style.map("TNotebook.Tab",
            background=[("selected", "#ff9800"), ("!selected", "#dee7ed")],
            foreground=[("selected", "#fff"), ("!selected", "#333")]
        )
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        style.configure("Treeview.Heading", font=(FONT_FAMILY, 17, "bold"))
        style.configure("Treeview", font=(FONT_FAMILY, 16), rowheight=60)
        style.configure("TButton", font=(FONT_FAMILY, 14))

        # Tạo notebook (tabs)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Tab import/export
        self.excel_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.excel_frame, text="📊 Import/Export")

        # Tab quản lý sản phẩm
        self.product_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.product_frame, text="📦 Quản lý sản phẩm")

        # Tab quét mã vạch
        self.scanner_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.scanner_frame, text="🔍 Quét mã vạch")

        self.setup_excel_tab()
        self.setup_product_tab()
        self.setup_scanner_tab()

        

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
            search_frame, textvariable=self.search_var, font=(FONT_FAMILY, 16)
        )
        search_entry.pack(side=LEFT, fill=X, expand=True, padx=5)

        # Treeview cho danh sách sản phẩm
        tree_frame = ttk.Frame(self.product_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        # Tạo Treeview với 3 cột: ID, mã vạch, tên sản phẩm, căn thẳng hàng
        columns = ("barcode", "name")
        self.product_tree = ttk.Treeview(
            tree_frame, columns=columns, show="tree headings", selectmode="extended"
        )
        # Thiết lập style header lớn, đậm
        style = ttk.Style()
        style.configure("Treeview.Heading", font=(FONT_FAMILY, 18, "bold"))
        style.configure("Treeview", font=(FONT_FAMILY, 16))
        style.configure("Treeview", rowheight=60)  # Tăng chiều cao dòng lên 60px để hiển thị rõ hơn
        # Thiết lập column widths và căn giữa/thẳng hàng
        self.product_tree.column("#0", width=80, anchor="center", stretch=False)
        self.product_tree.column("barcode", width=260, anchor="center", stretch=False, minwidth=220)
        self.product_tree.column("name", width=520, anchor="w", stretch=True, minwidth=320)
        # Đảm bảo header style lớn
        self.product_tree.heading("#0", text="ID", anchor="center")
        self.product_tree.heading("barcode", text="Mã vạch", anchor="center")
        self.product_tree.heading("name", text="Tên sản phẩm", anchor="w")
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
        # Lưu ảnh gốc để resize động
        self.img_left_src = Image.open(resource_path("Hy.png"))
        self.img_right_src = Image.open(resource_path("Dan.png"))

        # Tạo label cho ảnh
        self.label_left = tk.Label(self.scanner_frame, borderwidth=0)
        self.label_right = tk.Label(self.scanner_frame, borderwidth=0)
        self.label_left.place(relx=0, rely=0.5, anchor="w")
        self.label_right.place(relx=1, rely=0.5, anchor="e")

        # Nội dung trung tâm
        self.center_frame = ttk.Frame(self.scanner_frame)
        self.center_frame.place(relx=0.5, rely=0.5, anchor="center")

        title_label = ttk.Label(
            self.center_frame, text="🔍 QUÉT MÃ VẠCH SẢN PHẨM", font=(FONT_FAMILY, 24, "bold")
        )
        title_label.pack(pady=30)
        ttk.Label(
            self.center_frame, text="Nhập hoặc quét mã vạch:", font=(FONT_FAMILY, 16)
        ).pack(pady=10)
        self.barcode_var = tk.StringVar()
        barcode_entry = ttk.Entry(
            self.center_frame, textvariable=self.barcode_var, font=(FONT_FAMILY, 18), width=32
        )
        barcode_entry.pack(pady=12)
        barcode_entry.bind("<Return>", self.scan_barcode)
        barcode_entry.focus()
        ttk.Button(
            self.center_frame,
            text="🔍 Tìm kiếm",
            bootstyle=PRIMARY,
            command=self.scan_barcode,
            style="info.TButton",
            width=16
        ).pack(pady=18)

        def update_images(event=None):
            w = self.scanner_frame.winfo_width()
            h = self.scanner_frame.winfo_height()
            img_h = max(int(h * 0.9), 10)      # Đảm bảo >0
            img_w = max(int(w * 0.25), 10)     # Đảm bảo >0
            left_img = self.img_left_src.resize((img_w, img_h), Image.LANCZOS)
            right_img = self.img_right_src.resize((img_w, img_h), Image.LANCZOS)
            self.left_photo = ImageTk.PhotoImage(left_img)
            self.right_photo = ImageTk.PhotoImage(right_img)
            self.label_left.config(image=self.left_photo)
            self.label_left.image = self.left_photo
            self.label_right.config(image=self.right_photo)
            self.label_right.image = self.right_photo

        self.scanner_frame.bind("<Configure>", update_images)
        update_images()

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
        # Load lại từ database và sort theo ID tăng dần
        products = self.db.get_all_products()
        products = sorted(products, key=lambda x: x[0])  # x[0] là ID
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

        if products:
            for product in products:
                barcode = product[1]
                self.product_tree.insert(
                    "", "end", text=product[0], values=(barcode, product[2])
                )
        else:
            # Hiện thông báo không tìm thấy
            self.product_tree.insert(
                "", "end",
                text="",
                values=("", ""),
            )
            self.product_tree.heading("name", text="Tên sản phẩm", anchor="w")
            # Đổi màu và font cho dòng thông báo
            self.product_tree.item(self.product_tree.get_children()[0], tags=("notfound",))
            style = ttk.Style()
            style.configure("notfound.Treeview", foreground="blue", font=(FONT_FAMILY, 21, "normal"))
            self.product_tree.tag_configure("notfound", foreground="blue", font=(FONT_FAMILY, 21, "normal"))
            self.product_tree.set(self.product_tree.get_children()[0], "name", "Vui lòng nhập tìm kiếm chính xác hơn!")


    def add_product(self):
        """Thêm sản phẩm mới"""
        self.product_dialog()

    def delete_product(self):
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn sản phẩm cần xóa!")
            return

        products_to_delete = []
        for item_id in selected:
            item = self.product_tree.item(item_id)
            barcode = item["values"][0]
            product_name = item["values"][1] if len(item["values"]) > 1 else ""
            products_to_delete.append((barcode, product_name))

        # Tạo popup xác nhận custom
        confirm = tk.Toplevel(self.root)
        confirm.title("Xác nhận xoá sản phẩm")
        screen_width = confirm.winfo_screenwidth()
        screen_height = confirm.winfo_screenheight()
        w = min(800, int(screen_width * 0.5))
        h = min(400, int(screen_height * 0.3))
        confirm.geometry(f"{w}x{h}")
        confirm.transient(self.root)
        confirm.grab_set()
        confirm.resizable(True, True)
        confirm.minsize(600, 250)
        confirm.update_idletasks()
        x = (screen_width // 2) - (w // 2)
        y = (screen_height // 2) - (h // 2)
        confirm.geometry(f"{w}x{h}+{x}+{y}")

        names = "\n".join([f'- {name}' for _, name in products_to_delete])
        label = ttk.Label(
            confirm,
            text=f'Bạn có chắc muốn xoá {len(products_to_delete)} sản phẩm sau?\n{names}',
            font=(FONT_FAMILY, 18, "bold"),
            wraplength=550,
            anchor="center",
            justify="center"
        )
        label.pack(pady=50, padx=30)

        btn_frame = ttk.Frame(confirm)
        btn_frame.pack(pady=10)

        def do_delete():
            for barcode, _ in products_to_delete:
                self.db.delete_product(barcode)
            self.refresh_products()
            confirm.destroy()
            messagebox.showinfo("Thành công", f'Đã xoá {len(products_to_delete)} sản phẩm!')

        ttk.Button(
            btn_frame, text="Không", command=confirm.destroy, style="danger.TButton", width=12
        ).pack(side=LEFT, padx=10)
        ttk.Button(
            btn_frame, text="Xoá", command=do_delete, style="success.TButton", width=12
        ).pack(side=LEFT, padx=10)

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
        """Dialog thêm/sửa sản phẩm với quản lý nhiều đơn vị tính (đảm bảo Hủy không mất dữ liệu gốc)"""

        dialog = tk.Toplevel(self.root)
        dialog.title("Thêm sản phẩm" if barcode is None else "Sửa sản phẩm")
        
        # Tự động tính toán kích thước dựa trên màn hình
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        w = min(1200, int(screen_width * 0.75))  # 75% chiều rộng màn hình hoặc tối đa 1200px
        h = min(750, int(screen_height * 0.65))  # 65% chiều cao màn hình hoặc tối đa 750px
        
        dialog.geometry(f"{w}x{h}")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)  # Allow resizing
        dialog.minsize(1200, 800)    # Set larger minimum size
        
        # Căn giữa màn hình
        dialog.update_idletasks()
        x = (screen_width // 2) - (w // 2)
        y = (screen_height // 2) - (h // 2)
        dialog.geometry(f"{w}x{h}+{x}+{y}")

        barcode_var = tk.StringVar(value=values[0] if values else (barcode or ""))
        name_var = tk.StringVar(value=values[1] if values else "")

        # Lấy đơn vị tính gốc từ DB (biến tạm, không thao tác trực tiếp DB khi thêm/xoá)
        units_origin = list(self.db.get_units_by_barcode(barcode_var.get().strip()))
        units_temp = [list(u) for u in units_origin]  # Sao chép để thao tác tạm

        input_frame = ttk.Frame(dialog)
        input_frame.pack(fill=X, padx=30, pady=20)
        ttk.Label(input_frame, text="Mã vạch:", font=(FONT_FAMILY, 19, "bold")).grid(row=0, column=0, sticky="w", pady=8)
        barcode_entry = ttk.Entry(input_frame, textvariable=barcode_var, font=(FONT_FAMILY, 17))
        barcode_entry.grid(row=0, column=1, sticky="ew", padx=(16,0), pady=8)
        ttk.Label(input_frame, text="Tên sản phẩm:", font=(FONT_FAMILY, 19, "bold")).grid(row=1, column=0, sticky="w", pady=8)
        name_entry = ttk.Entry(input_frame, textvariable=name_var, font=(FONT_FAMILY, 17))
        name_entry.grid(row=1, column=1, sticky="ew", padx=(16,0), pady=8)
        input_frame.columnconfigure(1, weight=1)

        unit_frame = ttk.LabelFrame(dialog)
        unit_frame.pack(fill=BOTH, expand=True, padx=30, pady=20)
        ttk.Label(unit_frame, text="Đơn vị tính & Giá bán", font=(FONT_FAMILY, 19, "bold")).pack(anchor=W, padx=12, pady=12)

        add_unit_frame = ttk.Frame(unit_frame)
        add_unit_frame.pack(fill=X, padx=8, pady=10)
        add_unit_var = tk.StringVar()
        add_price_var = tk.StringVar()

        unit_entry = ttk.Entry(add_unit_frame, textvariable=add_unit_var, width=20, font=(FONT_FAMILY, 15))
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

        price_entry = ttk.Entry(add_unit_frame, textvariable=add_price_var, width=20, font=(FONT_FAMILY, 15))
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

        # Treeview đơn vị tính
        tree_container = ttk.Frame(unit_frame)
        tree_container.pack(fill=BOTH, expand=True, padx=10, pady=15)
        
        style = ttk.Style()
        style.configure("Treeview.Heading", font=(FONT_FAMILY, 16, "bold"))
        style.configure("Treeview", font=(FONT_FAMILY, 14), rowheight=35)
        unit_tree = ttk.Treeview(tree_container, columns=("unit", "price", "action"), show="headings", height=4)
        unit_tree.heading("unit", text="Đơn vị tính")
        unit_tree.heading("price", text="Giá bán (VND)")
        unit_tree.heading("action", text="Thao tác")
        unit_tree.column("unit", width=300, anchor="center")
        unit_tree.column("price", width=300, anchor="center")
        unit_tree.column("action", width=150, anchor="center")
        unit_tree.pack(side=LEFT, fill=BOTH, expand=True)
        unit_scroll = ttk.Scrollbar(tree_container, orient=VERTICAL, command=unit_tree.yview)
        unit_tree.configure(yscrollcommand=unit_scroll.set)
        unit_scroll.pack(side=RIGHT, fill=Y)

        def load_units():
            unit_tree.delete(*unit_tree.get_children())
            for idx, u in enumerate(units_temp):
                unit_tree.insert(
                    "", "end", iid=idx,
                    values=(u[2], f"{u[3]:,.0f}", "🗑️ Xoá")
                )

        def add_unit():
            barcode = barcode_var.get().strip()
            unit_raw = add_unit_var.get().strip()
            price_str = add_price_var.get().replace(".", "").replace(",", ".")
            # Validate: Đơn vị không được để trống hoặc là placeholder "Đơn vị"
            if not unit_raw or unit_raw == "Đơn vị":
                messagebox.showerror("Lỗi", "Cần nhập đơn vị tính!")
                return
            # Validate: Giá bán không được để trống
            if not price_str or price_str == "Giá bán":
                messagebox.showerror("Lỗi", "Cần nhập giá bán!")
                return
            # Validate: Giá bán phải là số và > 0
            try:
                price = float(price_str)
                if price <= 0:
                    raise ValueError
            except:
                messagebox.showerror("Lỗi", "Giá bán phải là số > 0!")
                return
            unit_check = unit_raw.lower()
            unit_save = unit_raw.capitalize()
            existed_units = [u[2].lower() for u in units_temp]
            if unit_check in existed_units:
                messagebox.showerror("Lỗi", "Trùng mã vạch + đơn vị!")
                return
            units_temp.append([None, barcode, unit_save, price])
            load_units()
            add_unit_var.set("")
            add_price_var.set("")
            restore_unit_placeholder(None)
            restore_price_placeholder(None)

        def delete_unit(idx):
            del units_temp[int(idx)]
            load_units()

        def show_delete_unit_popup(idx, unit_name):
            popup = tk.Toplevel(dialog)
            popup.title("Xác nhận xoá đơn vị tính")
            
            # Tự động tính toán kích thước dựa trên màn hình
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            w = min(700, int(screen_width * 0.45))  # 45% chiều rộng màn hình hoặc tối đa 700px
            h = min(350, int(screen_height * 0.25))  # 25% chiều cao màn hình hoặc tối đa 350px
            
            popup.geometry(f"{w}x{h}")
            popup.transient(dialog)
            popup.grab_set()
            popup.resizable(True, True)  # Allow resizing
            popup.minsize(550, 280)     # Set minimum size
            
            # Căn giữa màn hình
            popup.update_idletasks()
            x = (screen_width // 2) - (w // 2)
            y = (screen_height // 2) - (h // 2)
            popup.geometry(f"{w}x{h}+{x}+{y}")
            label = ttk.Label(
                popup,
                text=f'Bạn có chắc muốn xoá đơn vị tính "{unit_name}" này?',
                font=(FONT_FAMILY, 16, "bold"),
                wraplength=550,
                anchor="center",
                justify="center"
            )
            label.pack(pady=60, padx=30)
            btn_frame = ttk.Frame(popup)
            btn_frame.pack(pady=30)
            ttk.Button(
                btn_frame, text="Không", command=popup.destroy, style="danger.TButton", width=12
            ).pack(side=LEFT, padx=10)
            def do_delete():
                delete_unit(idx)
                popup.destroy()
                messagebox.showinfo("Thành công", f'Đã xoá đơn vị tính "{unit_name}"!')
            ttk.Button(
                btn_frame, text="Xoá", command=do_delete, style="success.TButton", width=12
            ).pack(side=LEFT, padx=10)

        unit_tree.bind("<Button-1>", lambda event: (
            lambda region, col, row: show_delete_unit_popup(row, unit_tree.item(row)["values"][0])
            if region == "cell" and col == "#3" and row else None
        )(
            unit_tree.identify("region", event.x, event.y),
            unit_tree.identify_column(event.x),
            unit_tree.identify_row(event.y)
        ))

        load_units()

        ttk.Button(
            add_unit_frame,
            text="➕",
            width=3,
            command=add_unit,
            style="success.TButton",
        ).pack(side=LEFT, padx=4)

        # Google Images
        def open_google_images():
            import webbrowser
            q = barcode_var.get().strip() or name_var.get().strip()
            if not q:
                messagebox.showinfo("Gợi ý", "Nhập mã vạch hoặc tên sản phẩm trước!")
                return
            url = f"https://www.google.com/search?tbm=isch&q={q}"
            webbrowser.open(url)

        # Google Images frame
        google_frame = ttk.Frame(dialog)
        google_frame.pack(fill=X, padx=30, pady=(15, 10))
        ttk.Label(
            google_frame, text="Gợi ý ảnh Google Images:", font=(FONT_FAMILY, 18, "bold")
        ).pack(side=LEFT)
        ttk.Button(
            google_frame,
            text="🔎 Mở Google Images",
            command=open_google_images,
            style="info.TButton",
        ).pack(side=LEFT, padx=10)

        # Buttons frame
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)

        def save_product():
            barcode = barcode_var.get().strip()
            name = name_var.get().strip()
            # Kiểm tra các trường input
            if not barcode or not name:
                messagebox.showerror("Lỗi", "Điền đủ mã vạch và tên sản phẩm!")
                return
            if not units_temp or any(not u[2] or not u[3] for u in units_temp):
                messagebox.showerror("Lỗi", "Sản phẩm phải có ít nhất 1 đơn vị tính và giá bán!")
                return
            existed = self.db.get_product_by_barcode(barcode)
            # Nếu đang sửa sản phẩm (chỉ khi values truyền vào từ nút Sửa, không phải Thêm mới)
            is_edit = values is not None and len(values) >= 2 and existed and barcode == values[0]
            if is_edit:
                self.db.update_product(barcode, name)
                self.db.delete_all_units_by_barcode(barcode)
                for u in units_temp:
                    self.db.add_unit(barcode, u[2], u[3])
                messagebox.showinfo("Thành công", "Đã cập nhật sản phẩm!")
                dialog.destroy()
                self.refresh_products()
                return
            # Nếu đang thêm mới sản phẩm (mã vạch chưa tồn tại)
            if not existed:
                added = self.db.add_product(barcode, name)
                if added:
                    self.db.delete_all_units_by_barcode(barcode)
                    for u in units_temp:
                        self.db.add_unit(barcode, u[2], u[3])
                    messagebox.showinfo("Thành công", "Đã thêm sản phẩm!")
                    dialog.destroy()
                    self.refresh_products()
                else:
                    messagebox.showerror("Lỗi", "Trùng mã vạch!")
                return
            # Nếu mã vạch đã tồn tại nhưng không phải sửa sản phẩm đó
            messagebox.showerror("Lỗi", "Trùng mã vạch!")

        ttk.Button(
            btn_frame,
            text="💾 Lưu",
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

    # ----

    # Scanner methods
    def scan_barcode(self, event=None):
        """Quét mã vạch và hiện popup thông tin sản phẩm nếu có"""
        barcode = self.barcode_var.get().strip()
        if not barcode:
            return

        # Xóa/thêm label thêm mới nếu có
        if hasattr(self, "add_new_label") and self.add_new_label:
            self.add_new_label.destroy()
            self.add_new_label = None

        product = self.db.get_product_by_barcode(barcode)
        units = self.db.get_units_by_barcode(barcode)

        if product:
            # Popup thông tin sản phẩm (layout giống sửa sản phẩm)
            popup = tk.Toplevel(self.root)
            popup.title("Thông tin sản phẩm")
            
            # Tự động tính toán kích thước dựa trên màn hình
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            w = min(1000, int(screen_width * 0.65))  # 65% chiều rộng màn hình hoặc tối đa 1000px
            h = min(600, int(screen_height * 0.55))  # 55% chiều cao màn hình hoặc tối đa 600px
            
            popup.geometry(f"{w}x{h}")
            popup.transient(self.root)
            popup.grab_set()
            popup.resizable(True, True)  # Allow resizing
            popup.minsize(1100, 700)     # Set larger minimum size
            
            # Căn giữa màn hình
            popup.update_idletasks()
            x = (screen_width // 2) - (w // 2)
            y = (screen_height // 2) - (h // 2)
            popup.geometry(f"{w}x{h}+{x}+{y}")

            # Thông tin barcode, tên sản phẩm (dùng grid cho thẳng hàng)
            info_frame = ttk.Frame(popup)
            info_frame.pack(fill=X, padx=30, pady=15)
            ttk.Label(info_frame, text="Mã vạch:", font=(FONT_FAMILY, 19, "bold")).grid(row=0, column=0, sticky="w", pady=8)
            ttk.Label(info_frame, text=product[1], font=(FONT_FAMILY, 17)).grid(row=0, column=1, sticky="ew", padx=(16,0), pady=8)
            ttk.Label(info_frame, text="Tên sản phẩm:", font=(FONT_FAMILY, 19, "bold")).grid(row=1, column=0, sticky="w", pady=8)
            ttk.Label(info_frame, text=product[2], font=(FONT_FAMILY, 17)).grid(row=1, column=1, sticky="ew", padx=(16,0), pady=8)
            info_frame.columnconfigure(1, weight=1)

            # Đơn vị tính & Giá bán
            unit_frame = ttk.LabelFrame(popup)
            unit_frame.pack(fill=BOTH, expand=True, padx=30, pady=15)
            ttk.Label(unit_frame, text="Đơn vị tính & Giá bán", font=(FONT_FAMILY, 19, "bold")).pack(anchor=W, padx=12, pady=12)

            # Create container for treeview and scrollbar
            tree_container = ttk.Frame(unit_frame)
            tree_container.pack(fill=BOTH, expand=True, padx=10, pady=15)

            style = ttk.Style()
            style.configure("Treeview.Heading", font=(FONT_FAMILY, 16, "bold"))
            style.configure("Treeview", font=(FONT_FAMILY, 14))
            style.configure("Treeview", rowheight=35)

            unit_tree = ttk.Treeview(
                tree_container, columns=("unit", "price"), show="headings", height=4
            )
            unit_tree.heading("unit", text="Đơn vị tính")
            unit_tree.heading("price", text="Giá bán (VND)")
            unit_tree.column("unit", width=400, anchor="center")
            unit_tree.column("price", width=400, anchor="center")
            unit_tree.pack(side=LEFT, fill=BOTH, expand=True)

            unit_scroll = ttk.Scrollbar(
                tree_container, orient=VERTICAL, command=unit_tree.yview
            )
            unit_tree.configure(yscrollcommand=unit_scroll.set)
            unit_scroll.pack(side=RIGHT, fill=Y)

            for u in units:
                unit_tree.insert("", "end", values=(u[2], f"{u[3]:,.0f}"))
            # Gợi ý Google Images (giống popup sửa sản phẩm)
            def open_google_images():
                import webbrowser
                q = product[1] or product[2]
                if not q:
                    messagebox.showinfo("Gợi ý", "Nhập mã vạch hoặc tên sản phẩm trước!")
                    return
                url = f"https://www.google.com/search?tbm=isch&q={q}"
                webbrowser.open(url)

            # Google Images frame
            google_frame = ttk.Frame(popup)
            google_frame.pack(fill=X, padx=30, pady=(15, 10))
            ttk.Label(
                google_frame, text="Gợi ý ảnh Google Images:", font=(FONT_FAMILY, 15)
            ).pack(side=LEFT)
            ttk.Button(
                google_frame,
                text="🔎 Mở Google Images",
                command=open_google_images,
                style="info.TButton",
            ).pack(side=LEFT, padx=10)

            # Nút Sửa và Đóng
            def open_edit():
                popup.destroy()
                self.product_dialog(barcode=barcode, values=[product[1], product[2]])

            btn_frame = ttk.Frame(popup)
            btn_frame.pack(pady=20)
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
            
            # Tự động tính toán kích thước dựa trên màn hình
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            w = min(1000, int(screen_width * 0.6))  # 60% chiều rộng màn hình hoặc tối đa 1000px
            h = min(600, int(screen_height * 0.5))  # 50% chiều cao màn hình hoặc tối đa 600px
            
            popup.geometry(f"{w}x{h}")
            popup.transient(self.root)
            popup.grab_set()
            popup.resizable(True, True)  # Allow resizing for better UX
            popup.minsize(800, 400)     # Set minimum size
            
            # Căn giữa màn hình
            popup.update_idletasks()
            x = (screen_width // 2) - (w // 2)
            y = (screen_height // 2) - (h // 2)
            popup.geometry(f"{w}x{h}+{x}+{y}")
            label = ttk.Label(
                popup,
                text=f"Không tìm thấy sản phẩm với mã vạch: {barcode}",
                font=(FONT_FAMILY, 18),
                wraplength=850,
                anchor="center",
                justify="center"
            )
            label.pack(pady=80)

            def open_add():
                popup.destroy()
                self.product_dialog(barcode=barcode, values=[barcode, "", "", "", ""])

            btn_frame = ttk.Frame(popup)
            btn_frame.pack(pady=40)
            
            add_btn = ttk.Button(
                btn_frame,
                text="➕ Thêm sản phẩm mới",
                command=open_add,
                style="success.TButton",
                width=20,
            )
            add_btn.pack(pady=20)
            close_btn = ttk.Button(
                btn_frame,
                text="Đóng",
                command=popup.destroy,
                style="danger.TButton",
                width=20,
            )
            close_btn.pack(pady=10)

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
            result = self.excel_handler.import_from_excel(file_path)
            if isinstance(result, tuple) and len(result) == 3:
                success, message, error_list = result
            else:
                success, message = result
                error_list = []
            self.log_message(message)
            if error_list:
                for err in error_list:
                    self.log_message(f"❌ {err}")
            if success:
                self.refresh_products()
                messagebox.showinfo("Thành công", message)
            else:
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
    app = HyCheckApp()
    app.run()
