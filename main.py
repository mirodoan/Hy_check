import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import platform
import sys
import os
import webbrowser

# ==================== DPI AWARENESS – ĐẶT TRƯỚC KHI IMPORT TKINTER ====================
if platform.system() == "Windows":
    try:
        from ctypes import windll
        try:
            windll.shcore.SetProcessDpiAwareness(2)  # Per-Monitor v2 (Windows 10 1703+)
        except:
            try:
                windll.shcore.SetProcessDpiAwareness(1)  # Per-Monitor v1
            except:
                pass
    except Exception as e:
        print(f"DPI Warning: {e}")

# ==================== FONT & IMAGE ====================
FONT_FAMILY = "Segoe UI" if platform.system() == "Windows" else "Helvetica"

from PIL import Image, ImageTk

# Tương thích Pillow 10.0.0+ (Image.LANCZOS deprecated)
try:
    RESAMPLE_FILTER = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE_FILTER = Image.LANCZOS

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def get_scaled_font(base_size):
    """Tính font size phù hợp với từng hệ điều hành"""
    if platform.system() == "Windows":
        return base_size
    else:  # macOS, Linux
        return base_size + 1  # macOS cần font lớn hơn 1pt

# ==================== TTKBOOTSTRAP ====================
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    BOOTSTRAP_AVAILABLE = True
except ImportError:
    BOOTSTRAP_AVAILABLE = False
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
        self.db = Database()
        self.excel_handler = ExcelHandler(self.db)

        if BOOTSTRAP_AVAILABLE:
            self.root = tb.Window(themename="cosmo")
        else:
            self.root = tk.Tk()
            
        self.root.title("HyCheck - Quản lý sản phẩm")
        
        # Zoomed (maximize) chỉ hoạt động hiệu quả trên Windows
        if platform.system() == "Windows":
            self.root.state('zoomed')
        else:
            self.root.geometry("1200x800")
        
        self.root.minsize(900, 600)

        # ==================== ✅ FIX: STYLE CHO NGƯỜI U60 - CHỮ TO, DỄ ĐỌC ====================
        style = ttk.Style()
        
        # 1. ✅ FIX: Bảng sản phẩm chính - CHỮ TO HƠN, ROWHEIGHT CAO HƠN
        style.configure("Product.Treeview", 
            font=(FONT_FAMILY, get_scaled_font(15)),
            rowheight=48
        )
        style.configure("Product.Treeview.Heading", 
            font=(FONT_FAMILY, get_scaled_font(18), "bold")
        )

        # 2. ✅ FIX: Bảng đơn vị tính trong dialog - CHỮ CỰC TO (QUAN TRỌNG NHẤT!)
        style.configure("Unit.Treeview", 
            font=(FONT_FAMILY, get_scaled_font(19)),  # ✅ TĂNG TỪ 16 LÊN 19
            rowheight=62  # ✅ TĂNG TỪ 52 LÊN 62
        )
        style.configure("Unit.Treeview.Heading", 
            font=(FONT_FAMILY, get_scaled_font(22), "bold")  # ✅ TĂNG TỪ 19 LÊN 22
        )

        # 3. ✅ FIX: Bảng scan - CHỮ CỰC TO (QUAN TRỌNG NHẤT!)
        style.configure("Scan.Treeview", 
            font=(FONT_FAMILY, get_scaled_font(22)),  # ✅ TĂNG TỪ 18 LÊN 22
            rowheight=68  # ✅ TĂNG TỪ 58 LÊN 68
        )
        style.configure("Scan.Treeview.Heading", 
            font=(FONT_FAMILY, get_scaled_font(25), "bold")  # ✅ TĂNG TỪ 21 LÊN 25
        )

        style.configure("TButton", font=(FONT_FAMILY, get_scaled_font(12)))
        style.configure("TNotebook.Tab", font=(FONT_FAMILY, get_scaled_font(16), "bold"), padding=[22, 12])

        self.setup_ui()

    def setup_ui(self):
        """Thiết lập giao diện chính"""
        style = ttk.Style()
        style.map("TNotebook.Tab",
            background=[("selected", "#ff9800"), ("!selected", "#dee7ed")],
            foreground=[("selected", "#fff"), ("!selected", "#333")]
        )
        self.notebook = ttk.Notebook(self.root)
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
        """✅ FIX: Tab quản lý sản phẩm - CHỮ TO, CỘT THẲNG HÀNG"""
        btn_frame = ttk.Frame(self.product_frame)
        btn_frame.pack(fill=X, padx=12, pady=8)

        ttk.Button(btn_frame, text="Thêm sản phẩm", bootstyle=SUCCESS, command=self.add_product).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Sửa", bootstyle=WARNING, command=self.edit_product).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Xóa", bootstyle=DANGER, command=self.delete_product).pack(side=LEFT, padx=5)
        ttk.Button(btn_frame, text="Refresh", command=self.refresh_products).pack(side=LEFT, padx=5)

        # Search
        search_frame = ttk.Frame(self.product_frame)
        search_frame.pack(fill=X, padx=12, pady=6)
        ttk.Label(search_frame, text="Tìm kiếm:", font=(FONT_FAMILY, get_scaled_font(14))).pack(side=LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self.on_search_change())
        ttk.Entry(search_frame, textvariable=self.search_var, font=(FONT_FAMILY, get_scaled_font(14))).pack(side=LEFT, fill=X, expand=True, padx=5)

        # ✅ FIX: Treeview - THẲNG HÀNG, CỘT GẦN NHAU
        tree_frame = ttk.Frame(self.product_frame)
        tree_frame.pack(fill=BOTH, expand=True, padx=12, pady=8)

        self.product_tree = ttk.Treeview(
            tree_frame, 
            columns=("barcode", "name"), 
            show="tree headings", 
            selectmode="extended", 
            style="Product.Treeview"
        )
        
        self.product_tree.column("#0", width=80, anchor="center")
        self.product_tree.column("barcode", width=250, anchor="center")
        self.product_tree.column("name", width=800, anchor="w", stretch=True)
        
        self.product_tree.heading("#0", text="ID")
        self.product_tree.heading("barcode", text="Mã vạch")
        self.product_tree.heading("name", text="Tên sản phẩm")

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.product_tree.yview)
        self.product_tree.configure(yscrollcommand=scroll.set)
        self.product_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)

        self.refresh_products()

    def setup_scanner_tab(self):
        # Lưu ảnh gốc để resize động
        self.img_left_src = Image.open(resource_path("assets/Hy.png"))
        self.img_right_src = Image.open(resource_path("assets/Dan.png"))

        # Tạo label cho ảnh
        self.label_left = tk.Label(self.scanner_frame, borderwidth=0)
        self.label_right = tk.Label(self.scanner_frame, borderwidth=0)
        self.label_left.place(relx=0, rely=0.5, anchor="w")
        self.label_right.place(relx=1, rely=0.5, anchor="e")

        # Nội dung trung tâm
        self.center_frame = ttk.Frame(self.scanner_frame)
        self.center_frame.place(relx=0.5, rely=0.5, anchor="center")

        title_label = ttk.Label(
            self.center_frame, text="🔍 QUÉT MÃ VẠCH SẢN PHẨM", 
            font=(FONT_FAMILY, get_scaled_font(20), "bold")
        )
        title_label.pack(pady=25)
        ttk.Label(
            self.center_frame, text="Nhập hoặc quét mã vạch:", 
            font=(FONT_FAMILY, get_scaled_font(14))
        ).pack(pady=8)
        self.barcode_var = tk.StringVar()
        barcode_entry = ttk.Entry(
            self.center_frame, textvariable=self.barcode_var, 
            font=(FONT_FAMILY, get_scaled_font(16)), width=30
        )
        barcode_entry.pack(pady=10)
        barcode_entry.bind("<Return>", self.scan_barcode)
        barcode_entry.focus()
        ttk.Button(
            self.center_frame,
            text="🔍 Tìm kiếm",
            bootstyle=PRIMARY,
            command=self.scan_barcode,
            style="info.TButton",
            width=14
        ).pack(pady=15)

        def update_images(event=None):
            w = self.scanner_frame.winfo_width()
            h = self.scanner_frame.winfo_height()
            img_h = max(int(h * 0.9), 10)
            img_w = max(int(w * 0.25), 10)
            
            left_img = self.img_left_src.resize((img_w, img_h), RESAMPLE_FILTER)
            right_img = self.img_right_src.resize((img_w, img_h), RESAMPLE_FILTER)
            
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

        self.log_text = tk.Text(log_frame, height=5, font=("Consolas", 10))
        log_scrollbar = ttk.Scrollbar(
            log_frame, orient=VERTICAL, command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        self.log_text.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)
        log_scrollbar.pack(side=RIGHT, fill=Y, pady=10)

    # Product management methods
    def refresh_products(self):
        """Refresh danh sách sản phẩm"""
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)
        products = self.db.get_all_products()
        products = sorted(products, key=lambda x: x[0])
        for product in products:
            barcode = product[1]
            self.product_tree.insert(
                "", "end", text=product[0], values=(barcode, product[2])
            )

    def on_search_change(self, *args):
        """Xử lý khi search text thay đổi"""
        search_term = self.search_var.get()
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)
        
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
            self.product_tree.insert("", "end", text="", values=("", ""))
            self.product_tree.heading("name", text="Tên sản phẩm", anchor="w")
            self.product_tree.item(self.product_tree.get_children()[0], tags=("notfound",))
            style = ttk.Style()
            style.configure("notfound.Treeview", foreground="blue", font=(FONT_FAMILY, get_scaled_font(16), "normal"))
            self.product_tree.tag_configure("notfound", foreground="blue", font=(FONT_FAMILY, get_scaled_font(16), "normal"))
            self.product_tree.set(self.product_tree.get_children()[0], "name", "Vui lòng nhập tìm kiếm chính xác hơn!")

    def add_product(self):
        """Thêm sản phẩm mới"""
        self.product_dialog()

    def delete_product(self):
        """✅ FIX CUỐI CÙNG: Xóa sản phẩm - HIỂN THỊ ĐẦY ĐỦ"""
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

        confirm = tk.Toplevel(self.root)
        confirm.title("Xác nhận xoá sản phẩm")
        confirm.transient(self.root)
        confirm.grab_set()
        confirm.resizable(True, True)
        
        num_products = len(products_to_delete)
        
        screen_height = confirm.winfo_screenheight()
        screen_width = confirm.winfo_screenwidth()
        
        # ✅ FIX: Usable height an toàn cho Windows
        usable_height = screen_height - 100 if platform.system() == "Windows" else screen_height - 60
        
        # ✅ FIX: Tính toán chính xác - FONT TO HƠN = CHIỀU CAO LỚN HƠN
        title_height = 90  # Tăng từ 80
        product_line_height = 38  # Tăng từ 32 (vì font 14pt cao hơn)
        products_area_height = num_products * product_line_height
        button_height = 110  # Tăng từ 100
        total_padding = 70  # Tăng từ 60
        
        needed_height = title_height + products_area_height + button_height + total_padding
        
        # ✅ FIX: Auto adjust với safety margin
        if needed_height > usable_height * 0.82:
            h = int(usable_height * 0.78)
            use_scrollbar = True
        else:
            h = min(needed_height, int(usable_height * 0.82))
            use_scrollbar = False
        
        w = 750  # Tăng từ 700 để chứa chữ to
        
        min_h = title_height + 100 + button_height + total_padding
        confirm.minsize(650, min(min_h, int(usable_height * 0.75)))
        
        x = (screen_width - w) // 2
        y = max(20, (screen_height - h) // 2 - 30)
        confirm.geometry(f"{w}x{h}+{x}+{y}")

        # ✅ FIX: LAYOUT - BUTTONS LUÔN HIỂN THỊ
        main_container = ttk.Frame(confirm)
        main_container.pack(fill=BOTH, expand=True)
        
        content_frame = ttk.Frame(main_container)
        content_frame.pack(fill=BOTH, expand=True, padx=28, pady=(22, 0))
        
        title_text = f'Bạn có chắc muốn xoá {num_products} sản phẩm sau?'
        title_label = ttk.Label(
            content_frame,
            text=title_text,
            font=(FONT_FAMILY, get_scaled_font(16), "bold"),
            anchor="center",
            justify="center"
        )
        title_label.pack(pady=(0, 20))
        
        if use_scrollbar:
            list_container = ttk.Frame(content_frame)
            list_container.pack(fill=BOTH, expand=True)
            
            canvas = tk.Canvas(list_container, highlightthickness=0)
            scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            for barcode, name in products_to_delete:
                ttk.Label(
                    scrollable_frame,
                    text=f'- {name}',
                    font=(FONT_FAMILY, get_scaled_font(14)),
                    anchor="w"
                ).pack(anchor="w", pady=5, padx=12)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
        else:
            list_frame = ttk.Frame(content_frame)
            list_frame.pack(fill=BOTH, expand=True)
            
            for barcode, name in products_to_delete:
                ttk.Label(
                    list_frame,
                    text=f'- {name}',
                    font=(FONT_FAMILY, get_scaled_font(14)),
                    anchor="w"
                ).pack(anchor="w", pady=5, padx=12)
        
        btn_container = ttk.Frame(main_container)
        btn_container.pack(fill=X, pady=(20, 30))
        
        btn_frame = ttk.Frame(btn_container)
        btn_frame.pack()

        def do_delete():
            for barcode, _ in products_to_delete:
                self.db.delete_product(barcode)
            self.refresh_products()
            confirm.destroy()
            self.root.focus_force()
            messagebox.showinfo("Thành công", f'Đã xoá {num_products} sản phẩm!', parent=self.root)

        ttk.Button(
            btn_frame, text="Không", command=confirm.destroy, style="danger.TButton", width=14
        ).pack(side=LEFT, padx=12)
        ttk.Button(
            btn_frame, text="Xoá", command=do_delete, style="success.TButton", width=14
        ).pack(side=LEFT, padx=12)

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
        """✅ FIX CUỐI CÙNG: Dialog thêm/sửa - HIỂN THỊ ĐẦY ĐỦ KHÔNG BỊ CROP"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Thêm sản phẩm" if barcode is None else "Sửa sản phẩm")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()

        is_macos = platform.system() == "Darwin"
        
        # ✅ FIX: Usable height chính xác
        usable_height = screen_height - 100 if not is_macos else screen_height - 60
        
        # ✅ FIX: Width đủ lớn để chứa font 19pt (TĂNG SIZE!)
        base_w = 1120  # ✅ TĂNG TỪ 1080 LÊN 1120
        w = min(base_w, int(screen_width * 0.90))
        
        # ✅ FIX: Height đủ lớn - TÍNH TOÁN DỰA TRÊN FONT SIZE MỚI
        # Input frame: 2 dòng × 60px (font 18pt) = 120px
        # Unit frame title + add controls: ~160px (TĂNG)
        # Unit treeview: 3 hàng × 62px (rowheight) = 186px (TĂNG)
        # Google frame: 60px
        # Buttons frame: 80px
        # Padding total: 160px
        # TOTAL: ~766px
        base_content_height = 800  # ✅ TĂNG TỪ 750 LÊN 800
        h = min(base_content_height, int(usable_height * 0.96))
        
        # ✅ FIX: Minsize an toàn - không crop content
        dialog.minsize(
            min(920, w - 120),
            min(650, int(usable_height * 0.82))
        )
        
        x = (screen_width - w) // 2
        y = max(10, (screen_height - h) // 2 - 25)
        dialog.geometry(f"{w}x{h}+{x}+{y}")

        barcode_var = tk.StringVar(value=str(values[0]) if values else (barcode or ""))
        name_var = tk.StringVar(value=values[1] if values else "")

        units_origin = list(self.db.get_units_by_barcode(barcode_var.get().strip()))
        units_temp = [list(u) for u in units_origin]

        # ✅ FIX: Input frame - spacing đủ lớn
        input_frame = ttk.Frame(dialog)
        input_frame.pack(fill=X, padx=38, pady=(24, 14))
        
        ttk.Label(input_frame, text="Mã vạch:", font=(FONT_FAMILY, get_scaled_font(18), "bold")).grid(row=0, column=0, sticky="w", pady=12)
        barcode_entry = ttk.Entry(input_frame, textvariable=barcode_var, font=(FONT_FAMILY, get_scaled_font(17)))
        barcode_entry.grid(row=0, column=1, sticky="ew", padx=(14,0), pady=12)
        
        ttk.Label(input_frame, text="Tên sản phẩm:", font=(FONT_FAMILY, get_scaled_font(18), "bold")).grid(row=1, column=0, sticky="w", pady=12)
        name_entry = ttk.Entry(input_frame, textvariable=name_var, font=(FONT_FAMILY, get_scaled_font(17)))
        name_entry.grid(row=1, column=1, sticky="ew", padx=(14,0), pady=12)
        input_frame.columnconfigure(1, weight=1)

        # ✅ FIX: Unit frame - spacing đủ lớn + FONT TO HƠN
        unit_frame = ttk.LabelFrame(dialog, text="Đơn vị tính & Giá bán", padding=(18, 16))  # ✅ TĂNG PADDING
        unit_frame.pack(fill=BOTH, expand=True, padx=38, pady=(0, 14))

        add_unit_frame = ttk.Frame(unit_frame)
        add_unit_frame.pack(fill=X, pady=(0, 16))  # ✅ TĂNG PADY
        add_unit_var = tk.StringVar()
        add_price_var = tk.StringVar()

        unit_entry = ttk.Entry(add_unit_frame, textvariable=add_unit_var, width=16, font=(FONT_FAMILY, get_scaled_font(18)))  # ✅ TĂNG FONT TỪ 16 LÊN 18
        unit_entry.pack(side=LEFT, padx=8)  # ✅ TĂNG PADX
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

        price_entry = ttk.Entry(add_unit_frame, textvariable=add_price_var, width=16, font=(FONT_FAMILY, get_scaled_font(18)))  # ✅ TĂNG FONT TỪ 16 LÊN 18
        price_entry.pack(side=LEFT, padx=8)  # ✅ TĂNG PADX
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

        tree_container = ttk.Frame(unit_frame)
        tree_container.pack(fill=BOTH, expand=True, pady=(0, 16))  # ✅ TĂNG PADY

        unit_tree = ttk.Treeview(
            tree_container,
            columns=("unit", "price", "action"),
            show="headings",
            height=3,
            style="Unit.Treeview"
        )

        unit_tree.heading("unit", text="Đơn vị tính")
        unit_tree.heading("price", text="Giá bán (VND)")
        unit_tree.heading("action", text="")
        
        # ✅ FIX: Column width phù hợp với font 19pt (TĂNG SIZE!)
        unit_tree.column("unit", width=380, anchor="center")  # ✅ TĂNG TỪ 340 LÊN 380
        unit_tree.column("price", width=380, anchor="center")  # ✅ TĂNG TỪ 340 LÊN 380
        unit_tree.column("action", width=180, anchor="center")  # ✅ TĂNG TỪ 170 LÊN 180

        unit_tree.pack(side="left", fill="both", expand=True)
        
        scroll_y = ttk.Scrollbar(tree_container, orient="vertical", command=unit_tree.yview)
        scroll_y.pack(side="right", fill="y")
        unit_tree.configure(yscrollcommand=scroll_y.set)

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
            
            if not unit_raw or unit_raw == "Đơn vị":
                messagebox.showerror("Lỗi", "Cần nhập đơn vị tính!", parent=dialog)
                return
            if not price_str or price_str == "Giá bán":
                messagebox.showerror("Lỗi", "Cần nhập giá bán!", parent=dialog)
                return
            
            try:
                price = float(price_str)
                if price <= 0:
                    raise ValueError
            except:
                messagebox.showerror("Lỗi", "Giá bán phải là số > 0!", parent=dialog)
                return
            
            unit_check = unit_raw.lower()
            unit_save = unit_raw.capitalize()
            existed_units = [u[2].lower() for u in units_temp]
            
            if unit_check in existed_units:
                messagebox.showerror("Lỗi", "Trùng mã vạch + đơn vị!", parent=dialog)
                return
            
            units_temp.append([None, barcode, unit_save, price])
            load_units()
            add_unit_var.set("")
            add_price_var.set("")
            unit_entry.delete(0, tk.END)
            unit_entry.insert(0, "Đơn vị")
            unit_entry.config(foreground="#888")
            price_entry.delete(0, tk.END)
            price_entry.insert(0, "Giá bán")
            price_entry.config(foreground="#888")

        def delete_unit(idx):
            del units_temp[int(idx)]
            load_units()

        def show_delete_unit_popup(idx, unit_name):
            """✅ FIX: Popup xóa unit - SIZE ĐẦY ĐỦ"""
            popup = tk.Toplevel(dialog)
            popup.title("Xác nhận xoá đơn vị tính")
            popup.transient(dialog)
            popup.grab_set()
            popup.resizable(True, True)
            
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            
            w = 680  # Tăng từ 640
            h = 260  # Tăng từ 240
            
            popup.minsize(600, 230)
            
            x = (screen_width - w) // 2
            y = max(25, (screen_height - h) // 2 - 35)
            popup.geometry(f"{w}x{h}+{x}+{y}")
            
            main_container = ttk.Frame(popup)
            main_container.pack(fill=BOTH, expand=True)
            
            message_frame = ttk.Frame(main_container)
            message_frame.pack(fill=BOTH, expand=True, padx=32, pady=(32, 20))
            
            label = ttk.Label(
                message_frame,
                text=f'Bạn có chắc muốn xoá đơn vị tính\n"{unit_name}" này?',
                font=(FONT_FAMILY, get_scaled_font(16), "bold"),
                anchor="center",
                justify="center"
            )
            label.pack()
            
            btn_container = ttk.Frame(main_container)
            btn_container.pack(fill=X, pady=(0, 32))
            
            btn_frame = ttk.Frame(btn_container)
            btn_frame.pack()
            
            ttk.Button(
                btn_frame, text="Không", command=popup.destroy, style="danger.TButton", width=14
            ).pack(side=LEFT, padx=12)
            
            def do_delete():
                delete_unit(idx)
                popup.destroy()
                dialog.focus_force()
                dialog.lift()
                messagebox.showinfo("Thành công", f'Đã xoá đơn vị tính "{unit_name}"!', parent=dialog)
                dialog.focus_force()
            
            ttk.Button(
                btn_frame, text="Xoá", command=do_delete, style="success.TButton", width=14
            ).pack(side=LEFT, padx=12)

        def on_unit_click(event):
            region = unit_tree.identify("region", event.x, event.y)
            if region != "cell":
                return
            col = unit_tree.identify_column(event.x)
            row = unit_tree.identify_row(event.y)
            if col == "#3" and row:
                try:
                    idx = int(row)
                    if 0 <= idx < len(units_temp):
                        unit_name = units_temp[idx][2]
                        show_delete_unit_popup(idx, unit_name)
                except:
                    pass

        unit_tree.bind("<Button-1>", on_unit_click)
        load_units()

        ttk.Button(
            add_unit_frame,
            text="➕",
            width=3,
            command=add_unit,
            style="success.TButton",
        ).pack(side=LEFT, padx=8)  # ✅ TĂNG PADX

        # Google Images
        def open_google_images():
            q = barcode_var.get().strip() or name_var.get().strip()
            if not q:
                messagebox.showinfo("Gợi ý", "Nhập mã vạch hoặc tên sản phẩm trước!", parent=dialog)
                return
            url = f"https://www.google.com/search?tbm=isch&q={q}"
            webbrowser.open(url)

        google_frame = ttk.Frame(dialog)
        google_frame.pack(fill=X, padx=38, pady=(0, 14))
        ttk.Label(
            google_frame, text="Gợi ý ảnh Google Images:", font=(FONT_FAMILY, get_scaled_font(13))
        ).pack(side=LEFT)
        ttk.Button(
            google_frame,
            text="🔎 Mở Google Images",
            command=open_google_images,
            style="info.TButton",
        ).pack(side=LEFT, padx=12)

        # Buttons
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=(0, 26))

        def save_product():
            barcode = barcode_var.get().strip()
            name = name_var.get().strip()
            
            if not barcode or not name:
                messagebox.showerror("Lỗi", "Điền đủ mã vạch và tên sản phẩm!", parent=dialog)
                return
            if not units_temp or any(not u[2] or not u[3] for u in units_temp):
                messagebox.showerror("Lỗi", "Sản phẩm phải có ít nhất 1 đơn vị tính và giá bán!", parent=dialog)
                return
            
            existed = self.db.get_product_by_barcode(barcode)
            is_edit = values is not None and len(values) >= 2 and existed and str(barcode) == str(values[0])
            
            if is_edit:
                self.db.update_product(barcode, name)
                self.db.delete_all_units_by_barcode(barcode)
                for u in units_temp:
                    self.db.add_unit(barcode, u[2], u[3])
                dialog.destroy()
                self.root.focus_force()
                messagebox.showinfo("Thành công", "Đã cập nhật sản phẩm!", parent=self.root)
                self.refresh_products()
                return
            
            if not existed:
                added = self.db.add_product(barcode, name)
                if added:
                    self.db.delete_all_units_by_barcode(barcode)
                    for u in units_temp:
                        self.db.add_unit(barcode, u[2], u[3])
                    dialog.destroy()
                    self.root.focus_force()
                    messagebox.showinfo("Thành công", "Đã thêm sản phẩm!", parent=self.root)
                    self.refresh_products()
                else:
                    messagebox.showerror("Lỗi", "Trùng mã vạch!", parent=dialog)
                return
            
            messagebox.showerror("Lỗi", "Trùng mã vạch!", parent=dialog)

        ttk.Button(
            btn_frame,
            text="💾 Lưu",
            bootstyle=SUCCESS,
            command=save_product,
            style="success.TButton",
            width=13,
        ).pack(side=LEFT, padx=12)
        ttk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            style="danger.TButton",
            width=13,
        ).pack(side=LEFT, padx=12)

    # Scanner methods
    def scan_barcode(self, event=None):
        """✅ FIX CUỐI CÙNG: Scan popup - HIỂN THỊ ĐẦY ĐỦ"""
        barcode = self.barcode_var.get().strip()
        if not barcode:
            return

        if hasattr(self, "add_new_label") and self.add_new_label:
            self.add_new_label.destroy()
            self.add_new_label = None

        product = self.db.get_product_by_barcode(barcode)
        units = self.db.get_units_by_barcode(barcode)

        if product:
            popup = tk.Toplevel(self.root)
            popup.title("Thông tin sản phẩm")
            popup.transient(self.root)
            popup.grab_set()
            popup.resizable(True, True)
            
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            
            num_units = len(units)
            is_macos = platform.system() == "Darwin"
            
            usable_height = screen_height - 100 if not is_macos else screen_height - 60
            
            # ✅ FIX: Tính toán chính xác - FONT TO HƠN = CHIỀU CAO LỚN HƠN
            # Info frame: 2 dòng × 70px (font 20pt) = 140px
            # Unit frame title: 80px (TĂNG)
            # Unit treeview: N × 68px (rowheight) (TĂNG!)
            # Google frame: 70px
            # Buttons: 90px
            # Padding: 140px
            base_height = 420  # ✅ TĂNG TỪ 400 LÊN 420
            unit_line_height = 74  # ✅ TĂNG TỪ 64 LÊN 74 (font 22pt cao hơn)
            content_height = base_height + (num_units * unit_line_height)
            
            max_h = min(780, int(usable_height * 0.96))  # ✅ TĂNG TỪ 720 LÊN 780
            h = min(content_height, max_h)
            
            w = min(1020, int(screen_width * 0.82))  # ✅ TĂNG TỪ 950 LÊN 1020
            
            popup.minsize(
                min(860, w - 90),
                min(540, int(usable_height * 0.80))
            )
            
            x = (screen_width - w) // 2
            y = max(10, (screen_height - h) // 2 - 25)
            popup.geometry(f"{w}x{h}+{x}+{y}")

            # ✅ FIX: Info frame - spacing lớn hơn
            info_frame = ttk.Frame(popup)
            info_frame.pack(fill=X, padx=40, pady=(28, 18))  # ✅ TĂNG PADDING
            
            ttk.Label(info_frame, text="Mã vạch:", font=(FONT_FAMILY, get_scaled_font(20), "bold")).grid(row=0, column=0, sticky="w", pady=14)
            ttk.Label(info_frame, text=product[1], font=(FONT_FAMILY, get_scaled_font(19))).grid(row=0, column=1, sticky="ew", padx=(14,0), pady=14)
            
            ttk.Label(info_frame, text="Tên sản phẩm:", font=(FONT_FAMILY, get_scaled_font(20), "bold")).grid(row=1, column=0, sticky="w", pady=14)
            ttk.Label(info_frame, text=product[2], font=(FONT_FAMILY, get_scaled_font(19))).grid(row=1, column=1, sticky="ew", padx=(14,0), pady=14)
            info_frame.columnconfigure(1, weight=1)

            # ✅ FIX: Unit frame - spacing lớn hơn + FONT TO HƠN
            unit_frame = ttk.LabelFrame(popup, text="", padding=(18, 16))  # ✅ TĂNG PADDING
            unit_frame.pack(fill=BOTH, expand=True, padx=40, pady=(0, 18))  # ✅ TĂNG PADDING

            title_label = ttk.Label(unit_frame, text="Đơn vị tính & Giá bán", font=(FONT_FAMILY, get_scaled_font(20), "bold"))  # ✅ TĂNG FONT TỪ 18 LÊN 20
            title_label.pack(anchor="w", padx=8, pady=(0, 16))  # ✅ TĂNG PADY

            tree_container = ttk.Frame(unit_frame)
            tree_container.pack(fill=BOTH, expand=True)

            unit_tree = ttk.Treeview(
                tree_container, 
                columns=("unit", "price"), 
                show="headings", 
                height=num_units,
                style="Scan.Treeview"
            )
            unit_tree.heading("unit", text="Đơn vị tính")
            unit_tree.heading("price", text="Giá bán (VND)")

            # ✅ FIX: Column width phù hợp font 22pt (TĂNG SIZE!)
            unit_tree.column("unit", width=460, anchor="center")  # ✅ TĂNG TỪ 410 LÊN 460
            unit_tree.column("price", width=460, anchor="center")  # ✅ TĂNG TỪ 410 LÊN 460
            unit_tree.pack(side=LEFT, fill=BOTH, expand=True)

            unit_scroll = ttk.Scrollbar(tree_container, orient=VERTICAL, command=unit_tree.yview)
            unit_tree.configure(yscrollcommand=unit_scroll.set)

            for u in units:
                unit_tree.insert("", "end", values=(u[2], f"{u[3]:,.0f}"))

            # Google Images
            def open_google_images():
                q = product[1] or product[2]
                if not q:
                    messagebox.showinfo("Gợi ý", "Nhập mã vạch hoặc tên sản phẩm trước!", parent=popup)
                    return
                url = f"https://www.google.com/search?tbm=isch&q={q}"
                webbrowser.open(url)

            google_frame = ttk.Frame(popup)
            google_frame.pack(fill=X, padx=40, pady=(0, 18))  # ✅ TĂNG PADDING
            ttk.Label(
                google_frame, text="Gợi ý ảnh Google Images:", font=(FONT_FAMILY, get_scaled_font(13))
            ).pack(side=LEFT)
            ttk.Button(
                google_frame,
                text="🔎 Mở Google Images",
                command=open_google_images,
                style="info.TButton",
            ).pack(side=LEFT, padx=12)

            # Buttons
            def open_edit():
                popup.destroy()
                self.product_dialog(barcode=barcode, values=[product[1], product[2]])

            btn_frame = ttk.Frame(popup)
            btn_frame.pack(pady=(0, 28))  # ✅ TĂNG PADY
            ttk.Button(
                btn_frame,
                text="✏️ Sửa sản phẩm",
                command=open_edit,
                style="success.TButton",
                width=15,
            ).pack(side=LEFT, padx=12)
            ttk.Button(
                btn_frame,
                text="Đóng",
                command=popup.destroy,
                style="danger.TButton",
                width=15,
            ).pack(side=LEFT, padx=12)
        else:
            # Not found popup
            popup = tk.Toplevel(self.root)
            popup.title("Không tìm thấy sản phẩm")
            popup.transient(self.root)
            popup.grab_set()
            popup.resizable(False, False)
            
            w = 740  # Tăng từ 700
            h = 280  # Tăng từ 260
            
            screen_width = popup.winfo_screenwidth()
            screen_height = popup.winfo_screenheight()
            x = (screen_width - w) // 2
            y = (screen_height - h) // 2
            popup.geometry(f"{w}x{h}+{x}+{y}")
            
            label = ttk.Label(
                popup,
                text=f"Không tìm thấy sản phẩm với mã vạch: {barcode}",
                font=(FONT_FAMILY, get_scaled_font(15)),
                wraplength=680,
                anchor="center",
                justify="center"
            )
            label.pack(pady=(45, 32))

            def open_add():
                popup.destroy()
                self.product_dialog(barcode=barcode)

            btn_frame = ttk.Frame(popup)
            btn_frame.pack(pady=(0, 35))
            
            ttk.Button(
                btn_frame,
                text="➕ Thêm sản phẩm mới",
                command=open_add,
                style="success.TButton",
                width=19,
            ).pack(pady=(0, 14))
            ttk.Button(
                btn_frame,
                text="Đóng",
                command=popup.destroy,
                style="danger.TButton",
                width=19,
            ).pack()

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