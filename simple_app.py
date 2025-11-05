#!/usr/bin/env python3
"""
KiotCheck - Simple Terminal Version (No Dependencies)
Chương trình quản lý sản phẩm đơn giản cho cửa hàng nhỏ
"""

import os
import sys
import csv
import json
from database import Database

class SimpleExcelHandler:
    def __init__(self, db: Database):
        self.db = db
    
    def import_from_csv(self, file_path):
        """Import sản phẩm từ CSV"""
        try:
            success_count = 0
            error_list = []
            
            with open(file_path, 'r', encoding='utf-8') as file:
                csv_reader = csv.DictReader(file)
                
                # Kiểm tra columns cần thiết
                required_cols = ['barcode', 'name', 'unit', 'price']
                if not all(col in csv_reader.fieldnames for col in required_cols):
                    missing = [col for col in required_cols if col not in csv_reader.fieldnames]
                    return False, f"Thiếu cột: {', '.join(missing)}"
                
                # Import từng dòng
                for row_num, row in enumerate(csv_reader, start=2):
                    try:
                        barcode = str(row['barcode']).strip()
                        name = str(row['name']).strip()
                        unit = str(row['unit']).strip()
                        price = float(row['price'])
                        
                        if self.db.add_product(barcode, name, unit, price):
                            success_count += 1
                        else:
                            error_list.append(f"Dòng {row_num}: Trùng mã vạch + đơn vị")
                            
                    except Exception as e:
                        error_list.append(f"Dòng {row_num}: {str(e)}")
            
            return True, f"Import thành công {success_count} sản phẩm. Lỗi: {len(error_list)}"
            
        except Exception as e:
            return False, f"Lỗi đọc file: {str(e)}"
    
    def export_to_csv(self, file_path):
        """Export sản phẩm ra CSV"""
        try:
            products = self.db.get_all_products()
            
            with open(file_path, 'w', newline='', encoding='utf-8') as file:
                fieldnames = ['barcode', 'name', 'unit', 'price']
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                
                writer.writeheader()
                for product in products:
                    writer.writerow({
                        'barcode': product[1],
                        'name': product[2],
                        'unit': product[3],
                        'price': product[4]
                    })
            
            return True, f"Export thành công {len(products)} sản phẩm"
            
        except Exception as e:
            return False, f"Lỗi export: {str(e)}"
    
    def create_template_csv(self, file_path):
        """Tạo template CSV mẫu"""
        try:
            template_data = [
                {'barcode': '8936036018622', 'name': 'Nước ngọt Coca Cola', 'unit': 'chai', 'price': 15000},
                {'barcode': '8934673123456', 'name': 'Bánh mì sandwich', 'unit': 'cái', 'price': 25000}
            ]
            
            with open(file_path, 'w', newline='', encoding='utf-8') as file:
                fieldnames = ['barcode', 'name', 'unit', 'price']
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                
                writer.writeheader()
                for row in template_data:
                    writer.writerow(row)
                    
            return True, f"Đã tạo template: {file_path}"
            
        except Exception as e:
            return False, f"Lỗi tạo template: {str(e)}"

class KiotCheckSimple:
    def __init__(self):
        self.db = Database()
        self.csv_handler = SimpleExcelHandler(self.db)
        
    def clear_screen(self):
        """Xóa màn hình"""
        os.system('clear' if os.name == 'posix' else 'cls')
        
    def print_header(self):
        """In header"""
        print("=" * 60)
        print("🏪 KIOTCHECK - QUẢN LÝ SẢN PHẨM (Phiên bản đơn giản)")
        print("=" * 60)
        
    def print_menu(self):
        """In menu chính"""
        print("\n📋 MENU CHÍNH:")
        print("1. 📦 Quản lý sản phẩm")
        print("2. 🔍 Quét mã vạch")
        print("3. 📊 Import/Export CSV") 
        print("4. ❌ Thoát")
        print("-" * 60)
        
    def product_menu(self):
        """Menu quản lý sản phẩm"""
        while True:
            self.clear_screen()
            self.print_header()
            print("\n📦 QUẢN LÝ SẢN PHẨM")
            print("1. Xem danh sách sản phẩm")
            print("2. Thêm sản phẩm mới")
            print("3. Tìm kiếm sản phẩm")
            print("4. Sửa sản phẩm")
            print("5. Xóa sản phẩm")
            print("6. Quay lại menu chính")
            
            choice = input("\n👉 Chọn chức năng (1-6): ").strip()
            
            if choice == "1":
                self.show_all_products()
            elif choice == "2":
                self.add_product()
            elif choice == "3":
                self.search_products()
            elif choice == "4":
                self.edit_product()
            elif choice == "5":
                self.delete_product()
            elif choice == "6":
                break
            else:
                print("❌ Lựa chọn không hợp lệ!")
                input("Nhấn Enter để tiếp tục...")
                
    def show_all_products(self):
        """Hiển thị tất cả sản phẩm"""
        products = self.db.get_all_products()
        
        print("\n📋 DANH SÁCH SẢN PHẨM:")
        print("-" * 80)
        print(f"{'ID':<5} {'Mã vạch':<15} {'Tên sản phẩm':<25} {'Đơn vị':<10} {'Giá':<15}")
        print("-" * 80)
        
        if products:
            for product in products:
                print(f"{product[0]:<5} {product[1]:<15} {product[2]:<25} {product[3]:<10} {product[4]:>13,.0f}")
        else:
            print("❌ Chưa có sản phẩm nào!")
            
        print("-" * 80)
        print(f"📊 Tổng số: {len(products)} sản phẩm")
        input("\nNhấn Enter để tiếp tục...")
        
    def add_product(self):
        """Thêm sản phẩm mới"""
        print("\n➕ THÊM SẢN PHẨM MỚI")
        print("-" * 40)
        
        try:
            barcode = input("Mã vạch: ").strip()
            if not barcode:
                print("❌ Mã vạch không được để trống!")
                input("Nhấn Enter để tiếp tục...")
                return
                
            name = input("Tên sản phẩm: ").strip()
            if not name:
                print("❌ Tên sản phẩm không được để trống!")
                input("Nhấn Enter để tiếp tục...")
                return
                
            unit = input("Đơn vị tính: ").strip()
            if not unit:
                print("❌ Đơn vị tính không được để trống!")
                input("Nhấn Enter để tiếp tục...")
                return
                
            price_str = input("Giá: ").strip()
            price = float(price_str)
            
            if price <= 0:
                print("❌ Giá phải lớn hơn 0!")
                input("Nhấn Enter để tiếp tục...")
                return
                
            if self.db.add_product(barcode, name, unit, price):
                print("✅ Đã thêm sản phẩm thành công!")
            else:
                print("❌ Lỗi: Trùng mã vạch + đơn vị tính!")
                
        except ValueError:
            print("❌ Giá phải là số!")
        except Exception as e:
            print(f"❌ Lỗi: {str(e)}")
            
        input("Nhấn Enter để tiếp tục...")
        
    def search_products(self):
        """Tìm kiếm sản phẩm"""
        search_term = input("\n🔍 Nhập từ khóa tìm kiếm: ").strip()
        
        if not search_term:
            print("❌ Từ khóa không được để trống!")
            input("Nhấn Enter để tiếp tục...")
            return
            
        products = self.db.search_products(search_term)
        
        print(f"\n🔍 KẾT QUẢ TÌM KIẾM: '{search_term}'")
        print("-" * 80)
        print(f"{'ID':<5} {'Mã vạch':<15} {'Tên sản phẩm':<25} {'Đơn vị':<10} {'Giá':<15}")
        print("-" * 80)
        
        if products:
            for product in products:
                print(f"{product[0]:<5} {product[1]:<15} {product[2]:<25} {product[3]:<10} {product[4]:>13,.0f}")
        else:
            print("❌ Không tìm thấy sản phẩm nào!")
            
        print("-" * 80)
        print(f"📊 Tìm thấy: {len(products)} sản phẩm")
        input("\nNhấn Enter để tiếp tục...")
        
    def edit_product(self):
        """Sửa sản phẩm"""
        try:
            product_id = int(input("\n✏️ Nhập ID sản phẩm cần sửa: ").strip())
            
            # Lấy thông tin sản phẩm hiện tại
            products = self.db.get_all_products()
            current_product = None
            
            for product in products:
                if product[0] == product_id:
                    current_product = product
                    break
                    
            if not current_product:
                print("❌ Không tìm thấy sản phẩm!")
                input("Nhấn Enter để tiếp tục...")
                return
                
            print(f"\n📝 THÔNG TIN HIỆN TẠI:")
            print(f"Mã vạch: {current_product[1]}")
            print(f"Tên: {current_product[2]}")
            print(f"Đơn vị: {current_product[3]}")
            print(f"Giá: {current_product[4]:,.0f}")
            
            print("\n✏️ NHẬP THÔNG TIN MỚI (Enter để giữ nguyên):")
            
            barcode = input(f"Mã vạch [{current_product[1]}]: ").strip()
            if not barcode:
                barcode = current_product[1]
                
            name = input(f"Tên sản phẩm [{current_product[2]}]: ").strip()
            if not name:
                name = current_product[2]
                
            unit = input(f"Đơn vị [{current_product[3]}]: ").strip()
            if not unit:
                unit = current_product[3]
                
            price_input = input(f"Giá [{current_product[4]:,.0f}]: ").strip()
            if price_input:
                price = float(price_input)
                if price <= 0:
                    print("❌ Giá phải lớn hơn 0!")
                    input("Nhấn Enter để tiếp tục...")
                    return
            else:
                price = current_product[4]
                
            if self.db.update_product(product_id, barcode, name, unit, price):
                print("✅ Đã cập nhật sản phẩm thành công!")
            else:
                print("❌ Lỗi: Trùng mã vạch + đơn vị tính!")
                
        except ValueError:
            print("❌ ID hoặc giá phải là số!")
        except Exception as e:
            print(f"❌ Lỗi: {str(e)}")
            
        input("Nhấn Enter để tiếp tục...")
        
    def delete_product(self):
        """Xóa sản phẩm"""
        try:
            product_id = int(input("\n🗑️ Nhập ID sản phẩm cần xóa: ").strip())
            
            confirm = input("❓ Bạn có chắc muốn xóa? (y/N): ").strip().lower()
            
            if confirm == 'y':
                self.db.delete_product(product_id)
                print("✅ Đã xóa sản phẩm!")
            else:
                print("❌ Đã hủy!")
                
        except ValueError:
            print("❌ ID phải là số!")
        except Exception as e:
            print(f"❌ Lỗi: {str(e)}")
            
        input("Nhấn Enter để tiếp tục...")
        
    def scanner_menu(self):
        """Menu quét mã vạch"""
        while True:
            self.clear_screen()
            self.print_header()
            print("\n🔍 QUÉT MÃ VẠCH")
            print("Nhập mã vạch để tìm kiếm (hoặc 'q' để thoát)")
            print("-" * 60)
            
            barcode = input("👉 Mã vạch: ").strip()
            
            if barcode.lower() == 'q':
                break
                
            if not barcode:
                continue
                
            products = self.db.get_product_by_barcode(barcode)
            
            print(f"\n🔍 KẾT QUẢ CHO MÃ VẠCH: {barcode}")
            print("-" * 60)
            
            if products:
                for product in products:
                    print(f"📦 Tên: {product[2]}")
                    print(f"📏 Đơn vị: {product[3]}")
                    print(f"💰 Giá: {product[4]:,.0f} VNĐ")
                    print("-" * 30)
                print(f"✅ Tìm thấy {len(products)} sản phẩm")
            else:
                print("❌ Không tìm thấy sản phẩm!")
                
            input("\nNhấn Enter để tiếp tục...")
            
    def csv_menu(self):
        """Menu import/export CSV"""
        while True:
            self.clear_screen()
            self.print_header()
            print("\n📊 IMPORT/EXPORT CSV")
            print("1. Import từ CSV")
            print("2. Export ra CSV")
            print("3. Tạo template mẫu")
            print("4. Quay lại menu chính")
            
            choice = input("\n👉 Chọn chức năng (1-4): ").strip()
            
            if choice == "1":
                self.import_csv()
            elif choice == "2":
                self.export_csv()
            elif choice == "3":
                self.create_template()
            elif choice == "4":
                break
            else:
                print("❌ Lựa chọn không hợp lệ!")
                input("Nhấn Enter để tiếp tục...")
                
    def import_csv(self):
        """Import từ CSV"""
        file_path = input("\n📥 Nhập đường dẫn file CSV: ").strip()
        
        if not file_path or not os.path.exists(file_path):
            print("❌ File không tồn tại!")
            input("Nhấn Enter để tiếp tục...")
            return
            
        print("⏳ Đang import...")
        success, message = self.csv_handler.import_from_csv(file_path)
        
        if success:
            print(f"✅ {message}")
        else:
            print(f"❌ {message}")
            
        input("Nhấn Enter để tiếp tục...")
        
    def export_csv(self):
        """Export ra CSV"""
        file_path = input("\n📤 Nhập đường dẫn file xuất (ví dụ: products.csv): ").strip()
        
        if not file_path:
            print("❌ Đường dẫn không được để trống!")
            input("Nhấn Enter để tiếp tục...")
            return
            
        if not file_path.endswith('.csv'):
            file_path += '.csv'
            
        print("⏳ Đang export...")
        success, message = self.csv_handler.export_to_csv(file_path)
        
        if success:
            print(f"✅ {message}")
        else:
            print(f"❌ {message}")
            
        input("Nhấn Enter để tiếp tục...")
        
    def create_template(self):
        """Tạo template mẫu"""
        file_path = input("\n📋 Nhập tên file template (ví dụ: template.csv): ").strip()
        
        if not file_path:
            file_path = "template.csv"
        elif not file_path.endswith('.csv'):
            file_path += '.csv'
            
        success, message = self.csv_handler.create_template_csv(file_path)
        
        if success:
            print(f"✅ {message}")
        else:
            print(f"❌ {message}")
            
        input("Nhấn Enter để tiếp tục...")
        
    def run(self):
        """Chạy ứng dụng"""
        try:
            while True:
                self.clear_screen()
                self.print_header()
                self.print_menu()
                
                choice = input("👉 Chọn chức năng (1-4): ").strip()
                
                if choice == "1":
                    self.product_menu()
                elif choice == "2":
                    self.scanner_menu()
                elif choice == "3":
                    self.csv_menu()
                elif choice == "4":
                    print("\n👋 Cảm ơn bạn đã sử dụng KiotCheck!")
                    break
                else:
                    print("❌ Lựa chọn không hợp lệ!")
                    input("Nhấn Enter để tiếp tục...")
                    
        except KeyboardInterrupt:
            print("\n\n👋 Tạm biệt!")
        except Exception as e:
            print(f"\n❌ Lỗi: {str(e)}")

if __name__ == "__main__":
    print("🚀 Khởi động KiotCheck...")
    app = KiotCheckSimple()
    app.run()