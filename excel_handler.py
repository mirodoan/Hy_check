from operator import index
import pandas as pd
from database import Database

class ExcelHandler:
    def __init__(self, db: Database):
        self.db = db
    
    def import_from_excel(self, file_path):
        """Import sản phẩm từ Excel/CSV theo mẫu: Mã vạch, Tên hàng, Giá bán, ĐVT"""
        try:
            # Đọc file Excel hoặc CSV
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, dtype={'Mã vạch': str})
            else:
                df = pd.read_excel(file_path, dtype={'Mã vạch': str})

            success_count = 0
            error_list = []

            # Đổi tên cột về chuẩn nếu có
            col_map = {
                'Mã vạch': 'barcode',
                'Tên hàng': 'name',
                'Giá bán': 'price',
                'ĐVT': 'unit'
            }
            df = df.rename(columns=col_map)

            required_cols = ['barcode', 'name', 'unit', 'price']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                return False, f"Thiếu cột: {', '.join(missing_cols)}"

            # Import từng dòng
            for index, row in df.iterrows():
                try:
                    barcode = str(row['barcode']).strip()
                    name = str(row['name']).strip()
                    # Đơn vị tính dùng lowercase để kiểm tra trùng, nhưng lưu thì capitalize
                    unit_raw = str(row['unit']).strip()
                    unit_check = unit_raw.lower()
                    unit_save = unit_raw.capitalize()
                    price_str = str(row['price']).replace('.', '').replace(',', '.')
                    price = float(price_str) if price_str else 0.0
                    price = int(price)

                    # Kiểm tra trùng barcode + đơn vị tính (dùng lowercase)
                    existed_units = [u[2].lower() for u in self.db.get_units_by_barcode(barcode)]
                    if unit_check in existed_units:
                        error_list.append(f"Dòng {index + 2}: Trùng mã vạch + đơn vị")
                        continue

                    # Thêm sản phẩm nếu chưa có
                    if not self.db.get_product_by_barcode(barcode):
                        self.db.add_product(barcode, name)
                    # Thêm đơn vị tính (lưu dạng capitalize)
                    if self.db.add_unit(barcode, unit_save, price):
                        success_count += 1
                    else:
                        error_list.append(f"Dòng {index + 2}: Trùng mã vạch + đơn vị")
                except Exception as e:
                    error_list.append(f"Dòng {index + 2}: {str(e)}")
            return True, f"Import thành công {success_count} đơn vị tính. " + \
                   (f"Lỗi ở {len(error_list)} dòng." if error_list else ""), error_list
        except Exception as e:
            return False, f"Lỗi đọc file: {str(e)}", []

    def export_to_excel(self, file_path):
        """Export sản phẩm ra Excel theo mẫu: Mã vạch, Tên hàng, Giá bán, ĐVT"""
        try:
            products = self.db.get_all_products()
            rows = []
            for prod in products:
                barcode = str(prod[1])  # Luôn là text
                name = prod[2]
                units = self.db.get_units_by_barcode(barcode)
                for unit_row in units:
                    unit = unit_row[2]
                    price = int(unit_row[3])  # Cắt phần thập phân
                    rows.append({
                        'Mã vạch': barcode,
                        'Tên hàng': name,
                        'ĐVT': unit,
                        'Giá bán': price
                    })
            df = pd.DataFrame(rows, columns=['Mã vạch', 'Tên hàng', 'Giá bán', 'ĐVT'])
            df.to_excel(file_path, index=False)
            return True, f"Export thành công {len(rows)} đơn vị tính"
        except Exception as e:
            return False, f"Lỗi export: {str(e)}"
    
    def get_template_data(self):
        """Tạo template Excel mẫu đúng yêu cầu"""
        template_data = {
            'Mã vạch': ['8936036018622', '8934673123456'],
            'Tên hàng': ['Nước ngọt Coca Cola', 'Bánh mì sandwich'],
            'Giá bán': [15000, 25000],
            'ĐVT': ['chai', 'cái']
        }
        return pd.DataFrame(template_data, columns=['Mã vạch', 'Tên hàng', 'Giá bán', 'ĐVT'])