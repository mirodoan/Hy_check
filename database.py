import sqlite3
import os
from pathlib import Path

class Database:
    def __init__(self, db_path="kiot_check.db"):
        self.db_path = db_path
        self.init_database()
    
    def get_connection(self):
        return sqlite3.connect(self.db_path)
    
    def init_database(self):
        """Tạo bảng products và units nếu chưa có"""
        conn = self.get_connection()
        cursor = conn.cursor()
        # Bảng sản phẩm
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                barcode TEXT NOT NULL UNIQUE,
                name TEXT NOT NULL
            )
        ''')
        # Bảng đơn vị tính
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS units (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                barcode TEXT NOT NULL,
                unit TEXT NOT NULL,
                price REAL NOT NULL,
                UNIQUE(barcode, unit),
                FOREIGN KEY(barcode) REFERENCES products(barcode) ON DELETE CASCADE
            )
        ''')
        conn.commit()
        conn.close()
    
    def add_product(self, barcode, name):
        """Thêm sản phẩm mới (chỉ thông tin chung)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO products (barcode, name)
                VALUES (?, ?)
            ''', (barcode, name))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False  # Trùng barcode
        finally:
            conn.close()

    def add_unit(self, barcode, unit, price):
        """Thêm đơn vị tính cho sản phẩm"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO units (barcode, unit, price)
                VALUES (?, ?, ?)
            ''', (barcode, unit, price))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False  # Trùng barcode + unit
        finally:
            conn.close()
    
    def get_all_products(self):
        """Lấy tất cả sản phẩm (không đơn vị tính)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM products ORDER BY name')
        products = cursor.fetchall()
        conn.close()
        return products

    def get_units_by_barcode(self, barcode):
        """Lấy tất cả đơn vị tính và giá của barcode"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM units WHERE barcode = ? ORDER BY unit', (barcode,))
        units = cursor.fetchall()
        conn.close()
        return units
    
    def search_products(self, search_term):
        """Tìm kiếm sản phẩm theo tên hoặc barcode"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT * FROM products 
            WHERE name LIKE ? OR barcode LIKE ?
            ORDER BY name
        ''', (f'%{search_term}%', f'%{search_term}%'))
        products = cursor.fetchall()
        conn.close()
        return products
    
    def get_product_by_barcode(self, barcode):
        """Lấy sản phẩm theo barcode"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM products WHERE barcode = ?', (barcode,))
        product = cursor.fetchone()
        conn.close()
        return product
    
    def update_product(self, barcode, name):
        """Cập nhật thông tin sản phẩm"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                UPDATE products SET name=? WHERE barcode=?
            ''', (name, barcode))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()

    def update_unit(self, unit_id, unit, price):
        """Cập nhật đơn vị tính"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                UPDATE units SET unit=?, price=? WHERE id=?
            ''', (unit, price, unit_id))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            conn.close()
    
    def delete_product(self, barcode):
        """Xóa sản phẩm và các đơn vị tính liên quan"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM products WHERE barcode=?', (barcode,))
        conn.commit()
        conn.close()

    def delete_unit(self, unit_id):
        """Xóa đơn vị tính"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM units WHERE id=?', (unit_id,))
        conn.commit()
        conn.close()