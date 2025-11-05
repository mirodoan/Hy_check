# KiotCheck - Quản lý sản phẩm đơn giản

## Mô tả

KiotCheck là ứng dụng quản lý sản phẩm dành cho cửa hàng nhỏ, hỗ trợ:

- Quản lý sản phẩm (CRUD)
- Quét mã vạch tìm kiếm
- Import/Export CSV
- Giao diện terminal đơn giản

## Cài đặt và chạy

### Yêu cầu hệ thống

- Python 3.6+
- Không cần thư viện bên ngoài (chỉ dùng built-in libraries)

### Cách chạy

```bash
python simple_app.py
```

## Cách sử dụng

### 1. Quản lý sản phẩm

- Xem danh sách tất cả sản phẩm
- Thêm sản phẩm mới
- Tìm kiếm theo tên hoặc mã vạch
- Sửa thông tin sản phẩm
- Xóa sản phẩm

### 2. Quét mã vạch

- Nhập mã vạch để tìm kiếm nhanh
- Hiển thị thông tin sản phẩm (tên, đơn vị, giá)

### 3. Import/Export CSV

- Import sản phẩm từ file CSV
- Export danh sách sản phẩm ra CSV
- Tạo template CSV mẫu

## Cấu trúc file CSV

```csv
barcode,name,unit,price
8936036018622,Nước ngọt Coca Cola,chai,15000
8934673123456,Bánh mì sandwich,cái,25000
```

## Files trong dự án

- `simple_app.py` - Ứng dụng chính (terminal interface)
- `database.py` - Quản lý database SQLite
- `main.py` - GUI version (cần tkinter)
- `sample_products.csv` - Dữ liệu mẫu
- `kiot_check.db` - Database SQLite (tự động tạo)

## Tính năng nổi bật

- ✅ Hoạt động offline hoàn toàn
- ✅ Không cần internet
- ✅ Giao diện đơn giản, dễ sử dụng
- ✅ Hỗ trợ máy quét mã vạch USB
- ✅ Import/Export dữ liệu dễ dàng
- ✅ Kiểm tra trùng lặp mã vạch + đơn vị

## Lưu ý

- Database được lưu trong file `kiot_check.db`
- Mỗi sản phẩm có thể có nhiều đơn vị tính khác nhau
- Không cho phép trùng mã vạch + đơn vị tính
