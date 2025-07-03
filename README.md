# RPA Tải Hóa Đơn FPT

## Hướng dẫn chạy

1. Cài thư viện:
```
pip install selenium pandas openpyxl
```

2. Tải ChromeDriver cùng version với Chrome:
https://chromedriver.chromium.org/downloads

3. Đặt `chromedriver.exe` cùng thư mục với `main_fpt.py`

4. Chạy script:
```
py main_fpt.py
```

## Đầu vào: `input.xlsx`
Gồm 4 cột đúng theo tài liệu FPT:
- STT
- Mã tra cứu
- Mã số thuế
- URL tra cứu

## Đầu ra:
- Thư mục `Invoices`: chứa các file hóa đơn XML tải về
- File `output.xlsx`: chứa thông tin tổng hợp từ các hóa đơn