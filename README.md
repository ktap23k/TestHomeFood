# TestHomeFood
Hướng dẫn dùng:
  1. Cài đặt python
  2. Cài đặt thư viện selenium, openpyxl. Code: "pip install selenium openpyxl"
  3. Cài bản Chrome phù hợp với bản chromediver (nếu dùng win đôi khi bắt add path, gg : add path one folder on windown)
  4. Tải về, giải nén, xong cd tới thư mục TestHomeFood
  5. run: "python Run.py" để test
      TestCaseGioHang()
      TestCaseTimKiem()
      TestCaseDatHang()
      là 3 testcase lớn, với đầu ra của TestCaseGioHang() là file outputGioHang.xlsx, TestCaseTimKiem() là file input có đầu vào, 
        TestCaseDatHang() là file outputThanhToan.xlsx
File main.py chứa code các testcase con của các testcase lớn
