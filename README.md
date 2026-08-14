# ePort Vessel Lookup

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-All%20rights%20reserved-lightgrey)

Công cụ dòng lệnh giúp tra cứu thông tin tàu/chuyến tại khu vực **Cát Lái** trên ePort Saigon Newport. Ứng dụng tự động điều khiển Microsoft Edge hoặc Google Chrome, hỗ trợ tìm nhiều tàu song song, hiển thị bảng tổng hợp và xuất kết quả ra Excel.

> **Tuyên bố miễn trừ:** Đây là dự án độc lập, không phải sản phẩm chính thức và không được Saigon Newport Corporation bảo trợ. Người dùng có trách nhiệm tuân thủ điều khoản sử dụng của website nguồn và các quy định liên quan.

## Tính năng

- Tự động chọn khu vực Cát Lái và tra cứu theo tên tàu.
- Hỗ trợ Edge, Chrome hoặc chế độ tự động ưu tiên Edge.
- Tra cứu tối đa ba tàu song song để rút ngắn thời gian chờ.
- Hiển thị Terminal, Agent, Vessel, Voyage, Berth/ATB và Departure trên bảng tổng hợp.
- Cho phép xem toàn bộ trường dữ liệu của từng chuyến tàu.
- Xuất báo cáo `.xlsx` có định dạng, bộ lọc và cố định hàng tiêu đề.
- Có chế độ hiển thị trình duyệt và chế độ chẩn đoán khi website thay đổi.
- Có thể đóng gói thành một file `.exe` bằng PyInstaller.

## Yêu cầu

- Windows 10 hoặc Windows 11.
- Python 3.10 trở lên.
- Microsoft Edge hoặc Google Chrome đã được cài đặt.
- Kết nối Internet tới [ePort Saigon Newport](https://eport.saigonnewport.com.vn/Ships).

`webdriver-manager` sẽ tự tải WebDriver tương thích trong lần chạy đầu tiên. Môi trường có proxy hoặc firewall doanh nghiệp có thể cần cho phép kết nối này.

## Bắt đầu nhanh

### 1. Tải mã nguồn

```powershell
git clone https://github.com/hugo-lcqh/ePort-Vessel-Lookup.git
cd ePort-Vessel-Lookup
```

### 2. Tạo môi trường Python riêng

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Nếu PowerShell không cho phép kích hoạt môi trường ảo, có thể gọi Python trực tiếp bằng `.\.venv\Scripts\python.exe`.

### 3. Chạy ứng dụng

```powershell
python .\eport_vessel_lookup.py
```

## Cách sử dụng

1. Chọn trình duyệt: tự động, Edge hoặc Chrome.
2. Chọn chạy ẩn, hiển thị trình duyệt hoặc chẩn đoán.
3. Nhập một tên tàu hoặc nhiều tên tàu, phân tách bằng dấu phẩy.
4. Xem bảng tổng hợp và chọn xem chi tiết nếu cần.
5. Nhấn `Enter` ở câu hỏi xuất Excel để lưu báo cáo.

Ví dụ đầu vào:

```text
EVER GIVEN, MSC OSCAR, EMMA MAERSK
```

Ứng dụng chỉ tra cứu theo tên tàu và khu vực được cố định là Cát Lái.

## Dữ liệu đầu ra

- Báo cáo Excel: `eport_ket_qua_YYYYMMDD_HHMMSS.xlsx`.
- File chẩn đoán khi không tìm thấy dữ liệu: `debug_<TEN_TAU>.html`.

Hai loại file này được Git bỏ qua. File chẩn đoán là bản HTML của trang tại thời điểm lỗi; hãy kiểm tra nội dung trước khi chia sẻ cho người khác.

## Đóng gói file EXE

Cách đơn giản nhất trên Windows:

```powershell
.\build_exe.bat
```

Hoặc chạy thủ công:

```powershell
python -m pip install -r requirements-dev.txt
python -m PyInstaller --clean --noconfirm ePortVesselLookup.spec
```

File hoàn tất nằm tại `dist\ePortVesselLookup.exe`. Build artifact không được commit vào repository; nếu cần phân phối binary, nên đăng dưới mục GitHub Releases. File `.exe` hiện chưa được ký số nên Windows SmartScreen hoặc phần mềm bảo mật có thể hiển thị cảnh báo.

## Cấu trúc dự án

```text
ePort-Vessel-Lookup/
├── eport_vessel_lookup.py   # Ứng dụng CLI và logic tra cứu
├── ePortVesselLookup.spec   # Cấu hình đóng gói PyInstaller
├── build_exe.bat            # Build tự động trên Windows
├── requirements.txt         # Dependency khi chạy
└── requirements-dev.txt     # Dependency dùng để build
```

Luồng xử lý chính: nhận tên tàu → mở trình duyệt bằng Selenium → thao tác DevExtreme trên website ePort → đọc bảng kết quả → hiển thị trên terminal hoặc xuất Excel.

## Xử lý sự cố

### Không mở được Edge hoặc Chrome

- Kiểm tra trình duyệt đã được cài đặt và có thể mở bình thường.
- Kiểm tra Internet, proxy và firewall có cho phép tải WebDriver.
- Thử chọn trực tiếp Edge hoặc Chrome thay vì chế độ tự động.

### Không tìm thấy dữ liệu dù website có kết quả

- Chạy lại với chế độ **Hiện browser** để quan sát thao tác.
- Dùng chế độ **Chẩn đoán** nếu giao diện website vừa thay đổi.
- Kiểm tra file `debug_<TEN_TAU>.html` được tạo sau lần tra cứu lỗi.

### Không lưu được Excel

Đóng file Excel cùng tên nếu file đang mở, rồi xuất lại. Theo mặc định, mỗi báo cáo có timestamp riêng để tránh ghi đè.

## Giới hạn và sử dụng có trách nhiệm

- Công cụ phụ thuộc vào cấu trúc HTML/DevExtreme của website nguồn và có thể cần cập nhật khi giao diện thay đổi.
- Kết quả phản ánh dữ liệu website cung cấp tại thời điểm tra cứu; hãy xác minh lại trước khi dùng cho quyết định vận hành quan trọng.
- Không tăng số luồng hoặc thực hiện truy vấn hàng loạt quá mức gây tải không cần thiết cho dịch vụ nguồn.
- Dự án không lưu thông tin đăng nhập và không yêu cầu API key.

## Đóng góp

Xem [CONTRIBUTING.md](CONTRIBUTING.md) để thiết lập môi trường phát triển, báo lỗi và gửi thay đổi. Các thay đổi đáng chú ý được ghi tại [CHANGELOG.md](CHANGELOG.md).

## Giấy phép

Bản quyền thuộc về `hugo-lcqh`. Dự án hiện không cấp giấy phép mã nguồn mở; xem [LICENSE](LICENSE) để biết chi tiết.
