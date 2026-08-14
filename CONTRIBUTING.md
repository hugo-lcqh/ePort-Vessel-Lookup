# Đóng góp cho ePort Vessel Lookup

Cảm ơn bạn đã quan tâm đến dự án. Hãy giữ thay đổi nhỏ, tập trung và có thể kiểm chứng.

## Báo lỗi

Trước khi tạo issue, hãy kiểm tra issue hiện có và thử lại bằng phiên bản mới nhất trên nhánh `main`.

Một báo cáo lỗi hữu ích nên có:

- Phiên bản Windows và Python.
- Trình duyệt được chọn cùng phiên bản trình duyệt.
- Các bước tái hiện lỗi.
- Kết quả mong đợi và kết quả thực tế.
- Thông báo lỗi đầy đủ hoặc ảnh chụp màn hình đã che dữ liệu nhạy cảm.
- File chẩn đoán nếu phù hợp, sau khi kiểm tra và loại bỏ thông tin không nên chia sẻ.

Không đăng mật khẩu, token, cookie, dữ liệu nội bộ hoặc thông tin cá nhân lên issue.

## Thiết lập môi trường phát triển

```powershell
git clone https://github.com/hugo-lcqh/ePort-Vessel-Lookup.git
cd ePort-Vessel-Lookup
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements-dev.txt
```

## Quy trình thay đổi

1. Tạo nhánh ngắn từ `main`, ví dụ `fix/result-parser` hoặc `docs/setup-guide`.
2. Chỉ thay đổi một vấn đề trong mỗi commit.
3. Dùng commit message theo dạng `fix: ...`, `feat: ...`, `docs: ...` hoặc `chore: ...`.
4. Kiểm tra thay đổi trước khi mở pull request.

## Kiểm tra tối thiểu

```powershell
python -m compileall -q eport_vessel_lookup.py
python -m pip check
python -m PyInstaller --clean --noconfirm ePortVesselLookup.spec
```

Với thay đổi liên quan đến tra cứu, hãy chạy thủ công cả chế độ ẩn và hiển thị trình duyệt. Không commit nội dung trong `build/`, `dist/`, báo cáo Excel hoặc file HTML chẩn đoán.

## Pull request checklist

- [ ] Thay đổi giải quyết một mục tiêu rõ ràng.
- [ ] Không chứa secret, dữ liệu cá nhân hoặc build artifact.
- [ ] Cú pháp Python và dependency check đều thành công.
- [ ] Build PyInstaller thành công nếu thay đổi ảnh hưởng đóng gói.
- [ ] README và CHANGELOG được cập nhật khi hành vi người dùng thay đổi.
- [ ] Đã mô tả cách kiểm chứng trong pull request.
