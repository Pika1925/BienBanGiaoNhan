# Biên bản giám định 190–191 (Windows)

Ứng dụng nhập liệu Windows, xuất Word theo hai mẫu đính kèm, mật khẩu mở ứng dụng: `024256`.

## Tạo EXE bằng GitHub Actions

1. Tạo repository **Private** trên GitHub. Không tải hồ sơ có dữ liệu cá nhân hoặc bí mật vụ án lên GitHub.
2. Tải và giải nén ZIP này. Đưa **toàn bộ nội dung thư mục dự án**, kể cả `.github/workflows/build-windows.yml`, lên **thư mục gốc** repository. Có thể dùng **Add file → Upload files**, hoặc Git.
3. Commit vào nhánh `main`.
4. Vào tab **Actions** → **Build Windows EXE** → **Run workflow** → **Run workflow** (hoặc đợi build tự chạy sau push).
5. Khi dấu kiểm màu xanh xuất hiện, mở lần chạy đó, kéo xuống **Artifacts**, tải `BienBanGiamDinh-Windows-x64.zip`.
6. Giải nén ZIP artifact, chạy `BienBanGiamDinh.exe` và nhập mật khẩu `024256`. Không cần cài Python trên máy sử dụng.

Lưu ý: artifact GitHub luôn được tải dưới dạng ZIP; bên trong có **một EXE** và file kiểm tra SHA256 (không cần thiết khi chạy).

## Kiểm tra tại máy Windows

Chạy `python -m unittest discover -s tests -v` để kiểm tra chức năng xuất Word. Có thể chạy `BUILD_WINDOWS.bat` để đóng gói tại máy. GitHub Actions chỉ kiểm tra EXE mở được 12 giây, **chưa thể xác nhận tương tác giao diện hay bản in giống 100%**. Kiểm tra hai file Word xuất ra trước khi sử dụng chính thức.

## Vấn đề bảo mật

Mật khẩu `024256` được nhúng trong ứng dụng nhằm hạn chế truy cập thông thường, **không phải cơ chế mã hóa dữ liệu**. Hồ sơ JSON và DOCX xuất ra không được mã hóa. Không lưu dữ liệu nhạy cảm lên repository.

## Chẩn đoán

Nếu chạy EXE không hiện cửa sổ, xem `BienBanGiamDinh_LOI.txt` trên Desktop, kiểm tra Windows Security và chạy bản debug bằng `BUILD_DEBUG_WINDOWS.bat`.
