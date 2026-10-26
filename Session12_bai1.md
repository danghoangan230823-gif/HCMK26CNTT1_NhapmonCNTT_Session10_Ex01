# Bài 1 : Chuẩn hóa quy trình đối soát giao dịch ví điện tử
## Phần 1. Phân tích:
### 1.Field (Cột / Trường dữ liệu):

Là các thuộc tính đơn lẻ mô tả thông tin chi tiết của một giao dịch, bao gồm:

`MaGiaoDich` (Mã định danh giao dịch độc lập, ví dụ: TXN001)

`TenNguoiDung` (Họ và tên khách hàng thực hiện giao dịch)

`SoTien` (Giá trị giao dịch tính bằng VNĐ)

`TrangThai` (Trạng thái giao dịch: Thành công, Thất bại, Đang chờ)

`NgayGiaoDich` (Thời gian thực hiện giao dịch)

### 2.Record (Hàng / Bản ghi):

Là một tập hợp các giá trị của các Field liên kết với nhau, đại diện hoàn chỉnh cho một giao dịch tài chính độc lập thành công hoặc thất bại trên hệ thống PayKX tại một thời điểm cụ thể.
## **Phần 2. Triển khai:** Bộ 3 file Word, Excel, PowerPoint.
### Tài liệu word:
Cấu trúc phân cấp chuẩn Heading 1, Heading 2 đầy đủ 5 phần (Giới thiệu, Tổng quan, Yêu cầu chức năng hoàn tiền & đối soát, Phi chức năng, Phụ lục)
### Bảng tính Excel:
Sheet DoiSoat_GiaoDich tích hợp hàm chuẩn hóa =TRIM(B2) để loại bỏ triệt để khoảng trắng thừa trước khi đối soát.

Tuyệt đối không sử dụng Merge Cell trong vùng dữ liệu và không để trống hàng giữa các bản ghi.

Sheet BangTraCuu hỗ trợ tra cứu mã giao dịch.
### Bản trình bày PowerPoint:
Sử dụng Assertion Title (Tiêu đề khẳng định kết quả, ví dụ: "Doanh thu tăng 15% nhờ ưu đãi nạp điện thoại").

Cỡ chữ nội dung tối thiểu 30pt, thiết kế tối giản, trực quan, chuyên nghiệp.
## Phần 3.Kiểm thử: Giả định bảng tra cứu bị lỗi #N/A
### 1.Kiểm tra khoảng trắng ẩn (Whitespace & Hidden Characters):

Nguyên nhân: Mã giao dịch ở bảng chính và bảng tra cứu có chứa khoảng trắng thừa ở đầu, cuối hoặc giữa (ví dụ: "TXN001 " khác với "TXN001").

Cách khắc phục: Áp dụng hàm =TRIM() cho cả hai cột mã giao dịch trước khi thực hiện dò tìm.

### 2.Kiểm tra định dạng dữ liệu (Data Type Mismatch):

Nguyên nhân: Một bên mã giao dịch được lưu dưới dạng chuỗi (Text) và bên còn lại lưu dưới dạng số (Number). Excel sẽ không khớp được giá trị dù ký tự giống nhau.

Cách khắc phục: Đồng bộ định dạng toàn bộ cột mã giao dịch về dạng Text (chọn định dạng General/Text hoặc dùng hàm TEXT()).

### 3.Kiểm tra độ chính xác của tham số Range_lookup (Exact Match):

Nguyên nhân: Thiếu hoặc truyền sai tham số cuối cùng của hàm VLOOKUP (FALSE hoặc 0 để tìm kiếm chính xác tuyệt đối). Nếu để trống hoặc đặt TRUE, hàm yêu cầu dữ liệu phải được sắp xếp tăng dần và dễ sinh lỗi #N/A.

Cách khắc phục: Đảm bảo cú pháp luôn là =VLOOKUP(Lookup_value, Table_array, Col_index_num, FALSE).

### 4.Kiểm tra giới hạn vùng dò tìm (Table_array):

Nguyên nhân: Cột chứa giá trị cần tìm nằm ở bên trái cột đầu tiên của vùng Table_array (hàm VLOOKUP chỉ tìm kiếm từ trái sang phải).

Cách khắc phục: Mở rộng hoặc sắp xếp lại thứ tự cột trong bảng tra cứu sao cho cột chứa mã giao dịch nằm ở vị trí đầu tiên bên trái (Cột số 1).
