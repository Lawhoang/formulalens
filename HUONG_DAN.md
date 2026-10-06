# FormulaLens – Mô tả tính năng và hướng dẫn sử dụng

FormulaLens là add-in Excel (Office.js) giúp kiểm tra, lần theo và rà soát công thức trong mô hình tài chính hoặc bảng tính phức tạp. Ý tưởng chức năng lấy cảm hứng từ các add-in kiểm toán công thức như Arixcel, nhưng toàn bộ mã nguồn được viết riêng.

> Lưu ý: phiên bản này chưa được kiểm thử trực tiếp trong Excel. Hãy thử trên một bản sao của workbook trước khi dùng cho file quan trọng.

## 1. Cài đặt (Excel desktop trên Windows)

**Yêu cầu:** Excel hỗ trợ ExcelApi 1.12 trở lên (Microsoft 365 hoặc Excel 2021/2019 bản mới), cài sẵn Node.js.

1. Mở **PowerShell** trong thư mục chứa các file (ví dụ `cd D:\Arixcel_Fake`).
2. Cài chứng chỉ HTTPS cục bộ: `npx office-addin-dev-certs install`. Khi Windows hỏi cài chứng chỉ, chọn **Yes**. Kiểm tra bằng lệnh `dir $env:USERPROFILE\.office-addin-dev-certs`, phải thấy `localhost.crt` và `localhost.key`.
3. Chạy server (trên PowerShell không dùng được ký hiệu `~`, phải dùng `$env:USERPROFILE`):
   ```
   npx http-server . -S -C "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt" -K "$env:USERPROFILE\.office-addin-dev-certs\localhost.key" -p 3000 --cors -c-1
   ```
   Tùy chọn `-c-1` tắt cache của server để panel luôn nhận file mới sau khi bạn ghi đè. Giữ cửa sổ này mở trong lúc dùng add-in. Mở `https://localhost:3000/taskpane.html` trong trình duyệt để kiểm tra, trang phải hiện giao diện panel và không báo lỗi chứng chỉ.
4. Nạp add-in bằng thư mục chia sẻ (cách này không cần `package.json`):
   - Chuột phải thư mục chứa `manifest.xml` → *Properties → Sharing → Share*, rồi ghi lại đường dẫn mạng dạng `\\TEN-MAY\Arixcel_Fake`.
   - Trong Excel: *File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs*. Dán đường dẫn vào ô *Catalog Url*, bấm *Add catalog*, tích *Show in Menu*, bấm OK rồi **khởi động lại Excel**.
   - Vào *Insert → Add-ins → My Add-ins → tab SHARED FOLDER*, chọn FormulaLens và bấm *Add*.
5. Trên ribbon sẽ xuất hiện tab **FormulaLens**.

## 2. Tổng quan giao diện

Panel bên phải có 5 tab: **Explorer, Formula map, Compare, Flow, Utilities**. Dòng trạng thái ở dưới cùng cho biết kết quả hoặc lỗi. Tab FormulaLens trên ribbon có các nút nhanh: Explore precedents, Explore dependents, Map formulas, Clear colours, Open panel.

Phím tắt:

| Phím | Chức năng |
|---|---|
| **Ctrl+Q** | Explore **dependents** |
| **Ctrl+Shift+Q** | Explore **precedents** |
| Nhấn lại cùng phím trên cùng ô | Xóa kết quả và màu tô |
| **Esc** (khi con trỏ ở trong panel) | Xóa màu tô, ẩn panel, bàn phím quay về sheet |
| **Alt+R** (khi con trỏ ở trong panel) | Refresh |

**Phím tắt mở và đưa con trỏ vào cây Explorer** (mặc định, có thể tắt bằng ô *Focus panel on shortcut*). Nếu Excel không chuyển bàn phím sang panel, hãy bấm chuột một lần vào vùng trống của panel, các lần sau dùng phím như bình thường. Khi tắt ô này, phím tắt chỉ tô màu trên sheet và không giành bàn phím.

Ctrl+Q mặc định là Quick Analysis của Excel; nếu bị trùng, hãy đổi phím trong `shortcuts.json`. Sau khi đổi manifest hoặc `shortcuts.json`, tắt hẳn Excel, xóa cache `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` và mở lại.

## 3. Các tính năng

### 3.1 Explorer – xem logic công thức
Chọn một ô rồi bấm **Precedents** (hoặc Ctrl+Shift+Q).

- Hiển thị địa chỉ, giá trị và công thức của ô.
- **Formula logic**: cây phân tích công thức thành hàm, đối số và toán hạng, mỗi nhánh kèm giá trị.
- **Cây precedents/dependents** (giống cửa sổ Explorer của Arixcel): gốc là ô đang chọn, các nhánh là ô hoặc vùng liên quan, mỗi dòng kèm giá trị. Bấm mũi tên ▸ để mở thêm nhiều tầng.
- **Công thức có liên kết**: các tham chiếu trong công thức hiện màu xanh, gạch chân. Bấm vào để chọn ô đó ngay trên sheet.
- **Bấm một dòng** trong cây hoặc một tham chiếu: Excel chọn ô đó trên sheet, dòng được tô sáng và ô công thức hiển thị công thức của ô vừa chọn. **Bấm đúp** để đặt ô đó làm gốc mới và phân tích tiếp. Bấm **Back** để quay lại gốc trước.
- Tích **Follow selection** để panel tự cập nhật mỗi khi bạn bấm sang ô khác trên sheet.

**Điều hướng bằng bàn phím trong cây** (khi con trỏ đang ở trong panel):

| Phím | Chức năng |
|---|---|
| ↑ / ↓ | Di chuyển giữa các dòng; Excel tự chọn ô tương ứng trên sheet |
| → | Mở nhánh (nếu đã mở thì xuống dòng con đầu tiên) |
| ← | Thu nhánh (nếu đã thu thì lên dòng cha) |
| **Phím cách** | Mở / thu nhánh |
| **Enter** | Đặt dòng đang chọn làm gốc mới (như bấm đúp) |
| Backspace | Quay lại gốc trước |
| Home / End | Về dòng đầu / dòng cuối |
| Esc | Xóa màu tô, ẩn panel, bàn phím về sheet (con trỏ Excel dừng ở ô cuối cùng bạn duyệt tới) |

**Màu tô (giống Arixcel):** ô gốc **hồng đậm**, precedents **xanh dương đậm**, dependents **xanh lá**; ô đang được duyệt trong cây chuyển sang màu **sáng hơn** (xanh dương nhạt hoặc xanh lá nhạt) rồi trở lại màu cũ khi bạn duyệt sang dòng khác. Add-in chưa vẽ viền quanh ô như Arixcel.

**Dependents** (hoặc Ctrl+Q) liệt kê những ô phụ thuộc vào ô đang chọn. Chọn nhiều ô cùng lúc để lần theo precedents hoặc dependents của cả nhóm.

**Evaluate sub-expressions (mới):** khi bật ô đánh dấu *Evaluate sub-expressions* (mặc định bật), mỗi hàm và biểu thức con trong cây đều hiển thị giá trị tính ra. Ví dụ với `=IF(A1>0, SUM(B1:B3), 0)` bạn sẽ thấy kết quả của `A1>0` và của `SUM(B1:B3)`.

Cách hoạt động: Office.js không có hàm Evaluate, nên add-in ghi tạm từng biểu thức vào các ô ở góc dưới bên phải của sheet (ô XFD1048576 và các ô phía trên), đọc kết quả rồi xóa ngay. Hệ quả:

- Nếu ô XFD1048576 đang có dữ liệu, hoặc sheet bị khóa, tính năng tự bỏ qua.
- Chỉ đánh giá tối đa 60 biểu thức mỗi lần.
- Nếu không muốn add-in ghi tạm vào sheet, bỏ chọn ô này.

### 3.2 Formula map – tô màu phát hiện rủi ro
Vào tab **Formula map**, bấm **Map active sheet** (hoặc nút Map formulas trên ribbon).

| Màu | Ý nghĩa |
|---|---|
| Tím | Công thức chứa liên kết ngoài (file khác) |
| Đỏ | Công thức khác với các ô lân cận trong khi hai ô hai bên lại giống nhau (nghi sai) |
| Vàng | Công thức chứa số gõ cứng (trừ 0 và 1) |
| Xanh lá | Công thức nhất quán |
| Xanh dương nhạt | Ô chứa số gõ cứng |

Chú giải kèm số lượng từng loại. Bấm **Clear map** (hoặc Clear colours trên ribbon) để khôi phục màu nền ban đầu. Giới hạn 10.000 ô trong vùng sử dụng.

### 3.3 Compare – so sánh hai sheet
Vào tab **Compare**, chọn Sheet A (bản gốc) và Sheet B (bản sửa), bấm **Compare**.

- Liệt kê từng ô khác nhau theo dạng `giá trị/công thức A → B`, tối đa 500 dòng hiển thị.
- Bấm một dòng để nhảy tới ô đó trong Sheet B.
- **Highlight in B** tô cam các ô khác biệt (tối đa 2.000 ô), **Clear highlight** để bỏ.
- Chỉ so sánh trong cùng một workbook, giới hạn 100.000 ô. Chưa tự căn hàng/cột khi có chèn thêm hàng hoặc cột.

### 3.4 Flow – Calculation flow (mới)
Cho biết dữ liệu chảy vào và ra khỏi một vùng tính toán.

1. Chọn vùng tính toán (tối đa 2.000 ô).
2. Vào tab **Flow**, bấm **Analyse selection**.

| Màu | Ý nghĩa |
|---|---|
| Xanh dương nhạt | Input trong vùng (số gõ cứng) |
| Vàng nhạt | Calculation trong vùng (công thức) |
| Tím | Output: ô trong vùng có ô phụ thuộc nằm ngoài vùng |
| Xanh lá | Nguồn bên ngoài vùng mà vùng đang dùng |
| Hồng | Ô bên ngoài vùng đang dùng kết quả của vùng |

Hai danh sách **Flows in from** và **Flows out to** liệt kê các địa chỉ bên ngoài, bấm vào để nhảy tới. Tính năng giúp trả lời nhanh: vùng này lấy số từ đâu, và kết quả đang đi tới đâu. Phần kiểm tra output chỉ xét 300 ô không trống đầu tiên. Bấm **Clear colours** để khôi phục.

### 3.5 Utilities – tiện ích
- Show all hidden sheets: hiện toàn bộ sheet ẩn.
- Toggle gridlines: bật/tắt đường lưới sheet hiện tại.
- Convert selected formulas to values: đổi công thức thành giá trị (có hỏi xác nhận, **không hoàn tác được**).
- Autofit columns: tự chỉnh độ rộng cột.
- Toggle automatic / manual calculation: đổi chế độ tính toán.

## 4. Lưu ý quan trọng

- **Undo bị xóa:** việc tô màu và ghi ô tạm làm Excel xóa lịch sử Ctrl+Z. Hãy lưu file trước khi dùng.
- **Màu nền:** add-in khôi phục màu nền cũ khi xóa. Ô vốn được tô trắng tường minh sẽ trở về không tô màu. Nếu đóng Excel mà chưa xóa màu, màu tạm sẽ còn lại trong file, vì vậy hãy bấm Clear trước khi lưu hoặc đóng.
- **Phạm vi:** chỉ làm việc trong workbook đang mở; không thay thế các công cụ kiểm toán chuyên nghiệp cho mô hình quan trọng.

## 5. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Lỗi `Could not find certificate ~/...` | Dùng `$env:USERPROFILE` thay cho `~` như bước 3 |
| Lỗi `package.json` khi chạy `office-addin-debugging` | Không dùng lệnh đó, nạp bằng thư mục chia sẻ như bước 4 |
| Không thấy tab FormulaLens | Kiểm tra server đang chạy ở `https://localhost:3000`, chứng chỉ đã cài, rồi mở lại Excel |
| Panel vẫn chạy bản cũ sau khi ghi đè file | Chạy server với `-c-1`, tắt hẳn Excel, xóa cache Wef rồi mở lại. Dòng trạng thái dưới cùng phải hiện `build 05.10-b` |
| Panel trắng | Mở `https://localhost:3000/taskpane.html` trong trình duyệt để kiểm tra chứng chỉ |
| Ctrl+Q không hoạt động | Phím có thể trùng với Excel: đổi phím trong `shortcuts.json`, xóa cache Wef, mở lại Excel. Vẫn dùng được nút trên ribbon |
| Báo lỗi API | Cần Excel hỗ trợ ExcelApi 1.12 (getDirectPrecedents/Dependents) |
| Giá trị biểu thức con trống | Ô XFD1048576 đang có dữ liệu, sheet bị khóa, hoặc biểu thức không hợp lệ |

## 6. Cấu trúc thư mục

- `manifest.xml` – khai báo add-in, ribbon, runtime dùng chung
- `taskpane.html` – giao diện và toàn bộ logic
- `shortcuts.json` – phím tắt (được `manifest.xml` tham chiếu)
- `icon-16/32/80.png` – biểu tượng
- `HUONG_DAN.md` – tài liệu này
