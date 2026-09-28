# BWPConvertTTNVN - Tiện ích Chuyển đổi Số thành Chữ Tiếng Việt cho Microsoft Excel

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Excel%202013--365-green.svg)]()
[![Build Pipeline](https://img.shields.io/badge/Build-Automated%20COM%20Pipeline-success.svg)]()

**BWPConvertTTNVN** là add-in Microsoft Excel chuyên nghiệp, hiệu năng cao dùng để đọc số, chuyển đổi số tiền và số thập phân thành chữ tiếng Việt chuẩn xác theo ngữ pháp tiếng Việt và chuẩn tài chính kế toán Việt Nam.

Toàn bộ mã nguồn VBA tuân thủ tiêu chuẩn 7-bit ASCII an toàn tuyệt đối và sử dụng cơ chế dựng ký tự Unicode UTF-16 (`ChrW`) lúc chạy, loại bỏ hoàn toàn lỗi hiển thị tiếng Việt (font chữ, bảng mã) trên mọi phiên bản Windows và gói ngôn ngữ Office.

---

## Mục lục

1. [Tính năng nổi bật](#tính-năng-nổi-bật)
2. [Yêu cầu hệ thống](#yêu-cầu-hệ-thống)
3. [Cài đặt](#cài-đặt)
   - [Cách 1: Cài đặt tự động bằng Install.vbs (Khuyên dùng)](#cách-1-cài-đặt-tự-động-bằng-installvbs-khuyên-dùng)
   - [Cách 2: Cài đặt thủ công trong Excel](#cách-2-cài-đặt-thủ-công-trong-excel)
4. [Hướng dẫn sử dụng giao diện Ribbon](#hướng-dẫn-sử-dụng-giao-diện-ribbon)
   - [Chuyển đổi nhanh (Quick Convert)](#chuyển-đổi-nhanh-quick-convert)
   - [Chuyển đổi hàng loạt theo vùng dữ liệu (Batch Conversion)](#chuyển-đổi-hàng-loạt-theo-vùng-dữ-liệu-batch-conversion)
   - [Tính năng Hoàn tác (Undo)](#tính-năng-hoàn-tác-undo)
   - [Cấu hình tùy chọn (Preferences / Settings)](#cấu-hình-tùy-chọn-preferences--settings)
5. [Tra cứu Hàm công thức Excel (UDF Reference)](#tra-cứu-hàm-công-thức-excel-udf-reference)
   - [Hàm BWPVND](#hàm-bwpvnd)
   - [Hàm BWPVNDUPPER](#hàm-bwpvndupper)
   - [Hàm BWPVNWORDS](#hàm-bwpvnwords)
   - [Bảng tham số hàm BWPVNWORDS](#bảng-tham-số-hàm-bwpvnwords)
   - [Các bí danh tương thích (Aliases)](#các-bí-danh-tương-thích-aliases)
6. [Bảo mật & Cam kết hoạt động Offline](#bảo-mật--cam-kết-hoạt-động-offline)
7. [Hướng dẫn Biên dịch & Đóng gói (Build Instructions)](#hướng-dẫn-biên-dịch--đóng-gói-build-instructions)
8. [Bản quyền & Giấy phép (License)](#bản-quyền--giấy-phép-license)

---

## Tính năng nổi bật

- **Độ chính xác và dải giá trị cực lớn:**
  - Hỗ trợ số nguyên và số thập phân lên đến 15 chữ số (`999,999,999,999,999` - gần một triệu tỷ).
  - Tự động bắt lỗi tràn số (`#VALUE!`) nếu vượt quá giới hạn an toàn IEEE 754.
- **Ngữ pháp tiếng Việt chuẩn xác:**
  - Xử lý hoàn hảo các quy tắc đọc số: *mười một*, *mười lăm*, *hai mươi mốt*, *hai mươi tư*, *hai mươi lăm*, *không trăm lẻ một*, *linh/lẻ*.
  - Đọc chính xác các lớp số lớn (*nghìn*, *triệu*, *tỷ*, *nghìn tỷ*), xử lý thông minh các lớp có cụm 3 số không (*000*).
- **Hỗ trợ đa tiền tệ:**
  - Tiền tệ trong nước: **VND** (*đồng chẵn*, *đồng*).
  - Ngoại tệ thông dụng: **USD** (*đô la Mỹ*, *cent*), **EUR** (*euro*, *cent*), **JPY** (*yên Nhật*), **GBP** (*bảng Anh*, *pence*), **CNY** (*nhân dân tệ*, *hào*, *xu*).
  - Chế độ đọc số thuần túy (không kèm tên đơn vị tiền tệ).
- **Tùy biến phong cách và phương ngữ vùng miền:**
  - Kiểu đọc số không: *lẻ* (miền Bắc) hoặc *linh* (miền Nam).
  - Kiểu đọc hàng nghìn: *nghìn* (miền Bắc) hoặc *ngàn* (miền Nam).
  - Kiểu đọc số 4: *bốn* hoặc *tư*.
  - Định dạng kiểu chữ: Viết hoa chữ cái đầu (*Một trăm...*), Viết Hoa Từng Từ (*Một Trăm...*), chữ thường (*một trăm...*), VIẾT HOA TOÀN BỘ (*MỘT TRĂM...*).
- **Giao diện hiện đại & tiện ích:**
  - Thẻ Ribbon riêng biệt `BWPConvertTTNVN` tích hợp sẵn trong Microsoft Excel.
  - Hộp thoại Quick Convert trực quan, xem trước kết quả trực tiếp (live preview).
  - Hỗ trợ chuyển đổi hàng loạt (Batch Processing) cực nhanh: xử lý 10,000 ô trong dưới 1.5 giây thông qua mảng nhớ 2D variant.
  - Cơ chế Hoàn tác (Undo) giao dịch đơn cấp bảo toàn công thức và giá trị nguyên bản.
  - Cơ chế chống ghi đè vùng dữ liệu (Overlap Guard) bảo vệ an toàn bảng tính.

---

## Yêu cầu hệ thống

- **Hệ điều hành:** Windows 10, Windows 11 (32-bit hoặc 64-bit).
- **Microsoft Office:** Excel 2013, Excel 2016, Excel 2019, Excel 2021 hoặc Microsoft 365 (Office 365).
- Không yêu cầu cài đặt thêm bất kỳ thư viện hoặc phần mềm bên ngoài nào.

---

## Cài đặt

### Cách 1: Cài đặt tự động bằng Install.vbs (Khuyên dùng)

Bộ cài đặt tự động được đóng gói dưới dạng tập lệnh VBScript không cần quyền quản trị viên (non-elevated):

1. Tải bộ cài `BWPConvertTTNVN-v1.0.0.zip` và giải nén vào một thư mục trên máy tính.
2. Lưu lại các file Excel đang mở và **đóng Microsoft Excel**.
3. Nhấp đúp chuột vào file `Install.vbs`.
4. Trình cài đặt sẽ tự động:
   - Kiểm tra nếu Excel đang mở sẽ nhắc nhở người dùng lưu tài liệu và thoát an toàn.
   - Kiểm tra mã băm SHA-256 đối chiếu với `BWPConvertTTNVN-v1.0.0.sha256` để đảm bảo file nguyên vẹn, không bị can thiệp.
   - Gỡ bỏ cờ bảo mật Mark of the Web (`Zone.Identifier`) để add-in không bị khóa macro.
   - Sao chép add-in vào thư mục `%APPDATA%\Microsoft\AddIns\`.
   - Đăng ký và kích hoạt add-in tự động vào Excel.
5. Mở Excel, bạn sẽ thấy thẻ **BWPConvertTTNVN** xuất hiện trên thanh công cụ Ribbon.

### Cách 2: Cài đặt thủ công trong Excel

1. Sao chép file `BWPConvertTTNVN.xlam` vào thư mục:
   ```text
   %APPDATA%\Microsoft\AddIns\
   (Ví dụ: C:\Users\<Tên_User>\AppData\Roaming\Microsoft\AddIns\)
   ```
2. Nhấp chuột phải vào file `BWPConvertTTNVN.xlam` vừa sao chép, chọn **Properties**, nếu có mục **Security: This file came from another computer and might be blocked**, hãy tích vào ô **Unblock** rồi bấm **OK**.
3. Mở Microsoft Excel.
4. Chọn **File** -> **Options** -> **Add-ins**.
5. Ở mục **Manage** cuối cửa sổ, chọn **Excel Add-ins** rồi bấm **Go...**.
6. Trong danh sách Add-ins, bấm **Browse...** và chọn `BWPConvertTTNVN.xlam` (hoặc tích chọn `BWPConvertTTNVN` nếu đã có sẵn trong danh sách).
7. Bấm **OK** để hoàn tất kích hoạt.

---

## Hướng dẫn sử dụng giao diện Ribbon

Khi kích hoạt add-in thành công, thanh Ribbon sẽ hiển thị tab `BWPConvertTTNVN` với các nhóm chức năng:

### Chuyển đổi nhanh (Quick Convert)
1. Chọn ô hoặc cột chứa số tiền cần chuyển đổi.
2. Bấm nút **Quick Convert** trên tab Ribbon (hoặc mở hộp thoại Chuyển đổi).
3. Kết quả chuyển đổi sẽ được tạo ngay tại cột liền kề bên phải một cách tự động và an toàn.

### Chuyển đổi hàng loạt theo vùng dữ liệu (Batch Conversion)
1. Bấm nút **Hộp thoại Chuyển đổi** trên Ribbon để mở cửa sổ `frmConvert`.
2. Chọn **Vùng nguồn (Source Range)** chứa danh sách số tiền cần đọc.
3. Chọn **Ô đích (Destination Cell)** là ô đầu tiên sẽ đặt kết quả.
4. Chọn loại tiền tệ (VND, USD, EUR,...), kiểu chữ và các tùy chọn ngữ pháp.
5. Xem trước chuỗi kết quả mẫu ở khung xem trước (Preview).
6. Bấm **Chuyển đổi**. BWPConvertTTNVN sẽ thực hiện chuyển đổi tốc độ cao bằng mảng dữ liệu trong bộ nhớ RAM, tự động bỏ qua các ô trống hoặc lỗi mà không ảnh hưởng tới dữ liệu xung quanh.

> **Lưu ý an toàn:** BWPConvertTTNVN tích hợp cơ chế chống trùng lặp (Overlap Guard). Nếu vùng đích vô tình giao cắt hoặc đè lên vùng nguồn trên cùng một trang tính, hệ thống sẽ cảnh báo và từ chối ghi đè nhằm bảo vệ tuyệt đối dữ liệu gốc của bạn.

### Tính năng Hoàn tác (Undo)
Nếu sau khi thực hiện chuyển đổi bạn muốn quay lại trạng thái ban đầu:
1. Bấm ngay nút **Hoàn tác (Undo)** trên thanh Ribbon.
2. Hệ thống khôi phục 100% công thức, giá trị và định dạng ô như trước khi chuyển đổi.

### Cấu hình tùy chọn (Preferences / Settings)
Bấm nút **Cài đặt (Preferences)** để mở hộp thoại `frmSettings`. Bạn có thể tùy chỉnh:
- **Kiểu đọc số không:** Chọn *lẻ* hoặc *linh*.
- **Kiểu đọc hàng nghìn:** Chọn *nghìn* hoặc *ngàn*.
- **Kiểu đọc số bốn:** Chọn *bốn* hoặc *tư*.
- **Hậu tố đồng:** Bật/tắt chữ *chẵn* khi số tiền tròn nghìn (*đồng chẵn* / *đồng*).
- **Kiểu chữ mặc định:** Viết hoa chữ đầu, viết hoa từng từ, chữ thường, chữ hoa.
- Các cài đặt này được lưu trữ an toàn trong Windows Registry (`HKCU\Software\BWPConvertTTNVN`) và áp dụng nhất quán trên toàn bộ sổ làm việc.

---

## Tra cứu Hàm công thức Excel (UDF Reference)

BWPConvertTTNVN cung cấp các hàm người dùng tự định nghĩa (UDF) có thể gõ trực tiếp vào bất kỳ ô tính nào như các hàm tích hợp sẵn của Excel:

### Hàm BWPVND
Hàm chuyên dụng đọc số tiền Việt Nam Đồng (mặc định thêm hậu tố *đồng chẵn* và viết hoa chữ cái đầu):
```excel
=BWPVND(A1)
```
- **Ví dụ 1:** `=BWPVND(125430000)`
  - Kết quả: `Một trăm hai mươi lăm triệu bốn trăm ba mươi nghìn đồng chẵn.`
- **Ví dụ 2:** `=BWPVND(0)`
  - Kết quả: `Không đồng chẵn.`
- **Ví dụ 3:** `=BWPVND(-150000)`
  - Kết quả: `Âm một trăm năm mươi nghìn đồng chẵn.`

### Hàm BWPVNDUPPER
Hàm chuyên dụng đọc số tiền VND và xuất ra chữ viết hoa toàn bộ (chuẩn kế toán cho các mẫu séc, hóa đơn VAT):
```excel
=BWPVNDUPPER(A1)
```
- **Ví dụ:** `=BWPVNDUPPER(50000000)`
  - Kết quả: `NĂM MƯƠI TRIỆU ĐỒNG CHẴN.`

### Hàm BWPVNWORDS
Hàm chuyển đổi đa năng toàn diện cho phép tùy biến loại tiền tệ, cách hiển thị và định dạng:
```excel
=BWPVNWORDS(value, [currency], [style], [unit])
```

#### Bảng tham số hàm BWPVNWORDS

| Tham số | Kiểu dữ liệu | Bắt buộc? | Mặc định | Diễn giải |
|:---|:---|:---:|:---:|:---|
| `value` | Number / Range | Có | - | Số cần chuyển đổi (tối đa 15 chữ số). |
| `currency` | Text | Không | `"VND"` | Mã tiền tệ: `"VND"`, `"USD"`, `"EUR"`, `"JPY"`, `"GBP"`, `"CNY"`. Đặt chuỗi rỗng `""` để đọc số thuần túy (không kèm đơn vị tiền). |
| `style` | Integer | Không | `0` | Định dạng kiểu chữ:<br>`0`: Viết hoa chữ đầu (*Một trăm hai mươi...*)<br>`1`: Viết Hoa Từng Từ (*Một Trăm Hai Mươi...*)<br>`2`: Chữ thường (*một trăm hai mươi...*)<br>`3`: CHỮ IN HOA (*MỘT TRĂM HAI MƯƠI...*) |
| `unit` | Boolean | Không | `TRUE` | `TRUE`: Kèm theo tên đơn vị tiền tệ và hậu tố.<br>`FALSE`: Chỉ đọc số, không xuất đơn vị tiền tệ. |

#### Một số ví dụ nâng cao với BWPVNWORDS:
- **Đọc số ngoại tệ USD:**
  ```excel
  =BWPVNWORDS(1250.5, "USD")
  ```
  -> `Một nghìn hai trăm năm mươi đô la Mỹ năm mươi cent.`

- **Đọc số thập phân thông thường (không kèm đơn vị tiền):**
  ```excel
  =BWPVNWORDS(125.05, "")
  ```
  -> `Một trăm hai mươi lăm phẩy không năm.`

- **Đọc số viết hoa từng từ:**
  ```excel
  =BWPVNWORDS(1500000, "VND", 1)
  ```
  -> `Một Triệu Năm Trăm Nghìn Đồng Chẵn.`

### Các bí danh tương thích (Aliases)
Để tiện lợi khi gõ công thức nhanh, add-in hỗ trợ các tên hàm ngắn:
- `=VND(A1)` tương đương `=BWPVND(A1)`
- `=VNDUPPER(A1)` tương đương `=BWPVNDUPPER(A1)`
- `=VNWORDS(A1, ...)` tương đương `=BWPVNWORDS(A1, ...)`

---

## Bảo mật & Cam kết hoạt động Offline

- **100% Offline:** Toàn bộ thuật toán xử lý dữ liệu hoàn toàn cục bộ trên máy tính của bạn thông qua VBA engine nội bộ của Microsoft Excel.
- **Không thu thập dữ liệu (No Telemetry):** Không gửi bất kỳ dữ liệu nào qua mạng Internet, không kết nối máy chủ ngoài.
- **An toàn cho môi trường Doanh nghiệp:** Tuân thủ các tiêu chuẩn bảo mật dữ liệu tài chính khắt khe nhất của ngân hàng, tổ chức kiểm toán và cơ quan thuế.

---

## Hướng dẫn Biên dịch & Đóng gói (Build Instructions)

Dành cho lập trình viên muốn tự biên dịch add-in từ mã nguồn:

### Công cụ yêu cầu:
- Windows 10/11 với Microsoft Excel bản quyền đã kích hoạt.
- Python 3.10 trở lên với thư viện `pywin32`.
- Thiết lập quyền truy cập VBA: Trong Excel, vào **Trust Center** -> **Trust Center Settings...** -> **Macro Settings** -> Tích chọn **Trust access to the VBA project object model** (AccessVBOM).

### Lệnh biên dịch Release đầy đủ:
Chạy lệnh bằng Python:
```bash
python scripts/build.py --release
```

Hoặc chạy lệnh bằng PowerShell (tự động bật tạm thời AccessVBOM nếu cần và hoàn nguyên sau khi build):
```powershell
.\scripts\build.ps1 -Release -EnableAccessVBOM
```

### Kết quả sau khi Build thành công:
Hệ thống sẽ thực hiện toàn bộ 7 giai đoạn kiểm định nghiêm ngặt:
1. `Stage 0`: Đồng bộ phiên bản từ `VERSION` và kiểm tra 100% 7-bit ASCII của các thành phần mã nguồn.
2. `Stage 1`: Kiểm tra môi trường COM và Registry.
3. `Stage 2`: Khởi động Excel COM cách ly với PID riêng, biên dịch các module và form thành file `.xlam`.
4. `Stage 3`: Đóng gói trực tiếp cấu trúc OpenXML và nhúng thanh Ribbon `customUI14.xml`.
5. `Stage 4`: Tái mở kiểm định add-in trên phiên Excel mới và chạy smoke test.
6. `Stage 5`: Chạy toàn bộ bộ test tự động 3 tầng (`tests/run_tests.py`).
7. `Stage 6`: Tạo mã băm SHA-256 (`dist/BWPConvertTTNVN-v1.0.0.sha256`) và đóng gói bộ cài hoàn chỉnh `dist/BWPConvertTTNVN-v1.0.0.zip`.

---

## Bản quyền & Giấy phép (License)

Dự án được phát hành theo giấy phép mã nguồn mở [MIT License](LICENSE).

Copyright (c) 2026 IT Leon. All rights reserved.
