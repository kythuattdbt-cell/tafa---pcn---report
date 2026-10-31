# Báo cáo Phòng Chăn nuôi TAFA

Hệ thống Python xử lý Excel/Word cục bộ, mặc định tháng **10/2026**, xuất PowerPoint BOD 14 slide cùng báo cáo Excel/Word. PDF là tùy chọn. Chưa sử dụng dữ liệu thực tế. Không gọi dịch vụ AI, không gửi file qua mạng và không tự sinh số liệu để bù dữ liệu thiếu.


## Windows 10/11 — Python 3.12, PowerShell, không cần Docker

Mã đọc/xuất Office dùng thư viện Python, không cần cài Excel để tạo `.xlsx` hoặc PowerPoint để tạo `.pptx`. Cần PowerPoint hoặc phần mềm tương thích để xem/chỉnh sửa. PDF là tùy chọn, cần LibreOffice; không có hoặc chuyển đổi thất bại vẫn giữ đầu ra PowerPoint.

Sau khi mã nguồn đã được đưa lên GitHub và bạn đã cài Git cùng Python 3.12 64-bit:

```powershell
git clone https://github.com/kythuattdbt-cell/tafa---pcn---report.git D:\TAFA_PCN_REPORT
Set-Location D:\TAFA_PCN_REPORT
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-lock.txt
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -m unittest discover -v
```

Nếu nhận bản ZIP mã nguồn, giải nén thư mục `TAFA_PCN_REPORT` vào ổ D: để có `D:\TAFA_PCN_REPORT`, bỏ lệnh `git clone` và chạy từ `Set-Location`. ZIP mã nguồn không chứa lịch sử `.git` hoặc cấu hình riêng; thao tác commit/PR cần thực hiện trong checkout Git riêng sau khi đồng bộ mã nguồn.

Gọi thẳng Python trong `.venv` nên không cần đổi ExecutionPolicy hoặc chạy Activate.ps1. Không sao chép virtualenv Linux sang Windows. `requirements-lock.txt` chốt bộ phiên bản đã kiểm thử trên Python 3.12/Linux; phải kiểm tra cài đặt và test trên laptop Windows trước khi dùng chính thức. Có thể dùng `requirements.txt` nếu cần giải quyết tương thích nền tảng và kiểm thử lại.

### Cấu hình OneDrive riêng trên laptop

`config.example.json` là mẫu công khai, **không chứa đường dẫn nội bộ**. Sao chép thành `config.local.json` (đã được `.gitignore` loại trừ):

```powershell
Copy-Item config.example.json config.local.json
notepad config.local.json
```

Trong `data_dir`, thay `${TAFA_DATA_DIR}` bằng đường dẫn OneDrive thật trên máy. JSON dùng dấu `/` hoặc `\\` để biểu diễn dấu phân cách Windows. Ví dụ chỉ minh họa cú pháp: `"D:/OneDrive/Thu_muc_du_lieu"`. Không đưa cấu hình riêng này lên GitHub.

Hoặc giữ `${TAFA_DATA_DIR}` và nhập đường dẫn qua biến môi trường trong cửa sổ PowerShell hiện tại:

```powershell
$env:TAFA_DATA_DIR = Read-Host 'Nhap duong dan thu muc OneDrive duoc cong ty cho phep'
```

Không tự động tìm tài khoản OneDrive hay gửi tệp ra ngoài. Đường dẫn có dấu tiếng Việt/khoảng trắng được hỗ trợ. Chọn **Always keep on this device / Luôn giữ trên thiết bị này** cho các file nguồn, chờ OneDrive đồng bộ và đóng file đang chỉnh sửa trước khi chạy. Chương trình đọc `.xlsx`, `.xlsm`, `.docx` trong thư mục được chọn và toàn bộ thư mục con khi `recursive=true`; giữ đường dẫn tương đối trong phụ lục để phân biệt các file trùng tên. Không đi theo thư mục symlink/junction để tránh quét ra ngoài phạm vi được chọn.

Thiết lập trong cấu hình riêng:

- `month`: `2026-10`.
- `recursive`: `true`.
- `sample`: `false` — chế độ chính thức.
- `confirm_real_data`: chỉ đổi thành `true` sau khi người phụ trách xác nhận nguồn là dữ liệu thật được phép xử lý.
- `output_dir`: ví dụ `output/2026-10-official`; thư mục đầu ra phải tách biệt và không lồng trong thư mục nguồn. Chọn thư mục mới cho mỗi lần chạy.
- `pdf`: `false` nếu chỉ cần PowerPoint; `true` nếu muốn thử xuất thêm PDF.
- `font`, `fallback_fonts`: ưu tiên Montserrat, dự phòng các font Windows phổ biến. Nếu không tìm được font yêu cầu, dùng Arial. Cài Montserrat trước khi chạy và mở lại terminal nếu cần; việc thay font có thể đổi cách xuống dòng, cần xem lại bản trình chiếu.
- `libreoffice_path`: `null` để tự tìm trên PATH hoặc vị trí cài chuẩn Windows; có thể đặt đường dẫn đầy đủ đến `soffice.exe` trong cấu hình riêng.

Đường dẫn tương đối trong JSON được tính từ thư mục chứa cấu hình. Tùy chọn CLI được ưu tiên hơn JSON; có `--no-pdf`, `--no-sample`, `--no-recursive` để tắt các tùy chọn tương ứng. File cấu hình UTF-8 có/không có BOM đều đọc được.

```powershell
.\.venv\Scripts\python.exe bod.py --config config.local.json
# Chọn tháng tiếp theo và thư mục đầu ra mới:
.\.venv\Scripts\python.exe bod.py --config config.local.json --month 2026-11 --output output/2026-11-official
```

Nếu chọn `output/2026-10-official`, kết quả chính nằm tại:

`D:\TAFA_PCN_REPORT\output\2026-10-official\Bao_cao_BOD_PCN_T10_2026.pptx`

Chương trình vẫn yêu cầu đúng cấu trúc sheet/cột được mô tả bên dưới. Không đoán ý nghĩa các báo cáo Excel/Word tùy ý trong OneDrive. Nếu file thật khác mẫu, cần chuẩn hóa bản sao vào thư mục nhập liệu được phép; không sửa nguồn gốc hoặc tự bù số liệu. Dữ liệu ngoài kỳ, khóa trùng và file không đọc được phải được đối soát trước khi xuất.

### Phân biệt thử nghiệm và chính thức

```powershell
# Chỉ tạo dữ liệu giả lập khi người dùng chủ động gọi trình tạo mẫu:
.\.venv\Scripts\python.exe make_bod_sample.py --directory data/demo --month 2026-10
.\.venv\Scripts\python.exe bod.py --data data/demo --output output/demo --month 2026-10 --sample
```

Chương trình **không tự tạo/tìm dữ liệu mẫu để thay nguồn thiếu**. Chế độ chính thức yêu cầu xác nhận nguồn, chặn tên file hoặc nội dung có các nhãn giả lập đã biết, kể cả khi đã đổi tên file, và dừng khi thiếu nhóm dữ liệu tháng báo cáo. Trộn nguồn thật và nguồn có nhãn mẫu cũng bị chặn. `report.py` áp dụng cùng cơ chế với `--confirm-real-data`; `--sample` phải được chỉ định rõ để chạy thử.

Không có cách xác thực tuyệt đối dữ liệu thật chỉ từ các con số hoặc phát hiện mọi bộ dữ liệu giả không gắn nhãn. Xác nhận nguồn và kiểm tra của người phụ trách vẫn bắt buộc. Các bài test chỉ dùng fixture giả lập trong thư mục tạm, không phải xác nhận dữ liệu sản xuất thực.

## Đưa riêng mã nguồn lên GitHub

Chỉ commit các tệp trong danh sách `$SourceFiles` bên dưới; không đưa tệp trong `data/` hoặc `output/` vào commit. Không dùng `git add .` hoặc `git add -f`. Cấu hình riêng, `.env`, khóa/chứng thư, virtualenv, Excel/Word, báo cáo, dữ liệu CSV và thư mục `data`/`output` (ngoại trừ `.gitkeep` trống) bị loại khỏi Git. `.gitignore` không nhận diện được nội dung bí mật đặt trong tệp mã nguồn hoặc tệp đã theo dõi; luôn duyệt bản diff.

Sau khi đã xem lại nội dung và quyết định commit, chạy từ thư mục dự án:

```powershell
git switch -c feature/bod-windows-onedrive
$SourceFiles = @(
  '.gitignore', 'README.md', 'config.example.json',
  'bod.py', 'report.py', 'local_runtime.py',
  'make_bod_sample.py', 'make_sample.py',
  'requirements.txt', 'requirements-lock.txt',
  'tests/__init__.py', 'tests/test_bod.py',
  'tests/test_report.py', 'tests/test_windows_runtime.py'
)
$SourceFiles | ForEach-Object { git add -- $_ }
git diff --cached --name-only
git diff --cached --check
git diff --cached
# Chỉ commit sau khi xác nhận danh sách và diff không chứa dữ liệu nội bộ:
git commit -m "Support Windows and configurable OneDrive BOD inputs"
```

Push và tạo Pull Request là bước riêng, chỉ thực hiện sau khi được xác nhận và có quyền GitHub:

```powershell
git push -u origin feature/bod-windows-onedrive
# Trên GitHub: Compare & pull request, base main; chỉ chọn nhánh mã nguồn vừa push.
```

Không upload dữ liệu qua PR, issue, release, artifact CI hoặc snapshot môi trường. Dữ liệu thật tiếp tục ở OneDrive/máy công ty và đầu ra nội bộ. Không chuyển file qua GitHub để tải về laptop.

## Cài đặt và chạy trên Linux

Khuyến nghị Python 3.12 (phiên bản đã kiểm thử). Từ thư mục kho mã:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python report.py --data data --output output --month 2026-10 --confirm-real-data
```

Chương trình đọc mọi workbook `.xlsx`, `.xlsm` trong `data` (thêm `--recursive` để đọc thư mục con; bỏ file khóa `~$`). `.xls` được báo lỗi và cần chuyển sang `.xlsx`. Macro không được thực thi. Không cần Microsoft Office. Dữ liệu có thể nằm trong nhiều file và nhiều sheet; tên sheet phải khớp mẫu bên dưới, không phân biệt hoa thường. Sheet ngoài mẫu được ghi cảnh báo và bỏ qua.

## Mẫu đầu vào bắt buộc

Mỗi sheet có một hàng tiêu đề ở dòng 1, mỗi dòng sau là một bản ghi; không gộp ô, không thêm dòng tổng/lũy kế. Sử dụng giá trị, không dùng công thức chưa có kết quả được Excel lưu. Số dùng kiểu số Excel, không nhập chuỗi phân cách hàng nghìn. `ngay` là ngày Excel hoặc chuỗi ISO `2026-10-31`. Tất cả ngày phải trong tháng được chọn. `khu` là A, B, C, D, E hoặc F; không tự tạo khu chưa được cung cấp.

| Sheet | Cột bắt buộc | Ý nghĩa / khóa duy nhất |
|---|---|---|
| `dan_ga` | `ngay, phap_nhan, khu, so_con` | Ảnh chụp đàn theo ngày/pháp nhân/khu; số con nguyên không âm |
| `trung` | `ngay, phap_nhan, khu, tong_trung, trung_dat` | Phát sinh ngày/pháp nhân/khu; số trứng nguyên, đạt không vượt tổng |
| `cam` | `ngay, phap_nhan, khu, loai_cam, khoi_luong_kg` | Phát sinh ngày/pháp nhân/khu/loại cám; đơn vị kg |
| `kpi` | `thang, ma_nv, ho_ten, chi_tieu, trong_so, muc_tieu, thuc_hien, chieu` | Khóa tháng/mã nhân viên/chỉ tiêu; `thang=2026-10`; `chieu=tang` hoặc `giam` |

Mỗi nhân viên có tổng `trong_so` bằng 100 (nhập 60 cho 60%, không nhập 0.6). Mục tiêu > 0; thực hiện không âm; KPI chiều giảm yêu cầu thực hiện > 0. Mã nhân viên phải thống nhất tên và nên lưu dạng văn bản để giữ số 0 đầu mã.

Dữ liệu khác cấu trúc cần được chuẩn hóa vào mẫu; chương trình không đoán tên cột hoặc ý nghĩa của sheet. Dữ liệu trùng khóa giữa các file bị chặn để tránh cộng lặp. Nếu cần phân tích theo ca, lô hoặc trại, tổng hợp có kiểm soát về khóa trên trước khi chạy.

## Phương pháp và kết quả

- Tổng đàn: lấy **ngày gần nhất chung của bộ dữ liệu**, tổng hợp theo pháp nhân/khu. Cảnh báo nếu chưa đến cuối tháng hoặc ngày chốt thiếu cặp pháp nhân/khu từng xuất hiện. Không cộng ảnh chụp nhiều ngày.
- Trứng: cộng phát sinh; tỷ lệ đạt = tổng đạt / tổng trứng × 100. Tổng bằng 0 thì tỷ lệ để trống. Chưa có chỉ tiêu kích cỡ, trọng lượng hay phân hạng nên không suy diễn chất lượng khác.
- Cám: cộng khối lượng kg theo pháp nhân/khu/loại. Không tự suy ra hệ số chuyển đổi thức ăn khi chưa có cơ sở phù hợp.
- KPI: chiều tăng = thực hiện/mục tiêu × 100; chiều giảm = mục tiêu/thực hiện × 100. Điểm tổng là tổng điểm × trọng số/100, không giới hạn trần. Đồng điểm đồng hạng kiểu 1, 1, 3. Cần phê duyệt chính sách này trước khi dùng quyết định nhân sự; dữ liệu giả lập chỉ minh họa.

Xuất `TAFA_2026-10.xlsx` (bảng tổng hợp, biểu đồ đàn, chi tiết KPI, ghi chú, kiểm tra dữ liệu) và `TAFA_2026-10.docx` (tóm tắt điều hành và bốn nhóm phân tích). Không ghi đè báo cáo đã tồn tại: chọn thư mục output mới. Báo cáo không xác nhận đủ dữ liệu mọi ngày hoặc mọi pháp nhân; người phụ trách phải đối soát phạm vi nguồn trước khi phát hành.

Nếu có lỗi, mã thoát là 2, chỉ tạo `loi_du_lieu.csv` và không tạo báo cáo mới. Báo cáo cũ trong output không đại diện cho lần chạy lỗi. Kiểm tra tệp lỗi, sửa nguồn và chạy lại. Tệp lỗi cũng là tài liệu nội bộ.

## Kiểm thử hoàn toàn bằng dữ liệu giả lập

```bash
python -m unittest discover -v
python make_sample.py --directory /tmp/tafa-demo/data
python report.py --data /tmp/tafa-demo/data --output /tmp/tafa-demo/output --sample
```

Chọn thư mục mẫu mới nếu file đã tồn tại. File mẫu có tên `MAU_GIA_LAP.xlsx`; báo cáo mẫu có hậu tố `MAU_GIA_LAP` và nhãn trong nội dung. Mẫu chỉ có khu A/B để minh họa, không đại diện hoạt động TAFA. Kiểm thử kiểm tra phép tổng hợp, ngày chốt, tỷ lệ có trọng số, KPI, đọc nhiều workbook, file Word/Excel và các dữ liệu lỗi (âm, vô hạn, sai tháng, sai khu, thiếu giá trị, trùng khóa, sai trọng số).

## Cung cấp Excel an toàn

1. Giữ Excel thực tế trên máy hoặc vùng lưu trữ nội bộ được TAFA cho phép. Đặt bản sao vào `data/` bằng kênh truyền tệp được tổ chức phê duyệt; không dán dữ liệu nhân sự/sản xuất vào chat hoặc issue/PR GitHub.
2. Chỉ dùng môi trường đám mây nếu TAFA cho phép xử lý dữ liệu tại đó. Nếu chưa được phép, tải mã chương trình về và chạy trên máy nội bộ. Không cần gửi file thực tế cho người phát triển.
3. `data/`, `output/`, mọi `.xlsx/.xlsm/.xls/.docx` và `.env` được `.gitignore` loại khỏi Git. Đây không phải biện pháp phân quyền hay mã hóa; không dùng `git add -f`. CSV dữ liệu đặt trong `data/` hoặc `output/`, không đặt bên ngoài hai thư mục này.
4. Trước khi commit, kiểm tra `git status --short` và `git diff --cached --name-only`. Không commit file nguồn, báo cáo, log chứa mã/tên nhân viên, hoặc snapshot môi trường có dữ liệu thực tế. Chia sẻ báo cáo qua kênh nội bộ có phân quyền và xóa bản tạm theo quy định TAFA.

Tạo thư mục `data/` để đặt nguồn khi cần; chương trình tự tạo thư mục `output` khi chạy. Các tệp trong hai thư mục này không được đưa vào commit mã nguồn. Chương trình không thực hiện commit hoặc push.

## PowerPoint BOD tự động (mới)

Chạy độc lập bằng `bod.py`; luồng Excel/Word cũ `report.py` vẫn sử dụng được. Dùng đủ bốn thư viện `pandas`, `openpyxl`, `python-docx`, `python-pptx`. Cài lại `requirements.txt` khi nâng cấp. Bản thử nghiệm chỉ sử dụng số liệu giả lập, chưa xác nhận trên biểu mẫu thực tế của TAFA.

```bash
# Từ thư mục dự án, sau khi kích hoạt môi trường Python:
python make_bod_sample.py --directory /tmp/tafa-bod-demo-moi/data --month 2026-10
python bod.py --data /tmp/tafa-bod-demo-moi/data --output /tmp/tafa-bod-demo-moi/output --month 2026-10 --sample --pdf
python -m unittest discover -v
```

Dữ liệu thật trong môi trường được công ty cho phép:

```bash
python bod.py --data data --output output/2026-10-chinh-thuc --month 2026-10 --confirm-real-data --pdf
# Tháng tiếp theo: thay bộ dữ liệu đến hết tháng đó, gồm dữ liệu từ đầu năm.
python bod.py --data data --output output/2026-11-chinh-thuc --month 2026-11 --confirm-real-data --pdf
```

Không dùng `--sample` cho dữ liệu thật. `--sample` chỉ gắn nhãn, không sinh hoặc thay số liệu. Chương trình không ghi đè đầu ra; chọn thư mục mới cho mỗi lần chạy. Không trộn dữ liệu mẫu vào `data` thực tế. Các file nguồn không bị sửa.

### Nguồn Excel và Word cho BOD

Bốn sheet `dan_ga`, `trung`, `cam`, `kpi` dùng cấu trúc ở trên, nay được phép chứa **từ tháng 1 đến tháng được chọn trong cùng năm**. Mỗi nhóm phải có dữ liệu trong tháng báo cáo. Bộ BOD không yêu cầu đủ lịch sử mới chạy: tháng/ngày thiếu được ghi rõ, không tự điền 0 cho tháng vắng dữ liệu. Muốn có lũy kế đầy đủ, cung cấp đầy đủ các phát sinh từ đầu năm theo từng phạm vi pháp nhân/khu. Không dùng số lũy kế đã cộng sẵn trong các sheet phát sinh.

Các sheet bổ sung không bắt buộc; thiếu sẽ ghi “Chưa đủ căn cứ” cho phân tích liên quan:

| Sheet | Cột bắt buộc | Quy tắc |
|---|---|---|
| `ga_de` | `ngay, phap_nhan, khu, so_ga_de` | Số gà đẻ thực tế trong ngày, số nguyên không âm; duy nhất theo ngày/pháp nhân/khu |
| `bien_dong` | `thang, phap_nhan, khu, dau_ky, nhap, xuat, chet, loai, cuoi_ky` | Số con nguyên không âm, theo tháng/pháp nhân/khu; đầu kỳ + nhập − xuất − chết − loại = cuối kỳ |
| `ke_hoach` | `thang, chi_so, ke_hoach, don_vi` | Một mục tiêu tổng toàn phạm vi nguồn cho tháng/chỉ số; không trộn kế hoạch theo trại với tổng thực tế |

`chi_so` được hỗ trợ: `tong_dan` (con), `tong_trung` (quả), `trung_dat` (quả), `cam_kg` (kg), `ty_le_de_pct` (%), `ty_le_dat_pct` (%). Nhập tỷ lệ 85 cho 85%, không nhập 0.85. Tháng dùng `YYYY-MM`; mục tiêu không âm, tỷ lệ không vượt 100. Mục tiêu bằng 0 vẫn thể hiện chênh lệch nhưng không chia để tính TH/KH. Chênh lệch tỷ lệ dùng **điểm phần trăm**, không phải tỷ lệ tăng trưởng tương đối.

Tỷ lệ đẻ = tổng trứng / tổng số gà đẻ-ngày × 100. Hai nguồn phải khớp hoàn toàn khóa ngày/pháp nhân/khu trong tháng; không dùng tổng đàn làm mẫu số. Số gà đẻ = 0 trong khi có trứng khiến tỷ lệ chưa đủ căn cứ. Bảng biến động được kiểm tra cân đối nội bộ và cảnh báo nếu không khớp ảnh chụp đàn cuối tháng. Không suy ra nguyên nhân hao hụt khi thiếu dữ liệu luồng.

Word `.docx` được đọc để lấy ngữ cảnh, nguyên nhân khai báo, giải pháp và kiến nghị. Để đưa nội dung vào slide, tạo bảng có hàng tiêu đề:

| thang | nhom | noi_dung | can_cu |
|---|---|---|---|
| 2026-10 | nguyen_nhan | Nội dung cần đối soát | Tên biên bản / nhật ký / hồ sơ làm căn cứ |

`nhom` nhận `ton_tai`, `nguyen_nhan`, `giai_phap`, `kien_nghi`. Chỉ nội dung đúng tháng được chọn mới lên slide; nội dung khác tháng được giữ làm tham khảo. Đoạn văn và bảng không đúng mẫu được giữ trong phụ lục để đối soát, không tự suy diễn số liệu hoặc nguyên nhân. **Số liệu tính toán lấy từ Excel**; nếu số chỉ nằm trong Word, cần người phụ trách chuẩn hóa vào sheet Excel tương ứng trước khi chạy. Chương trình không xác minh được tính đúng đắn của hồ sơ căn cứ; các nhận định Word luôn gắn nguồn và nhãn chưa xác minh.

### Cấu trúc 14 slide và thiết kế

1. Bìa báo cáo BOD.
2. Executive Summary.
3. Tổng đàn tại ngày chốt.
4. Biến động đàn.
5. Sản lượng trứng.
6. Tỷ lệ đẻ.
7. Chất lượng trứng.
8. Cám sử dụng.
9. KPI sản xuất, kế hoạch/thực tế.
10. KPI và xếp hạng nhân sự.
11. Lũy kế từ đầu năm đến tháng chọn, xu hướng thực tế/kế hoạch trứng và tổng cám.
12. Tồn tại và căn cứ nguyên nhân.
13. Giải pháp đề xuất.
14. Kiến nghị Ban Lãnh đạo.

Thiết kế 16:9, nền sáng, xanh lá `#145C43`, Gold `#C6A34B`, dùng Montserrat nếu máy có font; dự phòng lần lượt Aptos, Calibri, Arial, DejaVu Sans. Nhận diện từ tệp font và Windows Registry, không gọi lệnh hệ điều hành để liệt kê font. Đây là palette mặc định, chưa đối chiếu cẩm nang thương hiệu chính thức. Bảng, thẻ số, chữ và biểu đồ PowerPoint đều chỉnh sửa được; dữ liệu biểu đồ được nhúng trong file. Không dùng ảnh toàn slide.

Để giữ bố cục thoáng, biểu đồ nhóm trên 8 nhóm sẽ cộng nhóm còn lại thành “Các nhóm khác”; bảng nhân sự hiển thị tối đa 8 người, slide nội dung tối đa 4 mục. Nội dung dài được đánh dấu rút gọn và giữ đầy đủ trong phụ lục. Không bỏ số liệu khỏi tổng cộng. Điểm KPI nhân sự dùng công thức trong hướng dẫn cũ; “TH/KH” của KPI sản xuất chỉ là tỷ lệ so sánh. Cám thấp hơn kế hoạch không tự chứng minh hiệu quả; cần xét quy mô và sản lượng. Không tự đề xuất ngân sách, người chịu trách nhiệm, thời hạn hoặc mức cam kết khi thiếu căn cứ.

### Tệp xuất và kiểm tra bàn giao

Trong thư mục `--output`:

- `Bao_cao_BOD_PCN_T10_2026.pptx` — PowerPoint 14 slide; tên tự đổi theo tháng/năm.
- `Bao_cao_BOD_PCN_T10_2026.pdf` — có khi thêm `--pdf` và có LibreOffice/soffice. Xuất từ bản PPTX vừa tạo, kiểm tra SHA-256 của PPTX trước/sau để bảo đảm không sửa nội dung gốc. Có thể khác cách dựng font/biểu đồ giữa các phần mềm.
- `Doi_soat_BOD.xlsx` — số chi tiết, công thức/phương pháp, độ phủ ngày, kế hoạch, biến động, toàn bộ nội dung Word và danh sách nguồn kèm SHA-256.
- `Kiem_tra_BOD.json` — tháng, nhãn mẫu, số slide/biểu đồ, font, trạng thái PDF và giới hạn kiểm tra.
- `Loi_BOD.json` — chỉ khi lỗi dữ liệu; không tạo PowerPoint mới. Mã thoát 2. Báo cáo cũ ở thư mục khác không đại diện lần chạy lỗi.

Kiểm tra tự động: kiểu số, số âm/vô hạn, thiếu giá trị, trùng khóa, tháng, đơn vị, số trứng đạt, trọng số KPI, cân đối đàn, tổng cộng, số slide, đối tượng trong khung, ước lượng sức chứa khung chữ, mở lại PPTX và sự hiện diện biểu đồ nhúng. Kiểm thử còn đối chiếu số tổng và tỷ lệ với kỳ vọng độc lập, dữ liệu thiếu, đổi tháng và các đầu vào lỗi. Kiểm tra hình học/ước lượng chữ không thay thế việc xem bản trình chiếu cuối cùng trong PowerPoint; đặc biệt cần xem lại khi đổi font, nhập tên dài hoặc mở bằng phiên bản Office khác.

### Lấy PowerPoint về máy và dữ liệu nội bộ

Mở thư mục output trong trình duyệt tệp của môi trường và chọn tải xuống `.pptx`; nếu chat có đường dẫn tệp bàn giao, có thể nhấp trực tiếp. Để trình chiếu chỉ cần `.pptx` vì biểu đồ đã nhúng; phụ lục Excel dùng khi đối soát. Chia sẻ PDF/PPTX theo quyền truy cập nội bộ — chúng có thể chứa KPI nhân sự. Nếu giao diện không hỗ trợ tải xuống, dùng kênh chuyển tệp được công ty phê duyệt; không dùng GitHub làm nơi trung chuyển.

Chỉ đưa Excel/Word thật vào môi trường đã được TAFA cho phép; nếu chưa có phê duyệt đám mây, chạy mã trên máy nội bộ. `data/`, `output/`, `.pptx`, `.pdf`, `.xlsx`, `.docx` được bỏ qua bởi Git. File JSON/CSV nội bộ phải nằm trong `output/` hoặc `data/`. Không dùng `git add -f`, không gửi file qua issue/PR/chat, không publish snapshot chứa dữ liệu thực. `.gitignore` không mã hóa hoặc phân quyền: áp dụng quyền truy cập và chính sách lưu/xóa dữ liệu của công ty.

`requirements-lock.txt` lưu phiên bản đã kiểm thử trên Python 3.12 của môi trường này; để tái lập đúng bộ thư viện, dùng `python -m pip install -r requirements-lock.txt`. Không chứa thông tin xác thực. Nhãn nguồn trong Excel phụ lục và workbook biểu đồ được ghi như văn bản để tránh bị diễn giải thành công thức.
