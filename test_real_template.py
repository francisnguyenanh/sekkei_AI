import json
import time
from schemas.models import parse_request
from core.generator import ExcelGeneratorService

# 1. Tải JSON Template khổng lồ
with open('real_template_menu.json', 'r', encoding='utf-8') as f:
    template_data = json.load(f)

# 2. Tạo JSON Logic chứa dữ liệu nghiệp vụ
logic_data = {
    "single_values": {
        "system_name": "Hệ thống AI Generator",
        "project_number": "PRJ-2026-X",
        "customer_name": "Tập đoàn Demo",
        "create_date": "2026-03-30",
        "author_1": "Antigravity AI",
        "screen_id": "SCR-999",
        "screen_name": "Màn hình Danh sách Test",
        "screen_nav": "Đăng nhập -> Click Menu Chính -> Hiển thị",
        "screen_ver": "V1.0.0 (AI Generated)"
    },
    "table_data": {
        "menu_table_start": [
            # Data trải ra mảng 1D, chèn khoảng trống ("") vì Excel gốc bị gộp ô
            # B49(Text) -> C (bị gộp vào B) -> D(Icon) -> E..N (gộp) -> O(Role) -> R(Ribbon) -> AC (Link)
            ["Menu Quản lý User", "", "Icon_01", "", "", "", "", "", "", "", "", "", "", "", "Admin", "", "", "Quản lý Người Dùng", "", "", "", "", "", "", "", "", "", "", "", "", "", "Tới Màn hình QL User"],
            ["Menu Xem Báo cáo", "", "Icon_02", "", "", "", "", "", "", "", "", "", "", "", "User", "", "", "Báo cáo Doanh thu", "", "", "", "", "", "", "", "", "", "", "", "", "", "Tới Màn hình Báo cáo"]
        ]
    }
}

payload = {
    "template": template_data,
    "logic": logic_data
}

print("Bắt đầu Parse dữ liệu JSON (40.000 dòng)...")
start_time = time.time()
req = parse_request(payload)
print(f"Parse xong trong: {time.time() - start_time:.2f}s")

print("Bắt đầu vẽ Excel bằng thư viện Openpyxl...")
start_time = time.time()
svc = ExcelGeneratorService(req)
output_io = svc.generate()
print(f"Vẽ xong trong: {time.time() - start_time:.2f}s")

output_file = 'output_real_template_menu.xlsx'
with open(output_file, 'wb') as f:
    f.write(output_io.read())
    
print(f"\nThành công rực rỡ! Output lưu tại: {output_file}")
