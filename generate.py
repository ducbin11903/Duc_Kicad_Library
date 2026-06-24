import os
import sys
import csv
import subprocess
import shutil
import re

# ==============================================================================
# CẤU HÌNH THƯ VIỆN
# ==============================================================================
LIB_NAME = "Duc_Library"
SYMBOLS_DIR = "symbols"
FOOTPRINTS_DIR = os.path.join("footprints", f"{LIB_NAME}.pretty")
MODELS_DIR = os.path.join("3dmodels", f"{LIB_NAME}.3dshapes")
TEMP_DIR = "temp"
SYM_FILE = os.path.join(SYMBOLS_DIR, f"{LIB_NAME}.kicad_sym")

# ==============================================================================
# QUẢN LÝ THƯ MỤC
# ==============================================================================
def ensure_directories():
    """Tạo các thư mục cần thiết và khởi tạo file symbol library nếu chưa có."""
    dirs = [SYMBOLS_DIR, FOOTPRINTS_DIR, MODELS_DIR, TEMP_DIR]
    for d in dirs:
        os.makedirs(d, exist_ok=True)

    # Khởi tạo file thư viện KiCad 9 trống nếu nó chưa tồn tại
    if not os.path.exists(SYM_FILE):
        with open(SYM_FILE, 'w', encoding='utf-8') as f:
            f.write('(kicad_symbol_lib (version 20231120) (generator easyeda2kicad)\n)\n')

def clear_temp():
    """Xóa dọn thư mục temp để tránh sót file của lần import trước."""
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR, ignore_errors=True)
    os.makedirs(TEMP_DIR, exist_ok=True)

# ==============================================================================
# XỬ LÝ KICAD SYMBOL (S-EXPRESSION PARSER)
# ==============================================================================
def extract_symbol_blocks(filepath):
    """
    Đọc file .kicad_sym và trích xuất từng block (symbol "...") độc lập ở TOP-LEVEL.
    Bỏ qua các sub-symbol (như _1_1, _0_1) lồng bên trong để tránh duplicate part.
    Đồng thời tự động fix tiền tố thư viện footprint.
    """
    if not os.path.exists(filepath):
        return []
    
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    blocks = []
    pos = 0  # Vị trí con trỏ đọc file
    
    while True:
        # Tìm vị trí bắt đầu của một định nghĩa symbol
        start_idx = content.find('(symbol "', pos)
        if start_idx == -1:
            break
            
        stack = 0
        end_idx = -1
        # Duyệt từng ký tự để tìm dấu ngoặc đóng của symbol này
        for i in range(start_idx, len(content)):
            if content[i] == '(':
                stack += 1
            elif content[i] == ')':
                stack -= 1
                # Khi stack về 0, ta đã tìm thấy điểm kết thúc của symbol mẹ
                if stack == 0:
                    end_idx = i + 1
                    break
                    
        if end_idx != -1:
            block_content = content[start_idx:end_idx]
            
            # Tự động đổi tiền tố footprint từ "temp:" thành thư viện chính
            block_content = re.sub(
                r'\(property\s+"Footprint"\s+"[^:]+:([^"]+)"', 
                rf'(property "Footprint" "{LIB_NAME}:\1"', 
                block_content
            )
            
            # Trích xuất tên symbol (chỉ lấy ở dòng đầu tiên của block)
            name_match = re.match(r'\(symbol\s+"([^"]+)"', block_content)
            if name_match:
                name = name_match.group(1)
                
                # Trích xuất property LCSC Part
                lcsc_match = re.search(r'\(property\s+"LCSC Part"\s+"([^"]+)"', block_content)
                lcsc = lcsc_match.group(1) if lcsc_match else None
                
                blocks.append({
                    'name': name, 
                    'lcsc': lcsc, 
                    'content': block_content
                })
            
            # QUAN TRỌNG: Nhảy con trỏ qua toàn bộ block vừa tìm được
            # Điều này giúp bỏ qua các symbol con (_0_1, _1_1) lồng bên trong
            pos = end_idx
        else:
            break
            
    return blocks

def merge_symbol(block_content):
    """Nối block symbol mới vào file library chính mà không ghi đè."""
    with open(SYM_FILE, 'r', encoding='utf-8') as f:
        content = f.read()

    # Tìm dấu ngoặc đóng cuối cùng của thư viện chính
    last_paren_idx = content.rfind(')')
    if last_paren_idx != -1:
        # Chèn block mới vào trước dấu ngoặc đóng cuối cùng
        new_content = content[:last_paren_idx] + "\n  " + block_content + "\n" + content[last_paren_idx:]
        with open(SYM_FILE, 'w', encoding='utf-8') as f:
            f.write(new_content)

# ==============================================================================
# XỬ LÝ FOOTPRINT VÀ 3D MODEL
# ==============================================================================
def process_footprints(stats):
    """Copy footprint và fix lại đường dẫn 3D."""
    temp_fp_dir = os.path.join(TEMP_DIR, "temp.pretty")
    if not os.path.exists(temp_fp_dir):
        return

    # Biểu thức chính quy phát hiện khai báo 3D model và trích xuất đúng tên file
    model_pattern = re.compile(r'\(model\s+"[^"]*?([^"/]+\.(?:step|wrl|STEP|WRL))"')

    for fp_file in os.listdir(temp_fp_dir):
        if fp_file.endswith(".kicad_mod"):
            src = os.path.join(temp_fp_dir, fp_file)
            dst = os.path.join(FOOTPRINTS_DIR, fp_file)

            # CHỐNG TRÙNG FOOTPRINT
            if os.path.exists(dst):
                continue

            with open(src, 'r', encoding='utf-8') as f:
                content = f.read()

            # FIX ĐƯỜNG DẪN 3D SANG BIẾN MÔI TRƯỜNG KICAD_USER_LIB
            fixed_content = model_pattern.sub(r'(model "${KICAD_USER_LIB}/3dmodels/' + LIB_NAME + r'.3dshapes/\1"', content)

            with open(dst, 'w', encoding='utf-8') as f:
                f.write(fixed_content)

            print(f"  ✓ Footprint copied")
            stats['fp'] += 1

def process_3d_models(stats):
    """Copy file 3D (STEP/WRL). Bỏ qua nếu đã tồn tại."""
    temp_3d_dir = os.path.join(TEMP_DIR, "temp.3dshapes")
    if not os.path.exists(temp_3d_dir):
        return

    for mod_file in os.listdir(temp_3d_dir):
        if mod_file.lower().endswith((".step", ".wrl")):
            src = os.path.join(temp_3d_dir, mod_file)
            dst = os.path.join(MODELS_DIR, mod_file)

            ext = "STEP" if mod_file.lower().endswith(".step") else "WRL"

            # CHỐNG TRÙNG 3D MODEL
            if os.path.exists(dst):
                continue

            shutil.copy2(src, dst)
            print(f"  ✓ 3D {ext} copied")
            stats['3d'] += 1

# ==============================================================================
# HÀM MAIN
# ==============================================================================
def main():
    ensure_directories()

    # 1. ĐỌC parts.csv (Hỗ trợ chuẩn CSV nhiều cột)
    if not os.path.exists('parts.csv'):
        print("Lỗi: Không tìm thấy file parts.csv")
        return

    lines = []
    # Dùng utf-8-sig để bỏ qua ký tự BOM ẩn nếu file được lưu từ Excel
    with open('parts.csv', 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        header = next(reader, None)
        
        if header:
            headers = [h.strip().lower() for h in header]
            # Xác định vị trí cột 'lcsc'
            if 'lcsc' in headers:
                lcsc_idx = headers.index('lcsc')
                for row in reader:
                    if len(row) > lcsc_idx and row[lcsc_idx].strip():
                        lines.append(row[lcsc_idx].strip())
            else:
                # Nếu không có chữ lcsc trên header, lấy giá trị đầu tiên nếu bắt đầu bằng 'C'
                if header[0].strip().upper().startswith('C'):
                    lines.append(header[0].strip())
                for row in reader:
                    if row and row[0].strip():
                        lines.append(row[0].strip())

    if not lines:
        print("File parts.csv trống hoặc không chứa cột 'lcsc'.")
        return

    print(f"\n📦 Bắt đầu generate {len(lines)} parts...\n")

    stats = {'sym': 0, 'fp': 0, '3d': 0, 'success': 0, 'skip': 0, 'err': 0}

    # Nạp dữ liệu các part đã tồn tại để check trùng tốc độ cao
    existing_blocks = extract_symbol_blocks(SYM_FILE)
    existing_symbols = {b['name'] for b in existing_blocks}
    existing_lcsc = {b['lcsc'] for b in existing_blocks if b['lcsc']}

    # Cờ để quyết định dùng shell=True trên Windows
    use_shell = (os.name == 'nt')

    for lcsc_id in lines:
        clear_temp()

        # 2. IMPORT BẰNG easyeda2kicad
        result = subprocess.run([
            "easyeda2kicad",
            "--full",
            f"--lcsc_id={lcsc_id}",
            "--output",
            os.path.join("temp", "temp") # Cấu hình ép tạo file có tiền tố là temp trong thư mục temp
        ], capture_output=True, text=True, shell=use_shell)

        if result.returncode != 0:
            print(f"─── Lỗi tải: {lcsc_id}")
            print("  ✗ Lỗi từ thư viện easyeda2kicad")
            stats['err'] += 1
            continue

        temp_sym_file = os.path.join(TEMP_DIR, "temp.kicad_sym")
        imported_blocks = extract_symbol_blocks(temp_sym_file)

        if not imported_blocks:
            print(f"─── {lcsc_id}")
            print("  ✗ Không tìm thấy symbol hợp lệ trong file tải về")
            stats['err'] += 1
            continue

        # 3. PARSE VÀ XỬ LÝ
        for block in imported_blocks:
            sym_name = block['name']
            sym_lcsc = block['lcsc']

            print(f"─── {sym_name} ({lcsc_id})")

            # CHỐNG TRÙNG SYMBOL NAME
            if sym_name in existing_symbols:
                print(f"  ⚠ Symbol {sym_name} đã tồn tại")
                print("  ⚠ Skip")
                stats['skip'] += 1
                continue

            # CHỐNG TRÙNG LCSC PART
            if sym_lcsc and sym_lcsc in existing_lcsc:
                print(f"  ⚠ LCSC {sym_lcsc} đã tồn tại")
                print("  ⚠ Skip")
                stats['skip'] += 1
                continue

            # MERGE SYMBOL
            merge_symbol(block['content'])
            
            # Cập nhật ngay vào danh sách kiểm tra trùng nội bộ tránh trùng chéo trong cùng file csv
            existing_symbols.add(sym_name)
            if sym_lcsc:
                existing_lcsc.add(sym_lcsc)

            print("  ✓ Symbol imported")
            stats['sym'] += 1

            # QUẢN LÝ FOOTPRINT VÀ 3D
            process_footprints(stats)
            process_3d_models(stats)

            stats['success'] += 1
            
        print("") # Xuống dòng giữa các part

    # Xóa temp sau khi hoàn thành
    clear_temp()

    # LOG KẾT QUẢ CUỐI CÙNG
    print("────────────────────────────────")
    print(f"✓ Symbol mới      : {stats['sym']}")
    print(f"✓ Footprint mới   : {stats['fp']}")
    print(f"✓ 3D models mới   : {stats['3d']}")
    print(f"✓ Thành công      : {stats['success']}")
    print("")
    print(f"✗ Bỏ qua          : {stats['skip']}")
    print(f"✗ Lỗi             : {stats['err']}")

if __name__ == "__main__":
    main()