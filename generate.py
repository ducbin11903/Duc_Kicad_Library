#!/usr/bin/env python3
"""
Duc KiCad Library Generator
----------------------------
Cách dùng:
  1. Thêm LCSC part number vào parts.csv
  2. Chạy: python generate.py
  3. git add . && git commit && git push

Máy khác chỉ cần git pull, không cần chạy lại script.
"""

import subprocess
import shutil
import os
import csv
import glob
import sys

# ─── Config ───────────────────────────────────────────────────────────────────
PARTS_FILE     = "parts.csv"
OUTPUT_SYMBOL  = os.path.join("symbols", "Duc_Library.kicad_sym")
OUTPUT_FP_DIR  = os.path.join("footprints", "Duc_Library.pretty")
OUTPUT_3D_DIR  = os.path.join("3dmodels", "Duc_Library.3dshapes")
TEMP_DIR       = "_temp"

SYMBOL_HEADER  = """(kicad_symbol_lib
\t(version 20231120)
\t(generator "duc_kicad_lib")
\t(generator_version "1.0")
"""
SYMBOL_FOOTER  = ")\n"
# ──────────────────────────────────────────────────────────────────────────────


def read_parts():
    if not os.path.exists(PARTS_FILE):
        print(f"✗ Không tìm thấy {PARTS_FILE}")
        sys.exit(1)
    parts = []
    with open(PARTS_FILE, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["lcsc"].strip():
                parts.append(row)
    return parts


def convert_part(lcsc: str) -> bool:
    """Chạy easyeda2kicad cho 1 LCSC part, output vào TEMP_DIR."""
    temp_output = os.path.join(TEMP_DIR, lcsc)
    result = subprocess.run(
        ["easyeda2kicad", "--full", f"--lcsc_id={lcsc}", "--output", temp_output],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"  ✗ Lỗi khi convert {lcsc}:\n{result.stderr.strip()}")
        return False
    return True


def copy_footprints(lcsc: str):
    src = os.path.join(TEMP_DIR, f"{lcsc}.pretty")
    if os.path.exists(src):
        for f in os.listdir(src):
            shutil.copy2(os.path.join(src, f), OUTPUT_FP_DIR)
            print(f"  ✓ Footprint: {f}")


def copy_3d_models(lcsc: str):
    src = os.path.join(TEMP_DIR, f"{lcsc}.3dshapes")
    if os.path.exists(src):
        for f in os.listdir(src):
            shutil.copy2(os.path.join(src, f), OUTPUT_3D_DIR)
            print(f"  ✓ 3D Model: {f}")


def extract_symbols(filepath: str) -> list:
    """
    Trích xuất tất cả top-level (symbol ...) blocks từ file .kicad_sym.
    Dùng đếm ngoặc đơn để parse chính xác S-expression.
    """
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    symbols = []
    i = 0
    n = len(content)

    while i < n:
        # Tìm block (symbol ở top-level (sau tab hoặc newline)
        idx = content.find("\n\t(symbol ", i)
        if idx == -1:
            break

        start = idx + 1  # bỏ ký tự newline đầu
        depth = 0
        j = start

        while j < n:
            if content[j] == "(":
                depth += 1
            elif content[j] == ")":
                depth -= 1
                if depth == 0:
                    symbols.append(content[start : j + 1])
                    i = j + 1
                    break
            j += 1
        else:
            break

    return symbols


def merge_symbols():
    """Merge tất cả .kicad_sym trong TEMP_DIR thành 1 file duy nhất."""
    all_symbols = []

    for sym_file in sorted(glob.glob(os.path.join(TEMP_DIR, "*.kicad_sym"))):
        extracted = extract_symbols(sym_file)
        all_symbols.extend(extracted)
        print(f"  ✓ Symbol: {os.path.basename(sym_file)} ({len(extracted)} symbol)")

    if not all_symbols:
        print("  ✗ Không có symbol nào để merge!")
        return

    os.makedirs("symbols", exist_ok=True)
    with open(OUTPUT_SYMBOL, "w", encoding="utf-8") as f:
        f.write(SYMBOL_HEADER)
        for sym in all_symbols:
            f.write(sym + "\n")
        f.write(SYMBOL_FOOTER)

    print(f"\n✓ Merged {len(all_symbols)} symbols → {OUTPUT_SYMBOL}")


def fix_3d_model_paths():
    """
    Sửa path 3D model trong tất cả .kicad_mod để dùng biến KICAD_USER_LIB.
    easyeda2kicad thường dùng path tuyệt đối hoặc path không đúng.
    """
    kicad_mod_files = glob.glob(os.path.join(OUTPUT_FP_DIR, "*.kicad_mod"))
    for mod_file in kicad_mod_files:
        with open(mod_file, "r", encoding="utf-8") as f:
            content = f.read()

        # Tìm và thay thế bất kỳ path nào trỏ tới file .step hay .wrl
        import re
        # Thay path cũ dạng bất kỳ trỏ tới .step hoặc .wrl
        content = re.sub(
            r'"[^"]*?/([^"/]+\.(?:step|wrl))"',
            lambda m: f'"${{KICAD_USER_LIB}}/3dmodels/Duc_Library.3dshapes/{m.group(1)}"',
            content,
            flags=re.IGNORECASE,
        )

        with open(mod_file, "w", encoding="utf-8") as f:
            f.write(content)


def main():
    parts = read_parts()
    if not parts:
        print("✗ Không có part nào trong parts.csv")
        sys.exit(1)

    print(f"📦 Bắt đầu generate {len(parts)} parts...\n")

    # Tạo thư mục cần thiết
    os.makedirs(TEMP_DIR, exist_ok=True)
    os.makedirs(OUTPUT_FP_DIR, exist_ok=True)
    os.makedirs(OUTPUT_3D_DIR, exist_ok=True)

    failed = []

    for part in parts:
        lcsc = part["lcsc"].strip()
        name = part.get("name", lcsc).strip()
        category = part.get("category", "").strip()
        print(f"─── [{category}] {name} ({lcsc})")

        if not convert_part(lcsc):
            failed.append(lcsc)
            continue

        copy_footprints(lcsc)
        copy_3d_models(lcsc)

    # Merge tất cả symbol
    print("\n─── Merging symbols...")
    merge_symbols()

    # Fix 3D model paths
    print("\n─── Fixing 3D model paths...")
    fix_3d_model_paths()

    # Cleanup temp
    shutil.rmtree(TEMP_DIR, ignore_errors=True)

    # Tóm tắt
    print("\n" + "=" * 50)
    print(f"✓ Hoàn thành: {len(parts) - len(failed)}/{len(parts)} parts")
    if failed:
        print(f"✗ Thất bại: {', '.join(failed)}")
    print("\nBước tiếp theo:")
    print("  git add .")
    print('  git commit -m "Update library"')
    print("  git push")


if __name__ == "__main__":
    main()
