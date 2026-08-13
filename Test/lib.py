#!/usr/bin/env python3
"""
lib.py — Duc KiCad Library Manager
====================================
Cach dung:
  python lib.py          -> Download tu Queue (LCSC/EasyEDA) -> Library
  python lib.py custom   -> Tao symbol thu cong tu sheet CustomSym -> Library
"""

import os, sys, re, subprocess, shutil, platform
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==============================================================================
# CAU HINH
# ==============================================================================
LIB_PREFIX     = "Duc"
SYMBOLS_DIR    = "symbols"
FOOTPRINTS_DIR = os.path.join("footprints", f"{LIB_PREFIX}.pretty")
MODELS_DIR     = os.path.join("3dmodels",   f"{LIB_PREFIX}.3dshapes")
TEMP_DIR       = "temp"
PARTS_FILE     = "parts.xlsx"
SHEET_LIBRARY  = "Library"
SHEET_QUEUE    = "Queue"
SHEET_CUSTOM   = "CustomSym"
DEFAULT_CAT    = "Misc"

# Custom symbol constants
PIN_LENGTH  = 2.54    # 100 mils
PIN_SPACING = 2.54
BODY_MIN_W  = 10.16   # 400 mils
FONT_SIZE   = 1.27    # 50 mils

# ==============================================================================
# MAPPING
# ==============================================================================
CATEGORY_MAP = {
    "resistors":"Resistors","resistor":"Resistors","res":"Resistors","r":"Resistors",
    "capacitors":"Capacitors","capacitor":"Capacitors","cap":"Capacitors","c":"Capacitors",
    "inductors":"Inductors","inductor":"Inductors","ind":"Inductors","l":"Inductors",
    "diodes":"Diodes","diode":"Diodes","d":"Diodes",
    "transistors":"Transistors","transistor":"Transistors","mosfet":"Transistors",
    "bjt":"Transistors","q":"Transistors",
    "mcus":"MCUs","mcu":"MCUs","microcontroller":"MCUs",
    "ic":"ICs","ics":"ICs","logic":"ICs",
    "power":"Power","ldo":"Power","dcdc":"Power","regulator":"Power",
    "sensor":"Sensors","sensors":"Sensors",
    "connectors":"Connectors","connector":"Connectors","conn":"Connectors",
    "crystals":"Crystals","crystal":"Crystals","osc":"Crystals",
    "memory":"Memory","flash":"Memory","eeprom":"Memory",
    "rf":"RF","wireless":"RF","bluetooth":"RF","wifi":"RF",
    "analog":"Analog","opamp":"Analog","adc":"Analog","dac":"Analog",
    "interface":"Interface",
    "optocouplers":"Optocouplers","optocoupler":"Optocouplers",
    "misc":"Misc","other":"Misc",
}

PIN_TYPE_MAP = {
    "in":"input","input":"input","out":"output","output":"output",
    "io":"bidirectional","bidir":"bidirectional","bidirectional":"bidirectional",
    "pwr":"power_in","power":"power_in","power_in":"power_in",
    "vdd":"power_in","vcc":"power_in","gnd":"power_in",
    "pwr_out":"power_out","power_out":"power_out",
    "oc":"open_collector","open_collector":"open_collector",
    "oe":"open_emitter","open_emitter":"open_emitter",
    "nc":"no_connect","no_connect":"no_connect",
    "passive":"passive","p":"passive",
    "unspec":"unspecified","unspecified":"unspecified",
    "tri":"tri_state","tri_state":"tri_state",
    "clk":"clock","clock":"clock",
}
PIN_SIDE_MAP = {
    "l":"L","left":"L","r":"R","right":"R",
    "t":"T","top":"T","b":"B","bottom":"B",
}

def normalize_category(raw):
    key = raw.strip().lower()
    return CATEGORY_MAP.get(key, raw.strip().title() if raw.strip() else DEFAULT_CAT)

def sym_file_for(category):
    return os.path.join(SYMBOLS_DIR, f"{LIB_PREFIX}-{category}.kicad_sym")

# ==============================================================================
# EXCEL STYLES
# ==============================================================================
THIN = Border(
    left=Side(style="thin",color="D0D0D0"), right=Side(style="thin",color="D0D0D0"),
    top=Side(style="thin",color="D0D0D0"),  bottom=Side(style="thin",color="D0D0D0"),
)
FONT_NORMAL = Font(name="Arial", size=10)
ALIGN_CTR   = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT  = Alignment(horizontal="left",   vertical="center")
FILL_GRAY   = PatternFill("solid", fgColor="F2F2F2")
FILL_GREEN  = PatternFill("solid", fgColor="E2EFDA")
FILL_RED    = PatternFill("solid", fgColor="FFDDE0")

# ==============================================================================
# KIEM TRA FILE EXCEL DANG MO
# ==============================================================================
def check_file_not_open(filepath):
    if not os.path.exists(filepath):
        return
    if os.name == "nt":
        try:
            fd = os.open(filepath, os.O_RDWR)
            os.close(fd)
        except OSError:
            print("="*52)
            print("LOI: File parts.xlsx dang duoc mo boi Excel!")
            print("  -> Dong file Excel truoc, sau do chay lai script.")
            print("="*52)
            sys.exit(1)
    else:
        folder   = os.path.dirname(os.path.abspath(filepath))
        basename = os.path.basename(filepath)
        lockfile = os.path.join(folder, f".~lock.{basename}#")
        if os.path.exists(lockfile):
            print("="*52)
            print("LOI: File parts.xlsx dang duoc mo boi LibreOffice!")
            print("  -> Dong file truoc, sau do chay lai script.")
            print("="*52)
            sys.exit(1)

def open_excel_file(filepath):
    print(f"\nDang mo file {filepath}...")
    try:
        if platform.system() == "Windows":
            os.startfile(filepath)
        elif platform.system() == "Darwin":
            subprocess.call(["open", filepath])
        else:
            subprocess.call(["xdg-open", filepath])
    except Exception as e:
        print(f"  Khong the tu dong mo: {e}")

# ==============================================================================
# EXCEL HELPERS (DUNG CHUNG)
# ==============================================================================
def add_to_library(wb, lcsc, name, category, note, sym_ok, fp_ok, model_ok):
    ws   = wb[SHEET_LIBRARY]
    nrow = ws.max_row + 1
    now  = datetime.now().strftime("%Y-%m-%d %H:%M")
    vals = [lcsc, name, category, note,
            "v" if sym_ok else "x",
            "v" if fp_ok  else "x",
            "v" if model_ok else "x",
            now]
    for col, val in enumerate(vals, 1):
        cell = ws.cell(row=nrow, column=col, value=val)
        cell.font = FONT_NORMAL; cell.border = THIN
        if col in (5,6,7):
            cell.alignment = ALIGN_CTR
            cell.fill = FILL_GREEN if val == "v" else FILL_RED
        elif col == 8:
            cell.alignment = ALIGN_CTR; cell.fill = FILL_GRAY
        else:
            cell.alignment = ALIGN_LEFT; cell.fill = FILL_GRAY

def get_library_lcsc(wb):
    ws, result = wb[SHEET_LIBRARY], set()
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0]:
            result.add(str(row[0]).strip())
    return result

# ==============================================================================
# KICAD SYMBOL FILE HELPERS (DUNG CHUNG)
# ==============================================================================
def ensure_directories():
    for d in [SYMBOLS_DIR, FOOTPRINTS_DIR, MODELS_DIR, TEMP_DIR]:
        os.makedirs(d, exist_ok=True)

def ensure_sym_file(sym_file):
    if not os.path.exists(sym_file):
        with open(sym_file, "w", encoding="utf-8") as f:
            f.write("(kicad_symbol_lib (version 20231120) (generator duc_lib)\n)\n")

def clear_temp():
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR, ignore_errors=True)
    os.makedirs(TEMP_DIR, exist_ok=True)

def extract_symbol_blocks(filepath):
    if not os.path.exists(filepath):
        return []
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()
    blocks, pos = [], 0
    while True:
        start = content.find('(symbol "', pos)
        if start == -1: break
        stack, end = 0, -1
        for i in range(start, len(content)):
            if content[i] == "(":   stack += 1
            elif content[i] == ")":
                stack -= 1
                if stack == 0: end = i + 1; break
        if end == -1: break
        block = content[start:end]
        m = re.match(r'\(symbol\s+"([^"]+)"', block)
        if m:
            lm   = re.search(r'\(property\s+"LCSC Part"\s+"([^"]+)"', block)
            dm   = re.search(r'\(property\s+"Description"\s+"([^"]*)"', block)
            blocks.append({"name": m.group(1),
                           "lcsc": lm.group(1) if lm else None,
                           "desc": dm.group(1) if dm else "",
                           "content": block})
        pos = end
    return blocks

def merge_symbol(block, sym_file):
    with open(sym_file, "r", encoding="utf-8") as f:
        content = f.read()
    block = re.sub(
        r'\(property\s+"Footprint"\s+"[^:]+:([^"]+)"',
        rf'(property "Footprint" "{LIB_PREFIX}:\1"', block
    )
    last = content.rfind(")")
    if last != -1:
        with open(sym_file, "w", encoding="utf-8") as f:
            f.write(content[:last] + "\n  " + block + "\n" + content[last:])

def symbol_exists(sym_file, sym_name):
    if not os.path.exists(sym_file): return False
    with open(sym_file, "r", encoding="utf-8") as f:
        return f'(symbol "{sym_name}"' in f.read()

def load_all_existing():
    names, lcsc_map = set(), {}
    if os.path.exists(SYMBOLS_DIR):
        for fname in os.listdir(SYMBOLS_DIR):
            if fname.startswith(f"{LIB_PREFIX}-") and fname.endswith(".kicad_sym"):
                category = fname.replace(f"{LIB_PREFIX}-","").replace(".kicad_sym","")
                for b in extract_symbol_blocks(os.path.join(SYMBOLS_DIR, fname)):
                    names.add(b["name"])
                    if b["lcsc"]:
                        lcsc_map[b["lcsc"]] = {"name":b["name"],"desc":b["desc"],"category":category}
    return names, lcsc_map

# ==============================================================================
# QUEUE MODE — process_footprints / process_3d_models
# ==============================================================================
def process_footprints(stats):
    fp_dir  = os.path.join(TEMP_DIR, "temp.pretty")
    pattern = re.compile(r'\(model\s+"[^"]*?([^"/]+\.(?:step|wrl|STEP|WRL))"')
    added   = False
    if not os.path.exists(fp_dir): return False
    for fname in os.listdir(fp_dir):
        if not fname.endswith(".kicad_mod"): continue
        src = os.path.join(fp_dir, fname)
        dst = os.path.join(FOOTPRINTS_DIR, fname)
        if os.path.exists(dst): added = True; continue
        with open(src,"r",encoding="utf-8") as f: data = f.read()
        data = pattern.sub(r'(model "${KICAD_USER_LIB}/3dmodels/'+LIB_PREFIX+r'.3dshapes/\1"', data)
        with open(dst,"w",encoding="utf-8") as f: f.write(data)
        print(f"  + Footprint: {fname}")
        stats["fp"] += 1; added = True
    return added

def process_3d_models(stats):
    d3_dir = os.path.join(TEMP_DIR, "temp.3dshapes")
    added  = False
    if not os.path.exists(d3_dir): return False
    for fname in os.listdir(d3_dir):
        if not fname.lower().endswith((".step",".wrl")): continue
        src = os.path.join(d3_dir, fname)
        dst = os.path.join(MODELS_DIR, fname)
        if os.path.exists(dst): added = True; continue
        shutil.copy2(src, dst)
        print(f"  + 3D: {fname}")
        stats["3d"] += 1; added = True
    return added

def read_queue(wb):
    ws, parts = wb[SHEET_QUEUE], []
    for row in ws.iter_rows(min_row=2, values_only=True):
        lcsc = str(row[0]).strip() if row[0] else ""
        if not lcsc or lcsc == "None": continue
        parts.append({"lcsc":lcsc,
                      "name":     str(row[1]).strip() if row[1] else "",
                      "category": normalize_category(str(row[2]).strip() if row[2] else ""),
                      "note":     str(row[3]).strip() if row[3] else ""})
    return parts

def clear_queue(wb):
    ws = wb[SHEET_QUEUE]
    for row in ws.iter_rows(min_row=2):
        for cell in row: cell.value = None

def write_back_to_queue(wb, parts):
    ws = wb[SHEET_QUEUE]
    for row_idx, part in enumerate(parts, 2):
        vals = [part["lcsc"],part["name"],part["category"],part["note"]]
        for col,val in enumerate(vals,1):
            cell = ws.cell(row=row_idx,column=col,value=val)
            cell.font=FONT_NORMAL; cell.border=THIN
            cell.alignment=ALIGN_CTR if col==1 else ALIGN_LEFT

# ==============================================================================
# CUSTOM SYMBOL MODE — doc CustomSym sheet + tao symbol
# ==============================================================================
def read_custom_sym(wb):
    if SHEET_CUSTOM not in wb.sheetnames:
        print(f"LOI: Khong tim thay sheet '{SHEET_CUSTOM}' trong {PARTS_FILE}")
        print("  Chay script mot lan de tao sheet: python lib.py")
        sys.exit(1)
    ws = wb[SHEET_CUSTOM]

    def v(r,c):
        val = ws.cell(row=r,column=c).value
        return str(val).strip() if val is not None else ""

    sym_name    = v(2,2)
    raw_cat     = v(3,2)
    ref         = v(4,2) or "U"
    description = v(5,2)

    if not sym_name:
        print("LOI: Chua dien Symbol Name (o B2)"); sys.exit(1)

    category = CATEGORY_MAP.get(raw_cat.lower(), raw_cat.title() if raw_cat else "Misc")
    pins = []
    for row in ws.iter_rows(min_row=9, values_only=True):
        num   = str(row[0]).strip() if row[0] is not None else ""
        name  = str(row[1]).strip() if row[1] is not None else ""
        ptype = str(row[2]).strip() if row[2] is not None else "passive"
        side  = str(row[3]).strip() if row[3] is not None else "L"
        if not num or not name or num=="None" or name=="None": continue
        if name.startswith("Ten chan") or name.startswith("VDD, GND"): continue
        pins.append({"number":num,"name":name,
                     "type":ptype.lower(),
                     "side":PIN_SIDE_MAP.get(side.lower(),"L")})

    if not pins:
        print("LOI: Chua co pin nao (dien tu row 9)"); sys.exit(1)

    return {"name":sym_name,"category":category,"ref":ref,"description":description,"pins":pins}

def generate_kicad_symbol(sym_name, ref, description, pins):
    L=[p for p in pins if p["side"]=="L"]
    R=[p for p in pins if p["side"]=="R"]
    T=[p for p in pins if p["side"]=="T"]
    B=[p for p in pins if p["side"]=="B"]

    n_lr   = max(len(L),len(R),1)
    n_tb   = max(len(T),len(B),0)
    body_h = n_lr*PIN_SPACING + PIN_SPACING
    body_w = max(BODY_MIN_W, n_tb*PIN_SPACING+PIN_SPACING if n_tb else BODY_MIN_W)
    hw = body_w/2; hh = body_h/2

    def yp(n):
        if not n: return []
        s=(n-1)*PIN_SPACING; return [s/2-i*PIN_SPACING for i in range(n)]
    def xp(n):
        if not n: return []
        s=(n-1)*PIN_SPACING; return [-s/2+i*PIN_SPACING for i in range(n)]

    f  = lambda v: f"{v:.4f}"
    FS = f"{FONT_SIZE:.4f}"; PL = f"{PIN_LENGTH:.4f}"

    def prop(name,val,x=0.0,y=0.0,hide=False):
        h=" (hide yes)" if hide else ""
        return (f'  (property "{name}" "{val}" (at {f(x)} {f(y)} 0)\n'
                f'    (effects (font (size {FS} {FS})){h})\n  )')

    def pin_str(p,x,y,angle):
        pt=PIN_TYPE_MAP.get(p["type"],"unspecified")
        pn=p["name"].replace('"','\\"')
        return (f'    (pin {pt} line (at {f(x)} {f(y)} {angle}) (length {PL})\n'
                f'      (name "{pn}" (effects (font (size {FS} {FS}))))\n'
                f'      (number "{p["number"]}" (effects (font (size {FS} {FS}))))\n    )')

    lines = [
        f'(symbol "{sym_name}"',
        f'  (pin_names (offset 1.016))',
        f'  (exclude_from_sim no)','  (in_bom yes)','  (on_board yes)',
        prop("Reference",ref,0,hh+1.27),
        prop("Value",sym_name,0,-(hh+1.27)),
        prop("Footprint","",hide=True),
        prop("Datasheet","",hide=True),
        prop("Description",description,hide=True),
        f'  (symbol "{sym_name}_0_1"',
        f'    (rectangle (start {f(-hw)} {f(hh)}) (end {f(hw)} {f(-hh)})',
        f'      (stroke (width 0) (type default))',
        f'      (fill (type background))',
        f'    )',
        f'  )',
        f'  (symbol "{sym_name}_1_1"',
    ]
    for p,y in zip(L,yp(len(L))): lines.append(pin_str(p,-(hw+PIN_LENGTH),y,0))
    for p,y in zip(R,yp(len(R))): lines.append(pin_str(p,hw+PIN_LENGTH,y,180))
    for p,x in zip(T,xp(len(T))): lines.append(pin_str(p,x,hh+PIN_LENGTH,270))
    for p,x in zip(B,xp(len(B))): lines.append(pin_str(p,x,-(hh+PIN_LENGTH),90))
    lines+=["  )",")"]
    return "\n".join(lines)

# ==============================================================================
# MODE 1: QUEUE — download tu LCSC
# ==============================================================================
def run_queue():
    ensure_directories()
    wb = load_workbook(PARTS_FILE)

    if SHEET_LIBRARY not in wb.sheetnames or SHEET_QUEUE not in wb.sheetnames:
        print(f"Thieu sheet Library hoac Queue trong {PARTS_FILE}"); sys.exit(1)

    queue = read_queue(wb)
    if not queue:
        print("Queue trong — khong co gi can import."); return

    existing_names, existing_lcsc_map = load_all_existing()
    existing_lcsc_lib = get_library_lcsc(wb)

    already_have, to_download = [], []
    for part in queue:
        lcsc   = part["lcsc"]
        lcsc_c = lcsc if lcsc.upper().startswith("C") else f"C{lcsc}"
        info   = existing_lcsc_map.get(lcsc_c) or existing_lcsc_map.get(lcsc)
        if info:
            if not part["name"]:    part["name"]     = info["name"]
            if not part["note"]:    part["note"]      = info["desc"]
            if part["category"] == DEFAULT_CAT: part["category"] = info["category"]
            already_have.append(part)
        elif lcsc_c in existing_lcsc_lib or lcsc in existing_lcsc_lib:
            already_have.append(part)
        else:
            to_download.append(part)

    print(f"\nQueue hien tai  : {len(queue)} parts")
    print(f"Da co trong lib : {len(already_have)} - chuyen sang Library")
    print(f"Can tai moi     : {len(to_download)}\n")

    stats           = {"sym":0,"fp":0,"3d":0,"success":0,"skip":0,"err":0}
    processed_parts = []
    failed_parts    = []

    # Tim easyeda2kicad
    scripts_dir  = os.path.dirname(sys.executable)
    easyeda_name = "easyeda2kicad.exe" if os.name=="nt" else "easyeda2kicad"
    easyeda_cmd  = os.path.join(scripts_dir, easyeda_name)
    if not os.path.isfile(easyeda_cmd):
        found = shutil.which("easyeda2kicad")
        if found: easyeda_cmd = found
        else:
            print(f"LOI: Khong tim thay easyeda2kicad")
            print(f"  Chay: pip install easyeda2kicad")
            sys.exit(1)
    print(f"easyeda2kicad: {easyeda_cmd}\n")

    for i, part in enumerate(to_download, 1):
        lcsc_id  = part["lcsc"]
        note     = part["note"]
        category = part["category"]
        sym_file = sym_file_for(category)

        ensure_sym_file(sym_file)
        clear_temp()

        label = f" - {note}" if note else ""
        print(f"[{i:2}/{len(to_download)}] [{category}] {lcsc_id}{label}")

        result = subprocess.run(
            [easyeda_cmd,"--full",f"--lcsc_id={lcsc_id}",
             "--output",os.path.join("temp","temp")],
            capture_output=True, text=True
        )

        if result.returncode != 0:
            if result.stderr.strip(): print(f"  stderr: {result.stderr.strip()[:300]}")
            print(f"  FAIL - giu lai Queue de retry")
            stats["err"] += 1; failed_parts.append(part); print(); continue

        temp_sym = os.path.join(TEMP_DIR,"temp.kicad_sym")
        blocks   = extract_symbol_blocks(temp_sym)

        if not blocks:
            print(f"  FAIL - khong co symbol hop le")
            stats["err"] += 1; failed_parts.append(part); print(); continue

        sym_ok = False
        for block in blocks:
            bname=block["name"]; blcsc=block["lcsc"]; bcont=block["content"]
            if bname in existing_names:
                print(f"  skip '{bname}' da ton tai"); stats["skip"]+=1; sym_ok=True; continue
            if blcsc and blcsc in existing_lcsc_map:
                print(f"  skip LCSC {blcsc} da ton tai"); stats["skip"]+=1; sym_ok=True; continue
            if note:
                safe=note.replace('"','\\"')
                if '(property "Description"' in bcont:
                    bcont=re.sub(r'\(property\s+"Description"\s+"[^"]*"',
                                 f'(property "Description" "{safe}"',bcont,count=1)
                else:
                    bcont=re.sub(r'(\(symbol\s+"[^"]+"\n)',
                                 f'\\1  (property "Description" "{safe}" (at 0 0 0)'
                                 f' (effects (font (size 1.27 1.27)) (hide yes)))\n',
                                 bcont,count=1)
            merge_symbol(bcont, sym_file)
            existing_names.add(bname)
            if blcsc: existing_lcsc_map[blcsc]={"name":bname,"desc":note,"category":category}
            print(f"  + Symbol: {bname} -> {LIB_PREFIX}-{category}")
            stats["sym"]+=1; sym_ok=True

        fp_ok    = process_footprints(stats)
        model_ok = process_3d_models(stats)
        processed_parts.append((part, sym_ok, fp_ok, model_ok))
        stats["success"]+=1; print()

    clear_temp()

    # Cap nhat Excel
    for part, sym_ok, fp_ok, model_ok in processed_parts:
        add_to_library(wb, part["lcsc"], part["name"], part["category"],
                       part["note"], sym_ok, fp_ok, model_ok)

    for part in already_have:
        add_to_library(wb, part["lcsc"], part["name"], part["category"],
                       part["note"], True, True, True)

    # Dong bo o trong trong Library tu KiCad
    ws_lib = wb[SHEET_LIBRARY]
    for row_idx in range(2, ws_lib.max_row+1):
        lcsc_val = str(ws_lib.cell(row=row_idx,column=1).value or "").strip()
        info = existing_lcsc_map.get(lcsc_val)
        if info:
            n_cell = ws_lib.cell(row=row_idx,column=2)
            if not n_cell.value: n_cell.value = info["name"]
            c_cell = ws_lib.cell(row=row_idx,column=3)
            if not c_cell.value: c_cell.value = info["category"]
            d_cell = ws_lib.cell(row=row_idx,column=4)
            if not d_cell.value: d_cell.value = info["desc"]

    # Phuc hoi linh kien trong KiCad nhung khong co trong Library sheet
    current_lib = get_library_lcsc(wb)
    recovered   = 0
    for lcsc_k, info in existing_lcsc_map.items():
        if lcsc_k and lcsc_k not in current_lib:
            add_to_library(wb, lcsc_k, info["name"], info["category"], info["desc"], True,True,True)
            current_lib.add(lcsc_k); recovered+=1
    if recovered:
        print(f"  + Phuc hoi {recovered} linh kien tu KiCad vao Library sheet")

    clear_queue(wb)
    write_back_to_queue(wb, failed_parts)

    # Reset view
    for sheet in wb.worksheets:
        if sheet.sheet_view.selection:
            sheet.sheet_view.selection[0].activeCell="A1"
            sheet.sheet_view.selection[0].sqref="A1"
        sheet.sheet_view.topLeftCell="A1"
    if SHEET_QUEUE in wb.sheetnames:
        wb.active = wb.sheetnames.index(SHEET_QUEUE)

    wb.save(PARTS_FILE)

    moved = len(processed_parts)+len(already_have)+recovered
    print(f"Cap nhat {PARTS_FILE}: {moved} parts -> Library sheet")
    if failed_parts:
        print(f"  ! {len(failed_parts)} parts loi -> giu lai Queue")

    sym_files = sorted(f for f in os.listdir(SYMBOLS_DIR)
                       if f.startswith(f"{LIB_PREFIX}-") and f.endswith(".kicad_sym"))
    if sym_files:
        print("\nFiles thu vien:")
        for fname in sym_files:
            blocks   = extract_symbol_blocks(os.path.join(SYMBOLS_DIR,fname))
            nickname = fname.replace(".kicad_sym","")
            print(f"  {nickname:<28} ({len(blocks)} symbols)")

    print(f"\n{'─'*32}")
    print(f"Symbol moi    : {stats['sym']}")
    print(f"Footprint moi : {stats['fp']}")
    print(f"3D model moi  : {stats['3d']}")
    print(f"Thanh cong    : {stats['success']}")
    print(f"Bo qua (trung): {stats['skip']+len(already_have)}")
    print(f"Loi           : {stats['err']}")
    if stats["success"]>0:
        print("\n  git add .\n  git commit -m \"Update library\"\n  git push")

    open_excel_file(PARTS_FILE)

# ==============================================================================
# MODE 2: CUSTOM — tao symbol thu cong tu sheet CustomSym
# ==============================================================================
def run_custom():
    ensure_directories()
    wb = load_workbook(PARTS_FILE)

    if SHEET_LIBRARY not in wb.sheetnames:
        print(f"Thieu sheet Library trong {PARTS_FILE}"); sys.exit(1)

    sym_def  = read_custom_sym(wb)
    sym_name = sym_def["name"]
    category = sym_def["category"]
    pins     = sym_def["pins"]
    sym_file = sym_file_for(category)

    print(f"\nSymbol  : {sym_name}")
    print(f"Category: {category}  |  Ref: {sym_def['ref']}")
    print(f"File    : {sym_file}")
    print(f"Pins    : {len(pins)}")

    L=[p for p in pins if p["side"]=="L"]
    R=[p for p in pins if p["side"]=="R"]
    T=[p for p in pins if p["side"]=="T"]
    B=[p for p in pins if p["side"]=="B"]
    if L: print(f"  Left  ({len(L)}): "+", ".join(f"{p['name']}({p['number']})" for p in L))
    if R: print(f"  Right ({len(R)}): "+", ".join(f"{p['name']}({p['number']})" for p in R))
    if T: print(f"  Top   ({len(T)}): "+", ".join(f"{p['name']}({p['number']})" for p in T))
    if B: print(f"  Bot   ({len(B)}): "+", ".join(f"{p['name']}({p['number']})" for p in B))

    ensure_sym_file(sym_file)
    if symbol_exists(sym_file, sym_name):
        print(f"\nLOI: '{sym_name}' da ton tai trong {sym_file}")
        print("  Doi ten (B2) hoac xoa symbol cu truoc."); sys.exit(1)

    block = generate_kicad_symbol(sym_name, sym_def["ref"], sym_def["description"], pins)
    merge_symbol(block, sym_file)
    print(f"\n  + Symbol '{sym_name}' -> {sym_file}")

    add_to_library(wb, sym_name, sym_name, category, sym_def["description"], True, False, False)
    wb.save(PARTS_FILE)
    print(f"  + Library sheet updated")
    print(f"\nXong! Mo KiCad Symbol Editor de kiem tra va them footprint.")
    open_excel_file(PARTS_FILE)

# ==============================================================================
# ENTRY POINT
# ==============================================================================
def main():
    mode = sys.argv[1].lower() if len(sys.argv) > 1 else "queue"

    if mode in ("custom", "sym", "symbol"):
        print("="*50)
        print("MODE: Custom Symbol Creator")
        print("="*50)
        check_file_not_open(PARTS_FILE)
        if not os.path.exists(PARTS_FILE):
            print(f"Khong tim thay {PARTS_FILE}"); sys.exit(1)
        run_custom()

    elif mode in ("queue", "generate", "gen"):
        print("="*50)
        print("MODE: Queue Download (LCSC -> Library)")
        print("="*50)
        check_file_not_open(PARTS_FILE)
        if not os.path.exists(PARTS_FILE):
            print(f"Khong tim thay {PARTS_FILE}"); sys.exit(1)
        run_queue()

    else:
        print("Cach dung:")
        print("  python lib.py          -> Download tu Queue")
        print("  python lib.py custom   -> Tao symbol thu cong")

if __name__ == "__main__":
    main()
