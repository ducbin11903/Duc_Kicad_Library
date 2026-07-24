#!/usr/bin/env python3
"""
create_symbol.py
Tao KiCad symbol tu sheet CustomSym trong CustomSym.xlsx
Doc pin -> generate .kicad_sym -> add vao Library sheet trong parts.xlsx
"""

import os, sys
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

CUSTOM_FILE   = "parts.xlsx"
PARTS_FILE    = "parts.xlsx"
SHEET_CUSTOM  = "CustomSym"
SHEET_LIBRARY = "Library"
LIB_PREFIX    = "Duc"
SYMBOLS_DIR   = "symbols"
PIN_LENGTH    = 2.54
PIN_SPACING   = 2.54
BODY_MIN_W    = 10.16
FONT_SIZE     = 1.27

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
CATEGORY_MAP = {
    "resistors":"Resistors","resistor":"Resistors","res":"Resistors",
    "capacitors":"Capacitors","capacitor":"Capacitors","cap":"Capacitors",
    "inductors":"Inductors","inductor":"Inductors",
    "diodes":"Diodes","diode":"Diodes",
    "transistors":"Transistors","transistor":"Transistors","mosfet":"Transistors",
    "mcus":"MCUs","mcu":"MCUs","ic":"ICs","ics":"ICs","logic":"ICs",
    "power":"Power","ldo":"Power","dcdc":"Power",
    "sensor":"Sensors","sensors":"Sensors",
    "connectors":"Connectors","connector":"Connectors","conn":"Connectors",
    "crystals":"Crystals","crystal":"Crystals",
    "memory":"Memory","rf":"RF","analog":"Analog","interface":"Interface",
    "optocouplers":"Optocouplers","optocoupler":"Optocouplers",
    "misc":"Misc","other":"Misc",
}

THIN  = Border(left=Side(style="thin",color="D0D0D0"),right=Side(style="thin",color="D0D0D0"),
               top=Side(style="thin",color="D0D0D0"),bottom=Side(style="thin",color="D0D0D0"))
F_N   = Font(name="Arial",size=10)
A_C   = Alignment(horizontal="center",vertical="center")
A_L   = Alignment(horizontal="left",vertical="center")
F_OK  = PatternFill("solid",fgColor="E2EFDA")
F_NO  = PatternFill("solid",fgColor="FFDDE0")
F_GR  = PatternFill("solid",fgColor="F2F2F2")

def read_custom_sym():
    if not os.path.exists(CUSTOM_FILE):
        print(f"Khong tim thay {CUSTOM_FILE}"); sys.exit(1)
    wb = load_workbook(CUSTOM_FILE)
    if SHEET_CUSTOM not in wb.sheetnames:
        print(f"Khong tim thay sheet '{SHEET_CUSTOM}'"); sys.exit(1)
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
        if name.startswith("Ten chan"): continue
        pins.append({"number":num,"name":name,
                     "type":ptype.lower(),
                     "side":PIN_SIDE_MAP.get(side.lower(),"L")})

    if not pins:
        print("LOI: Chua co pin nao (dien tu row 9)"); sys.exit(1)

    return {"name":sym_name,"category":category,
            "ref":ref,"description":description,"pins":pins}

def generate_kicad_symbol(sym_name, ref, description, pins):
    L = [p for p in pins if p["side"]=="L"]
    R = [p for p in pins if p["side"]=="R"]
    T = [p for p in pins if p["side"]=="T"]
    B = [p for p in pins if p["side"]=="B"]

    n_lr   = max(len(L),len(R),1)
    n_tb   = max(len(T),len(B),0)
    body_h = n_lr*PIN_SPACING + PIN_SPACING
    body_w = max(BODY_MIN_W, n_tb*PIN_SPACING+PIN_SPACING if n_tb else BODY_MIN_W)
    hw = body_w/2;  hh = body_h/2

    def yp(n):
        if not n: return []
        s=(n-1)*PIN_SPACING; return [s/2-i*PIN_SPACING for i in range(n)]
    def xp(n):
        if not n: return []
        s=(n-1)*PIN_SPACING; return [-s/2+i*PIN_SPACING for i in range(n)]

    f  = lambda v: f"{v:.4f}"
    FS = f"{FONT_SIZE:.4f}";  PL = f"{PIN_LENGTH:.4f}"

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
        prop("Reference",   ref,          0,  hh+1.27),
        prop("Value",       sym_name,     0,-(hh+1.27)),
        prop("Footprint",   "",hide=True),
        prop("Datasheet",   "",hide=True),
        prop("Description", description,hide=True),
        f'  (symbol "{sym_name}_0_1"',
        f'    (rectangle (start {f(-hw)} {f(hh)}) (end {f(hw)} {f(-hh)})',
        f'      (stroke (width 0) (type default))',
        f'      (fill (type background))',
        f'    )',
        f'  )',
        f'  (symbol "{sym_name}_1_1"',
    ]
    for p,y in zip(L,yp(len(L))): lines.append(pin_str(p,-(hw+PIN_LENGTH),y,0))
    for p,y in zip(R,yp(len(R))): lines.append(pin_str(p, hw+PIN_LENGTH,  y,180))
    for p,x in zip(T,xp(len(T))): lines.append(pin_str(p,x, hh+PIN_LENGTH,270))
    for p,x in zip(B,xp(len(B))): lines.append(pin_str(p,x,-(hh+PIN_LENGTH),90))
    lines+=["  )",")"]
    return "\n".join(lines)

def sym_file_path(category):
    return os.path.join(SYMBOLS_DIR,f"{LIB_PREFIX}-{category}.kicad_sym")

def ensure_sym_file(path):
    os.makedirs(SYMBOLS_DIR,exist_ok=True)
    if not os.path.exists(path):
        with open(path,"w",encoding="utf-8") as f:
            f.write("(kicad_symbol_lib (version 20231120) (generator duc_lib)\n)\n")

def symbol_exists(path,name):
    if not os.path.exists(path): return False
    with open(path,"r",encoding="utf-8") as f:
        return f'(symbol "{name}"' in f.read()

def write_symbol(block,path):
    with open(path,"r",encoding="utf-8") as f: content=f.read()
    last=content.rfind(")")
    if last!=-1:
        with open(path,"w",encoding="utf-8") as f:
            f.write(content[:last]+"\n  "+block+"\n"+content[last:])

def add_to_library(sym_def):
    if not os.path.exists(PARTS_FILE):
        print(f"  Khong tim thay {PARTS_FILE} — bo qua Library sheet"); return
    wb=load_workbook(PARTS_FILE)
    if SHEET_LIBRARY not in wb.sheetnames:
        print("  Khong tim thay sheet Library — bo qua"); return
    ws=wb[SHEET_LIBRARY]
    nrow=ws.max_row+1
    now=datetime.now().strftime("%Y-%m-%d %H:%M")
    vals=[sym_def["name"],sym_def["name"],sym_def["category"],
          sym_def["description"],"v","x","x",now]
    for col,val in enumerate(vals,1):
        cell=ws.cell(row=nrow,column=col,value=val)
        cell.font=F_N; cell.border=THIN
        if col in(5,6,7):
            cell.alignment=A_C; cell.fill=F_OK if val=="v" else F_NO
        elif col==8:
            cell.alignment=A_C; cell.fill=F_GR
        else:
            cell.alignment=A_L; cell.fill=F_GR
    wb.save(PARTS_FILE)
    print(f"  + Library sheet updated ({PARTS_FILE})")

def main():
    print("="*50+"\nKiCad Symbol Creator\n"+"="*50)
    sym_def  = read_custom_sym()
    sym_name = sym_def["name"]
    category = sym_def["category"]
    pins     = sym_def["pins"]
    sym_file = sym_file_path(category)

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
    if symbol_exists(sym_file,sym_name):
        print(f"\nLOI: '{sym_name}' da ton tai trong {sym_file}")
        print("Doi ten (B2) hoac xoa symbol cu truoc."); sys.exit(1)

    block=generate_kicad_symbol(sym_name,sym_def["ref"],sym_def["description"],pins)
    write_symbol(block,sym_file)
    print(f"\n  + Symbol '{sym_name}' -> {sym_file}")
    add_to_library(sym_def)
    print(f"\nXong! Mo KiCad Symbol Editor de kiem tra va them footprint.")

if __name__=="__main__":
    main()
