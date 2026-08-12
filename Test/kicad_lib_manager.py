#!/usr/bin/env python3
"""
kicad_lib_manager.py — Duc KiCad Library Manager
Build EXE: pip install pyinstaller
           pyinstaller --onefile --windowed --name KiCad_Lib_Manager kicad_lib_manager.py
"""
import os, sys, re, subprocess, shutil, threading, queue as TQ, platform
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Constants ─────────────────────────────────────────────────────────────────
APP_TITLE   = "Duc KiCad Library Manager"
APP_VER     = "2.0"
PARTS_FILE  = "parts.xlsx"
SHEET_LIB   = "Library"
SHEET_Q     = "Queue"
SHEET_CUS   = "CustomSym"
LIB_PFX     = "Duc"
SYM_DIR     = "symbols"
FP_DIR      = os.path.join("footprints", f"{LIB_PFX}.pretty")
M3D_DIR     = os.path.join("3dmodels",   f"{LIB_PFX}.3dshapes")
TEMP_DIR    = "temp"
DEF_CAT     = "Misc"

PIN_LEN   = 2.54; PIN_SPC = 2.54; BODY_W = 10.16; FSYM = 1.27

CATS = ["Resistors","Capacitors","Inductors","Diodes","Transistors","MCUs","ICs",
        "Power","Sensors","Connectors","Crystals","Memory","RF","Analog",
        "Interface","Optocouplers","Misc"]
PTYPES = ["IN","OUT","IO","PWR","GND","PWR_OUT","PASSIVE","NC","OC","OE","UNSPEC","TRI","CLK"]
PSIDES = ["L","R","T","B"]
REFS   = ["U","R","C","L","Q","J","D","Y","SW","F","BT","TP","M","P"]

CATMAP = {
    "resistors":"Resistors","res":"Resistors","r":"Resistors",
    "capacitors":"Capacitors","cap":"Capacitors","c":"Capacitors",
    "inductors":"Inductors","ind":"Inductors","l":"Inductors",
    "diodes":"Diodes","diode":"Diodes","d":"Diodes",
    "transistors":"Transistors","mosfet":"Transistors","q":"Transistors",
    "mcus":"MCUs","mcu":"MCUs","ic":"ICs","ics":"ICs","logic":"ICs",
    "power":"Power","ldo":"Power","dcdc":"Power",
    "sensor":"Sensors","sensors":"Sensors",
    "connectors":"Connectors","connector":"Connectors","conn":"Connectors",
    "crystals":"Crystals","crystal":"Crystals",
    "memory":"Memory","rf":"RF","analog":"Analog","interface":"Interface",
    "optocouplers":"Optocouplers","misc":"Misc","other":"Misc",
}
PTMAP = {
    "in":"input","out":"output","io":"bidirectional","bidir":"bidirectional",
    "pwr":"power_in","power":"power_in","vdd":"power_in","vcc":"power_in","gnd":"power_in",
    "pwr_out":"power_out","oc":"open_collector","oe":"open_emitter",
    "nc":"no_connect","passive":"passive","p":"passive",
    "unspec":"unspecified","tri":"tri_state","clk":"clock","clock":"clock",
}

# Colors
CB="#1F4E79"; CO="#C55A11"; CG="#375623"; CGRAY="#F5F5F5"
CW="#FFFFFF"; COK="#2E7D32"; CERR="#C62828"; CWARN="#E65100"
CDARK="#263238"

def norm_cat(raw):
    return CATMAP.get(raw.strip().lower(), raw.strip().title() if raw.strip() else DEF_CAT)
def sym_path(cat): return os.path.join(SYM_DIR,f"{LIB_PFX}-{cat}.kicad_sym")
def find_easyeda():
    sc=os.path.dirname(sys.executable)
    n="easyeda2kicad.exe" if os.name=="nt" else "easyeda2kicad"
    c=os.path.join(sc,n)
    return c if os.path.isfile(c) else shutil.which("easyeda2kicad")

# ── Excel helpers ─────────────────────────────────────────────────────────────
THIN=Border(left=Side(style="thin",color="D0D0D0"),right=Side(style="thin",color="D0D0D0"),
            top=Side(style="thin",color="D0D0D0"),bottom=Side(style="thin",color="D0D0D0"))
FN=Font(name="Arial",size=10); AC=Alignment(horizontal="center",vertical="center")
AL=Alignment(horizontal="left",vertical="center")
FGR=PatternFill("solid",fgColor="F2F2F2"); FOK=PatternFill("solid",fgColor="E2EFDA")
FNO=PatternFill("solid",fgColor="FFDDE0")

def excel_add_lib_row(wb, lcsc, name, cat, note, s, f, m):
    ws=wb[SHEET_LIB]; nr=ws.max_row+1
    now=datetime.now().strftime("%Y-%m-%d %H:%M")
    for col,val in enumerate([lcsc,name,cat,note,"v"if s else"x","v"if f else"x","v"if m else"x",now],1):
        cell=ws.cell(row=nr,column=col,value=val); cell.font=FN; cell.border=THIN
        if col in(5,6,7): cell.alignment=AC; cell.fill=FOK if val=="v" else FNO
        elif col==8: cell.alignment=AC; cell.fill=FGR
        else: cell.alignment=AL; cell.fill=FGR

def save_queue_to_excel(rows):
    if not os.path.exists(PARTS_FILE): return
    try:
        wb=load_workbook(PARTS_FILE)
        if SHEET_Q not in wb.sheetnames: return
        ws=wb[SHEET_Q]
        for r in ws.iter_rows(min_row=2):
            for c in r: c.value=None
        for ri,(lcsc,name,cat,note) in enumerate(rows,2):
            for ci,v in enumerate([lcsc,name,cat,note],1):
                cell=ws.cell(row=ri,column=ci,value=v); cell.font=FN
                cell.border=THIN; cell.alignment=AC if ci==1 else AL
        wb.save(PARTS_FILE)
    except Exception: pass

def load_queue_from_excel():
    rows=[]
    if not os.path.exists(PARTS_FILE): return rows
    try:
        wb=load_workbook(PARTS_FILE)
        if SHEET_Q in wb.sheetnames:
            ws=wb[SHEET_Q]
            for row in ws.iter_rows(min_row=2,values_only=True):
                lcsc=str(row[0]).strip() if row[0] else ""
                if not lcsc or lcsc=="None": continue
                rows.append((lcsc,
                             str(row[1]).strip() if row[1] else "",
                             norm_cat(str(row[2]).strip() if row[2] else ""),
                             str(row[3]).strip() if row[3] else ""))
    except Exception: pass
    return rows

# ── KiCad helpers ─────────────────────────────────────────────────────────────
def ensure_dirs():
    for d in [SYM_DIR,FP_DIR,M3D_DIR,TEMP_DIR]: os.makedirs(d,exist_ok=True)
def ensure_sym(path):
    if not os.path.exists(path):
        with open(path,"w",encoding="utf-8") as f:
            f.write("(kicad_symbol_lib (version 20231120) (generator duc_lib)\n)\n")
def clear_temp():
    if os.path.exists(TEMP_DIR): shutil.rmtree(TEMP_DIR,ignore_errors=True)
    os.makedirs(TEMP_DIR,exist_ok=True)

def extract_blocks(filepath):
    if not os.path.exists(filepath): return []
    with open(filepath,"r",encoding="utf-8") as f: content=f.read()
    blocks,pos=[],0
    while True:
        start=content.find('(symbol "',pos)
        if start==-1: break
        stack,end=0,-1
        for i in range(start,len(content)):
            if content[i]=="(": stack+=1
            elif content[i]==")":
                stack-=1
                if stack==0: end=i+1; break
        if end==-1: break
        b=content[start:end]; m=re.match(r'\(symbol\s+"([^"]+)"',b)
        if m:
            lm=re.search(r'\(property\s+"LCSC Part"\s+"([^"]+)"',b)
            blocks.append({"name":m.group(1),"lcsc":lm.group(1) if lm else None,"content":b})
        pos=end
    return blocks

def merge_block(block,sym_file):
    with open(sym_file,"r",encoding="utf-8") as f: c=f.read()
    block=re.sub(r'\(property\s+"Footprint"\s+"[^:]+:([^"]+)"',
                 rf'(property "Footprint" "{LIB_PFX}:\1"',block)
    last=c.rfind(")")
    if last!=-1:
        with open(sym_file,"w",encoding="utf-8") as f:
            f.write(c[:last]+"\n  "+block+"\n"+c[last:])

def sym_exists(sf,name):
    if not os.path.exists(sf): return False
    with open(sf,"r",encoding="utf-8") as f: return f'(symbol "{name}"' in f.read()

def gen_kicad_sym(sym_name,ref,desc,pins):
    L=[p for p in pins if p["side"]=="L"]; R=[p for p in pins if p["side"]=="R"]
    T=[p for p in pins if p["side"]=="T"]; B=[p for p in pins if p["side"]=="B"]
    n=max(len(L),len(R),1); nt=max(len(T),len(B),0)
    bh=n*PIN_SPC+PIN_SPC; bw=max(BODY_W,nt*PIN_SPC+PIN_SPC if nt else BODY_W)
    hw=bw/2; hh=bh/2
    def yp(n): return [((n-1)*PIN_SPC/2)-i*PIN_SPC for i in range(n)] if n else []
    def xp(n): return [-(n-1)*PIN_SPC/2+i*PIN_SPC for i in range(n)] if n else []
    f=lambda v:f"{v:.4f}"; FS=f"{FSYM:.4f}"; PL=f"{PIN_LEN:.4f}"
    def pr(nm,val,x=0.,y=0.,hide=False):
        h=" (hide yes)" if hide else ""
        return f'  (property "{nm}" "{val}" (at {f(x)} {f(y)} 0)\n    (effects (font (size {FS} {FS})){h})\n  )'
    def ps(p,x,y,a):
        pt=PTMAP.get(p["type"],"unspecified"); pn=p["name"].replace('"','\\"')
        return(f'    (pin {pt} line (at {f(x)} {f(y)} {a}) (length {PL})\n'
               f'      (name "{pn}" (effects (font (size {FS} {FS}))))\n'
               f'      (number "{p["number"]}" (effects (font (size {FS} {FS}))))\n    )')
    lines=[f'(symbol "{sym_name}"','  (pin_names (offset 1.016))',
           '  (exclude_from_sim no)','  (in_bom yes)','  (on_board yes)',
           pr("Reference",ref,0,hh+1.27),pr("Value",sym_name,0,-(hh+1.27)),
           pr("Footprint","",hide=True),pr("Datasheet","",hide=True),pr("Description",desc,hide=True),
           f'  (symbol "{sym_name}_0_1"',
           f'    (rectangle (start {f(-hw)} {f(hh)}) (end {f(hw)} {f(-hh)})',
           f'      (stroke (width 0) (type default))','      (fill (type background))','    )','  )',
           f'  (symbol "{sym_name}_1_1"']
    for p,y in zip(L,yp(len(L))): lines.append(ps(p,-(hw+PIN_LEN),y,0))
    for p,y in zip(R,yp(len(R))): lines.append(ps(p,hw+PIN_LEN,y,180))
    for p,x in zip(T,xp(len(T))): lines.append(ps(p,x,hh+PIN_LEN,270))
    for p,x in zip(B,xp(len(B))): lines.append(ps(p,x,-(hh+PIN_LEN),90))
    lines+=["  ",")"]; return "\n".join(lines)

# ==============================================================================
# EDITABLE TREEVIEW (click-to-edit, Ctrl+V paste, context menu)
# ==============================================================================
class EditableTV(ttk.Treeview):
    """
    Treeview co kha nang:
    - Click vao o de edit truc tiep
    - Ctrl+V paste tu Excel (tab-separated) hoac text (comma/newline)
    - Tab / Enter de di chuyen giua cac o
    - Delete de xoa hang
    - Chuot phai: context menu
    """
    def __init__(self, parent, cols, hdrs, widths, combos=None, **kw):
        super().__init__(parent, columns=cols, show="headings",
                         selectmode="extended", **kw)
        self._cols   = cols
        self._combos = combos or {}   # col_name -> list of values
        self._edit_w = None
        self._edit_item = None
        self._edit_col  = None

        for col,hdr,w in zip(cols,hdrs,widths):
            self.heading(col,text=hdr,command=lambda c=col:self._sort(c))
            self.column(col,width=w,minwidth=40)

        # Alternating row colors
        self.tag_configure("even",background="#FAFAFA")
        self.tag_configure("odd", background=CGRAY)

        self.bind("<Button-1>",   self._on_click)
        self.bind("<Delete>",     self._on_del_key)
        self.bind("<Control-v>",  self._on_paste)
        self.bind("<Control-V>",  self._on_paste)
        self.bind("<Control-c>",  self._on_copy)
        self.bind("<Control-C>",  self._on_copy)
        self.bind("<Button-3>",   self._on_rclick)
        self.bind("<Escape>",     lambda e: self._cancel())

    # ── Editing ───────────────────────────────────────────────────────────────
    def _on_click(self, event):
        self._cancel()
        region=self.identify_region(event.x,event.y)
        if region!="cell": return
        item=self.identify_row(event.y)
        col=self.identify_column(event.x)
        if not item or not col: return
        col_idx=int(col.replace("#",""))-1
        if col_idx<0 or col_idx>=len(self._cols): return
        self._start_edit(item,col_idx)

    def _start_edit(self, item, col_idx):
        col=self._cols[col_idx]
        bbox=self.bbox(item,f"#{col_idx+1}")
        if not bbox: return
        x,y,w,h=bbox
        cur_val=self.set(item,col)
        self._edit_item=item; self._edit_col=col_idx

        if col in self._combos:
            self._edit_w=ttk.Combobox(self,values=self._combos[col],
                                       state="readonly",font=("Arial",9))
            self._edit_w.set(cur_val)
            self._edit_w.place(x=x,y=y,width=w,height=h)
            self._edit_w.focus_set()
            self._edit_w.bind("<<ComboboxSelected>>",lambda e:self._commit())
            self._edit_w.bind("<Escape>",lambda e:self._cancel())
            self._edit_w.bind("<Tab>",lambda e:self._move(1))
            self._edit_w.bind("<Return>",lambda e:self._commit())
        else:
            self._edit_w=tk.Entry(self,font=("Arial",9),
                                  bd=1,relief="solid",highlightthickness=1,
                                  highlightcolor=CB)
            self._edit_w.insert(0,cur_val)
            self._edit_w.select_range(0,"end")
            self._edit_w.place(x=x,y=y,width=w,height=h)
            self._edit_w.focus_set()
            self._edit_w.bind("<Return>",lambda e:self._move_row(1))
            self._edit_w.bind("<Tab>",lambda e:self._move(1))
            self._edit_w.bind("<Shift-Tab>",lambda e:self._move(-1))
            self._edit_w.bind("<Escape>",lambda e:self._cancel())
            self._edit_w.bind("<FocusOut>",lambda e:self._commit())

    def _commit(self):
        if self._edit_w and self._edit_item and self._edit_col is not None:
            col=self._cols[self._edit_col]
            self.set(self._edit_item,col,self._edit_w.get())
            self._renumber()
        self._cancel()

    def _cancel(self):
        if self._edit_w:
            try: self._edit_w.destroy()
            except: pass
            self._edit_w=None
        self._edit_item=None; self._edit_col=None

    def _move(self, direction):
        """Di chuyen sang cot ke tiep / truoc."""
        self._commit()
        if self._edit_item is None: return
        next_col=(self._edit_col+direction) % len(self._cols)
        self._start_edit(self._edit_item,next_col)

    def _move_row(self, direction):
        """Di chuyen xuong/len hang ke tiep (Enter)."""
        self._commit()
        if self._edit_item is None: return
        items=self.get_children()
        if not items: return
        idx=items.index(self._edit_item) if self._edit_item in items else 0
        next_idx=(idx+direction) % len(items)
        self._start_edit(items[next_idx], self._edit_col or 0)

    # ── Paste ────────────────────────────────────────────────────────────────
    def _on_paste(self, event):
        self._cancel()
        try: clip=self.clipboard_get()
        except: return
        lines=[l for l in clip.strip().split("\n") if l.strip()]
        for line in lines:
            # Tab-separated (from Excel) atau comma
            if "\t" in line:
                cells=[c.strip() for c in line.split("\t")]
            else:
                cells=[c.strip() for c in line.split(",")]
            # Pad/trim to column count
            while len(cells)<len(self._cols): cells.append("")
            cells=cells[:len(self._cols)]
            # Normalize category if col exists
            if len(self._cols)>2 and len(cells)>2:
                cells[2]=norm_cat(cells[2]) if cells[2] else DEF_CAT
            n=len(self.get_children())
            tag="even" if n%2==0 else "odd"
            self.insert("","end",values=cells,tags=(tag,))
        self._renumber()
        return "break"

    def _on_copy(self, event):
        sel=self.selection()
        if not sel: return
        rows=["\t".join(str(self.set(i,c)) for c in self._cols) for i in sel]
        self.clipboard_clear(); self.clipboard_append("\n".join(rows))

    # ── Delete ───────────────────────────────────────────────────────────────
    def _on_del_key(self, event):
        for item in self.selection(): self.delete(item)
        self._renumber()

    # ── Context menu ─────────────────────────────────────────────────────────
    def _on_rclick(self, event):
        self._cancel()
        item=self.identify_row(event.y)
        if item: self.selection_set(item)
        menu=tk.Menu(self,tearoff=0)
        menu.add_command(label="Add row (Enter)",command=self.add_empty_row)
        menu.add_command(label="Delete selected (Del)",command=lambda:[self.delete(i) for i in self.selection()])
        menu.add_separator()
        menu.add_command(label="Move up",  command=lambda:self._move_sel(-1))
        menu.add_command(label="Move down",command=lambda:self._move_sel(1))
        menu.add_separator()
        menu.add_command(label="Clear all",command=self.clear_all)
        menu.tk_popup(event.x_root,event.y_root)

    def _move_sel(self, direction):
        sel=self.selection()
        if not sel: return
        item=sel[0]; items=list(self.get_children())
        idx=items.index(item)
        new_idx=max(0,min(len(items)-1,idx+direction))
        ref=items[new_idx]
        if direction>0: self.move(item,"",new_idx)
        else: self.move(item,"",new_idx)
        self._renumber()

    # ── Sort ─────────────────────────────────────────────────────────────────
    def _sort(self, col):
        items=[(self.set(i,col),i) for i in self.get_children()]
        items.sort(); 
        for idx,(val,item) in enumerate(items):
            self.move(item,"",idx)
            self.item(item,tags=("even" if idx%2==0 else "odd",))

    # ── Public helpers ────────────────────────────────────────────────────────
    def add_empty_row(self, defaults=None):
        n=len(self.get_children()); tag="even" if n%2==0 else "odd"
        vals=defaults or [""]*len(self._cols)
        self.insert("","end",values=vals,tags=(tag,))
        self._renumber()

    def clear_all(self):
        self.delete(*self.get_children())

    def get_all_rows(self):
        return [tuple(self.set(i,c) for c in self._cols) for i in self.get_children()]

    def _renumber(self):
        pass  # override in subclass if needed

    def load_rows(self, rows):
        self.clear_all()
        for idx,row in enumerate(rows):
            tag="even" if idx%2==0 else "odd"
            vals=list(row)
            while len(vals)<len(self._cols): vals.append("")
            self.insert("","end",values=vals[:len(self._cols)],tags=(tag,))

# ── Queue-specific table ──────────────────────────────────────────────────────
class QueueTable(EditableTV):
    def __init__(self, parent, **kw):
        super().__init__(parent,
            cols   =("lcsc","name","category","note"),
            hdrs   =("LCSC","Name","Category","Note"),
            widths =(110,200,120,300),
            combos ={"category":CATS},
            **kw)
    def _renumber(self): pass  # no auto-number needed

# ── Pin table ─────────────────────────────────────────────────────────────────
class PinTable(EditableTV):
    def __init__(self, parent, **kw):
        super().__init__(parent,
            cols   =("num","name","type","side"),
            hdrs   =("#","Pin Name","Type","Side"),
            widths =(50,160,100,70),
            combos ={"type":PTYPES,"side":PSIDES},
            **kw)
    def add_empty_row(self, defaults=None):
        n=len(self.get_children())+1
        super().add_empty_row(defaults or [str(n),"","IN","L"])
    def get_pins(self):
        pins=[]
        for row in self.get_all_rows():
            num,name,ptype,side=row
            if num and name:
                pins.append({"number":num,"name":name,
                             "type":ptype.lower(),
                             "side":{"L":"L","R":"R","T":"T","B":"B"}.get(side.upper(),"L")})
        return pins

# ==============================================================================
# LOG BOX
# ==============================================================================
class LogBox(tk.Frame):
    def __init__(self, parent, height=7, **kw):
        super().__init__(parent,bg=CGRAY,**kw)
        self.txt=tk.Text(self,height=height,font=("Consolas",9),
                         bg="#1E1E1E",fg="#D4D4D4",state="disabled",
                         relief="flat",wrap="word",cursor="arrow")
        sb=ttk.Scrollbar(self,command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.pack(side="left",fill="both",expand=True)
        sb.pack(side="right",fill="y")
        for tag,fg in [("ok","#4EC9B0"),("err","#F48771"),
                       ("warn","#DCDCAA"),("info","#9CDCFE"),("dim","#808080")]:
            self.txt.tag_config(tag,foreground=fg)
    def log(self,msg,tag="info"):
        self.txt.configure(state="normal")
        self.txt.insert("end",msg+"\n",tag)
        self.txt.see("end"); self.txt.configure(state="disabled")
    def clear(self):
        self.txt.configure(state="normal")
        self.txt.delete("1.0","end")
        self.txt.configure(state="disabled")

# ==============================================================================
# STYLED BUTTON
# ==============================================================================
class Btn(tk.Button):
    def __init__(self,parent,text,cmd,color=CB,fg=CW,**kw):
        super().__init__(parent,text=text,command=cmd,bg=color,fg=fg,
                         font=("Arial",9,"bold"),relief="flat",
                         padx=10,pady=5,cursor="hand2",
                         activebackground=color,activeforeground=fg,**kw)
        def dark(c):
            h=c.lstrip("#"); r,g,b=int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
            return f"#{max(0,r-30):02X}{max(0,g-30):02X}{max(0,b-30):02X}"
        self.bind("<Enter>",lambda e:self.config(bg=dark(color)))
        self.bind("<Leave>",lambda e:self.config(bg=color))

# ==============================================================================
# QUEUE TAB
# ==============================================================================
class QueueTab(ttk.Frame):
    def __init__(self, parent, status_cb):
        super().__init__(parent)
        self._running=False
        self._status=status_cb
        self._build()
        self._load_excel()

    def _build(self):
        self.configure(style="Tab.TFrame")
        # Header
        hf=tk.Frame(self,bg=CGRAY); hf.pack(fill="x",padx=8,pady=(8,4))
        tk.Label(hf,text="📥  Queue Import",font=("Arial",12,"bold"),bg=CGRAY,fg=CB).pack(side="left")

        # Quick-add bar
        qa=tk.Frame(self,bg=CGRAY,bd=1,relief="groove"); qa.pack(fill="x",padx=8,pady=4)
        tk.Label(qa,text="Quick add:",font=("Arial",9,"bold"),bg=CGRAY).pack(side="left",padx=6,pady=6)
        self.v_quick=tk.StringVar()
        e=tk.Entry(qa,textvariable=self.v_quick,font=("Arial",10),width=18)
        e.pack(side="left",padx=4,pady=4)
        e.bind("<Return>",lambda ev:self._quick_add())
        tk.Label(qa,text="Name:",font=("Arial",9),bg=CGRAY).pack(side="left",padx=(8,2))
        self.v_qname=tk.StringVar()
        tk.Entry(qa,textvariable=self.v_qname,font=("Arial",10),width=18).pack(side="left",padx=2,pady=4)
        tk.Label(qa,text="Category:",font=("Arial",9),bg=CGRAY).pack(side="left",padx=(8,2))
        self.v_qcat=tk.StringVar(value="ICs")
        ttk.Combobox(qa,textvariable=self.v_qcat,values=CATS,width=14,state="readonly").pack(side="left",padx=2)
        Btn(qa,"+ Add",self._quick_add,color="#1976D2").pack(side="left",padx=8)
        tk.Label(qa,text="  💡 Ctrl+V de paste nhieu hang tu Excel",
                 font=("Arial",8,"italic"),bg=CGRAY,fg="#666").pack(side="right",padx=8)

        # Table
        tf=tk.Frame(self,bg=CGRAY); tf.pack(fill="both",expand=True,padx=8,pady=(0,4))
        self.table=QueueTable(tf,height=12)
        sb_y=ttk.Scrollbar(tf,command=self.table.yview)
        sb_x=ttk.Scrollbar(tf,orient="horizontal",command=self.table.xview)
        self.table.configure(yscrollcommand=sb_y.set,xscrollcommand=sb_x.set)
        sb_y.pack(side="right",fill="y"); sb_x.pack(side="bottom",fill="x")
        self.table.pack(fill="both",expand=True)

        # Toolbar below table
        bt=tk.Frame(self,bg=CGRAY); bt.pack(fill="x",padx=8,pady=(0,4))
        Btn(bt,"+ Row",  self.table.add_empty_row, color="#1976D2").pack(side="left",padx=2)
        Btn(bt,"Del Sel",self._del_selected,       color="#757575").pack(side="left",padx=2)
        Btn(bt,"Clear All",self.table.clear_all,   color="#9E9E9E").pack(side="left",padx=2)
        Btn(bt,"Load Excel",self._load_excel,      color="#607D8B").pack(side="left",padx=2)
        Btn(bt,"Save Excel",self._save_excel,      color="#455A64").pack(side="left",padx=2)
        self.lbl_count=tk.Label(bt,text="0 rows",font=("Arial",9),bg=CGRAY,fg="#666")
        self.lbl_count.pack(side="left",padx=12)

        # Import section
        imf=tk.Frame(self,bg=CB,pady=6); imf.pack(fill="x",padx=0)
        self.btn_imp=Btn(imf,"  📥  Import All from LCSC  ",self._start_import,color=CO,fg=CW)
        self.btn_imp.pack(side="left",padx=12)
        self.progress=ttk.Progressbar(imf,length=350,mode="determinate")
        self.progress.pack(side="left",padx=8,fill="x",expand=True)
        self.lbl_prog=tk.Label(imf,text="",font=("Arial",9,"bold"),bg=CB,fg=CW)
        self.lbl_prog.pack(side="right",padx=12)

        # Log
        tk.Label(self,text="Import Log:",font=("Arial",9,"bold"),bg=CGRAY).pack(anchor="w",padx=8)
        self.log=LogBox(self,height=6); self.log.pack(fill="x",padx=8,pady=(0,8))

        self.table.bind("<<TreeviewSelect>>",lambda e:self._upd_count())
        self._upd_count()

    def _upd_count(self):
        n=len(self.table.get_children())
        self.lbl_count.config(text=f"{n} rows")

    def _quick_add(self):
        lcsc=self.v_quick.get().strip()
        if not lcsc: return
        row=(lcsc,self.v_qname.get().strip(),
             norm_cat(self.v_qcat.get()),
             "")
        self.table.add_empty_row(list(row))
        self.v_quick.set(""); self.v_qname.set("")
        self._upd_count()

    def _del_selected(self):
        for i in self.table.selection(): self.table.delete(i)
        self._upd_count()

    def _load_excel(self):
        rows=load_queue_from_excel()
        self.table.load_rows(rows)
        self._upd_count()
        if rows: self.log.log(f"Loaded {len(rows)} rows from Excel.","dim")

    def _save_excel(self):
        rows=self.table.get_all_rows()
        save_queue_to_excel(rows)
        self.log.log(f"Saved {len(rows)} rows to Excel.","ok")

    def _start_import(self):
        if self._running: return
        rows=self.table.get_all_rows()
        rows=[r for r in rows if r[0].strip()]
        if not rows:
            messagebox.showinfo("","Them LCSC vao Queue truoc."); return
        ez=find_easyeda()
        if not ez:
            messagebox.showerror("Loi","Khong tim thay easyeda2kicad.\npip install easyeda2kicad")
            return
        self._running=True
        self.btn_imp.config(state="disabled",text="⏳  Importing...")
        self.log.clear(); self.progress["value"]=0; self.progress["maximum"]=len(rows)
        mq=TQ.Queue()
        def worker():
            ensure_dirs()
            ok_idx=[]; fail_rows=[]; wb=None
            if os.path.exists(PARTS_FILE):
                try: wb=load_workbook(PARTS_FILE)
                except: wb=None
            for idx,(lcsc,name,cat,note) in enumerate(rows,1):
                mq.put(("log",f"[{idx}/{len(rows)}] {lcsc}  {name}","info"))
                sf=sym_path(cat); ensure_sym(sf); clear_temp()
                res=subprocess.run([ez,"--full",f"--lcsc_id={lcsc}",
                                    "--output",os.path.join("temp","temp")],
                                   capture_output=True,text=True)
                if res.returncode!=0:
                    mq.put(("log",f"  ✗ {res.stderr.strip()[:100]}","err"))
                    fail_rows.append((lcsc,name,cat,note)); continue
                blocks=extract_blocks(os.path.join(TEMP_DIR,"temp.kicad_sym"))
                sym_ok=False
                existing={b["name"] for b in extract_blocks(sf)}
                for b in blocks:
                    if b["name"] not in existing:
                        merge_block(b["content"],sf); sym_ok=True
                fp_dir=os.path.join(TEMP_DIR,"temp.pretty")
                fp_ok=os.path.exists(fp_dir)
                if fp_ok:
                    for fn in os.listdir(fp_dir):
                        src=os.path.join(fp_dir,fn); dst=os.path.join(FP_DIR,fn)
                        if not os.path.exists(dst): shutil.copy2(src,dst)
                m3_dir=os.path.join(TEMP_DIR,"temp.3dshapes")
                m3_ok=os.path.exists(m3_dir)
                if m3_ok:
                    for fn in os.listdir(m3_dir):
                        src=os.path.join(m3_dir,fn); dst=os.path.join(M3D_DIR,fn)
                        if not os.path.exists(dst): shutil.copy2(src,dst)
                if wb and SHEET_LIB in wb.sheetnames:
                    excel_add_lib_row(wb,lcsc,name,cat,note,sym_ok,fp_ok,m3_ok)
                ok_idx.append(idx-1)
                mq.put(("log",f"  ✓ sym={sym_ok} fp={fp_ok} 3d={m3_ok}","ok"))
                mq.put(("prog",idx,None))
            clear_temp()
            if wb:
                ws=wb[SHEET_Q]
                for row in ws.iter_rows(min_row=2):
                    for c in row: c.value=None
                for ri,row in enumerate(fail_rows,2):
                    for ci,v in enumerate(row,1): ws.cell(row=ri,column=ci,value=v)
                try: wb.save(PARTS_FILE)
                except: pass
            mq.put(("done",ok_idx,fail_rows))
        threading.Thread(target=worker,daemon=True).start()
        self._poll(mq)

    def _poll(self,mq):
        try:
            while True:
                msg=mq.get_nowait()
                if msg[0]=="log": self.log.log(msg[1],msg[2])
                elif msg[0]=="prog":
                    self.progress["value"]=msg[1]
                    self.lbl_prog.config(text=f"{msg[1]}/{int(self.progress['maximum'])}")
                elif msg[0]=="done":
                    ok_idx,fail=msg[1],msg[2]
                    # Remove ok rows from table
                    children=self.table.get_children()
                    for i in sorted(ok_idx,reverse=True):
                        if i<len(children): self.table.delete(children[i])
                    self.log.log(f"Done: {len(ok_idx)} OK, {len(fail)} FAIL","ok" if not fail else "warn")
                    self.btn_imp.config(state="normal",text="  📥  Import All from LCSC  ")
                    self._running=False; self._upd_count()
                    self._status(f"Import done: {len(ok_idx)} OK, {len(fail)} failed")
                    return
        except TQ.Empty: pass
        self.after(100,self._poll,mq)

# ==============================================================================
# CUSTOM SYMBOL TAB
# ==============================================================================
class CustomTab(ttk.Frame):
    def __init__(self, parent, status_cb):
        super().__init__(parent)
        self._status=status_cb
        self._build()

    def _build(self):
        tk.Label(self,text="✏️  Custom Symbol Creator",font=("Arial",12,"bold"),
                 bg=CGRAY,fg=CG).pack(anchor="w",padx=8,pady=(8,4))

        # Metadata
        mf=tk.LabelFrame(self,text="Thong tin",font=("Arial",9,"bold"),
                          bg=CGRAY,fg=CB,padx=8,pady=6)
        mf.pack(fill="x",padx=8,pady=(0,6))

        r1=tk.Frame(mf,bg=CGRAY); r1.pack(fill="x",pady=2)
        for lbl,sv,w,widget in [
            ("Symbol Name",tk.StringVar(),20,"entry"),
            ("Category",   tk.StringVar(value="ICs"),14,"combo_cat"),
            ("Reference",  tk.StringVar(value="U"),6,"combo_ref"),
        ]:
            tk.Label(r1,text=lbl+":",font=("Arial",9,"bold"),bg=CGRAY).pack(side="left",padx=(8,2))
            if widget=="entry":
                self.v_name=sv
                tk.Entry(r1,textvariable=sv,width=w,font=("Arial",10)).pack(side="left",padx=(0,8))
            elif widget=="combo_cat":
                self.v_cat=sv
                ttk.Combobox(r1,textvariable=sv,values=CATS,width=w,state="readonly").pack(side="left",padx=(0,8))
            elif widget=="combo_ref":
                self.v_ref=sv
                ttk.Combobox(r1,textvariable=sv,values=REFS,width=w).pack(side="left",padx=(0,8))

        r2=tk.Frame(mf,bg=CGRAY); r2.pack(fill="x",pady=2)
        tk.Label(r2,text="Description:",font=("Arial",9,"bold"),bg=CGRAY).pack(side="left",padx=(8,2))
        self.v_desc=tk.StringVar()
        tk.Entry(r2,textvariable=self.v_desc,width=70,font=("Arial",10)).pack(side="left",fill="x",expand=True,padx=(0,8))

        # Pin table
        pf=tk.LabelFrame(self,text="Pin Table  (click de edit | Ctrl+V paste tu Excel | Tab di chuyen o)",
                          font=("Arial",9,"bold"),bg=CGRAY,fg=CB,padx=4,pady=4)
        pf.pack(fill="both",expand=True,padx=8,pady=(0,6))

        tb=tk.Frame(pf,bg=CGRAY); tb.pack(fill="x",pady=(0,4))
        Btn(tb,"+ Pin",   self._add_pin,   color="#1976D2").pack(side="left",padx=2)
        Btn(tb,"Del Sel", self._del_pin,   color="#757575").pack(side="left",padx=2)
        Btn(tb,"Clear",   self.pin_tbl_clear, color="#9E9E9E").pack(side="left",padx=2)
        tk.Label(tb,text="  💡 Paste cot pin tu Excel: so|ten|loai|side",
                 font=("Arial",8,"italic"),bg=CGRAY,fg="#666").pack(side="right",padx=8)

        tf2=tk.Frame(pf,bg=CGRAY); tf2.pack(fill="both",expand=True)
        self.pin_tbl=PinTable(tf2,height=12)
        sb=ttk.Scrollbar(tf2,command=self.pin_tbl.yview)
        self.pin_tbl.configure(yscrollcommand=sb.set)
        sb.pack(side="right",fill="y"); self.pin_tbl.pack(fill="both",expand=True)

        # Action
        af=tk.Frame(self,bg=CGRAY); af.pack(fill="x",padx=8,pady=(0,4))
        Btn(af,"✏️  Create Symbol",self._create,color=CG).pack(side="left")
        self.lbl_res=tk.Label(af,text="",font=("Arial",10,"bold"),bg=CGRAY)
        self.lbl_res.pack(side="left",padx=12)
        tk.Label(self,text="Log:",font=("Arial",9,"bold"),bg=CGRAY).pack(anchor="w",padx=8)
        self.log=LogBox(self,height=5); self.log.pack(fill="x",padx=8,pady=(0,8))

    def _add_pin(self): self.pin_tbl.add_empty_row()
    def _del_pin(self):
        for i in self.pin_tbl.selection(): self.pin_tbl.delete(i)
    def pin_tbl_clear(self): self.pin_tbl.clear_all()

    def _create(self):
        name=self.v_name.get().strip()
        cat =self.v_cat.get().strip()
        ref =self.v_ref.get().strip() or "U"
        desc=self.v_desc.get().strip()
        if not name: messagebox.showwarning("","Chua dien Symbol Name"); return
        pins=self.pin_tbl.get_pins()
        if not pins: messagebox.showwarning("","Chua co pin nao."); return

        ensure_dirs()
        sf=sym_path(cat); ensure_sym(sf)
        if sym_exists(sf,name):
            if not messagebox.askyesno("Da ton tai",f"'{name}' da co.\nGhi de?"): return
        try:
            block=gen_kicad_sym(name,ref,desc,pins)
            if sym_exists(sf,name):
                with open(sf,"r",encoding="utf-8") as f: c=f.read()
                s=c.find(f'(symbol "{name}"')
                if s!=-1:
                    stack=0; e=-1
                    for i in range(s,len(c)):
                        if c[i]=="(": stack+=1
                        elif c[i]==")":
                            stack-=1
                            if stack==0: e=i+1; break
                    if e!=-1:
                        c=c[:s]+c[e:]
                        last=c.rfind(")"); c=c[:last]+"\n  "+block+"\n"+c[last:]
                        with open(sf,"w",encoding="utf-8") as f: f.write(c)
            else:
                merge_block(block,sf)
            if os.path.exists(PARTS_FILE):
                wb=load_workbook(PARTS_FILE)
                if SHEET_LIB in wb.sheetnames:
                    excel_add_lib_row(wb,name,name,cat,desc,True,False,False)
                    wb.save(PARTS_FILE)
            self.lbl_res.config(text=f"✓ '{name}' tao xong!",fg=COK)
            self.log.log(f"Symbol '{name}' ({len(pins)} pins) -> {sf}","ok")
            L=sum(1 for p in pins if p["side"]=="L")
            R=sum(1 for p in pins if p["side"]=="R")
            T=sum(1 for p in pins if p["side"]=="T")
            B=sum(1 for p in pins if p["side"]=="B")
            self.log.log(f"  Layout: Left={L} Right={R} Top={T} Bot={B}","info")
            self._status(f"Symbol '{name}' created OK")
        except Exception as ex:
            self.lbl_res.config(text=f"✗ {ex}",fg=CERR)
            self.log.log(f"LOI: {ex}","err")

# ==============================================================================
# LIBRARY TAB
# ==============================================================================
class LibTab(ttk.Frame):
    def __init__(self, parent, status_cb):
        super().__init__(parent)
        self._status=status_cb
        self._build()

    def _build(self):
        hf=tk.Frame(self,bg=CGRAY); hf.pack(fill="x",padx=8,pady=8)
        tk.Label(hf,text="📚  Library",font=("Arial",12,"bold"),bg=CGRAY,fg=CB).pack(side="left")
        Btn(hf,"Refresh",self._load,color="#607D8B").pack(side="right")
        self.lbl_n=tk.Label(hf,text="",font=("Arial",9),bg=CGRAY,fg="#666")
        self.lbl_n.pack(side="right",padx=12)

        tf=tk.Frame(self,bg=CGRAY); tf.pack(fill="both",expand=True,padx=8)
        cols=("lcsc","name","cat","note","sym","fp","3d","date")
        hdrs=("LCSC","Name","Category","Note","Sym","FP","3D","Date")
        widths=(100,180,110,230,50,50,50,130)
        self.tree=ttk.Treeview(tf,columns=cols,show="headings",height=22)
        for col,hdr,w in zip(cols,hdrs,widths):
            self.tree.heading(col,text=hdr); self.tree.column(col,width=w,minwidth=40)
        self.tree.tag_configure("ok",background="#F1FFF1")
        sb=ttk.Scrollbar(tf,command=self.tree.yview)
        sbx=ttk.Scrollbar(tf,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=sb.set,xscrollcommand=sbx.set)
        sb.pack(side="right",fill="y"); sbx.pack(side="bottom",fill="x")
        self.tree.pack(fill="both",expand=True)

    def _load(self):
        self.tree.delete(*self.tree.get_children())
        if not os.path.exists(PARTS_FILE): return
        try:
            wb=load_workbook(PARTS_FILE); count=0
            if SHEET_LIB in wb.sheetnames:
                ws=wb[SHEET_LIB]
                for row in ws.iter_rows(min_row=2,values_only=True):
                    if not row[0]: continue
                    vals=[str(v or "").strip() for v in row[:8]]
                    tag="ok" if len(vals)>4 and vals[4]=="v" else ""
                    self.tree.insert("","end",values=vals,tags=(tag,)); count+=1
            self.lbl_n.config(text=f"{count} components")
            self._status(f"Library: {count} components")
        except Exception as e: self.lbl_n.config(text=f"Error: {e}")

# ==============================================================================
# MAIN APP
# ==============================================================================
class App:
    def __init__(self):
        self.root=tk.Tk()
        self.root.title(f"{APP_TITLE}  v{APP_VER}")
        self.root.geometry("1000x720"); self.root.minsize(820,600)
        self.root.configure(bg=CGRAY)
        if os.name=="nt":
            try: self.root.iconbitmap(default="")
            except: pass
        self._setup_style()
        self._build()

    def _setup_style(self):
        s=ttk.Style()
        try: s.theme_use("clam")
        except: pass
        s.configure("TFrame",background=CGRAY)
        s.configure("Tab.TFrame",background=CGRAY)
        s.configure("TNotebook",background=CDARK)
        s.configure("TNotebook.Tab",font=("Arial",10,"bold"),
                    padding=(14,7),background="#455A64",foreground=CW)
        s.map("TNotebook.Tab",
              background=[("selected",CB)],foreground=[("selected",CW)])
        s.configure("Treeview.Heading",background=CB,foreground=CW,
                    font=("Arial",9,"bold"),relief="flat")
        s.map("Treeview.Heading",background=[("active","#1565C0")])
        s.configure("Treeview",font=("Arial",9),rowheight=22,background=CW)

    def _build(self):
        # Header
        hb=tk.Frame(self.root,bg=CDARK,height=54); hb.pack(fill="x"); hb.pack_propagate(False)
        tk.Label(hb,text=f"  🔧  {APP_TITLE}",font=("Arial",14,"bold"),
                 bg=CDARK,fg=CW).pack(side="left",pady=10)
        tk.Label(hb,text=f"v{APP_VER}",font=("Arial",9),bg=CDARK,fg="#78909C").pack(side="right",padx=12)

        # Browse file
        ff=tk.Frame(hb,bg=CDARK); ff.pack(side="right",padx=8)
        tk.Label(ff,text="File:",font=("Arial",9),bg=CDARK,fg="#90CAF9").pack(side="left")
        self.v_file=tk.StringVar(value=PARTS_FILE)
        tk.Entry(ff,textvariable=self.v_file,width=24,font=("Arial",9)).pack(side="left",padx=4)
        tk.Button(ff,text="…",command=self._browse,font=("Arial",9),
                  bg="#1565C0",fg=CW,relief="flat",padx=6).pack(side="left")

        # Tabs
        self.nb=ttk.Notebook(self.root); self.nb.pack(fill="both",expand=True)
        self.tab_q  =QueueTab  (self.nb, self._set_status)
        self.tab_cus=CustomTab (self.nb, self._set_status)
        self.tab_lib=LibTab    (self.nb, self._set_status)
        self.nb.add(self.tab_q,   text="  📥  Queue Import  ")
        self.nb.add(self.tab_cus, text="  ✏️  Custom Symbol  ")
        self.nb.add(self.tab_lib, text="  📚  Library  ")
        self.nb.bind("<<NotebookTabChanged>>",self._on_tab)

        # Status bar
        sb2=tk.Frame(self.root,bg=CDARK,height=22); sb2.pack(fill="x",side="bottom")
        sb2.pack_propagate(False)
        self.lbl_stat=tk.Label(sb2,text="Ready",font=("Arial",8),bg=CDARK,fg="#78909C")
        self.lbl_stat.pack(side="left",padx=8)
        tk.Label(sb2,text="Click cell to edit  |  Ctrl+V paste  |  Del remove  |  Right-click menu",
                 font=("Arial",8),bg=CDARK,fg="#546E7A").pack(side="right",padx=8)

    def _browse(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx"),("All","*.*")])
        if p: self.v_file.set(p); global PARTS_FILE; PARTS_FILE=p

    def _on_tab(self, event):
        idx=self.nb.index(self.nb.select())
        if idx==2: self.tab_lib._load()

    def _set_status(self, msg):
        self.lbl_stat.config(text=msg)

    def run(self):
        self.tab_lib._load()
        self.root.mainloop()

# ── Entry ─────────────────────────────────────────────────────────────────────
if __name__=="__main__":
    if getattr(sys,"frozen",False):
        os.chdir(os.path.dirname(sys.executable))
    else:
        script_dir=os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)
    App().run()