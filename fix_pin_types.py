#!/usr/bin/env python3
"""
fix_pin_types.py
================
Rewrites the electrical type of pins inside KiCad .kicad_sym libraries.

easyeda2kicad usually imports every pin as `unspecified`, which makes KiCad ERC
complain on almost every net. Converting those pins to `passive` silences the
false warnings while keeping real ones.

Usage
-----
    python fix_pin_types.py                     preview only, changes nothing
    python fix_pin_types.py --apply             convert unspecified -> passive
    python fix_pin_types.py --apply --all       convert EVERY pin -> passive
    python fix_pin_types.py --apply --dir path  use another symbols folder
    python fix_pin_types.py --apply --keep-power    leave power pins untouched

A timestamped backup of every file is written before anything is modified.
"""

import os, re, sys, shutil, argparse
from datetime import datetime

# (pin <electrical_type> <graphic_style> (at ...
PIN_RE = re.compile(r'(\(pin\s+)([a-z_]+)(\s+[a-z_]+\s+\(at\b)')

VALID = {"input","output","bidirectional","tri_state","passive","free",
         "unspecified","power_in","power_out","open_collector",
         "open_emitter","no_connect"}

POWER = {"power_in","power_out"}


def script_dir():
    return os.path.dirname(os.path.abspath(__file__))


def locate_symbols(explicit=None):
    """
    Work out where the symbols folder is:
      1. --dir if it was given
      2. a 'symbols' folder next to this script
      3. walk up from the script, up to 4 levels
      4. the folder the terminal is in
    """
    if explicit:
        d=os.path.abspath(explicit)
        return d if os.path.isdir(d) else None

    here=script_dir()

    cand=os.path.join(here,"symbols")
    if os.path.isdir(cand): return cand

    # the script itself may live in symbols/
    if os.path.basename(here).lower()=="symbols": return here

    d=here
    for _ in range(4):
        parent=os.path.dirname(d)
        if parent==d: break
        d=parent
        cand=os.path.join(d,"symbols")
        if os.path.isdir(cand): return cand

    cand=os.path.join(os.getcwd(),"symbols")
    if os.path.isdir(cand): return cand
    return None


def find_symbol_files(folder):
    out=[]
    for dirpath,_,names in os.walk(folder):
        for n in sorted(names):
            if n.lower().endswith(".kicad_sym"):
                out.append(os.path.join(dirpath,n))
    return out


def convert(text, target="passive", only_unspecified=True, keep_power=True):
    """Return (new_text, counts) where counts maps old_type -> number changed."""
    counts={}

    def repl(m):
        head, etype, tail = m.group(1), m.group(2), m.group(3)
        if etype == target:
            return m.group(0)
        if only_unspecified and etype != "unspecified":
            return m.group(0)
        if keep_power and etype in POWER:
            return m.group(0)
        counts[etype] = counts.get(etype, 0) + 1
        return head + target + tail

    return PIN_RE.sub(repl, text), counts


def count_types(text):
    stats={}
    for m in PIN_RE.finditer(text):
        t=m.group(2)
        stats[t]=stats.get(t,0)+1
    return stats


def main():
    ap=argparse.ArgumentParser(description="Change pin electrical types in .kicad_sym files")
    ap.add_argument("--dir", default=None,
                    help="folder holding the .kicad_sym files "
                         "(found automatically if left out)")
    ap.add_argument("--apply", action="store_true", help="write the changes (otherwise preview)")
    ap.add_argument("--all", action="store_true", help="convert every pin, not just unspecified")
    ap.add_argument("--target", default="passive", help="type to convert to (default: passive)")
    ap.add_argument("--keep-power", dest="keep_power", action="store_true", default=True,
                    help="leave power_in / power_out alone (default)")
    ap.add_argument("--include-power", dest="keep_power", action="store_false",
                    help="convert power pins as well")
    args=ap.parse_args()

    if args.target not in VALID:
        print(f"Unknown pin type: {args.target}")
        print("Valid types: "+", ".join(sorted(VALID)))
        sys.exit(1)

    folder=locate_symbols(args.dir)
    if not folder:
        print("Could not find a 'symbols' folder.")
        print()
        print(f"  script location : {script_dir()}")
        print(f"  terminal folder : {os.getcwd()}")
        print()
        print("Point at it directly, for example:")
        print('    python fix_pin_types.py --dir "D:\\KiCad Extensions\\Duc-KiCad-Lib\\symbols"')
        sys.exit(1)

    files=find_symbol_files(folder)
    if not files:
        print(f"No .kicad_sym files under {folder}")
        sys.exit(1)

    mode = "every pin" if args.all else "unspecified pins only"
    print(f"Folder : {folder}")
    print(f"Files  : {len(files)}")
    print(f"Mode   : {mode}  ->  {args.target}")
    print(f"Power  : {'kept as is' if args.keep_power else 'converted too'}")
    print(f"Action : {'WRITE' if args.apply else 'PREVIEW (nothing will change)'}")
    print("-"*64)

    # backup
    backup=None
    if args.apply:
        stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        backup=os.path.join(os.path.dirname(folder), f"symbols_backup_{stamp}")
        shutil.copytree(folder, backup)
        print(f"Backup : {backup}")
        print("-"*64)

    grand={}
    changed_files=0

    for path in files:
        try:
            text=open(path, encoding="utf-8").read()
        except Exception as ex:
            print(f"  skip {os.path.basename(path)}: {ex}")
            continue

        before=count_types(text)
        new, counts = convert(text, args.target,
                              only_unspecified=not args.all,
                              keep_power=args.keep_power)
        total=sum(counts.values())
        name=os.path.basename(path)

        if total==0:
            summary=", ".join(f"{k}={v}" for k,v in sorted(before.items())) or "no pins"
            print(f"  {name:<34} unchanged   ({summary})")
            continue

        detail=", ".join(f"{k}->{args.target}: {v}" for k,v in sorted(counts.items()))
        print(f"  {name:<34} {total:>4} pins   {detail}")
        for k,v in counts.items(): grand[k]=grand.get(k,0)+v
        changed_files+=1

        if args.apply:
            open(path,"w",encoding="utf-8").write(new)

    print("-"*64)
    if not grand:
        print("Nothing to change.")
        return

    total=sum(grand.values())
    print(f"{total} pins in {changed_files} files:")
    for k,v in sorted(grand.items(), key=lambda x:-x[1]):
        print(f"    {k:<16} -> {args.target:<12} {v}")

    if args.apply:
        print()
        print("Done. Reload the libraries in KiCad to see the change.")
        print(f"Backup kept at: {backup}")
    else:
        print()
        print("This was a preview. Re-run with --apply to write the changes.")


if __name__=="__main__":
    main()