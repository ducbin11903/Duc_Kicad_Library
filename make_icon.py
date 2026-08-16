#!/usr/bin/env python3
"""
make_icon.py — convert a PNG/JPG logo into a multi-size Windows .ico

Usage:
    pip install pillow
    python make_icon.py logo.png
    python make_icon.py                 (looks for logo.png in this folder)

Produces: app.ico
"""
import sys, os
from PIL import Image

SIZES=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)]
OUT="app.ico"

def main():
    src=sys.argv[1] if len(sys.argv)>1 else "logo.png"
    if not os.path.exists(src):
        print(f"Not found: {src}")
        print("Save your logo as logo.png next to this script, or pass the path:")
        print("    python make_icon.py my_logo.png")
        sys.exit(1)

    img=Image.open(src).convert("RGBA")

    # Square-crop from the centre if the source is not square
    w,h=img.size
    if w!=h:
        side=min(w,h)
        left,top=(w-side)//2,(h-side)//2
        img=img.crop((left,top,left+side,top+side))
        print(f"Cropped {w}x{h} -> {side}x{side}")

    img.save(OUT,format="ICO",sizes=SIZES)
    print(f"Created {OUT}  ({os.path.getsize(OUT)//1024} KB)")
    print("Sizes:", ", ".join(f"{a}x{b}" for a,b in SIZES))

    # 256px PNG for the in-app title bar / about box
    img.resize((256,256),Image.LANCZOS).save("app_logo.png")
    print("Created app_logo.png (used inside the app window)")

if __name__=="__main__":
    main()
