# Duc-KiCad-Lib

Thư viện KiCad cá nhân — convert từ LCSC/EasyEDA.

## Cách thêm linh kiện mới

1. Mở `parts.csv`, thêm dòng mới:
```
C{số},Tên linh kiện,Category,Ghi chú
```

2. Chạy script:
```bash
python generate.py
```

3. Push lên git:
```bash
git add .
git commit -m "Add {tên linh kiện}"
git push
```

## Setup KiCad (làm 1 lần trên mỗi máy)

**1. Set Environment Variable**

Preferences > Configure Paths:
| Name | Path |
|---|---|
| `KICAD_USER_LIB` | `<path tới thư mục này>` |

**2. Add Symbol Library**

Preferences > Manage Symbol Libraries > Global Libraries:
```
${KICAD_USER_LIB}/symbols/Duc_Library.kicad_sym
```

**3. Add Footprint Library**

Preferences > Manage Footprint Libraries > Global Libraries:
```
${KICAD_USER_LIB}/footprints/Duc_Library.pretty
```

## Sync giữa 2 máy

```bash
# Trước khi làm việc
git pull

# Sau khi thêm linh kiện mới
git add .
git commit -m "Add ..."
git push
```
