import os
from fontTools.ttLib import TTCollection

# Windows 字型資料夾
font_dir = r"C:\Windows\Fonts"
out_dir = os.path.join(font_dir, "ttf_export")

os.makedirs(out_dir, exist_ok=True)

for file in os.listdir(font_dir):
    if file.lower().endswith(".ttc"):
        ttc_path = os.path.join(font_dir, file)
        try:
            ttc = TTCollection(ttc_path)
            for i, font in enumerate(ttc.fonts):
                base_name = os.path.splitext(file)[0]
                out_path = os.path.join(out_dir, f"{base_name}_{i}.ttf")
                font.save(out_path)
                print(f"✔ {ttc_path} → {out_path}")
        except Exception as e:
            print(f"⚠ 無法處理 {ttc_path}: {e}")

print(f"\n✅ 完成！輸出目錄：{out_dir}")