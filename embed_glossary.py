#!/usr/bin/env python3
"""
重新嵌入术语表到 app.py。

用法：
  python embed_glossary.py <glossary.xlsx> [app.py路径]

示例：
  python embed_glossary.py new_glossary.xlsx
  python embed_glossary.py new_glossary.xlsx ../translate-app/app.py

修改术语表后运行此脚本，再 git commit + push 即可更新线上的内置术语表。
"""

import sys
import os
import base64
import zlib
import json
from pathlib import Path

# 允许从同目录或上级 translate-app 目录导入
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR / "translate-app"))
sys.path.insert(0, str(SCRIPT_DIR.parent / "translate-app"))

try:
    from glossary_helper import load_glossary
except ImportError:
    print("❌ 找不到 glossary_helper.py，请确保它在同目录或 translate-app/ 下")
    sys.exit(1)


def embed_glossary(xlsx_path: str, app_py_path: str):
    """Generate compressed base64 of glossary and inject into app.py."""
    # 1. Load glossary
    print(f"📖 加载术语表：{xlsx_path}")
    data = load_glossary(xlsx_path)
    print(f"   共 {data['count']} 条术语")

    # 2. Compress and encode
    json_str = json.dumps(data, ensure_ascii=False)
    compressed = zlib.compress(json_str.encode("utf-8"))
    b64 = base64.b64encode(compressed).decode("ascii")
    print(f"   压缩后 {len(b64)} 字符")

    # 3. Read app.py
    print(f"📝 更新 {app_py_path} …")
    with open(app_py_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 4. Find and replace the embedded glossary
    marker_start = "_EMBEDDED_GLOSSARY_B64 = "
    if marker_start not in content:
        print("❌ app.py 中找不到 _EMBEDDED_GLOSSARY_B64，请确认文件正确")
        sys.exit(1)

    # Replace: find the triple-quoted string after the marker
    import re
    pattern = r'(_EMBEDDED_GLOSSARY_B64 = """)([^"]*)(""")'
    new_content = re.sub(pattern, f'_EMBEDDED_GLOSSARY_B64 = """{b64}"""', content)

    if new_content == content:
        print("❌ 替换失败，请检查 app.py 格式")
        sys.exit(1)

    # 5. Write back
    with open(app_py_path, "w", encoding="utf-8") as f:
        f.write(new_content)

    print(f"✅ 已更新 {app_py_path} 中的内置术语表")
    print()
    print("下一步：")
    print("  cd translate-app")
    print("  git add app.py")
    print('  git commit -m "chore: update glossary"')
    print("  git push")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python embed_glossary.py <glossary.xlsx> [app.py路径]")
        sys.exit(1)

    xlsx_path = sys.argv[1]

    if len(sys.argv) >= 3:
        app_py_path = sys.argv[2]
    else:
        # Auto-detect app.py
        candidates = [
            SCRIPT_DIR / "app.py",
            SCRIPT_DIR / "translate-app" / "app.py",
            SCRIPT_DIR.parent / "translate-app" / "app.py",
        ]
        app_py_path = None
        for c in candidates:
            if c.exists():
                app_py_path = str(c)
                break
        if app_py_path is None:
            print("❌ 找不到 app.py，请手动指定路径")
            print("用法：python embed_glossary.py <glossary.xlsx> <app.py路径>")
            sys.exit(1)

    if not os.path.exists(xlsx_path):
        print(f"❌ 找不到术语表文件：{xlsx_path}")
        sys.exit(1)

    if not os.path.exists(app_py_path):
        print(f"❌ 找不到 app.py：{app_py_path}")
        sys.exit(1)

    embed_glossary(xlsx_path, app_py_path)
