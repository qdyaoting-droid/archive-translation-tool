"""
ocr_fix.py — 针对乱码页面的轻量级 OCR 纠错

策略：
1. 检测 chunks/ 里每个 txt 文件的"乱码程度"
2. 乱码严重的页面 → 发给 Claude Haiku（便宜）做 OCR 纠错
3. 正常页面 → 保持不动，继续用 DeepSeek 翻译

乱码检测标准：
- 可打印 ASCII 字符比例 < 60%，或
- 包含大量连续替换字符（□、■、▪ 等），或
- 字符总数 < 30（空白页另行处理）

依赖：anthropic（pip install anthropic）
"""

import re
import anthropic
from pathlib import Path

# ============================================================
# 配置
# ============================================================
CHUNKS_DIR = Path("chunks")
CLAUDE_API_KEY = ""  # 填入你的 Anthropic API Key

# 乱码检测阈值
MIN_ASCII_RATIO = 0.60      # ASCII 可打印字符比例低于此值 → 乱码
GARBLE_CHARS = set("□■▪▫▬▭▮▯●◆◇◈◉◊○◌◍◎◐◑◒◓◔◕◖◗")

# ============================================================
# 乱码检测
# ============================================================
def garble_score(text: str) -> float:
    """返回 0~1 的乱码程度，越高越乱"""
    if not text or len(text.strip()) < 10:
        return 0.0

    total = len(text)
    printable_ascii = sum(1 for c in text if 32 <= ord(c) <= 126)
    garble_count = sum(1 for c in text if c in GARBLE_CHARS)

    ascii_ratio = printable_ascii / total
    garble_ratio = garble_count / total

    score = (1 - ascii_ratio) * 0.7 + garble_ratio * 0.3
    return score

def is_garbled(text: str, threshold: float = 0.35) -> bool:
    """判断一个 txt 页面是否乱码"""
    # 排除页码标记行后判断
    content = "\n".join(
        l for l in text.splitlines()
        if not l.strip().startswith("===== PAGE")
    )
    if len(content.strip()) < 30:
        return False  # 太短的页面可能是空白页，不处理
    return garble_score(content) > threshold

# ============================================================
# Claude Haiku OCR 纠错
# ============================================================
def fix_with_haiku(client: anthropic.Anthropic, garbled_text: str, page_hint: str = "") -> str:
    """用 Claude Haiku 纠正 OCR 乱码文本"""
    prompt = f"""以下是从历史档案扫描件中提取的文字，因OCR识别错误导致部分文字乱码或错误。
请尽力还原原始英文或法文文本，保持原文格式、段落和标点。
如果某段实在无法辨认，用 [原文模糊，无法辨认] 标注。
不要翻译，只做纠错。

{f'页面参考：{page_hint}' if page_hint else ''}

---
{garbled_text}
---

请输出纠错后的文本："""

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}]
    )
    return message.content[0].text

# ============================================================
# 主流程
# ============================================================
def main():
    if not CLAUDE_API_KEY:
        print("❌ 请先在脚本顶部填入 CLAUDE_API_KEY")
        return

    txt_files = sorted(CHUNKS_DIR.glob("*.txt"))
    if not txt_files:
        print("❌ chunks/ 目录下没有找到 txt 文件，请先运行 extract_pdf_text.py")
        return

    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)

    garbled_pages = []
    for f in txt_files:
        text = f.read_text(encoding="utf-8")
        if is_garbled(text):
            garbled_pages.append(f)

    print(f"共 {len(txt_files)} 页，检测到 {len(garbled_pages)} 页乱码")

    if not garbled_pages:
        print("✅ 无乱码页面，无需处理")
        return

    for page_file in garbled_pages:
        original = page_file.read_text(encoding="utf-8")
        print(f"  纠错中：{page_file.name}（乱码分数 {garble_score(original):.2f}）")

        fixed = fix_with_haiku(client, original, page_hint=page_file.name)

        # 保留页码标记行
        page_markers = [
            l for l in original.splitlines()
            if l.strip().startswith("===== PAGE")
        ]
        if page_markers:
            fixed = page_markers[0] + "\n\n" + fixed

        page_file.write_text(fixed, encoding="utf-8")
        print(f"  ✅ 已修复 → {page_file.name}")

    print("\n全部乱码页面已处理完毕，可继续运行翻译流程。")

if __name__ == "__main__":
    main()
