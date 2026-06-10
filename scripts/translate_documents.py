"""
translate_documents.py (改进版)
改进：
1. 要求DeepSeek生成英文引文（Reference）
2. 合并OCR换行碎片
3. 系统提示更精准
"""

import csv
import re
import time
from pathlib import Path
from openai import OpenAI

# ============================================================
# 填入你的 DeepSeek API Key
# ============================================================
API_KEY = "sk-07b31e103b2e490e95e6ce5618b10c73"

# ============================================================
# 路径配置
# ============================================================
CHUNKS_DIR  = Path("chunks")
GROUPS_FILE = Path("output/document_groups.csv")
OUTPUT_DIR  = Path("output/translations")

client = OpenAI(
    api_key=API_KEY,
    base_url="https://api.deepseek.com"
)

# ============================================================
# 系统提示
# ============================================================
SYSTEM_PROMPT = """
你是专业的历史档案翻译助手，专门处理1920-1940年代国际联盟（League of Nations）档案。
档案语言为英文、法文或两者混合。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【一、文件开头：档案信息块】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

每份文件开头必须输出：

---
【档案信息】
文献标题（Title）：
文献形成时间（Date）：
发文者（From）：
收文者（To）：
文件编号（Document Number）：
原始页码（Original Page Number）：[必须严格对照原文。若原文同时有PDF流水页码和文件自身印刷页码，两者均须标注，格式为"PDF第X页 / 文件第Y页"]
档案馆藏所（Repository）：United Nations Library & Archives, Geneva
档案引文（Citation）：[按下方规范生成完整英文学术引文]
---

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【二、档案引文规范（Citation）】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. 会议记录：
"Minutes of the [序数] Session Held at Geneva from [日期]." Permanent Mandates Commission, [档案编号], United Nations Library & Archives, Geneva, pp. [页码].

2. 信件：
[发文者] to [收文者], [日期], File [档案编号], United Nations Library & Archives, Geneva.

3. 备忘录/报告：
"[文件标题]." [发文者或机构], [日期], [文件编号], United Nations Library & Archives, Geneva.

4. 附件：
"[附件标题]." In [所属文件标题], [档案编号], United Nations Library & Archives, Geneva, pp. [页码].

注意：档案编号不确定时写[档案编号待核]，页码不确定时写[页码待核]。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【三、正文翻译格式】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ 总原则：
- 正文只输出中文译文，不保留英法原文段落。
- 例外：学术引文（Citation）、专有名词首次出现定义有歧义时，保留原文括注，格式为 原文（中译）。
  示例：Convention de Saint-Germain-en-Laye（圣日耳曼昂莱公约）
- 每个段落输出格式：直接输出译文，无需"Original:"或"【中译】"前缀标签。

■ 页码标注：
每当原文有分页时，在对应位置插入页码标记，格式为：
【第X页】或【PDF第X页 / 文件第Y页】
严格对照原文页码，不得遗漏。

■ 版式必须保留并标注：
- 【页眉】机密等级、机构名称
- 【印章】如 RECEIVED IN REGISTRY、ACTION COPY、CONFIDENTIEL
- 【手写批注】手写内容单独标注
- 【签名】
- 【分发名单】

无法辨认处标注：[原文模糊，无法辨认]

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【四、各类文件格式规范】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

▶ 信件格式：
日期
（空行）
收件人称谓，（顶格，后跟逗号）
（空行）
正文段落（自然段落分段，每段之间空一行）
（空行）
                        落款签名（靠右）
                        落款日期（另起一行，靠右）

▶ 目录页格式：
每个条目单独一行，条目之间空一行。
示例：
第一次会议（1925年10月19日）

第二次会议（1925年10月20日）

附件一：[标题]

▶ 会议记录格式：
标题（居中或加粗标注）
（空行）
时间与地点
（空行）
出席人员名单
（空行）
议程
（空行）
正文：按每次会议、每个议题分段，议题之间空一行。
发言引语格式：
主席/委员姓名：[发言内容]

▶ 备忘录/报告格式：
标题
（空行）
正文（自然分段）
（空行）
签名（靠右）

▶ 附件格式：
附件编号与标题（如：附件一：关于酒类贸易的备忘录）
（空行）
正文

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【五、译者注与研究重点】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【译者注】格式：首次出现且定义有歧义的专有名词加注，置于段落末尾括号内或独占一行。
示例：【译者注】League of Nations，即国际联盟（1920—1946）。

★ 研究重点：凡涉及以下主题，在该段落前单独一行加 ★ 标记：
- 酒类贸易管制、禁酒政策
- 托管制度、殖民治理
- 财政利益、税率
- 国际合作谈判
- 非洲殖民地管理

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【六、已知人物（无需加译者注）】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- V. Catastini / Vito Catastini → 维托·卡塔斯提尼，国际联盟托管科官员
- Sir Frederick D. Lugard → 弗雷德里克·卢加德爵士，常设托管委员会英国委员
- William Rappard / M. Rappard → 威廉·拉帕德，托管科主任
- Robert de Caix → 罗贝尔·德·凯，法国外交部亚洲司官员
- Dr. R. Hercod → 赫尔科德博士，国际反酒精局局长
- Sir Eric Drummond → 埃里克·德拉蒙德爵士，国际联盟秘书长
- Ken Harrada / K.H. → 哈拉达，国际联盟国际局科成员（日本籍）

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【七、质量要求】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. 同一段落内容必须合并，不得每行单独成段
2. 不得省略任何正文内容
3. 所有日期、编号、页码完整保留
4. 格式严格按文件类型规范排版，不得混乱
5. 页码标注必须精确，有双重页码时两者均须标注
"""

# ============================================================
# 合并OCR换行碎片
# ============================================================
def clean_ocr_text(text):
    """
    合并OCR导致的碎片换行
    规则：如果一行不以句号/问号/感叹号结尾，且下一行不是新段落，则合并
    """
    lines = text.split("\n")
    merged = []
    buffer = ""

    for line in lines:
        stripped = line.strip()

        # 空行 = 段落分隔
        if not stripped:
            if buffer:
                merged.append(buffer.strip())
                buffer = ""
            merged.append("")
            continue

        # 页码标记保留
        if stripped.startswith("===== PAGE"):
            if buffer:
                merged.append(buffer.strip())
                buffer = ""
            merged.append(stripped)
            continue

        # 判断是否应该合并到上一行
        if buffer:
            # 上一行以连字符结尾 = 单词被截断
            if buffer.endswith("-"):
                buffer = buffer[:-1] + stripped
            # 上一行以句末标点结尾 = 新段落
            elif buffer[-1] in ".!?:\"'»":
                merged.append(buffer.strip())
                buffer = stripped
            # 新行是大写开头且上一行较短 = 可能是新段落
            elif len(buffer) < 40 and stripped and stripped[0].isupper():
                merged.append(buffer.strip())
                buffer = stripped
            # 否则合并
            else:
                buffer += " " + stripped
        else:
            buffer = stripped

    if buffer:
        merged.append(buffer.strip())

    return "\n".join(merged)

# ============================================================
# 翻译单份文件
# ============================================================
def translate_document(text, doc_id, pages):
    # 先清理OCR换行
    cleaned_text = clean_ocr_text(text)

    print(f"  正在翻译... (原始 {len(text)} 字符 → 清理后 {len(cleaned_text)} 字符)")

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": (
                    f"请按规范翻译以下档案。\n"
                    f"原始页码范围：{pages}\n\n"
                    f"{cleaned_text}"
                )}
            ],
            max_tokens=8192,
            temperature=0.1
        )
        return response.choices[0].message.content

    except Exception as e:
        print(f"  ❌ 翻译失败：{e}")
        return None

# ============================================================
# 主流程
# ============================================================
def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with GROUPS_FILE.open("r", encoding="utf-8") as f:
        groups = list(csv.DictReader(f))

    print(f"共发现 {len(groups)} 份文件，开始翻译...\n")

    for group in groups:
        doc_id = int(group["document_id"])
        pages  = group["pages"].split("; ")
        start  = group["start_page"]
        end    = group["end_page"]

        print(f"【文件 {doc_id:03}】{start} → {end}  ({len(pages)} 页)")

        out_file = OUTPUT_DIR / f"doc_{doc_id:03}.txt"
        if out_file.exists():
            print(f"  ⏭ 已翻译，跳过\n")
            continue

        full_text = ""
        for page in pages:
            page_file = CHUNKS_DIR / page
            if page_file.exists():
                full_text += page_file.read_text(encoding="utf-8") + "\n\n"

        if len(full_text.strip()) < 50:
            print(f"  ⚠ 内容过少，跳过\n")
            continue

        result = translate_document(full_text, doc_id, ", ".join(pages))

        if result:
            out_file.write_text(result, encoding="utf-8")
            print(f"  ✅ 已保存 → {out_file}\n")
        else:
            print(f"  ❌ 翻译失败\n")

        time.sleep(1)

    print("全部完成！翻译结果在 output/translations/ 文件夹")

if __name__ == "__main__":
    main()
