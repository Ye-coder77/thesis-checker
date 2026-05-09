import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具", layout="wide")
st.title("📄 学位论文格式检测工具")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 分类函数（关键）
# ===========================

def classify_paragraph(text):
    """
    返回：'title1' / 'title2' / 'title3' / 'figure' / 'table' / 'body'
    """
    text = text.strip()

    # 一级标题（第X章）
    if re.match(r'^第\s*\d+\s*章', text):
        return "title1"

    # 一级标题（2. xxx）
    if re.match(r'^\d+\.\s+', text):
        return "title1"

    # 二级标题（1.1）
    if re.match(r'^\d+\.\d+\s*', text):
        return "title2"

    # 三级标题（1.1.1）
    if re.match(r'^\d+\.\d+\.\d+\s*', text):
        return "title3"

    # 图表
    if text.startswith("图"):
        return "figure"

    if text.startswith("表"):
        return "table"

    return "body"


def is_chinese(text):
    return any('\u4e00' <= c <= '\u9fff' for c in text)


def is_english_or_number(text):
    return any(c.isascii() and (c.isalpha() or c.isdigit()) for c in text)


# ===========================
# 检测函数（分类型规则）
# ===========================

def check_format(doc):
    results = {}
    summary = {
        "正文问题": 0,
        "标题问题": 0,
        "图表问题": 0,
        "字体问题": 0
    }

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        ptype = classify_paragraph(text)
        pf = para.paragraph_format
        errors = []

        # ================= 正文 =================
        if ptype == "body":
            # 行距
            ls = pf.line_spacing
            if ls is None:
                errors.append("正文未设置行距（应为20磅）")
                summary["正文问题"] += 1
            elif hasattr(ls, "pt") and ls.pt != 20:
                errors.append("正文行距应为20磅")
                summary["正文问题"] += 1

            # 首行缩进
            if pf.first_line_indent is None:
                errors.append("正文应首行缩进")
                summary["正文问题"] += 1

        # ================= 一级标题 =================
        elif ptype == "title1":
            # 居中
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("一级标题应居中")
                summary["标题问题"] += 1

            # 字体（黑体）
            for run in para.runs:
                if run.font.name and "黑体" not in run.font.name:
                    errors.append("一级标题应为黑体")
                    summary["标题问题"] += 1
                    break

        # ================= 二级标题 =================
        elif ptype == "title2":
            if para.alignment not in [None, WD_ALIGN_PARAGRAPH.LEFT]:
                errors.append("二级标题应左对齐")
                summary["标题问题"] += 1

            if not any(run.bold for run in para.runs):
                errors.append("二级标题应加粗")
                summary["标题问题"] += 1

        # ================= 三级标题 =================
        elif ptype == "title3":
            if not any(run.bold for run in para.runs):
                errors.append("三级标题应加粗")
                summary["标题问题"] += 1

        # ================= 图 =================
        elif ptype == "figure":
            errors.append("请确认图标题在图片下方")
            summary["图表问题"] += 1

        # ================= 表 =================
        elif ptype == "table":
            errors.append("请确认表标题在表格上方")
            summary["图表问题"] += 1

        # ================= 字体（通用） =================
        for run in para.runs:
            name = run.font.name
            size = run.font.size

            if name:
                if is_chinese(run.text) and "宋体" not in name:
                    errors.append("中文应为宋体")
                    summary["字体问题"] += 1

                if is_english_or_number(run.text) and "Times" not in name:
                    errors.append("英文应为Times New Roman")
                    summary["字体问题"] += 1

            if size and size.pt != 12 and ptype == "body":
                errors.append("正文应为12pt（小四）")
                summary["字体问题"] += 1

        if errors:
            results[i + 1] = {
                "text": text,
                "errors": list(set(errors))
            }

    return results, summary


# ===========================
# 标注论文
# ===========================

def highlight_doc(doc, results):
    for i, para in enumerate(doc.paragraphs):
        if i + 1 in results:
            for run in para.runs:
                run.font.highlight_color = 7

            note = "\n【检测问题】\n"
            for err in results[i + 1]["errors"]:
                note += f"- {err}\n"

            para.add_run(note)

    return doc


# ===========================
# 报告
# ===========================

def generate_report(results, summary):
    doc = Document()

    doc.add_heading("论文格式检测报告", 0)

    total = sum(summary.values())
    doc.add_paragraph(f"总问题数：{total}")

    for k, v in summary.items():
        doc.add_paragraph(f"{k}：{v}")

    doc.add_heading("详细问题", 1)

    for para, content in results.items():
        doc.add_paragraph(f"第{para}段：{content['text']}")
        for err in content["errors"]:
            doc.add_paragraph(f"- {err}")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name


# ===========================
# 主程序
# ===========================

if uploaded_file:
    doc = Document(uploaded_file)

    results, summary = check_format(doc)

    st.subheader("📊 检测结果")

    total = sum(summary.values())
    st.error(f"共发现 {total} 个问题")

    col1, col2, col3 = st.columns(3)
    col1.metric("正文", summary["正文问题"])
    col2.metric("标题", summary["标题问题"])
    col3.metric("图表", summary["图表问题"])

    st.divider()

    st.subheader("📍 问题详情")

    for para, content in results.items():
        with st.expander(f"第{para}段：{content['text'][:30]}"):
            for err in content["errors"]:
                st.write(f"👉 {err}")

    # 下载报告
    report_path = generate_report(results, summary)
    with open(report_path, "rb") as f:
        st.download_button("📥 下载检测报告", f, file_name="检测报告.docx")

    # 标注论文
    if st.button("🖍 生成标注版论文"):
        doc_marked = Document(uploaded_file)
        doc_marked = highlight_doc(doc_marked, results)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc_marked.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button("📥 下载标注论文", f, file_name="标注论文.docx")
