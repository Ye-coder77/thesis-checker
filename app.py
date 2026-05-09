import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具（规范版）", layout="wide")
st.title("📄 学位论文格式检测工具（规范版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 分类函数
# ===========================

def classify(text):
    text = text.strip()

    if text == "摘 要" or text == "摘要":
        return "cn_abstract_title"

    if text == "ABSTRACT":
        return "en_abstract_title"

    if text.startswith("关键词"):
        return "cn_keywords"

    if text.startswith("KEY WORDS"):
        return "en_keywords"

    if re.match(r'^第\s*\d+\s*章', text):
        return "title1"

    if re.match(r'^\d+\.\d+\.\d+', text):
        return "title3"

    if re.match(r'^\d+\.\d+', text):
        return "title2"

    if re.match(r'^（\d+）', text):
        return "title4"

    if text.startswith("图"):
        return "figure"

    if text.startswith("表"):
        return "table"

    if re.search(r'[a-zA-Z]', text):
        return "en_body"

    return "cn_body"


def is_chinese(text):
    return any('\u4e00' <= c <= '\u9fff' for c in text)


# ===========================
# 核心检测
# ===========================

def check(doc):
    results = {}
    summary = {
        "摘要问题": 0,
        "正文问题": 0,
        "标题问题": 0,
        "图表问题": 0,
        "字体问题": 0
    }

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        ptype = classify(text)
        pf = para.paragraph_format
        errors = []

        # ===== 中文摘要标题 =====
        if ptype == "cn_abstract_title":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("摘要标题应居中")
            for run in para.runs:
                if run.font.name and "黑体" not in run.font.name:
                    errors.append("摘要标题应为黑体")
                    break

        # ===== 英文摘要标题 =====
        if ptype == "en_abstract_title":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("ABSTRACT应居中")
            for run in para.runs:
                if run.font.name and "Times" not in run.font.name:
                    errors.append("ABSTRACT应为Times New Roman")
                    break

        # ===== 中文正文 =====
        if ptype == "cn_body":
            if pf.first_line_indent is None:
                errors.append("中文正文应首行缩进2字符")

            for run in para.runs:
                if run.font.name and "宋体" not in run.font.name:
                    errors.append("中文正文应为宋体")
                    break

        # ===== 英文正文 =====
        if ptype == "en_body":
            for run in para.runs:
                if run.font.name and "Times" not in run.font.name:
                    errors.append("英文正文应为Times New Roman")
                    break

        # ===== 一级标题 =====
        if ptype == "title1":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("一级标题必须居中")

            if not re.match(r'^第\d+章', text):
                errors.append("一级标题格式错误，应为“第X章”")

        # ===== 二级标题 =====
        if ptype == "title2":
            if para.alignment not in [None, WD_ALIGN_PARAGRAPH.LEFT]:
                errors.append("二级标题应左对齐")

        # ===== 三级标题 =====
        if ptype == "title3":
            if para.alignment not in [None, WD_ALIGN_PARAGRAPH.LEFT]:
                errors.append("三级标题应左对齐")

        # ===== 四级标题 =====
        if ptype == "title4":
            if not any(run.bold for run in para.runs):
                errors.append("四级标题应加粗")

        # ===== 图 =====
        if ptype == "figure":
            if not re.match(r'^图\d+[-\.]\d+', text):
                errors.append("图编号格式错误，应为图1-1或图1.1")

        # ===== 表 =====
        if ptype == "table":
            if not re.match(r'^表\d+\.\d+', text):
                errors.append("表编号格式错误，应为表1.1")

        # ===== 关键词 =====
        if ptype == "cn_keywords":
            if "；" not in text and "，" not in text:
                errors.append("关键词应使用分号或逗号分隔")

        if ptype == "en_keywords":
            if ";" not in text and "," not in text:
                errors.append("KEY WORDS应使用英文分号或逗号")

        if errors:
            results[i + 1] = {
                "text": text,
                "errors": errors
            }

    return results, summary


# ===========================
# 标注
# ===========================

def highlight(doc, results):
    for i, para in enumerate(doc.paragraphs):
        if i + 1 in results:
            for run in para.runs:
                run.font.highlight_color = 7

            note = "\n【格式问题】\n"
            for err in results[i + 1]["errors"]:
                note += f"- {err}\n"

            para.add_run(note)

    return doc


# ===========================
# 报告
# ===========================

def report(results):
    doc = Document()
    doc.add_heading("论文格式检测报告", 0)

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
    results, summary = check(doc)

    st.error(f"共发现 {len(results)} 处问题")

    for para, content in results.items():
        with st.expander(f"第{para}段：{content['text'][:30]}"):
            for err in content["errors"]:
                st.write(f"👉 {err}")

    # 报告
    path = report(results)
    with open(path, "rb") as f:
        st.download_button("📥 下载检测报告", f, file_name="检测报告.docx")

    # 标注
    if st.button("🖍 生成标注版论文"):
        doc2 = Document(uploaded_file)
        doc2 = highlight(doc2, results)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc2.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button("📥 下载标注论文", f, file_name="标注论文.docx")
