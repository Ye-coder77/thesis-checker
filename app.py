import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具（终极稳定版）", layout="wide")
st.title("📄 学位论文格式检测工具（终极稳定版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 分类函数（完全修复版）
# ===========================

def classify(text):
    text = text.strip()

    # ===== 摘要 =====
    if text in ["摘要", "摘 要"]:
        return "cn_abstract_title"

    if text == "ABSTRACT":
        return "en_abstract_title"

    if text.startswith("关键词"):
        return "cn_keywords"

    if text.startswith("KEY WORDS"):
        return "en_keywords"

    # ===== 标题（顺序必须这样）=====
    if re.match(r'^\d+\.\d+\.\d+', text):
        return "title3"

    if re.match(r'^\d+\.\d+', text):
        return "title2"

    if re.match(r'^第\d+章', text):
        return "title1"

    if re.match(r'^\d+\.\s*', text):
        return "wrong_title1"

    if re.match(r'^（\d+）', text):
        return "title4"

    # ===== 图表 =====
    if text.startswith("图"):
        return "figure"

    if text.startswith("表"):
        return "table"

    # ===== 正文 =====
    if re.search(r'[a-zA-Z]', text):
        return "en_body"

    return "cn_body"


# ===========================
# 核心检测
# ===========================

def check(doc):
    results = {}

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        ptype = classify(text)
        pf = para.paragraph_format
        errors = []

        # ========= 中文摘要标题 =========
        if ptype == "cn_abstract_title":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("摘要标题应居中")
            if " " not in text and "摘 要" not in text:
                errors.append("摘要两字中间应有空格")

        # ========= 英文摘要标题 =========
        if ptype == "en_abstract_title":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("ABSTRACT应居中")
            if not any(run.bold for run in para.runs):
                errors.append("ABSTRACT应加粗")

        # ========= 中文正文 =========
        if ptype == "cn_body":
            if pf.first_line_indent is None:
                errors.append("中文正文应首行缩进2字符")

        # ========= 英文正文 =========
        if ptype == "en_body":
            for run in para.runs:
                if run.font.name and "Times" not in run.font.name:
                    errors.append("英文正文应为Times New Roman")

        # ========= 一级标题 =========
        if ptype == "title1":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("一级标题必须居中")

        # ========= 错误一级标题 =========
        if ptype == "wrong_title1":
            errors.append("一级标题必须为“第X章”，不能写成“2. 标题”")

        # ========= 二级标题 =========
        if ptype == "title2":
            if para.alignment not in [None, WD_ALIGN_PARAGRAPH.LEFT]:
                errors.append("二级标题应左对齐")

        # ========= 三级标题 =========
        if ptype == "title3":
            if para.alignment not in [None, WD_ALIGN_PARAGRAPH.LEFT]:
                errors.append("三级标题应左对齐")

        # ========= 四级标题 =========
        if ptype == "title4":
            if not any(run.bold for run in para.runs):
                errors.append("四级标题应加粗")

        # ========= 图 =========
        if ptype == "figure":
            if not re.match(r'^图\d+[-\.]\d+', text):
                errors.append("图编号错误，应为图1-1或图1.1")

        # ========= 表 =========
        if ptype == "table":
            if not re.match(r'^表\d+\.\d+', text):
                errors.append("表编号错误，应为表1.1")

        # ========= 关键词 =========
        if ptype == "cn_keywords":
            if not ("；" in text or "，" in text):
                errors.append("关键词应使用中文分号或逗号")

        if ptype == "en_keywords":
            if not (";" in text or "," in text):
                errors.append("KEY WORDS应使用英文分号或逗号")

        if errors:
            results[i + 1] = {
                "text": text,
                "type": ptype,
                "errors": list(set(errors))
            }

    return results


# ===========================
# UI（分类展示）
# ===========================

if uploaded_file:
    doc = Document(uploaded_file)
    results = check(doc)

    categories = {
        "cn_abstract_title": "摘要",
        "en_abstract_title": "摘要",
        "cn_keywords": "摘要",
        "en_keywords": "摘要",
        "cn_body": "正文",
        "en_body": "正文",
        "title1": "标题",
        "title2": "标题",
        "title3": "标题",
        "title4": "标题",
        "figure": "图表",
        "table": "图表",
        "wrong_title1": "标题"
    }

    grouped = {"摘要": [], "正文": [], "标题": [], "图表": []}

    for para, content in results.items():
        group = categories.get(content["type"], "正文")
        grouped[group].append((para, content))

    st.error(f"⚠️ 共发现 {len(results)} 处问题")

    for group_name, items in grouped.items():
        if items:
            with st.expander(f"📌 {group_name}部分（{len(items)}项）"):
                for para, content in items:
                    st.markdown(f"**第{para}段：{content['text']}**")
                    for err in content["errors"]:
                        st.write(f"👉 {err}")
