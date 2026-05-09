import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具（专业版）", layout="wide")
st.title("📄 学位论文格式检测工具（专业版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 分类函数
# ===========================

def classify(text):
    text = text.strip()

    if text in ["摘要", "摘 要"]:
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

        # ===== 英文摘要标题 =====
        if ptype == "en_abstract_title":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("ABSTRACT应居中")
            for run in para.runs:
                if run.font.name and "Times" not in run.font.name:
                    errors.append("ABSTRACT应为Times New Roman")

        # ===== 中文正文 =====
        if ptype == "cn_body":
            if pf.first_line_indent is None:
                errors.append("中文正文应首行缩进")

            for run in para.runs:
                if run.font.name and "宋体" not in run.font.name:
                    errors.append("中文正文应为宋体")

        # ===== 英文正文 =====
        if ptype == "en_body":
            for run in para.runs:
                if run.font.name and "Times" not in run.font.name:
                    errors.append("英文正文应为Times New Roman")

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
                errors.append("图编号格式错误（图1-1 或 图1.1）")

        # ===== 表 =====
        if ptype == "table":
            if not re.match(r'^表\d+\.\d+', text):
                errors.append("表编号格式错误（表1.1）")

        # ===== 关键词 =====
        if ptype == "cn_keywords":
            if "；" not in text and "，" not in text:
                errors.append("关键词应使用中文分号或逗号")

        if ptype == "en_keywords":
            if ";" not in text and "," not in text:
                errors.append("KEY WORDS应使用英文分号或逗号")

        if errors:
            results[i + 1] = {
                "text": text,
                "type": ptype,
                "errors": list(set(errors))
            }

    return results


# ===========================
# 标注论文
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

def generate_report(results):
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
# UI 展示（专业版）
# ===========================

if uploaded_file:
    doc = Document(uploaded_file)
    results = check(doc)

    total = len(results)

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
        "table": "图表"
    }

    grouped = {"摘要": [], "正文": [], "标题": [], "图表": []}

    for para, content in results.items():
        group = categories.get(content["type"], "正文")
        grouped[group].append((para, content))

    # ===== 总览 =====
    st.markdown("## 📊 检测总览")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("摘要问题", len(grouped["摘要"]))
    col2.metric("正文问题", len(grouped["正文"]))
    col3.metric("标题问题", len(grouped["标题"]))
    col4.metric("图表问题", len(grouped["图表"]))

    st.error(f"⚠️ 共发现 {total} 处问题")

    st.divider()

    # ===== 卡片展示 =====
    def render_card(para, content):
        st.markdown(f"""
        <div style="
            border:1px solid #eee;
            border-radius:10px;
            padding:12px;
            margin-bottom:10px;
            background:#fafafa;
        ">
            <b>第{para}段：</b>{content['text'][:80]}<br>
            <span style="color:#d9534f;">⚠️ {len(content['errors'])}个问题</span>
        </div>
        """, unsafe_allow_html=True)

        for err in content["errors"]:
            st.write(f"👉 {err}")

    # ===== 分类展示 =====
    for group_name, items in grouped.items():
        if items:
            with st.expander(f"📌 {group_name}部分（{len(items)}项问题）"):
                for para, content in items:
                    render_card(para, content)

    # ===== 下载报告 =====
    report_path = generate_report(results)
    with open(report_path, "rb") as f:
        st.download_button("📥 下载检测报告", f, file_name="检测报告.docx")

    # ===== 标注 =====
    if st.button("🖍 生成标注版论文"):
        doc2 = Document(uploaded_file)
        doc2 = highlight(doc2, results)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc2.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button("📥 下载标注论文", f, file_name="标注论文.docx")
