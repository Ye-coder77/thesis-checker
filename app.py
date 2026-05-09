import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具（完整版）", layout="wide")
st.title("📄 学位论文格式检测工具（完整版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 工具函数
# ===========================
def get_font(run, para):
    return run.font.name or para.style.font.name

def get_size(run, para):
    if run.font.size:
        return run.font.size.pt
    if para.style.font.size:
        return para.style.font.size.pt
    return None

def get_spacing_before(para):
    if para.paragraph_format.space_before:
        return para.paragraph_format.space_before.pt
    if para.style and para.style.paragraph_format.space_before:
        return para.style.paragraph_format.space_before.pt
    return None

def get_spacing_after(para):
    if para.paragraph_format.space_after:
        return para.paragraph_format.space_after.pt
    if para.style and para.style.paragraph_format.space_after:
        return para.style.paragraph_format.space_after.pt
    return None

def get_line_spacing(para):
    pf = para.paragraph_format
    if pf.line_spacing_rule:
        return pf.line_spacing_rule
    if para.style and para.style.paragraph_format.line_spacing_rule:
        return para.style.paragraph_format.line_spacing_rule
    return None


# ===========================
# 分类
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

    if text == "目录":
        return "toc_title"
    if re.match(r'^\d+(\.\d+)*\s+.*\s+\d+$', text):
        return "toc_item"

    if text.startswith("图"):
        return "figure"
    if text.startswith("表"):
        return "table"

    if re.search(r'[a-zA-Z]', text):
        return "en_body"
    return "cn_body"


# ===========================
# 检测
# ===========================
def check(doc):
    results = {}

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        ptype = classify(text)
        errors = []
        pf = para.paragraph_format
        prev = doc.paragraphs[i-1].text.strip() if i > 0 else ""

        # ===== 一级标题 =====
        if ptype == "title1":
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                errors.append("一级标题应居中")

            before = get_spacing_before(para)
            after = get_spacing_after(para)

            if before is None or before < 12:
                errors.append("一级标题段前应空1行")

            if after is None or after < 12:
                errors.append("一级标题段后应空1行")

        # ===== 错误一级标题 =====
        if ptype == "wrong_title1":
            errors.append("一级标题必须为“第X章”")

        # ===== 中文正文 =====
        if ptype == "cn_body":
            for run in para.runs:
                if get_font(run, para) and "宋体" not in get_font(run, para):
                    errors.append("正文应为宋体")
                size = get_size(run, para)
                if size and abs(size - 12) > 0.5:
                    errors.append("正文应为小四（12pt）")

            if get_line_spacing(para) != 1:
                errors.append("正文应为1.5倍行距")

            if pf.first_line_indent is None:
                errors.append("正文应首行缩进2字符")

        # ===== 图 =====
        if ptype == "figure":
            if not re.match(r'^图\d+[-\.]\d+', text):
                errors.append("图编号错误")

        # ===== 表 =====
        if ptype == "table":
            if not re.match(r'^表\d+\.\d+', text):
                errors.append("表编号错误")

        if errors:
            results[i] = {
                "text": text,
                "errors": list(set(errors))
            }

    return results


# ===========================
# 生成报告
# ===========================
def generate_report(results):
    doc = Document()
    doc.add_heading("论文格式检测报告", 0)

    for i, content in results.items():
        doc.add_paragraph(f"第{i+1}段：{content['text']}")
        for err in content["errors"]:
            doc.add_paragraph(f"- {err}")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name


# ===========================
# 标注论文
# ===========================
def generate_marked_doc(original_doc, results):
    doc = original_doc

    for i, para in enumerate(doc.paragraphs):
        if i in results:
            # 高亮
            for run in para.runs:
                run.font.highlight_color = 7

            # 添加备注
            note = "\n【格式问题】\n"
            for err in results[i]["errors"]:
                note += f"- {err}\n"

            para.add_run(note)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name


# ===========================
# UI
# ===========================
if uploaded_file:
    doc = Document(uploaded_file)
    results = check(doc)

    st.error(f"⚠️ 共发现 {len(results)} 处问题")

    for i, content in results.items():
        with st.expander(f"第{i+1}段：{content['text'][:30]}"):
            for err in content["errors"]:
                st.write(f"👉 {err}")

    # 下载报告
    report_path = generate_report(results)
    with open(report_path, "rb") as f:
        st.download_button("📥 下载检测报告", f, file_name="检测报告.docx")

    # 标注论文
    if st.button("🖍 生成标注版论文"):
        doc2 = Document(uploaded_file)
        marked_path = generate_marked_doc(doc2, results)

        with open(marked_path, "rb") as f:
            st.download_button("📥 下载标注论文", f, file_name="标注论文.docx")
