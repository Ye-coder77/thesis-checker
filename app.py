import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具（专业版）", layout="wide")

st.title("📄 学位论文格式检测工具（测试版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 🔍 工具函数
# ===========================

def is_chinese(text):
    return any('\u4e00' <= c <= '\u9fff' for c in text)


def is_english_or_number(text):
    return any(c.isascii() and (c.isalpha() or c.isdigit()) for c in text)


def is_title_level1(text):
    return re.match(r'^第?\d+章', text)


def is_title_level2(text):
    return re.match(r'^\d+\.\d+', text)


def is_title_level3(text):
    return re.match(r'^\d+\.\d+\.\d+', text)


def is_figure(text):
    return text.startswith("图")


def is_table(text):
    return text.startswith("表")


def check_figure_number(text, chapter):
    return re.match(rf'^图{chapter}\.\d+', text)


def check_table_number(text, chapter):
    return re.match(rf'^表{chapter}\.\d+', text)


# ===========================
# 🔍 主检测函数（已修复所有None问题）
# ===========================

def check_format(doc):
    results = {}
    summary = {
        "字体错误": 0,
        "字号错误": 0,
        "行距错误": 0,
        "缩进错误": 0,
        "标题错误": 0,
        "图表错误": 0
    }

    current_chapter = 1

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        errors = []

        pf = para.paragraph_format

        # ================= 行距检测 =================
        ls = pf.line_spacing
        if ls is None:
            errors.append("未设置行距（应为20磅）")
            summary["行距错误"] += 1
        elif hasattr(ls, "pt") and ls.pt != 20:
            errors.append("行距不是20磅")
            summary["行距错误"] += 1

        # ================= 缩进检测 =================
        if pf.first_line_indent is None:
            errors.append("未设置首行缩进（应为2字符）")
            summary["缩进错误"] += 1

        # ================= 标题检测 =================
        if is_title_level1(text):
            current_chapter += 1
            if "第一章" in text:
                errors.append("一级标题不能用中文编号，应使用阿拉伯数字")
                summary["标题错误"] += 1

        if is_title_level2(text):
            if not text.startswith(f"{current_chapter}."):
                errors.append("二级标题编号不符合章编号")
                summary["标题错误"] += 1

        if is_title_level3(text):
            if not text.startswith(f"{current_chapter}."):
                errors.append("三级标题编号不符合章编号")
                summary["标题错误"] += 1

        # ================= 图表检测 =================
        if is_figure(text):
            if not check_figure_number(text, current_chapter):
                errors.append("图编号错误，应为 图X.X")
                summary["图表错误"] += 1

        if is_table(text):
            if not check_table_number(text, current_chapter):
                errors.append("表编号错误，应为 表X.X")
                summary["图表错误"] += 1

        # ================= 字体检测 =================
        for run in para.runs:
            font = run.font
            name = font.name

            if name is None:
                continue

            if is_chinese(run.text):
                if "宋体" not in name:
                    errors.append("中文应为宋体")
                    summary["字体错误"] += 1

            if is_english_or_number(run.text):
                if "Times" not in name:
                    errors.append("英文/数字应为Times New Roman")
                    summary["字体错误"] += 1

            # ================= 字号检测 =================
            size = font.size
            if size is None:
                continue

            if size.pt != 12:
                errors.append("正文应为12pt（小四）")
                summary["字号错误"] += 1

        if errors:
            results[i + 1] = {
                "text": text,
                "errors": list(set(errors))
            }

    return results, summary


# ===========================
# 📄 生成 Word 报告
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
            doc.add_paragraph(f" - {err}")

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_file.name)
    return temp_file.name


# ===========================
# 🚀 主逻辑
# ===========================

if uploaded_file:
    doc = Document(uploaded_file)

    results, summary = check_format(doc)

    st.success("检测完成")

    total = sum(summary.values())

    st.write(f"❌ 共发现 {total} 个问题")

    col1, col2 = st.columns(2)

    with col1:
        for k, v in summary.items():
            st.write(f"{k}：{v}")

    with col2:
        st.write("👇 详细问题")

    for para, content in results.items():
        st.error(f"第{para}段：{content['text']}")
        for err in content["errors"]:
            st.write(f"👉 {err}")

    # ================= 下载报告 =================
    report_path = generate_report(results, summary)

    with open(report_path, "rb") as f:
        st.download_button(
            "📥 下载Word检测报告",
            f,
            file_name="论文检测报告.docx"
        )
