import streamlit as st
from docx import Document
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具（专业版）", layout="wide")

st.title("📄 学位论文格式检测工具（专业版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 工具函数
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
# 检测函数（稳定版）
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

        # 行距
        ls = pf.line_spacing
        if ls is None:
            errors.append("未设置行距（应为20磅）")
            summary["行距错误"] += 1
        elif hasattr(ls, "pt") and ls.pt != 20:
            errors.append("行距不是20磅")
            summary["行距错误"] += 1

        # 缩进
        if pf.first_line_indent is None:
            errors.append("未设置首行缩进")
            summary["缩进错误"] += 1

        # 标题检测
        if is_title_level1(text):
            current_chapter += 1
            if "第一章" in text:
                errors.append("一级标题不能用中文编号")
                summary["标题错误"] += 1

        if is_title_level2(text):
            if not text.startswith(f"{current_chapter}."):
                errors.append("二级标题编号错误")
                summary["标题错误"] += 1

        if is_title_level3(text):
            if not text.startswith(f"{current_chapter}."):
                errors.append("三级标题编号错误")
                summary["标题错误"] += 1

        # 图表
        if is_figure(text):
            if not check_figure_number(text, current_chapter):
                errors.append("图编号应为 图X.X")
                summary["图表错误"] += 1

        if is_table(text):
            if not check_table_number(text, current_chapter):
                errors.append("表编号应为 表X.X")
                summary["图表错误"] += 1

        # 字体字号
        for run in para.runs:
            font = run.font
            name = font.name
            size = font.size

            if name:
                if is_chinese(run.text) and "宋体" not in name:
                    errors.append("中文应为宋体")
                    summary["字体错误"] += 1

                if is_english_or_number(run.text) and "Times" not in name:
                    errors.append("英文应为Times New Roman")
                    summary["字体错误"] += 1

            if size and size.pt != 12:
                errors.append("正文应为12pt（小四）")
                summary["字号错误"] += 1

        if errors:
            results[i] = {
                "text": text,
                "errors": list(set(errors))
            }

    return results, summary


# ===========================
# Word报告
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
        doc.add_paragraph(f"第{para+1}段：{content['text']}")
        for err in content["errors"]:
            doc.add_paragraph(f" - {err}")

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp.name)
    return temp.name


# ===========================
# 主程序
# ===========================

if uploaded_file:
    doc = Document(uploaded_file)

    results, summary = check_format(doc)

    st.success("检测完成")

    total = sum(summary.values())
    st.error(f"共发现 {total} 个问题")

    # 初始化定位
    if "focus_para" not in st.session_state:
        st.session_state["focus_para"] = None

    col1, col2 = st.columns([1, 2])

    # ===== 左侧：问题列表（按钮版）=====
    with col1:
        st.markdown("## ❌ 问题列表（点击定位）")

        for para, content in results.items():
            if st.button(f"👉 第{para+1}段", key=f"btn_{para}"):
                st.session_state["focus_para"] = para

            for err in content["errors"]:
                st.write(f" - {err}")

    # ===== 右侧：原文 =====
    with col2:
        st.markdown("## 📄 原文（定位 + 高亮）")

        focus_para = st.session_state.get("focus_para")

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()

            if i == focus_para:
                st.markdown(
                    f"<div style='background-color:#ffcccc;padding:12px;border-radius:8px'>"
                    f"<b>👉 当前定位：第{i+1}段</b><br>{text}</div>",
                    unsafe_allow_html=True
                )
            elif i in results:
                st.markdown(
                    f"<div style='background-color:#ffe6e6;padding:8px;border-radius:5px'>"
                    f"<b>第{i+1}段：</b>{text}</div>",
                    unsafe_allow_html=True
                )
            else:
                st.write(f"第{i+1}段：{text}")

    # ===== 下载报告 =====
    report_path = generate_report(results, summary)

    with open(report_path, "rb") as f:
        st.download_button(
            "📥 下载Word检测报告",
            f,
            file_name="论文检测报告.docx"
        )
