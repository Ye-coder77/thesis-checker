import streamlit as st
from docx import Document
import re

st.set_page_config(page_title="论文格式检测工具", layout="wide")

st.title("📄 学位论文格式检测工具（产品版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


def is_chinese(char):
    return '\u4e00' <= char <= '\u9fff'


def check_format(doc):
    paragraph_results = {}

    summary = {
        "字体错误": 0,
        "字号错误": 0,
        "行距错误": 0,
        "缩进错误": 0,
        "标题错误": 0,
        "图表错误": 0
    }

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        para_index = i + 1
        errors = set()

        # ========================
        # 标题检测
        # ========================
        if re.match(r'^第.+章', text):
            if not re.match(r'^第\d+章', text):
                errors.add("一级标题错误（应为第1章）")
                summary["标题错误"] += 1

        elif re.match(r'^\d+\.\d+\.\d+', text):
            pass

        elif re.match(r'^\d+\.\d+', text):
            pass

        elif re.match(r'^[一二三四五六七八九十]+、', text):
            errors.add("标题编号错误（不应使用中文编号）")
            summary["标题错误"] += 1

        # ========================
        # 图标题检测
        # ========================
        if text.startswith("图"):
            if not re.match(r'^图\d+\.\d+', text):
                errors.add("图编号错误（应为图1.1）")
                summary["图表错误"] += 1

            if i > 0 and doc.paragraphs[i - 1].text.strip():
                errors.add("图标题位置错误（应在图片下方）")
                summary["图表错误"] += 1

        # ========================
        # 表标题检测
        # ========================
        if text.startswith("表"):
            if not re.match(r'^表\d+\.\d+', text):
                errors.add("表编号错误（应为表1.1）")
                summary["图表错误"] += 1

            if i < len(doc.paragraphs) - 1 and doc.paragraphs[i + 1].text.strip():
                errors.add("表标题位置可能错误（应在表格上方）")
                summary["图表错误"] += 1

        # ========================
        # 正文检测
        # ========================
        for run in para.runs:
            font = run.font

            for char in run.text:
                if is_chinese(char):
                    if font.name and "宋体" not in font.name:
                        errors.add("字体错误（中文应为宋体）")
                        summary["字体错误"] += 1
                        break
                else:
                    if font.name and "Times New Roman" not in font.name:
                        errors.add("字体错误（英文应为Times New Roman）")
                        summary["字体错误"] += 1
                        break

            if font.size:
                if font.size.pt != 12:
                    errors.add("字号错误（应为小四 12pt）")
                    summary["字号错误"] += 1

        # 行距
        if para.paragraph_format.line_spacing:
            if para.paragraph_format.line_spacing.pt != 20:
                errors.add("行距错误（应为20磅）")
                summary["行距错误"] += 1
        else:
            errors.add("行距错误（未设置）")
            summary["行距错误"] += 1

        # 缩进
        if para.paragraph_format.first_line_indent:
            if para.paragraph_format.first_line_indent.pt < 20:
                errors.add("首行缩进错误（应为2字符）")
                summary["缩进错误"] += 1
        else:
            errors.add("首行缩进错误（未设置）")
            summary["缩进错误"] += 1

        if errors:
            paragraph_results[para_index] = {
                "text": text[:60],
                "errors": list(errors)
            }

    return paragraph_results, summary


# ========================
# 报告生成
# ========================
def generate_report(results, summary):
    doc = Document()

    doc.add_heading("论文格式检测报告", 0)

    total = sum(summary.values())
    doc.add_paragraph(f"总问题数：{total}")

    doc.add_heading("错误统计", level=1)
    for k, v in summary.items():
        doc.add_paragraph(f"{k}：{v}")

    doc.add_heading("详细问题", level=1)

    for para, content in results.items():
        doc.add_heading(f"第{para}段", level=2)
        doc.add_paragraph(f"内容预览：{content['text']}")

        for err in content["errors"]:
            doc.add_paragraph(f"- {err}")

    file_path = "论文检测报告.docx"
    doc.save(file_path)

    return file_path


# ========================
# 主程序
# ========================
if uploaded_file:
    doc = Document(uploaded_file)

    results, summary = check_format(doc)

    st.subheader("📊 检测总览")

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("字体错误", summary["字体错误"])
    col2.metric("字号错误", summary["字号错误"])
    col3.metric("行距错误", summary["行距错误"])
    col4.metric("缩进错误", summary["缩进错误"])
    col5.metric("标题错误", summary["标题错误"])
    col6.metric("图表错误", summary["图表错误"])

    total = sum(summary.values())

    if total == 0:
        st.success("🎉 未发现问题，格式完全正确！")
    else:
        st.error(f"❌ 共发现 {total} 个问题")

    st.divider()

    st.subheader("📍 问题定位")

    for para, content in results.items():
        st.markdown(f"### ❌ 第{para}段")
        st.caption(f"内容预览：{content['text']}")

        for err in content["errors"]:
            st.write("👉", err)

        st.divider()

    # ========================
    # 导出报告
    # ========================
    if results:
        if st.button("📄 生成检测报告"):
            report_path = generate_report(results, summary)

            with open(report_path, "rb") as f:
                st.download_button(
                    label="📥 下载检测报告",
                    data=f,
                    file_name="论文检测报告.docx"
                )
