import streamlit as st
from docx import Document

st.set_page_config(page_title="学位论文格式检测工具", layout="wide")

st.title("📄 学位论文格式检测工具（专业版）")

uploaded_file = st.file_uploader("请上传你的论文（Word文件）", type=["docx"])


def is_chinese(char):
    return '\u4e00' <= char <= '\u9fff'


def check_format(doc):
    paragraph_results = {}
    
    font_errors = 0
    size_errors = 0
    line_errors = 0
    indent_errors = 0

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        para_index = i + 1
        errors = set()

        for run in para.runs:
            font = run.font

            # 字体检测（中英文分开）
            for char in run.text:
                if is_chinese(char):
                    if font.name and "宋体" not in font.name:
                        errors.add("字体错误（中文应为宋体）")
                        font_errors += 1
                        break
                else:
                    if font.name and "Times New Roman" not in font.name:
                        errors.add("字体错误（英文/数字应为Times New Roman）")
                        font_errors += 1
                        break

            # 字号检测（小四 = 12pt）
            if font.size:
                if font.size.pt != 12:
                    errors.add("字号错误（应为小四号 12pt）")
                    size_errors += 1

        # 行距检测（20磅）
        if para.paragraph_format.line_spacing:
            if para.paragraph_format.line_spacing.pt != 20:
                errors.add("行距错误（应为20磅固定行距）")
                line_errors += 1
        else:
            errors.add("行距错误（未设置固定行距）")
            line_errors += 1

        # 首行缩进（约2字符 ≈ 21pt）
        if para.paragraph_format.first_line_indent:
            if para.paragraph_format.first_line_indent.pt < 20:
                errors.add("首行缩进错误（应为2字符）")
                indent_errors += 1
        else:
            errors.add("首行缩进错误（未设置）")
            indent_errors += 1

        if errors:
            paragraph_results[para_index] = {
                "text": text[:50],  # 显示前50字符方便定位
                "errors": list(errors)
            }

    summary = {
        "字体错误": font_errors,
        "字号错误": size_errors,
        "行距错误": line_errors,
        "缩进错误": indent_errors
    }

    return paragraph_results, summary


if uploaded_file:
    doc = Document(uploaded_file)

    results, summary = check_format(doc)

    st.subheader("📊 总览")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("字体错误", summary["字体错误"])
    col2.metric("字号错误", summary["字号错误"])
    col3.metric("行距错误", summary["行距错误"])
    col4.metric("缩进错误", summary["缩进错误"])

    total = sum(summary.values())

    if total == 0:
        st.success("🎉 未发现格式问题，格式完全正确！")
    else:
        st.error(f"❌ 共发现 {total} 个问题")

    st.divider()

    st.subheader("📍 问题段落定位（含解释）")

    for para, content in results.items():
        st.markdown(f"### ❌ 第{para}段（{len(content['errors'])}个问题）")

        st.caption(f"内容预览：{content['text']}")

        for err in content["errors"]:
            st.write("👉", err)

        st.divider()
