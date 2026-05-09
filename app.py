import streamlit as st
from docx import Document
import re

st.set_page_config(page_title="论文格式检测工具", layout="wide")

st.title("📄 学位论文格式检测工具（终极版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


def is_chinese(char):
    return '\u4e00' <= char <= '\u9fff'


def is_english(char):
    return bool(re.match(r'[A-Za-z]', char))


def check_format(doc):
    results = {}

    summary = {
        "标题错误": 0,
        "字体错误": 0,
        "字号错误": 0,
        "行距错误": 0,
        "缩进错误": 0
    }

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        idx = i + 1
        errors = set()

        # ========================
        # 标题层级识别
        # ========================

        # 一级标题
        if re.match(r'^第\d+章', text):
            for run in para.runs:
                # 字体
                if run.font.name and "黑体" not in run.font.name:
                    errors.add("一级标题字体应为黑体")
                    summary["标题错误"] += 1

                # 字号（允许16或14）
                if run.font.size:
                    if run.font.size.pt not in [16, 14]:
                        errors.add("一级标题字号应为小三或四号")
                        summary["标题错误"] += 1

            # 对齐（允许居中或左）
            if para.paragraph_format.alignment not in [None, 0, 1]:
                errors.add("一级标题对齐错误")
                summary["标题错误"] += 1

        # 二级标题
        elif re.match(r'^\d+\.\d+', text):
            for run in para.runs:
                # 字体
                if run.font.name:
                    if ("黑体" not in run.font.name and
                        not (run.font.name == "宋体" and run.bold)):
                        errors.add("二级标题字体应为黑体或宋体加粗")
                        summary["标题错误"] += 1

                # 字号
                if run.font.size:
                    if run.font.size.pt not in [14, 12]:
                        errors.add("二级标题字号应为四号或小四")
                        summary["标题错误"] += 1

        # 三级标题
        elif re.match(r'^\d+\.\d+\.\d+', text):
            for run in para.runs:
                if run.font.name:
                    if ("黑体" not in run.font.name and
                        not (run.font.name == "宋体" and run.bold)):
                        errors.add("三级标题字体应为黑体或宋体加粗")
                        summary["标题错误"] += 1

                if run.font.size:
                    if run.font.size.pt not in [12, 10.5]:
                        errors.add("三级标题字号应为小四或五号")
                        summary["标题错误"] += 1

        # 四级标题（简单规则）
        elif re.match(r'^\(\d+\)', text):
            for run in para.runs:
                if run.font.size:
                    if run.font.size.pt != 10.5:
                        errors.add("四级标题字号应为五号")
                        summary["标题错误"] += 1

        # 中文编号错误
        elif re.match(r'^[一二三四五六七八九十]+、', text):
            errors.add("标题编号错误（不应使用中文编号）")
            summary["标题错误"] += 1

        # ========================
        # 正文字体检测
        # ========================
        for run in para.runs:
            for char in run.text:
                if is_chinese(char):
                    if run.font.name and "宋体" not in run.font.name:
                        errors.add("中文应为宋体")
                        summary["字体错误"] += 1
                        break

                elif is_english(char):
                    if run.font.name and "Times New Roman" not in run.font.name:
                        errors.add("英文应为Times New Roman")
                        summary["字体错误"] += 1
                        break

            # 字号
            if run.font.size:
                if run.font.size.pt != 12:
                    errors.add("正文字号应为小四（12pt）")
                    summary["字号错误"] += 1

        # 行距
        if para.paragraph_format.line_spacing:
            if para.paragraph_format.line_spacing.pt != 20:
                errors.add("行距应为20磅")
                summary["行距错误"] += 1
        else:
            errors.add("未设置行距")
            summary["行距错误"] += 1

        # 缩进
        if para.paragraph_format.first_line_indent:
            if para.paragraph_format.first_line_indent.pt < 20:
                errors.add("首行缩进应为2字符")
                summary["缩进错误"] += 1
        else:
            errors.add("未设置首行缩进")
            summary["缩进错误"] += 1

        if errors:
            results[idx] = {
                "text": text[:60],
                "errors": list(errors)
            }

    return results, summary


# ========================
# 主程序
# ========================
if uploaded_file:
    doc = Document(uploaded_file)

    results, summary = check_format(doc)

    st.subheader("📊 总览")

    cols = st.columns(5)
    for i, (k, v) in enumerate(summary.items()):
        cols[i].metric(k, v)

    total = sum(summary.values())

    if total == 0:
        st.success("🎉 完全符合规范！")
    else:
        st.error(f"❌ 共发现 {total} 个问题")

    st.divider()

    st.subheader("📍 问题定位")

    for p, c in results.items():
        st.markdown(f"### ❌ 第{p}段")
        st.caption(f"{c['text']}")

        for e in c["errors"]:
            st.write("👉", e)

        st.divider()
