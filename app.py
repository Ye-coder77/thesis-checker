import streamlit as st
from docx import Document
import re
import tempfile

st.set_page_config(page_title="论文格式检测工具", layout="wide")

st.title("📄 学位论文格式检测工具")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ===========================
# 工具函数
# ===========================

def is_chinese(text):
    return any('\u4e00' <= c <= '\u9fff' for c in text)

def is_english_or_number(text):
    return any(c.isascii() and (c.isalpha() or c.isdigit()) for c in text)

def is_title_level1(text):
    return re.match(r'^第\d+章', text)

def is_title_level2(text):
    return re.match(r'^\d+\.\d+', text)

def is_title_level3(text):
    return re.match(r'^\d+\.\d+\.\d+', text)

def is_figure(text):
    return text.startswith("图")

def is_table(text):
    return text.startswith("表")


# ===========================
# 核心检测（已彻底修复误判）
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

    current_chapter = None

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        errors = []
        pf = para.paragraph_format

        # ===== 正确解析章号 =====
        if is_title_level1(text):
            match = re.search(r'\d+', text)
            if match:
                current_chapter = int(match.group())

        # ===== 行距 =====
        ls = pf.line_spacing
        if ls is None:
            errors.append("未设置行距（应为20磅）")
            summary["行距错误"] += 1
        elif hasattr(ls, "pt") and ls.pt != 20:
            errors.append("行距不是20磅")
            summary["行距错误"] += 1

        # ===== 缩进 =====
        if pf.first_line_indent is None:
            errors.append("未设置首行缩进")
            summary["缩进错误"] += 1

        # ===== 标题检测（修复版） =====
        if is_title_level1(text):
            if "第一章" in text:
                errors.append("一级标题不能使用中文编号")
                summary["标题错误"] += 1

        # 二级标题
        if is_title_level2(text):
            match = re.match(r'^(\d+)\.(\d+)', text)
            if match:
                chapter_num = int(match.group(1))
                if current_chapter is not None and chapter_num != current_chapter:
                    errors.append(f"二级标题编号错误（应属于第{current_chapter}章）")
                    summary["标题错误"] += 1

        # 三级标题
        if is_title_level3(text):
            match = re.match(r'^(\d+)\.(\d+)\.(\d+)', text)
            if match:
                chapter_num = int(match.group(1))
                if current_chapter is not None and chapter_num != current_chapter:
                    errors.append(f"三级标题编号错误（应属于第{current_chapter}章）")
                    summary["标题错误"] += 1

        # ===== 图表 =====
        if current_chapter:
            if is_figure(text):
                if not re.match(rf'^图{current_chapter}\.\d+', text):
                    errors.append("图编号应为 图X.X")
                    summary["图表错误"] += 1

            if is_table(text):
                if not re.match(rf'^表{current_chapter}\.\d+', text):
                    errors.append("表编号应为 表X.X")
                    summary["图表错误"] += 1

        # ===== 字体字号 =====
        for run in para.runs:
            name = run.font.name
            size = run.font.size

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
            results[i + 1] = {
                "text": text,
                "errors": list(set(errors))
            }

    return results, summary


# ===========================
# 标注论文（真正定位）
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
    col1.metric("字体", summary["字体错误"])
    col2.metric("标题", summary["标题错误"])
    col3.metric("图表", summary["图表错误"])

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
