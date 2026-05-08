import streamlit as st
from docx import Document
import re

st.title("📄 学位论文格式检测工具（专业版）")

uploaded_file = st.file_uploader("上传论文（.docx）", type=["docx"])


# ========================
# 工具函数
# ========================
def is_chinese(char):
    return '\u4e00' <= char <= '\u9fff'

def is_english_or_number(char):
    return bool(re.match(r'[A-Za-z0-9]', char))


# ========================
# 段落类型识别
# ========================
def get_para_type(text):
    text = text.strip()

    if text.startswith("第") and "章" in text:
        return "一级标题"
    elif re.match(r"\d+\.\d+\.\d+", text):
        return "三级标题"
    elif re.match(r"\d+\.\d+", text):
        return "二级标题"
    else:
        return "正文"


# ========================
# 字体检测（中英文分开）
# ========================
def check_font(paragraphs):
    issues = []

    for i, para in enumerate(paragraphs):
        if not para.text.strip():
            continue

        for run in para.runs:
            font = run.font.name
            text = run.text

            if not text:
                continue

            for char in text:
                if is_chinese(char):
                    if font and "宋" not in font:
                        issues.append({
                            "para": i+1,
                            "type": "字体错误",
                            "msg": f"中文应为宋体，当前为 {font}"
                        })
                        break

                elif is_english_or_number(char):
                    if font and "Times" not in font:
                        issues.append({
                            "para": i+1,
                            "type": "字体错误",
                            "msg": f"英文/数字应为Times New Roman，当前为 {font}"
                        })
                        break

    return issues


# ========================
# 字号检测
# ========================
def check_font_size(paragraphs):
    issues = []

    for i, para in enumerate(paragraphs):
        if not para.text.strip():
            continue

        para_type = get_para_type(para.text)

        for run in para.runs:
            size = run.font.size

            if size:
                pt = size.pt

                if para_type == "正文" and abs(pt - 12) > 0.5:
                    issues.append({
                        "para": i+1,
                        "type": "字号错误",
                        "msg": f"正文应为12pt，小四号"
                    })

                elif para_type == "一级标题" and abs(pt - 16) > 1:
                    issues.append({
                        "para": i+1,
                        "type": "字号错误",
                        "msg": f"一级标题应为16pt"
                    })

                elif para_type == "二级标题" and abs(pt - 14) > 1:
                    issues.append({
                        "para": i+1,
                        "type": "字号错误",
                        "msg": f"二级标题应为14pt"
                    })

    return issues


# ========================
# 行距检测
# ========================
def check_line_spacing(paragraphs):
    issues = []

    for i, para in enumerate(paragraphs):
        spacing = para.paragraph_format.line_spacing

        if spacing:
            try:
                if abs(spacing.pt - 20) > 1:
                    issues.append({
                        "para": i+1,
                        "type": "行距错误",
                        "msg": "应为20磅"
                    })
            except:
                pass
        else:
            issues.append({
                "para": i+1,
                "type": "行距错误",
                "msg": "未设置固定行距"
            })

    return issues


# ========================
# 首行缩进
# ========================
def check_indent(paragraphs):
    issues = []

    for i, para in enumerate(paragraphs):
        indent = para.paragraph_format.first_line_indent

        if indent:
            pt = indent.pt
            if pt < 20:
                issues.append({
                    "para": i+1,
                    "type": "缩进错误",
                    "msg": "首行应缩进2字符"
                })
        else:
            issues.append({
                "para": i+1,
                "type": "缩进错误",
                "msg": "未设置首行缩进"
            })

    return issues


# ========================
# 空行
# ========================
def check_empty_lines(paragraphs):
    issues = []

    for i, para in enumerate(paragraphs):
        if para.text.strip() == "":
            issues.append({
                "para": i+1,
                "type": "空行问题",
                "msg": "存在空行"
            })

    return issues


# ========================
# 主程序
# ========================
if uploaded_file:

    doc = Document(uploaded_file)
    paragraphs = doc.paragraphs

    st.success("文件上传成功，开始检测...")

    all_issues = (
        check_font(paragraphs) +
        check_font_size(paragraphs) +
        check_line_spacing(paragraphs) +
        check_indent(paragraphs) +
        check_empty_lines(paragraphs)
    )

    # ========================
    # 评分
    # ========================
    score = 100 - len(all_issues) * 2
    if score < 0:
        score = 0

    st.metric("📊 格式评分", score)

    # ========================
    # 分类统计
    # ========================
    type_count = {}
    for issue in all_issues:
        t = issue["type"]
        type_count[t] = type_count.get(t, 0) + 1

    st.subheader("📊 错误分类统计")
    for t, c in type_count.items():
        st.write(f"{t}：{c}处")

    # ========================
    # 总体问题
    # ========================
    if not all_issues:
        st.success("✅ 所有格式符合规范！")
    else:
        st.error(f"❌ 共发现 {len(all_issues)} 个问题")

    # ========================
    # 问题段落定位 + 解释
    # ========================
    st.subheader("📍 问题段落定位（含解释）")

    for i, para in enumerate(paragraphs):
        text = para.text.strip()

        if not text:
            continue

        para_issues = [x for x in all_issues if x["para"] == i+1]

        if para_issues:
            st.error(f"❌ 第{i+1}段：{text}")

            for p in para_issues:
                st.write(f"👉 {p['type']}：{p['msg']}")

        else:
            st.write(f"第{i+1}段：{text}")