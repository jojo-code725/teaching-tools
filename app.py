"""
教培工具网页版 — 学生档案生成 + 课后反馈生成
运行: streamlit run app.py
"""

import streamlit as st
import openpyxl
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from collections import defaultdict
import json, os, tempfile, zipfile, io

st.set_page_config(page_title="教培效率工具", page_icon="📚", layout="wide")
st.title("📚 教培效率工具 v1.0")

tab1, tab2 = st.tabs(["📋 学生档案生成", "✏️ 课后反馈生成"])

# ============================================================
# 共享工具函数
# ============================================================
PAGE_W = Cm(29.7); PAGE_H = Cm(21.0)
MARGIN_LR = Cm(2.54); MARGIN_TB = Cm(3.05)

TABLE_HEADERS = ["序号", "科目", "托管老师", "月考成绩", "班级排名", "总分", "学校排名",
                 "月度学习表现", "本月考情分析", "下一阶段学科建议", "备注"]

def set_page(doc):
    sec = doc.sections[0]
    sec.page_width = PAGE_W; sec.page_height = PAGE_H
    sec.left_margin = MARGIN_LR; sec.right_margin = MARGIN_LR
    sec.top_margin = MARGIN_TB; sec.bottom_margin = MARGIN_TB

def add_title(doc):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("一鸣学员档案"); run.font.size = Pt(14); run.bold = True

def add_info_line(doc, school, grade, name, enroll):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    text = f"学校：{school}    年级班级：{grade}         姓名：{name}       入学时间：{enroll}"
    run = p.add_run(text); run.font.size = Pt(14); run.bold = True

def add_subject_header(doc, name):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(name); run.font.size = Pt(12); run.bold = True

def add_label(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text); run.font.size = Pt(12); run.bold = True

def add_dual(doc, label, text):
    p = doc.add_paragraph()
    r1 = p.add_run(label); r1.font.size = Pt(12); r1.bold = True
    if text: p.add_run(text)

def add_body(doc, text):
    if not text: return
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run(text)

def add_body_lines(doc, text):
    if not text: return
    for line in text.replace("\r", "").split("\n"):
        line = line.strip()
        if not line: continue
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.add_run(line)

def add_summary_table(doc, subjects):
    n = len(subjects)
    table = doc.add_table(rows=n+1, cols=11)
    table.style = "Table Grid"; table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(TABLE_HEADERS):
        cell = table.rows[0].cells[j]; cell.text = ""
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h); run.font.size = Pt(9); run.bold = True
    for i, subj in enumerate(subjects):
        row = table.rows[i+1]
        vals = [str(i+1), subj.get("科目",""), subj.get("托管老师",""),
                subj.get("月考成绩",""), subj.get("班级排名",""),
                subj.get("总分",""), subj.get("学校排名",""),
                "", "", "", ""]
        for j, v in enumerate(vals):
            cell = row.cells[j]; cell.text = ""
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(v)

# ============================================================
# Tab 1: 学生档案生成
# ============================================================
with tab1:
    st.header("学生档案生成")
    st.markdown("上传 Excel（学生更新数据.xlsx）+ 基础信息库（JSON），一键生成全部档案。")

    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("📊 学生更新数据 Excel", type=["xlsx"], key="archive_excel")
    with col2:
        json_file = st.file_uploader("🗂 学生基础信息库 JSON", type=["json"], key="archive_json")

    if excel_file and json_file:
        try:
            base_info = json.loads(json_file.read())
            wb = openpyxl.load_workbook(excel_file)
            ws = wb["学生更新数据"]
            rows = list(ws.iter_rows(min_row=2, values_only=True))

            students = defaultdict(lambda: {"subjects": []})
            for row in rows:
                if not row or not row[0]: continue
                vals = [str(c).strip() if c else "" for c in row]
                name = vals[0]
                if not name: continue
                subj_name = vals[1]
                base = base_info.get(name, {})
                teacher = ""
                for bs in base.get("科目", []):
                    if bs["科目"] == subj_name:
                        teacher = bs["托管老师"]; break

                subj = {"科目": subj_name, "托管老师": teacher,
                        "月考成绩": vals[2], "班级排名": vals[3],
                        "总分": vals[4], "学校排名": vals[5],
                        "优点": vals[6], "不足": vals[7],
                        "考情分析": vals[8], "下一阶段建议": vals[9]}
                if not students[name].get("学校"):
                    students[name].update({"学校": base.get("学校",""),
                        "年级班级": base.get("年级班级",""),
                        "入学时间": base.get("入学时间",""), "name": name})
                students[name]["subjects"].append(subj)

            if not students:
                st.warning("未读取到学生数据，请检查 Excel 文件格式")
            else:
                if st.button("🚀 生成全部档案", type="primary", key="gen_archive"):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        zip_path = os.path.join(tmpdir, "档案汇总.zip")
                        with zipfile.ZipFile(zip_path, 'w') as zf:
                            for name, data in students.items():
                                doc = Document(); set_page(doc)
                                add_title(doc)
                                add_info_line(doc, data.get("学校",""), data.get("年级班级",""),
                                             name, data.get("入学时间",""))
                                st.session_state[f"_p_archive"] = doc.add_paragraph()
                                for subj in data["subjects"]:
                                    add_subject_header(doc, subj["科目"])
                                    add_label(doc, "月度学习表现：")
                                    add_dual(doc, "优点：", subj.get("优点",""))
                                    add_dual(doc, "不足：", subj.get("不足",""))
                                    add_label(doc, "考情分析：")
                                    add_body(doc, subj.get("考情分析",""))
                                    add_label(doc, "下一阶段建议：")
                                    add_body_lines(doc, subj.get("下一阶段建议",""))
                                    doc.add_paragraph()
                                add_summary_table(doc, data["subjects"])
                                safe = name.replace("/","_")
                                docx_path = os.path.join(tmpdir, f"{safe}.docx")
                                doc.save(docx_path)
                                zf.write(docx_path, f"{safe}.docx")

                        with open(zip_path, "rb") as f:
                            st.download_button("📥 下载全部档案（ZIP）", f.read(),
                                               "学生档案汇总.zip", "application/zip")
                        st.success(f"✅ 已生成 {len(students)} 份档案")
        except Exception as e:
            st.error(f"处理失败：{e}\n请确认文件格式正确")

# ============================================================
# Tab 2: 课后反馈生成
# ============================================================
with tab2:
    st.header("课后反馈生成")
    st.markdown("上传 Excel（课后反馈数据.xlsx），一键生成可直接发微信的文案。")

    fb_file = st.file_uploader("📊 课后反馈数据 Excel", type=["xlsx"], key="fb_excel")

    if fb_file:
        try:
            wb = openpyxl.load_workbook(fb_file)
            ws1 = wb["上课内容"]
            course_name = str(ws1.cell(row=2, column=2).value or "").strip()
            date = str(ws1.cell(row=3, column=2).value or "").strip()
            lesson = str(ws1.cell(row=4, column=2).value or "").strip()

            ws2 = wb["学生表现"]
            students_fb = []
            for row in ws2.iter_rows(min_row=2, values_only=True):
                n = str(row[0]).strip() if row[0] else ""
                notes = str(row[1]).strip() if row[1] else ""
                if n and notes: students_fb.append({"姓名": n, "表现": notes})

            if not lesson:
                st.warning("「上课内容」为空")
            elif not students_fb:
                st.warning("「学生表现」表中没有数据")
            else:
                if st.button("🚀 生成全部反馈", type="primary", key="gen_fb"):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        zip_path = os.path.join(tmpdir, "反馈汇总.zip")
                        with zipfile.ZipFile(zip_path, 'w') as zf:
                            all_text = ""
                            for s in students_fb:
                                txt = f"【课后反馈】{date}\n\n"
                                txt += f"📚 上节课主要内容：\n\n{lesson}\n\n"
                                txt += f"👦 {s['姓名']}课堂表现：\n\n{s['表现']}"
                                safe = s["姓名"].replace("/","_")
                                txt_path = os.path.join(tmpdir, f"{safe}.txt")
                                with open(txt_path, "w", encoding="utf-8") as tf:
                                    tf.write(txt)
                                zf.write(txt_path, f"{safe}.txt")
                                all_text += txt + "\n\n" + "—"*20 + "\n\n"

                            # 汇总文件
                            summary_path = os.path.join(tmpdir, "全部反馈汇总.txt")
                            with open(summary_path, "w", encoding="utf-8") as sf:
                                sf.write(all_text)
                            zf.write(summary_path, "全部反馈汇总.txt")

                        with open(zip_path, "rb") as f:
                            st.download_button("📥 下载全部反馈（ZIP）", f.read(),
                                               "课后反馈汇总.zip", "application/zip")
                        st.success(f"✅ 已生成 {len(students_fb)} 份反馈")

                        # 预览第一条
                        with st.expander("👁 预览第一条反馈"):
                            s = students_fb[0]
                            preview = f"【课后反馈】{date}\n\n"
                            preview += f"📚 上节课主要内容：\n\n{lesson}\n\n"
                            preview += f"👦 {s['姓名']}课堂表现：\n\n{s['表现']}"
                            st.text(preview)

        except Exception as e:
            st.error(f"处理失败：{e}\n请确认文件格式正确")

# ============================================================
# 侧边栏：使用说明
# ============================================================
with st.sidebar:
    st.markdown("### 📖 使用说明")
    st.markdown("""
    **学生档案生成：**
    1. 上传「学生更新数据.xlsx」
    2. 上传「学生基础信息库.json」
    3. 点击生成 → 下载 ZIP

    **课后反馈生成：**
    1. 上传「课后反馈数据.xlsx」
    2. 点击生成 → 下载 ZIP
    3. 打开 .txt 复制到微信

    ---
    **关于部署上线：**
    本地运行：`streamlit run app.py`
    免费部署：Streamlit Cloud
    """)
