import io
import zipfile
import streamlit as st
from core import (
    load_course_lookup,
    build_student_records,
    build_notice_docx,
    build_combined_notice_docx,
    format_level_period,
)

st.set_page_config(page_title="補課通知單產生器", page_icon="📋", layout="centered")

st.title("📋 補課通知單產生器")
st.caption("上傳簽到表與課程名稱，一鍵產出每位學員的補課通知單（Word）")

with st.expander("使用說明", expanded=False):
    st.markdown(
        """
        1. 上傳一份或多份「簽到表」Excel（檔名需包含班級關鍵字：**初級 / 中級 / 高級 / 研經**，
           以及日夜關鍵字：**日 / 夜**，例如「初級夜簽到表.xlsx」）
        2. 上傳一份「課程名稱」Excel：初級／中級／高級 各一個分頁（日夜共用），
           研經班日夜課表不同，需分開為「研經日」「研經夜」兩個分頁
        3. 按下「產生補課通知單」
        4. 檢查預覽結果，確認無誤後下載 ZIP（內含每位學員的 Word 通知單）
        """
    )

st.subheader("① 上傳簽到表（可多選）")
attendance_uploads = st.file_uploader(
    "簽到表 Excel（.xlsx）", type=["xlsx"], accept_multiple_files=True, key="attendance"
)

st.subheader("② 上傳課程名稱")
course_upload = st.file_uploader("課程名稱 Excel（.xlsx）", type=["xlsx"], key="course")

st.subheader("③ 補課期限與說明（會套用到本批全部通知單）")
col1, col2 = st.columns(2)
with col1:
    org_name = st.text_input("機構名稱", value="普印精舍")
with col2:
    deadline_text = st.text_input("補課期限（例如：9月2日）", value="9月2日")
hours_note = st.text_input(
    "報到時間說明",
    value="於周一至周日早上9:00~20:00至精舍櫃檯報到完成補課，",
)
late_note = st.text_input(
    "逾期說明",
    value="若超過補課期限，請事先預約補課時間，以利事前安排。",
)

st.subheader("④ 輸出方式")
output_mode = st.radio(
    "選擇通知單輸出方式",
    options=["每人一份（ZIP，各自獨立 Word 檔）", "合併成一份 Word（精簡版面，省紙、附裁切分隔線）"],
    index=0,
)

generate = st.button("🚀 產生補課通知單", type="primary",
                      disabled=not (attendance_uploads and course_upload))

if generate:
    course_bytes = course_upload.getvalue()
    course_lookup = load_course_lookup(course_bytes, course_upload.name)

    attendance_files = [(f.name, f.getvalue()) for f in attendance_uploads]
    students, warnings = build_student_records(attendance_files, course_lookup)

    if warnings:
        st.warning("以下項目請人工確認：\n\n" + "\n".join(f"- {w}" for w in warnings))

    if not students:
        st.error("沒有找到任何需要補課通知的學員，請確認檔案內容與格式。")
    else:
        st.success(f"共 {len(students)} 位學員需要補課通知")

        preview_rows = [
            {
                "姓名": s["name"],
                "法名": s.get("dharma_name") or "",
                "班級": format_level_period(s.get("level"), s.get("period")),
                "組別": s.get("group") or "",
                "缺曠堂數": s.get("absence_count"),
                "缺曠日期": "、".join(i["date"] for i in s["items"]),
            }
            for s in students
        ]
        st.dataframe(preview_rows, use_container_width=True)

        if output_mode.startswith("每人一份"):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for s in students:
                    docx_buf = build_notice_docx(
                        s, org_name=org_name, deadline_text=deadline_text,
                        hours_note=hours_note, late_note=late_note,
                    )
                    safe_name = f"補課通知單_{s['name']}.docx"
                    zf.writestr(safe_name, docx_buf.getvalue())
            zip_buf.seek(0)

            st.download_button(
                "⬇️ 下載全部補課通知單（ZIP）",
                data=zip_buf,
                file_name="補課通知單.zip",
                mime="application/zip",
                type="primary",
            )
        else:
            combined_buf = build_combined_notice_docx(
                students, org_name=org_name, deadline_text=deadline_text,
                hours_note=hours_note, late_note=late_note,
            )
            st.download_button(
                "⬇️ 下載合併版補課通知單（Word）",
                data=combined_buf,
                file_name="補課通知單_合併版.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
