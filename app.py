import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="DELTA 發票資料擷取工具", 
    page_icon="⚡", 
    layout="wide"
)
st.title("⚡ DELTA 發票資料自動轉 Excel 工具")
st.caption("自動精準擷取 Delta 發票：Invoice No、P.O. NO.、PART NO（鎖定第二組括號）、QTY，支援批次轉換與既有 Excel 累加！")


def parse_delta_invoice(file_bytes):
    """解析 DELTA 格式發票"""
    full_text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"

    # ================= 1. Invoice No =================
    # 匹配例如 "Invoice No : 9265356068"
    inv_match = re.search(
        r"Invoice\s*No\.?\s*[:：]?\s*([0-9A-Za-z]+)", 
        full_text, 
        re.IGNORECASE
    )
    invoice_no = inv_match.group(1).strip() if inv_match else ""

    # ================= 2. P.O. NO. =================
    # 策略 A: 優先鎖定 Delta 常見的 45 開頭 10 碼採購單號 (例如 4510194421)
    po_match = re.search(r"\b(45\d{8})\b", full_text)
    
    # 策略 B: 備援從 P.O. NO. 欄位後面的第一組 8~12 碼長數字
    if not po_match:
        po_match = re.search(
            r"P\.?O\.?\s*NO\.?[\s\S]*?\n\s*([0-9]{8,12})", 
            full_text, 
            re.IGNORECASE
        )
    po_no = po_match.group(1).strip() if po_match else ""

    # ================= 3. PART NO (鎖定第二組括號) =================
    part_no = ""
    # 策略 A: 找到像 (5126295776) (GTMA45PBK39R-V001) 這種兩組連續括號，直接取出第 2 組括號內的內容
    two_brackets_match = re.search(r"\([^)]+\)\s*\(([^)]+)\)", full_text)
    
    if two_brackets_match:
        part_no = two_brackets_match.group(1).strip()
    else:
        # 策略 B (備援): 若只有單一括號且帶有連字號 (例如 (GTMA45PBK39R-V001))
        single_bracket_match = re.search(r"\(([A-Za-z0-9\-]+-[A-Za-z0-9\-]+)\)", full_text)
        if single_bracket_match:
            part_no = single_bracket_match.group(1).strip()

    # ================= 4. QTY =================
    # 策略 A: 匹配 數字 + 單位 (例如 64PCE、64 PCS 等)
    qty_match = re.search(r"\b(\d+)\s*(PCE|PCS|EA|SET)\b", full_text, re.IGNORECASE)
    if qty_match:
        qty = f"{qty_match.group(1)}{qty_match.group(2).upper()}"
    else:
        # 策略 B: 備援從 QTY 標籤下方的數字抓取
        qty_fallback = re.search(r"QTY[\s\S]*?\n\s*(\d+)", full_text, re.IGNORECASE)
        qty = qty_fallback.group(1).strip() if qty_fallback else ""

    return {
        "Invoice No": invoice_no,
        "P.O. NO.": po_no,
        "PART NO": part_no,
        "QTY": qty,
    }


# ==================== 網頁介面排版 ====================
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 步驟 1：(選填) 上傳現有的 Excel 檔案")
    existing_file = st.file_uploader(
        "若要將新發票資料接續累加至既有 Excel，請在此上傳",
        type=["xlsx", "xls"],
        key="existing_excel",
    )

with col2:
    st.markdown("### 步驟 2：上傳 DELTA PDF 發票檔案")
    uploaded_files = st.file_uploader(
        "請選擇或拖曳 PDF 發票（支援多個檔案同時拖入）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files",
    )

if uploaded_files:
    records = []
    progress_bar = st.progress(0)

    for i, file in enumerate(uploaded_files):
        data = parse_delta_invoice(file.read())
        records.append(data)
        progress_bar.progress((i + 1) / len(uploaded_files))

    new_df = pd.DataFrame(records)

    # 合併既有 Excel 邏輯
    if existing_file:
        try:
            old_df = pd.read_excel(existing_file)
            final_df = pd.concat([old_df, new_df], ignore_index=True)
            # 依據 Invoice No 移除重複項（保留最新）
            if "Invoice No" in final_df.columns:
                final_df = final_df.drop_duplicates(subset=["Invoice No"], keep="last")
            st.success(f"🎉 成功解析 {len(records)} 筆發票，並已累加至原有的 Excel！目前累積共 {len(final_df)} 筆。")
        except Exception as e:
            st.error(f"讀取既有 Excel 失敗: {e}")
            final_df = new_df
    else:
        final_df = new_df
        st.success(f"🎉 成功解析 {len(records)} 筆 DELTA 發票資料！")

    # 預覽結果表格
    st.subheader("📊 擷取結果預覽")
    st.dataframe(final_df, use_container_width=True)

    # 產出下載檔案
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Delta_Invoices")
    excel_data = output.getvalue()

    st.download_button(
        label="📥 下載整理好的 Excel 檔案 (.xlsx)",
        data=excel_data,
        file_name="Delta_發票整理.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )