# --- helper.py (強化版) ---

import streamlit as st
import pandas as pd
import io
import re

def process_excel_file(df, search_column, comparison_items, keyword_to_fill, case_sensitive):
    """
    處理 DataFrame，根據比對清單和指定欄位進行標註。
    新增 case_sensitive 參數來控制是否區分大小寫。
    """
    df_processed = df.copy()
    remark_column = "備註欄"

    if remark_column not in df_processed.columns:
        df_processed[remark_column] = ""

    df_processed[search_column] = df_processed[search_column].astype(str)
    df_processed[remark_column] = df_processed[remark_column].astype(str)

    # 處理萬用字元
    regex_items = []
    for item in comparison_items:
        if item == '*':
            regex_items.append(re.escape(item))
        else:
            regex_items.append(re.escape(item).replace('\\*', '.*'))

    regex_pattern = '|'.join(regex_items)

    # 根據是否區分大小寫，設定 re.flags
    regex_flags = 0 if case_sensitive else re.IGNORECASE

    if not regex_pattern:
        mask = pd.Series([False] * len(df_processed), index=df_processed.index)
    else:
        # 使用 .str.contains() 進行模糊搜尋，並傳入 flags 參數
        mask = df_processed[search_column].str.contains(
            regex_pattern,
            na=False,
            regex=True,
            flags=regex_flags
        )

    df_processed.loc[mask, remark_column] = keyword_to_fill
    return df_processed, mask.sum()

def main():
    st.set_page_config(page_title="關鍵字搜尋工具", layout="centered")
    st.title("關鍵字搜尋工具")
    st.write("上傳 Excel 檔案，輸入多個關鍵字，即可對「包含」這些關鍵字的資料進行批次標註。")

    # 允許上傳 .xls 和 .xlsx
    uploaded_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("✅ 檔案上傳成功！以下是資料預覽：")
            st.dataframe(df.head())

            st.markdown("---") # 分隔線

            # 讓使用者選擇欄位和輸入比對資料
            col1, col2 = st.columns(2)
            with col1:
                search_column = st.selectbox("1. 請選擇要搜尋的欄位：", df.columns)
            with col2:
                # 【新功能】讓使用者選擇是否區分大小寫
                case_sensitive = st.toggle("區分英文大小寫", value=False)


            comparison_data = st.text_area(
                "2. 請輸入要搜尋的關鍵字 (一行一個)",
                height=150,
                help="請將所有要搜尋的關鍵字貼在此處。'*' 可作為萬用字元（例如 `11*`）。"
            )

            keyword_to_fill = st.text_input("3. 請輸入符合條件後，要標註的文字：")

            if st.button("🚀 開始標註", type="primary"):
                comparison_items = [item.strip() for item in comparison_data.split('\n') if item.strip()]

                if not comparison_items:
                    st.warning("⚠️ 請先在『關鍵字』框中輸入內容！")
                elif not keyword_to_fill:
                    st.warning("⚠️ 請先輸入要標註的文字！")
                else:
                    processed_df, match_count = process_excel_file(
                        df=df,
                        search_column=search_column,
                        comparison_items=comparison_items,
                        keyword_to_fill=keyword_to_fill,
                        case_sensitive=case_sensitive # 傳入新參數
                    )
                    st.success(f"處理完成！共找到並標註了 {match_count} 筆符合的資料。")
                    st.write("以下是處理後的資料預覽：")
                    st.dataframe(processed_df.head())

                    # ... (下載部分的程式碼保持不變) ...
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        processed_df.to_excel(writer, index=False, sheet_name='Sheet1')
                    processed_data = output.getvalue()
                    st.download_button(
                        label="📥 點此下載處理後的檔案",
                        data=processed_data,
                        file_name=f"{uploaded_file.name.split('.')[0]}_processed.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"處理檔案時發生錯誤：{e}")

if __name__ == "__main__":
    main()
