# --- helper.py (最終完整版) ---
# 功能包括：多工作表選擇、不區分大小寫、萬用字元支援、符合條件的資料置頂

import streamlit as st
import pandas as pd
import io
import re


def process_excel_file(df, search_column, comparison_items, keyword_to_fill, case_sensitive):
    """
    處理 DataFrame，根據比對清單和指定欄位進行標註。
    使用更穩定的排序方法。
    """
    df_processed = df.copy()
    remark_column = "備註欄"

    if remark_column not in df_processed.columns:
        df_processed[remark_column] = ""

    df_processed[search_column] = df_processed[search_column].astype(str)
    df_processed[remark_column] = df_processed[remark_column].astype(str)

    # --- 搜尋邏輯 (不變) ---
    regex_items = []
    for item in comparison_items:
        if item == '*':
            regex_items.append(re.escape(item))
        else:
            regex_items.append(re.escape(item).replace('\\*', '.*'))

    regex_pattern = '|'.join(regex_items)
    regex_flags = 0 if case_sensitive else re.IGNORECASE

    if not regex_pattern:
        mask = pd.Series([False] * len(df_processed), index=df_processed.index)
    else:
        mask = df_processed[search_column].str.contains(
            regex_pattern,
            na=False,
            regex=True,
            flags=regex_flags
        )

    # --- 標註邏輯 (不變) ---
    df_processed.loc[mask, remark_column] = keyword_to_fill

    # --- 【關鍵修正】更穩健的排序邏輯 ---
    # 1. 將 mask (True/False 序列) 新增為一個暫時的排序欄位。
    #    給它一個幾乎不可能與現有欄位重複的名稱。
    sort_col_name = "__sort_col__"
    df_processed[sort_col_name] = mask

    # 2. 根據這個新欄位進行降序排序 (True=1, False=0，所以 True 會排在前面)。
    df_sorted = df_processed.sort_values(by=sort_col_name, ascending=False, kind='mergesort')

    # 3. 刪除這個暫時的排序欄位，保持 DataFrame 的乾淨。
    df_sorted = df_sorted.drop(columns=[sort_col_name])
    # ------------------------------------

    return df_sorted, mask.sum()
def main():
    """
    Streamlit 應用程式的主函式。
    """
    st.set_page_config(page_title="Excel 關鍵字搜尋工具",page_icon="🧩", layout="centered")
    st.title(" 🧩 Excel 關鍵字搜尋工具")
    st.write("上傳 Excel 檔案，選擇工作表與欄位，輸入多個關鍵字進行批次標註。")

    # 檔案上傳元件，允許 .xls 和 .xlsx
    uploaded_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            # 偵測所有工作表名稱
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names

            selected_sheet = None

            # 如果有多個工作表，提供下拉選單讓使用者選擇
            if len(sheet_names) > 1:
                st.info(f"偵測到 {len(sheet_names)} 個工作表，請選擇您要處理的一個。")
                selected_sheet = st.selectbox(
                    "請選擇工作表 (Sheet)：",
                    sheet_names
                )
            else:
                # 如果只有一個工作表，就自動選擇它
                selected_sheet = sheet_names[0]

            # 只有在使用者選擇了工作表之後，才執行後續操作
            if selected_sheet:
                st.markdown("---")  # 畫一條分隔線

                # 讀取使用者指定的工作表資料
                df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
                st.write(f"✅ 已選擇工作表 **'{selected_sheet}'**！以下是資料預覽：")
                st.dataframe(df.head())

                # --- 使用者輸入介面 ---
                col1, col2 = st.columns(2)
                with col1:
                    search_column = st.selectbox("1. 請選擇要搜尋的欄位：", df.columns)
                with col2:
                    case_sensitive = st.toggle("區分英文大小寫", value=False)

                comparison_data = st.text_area(
                    "2. 請輸入要搜尋的關鍵字 (一行一個)",
                    height=150,
                    help="請將所有要搜尋的關鍵字貼在此處。'*' 可作為萬用字元。"
                )

                keyword_to_fill = st.text_input("3. 請輸入符合條件後，要標註的文字：")

                # --- 執行按鈕與後續邏輯 ---
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
                            case_sensitive=case_sensitive
                        )
                        st.success(f"處理完成！共找到並標註了 {match_count} 筆符合的資料。")
                        st.write("以下是處理後的資料預覽（符合條件的已置頂）：")
                        st.dataframe(processed_df)  # 顯示完整的、已排序的 DataFrame

                        # --- 下載檔案邏輯 ---
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            # 儲存已排序的 DataFrame
                            processed_df.to_excel(writer, index=False, sheet_name=selected_sheet)
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
