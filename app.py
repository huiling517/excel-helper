# --- 請將以下所有內容，完整貼到「helper.py」檔案中 ---

import streamlit as st
import pandas as pd
import io
import re  # 匯入正規表示式模組


def process_excel_file(df, search_column, comparison_items, keyword_to_fill):
    """
    處理 DataFrame，根據比對清單和指定欄位進行標註。
    支援 '*' 作為萬用字元，並修正單獨 '*' 會匹配所有的問題。
    """
    # 在副本上操作以確保安全
    df_processed = df.copy()
    remark_column = "備註欄"

    # 如果備註欄不存在，就新增一個
    if remark_column not in df_processed.columns:
        df_processed[remark_column] = ""

    # 將搜尋欄位和備註欄都轉為文字，避免比對時因資料型態不同而出錯
    df_processed[search_column] = df_processed[search_column].astype(str)
    df_processed[remark_column] = df_processed[remark_column].astype(str)

    # --- 核心修正：更智慧地處理萬用字元 ---
    regex_items = []
    for item in comparison_items:
        # 情況1：如果使用者輸入的關鍵字剛好就是一個單獨的 '*'
        # 我們將其視為尋找 '*' 這個符號本身，而不是萬用字元。
        if item == '*':
            # re.escape('*') 會將其轉換為 '\\*'，這在正規表示式中代表尋找字面上的 '*' 符號。
            regex_items.append(re.escape(item))
        # 情況2：對於所有其他關鍵字 (包含像 '11*' 這樣的組合)
        else:
            # 我們才套用萬用字元邏輯：
            # 1. re.escape(item): 先將所有內容轉為純文字，避免特殊符號干擾。
            # 2. .replace('\\*', '.*'): 再將裡面的 '*' 換成正規表示式的萬用字元 '.*'。
            regex_items.append(re.escape(item).replace('\\*', '.*'))

    # 將所有處理好的正規表示式用 '|' (OR) 串聯起來
    regex_pattern = '|'.join(regex_items)

    # 如果 regex_pattern 為空 (使用者沒輸入任何東西)，則不進行任何搜尋
    if not regex_pattern:
        mask = pd.Series([False] * len(df_processed), index=df_processed.index)
    else:
        # 使用 .str.contains() 進行模糊搜尋
        mask = df_processed[search_column].str.contains(regex_pattern, na=False, regex=True)

    # 將指定的關鍵字填入符合條件的行的備註欄
    df_processed.loc[mask, remark_column] = keyword_to_fill

    # 回傳處理好的 DataFrame 和符合的數量
    return df_processed, mask.sum()


def main():
    """
    Streamlit 應用程式的主函式。
    """
    st.set_page_config(page_title="關鍵字搜尋工具", layout="centered")
    st.title("關鍵字搜尋工具")
    st.write("上傳 Excel 檔案，輸入多個關鍵字，即可對「包含」這些關鍵字的資料進行批次標註。")

    # 檔案上傳元件
    uploaded_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            st.write("✅ 檔案上傳成功！以下是資料預覽：")
            st.dataframe(df.head())

            # 讓使用者選擇欄位和輸入比對資料
            search_column = st.selectbox("1. 請選擇要進行搜尋的欄位：", df.columns)

            # 使用 text_area 讓使用者輸入多行比對資料 (關鍵字)
            comparison_data = st.text_area(
                "2. 請輸入要搜尋的關鍵字 (一行一個)",
                height=150,
                # --- 修改說明文字，使其更清晰 ---
                help="請將所有要搜尋的關鍵字貼在此處。'*' 可作為萬用字元（例如 `11*`）。若要搜尋 '*' 符號本身，請單獨輸入一個 '*'。"
            )

            keyword = st.text_input("3. 請輸入符合條件後，要標註的文字：")

            # 執行按鈕
            if st.button("🚀 開始標註", type="primary"):
                # 將使用者輸入的多行文字，轉換成一個乾淨的 list
                comparison_items = [item.strip() for item in comparison_data.split('\n') if item.strip()]

                if not comparison_items:
                    st.warning("⚠️ 請先在『關鍵字』框中輸入內容！")
                elif not keyword:
                    st.warning("⚠️ 請先輸入要標註的文字！")
                else:
                    processed_df, match_count = process_excel_file(
                        df=df,
                        search_column=search_column,
                        comparison_items=comparison_items,
                        keyword_to_fill=keyword
                    )
                    st.success(f"處理完成！共找到並標註了 {match_count} 筆符合的資料。")
                    st.write("以下是處理後的資料預覽：")
                    st.dataframe(processed_df.head())

                    # --- 使用記憶體來準備下載檔案 ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        processed_df.to_excel(writer, index=False, sheet_name='Sheet1')
                    processed_data = output.getvalue()

                    # 下載按鈕
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
