# --- 請將以下所有內容，完整貼到「app.py」檔案中 ---

from flask import Flask, render_template, request, send_file
import pandas as pd
import io

# 建立 Flask 應用實例
app = Flask(__name__)


# --- 核心處理邏輯 (從 Streamlit 專案直接複製過來，完全不用改) ---
def process_excel_file(df, search_column, keyword):
    df_processed = df.copy()
    remark_column = "備註欄"
    if remark_column not in df_processed.columns:
        df_processed[remark_column] = ""
    df_processed[search_column] = df_processed[search_column].astype(str)
    df_processed[remark_column] = df_processed[remark_column].astype(str)
    mask = df_processed[search_column].str.contains(keyword, na=False)
    df_processed.loc[mask, remark_column] = keyword
    return df_processed, mask.sum()


# --- 建立路由 (Routes) ---

# 路由 1: 處理主頁面 (GET 請求)
@app.route('/', methods=['GET'])
def index():
    """當使用者第一次打開網頁時，顯示上傳表單。"""
    return render_template('index.html')


# 路由 2: 處理檔案上傳和處理 (POST 請求)
@app.route('/process', methods=['POST'])
def process():
    """當使用者點擊「開始標註」按鈕後，處理提交的表單。"""
    # 1. 檢查是否有檔案被上傳
    if 'excel_file' not in request.files:
        return "錯誤：沒有找到上傳的檔案。", 400

    file = request.files['excel_file']
    if file.filename == '':
        return "錯誤：未選擇任何檔案。", 400

    # 2. 獲取表單中的其他資料
    search_column = request.form.get('search_column')
    keyword = request.form.get('keyword')

    if not search_column or not keyword:
        return "錯誤：請填寫要搜尋的欄位和關鍵字。", 400

    # 3. 處理檔案
    if file:
        try:
            # 直接從上傳的檔案讀取 DataFrame
            df = pd.read_excel(file)

            # 呼叫我們的核心處理函數
            processed_df, match_count = process_excel_file(df, search_column, keyword)

            # --- 將處理好的 DataFrame 存入記憶體中，準備下載 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                processed_df.to_excel(writer, index=False, sheet_name='Sheet1')
            output.seek(0)  # 移動指標到記憶體緩衝區的開頭

            original_filename = file.filename.rsplit('.', 1)[0]
            new_filename = f"{original_filename}_processed.xlsx"

            # 4. 將記憶體中的檔案作為一個下載請求回傳給使用者
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=new_filename
            )

        except Exception as e:
            return f"處理檔案時發生錯誤：{e}", 500


# --- 啟動伺服器 ---
if __name__ == '__main__':
    # debug=True 讓您在開發時，修改程式碼後伺服器會自動重啟
    app.run(debug=True, host='0.0.0.0', port=5003)