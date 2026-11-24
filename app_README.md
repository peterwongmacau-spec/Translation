# MP3 語音轉 Word 逐字稿

這是一個使用 Streamlit 與 OpenAI Whisper 的語音轉文字網頁應用程式。

## 功能

- 上傳 MP3 檔案
- 選擇辨識語言與 Whisper 模型
- 將語音轉錄為文字
- 自動依中文標點切句與分段
- 產生排版後的 Word (.docx) 檔
- 直接從瀏覽器下載 Word 檔

## 部署到 Streamlit Community Cloud

1. 將 `app.py`、`requirements.txt` 與 `README.md` 上傳至 GitHub repository。
2. 前往 [Streamlit Community Cloud](https://share.streamlit.io/)。
3. 登入並連結 GitHub 帳號。
4. 點選 **Create app**。
5. 選擇你的 GitHub repository、分支與 `app.py`。
6. 點選 **Deploy**。

## 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 注意事項

- 建議上傳 MP3 格式檔案。
- 較大的 Whisper 模型較準確，但需要較多運算資源與時間。
- 請勿上傳含有敏感、私人或機密資訊的錄音。
