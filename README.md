# 圖片匯入 PowerPoint 工具

這是一個 Python 工具，將資料夾內的 PNG 圖片自動匯入到 PowerPoint 簡報中，圖片將依檔名中的數字排序，並自動裁切為 16:9 比例。

## 安裝

### 使用 Pip

```bash
pip install -r requirements.txt
```

### 使用 Poetry

```bash
poetry install
```

## 使用方法

執行以下指令啟動工具：

```bash
python main.py
```

1. 選擇輸入圖片資料夾。
2. 選擇輸出檔案位置（預設為輸入資料夾的 `output.pptx`）。
3. 點擊「開始處理」，完成後檔案會匯出到指定位置。

## 注意

- 僅支持 PNG 格式的圖片。
- 圖片會裁切為 16:9 比例，請確保重要內容位於圖片中心。
- 圖片會依檔名中的數字進行排序。

## LICENSE

本專案使用 MIT 授權條款。詳細資訊請參閱 [LICENSE](LICENSE) 。
