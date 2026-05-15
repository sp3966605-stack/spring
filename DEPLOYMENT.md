# 部署指南

## 本地開發環境

### 前置要求

- Python 3.8+
- pip 或 conda
- 現代網頁瀏覽器

### 快速開始

1. **克隆專案**
```bash
git clone https://github.com/Jason0317-0317/integrated-management-tool.git
cd integrated-management-tool
```

2. **建立虛擬環境（推薦）**
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

3. **安裝依賴**
```bash
pip install -r requirements.txt
```

4. **設定環境變數**

建立 `.env` 檔案：
```env
FINANCE_PASSWORD=your-strong-password-here
```

安裝 dotenv 支援：
```bash
pip install python-dotenv
```

在 `app.py` 頂部加入（如尚未加入）：
```python
from dotenv import load_dotenv
load_dotenv()
```

5. **啟動應用**
```bash
streamlit run app.py
```

應用會在 `http://localhost:8501` 啟動。

---

## 生產環境部署

### Streamlit Community Cloud（推薦免費方案）

1. **前往** [share.streamlit.io](https://share.streamlit.io)，使用 GitHub 帳號登入

2. **部署應用**
   - 選擇您的 GitHub 儲存庫
   - 主檔案設定為 `app.py`

3. **設定 Secrets（重要）**
   - 進入 App Settings → Secrets
   - 加入以下內容：
   ```toml
   FINANCE_PASSWORD = "your-strong-password-here"
   ```

4. **部署完成**，Streamlit Cloud 會提供公開 URL

---

### Railway 部署

1. **連接 GitHub**
   - 訪問 [Railway.app](https://railway.app)
   - 使用 GitHub 帳戶登入並連接儲存庫

2. **設定環境變數**
   - 在 Railway 儀表板 → Variables 中加入：
   ```
   FINANCE_PASSWORD=your-strong-password-here
   ```

3. **設定啟動指令**
   ```
   streamlit run app.py --server.port $PORT --server.address 0.0.0.0
   ```

4. Railway 會自動偵測 `requirements.txt` 並部署

---

### Render 部署

1. **連接 GitHub**
   - 訪問 [Render.com](https://render.com)
   - 建立 Web Service，選擇您的儲存庫

2. **設定**
   - Environment: Python
   - Start Command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`

3. **設定環境變數**
   - 在 Environment 頁籤加入 `FINANCE_PASSWORD`

---

### Docker 部署

1. **建立 Dockerfile**
```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

2. **建立 .dockerignore**
```
__pycache__
*.pyc
.git
.gitignore
venv/
.env
```

3. **構建和運行**
```bash
# 構建映像
docker build -t chunbulao-app:latest .

# 運行容器（透過環境變數傳入密碼）
docker run -p 8501:8501 -e FINANCE_PASSWORD=your-strong-password chunbulao-app:latest
```

---

## 環境變數配置

| 變數名稱 | 說明 | 必填 |
|---------|------|------|
| `FINANCE_PASSWORD` | 財務角色登入密碼 | **是** |

> **安全提醒**：不要將 `.env` 提交到 Git，請確認 `.gitignore` 中已包含 `.env`。

---

## 安全性建議

### 生產環境檢查清單

- [ ] **財務密碼已設定為環境變數**，未硬編碼於原始碼中
- [ ] `.env` 檔案已加入 `.gitignore`，確保不被提交至版本控制
- [ ] 使用強密碼（至少 12 字元，包含大小寫、數字、符號）
- [ ] 啟用 HTTPS（各雲端平台預設提供）
- [ ] 定期更換財務密碼並更新平台環境變數
- [ ] 定期更新 Python 依賴套件
- [ ] 監控應用日誌
- [ ] 設定備份策略

### 定期更新依賴

```bash
# 檢查過時的套件
pip list --outdated

# 更新 requirements.txt
pip freeze > requirements.txt
```

---

## 故障排除

**Q: 財務登入顯示「系統尚未設定財務密碼」**
- 確認環境變數 `FINANCE_PASSWORD` 已正確設定
- 本地開發請確認 `.env` 檔案存在且格式正確
- 雲端部署請至平台的環境變數/Secrets 設定頁面確認

**Q: 應用啟動失敗**
- 確認 `requirements.txt` 中所有套件已安裝
- 查看終端機錯誤訊息

**Q: 檔案上傳失敗**
- 確認檔案格式正確（Excel 或 CSV）
- 確認檔案包含必要欄位
- 預設上傳限制為 200MB；如需調整，在啟動指令加入 `--server.maxUploadSize=500`

**Q: 應用運行緩慢**
- Streamlit Community Cloud 免費方案資源有限，考慮升級或改用 Railway/Render
- 確認上傳的 Excel 檔案不過大

---

## 版本控制

使用 Git 標籤標記版本：

```bash
git tag -a v2.1.0 -m "Release version 2.1.0"
git push origin v2.1.0
```

---

如有問題，請參考 README.md 或提交 Issue。
