# Discord Mouse Controller 🎮🎤

一個簡潔美觀的 Discord 靜音/拒聽控制器，支援自訂滑鼠或鍵盤快捷鍵來快速切換 Discord 的語音狀態。

## ⚙️ Discord 設定

使用此應用程式前，您需要建立 Discord 應用程式以取得 API 憑證：

1. 前往 [Discord Developer Portal](https://discord.com/developers/applications)
2. 點擊 **New Application** 建立新應用程式
3. 在左側選單中選擇 **OAuth2**
4. 複製 **Client ID** 和 **Client Secret**
5. 在 **Redirects** 中添加 `http://127.0.0.1`
6. 將 Client ID 和 Client Secret 填入應用程式的設定中

## 📖 使用說明

1. **連接 Discord**：填入 Client ID 和 Client Secret 後點擊「連接 Discord」
2. **綁定按鍵**：點擊「拒聽」或「靜音」旁的按鈕，然後按下想要綁定的按鍵
3. **切換狀態**：
   - 使用綁定的按鍵快速切換
   - 或直接點擊介面上的卡片區域
4. **系統列模式**：開啟「關閉時縮小至系統列」可讓程式在背景運行

## 🔨 從原始碼打包

```bash
# 安裝 PyInstaller
pip install pyinstaller

# 打包成單一執行檔
pyinstaller --noconsole --onefile --add-data "web;web" --name "DiscordMouseController" --clean discord_mouse_rpc.py
```

打包完成後，執行檔會在 `dist/` 資料夾中。

## 📁 專案結構

```
discord-mic-toggle/
├── discord_mouse_rpc.py  # 主程式 (Python 後端)
├── web/
│   └── index.html        # 前端介面 (HTML/CSS/JS)
├── .gitignore
└── README.md
```

## 🔧 技術棧

- **後端**: Python, PyWebView, pypresence (Discord RPC)
- **前端**: HTML, CSS (Glassmorphism), JavaScript
- **輸入監聽**: pynput
- **系統列**: pystray, Pillow

## 📝 注意事項

- 請確保 Discord 桌面版已啟動才能連接
- 首次連接時需要在 Discord 中授權應用程式
- 設定會自動儲存在 `config.json` 中（請勿分享此檔案）

## 📄 授權

MIT License

---

**by Svemic** 💜
