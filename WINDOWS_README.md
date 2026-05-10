# easyMoney Windows Python 版

`easy_money_win.py` 是 `easyMoney.swift` 的 Windows 第一版迁移。它优先保留常用工作流：微信朋友圈窗口定位、刷新按钮标定、按用户头像定位动态、自动评论、LLM/豆包答题、SQLite 知识库查询。

这不是 Swift 版的逐行复刻。macOS 的 `AXUIElement`、`CGEvent`、`ScreenCaptureKit`、`Vision`、`CoreML` 在 Windows 上分别替换为 UI Automation、`pyautogui`、`mss`、可选 OpenCV/YOLO。第一版不默认启用 OCR，也不迁移 CoreML。

## 安装

1. 安装 Python 3.11 或更新版本，并勾选 “Add python.exe to PATH”。
2. 在 `E:\Murder` 打开 PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -r requirements.txt
```

如果 PowerShell 禁止激活 venv，可临时运行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

## 配置文件位置

Windows 版使用 `Path.home()`，并尽量兼容 Swift 版的文件名：

- `%USERPROFILE%\.wechat_refresh_offset`
- `%USERPROFILE%\.wechat_comment_config`
- `%USERPROFILE%\.wechat_avatar_tpl.png`
- `%USERPROFILE%\.wechat_avatar_offset`
- `%USERPROFILE%\.wechat_user_templates.json`
- `%USERPROFILE%\.wechat_kb.sqlite`
- `%USERPROFILE%\.easyMoney.env`
- `%USERPROFILE%\.easyMoney\doubaotext-prefix-cache.json`
- `%USERPROFILE%\.easyMoney\userPhoto\用户名.png`

如果你已有 macOS 版知识库，可把 `~/.wechat_kb.sqlite` 复制到 Windows 的 `%USERPROFILE%\.wechat_kb.sqlite`。

## LLM / 豆包配置

在项目目录或用户目录创建 `.easyMoney.env`：

```env
EASYMONEY_LLM_PROVIDER=doubao
ARK_API_KEY=你的火山或豆包 API Key
ARK_MODEL=doubao-seed-2-0-mini-260215
ARK_ENDPOINT=https://ark.cn-beijing.volces.com/api/v3/responses
EASYMONEY_LLM_TIMEOUT=18
```

也支持：

- `EASYMONEY_LLM_PROVIDER=ollama`
- `EASYMONEY_LLM_PROVIDER=openai`
- `OPENAI_API_KEY`
- `DOUBAO_API_KEY` / `ARK_API_KEY` / `VOLCENGINE_API_KEY`

## 首次标定流程

打开微信桌面版，并进入朋友圈窗口。

```powershell
python .\easy_money_win.py uia-dump --buttons-only
python .\easy_money_win.py locate
python .\easy_money_win.py comment-locate
python .\easy_money_win.py avatar-template-locate --name 方南 --set-default
python .\easy_money_win.py avatar-center-locate
python .\easy_money_win.py comment --text "好看！" --user 方南 --debug
```

确认 `--debug` 定位正确后，再执行真实评论：

```powershell
python .\easy_money_win.py comment --text "好看！" --user 方南
```

## 常用命令

```powershell
python .\easy_money_win.py run --interval 15
python .\easy_money_win.py comment --solve-question --user 方南
python .\easy_money_win.py comment --doubao --noLocal --user 方南
python .\easy_money_win.py comment --LLM --user 方南
python .\easy_money_win.py kb stats
python .\easy_money_win.py kb search "关键词"
python .\easy_money_win.py kb ask "问题" --store "商家名"
python .\easy_money_win.py llm ask "问题"
python .\easy_money_win.py doubao ask "朋友圈正文"
```

## 测试

安装依赖后可以运行：

```powershell
python -m py_compile .\easy_money_win.py
python -m unittest discover -s tests
```

## YOLO / 视觉能力

`--LLM --vision` 需要额外安装 `ultralytics` 并配置模型：

```powershell
python -m pip install ultralytics
```

`.easyMoney.env`：

```env
EASYMONEY_YOLO_MODEL=C:\path\to\best.pt
EASYMONEY_YOLO_CONF=0.25
```

没有模型或未安装 `ultralytics` 时，Windows 版会明确报错，而不会静默假装识别成功。

## Windows 注意事项

- 建议 Windows 缩放先用 100% 或 125%。脚本会启用 DPI aware，但微信 UI 与截图坐标仍可能受系统缩放影响。
- 如果 UIA 读不到朋友圈正文，`comment --text` 仍可用；`--solve-question`、`--doubao`、`--LLM` 需要正文时会提示失败。
- 如果头像匹配失败，重新运行 `avatar-template-locate`，确保鼠标放在头像中心；必要时使用 `user add <name> --template <path> --threshold 0.65` 降低阈值。
- `--debug` 不会发送评论，只移动鼠标并打印定位信息。
