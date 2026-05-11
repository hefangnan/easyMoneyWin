# easyMoney Windows Python 版

`easy_money_win.py` 是 `easyMoney.swift` 的 Windows 第一版迁移。它优先保留常用工作流：微信朋友圈窗口定位、按 `locate` 保存的刷新按钮坐标自动刷新、用 DXGI 截图 + OpenCV 头像模板匹配定位目标动态、自动评论、LLM/豆包答题、SQLite 知识库查询。

这不是 Swift 版的逐行复刻。macOS 的 `AXUIElement`、`CGEvent`、`ScreenCaptureKit`、`Vision`、`CoreML` 在 Windows 上分别替换为 UI Automation、Win32 输入、DXGI Desktop Duplication（`dxcam`，`mss` 兜底）、可选 OpenCV/YOLO。第一版不默认启用 OCR，也不迁移 CoreML。

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

截图后端默认优先使用 DXGI：

- `EASYMONEY_CAPTURE_BACKEND=auto`：默认，优先 `dxcam`，不可用时退回 `mss`
- `EASYMONEY_CAPTURE_BACKEND=dxgi`：强制 DXGI，初始化失败会报错
- `EASYMONEY_CAPTURE_BACKEND=mss`：强制旧截图后端
- `EASYMONEY_DXGI_OUTPUT=0`：选择 DXGI 输出编号，多显示器时可调整
- `EASYMONEY_AVATAR_STREAM=1`：头像匹配默认开启 DXGI 流采帧；设为 `0` 可退回逐次截图
- `EASYMONEY_DXGI_STREAM_FPS=240`：头像匹配流采帧目标帧率，只在 DXGI 后端生效

## 配置文件位置

Windows 版使用 `Path.home()`，并尽量兼容 Swift 版的文件名：

- `%USERPROFILE%\.wechat_comment_config`
- `%USERPROFILE%\.wechat_refresh_offset`
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
python .\easy_money_win.py comment --text "好看！" --user 方南 --debug
```

`locate` 会手动标定顶部“刷新”按钮，并保存窗口相对坐标到 `%USERPROFILE%\.wechat_refresh_offset`。运行后把鼠标移到刷新按钮中心，倒计时结束时脚本会保存坐标；后续刷新不再从 UIA 里找刷新按钮。

`avatar-template-locate --name 方南` 会把鼠标所在头像中心附近的模板保存到 `%USERPROFILE%\.easyMoney\userPhoto\方南.png`，并同时保存头像中心偏移。后续 `--user 方南` 表示用这个头像模板做视觉匹配；默认走 `center_square` 窄区域匹配，必要时可加 `--avatar-wide` 扫左侧更大区域。

确认 `--debug` 定位正确后，再执行真实评论：

```powershell
python .\easy_money_win.py comment --text "好看！" --user 方南
```

评论输入框使用 `Tab+Enter` 打开，并通过 Win32 原生输入事件减少 `pyautogui` 延迟。脚本发送评论时不会恢复原剪贴板内容。

## 常用命令

```powershell
python .\easy_money_win.py locate
python .\easy_money_win.py capture-info --backend dxgi
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
- 如果头像匹配失败，先确认 `avatar-template-locate --name 用户名` 保存的是头像中心，不是昵称或正文；也可以运行 `comment --text "测试" --user 用户名 --avatar-wide --debug` 扩大搜索范围。
- `--debug` 不会发送评论，只移动鼠标并打印定位信息。
