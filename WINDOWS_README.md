# easyMoney Windows Python 版

`easy_money_win.py` 是 `easyMoney.swift` 的 Windows 第一版迁移。它优先保留常用工作流：微信朋友圈窗口定位、按 `locate` 保存的刷新按钮坐标自动刷新、通过 UI Automation 读取朋友圈 `sns_list` 的第二条 `ListItem` 来匹配目标动态、自动评论、LLM/豆包答题、SQLite 知识库查询。

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
- `EASYMONEY_DXGI_STREAM_FPS=240`：DXGI 流采帧目标帧率，只在 DXGI 后端生效

## 配置文件位置

Windows 版使用 `Path.home()`，并尽量兼容 Swift 版的文件名：

- `%USERPROFILE%\.wechat_comment_config`
- `%USERPROFILE%\.wechat_refresh_offset`
- `%USERPROFILE%\.wechat_kb.sqlite`
- `%USERPROFILE%\.easyMoney.env`
- `%USERPROFILE%\.easyMoney\doubaotext-prefix-cache.json`

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
python .\easy_money_win.py uia-dump
python .\easy_money_win.py locate
python .\easy_money_win.py comment-locate
python .\easy_money_win.py comment --text "好看！" --user 方南 --debug
```

`locate` 会手动标定顶部“刷新”按钮，并保存窗口相对坐标到 `%USERPROFILE%\.wechat_refresh_offset`。运行后把鼠标移到刷新按钮中心，倒计时结束时脚本会保存坐标；后续刷新不再从 UIA 里找刷新按钮。

`uia-dump` 默认只展示朋友圈 `sns_list` 下第二条 `ListItem`，用于确认微信当前暴露出来的正文内容。后续 `--user 方南` 表示脚本会读取第二条 `ListItem`，并检查它的开头是否匹配 `方南`；不再需要保存用户图片或视觉模板。

确认 `--debug` 定位正确后，再执行真实评论：

```powershell
python .\easy_money_win.py comment --text "好看！" --user 方南
```

默认使用快捷键链路：点击动态右下角操作按钮后，通过 `Tab+Enter` 打开评论输入框，再通过 `Tab+Tab+Tab+Enter` 发送。评论文本优先走 `KEYEVENTF_UNICODE` 直接输入，逻辑对应 Swift 版的 `keyboardSetUnicodeString`；只有超长文本或控制字符等不适合直输的内容才回退剪贴板。脚本发送评论时不会恢复原剪贴板内容。

默认启用极限快捷键链路：把“移动到操作按钮 + 点击 + Tab+Enter + Unicode 文本输入 + Tab+Tab+Tab+Enter”全部合并为一次 `SendInput` 调用。它不走评论菜单坐标点击、不走发送按钮点击，但很依赖微信窗口响应时机。

```powershell
python .\easy_money_win.py comment --text "好看！" --user 方南
```

如果同目录存在 `easy_money_input.dll`，脚本会优先使用 C++ 原生热路径，进一步减少 Python/ctypes 组装事件的开销。构建方式：

```powershell
.\build_native_helper.bat
```

构建需要 Visual Studio Build Tools 的 `cl.exe`，或 MinGW-w64 的 `g++.exe`。想临时禁用 C++ 热路径可设置：

```powershell
$env:EASYMONEY_DISABLE_NATIVE_INPUT="1"
```

C++ 热路径默认用 `SetCursorPos` 移动鼠标，再用一次 `SendInput` 发送点击和快捷键。若要对比“鼠标移动也塞进 SendInput”的旧模式：

```powershell
$env:EASYMONEY_NATIVE_MOVE_MODE="sendinput"
```

如果需要回退到分段快捷键链路，可以关闭：

```powershell
$env:EASYMONEY_FUSE_ALL_INPUTS="0"
```

如需实验坐标点击打开评论菜单，可以显式指定：

```powershell
python .\easy_money_win.py comment --text "好看！" --user 方南 --open-click
```

如需实验坐标点击发送，可以显式指定：

```powershell
python .\easy_money_win.py comment --text "好看！" --user 方南 --submit-click
```

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
- 如果用户匹配失败，先运行 `python .\easy_money_win.py uia-dump`，确认第二条 `ListItem` 的开头就是你传给 `--user` 的用户名前缀。
- `--debug` 不会发送评论，只移动鼠标并打印定位信息。
