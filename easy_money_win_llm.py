from __future__ import annotations

import math

from easy_money_win_core import *
from easy_money_win_capture import CaptureBackend
from easy_money_win_input import InputBackend
from easy_money_win_uia import WindowBackend


SINGLE_IMAGE_FOCUS_KEYS = ("down", "tab", "tab", "tab")

def load_easy_money_dotenv() -> dict[str, str]:
    values: dict[str, str] = {}
    for path in [Path.cwd() / ".easyMoney.env", HOME / ".easyMoney.env"]:
        if not path.exists():
            continue
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("export "):
                line = line[len("export ") :].strip()
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
                value = value[1:-1]
            if key:
                values[key] = value
    return values


_DOTENV_CACHE: Optional[dict[str, str]] = None


def first_non_empty_env(keys: Iterable[str]) -> Optional[str]:
    global _DOTENV_CACHE
    for key in keys:
        value = os.environ.get(key, "").strip()
        if value:
            return value
    if _DOTENV_CACHE is None:
        _DOTENV_CACHE = load_easy_money_dotenv()
    for key in keys:
        value = (_DOTENV_CACHE.get(key) or "").strip()
        if value:
            return value
    return None


def load_local_llm_config() -> Optional[LocalLLMConfig]:
    raw_endpoint = first_non_empty_env(
        [
            "EASYMONEY_LLM_ENDPOINT",
            "WECHAT_LLM_ENDPOINT",
            "OMLX_ENDPOINT",
            "OLMX_ENDPOINT",
            "OLLAMA_HOST",
            "DOUBAO_ENDPOINT",
            "ARK_ENDPOINT",
            "VOLCENGINE_LLM_ENDPOINT",
        ]
    )
    inferred = None
    if raw_endpoint:
        if ":8000" in raw_endpoint or "/admin/chat" in raw_endpoint or "/v1/" in raw_endpoint:
            inferred = "openai"
        elif "volces.com" in raw_endpoint or "volcengine.com" in raw_endpoint:
            inferred = "doubao"
    provider = (
        first_non_empty_env(
            [
                "EASYMONEY_LLM_PROVIDER",
                "WECHAT_LLM_PROVIDER",
                "OMLX_PROVIDER",
                "OLMX_PROVIDER",
                "OLLAMA_PROVIDER",
                "DOUBAO_PROVIDER",
                "ARK_PROVIDER",
                "VOLCENGINE_LLM_PROVIDER",
            ]
        )
        or inferred
        or "ollama"
    ).lower()
    if provider in {"openai-compatible", "openai_compatible"}:
        provider = "openai"
    if provider not in {"ollama", "openai", "doubao"}:
        provider = "ollama"

    default_endpoint = {
        "ollama": "http://127.0.0.1:11434/api/chat",
        "openai": "http://127.0.0.1:11434/v1/chat/completions",
        "doubao": "https://ark.cn-beijing.volces.com/api/v3/responses",
    }[provider]
    endpoint = raw_endpoint or default_endpoint
    if endpoint.endswith("/admin/chat"):
        endpoint = endpoint[: -len("/admin/chat")] + "/v1/chat/completions"
    elif provider == "ollama" and endpoint.rstrip("/") in {"http://127.0.0.1:11434", "http://localhost:11434"}:
        endpoint = endpoint.rstrip("/") + "/api/chat"
    elif provider == "openai" and (
        endpoint.rstrip("/") in {"http://127.0.0.1:11434", "http://localhost:11434", "http://127.0.0.1:8000", "http://localhost:8000"}
        or endpoint.endswith("/v1")
    ):
        endpoint = endpoint.rstrip("/")
        endpoint = endpoint + "/chat/completions" if endpoint.endswith("/v1") else endpoint + "/v1/chat/completions"
    elif provider == "doubao" and (
        endpoint.rstrip("/") in {"https://ark.cn-beijing.volces.com", "https://ark.cn-beijing.volces.com/api/v3"}
        or endpoint.endswith("/api/v3")
    ):
        endpoint = endpoint.rstrip("/")
        endpoint = endpoint + "/responses" if endpoint.endswith("/api/v3") else endpoint + "/api/v3/responses"

    model = first_non_empty_env(
        [
            "EASYMONEY_LLM_MODEL",
            "WECHAT_LLM_MODEL",
            "OMLX_MODEL",
            "OLMX_MODEL",
            "OLLAMA_MODEL",
            "DOUBAO_MODEL",
            "ARK_MODEL",
            "VOLCENGINE_LLM_MODEL",
        ]
    )
    if not model and provider == "doubao":
        model = "doubao-seed-2-0-mini-260215"
    if not model:
        return None
    timeout = float(first_non_empty_env(["EASYMONEY_LLM_TIMEOUT", "WECHAT_LLM_TIMEOUT"]) or "18")
    api_key = first_non_empty_env(
        [
            "EASYMONEY_LLM_API_KEY",
            "WECHAT_LLM_API_KEY",
            "OMLX_API_KEY",
            "OLMX_API_KEY",
            "OPENAI_API_KEY",
            "DOUBAO_API_KEY",
            "ARK_API_KEY",
            "VOLCENGINE_API_KEY",
        ]
    )
    return LocalLLMConfig(provider=provider, endpoint=endpoint, model=model, api_key=api_key, timeout_seconds=max(5.0, timeout))


def generic_llm_system_prompt() -> str:
    return "你是一个可靠的中文助手。回答要直接、简洁。"


def doubao_question_solve_system_prompt() -> str:
    return "你在帮助用户根据朋友圈正文回答剧本杀/活动相关问题。只输出最终答案；不知道就输出“不知道”。"


def build_generic_llm_user_prompt(prompt: str, context: str = "") -> str:
    if context.strip():
        return f"上下文：\n{context.strip()}\n\n问题：\n{prompt.strip()}"
    return prompt.strip()


def build_doubao_question_prompt(post_text: str) -> str:
    return f"请根据下面朋友圈正文回答问题，只输出答案：\n\n{post_text.strip()}"


def parse_openai_compatible_response(root: dict[str, Any]) -> Optional[str]:
    try:
        content = root["choices"][0]["message"]["content"]
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, dict):
                    parts.append(str(item.get("text") or item.get("content") or ""))
            return "".join(parts).strip()
    except Exception:
        return None
    return None


def parse_responses_api_response(root: dict[str, Any]) -> Optional[str]:
    text = root.get("output_text")
    if isinstance(text, str) and text.strip():
        return text
    parts: list[str] = []
    for output in root.get("output") or []:
        if not isinstance(output, dict):
            continue
        for item in output.get("content") or []:
            if not isinstance(item, dict):
                continue
            value = item.get("text") or item.get("value")
            if isinstance(value, str):
                parts.append(value)
    return "".join(parts).strip() or None


def parse_ollama_response(root: dict[str, Any]) -> Optional[str]:
    message = root.get("message")
    if isinstance(message, dict) and isinstance(message.get("content"), str):
        return message["content"]
    if isinstance(root.get("response"), str):
        return root["response"]
    return None


def clean_llm_answer(text: str) -> str:
    answer = text.strip()
    answer = re.sub(r"^\s*(答案|回答)\s*[:：]\s*", "", answer)
    return answer.strip().strip('"').strip("'")


def request_llm_answer(
    config: LocalLLMConfig,
    system_prompt: str,
    user_prompt: str,
    image_data_urls: Optional[list[str]] = None,
) -> Optional[str]:
    requests = require_module("requests")
    images = [url for url in (image_data_urls or []) if url.strip()]
    headers = {"Content-Type": "application/json"}
    if config.api_key:
        headers["Authorization"] = f"Bearer {config.api_key}"

    if config.provider == "ollama":
        body: dict[str, Any] = {
            "model": config.model,
            "stream": False,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
    elif config.provider == "openai":
        content: Any = user_prompt
        if images:
            content = [{"type": "text", "text": user_prompt}]
            content.extend({"type": "image_url", "image_url": {"url": url}} for url in images)
        body = {
            "model": config.model,
            "stream": False,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content},
            ],
            "temperature": 0.0,
            "max_tokens": 256,
        }
    else:
        if "/responses" in config.endpoint:
            user_content: list[dict[str, Any]] = [{"type": "input_text", "text": user_prompt}]
            user_content[0:0] = [{"type": "input_image", "image_url": url, "detail": "auto"} for url in images]
            body = {
                "model": config.model,
                "stream": False,
                "max_output_tokens": 256,
                "thinking": {"type": "disabled"},
                "input": [
                    {"type": "message", "role": "system", "content": system_prompt},
                    {"type": "message", "role": "user", "content": user_content},
                ],
            }
        else:
            content = [{"type": "text", "text": user_prompt}]
            content.extend({"type": "image_url", "image_url": {"url": url}} for url in images)
            body = {
                "model": config.model,
                "stream": False,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": content},
                ],
                "temperature": 0.0,
                "max_tokens": 256,
            }

    try:
        response = requests.post(config.endpoint, headers=headers, json=body, timeout=config.timeout_seconds)
    except Exception as exc:
        print(f"  LLM 请求失败: {exc}")
        return None
    if not 200 <= response.status_code <= 299:
        print(f"  LLM 请求失败: HTTP {response.status_code}")
        print(f"  响应: {response.text[:1000]}")
        return None
    try:
        root = response.json()
    except Exception:
        print(f"  LLM 响应不是 JSON: {response.text[:1000]}")
        return None
    if config.provider == "ollama":
        answer = parse_ollama_response(root)
    elif "/responses" in config.endpoint:
        answer = parse_responses_api_response(root)
    else:
        answer = parse_openai_compatible_response(root)
    return clean_llm_answer(answer) if answer else None


def ask_local_llm(prompt: str, context: str = "") -> Optional[str]:
    config = load_local_llm_config()
    if not config:
        print_ts("LLM 配置缺失：请设置 EASYMONEY_LLM_MODEL 或 DOUBAO/ARK_MODEL")
        return None
    return request_llm_answer(config, generic_llm_system_prompt(), build_generic_llm_user_prompt(prompt, context))


def ask_doubao_to_solve_post(post_text: str, image_data_urls: Optional[list[str]] = None) -> Optional[SolvedQuestion]:
    config = load_local_llm_config()
    if not config:
        print_ts("豆包/LLM 配置缺失：请检查 .easyMoney.env")
        return None
    answer = request_llm_answer(config, doubao_question_solve_system_prompt(), build_doubao_question_prompt(post_text), image_data_urls=image_data_urls)
    if not answer:
        return None
    return SolvedQuestion(answer=answer, evidence="LLM", confidence=0.62, source=config.provider)


def image_to_data_url(image: Any, max_side: int = 1280, quality: int = 78) -> str:
    from io import BytesIO

    img = image.convert("RGB")
    if max(img.size) > max_side:
        scale = max_side / max(img.size)
        img = img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))))
    buf = BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def capture_post_image(post: MomentPostResolution, window_rect: Rect) -> Any:
    capture = CaptureBackend()
    region = post.body_frame.expanded(12, 24).clamp_to(window_rect)
    try:
        return capture.screenshot(region)
    finally:
        capture.close()


def direct_uia_inline_image_rects(body_frame: Rect, image_count: int, window_rect: Rect) -> list[Rect]:
    count = max(1, min(image_count, 9))
    side = 120.0
    gap = 4.0
    left = body_frame.left + 76.0
    bottom = body_frame.bottom - 32.0
    window_inner = window_rect.inset(8, 8)

    rects: list[Rect] = []
    for index in range(count):
        if count in {2, 3}:
            col = index
            row = 0
            rows_above_bottom = 0
        elif count == 4:
            col = index % 2
            row = index // 2
            rows_above_bottom = 1 - row
        elif count in {5, 6}:
            col = index if index < 3 else index - 3
            row = 0 if index < 3 else 1
            rows_above_bottom = 1 - row
        else:
            col = index % 3
            row = index // 3
            rows_above_bottom = 2 - row

        rect_bottom = bottom - rows_above_bottom * (side + gap)
        rect = Rect(
            left + col * (side + gap),
            rect_bottom - side,
            left + col * (side + gap) + side,
            rect_bottom,
        ).clamp_to(window_inner)
        if rect.width > 24 and rect.height > 24:
            rects.append(rect)
    return rects


def inline_image_rows(image_count: int) -> list[list[int]]:
    count = max(1, min(image_count, 9))
    if count == 2:
        return [[0, 1]]
    if count == 3:
        return [[0, 1, 2]]
    if count == 4:
        return [[0, 1], [2, 3]]
    if count == 5:
        return [[0, 1, 2], [3, 4]]
    if count == 6:
        return [[0, 1, 2], [3, 4, 5]]
    if count == 7:
        return [[0, 1, 2], [3, 4, 5], [6]]
    if count == 8:
        return [[0, 1, 2], [3, 4, 5], [6, 7]]
    if count == 9:
        return [[0, 1, 2], [3, 4, 5], [6, 7, 8]]
    return [[0]]


def image_crop_box_from_screen_rect(screen_rect: Rect, source_frame: Rect, image: Any) -> tuple[int, int, int, int]:
    scale_x = image.width / max(1.0, source_frame.width)
    scale_y = image.height / max(1.0, source_frame.height)
    left = math.floor((screen_rect.left - source_frame.left) * scale_x)
    top = math.floor((screen_rect.top - source_frame.top) * scale_y)
    right = math.ceil((screen_rect.right - source_frame.left) * scale_x)
    bottom = math.ceil((screen_rect.bottom - source_frame.top) * scale_y)
    return (
        max(0, min(image.width, left)),
        max(0, min(image.height, top)),
        max(0, min(image.width, right)),
        max(0, min(image.height, bottom)),
    )


def inline_image_looks_loaded(image: Any) -> bool:
    rgb = image.convert("RGB")
    width, height = rgb.size
    step_x = max(1, width // 32)
    step_y = max(1, height // 32)
    pixels = rgb.load()
    samples = 0
    bright_low_chroma = 0
    dark_or_colorful = 0
    brightness_sum = 0.0
    brightness_squared_sum = 0.0

    for y in range(0, height, step_y):
        for x in range(0, width, step_x):
            r, g, b = pixels[x, y]
            max_rgb = max(r, g, b)
            min_rgb = min(r, g, b)
            chroma = max_rgb - min_rgb
            brightness = (r + g + b) / 3.0
            samples += 1
            brightness_sum += brightness
            brightness_squared_sum += brightness * brightness
            if brightness >= 224 and chroma <= 18:
                bright_low_chroma += 1
            if brightness <= 210 or chroma >= 24:
                dark_or_colorful += 1

    if samples <= 0:
        return False
    bright_low_chroma_ratio = bright_low_chroma / samples
    dark_or_colorful_ratio = dark_or_colorful / samples
    mean = brightness_sum / samples
    variance = max(0.0, brightness_squared_sum / samples - mean * mean)
    stddev = math.sqrt(variance)
    return not (bright_low_chroma_ratio >= 0.92 and dark_or_colorful_ratio <= 0.05 and stddev <= 18)


def inline_image_load_timeout_ms() -> int:
    timeout_ms = max(
        80,
        min(
            int(first_non_empty_env(["EASYMONEY_AX_IMAGE_LOAD_TIMEOUT_MS", "EASYMONEY_DOUBAO_AX_IMAGE_LOAD_TIMEOUT_MS"]) or "1200"),
            8000,
        ),
    )
    return timeout_ms


def inline_image_load_interval_ms() -> int:
    return max(
        0,
        min(
            int(first_non_empty_env(["EASYMONEY_AX_IMAGE_LOAD_INTERVAL_MS", "EASYMONEY_DOUBAO_AX_IMAGE_LOAD_INTERVAL_MS"]) or "0"),
            1000,
        ),
    )


def rect_intersects(lhs: Rect, rhs: Rect) -> bool:
    return lhs.left < rhs.right and lhs.right > rhs.left and lhs.top < rhs.bottom and lhs.bottom > rhs.top


def crop_loaded_inline_image_regions(rects: list[Rect], window_rect: Rect, label: str) -> list[Any]:
    if not rects:
        return []
    timeout_ms = inline_image_load_timeout_ms()
    interval_ms = inline_image_load_interval_ms()
    deadline = time.perf_counter() + timeout_ms / 1000.0
    loaded: list[Any | None] = [None] * len(rects)
    attempts = 0
    capture = CaptureBackend()
    try:
        while time.perf_counter() < deadline and any(image is None for image in loaded):
            full_window_image = capture.screenshot_stream(window_rect)
            attempts += 1
            for index, screen_rect in enumerate(rects):
                if loaded[index] is not None:
                    continue
                left, top, right, bottom = image_crop_box_from_screen_rect(screen_rect, window_rect, full_window_image)
                if right - left <= 1 or bottom - top <= 1:
                    continue
                cropped = full_window_image.crop((left, top, right, bottom))
                if inline_image_looks_loaded(cropped):
                    loaded[index] = cropped
            if any(image is None for image in loaded) and interval_ms > 0:
                time.sleep(interval_ms / 1000.0)
    finally:
        capture.close()

    images = [image for image in loaded if image is not None]
    if len(images) < len(rects):
        print_ts(f"  {label}加载等待超时: 已加载{len(images)}/{len(rects)}，尝试{attempts}帧，等待{timeout_ms}ms")
    else:
        print_ts(f"  {label}加载完成: {len(images)}/{len(rects)}，尝试{attempts}帧")
    return images


def capture_uia_inline_image_regions(post: MomentPostResolution, window_rect: Rect, image_count: int) -> list[Any]:
    rects = direct_uia_inline_image_rects(post.body_frame, image_count, window_rect)
    return crop_loaded_inline_image_regions(rects, window_rect, "UIA图片")


def focused_element_description(backend: WindowBackend, element: Any, rect: Rect) -> str:
    name = backend._safe_text(element).strip()
    control_type = backend._control_type(element)
    class_name = backend._class_name(element)
    return f"type={control_type or '?'} name={name or '(空)'} class={class_name or '(空)'} rect={rect.describe()}"


def focused_element_is_image_button(backend: WindowBackend, element: Any) -> bool:
    name = backend._safe_text(element).strip()
    control_type = backend._control_type(element).strip().lower()
    class_name = backend._class_name(element)
    is_button = control_type in {"button", "按钮"} or "button" in control_type
    return is_button and (name == "图片" or class_name == "mmui::XMouseEventView")


def locate_single_uia_inline_image_rect(post: MomentPostResolution, window_rect: Rect) -> Rect:
    backend = WindowBackend()
    input_backend = InputBackend()

    input_backend.prepare_key_sequence(SINGLE_IMAGE_FOCUS_KEYS)
    key_gap_ms = max(0, min(int(first_non_empty_env(["EASYMONEY_SINGLE_IMAGE_KEY_GAP_MS"]) or "30"), 1000))
    input_backend.press_sequence(SINGLE_IMAGE_FOCUS_KEYS, gap=key_gap_ms / 1000.0)

    focus_wait_ms = max(0, min(int(first_non_empty_env(["EASYMONEY_SINGLE_IMAGE_FOCUS_WAIT_MS"]) or "120"), 3000))
    if focus_wait_ms > 0:
        time.sleep(focus_wait_ms / 1000.0)

    automation, _ = backend._ensure_automation()
    focused = automation.GetFocusedElement()
    if focused is None:
        raise EasyMoneyError("单图键盘定位失败: 未读取到当前焦点元素")
    if backend._safe_text(focused).strip() != "图片":
        input_backend.prepare_key_sequence(("tab",))
        input_backend.press_sequence(("tab",), gap=key_gap_ms / 1000.0)
        if focus_wait_ms > 0:
            time.sleep(focus_wait_ms / 1000.0)
        focused = automation.GetFocusedElement()
        if focused is None:
            raise EasyMoneyError("单图键盘定位失败: 兜底Tab后未读取到当前焦点元素")
    focus_rect = backend.rect(focused)
    if focus_rect is None:
        raise EasyMoneyError("单图键盘定位失败: 当前焦点元素没有有效矩形")
    if not focused_element_is_image_button(backend, focused):
        raise EasyMoneyError(f"单图键盘定位失败: 当前焦点不是图片按钮 ({focused_element_description(backend, focused, focus_rect)})")
    if not rect_intersects(focus_rect, post.body_frame):
        raise EasyMoneyError(
            "单图键盘定位失败: 当前焦点图片不在目标动态区域内 "
            f"(focus={focus_rect.describe()} post={post.body_frame.describe()})"
        )

    crop_rect = focus_rect.clamp_to(window_rect)
    if crop_rect.width <= 1 or crop_rect.height <= 1:
        raise EasyMoneyError(f"单图键盘定位失败: 焦点截图区域无效 {crop_rect.describe()}")
    print_ts(f"  单图键盘定位成功: {focused_element_description(backend, focused, crop_rect)}")
    return crop_rect


def capture_single_uia_inline_image_region(post: MomentPostResolution, window_rect: Rect) -> Any | None:
    rect = locate_single_uia_inline_image_rect(post, window_rect)
    images = crop_loaded_inline_image_regions([rect], window_rect, "UIA单图")
    return images[0] if images else None


def stitch_inline_images(images: list[Any], image_count: int) -> Any | None:
    rows = inline_image_rows(image_count)
    if not images or not rows:
        return None
    tile_width = min(image.width for image in images)
    tile_height = min(image.height for image in images)
    if tile_width <= 0 or tile_height <= 0:
        return None
    columns = max(len(row) for row in rows)
    image_module = require_module("PIL.Image", "Pillow")
    stitched = image_module.new("RGB", (columns * tile_width, len(rows) * tile_height), "white")
    for row_index, row in enumerate(rows):
        for column_index, image_index in enumerate(row):
            if image_index < 0 or image_index >= len(images):
                continue
            image = images[image_index].convert("RGB")
            if image.size != (tile_width, tile_height):
                image = image.resize((tile_width, tile_height))
            stitched.paste(image, (column_index * tile_width, row_index * tile_height))
    return stitched

def ctrl_left_click(point: Point) -> None:
    """模拟 Ctrl+鼠标左键单击（用于点击朋友圈图片触发预览等操作）"""
    input_backend = InputBackend()
    if input_backend.native and input_backend.user32 is not None:
        ctrl_vk = InputBackend._vk("ctrl")
        input_backend.user32.keybd_event(ctrl_vk, 0, 0, 0)
        time.sleep(0.02)
        input_backend.click(point)
        time.sleep(0.02)
        input_backend.user32.keybd_event(ctrl_vk, 0, 0x0002, 0)
    else:
        import pyautogui
        pyautogui.keyDown("ctrl")
        input_backend.click(point)
        pyautogui.keyUp("ctrl")


def save_one_picture(post: MomentPostResolution, window_rect: Rect) -> Any:
    """图片数量为 1 时，Ctrl+左键单击预览图片，截图"图片和视频"窗口（去掉顶部40px），返回 PIL Image 或 None"""
    target_point = Point(post.body_frame.left + 100, post.action_point.y - 70)
    print(target_point.x, target_point.y)
    ctrl_left_click(target_point)
    # 查找微信"图片和视频"窗口（带重试等待窗口出现）
    import ctypes
    user32 = ctypes.windll.user32
    hwnd = None
    for _ in range(10):
        time.sleep(0.15)
        hwnd = user32.FindWindowW(None, "图片和视频")
        if hwnd:
            break
    if not hwnd:
        print_ts(f"  '图片和视频' 窗口未找到")
        return None
    rect_struct = ctypes.wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect_struct))
    pic_rect = Rect(rect_struct.left, rect_struct.top, rect_struct.right, rect_struct.bottom)
    print_ts(f"  '图片和视频' 窗口找到 (HWND={hwnd:#x})")
    print_ts(f"    位置: ({pic_rect.left}, {pic_rect.top}) - ({pic_rect.right}, {pic_rect.bottom})")
    print_ts(f"    尺寸: {pic_rect.width}x{pic_rect.height}")
    # 切换到"图片和视频"窗口
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.3)
    # 截图窗口区域，去掉顶部 40px 的标题栏/工具栏
    crop_rect = Rect(pic_rect.left, pic_rect.top + 40, pic_rect.right, pic_rect.bottom)
    capture = CaptureBackend()
    try:
        image = capture.screenshot(crop_rect)
    finally:
        capture.close()
    print_ts(f"  窗口截图成功: {image.width}x{image.height}（已裁切顶部40px）")
    return image
def capture_vision_image_data_urls(
    post: MomentPostResolution,
    window_rect: Rect,
    save_path: Optional[Path] = None,
) -> list[str]:
    image_count = post.inline_image_count or extract_inline_image_count(post.text)
    if image_count is not None and 2 <= image_count <= 9:
        images = capture_uia_inline_image_regions(post, window_rect, image_count)
        if len(images) != image_count:
            raise EasyMoneyError(f"--LLM --vision UIA直接裁剪数量不完整: {len(images)}/{image_count}")
        stitched = stitch_inline_images(images, image_count)
        if stitched is None:
            raise EasyMoneyError("--LLM --vision UIA裁剪图拼接失败")
        image = stitched
        print_ts(f"  UIA裁剪图已去除4px间隔并拼接: {image.width}x{image.height}")
    elif image_count is not None and image_count == 1:
        image = capture_single_uia_inline_image_region(post, window_rect)
        if image is None:
            raise EasyMoneyError("--LLM --vision UIA单图裁剪未加载完成")
    else:
        if image_count is None:
            print_ts("  UIA未检测到图片数量，使用整条动态区域截图")
        else:
            print_ts(f"  UIA检测到该动态包含{image_count}张图片，使用整条动态区域截图")
        image = capture_post_image(post, window_rect)
    if save_path is not None:
        ensure_parent(save_path)
        image.save(save_path)
    return [image_to_data_url(image)]

