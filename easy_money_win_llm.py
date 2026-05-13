from __future__ import annotations

import math

from easy_money_win_core import *
from easy_money_win_capture import CaptureBackend

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


def capture_uia_inline_image_regions(post: MomentPostResolution, window_rect: Rect, image_count: int) -> list[Any]:
    rects = direct_uia_inline_image_rects(post.body_frame, image_count, window_rect)
    if not rects:
        return []
    timeout_ms = max(
        80,
        min(
            int(first_non_empty_env(["EASYMONEY_AX_IMAGE_LOAD_TIMEOUT_MS", "EASYMONEY_DOUBAO_AX_IMAGE_LOAD_TIMEOUT_MS"]) or "1200"),
            8000,
        ),
    )
    interval_ms = max(
        0,
        min(
            int(first_non_empty_env(["EASYMONEY_AX_IMAGE_LOAD_INTERVAL_MS", "EASYMONEY_DOUBAO_AX_IMAGE_LOAD_INTERVAL_MS"]) or "0"),
            1000,
        ),
    )
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
        print_ts(f"  UIA图片加载等待超时: 已加载{len(images)}/{len(rects)}，尝试{attempts}帧，等待{timeout_ms}ms")
    else:
        print_ts(f"  UIA图片加载完成: {len(images)}/{len(rects)}，尝试{attempts}帧")
    return images


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

