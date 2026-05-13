from __future__ import annotations

from easy_money_win_core import *

class CaptureBackend:
    def __init__(self, backend: Optional[str] = None) -> None:
        self.Image = require_module("PIL.Image", "Pillow")
        requested_backend = (backend or os.environ.get("EASYMONEY_CAPTURE_BACKEND") or "auto").strip().lower()
        self.backend = requested_backend
        self._allow_mss_fallback = requested_backend == "auto"
        self._dxcam_mod = None
        self._dx_camera = None
        self._dx_stream_region: Optional[tuple[int, int, int, int]] = None
        self.mss_mod = None
        self._sct = None
        dxgi_ready = False
        if self.backend in {"auto", "dxgi", "dxcam"}:
            try:
                self._dxcam_mod = importlib.import_module("dxcam")
                output_idx = int(os.environ.get("EASYMONEY_DXGI_OUTPUT", "0"))
                processor_backend = (os.environ.get("EASYMONEY_DXGI_PROCESSOR_BACKEND") or "numpy").strip().lower()
                try:
                    self._dx_camera = self._dxcam_mod.create(
                        output_idx=output_idx,
                        output_color="RGB",
                        processor_backend=processor_backend,
                    )
                except TypeError:
                    self._dx_camera = self._dxcam_mod.create(output_idx=output_idx, output_color="RGB")
                if self._dx_camera is not None:
                    self.backend = "dxgi"
                    dxgi_ready = True
            except ImportError as exc:
                if self.backend in {"dxgi", "dxcam"}:
                    raise EasyMoneyError("缺少 DXGI 依赖 `dxcam`，请运行: python -m pip install -r requirements.txt") from exc
            except Exception as exc:
                if self.backend in {"dxgi", "dxcam"}:
                    raise EasyMoneyError(f"DXGI 捕获初始化失败: {exc}") from exc
        if dxgi_ready and not self._allow_mss_fallback:
            return
        self._init_mss(required=not dxgi_ready)
        if not dxgi_ready:
            self.backend = "mss"

    def _init_mss(self, required: bool = True) -> None:
        try:
            self.mss_mod = require_module("mss")
        except EasyMoneyError:
            if required:
                raise
            return
        if hasattr(self.mss_mod, "MSS"):
            self._sct = self.mss_mod.MSS()
        else:
            self._sct = self.mss_mod.mss()

    def _grab_mss(self, rect: Rect) -> CaptureFrame:
        if self._sct is None:
            raise CaptureUnavailable("截图后端未初始化")
        try:
            shot = self._sct.grab(rect.to_mss())
        except Exception as exc:
            raise CaptureUnavailable(f"MSS 截图失败: {rect.describe()} ({exc})") from exc
        return CaptureFrame(width=int(shot.width), height=int(shot.height), rgb=shot.rgb)

    def _grab_mss_fallback(self, rect: Rect, exc: Exception) -> CaptureFrame:
        if not self._allow_mss_fallback:
            raise CaptureUnavailable(
                f"DXGI 截图区域无效: {rect.describe()} ({exc})；"
                "窗口可能在副屏或跨屏区域，请移动到主屏，或设置 EASYMONEY_CAPTURE_BACKEND=mss"
            ) from exc
        if self._sct is None:
            self._init_mss(required=True)
        self._stop_dx_stream()
        return self._grab_mss(rect)

    def grab(self, rect: Rect) -> CaptureFrame:
        if rect.width <= 0 or rect.height <= 0:
            raise EasyMoneyError(f"截图区域无效: {rect.describe()}")
        if self.backend == "dxgi" and self._dx_camera is not None:
            region = (
                int(round(rect.left)),
                int(round(rect.top)),
                int(round(rect.right)),
                int(round(rect.bottom)),
            )
            if getattr(self._dx_camera, "is_capturing", False) and self._dx_stream_region != region:
                self._stop_dx_stream()
            try:
                frame = self._dx_camera.grab(region=region)
            except Exception as exc:
                return self._grab_mss_fallback(rect, exc)
            if frame is None:
                raise EasyMoneyError(f"DXGI 截图失败: {rect.describe()}")
            height, width = int(frame.shape[0]), int(frame.shape[1])
            return CaptureFrame(width=width, height=height, rgb=frame.tobytes())
        return self._grab_mss(rect)

    def grab_stream(self, rect: Rect) -> CaptureFrame:
        if self.backend != "dxgi" or self._dx_camera is None:
            return self.grab(rect)
        if rect.width <= 0 or rect.height <= 0:
            raise EasyMoneyError(f"截图区域无效: {rect.describe()}")
        region = (
            int(round(rect.left)),
            int(round(rect.top)),
            int(round(rect.right)),
            int(round(rect.bottom)),
        )
        if self._dx_stream_region != region or not getattr(self._dx_camera, "is_capturing", False):
            self._stop_dx_stream()
            fps = int(os.environ.get("EASYMONEY_DXGI_STREAM_FPS", "240"))
            try:
                self._dx_camera.start(region=region, target_fps=fps, video_mode=True)
            except Exception as exc:
                return self._grab_mss_fallback(rect, exc)
            self._dx_stream_region = region
        try:
            frame = self._dx_camera.get_latest_frame(copy=True)
        except Exception as exc:
            return self._grab_mss_fallback(rect, exc)
        if frame is None:
            self._stop_dx_stream()
            try:
                return self.grab(rect)
            except Exception as exc:
                return self._grab_mss_fallback(rect, exc)
        height, width = int(frame.shape[0]), int(frame.shape[1])
        return CaptureFrame(width=width, height=height, rgb=frame.tobytes())

    def _stop_dx_stream(self) -> None:
        stop = getattr(self._dx_camera, "stop", None)
        if callable(stop):
            try:
                stop()
            except Exception:
                pass
        self._dx_stream_region = None

    def screenshot(self, rect: Rect):
        shot = self.grab(rect)
        return self.Image.frombytes("RGB", (shot.width, shot.height), shot.rgb)

    def screenshot_stream(self, rect: Rect):
        shot = self.grab_stream(rect)
        return self.Image.frombytes("RGB", (shot.width, shot.height), shot.rgb)

    def save(self, image: Any, path: Path) -> Path:
        ensure_parent(path)
        image.save(path)
        return path

    def close(self) -> None:
        self._stop_dx_stream()
        close = getattr(self._sct, "close", None)
        if callable(close):
            close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass


def quick_capture_fingerprint(capture: CaptureBackend, rect: Rect) -> Optional[int]:
    try:
        shot = capture.grab(rect)
    except Exception:
        return None
    width, height = int(shot.width), int(shot.height)
    if width <= 0 or height <= 0:
        return None
    data = shot.rgb
    samples: list[int] = []
    for row in range(8):
        y = min(height - 1, max(0, int((row + 0.5) * height / 8)))
        for col in range(8):
            x = min(width - 1, max(0, int((col + 0.5) * width / 8)))
            idx = (y * width + x) * 3
            r, g, b = data[idx], data[idx + 1], data[idx + 2]
            samples.append((int(r) * 30 + int(g) * 59 + int(b) * 11) // 100)
    avg = sum(samples) / max(1, len(samples))
    fingerprint = 0
    for value in samples:
        fingerprint = (fingerprint << 1) | (1 if value >= avg else 0)
    return fingerprint


def fingerprint_distance(lhs: int, rhs: int) -> int:
    return (lhs ^ rhs).bit_count()


def wait_for_region_refresh(
    capture: CaptureBackend,
    region: Rect,
    baseline_fingerprint: Optional[int],
    timeout_seconds: float = COMMENT_REFRESH_WAIT_SECONDS,
) -> bool:
    deadline = time.perf_counter() + max(0.001, timeout_seconds)
    next_check = time.perf_counter()
    while time.perf_counter() < deadline:
        now = time.perf_counter()
        if now >= next_check:
            current = quick_capture_fingerprint(capture, region)
            if current is not None:
                if baseline_fingerprint is None:
                    return True
                if fingerprint_distance(baseline_fingerprint, current) >= 10:
                    return True
            next_check = now + COMMENT_REFRESH_CAPTURE_INTERVAL_SECONDS
        else:
            time.sleep(COMMENT_REFRESH_IDLE_SECONDS)
    return False


def refresh_observation_region(window_rect: Rect) -> Rect:
    return Rect(
        window_rect.left,
        window_rect.top + min(60, max(0, window_rect.height * 0.08)),
        window_rect.left + max(1, window_rect.width / 7),
        window_rect.top + max(80, window_rect.height * 0.62),
    ).clamp_to(window_rect)


_UIA_CONTROL_TYPE_NAMES: Optional[dict[int, str]] = None


