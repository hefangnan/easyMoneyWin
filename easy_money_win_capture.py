from __future__ import annotations

from easy_money_win_core import *

class CaptureBackend:
    def __init__(self, backend: Optional[str] = None) -> None:
        self.Image = require_module("PIL.Image", "Pillow")
        requested_backend = (backend or os.environ.get("EASYMONEY_CAPTURE_BACKEND") or "dxgi").strip().lower()
        if requested_backend not in {"dxgi", "dxcam"}:
            raise EasyMoneyError("截图后端只支持 DXGI；请移除 EASYMONEY_CAPTURE_BACKEND 或设置为 dxgi")
        self.backend = "dxgi"
        self._dxcam_mod = None
        self._dx_camera = None
        self._dx_stream_region: Optional[tuple[int, int, int, int]] = None
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
            if self._dx_camera is None:
                raise EasyMoneyError("DXGI 捕获初始化失败: dxcam.create 返回空对象")
        except ImportError as exc:
            raise EasyMoneyError("缺少 DXGI 依赖 `dxcam`，请运行: python -m pip install -r requirements.txt") from exc
        except EasyMoneyError:
            raise
        except Exception as exc:
            raise EasyMoneyError(f"DXGI 捕获初始化失败: {exc}") from exc

    def grab(self, rect: Rect) -> CaptureFrame:
        if rect.width <= 0 or rect.height <= 0:
            raise EasyMoneyError(f"截图区域无效: {rect.describe()}")
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
            raise CaptureUnavailable(f"DXGI 截图失败: {rect.describe()} ({exc})") from exc
        if frame is None:
            raise EasyMoneyError(f"DXGI 截图失败: {rect.describe()}")
        height, width = int(frame.shape[0]), int(frame.shape[1])
        return CaptureFrame(width=width, height=height, rgb=frame.tobytes())

    def grab_stream(self, rect: Rect) -> CaptureFrame:
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
                raise CaptureUnavailable(f"DXGI 流截图启动失败: {rect.describe()} ({exc})") from exc
            self._dx_stream_region = region
        try:
            frame = self._dx_camera.get_latest_frame(copy=True)
        except Exception as exc:
            raise CaptureUnavailable(f"DXGI 流截图失败: {rect.describe()} ({exc})") from exc
        if frame is None:
            self._stop_dx_stream()
            return self.grab(rect)
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

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass


def refresh_observation_region(window_rect: Rect) -> Rect:
    return Rect(
        window_rect.left,
        window_rect.top + min(60, max(0, window_rect.height * 0.08)),
        window_rect.left + max(1, window_rect.width / 7),
        window_rect.top + max(80, window_rect.height * 0.62),
    ).clamp_to(window_rect)


_UIA_CONTROL_TYPE_NAMES: Optional[dict[int, str]] = None


