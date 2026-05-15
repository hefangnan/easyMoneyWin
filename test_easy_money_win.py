import io
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

import easy_money_win as em
import easy_money_win_capture as em_capture
import easy_money_win_commands as em_commands
import easy_money_win_core as em_core
import easy_money_win_llm as em_llm
import easy_money_win_uia as em_uia


class EasyMoneyWinTests(unittest.TestCase):
    def test_point_and_comment_config_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_path = em_core.CONFIG_COMMENT
            try:
                em_core.CONFIG_COMMENT = Path(tmp) / ".wechat_comment_config"
                config = em.CommentConfig(
                    comment_from_action=em.Point(10, 20),
                    send_x_ratio=0.82,
                    send_from_action=em.Point(30, 40),
                    fixed_send_action_y_threshold=500,
                    fixed_send_window_offset=em.Point(700, 650),
                )
                em.save_comment_config(config)
                loaded = em.load_comment_config()
                self.assertIsNotNone(loaded)
                self.assertEqual(loaded.comment_from_action, em.Point(10, 20))
                self.assertEqual(loaded.send_from_action, em.Point(30, 40))
                self.assertEqual(loaded.fixed_send_window_offset, em.Point(700, 650))
            finally:
                em_core.CONFIG_COMMENT = old_path

    def test_dotenv_loading(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = os.getcwd()
            old_cache = em_llm._DOTENV_CACHE
            try:
                os.chdir(tmp)
                Path(".easyMoney.env").write_text(
                    "export EASYMONEY_LLM_PROVIDER=doubao\nARK_MODEL='demo-model'\n",
                    encoding="utf-8",
                )
                em_llm._DOTENV_CACHE = None
                self.assertEqual(em.first_non_empty_env(["EASYMONEY_LLM_PROVIDER"]), "doubao")
                self.assertEqual(em.first_non_empty_env(["ARK_MODEL"]), "demo-model")
            finally:
                os.chdir(old_cwd)
                em_llm._DOTENV_CACHE = old_cache

    def test_llm_response_parsers(self):
        self.assertEqual(
            em.parse_openai_compatible_response({"choices": [{"message": {"content": "答案"}}]}),
            "答案",
        )
        self.assertEqual(
            em.parse_ollama_response({"message": {"content": "本地答案"}}),
            "本地答案",
        )
        self.assertEqual(
            em.parse_responses_api_response({"output_text": "豆包答案"}),
            "豆包答案",
        )

    def test_window_backend_uses_pywinauto_desktop(self):
        old_require_module = em_uia.require_module
        calls: list[tuple[str, object]] = []

        class FakePywinauto:
            @staticmethod
            def Desktop(backend: str) -> object:
                calls.append(("Desktop", backend))
                return object()

        class FakeTimingsModule:
            class Timings:
                window_find_timeout = 5.0

        def fake_require_module(module_name: str, pip_name: str | None = None) -> object:
            calls.append((module_name, pip_name))
            if module_name == "pywinauto":
                return FakePywinauto
            if module_name == "pywinauto.timings":
                return FakeTimingsModule
            raise AssertionError(f"unexpected module: {module_name}")

        try:
            em_uia.require_module = fake_require_module
            backend = em.WindowBackend()

            self.assertIsNotNone(backend.desktop)
            self.assertIn(("pywinauto", None), calls)
            self.assertIn(("Desktop", "uia"), calls)
            self.assertNotIn(("uiautomation", None), calls)
        finally:
            em_uia.require_module = old_require_module

    def test_sns_list_lookup_prefers_fast_path(self):
        sentinel = object()

        class FakeBackend:
            def __init__(self) -> None:
                self.iter_calls = 0

            def find_sns_list_fast(self, root: object) -> object:
                return sentinel

            def iter_tree(self, root: object, max_depth: int = 8):
                self.iter_calls += 1
                return iter(())

        backend = FakeBackend()
        self.assertIs(em.find_sns_list_control(backend, object()), sentinel)
        self.assertEqual(backend.iter_calls, 0)

    def test_list_item_lookup_prefers_direct_children(self):
        class FakeNode:
            def __init__(self, kind: str) -> None:
                self.kind = kind

        class FakeBackend:
            def listitem_children(self, root: object, limit=None):
                return [FakeNode("listitem"), FakeNode("listitem")]

            def children(self, root: object):
                raise AssertionError("children traversal should not run when direct list items are available")

            def _control_identity(self, node: FakeNode):
                return id(node)

            def _control_type(self, node: FakeNode):
                return node.kind

        items = em.find_list_items_under_control(FakeBackend(), object(), limit=1)
        self.assertEqual(len(items), 1)

    def test_list_item_lookup_reuses_successful_strategy(self):
        class FakeNode:
            def __init__(self, kind: str) -> None:
                self.kind = kind

        root = object()

        class FakeBackend:
            def __init__(self) -> None:
                self._list_item_strategy_cache: dict[object, str] = {}
                self.calls: list[str] = []

            def listitem_children(self, root: object, limit=None):
                self.calls.append("listitem_children")
                return []

            def children(self, root: object):
                self.calls.append("children")
                return [FakeNode("listitem")]

            def iter_tree(self, root: object, max_depth: int = 3):
                raise AssertionError("tree traversal should not run after children succeeds")

            def _control_identity(self, node: object):
                return node if node is root else id(node)

            def _control_type(self, node: FakeNode):
                return node.kind

        backend = FakeBackend()
        self.assertEqual(len(em.find_list_items_under_control(backend, root, limit=1)), 1)
        self.assertEqual(backend.calls, ["listitem_children", "children"])

        backend.calls = []
        self.assertEqual(len(em.find_list_items_under_control(backend, root, limit=1)), 1)
        self.assertEqual(backend.calls, ["children"])

    def test_default_dump_uses_sns_list_fast_path(self):
        class FakeNode:
            def __init__(self, kind: str, name: str = "", automation_id: str = "") -> None:
                self.kind = kind
                self.name = name
                self.automation_id = automation_id

        sns_list = FakeNode("list", "朋友圈", "sns_list")

        class FakeBackend:
            def __init__(self) -> None:
                self.item_limit = None

            def find_sns_list_fast(self, root: object) -> FakeNode:
                return sns_list

            def iter_tree(self, root: object, max_depth: int = 8):
                raise AssertionError("raw tree traversal should not run for default sns_list dump")

            def _control_type(self, node: FakeNode) -> str:
                return node.kind

            def _safe_text(self, node: FakeNode) -> str:
                return node.name

            def _automation_id(self, node: FakeNode) -> str:
                return node.automation_id

            def rect(self, node: FakeNode):
                return em.Rect(0, 0, 100, 20)

            def listitem_children(self, root: FakeNode, limit=None):
                self.item_limit = limit
                return [FakeNode("listitem", "ignored"), FakeNode("listitem", "fn post")]

            def _control_identity(self, node: FakeNode):
                return id(node)

        backend = FakeBackend()
        found_sns, found_item, item_count = em.dump_named_list_contents(
            backend,
            object(),
            "朋友圈",
            quiet=True,
            item_index=2,
            item_limit=1,
        )
        self.assertTrue(found_sns)
        self.assertTrue(found_item)
        self.assertEqual(item_count, 1)
        self.assertEqual(backend.item_limit, 2)

    def test_comment_uses_initial_window_rect_for_refresh_retry(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_comment_config = em_core.CONFIG_COMMENT
            old_refresh_config = em_core.CONFIG_REFRESH
            old_commands_refresh_config = em_commands.CONFIG_REFRESH
            old_window_backend = em_commands.WindowBackend
            old_input_backend = em_commands.InputBackend
            clicks: list[em.Point] = []
            window_backends: list[object] = []

            class FakeWindowBackend:
                def __init__(self) -> None:
                    self.rect_calls = 0
                    window_backends.append(self)

                def moments_window(self) -> object:
                    return object()

                def activate(self, win: object) -> None:
                    pass

                def rect(self, win: object) -> em.Rect | None:
                    self.rect_calls += 1
                    if self.rect_calls == 1:
                        return em.Rect(100, 200, 500, 1000)
                    return None

            class FakeInputBackend:
                def click(self, point: em.Point) -> None:
                    clicks.append(point)

            try:
                root = Path(tmp)
                em_core.CONFIG_COMMENT = root / ".wechat_comment_config"
                em_core.CONFIG_REFRESH = root / ".wechat_refresh_point"
                em_commands.CONFIG_REFRESH = em_core.CONFIG_REFRESH
                em.save_comment_config(
                    em.CommentConfig(
                        comment_from_action=em.Point(-66, -3),
                        send_x_ratio=0.859375,
                        send_from_action=em.Point(-33, 99),
                    )
                )
                em.save_point(em_core.CONFIG_REFRESH, em.Point(184, 186))
                em_commands.WindowBackend = FakeWindowBackend
                em_commands.InputBackend = FakeInputBackend

                with redirect_stdout(io.StringIO()):
                    with self.assertRaises(em.EasyMoneyError):
                        em.cmd_comment(["--text", "1", "--user", "fn", "--rounds", "2"])

                self.assertEqual([(int(point.x), int(point.y)) for point in clicks], [(284, 386)])
                self.assertEqual(window_backends[0].rect_calls, 1)
            finally:
                em_core.CONFIG_COMMENT = old_comment_config
                em_core.CONFIG_REFRESH = old_refresh_config
                em_commands.CONFIG_REFRESH = old_commands_refresh_config
                em_commands.WindowBackend = old_window_backend
                em_commands.InputBackend = old_input_backend

    def test_parse_comment_options_collects_comment_modes(self):
        options = em.parse_comment_options([
            "--text",
            "好看",
            "--user",
            "fn",
            "--submit-mode",
            "keys",
            "--rounds",
            "2",
            "--timing-detail",
        ])

        self.assertEqual(options.comment_text, "好看")
        self.assertEqual(options.requested_user, "fn")
        self.assertEqual(options.submit_mode, "keys")
        self.assertEqual(options.rounds, 2)
        self.assertTrue(options.timing_detail)

    def test_parse_comment_options_collects_vision_save_path(self):
        options = em.parse_comment_options([
            "--LLM",
            "--vision",
            "--save-vision-image",
            "--vision-output",
            "vision.png",
            "--user",
            "fn",
        ])

        self.assertTrue(options.use_llm)
        self.assertTrue(options.use_vision)
        self.assertTrue(options.save_vision_image)
        self.assertEqual(options.vision_save_path, Path("vision.png"))

    def test_parse_comment_options_rejects_vision_save_without_vision(self):
        with self.assertRaises(em.EasyMoneyError):
            em.parse_comment_options(["--LLM", "--save-vision-image", "--user", "fn"])

    def test_extract_inline_image_count_supports_chinese_and_digits(self):
        self.assertEqual(em.extract_inline_image_count("正文\n包含3张图片"), 3)
        self.assertEqual(em.extract_inline_image_count("这条包含9张图片"), 9)
        self.assertIsNone(em.extract_inline_image_count("这条包含两张图片"))
        self.assertIsNone(em.extract_inline_image_count("包含12张图片"))
        self.assertIsNone(em.extract_inline_image_count("这条含两张图片"))
        self.assertIsNone(em.extract_inline_image_count("没有图片提示"))

    def test_direct_uia_inline_image_rects_match_swift_grid(self):
        body = em.Rect(100, 200, 500, 600)
        window = em.Rect(0, 0, 1000, 1000)

        rects = em.direct_uia_inline_image_rects(body, 5, window)

        self.assertEqual(len(rects), 5)
        self.assertEqual(rects[0], em.Rect(176, 324, 296, 444))
        self.assertEqual(rects[2], em.Rect(424, 324, 544, 444))
        self.assertEqual(rects[3], em.Rect(176, 448, 296, 568))

    def test_single_uia_inline_image_region_uses_keyboard_focus_and_loaded_check(self):
        old_window_backend = em_llm.WindowBackend
        old_input_backend = em_llm.InputBackend
        old_capture_backend = em_llm.CaptureBackend
        old_loaded_check = em_llm.inline_image_looks_loaded
        old_env = {
            key: os.environ.get(key)
            for key in (
                "EASYMONEY_SINGLE_IMAGE_ACTIVATE_WAIT_MS",
                "EASYMONEY_SINGLE_IMAGE_FOCUS_WAIT_MS",
                "EASYMONEY_SINGLE_IMAGE_KEY_GAP_MS",
            )
        }
        events: list[tuple[str, object]] = []
        focused = object()
        cropped = object()
        test_case = self

        class FakeAutomation:
            def GetFocusedElement(self) -> object:
                events.append(("focused", None))
                return focused

        class FakeWindowBackend:
            def moments_window(self) -> object:
                events.append(("moments_window", None))
                return object()

            def activate(self, win: object) -> None:
                events.append(("activate", None))

            def _ensure_automation(self):
                return FakeAutomation(), object()

            def rect(self, element: object) -> em.Rect:
                test_case.assertIs(element, focused)
                return em.Rect(176, 324, 296, 444)

            def _safe_text(self, element: object) -> str:
                return "图片"

            def _control_type(self, element: object) -> str:
                return "按钮"

            def _class_name(self, element: object) -> str:
                return "mmui::XMouseEventView"

        class FakeInputBackend:
            def prepare_key_sequence(self, keys) -> None:
                events.append(("prepare", tuple(keys)))

            def press_sequence(self, keys, gap: float = 0.0) -> None:
                events.append(("press", (tuple(keys), gap)))

        class FakeWindowImage:
            width = 1000
            height = 1000

            def crop(self, box):
                events.append(("crop", box))
                return cropped

        class FakeCaptureBackend:
            def screenshot_stream(self, rect: em.Rect) -> FakeWindowImage:
                events.append(("screenshot_stream", rect))
                return FakeWindowImage()

            def close(self) -> None:
                events.append(("close", None))

        try:
            os.environ["EASYMONEY_SINGLE_IMAGE_ACTIVATE_WAIT_MS"] = "0"
            os.environ["EASYMONEY_SINGLE_IMAGE_FOCUS_WAIT_MS"] = "0"
            os.environ["EASYMONEY_SINGLE_IMAGE_KEY_GAP_MS"] = "0"
            em_llm.WindowBackend = FakeWindowBackend
            em_llm.InputBackend = FakeInputBackend
            em_llm.CaptureBackend = FakeCaptureBackend
            em_llm.inline_image_looks_loaded = lambda image: image is cropped

            post = em.MomentPostResolution(
                body_frame=em.Rect(100, 200, 500, 600),
                action_point=em.Point(300, 560),
                text="Doudo 包含1张图片",
                source="test",
                inline_image_count=1,
            )
            image = em_llm.capture_single_uia_inline_image_region(post, em.Rect(0, 0, 1000, 1000))

            self.assertIs(image, cropped)
            self.assertIn(("prepare", ("down", "tab", "tab", "tab")), events)
            self.assertIn(("press", (("down", "tab", "tab", "tab"), 0.0)), events)
            self.assertIn(("crop", (176, 324, 296, 444)), events)
        finally:
            em_llm.WindowBackend = old_window_backend
            em_llm.InputBackend = old_input_backend
            em_llm.CaptureBackend = old_capture_backend
            em_llm.inline_image_looks_loaded = old_loaded_check
            for key, value in old_env.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

    def test_single_uia_inline_image_region_tabs_once_more_when_focus_name_is_not_image(self):
        old_window_backend = em_llm.WindowBackend
        old_input_backend = em_llm.InputBackend
        old_capture_backend = em_llm.CaptureBackend
        old_loaded_check = em_llm.inline_image_looks_loaded
        old_env = {
            key: os.environ.get(key)
            for key in (
                "EASYMONEY_SINGLE_IMAGE_FOCUS_WAIT_MS",
                "EASYMONEY_SINGLE_IMAGE_KEY_GAP_MS",
            )
        }
        events: list[tuple[str, object]] = []
        first_focus = object()
        image_focus = object()
        cropped = object()
        test_case = self

        class FakeAutomation:
            def __init__(self) -> None:
                self.calls = 0

            def GetFocusedElement(self) -> object:
                self.calls += 1
                focused = first_focus if self.calls == 1 else image_focus
                events.append(("focused", focused))
                return focused

        automation = FakeAutomation()

        class FakeWindowBackend:
            def _ensure_automation(self):
                return automation, object()

            def rect(self, element: object) -> em.Rect:
                test_case.assertIs(element, image_focus)
                return em.Rect(180, 330, 300, 450)

            def _safe_text(self, element: object) -> str:
                return "评论" if element is first_focus else "图片"

            def _control_type(self, element: object) -> str:
                return "按钮"

            def _class_name(self, element: object) -> str:
                return "mmui::XMouseEventView" if element is image_focus else ""

        class FakeInputBackend:
            def prepare_key_sequence(self, keys) -> None:
                events.append(("prepare", tuple(keys)))

            def press_sequence(self, keys, gap: float = 0.0) -> None:
                events.append(("press", (tuple(keys), gap)))

        class FakeWindowImage:
            width = 1000
            height = 1000

            def crop(self, box):
                events.append(("crop", box))
                return cropped

        class FakeCaptureBackend:
            def screenshot_stream(self, rect: em.Rect) -> FakeWindowImage:
                return FakeWindowImage()

            def close(self) -> None:
                pass

        try:
            os.environ["EASYMONEY_SINGLE_IMAGE_FOCUS_WAIT_MS"] = "0"
            os.environ["EASYMONEY_SINGLE_IMAGE_KEY_GAP_MS"] = "0"
            em_llm.WindowBackend = FakeWindowBackend
            em_llm.InputBackend = FakeInputBackend
            em_llm.CaptureBackend = FakeCaptureBackend
            em_llm.inline_image_looks_loaded = lambda image: image is cropped

            post = em.MomentPostResolution(
                body_frame=em.Rect(100, 200, 500, 600),
                action_point=em.Point(300, 560),
                text="Doudo 包含1张图片",
                source="test",
                inline_image_count=1,
            )
            image = em_llm.capture_single_uia_inline_image_region(post, em.Rect(0, 0, 1000, 1000))

            self.assertIs(image, cropped)
            self.assertIn(("press", (("down", "tab", "tab", "tab"), 0.0)), events)
            self.assertIn(("press", (("tab",), 0.0)), events)
            self.assertIn(("crop", (180, 330, 300, 450)), events)
        finally:
            em_llm.WindowBackend = old_window_backend
            em_llm.InputBackend = old_input_backend
            em_llm.CaptureBackend = old_capture_backend
            em_llm.inline_image_looks_loaded = old_loaded_check
            for key, value in old_env.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

    def test_parse_comment_options_rejects_removed_local_kb_flags(self):
        with self.assertRaises(em.EasyMoneyError):
            em.parse_comment_options(["--text", "好看", "--user", "fn", "--noLocal"])

    def test_parse_comment_options_rejects_removed_comment_flags(self):
        for args in (
            ["--text", "好看", "--user", "fn", "--open-click"],
            ["--text", "好看", "--user", "fn", "--open-key-sequence", "tab,enter"],
            ["--text", "好看", "--user", "fn", "--submit-enter"],
            ["--text", "好看", "--user", "fn", "--submit-keys", "enter"],
            ["--text", "好看", "--user", "fn", "--fast"],
            ["--text", "好看", "--user", "fn", "--fast-send"],
            ["--text", "好看", "--user", "fn", "--slove-question"],
            ["--text", "好看", "--user", "fn", "--yolo-debug"],
            ["--text", "好看", "--user", "fn", "--save-yolo-images"],
        ):
            with self.subTest(args=args):
                with self.assertRaises(em.EasyMoneyError):
                    em.parse_comment_options(args)

    def test_comment_prefers_direct_input_when_backend_supports_text(self):
        class FakeInputBackend:
            def can_type_text_directly(self, text: str) -> bool:
                return True

        input_backend = FakeInputBackend()
        self.assertTrue(em_commands.should_type_comment_text_directly("1", input_backend))
        self.assertTrue(em_commands.should_type_comment_text_directly("6.7暗夜降至整车锁晓星", input_backend))

    def test_comment_sends_with_default_click_after_typing(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_comment_config = em_core.CONFIG_COMMENT
            old_window_backend = em_commands.WindowBackend
            old_input_backend = em_commands.InputBackend
            old_refresh_point = em_commands.refresh_point_from_saved_offset
            old_resolve_second = em_commands.resolve_second_uia_list_item_post
            old_submit_mode = os.environ.pop("EASYMONEY_SUBMIT_MODE", None)
            events: list[tuple[str, object]] = []
            window_backends: list[object] = []
            include_text_values: list[bool] = []

            class FakeWindowBackend:
                def __init__(self) -> None:
                    self.moments_window_calls = 0
                    window_backends.append(self)

                def moments_window(self) -> object:
                    self.moments_window_calls += 1
                    return object()

                def activate(self, win: object) -> None:
                    events.append(("activate", None))

                def rect(self, win: object) -> em.Rect:
                    return em.Rect(0, 0, 400, 800)

            class FakeInputBackend:
                def click(self, point: em.Point, clicks: int = 1, interval: float = 0.04) -> None:
                    events.append(("click", (int(point.x), int(point.y), clicks)))

                def press_sequence(self, keys, gap: float = 0.0) -> None:
                    events.append(("press_sequence", tuple(keys)))

                def prepare_key_sequence(self, keys) -> None:
                    events.append(("prepare_key_sequence", tuple(keys)))

                def press_sequence_atomic(self, keys) -> None:
                    events.append(("press_sequence_atomic", tuple(keys)))

                def can_type_directly(self, text: str) -> bool:
                    return True

                def type_text_directly(self, text: str) -> str:
                    events.append(("type", text))
                    return "直接键盘输入"

            try:
                root = Path(tmp)
                em_core.CONFIG_COMMENT = root / ".wechat_comment_config"
                em.save_comment_config(
                    em.CommentConfig(
                        comment_from_action=em.Point(-66, -3),
                        send_x_ratio=0.859375,
                        send_from_action=em.Point(-33, 99),
                    )
                )
                em_commands.WindowBackend = FakeWindowBackend
                em_commands.InputBackend = FakeInputBackend
                em_commands.refresh_point_from_saved_offset = lambda backend: em.Point(153, 43)

                def fake_resolve_second(*args, **kwargs):
                    include_text = kwargs.get("include_text", True)
                    include_text_values.append(include_text)
                    return em.UIAListItemResolution(
                        item_index=1,
                        body_frame=em.Rect(80, 100, 340, 170),
                        action_point=em.Point(300, 160),
                        text="fn post" if include_text else "",
                        expected_user_id="fn",
                        detected_prefix="fn",
                        elapsed_ms=3,
                    )

                em_commands.resolve_second_uia_list_item_post = fake_resolve_second

                with redirect_stdout(io.StringIO()):
                    em.cmd_comment(["--text", "1", "--user", "fn", "--rounds", "1"])

                self.assertEqual(events[-6:], [
                    ("activate", None),
                    ("prepare_key_sequence", ("tab", "enter")),
                    ("click", (300, 160, 1)),
                    ("press_sequence_atomic", ("tab", "enter")),
                    ("type", "1"),
                    ("click", (267, 259, 1)),
                ])
                self.assertEqual([event for event in events if event[0] == "click"], [
                    ("click", (300, 160, 1)),
                    ("click", (267, 259, 1)),
                ])
                self.assertEqual(window_backends[0].moments_window_calls, 1)
                self.assertEqual(include_text_values, [False])
            finally:
                em_core.CONFIG_COMMENT = old_comment_config
                em_commands.WindowBackend = old_window_backend
                em_commands.InputBackend = old_input_backend
                em_commands.refresh_point_from_saved_offset = old_refresh_point
                em_commands.resolve_second_uia_list_item_post = old_resolve_second
                if old_submit_mode is None:
                    os.environ.pop("EASYMONEY_SUBMIT_MODE", None)
                else:
                    os.environ["EASYMONEY_SUBMIT_MODE"] = old_submit_mode

    def test_capture_backend_falls_back_to_mss_for_dxgi_invalid_region(self):
        backend = object.__new__(em.CaptureBackend)
        backend.backend = "dxgi"
        backend._allow_mss_fallback = True
        backend._dx_stream_region = None
        backend.mss_mod = object()

        class FakeDxCamera:
            is_capturing = False

            def start(self, *args, **kwargs) -> None:
                raise ValueError("Invalid Region: Region should be in 1920x1080")

            def stop(self) -> None:
                pass

        class FakeShot:
            width = 4
            height = 3
            rgb = bytes([10, 20, 30] * 12)

        class FakeSct:
            def grab(self, region: dict[str, int]) -> FakeShot:
                self.region = region
                return FakeShot()

        backend._dx_camera = FakeDxCamera()
        backend._sct = FakeSct()

        frame = backend.grab_stream(em.Rect(-20, 10, 20, 13))

        self.assertEqual((frame.width, frame.height), (4, 3))
        self.assertEqual(backend._sct.region["left"], -20)

    def test_capture_backend_uses_numpy_dxgi_processor(self):
        old_import_module = em_capture.importlib.import_module
        calls: list[dict[str, object]] = []

        class FakeDxCamera:
            is_capturing = False

            def stop(self) -> None:
                pass

        class FakeDxcam:
            @staticmethod
            def create(**kwargs) -> FakeDxCamera:
                calls.append(kwargs)
                return FakeDxCamera()

        def fake_import_module(name: str):
            if name == "dxcam":
                return FakeDxcam
            return old_import_module(name)

        try:
            em_capture.importlib.import_module = fake_import_module
            em_capture.CaptureBackend("dxgi")
        finally:
            em_capture.importlib.import_module = old_import_module

        self.assertEqual(calls[0]["processor_backend"], "numpy")

if __name__ == "__main__":
    unittest.main()
