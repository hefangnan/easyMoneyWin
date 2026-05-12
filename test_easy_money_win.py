import io
import os
import tempfile
import unittest
from contextlib import closing, redirect_stdout
from pathlib import Path

import easy_money_win as em


class EasyMoneyWinTests(unittest.TestCase):
    def test_point_and_comment_config_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_path = em.CONFIG_COMMENT
            try:
                em.CONFIG_COMMENT = Path(tmp) / ".wechat_comment_config"
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
                em.CONFIG_COMMENT = old_path

    def test_dotenv_loading(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = os.getcwd()
            old_cache = em._DOTENV_CACHE
            try:
                os.chdir(tmp)
                Path(".easyMoney.env").write_text(
                    "export EASYMONEY_LLM_PROVIDER=doubao\nARK_MODEL='demo-model'\n",
                    encoding="utf-8",
                )
                em._DOTENV_CACHE = None
                self.assertEqual(em.first_non_empty_env(["EASYMONEY_LLM_PROVIDER"]), "doubao")
                self.assertEqual(em.first_non_empty_env(["ARK_MODEL"]), "demo-model")
            finally:
                os.chdir(old_cwd)
                em._DOTENV_CACHE = old_cache

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

    def test_sqlite_schema_and_kb_ask(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_db = em.CONFIG_KB
            try:
                em.CONFIG_KB = Path(tmp) / ".wechat_kb.sqlite"
                with closing(em.open_knowledge_db(create=True)) as conn:
                    conn.execute(
                        "INSERT INTO scripts(title, normalized_title) VALUES (?, ?)",
                        ("测试剧本", em.normalize_text("测试剧本")),
                    )
                    script_id = conn.execute("SELECT id FROM scripts WHERE title='测试剧本'").fetchone()["id"]
                    conn.execute(
                        "INSERT INTO qa_pairs(script_id, normalized_question, question, answer, evidence) VALUES (?, ?, ?, ?, ?)",
                        (
                            script_id,
                            em.normalize_text("测试剧本的凶手是谁"),
                            "测试剧本的凶手是谁",
                            "张三",
                            "测试证据",
                        ),
                    )
                    conn.commit()
                solved = em.solve_question_from_context("测试剧本的凶手是谁")
                self.assertIsNotNone(solved)
                self.assertEqual(solved.answer, "张三")
            finally:
                em.CONFIG_KB = old_db

    def test_window_backend_uses_pywinauto_desktop(self):
        old_require_module = em.require_module
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
            em.require_module = fake_require_module
            backend = em.WindowBackend()

            self.assertIsNotNone(backend.desktop)
            self.assertIn(("pywinauto", None), calls)
            self.assertIn(("Desktop", "uia"), calls)
            self.assertNotIn(("uiautomation", None), calls)
        finally:
            em.require_module = old_require_module

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

    def test_comment_aborts_when_window_rect_unavailable_before_refresh(self):
        old_window_backend = em.WindowBackend
        old_input_backend = em.InputBackend
        old_refresh_point = em.refresh_point_from_saved_offset
        clicks: list[em.Point] = []

        class FakeWindowBackend:
            def __init__(self) -> None:
                self.rect_calls = 0

            def moments_window(self) -> object:
                return object()

            def activate(self, win: object) -> None:
                pass

            def rect(self, win: object) -> em.Rect | None:
                self.rect_calls += 1
                if self.rect_calls == 1:
                    return em.Rect(0, 0, 400, 800)
                return None

        class FakeInputBackend:
            def click(self, point: em.Point) -> None:
                clicks.append(point)

        try:
            em.WindowBackend = FakeWindowBackend
            em.InputBackend = FakeInputBackend
            em.refresh_point_from_saved_offset = lambda backend: em.Point(184, 186)

            with redirect_stdout(io.StringIO()):
                with self.assertRaises(em.WindowPositionUnavailable):
                    em.cmd_comment(["--text", "1", "--user", "fn", "--rounds", "2"])

            self.assertEqual(clicks, [])
        finally:
            em.WindowBackend = old_window_backend
            em.InputBackend = old_input_backend
            em.refresh_point_from_saved_offset = old_refresh_point

    def test_comment_sends_with_tab_shortcut_after_typing(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_comment_config = em.CONFIG_COMMENT
            old_window_backend = em.WindowBackend
            old_input_backend = em.InputBackend
            old_refresh_point = em.refresh_point_from_saved_offset
            old_resolve_second = em.resolve_second_uia_list_item_post
            events: list[tuple[str, object]] = []
            window_backends: list[object] = []

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
                em.CONFIG_COMMENT = root / ".wechat_comment_config"
                em.save_comment_config(
                    em.CommentConfig(
                        comment_from_action=em.Point(-66, -3),
                        send_x_ratio=0.859375,
                        send_from_action=em.Point(-33, 99),
                    )
                )
                em.WindowBackend = FakeWindowBackend
                em.InputBackend = FakeInputBackend
                em.refresh_point_from_saved_offset = lambda backend: em.Point(153, 43)
                em.resolve_second_uia_list_item_post = lambda *args, **kwargs: em.UIAListItemResolution(
                    item_index=1,
                    body_frame=em.Rect(80, 100, 340, 170),
                    action_point=em.Point(300, 160),
                    text="fn post",
                    expected_user_id="fn",
                    detected_prefix="fn",
                    elapsed_ms=3,
                )

                with redirect_stdout(io.StringIO()):
                    em.cmd_comment(["--text", "1", "--user", "fn", "--rounds", "1"])

                self.assertEqual(events[-6:], [
                    ("prepare_key_sequence", ("tab", "enter")),
                    ("prepare_key_sequence", ("tab", "tab", "tab", "enter")),
                    ("click", (300, 160, 1)),
                    ("press_sequence_atomic", ("tab", "enter")),
                    ("type", "1"),
                    ("press_sequence_atomic", ("tab", "tab", "tab", "enter")),
                ])
                self.assertEqual([event for event in events if event[0] == "click"], [("click", (300, 160, 1))])
                self.assertEqual(window_backends[0].moments_window_calls, 1)
            finally:
                em.CONFIG_COMMENT = old_comment_config
                em.WindowBackend = old_window_backend
                em.InputBackend = old_input_backend
                em.refresh_point_from_saved_offset = old_refresh_point
                em.resolve_second_uia_list_item_post = old_resolve_second

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

if __name__ == "__main__":
    unittest.main()
