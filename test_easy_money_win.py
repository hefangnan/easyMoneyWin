import os
import tempfile
import unittest
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
                with em.open_knowledge_db(create=True) as conn:
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


if __name__ == "__main__":
    unittest.main()
