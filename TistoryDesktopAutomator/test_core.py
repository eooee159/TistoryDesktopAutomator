"""Core logic tests (will run on Windows after installing requirements)."""
import tempfile
import unittest
from pathlib import Path
from tistory_desktop import body_match, blog_post_address, compact, content_key, save_result, read_articles


class CoreTests(unittest.TestCase):
    def test_body_short(self):
        self.assertTrue(body_match("가나다\n 라마", "메뉴 가나다 라마 광고")[0])
        self.assertFalse(body_match("가나다라마바사", "가나다라")[0])

    def test_body_long(self):
        body = "A" * 95 + "B" * 95 + "C" * 95 + "D" * 95
        self.assertTrue(body_match(body, "제목" + body)[0])
        self.assertFalse(body_match(body, "B" * 95)[0])

    def test_url(self):
        self.assertTrue(blog_post_address("https://a.tistory.com/321", "https://a.tistory.com"))
        self.assertFalse(blog_post_address("https://a.tistory.com/manage/posts", "https://a.tistory.com"))
        self.assertFalse(blog_post_address("https://evil.tistory.com/321", "https://a.tistory.com"))

    def test_key(self):
        one = content_key("https://a.tistory.com/", "a", "b")
        two = content_key("https://a.tistory.com/", "a", "c")
        self.assertNotEqual(one, two)

    def test_no_sample_publish(self):
        path = Path(__file__).parent / "샘플_게시글.xlsx"
        self.assertEqual(read_articles(str(path), "https://a.tistory.com/"), [])

    def test_report_write(self):
        with tempfile.TemporaryDirectory() as tmp:
            save_result(Path(tmp), {"엑셀행": 2, "상태": "VERIFIED"})
            self.assertIn("VERIFIED", (Path(tmp) / "report.csv").read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
