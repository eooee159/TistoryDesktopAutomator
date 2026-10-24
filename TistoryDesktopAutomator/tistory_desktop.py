# -*- coding: utf-8 -*-
"""Tistory Desktop Automator — GUI-only (mouse/keyboard, no blog API)."""
from __future__ import annotations

import csv
import hashlib
import json
import queue
import re
import sys
import threading
import time
import traceback
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import openpyxl
    import pyautogui as pg
    import pyperclip
    import pygetwindow as gw
except ImportError as exc:
    raise SystemExit(
        f"필수 패키지가 없습니다: {exc}\n설치.bat을 실행하거나 pip install -r requirements.txt 를 실행하세요."
    )

BASE = Path(sys.argv[0]).resolve().parent
CONFIG_PATH = BASE / "settings.json"
SCREEN_DIR = BASE / "results"
CONFIG_KEYS = ["title", "body", "complete", "public", "publish", "latest_view"]
POINT_NAMES = {
    "title": "제목 입력칸",
    "body": "본문 입력칸",
    "complete": "완료 버튼 (1차)",
    "public": "공개 선택 항목 (선택)",
    "publish": "발행 버튼 (최종)",
    "latest_view": "글 관리: 최신 글 '보기' 링크",
}
REPORT_COLUMNS = ["처리시각", "엑셀행", "제목", "상태", "게시글URL", "제목검증", "본문검증", "캡처폴더", "비고"]


def compact(text: str) -> str:
    return re.sub(r"\s+", "", str(text or ""))


def clean_url(url: str) -> str:
    value = (url or "").strip()
    parts = urlsplit(value)
    if parts.scheme not in ("http", "https") or not parts.netloc:
        raise ValueError("URL은 http:// 또는 https://로 시작하는 유효한 주소여야 합니다.")
    if parts.username or parts.password:
        raise ValueError("아이디/비밀번호가 포함된 URL은 사용할 수 없습니다.")
    return value


def slug_title(title: str) -> str:
    result = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", title).strip(" .")
    return (result or "post")[:40]


def content_key(home_url: str, title: str, body: str) -> str:
    return hashlib.sha256((urlsplit(home_url).netloc.lower()+"\n"+title+"\n"+body).encode("utf-8")).hexdigest()


def body_match(body: str, page_text: str) -> tuple[bool, str]:
    source, rendered = compact(body), compact(page_text)
    if not source:
        return False, "원본 본문 없음"
    # Short articles: full match; long articles: three independent samples.
    if len(source) <= 220:
        return source in rendered, "본문 전체" if source in rendered else "본문 전체 불일치"
    width = min(120, max(65, len(source)//12))
    starts = [0, max(0, (len(source)-width)//2), max(0, len(source)-width)]
    checks = [source[p:p+width] in rendered for p in starts]
    return all(checks), f"본문 시작/중간/끝 {sum(checks)}/3 일치"


def blog_post_address(url: str, home_url: str) -> bool:
    try:
        current = urlsplit(clean_url(url))
        home = urlsplit(clean_url(home_url))
        if current.netloc.lower() != home.netloc.lower():
            return False
        path = current.path.strip("/")
        if not path or path.startswith("manage/") or path.startswith("admin/"):
            return False
        return True
    except ValueError:
        return False


@dataclass
class Article:
    excel_row: int
    title: str
    body: str
    key: str


def read_articles(path: str, home_url: str) -> list[Article]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        try:
            headers = [str(x).strip() if x is not None else "" for x in next(rows)]
        except StopIteration:
            raise ValueError("엑셀 파일에 데이터가 없습니다.")
        header_map = {key: index for index, key in enumerate(headers)}
        required = ("제목", "본문", "실행")
        missing = [x for x in required if x not in header_map]
        if missing:
            raise ValueError("엑셀 1행에서 다음 열을 찾지 못했습니다: " + ", ".join(missing))
        result = []
        for row_no, values in enumerate(rows, start=2):
            def get(name):
                i = header_map[name]
                return values[i] if i < len(values) else None
            if str(get("실행") or "").strip().upper() not in ("Y", "YES", "1", "TRUE", "발행"):
                continue
            title = str(get("제목") or "").strip()
            body = str(get("본문") or "").strip()
            if not title or not body:
                raise ValueError(f"엑셀 {row_no}행: 제목/본문 중 비어 있는 값이 있습니다.")
            if len(title) > 300:
                raise ValueError(f"엑셀 {row_no}행: 제목이 300자를 초과합니다.")
            result.append(Article(row_no, title, body, content_key(home_url, title, body)))
        return result
    finally:
        wb.close()


def read_json(path: Path, fallback: dict) -> dict:
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        pass
    return fallback


def atomic_json(path: Path, value: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def save_result(root: Path, record: dict):
    root.mkdir(parents=True, exist_ok=True)
    target = root / "report.csv"
    new_file = not target.exists()
    with target.open("a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=REPORT_COLUMNS)
        if new_file:
            writer.writeheader()
        writer.writerow({k: record.get(k, "") for k in REPORT_COLUMNS})


class Stopped(Exception):
    pass


class ScreenDriver:
    """No browser APIs/DOM/HTTP requests: only physical desktop controls."""
    def __init__(self, stop_event: threading.Event, logger, step_delay: float):
        self.stop_event = stop_event
        self.log = logger
        self.step_delay = step_delay
        pg.FAILSAFE = True  # Move mouse to a screen corner to stop.
        pg.PAUSE = 0.16

    def wait(self, seconds: float):
        until = time.monotonic() + max(seconds, 0)
        while time.monotonic() < until:
            if self.stop_event.is_set():
                raise Stopped("사용자가 중지했습니다.")
            time.sleep(min(0.15, max(0, until - time.monotonic())))

    def click(self, loc: list[int], label: str):
        if self.stop_event.is_set():
            raise Stopped("사용자가 중지했습니다.")
        self.log(f"클릭: {label} ({loc[0]}, {loc[1]})")
        pg.click(*loc)
        self.wait(self.step_delay)

    def paste(self, value: str):
        if self.stop_event.is_set():
            raise Stopped("사용자가 중지했습니다.")
        pyperclip.copy(value)
        self.wait(.18)
        pg.hotkey("ctrl", "v")
        self.wait(self.step_delay)

    def address(self) -> str:
        marker = "WAIT_URL_" + uuid.uuid4().hex
        pyperclip.copy(marker)
        pg.hotkey("ctrl", "l")
        self.wait(.2)
        pg.hotkey("ctrl", "c")
        self.wait(.35)
        address = pyperclip.paste().strip()
        pg.press("esc")
        return "" if address == marker else address

    def open_url(self, url: str):
        clean_url(url)
        pg.hotkey("ctrl", "l")
        self.paste(url)
        pg.press("enter")
        self.wait(max(2, self.step_delay * 3))

    def page_text(self, click_center: bool = True) -> str:
        marker = "WAIT_PAGE_" + uuid.uuid4().hex
        pyperclip.copy(marker)
        if click_center:
            pg.click(int(pg.size().width * .50), int(pg.size().height * .55))
        pg.hotkey("ctrl", "a")
        self.wait(.3)
        pg.hotkey("ctrl", "c")
        self.wait(.45)
        value = pyperclip.paste()
        pg.press("right")  # Clear selection to avoid accidental replacement later.
        return "" if value == marker else value

    def screenshot(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        pg.screenshot().save(path)

    def screenshots_of_post(self, folder: Path, max_count: int):
        pg.hotkey("ctrl", "home")
        self.wait(.4)
        for i in range(max_count):
            self.screenshot(folder / f"published_{i+1:02d}.png")
            if i < max_count - 1:
                pg.press("pagedown")
                self.wait(.55)

    def focus_chrome(self):
        try:
            active = gw.getActiveWindow()
            found = [w for w in gw.getAllWindows() if "Chrome" in w.title and w.width > 500]
            if found:
                window = active if active in found else (next((w for w in found if "티스토리" in w.title or "Tistory" in w.title), found[0]))
                if window.isMinimized:
                    window.restore()
                window.activate()
                self.wait(.55)
                return
        except Exception as exc:
            self.log("크롬 자동 활성화 실패: " + str(exc))
        self.log("크롬을 수동으로 활성화했는지 확인하세요.")


class Application:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("티스토리 엑셀 자동 글쓰기 · 화면 조작")
        self.root.geometry("900x760")
        self.root.minsize(790, 680)
        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.running = False
        self.config = read_json(CONFIG_PATH, {})
        self.vars = {}
        defaults = {
            "xlsx": "", "home": "https://내블로그.tistory.com/",
            "writer": "", "manage": "", "delay": "2.5", "max_screens": "3",
        }
        for key, default in defaults.items():
            self.vars[key] = tk.StringVar(value=str(self.config.get(key, default)))
        self.test_mode = tk.BooleanVar(value=True)
        self.use_public = tk.BooleanVar(value=bool(self.config.get("use_public", False)))
        self.points = self.config.get("points", {})
        self.res = self.config.get("screen_size")
        self.make_ui()
        self.root.after(120, self.drain_queue)

    def make_ui(self):
        parent = ttk.Frame(self.root, padding=14)
        parent.pack(fill="both", expand=True)
        ttk.Label(parent, text="티스토리 데스크톱 글쓰기 자동화", font=("맑은 고딕", 17, "bold")).pack(anchor="w", pady=(0, 5))
        ttk.Label(parent, text="API 없이 실제 크롬 화면의 마우스·키보드를 제어합니다. 로그인은 직접 하세요.", foreground="#666").pack(anchor="w", pady=(0, 10))
        tabs = ttk.Notebook(parent)
        tabs.pack(fill="both", expand=True)
        run_frame, cal_frame, help_frame = ttk.Frame(tabs, padding=13), ttk.Frame(tabs, padding=13), ttk.Frame(tabs, padding=13)
        tabs.add(run_frame, text="실행")
        tabs.add(cal_frame, text="클릭 위치 설정")
        tabs.add(help_frame, text="사용 설명")
        self.make_run_ui(run_frame)
        self.make_cal_ui(cal_frame)
        self.make_help_ui(help_frame)

    def row_entry(self, master, label, var, browse=False):
        row = ttk.Frame(master)
        row.pack(fill="x", pady=6)
        ttk.Label(row, text=label, width=18).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=(3, 5))
        if browse:
            ttk.Button(row, text="찾아보기", command=self.select_xlsx).pack(side="left")

    def make_run_ui(self, pane):
        self.row_entry(pane, "엑셀 파일(.xlsx)", self.vars["xlsx"], browse=True)
        self.row_entry(pane, "블로그 홈 URL", self.vars["home"])
        self.row_entry(pane, "글쓰기 URL", self.vars["writer"])
        self.row_entry(pane, "글 관리 URL", self.vars["manage"])
        ttk.Label(pane, text="※ URL은 로그인 후 크롬 주소창에서 직접 복사하세요. 블로그 홈/글쓰기/글 관리 세 곳이 필요합니다.", foreground="#666").pack(anchor="w", pady=(1, 9))

        options = ttk.Frame(pane)
        options.pack(fill="x", pady=5)
        ttk.Label(options, text="단계 지연(초)").pack(side="left")
        ttk.Entry(options, textvariable=self.vars["delay"], width=7).pack(side="left", padx=(6, 18))
        ttk.Label(options, text="게시글 캡처(장)").pack(side="left")
        ttk.Combobox(options, textvariable=self.vars["max_screens"], width=5, values=["1", "2", "3", "4", "5", "6"], state="readonly").pack(side="left", padx=6)
        ttk.Checkbutton(pane, text="테스트 모드: 첫 Y행만 입력하고 발행하지 않음 (최초 권장)", variable=self.test_mode).pack(anchor="w", pady=4)
        ttk.Checkbutton(pane, text="발행 옵션에서 공개 항목 자동 선택 (공개 위치 별도 지정 필요)", variable=self.use_public).pack(anchor="w", pady=4)

        btn = ttk.Frame(pane)
        btn.pack(fill="x", pady=12)
        ttk.Button(btn, text="엑셀 검사", command=self.preview).pack(side="left", padx=(0, 9))
        ttk.Button(btn, text="설정 저장", command=self.persist).pack(side="left", padx=(0, 9))
        self.start_btn = ttk.Button(btn, text="▶ 자동 작성 시작", command=self.start)
        self.start_btn.pack(side="left", padx=(0, 9))
        ttk.Button(btn, text="■ 중지 요청", command=self.stop).pack(side="left", padx=(0, 9))
        ttk.Button(btn, text="결과 폴더 열기", command=self.open_output).pack(side="left")
        ttk.Label(pane, text="실행 로그", font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(10, 5))
        log_frame = ttk.Frame(pane)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=12, state="disabled", wrap="word", font=("맑은 고딕", 9))
        scroller = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroller.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroller.pack(side="right", fill="y")
        self.status = tk.StringVar(value="준비 중 | 우선 클릭 위치 설정 후 테스트 모드를 실행하세요.")
        ttk.Label(pane, textvariable=self.status, foreground="#035b9a").pack(anchor="w", pady=(7, 0))

    def make_cal_ui(self, pane):
        ttk.Label(pane, text="먼저 크롬을 최대화하고, 배율 100%로 설정하세요.", font=("맑은 고딕", 11, "bold")).pack(anchor="w", pady=(0, 8))
        msg = (
            "위치 지정 버튼을 누르면 이 창이 5초간 숨겨집니다. 크롬으로 이동해 해당 입력칸/버튼 위에 "
            "마우스 포인터를 올려 두세요. 5초 후 그 좌표를 저장합니다. 클릭은 하지 않습니다.\n"
            "'최종 발행' 버튼은 티스토리 글쓰기에서 '완료'를 눌러 열린 발행 옵션 화면에서 지정하세요.\n"
            "'최신 글 보기'는 티스토리 글 관리 화면에서 첫 번째 글의 '보기' 링크로 지정하세요."
        )
        ttk.Label(pane, text=msg, wraplength=780, justify="left").pack(anchor="w", pady=(0, 12))
        self.point_labels = {}
        for key in CONFIG_KEYS:
            row = ttk.Frame(pane)
            row.pack(fill="x", pady=7)
            ttk.Label(row, text=POINT_NAMES[key], width=32).pack(side="left")
            self.point_labels[key] = tk.StringVar(value=self.describe_point(key))
            ttk.Label(row, textvariable=self.point_labels[key], width=20).pack(side="left")
            ttk.Button(row, text="5초 후 현재 마우스 위치 저장", command=lambda k=key: self.capture(k)).pack(side="left", padx=8)
        ttk.Separator(pane).pack(fill="x", pady=15)
        ttk.Label(pane, text="화면 해상도/Windows 배율/크롬 줌/브라우저 위치가 바뀌면 클릭 위치를 다시 지정하세요.", foreground="#a33").pack(anchor="w")
        ttk.Button(pane, text="설정 저장", command=self.persist).pack(anchor="w", pady=15)

    def make_help_ui(self, pane):
        notes = (
            "[1] 샘플_게시글.xlsx 파일의 제목, 본문을 입력하고 실행 열을 Y로 변경합니다.\n\n"
            "[2] 크롬을 열어 티스토리에 직접 로그인합니다. 크롬을 최대화하고 확대 비율을 100%로 고정합니다.\n\n"
            "[3] 크롬 글쓰기 화면/글 관리 화면의 URL을 각각 복사해 실행 탭에 넣습니다.\n\n"
            "[4] 클릭 위치 설정 탭에서 각 항목에 마우스를 올려 지정합니다. '완료' 후 나타나는 발행창 위치도 지정합니다.\n\n"
            "[5] '테스트 모드'를 유지한 채 첫 Y행만 작성해 미리보기 캡처를 확인합니다.\n\n"
            "[6] 정상 입력이 확인되면 '테스트 모드'를 해제하고 실행합니다. 실행 중 마우스/키보드를 사용하지 마세요.\n\n"
            "[7] 발행 완료 후 글 관리 목록의 맨 위 '보기' 링크를 열어 실제 페이지에서 제목과 본문을 검증합니다.\n\n"
            "[8] results/<블로그>/<실행일자>/report.csv 및 각 게시글 PNG를 확인하세요.\n\n"
            "[중지] 중지 버튼 또는 화면 왼쪽 위 모서리로 마우스를 이동하면 중단할 수 있습니다.\n\n"
            "[주의] 발행 시도 후 검증에 실패한 건은 재발행하지 않습니다. 티스토리 글 관리에서 확인한 뒤 수동 처리하세요.\n"
            "공개 여부, 카테고리, 태그 등은 현재 버전에서 세부 자동설정하지 않습니다.\n"
            "비밀번호, 2단계 인증, 보안문자는 자동으로 입력하거나 우회하지 않습니다.\n"
            "본문의 서식/이미지는 단순 텍스트 붙여넣기 버전이며 HTML 렌더링은 지원하지 않습니다."
        )
        ttk.Label(pane, text=notes, justify="left", wraplength=790).pack(anchor="w")

    def describe_point(self, key):
        pos = self.points.get(key)
        return f"X={pos[0]} / Y={pos[1]}" if pos else "미지정"

    def select_xlsx(self):
        path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel Workbook", "*.xlsx")])
        if path:
            self.vars["xlsx"].set(path)

    def capture(self, key):
        if self.running:
            messagebox.showwarning("실행 중", "자동화 실행이 끝난 뒤 위치를 설정하세요.")
            return
        self.root.withdraw()
        def worker():
            time.sleep(5)
            pos = pg.position()
            size = pg.size()
            self.root.after(0, lambda: finish(pos, size))
        def finish(pos, size):
            self.points[key] = [int(pos.x), int(pos.y)]
            self.res = [int(size.width), int(size.height)]
            self.point_labels[key].set(self.describe_point(key))
            self.root.deiconify()
            self.root.lift()
            self.log(f"좌표 지정 완료: {POINT_NAMES[key]} {self.points[key]}")
        threading.Thread(target=worker, daemon=True).start()

    def settings(self):
        cfg = {k: var.get().strip() for k, var in self.vars.items()}
        cfg.update(points=self.points, screen_size=self.res, use_public=self.use_public.get())
        return cfg

    def persist(self):
        atomic_json(CONFIG_PATH, self.settings())
        self.log("설정을 settings.json 파일에 저장했습니다.")

    def preview(self):
        try:
            home = clean_url(self.vars["home"].get())
            articles = read_articles(self.vars["xlsx"].get(), home)
            self.log(f"엑셀 검사 완료: 발행 대상 Y {len(articles)}건")
            for a in articles[:8]:
                self.log(f"  {a.excel_row}행 | {a.title[:65]}")
            if not articles:
                messagebox.showinfo("발행 대상 없음", "엑셀 실행 열을 Y로 지정한 행이 없습니다.")
        except Exception as exc:
            messagebox.showerror("엑셀 검사 실패", str(exc))

    def validate(self):
        cfg = self.settings()
        if not cfg["xlsx"] or not Path(cfg["xlsx"]).is_file():
            raise ValueError("존재하는 엑셀 파일을 지정하세요.")
        for field in ("home", "writer", "manage"):
            clean_url(cfg[field])
        # Do not accept obvious mismatched blog hosts; writer or manage may use tistory.com central domain.
        if float(cfg["delay"]) < 1 or float(cfg["delay"]) > 30:
            raise ValueError("단계 지연은 1~30초로 입력하세요.")
        if not self.test_mode.get() and not cfg["points"].get("latest_view"):
            raise ValueError("발행 모드에는 글 관리 최신 글 '보기' 좌표가 필요합니다.")
        required = ["title", "body"] if self.test_mode.get() else ["title", "body", "complete", "publish", "latest_view"]
        if cfg["use_public"] and not self.test_mode.get():
            required.append("public")
        for key in required:
            if key not in cfg["points"]:
                raise ValueError(f"클릭 위치 미지정: {POINT_NAMES[key]}")
        if cfg["screen_size"] != list(pg.size()):
            raise ValueError("좌표 지정 당시 화면 해상도와 현재 해상도가 다릅니다. 위치를 다시 지정하세요.")
        articles = read_articles(cfg["xlsx"], cfg["home"])
        if not articles:
            raise ValueError("엑셀 '실행' 열이 Y인 게시글이 없습니다.")
        return cfg, articles

    def start(self):
        if self.running:
            return
        try:
            cfg, articles = self.validate()
        except Exception as exc:
            messagebox.showerror("실행 전 확인", str(exc))
            return
        if not self.test_mode.get():
            msg = (f"티스토리에 총 {len(articles)}건을 실제 발행 시도합니다.\n\n"
                   "최초에는 테스트 모드에서 화면 입력 상태를 검증하세요.\n"
                   "발행 후 검증 실패 시 자동 재시도하지 않습니다.\n\n계속하시겠습니까?")
            if not messagebox.askyesno("실제 발행 확인", msg):
                return
        self.persist()
        self.running = True
        self.start_btn.configure(state="disabled")
        self.stop_event.clear()
        self.status.set("진행 중: 크롬을 건드리지 마세요. (긴급 중단: 좌측 상단 모서리)")
        threading.Thread(target=self.worker, args=(cfg, articles, self.test_mode.get()), daemon=True).start()

    def stop(self):
        self.stop_event.set()
        self.log("중지 요청을 접수했습니다. 현재 조작 후 중단합니다.")

    def open_output(self):
        SCREEN_DIR.mkdir(parents=True, exist_ok=True)
        import os
        os.startfile(str(SCREEN_DIR))

    def log(self, line):
        self.events.put(("log", line))

    def drain_queue(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", f"[{datetime.now():%H:%M:%S}] {payload}\n")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                elif kind == "done":
                    self.running = False
                    self.start_btn.configure(state="normal")
                    self.status.set(payload)
                    self.root.deiconify()
                    self.root.lift()
        except queue.Empty:
            pass
        self.root.after(120, self.drain_queue)

    def worker(self, cfg, articles, test_mode):
        home = cfg["home"]
        slug = re.sub(r"[^a-zA-Z0-9가-힣_.-]", "_", urlsplit(home).netloc.lower())
        current_dir = SCREEN_DIR / slug / datetime.now().strftime("%Y%m%d")
        current_dir.mkdir(parents=True, exist_ok=True)
        ledger_path = SCREEN_DIR / slug / "ledger.json"
        ledger = read_json(ledger_path, {})
        driver = ScreenDriver(self.stop_event, self.log, float(cfg["delay"]))
        self.log("크롬에서 로그인된 티스토리 탭을 준비하세요. 3초 후 시작합니다.")
        self.root.after(0, self.root.withdraw)
        errors = 0
        try:
            driver.wait(3)
            driver.focus_chrome()
            count = 1 if test_mode else len(articles)
            for index, article in enumerate(articles[:count], 1):
                if self.stop_event.is_set():
                    raise Stopped("중지됨")
                old = ledger.get(article.key, {})
                if not test_mode and old.get("status") not in (None, "FAILED_BEFORE_PUBLISH"):
                    self.log(f"건너뜀: {article.excel_row}행 (이전 발행 시도 기록: {old.get('status')})")
                    continue
                name = f"row{article.excel_row:04d}_{article.key[:10]}_{slug_title(article.title)}"
                folder = current_dir / name
                folder.mkdir(parents=True, exist_ok=True)
                url = ""
                title_ok = False
                body_ok = False
                published_attempt = False
                status = "FAILED_BEFORE_PUBLISH"
                note = ""
                self.log(f"[{index}/{count}] {article.excel_row}행: {article.title}")
                try:
                    driver.open_url(cfg["writer"])
                    # Refuse if URL appears to be login instead of writer.
                    writer_now = driver.address()
                    if "login" in urlsplit(writer_now).path.lower():
                        raise RuntimeError("로그인 화면입니다. 직접 로그인한 후 다시 실행하세요.")
                    if urlsplit(writer_now).netloc.lower() != urlsplit(cfg["writer"]).netloc.lower():
                        raise RuntimeError("글쓰기 페이지와 다른 사이트로 이동했습니다.")
                    if urlsplit(writer_now).path.rstrip("/") != urlsplit(cfg["writer"]).path.rstrip("/"):
                        raise RuntimeError("예상한 글쓰기 페이지 주소가 아닙니다. 수동 확인하세요.")
                    driver.click(cfg["points"]["title"], "제목")
                    driver.paste(article.title)
                    driver.click(cfg["points"]["body"], "본문")
                    driver.paste(article.body)
                    driver.wait(.7)
                    driver.screenshot(folder / "draft.png")
                    if not test_mode:
                        # Clipboard-based editor verification before irreversible publishing.
                        driver.click(cfg["points"]["title"], "제목 입력 확인")
                        typed_title = driver.page_text(click_center=False)
                        driver.click(cfg["points"]["body"], "본문 입력 확인")
                        typed_body = driver.page_text(click_center=False)
                        if compact(article.title) not in compact(typed_title):
                            raise RuntimeError("발행 전 제목 검증 실패 (클릭 위치/붙여넣기 확인 필요).")
                        pre_body_ok, pre_body_why = body_match(article.body, typed_body)
                        if not pre_body_ok:
                            raise RuntimeError("발행 전 본문 검증 실패: " + pre_body_why)
                        self.log("발행 전 확인 완료: 제목·본문 원문 일치")
                    # Do not publish in dry-run mode.
                    if test_mode:
                        status = "TEST_ONLY_NOT_PUBLISHED"
                        note = "테스트: 입력 화면만 캡처함. 실제 발행은 하지 않음."
                        self.log("테스트 입력 완료 / 발행하지 않았습니다.")
                    else:
                        driver.click(cfg["points"]["complete"], "완료 (1차)")
                        driver.wait(.8)
                        driver.screenshot(folder / "publish_dialog.png")
                        if cfg["use_public"]:
                            driver.click(cfg["points"]["public"], "공개 설정")
                        # Persist BEFORE the irreversible click so crash cannot lead to duplicate auto-post.
                        status = "PUBLISH_ATTEMPTED_REVIEW_IF_INTERRUPTED"
                        ledger[article.key] = {"status": status, "row": article.excel_row,
                                               "title": article.title, "when": datetime.now().isoformat(),
                                               "evidence": str(folder)}
                        atomic_json(ledger_path, ledger)
                        published_attempt = True
                        driver.click(cfg["points"]["publish"], "발행 (최종)")
                        driver.wait(3)
                        driver.screenshot(folder / "post_click.png")
                        # Independently open management, then first article view.
                        driver.open_url(cfg["manage"])
                        driver.wait(1)
                        driver.screenshot(folder / "manage_list.png")
                        listing = driver.page_text()
                        if compact(article.title) not in compact(listing):
                            raise RuntimeError("글 관리 목록에서 이번 제목을 찾지 못했습니다. 발행 여부를 수동 확인하세요.")
                        driver.click(cfg["points"]["latest_view"], "최신 글 보기")
                        driver.wait(2)
                        url = driver.address()
                        # If '보기' opened a new background tab, switch to it once.
                        if not blog_post_address(url, home):
                            pg.hotkey("ctrl", "tab")
                            driver.wait(1.5)
                            url = driver.address()
                        if not blog_post_address(url, home):
                            raise RuntimeError("글 관리의 '보기' 링크에서 실제 게시글 주소로 이동하지 않았습니다.")
                        # Always save evidence even if text extraction fails.
                        driver.screenshots_of_post(folder, int(cfg["max_screens"]))
                        pg.hotkey("ctrl", "home")
                        driver.wait(.3)
                        actual_page = driver.page_text()
                        title_ok = compact(article.title) in compact(actual_page)
                        body_ok, explanation = body_match(article.body, actual_page)
                        note = explanation
                        if not (title_ok and body_ok):
                            raise RuntimeError("게시글 URL은 확인했으나 제목 또는 본문이 원본과 다릅니다: " + explanation)
                        status = "VERIFIED"
                        self.log(f"검증 성공: {url}")
                except (Exception, pg.FailSafeException) as exc:
                    errors += 1
                    note = f"{type(exc).__name__}: {exc}"
                    if published_attempt:
                        status = "REVIEW_REQUIRED_NO_RETRY"
                    try:
                        driver.screenshot(folder / "error_screen.png")
                    except Exception:
                        pass
                    self.log(f"오류: {note}")
                finally:
                    if not test_mode:
                        ledger[article.key] = {
                            "status": status, "row": article.excel_row,
                            "title": article.title, "when": datetime.now().isoformat(),
                            "post_url": url, "evidence": str(folder), "reason": note
                        }
                        atomic_json(ledger_path, ledger)
                    record = {
                        "처리시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "엑셀행": article.excel_row, "제목": article.title,
                        "상태": status, "게시글URL": url, "제목검증": str(title_ok),
                        "본문검증": str(body_ok), "캡처폴더": str(folder), "비고": note,
                    }
                    save_result(current_dir, record)
                if status != "VERIFIED" and not test_mode:
                    self.log("안전상 다음 글로 진행하지 않습니다. 티스토리 글 관리에서 확인 후 수동 조치하세요.")
                    break
        except Stopped as exc:
            self.log(str(exc))
        except Exception as exc:
            errors += 1
            self.log("전체 작업 중단: " + str(exc))
            self.log(traceback.format_exc(limit=4))
        finally:
            self.events.put(("done", f"완료/중단 · 오류 {errors}건 · 결과 저장: {current_dir}"))


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    Application(root)
    root.mainloop()


if __name__ == "__main__":
    main()
