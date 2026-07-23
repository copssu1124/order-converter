# -*- coding: utf-8 -*-
"""주문서 변환기 GUI v4.0 — 토스형 미니멀(A) + 단계·큰 버튼(F) 디자인.
   ★ 변환 엔진·자동업데이트·발주서 분리 로직은 v3.5 그대로. 화면(UI)만 교체."""
import os
import re
import sys
import json
import queue
import ctypes
import threading
import traceback
import tempfile
import subprocess
import webbrowser
import urllib.request
import tkinter as tk
from tkinter import filedialog
import tkinter.font as tkfont

import step3_convert as engine

VERSION = "4.1"                 # ★ 버전은 이 한 곳에서만 관리
KAKAO = "https://open.kakao.com/o/gyxhX4zi"
CREDIT = "Developed by JANG JUNG WOO · JJ COMPANY"
GITHUB_REPO = "copssu1124/order-converter"
RELEASES_URL = "https://github.com/copssu1124/order-converter/releases/latest"

# ── 팔레트 (토스 미니멀 + 초록 브랜드) — 이름은 기존 유지, 값만 재조정 ──
BG      = "#ffffff"     # 페이지 바탕(흰색)
CARD    = "#ffffff"     # 카드/모달 바탕
INK     = "#191f28"     # 본문 잉크
INK2    = "#191f28"
MUTED   = "#8b95a1"     # 흐린 글씨
LINE    = "#eef1f4"     # 얇은 구분선
FBG     = "#f5f7f9"     # 인셋(파일 행) 배경
AC      = "#12a45f"     # 초록 브랜드
AC2     = "#0e8a4f"
AC_DK   = "#0c7442"
PILLBG  = "#eef1f4"
OKC     = "#12a45f"
OKBG    = "#e7f6ee"
WARNC   = "#b06f14"
WARNBG  = "#fdf3e3"
ERRC    = "#c0392b"
DISBG   = "#dfe4e9"
DISFG   = "#9aa2ab"
BADGE_BG = "#d6efe1"    # 단계 미완료 배지 — 옅은 초록(죽은 회색 대신)
BADGE_FG = "#2ba46b"
BTN_DIS  = "#a9dcc2"    # 비활성 초록 버튼 — 옅은 초록
BTN_DISF = "#ffffff"

_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def _번호(n):
    return _CIRCLED[n - 1] if 1 <= n <= len(_CIRCLED) else "(%d)" % n


def _버전튜플(s):
    try:
        s = str(s).strip().lstrip("vV")
        parts = re.findall(r"\d+", s)
        return tuple(int(x) for x in parts[:2]) if parts else None
    except Exception:
        return None


def _검증_다운로드(path, expected_size):
    """받은 업데이트 파일이 정상 exe인지 검증 → (ok, 사유)."""
    try:
        size = os.path.getsize(path)
    except OSError as e:
        return False, "파일 확인 실패: %s" % e
    if expected_size and size != expected_size:
        return False, "다운로드가 중간에 끊겼어요(%d/%d바이트)" % (size, expected_size)
    if size < 1048576:
        return False, "받은 파일이 너무 작습니다"
    try:
        with open(path, "rb") as f:
            if f.read(2) != b"MZ":
                return False, "받은 파일이 정상 실행파일이 아니에요"
    except OSError as e:
        return False, "파일 읽기 실패: %s" % e
    return True, ""


def app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def ui_dir():
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "ui")


def load_fonts():
    """Pretendard OTF를 개인폰트로 로드(설치 불필요). 실패 시 맑은 고딕 폴백."""
    try:
        FR_PRIVATE = 0x10
        for f in ("Pretendard-Regular.otf", "Pretendard-SemiBold.otf", "Pretendard-Bold.otf"):
            p = os.path.join(ui_dir(), f)
            if os.path.exists(p):
                ctypes.windll.gdi32.AddFontResourceExW(p, FR_PRIVATE, 0)
    except Exception:
        pass


def _round_rect(cv, x1, y1, x2, y2, r, **kw):
    r = max(0, min(r, (x2 - x1) // 2, (y2 - y1) // 2))
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


class RoundBtn(tk.Canvas):
    """둥근 버튼 (Canvas). primary=초록, ghost=흰 알약, danger=빨강."""
    def __init__(self, parent, text, command, kind="ghost", height=46,
                 fontobj=None, radius=11, **kw):
        bgp = kw.pop("bgparent", CARD)
        super().__init__(parent, height=height, bg=bgp, highlightthickness=0, **kw)
        self.text = text
        self.command = command
        self.kind = kind
        self.height = height
        self.radius = radius
        self.font = fontobj
        self._enabled = True
        self._hover = False
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda e: self._set_hover(True))
        self.bind("<Leave>", lambda e: self._set_hover(False))
        self.config(cursor="hand2")

    def _colors(self):
        if not self._enabled:
            if self.kind == "primary":
                return BTN_DIS, BTN_DISF          # 비활성도 옅은 초록으로(회색 X)
            return DISBG, DISFG
        if self.kind == "primary":
            return (AC2 if self._hover else AC), "#ffffff"
        if self.kind == "danger":
            return ("#b53d2e" if self._hover else "#c0392b"), "#ffffff"
        return (FBG if self._hover else CARD), "#33413a"

    def _draw(self):
        self.delete("all")
        w = self.winfo_width(); h = self.height
        if w < 4:
            return
        fill, fg = self._colors()
        _round_rect(self, 1, 1, w - 1, h - 1, self.radius, fill=fill,
                    outline=("" if self.kind in ("primary", "danger") else "#d3dae1"),
                    width=1)
        self.create_text(w // 2, h // 2, text=self.text, fill=fg,
                         font=self.font or ("맑은 고딕", 11, "bold"))

    def _set_hover(self, v):
        self._hover = v
        self._draw()

    def _click(self, _e):
        if self._enabled and self.command:
            self.command()

    def set_text(self, t):
        self.text = t
        self._draw()

    def set_enabled(self, v):
        self._enabled = bool(v)
        self.config(cursor="hand2" if v else "arrow")
        self._draw()


GUIDE_SECTIONS = [
    ("① 주문서 변환",
     "주문서(.xls)와 매핑표를 고르고 [변환 실행]을 누르면 결과 파일이 '변환결과' 폴더에 "
     "새로 만들어집니다. 원본·매핑표는 절대 수정하지 않아요. 매핑표는 같은 폴더에 있으면 자동으로 잡힙니다.\n"
     "※ ① 탭에는 '원본 주문서'만 넣어주세요. 이미 변환된 결과 파일을 넣으면 안내하고 멈춥니다."),
    ("② 택배사 분리",
     "①에서 만든 결과를 엑셀로 검증·보완한 뒤 불러오고, 발화주명(보내는 회사)을 적고 실행하면 "
     "택배사별 발주서가 '분리출력' 폴더에 만들어집니다.\n"
     "양식: 씨제이(R열 운임)·대신/대신낱개(연두=낱개)·대신택배·천일(K열 운임)·로젠(G열)·"
     "원준(L열)·위플(F열)·올담·카몬드 / 그 외는 '기타' 파일"),
    ("📑 주문서 규칙",
     "· 수령자·상품명이 빈 행에서 멈춥니다\n· 옵션(선택사항) 우선, 없으면 상품명 사용\n"
     "· 송장·배송번호 등 빈칸은 그대로 보존"),
    ("🗂 매핑표 규칙",
     "· 상품명변경: 옵션명 → 회사상품명\n· 판매비변경: 회사상품명*수량 → 배송비(등급*택배사 / 순수금액 / 천일박스 / …수정)\n"
     "· 택배비 시트: 천일·씨제이·로젠·대신택배·위플·원준·카몬드·올담·용차 + 수정택배비\n"
     "· 빨간 행: 매핑표에 키가 없으면 → 매핑표에 추가하면 해결 / 키가 있는데도 빨강 → 프로그램 문의"),
]
COLOR_LEGEND = [
    ("#DDEBF7", "파랑", "일반 택배사 (천일·씨제이 등)"),
    ("#FFC000", "주황", "다모아"),
    ("#FF0000", "빨강", "수동확인 필요 (비고·리포트에 사유)"),
    ("#FFFF00", "노랑", "배송비 빈칸 (확인 필요)"),
    ("#FFCCCC", "연빨강", "배송비합계 / 진빨강=단가×수량 곱셈"),
    ("#E6CCF5", "연보라", "수정택배비 사용 (택배사 '수정')"),
    ("#92D050", "연두", "대신 발주서의 대신낱개 행"),
    ("#FFC7CE", "분홍", "주소 중복 (같은 주소 2건 이상)"),
]


class NumBadge(tk.Canvas):
    """단계 번호 배지. set_done(True)면 초록, 아니면 회색."""
    def __init__(self, parent, num, fonts, bg=BG):
        super().__init__(parent, width=30, height=30, bg=bg, highlightthickness=0, bd=0)
        self.num = num
        self.fonts = fonts
        self._done = False
        self._draw()

    def set_done(self, done=True):
        self._done = done
        self._draw()

    def _draw(self):
        self.delete("all")
        fill, fg = (AC, "#ffffff") if self._done else (BADGE_BG, BADGE_FG)
        self.create_oval(2, 2, 28, 28, fill=fill, outline="")
        self.create_text(15, 16, text=str(self.num), fill=fg, font=self.fonts["badge"])


class FileRow(tk.Canvas):
    """① 번호 배지 + 카테고리/파일명 + 오른쪽 버튼, 둥근 회색 인셋. 행 전체 클릭 가능."""
    H = 78

    def __init__(self, parent, num, category, fonts, command):
        super().__init__(parent, height=self.H, bg=BG, highlightthickness=0, bd=0)
        self.num = num
        self.category = category
        self.fonts = fonts
        self.command = command
        self.filename = ""
        self.placeholder = "파일을 선택하세요"
        self.sub = ""
        self.sub_ok = False
        self._done = False
        self.btn = RoundBtn(self, "선택", command, kind="ghost", height=38,
                            fontobj=fonts["btn"], radius=11, bgparent=FBG, width=74)
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>", lambda e: self.command())

    def set_file(self, name, sub="", ok=False):
        self.filename = name
        self.sub = sub
        self.sub_ok = ok
        self.btn.set_text("변경")
        self._draw()

    def set_done(self, done=True):
        self._done = done
        self._draw()

    @staticmethod
    def _fit(text, font, maxw):
        if maxw <= 0 or font.measure(text) <= maxw:
            return text
        while text and font.measure(text + "…") > maxw:
            text = text[:-1]
        return text + "…"

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.H
        if w < 4:
            return
        _round_rect(self, 1, 4, w - 1, h - 4, 16, fill=FBG, outline=LINE, width=1)
        cy = h // 2
        r = 15
        bx = 26
        fill, fg = (AC, "#ffffff") if self._done else (BADGE_BG, BADGE_FG)
        self.create_oval(bx - r, cy - r, bx + r, cy + r, fill=fill, outline="")
        self.create_text(bx, cy + 1, text=str(self.num), fill=fg, font=self.fonts["badge"])
        tx = 60
        # 윗줄: 카테고리 + (안내문구)
        self.create_text(tx, cy - 13, text=self.category, anchor="w", fill=MUTED, font=self.fonts["cat"])
        if self.sub:
            sx = tx + self.fonts["cat"].measure(self.category) + 10
            self.create_text(sx, cy - 13, text=self.sub, anchor="w",
                             fill=(OKC if self.sub_ok else MUTED), font=self.fonts["note"])
        # 아랫줄: 파일명 (버튼과 겹치지 않게 말줄임)
        name = self.filename or self.placeholder
        is_ph = not self.filename
        avail = (w - 12 - 74 - 12) - tx      # 버튼(폭74)·여백 제외한 가용 폭
        name = self._fit(name, self.fonts["file"], avail)
        self.create_text(tx, cy + 8, text=name, anchor="w",
                         fill=(MUTED if is_ph else INK), font=self.fonts["file"])
        self.create_window(w - 12, cy, anchor="e", window=self.btn)


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.input_file = None
        self.mapping_file = None
        self.output_file = None
        self.conv_file = None
        self.split_dir = None
        self.log_queue = queue.Queue()
        self._busy = False

        # 폰트
        fams = tkfont.families()
        fam = "Pretendard" if "Pretendard" in fams else "맑은 고딕"
        semi = "Pretendard SemiBold" if "Pretendard SemiBold" in fams else fam
        self.fam = fam
        self.F = lambda s, b=False: (fam, s, "bold") if b else (fam, s)

        def _fnt(size, w="normal", family=None):
            return tkfont.Font(family=family or fam, size=size, weight=w)
        if semi != fam:
            self.rowfonts = {"badge": _fnt(11, "bold"), "cat": _fnt(9, family=semi),
                             "file": _fnt(13, family=semi), "note": _fnt(9, family=semi),
                             "btn": _fnt(10, family=semi)}
        else:
            self.rowfonts = {"badge": _fnt(11, "bold"), "cat": _fnt(9),
                             "file": _fnt(12, "bold"), "note": _fnt(9, "bold"),
                             "btn": _fnt(9, "bold")}

        root.title("주문서 변환기 · JJ COMPANY   v" + VERSION)
        W, H = 560, 880
        try:
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            root.geometry("%dx%d+%d+%d" % (W, H, max(0, (sw - W) // 2), max(0, (sh - H) // 2 - 30)))
        except Exception:
            root.geometry("%dx%d" % (W, H))
        root.minsize(520, 720)
        root.configure(bg=BG)
        root.rowconfigure(2, weight=1)
        root.columnconfigure(0, weight=1)

        self._build_header()
        self._build_tabs()
        self._build_body()
        self._build_footer()

        auto = engine.find_mapping_file(app_dir())
        if auto:
            self.mapping_file = auto
            self._set_file("map", os.path.basename(auto), "자동으로 찾았어요 ✓", ok=True)
        self.log("주문서를 불러오고 [변환 실행]을 누르세요.")
        self._refresh_conv()
        self._drain_log()
        threading.Thread(target=self._check_update, daemon=True).start()

    # ═══════════ 헤더 ═══════════
    def _build_header(self):
        h = tk.Frame(self.root, bg=BG)
        h.grid(row=0, column=0, sticky="ew")
        h.columnconfigure(1, weight=1)

        left = tk.Frame(h, bg=BG)
        left.grid(row=0, column=0, sticky="w", padx=(26, 0), pady=(20, 0))
        tk.Label(left, text="주문서 변환기", bg=BG, fg=INK, font=self.F(15, True)).pack(side="left")

        right = tk.Frame(h, bg=BG)
        right.grid(row=0, column=2, sticky="e", padx=(0, 24), pady=(20, 0))
        self.verlbl = tk.Label(right, text="v%s · 확인 중…" % VERSION, bg=BG, fg=MUTED, font=self.F(9))
        self.verlbl.pack(side="right")
        q = tk.Label(right, text="문의", bg=BG, fg=MUTED, font=self.F(9), cursor="hand2")
        q.pack(side="right", padx=(0, 16))
        q.bind("<Button-1>", lambda e: webbrowser.open(KAKAO))
        q.bind("<Enter>", lambda e: q.config(fg=INK))
        q.bind("<Leave>", lambda e: q.config(fg=MUTED))
        hp = tk.Label(right, text="설명서", bg=BG, fg=AC, font=self.F(9, True), cursor="hand2")
        hp.pack(side="right", padx=(0, 14))
        hp.bind("<Button-1>", lambda e: self.open_help())
        hp.bind("<Enter>", lambda e: hp.config(fg=AC2))
        hp.bind("<Leave>", lambda e: hp.config(fg=AC))

    # ═══════════ 탭 (토스 언더라인) ═══════════
    def _build_tabs(self):
        bar = tk.Frame(self.root, bg=BG)
        bar.grid(row=1, column=0, sticky="ew")
        tk.Frame(bar, bg=LINE, height=1).pack(fill="x", side="bottom")
        inner = tk.Frame(bar, bg=BG)
        inner.pack(anchor="w", padx=26, pady=(15, 0))
        self.tab_btns = {}
        for key, txt in (("conv", "주문서 변환"), ("split", "택배사 분리")):
            f = tk.Frame(inner, bg=BG)
            f.pack(side="left", padx=(0, 26))
            lb = tk.Label(f, text=txt, bg=BG, fg=MUTED, font=self.F(12), cursor="hand2")
            lb.pack()
            ul = tk.Frame(f, bg=BG, height=3)
            ul.pack(fill="x", pady=(8, 0))
            lb.bind("<Button-1>", lambda e, k=key: self.show_tab(k))
            self.tab_btns[key] = (lb, ul)

    # ═══════════ 본문 ═══════════
    def _build_body(self):
        wrap = tk.Frame(self.root, bg=BG)
        wrap.grid(row=2, column=0, sticky="nsew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        cv = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        cv.grid(row=0, column=0, sticky="nsew")
        sb = tk.Scrollbar(wrap, orient="vertical", command=cv.yview)
        sb.grid(row=0, column=1, sticky="ns")
        cv.configure(yscrollcommand=sb.set)
        inner = tk.Frame(cv, bg=BG)
        cv.create_window((0, 0), window=inner, anchor="nw", tags="inner")

        def _resize(e):
            cv.itemconfig("inner", width=e.width)
            cv.configure(scrollregion=cv.bbox("all"))
        cv.bind("<Configure>", _resize)
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-e.delta / 120), "units"))

        col = tk.Frame(inner, bg=BG)
        col.pack(fill="x", padx=26)
        self._rows = {}
        self.pane_conv = tk.Frame(col, bg=BG)
        self.pane_split = tk.Frame(col, bg=BG)
        self._build_conv(self.pane_conv)
        self._build_split(self.pane_split)
        self._build_log(col)
        self.show_tab("conv")

    def _set_file(self, kind, name, sub, ok=False):
        self._rows[kind].set_file(name, sub, ok)

    def _mark_done(self, obj, done=True):
        obj.set_done(done)

    def _build_conv(self, p):
        tk.Label(p, text="두 파일만 고르면 끝이에요", bg=BG, fg=INK,
                 font=self.F(19, True)).pack(anchor="w", pady=(20, 2))
        tk.Label(p, text="주문서와 매핑표를 고르고 아래 초록 버튼을 누르세요",
                 bg=BG, fg=MUTED, font=self.F(10)).pack(anchor="w", pady=(0, 14))

        row_o = FileRow(p, 1, "주문서", self.rowfonts, self.pick_input)
        row_o.pack(fill="x", pady=(0, 9))
        self._rows["order"] = row_o
        self._c1 = (row_o,)

        row_m = FileRow(p, 2, "매핑표", self.rowfonts, self.pick_mapping)
        row_m.pack(fill="x", pady=(0, 14))
        self._rows["map"] = row_m
        self._c2 = (row_m,)

        self.btn_conv = RoundBtn(p, "③   변환 실행", self.run_convert, kind="primary",
                                 height=56, fontobj=self.F(15, True), bgparent=BG, radius=15)
        self.btn_conv.pack(fill="x")
        self.btn_conv.set_enabled(False)

        self.result_box = tk.Frame(p, bg=BG)
        self.result_box.pack(fill="x")          # 결과 영역은 항상 표시(변환 전엔 – )
        tk.Frame(self.result_box, bg=LINE, height=1).pack(fill="x", pady=(18, 0))
        strip = tk.Frame(self.result_box, bg=BG)
        strip.pack(fill="x", pady=(12, 0))
        self.tiles = {}
        for i, (key, label, colr) in enumerate(
                (("total", "총 주문", INK), ("good", "정상", OKC), ("issue", "확인 필요", WARNC))):
            c = tk.Frame(strip, bg=BG)
            c.grid(row=0, column=i, sticky="nsew")
            strip.grid_columnconfigure(i, weight=1)
            v = tk.Label(c, text="–", bg=BG, fg=colr, font=self.F(25, True))
            v.pack()
            tk.Label(c, text=label, bg=BG, fg=MUTED, font=self.F(9)).pack()
            self.tiles[key] = v
        self.issue_wrap = tk.Frame(self.result_box, bg=BG)
        self.issue_wrap.pack(fill="x", pady=(12, 0))
        self.btn_open = RoundBtn(self.result_box, "결과 폴더 열기", self.open_result_folder,
                                 kind="ghost", height=44, fontobj=self.F(10, True), bgparent=BG, radius=12)
        # ↑ '결과 폴더 열기'는 변환이 끝난 뒤에만 노출(_show_result에서 pack)

    def _build_split(self, p):
        tk.Label(p, text="택배사별 발주서 만들기", bg=BG, fg=INK,
                 font=self.F(19, True)).pack(anchor="w", pady=(20, 2))
        tk.Label(p, text="①에서 만든 결과를 엑셀로 검증·보완한 뒤 불러오세요",
                 bg=BG, fg=MUTED, font=self.F(10)).pack(anchor="w", pady=(0, 14))

        row_c = FileRow(p, 1, "변환완료 파일", self.rowfonts, self.pick_conv)
        row_c.pack(fill="x", pady=(0, 12))
        self._rows["conv"] = row_c
        self._s1 = (row_c,)

        er = tk.Frame(p, bg=FBG, highlightbackground=LINE, highlightthickness=1)
        er.pack(fill="x", pady=(0, 14))
        badge2 = NumBadge(er, 2, self.rowfonts, bg=FBG)
        badge2.pack(side="left", padx=(11, 0), pady=13)
        self._s2 = (badge2,)
        mid = tk.Frame(er, bg=FBG)
        mid.pack(side="left", fill="x", expand=True, padx=(10, 14))
        tk.Label(mid, text="발화주명  (보내는 회사)", bg=FBG, fg=MUTED,
                 font=self.rowfonts["cat"]).pack(anchor="w", pady=(11, 0))
        self.entry_발화 = tk.Entry(mid, font=self.F(12), relief="flat", bg="#ffffff",
                                   highlightthickness=1, highlightbackground="#dbe0e6",
                                   highlightcolor=AC)
        self.entry_발화.pack(fill="x", pady=(3, 11), ipady=5)
        self.entry_발화.bind("<KeyRelease>", lambda e: self._refresh_split())

        self.btn_split = RoundBtn(p, "③   택배사 분리 실행", self.run_split, kind="primary",
                                  height=56, fontobj=self.F(15, True), bgparent=BG, radius=15)
        self.btn_split.pack(fill="x")
        self.btn_split.set_enabled(False)

        self.split_result = tk.Frame(p, bg=BG)
        self.carrier_wrap = tk.Frame(self.split_result, bg=BG)
        self.carrier_wrap.pack(fill="x", pady=(14, 0))
        self.btn_split_open = RoundBtn(self.split_result, "분리 폴더 열기", self.open_split_folder,
                                       kind="ghost", height=44, fontobj=self.F(10, True),
                                       bgparent=BG, radius=12)
        self.btn_split_open.pack(fill="x", pady=(6, 4))

    def _build_log(self, parent):
        self.logcard = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        self.logcard.pack(fill="x", pady=(6, 8))
        head = tk.Frame(self.logcard, bg=CARD, cursor="hand2")
        head.pack(fill="x")
        tk.Label(head, text="진행 상황", bg=CARD, fg="#5b636d", font=self.F(10, True)).pack(side="left", padx=14, pady=10)
        self.log_arrow = tk.Label(head, text="▾", bg=CARD, fg=MUTED, font=self.F(10))
        self.log_arrow.pack(side="right", padx=14)
        self.logbody = tk.Frame(self.logcard, bg=CARD)
        from tkinter import scrolledtext
        self.logbox = scrolledtext.ScrolledText(self.logbody, font=(self.fam, 10), height=9,
                                                 state="disabled", wrap="word", relief="flat",
                                                 bg="#fbfdfc", fg="#42544a", padx=12, pady=6,
                                                 highlightthickness=0, spacing1=1, spacing3=1)
        self.logbox.pack(fill="both", expand=True, padx=6, pady=(0, 8))
        self._log_open = True
        self.logbody.pack(fill="both")
        for w in (head, self.log_arrow):
            w.bind("<Button-1>", lambda e: self._toggle_log())

    def _toggle_log(self):
        self._log_open = not self._log_open
        if self._log_open:
            self.logbody.pack(fill="both")
            self.log_arrow.config(text="▾")
        else:
            self.logbody.pack_forget()
            self.log_arrow.config(text="▸")

    def _build_footer(self):
        f = tk.Frame(self.root, bg=BG)
        f.grid(row=3, column=0, sticky="ew")
        tk.Label(f, text=CREDIT, bg=BG, fg="#9aa8a0", font=self.F(8)).pack(pady=10)

    def show_tab(self, key):
        self.pane_conv.pack_forget()
        self.pane_split.pack_forget()
        pane = self.pane_conv if key == "conv" else self.pane_split
        if getattr(self, "logcard", None):        # 로그 카드는 항상 아래에 유지
            pane.pack(before=self.logcard, fill="x")
        else:
            pane.pack(fill="x")
        for k, (lb, ul) in self.tab_btns.items():
            on = (k == key)
            lb.config(fg=(INK if on else MUTED), font=self.F(12, on))
            ul.config(bg=(AC if on else BG))

    # ═══════════ 로그 ═══════════
    def log(self, msg):
        self.log_queue.put(str(msg))

    def _drain_log(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.logbox.config(state="normal")
                self.logbox.insert("end", msg + "\n")
                self.logbox.see("end")
                self.logbox.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log)

    # ═══════════ 파일 선택 ═══════════
    def _warn_if_result(self, path):
        try:
            cols = [str(c) for c in engine.pd.read_excel(path, nrows=0).columns]
            if '배송비합계' in cols:
                self.log("⚠️ 이 파일은 이미 변환된 결과로 보여요 — ①탭엔 원본 주문서, "
                         "결과 파일은 ②택배사 분리 탭에 넣어주세요.")
        except Exception:
            pass

    def pick_input(self):
        start = os.path.dirname(self.input_file) if self.input_file else app_dir()
        path = filedialog.askopenfilename(title="주문서(.xls) 선택", initialdir=start,
                                          filetypes=[("엑셀 주문서", "*.xls *.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self.input_file = path
        self._set_file("order", os.path.basename(path), "")
        self._mark_done(*self._c1)
        self.log("주문서 선택: " + os.path.basename(path))
        self._warn_if_result(path)
        if not self.mapping_file:
            am = engine.find_mapping_file(os.path.dirname(path))
            if am:
                self.mapping_file = am
                self._set_file("map", os.path.basename(am), "자동으로 찾았어요 ✓", ok=True)
                self._mark_done(*self._c2)
                self.log("매핑표 자동 탐색: " + os.path.basename(am))
        self._refresh_conv()

    def pick_mapping(self):
        start = os.path.dirname(self.mapping_file) if self.mapping_file else app_dir()
        path = filedialog.askopenfilename(title="매핑표(.xlsx) 선택", initialdir=start,
                                          filetypes=[("엑셀 매핑표", "*.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self.mapping_file = path
        self._set_file("map", os.path.basename(path), "")
        self._mark_done(*self._c2)
        self.log("매핑표 선택: " + os.path.basename(path))
        self._refresh_conv()

    def pick_conv(self):
        start = os.path.join(app_dir(), "변환결과")
        if not os.path.isdir(start):
            start = app_dir()
        path = filedialog.askopenfilename(title="변환완료 파일(.xlsx) 선택", initialdir=start,
                                          filetypes=[("엑셀", "*.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self.conv_file = path
        self._set_file("conv", os.path.basename(path), "")
        self._mark_done(*self._s1)
        self.log("분리 대상: " + os.path.basename(path))
        self._refresh_split()

    def _refresh_conv(self):
        if self.input_file:
            self._mark_done(*self._c1)
        if self.mapping_file:
            self._mark_done(*self._c2)
        self.btn_conv.set_enabled(bool(self.input_file and self.mapping_file) and not self._busy)

    def _refresh_split(self):
        if self.conv_file:
            self._mark_done(*self._s1)
        self._mark_done(*self._s2, done=bool(self.entry_발화.get().strip()))
        self.btn_split.set_enabled(bool(self.conv_file) and not self._busy)

    # ═══════════ 변환 ═══════════
    def run_convert(self):
        if self._busy or not self.input_file or not self.mapping_file:
            return
        self._busy = True
        self.btn_conv.set_enabled(False)
        self.btn_conv.set_text("변환 중…")
        self.log("─" * 30)
        self.log("변환을 시작합니다...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            결과폴더 = os.path.join(app_dir(), "변환결과")
            os.makedirs(결과폴더, exist_ok=True)
            base = os.path.splitext(os.path.basename(self.input_file))[0]
            out_path = os.path.join(결과폴더, base + ".xlsx")
            out, 사유 = engine.convert_v2(self.input_file, self.mapping_file, out_path, log=self.log)
            self.output_file = out
            총, 정상, 이슈 = engine.상세리포트(out, 사유)
            self._리포트출력(총, 정상, 이슈)
            self.log("")
            self.log("💾 저장 완료: " + out)
            self.root.after(0, lambda: self._show_result(총, 정상, 이슈))
        except Exception as ex:
            self.log("─" * 30)
            self.log("❌ 오류: " + str(ex))
            if "이미 변환된" not in str(ex):
                self.log(traceback.format_exc())
        finally:
            self._busy = False
            self.root.after(0, lambda: (self.btn_conv.set_text("③   변환 실행"), self._refresh_conv()))

    def _show_result(self, 총, 정상, 이슈):
        self.tiles["total"].config(text=str(총))
        self.tiles["good"].config(text=str(정상))
        self.tiles["issue"].config(text=str(len(이슈)))
        for w in self.issue_wrap.winfo_children():
            w.destroy()
        for it in 이슈[:6]:
            b = tk.Frame(self.issue_wrap, bg=WARNBG)
            b.pack(fill="x", pady=(0, 6))
            상품 = str(it.get('상품명') or '').strip()[:40]
            head = "엑셀 %d행 · %s" % (it['행'], 상품)
            tk.Label(b, text=head, bg=WARNBG, fg=WARNC, font=self.F(9, True), anchor="w",
                     wraplength=460, justify="left").pack(anchor="w", padx=10, pady=(7, 0))
            tk.Label(b, text=it['문제'] + "  →  " + it['해결'], bg=WARNBG, fg="#7d6a45",
                     font=self.F(9), anchor="w", wraplength=460, justify="left").pack(anchor="w", padx=10, pady=(0, 7))
        if len(이슈) > 6:
            tk.Label(self.issue_wrap, text="…외 %d건 (아래 진행 상황 참고)" % (len(이슈) - 6),
                     bg=BG, fg=MUTED, font=self.F(9)).pack(anchor="w", pady=(0, 4))
        self.btn_open.pack(fill="x", pady=(6, 4))
        self.result_box.pack(fill="x")

    def _리포트출력(self, 총, 정상, 이슈):
        self.log("")
        self.log("════════════ 변환 완료!  총 %d건 ════════════" % 총)
        self.log("   ✅ 잘 된 것 : %d건   ⚠️ 확인 필요 : %d건" % (정상, len(이슈)))
        if not 이슈:
            self.log("✅ 모두 정상으로 변환됐어요!")
            return
        self.log("")
        for i, it in enumerate(이슈, 1):
            상품 = str(it.get('상품명') or '').strip()
            수량 = it.get('수량')
            제목 = '%s 엑셀 %d행 — "%s"' % (_번호(i), it['행'], 상품[:42])
            if 수량 not in (None, ''):
                제목 += "  (수량 %s개)" % 수량
            self.log(제목)
            self.log("     문제 : " + it['문제'])
            self.log("     해결 : " + it['해결'])

    def open_result_folder(self):
        if self.output_file and os.path.exists(self.output_file):
            try:
                os.startfile(os.path.dirname(self.output_file))
            except OSError as ex:
                self.log("폴더 열기 실패: " + str(ex))

    # ═══════════ 발주서 분리 ═══════════
    def _form_dirs(self):
        dirs = [os.path.join(app_dir(), "출력양식"),
                os.path.join(os.path.dirname(app_dir()), "출력양식")]
        base = getattr(sys, "_MEIPASS", None)
        if base:
            dirs.append(os.path.join(base, "출력양식"))
        return dirs

    def run_split(self):
        if self._busy or not self.conv_file:
            return
        발화 = self.entry_발화.get().strip()
        self._busy = True
        self.btn_split.set_enabled(False)
        self.btn_split.set_text("분리 중…")
        self.split_result.pack_forget()
        self.log("─" * 30)
        self.log("발주서 분리를 시작합니다... (발화주명: %s)" % (발화 or "(빈칸)"))
        threading.Thread(target=self._split_worker, args=(발화,), daemon=True).start()

    def _split_worker(self, 발화):
        try:
            out_dir = os.path.join(app_dir(), "분리출력")
            res = engine.발주서분리(self.conv_file, 발화, self._form_dirs(), out_dir, log=self.log)
            self.split_dir = out_dir
            self.log("─" * 30)
            for 택배사, n, p in res:
                self.log("  %s : %d건 → %s" % (택배사, n, os.path.basename(p)))
            self.log("✅ 분리 완료: " + out_dir)
            self.root.after(0, lambda: self._show_split(res))
        except engine.양식파일없음 as ex:
            self.log("─" * 30)
            self.log("⚠️ " + str(ex))
        except Exception as ex:
            self.log("─" * 30)
            self.log("❌ 분리 오류: " + str(ex))
            self.log(traceback.format_exc())
        finally:
            self._busy = False
            self.root.after(0, lambda: (self.btn_split.set_text("③   택배사 분리 실행"), self._refresh_split()))

    def _show_split(self, res):
        for w in self.carrier_wrap.winfo_children():
            w.destroy()
        for 택배사, n, p in res:
            r = tk.Frame(self.carrier_wrap, bg=FBG)
            r.pack(fill="x", pady=(0, 6))
            tk.Label(r, text="📦 %s" % 택배사, bg=FBG, fg=INK, font=self.F(10), anchor="w").pack(side="left", padx=13, pady=9)
            tk.Label(r, text="%d건" % n, bg=FBG, fg=INK2, font=self.F(10, True)).pack(side="right", padx=13)
        # 총 합계 (전체 건수)
        총건수 = sum(n for _, n, _ in res)
        tk.Frame(self.carrier_wrap, bg=LINE, height=1).pack(fill="x", pady=(6, 0))
        tot = tk.Frame(self.carrier_wrap, bg=BG)
        tot.pack(fill="x", pady=(9, 2))
        tk.Label(tot, text="총 합계", bg=BG, fg=INK, font=self.F(12, True)).pack(side="left", padx=6)
        tk.Label(tot, text="%d건" % 총건수, bg=BG, fg=AC, font=self.F(16, True)).pack(side="right", padx=6)
        self.split_result.pack(fill="x")

    def open_split_folder(self):
        if self.split_dir and os.path.exists(self.split_dir):
            try:
                os.startfile(self.split_dir)
            except OSError as ex:
                self.log("폴더 열기 실패: " + str(ex))

    # ═══════════ 설명서 모달 ═══════════
    def open_help(self):
        win = tk.Toplevel(self.root)
        win.title("설명서")
        win.configure(bg=BG)
        win.transient(self.root)
        win.geometry("480x680")
        win.minsize(430, 560)
        cv = tk.Canvas(win, bg=BG, highlightthickness=0)
        cv.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(win, orient="vertical", command=cv.yview)
        sb.pack(side="right", fill="y")
        cv.configure(yscrollcommand=sb.set)
        body = tk.Frame(cv, bg=BG)
        wid = cv.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfigure(wid, width=e.width))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-e.delta / 120), "units"))

        pad = tk.Frame(body, bg=BG)
        pad.pack(fill="both", expand=True, padx=26, pady=22)
        tk.Label(pad, text="📖 설명서", bg=BG, fg=INK2, font=self.F(16, True)).pack(anchor="w")
        tk.Label(pad, text="주문서 변환기 사용법과 결과 색상 안내", bg=BG, fg=MUTED,
                 font=self.F(10)).pack(anchor="w", pady=(2, 8))
        for title, text in GUIDE_SECTIONS:
            tk.Label(pad, text=title, bg=BG, fg=INK2, font=self.F(11, True)).pack(anchor="w", pady=(13, 3))
            tk.Label(pad, text=text, bg=BG, fg="#4a5560", font=self.F(9), justify="left",
                     wraplength=396, anchor="w").pack(anchor="w")
        tk.Frame(pad, bg=LINE, height=1).pack(fill="x", pady=(16, 12))
        tk.Label(pad, text="🎨 결과 파일 색상", bg=BG, fg=INK2, font=self.F(11, True)).pack(anchor="w", pady=(0, 6))
        for hexc, nm, desc in COLOR_LEGEND:
            r = tk.Frame(pad, bg=BG)
            r.pack(fill="x", pady=3, anchor="w")
            sw = tk.Frame(r, bg=hexc, width=22, height=16,
                          highlightthickness=1, highlightbackground="#d5d5d5")
            sw.pack(side="left", padx=(0, 10))
            sw.pack_propagate(False)
            tk.Label(r, text=nm, bg=BG, fg=INK, font=self.F(9, True), width=6, anchor="w").pack(side="left")
            tk.Label(r, text=desc, bg=BG, fg="#4a5560", font=self.F(9),
                     wraplength=300, justify="left").pack(side="left")
        win.grab_set()

    # ═══════════ 자동 업데이트 (v3.5 로직 유지) ═══════════
    def _check_update(self):
        try:
            url = "https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO
            req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json",
                                                       "User-Agent": "order-converter"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            최신 = _버전튜플(str(data.get("tag_name", "")))
            현재 = _버전튜플(VERSION)
            dl = None
            for a in (data.get("assets") or []):
                if str(a.get("name", "")).lower().endswith(".exe"):
                    dl = a.get("browser_download_url")
                    break
            if 최신 and 현재 and 최신 > 현재:
                disp = "%d.%d" % 최신
                self.root.after(0, lambda: (self._setver("v%s 있음" % disp, AC2),
                                            self._update_popup(disp, dl)))
            else:
                self.root.after(0, lambda: self._setver("최신 ✓", OKC))
        except Exception:
            self.root.after(0, lambda: self._setver("오프라인", MUTED))

    def _setver(self, text, color):
        self.verlbl.config(text="v%s · %s" % (VERSION, text), fg=color)

    def _update_popup(self, new_ver, download_url):
        try:
            win = tk.Toplevel(self.root)
            win.title("업데이트")
            win.configure(bg=CARD)
            win.transient(self.root)
            win.resizable(False, False)
            tk.Label(win, text="🎉 새 버전 v%s 이(가) 나왔어요!" % new_ver, bg=CARD, fg=INK2,
                     font=self.F(14, True)).pack(padx=40, pady=(24, 4))
            tk.Label(win, text="지금 업데이트하면 자동으로 교체 후 다시 시작됩니다.", bg=CARD, fg=MUTED,
                     font=self.F(10)).pack(pady=(0, 10))
            status = tk.Label(win, text="", bg=CARD, fg=AC2, font=self.F(9, True))
            status.pack(pady=(0, 6))
            bar = tk.Frame(win, bg=CARD)
            bar.pack(pady=(0, 20), padx=30, fill="x")
            btn_later = RoundBtn(bar, "나중에", win.destroy, kind="ghost", height=42,
                                 fontobj=self.F(10, True), bgparent=CARD, width=110)
            btn_now = RoundBtn(bar, "지금 업데이트", None, kind="primary", height=42,
                               fontobj=self.F(10, True), bgparent=CARD)

            def 지금():
                btn_now.set_enabled(False)
                btn_later.set_enabled(False)
                threading.Thread(target=self._do_update,
                                 args=(download_url, win, status, btn_later),
                                 daemon=True).start()
            btn_now.command = 지금
            btn_later.pack(side="left")
            btn_now.pack(side="left", fill="x", expand=True, padx=(10, 0))
            win.update_idletasks()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            ww, wh = win.winfo_width(), win.winfo_height()
            win.geometry("+%d+%d" % (rx + (rw - ww) // 2, ry + (rh - wh) // 3))
            win.grab_set()
        except Exception:
            pass

    def _set_status(self, status, text, color=None):
        self.root.after(0, lambda: status.config(text=text, fg=color or AC2))

    def _manual_fallback(self, win, status, msg, btn_later=None):
        def show():
            status.config(text=msg, fg=ERRC)
            link = tk.Label(win, text="👉 수동 다운로드 (Releases 페이지 열기)", bg=CARD,
                            fg="#2b6de8", cursor="hand2", font=self.F(9, True))
            link.pack(pady=(2, 14))
            link.bind("<Button-1>", lambda e: webbrowser.open(RELEASES_URL))
            if btn_later is not None:
                btn_later.set_enabled(True)
        self.root.after(0, show)

    def _do_update(self, download_url, win, status, btn_later):
        if not getattr(sys, "frozen", False):
            self._manual_fallback(win, status, "개발 모드(.py 실행)에선 자동 교체가 안 돼요. 수동으로 받아주세요.", btn_later)
            return
        if not download_url:
            self._manual_fallback(win, status, "다운로드 주소를 찾지 못했어요. 수동으로 받아주세요.", btn_later)
            return
        target = sys.executable
        try:
            _t = os.path.join(os.path.dirname(target), ".__upd_test__")
            with open(_t, "w") as f:
                f.write("x")
            os.remove(_t)
        except Exception:
            self._manual_fallback(win, status, "이 위치는 자동 교체 권한이 없어요(관리자 폴더). 수동으로 받아주세요.", btn_later)
            return
        tmp = os.path.join(tempfile.gettempdir(), "OrderConverter_update.exe")
        part = tmp + ".part"
        try:
            self._set_status(status, "다운로드 준비 중...")
            req = urllib.request.Request(download_url, headers={"User-Agent": "order-converter"})
            with urllib.request.urlopen(req, timeout=30) as resp:
                total = int(resp.headers.get("Content-Length", 0) or 0)
                done = 0
                with open(part, "wb") as f:
                    while True:
                        chunk = resp.read(262144)
                        if not chunk:
                            break
                        f.write(chunk)
                        done += len(chunk)
                        if total:
                            self._set_status(status, "다운로드 중... %d%%" % int(done * 100 / total))
                        else:
                            self._set_status(status, "다운로드 중... %d MB" % (done // 1048576))
            ok, 사유 = _검증_다운로드(part, total)
            if not ok:
                raise RuntimeError(사유)
            os.replace(part, tmp)
            self._set_status(status, "교체 후 자동 재시작합니다...", OKC)
            self._launch_replace_bat(tmp, target)
            self.root.after(400, self._quit_for_update)
        except Exception as ex:
            try:
                if os.path.exists(part):
                    os.remove(part)
            except OSError:
                pass
            self._manual_fallback(win, status, "자동 업데이트 실패: %s" % str(ex)[:60], btn_later)

    def _quit_for_update(self):
        try:
            self.root.destroy()
        finally:
            os._exit(0)

    def _launch_replace_bat(self, tmp_exe, target_exe):
        bat = os.path.join(tempfile.gettempdir(), "OrderConverter_update.bat")
        bak = target_exe + ".bak"
        exe_name = os.path.basename(target_exe)
        lines = [
            "@echo off",
            'copy /y "%s" "%s" >nul 2>&1' % (target_exe, bak),
            'if not exist "%s" goto norepl' % bak,
            "set /a n=0",
            ":retry",
            "ping -n 2 127.0.0.1 >nul",
            'move /y "%s" "%s" >nul 2>&1' % (tmp_exe, target_exe),
            "if not errorlevel 1 goto launch",
            "set /a n+=1",
            "if %n% lss 40 goto retry",
            'start "" "%s"' % target_exe,
            "goto cleanup",
            ":norepl",
            'start "" "%s"' % target_exe,
            "goto cleanup",
            ":launch",
            "ping -n 9 127.0.0.1 >nul",
            'start "" "%s"' % target_exe,
            "set /a m=0",
            ":check",
            "ping -n 3 127.0.0.1 >nul",
            'tasklist /fi "imagename eq %s" 2>nul | find /i "%s" >nul' % (exe_name, exe_name),
            "if not errorlevel 1 goto cleanup",
            "set /a m+=1",
            "if %m% lss 4 goto check",
            'if exist "%s" move /y "%s" "%s" >nul 2>&1' % (bak, bak, target_exe),
            'if exist "%s" start "" "%s"' % (target_exe, target_exe),
            ":cleanup",
            'if exist "%s" del "%s" >nul 2>&1' % (bak, bak),
            'del "%~f0"',
        ]
        with open(bat, "w", encoding="mbcs", errors="replace") as f:
            f.write("\r\n".join(lines) + "\r\n")
        DETACHED = 0x00000008
        subprocess.Popen(["cmd", "/c", bat], creationflags=DETACHED, close_fds=True)


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    load_fonts()
    root = tk.Tk()
    ConverterApp(root)
    if os.environ.get("ORDERCONV_SMOKE") == "1":
        root.after(4000, root.destroy)
    root.mainloop()
    if os.environ.get("ORDERCONV_SMOKE") == "1":
        print("SMOKE-OK", flush=True)


if __name__ == '__main__':
    main()
