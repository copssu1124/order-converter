# -*- coding: utf-8 -*-
"""주문서 변환기 — CustomTkinter 리디자인 UI (파랑/네이비 테마. '주문서 변환기 리디자인.pdf' 5p 스펙표 값 그대로).
   변환/분리/집계/업데이트 로직은 gui.py(ConverterApp)의 메서드를 그대로 빌려 쓰고, 화면 접점만 여기서 새로 구현한다.
   폰트는 px(음수) = DPI 인식 픽셀."""
import os
import sys
import ctypes
import queue
import threading
import datetime
import time
import webbrowser
import tkinter as tk
import tkinter.font as tkfont
import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont

import gui as G0                      # 기존 프로그램: 로직·상수(VERSION, GITHUB_REPO, KAKAO …)의 단일 출처
engine = G0.engine
VERSION = G0.VERSION

# ── 색 토큰 (PDF 5p '색' + '위젯 값' 그대로) ──
BG = "#F3F5F8"; CARD = "#FFFFFF"; CARD_BD = "#E1E6EE"; CTRL_BD = "#D3DAE6"
INK = "#141D2E"; SUB = "#5A6579"; FAINT = "#96A3BC"; VALNONE = "#C4CBD6"
BRAND = "#2563C9"; BRAND_HOVER = "#1E52A6"; BRAND_DIS = "#C3CEDE"
SEL_BG = "#F2F6FD"; SEL_BD = "#C9DBF7"
SIDEBAR = "#131C2E"; SIDE_SEL = "#22304C"; SIDE_HOVER = "#1B2740"
SIDE_TXT = "#96A3BC"; SIDE_TXT_SEL = "#FFFFFF"; ACCENT = "#6EA8FF"
SIDE_DIV = "#212D43"; SIDE_SUB = "#637088"; CREDIT = "#474F62"      # 스펙표에 없어 목업에서 샘플링
BTN_TXT = "#2B3547"; HDR_TXT = "#3C4757"; HDR_BD = "#E1E6EE"
REQ = "#B42318"; OPT = "#7D879B"
WARN_TXT = "#B45309"; WARN_BG = "#FDF6EC"; WARN_BD = "#F0D8B4"; WARN_BTN = "#9A4E0E"; WARN_BTN_H = "#7D3F0B"
TOTAL_BG = "#F7F9FC"; LOG_DIV = "#F0F1F3"; ROW_DIV = "#EAEEF5"; TS = "#A3ACBA"   # 목업 샘플링 / 스펙 구분선·시각색
CHART = ("#2563C9", "#5B8FE0", "#9BBCEF")                                        # 스펙: 차트/목록 점 색
BW = 1   # 스펙: 1px 테두리. (자식 위젯이 테두리 위를 덮지 않도록 카드 안쪽 여백에 BW 적용)
FONT_PLUS = int(os.environ.get("CTK_FONT_PLUS", "2"))      # 글자 크기 보정: 스펙보다 +2px (사용자 요청: 가독성)
UI_SCALE = float(os.environ.get("CTK_SCALE", "1"))         # 전체 확대(창 포함, 비율 유지) 기본 1.0
SENDERS = ["제이제이컴퍼니", "아이스앤팩", "다다쇼핑", "다모아패키지"]

HELP_SECTIONS = [
    ("주문서 변환",
     "주문서(.xls)와 매핑표(1차)를 고르고 [변환 실행]을 누르세요. 매핑표는 프로그램 폴더에 있으면 자동으로 잡힙니다.\n"
     "매핑표에 '수정데이터' 탭이 있으면 3번 칸이 자동으로 채워지고, 1차 변환 뒤 발주용(2차)까지 한 번에 만듭니다.\n"
     "결과는 저장 폴더 안에 날짜 › 시간 › 변환결과 순서로 정리되고, 원본과 매핑표는 수정하지 않습니다.\n"
     "※ 이미 변환된 결과 파일을 넣으면 안내하고 멈춥니다. 결과로 발주서를 만들려면 택배사 분리로 가세요."),
    ("택배사 분리",
     "방금 변환한 결과가 자동으로 들어갑니다(다른 파일을 골라도 됩니다). 보내는 회사(발송인)를 고르고 [택배사 분리 실행]을 누르면\n"
     "같은 세션 폴더의 '분리출력'에 택배사별 발주서가 만들어집니다. 발송인의 이름·전화(로젠은 주소까지)가 발주서에 들어갑니다.\n"
     "양식: 씨제이(R열 운임) · 대신/대신낱개(연두=낱개) · 대신택배 · 천일(K열) · 로젠(G열) · 원준(L열) · 위플(F열) · 올담 · 카몬드 / 그 외는 '기타'"),
    ("매입·매출 집계",
     "저장 폴더의 최근 변환결과가 자동으로 들어갑니다. [집계 실행]을 누르면 매입(출고지별)과 매출(별칭별) 파일이 '매입매출' 폴더에 저장되고,\n"
     "오른쪽 표에서 매입/매출을 전환해 건수·금액·합계를 바로 볼 수 있습니다. 출고지 열이 없는 주문서는 매출만 집계됩니다."),
    ("저장 폴더",
     "처음 한 번 [변경]으로 지정하면 기억합니다. 모든 결과는 그 폴더 안에 날짜·시간별로 자동 정리되어 나중에 찾기 쉽습니다."),
    ("주문서 규칙",
     "· 수령자·상품명이 빈 행에서 멈춥니다\n· 옵션(선택사항) 우선, 없으면 상품명 사용\n· 송장·배송번호 등 빈칸은 그대로 보존\n"
     "· 스마트스토어·샵마인 주문서는 머리글 이름으로 자동 인식"),
    ("매핑표 규칙",
     "· 상품명변경: 옵션명 → 회사상품명\n· 판매비변경: 회사상품명*수량 → 배송비(등급*택배사 / 순수금액 / 천일박스 / …수정)\n"
     "· 택배비 시트: 천일·씨제이·로젠·대신택배·위플·원준·카몬드·올담·용차 + 수정택배비\n"
     "· 수정데이터: 기존품명 → 수정품명(묶음/낱개)·운임·등급·택배사 — 2차 재변환에 사용\n"
     "· 빨간 행: 매핑표에 키가 없으면 → 매핑표에 추가하면 해결 / 키가 있는데도 빨강 → 프로그램 문의"),
]


def _enable_dpi():
    # Per-Monitor-v2 DPI 인식(가장 선명). 창 생성 전에 호출해야 함.
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def load_fonts():
    # Pretendard는 TTF("Pretendard JJ")로만 로드. OTF는 tkinter(GDI)에서 작은 크기가 흐릿해 사용 금지.
    # 이름을 'Pretendard JJ'로 고유화해 시스템 설치 Pretendard(OTF)와의 충돌도 차단. 없으면 맑은 고딕.
    try:
        for fn in ("PretendardJJ-Regular.ttf", "PretendardJJ-SemiBold.ttf", "PretendardJJ-Bold.ttf"):
            p = os.path.join(G0.ui_dir(), fn)
            if os.path.exists(p):
                ctypes.windll.gdi32.AddFontResourceExW(p, 0x10, 0)
    except Exception:
        pass


def _hairline(parent, color):
    # CTkFrame은 height=1을 그리지 않으므로 순수 tk.Frame으로 정확히 1px 선을 그림
    return tk.Frame(parent, height=1, bg=color, bd=0, highlightthickness=0)


class _Btn:
    """기존 로직(run_convert/_worker 등)이 부르는 set_enabled/set_text 를 CTkButton에 연결하는 어댑터."""
    def __init__(self, btn, idle_text, restyle=None):
        self.btn, self.idle, self.restyle = btn, idle_text, restyle

    def set_enabled(self, v):
        self.btn.configure(state="normal" if v else "disabled")
        if self.restyle:
            self.restyle(bool(v))

    def set_text(self, t):
        self.btn.configure(text=t if "중" in str(t) else self.idle())


class _Slot:
    """파일 줄 (스펙: height=68, r10, border 1 / 빈: #FFF·#D3DAE6 / 선택: #F2F6FD·#C9DBF7)."""
    def __init__(self, app, parent, num, label, req, placeholder, on_pick=None, on_clear=None):
        self.app, self.num, self.placeholder, self.filled = app, num, placeholder, False
        self.fr = ctk.CTkFrame(parent, height=68, corner_radius=10, border_width=BW, fg_color=CARD, border_color=CTRL_BD)
        self.fr.pack(fill="x", pady=(0, 7))
        self.fr.pack_propagate(False)
        self.badge = ctk.CTkLabel(self.fr, text="", image=app._badge(num, False))
        self.badge.pack(side="left", padx=(16, 13))
        mid = ctk.CTkFrame(self.fr, fg_color="transparent")
        mid.pack(side="left", fill="x", expand=True)
        lab = ctk.CTkFrame(mid, fg_color="transparent")
        lab.pack(anchor="w")
        ctk.CTkLabel(lab, text=label + "  ", font=app.f(12), text_color=SUB).pack(side="left")
        ctk.CTkLabel(lab, text=req, font=app.f(12, "bold"), text_color=REQ if req == "필수" else OPT).pack(side="left")
        self.sub = ctk.CTkLabel(lab, text="", font=app.f(12), text_color=BRAND)
        self.sub.pack(side="left", padx=(8, 0))
        self.name = ctk.CTkLabel(mid, text=placeholder, font=app.f(14), text_color=FAINT)
        self.name.pack(anchor="w", pady=(2, 0))
        self.btn = app._sec_btn(self.fr, "선택")
        if on_pick:
            self.btn.configure(command=on_pick)
        self.btn.pack(side="right", padx=16)
        self.x = ctk.CTkButton(self.fr, text="✕", width=28, height=28, corner_radius=8, fg_color="transparent",
                               hover_color=BG, text_color=FAINT, font=app.f(13), command=on_clear)

    def set(self, filename, sub=""):
        self.filled = True
        self.fr.configure(fg_color=SEL_BG, border_color=SEL_BD)
        self.badge.configure(image=self.app._badge(self.num, True))
        shown = filename if len(filename) <= 32 else filename[:15] + "…" + filename[-14:]
        self.name.configure(text=shown, font=self.app.f(14, "bold"), text_color=INK)
        self.sub.configure(text=sub or "")
        self.btn.configure(text="변경")
        self.x.pack(side="right")

    def clear(self):
        self.filled = False
        self.fr.configure(fg_color=CARD, border_color=CTRL_BD)
        self.badge.configure(image=self.app._badge(self.num, False))
        self.name.configure(text=self.placeholder, font=self.app.f(14), text_color=FAINT)
        self.sub.configure(text="")
        self.btn.configure(text="선택")
        self.x.pack_forget()


class App(ctk.CTk):
    def __init__(self):
        _enable_dpi()          # super().__init__() 가 창을 만들기 전에 DPI 인식 설정
        super().__init__()
        ctk.set_appearance_mode("light")
        if UI_SCALE != 1:
            ctk.set_widget_scaling(UI_SCALE)
            ctk.set_window_scaling(UI_SCALE)
        # ── 상태 (gui.ConverterApp 과 동일한 이름: 빌려 쓰는 로직이 그대로 동작) ──
        self.root = self
        self.input_file = self.mapping_file = self.sus_file = self.output_file = None
        self.conv_file = self.split_dir = self.maemae_file = self.maemae_dir = None
        self.log_queue = queue.Queue()
        self._busy = False
        self._c1 = self._c2 = self._c3 = self._s1 = self._s2 = self._m1 = (None,)   # 옛 배지 참조(무시)
        self.발송인_var = tk.StringVar(value=SENDERS[0])
        self.last_issues = []
        self.folder_lbls, self.save_cards, self.tiles, self.slots, self.logs = [], [], {}, {}, {}
        self._log_target = None
        self._load_config()

        fams = tkfont.families()
        malgun = "맑은 고딕" if "맑은 고딕" in fams else "Malgun Gothic"
        self.reg = "Pretendard JJ" if "Pretendard JJ" in fams else malgun
        self.semi = "Pretendard JJ SemiBold" if "Pretendard JJ SemiBold" in fams else self.reg
        self._icons = {}
        self._fonts = {}
        self.title("주문서 변환기 · JJ COMPANY   v" + VERSION)
        self.geometry("1100x720")          # 스펙
        self.minsize(1040, 720)            # 스펙
        self.configure(fg_color=BG)
        try:
            self.iconbitmap(os.path.join(G0.ui_dir(), "appicon.ico"))
        except Exception:
            pass
        try:
            self._appicon = tk.PhotoImage(file=os.path.join(G0.ui_dir(), "appicon.png"))
            self.iconphoto(True, self._appicon)
        except Exception:
            pass
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_content()
        self._select(os.environ.get("CTK_SCREEN", "conv"))

        # ── 시작 동작 (기존과 동일) ──
        self._update_folderbar()
        auto = engine.find_mapping_file(G0.app_dir())
        if auto:
            self.mapping_file = auto
            self._set_file("map", os.path.basename(auto), "자동으로 찾았어요 ✓", ok=True)
            self._auto_sus(auto)
        self._refresh_conv()
        self._drain_log()
        self.bind("<Control-Return>", lambda e: self.run_convert())
        self.after(120, self._apply_taskbar_icon)
        self._bring_to_front()
        self._show_update_banner_if_needed()
        threading.Thread(target=self._check_update, daemon=True).start()

    # 폰트: size 음수 = 픽셀(스펙 px와 1:1). CTk가 DPI 배율을 곱함. w: r / semi / bold
    def f(self, px, w="r", plus=True):
        # plus=False: 로그·집계 표처럼 정보 밀도가 중요한 곳은 스펙 크기 그대로
        key = (px, w, plus)
        cached = self._fonts.get(key)
        if cached is not None:
            return cached
        px = px + (FONT_PLUS if plus else 0)
        if w == "semi":
            fnt = ctk.CTkFont(family=self.semi, size=-px)
        elif w == "bold":
            fnt = ctk.CTkFont(family=self.reg, size=-px, weight="bold")
        else:
            fnt = ctk.CTkFont(family=self.reg, size=-px)
        self._fonts[key] = fnt
        return fnt

    # 라인 아이콘 (스펙: 이모지 금지, 20px 이미지를 CTkImage로). 44px 캔버스에 그린 뒤 다운스케일.
    def icon(self, kind, color, px=20):
        key = (kind, color, px)
        if key in self._icons:
            return self._icons[key]
        S = 44; w = 3
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        if kind == "convert":
            d.line([(9, 17), (33, 17)], fill=color, width=w)
            d.line([(27, 11), (34, 17)], fill=color, width=w); d.line([(27, 23), (34, 17)], fill=color, width=w)
            d.line([(11, 29), (35, 29)], fill=color, width=w)
            d.line([(17, 23), (10, 29)], fill=color, width=w); d.line([(17, 35), (10, 29)], fill=color, width=w)
        elif kind == "truck":
            d.rectangle([(6, 15), (25, 30)], outline=color, width=w)
            d.line([(25, 21), (33, 21)], fill=color, width=w)
            d.line([(33, 21), (38, 26)], fill=color, width=w)
            d.line([(38, 26), (38, 30)], fill=color, width=w)
            d.line([(25, 30), (38, 30)], fill=color, width=w)
            d.ellipse([(10, 29), (18, 37)], outline=color, width=w)
            d.ellipse([(28, 29), (36, 37)], outline=color, width=w)
        elif kind == "chart":
            d.line([(8, 35), (37, 35)], fill=color, width=w)
            d.rectangle([(11, 25), (17, 34)], outline=color, width=w)
            d.rectangle([(20, 16), (26, 34)], outline=color, width=w)
            d.rectangle([(29, 21), (35, 34)], outline=color, width=w)
        elif kind == "doc":
            d.line([(13, 6), (13, 38), (31, 38), (31, 14), (23, 6), (13, 6)], fill=color, width=w, joint="curve")
            d.line([(23, 6), (23, 14), (31, 14)], fill=color, width=w)
            for yy in (20, 26, 32):
                d.line([(18, yy), (26, yy)], fill=color, width=2)
        elif kind == "chev":
            d.line([(13, 17), (22, 27), (31, 17)], fill=color, width=3, joint="curve")
        elif kind == "dot":
            d.ellipse([10, 10, 34, 34], fill=color)
        elif kind == "folder":
            d.line([(7, 33), (7, 16), (16, 16), (19, 20), (37, 20), (37, 33), (7, 33)],
                   fill=color, width=w, joint="curve")
        im = ctk.CTkImage(light_image=img, dark_image=img, size=(px, px))
        self._icons[key] = im
        return im

    # 번호 뱃지: 원 + 숫자를 PIL 4배 슈퍼샘플링으로 구움(매끈한 원). 빈 상태 = 테두리 #D3DAE6 + 배경 흰색(스펙)
    def _badge(self, num, sel, dia=26):
        key = ("badge", num, sel, dia)
        if key in self._icons:
            return self._icons[key]
        S = dia * 4
        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        if sel:
            d.ellipse([2, 2, S - 2, S - 2], fill=BRAND); tcol = "white"
        else:
            d.ellipse([3, 3, S - 3, S - 3], fill=CARD, outline=CTRL_BD, width=4); tcol = FAINT
        try:
            fnt = ImageFont.truetype(os.path.join(G0.ui_dir(), "PretendardJJ-SemiBold.ttf"), int(S * 0.44))
        except Exception:
            fnt = ImageFont.load_default()
        t = str(num)
        bb = d.textbbox((0, 0), t, font=fnt)
        d.text(((S - (bb[2] - bb[0])) / 2 - bb[0], (S - (bb[3] - bb[1])) / 2 - bb[1]), t, font=fnt, fill=tcol)
        im = ctk.CTkImage(light_image=img, dark_image=img, size=(dia, dia))
        self._icons[key] = im
        return im

    # ── 사이드바 (스펙: CTkFrame(width=232, corner_radius=0, fg_color=#131C2E)) ──
    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=232, corner_radius=0, fg_color=SIDEBAR)
        sb.grid(row=0, column=0, sticky="ns")
        sb.grid_propagate(False)
        brand = ctk.CTkFrame(sb, fg_color="transparent")
        brand.pack(fill="x", padx=18, pady=(13, 20))
        ctk.CTkLabel(brand, text="JJ", width=34, height=34, corner_radius=9,
                     fg_color=BRAND, text_color="white", font=self.f(15, "bold")).pack(side="left")
        tx = ctk.CTkFrame(brand, fg_color="transparent")
        tx.pack(side="left", padx=(11, 0))
        ctk.CTkLabel(tx, text="주문서 변환기", font=self.f(16, "bold"), text_color="white").pack(anchor="w")
        ctk.CTkLabel(tx, text="JJ COMPANY", font=self.f(11), text_color=SIDE_SUB).pack(anchor="w", pady=(1, 0))
        # 내비 항목 (스펙: height=42, corner_radius=9, anchor=w, hover #1B2740, text #96A3BC / 선택 fg #22304C text #FFF)
        self.nav = {}
        items = (("conv", "convert", "주문서 변환"), ("split", "truck", "택배사 분리"), ("maemae", "chart", "매입·매출"))
        for key, ic, label in items:
            row = ctk.CTkFrame(sb, height=46, fg_color="transparent")
            row.pack(fill="x", pady=1)
            row.pack_propagate(False)
            # 선택 표시바 (스펙: CTkFrame(width=3, height=20, fg_color=#6EA8FF) 를 버튼 왼쪽에 place)
            bar = ctk.CTkFrame(row, width=3, height=20, corner_radius=2, fg_color="transparent")
            bar.place(x=0, rely=0.5, anchor="w")
            btn = ctk.CTkButton(row, text="  " + label, image=self.icon(ic, SIDE_TXT),
                                anchor="w", compound="left", height=42, corner_radius=9,
                                fg_color="transparent", hover_color=SIDE_HOVER, text_color=SIDE_TXT,
                                font=self.f(14, "semi"), command=lambda k=key: self._select(k))
            btn.pack(fill="both", expand=True, padx=(12, 12), pady=2)
            self.nav[key] = (bar, btn, ic)
        ctk.CTkFrame(sb, fg_color="transparent").pack(fill="both", expand=True)
        _hairline(sb, SIDE_DIV).pack(fill="x", padx=18, pady=(0, 12))
        chip = ctk.CTkFrame(sb, height=34, corner_radius=8, fg_color=SIDE_HOVER)
        chip.pack(fill="x", padx=18)
        chip.pack_propagate(False)
        ctk.CTkLabel(chip, text="버전 v" + VERSION, font=self.f(13), text_color=SIDE_TXT).pack(side="left", padx=12)
        self.ver_lbl = ctk.CTkLabel(chip, text="확인 중…", font=self.f(13, "semi"), text_color=FAINT)
        self.ver_lbl.pack(side="right", padx=12)
        ctk.CTkLabel(sb, text="JANG JUNG WOO · JJ COMPANY", font=self.f(10), text_color=CREDIT).pack(pady=(12, 14))

    # ── 콘텐츠(헤더 + 화면) ──
    def _build_content(self):
        c = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        c.grid(row=0, column=1, sticky="nsew")
        c.grid_columnconfigure(0, weight=1)
        c.grid_rowconfigure(1, weight=1)
        hdr = ctk.CTkFrame(c, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=30, pady=(26, 6))
        hdr.grid_columnconfigure(0, weight=1)
        hdr.grid_rowconfigure(0, minsize=72)     # 설명문이 1줄이든 2줄이든 헤더 높이 고정 → 화면 전환 때 본문이 안 튐
        left = ctk.CTkFrame(hdr, fg_color="transparent")
        left.grid(row=0, column=0, sticky="w")
        left.grid_propagate(True)
        self.h_title = ctk.CTkLabel(left, text="", font=self.f(22, "bold"), text_color=INK)      # 화면 제목 22px Bold
        self.h_title.pack(anchor="w")
        self.h_sub = ctk.CTkLabel(left, text="", font=self.f(13), text_color=SUB,               # 설명 문장 12.5px
                                  wraplength=540, justify="left", anchor="w")
        self.h_sub.pack(anchor="w", pady=(3, 0))
        right = ctk.CTkFrame(hdr, fg_color="transparent")
        right.grid(row=0, column=1, sticky="e")
        # 헤더 버튼 (스펙: height=30, corner_radius=7, border #E1E6EE, text #3C4757). 매뉴얼만 문서 아이콘.
        cmds = {"매뉴얼": self.open_manual, "설명서": self.open_help, "문의": lambda: webbrowser.open(G0.KAKAO)}
        for t in ("매뉴얼", "설명서", "문의"):
            img = self.icon("doc", HDR_TXT, 15) if t == "매뉴얼" else None
            ctk.CTkButton(right, text=t, image=img, compound="left", height=32,
                          width=86 if t == "매뉴얼" else 70, corner_radius=7, border_width=BW,
                          border_color=HDR_BD, fg_color=CARD, hover_color=BG, text_color=HDR_TXT,
                          font=self.f(13, "semi"), command=cmds[t]).pack(side="left", padx=(7, 0))
        self.body = ctk.CTkFrame(c, fg_color="transparent")
        self.body.grid(row=1, column=0, sticky="nsew", padx=30, pady=(10, 24))
        self.body.grid_rowconfigure(0, weight=1)
        self.body.grid_columnconfigure(0, weight=1)
        self.screens = {"conv": self._build_conv(), "split": self._build_split(), "maemae": self._build_maemae()}
        for fr in self.screens.values():        # 세 화면을 같은 칸에 겹쳐 두고 tkraise 로 전환(재배치·재그리기 없음 → 빠름)
            fr.grid(row=0, column=0, sticky="nsew")

    # ── 공통 부품 ──
    def _card(self, parent):
        # 스펙: CTkFrame(corner_radius=12, fg #FFFFFF, border_width=1, border #E1E6EE) · 안쪽 여백 18
        return ctk.CTkFrame(parent, corner_radius=12, fg_color=CARD, border_width=BW, border_color=CARD_BD)

    def _sec_btn(self, parent, text, height=34, width=62, command=None):
        # 보조 버튼 (스펙: height=34, corner_radius=8, fg #FFF, hover #F3F5F8, border #D3DAE6, text #2B3547)
        return ctk.CTkButton(parent, text=text, width=width, height=height, corner_radius=8,
                             fg_color=CARD, hover_color=BG, border_width=BW, border_color=CTRL_BD,
                             text_color=BTN_TXT, font=self.f(13, "semi"), command=command)

    def _slot(self, parent, num, label, req, filename="", placeholder="파일을 선택하세요", on_pick=None, on_clear=None):
        s = _Slot(self, parent, num, label, req, placeholder, on_pick, on_clear)
        if filename:
            s.set(filename)
        return s

    def _tile(self, parent, key, label):
        t = ctk.CTkFrame(parent, height=92, corner_radius=12, fg_color=CARD, border_width=BW, border_color=CARD_BD)
        t.pack_propagate(False)
        lb = ctk.CTkLabel(t, text=label, font=self.f(12, "semi"), text_color=SUB)
        lb.pack(anchor="w", padx=15, pady=(14, 0))
        val = ctk.CTkLabel(t, text="—", font=self.f(26, "bold"), text_color=VALNONE)        # 숫자 26px Bold
        val.pack(anchor="w", padx=15)
        self.tiles[key] = (t, lb, val)
        return t

    def _set_tile(self, key, value, kind=""):
        t, lb, val = self.tiles[key]
        fg, bd, lc, vc = CARD, CARD_BD, SUB, INK
        if kind == "good":
            fg, bd, lc, vc = SEL_BG, SEL_BD, BRAND, BRAND
        elif kind == "warn":
            fg, bd, lc, vc = WARN_BG, WARN_BD, WARN_TXT, WARN_TXT
        t.configure(fg_color=fg, border_color=bd)
        lb.configure(text_color=lc)
        val.configure(text=value, text_color=vc if value != "—" else VALNONE)

    def _save_card(self, parent):
        sf = self._card(parent)
        sf.pack(fill="x", pady=(13, 0))
        si = ctk.CTkFrame(sf, fg_color="transparent")
        si.pack(fill="x", padx=18, pady=11)
        ctk.CTkLabel(si, text="", image=self.icon("folder", "#98A1B3", 22)).pack(side="left", padx=(2, 12))
        sm = ctk.CTkFrame(si, fg_color="transparent")
        sm.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(sm, text="저장 폴더", font=self.f(12), text_color=SUB).pack(anchor="w")
        lbl = ctk.CTkLabel(sm, text="", font=self.f(14, "bold"), text_color=INK)
        lbl.pack(anchor="w", pady=(1, 0))
        self._sec_btn(si, "변경", height=32, width=58, command=self.pick_folder).pack(side="right")   # 스펙: 저장 폴더 옆 32
        self.folder_lbls.append(lbl)
        self.save_cards.append(sf)
        return sf

    def _card_header(self, card, title, right_text, right_color, right_w="r"):
        h = ctk.CTkFrame(card, fg_color="transparent")
        h.pack(fill="x", padx=18, pady=(6, 7))
        ctk.CTkLabel(h, text=title, font=self.f(14, "bold"), text_color=INK).pack(side="left")
        st = ctk.CTkLabel(h, text=right_text, font=self.f(12, right_w), text_color=right_color)
        st.pack(side="right")
        _hairline(card, LOG_DIV).pack(fill="x", padx=BW)
        return st

    def _result_row(self, parent, dot, name, count, last=False):
        r = ctk.CTkFrame(parent, height=50, fg_color="transparent")      # 목업 행 간격 48~50
        r.pack(fill="x", padx=18)
        r.pack_propagate(False)
        ctk.CTkLabel(r, text="", image=self.icon("dot", dot, 10)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(r, text=name, font=self.f(14, "semi"), text_color=INK).pack(side="left")
        ctk.CTkLabel(r, text="›", font=self.f(18), text_color=FAINT).pack(side="right")
        ctk.CTkLabel(r, text=count, font=self.f(14, "semi"), text_color=INK).pack(side="right", padx=(0, 14))
        if not last:
            _hairline(parent, ROW_DIV).pack(fill="x", padx=18)

    def _table_row(self, parent, cols, bold=False, bg="transparent", header=False):
        r = ctk.CTkFrame(parent, height=34 if header else 46, fg_color=bg, corner_radius=0)   # 스펙 rowheight=46
        r.pack(fill="x", padx=BW)
        r.pack_propagate(False)
        r.grid_propagate(False)          # 자식이 grid라 이것도 꺼야 height가 유지됨
        r.grid_columnconfigure(0, weight=1)
        r.grid_columnconfigure(1, minsize=70)
        r.grid_columnconfigure(2, minsize=110)
        r.grid_rowconfigure(0, weight=1)
        for i, txt in enumerate(cols):
            if header:
                fnt, col = self.f(12, plus=False), FAINT
            else:
                fnt = self.f(14, "bold", plus=False) if (bold or i == 2) else self.f(14, "semi" if i == 0 else "r", plus=False)
                col = INK
            ctk.CTkLabel(r, text=txt, font=fnt, text_color=col, anchor="w" if i == 0 else "e").grid(
                row=0, column=i, sticky="w" if i == 0 else "e", padx=(16, 0) if i == 0 else (0, 16 if i == 2 else 0))
        return r

    # 로그 영역 (스펙: CTkTextbox r0 border0 흰배경 · 시각은 Consolas 태그 #A3ACBA · 경고 태그 #B45309)
    def _log_box(self, parent, lines=(), height=None, show=True):
        tb = ctk.CTkTextbox(parent, corner_radius=0, border_width=0, fg_color=CARD, text_color=INK,
                            font=self.f(13, plus=False), activate_scrollbars=False, height=height or 200)
        if show:
            tb.pack(fill="x" if height else "both", expand=not height, padx=12, pady=(6, 12))   # r12 모서리를 덮지 않게
        tb._textbox.configure(spacing3=11, padx=6)                                 # ponytail: 줄간격은 tk.Text 옵션 직접
        tb._textbox.tag_config("ts", foreground=TS, font=("Consolas", -12))          # ponytail: CTk가 tag font를 막음
        tb.tag_config("warn", foreground=WARN_TXT)
        for ts, msg in lines:
            tb.insert("end", ts + "   ", "ts")
            tb.insert("end", msg + "\n", ("warn",) if msg.startswith("!") else ())
        tb.configure(state="disabled")
        return tb

    def _screen(self):
        scr = ctk.CTkFrame(self.body, fg_color="transparent")
        scr.grid_columnconfigure(0, minsize=468)
        scr.grid_columnconfigure(1, weight=1)
        scr.grid_rowconfigure(0, weight=1)
        left = ctk.CTkFrame(scr, fg_color="transparent", width=468)
        left.grid(row=0, column=0, sticky="nsew")
        left.grid_propagate(False)       # 긴 경로·파일명이 왼쪽 열을 늘려 오른쪽을 밀어내지 않도록 폭 고정
        right = ctk.CTkFrame(scr, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(22, 0))
        right.grid_columnconfigure(0, weight=1)
        return scr, left, right

    def _main_btn(self, parent, text, icon, command=None, pady=(13, 0)):
        # 주 버튼 (스펙: height=54, corner_radius=11, fg #2563C9, hover #1E52A6, 글자 16px Bold)
        b = ctk.CTkButton(parent, text="   " + text, image=self.icon(icon, "white", 20), compound="left",
                          height=54, corner_radius=11, fg_color=BRAND, hover_color=BRAND_HOVER, text_color="white",
                          font=self.f(16, "bold"), command=command)
        b.pack(fill="x", pady=pady)
        return b

    # ── ② 주문서 변환 화면 (PDF 1p 빈 상태 → 2p 완료 상태) ──
    def _build_conv(self):
        scr, left, right = self._screen()
        fc = self._card(left)
        fc.pack(fill="x")
        fh = ctk.CTkFrame(fc, fg_color="transparent")
        fh.pack(fill="x", padx=18, pady=(14, 10))
        ctk.CTkLabel(fh, text="불러올 파일", font=self.f(14, "bold"), text_color=INK).pack(side="left")   # 카드 제목 14 Bold
        self.count_lbl = ctk.CTkLabel(fh, text="0 / 3 선택됨", font=self.f(13, "semi"), text_color=BRAND)
        self.count_lbl.pack(side="right")
        slots = ctk.CTkFrame(fc, fg_color="transparent")
        slots.pack(fill="x", padx=18, pady=(0, 14))
        self.slots["order"] = self._slot(slots, 1, "주문서", "필수", on_pick=self.pick_input, on_clear=lambda: self._clear("order"))
        self.slots["map"] = self._slot(slots, 2, "매핑표 (1차)", "필수", on_pick=self.pick_mapping, on_clear=lambda: self._clear("map"))
        self.slots["sus"] = self._slot(slots, 3, "수정데이터 (2차)", "선택", placeholder="넣으면 발주용까지 한 번에 만듭니다",
                                       on_pick=self.pick_sus, on_clear=lambda: self._clear("sus"))
        self._save_card(left)
        self.btn_conv_w = self._main_btn(left, "변환 실행", "convert", command=self.run_convert)
        self.btn_conv = _Btn(self.btn_conv_w, lambda: "   다시 변환" if self.output_file else "   변환 실행", self._style_conv_btn)
        self.conv_caption = ctk.CTkLabel(left, text="단축키 Ctrl + Enter", font=self.f(12), text_color=OPT)
        self.conv_caption.pack(pady=(10, 0))

        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure((0, 1, 2), weight=1)
        for i, (key, lb) in enumerate((("total", "총 주문"), ("good", "정상"), ("issue", "확인 필요"))):
            self._tile(right, key, lb).grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 11, 0))
        logc = self._card(right)
        logc.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(17, 0))
        self.log_status = self._card_header(logc, "진행 상황", "● 대기 중", FAINT)
        self.log_empty = ctk.CTkFrame(logc, fg_color="transparent")
        self.log_empty.pack(fill="both", expand=True, padx=12, pady=(0, 12))   # 둥근 모서리(r12)·테두리를 덮지 않게
        box = ctk.CTkFrame(self.log_empty, fg_color="transparent")
        box.pack(expand=True)
        ctk.CTkLabel(box, text="", image=self.icon("doc", VALNONE, 34)).pack(pady=(0, 10))
        ctk.CTkLabel(box, text="파일을 고르고 [변환 실행]을 누르면", font=self.f(14), text_color=OPT).pack()
        ctk.CTkLabel(box, text="여기에 진행 상황이 표시됩니다.", font=self.f(14), text_color=OPT).pack()
        self.logbox = self._log_box(logc, show=False)
        self.logs["conv"] = {"status": self.log_status, "empty": self.log_empty, "tb": self.logbox, "height": None}
        # 확인 필요 박스 (PDF 2p) — 결과 있을 때만 표시
        self.warn_box = ctk.CTkFrame(logc, fg_color=WARN_BG, border_width=BW, border_color=WARN_BD, corner_radius=10)
        self.warn_title = ctk.CTkLabel(self.warn_box, text="", font=self.f(13, "bold"), text_color=WARN_TXT)
        self.warn_title.pack(anchor="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(self.warn_box, text="매핑표에 없는 상품이거나 수량·주소를 확인해야 하는 주문입니다. "
                                         "매핑표에 추가한 뒤 다시 변환하거나, 결과 파일에서 직접 채워 주세요.",
                     font=self.f(12), text_color="#7A4A12", justify="left", anchor="w", wraplength=268).pack(anchor="w", padx=14)
        self.warn_btn = ctk.CTkButton(self.warn_box, text="목록 보기", height=32, corner_radius=8, fg_color=WARN_BTN,
                                      hover_color=WARN_BTN_H, text_color="white", font=self.f(13, "semi"),
                                      command=self._show_issue_list)
        self.warn_btn.pack(anchor="w", padx=14, pady=(10, 14))
        self.act_bar = ctk.CTkFrame(logc, fg_color="transparent")
        self.act_bar.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(self.act_bar, text="  결과 폴더 열기", image=self.icon("folder", "white", 18), compound="left", height=40,
                      corner_radius=8, fg_color=BRAND, hover_color=BRAND_HOVER, text_color="white",
                      font=self.f(13, "semi"), command=self.open_result_folder).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ctk.CTkButton(self.act_bar, text="  엑셀로 열기", image=self.icon("doc", HDR_TXT, 16), compound="left", height=40,
                      corner_radius=8, fg_color=CARD, hover_color=BG, border_width=BW, border_color=CTRL_BD,
                      text_color=BTN_TXT, font=self.f(13, "semi"), command=self._open_excel).grid(row=0, column=1, sticky="ew", padx=(4, 0))
        return scr

    # 진행 상황 카드 (헤더 + 빈 상태 + 숨긴 로그 상자). key별로 self.logs에 등록 → _append_log 가 채움
    def _log_card(self, parent, key, height=None):
        card = self._card(parent)
        status = self._card_header(card, "진행 상황", "● 대기 중", FAINT)
        empty = ctk.CTkFrame(card, fg_color="transparent", height=height or 0)
        empty.pack(fill="x" if height else "both", expand=not height, padx=12, pady=(0, 12))
        if height:
            empty.pack_propagate(False)
        box = ctk.CTkFrame(empty, fg_color="transparent")
        box.pack(expand=True)
        if not height:
            ctk.CTkLabel(box, text="", image=self.icon("doc", VALNONE, 34)).pack(pady=(0, 10))
        ctk.CTkLabel(box, text="실행하면 여기에 진행 상황이 표시됩니다.", font=self.f(14 if not height else 13), text_color=OPT).pack()
        tb = self._log_box(card, show=False, height=height)
        self.logs[key] = {"status": status, "empty": empty, "tb": tb, "height": height}
        return card

    # ── ③ 택배사 분리 화면 (PDF 3p) ──
    def _build_split(self):
        scr, left, right = self._screen()
        fc = self._card(left)
        fc.pack(fill="x")
        ctk.CTkLabel(fc, text="불러올 파일", font=self.f(14, "bold"), text_color=INK).pack(anchor="w", padx=18, pady=(14, 10))
        slots = ctk.CTkFrame(fc, fg_color="transparent")
        slots.pack(fill="x", padx=18, pady=(0, 14))
        self.slots["conv"] = self._slot(slots, 1, "변환결과 파일", "필수", on_pick=self.pick_conv, on_clear=lambda: self._clear("conv"))
        # 보내는 회사 (발송인) — 스펙 46/r9/#D3DAE6 상자. 클릭하면 CTk 드롭다운 메뉴
        sc = self._card(left)
        sc.pack(fill="x", pady=(13, 0))
        ctk.CTkLabel(sc, text="보내는 회사 (발송인)", font=self.f(14, "bold"), text_color=INK).pack(anchor="w", padx=18, pady=(11, 0))
        ctk.CTkLabel(sc, text="택배사로 보낼 발주서의 발송인 정보로 들어갑니다.", font=self.f(12), text_color=SUB).pack(anchor="w", padx=18, pady=(2, 9))
        cb = ctk.CTkFrame(sc, height=46, corner_radius=9, border_width=BW, border_color=CTRL_BD, fg_color=CARD)
        cb.pack(fill="x", padx=18, pady=(0, 9))
        cb.pack_propagate(False)
        self.sender_box = cb
        self.sender_lbl = ctk.CTkLabel(cb, text=self.발송인_var.get(), font=self.f(14, "semi"), text_color=INK)
        self.sender_lbl.pack(side="left", padx=16)
        chev = ctk.CTkLabel(cb, text="", image=self.icon("chev", SUB, 16))
        chev.pack(side="right", padx=16)
        for w in (cb, self.sender_lbl, chev):
            w.bind("<Button-1>", lambda e: self._open_sender_menu())
        self._save_card(left)
        self.btn_split_w = self._main_btn(left, "택배사 분리 실행", "truck", command=self.run_split, pady=(8, 0))
        self.btn_split = _Btn(self.btn_split_w, lambda: "   택배사 분리 실행",
                              lambda on: self._style_main_btn(self.btn_split_w, on))
        ctk.CTkLabel(left, text="택배사별로 파일이 각각 저장됩니다.", font=self.f(12), text_color=OPT).pack(pady=(10, 0))

        right.grid_rowconfigure(1, weight=1)
        rc = self._card(right)
        rc.grid(row=0, column=0, sticky="ew")
        self.split_count = self._card_header(rc, "만들어진 발주서", "—", FAINT)
        self.split_rows = ctk.CTkFrame(rc, fg_color="transparent")
        self.split_rows.pack(fill="x")
        ctk.CTkLabel(self.split_rows, text="분리 실행 후 택배사별 발주서가 여기에 표시됩니다.",
                     font=self.f(13), text_color=OPT).pack(pady=22)
        self.btn_split_open = ctk.CTkButton(rc, text="  결과 폴더 열기", image=self.icon("folder", "white", 18), compound="left",
                                            height=38, corner_radius=8, fg_color=BRAND, hover_color=BRAND_HOVER,
                                            text_color="white", font=self.f(13, "semi"), command=self.open_split_folder)
        self._log_card(right, "split").grid(row=1, column=0, sticky="nsew", pady=(17, 0))
        return scr

    # ── ④ 매입·매출 집계 화면 (PDF 4p) ──
    def _build_maemae(self):
        scr, left, right = self._screen()
        fc = self._card(left)
        fc.pack(fill="x")
        ctk.CTkLabel(fc, text="불러올 파일", font=self.f(14, "bold"), text_color=INK).pack(anchor="w", padx=18, pady=(14, 10))
        slots = ctk.CTkFrame(fc, fg_color="transparent")
        slots.pack(fill="x", padx=18, pady=(0, 14))
        self.slots["maemae"] = self._slot(slots, 1, "변환결과 파일", "필수", on_pick=self.pick_maemae,
                                          on_clear=lambda: self._clear("maemae"))
        self._save_card(left)
        self.btn_maemae_w = self._main_btn(left, "집계 실행", "chart", command=self.run_maemae, pady=(8, 0))
        self.btn_maemae = _Btn(self.btn_maemae_w, lambda: "   집계 실행",
                               lambda on: self._style_main_btn(self.btn_maemae_w, on))
        ctk.CTkLabel(left, text="매입·매출 시트가 한 파일로 저장됩니다.", font=self.f(12), text_color=OPT).pack(pady=(10, 0))
        self._log_card(left, "maemae", height=96).pack(fill="x", pady=(6, 0))

        right.grid_rowconfigure(0, weight=1)
        rc = self._card(right)
        rc.grid(row=0, column=0, sticky="nsew")
        self.maemae_count = self._card_header(rc, "집계 결과", "—", FAINT)
        seg = ctk.CTkFrame(rc, fg_color="transparent")
        seg.pack(fill="x", padx=14, pady=(12, 6))
        seg.grid_columnconfigure((0, 1), weight=1)
        self.seg_btns = {}
        for i, (mode, label) in enumerate((("매입", "매입 · 출고지별"), ("매출", "매출 · 별칭별"))):
            b = ctk.CTkButton(seg, text=label, height=36, corner_radius=8, font=self.f(13, "semi"),
                              command=lambda m=mode: self._show_mode(m))
            b.grid(row=0, column=i, sticky="ew", padx=(0, 4) if i == 0 else 0)
            self.seg_btns[mode] = b
        # 표는 그룹 수가 많아질 수 있어 스크롤 프레임 안에 (행높이 46, 구분선 #EAEEF5 — 스펙)
        self.table = ctk.CTkScrollableFrame(rc, fg_color="transparent", corner_radius=0,
                                            scrollbar_button_color=CTRL_BD, scrollbar_button_hover_color=FAINT)
        self.table.pack(fill="both", expand=True, padx=BW)
        self.maemae_data, self.maemae_mode = {}, "매입"
        self._show_mode("매입")
        bb = ctk.CTkFrame(rc, fg_color="transparent")
        bb.pack(fill="x", padx=14, pady=(8, 14))
        bb.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(bb, text="  결과 폴더 열기", image=self.icon("folder", "white", 18), compound="left", height=40,
                      corner_radius=8, fg_color=BRAND, hover_color=BRAND_HOVER, text_color="white",
                      font=self.f(13, "semi"), command=self.open_maemae_folder).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ctk.CTkButton(bb, text="  엑셀로 열기", image=self.icon("doc", HDR_TXT, 16), compound="left", height=40,
                      corner_radius=8, fg_color=CARD, hover_color=BG, border_width=BW, border_color=CTRL_BD,
                      text_color=BTN_TXT, font=self.f(13, "semi"), command=self._open_maemae_excel).grid(row=0, column=1, sticky="ew", padx=(4, 0))
        return scr

    def _select(self, key):
        titles = {
            "conv": ("주문서 변환", "주문서와 매핑표를 고르면 변환합니다. 수정데이터를 함께 넣으면 발주용으로 한 번 더 변환합니다."),
            "split": ("택배사 분리", "변환한 결과를 엑셀에서 검증·보완한 뒤 불러오면 택배사별 발주서로 나눕니다."),
            "maemae": ("매입·매출 집계", "변환결과를 매입은 출고지별로, 매출은 별칭별로 묶어 집계합니다."),
        }
        self._cur = key
        if key == "split" and hasattr(self, "split_rows"):
            self._auto_split_file()
        if key == "maemae" and hasattr(self, "table"):
            self._auto_maemae_file()
        t, s = titles[key]
        self.h_title.configure(text=t)
        self.h_sub.configure(text=s)
        prev = getattr(self, "_prev_key", None)
        self._prev_key = key
        for k in (self.nav.keys() if prev is None else {prev, key}):    # 바뀐 메뉴 두 개만 다시 그림
            bar, btn, ic = self.nav[k]
            on = (k == key)
            bar.configure(fg_color=ACCENT if on else "transparent")
            btn.configure(fg_color=SIDE_SEL if on else "transparent",
                          text_color=SIDE_TXT_SEL if on else SIDE_TXT,
                          font=self.f(14, "bold" if on else "semi"),
                          image=self.icon(ic, ACCENT if on else SIDE_TXT))
        self.screens[key].tkraise()

    # ═══════════ 기존 로직이 만지는 화면 접점 (gui.py 의 옛 위젯 → 새 UI) ═══════════
    def _set_file(self, kind, name, sub="", ok=False):
        s = self.slots.get(kind)
        if s:
            s.set(name, sub)

    def _mark_done(self, *a, **k):
        pass

    def _clear(self, kind):
        attr = {"order": "input_file", "map": "mapping_file", "sus": "sus_file", "conv": "conv_file", "maemae": "maemae_file"}[kind]
        setattr(self, attr, None)
        self.slots[kind].clear()
        {"conv": self._refresh_split, "maemae": self._refresh_maemae}.get(kind, self._refresh_conv)()

    def _refresh_conv(self):
        n = sum(bool(x) for x in (self.input_file, self.mapping_file, self.sus_file))
        self.count_lbl.configure(text="%d / 3 선택됨" % n)
        self.btn_conv.set_enabled(bool(self.input_file and self.mapping_file) and not self._busy)

    def _style_main_btn(self, b, enabled):
        if enabled:
            b.configure(fg_color=BRAND, hover_color=BRAND_HOVER)
        else:
            b.configure(fg_color=BRAND_DIS, hover_color=BRAND_DIS)     # 스펙 비활성 #C3CEDE

    # ── ③ 택배사 분리 접점 ──
    def _refresh_split(self):
        self.btn_split.set_enabled(bool(self.conv_file) and not self._busy)

    # ── ④ 매입·매출 접점 ──
    def _refresh_maemae(self):
        self.btn_maemae.set_enabled(bool(self.maemae_file) and not self._busy)

    def run_maemae(self):
        if self._busy or not self.maemae_file:
            return
        if not self._require_folder():
            return
        self._busy = True
        self._log_target = "maemae"
        self.btn_maemae.set_enabled(False)
        self.btn_maemae.set_text("집계 중…")
        self.log("─" * 30)
        self.log("매입·매출 집계를 시작합니다...")
        threading.Thread(target=self._maemae_worker, daemon=True).start()

    def _maemae_worker(self):
        try:
            out_dir = self._sess_out(self.maemae_file, "매입매출")
            stamp = datetime.datetime.now().strftime("%H%M%S")
            결과 = {}
            for mode in ('매입', '매출'):
                out_path = os.path.join(out_dir, "%s_%s.xlsx" % (mode, stamp))
                rows = []
                try:
                    시트수, 총 = engine.매입매출집계(self.maemae_file, mode, out_path, log=self.log, breakdown=rows)
                    결과[mode] = {"rows": rows, "total": 총, "sheets": 시트수, "path": out_path}
                except Exception as ex:
                    self.log("⚠️ %s 건너뜀: %s" % (mode, str(ex)))
            self.maemae_dir = out_dir
            self.log("💾 저장 완료: " + out_dir)
            self.root.after(0, lambda: self._show_maemae(결과))
        except Exception as ex:
            self.log("❌ 오류: " + str(ex))
            self.log(G0.traceback.format_exc())
        finally:
            self._busy = False
            self.root.after(0, lambda: (self.btn_maemae.set_text("집계 실행"), self._refresh_maemae()))

    def _show_maemae(self, 결과):
        self.maemae_data = 결과
        base = 결과.get("매입") or 결과.get("매출")
        self.maemae_count.configure(text=("%s건" % format(sum(n for _, n, _ in base["rows"]), ",")) if base else "—")
        if 결과 and self.maemae_mode not in 결과:
            self.maemae_mode = next(iter(결과))
        self._show_mode(self.maemae_mode)
        st = self.logs["maemae"]["status"]
        if 결과:
            st.configure(text="● 완료", text_color=BRAND)
        else:
            st.configure(text="● 오류", text_color=REQ)
            self.log("⚠️ 집계할 게 없었어요. (출고지/별칭 열을 확인해 주세요)")

    def _show_mode(self, mode):
        self.maemae_mode = mode
        for m, b in self.seg_btns.items():
            on = (m == mode)
            b.configure(fg_color=BRAND if on else BG, hover_color=BRAND_HOVER if on else "#E9EDF3",
                        text_color="white" if on else HDR_TXT)
        for w in self.table.winfo_children():
            w.destroy()
        self._table_row(self.table, ("출고지" if mode == "매입" else "별칭", "건수", "금액"), header=True)
        _hairline(self.table, ROW_DIV).pack(fill="x")
        d = self.maemae_data.get(mode)
        if not d:
            ctk.CTkLabel(self.table, text="집계 실행 후 결과가 여기에 표시됩니다.", font=self.f(13), text_color=OPT).pack(pady=24)
            return
        for name, n, amt in d["rows"]:
            self._table_row(self.table, (name, format(n, ","), format(int(amt), ",")))
            _hairline(self.table, ROW_DIV).pack(fill="x")
        self._table_row(self.table, ("합계", format(sum(n for _, n, _ in d["rows"]), ","), format(int(d["total"]), ",")),
                        bold=True, bg=TOTAL_BG)
        _hairline(self.table, ROW_DIV).pack(fill="x")

    def _open_maemae_excel(self):
        d = self.maemae_data.get(self.maemae_mode)
        if d and os.path.exists(d["path"]):
            try:
                os.startfile(d["path"])
            except OSError as ex:
                self.log("엑셀 열기 실패: " + str(ex))

    def _auto_split_file(self):
        """③ 진입 시 방금 변환한 결과가 있으면 자동으로 넣어줌."""
        if self.conv_file or not (self.output_file and os.path.exists(self.output_file)):
            return
        self.conv_file = self.output_file
        self._set_file("conv", os.path.basename(self.output_file), "방금 변환한 결과 ✓")
        self._refresh_split()

    def _open_sender_menu(self):
        b = self.sender_box
        x, y = b.winfo_rootx(), b.winfo_rooty() + b.winfo_height() + 2
        try:
            from customtkinter.windows.widgets.dropdown_menu import DropdownMenu
            if not hasattr(self, "_sender_menu"):
                self._sender_menu = DropdownMenu(master=self, values=SENDERS, command=self._pick_sender,
                                                 fg_color=CARD, hover_color=BG, text_color=INK, font=self.f(14))
            self._sender_menu.open(x, y)
        except Exception:
            m = tk.Menu(self, tearoff=0, font=self.f(13))
            for v in SENDERS:
                m.add_command(label=v, command=lambda v=v: self._pick_sender(v))
            m.tk_popup(x, y)

    def _pick_sender(self, v):
        self.발송인_var.set(v)
        self.sender_lbl.configure(text=v)

    def run_split(self):
        if self._busy or not self.conv_file:
            return
        if not self._require_folder():
            return
        발화 = self.발송인_var.get()
        self._busy = True
        self._log_target = "split"
        self.btn_split.set_enabled(False)
        self.btn_split.set_text("분리 중…")
        self.log("─" * 30)
        self.log("발주서 분리를 시작합니다... (발송인: %s)" % (발화 or "(빈칸)"))
        threading.Thread(target=self._split_worker, args=(발화,), daemon=True).start()

    def _show_split(self, res):
        for w in self.split_rows.winfo_children():
            w.destroy()
        for i, (택배사, n, p) in enumerate(res):
            self._result_row(self.split_rows, CHART[i % len(CHART)], 택배사, "%d건" % n, last=(i == len(res) - 1))
        self.split_count.configure(text="%s건" % format(sum(n for _, n, _ in res), ","))
        self.btn_split_open.pack(fill="x", padx=16, pady=(26, 14))
        self.logs["split"]["status"].configure(text="● 완료", text_color=BRAND)

    def _style_conv_btn(self, enabled):
        b = self.btn_conv_w
        if not enabled:
            b.configure(fg_color=BRAND_DIS, hover_color=BRAND_DIS, text_color="white", border_width=0,
                        image=self.icon("convert", "white", 20))
        elif self.output_file:      # 다시 변환 = 외곽선 버튼 (PDF 2p)
            b.configure(fg_color=CARD, hover_color=BG, text_color=BRAND, border_width=2, border_color=BRAND,
                        image=self.icon("convert", BRAND, 20))
        else:
            b.configure(fg_color=BRAND, hover_color=BRAND_HOVER, text_color="white", border_width=0,
                        image=self.icon("convert", "white", 20))

    def run_convert(self):
        self._t0 = time.time()
        self._log_target = "conv"
        self.warn_box.pack_forget()
        self.act_bar.pack_forget()
        G0.ConverterApp.run_convert(self)

    def _update_folderbar(self):
        b = self.config.get('저장폴더')
        ok = bool(b) and os.path.isdir(b)
        shown = (b if len(b) <= 30 else "…" + b[-28:]) if ok else "저장 폴더를 지정하세요"
        for lbl in self.folder_lbls:
            lbl.configure(text=shown, text_color=INK if ok else REQ)
        for c in self.save_cards:
            c.configure(border_color=CARD_BD)

    def _blink_folderbar(self, n=8):
        for c in self.save_cards:
            c.configure(border_color=REQ)
        self.after(1500, self._stop_blink)

    def _stop_blink(self):
        for c in self.save_cards:
            c.configure(border_color=CARD_BD)

    def _drain_log(self):
        try:
            while True:
                self._append_log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(100, self._drain_log)

    def _append_log(self, msg):
        m = str(msg)
        L = self.logs.get(self._log_target or self._cur) or self.logs["conv"]
        tb, empty, status, h = L["tb"], L["empty"], L["status"], L["height"]
        if empty.winfo_manager() and (self._busy or m.startswith(("─", "⚠", "❌"))):
            empty.pack_forget()
            tb.pack(fill="x" if h else "both", expand=not h, padx=12, pady=(6, 12))
            status.configure(text="● 진행 중", text_color=BRAND)
        if m.strip("─═ ") == "":      # 옛 UI용 구분선은 새 로그에선 생략(줄바꿈으로 어색하게 감김)
            return
        tb.configure(state="normal")
        tb.insert("end", datetime.datetime.now().strftime("%H:%M:%S") + "   ", "ts")
        tb.insert("end", m.strip() + "\n", ("warn",) if m.lstrip().startswith(("⚠", "❌")) else ())
        tb.see("end")
        tb.configure(state="disabled")
        if m.startswith("❌"):
            status.configure(text="● 오류", text_color=REQ)

    def _show_result(self, 총, 정상, 이슈):
        self.last_issues = 이슈
        self._set_tile("total", format(총, ","))
        self._set_tile("good", format(정상, ","), "good")
        self._set_tile("issue", str(len(이슈)), "warn" if 이슈 else "")
        self.log_status.configure(text="● 완료", text_color=BRAND)
        now = datetime.datetime.now()
        self.conv_caption.configure(text="방금 변환: %s %d:%02d · %.1f초 소요" % (
            "오전" if now.hour < 12 else "오후", now.hour % 12 or 12, now.minute, time.time() - getattr(self, "_t0", time.time())))
        self.warn_box.pack_forget()
        self.act_bar.pack_forget()
        if 이슈:
            self.warn_title.configure(text="⚠  확인이 필요한 %d건" % len(이슈))
            self.warn_btn.configure(text="%d건 목록 보기" % len(이슈))
            self.warn_box.pack(fill="x", padx=14, pady=(0, 12))
        self.act_bar.pack(fill="x", padx=14, pady=(0, 14))

    def _show_issue_list(self):
        win = ctk.CTkToplevel(self)
        win.title("확인이 필요한 항목")
        win.geometry("700x540")
        win.configure(fg_color=BG)
        card = self._card(win)
        card.pack(fill="both", expand=True, padx=16, pady=16)
        self._card_header(card, "확인이 필요한 %d건" % len(self.last_issues), "행 번호는 엑셀 기준", FAINT)
        lines = []
        for it in self.last_issues:
            행 = str(it.get('행') or '-')
            상품 = str(it.get('상품명') or '').strip()[:40]
            lines.append((("%s행" % 행).rjust(6), "%s\n         문제: %s\n         해결: %s" % (상품, it.get('문제', ''), it.get('해결', ''))))
        self._log_box(card, lines)
        win.lift()

    def open_help(self):
        """설명서 창 (새 디자인: 카드형 + 스크롤). 내용은 HELP_SECTIONS + 결과 색상 범례."""
        win = ctk.CTkToplevel(self)
        win.title("설명서")
        win.geometry("580x740")
        win.minsize(500, 560)
        win.configure(fg_color=BG)
        try:
            win.after(200, lambda: win.iconbitmap(os.path.join(G0.ui_dir(), "appicon.ico")))
        except Exception:
            pass
        sf = ctk.CTkScrollableFrame(win, fg_color="transparent", corner_radius=0,
                                    scrollbar_button_color=CTRL_BD, scrollbar_button_hover_color=FAINT)
        sf.pack(fill="both", expand=True, padx=12, pady=12)
        ctk.CTkLabel(sf, text="설명서", font=self.f(22, "bold"), text_color=INK).pack(anchor="w", padx=6, pady=(4, 0))
        ctk.CTkLabel(sf, text="주문서 변환기 사용법과 결과 색상 안내 · 자세한 내용은 [매뉴얼]을 열어 주세요.",
                     font=self.f(13), text_color=SUB, wraplength=480, justify="left").pack(anchor="w", padx=6, pady=(2, 10))
        for title, text in HELP_SECTIONS:
            c = self._card(sf)
            c.pack(fill="x", padx=2, pady=(0, 10))
            ctk.CTkLabel(c, text=title, font=self.f(14, "bold"), text_color=INK).pack(anchor="w", padx=16, pady=(12, 4))
            ctk.CTkLabel(c, text=text, font=self.f(13, plus=False), text_color="#3C4757", justify="left", anchor="w",
                         wraplength=470).pack(anchor="w", padx=16, pady=(0, 12))
        c = self._card(sf)
        c.pack(fill="x", padx=2, pady=(0, 10))
        ctk.CTkLabel(c, text="결과 파일 색상", font=self.f(14, "bold"), text_color=INK).pack(anchor="w", padx=16, pady=(12, 6))
        for hexc, nm, desc in G0.COLOR_LEGEND:
            r = ctk.CTkFrame(c, fg_color="transparent")
            r.pack(fill="x", padx=16, pady=3)
            ctk.CTkFrame(r, width=22, height=16, corner_radius=3, fg_color=hexc,
                         border_width=1, border_color=CTRL_BD).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(r, text=nm, font=self.f(13, "bold"), text_color=INK, width=56, anchor="w").pack(side="left")
            ctk.CTkLabel(r, text=desc, font=self.f(13, plus=False), text_color="#3C4757", anchor="w", justify="left",
                         wraplength=330).pack(side="left")
        ctk.CTkFrame(c, fg_color="transparent", height=8).pack()
        win.lift()

    def _open_excel(self):
        if self.output_file and os.path.exists(self.output_file):
            try:
                os.startfile(self.output_file)
            except OSError as ex:
                self.log("엑셀 열기 실패: " + str(ex))

    def _setver(self, text, color=None):
        if "최신" in text:
            t, c = "✓ 최신", ACCENT
        elif "있음" in text:
            t, c = text, "#FFB84D"
        else:
            t, c = text, FAINT
        self.ver_lbl.configure(text=t, text_color=c)


# gui.py(ConverterApp)의 로직 메서드를 그대로 빌려온다 (self.root / self.log / 상태값만 사용하는 것들)
for _n in ("_config_path", "_load_config", "_save_config", "pick_folder", "_base_ok", "_require_folder",
           "_session_dir", "_new_session_out", "_sess_out", "log", "_warn_if_result",
           "pick_input", "pick_mapping", "_auto_sus", "pick_sus", "_worker", "_리포트출력", "open_result_folder",
           "pick_conv", "_form_dirs", "_split_worker", "open_split_folder",
           "pick_maemae", "_auto_maemae_file", "_latest_변환결과", "open_maemae_folder",
           "_manual_path", "open_manual",
           "_check_update", "_update_popup", "_set_status", "_manual_fallback", "_do_update", "_quit_for_update",
           "_launch_replace_bat", "_apply_taskbar_icon", "_bring_to_front", "_show_update_banner_if_needed"):
    setattr(App, _n, getattr(G0.ConverterApp, _n))


def main():
    try:    # 작업표시줄이 이 앱 고유 아이콘을 쓰도록
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("JJCompany.OrderConverter")
    except Exception:
        pass
    load_fonts()
    app = App()
    if os.environ.get("CTK_SMOKE") == "1":
        app.after(3500, app.destroy)
    app.mainloop()


if __name__ == "__main__":
    main()
