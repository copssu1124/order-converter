# -*- coding: utf-8 -*-
"""주문서 변환기 GUI — step3_convert 의 convert()/분석함수를 호출 (변환 로직 미수정)."""
import os
import re
import sys
import json
import queue
import threading
import traceback
import tempfile
import subprocess
import webbrowser
import urllib.request
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import tkinter.font as tkfont

import step3_convert as engine

VERSION = "3.3"                 # ★ 버전은 이 한 곳에서만 관리
KAKAO = "https://open.kakao.com/o/gyxhX4zi"
CREDIT = "Developed by JANG JUNG WOO · JJ COMPANY"
GITHUB_REPO = "copssu1124/order-converter"
RELEASES_URL = "https://github.com/copssu1124/order-converter/releases/latest"

_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def _번호(n):
    return _CIRCLED[n - 1] if 1 <= n <= len(_CIRCLED) else "(%d)" % n


def _버전튜플(s):
    """'v2.7' / '2.7' → (2, 7). 실패 시 None."""
    try:
        s = str(s).strip().lstrip("vV")
        parts = re.findall(r"\d+", s)
        if not parts:
            return None
        return tuple(int(x) for x in parts[:2])
    except Exception:
        return None


def _검증_다운로드(path, expected_size):
    """받은 업데이트 파일이 정상 exe인지 검증 → (ok, 사유).
       ① Content-Length 일치(부분 다운로드 차단) ② 최소 1MB ③ PE 시그니처(MZ)."""
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


GUIDE = (
    "📦 주문서 변환기 사용법\n"
    "──────────────────────\n"
    "① [주문서 불러오기]\n"
    "    변환할 .xls 주문서를 고릅니다.\n\n"
    "② [매핑표 불러오기]\n"
    "    같은 폴더에 있으면 자동으로 잡힙니다.\n"
    "    없거나 바꾸려면 직접 선택하세요.\n\n"
    "③ [변환 실행]\n"
    "    결과 파일이 주문서와 같은 폴더에\n"
    "    새로 만들어집니다.\n"
    "    (원본·매핑표는 절대 수정 안 함)\n\n"
    "④ [결과 폴더 열기]\n"
    "    변환이 끝나면 결과 폴더를 엽니다.\n\n"
    "🎨 결과 파일 색상 의미\n"
    "──────────────────────\n"
    "[택배사(L)열]\n"
    "  · 파랑 = 일반 택배사(천일·씨제이 등)\n"
    "  · 주황 = 다모아\n"
    "  · 빨강 = 수동확인 필요\n"
    "[배송비(Q)열]\n"
    "  · 초록 = 배송비 있음\n"
    "  · 노랑 = 빈칸(확인 필요)\n"
    "[배송비합계(R)열]\n"
    "  · 연빨강 = 합계 / 진빨강 = 단가×수량 곱셈한 합계\n"
    "[택배사(L)·택배등급(S)열]\n"
    "  · 연보라 = 수정택배비 사용(택배사 '수정' · 검토 필요)\n"
    "[주소]\n"
    "  · 연한 빨강 = 주소 중복(같은 주소 2건 이상)\n\n"
    "📋 결과 '비고'열\n"
    "──────────────────────\n"
    "수동확인(빨강) 행에만 사유가 적힙니다.\n"
    "  · 상품명 미등록\n"
    "  · 택배비 미등록(회사상품명*수량)\n\n"
    "📑 주문서 규칙\n"
    "──────────────────────\n"
    "· 수령자·상품명이 빈 행에서 멈춤\n"
    "· 옵션(선택사항) 우선, 없으면 상품명 사용\n"
    "· 송장·배송번호 등 빈칸은 그대로 보존\n\n"
    "🗂 매핑표 규칙\n"
    "──────────────────────\n"
    "· 시트 4개: 상품명변경/판매비변경/\n"
    "  천일택배비/씨제이택배비\n"
    "· 상품명변경: 옵션명 → 회사상품명\n"
    "· 판매비변경: 회사상품명*수량 → 배송비\n"
)


GUIDE2 = (
    "🚚 택배사 분리 사용법\n"
    "──────────────────────\n"
    "① [변환완료 파일 불러오기]\n"
    "    ① 변환 탭에서 만든 결과 파일을\n"
    "    엑셀에서 검증·빈칸 보완한 뒤 고릅니다.\n\n"
    "② [발화주명] 입력\n"
    "    보내는 회사명을 적습니다.\n"
    "    (대신 발주서 '발화주명' 칸에 들어감)\n\n"
    "③ [택배사 분리 실행]\n"
    "    택배사별로 나눠 각 양식에 채운\n"
    "    파일들을 '분리출력' 폴더에 만듭니다.\n\n"
    "④ [분리 폴더 열기]\n"
    "    분리 결과 폴더를 엽니다.\n\n"
    "📦 택배사 → 양식\n"
    "──────────────────────\n"
    "· 씨제이 → CJ 양식\n"
    "· 대신·대신낱개 → 대신 양식\n"
    "· 대신택배 → 대신택배(별도 파일)\n"
    "· 로젠 → 로젠 양식\n"
    "· 천일 → 천일 양식\n"
    "· 원준 → 원준 양식 (L열 총운임)\n"
    "· 위플 → 위플 양식 (F열 총운임)\n"
    "· 로젠 → 로젠 양식 (G열 총운임)\n"
    "· 올담 → 올담 양식\n"
    "· 카몬드 → 카몬드 양식\n"
    "· 양식 없음 → '기타' 파일로 모음\n"
    "· 대신 파일: 연두칸 = 대신낱개 행\n\n"
    "⚠ '출력양식' 폴더가 프로그램과\n"
    "   같은(또는 상위) 폴더에 있어야 합니다.\n"
)


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.input_file = None
        self.mapping_file = None
        self.output_file = None
        self.conv_file = None       # 발주서 분리 대상(변환완료 파일)
        self.split_dir = None
        self.log_queue = queue.Queue()

        # ── 전체 UI 폰트: 맑은 고딕으로 통일 + 크게 ──
        for _nm, _sz in (("TkDefaultFont", 11), ("TkTextFont", 11),
                         ("TkMenuFont", 11), ("TkHeadingFont", 12),
                         ("TkFixedFont", 11), ("TkIconFont", 11),
                         ("TkTooltipFont", 11)):
            try:
                tkfont.nametofont(_nm).configure(family="맑은 고딕", size=_sz)
            except Exception:
                pass
        try:
            _style = ttk.Style()
            _style.configure("TNotebook.Tab", font=("맑은 고딕", 13, "bold"),
                             padding=(18, 9))
        except Exception:
            pass

        root.title("주문서 변환기 · JJ COMPANY   v" + VERSION)
        root.geometry("1040x820")
        root.minsize(940, 720)

        root.rowconfigure(0, weight=3)   # 탭(노트북)
        root.rowconfigure(1, weight=3)   # 공용 로그(크게)
        root.rowconfigure(2, weight=0)   # 하단
        root.columnconfigure(0, weight=1)

        nb = ttk.Notebook(root)
        nb.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 4))

        # ═══════════ 탭 ①  주문서 변환 ═══════════
        tab1 = tk.Frame(nb)
        nb.add(tab1, text="    ①  주문서 변환    ")
        tab1.rowconfigure(0, weight=1)
        tab1.columnconfigure(0, weight=1)   # 설명서
        tab1.columnconfigure(1, weight=2)   # 작업

        left1 = tk.LabelFrame(tab1, text="📖 설명서", font=("맑은 고딕", 12, "bold"))
        left1.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=4)
        left1.rowconfigure(0, weight=1)
        left1.columnconfigure(0, weight=1)
        guide = scrolledtext.ScrolledText(left1, width=34, wrap="word",
                                          font=("맑은 고딕", 11), bg="#FAFAFA")
        guide.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        guide.insert("1.0", GUIDE)
        guide.config(state="disabled")

        btns = tk.Frame(tab1)
        btns.grid(row=0, column=1, sticky="new", pady=4)
        btns.columnconfigure(0, weight=1)

        self.lbl_input = tk.Label(btns, text="주문서: (선택 안 됨)", anchor="w", fg="#555")
        self.lbl_mapping = tk.Label(btns, text="매핑표: (자동 탐색)", anchor="w", fg="#555")

        tk.Button(btns, text="①  주문서 불러오기", height=2,
                  command=self.pick_input).grid(row=0, column=0, sticky="ew", pady=2)
        self.lbl_input.grid(row=1, column=0, sticky="ew", padx=2)
        tk.Button(btns, text="②  매핑표 불러오기", height=2,
                  command=self.pick_mapping).grid(row=2, column=0, sticky="ew", pady=2)
        self.lbl_mapping.grid(row=3, column=0, sticky="ew", padx=2)

        self.btn_run = tk.Button(btns, text="③  변환 실행", height=2,
                                 bg="#2E7D32", fg="white", font=("맑은 고딕", 13, "bold"),
                                 command=self.run_convert)
        self.btn_run.grid(row=4, column=0, sticky="ew", pady=(8, 2))
        self.btn_open = tk.Button(btns, text="④  결과 폴더 열기", height=2,
                                  state="disabled", command=self.open_result_folder)
        self.btn_open.grid(row=5, column=0, sticky="ew", pady=2)

        # ═══════════ 탭 ②  택배사 분리 ═══════════
        tab2 = tk.Frame(nb)
        nb.add(tab2, text="    ②  택배사 분리    ")
        tab2.rowconfigure(0, weight=1)
        tab2.columnconfigure(0, weight=1)
        tab2.columnconfigure(1, weight=2)

        left2 = tk.LabelFrame(tab2, text="📖 택배사 분리 안내", font=("맑은 고딕", 12, "bold"))
        left2.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=4)
        left2.rowconfigure(0, weight=1)
        left2.columnconfigure(0, weight=1)
        guide2 = scrolledtext.ScrolledText(left2, width=34, wrap="word",
                                           font=("맑은 고딕", 11), bg="#FAFAFA")
        guide2.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        guide2.insert("1.0", GUIDE2)
        guide2.config(state="disabled")

        sp = tk.Frame(tab2)
        sp.grid(row=0, column=1, sticky="new", pady=4)
        sp.columnconfigure(0, weight=1)

        tk.Button(sp, text="①  변환완료 파일 불러오기", height=2,
                  bg="#E3F2FD", font=("맑은 고딕", 13, "bold"),
                  command=self.pick_conv).grid(row=0, column=0, sticky="ew", pady=2)
        self.lbl_conv = tk.Label(sp, text="분리할 파일: (선택 안 됨)", anchor="w", fg="#555")
        self.lbl_conv.grid(row=1, column=0, sticky="ew", padx=2)

        tk.Label(sp, text="②  발화주명 (보내는 회사)", anchor="w",
                 font=("맑은 고딕", 12, "bold")).grid(row=2, column=0, sticky="ew",
                                                      padx=2, pady=(10, 0))
        self.entry_발화 = tk.Entry(sp, font=("맑은 고딕", 12))
        self.entry_발화.grid(row=3, column=0, sticky="ew", padx=2, pady=(2, 4))

        self.btn_split = tk.Button(sp, text="③  택배사 분리 실행", height=2,
                                   bg="#1565C0", fg="white", font=("맑은 고딕", 13, "bold"),
                                   command=self.run_split)
        self.btn_split.grid(row=4, column=0, sticky="ew", pady=(8, 2))
        self.btn_split_open = tk.Button(sp, text="④  분리 폴더 열기", height=2,
                                        state="disabled", command=self.open_split_folder)
        self.btn_split_open.grid(row=5, column=0, sticky="ew", pady=2)

        # ═══════════ 공용 로그 (크게·또렷) ═══════════
        logframe = tk.LabelFrame(root, text="📂 진행 상황 · 오류 리포트",
                                 font=("맑은 고딕", 13, "bold"))
        logframe.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 4))
        logframe.rowconfigure(0, weight=1)
        logframe.columnconfigure(0, weight=1)
        self.logbox = scrolledtext.ScrolledText(logframe, font=("맑은 고딕", 12),
                                                state="disabled", wrap="word",
                                                height=16, spacing1=1, spacing3=1)
        self.logbox.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        # ═══════════ 하단: 제작자 + 카카오톡 ═══════════
        bottom = tk.Frame(root, bg="#ECEFF1")
        bottom.grid(row=2, column=0, sticky="ew")
        tk.Label(bottom, text=CREDIT, bg="#ECEFF1", fg="#37474F",
                 font=("맑은 고딕", 10)).pack(side="left", padx=12, pady=6)
        link = tk.Label(bottom, text="💬 카카오톡 문의", bg="#ECEFF1", fg="#1565C0",
                        cursor="hand2", font=("맑은 고딕", 10, "underline"))
        link.pack(side="right", padx=12, pady=6)
        link.bind("<Button-1>", lambda e: webbrowser.open(KAKAO))

        auto_map = engine.find_mapping_file(app_dir())
        if auto_map:
            self.mapping_file = auto_map
            self.lbl_mapping.config(text="매핑표: " + os.path.basename(auto_map) + " (자동)", fg="#1565C0")

        self.log("주문서를 불러오고 [변환 실행]을 누르세요.")
        self._drain_log()
        self._start_update_check()       # 백그라운드 최신버전 확인(알림형)

    # ── 로그 ──
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

    # ── 파일 선택 ──
    def pick_input(self):
        start = os.path.dirname(self.input_file) if self.input_file else app_dir()
        path = filedialog.askopenfilename(
            title="주문서(.xls) 선택", initialdir=start,
            filetypes=[("엑셀 주문서", "*.xls *.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self.input_file = path
        self.lbl_input.config(text="주문서: " + os.path.basename(path), fg="#1565C0")
        self.log("주문서 선택: " + os.path.basename(path))
        if not self.mapping_file:
            am = engine.find_mapping_file(os.path.dirname(path))
            if am:
                self.mapping_file = am
                self.lbl_mapping.config(text="매핑표: " + os.path.basename(am) + " (자동)", fg="#1565C0")
                self.log("매핑표 자동 탐색: " + os.path.basename(am))

    def pick_mapping(self):
        start = os.path.dirname(self.mapping_file) if self.mapping_file else app_dir()
        path = filedialog.askopenfilename(
            title="매핑표(.xlsx) 선택", initialdir=start,
            filetypes=[("엑셀 매핑표", "*.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self.mapping_file = path
        self.lbl_mapping.config(text="매핑표: " + os.path.basename(path), fg="#1565C0")
        self.log("매핑표 선택: " + os.path.basename(path))

    # ── 변환 ──
    def run_convert(self):
        if not self.input_file:
            self.log("[오류] 먼저 주문서를 불러오세요.")
            return
        if not self.mapping_file:
            self.log("[오류] 매핑표를 찾을 수 없습니다. [매핑표 불러오기]로 지정하세요.")
            return
        self.btn_run.config(state="disabled", text="변환 중...")
        self.btn_open.config(state="disabled")
        self.log("─" * 38)
        self.log("변환을 시작합니다...")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            # 출력 파일명 = 입력 파일명 (.xlsx), 결과 폴더에 저장(원본 미덮어쓰기)
            결과폴더 = os.path.join(app_dir(), "변환결과")
            os.makedirs(결과폴더, exist_ok=True)
            base = os.path.splitext(os.path.basename(self.input_file))[0]
            out_path = os.path.join(결과폴더, base + ".xlsx")
            out, 사유 = engine.convert_v2(self.input_file, self.mapping_file,
                                          out_path, log=self.log)
            self.output_file = out
            총, 정상, 이슈 = engine.상세리포트(out, 사유)
            self._리포트출력(총, 정상, 이슈)
            self.log("")
            self.log("💾 저장 완료: " + out)
            self.root.after(0, lambda: self.btn_open.config(state="normal"))
        except Exception as ex:
            self.log("─" * 38)
            self.log("❌ 오류: " + str(ex))
            self.log(traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.btn_run.config(state="normal", text="③  변환 실행"))

    def open_result_folder(self):
        if not self.output_file or not os.path.exists(self.output_file):
            self.log("[오류] 결과 파일이 없습니다.")
            return
        try:
            os.startfile(os.path.dirname(self.output_file))
        except Exception as ex:
            self.log("폴더 열기 실패: " + str(ex))

    # ── 변환 결과 리포트 (친절한 안내) ──
    def _리포트출력(self, 총, 정상, 이슈):
        self.log("")
        self.log("════════════════════════════")
        self.log("   변환 완료!     총 %d건" % 총)
        self.log("   ✅ 잘 된 것 : %d건" % 정상)
        self.log("   ⚠️ 확인이 필요한 것 : %d건" % len(이슈))
        self.log("════════════════════════════")
        if not 이슈:
            self.log("")
            self.log("✅ 모두 정상으로 변환됐어요!")
            return
        self.log("")
        self.log("[확인이 필요한 항목]")
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
            self.log("")

    # ── 자동 업데이트 (원클릭 반자동: 다운로드 → 교체 → 재시작) ──
    def _start_update_check(self):
        threading.Thread(target=self._check_update, daemon=True).start()

    def _check_update(self):
        try:
            url = "https://api.github.com/repos/%s/releases/latest" % GITHUB_REPO
            req = urllib.request.Request(url, headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": "order-converter"})
            with urllib.request.urlopen(req, timeout=3) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            최신 = _버전튜플(str(data.get("tag_name", "")))
            현재 = _버전튜플(VERSION)
            dl = None
            assets = data.get("assets") or []
            for a in assets:                       # .exe 자산의 직접 다운로드 주소
                if str(a.get("name", "")).lower().endswith(".exe"):
                    dl = a.get("browser_download_url"); break
            if not dl and assets:
                dl = assets[0].get("browser_download_url")
            if 최신 and 현재 and 최신 > 현재:
                disp = "%d.%d" % 최신
                self.root.after(0, lambda: self._update_popup(disp, dl))
        except Exception:
            pass    # 네트워크 불가/타임아웃/404(Releases 없음) → 조용히 무시

    def _update_popup(self, new_ver, download_url):
        try:
            win = tk.Toplevel(self.root)
            win.title("업데이트")
            win.transient(self.root)
            win.resizable(False, False)
            tk.Label(win, text="🎉 새 버전 v%s 이(가) 나왔어요!" % new_ver,
                     font=("맑은 고딕", 14, "bold")).pack(padx=40, pady=(22, 4))
            tk.Label(win, text="지금 업데이트할까요?",
                     font=("맑은 고딕", 12)).pack(pady=(0, 10))
            status = tk.Label(win, text="", font=("맑은 고딕", 11), fg="#1565C0")
            status.pack(pady=(0, 6))
            bar = tk.Frame(win)
            bar.pack(pady=(0, 18))

            btn_later = tk.Button(bar, text="나중에", width=10, height=1,
                                  font=("맑은 고딕", 12), command=win.destroy)

            def 지금():
                btn_now.config(state="disabled")
                btn_later.config(state="disabled")
                threading.Thread(target=self._do_update,
                                 args=(download_url, win, status, btn_later),
                                 daemon=True).start()

            btn_now = tk.Button(bar, text="지금 업데이트", width=12, height=1,
                                bg="#2E7D32", fg="white", font=("맑은 고딕", 12, "bold"),
                                command=지금)
            btn_now.pack(side="left", padx=8)
            btn_later.pack(side="left", padx=8)
            win.update_idletasks()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            ww, wh = win.winfo_width(), win.winfo_height()
            win.geometry("+%d+%d" % (rx + (rw - ww) // 2, ry + (rh - wh) // 3))
            win.grab_set()
        except Exception:
            pass

    def _set_status(self, status, text, color="#1565C0"):
        self.root.after(0, lambda: status.config(text=text, fg=color))

    def _manual_fallback(self, win, status, msg, btn_later=None):
        def show():
            status.config(text=msg, fg="#C62828")
            link = tk.Label(win, text="👉 수동 다운로드 (Releases 페이지 열기)",
                            fg="#1565C0", cursor="hand2",
                            font=("맑은 고딕", 11, "underline"))
            link.pack(pady=(2, 14))
            link.bind("<Button-1>", lambda e: webbrowser.open(RELEASES_URL))
            if btn_later is not None:
                btn_later.config(state="normal")
        self.root.after(0, show)

    def _do_update(self, download_url, win, status, btn_later):
        # 1) exe(frozen)일 때만 자가교체 가능
        if not getattr(sys, "frozen", False):
            self._manual_fallback(win, status,
                "개발 모드(.py 실행)에선 자동 교체가 안 돼요. 수동으로 받아주세요.", btn_later)
            return
        if not download_url:
            self._manual_fallback(win, status,
                "다운로드 주소를 찾지 못했어요. 수동으로 받아주세요.", btn_later)
            return
        target = sys.executable
        # 2) 교체 권한(쓰기) 확인 — 관리자 폴더(Program Files 등)면 폴백
        try:
            _t = os.path.join(os.path.dirname(target), ".__upd_test__")
            with open(_t, "w") as f:
                f.write("x")
            os.remove(_t)
        except Exception:
            self._manual_fallback(win, status,
                "이 위치는 자동 교체 권한이 없어요(관리자 폴더). 수동으로 받아주세요.", btn_later)
            return
        # 3) 다운로드 (.part로 받아 검증 통과분만 최종 파일로 확정 — 손상/부분 파일 차단)
        tmp = os.path.join(tempfile.gettempdir(), "OrderConverter_update.exe")
        part = tmp + ".part"
        try:
            self._set_status(status, "다운로드 준비 중...")
            req = urllib.request.Request(download_url,
                                         headers={"User-Agent": "order-converter"})
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
            ok, 사유 = _검증_다운로드(part, total)    # 길이일치·최소크기·MZ 시그니처
            if not ok:
                raise RuntimeError(사유)
            os.replace(part, tmp)                      # 검증 통과분만 최종 파일로 확정
            # 4) 배치로 교체 + 재시작 예약 후 현재 프로그램 종료
            self._set_status(status, "교체 후 자동 재시작합니다...", "#2E7D32")
            self._launch_replace_bat(tmp, target)
            self.root.after(400, self._quit_for_update)
        except Exception as ex:
            try:
                if os.path.exists(part):
                    os.remove(part)
            except OSError:
                pass
            self._manual_fallback(win, status,
                "자동 업데이트 실패: %s" % str(ex)[:60], btn_later)

    def _quit_for_update(self):
        try:
            self.root.destroy()
        finally:
            os._exit(0)

    def _launch_replace_bat(self, tmp_exe, target_exe):
        """실행 중 exe는 잠겨 있어 못 덮어씀 → 배치가 종료 대기 후 교체·재시작.
           안전망: 교체 전 원본을 .bak로 백업하고, 새 버전이 안 뜨면 원본을 복원(브릭 방지)."""
        bat = os.path.join(tempfile.gettempdir(), "OrderConverter_update.bat")
        bak = target_exe + ".bak"
        exe_name = os.path.basename(target_exe)
        lines = [
            "@echo off",
            'copy /y "%s" "%s" >nul 2>&1' % (target_exe, bak),   # 원본 백업(실행중에도 읽기복사 가능)
            'if not exist "%s" goto norepl' % bak,               # 백업 실패(디스크풀 등) → 교체 포기(브릭 방지)
            "set /a n=0",
            ":retry",
            "ping -n 2 127.0.0.1 >nul",                          # 약 1초 대기(=종료 대기)
            'move /y "%s" "%s" >nul 2>&1' % (tmp_exe, target_exe),
            "if not errorlevel 1 goto launch",
            "set /a n+=1",
            "if %n% lss 40 goto retry",                          # 최대 ~80초 재시도(무한루프 방지)
            'start "" "%s"' % target_exe,                        # 교체 실패 → 원본(미교체) 그대로 실행
            "goto cleanup",
            ":norepl",                                           # 백업 못 만듦 → 원본 그대로 실행(미교체)
            'start "" "%s"' % target_exe,
            "goto cleanup",
            ":launch",
            "ping -n 9 127.0.0.1 >nul",                          # 교체 직후 ~8초 대기(백신 검사 완료) → 첫 실행 DLL오류 완화
            'start "" "%s"' % target_exe,                        # 새 버전 실행
            "set /a m=0",
            ":check",                                            # 새 버전 기동 확인(최대 ~10초)
            "ping -n 3 127.0.0.1 >nul",
            'tasklist /fi "imagename eq %s" 2>nul | find /i "%s" >nul' % (exe_name, exe_name),
            "if not errorlevel 1 goto cleanup",
            "set /a m+=1",
            "if %m% lss 4 goto check",
            'if exist "%s" move /y "%s" "%s" >nul 2>&1' % (bak, bak, target_exe),  # 기동 실패 → 롤백
            'if exist "%s" start "" "%s"' % (target_exe, target_exe),
            ":cleanup",
            'if exist "%s" del "%s" >nul 2>&1' % (bak, bak),     # 백업 정리
            'del "%~f0"',                                        # 배치 자기삭제
        ]
        # 한글 경로를 cmd가 그대로 읽도록 시스템 ANSI 코드페이지(mbcs)로 기록
        with open(bat, "w", encoding="mbcs", errors="replace") as f:
            f.write("\r\n".join(lines) + "\r\n")
        DETACHED = 0x00000008                        # DETACHED_PROCESS (부모 종료 후에도 생존)
        subprocess.Popen(["cmd", "/c", bat], creationflags=DETACHED, close_fds=True)

    # ── ② 발주서 분리 ──
    def pick_conv(self):
        start = os.path.join(app_dir(), "변환결과")
        if not os.path.isdir(start):
            start = app_dir()
        path = filedialog.askopenfilename(
            title="변환완료 파일(.xlsx) 선택", initialdir=start,
            filetypes=[("엑셀", "*.xlsx"), ("모든 파일", "*.*")])
        if not path:
            return
        self.conv_file = path
        self.lbl_conv.config(text=os.path.basename(path), fg="#1565C0")
        self.log("분리 대상: " + os.path.basename(path))

    def run_split(self):
        if not self.conv_file:
            self.log("[오류] 먼저 [변환완료 파일 불러오기]로 검증된 파일을 고르세요.")
            return
        발화 = self.entry_발화.get().strip()
        self.btn_split.config(state="disabled", text="분리 중...")
        self.btn_split_open.config(state="disabled")
        self.log("─" * 38)
        self.log("발주서 분리를 시작합니다... (발화주명: %s)" % (발화 or "(빈칸)"))
        threading.Thread(target=self._split_worker, args=(발화,), daemon=True).start()

    def _form_dirs(self):
        """출력양식 후보 폴더(앞쪽 우선): exe 옆 → 상위 폴더 → exe 번들(_MEIPASS)."""
        dirs = [os.path.join(app_dir(), "출력양식"),
                os.path.join(os.path.dirname(app_dir()), "출력양식")]
        base = getattr(sys, "_MEIPASS", None)      # 번들된 출력양식(파일 누락 시 보충)
        if base:
            dirs.append(os.path.join(base, "출력양식"))
        return dirs

    def _split_worker(self, 발화):
        try:
            out_dir = os.path.join(app_dir(), "분리출력")
            res = engine.발주서분리(self.conv_file, 발화, self._form_dirs(),
                                    out_dir, log=self.log)
            self.split_dir = out_dir
            self.log("─" * 38)
            for 택배사, n, p in res:
                self.log("  %s : %d건 → %s" % (택배사, n, os.path.basename(p)))
            self.log("✅ 분리 완료: " + out_dir)
            self.root.after(0, lambda: self.btn_split_open.config(state="normal"))
        except engine.양식파일없음 as ex:          # 양식 없음 → 친절한 안내(트레이스백 없이)
            self.log("─" * 38)
            self.log("⚠️ " + str(ex))
        except Exception as ex:
            self.log("─" * 38)
            self.log("❌ 분리 오류: " + str(ex))
            self.log(traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.btn_split.config(state="normal", text="발주서 분리 실행"))

    def open_split_folder(self):
        if self.split_dir and os.path.exists(self.split_dir):
            try:
                os.startfile(self.split_dir)
            except Exception as ex:
                self.log("폴더 열기 실패: " + str(ex))
        else:
            self.log("[오류] 분리출력 폴더가 없습니다.")


def main():
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
