# -*- coding: utf-8 -*-
"""주문서 변환기 GUI — step3_convert 의 convert()/분석함수를 호출 (변환 로직 미수정)."""
import os
import sys
import queue
import threading
import traceback
import webbrowser
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

import step3_convert as engine

KAKAO = "https://open.kakao.com/o/gyxhX4zi"
CREDIT = "Developed by Jeong-woo Jang · JJ COMPANY"


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
    "  · 빨강계열 = 합계 (진할수록 묶음 2개↑)\n"
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
    "· 원준 → 원준 양식\n"
    "· 위플 → 위플 양식\n"
    "· 양식 없음 → '기타' 파일로 모음\n\n"
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

        root.title("주문서 변환기 · JJ COMPANY")
        root.geometry("940x660")
        root.minsize(860, 560)

        root.rowconfigure(0, weight=3)   # 탭(노트북)
        root.rowconfigure(1, weight=2)   # 공용 로그
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

        left1 = tk.LabelFrame(tab1, text="📖 설명서", font=("맑은 고딕", 10, "bold"))
        left1.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=4)
        left1.rowconfigure(0, weight=1)
        left1.columnconfigure(0, weight=1)
        guide = scrolledtext.ScrolledText(left1, width=38, wrap="word",
                                          font=("맑은 고딕", 9), bg="#FAFAFA")
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
                                 bg="#2E7D32", fg="white", font=("맑은 고딕", 11, "bold"),
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

        left2 = tk.LabelFrame(tab2, text="📖 택배사 분리 안내", font=("맑은 고딕", 10, "bold"))
        left2.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=4)
        left2.rowconfigure(0, weight=1)
        left2.columnconfigure(0, weight=1)
        guide2 = scrolledtext.ScrolledText(left2, width=38, wrap="word",
                                           font=("맑은 고딕", 9), bg="#FAFAFA")
        guide2.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        guide2.insert("1.0", GUIDE2)
        guide2.config(state="disabled")

        sp = tk.Frame(tab2)
        sp.grid(row=0, column=1, sticky="new", pady=4)
        sp.columnconfigure(0, weight=1)

        tk.Button(sp, text="①  변환완료 파일 불러오기", height=2,
                  bg="#E3F2FD", font=("맑은 고딕", 11, "bold"),
                  command=self.pick_conv).grid(row=0, column=0, sticky="ew", pady=2)
        self.lbl_conv = tk.Label(sp, text="분리할 파일: (선택 안 됨)", anchor="w", fg="#555")
        self.lbl_conv.grid(row=1, column=0, sticky="ew", padx=2)

        tk.Label(sp, text="②  발화주명 (보내는 회사)", anchor="w",
                 font=("맑은 고딕", 10, "bold")).grid(row=2, column=0, sticky="ew",
                                                      padx=2, pady=(10, 0))
        self.entry_발화 = tk.Entry(sp, font=("맑은 고딕", 11))
        self.entry_발화.grid(row=3, column=0, sticky="ew", padx=2, pady=(2, 4))

        self.btn_split = tk.Button(sp, text="③  택배사 분리 실행", height=2,
                                   bg="#1565C0", fg="white", font=("맑은 고딕", 11, "bold"),
                                   command=self.run_split)
        self.btn_split.grid(row=4, column=0, sticky="ew", pady=(8, 2))
        self.btn_split_open = tk.Button(sp, text="④  분리 폴더 열기", height=2,
                                        state="disabled", command=self.open_split_folder)
        self.btn_split_open.grid(row=5, column=0, sticky="ew", pady=2)

        # ═══════════ 공용 로그 ═══════════
        logframe = tk.LabelFrame(root, text="📂 진행 상황 · 오류 리포트",
                                 font=("맑은 고딕", 10, "bold"))
        logframe.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 4))
        logframe.rowconfigure(0, weight=1)
        logframe.columnconfigure(0, weight=1)
        self.logbox = scrolledtext.ScrolledText(logframe, font=("Consolas", 9),
                                                state="disabled", wrap="word", height=8)
        self.logbox.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        # ═══════════ 하단: 제작자 + 카카오톡 ═══════════
        bottom = tk.Frame(root, bg="#ECEFF1")
        bottom.grid(row=2, column=0, sticky="ew")
        tk.Label(bottom, text=CREDIT, bg="#ECEFF1", fg="#37474F",
                 font=("Segoe UI", 9)).pack(side="left", padx=12, pady=6)
        link = tk.Label(bottom, text="💬 카카오톡 문의", bg="#ECEFF1", fg="#1565C0",
                        cursor="hand2", font=("맑은 고딕", 9, "underline"))
        link.pack(side="right", padx=12, pady=6)
        link.bind("<Button-1>", lambda e: webbrowser.open(KAKAO))

        auto_map = engine.find_mapping_file(app_dir())
        if auto_map:
            self.mapping_file = auto_map
            self.lbl_mapping.config(text="매핑표: " + os.path.basename(auto_map) + " (자동)", fg="#1565C0")

        self.log("주문서를 불러오고 [변환 실행]을 누르세요.")
        self._drain_log()

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
            self.log("─" * 38)
            if 사유:
                self.log("⚠ 수동확인(빨강) %d건 — 결과 '비고'열에도 기입됨:" % len(사유))
                for excel_row, 상품, reason in 사유:
                    self.log("  · 행%d  %s\n       → %s" % (excel_row, str(상품)[:28], reason))
            else:
                self.log("✅ 수동확인 0건 — 모두 정상 매칭되었습니다.")
            self.log("─" * 38)
            self.log("✅ 저장 완료: " + out)
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

    def _split_worker(self, 발화):
        try:
            양식폴더 = os.path.join(app_dir(), "출력양식")
            if not os.path.isdir(양식폴더):     # exe가 dist면 상위 폴더의 출력양식
                양식폴더 = os.path.join(os.path.dirname(app_dir()), "출력양식")
            out_dir = os.path.join(app_dir(), "분리출력")
            res = engine.발주서분리(self.conv_file, 발화, 양식폴더, out_dir, log=self.log)
            self.split_dir = out_dir
            self.log("─" * 38)
            for 택배사, n, p in res:
                self.log("  %s : %d건 → %s" % (택배사, n, os.path.basename(p)))
            self.log("✅ 분리 완료: " + out_dir)
            self.root.after(0, lambda: self.btn_split_open.config(state="normal"))
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
