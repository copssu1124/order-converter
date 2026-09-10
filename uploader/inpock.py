#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""인포크링크에 상품 링크를 올린다.

  python3 inpock.py login              최초 1회. 로그인하고 세션을 저장한다.
  python3 inpock.py add -j 상품.json    상품(링크)을 하나 추가한다.
  python3 inpock.py check              세션이 아직 살아있는지 본다.

아이디·비밀번호는 코드에 적지 않는다. 환경변수로 주거나, 없으면 물어본다.
  export INPOCK_ID=아이디
  export INPOCK_PW=비밀번호
"""

import argparse
import getpass
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import Session, pause, type_like_human, state_path   # noqa: E402

HOME = "https://link.inpock.co.kr"
LOGIN = HOME + "/user/login"

# 화면이 바뀌어도 버티도록 여러 후보를 순서대로 시도한다.
ID_HINTS = ["아이디", "이메일", "ID", "email", "userId", "username"]
PW_HINTS = ["비밀번호", "패스워드", "password", "pw"]
LOGIN_BTN = ["로그인", "Log in", "Login", "Sign in"]


def first_visible(page, finders, what):
    """후보를 순서대로 찾아 처음 보이는 것을 돌려준다."""
    for make in finders:
        try:
            loc = make()
            if loc.count() and loc.first.is_visible():
                return loc.first
        except Exception:
            continue
    return None


def find_input(page, hints, input_type=None):
    def by_ph(h):
        return lambda: page.get_by_placeholder(h, exact=False)

    def by_label(h):
        return lambda: page.get_by_label(h, exact=False)

    finders = [by_ph(h) for h in hints] + [by_label(h) for h in hints]
    if input_type:
        finders.append(lambda: page.locator(f"input[type={input_type}]"))
    for h in hints:
        finders.append(lambda h=h: page.locator(f"input[name*='{h}' i]"))
    return first_visible(page, finders, hints[0])


def find_button(page, texts):
    finders = []
    for t in texts:
        finders.append(lambda t=t: page.get_by_role("button", name=t, exact=False))
        finders.append(lambda t=t: page.get_by_text(t, exact=True))
    finders.append(lambda: page.locator("button[type=submit]"))
    return first_visible(page, finders, texts[0])


def logged_in(page):
    """로그인된 화면인지 본다. 로그인 폼이 안 보이면 들어간 것으로 본다."""
    pause(0.8, 1.5)
    if "/login" in page.url:
        return False
    return find_input(page, PW_HINTS, "password") is None


def do_login(s, user, pw):
    page = s.page
    page.goto(LOGIN, wait_until="domcontentloaded")
    pause(1.0, 2.0)

    if logged_in(page):
        print("이미 로그인되어 있습니다.")
        return True

    idbox = find_input(page, ID_HINTS)
    pwbox = find_input(page, PW_HINTS, "password")
    if not idbox or not pwbox:
        s.stop("로그인 입력칸을 못 찾았습니다. explore.py 로 화면을 먼저 확인해주세요.",
               "login_no_fields")

    type_like_human(idbox, user)
    pause()
    type_like_human(pwbox, pw)
    pause(0.5, 1.2)

    btn = find_button(page, LOGIN_BTN)
    if btn:
        btn.click()
    else:
        pwbox.press("Enter")

    try:
        page.wait_for_load_state("networkidle", timeout=25000)
    except Exception:
        pass
    pause(1.5, 2.5)

    if not logged_in(page):
        s.stop("로그인이 안 됐습니다. 아이디·비밀번호나 추가 인증(문자·캡차)을 확인해주세요.",
               "login_failed")
    print("로그인 성공.")
    return True


def cmd_login(a):
    user = os.environ.get("INPOCK_ID") or input("인포크 아이디: ").strip()
    pw = os.environ.get("INPOCK_PW") or getpass.getpass("인포크 비밀번호: ")
    with Session("inpock", headless=False) as s:
        do_login(s, user, pw)
        s.save()
        print("\n다음부터는 로그인 없이 바로 씁니다.")
        input("창을 닫으려면 Enter > ")


def cmd_check(a):
    sp = state_path("inpock")
    if not sp.exists():
        sys.exit("저장된 세션이 없습니다. 먼저 `python3 inpock.py login` 을 하세요.")
    with Session("inpock", headless=False) as s:
        s.page.goto(HOME, wait_until="domcontentloaded")
        ok = logged_in(s.page)
        print("세션 살아있음 ✓" if ok else "세션 만료됨 — 다시 login 하세요")
        if ok:
            s.save()
        sys.exit(0 if ok else 2)


def cmd_add(a):
    spec = json.loads(Path(a.json).read_text(encoding="utf-8"))
    for key in ("title", "url"):
        if not spec.get(key):
            sys.exit(f"상품 json에 '{key}' 가 필요합니다.")
    img = spec.get("image")
    if img and not Path(img).exists():
        sys.exit(f"이미지 파일이 없습니다: {img}")

    with Session("inpock", headless=False) as s:
        page = s.page
        page.goto(HOME, wait_until="domcontentloaded")
        pause(1.0, 2.0)
        if not logged_in(page):
            s.stop("로그인이 풀렸습니다. `python3 inpock.py login` 을 다시 하세요.",
                   "add_not_logged_in")

        # 여기서부터는 인포크링크 화면 구조에 맞춰 채워 넣는다.
        # explore.py 로 뽑은 버튼·입력칸 이름이 확정되면 그대로 옮겨 적는다.
        s.stop("상품 추가 화면의 버튼 이름이 아직 확정되지 않았습니다.\n"
               "  python3 explore.py inpock https://link.inpock.co.kr --keep\n"
               "  로 '상품 추가' 화면까지 이동한 뒤 Enter 를 눌러 결과를 보내주세요.",
               "add_todo")


def main():
    ap = argparse.ArgumentParser(description="인포크링크 자동 등록")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("login", help="최초 1회 로그인하고 세션 저장")
    sub.add_parser("check", help="세션이 살아있는지 확인")
    p = sub.add_parser("add", help="상품 링크 추가")
    p.add_argument("-j", "--json", required=True, help="상품 정보 json")
    a = ap.parse_args()
    {"login": cmd_login, "check": cmd_check, "add": cmd_add}[a.cmd](a)


if __name__ == "__main__":
    main()
