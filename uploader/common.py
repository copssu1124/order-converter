#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""업로드 자동화 공용 도구.

원칙
----
1. **로그인은 최초 한 번만.** 로그인 결과(쿠키·로컬스토리지)를 파일로 저장해
   다음부터 재사용한다. 매번 로그인하면 그 자체가 이상 신호로 잡힌다.
2. **사람 속도로.** 클릭 사이에 불규칙한 간격을 둔다. 기계처럼 0초 간격으로
   움직이는 것이 가장 흔한 탐지 신호다.
3. **화면을 켜고 돌린다.** headless는 탐지가 쉽다. 기본은 창을 띄운다.
4. **막히면 멈춘다.** 못 찾은 요소를 추측해서 아무 데나 누르지 않는다.
   화면을 저장하고 사람에게 넘긴다.
"""

import os
import random
import sys
import time
from pathlib import Path

STATE_DIR = Path(os.environ.get("SHORTS_UPLOAD_HOME",
                                Path.home() / ".shorts-upload"))
SHOT_DIR = STATE_DIR / "screenshots"


def _need_playwright():
    try:
        from playwright.sync_api import sync_playwright        # noqa: F401
        return
    except ImportError:
        sys.exit("Playwright가 없습니다.\n"
                 "  pip install playwright\n"
                 "  playwright install chromium")


def state_path(site):
    """사이트별 로그인 세션 파일 경로."""
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    return STATE_DIR / f"{site}.session.json"


def pause(low=0.6, high=1.8):
    """사람이 화면을 보는 정도의 간격."""
    time.sleep(random.uniform(low, high))


def type_like_human(locator, text, low=0.04, high=0.13):
    """한 글자씩 불규칙한 속도로 입력한다."""
    locator.click()
    pause(0.2, 0.5)
    for ch in text:
        locator.type(ch)
        time.sleep(random.uniform(low, high))


class Session:
    """로그인 상태를 유지하는 브라우저 한 개.

    with Session("inpock") as s:
        page = s.page
        ...
    저장된 세션이 있으면 그대로 열고, 없으면 빈 상태로 연다.
    작업이 끝나면 세션을 다시 저장한다.
    """

    def __init__(self, site, headless=False, slow_mo=60, locale="ko-KR"):
        _need_playwright()
        self.site = site
        self.headless = headless
        self.slow_mo = slow_mo
        self.locale = locale
        self._pw = self._browser = self._ctx = self.page = None

    def __enter__(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=self.headless, slow_mo=self.slow_mo,
            args=["--disable-blink-features=AutomationControlled"])
        sp = state_path(self.site)
        kw = dict(locale=self.locale, timezone_id="Asia/Seoul",
                  viewport={"width": 1440, "height": 900})
        if sp.exists():
            kw["storage_state"] = str(sp)
            print(f"[{self.site}] 저장된 로그인 세션을 씁니다 ({sp})")
        else:
            print(f"[{self.site}] 저장된 세션이 없습니다. 이번에 로그인합니다.")
        self._ctx = self._browser.new_context(**kw)
        # navigator.webdriver 흔적을 지운다
        self._ctx.add_init_script(
            "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
        self.page = self._ctx.new_page()
        self.page.set_default_timeout(20000)
        return self

    def save(self):
        if self._ctx:
            self._ctx.storage_state(path=str(state_path(self.site)))
            print(f"[{self.site}] 로그인 세션을 저장했습니다.")

    def shot(self, name):
        """화면을 저장한다. 막혔을 때 무엇을 봤는지 남기기 위한 것."""
        SHOT_DIR.mkdir(parents=True, exist_ok=True)
        p = SHOT_DIR / f"{self.site}_{name}_{int(time.time())}.png"
        try:
            self.page.screenshot(path=str(p), full_page=True)
            print(f"  화면 저장: {p}")
        except Exception as e:
            print(f"  화면 저장 실패: {e}")
        return p

    def stop(self, why, name="stopped"):
        """추측하지 않고 멈춘다."""
        print(f"\n[멈춤] {why}")
        self.shot(name)
        raise SystemExit(1)

    def __exit__(self, *exc):
        try:
            if self._ctx:
                self._ctx.close()
        finally:
            if self._browser:
                self._browser.close()
            if self._pw:
                self._pw.stop()
