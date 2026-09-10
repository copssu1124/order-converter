#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""화면에 무엇이 있는지 훑어서 적어준다.

자동화 코드는 "어느 버튼을 누를지"를 알아야 하는데, 사이트마다 다르고
로그인해야만 보이는 화면도 많다. 그래서 먼저 이 도구로 화면을 열고,
사람이 직접 원하는 화면까지 간 뒤 Enter를 누르면 그 화면의 입력창·버튼을
전부 적어준다. 그 목록을 보고 자동화 코드를 정확히 쓴다.

사용법
------
    python3 explore.py inpock https://link.inpock.co.kr
    (창이 열리면 직접 로그인하고 원하는 화면까지 이동 → 터미널에서 Enter)

결과는 ~/.shorts-upload/explore/ 에 저장된다. 비밀번호는 적지 않는다.
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import Session, STATE_DIR                      # noqa: E402

DUMP_JS = r"""
() => {
  const vis = el => {
    const r = el.getBoundingClientRect();
    const s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
  };
  const txt = el => (el.innerText || el.value || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  const out = { url: location.href, title: document.title,
                inputs: [], buttons: [], links: [], headings: [] };
  for (const el of document.querySelectorAll('input,textarea,select')) {
    if (!vis(el)) continue;
    out.inputs.push({ tag: el.tagName.toLowerCase(), type: el.type || '',
                      name: el.name || '', id: el.id || '',
                      placeholder: el.placeholder || '',
                      label: (el.labels && el.labels[0]) ? txt(el.labels[0]) : '',
                      aria: el.getAttribute('aria-label') || '' });
  }
  for (const el of document.querySelectorAll('button,[role=button],a[class*=btn],input[type=submit]')) {
    if (!vis(el)) continue;
    const t = txt(el);
    if (t) out.buttons.push({ text: t, aria: el.getAttribute('aria-label') || '',
                              id: el.id || '' });
  }
  for (const el of document.querySelectorAll('a[href]')) {
    if (!vis(el)) continue;
    const t = txt(el);
    if (t) out.links.push({ text: t, href: el.getAttribute('href') });
  }
  for (const el of document.querySelectorAll('h1,h2,h3')) {
    if (vis(el) && txt(el)) out.headings.push(txt(el));
  }
  const uniq = (a, k) => { const s = new Set(); return a.filter(x => {
    const v = JSON.stringify(k ? x[k] : x); if (s.has(v)) return false; s.add(v); return true; }); };
  out.buttons = uniq(out.buttons).slice(0, 60);
  out.links = uniq(out.links).slice(0, 60);
  return out;
}
"""


def main():
    ap = argparse.ArgumentParser(description="화면 구조 훑어보기")
    ap.add_argument("site", help="세션 이름 (inpock, youtube, tiktok ...)")
    ap.add_argument("url", help="처음 열 주소")
    ap.add_argument("--keep", action="store_true", help="로그인 세션을 저장한다")
    a = ap.parse_args()

    out_dir = STATE_DIR / "explore"
    out_dir.mkdir(parents=True, exist_ok=True)

    with Session(a.site) as s:
        s.page.goto(a.url, wait_until="domcontentloaded")
        print("\n창이 열렸습니다.")
        print("  1) 직접 로그인하고, 자동화하려는 화면까지 이동하세요.")
        print("  2) 다 됐으면 여기 터미널에서 Enter 를 누르세요.")
        print("  (여러 화면을 남기려면 화면마다 Enter, 끝내려면 q + Enter)\n")

        n = 0
        while True:
            cmd = input("Enter=지금 화면 기록 / q=종료 > ").strip().lower()
            if cmd == "q":
                break
            n += 1
            try:
                info = s.page.evaluate(DUMP_JS)
            except Exception as e:
                print("  읽기 실패:", e)
                continue
            f = out_dir / f"{a.site}_{n:02d}.json"
            f.write_text(json.dumps(info, ensure_ascii=False, indent=2),
                         encoding="utf-8")
            s.shot(f"explore_{n:02d}")
            print(f"  기록됨: {f}")
            print(f"    주소  : {info['url']}")
            print(f"    입력칸 {len(info['inputs'])}개 / 버튼 {len(info['buttons'])}개")
            for i in info["inputs"][:8]:
                print(f"      · input {i['type']:<10} name={i['name'][:20]:<20} "
                      f"placeholder={i['placeholder'][:24]}")
            for b in info["buttons"][:10]:
                print(f"      · 버튼 「{b['text']}」")

        if a.keep:
            s.save()
    print(f"\n결과 폴더: {out_dir}")
    print("이 폴더의 json 파일과 스크린샷을 저한테 보내주시면 자동화 코드를 정확히 씁니다.")


if __name__ == "__main__":
    main()
