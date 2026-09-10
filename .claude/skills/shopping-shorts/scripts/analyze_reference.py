#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
잘 나가는 쇼츠(레퍼런스)를 넣으면 그 스타일을 수치로 뽑아준다.

"이런 느낌으로" 같은 감이 아니라, 컷 길이·자막 크기·자막 위치·움직임량을
실제로 재서 make_shorts.py 에 그대로 넣을 수 있는 설정값으로 돌려준다.

사용법:
    python3 analyze_reference.py 레퍼런스1.mp4 레퍼런스2.mp4 ...
"""

import argparse
import json
import statistics
import subprocess
import sys

import cv2
import numpy as np

try:
    from detect_subtitle import detect, stroke_mask
except ImportError:
    import os
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from detect_subtitle import detect, stroke_mask

OUT_W, OUT_H = 1080, 1920


def frame_series(path, step=2, size=(180, 320)):
    """프레임 간 변화량 곡선. 컷 전환과 움직임량을 모두 여기서 뽑는다."""
    cap = cv2.VideoCapture(path)
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    prev, vals, idx = None, [], 0
    while True:
        ok = cap.grab()
        if not ok:
            break
        if idx % step == 0:
            ok, f = cap.retrieve()
            if ok:
                g = cv2.cvtColor(cv2.resize(f, size), cv2.COLOR_BGR2GRAY).astype("float32")
                if prev is not None:
                    vals.append(float(np.abs(g - prev).mean()))
                prev = g
        idx += 1
    cap.release()
    return np.array(vals), fps / step


def cut_times(vals, rate, factor=3.2, floor=12.0, min_gap=0.35):
    """변화량이 확 튀는 지점을 컷 전환으로 본다."""
    if len(vals) < 5:
        return []
    med = float(np.median(vals))
    thr = max(med * factor, floor)
    times, last = [], -99.0
    for i, v in enumerate(vals):
        t = (i + 1) / rate
        if v >= thr and t - last >= min_gap:
            times.append(round(t, 2))
            last = t
    return times


def subtitle_style(path, band):
    """자막 띠 안을 들여다보고 글자 크기와 배경 판 여부를 잰다."""
    cap = cv2.VideoCapture(path)
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 1)
    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    dark_ratios, heights = [], []
    for i in np.linspace(total * 0.1, total * 0.9, 12).astype(int):
        cap.set(cv2.CAP_PROP_POS_FRAMES, int(i))
        ok, f = cap.read()
        if not ok:
            continue
        g = cv2.cvtColor(f, cv2.COLOR_BGR2GRAY)[band["y"]:band["y"] + band["h"],
                                                band["x"]:band["x"] + band["w"]]
        if g.size == 0:
            continue
        dark_ratios.append(float((g < 80).mean()))
        # 글자(밝은 획)가 실제로 차지하는 세로 높이
        rows = np.where(stroke_mask(g).mean(axis=1) > 0.02)[0]
        if len(rows) > 1:
            heights.append(int(rows[-1] - rows[0] + 1))
    cap.release()
    if not heights:
        return None
    glyph_h = statistics.median(heights)
    # 화면 세로를 1920으로 환산했을 때의 글자 높이 → 폰트 크기는 대략 그 1.35배
    scaled = glyph_h * OUT_H / h
    return {
        "글자높이_1920환산": round(scaled),
        "추정_폰트크기": round(scaled * 1.35),
        "박스배경_추정": statistics.median(dark_ratios) > 0.42 if dark_ratios else False,
        "띠안_어두운비율": round(statistics.median(dark_ratios), 3) if dark_ratios else None,
    }


def audio_stats(path, ffmpeg="ffmpeg"):
    """무음 구간이 얼마나 되는지 (BGM이 깔려 있으면 거의 0에 가깝다)."""
    p = subprocess.run([ffmpeg, "-hide_banner", "-i", path,
                        "-af", "silencedetect=noise=-45dB:d=0.3", "-f", "null", "-"],
                       capture_output=True, text=True)
    tot = 0.0
    for line in p.stderr.splitlines():
        if "silence_duration:" in line:
            try:
                tot += float(line.split("silence_duration:")[1].strip())
            except ValueError:
                pass
    return round(tot, 2)


def analyze(path, ffmpeg="ffmpeg"):
    cap = cv2.VideoCapture(path)
    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    cap.release()
    dur = n / fps if fps else 0.0

    vals, rate = frame_series(path)
    cuts = cut_times(vals, rate)
    lens = [round(b - a, 2) for a, b in zip([0.0] + cuts, cuts + [dur])]
    lens = [x for x in lens if x > 0.15]

    det = detect(path, samples=36)
    bands = det.get("bands", [])
    style = subtitle_style(path, bands[0]) if bands else None

    return {
        "파일": path.split("/")[-1][:46],
        "길이": round(dur, 1),
        "해상도": f"{w}x{h}",
        "컷수": len(cuts) + 1,
        "컷길이_중앙값": round(statistics.median(lens), 2) if lens else None,
        "컷길이_평균": round(statistics.mean(lens), 2) if lens else None,
        "첫3초_컷수": sum(1 for c in cuts if c <= 3.0) + 1,
        "평균_움직임": round(float(vals.mean()), 2),
        "자막_세로위치_%": round(bands[0]["y_ratio"] * 100, 1) if bands else None,
        "자막_가로폭_%": round(bands[0]["w_ratio"] * 100, 1) if bands else None,
        "자막스타일": style,
        "무음_총합_초": audio_stats(path, ffmpeg),
    }


def main():
    ap = argparse.ArgumentParser(description="레퍼런스 쇼츠 스타일 수치화")
    ap.add_argument("videos", nargs="+")
    ap.add_argument("--ffmpeg", default="ffmpeg")
    a = ap.parse_args()

    rows = []
    for v in a.videos:
        try:
            r = analyze(v, a.ffmpeg)
        except Exception as e:
            print(f"[건너뜀] {v}: {e}", file=sys.stderr)
            continue
        rows.append(r)
        print(json.dumps(r, ensure_ascii=False))

    if len(rows) < 2:
        return

    def med(key, sub=None):
        xs = []
        for r in rows:
            v = (r.get("자막스타일") or {}).get(key) if sub else r.get(key)
            if isinstance(v, (int, float)):
                xs.append(v)
        return round(statistics.median(xs), 2) if xs else None

    summary = {
        "레퍼런스_수": len(rows),
        "길이_중앙값": med("길이"),
        "컷길이_중앙값": med("컷길이_중앙값"),
        "첫3초_컷수_중앙값": med("첫3초_컷수"),
        "평균_움직임_중앙값": med("평균_움직임"),
        "자막_세로위치_%_중앙값": med("자막_세로위치_%"),
        "추정_폰트크기_중앙값": med("추정_폰트크기", sub=True),
        "박스배경_쓴_비율": round(
            sum(1 for r in rows if (r.get("자막스타일") or {}).get("박스배경_추정")) / len(rows), 2),
    }
    print("\n=== 종합 ===")
    print(json.dumps(summary, ensure_ascii=False, indent=2))

    print("\n=== 이 값을 스펙에 넣으면 된다 ===")
    print(json.dumps({
        "style": {
            "font": "Pretendard JJ",
            "size": summary["추정_폰트크기_중앙값"],
            "box": summary["박스배경_쓴_비율"] >= 0.5,
            "pad": 18,
            "box_color": "&H73000000",
            "margin_v": "auto",
        },
        "목표_컷길이": summary["컷길이_중앙값"],
        "목표_평균움직임": summary["평균_움직임_중앙값"],
        "목표_길이": summary["길이_중앙값"],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
