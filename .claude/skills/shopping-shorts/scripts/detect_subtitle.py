#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
영상에 박혀 있는 자막(하드서브) 위치를 자동으로 찾는다.

원리
----
쇼츠에 박히는 자막은 거의 예외 없이 **밝은 글씨 + 어두운 외곽선**이다.
자연스러운 사진에는 거의 없는 조합이라, 이 신호만 세면 자막이 있는 줄을
정확히 골라낼 수 있다. OCR도 GPU도 필요 없다.

  1) 영상에서 프레임을 고르게 뽑는다
  2) 각 픽셀이 "밝은데 바로 옆에 아주 어두운 픽셀이 있는가"를 본다  → 획 후보
  3) 가로줄마다 획 후보 비율을 세고, 프레임 전체의 평균을 낸다
  4) 비율이 튀는 구간을 자막 띠로 본다
  5) 그 띠 안에서 좌우 범위도 같은 방법으로 잘라낸다

사용법
------
    python3 detect_subtitle.py 영상.mp4              # 결과를 JSON으로 출력
    python3 detect_subtitle.py 영상.mp4 --preview 미리보기.png
"""

import argparse
import json
import sys

import cv2
import numpy as np


def stroke_mask(gray, bright=190, dark=90, reach=5):
    """밝은 획 + 바로 옆 어두운 외곽선 조합을 True로 표시한다."""
    bright_px = gray >= bright
    dark_px = (gray <= dark).astype(np.uint8)
    # 좌우 reach 픽셀 안에 어두운 점이 있는지 (가로 방향 팽창)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (reach * 2 + 1, 3))
    dark_near = cv2.dilate(dark_px, kernel) > 0
    return bright_px & dark_near


def row_scores(video, samples=48, bright=190, dark=90):
    """가로줄별 획 비율 평균과, 영상 크기를 돌려준다."""
    cap = cv2.VideoCapture(video)
    if not cap.isOpened():
        raise RuntimeError("영상을 열 수 없음: " + video)
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    if total <= 0:
        total = 1

    idxs = np.linspace(0, max(0, total - 1), min(samples, max(1, total))).astype(int)
    acc = np.zeros(h, dtype=np.float64)
    col_acc = np.zeros(w, dtype=np.float64)
    used = 0
    for i in idxs:
        cap.set(cv2.CAP_PROP_POS_FRAMES, int(i))
        ok, frame = cap.read()
        if not ok:
            continue
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        m = stroke_mask(gray, bright, dark)
        acc += m.mean(axis=1)
        col_acc += m.mean(axis=0)
        used += 1
    cap.release()
    if used == 0:
        raise RuntimeError("읽을 수 있는 프레임이 없음")
    return acc / used, col_acc / used, w, h


def find_bands(score, min_rows, factor=4.0, floor=0.010):
    """점수가 튀는 구간(자막 띠)들을 찾는다."""
    base = float(np.median(score))
    thr = max(base * factor, floor)
    hot = score >= thr

    bands, start = [], None
    for i, v in enumerate(hot):
        if v and start is None:
            start = i
        elif not v and start is not None:
            if i - start >= min_rows:
                bands.append((start, i))
            start = None
    if start is not None and len(hot) - start >= min_rows:
        bands.append((start, len(hot)))

    # 사이가 가까운 띠는 하나로 합친다 (두 줄 자막)
    merged = []
    for b in bands:
        if merged and b[0] - merged[-1][1] <= min_rows * 2:
            merged[-1] = (merged[-1][0], b[1])
        else:
            merged.append(list(b))
            merged[-1] = tuple(merged[-1])
            continue
    return [tuple(b) for b in merged], thr, base


def col_extent(video, y0, y1, samples=48, bright=190, dark=90, ratio=0.02):
    """띠 안에서 글자가 실제로 있는 좌우 범위를 찾는다."""
    cap = cv2.VideoCapture(video)
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 1)
    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
    idxs = np.linspace(0, max(0, total - 1), min(samples, max(1, total))).astype(int)
    acc = np.zeros(w, dtype=np.float64)
    used = 0
    for i in idxs:
        cap.set(cv2.CAP_PROP_POS_FRAMES, int(i))
        ok, frame = cap.read()
        if not ok:
            continue
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)[y0:y1]
        acc += stroke_mask(gray, bright, dark).mean(axis=0)
        used += 1
    cap.release()
    if used == 0:
        return 0, w
    acc /= used
    hot = np.where(acc >= max(acc.max() * 0.15, ratio))[0]
    if len(hot) == 0:
        return 0, w
    return int(hot[0]), int(hot[-1]) + 1


def find_static_overlays(video, samples=32, std_max=6.0, min_area_ratio=0.0006,
                         edge_min=18.0):
    """영상 내내 픽셀이 거의 안 변하는 덧씌운 요소(워터마크·로고)를 찾는다.

    자막은 내용이 바뀌지만 워터마크는 처음부터 끝까지 같은 자리에 같은 모양으로
    박혀 있다. 그래서 프레임들 사이의 픽셀 변화량이 거의 0인데 주변과 대비는
    뚜렷한 곳을 고르면 워터마크만 남는다.
    """
    cap = cv2.VideoCapture(video)
    if not cap.isOpened():
        raise RuntimeError("영상을 열 수 없음: " + video)
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 1)
    idxs = np.linspace(0, max(0, total - 1), min(samples, max(1, total))).astype(int)

    frames = []
    for i in idxs:
        cap.set(cv2.CAP_PROP_POS_FRAMES, int(i))
        ok, f = cap.read()
        if ok:
            frames.append(cv2.cvtColor(f, cv2.COLOR_BGR2GRAY).astype(np.float32))
    cap.release()
    if len(frames) < 4:
        return []

    # 워터마크는 대개 반투명이라 픽셀 값 자체는 배경 따라 변한다.
    # 하지만 "글자 윤곽"은 모든 프레임에서 같은 자리에 그대로 남는다.
    # 그래서 픽셀값이 아니라 윤곽선이 몇 %의 프레임에 나타나는지를 센다.
    h, w = frames[0].shape
    k3 = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    presence = np.zeros((h, w), dtype=np.float32)
    for f in frames:
        g = cv2.morphologyEx(f.astype(np.uint8), cv2.MORPH_GRADIENT, k3)
        presence += (g >= edge_min).astype(np.float32)
    presence /= len(frames)

    cand = (presence >= 0.92).astype(np.uint8) * 255
    cand = cv2.morphologyEx(cand, cv2.MORPH_CLOSE,
                            cv2.getStructuringElement(cv2.MORPH_RECT, (25, 9)))
    cand = cv2.morphologyEx(cand, cv2.MORPH_OPEN,
                            cv2.getStructuringElement(cv2.MORPH_RECT, (9, 5)))

    n, _lab, stats, _cent = cv2.connectedComponentsWithStats(cand, 8)
    out = []
    min_area = h * w * min_area_ratio
    for i in range(1, n):
        x, y, bw, bh, area = stats[i]
        if area < min_area or bh < h * 0.008 or bw < w * 0.02:
            continue
        if bw > w * 0.95 and bh > h * 0.95:      # 화면 전체는 배경이다
            continue
        out.append({"x": int(x), "y": int(y), "w": int(bw), "h": int(bh),
                    "area": int(area)})
    out.sort(key=lambda b: -b["area"])
    return out[:6]


def detect(video, samples=48, pad_ratio=0.012, min_row_ratio=0.012):
    rows, _cols, w, h = row_scores(video, samples)
    min_rows = max(6, int(h * min_row_ratio))
    bands, thr, base = find_bands(rows, min_rows)

    pad = int(h * pad_ratio)
    out = []
    for y0, y1 in bands:
        x0, x1 = col_extent(video, y0, y1, samples)
        xp = int(w * pad_ratio)
        X0 = max(0, x0 - xp)
        X1 = min(w, x1 + xp)
        Y0 = max(0, y0 - pad)
        Y1 = min(h, y1 + pad)
        out.append({
            "x": X0, "y": Y0, "w": X1 - X0, "h": Y1 - Y0,
            "x_ratio": round(X0 / w, 4), "y_ratio": round(Y0 / h, 4),
            "w_ratio": round((X1 - X0) / w, 4), "h_ratio": round((Y1 - Y0) / h, 4),
            "score": round(float(rows[y0:y1].mean()), 5),
            "zone": "상단" if Y1 < h * 0.25 else ("하단" if Y0 > h * 0.55 else "중앙"),
        })
    out.sort(key=lambda b: -b["score"])

    # 워터마크는 밝기가 아니라 "안 변한다"로 따로 찾고, 자막 띠와 겹치면 버린다.
    marks = []
    for m in find_static_overlays(video):
        overlap = any(not (m["y"] + m["h"] <= b["y"] or m["y"] >= b["y"] + b["h"])
                      for b in out)
        if overlap:
            continue
        xp, yp = int(w * pad_ratio), int(h * pad_ratio)
        X0, Y0 = max(0, m["x"] - xp), max(0, m["y"] - yp)
        X1, Y1 = min(w, m["x"] + m["w"] + xp), min(h, m["y"] + m["h"] + yp)
        marks.append({
            "x": X0, "y": Y0, "w": X1 - X0, "h": Y1 - Y0,
            "x_ratio": round(X0 / w, 4), "y_ratio": round(Y0 / h, 4),
            "w_ratio": round((X1 - X0) / w, 4), "h_ratio": round((Y1 - Y0) / h, 4),
            "zone": "상단" if Y1 < h * 0.3 else ("하단" if Y0 > h * 0.7 else "중앙"),
        })

    return {"width": w, "height": h, "threshold": round(thr, 5),
            "baseline": round(base, 5), "bands": out, "watermarks": marks,
            "regions": out + marks}


def preview(video, result, out_png):
    """찾은 영역을 빨간 상자로 그려서 저장한다 (눈으로 확인용)."""
    cap = cv2.VideoCapture(video)
    total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 1)
    cap.set(cv2.CAP_PROP_POS_FRAMES, total // 3)
    ok, frame = cap.read()
    cap.release()
    if not ok:
        return False
    for b in result["bands"]:
        cv2.rectangle(frame, (b["x"], b["y"]), (b["x"] + b["w"], b["y"] + b["h"]),
                      (0, 0, 255), 4)
    for m in result.get("watermarks", []):
        cv2.rectangle(frame, (m["x"], m["y"]), (m["x"] + m["w"], m["y"] + m["h"]),
                      (0, 200, 255), 4)
    cv2.imwrite(out_png, frame)
    return True


def main():
    ap = argparse.ArgumentParser(description="영상에 박힌 자막 위치 자동 탐지")
    ap.add_argument("video")
    ap.add_argument("--samples", type=int, default=48, help="분석할 프레임 수")
    ap.add_argument("--preview", help="찾은 영역을 표시한 png 저장 경로")
    a = ap.parse_args()

    r = detect(a.video, a.samples)
    print(json.dumps(r, ensure_ascii=False, indent=2))
    if a.preview:
        print("미리보기:", a.preview, preview(a.video, r, a.preview), file=sys.stderr)


if __name__ == "__main__":
    main()
