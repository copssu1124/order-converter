#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
쇼핑쇼츠 자동 생성기 (1080x1920 / 9:16)

스펙 JSON 하나를 받아서 아래를 자동으로 처리한다.
  1) 장면별 나레이션을 TTS로 합성 (edge-tts, 무료)
  2) 나레이션 길이에 맞춰 장면 길이를 자동 결정
  3) 이미지/영상 소재를 9:16으로 채우고 줌 효과 부여
  4) 흰 굵은 글씨 + 검은 외곽선 자막을 영상에 구움 (libass)
  5) 나레이션을 합쳐 최종 mp4로 인코딩

사용법:
    python3 make_shorts.py 스펙파일.json

필요한 것:  pip install edge-tts imageio-ffmpeg
"""

import asyncio
import json
import os
import re
import ssl
import subprocess
import sys
import tempfile

# ── 프록시 환경(사내망/컨테이너)에서 TTS가 인증서 오류로 막히는 경우 대비 ──────
_CA_BUNDLE = os.environ.get("SHORTS_CA_BUNDLE") or "/root/.ccr/ca-bundle.crt"
_PROXY = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")


def ffmpeg_exe():
    """ffmpeg 실행 파일 경로. 시스템에 없으면 imageio-ffmpeg 내장본을 쓴다."""
    for cand in ("ffmpeg", "/usr/bin/ffmpeg", "/usr/local/bin/ffmpeg"):
        try:
            subprocess.run([cand, "-version"], capture_output=True, check=True)
            return cand
        except Exception:
            pass
    import imageio_ffmpeg
    return imageio_ffmpeg.get_ffmpeg_exe()


FFMPEG = ffmpeg_exe()


def run(args):
    """ffmpeg 실행. 실패하면 에러 로그를 그대로 올린다."""
    p = subprocess.run([FFMPEG, "-y", "-hide_banner", "-loglevel", "error"] + args,
                       capture_output=True, text=True)
    if p.returncode != 0:
        raise RuntimeError("ffmpeg 실패\n" + " ".join(args) + "\n" + p.stderr[-3000:])


def probe_duration(path):
    """ffprobe 없이 ffmpeg 출력에서 재생 길이(초)를 읽는다."""
    p = subprocess.run([FFMPEG, "-hide_banner", "-i", path],
                       capture_output=True, text=True)
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+\.\d+)", p.stderr)
    if not m:
        raise RuntimeError("길이를 못 읽음: " + path)
    h, mi, s = m.groups()
    return int(h) * 3600 + int(mi) * 60 + float(s)


# ── 1) TTS ────────────────────────────────────────────────────────────────────
async def _tts_one(text, voice, rate, out_path):
    import edge_tts
    import edge_tts.communicate as _c
    import edge_tts.voices as _v
    if os.path.exists(_CA_BUNDLE):
        ctx = ssl.create_default_context(cafile=_CA_BUNDLE)
        _c._SSL_CTX = ctx
        _v._SSL_CTX = ctx
    kw = {"proxy": _PROXY} if _PROXY else {}
    await edge_tts.Communicate(text, voice, rate=rate, **kw).save(out_path)


def trim_silence(src, out_path, threshold="-45dB"):
    """앞뒤 무음을 잘라낸다.

    edge-tts가 내주는 mp3는 앞뒤에 0.5초 안팎의 빈 소리가 붙어 있다.
    그대로 이어붙이면 컷마다 1초 넘게 말이 끊겨서 시청자가 바로 이탈한다.
    앞을 자르고, 뒤집어서 다시 앞을 자른 뒤, 되돌리는 방식으로 양쪽을 없앤다.
    """
    one = (f"silenceremove=start_periods=1:start_duration=0:"
           f"start_threshold={threshold}:detection=peak")
    run(["-i", src, "-af", f"{one},areverse,{one},areverse",
         "-ar", "44100", "-ac", "2", out_path])
    return probe_duration(out_path)


def tts(text, voice, rate, out_path):
    """나레이션을 합성하고 앞뒤 무음까지 제거한 wav 경로와 길이를 돌려준다."""
    raw = out_path + ".raw.mp3"
    asyncio.run(_tts_one(text, voice, rate, raw))
    return trim_silence(raw, out_path)


# ── 2) 장면 영상 ──────────────────────────────────────────────────────────────
VIDEO_EXT = (".mp4", ".mov", ".mkv", ".webm", ".avi")

_MASK_CACHE = {}


def auto_regions(src):
    """원본에 박힌 자막·워터마크 위치를 자동으로 찾는다 (같은 파일은 한 번만)."""
    if src in _MASK_CACHE:
        return _MASK_CACHE[src]
    try:
        from detect_subtitle import detect
    except ImportError:
        import os as _os, sys as _sys
        _sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
        from detect_subtitle import detect
    found = detect(src).get("regions", [])
    _MASK_CACHE[src] = found
    for r in found:
        print(f"      가림: {r['zone']} x={r['x']} y={r['y']} {r['w']}x{r['h']}")
    return found


def mask_filter(regions, mode="blur", strength=28):
    """원본 위의 지정 영역을 가리는 필터 문자열을 만든다.

    blur = 뭉갠다(자연스러움) / box = 검은 상자로 덮는다(확실함)
    """
    if not regions:
        return ""
    if mode == "box":
        return ",".join(
            f"drawbox=x={r['x']}:y={r['y']}:w={r['w']}:h={r['h']}:color=black@1:t=fill"
            for r in regions)

    # 워터마크처럼 작은 것은 delogo가 주변 화면을 끌어와 메워서 훨씬 덜 티난다.
    # 자막 띠처럼 큰 것은 delogo가 뭉개지므로 블러로 처리한다.
    small = [r for r in regions if r["w"] * r["h"] < 1080 * 1920 * 0.03]
    big = [r for r in regions if r not in small]
    head_chain = ",".join(
        f"delogo=x={max(1, r['x'])}:y={max(1, r['y'])}:w={r['w']}:h={r['h']}"
        for r in small)

    if not big:
        return head_chain
    parts, cur = [], None
    for i, r in enumerate(big):
        a, b, c = f"bg{i}", f"rg{i}", f"bl{i}"
        head = f"[{cur}]" if cur else ""
        parts.append(f"{head}split=2[{a}][{b}]")
        # boxblur 반경은 잘라낸 영역 크기를 넘을 수 없다(색차 평면 기준). 맞춰 줄인다.
        radius = max(2, min(strength, min(r["w"], r["h"]) // 4))
        parts.append(f"[{b}]crop={r['w']}:{r['h']}:{r['x']}:{r['y']},"
                     f"boxblur={radius}:2,gblur=sigma={max(6, strength)}[{c}]")
        cur = f"ov{i}"
        parts.append(f"[{a}][{c}]overlay={r['x']}:{r['y']}[{cur}]")
    chain = ";".join(parts) + f";[{cur}]null"
    return (head_chain + "," + chain) if head_chain else chain


def scene_video(src, duration, out_path, zoom_in=True, zoom_amount=0.14, fps=30,
                mask=None):
    """이미지 또는 동영상 한 컷을 1080x1920으로 만든다.

    이미지면 천천히 줌(켄 번스), 동영상이면 필요한 구간만 잘라 9:16으로 채운다.
    """
    frames = max(2, int(round(duration * fps)))
    is_video = os.path.splitext(src)[1].lower() in VIDEO_EXT

    # 원본에 박힌 자막·워터마크 가리기 (9:16으로 자르기 전에, 원본 좌표로 처리)
    pre = ""
    if mask:
        regions = mask.get("regions")
        if regions in ("auto", None):
            regions = auto_regions(src)
        pre = mask_filter(regions, mask.get("mode", "blur"),
                          int(mask.get("strength", 28)))
        if pre:
            pre += ","

    if is_video:
        vf = (f"{pre}scale=1080:1920:force_original_aspect_ratio=increase,"
              f"crop=1080:1920,fps={fps},setsar=1,format=yuv420p")
        args = ["-stream_loop", "-1", "-i", src, "-t", f"{duration:.3f}", "-an", "-vf", vf]
    else:
        # 줌 화질 손실을 막으려고 2배로 키운 뒤 zoompan으로 잘라낸다.
        k = zoom_amount / frames
        if zoom_in:
            z = f"min(1+{k:.8f}*on,{1 + zoom_amount:.4f})"
        else:
            z = f"max({1 + zoom_amount:.4f}-{k:.8f}*on,1.0)"
        vf = (f"scale=2160:3840:force_original_aspect_ratio=increase,crop=2160:3840,"
              f"zoompan=z='{z}':d=1:x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':"
              f"s=1080x1920:fps={fps},setsar=1,format=yuv420p")
        args = ["-loop", "1", "-framerate", str(fps), "-t", f"{duration:.3f}", "-i", src, "-vf", vf]

    run(args + ["-c:v", "libx264", "-preset", "veryfast", "-crf", "18",
                "-pix_fmt", "yuv420p", "-r", str(fps), out_path])


# ── 3) 장면 오디오 ────────────────────────────────────────────────────────────
def scene_audio(voice_wav, duration, lead, out_path):
    """나레이션 앞에 lead초 여백을 주고, 장면 길이만큼 무음으로 채운다."""
    chain = []
    if lead > 0:
        delay_ms = int(lead * 1000)
        chain.append(f"adelay={delay_ms}|{delay_ms}")
    chain += ["apad", "aformat=sample_fmts=s16:channel_layouts=stereo"]
    run(["-i", voice_wav, "-af", ",".join(chain),
         "-t", f"{duration:.3f}", "-ar", "44100", "-ac", "2", out_path])


# ── 4) 자막 ───────────────────────────────────────────────────────────────────
ASS_HEADER = """[Script Info]
ScriptType: v4.00+
PlayResX: 1080
PlayResY: 1920
WrapStyle: 2
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Main,{font},{size},&H00FFFFFF,&H000000FF,&H00000000,&H64000000,-1,0,0,0,100,100,0,0,1,{outline},{shadow},2,50,50,{margin_v},1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""


def ass_time(t):
    h = int(t // 3600)
    m = int((t % 3600) // 60)
    s = t % 60
    return f"{h}:{m:02d}:{s:05.2f}"


def build_ass(cues, style, out_path):
    body = ASS_HEADER.format(
        font=style.get("font", "Pretendard JJ"),
        size=style.get("size", 64),
        outline=style.get("outline", 5),
        shadow=style.get("shadow", 2),
        margin_v=style.get("margin_v", 520),
    )
    for start, end, text in cues:
        text = text.replace("\n", "\\N")
        body += f"Dialogue: 0,{ass_time(start)},{ass_time(end)},Main,,0,0,0,,{text}\n"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(body)


# ── 5) 전체 조립 ──────────────────────────────────────────────────────────────
def build(spec, workdir):
    voice = spec.get("voice", "ko-KR-SunHiNeural")
    rate = spec.get("rate", "+15%")
    # 쇼츠는 말이 끊기는 순간 이탈한다. 기본값은 거의 붙여 읽는 쪽으로 잡는다.
    gap = float(spec.get("gap", 0.05))       # 장면 사이 여백(초)
    lead = float(spec.get("lead", 0.0))      # 컷 바뀌고 말이 시작될 때까지 여백(초)
    style = spec.get("style", {})
    fps = int(spec.get("fps", 30))
    scenes = spec["scenes"]

    # 원본 자막이 있던 자리에 한글 자막을 얹으면, 가리느라 생긴 자국이 글자에
    # 덮여서 훨씬 덜 티난다. margin_v를 "auto"로 두면 그 위치를 자동으로 맞춘다.
    if str(style.get("margin_v", "")).lower() == "auto":
        band = None
        for sc in scenes:
            if os.path.splitext(sc["src"])[1].lower() not in VIDEO_EXT:
                continue
            if not sc.get("mask", spec.get("mask")):
                continue
            for r in auto_regions(sc["src"]):
                if r["zone"] == "하단" and (band is None or r["h"] > band["h"]):
                    band = r
            break
        style = dict(style)
        style["margin_v"] = max(140, 1920 - (band["y"] + band["h"])) if band else 500
        print(f"      자막 위치 자동: 아래에서 {style['margin_v']}px")

    seg_videos, seg_audios, cues = [], [], []
    t = 0.0
    for i, sc in enumerate(scenes):
        voice_wav = os.path.join(workdir, f"tts{i:02d}.wav")
        say = sc.get("say", sc["text"])       # 읽는 문장과 자막을 따로 둘 수 있다
        dur_tts = tts(say, sc.get("voice", voice), sc.get("rate", rate), voice_wav)
        # 장면 길이를 프레임 단위로 딱 맞춘다. 안 맞추면 컷·음성·자막이 조금씩
        # 어긋난 채로 쌓여서 뒤로 갈수록 자막이 밀린다.
        dur = round((dur_tts + lead + gap) * fps) / fps

        v = os.path.join(workdir, f"v{i:02d}.mp4")
        scene_mask = sc.get("mask", spec.get("mask"))
        if scene_mask in ("auto", True):
            scene_mask = {"regions": "auto"}
        scene_video(sc["src"], dur, v, zoom_in=(i % 2 == 0),
                    zoom_amount=float(sc.get("zoom", 0.14)), fps=fps,
                    mask=scene_mask)
        seg_videos.append(v)

        a = os.path.join(workdir, f"a{i:02d}.wav")
        scene_audio(voice_wav, dur, lead, a)
        seg_audios.append(a)

        # 자막은 장면 전체를 덮는다. 말이 끝났다고 자막까지 지우면 화면이 빈다.
        cues.append((t, t + dur, sc["text"]))
        t += dur
        print(f"  [{i + 1}/{len(scenes)}] {dur:4.2f}s  {sc['text']}")

    # 영상 이어붙이기
    lst = os.path.join(workdir, "v.txt")
    with open(lst, "w") as f:
        for p in seg_videos:
            f.write(f"file '{p}'\n")
    body = os.path.join(workdir, "body.mp4")
    run(["-f", "concat", "-safe", "0", "-i", lst, "-c", "copy", body])

    # 나레이션 이어붙이기
    lst_a = os.path.join(workdir, "a.txt")
    with open(lst_a, "w") as f:
        for p in seg_audios:
            f.write(f"file '{p}'\n")
    narr = os.path.join(workdir, "narr.wav")
    run(["-f", "concat", "-safe", "0", "-i", lst_a, "-c", "copy", narr])

    # 자막 굽고 최종 인코딩
    ass = os.path.join(workdir, "sub.ass")
    build_ass(cues, style, ass)
    fontsdir = style.get("fontsdir", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ui"))
    vf = f"subtitles={ass}:fontsdir={os.path.abspath(fontsdir)}"

    out = spec.get("output", "shorts_out.mp4")
    run(["-i", body, "-i", narr, "-vf", vf,
         "-c:v", "libx264", "-preset", "medium", "-crf", "20", "-pix_fmt", "yuv420p",
         "-c:a", "aac", "-b:a", "128k", "-shortest", "-movflags", "+faststart", out])
    return out, t


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    spec_path = sys.argv[1]
    with open(spec_path, encoding="utf-8") as f:
        spec = json.load(f)

    # 스펙 파일 위치 기준으로 소재 경로를 푼다.
    base = os.path.dirname(os.path.abspath(spec_path))
    for sc in spec["scenes"]:
        if not os.path.isabs(sc["src"]):
            sc["src"] = os.path.join(base, sc["src"])

    print(f"장면 {len(spec['scenes'])}개 렌더 시작")
    with tempfile.TemporaryDirectory(prefix="shorts_") as wd:
        out, total = build(spec, wd)
    print(f"\n완료: {out}  ({total:.1f}초)")


if __name__ == "__main__":
    main()
