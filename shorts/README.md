# 쇼핑쇼츠 생성기

상품 소재(이미지 또는 영상 클립)와 나레이션 문구만 주면 **1080x1920 세로 쇼츠**를
자동으로 뽑아준다. 유료 프로그램 없이 전부 무료 도구로 돌아간다.

## 무엇이 자동인가

| 단계 | 처리 |
|---|---|
| 나레이션 음성 | edge-tts (한국어 3종, 무료·API 키 불필요) |
| 장면 길이 | 나레이션 길이에 맞춰 자동 계산 |
| 화면 | 9:16으로 꽉 채우고, 이미지는 켄 번스 줌 / 영상은 구간 컷 |
| 자막 | 흰 굵은 글씨 + 검은 외곽선, 말하는 구간에만 표시 (libass로 영상에 구움) |
| 최종 | H.264 + AAC, 유튜브·인스타 바로 업로드 가능 |

## 설치

```bash
pip install edge-tts imageio-ffmpeg
```

시스템에 ffmpeg가 있으면 그걸 쓰고, 없으면 imageio-ffmpeg 내장본을 자동으로 쓴다.
ffmpeg는 `libass`가 포함된 빌드여야 자막이 구워진다.

## 실행

```bash
python3 make_shorts.py sample_thermos.json
```

## 스펙 파일

```jsonc
{
  "output": "결과.mp4",
  "voice": "ko-KR-SunHiNeural",   // 여성. 남성은 ko-KR-InJoonNeural
  "rate": "+18%",                  // 말 속도
  "gap": 0.26,                     // 컷 끝나고 다음 컷까지 여백(초)
  "lead": 0.10,                    // 컷 바뀌고 말이 시작될 때까지 여백(초)
  "style": {
    "font": "Pretendard JJ",
    "size": 92,                    // 자막 크기
    "outline": 7,                  // 검은 외곽선 두께
    "shadow": 3,
    "margin_v": 500                // 화면 아래에서 자막까지 거리(px)
  },
  "scenes": [
    { "src": "assets/00.jpg", "text": "화면에 보일 자막", "say": "읽을 문장(생략 시 text와 동일)", "zoom": 0.12 }
  ]
}
```

- `src`는 이미지(jpg/png)와 영상(mp4/mov/mkv/webm) 둘 다 된다.
  영상이면 필요한 길이만큼 잘라 쓰고, 모자라면 반복한다.
- `text`는 6~10자로 짧게 끊는 것이 쇼츠 자막의 기본이다.
- `say`를 따로 두면 자막과 읽는 문장을 다르게 할 수 있다(숫자·단위 읽기 교정용).

## 폰트

`../ui/PretendardJJ-*.ttf` 를 사용한다. 다른 폰트를 쓰려면 `style.fontsdir`에
폰트 폴더 경로를, `style.font`에 폰트 패밀리 이름을 넣는다.

## 샘플에 쓴 이미지 출처

`assets/` 안의 사진은 전부 [StockSnap.io](https://stocksnap.io) 의 **CC0(퍼블릭 도메인)**
이미지를 [Openverse](https://openverse.org) API로 받아온 것이다. 상업적 이용과 재배포에
제약이 없다. 실제 제작에서는 이 자리에 판매할 상품의 실제 소재를 넣는다.

## 주의

남이 만든 영상을 받아서 그대로 쓰면 저작권 문제와 함께 유튜브 **재사용 콘텐츠 정책**에
걸려 수익 창출 심사에서 거절될 수 있다. 상품 제조사가 제공한 소재, 직접 촬영한 영상,
라이선스가 명확한 스톡 영상을 쓰는 것이 안전하다.
