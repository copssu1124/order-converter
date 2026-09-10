# 업로드 자동화

만든 쇼츠를 인포크링크·유튜브·틱톡·인스타에 올린다.

## 플랫폼별로 방식이 다르다 (2026년 9월 확인)

| 플랫폼 | 방식 | 이유 |
|---|---|---|
| 인포크링크 | Playwright | 공개 API가 없다 |
| 유튜브 | Playwright | 공식 API는 **심사 전이면 올린 영상이 비공개로 잠긴다** |
| 틱톡 | Playwright | 공식 API는 **심사 전이면 나만 보기(SELF_ONLY)로만 올라간다** |
| 인스타·페북 | **Graph API** | 브라우저 자동화는 메타가 금지한다. 계정이 정지될 수 있다 |

**인스타는 Playwright로 하지 않는다.** 메타는 브라우저를 조작하는 자동화를
명시적으로 금지하고, 걸리면 인스타·페북이 함께 정지된다. 대신 **내 계정에만
올리는 앱은 심사가 필요 없다** — 메타 앱을 개발 모드로 두고 내 인스타를
Instagram Tester 로 등록하면 그대로 게시된다. 계정은 프로(비즈니스/크리에이터)여야 한다.

## 설치

```bash
pip install playwright
playwright install chromium
```

## 원칙

1. **로그인은 최초 한 번.** 세션을 `~/.shorts-upload/` 에 저장해 재사용한다.
   매번 로그인하는 것 자체가 이상 신호다.
2. **사람 속도로.** 클릭·입력 사이에 불규칙한 간격을 둔다.
3. **창을 띄우고 돌린다.** headless 는 탐지가 쉽다.
4. **막히면 멈춘다.** 못 찾은 버튼을 추측해서 아무 데나 누르지 않는다.
   화면을 `~/.shorts-upload/screenshots/` 에 저장하고 멈춘다.
5. **내 컴퓨터에서 돌린다.** 서버·클라우드 IP로 로그인하면 그 자체가 위험하다.

## 쓰는 순서

### 1) 화면 구조 확인 (사이트마다 한 번)

자동화는 "어느 버튼을 누를지"를 알아야 한다. 로그인해야 보이는 화면이 많아서,
먼저 직접 들어가 본 뒤 그 화면의 입력칸·버튼 목록을 뽑는다.

```bash
python3 explore.py inpock https://link.inpock.co.kr --keep
```

창이 열리면 직접 로그인하고 **상품 추가 화면까지** 이동한 뒤, 터미널에서 Enter.
결과는 `~/.shorts-upload/explore/` 에 json 과 스크린샷으로 남는다.
**비밀번호는 기록하지 않는다.**

### 2) 인포크링크

```bash
export INPOCK_ID=아이디
export INPOCK_PW=비밀번호

python3 inpock.py login          # 최초 1회
python3 inpock.py check          # 세션이 살아있는지
python3 inpock.py add -j 상품.json
```

`상품.json`

```jsonc
{
  "title": "수납하마 폭좁은 틈새수납장",
  "url":   "https://link.coupang.com/a/xxxxx",
  "image": "썸네일.jpg",
  "price": 39900
}
```

비밀번호를 환경변수에 두기 싫으면 그냥 비워두면 물어본다.
**세션 파일과 `.env` 는 절대 저장소에 올리지 않는다.**

## 아직 안 만든 것

- `inpock.py add` 의 상품 등록 부분 — 화면 구조 확인 후 채운다
- `youtube.py` — studio.youtube.com 업로드
- `instagram.py` — Graph API (컨테이너 생성 → 상태 확인 → 게시)
- `tiktok.py` — 마지막. 가장 조심해야 한다
