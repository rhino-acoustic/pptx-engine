# pptx-engine

**순수 Go로 만든 양방향 HTML ↔ PPTX 변환기. 의존성 제로. 단일 바이너리.**

```
PPTX ⇄ HTML ⇄ PPTX
```

> 🇺🇸 [English](./README.md)

## 기능

| 명령어 | 입력 | 출력 | 설명 |
|---|---|---|---|
| `pptx2html` | `.pptx` | `.html` | PPTX를 픽셀 정확도의 HTML로 변환 |
| `roundtrip` | `.html` | `.pptx` | pptx2html 출력물을 다시 PPTX로 역변환 |
| `anyhtml2pptx` | **임의 `.html`** | `.pptx` | 아무 HTML이나 PPTX로 번역 |

## 왜 만들었나

- **Node.js 불필요** — PptxGenJS, npm, node_modules 없음
- **\.NET 불필요** — Aspose, COM, Office interop 없음
- **브라우저 불필요** — Puppeteer, 헤드리스 Chrome 없음
- **단일 바이너리** — `go build` 한 번이면 끝 (~5MB)
- **양방향** — 같은 엔진으로 두 방향 모두 지원
- **한국어 폰트** — Freesentation, Pretendard, Noto Sans KR 내장 매핑

## 빠른 시작

```bash
go build ./cmd/pptx2html/
go build ./cmd/roundtrip/
go build ./cmd/anyhtml2pptx/

# PPTX → HTML
./pptx2html -input 발표자료.pptx -output 슬라이드.html -images slide_images/

# HTML → PPTX (왕복 변환)
./roundtrip -html 슬라이드.html -stub 템플릿.pptx -out 출력.pptx

# 임의 HTML → PPTX
./anyhtml2pptx -html 보고서.html -stub 템플릿.pptx -out 보고서.pptx
```

## 아키텍처

```
               pptx2html                    roundtrip / anyhtml2pptx
  ┌────────┐  ──────────▶  ┌────────┐  ──────────────▶  ┌────────┐
  │  PPTX  │               │  HTML  │                    │  PPTX  │
  │ (.pptx)│  ◀──────────  │ (.html)│  ◀──────────────  │ (.pptx)│
  └────────┘   roundtrip   └────────┘    anyhtml2pptx    └────────┘
```

### 모듈 구조

| 패키지 | 파일 | 역할 |
|---|---|---|
| `internal/reverse` | `pptxparser.go` | Go 네이티브 PPTX/ZIP/XML 파서 |
| `internal/reverse` | `html.go` | 파싱된 PPTX를 HTML로 렌더링 |
| `internal/reverse` | `htmlparser.go` | HTML → ParsedDoc 역파서 |
| `internal/compiler` | `compiler.go` | DOM/CSS → PptxElement AST 컴파일러 |
| `internal/compiler` | `html2pptx.go` | ParsedDoc → PptxElement 어댑터 |
| `internal/compiler` | `text.go` | 타이포그래피 매핑 (CSS ↔ OOXML) |
| `internal/openxml` | `builder.go` | Go 네이티브 PPTX ZIP/XML 빌더 |
| `internal/mapper` | `style.go` | CSS ↔ OOXML 스타일 브릿지 |

## 지원 요소

### PPTX → HTML
도형, 텍스트(전체 타이포그래피), 이미지(크롭 포함), 테이블(셀 병합), 그라데이션, 테두리, 그림자, 회전, 그룹 도형, 차트 SVG, 테마/마스터/레이아웃 상속.

### 임의 HTML → PPTX
`h1`-`h6`, `p`, `div`, `span`, `ul`/`ol`, `table`, `img`, `hr`, `strong`, `em`, `a`. CSS inline + `<style>` 블록 + class/id 셀렉터. 높이 기반 자동 페이지네이션, heading 단위 슬라이드 분리.

## Stub PPTX

`roundtrip`과 `anyhtml2pptx`는 `-stub` PPTX 파일이 필요합니다. 테마, 슬라이드 마스터, 레이아웃을 제공합니다. 아무 `.pptx` 파일이나 템플릿으로 사용 가능합니다.

## 단위 체계

```
1 inch = 914,400 EMU
1 pt   = 12,700 EMU
1 px   = 9,525 EMU (96 DPI)
```

## 개발 도구

이 프로젝트는 [NeuronFS](https://github.com/rhino-acoustic/NeuronFS)를 통해 설계·구현되었습니다. NeuronFS는 AI 에이전트를 위한 영속적 신경 메모리 시스템으로, 결정론적 쿼크 단위 작업 분해와 무환각 코드 생성을 보장합니다.

## 라이센스

MIT — [LICENSE](./LICENSE) 참조
