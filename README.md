# pptx-engine

**Bidirectional HTML ↔ PPTX converter in pure Go. Zero dependencies. Single binary.**

```
PPTX ⇄ HTML ⇄ PPTX
```

> 🇰🇷 [한국어 문서](./README_KR.md)

## What it does

| Command | Input | Output | Description |
|---|---|---|---|
| `pptx2html` | `.pptx` | `.html` | PPTX to pixel-accurate HTML |
| `roundtrip` | `.html` | `.pptx` | Lossless round-trip back to PPTX |
| `anyhtml2pptx` | **Any `.html`** | `.pptx` | Universal HTML to PPTX translator |

## Why this exists

- **No Node.js** — no PptxGenJS, no npm, no node_modules
- **No .NET** — no Aspose, no COM, no Office interop
- **No browser** — no Puppeteer, no headless Chrome
- **One binary** — `go build` and ship (~5MB)
- **Bidirectional** — both directions, same engine
- **CJK ready** — Korean/Japanese/Chinese font mapping built-in

## Quick Start

```bash
go build ./cmd/pptx2html/
go build ./cmd/roundtrip/
go build ./cmd/anyhtml2pptx/

# PPTX → HTML
./pptx2html -input slides.pptx -output slides.html -images slide_images/

# HTML → PPTX (round-trip)
./roundtrip -html slides.html -stub template.pptx -out output.pptx

# Any HTML → PPTX
./anyhtml2pptx -html report.html -stub template.pptx -out report.pptx
```

## Architecture

```
               pptx2html                    roundtrip / anyhtml2pptx
  ┌────────┐  ──────────▶  ┌────────┐  ──────────────▶  ┌────────┐
  │  PPTX  │               │  HTML  │                    │  PPTX  │
  │ (.pptx)│  ◀──────────  │ (.html)│  ◀──────────────  │ (.pptx)│
  └────────┘   roundtrip   └────────┘    anyhtml2pptx    └────────┘
```

### Modules

| Package | File | Purpose |
|---|---|---|
| `internal/reverse` | `pptxparser.go` | Native PPTX/ZIP/XML parser |
| `internal/reverse` | `html.go` | HTML generator from parsed PPTX |
| `internal/reverse` | `htmlparser.go` | HTML → ParsedDoc reverse parser |
| `internal/compiler` | `compiler.go` | DOM/CSS → PptxElement AST compiler |
| `internal/compiler` | `html2pptx.go` | ParsedDoc → PptxElement adapter |
| `internal/compiler` | `text.go` | Typography mapping (CSS ↔ OOXML) |
| `internal/openxml` | `builder.go` | Go-native PPTX ZIP/XML builder |
| `internal/mapper` | `style.go` | CSS ↔ OOXML style bridge |

## Supported Elements

### PPTX → HTML
Shapes, text (full typography), images (with crop), tables (cell merge), gradients, borders, shadows, rotation, group shapes, chart SVG, theme/master/layout inheritance.

### Any HTML → PPTX
`h1`-`h6`, `p`, `div`, `span`, `ul`/`ol`, `table`, `img`, `hr`, `strong`, `em`, `a`. CSS inline styles + `<style>` blocks + class/id selectors. Auto-pagination with heading-based slide breaks.

## Stub PPTX

Both `roundtrip` and `anyhtml2pptx` require a `-stub` PPTX file. This provides the theme, slide masters, and layouts. Use any existing `.pptx` as a template.

## Unit System

```
1 inch = 914,400 EMU
1 pt   = 12,700 EMU
1 px   = 9,525 EMU (96 DPI)
```

## Built With

This project was architected and built using [NeuronFS](https://github.com/rhino-acoustic/NeuronFS) — a persistent neural memory system for AI agents that ensures deterministic, quark-level task decomposition and zero-hallucination code generation.

## License

MIT — see [LICENSE](./LICENSE)
