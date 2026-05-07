package compiler

import (
	"sort"

	"github.com/rhino-acoustic/pptx-engine/internal/mapper"
)

// RawNode represents the unparsed CSS/DOM node from the browser scraper.
type RawNode struct {
	Type            string    `json:"type"` // "box", "text", "image"
	X               float64   `json:"x"`
	Y               float64   `json:"y"`
	W               float64   `json:"w"`
	H               float64   `json:"h"`
	Rotation        float64   `json:"rotation,omitempty"`
	ZIndex          int       `json:"zIndex,omitempty"`
	Overflow        string    `json:"overflow,omitempty"`
	BgColor         string    `json:"bgColor,omitempty"`
	BorderColor     string    `json:"borderColor,omitempty"`
	BorderStyle     string    `json:"borderStyle,omitempty"`
	BorderWidth     float64   `json:"borderWidth,omitempty"`
	BrCorners       []float64 `json:"brCorners,omitempty"`
	HasShadow       bool      `json:"hasShadow,omitempty"`
	BgGradient      string    `json:"bgGradient,omitempty"`
	Opacity         float64   `json:"opacity,omitempty"`
	Text            string    `json:"text,omitempty"`
	Color           string    `json:"color,omitempty"`
	Size            float64   `json:"size,omitempty"`
	FontWeight      string    `json:"fontWeight,omitempty"`
	FontFamily      string    `json:"fontFamily,omitempty"`
	TextAlign       string    `json:"textAlign,omitempty"`
	LineHeight      string    `json:"lineHeight,omitempty"`
	LetterSpacing   string    `json:"letterSpacing,omitempty"`
	FontStyle       string    `json:"fontStyle,omitempty"`
	ImageUrl        string    `json:"imageUrl,omitempty"`
	ObjectFit       string    `json:"objectFit,omitempty"`
	ImageRId        string    `json:"-"` // set by main.go, not from JSON
}

// PptxElement is the compiled AST node ready for rendering.
type PptxElement struct {
	Type       string                  `json:"type"` // "shape", "text", "image", "table"
	X          float64                 `json:"x"`
	Y          float64                 `json:"y"`
	W          float64                 `json:"w"`
	H          float64                 `json:"h"`
	Rotation   float64                 `json:"rotation,omitempty"`
	ZIndex     int                     `json:"zIndex"`
	Overflow   string                  `json:"overflow,omitempty"`
	Shape      mapper.ShapeConfig      `json:"shape,omitempty"`
	Border     *mapper.BorderConfig    `json:"border,omitempty"`
	Shadow     *mapper.ShadowConfig    `json:"shadow,omitempty"`
	Fill       map[string]interface{}  `json:"fill,omitempty"`
	Gradient   *mapper.GradientConfig  `json:"gradient,omitempty"`
	Table      *mapper.TableConfig     `json:"table,omitempty"`
	Image      *mapper.ImageConfig     `json:"image,omitempty"`
	Text       string                  `json:"text,omitempty"`
	TextConfig *TextConfig             `json:"textConfig,omitempty"`
	Animation  *mapper.AnimationConfig `json:"animation,omitempty"`
	ImageRId   string                  `json:"-"` // relationship ID for embedded image
}

// CompileNode transforms a raw CSS/DOM node into a PptxElement.
func CompileNode(raw RawNode) PptxElement {
	// Conversion factor: 96 DPI (Standard PPTX resolution)
	// 1 inch = 96 px. This ensures perfect 1:1 circle/aspect ratio.
	px2in := 1.0 / 96.0

	el := PptxElement{
		X:        raw.X * px2in,
		Y:        raw.Y * px2in,
		W:        raw.W * px2in,
		H:        raw.H * px2in,
		Rotation: raw.Rotation,
		ZIndex:   raw.ZIndex,
		Overflow: raw.Overflow,
	}

	switch raw.Type {
	case "box":
		el.Type = "shape"
		el.Shape = mapper.MapShapeRadius(raw.W, raw.H, raw.BrCorners)
		// Convert RectRadius from px to inches to match el.W and el.H
		if el.Shape.RectRadius > 0 {
			el.Shape.RectRadius *= px2in
		}
		el.Border = mapper.MapBorder(raw.BorderColor, raw.BorderStyle, raw.BorderWidth)
		el.Shadow = mapper.MapShadow(raw.HasShadow)
		
		// F1: TranslateFill — gradient A등급 → solid 폴백 (dash2pptx.mjs L1275-1313)
		opacity := raw.Opacity
		if opacity == 0 {
			opacity = 1.0
		}
		// A등급: 네이티브 그라데이션 시도
		el.Gradient = mapper.TranslateGradient(raw.BgGradient)
		if el.Gradient == nil {
			// 폴백: solid color
			fillColor, fillTrans := mapper.TranslateFill(raw.BgColor, raw.BgGradient, raw.HasShadow, opacity)
			if fillColor != "" {
				el.Fill = map[string]interface{}{"color": fillColor}
				if fillTrans > 0 {
					el.Fill["transparency"] = float64(fillTrans) / 100.0
				}
			}
		}

	case "image":
		el.Type = "image"
		el.Shape = mapper.MapShapeRadius(raw.W, raw.H, raw.BrCorners)
		imgCfg := mapper.MapImageFill(raw.ImageUrl, raw.ObjectFit, raw.W, raw.H, el.Shape)
		el.Image = &imgCfg

	case "text":
		el.Type = "text"
		el.Text = raw.Text
		
		hexText, textAlpha := mapper.ParseColorAndAlpha(raw.Color)
		if hexText == "" {
			hexText = "000000"
		}
		raw.Color = hexText
		
		// Use full text formatting with all CSS properties (dash2pptx.mjs parity)
		tc := MapTextFormattingFull(raw)
		tc.ColorAlpha = textAlpha
		el.TextConfig = &tc
	}

	// Global Animation Mapping (if element has CSS animation)
	el.Animation = mapper.MapAnimation("fade-up", el.Y)

	return el
}

// CompileSlide compiles all raw nodes for a slide and applies post-processing.
// This is the main entry point that replaces individual CompileNode calls.
// Implements dash2pptx.mjs renderSlide Phase 1-5 pipeline.
func CompileSlide(rawNodes []RawNode) []PptxElement {
	elements := make([]PptxElement, 0, len(rawNodes))

	// Phase 1: Compile all nodes
	for _, raw := range rawNodes {
		el := CompileNode(raw)
		elements = append(elements, el)
	}

	// Separate boxes and texts for post-processing
	var boxes []PptxElement
	var texts []*PptxElement
	for i := range elements {
		switch elements[i].Type {
		case "shape":
			boxes = append(boxes, elements[i])
		case "text":
			texts = append(texts, &elements[i])
		}
	}

	// Phase 1.5: C1 — 컬럼 감지 (dash2pptx.mjs L895-916)
	// x 좌표 클러스터링 (±0.31inch) → 컬럼 경계 추출
	if len(texts) >= 5 {
		xVals := make([]float64, 0, len(texts))
		for _, t := range texts {
			xVals = append(xVals, t.X)
		}
		sort.Float64s(xVals)
		type cluster struct{ vals []float64 }
		clusters := []cluster{{vals: []float64{xVals[0]}}}
		for i := 1; i < len(xVals); i++ {
			last := &clusters[len(clusters)-1]
			if xVals[i]-last.vals[len(last.vals)-1] < 0.31 {
				last.vals = append(last.vals, xVals[i])
			} else {
				if len(last.vals) < 2 {
					clusters = clusters[:len(clusters)-1]
				}
				clusters = append(clusters, cluster{vals: []float64{xVals[i]}})
			}
		}
		if len(clusters) > 0 && len(clusters[len(clusters)-1].vals) < 2 {
			clusters = clusters[:len(clusters)-1]
		}
		// 컬럼 경계를 기반으로 텍스트 w 클리핑
		if len(clusters) >= 2 {
			for ci := 0; ci < len(clusters)-1; ci++ {
				colRight := clusters[ci+1].vals[0] - 0.05
				for _, t := range texts {
					minX := clusters[ci].vals[0] - 0.31
					maxX := clusters[ci].vals[len(clusters[ci].vals)-1] + 0.31
					if t.X >= minX && t.X <= maxX {
						maxW := colRight - t.X
						if maxW > 0.5 && t.W > maxW {
							t.W = maxW
						}
					}
				}
			}
		}
	}

	// Phase 2: E4 — h 최소 보장 (dash2pptx.mjs L1241)
	for _, t := range texts {
		if t.TextConfig != nil {
			t.H = EnsureMinHeight(t.H, t.TextConfig.FontSize)
		}
	}

	// Phase 3: E3 — w 골디락스존 (dash2pptx.mjs L1208-1237)
	// (비활성화: 텍스트 박스 크기를 임의로 줄이면 줄바꿈 및 normAutofit으로 인한 폰트 크기 왜곡이 발생함)
	/*
	for _, t := range texts {
		if t.TextConfig == nil {
			continue
		}
		minW := EstimateTextWidthInch(t.Text, t.TextConfig.FontSize)
		// ... 생략 (원본 박스 크기 유지)
	}
	*/

	// Phase 4: E7 — 텍스트 대비 보정 (비활성화: 원본 CSS 색상을 훼손함)
	/*
	for _, t := range texts {
		if t.TextConfig == nil {
			continue
		}
		bgColor := findNearbyBg(t, boxes)
		t.TextConfig.Color = AdjustTextContrast(t.TextConfig.Color, bgColor)
	}
	*/

	// Phase 5: E8 — wrap 판정 (SSOT: 웹에서 1줄이었다면 PPT에서도 1줄이어야 함)
	for _, t := range texts {
		if t.TextConfig == nil {
			continue
		}
		// t.H (inch) * 96 = DOM Height (px)
		// FontSize (pt) / 0.75 = DOM FontSize (px)
		domH := t.H * 96.0
		domFontSize := t.TextConfig.FontSize / 0.75
		
		// If DOM height is less than 1.5x font size, it was a single line in HTML.
		if domH < domFontSize * 1.5 {
			t.TextConfig.Wrap = false
		} else {
			t.TextConfig.Wrap = true
		}
	}

	// Phase 6: D3 — 텍스트-in-박스 정렬 (비활성화: 원본 CSS의 좌우정렬(Align) 설정을 파괴함)
	/*
	for _, bp := range boxes {
		var insideTexts []*PptxElement
		for _, t := range texts {
			if t.X >= bp.X && t.X+t.W <= bp.X+bp.W+0.1 &&
				t.Y >= bp.Y && t.Y+t.H <= bp.Y+bp.H+0.1 {
				insideTexts = append(insideTexts, t)
			}
		}
		if len(insideTexts) == 1 {
			// Center align the single text inside box
			t := insideTexts[0]
			t.X = bp.X + (bp.W-t.W)/2
			t.Y = bp.Y + (bp.H-t.H)/2
			if t.TextConfig != nil {
				t.TextConfig.Align = "center"
				t.TextConfig.Valign = "middle"
			}
		}
	}
	*/

	// Phase 7: Boundary clipping
	slideW := 13.333
	slideH := 7.5
	for i := range elements {
		el := &elements[i]
		if el.X+el.W > slideW {
			el.W = slideW - el.X
		}
		if el.Y+el.H > slideH {
			el.H = slideH - el.Y
		}
		if el.W < 0.05 || el.H < 0.03 {
			el.Type = "" // mark for removal
		}
	}

	// Filter out empty elements
	result := make([]PptxElement, 0, len(elements))
	for _, el := range elements {
		if el.Type != "" {
			result = append(result, el)
		}
	}

	// Phase 8: Z-Index sorting (stable sort to preserve DOM order when zIndex is same)
	sort.SliceStable(result, func(i, j int) bool {
		return result[i].ZIndex < result[j].ZIndex
	})

	return result
}

// findNearbyBg finds the nearest background color for a text element.
// Ported from dash2pptx.mjs findNearbyBackgroundHex (L1371-1390)
func findNearbyBg(t *PptxElement, boxes []PptxElement) string {
	padded := struct{ x, y, w, h float64 }{
		t.X - 0.03, t.Y - 0.03, t.W + 0.06, t.H + 0.06,
	}

	type candidate struct {
		color   string
		overlap float64
		area    float64
	}
	var best *candidate

	for _, bp := range boxes {
		if bp.Fill == nil {
			continue
		}
		c, ok := bp.Fill["color"].(string)
		if !ok || c == "" {
			continue
		}

		// Calculate overlap
		ox := max(0, min(padded.x+padded.w, bp.X+bp.W)-max(padded.x, bp.X))
		oy := max(0, min(padded.y+padded.h, bp.Y+bp.H)-max(padded.y, bp.Y))
		overlap := ox * oy
		if overlap <= 0 {
			continue
		}

		area := bp.W * bp.H
		if best == nil || overlap > best.overlap || (overlap == best.overlap && area < best.area) {
			best = &candidate{color: c, overlap: overlap, area: area}
		}
	}

	if best != nil {
		return best.color
	}
	return "FFFFFF"
}

func abs(x float64) float64 {
	if x < 0 {
		return -x
	}
	return x
}

func min(a, b float64) float64 {
	if a < b {
		return a
	}
	return b
}

func max(a, b float64) float64 {
	if a > b {
		return a
	}
	return b
}

