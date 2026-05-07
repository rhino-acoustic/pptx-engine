package compiler

import (
	"math"
	"strconv"
	"strings"
	"unicode"
)

// TextConfig represents PPTX text formatting.
type TextConfig struct {
	FontFace      string    `json:"fontFace"`
	FontSize      float64   `json:"fontSize"`
	Color         string    `json:"color"`
	ColorAlpha    int       `json:"colorAlpha,omitempty"` // 0-100, 0=opaque, 100=fully transparent
	Bold          bool      `json:"bold"`
	Italic        bool      `json:"italic"`
	Align         string    `json:"align"`
	Valign        string    `json:"valign"`
	LineHeight    float64   `json:"lineHeight,omitempty"`
	LetterSpacing float64   `json:"letterSpacing,omitempty"`
	Margin        []float64 `json:"margin,omitempty"`
	Wrap          bool      `json:"wrap"`
}

// MapTextFormatting maps CSS typography to PPTX TextConfig.
// Ported from dash2pptx.mjs Goldilocks zone tuning.
func MapTextFormatting(fontFamily string, fontSize float64, color string, fontWeight string, align string) TextConfig {
	cfg := TextConfig{
		FontFace: MapFontFamily(fontFamily),
		Color:    color,
		Align:    "left",
		Valign:   "middle",
		Wrap:     true,
		Margin:   []float64{0, 0, 0, 0},
	}

	// Font Size: CSS px → PPTX pt = ×0.75 (96dpi → 72dpi)
	scaledSize := fontSize * 0.75
	cfg.FontSize = scaledSize

	// Bold detection — dash2pptx.mjs L1205
	w, _ := strconv.Atoi(fontWeight)
	if w >= 700 || fontWeight == "bold" {
		cfg.Bold = true
	}

	// Align
	switch align {
	case "center":
		cfg.Align = "center"
	case "right":
		cfg.Align = "right"
	}

	return cfg
}

// MapTextFormattingFull maps all CSS text properties including italic, lineHeight, letterSpacing.
func MapTextFormattingFull(raw RawNode) TextConfig {
	cfg := MapTextFormatting(raw.FontFamily, raw.Size, raw.Color, raw.FontWeight, raw.TextAlign)

	// Italic
	if raw.FontStyle == "italic" {
		cfg.Italic = true
	}

	// LetterSpacing: CSS "normal" or "1.5px" → centi-pt (100ths of a point)
	if raw.LetterSpacing != "" && raw.LetterSpacing != "normal" {
		ls := strings.TrimSuffix(raw.LetterSpacing, "px")
		if v, err := strconv.ParseFloat(ls, 64); err == nil {
			// 1px ≈ 0.75pt, PPTX uses hundredths of a point
			cfg.LetterSpacing = v * 0.75 * 100
		}
	}

	// LineHeight: CSS "normal" or "24px" or "1.5"
	if raw.LineHeight != "" && raw.LineHeight != "normal" {
		lh := strings.TrimSuffix(raw.LineHeight, "px")
		if v, err := strconv.ParseFloat(lh, 64); err == nil {
			if v > 5 {
				// absolute px value — convert to ratio
				cfg.LineHeight = v / raw.Size * 100 // percentage for PPTX spcPct
			} else {
				// ratio (e.g., 1.5)
				cfg.LineHeight = v * 100
			}
		}
	}

	return cfg
}

// EstimateTextWidthInch calculates minimum text width using CJK/Latin ratio.
// Ported from dash2pptx.mjs L1211-1216 Goldilocks zone.
// CJK: fs * 0.92 / 72 * 1.2 inch per char
// Latin: fs * 0.486 / 72 * 1.2 inch per char
func EstimateTextWidthInch(text string, fontSize float64) float64 {
	var w float64
	for _, ch := range text {
		if IsCJK(ch) {
			w += fontSize * 0.92 / 72.0 * 1.2
		} else {
			w += fontSize * 0.486 / 72.0 * 1.2
		}
	}
	if w < 0.5 {
		w = 0.5
	}
	return w
}

// IsCJK returns true if the rune is a CJK character (> U+2E80).
func IsCJK(r rune) bool {
	return r > 0x2E80 && !unicode.IsSpace(r)
}

// === E4: h 최소 보장 (dash2pptx.mjs L1241) ===
// h 최소: max(fs/72, 0.18) inch
func EnsureMinHeight(h float64, fontSize float64) float64 {
	lineH := fontSize / 72.0
	if lineH < 0.18 {
		lineH = 0.18
	}
	if h < lineH {
		return lineH
	}
	return h
}

// === E7: 텍스트 대비 보정 (dash2pptx.mjs L1391-1398) ===
// 흰 배경 위 흰 글씨 → 검정, 어두운 배경 위 어두운 글씨 → 흰색
func AdjustTextContrast(textColor string, bgColor string) string {
	textLum := Luminance(textColor)
	bgLum := Luminance(bgColor)
	if textLum < 0 || bgLum < 0 {
		return textColor
	}
	if textLum < 0.45 && bgLum < 0.45 {
		return "FFFFFF"
	}
	if textLum > 0.55 && bgLum > 0.55 {
		return "333333"
	}
	return textColor
}

// === G2: Luminance (dash2pptx.mjs L1361-1365) ===
// WCAG 상대 밝기 계산. 6자리 hex 입력.
func Luminance(hexColor string) float64 {
	if len(hexColor) != 6 {
		return -1
	}
	rgb := make([]float64, 3)
	for i := 0; i < 3; i++ {
		v64, err := strconv.ParseInt(hexColor[i*2:i*2+2], 16, 64)
		if err != nil {
			return -1
		}
		c := float64(v64) / 255.0
		if c <= 0.03928 {
			rgb[i] = c / 12.92
		} else {
			rgb[i] = math.Pow((c+0.055)/1.055, 2.4)
		}
	}
	return 0.2126*rgb[0] + 0.7152*rgb[1] + 0.0722*rgb[2]
}

// === E8: wrap 판정 (dash2pptx.mjs L1337-1346) ===
// 텍스트 예상 폭이 할당 w의 95%를 초과하면 wrap
func ShouldWrap(text string, fontSize float64, widthInch float64) bool {
	var estimatedW float64
	for _, ch := range text {
		if IsCJK(ch) {
			estimatedW += fontSize * 0.92 / 72.0
		} else {
			estimatedW += fontSize * 0.486 / 72.0
		}
	}
	return estimatedW > widthInch*0.95
}

// MapFontFamily simplifies browser font stacks to PPTX compatible presentation fonts.
// Ported from dash2pptx.mjs FONT_MAP (L10-23), expanded for production coverage.
func MapFontFamily(fontFamily string) string {
	ff := strings.ToLower(fontFamily)

	// 1. Brand specific Korean fonts
	if strings.Contains(ff, "pretendard") {
		return "Pretendard"
	}
	if strings.Contains(ff, "suit") || strings.Contains(ff, "spoqa") {
		return "Noto Sans KR"
	}
	if strings.Contains(ff, "noto sans kr") || strings.Contains(ff, "noto sans") {
		return "Noto Sans KR"
	}
	if strings.Contains(ff, "nanum gothic") || strings.Contains(ff, "nanumgothic") {
		return "나눔고딕"
	}
	if strings.Contains(ff, "nanum myeongjo") || strings.Contains(ff, "nanummyeongjo") {
		return "나눔명조"
	}
	if strings.Contains(ff, "nanum") {
		return "나눔고딕"
	}
	// Apple must come before generic "gothic" check
	if strings.Contains(ff, "apple sd") || strings.Contains(ff, "apple gothic") {
		return "Noto Sans KR"
	}
	if strings.Contains(ff, "malgun") || strings.Contains(ff, "맑은") || strings.Contains(ff, "gothic") {
		return "맑은 고딕"
	}
	if strings.Contains(ff, "dotum") || strings.Contains(ff, "돋움") {
		return "돋움"
	}
	if strings.Contains(ff, "gulim") || strings.Contains(ff, "굴림") {
		return "굴림"
	}
	if strings.Contains(ff, "batang") || strings.Contains(ff, "바탕") {
		return "바탕"
	}

	// 2. English presentation fonts
	if strings.Contains(ff, "montserrat") {
		return "Montserrat"
	}
	if strings.Contains(ff, "poppins") {
		return "Poppins"
	}
	if strings.Contains(ff, "dm sans") {
		return "DM Sans"
	}
	if strings.Contains(ff, "outfit") {
		return "Outfit"
	}
	if strings.Contains(ff, "raleway") {
		return "Raleway"
	}
	if strings.Contains(ff, "playfair") {
		return "Playfair Display"
	}
	if strings.Contains(ff, "instrument serif") || strings.Contains(ff, "georgia") {
		return "Georgia"
	}
	if strings.Contains(ff, "times") {
		return "Times New Roman"
	}
	if strings.Contains(ff, "courier") || strings.Contains(ff, "monospace") || strings.Contains(ff, "consolas") {
		return "Consolas"
	}
	if strings.Contains(ff, "roboto") || strings.Contains(ff, "inter") ||
		strings.Contains(ff, "segoe") || strings.Contains(ff, "helvetica") ||
		strings.Contains(ff, "sf pro") || strings.Contains(ff, "open sans") ||
		strings.Contains(ff, "lato") || strings.Contains(ff, "source sans") ||
		strings.Contains(ff, "ubuntu") || strings.Contains(ff, "nunito") {
		return "Arial"
	}
	if strings.Contains(ff, "serif") && !strings.Contains(ff, "sans") {
		return "Times New Roman"
	}

	// 3. System/fallback → Korean safe font
	if strings.Contains(ff, "sans-serif") || strings.Contains(ff, "system-ui") ||
		strings.Contains(ff, "ui-sans") || strings.Contains(ff, "blink") {
		return "Noto Sans KR"
	}

	return "Noto Sans KR"
}

