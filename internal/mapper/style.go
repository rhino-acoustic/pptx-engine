package mapper

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
)

// BorderConfig represents PPTX line options.
type BorderConfig struct {
	Color    string  `json:"color"`
	Width    float64 `json:"width"`
	DashType string  `json:"dashType,omitempty"`
}

// ShadowConfig represents PPTX shadow options.
type ShadowConfig struct {
	Type      string  `json:"type"`
	Blur      int     `json:"blur"`
	Offset    int     `json:"offset"`
	Direction int     `json:"direction"` // 0-21600000 (OpenXML 60000ths of a degree)
	Color     string  `json:"color"`
	Opacity   float64 `json:"opacity"`
}

// MapBorder maps CSS border properties to PPTX line configuration.
func MapBorder(borderColor, borderStyle string, borderWidth float64) *BorderConfig {
	if borderWidth <= 0 {
		return nil
	}
	hex, transparency := ParseColorAndAlpha(borderColor)
	if hex == "" || transparency >= 100 {
		return nil
	}

	cfg := &BorderConfig{
		Color: hex,
		Width: math.Max(0.3, borderWidth*0.4), // Scale down width for PPTX visual parity
	}
	// For simplicity, we skip border transparency in this struct, but it could be added if needed

	if borderStyle == "dashed" {
		cfg.DashType = "dash"
	} else if borderStyle == "dotted" {
		cfg.DashType = "sysDot"
	}

	return cfg
}

// MapShadow maps CSS box-shadow presence to a PPTX shadow approximation.
func MapShadow(hasShadow bool) *ShadowConfig {
	if !hasShadow {
		return nil
	}
	// Default soft outer shadow for UI elements
	return &ShadowConfig{
		Type:      "outer",
		Blur:      4,
		Offset:    2,
		Direction: 2700000, // 45 degrees
		Color:     "000000",
		Opacity:   0.15,
	}
}

// MapShadowFromCSS parses CSS box-shadow string and returns ShadowConfig.
// Format: "offsetX offsetY blur spread color" e.g. "2px 4px 8px 0px rgba(0,0,0,0.3)"
func MapShadowFromCSS(boxShadow string) *ShadowConfig {
	if boxShadow == "" || boxShadow == "none" {
		return nil
	}

	// Extract numeric values (px)
	pxRe := regexp.MustCompile(`(-?\d+(?:\.\d+)?)px`)
	pxMatches := pxRe.FindAllStringSubmatch(boxShadow, -1)

	offsetX, offsetY, blur := 2.0, 2.0, 4.0
	if len(pxMatches) >= 1 {
		offsetX, _ = strconv.ParseFloat(pxMatches[0][1], 64)
	}
	if len(pxMatches) >= 2 {
		offsetY, _ = strconv.ParseFloat(pxMatches[1][1], 64)
	}
	if len(pxMatches) >= 3 {
		blur, _ = strconv.ParseFloat(pxMatches[2][1], 64)
	}

	// Combined offset (diagonal distance)
	offset := math.Sqrt(offsetX*offsetX + offsetY*offsetY)
	
	// Calculate direction in degrees (0 to 360), mapping atan2(y, x) to PPTX EMU (1 deg = 60000 EMU)
	dirRad := math.Atan2(offsetY, offsetX)
	dirDeg := dirRad * 180.0 / math.Pi
	if dirDeg < 0 {
		dirDeg += 360.0
	}
	direction := int(math.Round(dirDeg * 60000.0))

	// Extract color and opacity
	color := "000000"
	opacity := 0.25
	hex, trans := ParseColorAndAlpha(boxShadow)
	if hex != "" {
		color = hex
		opacity = float64(100-trans) / 100.0
	}

	return &ShadowConfig{
		Type:      "outer",
		Blur:      int(blur),
		Offset:    int(offset),
		Direction: direction,
		Color:     color,
		Opacity:   opacity,
	}
}

// ParseColorAndAlpha converts CSS rgb/rgba string to 6-char HEX and a transparency percentage (0-100).
func ParseColorAndAlpha(colorStr string) (string, int) {
	colorStr = strings.TrimSpace(colorStr)
	
	// Transparent keyword
	if colorStr == "transparent" || colorStr == "rgba(0, 0, 0, 0)" {
		return "", 100
	}

	// Hex with alpha (e.g. #FF000080)
	if strings.HasPrefix(colorStr, "#") {
		h := colorStr[1:]
		if len(h) == 3 {
			return strings.ToUpper(fmt.Sprintf("%c%c%c%c%c%c", h[0], h[0], h[1], h[1], h[2], h[2])), 0
		}
		if len(h) == 8 {
			alphaHex := h[6:8]
			alphaInt, _ := strconv.ParseInt(alphaHex, 16, 64)
			transparency := 100 - int((float64(alphaInt)/255.0)*100.0)
			return strings.ToUpper(h[:6]), transparency
		}
		if len(h) >= 6 {
			return strings.ToUpper(h[:6]), 0
		}
	}

	// rgba parsing
	re := regexp.MustCompile(`rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+))?\s*\)`)
	matches := re.FindStringSubmatch(colorStr)
	if len(matches) >= 4 {
		r, _ := strconv.Atoi(matches[1])
		g, _ := strconv.Atoi(matches[2])
		b, _ := strconv.Atoi(matches[3])
		
		hex := fmt.Sprintf("%02X%02X%02X", r, g, b)
		transparency := 0
		
		if len(matches) > 4 && matches[4] != "" {
			alphaFloat, _ := strconv.ParseFloat(matches[4], 64)
			transparency = 100 - int(alphaFloat*100.0)
		}
		return hex, transparency
	}

	return "", 0
}

// TranslateFill converts CSS background (solid or gradient) to PPTX fill.
// Ported from dash2pptx.mjs translateFill (L1275-1313)
func TranslateFill(bgColor string, bgGradient string, hasShadow bool, opacity float64) (string, int) {
	if opacity == 0 {
		opacity = 1.0
	}

	// B등급: gradient → 두 색상 50% 블렌드 solid
	if bgGradient != "" && (strings.HasPrefix(bgGradient, "linear-gradient") || strings.HasPrefix(bgGradient, "radial-gradient")) {
		re := regexp.MustCompile(`linear-gradient\(\s*(\d+)deg\s*,\s*(#[0-9a-fA-F]{3,8})\s*(?:\d+%?)?\s*,\s*(#[0-9a-fA-F]{3,8})`)
		m := re.FindStringSubmatch(bgGradient)
		if len(m) >= 4 {
			r1, g1, b1 := parseHex(m[2])
			r2, g2, b2 := parseHex(m[3])
			blended := fmt.Sprintf("%02X%02X%02X", (r1+r2)/2, (g1+g2)/2, (b1+b2)/2)
			return blended, 0
		}
		// 폴백: 첫 HEX 추출
		fc := regexp.MustCompile(`#([0-9a-fA-F]{6})`).FindStringSubmatch(bgGradient)
		if len(fc) >= 2 {
			return strings.ToUpper(fc[1]), 0
		}
		// 폴백: rgb()
		rc := regexp.MustCompile(`rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)`).FindStringSubmatch(bgGradient)
		if len(rc) >= 4 {
			r, _ := strconv.Atoi(rc[1])
			g, _ := strconv.Atoi(rc[2])
			b, _ := strconv.Atoi(rc[3])
			return fmt.Sprintf("%02X%02X%02X", r, g, b), 0
		}
	}

	// A등급: solid color
	hex, bgTrans := ParseColorAndAlpha(bgColor)
	if hex != "" {
		// 불투명 흰색은 스킵 (빈 배경), 반투명 흰색은 유효
		if hex == "FFFFFF" && bgTrans == 0 {
			// fall through to shadow check
		} else {
			finalTrans := int(math.Round(100 - float64(100-bgTrans)/100.0*opacity*100))
			if finalTrans < 0 {
				finalTrans = 0
			}
			return hex, finalTrans
		}
	}

	// shadow 있으면 흰색 배경 부여
	if hasShadow {
		finalTrans := int(math.Round(100 - 1*opacity*100))
		if finalTrans < 0 {
			finalTrans = 0
		}
		return "FFFFFF", finalTrans
	}

	return "", 0
}

// parseHex converts "#RGB" or "#RRGGBB" to (r,g,b) ints
func parseHex(h string) (int, int, int) {
	h = strings.TrimPrefix(h, "#")
	if len(h) == 3 {
		h = string([]byte{h[0], h[0], h[1], h[1], h[2], h[2]})
	}
	if len(h) < 6 {
		return 0, 0, 0
	}
	r, _ := strconv.ParseInt(h[0:2], 16, 64)
	g, _ := strconv.ParseInt(h[2:4], 16, 64)
	b, _ := strconv.ParseInt(h[4:6], 16, 64)
	return int(r), int(g), int(b)
}

// GradientStop represents a single color stop in a gradient.
type GradientStop struct {
	Position int    `json:"position"` // 0-100000 (OpenXML thousandths)
	Color    string `json:"color"`    // 6-char hex
	Alpha    int    `json:"alpha,omitempty"` // 0-100000 (0 means opaque or not parsed)
}

// GradientConfig represents an OpenXML <a:gradFill> configuration.
type GradientConfig struct {
	Angle int            `json:"angle"` // degrees (CSS convention, converted to OpenXML 60000ths)
	Stops []GradientStop `json:"stops"`
}

// TranslateGradient parses CSS linear-gradient/radial-gradient into OpenXML-ready GradientConfig.
// Returns nil if input is not a valid gradient (caller falls back to TranslateFill solid).
// Handles: deg angles, keyword directions (to right, to bottom left), rgba(), 3/6/8-char hex.
func TranslateGradient(bgGradient string) *GradientConfig {
	if bgGradient == "" {
		return nil
	}

	isLinear := strings.HasPrefix(bgGradient, "linear-gradient")
	isRadial := strings.HasPrefix(bgGradient, "radial-gradient")
	if !isLinear && !isRadial {
		return nil
	}

	// Default angle
	angle := 180 // top-to-bottom

	if isLinear {
		// 1. Try "NNdeg"
		angleRe := regexp.MustCompile(`linear-gradient\(\s*(\d+)deg`)
		am := angleRe.FindStringSubmatch(bgGradient)
		if len(am) >= 2 {
			angle, _ = strconv.Atoi(am[1])
		} else {
			// 2. Try keyword directions: "to right", "to bottom left", etc.
			kwRe := regexp.MustCompile(`linear-gradient\(\s*(to\s+[a-z\s]+)\s*,`)
			km := kwRe.FindStringSubmatch(bgGradient)
			if len(km) >= 2 {
				dir := strings.TrimSpace(km[1])
				switch dir {
				case "to right":
					angle = 90
				case "to left":
					angle = 270
				case "to bottom":
					angle = 180
				case "to top":
					angle = 0
				case "to bottom right":
					angle = 135
				case "to bottom left":
					angle = 225
				case "to top right":
					angle = 45
				case "to top left":
					angle = 315
				}
			}
		}
	}
	// radial-gradient: treat as 180° (OpenXML doesn't have native radial, render as linear fallback)

	// Extract all color stops: #hex, rgb(), or rgba() with optional position %
	stopRe := regexp.MustCompile(`(#[0-9a-fA-F]{3,8}|rgba?\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*(?:,\s*[\d.]+)?\s*\))\s*(\d+%)?`)
	matches := stopRe.FindAllStringSubmatch(bgGradient, -1)

	if len(matches) < 2 {
		return nil
	}

	stops := make([]GradientStop, 0, len(matches))
	for i, m := range matches {
		color := ""
		alpha := 100000 // default 100% opaque
		if strings.HasPrefix(m[1], "#") {
			r, g, b := parseHex(m[1])
			color = fmt.Sprintf("%02X%02X%02X", r, g, b)
		} else {
			// rgb() or rgba() — extract R,G,B, and optional alpha
			rc := regexp.MustCompile(`rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)`).FindStringSubmatch(m[1])
			if len(rc) >= 4 {
				r, _ := strconv.Atoi(rc[1])
				g, _ := strconv.Atoi(rc[2])
				b, _ := strconv.Atoi(rc[3])
				color = fmt.Sprintf("%02X%02X%02X", r, g, b)
				if len(rc) >= 5 && rc[4] != "" {
					a, _ := strconv.ParseFloat(rc[4], 64)
					alpha = int(a * 100000)
				}
			}
		}
		if color == "" {
			continue
		}

		pos := 0
		if m[2] != "" {
			pct, _ := strconv.Atoi(strings.TrimSuffix(m[2], "%"))
			pos = pct * 1000 // OpenXML: thousandths of percent
		} else {
			// Auto-distribute positions evenly
			if len(matches) > 1 {
				pos = (i * 100000) / (len(matches) - 1)
			}
		}

		stops = append(stops, GradientStop{Position: pos, Color: color, Alpha: alpha})
	}

	if len(stops) < 2 {
		return nil
	}

	return &GradientConfig{
		Angle: angle,
		Stops: stops,
	}
}


