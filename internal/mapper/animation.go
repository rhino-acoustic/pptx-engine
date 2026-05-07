package mapper

import (
	"strings"
)

// AnimationConfig defines PPTX slide element animation properties.
type AnimationConfig struct {
	Type  string  `json:"type"`  // "fade", "fly", "appear"
	Delay float64 `json:"delay"` // seconds
	Speed string  `json:"speed"` // "fast", "medium", "slow"
}

// MapAnimation converts CSS animation attributes into PPTX animations.
// It also infers a cascading delay based on the Y-coordinate to recreate natural web scroll reveals.
func MapAnimation(cssAnimationName string, yPosition float64) *AnimationConfig {
	if cssAnimationName == "" || cssAnimationName == "none" {
		return nil
	}

	cfg := &AnimationConfig{
		Type:  "fade",
		Speed: "medium",
	}

	name := strings.ToLower(cssAnimationName)

	if strings.Contains(name, "fly") || strings.Contains(name, "up") || strings.Contains(name, "down") {
		cfg.Type = "fly"
	} else if strings.Contains(name, "appear") || strings.Contains(name, "zoom") {
		cfg.Type = "appear"
	}

	// PptxGenJS issue workaround: Synchronous overlapping animations often glitch.
	// We create a cascading effect based on Y position (elements lower on the page appear later).
	// Y is assumed to be in inches or normalized coordinates.
	// 1 inch down roughly adds 0.2s delay.
	cfg.Delay = 0.5 + (yPosition * 0.2)

	return cfg
}
