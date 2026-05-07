package mapper

import (
	"math"
)

// ShapeConfig represents the mapped PPTX shape attributes.
type ShapeConfig struct {
	ShapeType  string  `json:"shapeType"`
	RectRadius float64 `json:"rectRadius,omitempty"`
	FlipH      bool    `json:"flipH,omitempty"`
	FlipV      bool    `json:"flipV,omitempty"`
	Rotate     float64 `json:"rotate,omitempty"`
}

// MapShapeRadius maps CSS border-radius array to PPTX shape attributes.
// brCorners: [TopLeft, TopRight, BottomRight, BottomLeft]
func MapShapeRadius(w, h float64, brCorners []float64) ShapeConfig {
	cfg := ShapeConfig{ShapeType: "rect"}
	if len(brCorners) != 4 {
		return cfg
	}

	var maxR float64
	for _, r := range brCorners {
		if r > maxR {
			maxR = r
		}
	}

	// Pixel based isCircle logic: max radius >= min(w, h) * 0.45
	isCircle := maxR >= math.Min(w, h)*0.45

	if isCircle && math.Abs(w-h) < math.Max(w, h)*0.2 {
		cfg.ShapeType = "ellipse"
		return cfg
	}

	if maxR <= 0 {
		return cfg
	}

	rThreshold := math.Max(2.0, maxR*0.5)
	rounded := make([]bool, 4)
	count := 0
	for i, r := range brCorners {
		if r > rThreshold {
			rounded[i] = true
			count++
		}
	}

	// Clamp maxR to half of the shorter side (PPTX max roundRect radius)
	halfMin := math.Min(w, h) / 2
	if maxR > halfMin {
		maxR = halfMin
	}

	switch count {
	case 4:
		cfg.ShapeType = "roundRect"
		if maxR > 0 && halfMin > 0 {
			cfg.RectRadius = maxR / math.Min(w, h) // normalize to 0..0.5 (PPTX range)
		}
	case 1:
		cfg.ShapeType = "round1Rect"
		if rounded[1] {
			cfg.FlipH = true
		} else if rounded[2] {
			cfg.FlipH = true
			cfg.FlipV = true
		} else if rounded[3] {
			cfg.FlipV = true
		}
	case 2:
		if (rounded[0] && rounded[2]) || (rounded[1] && rounded[3]) {
			cfg.ShapeType = "round2DiagRect"
			if rounded[1] {
				cfg.FlipH = true
			}
		} else {
			cfg.ShapeType = "round2SameRect"
			if rounded[1] && rounded[2] {
				cfg.Rotate = 90
			} else if rounded[2] && rounded[3] {
				cfg.FlipV = true
			} else if rounded[3] && rounded[0] {
				cfg.Rotate = 270
			}
		}
	case 3:
		cfg.ShapeType = "roundRect" // Fallback
	}

	return cfg
}
