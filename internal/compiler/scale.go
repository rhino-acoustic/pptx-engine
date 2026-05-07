package compiler

import (
	"math"
)

// SlideConfig defines the target PPT slide dimensions and desired margins (in inches).
type SlideConfig struct {
	Width       float64
	Height      float64
	MarginX     float64 // Minimum horizontal margin
	MarginY     float64 // Minimum vertical margin
	BaseScaling float64 // Optional multiplier
}

// AutoScaleElements recalculates X, Y, W, H and crucially, Font Sizes and Internal Margins,
// completely bypassing the PPTX text box margin scaling bugs.
func AutoScaleElements(nodes []PptxElement, cfg SlideConfig) []PptxElement {
	if len(nodes) == 0 {
		return nodes
	}

	minX, minY := math.MaxFloat64, math.MaxFloat64
	maxX, maxY := -math.MaxFloat64, -math.MaxFloat64

	for _, n := range nodes {
		if n.X < minX {
			minX = n.X
		}
		if n.Y < minY {
			minY = n.Y
		}
		if n.X+n.W > maxX {
			maxX = n.X + n.W
		}
		if n.Y+n.H > maxY {
			maxY = n.Y + n.H
		}
	}

	origWidth := maxX - minX
	origHeight := maxY - minY

	if origWidth <= 0 || origHeight <= 0 {
		return nodes
	}

	availW := cfg.Width - (cfg.MarginX * 2)
	availH := cfg.Height - (cfg.MarginY * 2)

	scaleW := availW / origWidth
	scaleH := availH / origHeight
	scale := math.Min(scaleW, scaleH)
	
	if cfg.BaseScaling > 0 {
		scale *= cfg.BaseScaling
	}

	finalW := origWidth * scale
	finalH := origHeight * scale

	offsetX := cfg.MarginX + (availW-finalW)/2.0
	offsetY := cfg.MarginY + (availH-finalH)/2.0

	scaledNodes := make([]PptxElement, len(nodes))
	for i, n := range nodes {
		scaledNodes[i] = n
		scaledNodes[i].X = ((n.X - minX) * scale) + offsetX
		scaledNodes[i].Y = ((n.Y - minY) * scale) + offsetY
		scaledNodes[i].W = n.W * scale
		scaledNodes[i].H = n.H * scale

		// PptxGenJS Bug Fix: Scale internal padding/margins & font sizes dynamically
		// PPTX text margins are usually [t, r, b, l] in points.
		// If Text properties exist, scale them down proportionally to prevent overflow drift.
		if n.TextConfig != nil {
			// Deep copy to prevent mutation of shared AST
			tc := *n.TextConfig
			tc.FontSize = tc.FontSize * scale
			
			// Typography ratio protection (Presentation minimal legibility bounds)
			if tc.FontSize < 8.0 {
				tc.FontSize = 8.0 // Force readable minimum
			}
			
			// Map internal margins dynamically [T, R, B, L] based on element scale
			scaledPad := 2.0 * scale // 2pt base margin scaled
			tc.Margin = []float64{scaledPad, scaledPad * 2, scaledPad, scaledPad * 2}
			
			scaledNodes[i].TextConfig = &tc
		}
	}

	return scaledNodes
}
