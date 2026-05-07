package mapper

import (
	"testing"
)

func TestMapShapeRadius(t *testing.T) {
	tests := []struct {
		name      string
		w         float64
		h         float64
		brCorners []float64
		expected  ShapeConfig
	}{
		{
			name:      "Perfect Circle",
			w:         200,
			h:         200,
			brCorners: []float64{100, 100, 100, 100},
			expected:  ShapeConfig{ShapeType: "ellipse"},
		},
		{
			name:      "Pill Shape",
			w:         300,
			h:         60,
			brCorners: []float64{9999, 9999, 9999, 9999},
			expected:  ShapeConfig{ShapeType: "roundRect", RectRadius: 0.5},
		},
		{
			name:      "One Corner Rounded (TL)",
			w:         100,
			h:         100,
			brCorners: []float64{20, 0, 0, 0},
			expected:  ShapeConfig{ShapeType: "round1Rect"},
		},
		{
			name:      "One Corner Rounded (TR)",
			w:         100,
			h:         100,
			brCorners: []float64{0, 20, 0, 0},
			expected:  ShapeConfig{ShapeType: "round1Rect", FlipH: true},
		},
		{
			name:      "Two Same Corners Rounded (Top)",
			w:         100,
			h:         100,
			brCorners: []float64{20, 20, 0, 0},
			expected:  ShapeConfig{ShapeType: "round2SameRect"},
		},
		{
			name:      "Two Diag Corners Rounded (TL, BR)",
			w:         100,
			h:         100,
			brCorners: []float64{20, 0, 20, 0},
			expected:  ShapeConfig{ShapeType: "round2DiagRect"},
		},
		{
			name:      "Sharp Rectangle",
			w:         100,
			h:         100,
			brCorners: []float64{0, 0, 0, 0},
			expected:  ShapeConfig{ShapeType: "rect"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := MapShapeRadius(tt.w, tt.h, tt.brCorners)
			if got != tt.expected {
				t.Errorf("expected %+v, got %+v", tt.expected, got)
			}
		})
	}
}
