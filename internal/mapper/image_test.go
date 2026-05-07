package mapper

import (
	"testing"
)

func TestMapImageFill(t *testing.T) {
	// Dummy shape config for a pill shape
	pillShape := ShapeConfig{ShapeType: "roundRect", RectRadius: 0.5}

	tests := []struct {
		name      string
		url       string
		objectFit string
		targetW   float64
		targetH   float64
		shapeCfg  ShapeConfig
		expected  ImageConfig
	}{
		{
			name:      "Object-Fit Cover with Rounding",
			url:       "https://example.com/img.jpg",
			objectFit: "cover",
			targetW:   300,
			targetH:   60,
			shapeCfg:  pillShape,
			expected: ImageConfig{
				Path: "https://example.com/img.jpg",
				Sizing: ImageSizingConfig{
					Type: "cover",
					W:    300,
					H:    60,
				},
				Shape: pillShape,
			},
		},
		{
			name:      "Default (No specific object-fit)",
			url:       "https://example.com/logo.png",
			objectFit: "",
			targetW:   100,
			targetH:   100,
			shapeCfg:  ShapeConfig{ShapeType: "rect"},
			expected: ImageConfig{
				Path: "https://example.com/logo.png",
				Shape: ShapeConfig{ShapeType: "rect"},
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := MapImageFill(tt.url, tt.objectFit, tt.targetW, tt.targetH, tt.shapeCfg)
			if got != tt.expected {
				t.Errorf("expected %+v, got %+v", tt.expected, got)
			}
		})
	}
}
