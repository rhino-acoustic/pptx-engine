package mapper

import (
	"testing"
)

func TestMapBorder(t *testing.T) {
	tests := []struct {
		name        string
		color       string
		style       string
		width       float64
		expectNil   bool
		expectedDsh string
		expectedWdt float64
	}{
		{"Valid solid border", "rgb(255, 0, 0)", "solid", 5.0, false, "", 2.0},
		{"Dashed border", "#00FF00", "dashed", 2.0, false, "dash", 0.8},
		{"Dotted border", "rgba(0, 0, 255, 0.5)", "dotted", 1.0, false, "sysDot", 0.4},
		{"Transparent border", "transparent", "solid", 2.0, true, "", 0},
		{"Zero width border", "#FFFFFF", "solid", 0, true, "", 0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := MapBorder(tt.color, tt.style, tt.width)
			if tt.expectNil {
				if got != nil {
					t.Errorf("expected nil, got %+v", got)
				}
				return
			}
			if got == nil {
				t.Fatalf("expected border config, got nil")
			}
			if got.DashType != tt.expectedDsh {
				t.Errorf("expected DashType %q, got %q", tt.expectedDsh, got.DashType)
			}
			if got.Width != tt.expectedWdt {
				t.Errorf("expected Width %v, got %v", tt.expectedWdt, got.Width)
			}
		})
	}
}

func TestParseColorAndAlpha(t *testing.T) {
	tests := []struct {
		input    string
		expected string
		trans    int
	}{
		{"rgb(255, 128, 0)", "FF8000", 0},
		{"rgba(0, 255, 0, 0.5)", "00FF00", 50},
		{"rgba(0, 0, 0, 0)", "", 100},
		{"#ABC", "AABBCC", 0},
		{"#123456", "123456", 0},
		{"#12345678", "123456", 53}, // 0x78 = 120, 120/255 = 0.47 -> 1-0.47 = 0.53 -> 53
		{"transparent", "", 100},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got, trans := ParseColorAndAlpha(tt.input)
			if got != tt.expected {
				t.Errorf("expected hex %q, got %q", tt.expected, got)
			}
			if trans != tt.trans {
				t.Errorf("expected transparency %d, got %d", tt.trans, trans)
			}
		})
	}
}
