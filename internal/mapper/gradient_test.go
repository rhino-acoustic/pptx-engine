package mapper

import (
	"testing"
)

func TestTranslateGradientDeg(t *testing.T) {
	grad := TranslateGradient("linear-gradient(135deg, #FF0000, #0000FF)")
	if grad == nil {
		t.Fatal("expected gradient, got nil")
	}
	if grad.Angle != 135 {
		t.Errorf("angle: got %d, want 135", grad.Angle)
	}
	if len(grad.Stops) != 2 {
		t.Fatalf("stops: got %d, want 2", len(grad.Stops))
	}
	if grad.Stops[0].Color != "FF0000" {
		t.Errorf("stop[0] color: got %s, want FF0000", grad.Stops[0].Color)
	}
	if grad.Stops[1].Color != "0000FF" {
		t.Errorf("stop[1] color: got %s, want 0000FF", grad.Stops[1].Color)
	}
	if grad.Stops[0].Position != 0 {
		t.Errorf("stop[0] pos: got %d, want 0", grad.Stops[0].Position)
	}
	if grad.Stops[1].Position != 100000 {
		t.Errorf("stop[1] pos: got %d, want 100000", grad.Stops[1].Position)
	}
}

func TestTranslateGradientKeyword(t *testing.T) {
	grad := TranslateGradient("linear-gradient(to right, #E00842, #FF6B35)")
	if grad == nil {
		t.Fatal("expected gradient, got nil")
	}
	if grad.Angle != 90 {
		t.Errorf("angle: got %d, want 90 (to right)", grad.Angle)
	}
	if grad.Stops[0].Color != "E00842" {
		t.Errorf("stop[0] color: got %s, want E00842", grad.Stops[0].Color)
	}
}

func TestTranslateGradientRgba(t *testing.T) {
	grad := TranslateGradient("linear-gradient(180deg, rgba(255,0,0,0.8), rgba(0,0,255,0.5))")
	if grad == nil {
		t.Fatal("expected gradient, got nil")
	}
	if len(grad.Stops) != 2 {
		t.Fatalf("stops: got %d, want 2", len(grad.Stops))
	}
	if grad.Stops[0].Color != "FF0000" {
		t.Errorf("stop[0]: got %s, want FF0000", grad.Stops[0].Color)
	}
}

func TestTranslateGradientRadial(t *testing.T) {
	grad := TranslateGradient("radial-gradient(circle, #FF0000, #00FF00, #0000FF)")
	if grad == nil {
		t.Fatal("radial should parse as fallback")
	}
	if len(grad.Stops) != 3 {
		t.Fatalf("stops: got %d, want 3", len(grad.Stops))
	}
	if grad.Angle != 180 {
		t.Errorf("radial fallback angle: got %d, want 180", grad.Angle)
	}
}

func TestTranslateGradientNil(t *testing.T) {
	if TranslateGradient("") != nil {
		t.Error("empty string should return nil")
	}
	if TranslateGradient("solid-color") != nil {
		t.Error("non-gradient should return nil")
	}
	if TranslateGradient("linear-gradient(135deg, #FF0000)") != nil {
		t.Error("single color should return nil")
	}
}

func TestMapShadowFromCSS(t *testing.T) {
	// Standard box-shadow
	s := MapShadowFromCSS("2px 4px 8px 0px rgba(0,0,0,0.3)")
	if s == nil {
		t.Fatal("expected shadow")
	}
	if s.Blur != 8 {
		t.Errorf("blur: got %d, want 8", s.Blur)
	}
	if s.Color != "000000" {
		t.Errorf("color: got %s, want 000000", s.Color)
	}

	// none
	if MapShadowFromCSS("none") != nil {
		t.Error("'none' should return nil")
	}
	if MapShadowFromCSS("") != nil {
		t.Error("empty should return nil")
	}
}

func TestTranslateGradientPositions(t *testing.T) {
	grad := TranslateGradient("linear-gradient(90deg, #FF0000 0%, #00FF00 50%, #0000FF 100%)")
	if grad == nil {
		t.Fatal("expected gradient")
	}
	if len(grad.Stops) != 3 {
		t.Fatalf("stops: got %d, want 3", len(grad.Stops))
	}
	if grad.Stops[0].Position != 0 {
		t.Errorf("stop[0] pos: got %d, want 0", grad.Stops[0].Position)
	}
	if grad.Stops[1].Position != 50000 {
		t.Errorf("stop[1] pos: got %d, want 50000", grad.Stops[1].Position)
	}
	if grad.Stops[2].Position != 100000 {
		t.Errorf("stop[2] pos: got %d, want 100000", grad.Stops[2].Position)
	}
}
