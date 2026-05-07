package compiler

import "testing"

func TestMapFontFamilyKorean(t *testing.T) {
	cases := []struct{ input, want string }{
		{"Pretendard", "Pretendard"},
		{"Noto Sans KR", "Noto Sans KR"},
		{"nanum gothic", "나눔고딕"},
		{"NanumMyeongjo", "나눔명조"},
		{"맑은 고딕", "맑은 고딕"},
		{"Malgun Gothic", "맑은 고딕"},
		{"돋움", "돋움"},
		{"Dotum", "돋움"},
		{"굴림", "굴림"},
		{"Gulim", "굴림"},
		{"바탕", "바탕"},
		{"Batang", "바탕"},
		{"Apple SD Gothic", "Noto Sans KR"},
		{"Spoqa Han Sans", "Noto Sans KR"},
		{"SUIT Variable", "Noto Sans KR"},
	}
	for _, c := range cases {
		got := MapFontFamily(c.input)
		if got != c.want {
			t.Errorf("MapFontFamily(%q) = %q, want %q", c.input, got, c.want)
		}
	}
}

func TestMapFontFamilyEnglish(t *testing.T) {
	cases := []struct{ input, want string }{
		{"Montserrat", "Montserrat"},
		{"Poppins", "Poppins"},
		{"DM Sans", "DM Sans"},
		{"Outfit", "Outfit"},
		{"Raleway", "Raleway"},
		{"Playfair Display", "Playfair Display"},
		{"Georgia", "Georgia"},
		{"Times New Roman", "Times New Roman"},
		{"Courier New", "Consolas"},
		{"monospace", "Consolas"},
		{"Roboto", "Arial"},
		{"Inter", "Arial"},
		{"Segoe UI", "Arial"},
		{"Helvetica Neue", "Arial"},
		{"Open Sans", "Arial"},
		{"Lato", "Arial"},
		{"Ubuntu", "Arial"},
		{"Nunito", "Arial"},
	}
	for _, c := range cases {
		got := MapFontFamily(c.input)
		if got != c.want {
			t.Errorf("MapFontFamily(%q) = %q, want %q", c.input, got, c.want)
		}
	}
}

func TestMapFontFamilyFallback(t *testing.T) {
	cases := []struct{ input, want string }{
		{"sans-serif", "Noto Sans KR"},
		{"system-ui", "Noto Sans KR"},
		{"-apple-system, BlinkMacSystemFont, sans-serif", "Noto Sans KR"},
		{"unknown-font-xyz", "Noto Sans KR"},
	}
	for _, c := range cases {
		got := MapFontFamily(c.input)
		if got != c.want {
			t.Errorf("MapFontFamily(%q) = %q, want %q", c.input, got, c.want)
		}
	}
}
