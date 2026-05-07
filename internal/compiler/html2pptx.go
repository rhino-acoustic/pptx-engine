package compiler

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/rhino-acoustic/pptx-engine/internal/mapper"
)

// ParsedShapeAdapter is a minimal adapter interface matching reverse.ParsedShape fields
// to avoid circular imports. We use a struct instead.
type HTMLShape struct {
	Type            string
	X, Y, W, H     float64
	Rotation        float64
	ShapeType       string
	FillType        string
	FillColor       string
	FillTransparency float64
	GradientAngle   float64
	GradientStops   []HTMLGradStop
	BorderColor     string
	BorderWidth     float64
	BorderRadius    float64
	HasShadow       bool
	Valign          string
	PadL, PadR      float64
	PadT, PadB      float64
	ImagePath       string
	ImageRId        string
	CropL, CropT    float64
	CropR, CropB    float64
	HasText         bool
	TextRuns        []HTMLTextRun
	TableRows       int
	TableCols       int
	TableColWidths  []float64
	TableRowHeights []float64
	TableData       []json.RawMessage
	ChartSVG        string
}

type HTMLGradStop struct {
	Color    string
	Position int
	Alpha    float64
}

type HTMLTextRun struct {
	Text      string
	Font      string
	Size      float64
	Bold      int
	Italic    int
	Underline int
	Color     string
	Align     int
	Bullet    int
}

// CompileHTMLSlide converts parsed HTML shapes into PptxElements ready for OpenXML building.
func CompileHTMLSlide(shapes []HTMLShape) []PptxElement {
	elements := make([]PptxElement, 0, len(shapes))

	for si, s := range shapes {
		el := PptxElement{
			X:        s.X / 72.0,  // pt → inches
			Y:        s.Y / 72.0,
			W:        s.W / 72.0,
			H:        s.H / 72.0,
			Rotation: s.Rotation,
			ZIndex:   si + 1,
		}

		// Determine element type
		if s.ImagePath != "" && !s.HasText {
			el.Type = "image"
			el.ImageRId = s.ImageRId
			el.Shape = mapper.ShapeConfig{ShapeType: s.ShapeType}
		} else if s.TableRows > 0 && len(s.TableData) > 0 {
			el.Type = "table"
			el.Table = compileHTMLTable(s, el.W, el.H)
		} else if s.HasText && len(s.TextRuns) > 0 {
			el.Type = "text"
			// Combine text runs
			var textParts []string
			for _, r := range s.TextRuns {
				textParts = append(textParts, r.Text)
			}
			el.Text = strings.Join(textParts, "")

			// Use first run for primary formatting
			first := s.TextRuns[0]
			tc := TextConfig{
				FontFace: first.Font,
				FontSize: first.Size,       // already in pt
				Color:    cleanHexColor(first.Color),
				Bold:     first.Bold == -1,
				Italic:   first.Italic == -1,
				Wrap:     true,
			}

			// Align
			switch first.Align {
			case 2: tc.Align = "center"
			case 3: tc.Align = "right"
			default: tc.Align = "left"
			}

			// Valign
			tc.Valign = "top"
			if s.Valign == "ctr" { tc.Valign = "middle" }
			if s.Valign == "b" { tc.Valign = "bottom" }

			el.TextConfig = &tc
		} else {
			el.Type = "shape"
			el.Shape = mapper.ShapeConfig{ShapeType: s.ShapeType}
			if s.BorderRadius > 0 {
				el.Shape.ShapeType = "roundRect"
				el.Shape.RectRadius = s.BorderRadius
			}
		}

		// Fill
		if s.FillType == "solid" && s.FillColor != "" {
			hex := cleanHexColor(s.FillColor)
			el.Fill = map[string]interface{}{"color": hex}
			if s.FillTransparency > 0 {
				el.Fill["transparency"] = s.FillTransparency
			}
		} else if s.FillType == "gradient" && len(s.GradientStops) >= 2 {
			stops := make([]mapper.GradientStop, 0, len(s.GradientStops))
			for _, gs := range s.GradientStops {
				stops = append(stops, mapper.GradientStop{
					Color:    cleanHexColor(gs.Color),
					Position: gs.Position,
					Alpha:    int(gs.Alpha * 100000),
				})
			}
			el.Gradient = &mapper.GradientConfig{
				Angle: int(s.GradientAngle),
				Stops: stops,
			}
		}

		// Border
		if s.BorderWidth > 0 && s.BorderColor != "" {
			el.Border = &mapper.BorderConfig{
				Color: cleanHexColor(s.BorderColor),
				Width: s.BorderWidth,
			}
		}

		// Shadow
		if s.HasShadow {
			el.Shadow = mapper.MapShadow(true)
		}

		elements = append(elements, el)
	}

	return elements
}

func compileHTMLTable(s HTMLShape, elW, elH float64) *mapper.TableConfig {
	tc := &mapper.TableConfig{
		HasHeader: true,
	}

	// Column widths in EMU
	for _, cw := range s.TableColWidths {
		tc.ColWidths = append(tc.ColWidths, int64(cw/72.0*914400))
	}

	// Row heights in EMU
	for _, rh := range s.TableRowHeights {
		tc.RowHeights = append(tc.RowHeights, int64(rh/72.0*914400))
	}

	// Parse table data
	for _, rowRaw := range s.TableData {
		var cells []struct {
			Text      string  `json:"text"`
			Color     string  `json:"color"`
			Bold      int     `json:"bold"`
			Size      float64 `json:"size"`
			Bg        string  `json:"bg"`
			Align     string  `json:"align"`
			VAlign    string  `json:"valign"`
			MergeSkip int     `json:"mergeSkip"`
			VertMerge int     `json:"vertMerge"`
		}
		json.Unmarshal(rowRaw, &cells)

		var row []mapper.TableCell
		for _, c := range cells {
			cell := mapper.TableCell{
				Text:      c.Text,
				Color:     cleanHexColor(c.Color),
				Bold:      c.Bold == -1,
				FontSize:  int(c.Size * 100),
				FillColor: cleanHexColor(c.Bg),
				Align:     c.Align,
				ColSpan:   0,
			}
			if cell.FontSize == 0 { cell.FontSize = 1200 }
			if cell.Color == "" { cell.Color = "000000" }
			if cell.FillColor == "" { cell.FillColor = "FFFFFF" }
			if cell.Align == "" { cell.Align = "l" }
			row = append(row, cell)
		}
		tc.Rows = append(tc.Rows, row)
	}

	return tc
}

// cleanHexColor normalizes color strings to 6-char hex without #
func cleanHexColor(c string) string {
	c = strings.TrimSpace(c)
	c = strings.TrimPrefix(c, "#")
	if len(c) == 3 {
		c = fmt.Sprintf("%c%c%c%c%c%c", c[0], c[0], c[1], c[1], c[2], c[2])
	}
	if len(c) > 6 { c = c[:6] }
	return strings.ToUpper(c)
}
