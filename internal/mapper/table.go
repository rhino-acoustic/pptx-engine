package mapper

// TableCell represents a single cell in a PPTX table.
type TableCell struct {
	Text      string `json:"text"`
	FontSize  int    `json:"fontSize,omitempty"`  // hundredths of a point (e.g. 1400 = 14pt)
	Bold      bool   `json:"bold,omitempty"`
	Color     string `json:"color,omitempty"`     // 6-char hex text color
	FillColor string `json:"fillColor,omitempty"` // 6-char hex background
	Align     string `json:"align,omitempty"`     // "l", "ctr", "r"
	ColSpan   int    `json:"colSpan,omitempty"`
	RowSpan   int    `json:"rowSpan,omitempty"`
}

// TableConfig represents an OpenXML <a:tbl> configuration.
type TableConfig struct {
	Rows     [][]TableCell `json:"rows"`
	ColWidths []int64      `json:"colWidths,omitempty"` // EMU per column
	RowHeights []int64     `json:"rowHeights,omitempty"` // EMU per row
	HasHeader bool         `json:"hasHeader,omitempty"`
}
