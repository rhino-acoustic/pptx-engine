package reverse

import (
	"archive/zip"
	"fmt"
	"io"
	"math"
	"regexp"
	"strconv"
	"strings"
)

// parseChartFrame handles chart graphicFrame elements
func parseChartFrame(gfXML string, id *int, fileMap map[string]*zip.File, slideIdx int) *ParsedShape {
	shp := &ParsedShape{ID: *id, Type: "CHART"}
	*id++

	if m := regexp.MustCompile(`name="([^"]+)"`).FindStringSubmatch(gfXML); m != nil {
		shp.Name = m[1]
	}

	parseXfrm(gfXML, shp)

	// Extract chart relationship ID
	rIdRe := regexp.MustCompile(`<c:chart[^>]*r:id="([^"]+)"`)
	rIdM := rIdRe.FindStringSubmatch(gfXML)
	if rIdM == nil {
		return shp
	}
	rId := rIdM[1]

	if fileMap == nil {
		return shp
	}

	// Find slide rels file
	slideRelsPath := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideIdx)
	relsFile, ok := fileMap[slideRelsPath]
	if !ok {
		return shp
	}

	relsData, err := readZipFileChart(relsFile)
	if err != nil {
		return shp
	}

	// Find chart target from rels
	targetRe := regexp.MustCompile(`Id="` + rId + `"[^>]*Target="([^"]*)"`)
	targetM := targetRe.FindStringSubmatch(string(relsData))
	if targetM == nil {
		return shp
	}

	// Resolve chart path (Target is like "../charts/chart10.xml")
	chartTarget := targetM[1]
	chartPath := "ppt/" + strings.TrimPrefix(chartTarget, "../")

	chartFile, ok := fileMap[chartPath]
	if !ok {
		return shp
	}

	chartData, err := readZipFileChart(chartFile)
	if err != nil {
		return shp
	}

	// Parse chart XML and render SVG
	cd := parseChartXMLContent(string(chartData))
	if cd != nil && len(cd.Series) > 0 {
		shp.ChartSVG = renderChartSVG(cd, shp.W, shp.H)
	}

	return shp
}

func readZipFileChart(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

// ChartData holds parsed chart information for SVG rendering
type ChartData struct {
	Type       string // "bar", "line", "combo", "pie"
	Categories []string
	Series     []ChartSeries
	ValAxes    []ChartAxis
	CatAxes    []ChartAxis
	HasLegend  bool
	LegendPos  string
}

type ChartSeries struct {
	Name       string
	ChartType  string // "bar" or "line"
	Values     []float64
	Labels     []string
	Color      string
	LabelColor string
	ShowVal    bool
}

type ChartAxis struct {
	ID       string
	Position string // "l", "r", "b", "t"
	Delete   bool
	NumFmt   string
}

// parseChartXMLContent parses chart XML content string
func parseChartXMLContent(content string) *ChartData {

	cd := &ChartData{}

	// Detect chart types
	hasBar := strings.Contains(content, "<c:barChart>") || strings.Contains(content, "<c:barChart ")
	hasLine := strings.Contains(content, "<c:lineChart>") || strings.Contains(content, "<c:lineChart ")
	hasPie := strings.Contains(content, "<c:pieChart>") || strings.Contains(content, "<c:pieChart ")
	hasDoughnut := strings.Contains(content, "<c:doughnutChart>")

	if hasBar && hasLine {
		cd.Type = "combo"
	} else if hasBar {
		cd.Type = "bar"
	} else if hasLine {
		cd.Type = "line"
	} else if hasPie || hasDoughnut {
		cd.Type = "pie"
	} else {
		cd.Type = "bar" // default
	}

	// Parse series from barChart
	if hasBar {
		barBlock := extractBlock(content, "c:barChart")
		cd.Series = append(cd.Series, parseSeriesBlocks(barBlock, "bar")...)
	}

	// Parse series from lineChart
	if hasLine {
		lineBlock := extractBlock(content, "c:lineChart")
		cd.Series = append(cd.Series, parseSeriesBlocks(lineBlock, "line")...)
	}

	// Parse series from pieChart
	if hasPie {
		pieBlock := extractBlock(content, "c:pieChart")
		cd.Series = append(cd.Series, parseSeriesBlocks(pieBlock, "bar")...) // render as bars for simplicity
	}

	// Parse categories from first series that has them
	for _, s := range cd.Series {
		if len(s.Labels) > 0 && len(cd.Categories) == 0 {
			cd.Categories = s.Labels
		}
	}

	// Legend
	if strings.Contains(content, "<c:legend>") {
		cd.HasLegend = true
		if m := regexp.MustCompile(`<c:legendPos val="([^"]+)"`).FindStringSubmatch(content); m != nil {
			cd.LegendPos = m[1]
		}
	}

	return cd
}

func extractBlock(content, tag string) string {
	open := "<" + tag + ">"
	altOpen := "<" + tag + " "
	close := "</" + tag + ">"

	start := strings.Index(content, open)
	if start == -1 {
		start = strings.Index(content, altOpen)
	}
	if start == -1 {
		return ""
	}
	end := strings.Index(content[start:], close)
	if end == -1 {
		return ""
	}
	return content[start : start+end+len(close)]
}

func parseSeriesBlocks(block, chartType string) []ChartSeries {
	var series []ChartSeries
	serRe := regexp.MustCompile(`(?s)<c:ser>(.*?)</c:ser>`)
	for _, m := range serRe.FindAllStringSubmatch(block, -1) {
		s := parseSingleSeries(m[1], chartType)
		series = append(series, s)
	}
	return series
}

func parseSingleSeries(serXML, chartType string) ChartSeries {
	s := ChartSeries{ChartType: chartType}

	// Series name
	txRe := regexp.MustCompile(`(?s)<c:tx>.*?<c:v>(.*?)</c:v>`)
	if m := txRe.FindStringSubmatch(serXML); m != nil {
		s.Name = strings.TrimSpace(m[1])
	}

	// Color
	fillRe := regexp.MustCompile(`<c:spPr>.*?<a:solidFill>.*?<a:srgbClr val="([A-Fa-f0-9]{6})"`)
	if m := fillRe.FindStringSubmatch(serXML); m != nil {
		s.Color = "#" + m[1]
	}
	// scheme color fallback
	if s.Color == "" {
		schemeRe := regexp.MustCompile(`<c:spPr>.*?<a:solidFill>.*?<a:schemeClr val="([^"]+)"`)
		if m := schemeRe.FindStringSubmatch(serXML); m != nil {
			s.Color = schemeColorToHex(m[1])
		}
	}
	// Line color
	if chartType == "line" && s.Color == "" {
		lineRe := regexp.MustCompile(`<a:ln[^>]*>.*?<a:solidFill>.*?<a:srgbClr val="([A-Fa-f0-9]{6})"`)
		if m := lineRe.FindStringSubmatch(serXML); m != nil {
			s.Color = "#" + m[1]
		}
	}
	if s.Color == "" {
		if chartType == "bar" {
			s.Color = "#999999"
		} else {
			s.Color = "#0CB488"
		}
	}

	// Label color
	lblClrRe := regexp.MustCompile(`<c:dLbls>.*?<c:txPr>.*?<a:solidFill>.*?<a:srgbClr val="([A-Fa-f0-9]{6})"`)
	if m := lblClrRe.FindStringSubmatch(serXML); m != nil {
		s.LabelColor = "#" + m[1]
	} else {
		s.LabelColor = "#555555"
	}

	// Show values
	if strings.Contains(serXML, `<c:showVal val="1"`) {
		s.ShowVal = true
	}

	// Category labels — strCache first, then numCache fallback
	catRe := regexp.MustCompile(`(?s)<c:cat>.*?<c:strCache>(.*?)</c:strCache>`)
	if m := catRe.FindStringSubmatch(serXML); m != nil {
		ptRe := regexp.MustCompile(`<c:v>(.*?)</c:v>`)
		for _, pt := range ptRe.FindAllStringSubmatch(m[1], -1) {
			s.Labels = append(s.Labels, strings.TrimSpace(pt[1]))
		}
	} else {
		// numCache fallback (e.g. year numbers as categories)
		catNumRe := regexp.MustCompile(`(?s)<c:cat>.*?<c:numCache>(.*?)</c:numCache>`)
		if m := catNumRe.FindStringSubmatch(serXML); m != nil {
			ptRe := regexp.MustCompile(`<c:v>(.*?)</c:v>`)
			for _, pt := range ptRe.FindAllStringSubmatch(m[1], -1) {
				val := strings.TrimSpace(pt[1])
				// Convert float-like year strings to clean integers
				if f, err := strconv.ParseFloat(val, 64); err == nil && f > 1900 && f < 2100 {
					s.Labels = append(s.Labels, fmt.Sprintf("%.0f", f))
				} else {
					s.Labels = append(s.Labels, val)
				}
			}
		}
	}

	// Values
	valRe := regexp.MustCompile(`(?s)<c:val>.*?<c:numCache>(.*?)</c:numCache>`)
	if m := valRe.FindStringSubmatch(serXML); m != nil {
		ptRe := regexp.MustCompile(`<c:v>(.*?)</c:v>`)
		for _, pt := range ptRe.FindAllStringSubmatch(m[1], -1) {
			v, _ := strconv.ParseFloat(pt[1], 64)
			s.Values = append(s.Values, v)
		}
		// Number format — check val's own formatCode
		fmtRe := regexp.MustCompile(`<c:formatCode>(.*?)</c:formatCode>`)
		if fm := fmtRe.FindStringSubmatch(m[1]); fm != nil {
			if strings.Contains(fm[1], "%") {
				// Values are fractions, convert to percentage display
				for i := range s.Values {
					s.Values[i] = s.Values[i] * 100
				}
			}
		} else if chartType == "line" && len(s.Values) > 0 {
			// Heuristic: if all values are in -1..1 range, treat as percentage
			allFraction := true
			for _, v := range s.Values {
				if v < -1.5 || v > 1.5 {
					allFraction = false
					break
				}
			}
			if allFraction {
				for i := range s.Values {
					s.Values[i] = s.Values[i] * 100
				}
			}
		}
	}

	return s
}

func schemeColorToHex(scheme string) string {
	colors := map[string]string{
		"bg1":   "#A6A6A6",
		"tx1":   "#404040",
		"bg2":   "#EEECE1",
		"tx2":   "#1F497D",
		"accent1": "#4472C4",
		"accent2": "#ED7D31",
		"accent3": "#A5A5A5",
		"accent4": "#FFC000",
		"accent5": "#5B9BD5",
		"accent6": "#70AD47",
	}
	if c, ok := colors[scheme]; ok {
		return c
	}
	return "#999999"
}

// renderChartSVG generates an SVG string for the chart
func renderChartSVG(cd *ChartData, width, height float64) string {
	if cd == nil || len(cd.Series) == 0 {
		return ""
	}

	// Margins — proportional to chart size
	marginL := math.Min(60.0, width*0.12)
	marginR := math.Min(60.0, width*0.12)
	marginT := math.Min(20.0, height*0.06)
	marginB := math.Min(60.0, height*0.15)
	if cd.HasLegend && cd.LegendPos == "b" {
		marginB = math.Min(80.0, height*0.2)
	}

	// Check if axes are deleted (hidden) → reduce margins
	leftAxisHidden := false
	rightAxisHidden := true // default: no right axis
	for _, ax := range cd.ValAxes {
		if ax.Position == "l" && ax.Delete {
			leftAxisHidden = true
		}
		if ax.Position == "r" {
			rightAxisHidden = ax.Delete
		}
	}
	if leftAxisHidden {
		marginL = math.Min(15.0, width*0.04)
	}
	if rightAxisHidden {
		marginR = math.Min(15.0, width*0.04)
	}

	// Font scaling for small charts
	baseFontSize := 11.0
	labelFontSize := 10.0
	axisFontSize := 10.0
	if height < 200 {
		baseFontSize = math.Max(7, height/25)
		labelFontSize = math.Max(6, height/28)
		axisFontSize = math.Max(6, height/30)
	}

	plotW := width - marginL - marginR
	plotH := height - marginT - marginB

	var svg strings.Builder
	svg.WriteString(fmt.Sprintf(`<svg xmlns="http://www.w3.org/2000/svg" width="100%%" height="100%%" viewBox="0 0 %.0f %.0f" preserveAspectRatio="xMidYMid meet" style="font-family:'Malgun Gothic',sans-serif;font-size:%.0fpx;">`, width, height, baseFontSize))

	// Find min/max values for bar series (left axis)
	maxBarVal := 0.0
	minBarVal := 0.0
	for _, s := range cd.Series {
		if s.ChartType == "bar" {
			for _, v := range s.Values {
				if v > maxBarVal {
					maxBarVal = v
				}
				if v < minBarVal {
					minBarVal = v
				}
			}
		}
	}
	if maxBarVal == 0 && minBarVal == 0 {
		maxBarVal = 100
	}
	// Round up/down to nice numbers
	maxBarVal = niceMax(maxBarVal)
	if minBarVal < 0 {
		minBarVal = -niceMax(-minBarVal)
	}
	barRange := maxBarVal - minBarVal

	// Find min/max value for line series (right axis, percentage)
	maxLineVal := 0.0
	minLineVal := 0.0
	hasLine := false
	for _, s := range cd.Series {
		if s.ChartType == "line" {
			hasLine = true
			for _, v := range s.Values {
				if v > maxLineVal {
					maxLineVal = v
				}
				if v < minLineVal {
					minLineVal = v
				}
			}
		}
	}
	if maxLineVal == 0 && minLineVal == 0 {
		maxLineVal = 100
	}
	maxLineVal = niceMax(maxLineVal)
	if minLineVal < 0 {
		minLineVal = -niceMax(-minLineVal)
	}
	lineRange := maxLineVal - minLineVal

	nCats := len(cd.Categories)
	if nCats == 0 {
		nCats = 1
	}
	catW := plotW / float64(nCats)

	// Left axis gridlines & labels
	nTicks := 5
	if height > 250 {
		nTicks = 7
	}
	for i := 0; i <= nTicks; i++ {
		val := minBarVal + barRange*float64(i)/float64(nTicks)
		y := marginT + plotH - (plotH * float64(i) / float64(nTicks))
		svg.WriteString(fmt.Sprintf(`<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#E0E0E0" stroke-width="0.5"/>`, marginL, y, width-marginR, y))
		if !leftAxisHidden {
			label := formatAxisVal(val)
			svg.WriteString(fmt.Sprintf(`<text x="%.1f" y="%.1f" text-anchor="end" fill="#777" font-size="%.0f">%s</text>`, marginL-3, y+3, axisFontSize, label))
		}
	}

	// Right axis labels (if line series exists)
	if hasLine {
		for i := 0; i <= nTicks; i++ {
			val := minLineVal + lineRange*float64(i)/float64(nTicks)
			y := marginT + plotH - (plotH * float64(i) / float64(nTicks))
			svg.WriteString(fmt.Sprintf(`<text x="%.1f" y="%.1f" text-anchor="start" fill="#777" font-size="%.0f">%.1f%%</text>`, width-marginR+3, y+3, axisFontSize, val))
		}
	}

	// Category labels
	for i, cat := range cd.Categories {
		x := marginL + float64(i)*catW + catW/2
		y := marginT + plotH + baseFontSize + 4
		svg.WriteString(fmt.Sprintf(`<text x="%.1f" y="%.1f" text-anchor="middle" fill="#555" font-size="%.0f" font-weight="bold">%s</text>`, x, y, baseFontSize, xmlEscape(cat)))
	}

	// Count bar series for grouping
	barSeriesCount := 0
	for _, s := range cd.Series {
		if s.ChartType == "bar" {
			barSeriesCount++
		}
	}

	// Bar rendering
	gap := catW * 0.15
	barAreaW := catW - gap*2
	barW := barAreaW / float64(barSeriesCount)
	if barSeriesCount == 0 {
		barW = barAreaW
	}
	barIdx := 0
	for _, s := range cd.Series {
		if s.ChartType != "bar" {
			continue
		}
		for i, v := range s.Values {
			if i >= nCats {
				break
			}
			// Position relative to full range (supports negative values)
			valFrac := (v - minBarVal) / barRange
			zeroFrac := (0 - minBarVal) / barRange
			x := marginL + float64(i)*catW + gap + float64(barIdx)*barW
			if v >= 0 {
				barH := (valFrac - zeroFrac) * plotH
				y := marginT + plotH*(1-valFrac)
				if barH < 0.5 { barH = 0.5 }
				svg.WriteString(fmt.Sprintf(`<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" rx="1"/>`,
					x, y, barW*0.85, barH, s.Color))
				if s.ShowVal {
					label := formatBarVal(v)
					svg.WriteString(fmt.Sprintf(`<text x="%.1f" y="%.1f" text-anchor="middle" fill="%s" font-size="%.0f" font-weight="bold">%s</text>`,
						x+barW*0.425, y-2, s.LabelColor, labelFontSize, label))
				}
			} else {
				// Negative bar: grows downward from zero line
				barH := (zeroFrac - valFrac) * plotH
				zeroY := marginT + plotH*(1-zeroFrac)
				if barH < 0.5 { barH = 0.5 }
				svg.WriteString(fmt.Sprintf(`<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" rx="1"/>`,
					x, zeroY, barW*0.85, barH, s.Color))
				if s.ShowVal {
					label := formatBarVal(v)
					svg.WriteString(fmt.Sprintf(`<text x="%.1f" y="%.1f" text-anchor="middle" fill="%s" font-size="%.0f" font-weight="bold">%s</text>`,
						x+barW*0.425, zeroY+barH+labelFontSize, s.LabelColor, labelFontSize, label))
				}
			}
		}
		barIdx++
	}

	// Line rendering
	for _, s := range cd.Series {
		if s.ChartType != "line" {
			continue
		}
		var points []string
		for i, v := range s.Values {
			if i >= nCats {
				break
			}
			x := marginL + float64(i)*catW + catW/2
			lineFrac := (v - minLineVal) / lineRange
			y := marginT + plotH*(1-lineFrac)
			points = append(points, fmt.Sprintf("%.1f,%.1f", x, y))
		}
		if len(points) > 1 {
			svg.WriteString(fmt.Sprintf(`<polyline points="%s" fill="none" stroke="%s" stroke-width="2.5"/>`,
				strings.Join(points, " "), s.Color))
		}
		// Markers and labels
		for i, v := range s.Values {
			if i >= nCats {
				break
			}
			x := marginL + float64(i)*catW + catW/2
			lineFrac := (v - minLineVal) / lineRange
			y := marginT + plotH*(1-lineFrac)
			svg.WriteString(fmt.Sprintf(`<circle cx="%.1f" cy="%.1f" r="4" fill="%s" stroke="white" stroke-width="1"/>`, x, y, s.Color))
			if s.ShowVal {
				label := fmt.Sprintf("%.1f%%", v)
				svg.WriteString(fmt.Sprintf(`<text x="%.1f" y="%.1f" text-anchor="middle" fill="%s" font-size="%.0f" font-weight="bold">%s</text>`,
					x, y-8, s.LabelColor, labelFontSize, label))
			}
		}
	}

	// Legend
	if cd.HasLegend {
		legendY := height - math.Max(8, height*0.04)
		legendX := width / 2.0
		charW := baseFontSize * 0.7
		totalW := 0.0
		for _, s := range cd.Series {
			totalW += float64(len(s.Name))*charW + 25
		}
		startX := legendX - totalW/2
		for _, s := range cd.Series {
			if s.ChartType == "line" {
				svg.WriteString(fmt.Sprintf(`<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="%s" stroke-width="2"/>`, startX, legendY-3, startX+12, legendY-3, s.Color))
			} else {
				svg.WriteString(fmt.Sprintf(`<rect x="%.1f" y="%.1f" width="10" height="8" fill="%s"/>`, startX, legendY-8, s.Color))
			}
			svg.WriteString(fmt.Sprintf(`<text x="%.1f" y="%.1f" fill="#555" font-size="%.0f" font-weight="bold">%s</text>`, startX+14, legendY, baseFontSize, xmlEscape(s.Name)))
			startX += float64(len(s.Name))*charW + 25
		}
	}

	svg.WriteString(`</svg>`)
	return svg.String()
}

func niceMax(v float64) float64 {
	if v <= 0 {
		return 100
	}
	// Add ~15% headroom then round to a nice number
	target := v * 1.15
	exp := math.Floor(math.Log10(target))
	base := math.Pow(10, exp)
	frac := target / base
	if frac <= 1.2 {
		return 1.2 * base
	} else if frac <= 1.5 {
		return 1.5 * base
	} else if frac <= 2.0 {
		return 2.0 * base
	} else if frac <= 2.5 {
		return 2.5 * base
	} else if frac <= 3.0 {
		return 3.0 * base
	} else if frac <= 4.0 {
		return 4.0 * base
	} else if frac <= 5.0 {
		return 5.0 * base
	} else if frac <= 7.5 {
		return 7.5 * base
	}
	return 10.0 * base
}

func formatAxisVal(v float64) string {
	if v == 0 {
		return "-"
	}
	if v >= 1 && v == math.Floor(v) {
		return fmt.Sprintf("%.1f", v)
	}
	return fmt.Sprintf("%.1f", v)
}

func formatBarVal(v float64) string {
	if v >= 100 {
		return fmt.Sprintf("%.0f", v)
	}
	if v >= 10 {
		return fmt.Sprintf("%.1f", v)
	}
	return fmt.Sprintf("%.1f", v)
}

func xmlEscape(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	return s
}

