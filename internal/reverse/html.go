package reverse

import (
	"encoding/json"
	"fmt"
	"html"
	"regexp"
	"sort"
	"strings"
	"strconv"
)

type ComTextRun struct {
	Text      string  `json:"text"`
	Font      string  `json:"font"`
	Size      float64 `json:"size"`
	Bold      int     `json:"bold"`
	Italic    int     `json:"italic"`
	Underline int     `json:"underline"`
	Color     string  `json:"color"`
	Align     int     `json:"align"`
	Bullet    int     `json:"bullet"`
	LineSpace float64 `json:"lineSpace,omitempty"`
	SpaceAft  float64 `json:"spaceAft,omitempty"`
	Outline   string  `json:"outline,omitempty"`
}

type TableCell struct {
	Text      string  `json:"text"`
	Color     string  `json:"color"`
	Bold      int     `json:"bold"`
	Size      float64 `json:"size"`
	Bg        string  `json:"bg"`
	MarL      float64 `json:"marL,omitempty"`
	MarR      float64 `json:"marR,omitempty"`
	MarT      float64 `json:"marT,omitempty"`
	MarB      float64 `json:"marB,omitempty"`
	Align     string  `json:"align,omitempty"`
	VAlign    string  `json:"valign,omitempty"`
	BdrT      string  `json:"bdrT,omitempty"`
	BdrB      string  `json:"bdrB,omitempty"`
	BdrL      string  `json:"bdrL,omitempty"`
	BdrR      string  `json:"bdrR,omitempty"`
	MergeSkip int     `json:"mergeSkip"`
	VertMerge int     `json:"vertMerge"`
}

type ComShape struct {
	ID               int               `json:"id"`
	Name             string            `json:"name"`
	Type             string            `json:"type"`
	PhType           string            `json:"phType,omitempty"`
	X                float64           `json:"x"`
	Y                float64           `json:"y"`
	W                float64           `json:"w"`
	H                float64           `json:"h"`
	Rotation         float64           `json:"rotation,omitempty"`
	ShapeType        string            `json:"shapeType,omitempty"`
	Valign           string            `json:"valign,omitempty"`
	PadL             float64           `json:"padL,omitempty"`
	PadR             float64           `json:"padR,omitempty"`
	PadT             float64           `json:"padT,omitempty"`
	PadB             float64           `json:"padB,omitempty"`
	FillType         string            `json:"fillType"`
	FillColor        string            `json:"fillColor"`
	FillTransparency float64           `json:"fillTransparency"`
	GradientStops    []GradientStop    `json:"gradientStops,omitempty"`
	GradientAngle    float64           `json:"gradientAngle,omitempty"`
	BorderColor      string            `json:"borderColor,omitempty"`
	BorderWidth      float64           `json:"borderWidth,omitempty"`
	BorderRadius     float64           `json:"borderRadius,omitempty"`
	HasShadow        bool              `json:"hasShadow,omitempty"`
	ShadowColor      string            `json:"shadowColor,omitempty"`
	ImagePath        string            `json:"imagePath"`
	CropL            float64           `json:"cropL,omitempty"`
	CropT            float64           `json:"cropT,omitempty"`
	CropR            float64           `json:"cropR,omitempty"`
	CropB            float64           `json:"cropB,omitempty"`
	TableRows        int               `json:"tableRows"`
	TableCols        int               `json:"tableCols"`
	TableColWidths   []float64         `json:"tableColWidths,omitempty"`
	TableRowHeights  []float64         `json:"tableRowHeights,omitempty"`
	TableData        []json.RawMessage `json:"tableData"`
	HasText          bool              `json:"hasText"`
	TextRuns         []ComTextRun      `json:"textRuns"`
	ChartSVG         string            `json:"chartSVG,omitempty"`
}

type ComSlide struct {
	SlideIndex      int        `json:"slideIndex"`
	SlideImage      string     `json:"slideImage"`
	BackgroundColor string     `json:"backgroundColor"`
	BackgroundImage string     `json:"backgroundImage"`
	Shapes          []ComShape `json:"shapes"`
}

type ComDoc struct {
	SlideWidth  float64    `json:"slideWidth"`
	SlideHeight float64    `json:"slideHeight"`
	SlideCount  int        `json:"slideCount"`
	Slides      []ComSlide `json:"slides"`
}

func GenerateHTML(doc *ComDoc) string {
	var builder strings.Builder
	pt2px := 1.333

	builder.WriteString("<!DOCTYPE html>\n<html>\n<head>\n")
	builder.WriteString("<meta charset=\"UTF-8\">\n")
	builder.WriteString("<style>\n")
	builder.WriteString("  * { margin:0; padding:0; box-sizing:border-box; }\n")
	builder.WriteString("  body { background:#0a0a0a; font-family:'Segoe UI',Arial,sans-serif; display:flex; flex-direction:column; gap:20px; padding:20px 0; justify-content:flex-start; align-items:center; min-height:100vh; }\n")
	builder.WriteString(fmt.Sprintf("  .slide { position:relative; width:%.0fpx; height:%.0fpx; background:white; overflow:hidden; box-shadow:0 10px 40px rgba(0,0,0,0.7); }\n",
		float64(doc.SlideWidth)*pt2px, float64(doc.SlideHeight)*pt2px))
	// Remove overflow:hidden from shape
	builder.WriteString("  .shape { position:absolute; box-sizing:border-box; overflow:visible; }\n")
	builder.WriteString("  .shape img { width:100%; height:100%; object-fit:cover; display:block; }\n")
	builder.WriteString("  .para { margin:0 0 3px 0; }\n")
	// Remove height:100% from table
	builder.WriteString("  table.ppt-table { border-collapse:collapse; width:100%; font-size:8px; table-layout:fixed; }\n")
	// Remove overflow:hidden from td
	builder.WriteString("  table.ppt-table td { border:none; padding:1px 3px; vertical-align:middle; word-break:break-word; word-wrap:break-word; overflow:visible; }\n")
	builder.WriteString("</style>\n</head>\n<body>\n")

	for _, slide := range doc.Slides {
		bgStyle := "background:white;"
		if slide.BackgroundColor != "" && slide.BackgroundColor != "#000000" {
			bgStyle = fmt.Sprintf("background:%s;", fixBGR(slide.BackgroundColor))
		}
		// 배경 이미지가 있으면 CSS background-image로 렌더링
		if slide.BackgroundImage != "" {
			imgSrc := strings.ReplaceAll(slide.BackgroundImage, "\\\\", "/")
			imgSrc = strings.ReplaceAll(imgSrc, "\\", "/")
			parts := strings.Split(imgSrc, "/")
			imgFile := parts[len(parts)-1]
			bgStyle = fmt.Sprintf("background:url('../slide_images/%s') center/cover no-repeat;", imgFile)
		}
		builder.WriteString(fmt.Sprintf("  <div class=\"slide\" style=\"%s\">\n", bgStyle))

		// Shape 중복 제거: 동일 좌표(X,Y,W,H)는 마지막(Slide) 것만 사용
		seen := make(map[string]int) // key -> last index
		for i, shape := range slide.Shapes {
			key := fmt.Sprintf("%.0f,%.0f,%.0f,%.0f", shape.X, shape.Y, shape.W, shape.H)
			seen[key] = i
		}
		skipSet := make(map[int]bool)
		for i, shape := range slide.Shapes {
			key := fmt.Sprintf("%.0f,%.0f,%.0f,%.0f", shape.X, shape.Y, shape.W, shape.H)
			if seen[key] != i {
				skipSet[i] = true // 이전(Master/Layout) shape 스킵
			}
		}

		for si, shape := range slide.Shapes {
			if skipSet[si] {
				continue
			}
			x := shape.X * pt2px
			y := shape.Y * pt2px
			w := shape.W * pt2px
			h := shape.H * pt2px

			style := fmt.Sprintf("left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;z-index:%d;", x, y, w, h, si+1)

			// Preset geometry clip-path
			switch shape.ShapeType {
			case "triangle":
				style += "clip-path:polygon(50% 0%, 0% 100%, 100% 100%);"
			case "rtTriangle":
				style += "clip-path:polygon(0% 0%, 0% 100%, 100% 100%);"
			case "parallelogram":
				style += "clip-path:polygon(25% 0%, 100% 0%, 75% 100%, 0% 100%);"
			case "trapezoid":
				style += "clip-path:polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%);"
			case "diamond":
				style += "clip-path:polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%);"
			case "pentagon":
				style += "clip-path:polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%);"
			case "hexagon":
				style += "clip-path:polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%);"
			case "chevron":
				style += "clip-path:polygon(0% 0%, 85% 0%, 100% 50%, 85% 100%, 0% 100%, 15% 50%);"
			case "homePlate":
				style += "clip-path:polygon(0% 0%, 85% 0%, 100% 50%, 85% 100%, 0% 100%);"
			case "ellipse":
				style += "border-radius:50%;"
			case "roundRect":
				style += "border-radius:12px;"
			}

			// 배경색
			bgColor := "transparent"
			if shape.FillType == "solid" && shape.FillColor != "" {
				if shape.FillColor == "#000000" {
					// 검정 배경 = 이미지 유실 가능성 → 투명 처리
					bgColor = "transparent"
				} else if shape.FillTransparency < 0.95 {
					alpha := 1.0 - shape.FillTransparency
					bgColor = fixBGRA(shape.FillColor, alpha)
				}
			}
			style += fmt.Sprintf("background:%s;", bgColor)

			// ===== 이미지 (PICTURE + GRAPHIC 로고/사진) =====
			if shape.ImagePath != "" && (!shape.HasText || len(shape.TextRuns) == 0) {
				// 이미지 shape는 배경 투명 강제 (SVG 로고 등이 배경색에 가려지지 않도록)
				imgStyle := fmt.Sprintf("left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;z-index:%d;background:transparent;overflow:hidden;",
					shape.X*pt2px, shape.Y*pt2px, shape.W*pt2px, shape.H*pt2px, si+1)
				builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", imgStyle))
				imgSrc := strings.ReplaceAll(shape.ImagePath, "\\\\", "/")
				imgSrc = strings.ReplaceAll(imgSrc, "\\", "/")
				parts := strings.Split(imgSrc, "/")
				imgFile := parts[len(parts)-1]

				// [Quark 1] Crop (srcRect) 계산
				vw := 1.0 - shape.CropL - shape.CropR
				vh := 1.0 - shape.CropT - shape.CropB
				if vw <= 0 || vh <= 0 {
					vw, vh = 1.0, 1.0
				}
				imgW := 100.0 / vw
				imgH := 100.0 / vh
				imgL := -shape.CropL * imgW
				imgT := -shape.CropT * imgH
				innerImgStyle := fmt.Sprintf("position:absolute;width:%.2f%%;height:%.2f%%;left:%.2f%%;top:%.2f%%;", imgW, imgH, imgL, imgT)

				builder.WriteString(fmt.Sprintf("      <img src=\"../slide_images/%s\" alt=\"%s\" style=\"%s\">\n", imgFile, html.EscapeString(shape.Name), innerImgStyle))
				builder.WriteString("    </div>\n")
				continue
			}

			// ===== 차트 SVG =====
			if shape.ChartSVG != "" {
				builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", style))
				builder.WriteString(shape.ChartSVG)
				builder.WriteString("\n    </div>\n")
				continue
			}

			// ===== 테이블 (colspan + rowspan + 셀 배경색) =====
			if shape.TableRows > 0 && len(shape.TableData) > 0 {
				builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", style))
				builder.WriteString("      <table class=\"ppt-table\">\n")

				// colgroup — 열 폭 반영 (절대 px)
				if len(shape.TableColWidths) > 0 {
					builder.WriteString("      <colgroup>\n")
					for _, cw := range shape.TableColWidths {
						builder.WriteString(fmt.Sprintf("        <col style=\"width:%.0fpx\">\n", cw))
					}
					builder.WriteString("      </colgroup>\n")
				}

				// 1단계: 모든 행 파싱
				allRows := make([][]TableCell, 0)
				for _, rowRaw := range shape.TableData {
					var cells []TableCell
					if err := json.Unmarshal(rowRaw, &cells); err != nil {
						continue
					}
					allRows = append(allRows, cells)
				}

				// 2단계: rowspan 계산용 skip 맵 (row,col) → true면 스킵
				skipCell := make(map[[2]int]bool)

				// border fallback: 테이블 전체에 border가 없으면 기본 가로 border 적용
				hasBorder := false
				for _, r := range allRows {
					for _, c := range r {
						if c.BdrT != "" || c.BdrB != "" || c.BdrL != "" || c.BdrR != "" {
							hasBorder = true
							break
						}
					}
					if hasBorder { break }
				}
				if !hasBorder && len(allRows) > 1 {
					// 기본 가로 border: 각 행 하단에 연한 회색 1px
					for ri := range allRows {
						for ci := range allRows[ri] {
							if ri < len(allRows)-1 {
								allRows[ri][ci].BdrB = "0.5pt solid #D9D9D9"
							}
							if ri > 0 {
								allRows[ri][ci].BdrT = "0.5pt solid #D9D9D9"
							}
						}
					}
				}

				for ri, row := range allRows {
					trStyle := ""
					if ri < len(shape.TableRowHeights) {
						trStyle = fmt.Sprintf(" style=\"height:%.0fpx\"", shape.TableRowHeights[ri])
					}
					builder.WriteString(fmt.Sprintf("        <tr%s>\n", trStyle))
					for ci := 0; ci < len(row); ci++ {
						cell := row[ci]
						// horizontal merge 스킵
						if cell.MergeSkip == 1 {
							continue
						}
						// vertical merge 스킵
						if skipCell[[2]int{ri, ci}] {
							continue
						}

						// colspan 계산
						colspan := 1
						for j := ci + 1; j < len(row); j++ {
							if row[j].MergeSkip == 1 {
								colspan++
							} else {
								break
							}
						}

						// rowspan 계산
						rowspan := 1
						if cell.VertMerge == 0 {
							for rj := ri + 1; rj < len(allRows); rj++ {
								if ci < len(allRows[rj]) && allRows[rj][ci].VertMerge == 1 {
									rowspan++
									// 이 셀과 같은 colspan 범위도 스킵
									for cc := ci; cc < ci+colspan && cc < len(allRows[rj]); cc++ {
										skipCell[[2]int{rj, cc}] = true
									}
								} else {
									break
								}
							}
						}
						if cell.VertMerge == 1 {
							continue // 이미 상위 셀의 rowspan에 포함됨
						}

						color := fixBGR(cell.Color)
						if color == "#000000" || color == "" {
							color = "#333"
						}
						cs := fmt.Sprintf("color:%s;", color)
						if cell.Bold == -1 {
							cs += "font-weight:bold;"
						}
						if cell.Size > 0 {
							cs += fmt.Sprintf("font-size:%.1fpx;", cell.Size*pt2px*0.96)
						}
						if cell.Align == "ctr" { cs += "text-align:center;" } else if cell.Align == "r" { cs += "text-align:right;" }
						if cell.VAlign == "ctr" { cs += "vertical-align:middle;" } else if cell.VAlign == "b" { cs += "vertical-align:bottom;" } else { cs += "vertical-align:top;" }
						if cell.Bg != "" {
							cs += fmt.Sprintf("background:%s;", fixBGR(cell.Bg))
						}
						if cell.BdrT != "" { cs += fmt.Sprintf("border-top:%s;", resolveSchemeInCSS(cell.BdrT, nil)) }
						if cell.BdrB != "" { cs += fmt.Sprintf("border-bottom:%s;", resolveSchemeInCSS(cell.BdrB, nil)) }
						if cell.BdrL != "" { cs += fmt.Sprintf("border-left:%s;", resolveSchemeInCSS(cell.BdrL, nil)) }
						if cell.BdrR != "" { cs += fmt.Sprintf("border-right:%s;", resolveSchemeInCSS(cell.BdrR, nil)) }
						// [Quark 2] Apply cell padding
						padL, padR, padT, padB := 5.0, 5.0, 2.0, 2.0
						if cell.MarL != -1 && cell.MarL != 0 { padL = cell.MarL }
						if cell.MarR != -1 && cell.MarR != 0 { padR = cell.MarR }
						if cell.MarT != -1 && cell.MarT != 0 { padT = cell.MarT }
						if cell.MarB != -1 && cell.MarB != 0 { padB = cell.MarB }
						cs += fmt.Sprintf("padding:%.1fpx %.1fpx %.1fpx %.1fpx;", padT, padR, padB, padL)
						spanAttr := ""
						if colspan > 1 {
							spanAttr += fmt.Sprintf(" colspan=\"%d\"", colspan)
						}
						if rowspan > 1 {
							spanAttr += fmt.Sprintf(" rowspan=\"%d\"", rowspan)
						}
						// 셀 텍스트 내 줄바꿈을 <br>로 변환 (P0-1: S03/S14/S15 수정)
						cellHTML := strings.ReplaceAll(html.EscapeString(cell.Text), "\n", "<br>")
						cellHTML = strings.ReplaceAll(cellHTML, "\r", "")
						builder.WriteString(fmt.Sprintf("          <td%s style=\"%s\">%s</td>\n", spanAttr, cs, cellHTML))
					}
					builder.WriteString("        </tr>\n")
				}
				builder.WriteString("      </table>\n")
				builder.WriteString("    </div>\n")
				continue
			}

			// ===== 텍스트 =====
			if !shape.HasText || len(shape.TextRuns) == 0 {
				// 배경색 있는 도형만 렌더링 (투명이면 스킵 → 노이즈 제거)
				if bgColor != "transparent" {
					builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\"></div>\n", style))
				}
				continue
			}

			jc := "flex-start" // default top
			if shape.Valign == "ctr" { jc = "center" } else if shape.Valign == "b" { jc = "flex-end" }
			padL, padR, padT, padB := 0.0, 0.0, 0.0, 0.0
			if shape.PadL != -1 && shape.PadL != 0 { padL = shape.PadL }
			if shape.PadR != -1 && shape.PadR != 0 { padR = shape.PadR }
			if shape.PadT != -1 && shape.PadT != 0 { padT = shape.PadT }
			if shape.PadB != -1 && shape.PadB != 0 { padB = shape.PadB }

			style += fmt.Sprintf("display:flex;flex-direction:column;justify-content:%s;padding:%.1fpx %.1fpx %.1fpx %.1fpx;line-height:1.35;",
				jc, padT, padR, padB, padL)
			builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", style))

			var para []ComTextRun
			for i, run := range shape.TextRuns {
				para = append(para, run)
				if i == len(shape.TextRuns)-1 || strings.Contains(run.Text, "\n") {
					align := "left"
					if len(para) > 0 {
						switch para[0].Align {
						case 2: align = "center"
						case 3: align = "right"
						}
					}
					builder.WriteString(fmt.Sprintf("      <p class=\"para\" style=\"text-align:%s;\">\n", align))
					for _, pr := range para {
						// 줄바꿈을 <br>로 변환 (P0-1b: 텍스트 run 줄바꿈 보존)
						txt := strings.TrimRight(pr.Text, "\n\r")
						if txt == "" { continue }
						txt = strings.ReplaceAll(txt, "\r\n", "\n")
						lines := strings.Split(txt, "\n")
						txt = strings.Join(lines, "<br>")

						color := fixBGR(pr.Color)
						if color == "" { color = "#333" }
						// 흰 배경 위의 검정 텍스트는 살리되, 투명 배경의 검정 텍스트는 진회색으로
						if color == "#000000" && bgColor == "transparent" {
							color = "#333"
						}

						ss := fmt.Sprintf("font-family:'%s';font-size:%.1fpx;color:%s;",
							pr.Font, pr.Size*pt2px, color)
						if pr.Bold == -1 { ss += "font-weight:bold;" }
						if pr.Italic == -1 { ss += "font-style:italic;" }
						if pr.Underline == -1 { ss += "text-decoration:underline;" }
						// 흰색 텍스트에 CSS outline 추가 (HTML에서 가독성 보정용)
						if pr.Outline != "" && pr.Outline != "#FFFFFF" && pr.Outline != "#ffffff" {
							ss += fmt.Sprintf("-webkit-text-stroke:1px %s;paint-order:stroke fill;", fixBGR(pr.Outline))
						} else if color == "#FFFFFF" || color == "#ffffff" {
							ss += "-webkit-text-stroke:1px #333;paint-order:stroke fill;"
						}

						// 불릿 기호 추가
						prefix := ""
						if pr.Bullet != 0 {
							prefix = "• "
						}

						builder.WriteString(fmt.Sprintf("        <span style=\"%s\">%s%s</span>\n", ss, prefix, html.EscapeString(txt)))
					}
					builder.WriteString("      </p>\n")
					para = nil
				}
			}
			builder.WriteString("    </div>\n")
		}
		builder.WriteString("  </div>\n")
	}

	builder.WriteString("</body>\n</html>")
	return builder.String()
}

func fixBGR(hex string) string {
	if len(hex) == 7 && hex[0] == '#' {
		return "#" + hex[5:7] + hex[3:5] + hex[1:3]
	}
	return hex
}

func fixBGRA(hex string, alpha float64) string {
	if len(hex) == 7 && hex[0] == '#' {
		var r, g, b int
		fmt.Sscanf(hex, "#%02x%02x%02x", &b, &g, &r)
		return fmt.Sprintf("rgba(%d,%d,%d,%.2f)", r, g, b, alpha)
	}
	return hex
}

// resolveSchemeInCSS replaces "scheme:xxx" in a CSS border value with actual hex color
func resolveSchemeInCSS(cssVal string, themeColors map[string]string) string {
	if !strings.Contains(cssVal, "scheme:") { return cssVal }
	// Match both scheme:key and scheme:key:lum:mod:off
	re := regexp.MustCompile(`scheme:([a-zA-Z0-9]+)(?::lum:([0-9.]+):([0-9.]+))?`)
	return re.ReplaceAllStringFunc(cssVal, func(m string) string {
		matches := re.FindStringSubmatch(m)
		key := matches[1]
		
		color := "#000000"
		if c, ok := themeColors[key]; ok {
			color = c
		} else {
			// fallback defaults
			switch key {
			case "bg1", "lt1": color = "#FFFFFF"
			case "tx1", "dk1": color = "#000000"
			case "bg2", "lt2": color = "#EEEEEE"
			case "tx2", "dk2": color = "#666666"
			}
		}

		if matches[2] != "" && matches[3] != "" {
			lumMod, _ := strconv.ParseFloat(matches[2], 64)
			lumOff, _ := strconv.ParseFloat(matches[3], 64)
			lumMod = lumMod / 100000.0
			lumOff = lumOff / 100000.0
			
			if len(color) == 7 && color[0] == '#' {
				r, _ := strconv.ParseInt(color[1:3], 16, 64)
				g, _ := strconv.ParseInt(color[3:5], 16, 64)
				b, _ := strconv.ParseInt(color[5:7], 16, 64)
				
				clamp := func(v float64) int {
					if v < 0 { return 0 }
					if v > 255 { return 255 }
					return int(v + 0.5)
				}
				nr := clamp(float64(r)*lumMod + 255.0*lumOff)
				ng := clamp(float64(g)*lumMod + 255.0*lumOff)
				nb := clamp(float64(b)*lumMod + 255.0*lumOff)
				color = fmt.Sprintf("#%02X%02X%02X", nr, ng, nb)
			}
		}
		return color
	})
}

// GenerateHTMLFromParsed renders HTML from Go native ParsedDoc (no BGR fix needed — XML colors are RGB)
// [RETROACTIVE PARITY] VBS COM 렌더러(GenerateHTML)의 모든 노하우를 전수 이식한 버전
func GenerateHTMLFromParsed(doc *ParsedDoc) string {
	var builder strings.Builder
	themeColors := doc.ThemeColors
	if themeColors == nil { themeColors = map[string]string{} }
	pt2px := 1.333

	builder.WriteString("<!DOCTYPE html>\n<html>\n<head>\n")
	builder.WriteString("<meta charset=\"UTF-8\">\n")
	builder.WriteString("<style>\n")
	// Freesentation 웹폰트
	builder.WriteString("  @font-face { font-family:'Freesentation'; src:url('https://cdn.jsdelivr.net/gh/projectnoonnu/2404@1.0/Freesentation-4Regular.woff2') format('woff2'); font-weight:400; font-display:swap; }\n")
	builder.WriteString("  @font-face { font-family:'Freesentation'; src:url('https://cdn.jsdelivr.net/gh/projectnoonnu/2404@1.0/Freesentation-7Bold.woff2') format('woff2'); font-weight:700; font-display:swap; }\n")
	builder.WriteString("  @font-face { font-family:'Freesentation'; src:url('https://cdn.jsdelivr.net/gh/projectnoonnu/2404@1.0/Freesentation-6SemiBold.woff2') format('woff2'); font-weight:600; font-display:swap; }\n")
	builder.WriteString("  @font-face { font-family:'Freesentation'; src:url('https://cdn.jsdelivr.net/gh/projectnoonnu/2404@1.0/Freesentation-9Black.woff2') format('woff2'); font-weight:900; font-display:swap; }\n")
	builder.WriteString("  * { margin:0; padding:0; box-sizing:border-box; }\n")
	builder.WriteString("  body { background:#0a0a0a; font-family:'Freesentation','Segoe UI',Arial,sans-serif; display:flex; flex-direction:column; gap:20px; padding:20px 0; justify-content:flex-start; align-items:center; min-height:100vh; }\n")
	builder.WriteString(fmt.Sprintf("  .slide { position:relative; width:%.0fpx; height:%.0fpx; background:white; overflow:hidden; box-shadow:0 10px 40px rgba(0,0,0,0.7); }\n",
		doc.SlideWidth*pt2px, doc.SlideHeight*pt2px))
	// Remove overflow:hidden from shape so text can overflow natively like PPTX when autofit is off
	builder.WriteString("  .shape { position:absolute; box-sizing:border-box; overflow:visible; }\n")
	builder.WriteString("  .shape img { width:100%; height:100%; object-fit:cover; display:block; }\n")
	builder.WriteString("  .para { margin:0 0 3px 0; }\n")
	// Remove height:100% from table so it expands to fit text content
	builder.WriteString("  table.ppt-table { border-collapse:collapse; width:100%; font-size:8px; table-layout:fixed; }\n")
	// Remove overflow:hidden from td and use break-word
	builder.WriteString("  table.ppt-table td { border:none; padding:1px 3px; vertical-align:middle; word-break:break-word; word-wrap:break-word; overflow:visible; }\n")
	builder.WriteString("</style>\n</head>\n<body>\n")

	for _, slide := range doc.Slides {
		slideBgColor := "white"
		if slide.BackgroundColor != "" {
			slideBgColor = slide.BackgroundColor
		}
		// [Q1-D] (휴리스틱 제거) 가장 큰 shape의 fill 색상으로 배경을 덮어쓰는 로직은, 배경색 자체를 오염시키므로 제거함.
		bgStyle := fmt.Sprintf("background:%s;", slideBgColor)
		builder.WriteString(fmt.Sprintf("  <div class=\"slide\" style=\"%s\">\n", bgStyle))

		// Shape 중복 제거: 동일 좌표(X,Y,W,H)는 마지막(Slide) 것만 사용
		// pptxparser.go의 extractShapes에서 원본 XML 순서대로 정렬되어 있으므로
		// 추가 z-order 재정규화 없이 그대로 사용
		seen := make(map[string]int)
		for i, shape := range slide.Shapes {
			if shape.HasText { continue }
			key := fmt.Sprintf("%.0f,%.0f,%.0f,%.0f", shape.X, shape.Y, shape.W, shape.H)
			seen[key] = i
		}
		skipSet := make(map[int]bool)
		for i, shape := range slide.Shapes {
			if shape.HasText { continue }
			key := fmt.Sprintf("%.0f,%.0f,%.0f,%.0f", shape.X, shape.Y, shape.W, shape.H)
			if seen[key] != i {
				skipSet[i] = true
			}
		}

		for si, shape := range slide.Shapes {
			if skipSet[si] {
				continue
			}
			x := shape.X * pt2px
			y := shape.Y * pt2px
			w := shape.W * pt2px
			if shape.HasText {
				// 원본 PPTX 레이아웃 파리티 유지를 위해 폭 보정을 제거 (overflow:visible 이므로 자동 처리됨)
				// w = w * 1.15
			}
			h := shape.H * pt2px

			if w == 0 && h == 0 {
				continue
			}

			style := fmt.Sprintf("left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;z-index:%d;", x, y, w, h, si+1)

			// Rotation
			if shape.Rotation != 0 {
				style += fmt.Sprintf("transform:rotate(%.1fdeg);", shape.Rotation)
			}
			// Border
			if shape.BorderWidth > 0 && shape.BorderColor != "" {
				style += fmt.Sprintf("border:%.1fpx solid %s;", shape.BorderWidth*pt2px, shape.BorderColor)
			}
			// Preset geometry clip-path
			switch shape.ShapeType {
			case "triangle":
				style += "clip-path:polygon(50% 0%, 0% 100%, 100% 100%);"
			case "rtTriangle":
				style += "clip-path:polygon(0% 0%, 0% 100%, 100% 100%);"
			case "parallelogram":
				style += "clip-path:polygon(25% 0%, 100% 0%, 75% 100%, 0% 100%);"
			case "trapezoid":
				style += "clip-path:polygon(20% 0%, 80% 0%, 100% 100%, 0% 100%);"
			case "diamond":
				style += "clip-path:polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%);"
			case "pentagon":
				style += "clip-path:polygon(50% 0%, 100% 38%, 82% 100%, 18% 100%, 0% 38%);"
			case "hexagon":
				style += "clip-path:polygon(25% 0%, 75% 0%, 100% 50%, 75% 100%, 25% 100%, 0% 50%);"
			case "chevron":
				style += "clip-path:polygon(0% 0%, 85% 0%, 100% 50%, 85% 100%, 0% 100%, 15% 50%);"
			case "homePlate":
				style += "clip-path:polygon(0% 0%, 85% 0%, 100% 50%, 85% 100%, 0% 100%);"
			case "ellipse":
				style += "border-radius:50%;"
			case "roundRect":
				style += "border-radius:12px;"
			default:
				if shape.BorderRadius > 0 {
					minDim := w
					if h < w { minDim = h }
					radiusPx := shape.BorderRadius * minDim
					style += fmt.Sprintf("border-radius:%.0fpx;", radiusPx)
				}
			}
			// Shadow
			if shape.HasShadow {
				style += fmt.Sprintf("box-shadow: 2px 2px 5px %s;", shape.ShadowColor)
			}

			// Fill
			if shape.FillColor != "" {
				shape.FillColor = resolveSchemeInCSS(shape.FillColor, themeColors)
			}
			
			bgColor := "transparent"
			if shape.FillType == "none" {
				// noFill — transparent
			} else if shape.FillType == "solid" && shape.FillColor != "" {
				if (shape.FillColor == "#FFFFFF" || shape.FillColor == "#ffffff") && !shape.HasText && shape.FillTransparency == 0 {
					bgColor = "transparent"
				} else if shape.FillColor == "#000000" {
					bgColor = "transparent"
				} else if shape.FillTransparency > 0 && shape.FillTransparency < 0.95 {
					alpha := 1.0 - shape.FillTransparency
					var r, g, b int
					fmt.Sscanf(shape.FillColor, "#%02x%02x%02x", &r, &g, &b)
					bgColor = fmt.Sprintf("rgba(%d,%d,%d,%.2f)", r, g, b, alpha)
				} else if shape.FillTransparency >= 0.95 {
					bgColor = "transparent"
				} else {
					bgColor = shape.FillColor
				}
			} else if shape.FillType == "gradient" && len(shape.GradientStops) >= 2 {
				// Sort stops by position for correct CSS rendering
				sortedStops := make([]GradientStop, len(shape.GradientStops))
				copy(sortedStops, shape.GradientStops)
				sort.Slice(sortedStops, func(i, j int) bool {
					return sortedStops[i].Position < sortedStops[j].Position
				})
				stops := ""
				for i, gs := range sortedStops {
					pct := float64(gs.Position) / 1000.0
					if i > 0 { stops += "," }
					colorStr := gs.Color
					// Resolve scheme colors to hex
					colorStr = resolveSchemeInCSS(colorStr, themeColors)
					if gs.Alpha >= 0.0 && gs.Alpha < 1.0 && len(colorStr) == 7 && colorStr[0] == '#' {
						// Convert #RRGGBB to rgba(...)
						r, _ := strconv.ParseInt(colorStr[1:3], 16, 64)
						g, _ := strconv.ParseInt(colorStr[3:5], 16, 64)
						b, _ := strconv.ParseInt(colorStr[5:7], 16, 64)
						colorStr = fmt.Sprintf("rgba(%d,%d,%d,%.2f)", r, g, b, gs.Alpha)
					}
					stops += fmt.Sprintf("%s %.0f%%", colorStr, pct)
				}
				bgColor = fmt.Sprintf("linear-gradient(%.0fdeg,%s)", shape.GradientAngle, stops)
			} else if shape.FillType == "gradient" && shape.FillColor != "" {
				bgColor = shape.FillColor
			}
			style += fmt.Sprintf("background:%s;", bgColor)

			// [RETROACTIVE] Image — 상대경로 + 배경 투명 강제
			if shape.ImagePath != "" && (!shape.HasText || len(shape.TextRuns) == 0) {
				imgSrc := shape.ImagePath
				parts := strings.Split(strings.ReplaceAll(imgSrc, "\\", "/"), "/")
				imgFile := parts[len(parts)-1]
				imgStyle := fmt.Sprintf("left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;z-index:%d;background:transparent;overflow:hidden;",
					shape.X*pt2px, shape.Y*pt2px, shape.W*pt2px, shape.H*pt2px, si+1)
				builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", imgStyle))

				// [Quark 1] Crop (srcRect) 계산
				vw := 1.0 - shape.CropL - shape.CropR
				vh := 1.0 - shape.CropT - shape.CropB
				if vw <= 0 || vh <= 0 {
					vw, vh = 1.0, 1.0
				}
				imgW := 100.0 / vw
				imgH := 100.0 / vh
				imgL := -shape.CropL * imgW
				imgT := -shape.CropT * imgH
				innerImgStyle := fmt.Sprintf("position:absolute;width:%.2f%%;height:%.2f%%;left:%.2f%%;top:%.2f%%;", imgW, imgH, imgL, imgT)

				builder.WriteString(fmt.Sprintf("      <img src=\"../%s/%s\" alt=\"%s\" style=\"%s\">\n",
					doc.ImageDir, imgFile, html.EscapeString(shape.Name), innerImgStyle))
				builder.WriteString("    </div>\n")
				continue
			}

			// ===== 차트 SVG =====
			if shape.ChartSVG != "" {
				builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", style))
				builder.WriteString(shape.ChartSVG)
				builder.WriteString("\n    </div>\n")
				continue
			}

			// [RETROACTIVE #6] 테이블 (colspan + rowspan + 셀 배경색)
			if shape.TableRows > 0 && len(shape.TableData) > 0 {
				builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", style))
				builder.WriteString("      <table class=\"ppt-table\">\n")

				// colgroup — 열 폭 반영 (절대 px)
				if len(shape.TableColWidths) > 0 {
					builder.WriteString("      <colgroup>\n")
					for _, cw := range shape.TableColWidths {
						builder.WriteString(fmt.Sprintf("        <col style=\"width:%.0fpx\">\n", cw))
					}
					builder.WriteString("      </colgroup>\n")
				}

				allRows := make([][]TableCell, 0)
				for _, rowRaw := range shape.TableData {
					var cells []TableCell
					if err := json.Unmarshal(rowRaw, &cells); err != nil {
						continue
					}
					allRows = append(allRows, cells)
				}
				skipCell := make(map[[2]int]bool)
				for ri, row := range allRows {
					trStyle := ""
					if ri < len(shape.TableRowHeights) {
						trStyle = fmt.Sprintf(" style=\"height:%.0fpx\"", shape.TableRowHeights[ri])
					}
					builder.WriteString(fmt.Sprintf("        <tr%s>\n", trStyle))
					for ci := 0; ci < len(row); ci++ {
						cell := row[ci]
						if cell.MergeSkip == 1 { continue }
						if skipCell[[2]int{ri, ci}] { continue }
						colspan := 1
						for j := ci + 1; j < len(row); j++ {
							if row[j].MergeSkip == 1 { colspan++ } else { break }
						}
						rowspan := 1
						if cell.VertMerge == 0 {
							for rj := ri + 1; rj < len(allRows); rj++ {
								if ci < len(allRows[rj]) && allRows[rj][ci].VertMerge == 1 {
									rowspan++
									for cc := ci; cc < ci+colspan && cc < len(allRows[rj]); cc++ {
										skipCell[[2]int{rj, cc}] = true
									}
								} else { break }
							}
						}
						if cell.VertMerge == 1 { continue }
						color := cell.Color
						if color == "#000000" || color == "" { color = "#333" }
						cs := fmt.Sprintf("color:%s;", color)
						if cell.Bold == -1 { cs += "font-weight:bold;" }
						if cell.Size > 0 { cs += fmt.Sprintf("font-size:%.1fpx;", cell.Size*pt2px*0.96) }
						if cell.Align == "ctr" { cs += "text-align:center;" } else if cell.Align == "r" { cs += "text-align:right;" }
						if cell.VAlign == "ctr" { cs += "vertical-align:middle;" } else if cell.VAlign == "b" { cs += "vertical-align:bottom;" } else { cs += "vertical-align:top;" }
						if cell.Bg != "" { cs += fmt.Sprintf("background:%s;", cell.Bg) }
						if cell.BdrT != "" { cs += fmt.Sprintf("border-top:%s;", resolveSchemeInCSS(cell.BdrT, themeColors)) }
						if cell.BdrB != "" { cs += fmt.Sprintf("border-bottom:%s;", resolveSchemeInCSS(cell.BdrB, themeColors)) }
						if cell.BdrL != "" { cs += fmt.Sprintf("border-left:%s;", resolveSchemeInCSS(cell.BdrL, themeColors)) }
						if cell.BdrR != "" { cs += fmt.Sprintf("border-right:%s;", resolveSchemeInCSS(cell.BdrR, themeColors)) }
						// [Quark 2] Apply cell padding
						padL, padR, padT, padB := 5.0, 5.0, 2.0, 2.0
						if cell.MarL != -1 && cell.MarL != 0 { padL = cell.MarL }
						if cell.MarR != -1 && cell.MarR != 0 { padR = cell.MarR }
						if cell.MarT != -1 && cell.MarT != 0 { padT = cell.MarT }
						if cell.MarB != -1 && cell.MarB != 0 { padB = cell.MarB }
						cs += fmt.Sprintf("padding:%.1fpx %.1fpx %.1fpx %.1fpx;", padT, padR, padB, padL)
						spanAttr := ""
						if colspan > 1 { spanAttr += fmt.Sprintf(" colspan=\"%d\"", colspan) }
						if rowspan > 1 { spanAttr += fmt.Sprintf(" rowspan=\"%d\"", rowspan) }
						cellHTML := strings.ReplaceAll(html.EscapeString(cell.Text), "\n", "<br>")
						cellHTML = strings.ReplaceAll(cellHTML, "\r", "")
						builder.WriteString(fmt.Sprintf("          <td%s style=\"%s\">%s</td>\n", spanAttr, cs, cellHTML))
					}
					builder.WriteString("        </tr>\n")
				}
				builder.WriteString("      </table>\n")
				builder.WriteString("    </div>\n")
				continue
			}

			// 텍스트 없는 배경 도형
			if !shape.HasText || len(shape.TextRuns) == 0 {
				if bgColor != "transparent" {
					builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\"></div>\n", style))
				}
				continue
			}

			// [RETROACTIVE] min-height + padding 제거하여 박스 공간 확보
			jc := "flex-start"
			if shape.Valign == "ctr" { jc = "center" } else if shape.Valign == "b" { jc = "flex-end" }
			padL, padR, padT, padB := 0.0, 0.0, 0.0, 0.0
			if shape.PadL != -1 && shape.PadL != 0 { padL = shape.PadL }
			if shape.PadR != -1 && shape.PadR != 0 { padR = shape.PadR }
			if shape.PadT != -1 && shape.PadT != 0 { padT = shape.PadT }
			if shape.PadB != -1 && shape.PadB != 0 { padB = shape.PadB }

			style += fmt.Sprintf("display:flex;flex-direction:column;justify-content:%s;padding:%.1fpx %.1fpx %.1fpx %.1fpx;line-height:1.35;min-height:20px;",
				jc, padT, padR, padB, padL)
			if shape.TextWrap == "none" {
				style += "white-space:nowrap;"
			} else {
				style += "word-break:keep-all; overflow-wrap:break-word;"
			}
			builder.WriteString(fmt.Sprintf("    <div class=\"shape\" style=\"%s\">\n", style))

			var para []TextRun
			for i, run := range shape.TextRuns {
				para = append(para, run)
				if i == len(shape.TextRuns)-1 || strings.Contains(run.Text, "\n") {
					align := "left"
					if len(para) > 0 {
						switch para[0].Align {
						case 2: align = "center"
						case 3: align = "right"
						}
					}
					// 행간/여백: XML에 명시된 경우에만 인라인 추가, 아니면 기존 .para CSS 유지
					paraStyle := fmt.Sprintf("text-align:%s;", align)
					if len(para) > 0 && para[0].LineSpace > 0 && para[0].LineSpace != 1.0 {
						paraStyle += fmt.Sprintf("line-height:%.2f;", para[0].LineSpace)
					}
					if len(para) > 0 && para[0].SpaceAft > 0 {
						paraStyle += fmt.Sprintf("margin-bottom:%.1fpx;", para[0].SpaceAft*1.33)
					}
					builder.WriteString(fmt.Sprintf("      <p class=\"para\" style=\"%s\">\n", paraStyle))
					for _, pr := range para {
						txt := strings.TrimRight(pr.Text, "\n\r")
						if txt == "" { continue }
						txt = strings.ReplaceAll(txt, "\r\n", "\n")
						lines := strings.Split(txt, "\n")
						txt = strings.Join(lines, "<br>")

						color := pr.Color
						if color == "" { color = "#333" }
						color = resolveSchemeInCSS(color, themeColors)
						if len(color) == 6 && color[0] != '#' {
							color = "#" + color
						}
						// 텍스트 보정용 effectiveBgColor 계산
						effectiveBg := slideBgColor
						if effectiveBg == "white" || effectiveBg == "#FFFFFF" || effectiveBg == "" {
							maxArea := 0.0
							for _, bgShp := range slide.Shapes {
								area := bgShp.W * bgShp.H
								if area > maxArea && bgShp.FillColor != "" && bgShp.FillColor != "transparent" && bgShp.FillColor != "#FFFFFF" {
									if bgShp.FillType == "solid" || bgShp.FillType == "gradient" {
										maxArea = area
										effectiveBg = bgShp.FillColor
									}
								}
							}
						}
						// slide 배경이 밝을 때만 흰색 텍스트 보정
						slideBgIsLight := effectiveBg == "white" || effectiveBg == "#FFFFFF" || effectiveBg == ""
						if bgColor == "transparent" && slideBgIsLight {
							if color == "#FFFFFF" || color == "#ffffff" {
								color = "#333"
							}
						}

						// pt → px: ×1.333 (72dpi → 96dpi) + 브라우저 렌더링 글꼴 너비 보정 (0.96)
						fontScale := 0.96
						fontSize := pr.Size * pt2px * fontScale

						// Freesentation 폰트 매핑
						fontName := pr.Font
						if fontName == "Arial" || fontName == "Calibri" || fontName == "" {
							fontName = "Freesentation"
						}

						ss := fmt.Sprintf("font-family:'%s','Freesentation',sans-serif;font-size:%.1fpx;color:%s;",
							fontName, fontSize, color)
						if pr.Bold == -1 { ss += "font-weight:bold;" }
						if pr.Italic == -1 { ss += "font-style:italic;" }
						if pr.Underline == -1 { ss += "text-decoration:underline;" }

						// [RETROACTIVE #9] 흰색 텍스트 outline (빨간 배경 위의 흰 글씨 등)
						if (color == "#FFFFFF" || color == "#ffffff") && bgColor != "transparent" {
							ss += "-webkit-text-stroke:1px rgba(0,0,0,0.1);paint-order:stroke fill;"
						}

						// [RETROACTIVE #8] 불릿 기호
						prefix := ""
						if pr.Bullet == -1 || pr.Bullet == 1 {
							prefix = "• "
						}

						builder.WriteString(fmt.Sprintf("        <span style=\"%s\">%s%s</span>\n", ss, prefix, html.EscapeString(txt)))
					}
					builder.WriteString("      </p>\n")
					para = nil
				}
			}
			builder.WriteString("    </div>\n")
		}
		builder.WriteString("  </div>\n")
	}

	builder.WriteString("</body>\n</html>")
	return builder.String()
}

