package reverse

import (
	"archive/zip"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"

	"regexp"
	"sort"
	"strconv"
	"strings"
)

// ===== SSOT PPTX Native Parser =====
// ZIP → XML → rels → media 직접 파싱
// VBS COM 의존성 제거

// === Output Types (html.go와 공유) ===

type TextRun struct {
	Text      string  `json:"text"`
	Font      string  `json:"font"`
	Size      float64 `json:"size"`
	Bold      int     `json:"bold"`
	Italic    int     `json:"italic"`
	Underline int     `json:"underline"`
	Color     string  `json:"color"`
	Align     int     `json:"align"`
	Bullet    int     `json:"bullet"`
	LineSpace float64 `json:"lineSpace,omitempty"` // 행간 비율 (기본 1.0)
	SpaceAft  float64 `json:"spaceAft,omitempty"`  // 문단 후 여백
}

type TableCellData struct {
	Text   string  `json:"text"`
	Color  string  `json:"color"`
	Bold   int     `json:"bold"`
	Size   float64 `json:"size"`
	Bg     string  `json:"bg,omitempty"`
	MarL   float64 `json:"marL,omitempty"`
	MarR   float64 `json:"marR,omitempty"`
	MarT   float64 `json:"marT,omitempty"`
	MarB   float64 `json:"marB,omitempty"`
	VAlign string  `json:"valign,omitempty"`
	Align  string  `json:"align,omitempty"`
	BdrT   string  `json:"bdrT,omitempty"` // top border css (e.g. "1px solid #000")
	BdrB   string  `json:"bdrB,omitempty"`
	BdrL   string  `json:"bdrL,omitempty"`
	BdrR   string  `json:"bdrR,omitempty"`
	MergeSkip int  `json:"mergeSkip"`  // hMerge="1" → 수평 병합 대상 (skip)
	VertMerge int  `json:"vertMerge"`  // vMerge="1" → 수직 병합 대상
}

type GradientStop struct {
	Color    string  `json:"color"`
	Position int     `json:"position"` // 0-100000
	Alpha    float64 `json:"alpha"`    // 0.0 - 1.0 (default 1.0)
}

type ParsedShape struct {
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
	XmlOffset        int               `json:"-"`
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
	QuarantinedMeta  map[string]string `json:"quarantinedMeta,omitempty"` // 렌더링 불가 데이터(3D, Anim 등) 보존용
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
	TextWrap         string            `json:"textWrap"`
	TextRuns         []TextRun         `json:"textRuns"`
	Children         []ParsedShape     `json:"children,omitempty"` // 그룹 내 자식 도형들
	ChartSVG         string            `json:"chartSVG,omitempty"` // Rendered chart SVG
}

type ParsedSlide struct {
	SlideIndex      int           `json:"slideIndex"`
	Layout          string        `json:"layout"`
	BackgroundColor string        `json:"backgroundColor"`
	Shapes          []ParsedShape `json:"shapes"`
}

type ParsedDoc struct {
	File        string            `json:"file"`
	ImageDir    string            `json:"imageDir"`
	SlideWidth  float64           `json:"slideWidth"`
	SlideHeight float64           `json:"slideHeight"`
	SlideCount  int               `json:"slideCount"`
	Slides      []ParsedSlide     `json:"slides"`
	ThemeColors map[string]string `json:"themeColors,omitempty"`
}

// === XML structs ===

type xmlRelationships struct {
	XMLName xml.Name        `xml:"Relationships"`
	Rels    []xmlRelationship `xml:"Relationship"`
}

type xmlRelationship struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// EMU to pt conversion (1pt = 12700 EMU)
func emuToPt(emu int64) float64 {
	return float64(emu) / 12700.0
}

// Parse PPTX file natively via ZIP + XML
func ParsePPTX(pptxPath string, mediaOutputDir string) (*ParsedDoc, error) {
	r, err := zip.OpenReader(pptxPath)
	if err != nil {
		return nil, fmt.Errorf("cannot open PPTX: %w", err)
	}
	defer r.Close()

	// Build file map for quick lookup
	fileMap := make(map[string]*zip.File)
	for _, f := range r.File {
		fileMap[f.Name] = f
	}

	// Extract all media files
	os.MkdirAll(mediaOutputDir, 0755)
	for name, f := range fileMap {
		if strings.HasPrefix(name, "ppt/media/") {
			fname := filepath.Base(name)
			outPath := filepath.Join(mediaOutputDir, fname)
			if err := extractFile(f, outPath); err != nil {
				fmt.Fprintf(os.Stderr, "WARN: failed to extract %s: %v\n", name, err)
			}
		}
	}

	// Parse presentation.xml for slide dimensions
	doc := &ParsedDoc{File: pptxPath, ImageDir: mediaOutputDir}
	if presFile, ok := fileMap["ppt/presentation.xml"]; ok {
		presData, _ := readZipFile(presFile)
		doc.SlideWidth, doc.SlideHeight = parseSlideDimensions(presData)
	}
	if doc.SlideWidth == 0 {
		doc.SlideWidth = 960  // default 10" * 96dpi
		doc.SlideHeight = 540
	}

	// Parse theme color scheme
	themeColors := parseThemeColors(fileMap)

	// Discover slides
	slideNums := discoverSlides(fileMap)
	doc.SlideCount = len(slideNums)

	for _, si := range slideNums {
		slidePath := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		slideRelsPath := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", si)

		slideFile, ok := fileMap[slidePath]
		if !ok {
			continue
		}
		slideData, _ := readZipFile(slideFile)

		// Parse slide rels
		slideRels := parseRels(fileMap, slideRelsPath)

		// Find layout path from rels
		layoutPath := ""
		for _, rel := range slideRels {
			if strings.Contains(rel.Type, "slideLayout") {
				layoutPath = resolveRelPath("ppt/slides/", rel.Target)
				break
			}
		}

		// Parse layout rels for media
		layoutRelsPath := ""
		if layoutPath != "" {
			layoutRelsPath = relsPathFor(layoutPath)
		}
		layoutRels := parseRels(fileMap, layoutRelsPath)

		// Parse master rels (from layout)
		masterPath := ""
		for _, rel := range layoutRels {
			if strings.Contains(rel.Type, "slideMaster") {
				masterPath = resolveRelPath(filepath.Dir(layoutPath)+"/", rel.Target)
				break
			}
		}
		masterRelsPath := ""
		if masterPath != "" {
			masterRelsPath = relsPathFor(masterPath)
		}
		masterRels := parseRels(fileMap, masterRelsPath)

		// Build rId → media file maps
		slideMediaMap := buildMediaMap(slideRels)
		layoutMediaMap := buildMediaMap(layoutRels)
		masterMediaMap := buildMediaMap(masterRels)

		// Parse shapes from all layers
		parsed := ParsedSlide{
			SlideIndex: si,
		}

		globalID := 1

		// 1. Master shapes — 이미지/배경 도형만 가져오고, 플레이스홀더 텍스트는 제거
		if masterPath != "" {
			if masterFile, ok := fileMap[masterPath]; ok {
				masterData, _ := readZipFile(masterFile)
				resolveSchemeColors(masterData, themeColors)
				shapes := extractShapes(masterData, masterMediaMap, mediaOutputDir, &globalID, si, fileMap)
				for _, s := range shapes {
					if s.PhType == "" {
						parsed.Shapes = append(parsed.Shapes, s)
					} else if s.ImagePath != "" || s.FillType == "solid" || s.FillType == "gradient" {
						s.HasText = false
						s.TextRuns = nil
						parsed.Shapes = append(parsed.Shapes, s)
					}
				}
			}
		}

		// 2. Layout shapes — 비-플레이스홀더 도형 전부 가져오기 + PLACEHOLDER 스타일 맵 빌드
		phStyleMap := make(map[string]ParsedShape) // phType → 레이아웃의 기본 스타일
		if layoutPath != "" {
			if layoutFile, ok := fileMap[layoutPath]; ok {
				layoutData, _ := readZipFile(layoutFile)
				resolveSchemeColors(layoutData, themeColors)
				shapes := extractShapes(layoutData, layoutMediaMap, mediaOutputDir, &globalID, si, fileMap)
				for _, s := range shapes {
					if s.PhType == "" {
						parsed.Shapes = append(parsed.Shapes, s)
					} else if s.ImagePath != "" || s.FillType == "solid" || s.FillType == "gradient" {
						s.HasText = false
						s.TextRuns = nil
						parsed.Shapes = append(parsed.Shapes, s)
					}
				}
				// [Q1-A] fillRef 도형 별도 파싱 — parseFill이 noFill로 판정한 도형 중
				// p:style에 fillRef가 있고 placeholder가 아닌 것을 배경 도형으로 추가
				layoutStr2 := string(layoutData)
				spBlocks := regexp.MustCompile(`(?s)<p:sp>.*?</p:sp>`).FindAllString(layoutStr2, -1)
				for _, block := range spBlocks {
					// placeholder는 건너뜀
					if strings.Contains(block, `<p:ph `) {
						continue
					}
					// p:style 블록에서 fillRef 찾기
					fillRefRe := regexp.MustCompile(`(?s)<a:fillRef idx="(\d+)"[^>]*>.*?<a:schemeClr val="([^"]+)"`)
					styleRe := regexp.MustCompile(`(?s)(?s)<p:style[^>]*>(.*?)</p:style>`)
					sm := styleRe.FindStringSubmatch(block)
					if sm == nil {
						continue
					}
					fm := fillRefRe.FindStringSubmatch(sm[1])
					if fm == nil {
						continue
					}
					idx, _ := strconv.Atoi(fm[1])
					if idx == 0 {
						continue
					}
					schemeName := fm[2]
					color, ok := themeColors[schemeName]
					if !ok {
						continue
					}
					// 이미 extractShapes에서 이 도형이 solid/gradient로 추가되었으면 건너뜀
					// (중복 방지: name으로 체크)
					nameRe := regexp.MustCompile(`name="([^"]*)"`)
					nm := nameRe.FindStringSubmatch(block)
					shapeName := ""
					if nm != nil {
						shapeName = nm[1]
					}
					alreadyAdded := false
					for _, existing := range parsed.Shapes {
						if existing.Name == shapeName && existing.FillType != "" && existing.FillType != "none" {
							alreadyAdded = true
							break
						}
					}
					if alreadyAdded {
						continue
					}
					// ParsedShape 생성 — 배경이므로 맨 앞에 삽입 (z-order)
					var fillRefShape ParsedShape
					fillRefShape.ID = globalID
					globalID++
					fillRefShape.Type = "AUTO"
					fillRefShape.Name = shapeName
					fillRefShape.FillType = "solid"
					fillRefShape.FillColor = color
					parseXfrm(block, &fillRefShape)
					fillRefShape.HasText = false
					parsed.Shapes = append([]ParsedShape{fillRefShape}, parsed.Shapes...)
				}
				// [Q1-C] 마스터/레이아웃 PLACEHOLDER의 xfrm + defRPr에서 기본 스타일 추출
				// 1. 마스터에서 먼저 추출 (sldNum 등 공통 요소)
				if masterPath != "" {
					if masterFile, ok := fileMap[masterPath]; ok {
						masterData, _ := readZipFile(masterFile)
						masterStr := string(masterData)
						phBlocks := regexp.MustCompile(`(?s)<p:sp>.*?</p:sp>`).FindAllString(masterStr, -1)
						for _, block := range phBlocks {
							phType := ""
							if strings.Contains(block, `type="title"`) {
								phType = "title"
							} else if strings.Contains(block, `type="ctrTitle"`) {
								phType = "ctrTitle"
							} else if strings.Contains(block, `type="subTitle"`) {
								phType = "subTitle"
							} else if strings.Contains(block, `type="body"`) {
								phType = "body"
							} else if strings.Contains(block, `type="sldNum"`) {
								phType = "sldNum"
							} else if strings.Contains(block, `type="dt"`) {
								phType = "dt"
							} else if strings.Contains(block, `type="ftr"`) {
								phType = "ftr"
							}
							if phType == "" {
								continue
							}
							var phShape ParsedShape
							parseXfrm(block, &phShape)
							if m := regexp.MustCompile(`<a:defRPr[^>]*\ssz="(\d+)"`).FindStringSubmatch(block); m != nil {
								sz, _ := strconv.Atoi(m[1])
								phShape.FillType = fmt.Sprintf("%.1f", float64(sz)/100.0)
							}
							if regexp.MustCompile(`<a:defRPr[^>]*\sb="1"`).MatchString(block) {
								phShape.ShapeType = "bold"
							}
							if m := regexp.MustCompile(`<a:defRPr[^>]*>.*?<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"`).FindStringSubmatch(block); m != nil {
								phShape.FillColor = "#" + m[1]
							} else if m := regexp.MustCompile(`(?s)<a:defRPr[^>]*>.*?<a:solidFill>.*?<a:schemeClr val="([^"]+)"`).FindStringSubmatch(block); m != nil {
								if c, ok := themeColors[m[1]]; ok {
									phShape.FillColor = c
								}
							}
							phStyleMap[phType] = phShape
						}
					}
				}

				// 2. 레이아웃에서 덮어쓰기 (레이아웃에 오버라이드된 요소)
				layoutStr := string(layoutData)
				phBlocks := regexp.MustCompile(`(?s)<p:sp>.*?</p:sp>`).FindAllString(layoutStr, -1)
				for _, block := range phBlocks {
					phType := ""
					if strings.Contains(block, `type="title"`) {
						phType = "title"
					} else if strings.Contains(block, `type="ctrTitle"`) {
						phType = "ctrTitle"
					} else if strings.Contains(block, `type="subTitle"`) {
						phType = "subTitle"
					} else if strings.Contains(block, `type="body"`) {
						phType = "body"
					} else if strings.Contains(block, `type="sldNum"`) {
						phType = "sldNum"
					} else if strings.Contains(block, `type="dt"`) {
						phType = "dt"
					} else if strings.Contains(block, `type="ftr"`) {
						phType = "ftr"
					}
					if phType == "" {
						continue
					}
					// 마스터 속성을 복사한 뒤 덮어쓰기
					phShape := phStyleMap[phType]
					parseXfrm(block, &phShape)
					// defRPr에서 기본 폰트 크기/색상/bold 추출
					if m := regexp.MustCompile(`<a:defRPr[^>]*\ssz="(\d+)"`).FindStringSubmatch(block); m != nil {
						sz, _ := strconv.Atoi(m[1])
						phShape.FillType = fmt.Sprintf("%.1f", float64(sz)/100.0) // 임시 저장
					}
					// [Q1-C] bold 추출
					if regexp.MustCompile(`<a:defRPr[^>]*\sb="1"`).MatchString(block) {
						phShape.ShapeType = "bold" // 임시 플래그
					}
					// 색상 — srgbClr 직접 지정
					if m := regexp.MustCompile(`<a:defRPr[^>]*>.*?<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"`).FindStringSubmatch(block); m != nil {
						phShape.FillColor = "#" + m[1]
					} else if m := regexp.MustCompile(`(?s)<a:defRPr[^>]*>.*?<a:solidFill>.*?<a:schemeClr val="([^"]+)"`).FindStringSubmatch(block); m != nil {
						// [Q1-B] schemeClr → themeColors로 해석
						if c, ok := themeColors[m[1]]; ok {
							phShape.FillColor = c
						}
					}
					phStyleMap[phType] = phShape
				}
			}
		}

		// 3. Slide shapes — 전부 유지 + 레이아웃 PLACEHOLDER 스타일 상속
		shapes := extractShapes(slideData, slideMediaMap, mediaOutputDir, &globalID, si, fileMap)
		for i := range shapes {
			s := &shapes[i]
			if s.Type != "PLACEHOLDER" {
				continue
			}
			// PhType으로 레이아웃 스타일 직접 매칭
			ph, ok := phStyleMap[s.PhType]
			if !ok {
				continue
			}
			// 좌표 상속
			if s.W == 0 && s.H == 0 && ph.W > 0 {
				s.X = ph.X
				s.Y = ph.Y
				s.W = ph.W
				s.H = ph.H
			}
			// 폰트 크기 상속 (defRPr sz)
			if ph.FillType != "" {
				defSize := 0.0
				fmt.Sscanf(ph.FillType, "%f", &defSize)
				if defSize > 0 {
					for j := range s.TextRuns {
						if s.TextRuns[j].Size == 11 {
							s.TextRuns[j].Size = defSize
						}
					}
				}
			}
			// [Q1-C] bold 상속
			if ph.ShapeType == "bold" {
				for j := range s.TextRuns {
					if s.TextRuns[j].Bold == 0 {
						s.TextRuns[j].Bold = -1
					}
				}
			}
			// 색상 상속
			if ph.FillColor != "" {
				for j := range s.TextRuns {
					if s.TextRuns[j].Color == "" {
						s.TextRuns[j].Color = ph.FillColor
					}
				}
			}
		}
		parsed.Shapes = append(parsed.Shapes, shapes...)

		// [PARITY] 슬라이드에 전체화면 배경 도형(이미지/그라디언트)이 있으면
		// 마스터/레이아웃 상속 도형은 실제 PPT에서도 가려지므로 제거
		slideArea := doc.SlideWidth * doc.SlideHeight
		if slideArea > 0 {
			hasFullCover := false
			for _, s := range shapes {
				shapeArea := s.W * s.H
				if shapeArea / slideArea >= 0.85 && (s.ImagePath != "" || s.FillType == "gradient" || (s.FillType == "solid" && s.FillColor != "" && s.FillColor != "transparent")) {
					hasFullCover = true
					break
				}
			}
			if hasFullCover {
				// 마스터/레이아웃에서 온 도형 제거 (슬라이드 도형만 유지)
				masterLayoutCount := len(parsed.Shapes) - len(shapes)
				if masterLayoutCount > 0 {
					parsed.Shapes = parsed.Shapes[masterLayoutCount:]
				}
			}
		}

		// Background color: slide → layout → master chain
		parsed.BackgroundColor = extractBgColor(slideData)
		if parsed.BackgroundColor == "" && layoutPath != "" {
			if lf, ok := fileMap[layoutPath]; ok {
				ld, _ := readZipFile(lf)
				parsed.BackgroundColor = extractBgColor(ld)
			}
		}
		if parsed.BackgroundColor == "" && masterPath != "" {
			if mf, ok := fileMap[masterPath]; ok {
				md, _ := readZipFile(mf)
				parsed.BackgroundColor = extractBgColor(md)
			}
		}

		// Resolve scheme: references to actual colors
		if strings.HasPrefix(parsed.BackgroundColor, "scheme:") {
			key := strings.TrimPrefix(parsed.BackgroundColor, "scheme:")
			if c, ok := themeColors[key]; ok {
				parsed.BackgroundColor = c
			}
		}
		for i := range parsed.Shapes {
			s := &parsed.Shapes[i]
		if strings.HasPrefix(s.FillColor, "scheme:") {
				raw := strings.TrimPrefix(s.FillColor, "scheme:")
				// scheme:bg1:lum:95000:0 형식 처리
				if strings.Contains(raw, ":lum:") {
					parts := strings.Split(raw, ":lum:")
					key := parts[0]
					lumParts := strings.Split(parts[1], ":")
					lumMod := 100000.0
					lumOff := 0.0
					if len(lumParts) >= 1 { lumMod, _ = strconv.ParseFloat(lumParts[0], 64) }
					if len(lumParts) >= 2 { lumOff, _ = strconv.ParseFloat(lumParts[1], 64) }
					if baseColor, ok := themeColors[key]; ok {
						s.FillColor = applyLumModOff(baseColor, lumMod/100000.0, lumOff/100000.0)
					}
				} else {
					if c, ok := themeColors[raw]; ok {
						s.FillColor = c
					}
				}
			}
			for j := range s.TextRuns {
				if strings.HasPrefix(s.TextRuns[j].Color, "scheme:") {
					key := strings.TrimPrefix(s.TextRuns[j].Color, "scheme:")
					if c, ok := themeColors[key]; ok {
						s.TextRuns[j].Color = c
					}
				}
			}
			// Replace scheme in table data
			if len(s.TableData) > 0 {
				for j, td := range s.TableData {
					tdStr := string(td)
					for k, v := range themeColors {
						tdStr = strings.ReplaceAll(tdStr, `"scheme:`+k+`"`, `"`+v+`"`)
					}
					s.TableData[j] = json.RawMessage(tdStr)
				}
			}
		}

		doc.Slides = append(doc.Slides, parsed)
	}

	doc.ThemeColors = themeColors

	return doc, nil
}

func discoverSlides(fileMap map[string]*zip.File) []int {
	re := regexp.MustCompile(`ppt/slides/slide(\d+)\.xml$`)
	var nums []int
	for name := range fileMap {
		if m := re.FindStringSubmatch(name); m != nil {
			n, _ := strconv.Atoi(m[1])
			nums = append(nums, n)
		}
	}
	sort.Ints(nums)
	return nums
}

func parseRels(fileMap map[string]*zip.File, path string) []xmlRelationship {
	if path == "" {
		return nil
	}
	f, ok := fileMap[path]
	if !ok {
		return nil
	}
	data, err := readZipFile(f)
	if err != nil {
		return nil
	}
	var rels xmlRelationships
	xml.Unmarshal(data, &rels)
	return rels.Rels
}

func buildMediaMap(rels []xmlRelationship) map[string]string {
	m := make(map[string]string)
	for _, rel := range rels {
		if strings.Contains(rel.Type, "image") || strings.Contains(rel.Type, "media") {
			target := filepath.Base(rel.Target)
			m[rel.ID] = target
		}
	}
	return m
}

func resolveRelPath(base, target string) string {
	// target is like "../slideLayouts/slideLayout1.xml"
	combined := filepath.Join(base, target)
	// Normalize
	combined = filepath.ToSlash(combined)
	// Clean ../
	parts := strings.Split(combined, "/")
	var result []string
	for _, p := range parts {
		if p == ".." && len(result) > 0 {
			result = result[:len(result)-1]
		} else if p != "." && p != "" {
			result = append(result, p)
		}
	}
	return strings.Join(result, "/")
}

func relsPathFor(xmlPath string) string {
	dir := filepath.ToSlash(filepath.Dir(xmlPath))
	base := filepath.Base(xmlPath)
	return dir + "/_rels/" + base + ".rels"
}

// findBalancedTags finds all top-level balanced occurrences of <tagName>...</tagName>
// Returns slice of [start, end] byte offsets. Handles nested tags correctly.
func findBalancedTags(xml string, tagName string) [][2]int {
	openTag := "<" + tagName
	closeTag := "</" + tagName + ">"
	var results [][2]int

	i := 0
	for i < len(xml) {
		// Find next opening tag
		start := strings.Index(xml[i:], openTag)
		if start == -1 {
			break
		}
		start += i
		// Verify it's a proper tag (followed by > or space)
		afterTag := start + len(openTag)
		if afterTag >= len(xml) {
			break
		}
		ch := xml[afterTag]
		if ch != '>' && ch != ' ' && ch != '\n' && ch != '\r' && ch != '\t' {
			i = afterTag
			continue
		}

		// Count depth
		depth := 1
		searchFrom := afterTag
		for depth > 0 && searchFrom < len(xml) {
			nextOpen := strings.Index(xml[searchFrom:], openTag)
			nextClose := strings.Index(xml[searchFrom:], closeTag)

			if nextClose == -1 {
				break // Malformed XML
			}
			nextClose += searchFrom

			if nextOpen != -1 {
				nextOpen += searchFrom
				// Verify it's a proper tag (not e.g. <p:grpSpPr> when searching <p:grpSp>)
				na := nextOpen + len(openTag)
				validOpen := na < len(xml) && (xml[na] == '>' || xml[na] == ' ' || xml[na] == '\n' || xml[na] == '\r' || xml[na] == '\t')

				if nextOpen < nextClose {
					if validOpen {
						depth++
						searchFrom = nextOpen + len(openTag)
						continue
					} else {
						// Invalid open tag (e.g. <p:grpSpPr>), skip past it and re-search
						searchFrom = na
						continue
					}
				}
			}

			depth--
			if depth == 0 {
				end := nextClose + len(closeTag)
				results = append(results, [2]int{start, end})
			}
			searchFrom = nextClose + len(closeTag)
		}
		i = searchFrom
	}
	return results
}
func preprocessXML(xmlStr string) string {
	// [Quark 11] Flatten mc:AlternateContent Fallback
	// 차트나 스마트아트가 VML/기본 도형으로 캐시된 호환성 블록을 최상단으로 끌어올림 (승격)
	fallbackRe := regexp.MustCompile(`(?s)<mc:AlternateContent[^>]*>.*?<mc:Fallback[^>]*>(.*?)</mc:Fallback>.*?</mc:AlternateContent>`)
	xmlStr = fallbackRe.ReplaceAllString(xmlStr, "$1")

	// [Quark 8 & 11] Recursive Group Coordinate Flattening (Side-effect free math)
	// 중첩된(Nested) <p:grpSp>를 모두 풀기 위해 변경 사항이 없을 때까지 반복
	// MUST run BEFORE placeholder flattening to ensure balanced tags
	for i := 0; i < 5; i++ { // Limit recursion to 5
		grpBlocks := findBalancedTags(xmlStr, "p:grpSp")
		if len(grpBlocks) == 0 {
			break
		}

		// Replace from back to front to preserve offsets
		for j := len(grpBlocks) - 1; j >= 0; j-- {
			block := grpBlocks[j]
			grpBlock := xmlStr[block[0]:block[1]]
			replacement := func() string {
			// Extract parent coordinates
			offRe := regexp.MustCompile(`<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"`)
			extRe := regexp.MustCompile(`<a:ext\s+cx="(\d+)"\s+cy="(\d+)"`)
			chOffRe := regexp.MustCompile(`<a:chOff\s+x="(-?\d+)"\s+y="(-?\d+)"`)
			chExtRe := regexp.MustCompile(`<a:chExt\s+cx="(\d+)"\s+cy="(\d+)"`)

			// Only get the grpSpPr block to avoid matching child coordinates as parent
			grpPrRe := regexp.MustCompile(`(?s)<p:grpSpPr>(.*?)</p:grpSpPr>`)
			grpPrMatch := grpPrRe.FindString(grpBlock)
			
			pOffM := offRe.FindStringSubmatch(grpPrMatch)
			pExtM := extRe.FindStringSubmatch(grpPrMatch)
			pChOffM := chOffRe.FindStringSubmatch(grpPrMatch)
			pChExtM := chExtRe.FindStringSubmatch(grpPrMatch)

			if pOffM == nil || pExtM == nil || pChOffM == nil || pChExtM == nil {
				return grpBlock // fallback if missing transforms
			}

			pX, _ := strconv.ParseFloat(pOffM[1], 64)
			pY, _ := strconv.ParseFloat(pOffM[2], 64)
			pCx, _ := strconv.ParseFloat(pExtM[1], 64)
			pCy, _ := strconv.ParseFloat(pExtM[2], 64)
			chX, _ := strconv.ParseFloat(pChOffM[1], 64)
			chY, _ := strconv.ParseFloat(pChOffM[2], 64)
			chCx, _ := strconv.ParseFloat(pChExtM[1], 64)
			chCy, _ := strconv.ParseFloat(pChExtM[2], 64)

			if chCx == 0 { chCx = 1 }
			if chCy == 0 { chCy = 1 }
			scaleX := pCx / chCx
			scaleY := pCy / chCy


			// Flatten function for internal elements
			flattenFunc := func(childBlock string) string {
				// Replace off
				childBlock = offRe.ReplaceAllStringFunc(childBlock, func(m string) string {
					subM := offRe.FindStringSubmatch(m)
					cX, _ := strconv.ParseFloat(subM[1], 64)
					cY, _ := strconv.ParseFloat(subM[2], 64)
					newX := int64(pX + (cX - chX) * scaleX)
					newY := int64(pY + (cY - chY) * scaleY)
					return fmt.Sprintf(`<a:off x="%d" y="%d"`, newX, newY)
				})
				// Replace ext
				childBlock = extRe.ReplaceAllStringFunc(childBlock, func(m string) string {
					subM := extRe.FindStringSubmatch(m)
					cCx, _ := strconv.ParseFloat(subM[1], 64)
					cCy, _ := strconv.ParseFloat(subM[2], 64)
					newCx := int64(cCx * scaleX)
					newCy := int64(cCy * scaleY)
					return fmt.Sprintf(`<a:ext cx="%d" cy="%d"`, newCx, newCy)
				})
				return childBlock
			}

			// Extract and flatten children
			var flattenedChildren string

			// First, find inner grpSp blocks and extract them separately

			grpPrEndRe := regexp.MustCompile(`(?s)</p:grpSpPr>`)
			if grpPrEnd := grpPrEndRe.FindStringIndex(grpBlock); grpPrEnd != nil {
				innerContent := grpBlock[grpPrEnd[1]:]
				// Remove the closing </p:grpSp> of the parent
				if idx := strings.LastIndex(innerContent, "</p:grpSp>"); idx >= 0 {
					innerContent = innerContent[:idx]
				}
				innerGrps := findBalancedTags(innerContent, "p:grpSp")
				// Fallback: if findBalancedTags missed inner groups due to unbalanced XML
				// from prior preprocessing, try regex approach
				if len(innerGrps) == 0 && strings.Contains(innerContent, "<p:grpSp>") {
					// Greedy regex to find inner grpSp blocks
					innerGrpRe := regexp.MustCompile(`(?s)<p:grpSp>.*</p:grpSp>`)
					if m := innerGrpRe.FindStringIndex(innerContent); m != nil {
						innerGrps = [][2]int{{m[0], m[1]}}
					} else {
						// No closing tag found — extract from <p:grpSp> to end
						startIdx := strings.Index(innerContent, "<p:grpSp>")
						if startIdx >= 0 {
							innerGrps = [][2]int{{startIdx, len(innerContent)}}
						}
					}
				}
				// Collect ALL children (sp, pic, graphicFrame, grpSp) with XML offsets
				// and sort by offset to preserve original PPT layer order (z-index)
				type childEntry struct {
					offset  int
					content string
				}
				var orderedChildren []childEntry

				// grpSp children - transform only grpSpPr
				for _, ig := range innerGrps {
					child := innerContent[ig[0]:ig[1]]
					grpPrInner := regexp.MustCompile(`(?s)<p:grpSpPr>(.*?)</p:grpSpPr>`)
					if prM := grpPrInner.FindStringIndex(child); prM != nil {
						prBlock := child[prM[0]:prM[1]]
						transformedPr := flattenFunc(prBlock)
						child = child[:prM[0]] + transformedPr + child[prM[1]:]
					}
					orderedChildren = append(orderedChildren, childEntry{ig[0], child})
				}

				// Helper: check if offset is inside any grpSp range
				isInsideGrpSp := func(start int) bool {
					for _, r := range innerGrps {
						if start >= r[0] && start < r[1] {
							return true
						}
					}
					return false
				}

				// sp, pic, graphicFrame children (only direct, not inside grpSp)
				for _, tag := range []string{"p:sp", "p:pic", "p:graphicFrame"} {
					matches := findBalancedTags(innerContent, tag)
					for _, m := range matches {
						if !isInsideGrpSp(m[0]) {
							orderedChildren = append(orderedChildren, childEntry{m[0], flattenFunc(innerContent[m[0]:m[1]])})
						}
					}
				}

				// Sort by XML offset to preserve layer order
				sort.Slice(orderedChildren, func(i, j int) bool {
					return orderedChildren[i].offset < orderedChildren[j].offset
				})

				// Build flattenedChildren in correct order
				flattenedChildren = ""
				for _, c := range orderedChildren {
					flattenedChildren += c.content + "\n"
				}
			}
			
			// We return ONLY the flattened children, effectively destroying the <p:grpSp> shell
			// This completely solves nested rendering and coordinate inheritance bugs!
			if flattenedChildren != "" {
				return flattenedChildren
			}
			return grpBlock
			}()
			xmlStr = xmlStr[:block[0]] + replacement + xmlStr[block[1]:]
		}
	}
	
	// [Quark 7] Placeholder Flattening (Side-effect free)
	// Runs AFTER group flattening to preserve tag balance during grpSp processing
	spRe := regexp.MustCompile(`(?s)<p:sp\b[^>]*>.*?</p:sp>`)
	xmlStr = spRe.ReplaceAllStringFunc(xmlStr, func(spBlock string) string {
		if strings.Contains(spBlock, "<p:ph") {
			if strings.Contains(spBlock, "<p:spPr/>") {
				return strings.Replace(spBlock, "<p:spPr/>", "<p:spPr><a:noFill/></p:spPr>", 1)
			}
			spPrRe := regexp.MustCompile(`(?s)<p:spPr[^>]*>.*?</p:spPr>`)
			if spPrMatch := spPrRe.FindString(spBlock); spPrMatch != "" {
				if !strings.Contains(spPrMatch, "Fill") && !strings.Contains(spPrMatch, "schemeClr") {
					newSpPr := strings.Replace(spPrMatch, "</p:spPr>", "<a:noFill/></p:spPr>", 1)
					return strings.Replace(spBlock, spPrMatch, newSpPr, 1)
				}
			}
		}
		return spBlock
	})

	return xmlStr
}

func extractShapes(xmlData []byte, mediaMap map[string]string, mediaDir string, globalID *int, slideIdx int, fileMap map[string]*zip.File) []ParsedShape {
	var shapes []ParsedShape
	
	// [전처리 전] 원본 XML에서 도형 순서를 기록 (그룹 내 자식 포함)
	// grpSp 내 자식은 그룹의 순서를 상속 (그룹 위치 + 자식 내 순서)
	rawContent := string(xmlData)
	origOrder := map[string]int{} // name → 순서
	orderIdx := 0
	nameRe := regexp.MustCompile(`name="([^"]+)"`)
	// 그룹 내 자식 순서도 포함하여 모든 cNvPr name을 순서대로 추출
	for _, m := range nameRe.FindAllStringSubmatch(rawContent, -1) {
		if _, exists := origOrder[m[1]]; !exists {
			origOrder[m[1]] = orderIdx
			orderIdx++
		}
	}
	
	content := preprocessXML(rawContent)

	// preprocessXML에서 grpSp가 이미 평탄화되어 자식 sp/pic이 content에 직접 존재
	spPattern := regexp.MustCompile(`(?s)<p:sp\b[^>]*>.*?</p:sp>`)
	picPattern := regexp.MustCompile(`(?s)<p:pic\b[^>]*>.*?</p:pic>`)
	gfPattern := regexp.MustCompile(`(?s)<p:graphicFrame\b[^>]*>.*?</p:graphicFrame>`)

	// sp (text/auto shapes)
	for _, idx := range spPattern.FindAllStringIndex(content, -1) {
		match := content[idx[0]:idx[1]]
		shp := parseSpElement(match, mediaMap, mediaDir, globalID)
		if shp != nil {
			shp.XmlOffset = idx[0]
			shapes = append(shapes, *shp)
		}
	}

	// pic (picture shapes)
	for _, idx := range picPattern.FindAllStringIndex(content, -1) {
		match := content[idx[0]:idx[1]]
		shp := parsePicElement(match, mediaMap, mediaDir, globalID)
		if shp != nil {
			shp.XmlOffset = idx[0]
			shapes = append(shapes, *shp)
		}
	}

	// graphicFrame (tables)
	for _, idx := range gfPattern.FindAllStringIndex(content, -1) {
		match := content[idx[0]:idx[1]]
		shp := parseGraphicFrame(match, globalID, fileMap, slideIdx)
		if shp != nil {
			shp.XmlOffset = idx[0]
			shapes = append(shapes, *shp)
		}
	}

	// [Quark 6] Z-index 정렬: XML 문서 내 등장 순서(XmlOffset) 기준 정렬
	// 이름 기반 정렬은 동명 도형(예: "Text Box 52" 중복)에서 순서가 깨지므로 폐기
	sort.SliceStable(shapes, func(i, j int) bool {
		return shapes[i].XmlOffset < shapes[j].XmlOffset
	})

	return shapes
}

func extractInnerGroupShapes(grpXML string, mediaMap map[string]string, mediaDir string, globalID *int) []ParsedShape {
	var shapes []ParsedShape

	// Inner sp
	spRe := regexp.MustCompile(`(?s)<p:sp\b[^>]*>.*?</p:sp>`)
	for _, m := range spRe.FindAllString(grpXML, -1) {
		shp := parseSpElement(m, mediaMap, mediaDir, globalID)
		if shp != nil {
			shapes = append(shapes, *shp)
		}
	}

	// Inner pic
	picRe := regexp.MustCompile(`(?s)<p:pic\b[^>]*>.*?</p:pic>`)
	for _, m := range picRe.FindAllString(grpXML, -1) {
		shp := parsePicElement(m, mediaMap, mediaDir, globalID)
		if shp != nil {
			shapes = append(shapes, *shp)
		}
	}

	// Inner graphicFrame (tables)
	gfRe := regexp.MustCompile(`(?s)<p:graphicFrame\b[^>]*>.*?</p:graphicFrame>`)
	for _, m := range gfRe.FindAllString(grpXML, -1) {
		shp := parseGraphicFrame(m, globalID, nil, 0)
		if shp != nil {
			shapes = append(shapes, *shp)
		}
	}

	return shapes
}

func parseGraphicFrame(gfXML string, id *int, fileMap map[string]*zip.File, slideIdx int) *ParsedShape {
	// Handle charts
	if strings.Contains(gfXML, "drawingml/2006/chart") {
		return parseChartFrame(gfXML, id, fileMap, slideIdx)
	}
	// Only handle tables (a:tbl)
	if !strings.Contains(gfXML, "<a:tbl") {
		return nil
	}

	shp := &ParsedShape{ID: *id, Type: "TABLE"}
	*id++

	if m := regexp.MustCompile(`name="([^"]+)"`).FindStringSubmatch(gfXML); m != nil {
		shp.Name = m[1]
	}

	parseXfrm(gfXML, shp)

	// Parse column widths from a:gridCol (EMU → px: 1px = 9525 EMU)
	gridColRe := regexp.MustCompile(`<a:gridCol w="(\d+)"`)
	for _, m := range gridColRe.FindAllStringSubmatch(gfXML, -1) {
		w, _ := strconv.Atoi(m[1])
		shp.TableColWidths = append(shp.TableColWidths, float64(w)/9525.0)
	}

	// Parse table rows and cells
	rowRe := regexp.MustCompile(`(?s)<a:tr\b[^>]*>(.*?)</a:tr>`)
	rowHRe := regexp.MustCompile(`<a:tr[^>]*\bh="(\d+)"`)
	cellRe := regexp.MustCompile(`(?s)<a:tc\b[^>]*>(.*?)</a:tc>`)
	textRe := regexp.MustCompile(`(?s)<a:t>(.*?)</a:t>`)
	colorRe := regexp.MustCompile(`<a:srgbClr val="([0-9A-Fa-f]{6})"`)
	boldRe := regexp.MustCompile(`\bb="1"`)
	sizeRe := regexp.MustCompile(`\bsz="(\d+)"`)

	rows := rowRe.FindAllStringSubmatch(gfXML, -1)
	shp.TableRows = len(rows)

	// Extract row heights
	for _, m := range rowHRe.FindAllStringSubmatch(gfXML, -1) {
		h, _ := strconv.Atoi(m[1])
		shp.TableRowHeights = append(shp.TableRowHeights, float64(h)/9525.0)
	}

	for _, rowMatch := range rows {
		rowXML := rowMatch[1]
		cells := cellRe.FindAllStringSubmatch(rowXML, -1)
		if shp.TableCols == 0 {
			shp.TableCols = len(cells)
		}

		var rowCells []TableCellData
		for _, cellMatch := range cells {
			cellXML := cellMatch[1]
			cellText := ""
			// 줄바꿈 처리 및 bullet 파싱
			cellXMLBr := regexp.MustCompile(`<a:br\b[^>]*/>|(?s)<a:br\b[^>]*>.*?</a:br>`).ReplaceAllString(cellXML, "<a:t>\n</a:t>")
			paraBulletRe := regexp.MustCompile(`(?s)<a:p\b[^>]*>(.*?)</a:p>`)
			buCharRe := regexp.MustCompile(`<a:buChar\s+char="([^"]+)"`)
			buNoneRe := regexp.MustCompile(`<a:buNone`)
			cellParas := paraBulletRe.FindAllStringSubmatch(cellXMLBr, -1)
			if len(cellParas) > 0 {
				for pi, para := range cellParas {
					paraContent := para[1]
					prefix := ""
					if bm := buCharRe.FindStringSubmatch(paraContent); bm != nil {
						if !buNoneRe.MatchString(paraContent) {
							prefix = bm[1] + " "
						}
					}
					paraText := ""
					for _, t := range textRe.FindAllStringSubmatch(paraContent, -1) {
						paraText += t[1]
					}
					if prefix != "" && len(strings.TrimSpace(paraText)) > 0 {
						paraText = prefix + paraText
					}
					cellText += paraText
					if pi < len(cellParas)-1 {
						cellText += "\n"
					}
				}
			} else {
				cellXMLBr = regexp.MustCompile(`</a:p>\s*<a:p\b`).ReplaceAllString(cellXMLBr, "</a:p>\n<a:p")
				for _, t := range textRe.FindAllStringSubmatch(cellXMLBr, -1) {
					cellText += t[1]
				}
			}
			// Decode XML entities
			cellText = strings.ReplaceAll(cellText, "&amp;", "&")
			cellText = strings.ReplaceAll(cellText, "&lt;", "<")
			cellText = strings.ReplaceAll(cellText, "&gt;", ">")

			cell := TableCellData{Text: cellText}

			// Text color and bold: from a:rPr only
			rPrRe := regexp.MustCompile(`(?s)<a:rPr[^>]*>(.*?)</a:rPr>`)
			schemeRe := regexp.MustCompile(`<a:schemeClr val="([^"]+)"`)
			if rPr := rPrRe.FindStringSubmatch(cellXML); rPr != nil {
				if m := colorRe.FindStringSubmatch(rPr[1]); m != nil {
					cell.Color = "#" + m[1]
				} else if m := schemeRe.FindStringSubmatch(rPr[1]); m != nil {
					cell.Color = "scheme:" + m[1]
				}
				if boldRe.MatchString(rPr[0]) {
					cell.Bold = -1
				}
			}
			
			// 텍스트 정렬 (가로: a:pPr algn, 세로: a:tcPr anchor)
			algnRe := regexp.MustCompile(`<a:pPr[^>]*\balgn="([a-zA-Z]+)"`)
			if am := algnRe.FindStringSubmatch(cellXML); am != nil {
				cell.Align = am[1]
			}
			valignRe := regexp.MustCompile(`<a:tcPr[^>]*\banchor="([a-zA-Z]+)"`)
			if vm := valignRe.FindStringSubmatch(cellXML); vm != nil {
				cell.VAlign = vm[1]
			}
			
			if m := sizeRe.FindStringSubmatch(cellXML); m != nil {
				sz, _ := strconv.Atoi(m[1])
				cell.Size = float64(sz) / 100.0
			}

			// Cell background: tcPr direct child solidFill only (not inside lnL/lnR/lnT/lnB)
			tcPrRe := regexp.MustCompile(`(?s)<a:tcPr([^>]*)>(.*?)</a:tcPr>`)
			if tcPrMatch := tcPrRe.FindStringSubmatch(cellXML); tcPrMatch != nil {
				attrStr := tcPrMatch[1]
				tcPrContent := tcPrMatch[2]

				// [Quark 2] Parse cell margin/padding
				if m := regexp.MustCompile(`\bmarL="(\d+)"`).FindStringSubmatch(attrStr); m != nil {
					v, _ := strconv.ParseFloat(m[1], 64)
					cell.MarL = v / 9525.0
				} else { cell.MarL = -1 } // -1 means default
				if m := regexp.MustCompile(`\bmarR="(\d+)"`).FindStringSubmatch(attrStr); m != nil {
					v, _ := strconv.ParseFloat(m[1], 64)
					cell.MarR = v / 9525.0
				} else { cell.MarR = -1 }
				if m := regexp.MustCompile(`\bmarT="(\d+)"`).FindStringSubmatch(attrStr); m != nil {
					v, _ := strconv.ParseFloat(m[1], 64)
					cell.MarT = v / 9525.0
				} else { cell.MarT = -1 }
				if m := regexp.MustCompile(`\bmarB="(\d+)"`).FindStringSubmatch(attrStr); m != nil {
					v, _ := strconv.ParseFloat(m[1], 64)
					cell.MarB = v / 9525.0
				} else { cell.MarB = -1 }

				// Parse Borders: tcPr 직속 lnL/lnR/lnT/lnB (이 PPTX의 실제 구조)
				if m := regexp.MustCompile(`(?s)<a:lnL([^>]*)>(.*?)</a:lnL>`).FindStringSubmatch(tcPrContent); m != nil {
					cell.BdrL = extractBorderCssFromLn(m[1], m[2])
				}
				if m := regexp.MustCompile(`(?s)<a:lnR([^>]*)>(.*?)</a:lnR>`).FindStringSubmatch(tcPrContent); m != nil {
					cell.BdrR = extractBorderCssFromLn(m[1], m[2])
				}
				if m := regexp.MustCompile(`(?s)<a:lnT([^>]*)>(.*?)</a:lnT>`).FindStringSubmatch(tcPrContent); m != nil {
					cell.BdrT = extractBorderCssFromLn(m[1], m[2])
				}
				if m := regexp.MustCompile(`(?s)<a:lnB([^>]*)>(.*?)</a:lnB>`).FindStringSubmatch(tcPrContent); m != nil {
					cell.BdrB = extractBorderCssFromLn(m[1], m[2])
				}

				// Remove borders to safely find cell background fill
				tcPrContent = regexp.MustCompile(`(?s)<a:ln[A-Za-z]*[^>]*>.*?</a:ln[A-Za-z]*>`).ReplaceAllString(tcPrContent, "")
				
				srgbBgRe := regexp.MustCompile(`(?s)<a:solidFill>.*?<a:srgbClr val="([0-9A-Fa-f]{6})"`)
				schemeBgRe := regexp.MustCompile(`(?s)<a:solidFill>.*?<a:schemeClr val="([^"]+)"`)
				
				if bgM := srgbBgRe.FindStringSubmatch(tcPrContent); bgM != nil {
					cell.Bg = "#" + bgM[1]
				} else if bgM := schemeBgRe.FindStringSubmatch(tcPrContent); bgM != nil {
					cell.Bg = "scheme:" + bgM[1]
				}
			}
			rowCells = append(rowCells, cell)
		}

		data, _ := json.Marshal(rowCells)
		shp.TableData = append(shp.TableData, json.RawMessage(data))
	}

	shp.HasText = true
	return shp
}

func parseSpElement(spXML string, mediaMap map[string]string, mediaDir string, id *int) *ParsedShape {
	shp := &ParsedShape{ID: *id, Type: "AUTO"}
	*id++

	// Name
	if m := regexp.MustCompile(`name="([^"]+)"`).FindStringSubmatch(spXML); m != nil {
		shp.Name = m[1]
	}

	// Position/Size from xfrm
	parseXfrm(spXML, shp)

	// Fill color
	parseFill(spXML, shp)
	
	// Shadow
	parseShadow(spXML, shp)
	
	// Quarantined (Unhandled) Metadata
	parseQuarantinedMeta(spXML, shp)

	// Image via blipFill
	if rId := extractBlipRId(spXML); rId != "" {
		if mediaFile, ok := mediaMap[rId]; ok {
			absPath := filepath.Join(mediaDir, mediaFile)
			shp.ImagePath = filepath.ToSlash(absPath)
			shp.Type = "PICTURE"
		}
	}

	// Text
	parseTextBody(spXML, shp)

	// Placeholder type detection + 좌표 없으면 기본 좌표 할당
	phType := ""
	if strings.Contains(spXML, `type="title"`) {
		phType = "title"
	} else if strings.Contains(spXML, `type="ctrTitle"`) {
		phType = "ctrTitle"
	} else if strings.Contains(spXML, `type="body"`) {
		phType = "body"
	} else if strings.Contains(spXML, `type="subTitle"`) {
		phType = "subTitle"
	} else if strings.Contains(spXML, `type="sldNum"`) {
		phType = "sldNum"
	}
	if phType != "" {
		shp.Type = "PLACEHOLDER"
		shp.PhType = phType
		// 좌표/폰트 크기는 ParsePPTX에서 레이아웃 XML로부터 상속됨
	}

	return shp
}

func parsePicElement(picXML string, mediaMap map[string]string, mediaDir string, id *int) *ParsedShape {
	shp := &ParsedShape{ID: *id, Type: "PICTURE"}
	*id++

	if m := regexp.MustCompile(`name="([^"]+)"`).FindStringSubmatch(picXML); m != nil {
		shp.Name = m[1]
	}

	parseXfrm(picXML, shp)

	// [Quark 1] Parse srcRect for image cropping
	srcRectRe := regexp.MustCompile(`(?s)<a:srcRect([^>]+)/>`)
	if srcM := srcRectRe.FindStringSubmatch(picXML); srcM != nil {
		attrStr := srcM[1]
		if lM := regexp.MustCompile(`\bl="(\d+)"`).FindStringSubmatch(attrStr); lM != nil {
			v, _ := strconv.ParseFloat(lM[1], 64)
			shp.CropL = v / 100000.0
		}
		if tM := regexp.MustCompile(`\bt="(\d+)"`).FindStringSubmatch(attrStr); tM != nil {
			v, _ := strconv.ParseFloat(tM[1], 64)
			shp.CropT = v / 100000.0
		}
		if rM := regexp.MustCompile(`\br="(\d+)"`).FindStringSubmatch(attrStr); rM != nil {
			v, _ := strconv.ParseFloat(rM[1], 64)
			shp.CropR = v / 100000.0
		}
		if bM := regexp.MustCompile(`\bb="(\d+)"`).FindStringSubmatch(attrStr); bM != nil {
			v, _ := strconv.ParseFloat(bM[1], 64)
			shp.CropB = v / 100000.0
		}
	}

	if rId := extractBlipRId(picXML); rId != "" {
		if mediaFile, ok := mediaMap[rId]; ok {
			absPath := filepath.Join(mediaDir, mediaFile)
			shp.ImagePath = filepath.ToSlash(absPath)

			// [clrChange] Color-to-transparent image preprocessing
			clrChangeRe := regexp.MustCompile(`(?s)<a:clrChange>.*?<a:clrFrom>.*?val="([A-Fa-f0-9]{6})".*?</a:clrFrom>.*?<a:clrTo>.*?val="([A-Fa-f0-9]{6})".*?(?:alpha val="(\d+)")?.*?</a:clrTo>.*?</a:clrChange>`)
			if clrM := clrChangeRe.FindStringSubmatch(picXML); clrM != nil {
				fromHex := clrM[1]
				toAlphaStr := clrM[3]
				toAlpha := 0 // default: fully transparent
				if toAlphaStr != "" {
					if v, err := strconv.Atoi(toAlphaStr); err == nil {
						toAlpha = v * 255 / 100000 // PPT uses 0-100000 scale
					}
				}
				// Apply color change to image file
				newPath := applyClrChange(absPath, fromHex, uint8(toAlpha))
				if newPath != "" {
					shp.ImagePath = filepath.ToSlash(newPath)
				}
			}
		}
	}

	return shp
}


func parseGrpElement(grpXML string, mediaMap map[string]string, mediaDir string, id *int) *ParsedShape {
	shp := &ParsedShape{ID: *id, Type: "GROUP"}
	*id++

	if m := regexp.MustCompile(`name="([^"]+)"`).FindStringSubmatch(grpXML); m != nil {
		shp.Name = m[1]
	}

	parseXfrm(grpXML, shp)

	// 그룹 좌표계 추출 (chOff/chExt)
	grpSpPrRe := regexp.MustCompile(`(?s)<p:grpSpPr[^>]*>(.*?)</p:grpSpPr>`)
	grpSpPrM := grpSpPrRe.FindStringSubmatch(grpXML)
	
	// 그룹 world 좌표 (이미 parseXfrm에서 off/ext로 설정됨)
	worldOffX := shp.X
	worldOffY := shp.Y
	worldExtW := shp.W
	worldExtH := shp.H
	
	// child 좌표계
	chOffX, chOffY := 0.0, 0.0
	chExtW, chExtH := worldExtW, worldExtH
	if grpSpPrM != nil {
		chOffRe := regexp.MustCompile(`<a:chOff\s+x="(-?\d+)"\s+y="(-?\d+)"`)
		chExtRe := regexp.MustCompile(`<a:chExt\s+cx="(\d+)"\s+cy="(\d+)"`)
		if m := chOffRe.FindStringSubmatch(grpSpPrM[1]); m != nil {
			cx, _ := strconv.ParseFloat(m[1], 64)
			cy, _ := strconv.ParseFloat(m[2], 64)
			chOffX = cx / 12700.0 * 1.333333
			chOffY = cy / 12700.0 * 1.333333
		}
		if m := chExtRe.FindStringSubmatch(grpSpPrM[1]); m != nil {
			cw, _ := strconv.ParseFloat(m[1], 64)
			ch, _ := strconv.ParseFloat(m[2], 64)
			chExtW = cw / 12700.0 * 1.333333
			chExtH = ch / 12700.0 * 1.333333
		}
	}

	// 좌표 변환 함수: child → world
	scaleX := 1.0
	scaleY := 1.0
	if chExtW > 0 { scaleX = worldExtW / chExtW }
	if chExtH > 0 { scaleY = worldExtH / chExtH }
	
	transformChild := func(child *ParsedShape) {
		child.X = (child.X - chOffX) * scaleX + worldOffX
		child.Y = (child.Y - chOffY) * scaleY + worldOffY
		child.W = child.W * scaleX
		child.H = child.H * scaleY
	}

	// 자식 p:pic 추출
	picRe := regexp.MustCompile(`(?s)<p:pic>(.*?)</p:pic>`)
	picMatches := picRe.FindAllString(grpXML, -1)
	for _, picXML := range picMatches {
		childShp := parsePicElement(picXML, mediaMap, mediaDir, id)
		if childShp != nil && childShp.ImagePath != "" {
			transformChild(childShp)
			shp.Children = append(shp.Children, *childShp)
		}
	}

	// 자식 p:sp 추출
	spRe := regexp.MustCompile(`(?s)<p:sp>(.*?)</p:sp>`)
	spMatches := spRe.FindAllString(grpXML, -1)
	for _, spXML := range spMatches {
		childShp := parseSpElement(spXML, mediaMap, mediaDir, id)
		if childShp != nil {
			transformChild(childShp)
			shp.Children = append(shp.Children, *childShp)
		}
	}

	// Fallback: 자식이 없으면 기존 방식으로 텍스트/이미지 추출
	if len(shp.Children) == 0 {
		if rId := extractBlipRId(grpXML); rId != "" {
			if mediaFile, ok := mediaMap[rId]; ok {
				absPath := filepath.Join(mediaDir, mediaFile)
				shp.ImagePath = filepath.ToSlash(absPath)
			}
		}
		parseTextBody(grpXML, shp)
	}

	return shp
}

func parseXfrm(xmlStr string, shp *ParsedShape) {
	// Find <a:off x="..." y="..."/> and <a:ext cx="..." cy="..."/>
	offRe := regexp.MustCompile(`<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"`)
	extRe := regexp.MustCompile(`<a:ext\s+cx="(\d+)"\s+cy="(\d+)"`)

	if m := offRe.FindStringSubmatch(xmlStr); m != nil {
		x, _ := strconv.ParseInt(m[1], 10, 64)
		y, _ := strconv.ParseInt(m[2], 10, 64)
		shp.X = emuToPt(x)
		shp.Y = emuToPt(y)
	}
	if m := extRe.FindStringSubmatch(xmlStr); m != nil {
		cx, _ := strconv.ParseInt(m[1], 10, 64)
		cy, _ := strconv.ParseInt(m[2], 10, 64)
		shp.W = emuToPt(cx)
		shp.H = emuToPt(cy)
	}

	// Rotation: rot="5400000" → 90 degrees (units: 60000ths of a degree)
	rotRe := regexp.MustCompile(`<a:xfrm[^>]*\brot="(-?\d+)"`)
	if m := rotRe.FindStringSubmatch(xmlStr); m != nil {
		rot, _ := strconv.ParseInt(m[1], 10, 64)
		shp.Rotation = float64(rot) / 60000.0
	}
}

func parseFill(xmlStr string, shp *ParsedShape) {
	// Only look at spPr block for shape fill (not text run fills)
	spPrRe := regexp.MustCompile(`(?s)<p:spPr[^>]*>(.*?)</p:spPr>`)
	spPrMatch := spPrRe.FindStringSubmatch(xmlStr)
	if spPrMatch == nil {
		// Try matching empty spPr like <p:spPr/>
		if regexp.MustCompile(`<p:spPr\b[^>]*/>`).MatchString(xmlStr) {
			shp.FillType = "none"
		}
		return
	}
	spPrBlock := spPrMatch[1]

	// Remove line (border) blocks so their internal noFill/solidFill doesn't interfere
	borderlessSpPr := regexp.MustCompile(`(?s)<a:ln.*?</a:ln>`).ReplaceAllString(spPrBlock, "")

	// Check solidFill first
	srgbRe := regexp.MustCompile(`(?s)<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"`)
	schemeRe := regexp.MustCompile(`(?s)<a:solidFill>.*?<a:schemeClr val="([^"]+)"`)
	
	if m := srgbRe.FindStringSubmatch(borderlessSpPr); m != nil {
		shp.FillType = "solid"
		shp.FillColor = "#" + m[1]
		// Extract alpha (transparency)
		alphaRe := regexp.MustCompile(`<a:alpha val="(\d+)"`)
		if am := alphaRe.FindStringSubmatch(borderlessSpPr); am != nil {
			alphaVal, _ := strconv.ParseFloat(am[1], 64)
			shp.FillTransparency = 1.0 - (alphaVal / 100000.0)
		}
	} else if m := schemeRe.FindStringSubmatch(borderlessSpPr); m != nil {
		shp.FillType = "solid"
		// lumMod/lumOff 변형자 감지 → 직접 색상 계산
		fillBlock := borderlessSpPr[strings.Index(borderlessSpPr, "<a:solidFill>"):]
		if endIdx := strings.Index(fillBlock, "</a:solidFill>"); endIdx > 0 {
			fillBlock = fillBlock[:endIdx+len("</a:solidFill>")]
		}
		lumModRe := regexp.MustCompile(`<a:lumMod val="(\d+)"`)
		lumOffRe := regexp.MustCompile(`<a:lumOff val="(\d+)"`)
		hasLumMod := lumModRe.FindStringSubmatch(fillBlock)
		hasLumOff := lumOffRe.FindStringSubmatch(fillBlock)
		if hasLumMod != nil || hasLumOff != nil {
			// scheme 색상을 일단 "scheme:XX"로 저장하되, lumMod/Off 정보를 함께 전달
			lumMod := 100000.0
			lumOff := 0.0
			if hasLumMod != nil {
				lumMod, _ = strconv.ParseFloat(hasLumMod[1], 64)
			}
			if hasLumOff != nil {
				lumOff, _ = strconv.ParseFloat(hasLumOff[1], 64)
			}
			// lumMod/lumOff를 직접 적용하여 hex 색상 생성
			shp.FillColor = fmt.Sprintf("scheme:%s:lum:%.0f:%.0f", m[1], lumMod, lumOff)
		} else {
			shp.FillColor = "scheme:" + m[1]
		}
	} else if strings.Contains(borderlessSpPr, "<a:noFill") {
		shp.FillType = "none"
	} else {
		// spPr에 명시적 fill 없음 → p:style의 fillRef를 폴백으로 확인
		fillRefResolved := false
		styleRe := regexp.MustCompile(`(?s)<p:style[^>]*>(.*?)</p:style>`)
		if sm := styleRe.FindStringSubmatch(xmlStr); sm != nil {
			fillRefRe := regexp.MustCompile(`(?s)<a:fillRef idx="(\d+)"[^>]*>.*?<a:schemeClr val="([^"]+)"`)
			if fm := fillRefRe.FindStringSubmatch(sm[1]); fm != nil {
				idx, _ := strconv.Atoi(fm[1])
				if idx > 0 {
					shp.FillType = "solid"
					shp.FillColor = "scheme:" + fm[2]
					fillRefResolved = true
				}
			}
		}
		if !fillRefResolved {
			shp.FillType = "none" // default to none if no fill is explicitly set
		}
	}

		// Gradient fill detection — extract all stops + angle
		if strings.Contains(spPrBlock, "<a:gradFill") {
			gsBlockRe := regexp.MustCompile(`(?s)<a:gs\s+pos="(\d+)"[^>]*>(.*?)</a:gs>`)
			colorRe := regexp.MustCompile(`<a:srgbClr\s+val="([0-9A-Fa-f]{6})"`)
			alphaRe := regexp.MustCompile(`<a:alpha\s+val="(\d+)"`)
			gsBlocks := gsBlockRe.FindAllStringSubmatch(spPrBlock, -1)
			if len(gsBlocks) >= 2 {
				shp.FillType = "gradient"
				shp.FillColor = ""
				for _, gsBlock := range gsBlocks {
					pos, _ := strconv.Atoi(gsBlock[1])
					inner := gsBlock[2]
					colorM := colorRe.FindStringSubmatch(inner)
					color := ""
					if colorM != nil {
						color = "#" + colorM[1]
					} else {
						// Fallback: try schemeClr
						schemeRe := regexp.MustCompile(`<a:schemeClr\s+val="([^"]+)"`)
						if sm := schemeRe.FindStringSubmatch(inner); sm != nil {
							color = "scheme:" + sm[1]
						}
					}
					if color == "" {
						continue
					}
					stop := GradientStop{
						Color:    color,
						Position: pos,
						Alpha:    1.0,
					}
					if am := alphaRe.FindStringSubmatch(inner); am != nil {
						alphaVal, _ := strconv.ParseFloat(am[1], 64)
						stop.Alpha = alphaVal / 100000.0
					}
					shp.GradientStops = append(shp.GradientStops, stop)
				}
				if len(shp.GradientStops) > 0 {
					shp.FillColor = shp.GradientStops[0].Color
				}
				angRe := regexp.MustCompile(`<a:lin ang="(\d+)"`)
				if am := angRe.FindStringSubmatch(spPrBlock); am != nil {
					ang, _ := strconv.ParseFloat(am[1], 64)
					shp.GradientAngle = float64(int(90-ang/60000+360) % 360)
				}
			}
		}

	// ShapeType (prstGeom prst="roundRect" / "ellipse" etc.)
	geomRe := regexp.MustCompile(`<a:prstGeom prst="([^"]+)"`)
	if m := geomRe.FindStringSubmatch(spPrBlock); m != nil {
		shp.ShapeType = m[1]
	}

	// Border radius (roundRect adj value → 0.0~1.0)
	if shp.ShapeType == "roundRect" {
		adjRe := regexp.MustCompile(`fmla="val (\d+)"`)
		if m := adjRe.FindStringSubmatch(spPrBlock); m != nil {
			adj, _ := strconv.ParseFloat(m[1], 64)
			shp.BorderRadius = adj / 100000.0
		} else {
			shp.BorderRadius = 0.167
		}
	}

	// Border (a:ln) — color and width
	lnSrgbRe := regexp.MustCompile(`<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"`)
	lnSchemeRe := regexp.MustCompile(`<a:solidFill>\s*<a:schemeClr val="([^"]+)"`)
	lnRe := regexp.MustCompile(`<a:ln\s+w="(\d+)"[^>]*>`)
	if m := lnRe.FindStringSubmatch(spPrBlock); m != nil {
		w, _ := strconv.ParseInt(m[1], 10, 64)
		shp.BorderWidth = float64(w) / 12700.0
		lnBlock := spPrBlock[strings.Index(spPrBlock, "<a:ln"):]
		if endIdx := strings.Index(lnBlock, "</a:ln>"); endIdx > 0 {
			lnBlock = lnBlock[:endIdx]
		}
		if cm := lnSrgbRe.FindStringSubmatch(lnBlock); cm != nil {
			shp.BorderColor = "#" + cm[1]
		} else if cm := lnSchemeRe.FindStringSubmatch(lnBlock); cm != nil {
			shp.BorderColor = "scheme:" + cm[1]
		}
	}
}

func extractBlipRId(xml string) string {
	re := regexp.MustCompile(`r:embed="(rId\d+)"`)
	if m := re.FindStringSubmatch(xml); m != nil {
		return m[1]
	}
	return ""
}

func parseTextBody(xmlStr string, shp *ParsedShape) {
	// 텍스트 상자의 줄바꿈 설정(Wrap), 수직 정렬(anchor), 안쪽 여백(lIns/tIns/rIns/bIns) 확인
	bodyPrRe := regexp.MustCompile(`(?s)<a:bodyPr([^>]*)>`)
	if m := bodyPrRe.FindStringSubmatch(xmlStr); m != nil {
		attrStr := m[1]
		if wrapM := regexp.MustCompile(`\bwrap="([^"]+)"`).FindStringSubmatch(attrStr); wrapM != nil {
			shp.TextWrap = wrapM[1]
		}
		if anchorM := regexp.MustCompile(`\banchor="([^"]+)"`).FindStringSubmatch(attrStr); anchorM != nil {
			shp.Valign = anchorM[1] // "t", "ctr", "b"
		} else { shp.Valign = "t" } // default top

		// Parse padding (EMU to px)
		if lM := regexp.MustCompile(`\blIns="(\d+)"`).FindStringSubmatch(attrStr); lM != nil {
			v, _ := strconv.ParseFloat(lM[1], 64)
			shp.PadL = v / 9525.0
		} else { shp.PadL = -1 }
		if rM := regexp.MustCompile(`\brIns="(\d+)"`).FindStringSubmatch(attrStr); rM != nil {
			v, _ := strconv.ParseFloat(rM[1], 64)
			shp.PadR = v / 9525.0
		} else { shp.PadR = -1 }
		if tM := regexp.MustCompile(`\btIns="(\d+)"`).FindStringSubmatch(attrStr); tM != nil {
			v, _ := strconv.ParseFloat(tM[1], 64)
			shp.PadT = v / 9525.0
		} else { shp.PadT = -1 }
		if bM := regexp.MustCompile(`\bbIns="(\d+)"`).FindStringSubmatch(attrStr); bM != nil {
			v, _ := strconv.ParseFloat(bM[1], 64)
			shp.PadB = v / 9525.0
		} else { shp.PadB = -1 }
	}

	// Paragraph 경계를 정확히 보존하기 위해 <a:p> 단위로 순회
	// (?s) for multi-line XML matching
	paraRe := regexp.MustCompile(`(?s)<a:p\b[^>]*>(.*?)</a:p>`)
	runRe := regexp.MustCompile(`(?s)(?:<a:r>|<a:fld\b[^>]*>)(.*?)(?:</a:r>|</a:fld>)`)
	textRe := regexp.MustCompile(`(?s)<a:t>(.*?)</a:t>`)
	fontRe := regexp.MustCompile(`typeface="([^"]+)"`)
	sizeRe := regexp.MustCompile(`sz="(\d+)"`)
	boldRe := regexp.MustCompile(`\bb="1"`)
	italicRe := regexp.MustCompile(`\bi="1"`)
	uRe := regexp.MustCompile(`\bu="sng"`)
	colorRe := regexp.MustCompile(`(?s)<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"`)
	schemeRe := regexp.MustCompile(`(?s)<a:solidFill>.*?<a:schemeClr val="([^"]+)"`)
	alignRe := regexp.MustCompile(`algn="(l|ctr|r)"`)

	// Extract shape-level fontRef color
	spFontRefColor := ""
	fontRefRe := regexp.MustCompile(`(?s)<p:style>.*?<a:fontRef\b[^>]*>.*?<a:schemeClr val="([^"]+)".*?</a:fontRef>`)
	if m := fontRefRe.FindStringSubmatch(xmlStr); m != nil {
		spFontRefColor = "scheme:" + m[1]
	}

	paras := paraRe.FindAllStringSubmatch(xmlStr, -1)
	for pi, pm := range paras {
		paraContent := pm[1]

		// Extract paragraph-level default color from <a:defRPr>
		defRPrColor := ""
		defRPrRe := regexp.MustCompile(`(?s)<a:defRPr[^>]*>(.*?)</a:defRPr>`)
		if m := defRPrRe.FindStringSubmatch(paraContent); m != nil {
			defRPrBlock := m[1]
			if c := colorRe.FindStringSubmatch(defRPrBlock); c != nil {
				defRPrColor = "#" + c[1]
			} else if c := schemeRe.FindStringSubmatch(defRPrBlock); c != nil {
				schemeColor := c[1]
				lumModRe := regexp.MustCompile(`<a:lumMod val="(\d+)"`)
				lumOffRe := regexp.MustCompile(`<a:lumOff val="(\d+)"`)
				hasLumMod := lumModRe.FindStringSubmatch(defRPrBlock)
				hasLumOff := lumOffRe.FindStringSubmatch(defRPrBlock)
				if hasLumMod != nil || hasLumOff != nil {
					lumMod := 100000.0
					lumOff := 0.0
					if hasLumMod != nil { lumMod, _ = strconv.ParseFloat(hasLumMod[1], 64) }
					if hasLumOff != nil { lumOff, _ = strconv.ParseFloat(hasLumOff[1], 64) }
					defRPrColor = fmt.Sprintf("scheme:%s:lum:%.0f:%.0f", schemeColor, lumMod, lumOff)
				} else {
					defRPrColor = "scheme:" + schemeColor
				}
			}
		}

		// Paragraph 정렬 및 여백 파싱
		align := 1 // default left
		if m := alignRe.FindStringSubmatch(paraContent); m != nil {
			switch m[1] {
			case "ctr": align = 2
			case "r": align = 3
			}
		}
		
		lineSpace := 1.0
		lnSpcRe := regexp.MustCompile(`<a:lnSpc>\s*<a:spcPct val="(\d+)"`)
		if m := lnSpcRe.FindStringSubmatch(paraContent); m != nil {
			pct, _ := strconv.Atoi(m[1])
			lineSpace = float64(pct) / 100000.0 // 100000 = 100%
		}

		spaceAft := 0.0
		spcAftRe := regexp.MustCompile(`<a:spcAft>\s*<a:spcPts val="(\d+)"`)
		if m := spcAftRe.FindStringSubmatch(paraContent); m != nil {
			pts, _ := strconv.Atoi(m[1])
			spaceAft = float64(pts) / 100.0 // 100 = 1pt
		}

		// Bullet detection
		hasBullet := false
		buCharShapeRe := regexp.MustCompile(`<a:buChar\s+char="([^"]+)"`)
		buNoneShapeRe := regexp.MustCompile(`<a:buNone`)
		if buCharShapeRe.MatchString(paraContent) && !buNoneShapeRe.MatchString(paraContent) {
			hasBullet = true
		}

		runs := runRe.FindAllString(paraContent, -1)
		for ri, runMatch := range runs {
			run := TextRun{Align: align, LineSpace: lineSpace, SpaceAft: spaceAft}
			if hasBullet && ri == 0 {
				run.Bullet = 1
			}

			if m := textRe.FindStringSubmatch(runMatch); m != nil {
				run.Text = m[1]
			}
			// Decode XML entities
			run.Text = strings.ReplaceAll(run.Text, "&amp;", "&")
			run.Text = strings.ReplaceAll(run.Text, "&lt;", "<")
			run.Text = strings.ReplaceAll(run.Text, "&gt;", ">")
			if run.Text == "" {
				continue
			}

			if m := fontRe.FindStringSubmatch(runMatch); m != nil {
				run.Font = m[1]
			}
			if run.Font == "" || run.Font == "+mj-lt" || run.Font == "+mn-lt" {
				run.Font = "Arial"
			}

			if m := sizeRe.FindStringSubmatch(runMatch); m != nil {
				sz, _ := strconv.Atoi(m[1])
				run.Size = float64(sz) / 100.0
			}
			if run.Size == 0 {
				run.Size = 11
			}

			if boldRe.MatchString(runMatch) {
				run.Bold = -1
			}
			if italicRe.MatchString(runMatch) {
				run.Italic = -1
			}
			if uRe.MatchString(runMatch) {
				run.Underline = -1
			}

			if m := colorRe.FindStringSubmatch(runMatch); m != nil {
				run.Color = "#" + m[1]
			} else if m := schemeRe.FindStringSubmatch(runMatch); m != nil {
				schemeColor := m[1]
				lumModRe := regexp.MustCompile(`<a:lumMod val="(\d+)"`)
				lumOffRe := regexp.MustCompile(`<a:lumOff val="(\d+)"`)
				hasLumMod := lumModRe.FindStringSubmatch(runMatch)
				hasLumOff := lumOffRe.FindStringSubmatch(runMatch)
				
				if hasLumMod != nil || hasLumOff != nil {
					lumMod := 100000.0
					lumOff := 0.0
					if hasLumMod != nil {
						lumMod, _ = strconv.ParseFloat(hasLumMod[1], 64)
					}
					if hasLumOff != nil {
						lumOff, _ = strconv.ParseFloat(hasLumOff[1], 64)
					}
					run.Color = fmt.Sprintf("scheme:%s:lum:%.0f:%.0f", schemeColor, lumMod, lumOff)
				} else {
					run.Color = "scheme:" + schemeColor
				}
			}

			// Fallback colors if run has no explicit color
			if run.Color == "" {
				if defRPrColor != "" {
					run.Color = defRPrColor
				} else if spFontRefColor != "" {
					run.Color = spFontRefColor
				}
			}

			// Paragraph 경계: 마지막 paragraph가 아니고, 이 run이 paragraph의 마지막이면 \n 추가
			if pi < len(paras)-1 && ri == len(runs)-1 {
				run.Text += "\n"
			}

			shp.TextRuns = append(shp.TextRuns, run)
			shp.HasText = true
		}
	}
}

func extractBgColor(xmlData []byte) string {
	content := string(xmlData)
	// 1. <p:bg> 안의 solidFill — srgbClr
	bgRe := regexp.MustCompile(`(?s)<p:bg>.*?</p:bg>`)
	colorRe := regexp.MustCompile(`<a:srgbClr val="([0-9A-Fa-f]{6})"`)
	if m := bgRe.FindString(content); m != "" {
		if c := colorRe.FindStringSubmatch(m); c != nil {
			return "#" + c[1]
		}
		// schemeClr in bg
		schemeRe := regexp.MustCompile(`<a:schemeClr val="([^"]+)"`)
		if c := schemeRe.FindStringSubmatch(m); c != nil {
			return "scheme:" + c[1]
		}
	}
	return ""
}

func parseShadow(spXML string, shp *ParsedShape) {
	shp.HasShadow = false
	// effectLst 확인
	effectRe := regexp.MustCompile(`(?s)<a:effectLst([^>]*)>(.*?)</a:effectLst>`)
	if effectM := effectRe.FindStringSubmatch(spXML); effectM != nil {
		effectContent := effectM[2]
		shdwBlockRe := regexp.MustCompile(`(?s)<a:outerShdw([^>]*)>(.*?)</a:outerShdw>`)
		if m := shdwBlockRe.FindStringSubmatch(effectContent); m != nil {
			shp.HasShadow = true
			shp.ShadowColor = "rgba(0,0,0,0.3)" // fallback
			colorRe := regexp.MustCompile(`<a:srgbClr val="([0-9A-Fa-f]{6})"`)
			if cm := colorRe.FindStringSubmatch(m[2]); cm != nil {
				shp.ShadowColor = "#" + cm[1]
			}
		} else if strings.Contains(effectContent, "<a:prstShdw") {
			// Preset shadow fallback
			shp.HasShadow = true
			shp.ShadowColor = "rgba(0,0,0,0.2)"
		}
	}
}

func parseQuarantinedMeta(spXML string, shp *ParsedShape) {
	if shp.QuarantinedMeta == nil {
		shp.QuarantinedMeta = make(map[string]string)
	}
	
	// 1. 3D Scene / 3D Shape Properties
	if m := regexp.MustCompile(`(?s)<a:scene3d([^>]*)>(.*?)</a:scene3d>`).FindString(spXML); m != "" {
		shp.QuarantinedMeta["scene3d"] = m
	}
	if m := regexp.MustCompile(`(?s)<a:sp3d([^>]*)>(.*?)</a:sp3d>`).FindString(spXML); m != "" {
		shp.QuarantinedMeta["sp3d"] = m
	}
	
	// 2. Unhandled Effect Lists (Glow, Reflection, SoftEdge 등)
	if m := regexp.MustCompile(`(?s)<a:effectLst([^>]*)>(.*?)</a:effectLst>`).FindStringSubmatch(spXML); m != nil {
		if !strings.Contains(m[2], "outerShdw") && !strings.Contains(m[2], "prstShdw") && m[2] != "" {
			shp.QuarantinedMeta["effectLst_raw"] = m[0]
		}
	}
	
	// 3. WordArt Warp
	if m := regexp.MustCompile(`(?s)<a:prstTxWarp([^>]*)>(.*?)</a:prstTxWarp>`).FindString(spXML); m != "" {
		shp.QuarantinedMeta["prstTxWarp"] = m
	}
	
	// 4. Animation Effects (p:animEffect)
	if m := regexp.MustCompile(`(?s)<p:animEffect([^>]*)>(.*?)</p:animEffect>`).FindString(spXML); m != "" {
		shp.QuarantinedMeta["animEffect"] = m
	}
}

func extractBorderCssFromLn(attrStr string, bodyStr string) string {
	// noFill이면 테두리 없음
	if strings.Contains(bodyStr, "<a:noFill/>") { return "" }
	
	w := 1.0
	color := "#000000"
	
	// w 속성은 태그 어트리뷰트에 있음: <a:lnL w="6350" ...>
	wRe := regexp.MustCompile(`w="(\d+)"`)
	if m := wRe.FindStringSubmatch(attrStr); m != nil {
		v, _ := strconv.ParseFloat(m[1], 64)
		w = v / 12700.0 // EMU to pt
		if w < 0.5 { w = 0.5 }
	}
	
	// 색상은 바디 콘텐츠에 있음
	colorRe := regexp.MustCompile(`<a:srgbClr val="([0-9A-Fa-f]{6})"`)
	schemeRe := regexp.MustCompile(`<a:schemeClr val="([^"]+)"`)
	
	if m := colorRe.FindStringSubmatch(bodyStr); m != nil {
		color = "#" + m[1]
	} else if m := schemeRe.FindStringSubmatch(bodyStr); m != nil {
		color = "scheme:" + m[1]
	}
	return fmt.Sprintf("%.1fpt solid %s", w, color)
}

func extractBorderCss(lnBlock string) string {
	if lnBlock == "" { return "" }
	w := 1.0
	color := "#000000"
	lnRe := regexp.MustCompile(`<a:ln\b[^>]*w="(\d+)"`)
	if m := lnRe.FindStringSubmatch(lnBlock); m != nil {
		v, _ := strconv.ParseFloat(m[1], 64)
		w = v / 12700.0 // emu to pt
	}
	if strings.Contains(lnBlock, "<a:noFill/>") { return "" }
	
	colorRe := regexp.MustCompile(`<a:srgbClr val="([0-9A-Fa-f]{6})"`)
	schemeRe := regexp.MustCompile(`<a:schemeClr val="([^"]+)"`)
	
	if m := colorRe.FindStringSubmatch(lnBlock); m != nil {
		color = "#" + m[1]
	} else if m := schemeRe.FindStringSubmatch(lnBlock); m != nil {
		color = "scheme:" + m[1]
	}
	return fmt.Sprintf("%.1fpt solid %s", w, color)
}

// parseThemeColors extracts the color scheme from theme XML
// applyLumModOff applies OOXML lumMod/lumOff color transforms to a hex color
// lumMod: brightness multiplier (0.0~1.0), lumOff: brightness offset (0.0~1.0)
// e.g., bg1(#FFFFFF) + lumMod=0.95 → #F2F2F2
func applyLumModOff(hexColor string, lumMod float64, lumOff float64) string {
	if len(hexColor) != 7 || hexColor[0] != '#' {
		return hexColor
	}
	r, _ := strconv.ParseInt(hexColor[1:3], 16, 64)
	g, _ := strconv.ParseInt(hexColor[3:5], 16, 64)
	b, _ := strconv.ParseInt(hexColor[5:7], 16, 64)
	
	// Apply: newVal = val * lumMod + 255 * lumOff
	clamp := func(v float64) int {
		if v < 0 { return 0 }
		if v > 255 { return 255 }
		return int(v + 0.5)
	}
	nr := clamp(float64(r)*lumMod + 255.0*lumOff)
	ng := clamp(float64(g)*lumMod + 255.0*lumOff)
	nb := clamp(float64(b)*lumMod + 255.0*lumOff)
	
	return fmt.Sprintf("#%02X%02X%02X", nr, ng, nb)
}

func parseThemeColors(fileMap map[string]*zip.File) map[string]string {
	colors := map[string]string{
		"bg1": "#FFFFFF", "bg2": "#EEEEEE",
		"tx1": "#000000", "tx2": "#666666",
		"dk1": "#000000", "dk2": "#333333",
		"lt1": "#FFFFFF", "lt2": "#EEEEEE",
	}
	
	// Collect theme files and sort them to ensure theme1.xml is picked first
	var themeFiles []string
	for name := range fileMap {
		if strings.HasPrefix(name, "ppt/theme/theme") && strings.HasSuffix(name, ".xml") {
			themeFiles = append(themeFiles, name)
		}
	}
	sort.Strings(themeFiles)
	
	// Find theme file
	for _, name := range themeFiles {
		f := fileMap[name]
		data, err := readZipFile(f)
		if err != nil {
			continue
		}
		xml := string(data)
		// Extract clrScheme block
		schemeRe := regexp.MustCompile(`(?s)<a:clrScheme[^>]*>(.*?)</a:clrScheme>`)
		if m := schemeRe.FindStringSubmatch(xml); m != nil {
			schemeBlock := m[1]
			// Parse each color element: <a:dk1><a:srgbClr val="333C47"/></a:dk1>
			colorNames := []string{"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"}
			for _, cn := range colorNames {
				re := regexp.MustCompile(`(?s)<a:` + cn + `>.*?<a:srgbClr val="([0-9A-Fa-f]{6})".*?</a:` + cn + `>`)
				if cm := re.FindStringSubmatch(schemeBlock); cm != nil {
					colors[cn] = "#" + cm[1]
				}
			}
			// Map bg1→lt1, bg2→lt2, tx1→dk1, tx2→dk2
			colors["bg1"] = colors["lt1"]
			colors["bg2"] = colors["lt2"]
			colors["tx1"] = colors["dk1"]
			colors["tx2"] = colors["dk2"]
		}
		break
	}
	return colors
}

// resolveSchemeColors replaces "scheme:xxx" references in XML data with actual colors
// This is a pre-processing step — we don't modify the XML, but we'll resolve during fill parsing
func resolveSchemeColors(xmlData []byte, colors map[string]string) {
	// This is handled at shape level in the parseFill function
	// Theme colors are passed through the pipeline via the schemeClr prefix
	_ = xmlData
	_ = colors
}

func parseSlideDimensions(presData []byte) (float64, float64) {
	// <p:sldSz cx="12192000" cy="6858000"/>
	re := regexp.MustCompile(`<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"`)
	if m := re.FindStringSubmatch(string(presData)); m != nil {
		cx, _ := strconv.ParseInt(m[1], 10, 64)
		cy, _ := strconv.ParseInt(m[2], 10, 64)
		return emuToPt(cx), emuToPt(cy)
	}
	return 0, 0
}

func readZipFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

func extractFile(f *zip.File, outPath string) error {
	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()

	out, err := os.Create(outPath)
	if err != nil {
		return err
	}
	defer out.Close()

	_, err = io.Copy(out, rc)
	return err
}
