package main

import (
	"archive/zip"
	"flag"
	"fmt"
	"io"
	"log"
	"math"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"

	"golang.org/x/net/html"
	"golang.org/x/net/html/atom"

	"github.com/rhino-acoustic/pptx-engine/internal/openxml"
	"github.com/rhino-acoustic/pptx-engine/internal/compiler"
	"github.com/rhino-acoustic/pptx-engine/internal/mapper"
)

// ── config ──
const (
	slideW    = 13.333 // inches (widescreen 16:9)
	slideH    = 7.5
	marginIn  = 0.5
	contentW  = slideW - 2*marginIn
	maxSlideH = slideH - 2*marginIn // usable height per slide
)

func main() {
	htmlFile := flag.String("html", "", "Input HTML file (any HTML)")
	stubFile := flag.String("stub", "", "Stub PPTX for theme/layout")
	outFile := flag.String("out", "translated.pptx", "Output PPTX")
	flag.Parse()

	if *htmlFile == "" || *stubFile == "" {
		log.Fatal("Usage: anyhtml2pptx -html <file.html> -stub <template.pptx> -out <output.pptx>")
	}

	f, err := os.Open(*htmlFile)
	if err != nil {
		log.Fatalf("Cannot open HTML: %v", err)
	}
	defer f.Close()

	doc, err := html.Parse(f)
	if err != nil {
		log.Fatalf("HTML parse error: %v", err)
	}

	// Extract <style> blocks for class→style resolution
	styleMap := extractStyleBlocks(doc)

	// Walk DOM, collect visual blocks
	var blocks []visualBlock
	walkDOM(doc, styleMap, &blocks, 0)
	log.Printf("Extracted %d visual blocks from HTML", len(blocks))

	// Paginate blocks into slides
	slides := paginate(blocks)
	log.Printf("Paginated into %d slides", len(slides))

	// Build PPTX
	buildPPTX(*stubFile, *outFile, slides, *htmlFile)
	log.Printf("✅ DONE: %d slides → %s", len(slides), *outFile)
}

// ── types ──

type visualBlock struct {
	Type     string // "text", "image", "heading", "table", "shape", "hr"
	Text     string
	Runs     []textRun
	Level    int     // heading level (1-6) or indent level
	FontSize float64 // pt
	Font     string  // font-family
	Bold     bool
	Italic   bool
	Color    string  // hex
	BgColor  string
	Align    string
	ImgSrc   string
	ImgW     float64 // inches (estimated)
	ImgH     float64
	TableRows [][]string
	TableColW []float64
}

type textRun struct {
	Text      string
	Bold      bool
	Italic    bool
	Underline bool
	Color     string
	FontSize  float64
	Font      string
	Link      string
}

// ── DOM walker ──

func walkDOM(n *html.Node, styles map[string]string, blocks *[]visualBlock, depth int) {
	if n.Type == html.ElementNode {
		tag := n.DataAtom

		// Skip invisible
		style := getInlineStyle(n, styles)
		if strings.Contains(style, "display:none") || strings.Contains(style, "visibility:hidden") {
			return
		}

		switch tag {
		case atom.H1, atom.H2, atom.H3, atom.H4, atom.H5, atom.H6:
			level := int(n.Data[1] - '0')
			text := collectText(n)
			if text != "" {
				*blocks = append(*blocks, visualBlock{
					Type: "heading", Text: text, Level: level,
					FontSize: headingSize(level),
					Font: extractFontFamily(style),
					Bold: true,
					Color: extractColor(style),
					BgColor: extractBgColor(style),
				})
			}
			return

		case atom.P, atom.Div, atom.Span, atom.Section, atom.Article, atom.Main, atom.Aside, atom.Header, atom.Footer, atom.Nav:
			// Check if it's a simple text container or has children
			if isLeafText(n) {
				runs := collectRuns(n, styles)
				text := collectText(n)
				if strings.TrimSpace(text) != "" {
					blk := visualBlock{
						Type: "text", Text: text, Runs: runs,
						FontSize: extractFontSize(style),
						Font:     extractFontFamily(style),
						Bold:     strings.Contains(style, "font-weight:bold") || strings.Contains(style, "font-weight:700"),
						Italic:   strings.Contains(style, "font-style:italic"),
						Color:    extractColor(style),
						BgColor:  extractBgColor(style),
						Align:    extractAlign(style),
					}
					*blocks = append(*blocks, blk)
				}
				return
			}
			// else: recurse into children

		case atom.Img:
			src := getAttr(n, "src")
			if src != "" {
				*blocks = append(*blocks, visualBlock{
					Type: "image", ImgSrc: src,
					ImgW: 4.0, ImgH: 3.0, // default, will try to read actual
				})
			}
			return

		case atom.Table:
			tbl := parseHTMLTable(n)
			if len(tbl.TableRows) > 0 {
				*blocks = append(*blocks, tbl)
			}
			return

		case atom.Hr:
			*blocks = append(*blocks, visualBlock{Type: "hr"})
			return

		case atom.Ul, atom.Ol:
			items := collectListItems(n, tag == atom.Ol, styles)
			for _, item := range items {
				*blocks = append(*blocks, item)
			}
			return

		case atom.Br:
			*blocks = append(*blocks, visualBlock{Type: "text", Text: "\n", FontSize: 10})
			return

		case atom.Script, atom.Style, atom.Link, atom.Meta, atom.Head:
			return
		}
	}

	for c := n.FirstChild; c != nil; c = c.NextSibling {
		walkDOM(c, styles, blocks, depth+1)
	}
}

func isLeafText(n *html.Node) bool {
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if c.Type == html.ElementNode {
			tag := c.DataAtom
			// Inline elements are part of text
			if tag == atom.Span || tag == atom.Strong || tag == atom.B ||
				tag == atom.Em || tag == atom.I || tag == atom.U ||
				tag == atom.A || tag == atom.Code || tag == atom.Small ||
				tag == atom.Sub || tag == atom.Sup || tag == atom.Br {
				continue
			}
			return false // has block child
		}
	}
	return true
}

func collectText(n *html.Node) string {
	if n.Type == html.TextNode {
		return n.Data
	}
	var sb strings.Builder
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if c.Type == html.ElementNode && c.DataAtom == atom.Br {
			sb.WriteString("\n")
			continue
		}
		sb.WriteString(collectText(c))
	}
	return sb.String()
}

func collectRuns(n *html.Node, styles map[string]string) []textRun {
	var runs []textRun
	collectRunsInner(n, styles, &runs, false, false, false, "", 0, "")
	return runs
}

func collectRunsInner(n *html.Node, styles map[string]string, runs *[]textRun, bold, italic, underline bool, color string, fontSize float64, font string) {
	if n.Type == html.TextNode {
		text := n.Data
		if strings.TrimSpace(text) != "" || text == " " {
			*runs = append(*runs, textRun{
				Text: text, Bold: bold, Italic: italic, Underline: underline,
				Color: color, FontSize: fontSize, Font: font,
			})
		}
		return
	}
	if n.Type == html.ElementNode {
		style := getInlineStyle(n, styles)
		b, i, u := bold, italic, underline
		c, fs, ff := color, fontSize, font

		switch n.DataAtom {
		case atom.Strong, atom.B:
			b = true
		case atom.Em, atom.I:
			i = true
		case atom.U:
			u = true
		case atom.Br:
			*runs = append(*runs, textRun{Text: "\n"})
			return
		}

		if strings.Contains(style, "font-weight:bold") || strings.Contains(style, "font-weight:700") {
			b = true
		}
		if strings.Contains(style, "font-style:italic") {
			i = true
		}
		if ec := extractColor(style); ec != "" {
			c = ec
		}
		if ef := extractFontSize(style); ef > 0 {
			fs = ef
		}

		for child := n.FirstChild; child != nil; child = child.NextSibling {
			collectRunsInner(child, styles, runs, b, i, u, c, fs, ff)
		}
	}
}

func collectListItems(n *html.Node, ordered bool, styles map[string]string) []visualBlock {
	var items []visualBlock
	idx := 1
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if c.Type == html.ElementNode && c.DataAtom == atom.Li {
			text := collectText(c)
			prefix := "• "
			if ordered {
				prefix = fmt.Sprintf("%d. ", idx)
				idx++
			}
			items = append(items, visualBlock{
				Type: "text", Text: prefix + strings.TrimSpace(text),
				FontSize: 12, Level: 1,
			})
		}
	}
	return items
}

func parseHTMLTable(n *html.Node) visualBlock {
	blk := visualBlock{Type: "table"}
	// find all tr
	var walk func(*html.Node)
	walk = func(node *html.Node) {
		if node.Type == html.ElementNode && node.DataAtom == atom.Tr {
			var row []string
			for td := node.FirstChild; td != nil; td = td.NextSibling {
				if td.Type == html.ElementNode && (td.DataAtom == atom.Td || td.DataAtom == atom.Th) {
					row = append(row, strings.TrimSpace(collectText(td)))
				}
			}
			if len(row) > 0 {
				blk.TableRows = append(blk.TableRows, row)
			}
			return
		}
		for c := node.FirstChild; c != nil; c = c.NextSibling {
			walk(c)
		}
	}
	walk(n)
	return blk
}

// ── CSS helpers ──

func extractStyleBlocks(doc *html.Node) map[string]string {
	m := make(map[string]string)
	var walk func(*html.Node)
	walk = func(n *html.Node) {
		if n.Type == html.ElementNode && n.DataAtom == atom.Style {
			if n.FirstChild != nil {
				parseCSS(n.FirstChild.Data, m)
			}
		}
		for c := n.FirstChild; c != nil; c = c.NextSibling {
			walk(c)
		}
	}
	walk(doc)
	return m
}

func parseCSS(css string, m map[string]string) {
	re := regexp.MustCompile(`([.#]?[a-zA-Z][\w-]*)\s*\{([^}]+)\}`)
	for _, match := range re.FindAllStringSubmatch(css, -1) {
		selector := strings.TrimSpace(match[1])
		body := strings.TrimSpace(match[2])
		m[selector] = body
	}
}

func getInlineStyle(n *html.Node, styles map[string]string) string {
	inline := getAttr(n, "style")
	// Merge class styles
	cls := getAttr(n, "class")
	if cls != "" {
		for _, c := range strings.Fields(cls) {
			if s, ok := styles["."+c]; ok {
				inline = s + ";" + inline
			}
		}
	}
	// Merge id styles
	id := getAttr(n, "id")
	if id != "" {
		if s, ok := styles["#"+id]; ok {
			inline = s + ";" + inline
		}
	}
	return inline
}

func getAttr(n *html.Node, key string) string {
	for _, a := range n.Attr {
		if a.Key == key {
			return a.Val
		}
	}
	return ""
}

func extractColor(style string) string {
	re := regexp.MustCompile(`(?:^|;)\s*color:\s*(#[0-9a-fA-F]{3,6}|rgb[^;]+)`)
	if m := re.FindStringSubmatch(style); m != nil {
		return normalizeColor(m[1])
	}
	return ""
}

func extractBgColor(style string) string {
	re := regexp.MustCompile(`background(?:-color)?:\s*(#[0-9a-fA-F]{3,6}|rgb[^;]+)`)
	if m := re.FindStringSubmatch(style); m != nil {
		return normalizeColor(m[1])
	}
	return ""
}

func normalizeColor(c string) string {
	c = strings.TrimSpace(c)
	if strings.HasPrefix(c, "rgb") {
		re := regexp.MustCompile(`(\d+)`)
		parts := re.FindAllString(c, 3)
		if len(parts) == 3 {
			r, _ := strconv.Atoi(parts[0])
			g, _ := strconv.Atoi(parts[1])
			b, _ := strconv.Atoi(parts[2])
			return fmt.Sprintf("%02X%02X%02X", r, g, b)
		}
	}
	c = strings.TrimPrefix(c, "#")
	if len(c) == 3 {
		c = string([]byte{c[0], c[0], c[1], c[1], c[2], c[2]})
	}
	return strings.ToUpper(c)
}

func extractFontSize(style string) float64 {
	re := regexp.MustCompile(`font-size:\s*([\d.]+)(px|pt|em|rem)`)
	if m := re.FindStringSubmatch(style); m != nil {
		v, _ := strconv.ParseFloat(m[1], 64)
		switch m[2] {
		case "px":
			return v * 0.75 // px to pt
		case "em", "rem":
			return v * 12 // assume 12pt base
		}
		return v
	}
	return 0
}

func extractFontFamily(style string) string {
	re := regexp.MustCompile(`font-family:\s*([^;]+)`)
	if m := re.FindStringSubmatch(style); m != nil {
		fontStr := strings.TrimSpace(m[1])
		fonts := strings.Split(fontStr, ",")
		if len(fonts) > 0 {
			firstFont := strings.Trim(strings.TrimSpace(fonts[0]), `"'`)
			return firstFont
		}
	}
	return ""
}

func extractAlign(style string) string {
	re := regexp.MustCompile(`text-align:\s*(\w+)`)
	if m := re.FindStringSubmatch(style); m != nil {
		return m[1]
	}
	return ""
}

func headingSize(level int) float64 {
	sizes := []float64{0, 32, 26, 22, 18, 16, 14}
	if level >= 1 && level <= 6 {
		return sizes[level]
	}
	return 14
}

// ── pagination ──

func paginate(blocks []visualBlock) [][]visualBlock {
	var slides [][]visualBlock
	var current []visualBlock
	usedH := 0.0

	for _, blk := range blocks {
		h := estimateHeight(blk)

		// Force new slide on h1/h2 if we already have content
		if (blk.Type == "heading" && blk.Level <= 2) && len(current) > 0 {
			slides = append(slides, current)
			current = nil
			usedH = 0
		}

		if usedH+h > maxSlideH && len(current) > 0 {
			slides = append(slides, current)
			current = nil
			usedH = 0
		}

		current = append(current, blk)
		usedH += h
	}
	if len(current) > 0 {
		slides = append(slides, current)
	}
	if len(slides) == 0 {
		slides = append(slides, nil) // at least one empty slide
	}
	return slides
}

func estimateHeight(blk visualBlock) float64 {
	switch blk.Type {
	case "heading":
		return blk.FontSize/72.0 + 0.2
	case "text":
		lines := math.Ceil(float64(len(blk.Text)) / 80.0) // ~80 chars per line
		fs := blk.FontSize
		if fs == 0 { fs = 12 }
		return lines * (fs / 72.0) * 1.4
	case "image":
		return blk.ImgH + 0.2
	case "table":
		return float64(len(blk.TableRows)) * 0.35
	case "hr":
		return 0.15
	default:
		return 0.3
	}
}

// ── PPTX builder ──

func buildPPTX(stubPath, outPath string, slides [][]visualBlock, htmlPath string) {
	stub, err := zip.OpenReader(stubPath)
	if err != nil {
		log.Fatalf("Cannot open stub: %v", err)
	}
	defer stub.Close()

	outF, _ := os.Create(outPath)
	defer outF.Close()
	zw := zip.NewWriter(outF)
	defer zw.Close()

	htmlDir := filepath.Dir(htmlPath)

	// Copy non-slide entries from stub
	for _, f := range stub.File {
		if strings.Contains(f.Name, "ppt/slides/slide") ||
			strings.Contains(f.Name, "ppt/slides/_rels/slide") ||
			f.Name == "ppt/presentation.xml" ||
			f.Name == "ppt/_rels/presentation.xml.rels" ||
			strings.HasSuffix(f.Name, "[Content_Types].xml") {
			continue
		}
		src, _ := f.Open()
		dst, _ := zw.Create(f.Name)
		io.Copy(dst, src)
		src.Close()
	}

	var slideRefs []string

	for si, slideBlocks := range slides {
		slideNum := si + 1
		var elements []compiler.PptxElement
		var imgRels []string
		imgRId := 10
		curY := marginIn

		for _, blk := range slideBlocks {
			h := estimateHeight(blk)

			switch blk.Type {
			case "heading":
				fs := blk.FontSize
				if fs == 0 { fs = 24 }
				fontFace := "Arial"
				if blk.Font != "" { fontFace = blk.Font }
				el := compiler.PptxElement{
					Type: "text",
					X:    marginIn,
					Y:    curY,
					W:    contentW,
					H:    h,
					Text: blk.Text,
					TextConfig: &compiler.TextConfig{
						FontFace: fontFace,
						FontSize: fs,
						Bold:     true,
						Color:    safeColor(blk.Color, "000000"),
						Align:    safeAlign(blk.Align),
						Valign:   "top",
						Wrap:     true,
					},
				}
				if blk.BgColor != "" {
					el.Fill = map[string]interface{}{"color": blk.BgColor}
				}
				elements = append(elements, el)

			case "text":
				fs := blk.FontSize
				if fs == 0 { fs = 12 }
				fontFace := "Arial"
				if blk.Font != "" { fontFace = blk.Font }
				el := compiler.PptxElement{
					Type: "text",
					X:    marginIn + float64(blk.Level)*0.3,
					Y:    curY,
					W:    contentW - float64(blk.Level)*0.3,
					H:    h,
					Text: blk.Text,
					TextConfig: &compiler.TextConfig{
						FontFace: fontFace,
						FontSize: fs,
						Bold:     blk.Bold,
						Italic:   blk.Italic,
						Color:    safeColor(blk.Color, "333333"),
						Align:    safeAlign(blk.Align),
						Valign:   "top",
						Wrap:     true,
					},
				}
				if blk.BgColor != "" {
					el.Fill = map[string]interface{}{"color": blk.BgColor}
				}
				elements = append(elements, el)

			case "image":
				imgPath := blk.ImgSrc
				if !strings.HasPrefix(imgPath, "http") {
					imgPath = filepath.Join(htmlDir, imgPath)
				}
				if data, err := readImage(imgPath); err == nil {
					ext := filepath.Ext(imgPath)
					if ext == "" { ext = ".png" }
					mediaName := fmt.Sprintf("ppt/media/anyhtml_s%d_i%d%s", slideNum, imgRId, ext)
					dst, _ := zw.Create(mediaName)
					dst.Write(data)

					rId := fmt.Sprintf("rId%d", imgRId)
					imgRels = append(imgRels, fmt.Sprintf(
						`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/anyhtml_s%d_i%d%s"/>`,
						rId, slideNum, imgRId, ext))

					el := compiler.PptxElement{
						Type:     "image",
						X:        marginIn,
						Y:        curY,
						W:        math.Min(blk.ImgW, contentW),
						H:        blk.ImgH,
						ImageRId: rId,
					}
					elements = append(elements, el)
					imgRId++
				}

			case "table":
				tbl := &mapper.TableConfig{HasHeader: true}
				cols := 0
				for _, row := range blk.TableRows {
					if len(row) > cols { cols = len(row) }
				}
				colW := contentW / float64(cols)
				for i := 0; i < cols; i++ {
					tbl.ColWidths = append(tbl.ColWidths, int64(colW*914400))
				}
				for ri, row := range blk.TableRows {
					var cells []mapper.TableCell
					for _, cell := range row {
						tc := mapper.TableCell{
							Text: cell, FontSize: 1000, Color: "333333",
							FillColor: "FFFFFF", Align: "l",
							FontFace: "Arial",
						}
						if ri == 0 {
							tc.Bold = true
							tc.FillColor = "F2F2F2"
						}
						cells = append(cells, tc)
					}
					tbl.Rows = append(tbl.Rows, cells)
					tbl.RowHeights = append(tbl.RowHeights, int64(0.35*914400))
				}

				el := compiler.PptxElement{
					Type: "table", X: marginIn, Y: curY,
					W: contentW, H: h, Table: tbl,
				}
				elements = append(elements, el)

			case "hr":
				el := compiler.PptxElement{
					Type: "shape", X: marginIn, Y: curY + 0.05,
					W: contentW, H: 0.02,
					Fill: map[string]interface{}{"color": "CCCCCC"},
					Shape: mapper.ShapeConfig{ShapeType: "rect"},
				}
				elements = append(elements, el)
			}

			curY += h
		}

		// Slide number footer
		elements = append(elements, compiler.PptxElement{
			Type: "text", X: slideW - 1.5, Y: slideH - 0.5, W: 1.0, H: 0.3,
			Text: fmt.Sprintf("%d / %d", slideNum, len(slides)),
			TextConfig: &compiler.TextConfig{
				FontSize: 9, Color: "999999", Align: "right", Valign: "bottom",
				FontFace: "Arial", Wrap: false,
			},
		})

		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		openxml.BuildPPTXSlide(zw, elements, slideName)

		relsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
` + strings.Join(imgRels, "\n") + `
</Relationships>`
		relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
		rd, _ := zw.Create(relsName)
		io.WriteString(rd, relsXML)

		slideRefs = append(slideRefs, fmt.Sprintf(`<p:sldId id="%d" r:id="rId%d"/>`, 256+slideNum, 100+slideNum))
		log.Printf("Slide %d: %d elements", slideNum, len(elements))
	}

	// presentation.xml + rels + content_types (reuse from roundtrip logic)
	updatePres(zw, slideRefs, stub)
	updateCT(zw, len(slides), stub)
}

func readImage(path string) ([]byte, error) {
	if strings.HasPrefix(path, "http") {
		resp, err := http.Get(path)
		if err != nil { return nil, err }
		defer resp.Body.Close()
		return io.ReadAll(resp.Body)
	}
	return os.ReadFile(path)
}

func safeColor(c, fallback string) string {
	if c == "" { return fallback }
	return strings.TrimPrefix(c, "#")
}

func safeAlign(a string) string {
	switch a {
	case "center": return "center"
	case "right": return "right"
	default: return "left"
	}
}

func updatePres(zw *zip.Writer, refs []string, stub *zip.ReadCloser) {
	for _, f := range stub.File {
		if f.Name != "ppt/presentation.xml" { continue }
		src, _ := f.Open()
		data, _ := io.ReadAll(src)
		src.Close()
		content := string(data)
		s := strings.Index(content, "<p:sldIdLst>")
		e := strings.Index(content, "</p:sldIdLst>")
		if s != -1 && e != -1 {
			content = content[:s] + "<p:sldIdLst>" + strings.Join(refs, "\n") + "</p:sldIdLst>" + content[e+len("</p:sldIdLst>"):]
		}
		dst, _ := zw.Create("ppt/presentation.xml")
		io.WriteString(dst, content)

		for _, rf := range stub.File {
			if rf.Name != "ppt/_rels/presentation.xml.rels" { continue }
			rs, _ := rf.Open()
			rd, _ := io.ReadAll(rs)
			rs.Close()
			rc := string(rd)
			slideRelRe := regexp.MustCompile(`<Relationship[^>]*relationships/slide"[^>]*/>`)
			rc = slideRelRe.ReplaceAllString(rc, "")
			for i := range refs {
				rel := fmt.Sprintf(`<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>`, 100+i+1, i+1)
				rc = strings.Replace(rc, "</Relationships>", rel+"\n</Relationships>", 1)
			}
			d, _ := zw.Create("ppt/_rels/presentation.xml.rels")
			io.WriteString(d, rc)
		}
		return
	}
}

func updateCT(zw *zip.Writer, count int, stub *zip.ReadCloser) {
	for _, f := range stub.File {
		if !strings.HasSuffix(f.Name, "[Content_Types].xml") { continue }
		src, _ := f.Open()
		data, _ := io.ReadAll(src)
		src.Close()
		c := string(data)

		// Add default image content types if missing
		defaults := []string{
			`<Default Extension="png" ContentType="image/png"/>`,
			`<Default Extension="jpeg" ContentType="image/jpeg"/>`,
			`<Default Extension="jpg" ContentType="image/jpeg"/>`,
			`<Default Extension="gif" ContentType="image/gif"/>`,
		}
		for _, d := range defaults {
			if !strings.Contains(c, d) {
				c = strings.Replace(c, "</Types>", d+"\n</Types>", 1)
			}
		}

		for i := 1; i <= count; i++ {
			ov := fmt.Sprintf(`<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`, i)
			if !strings.Contains(c, ov) {
				c = strings.Replace(c, "</Types>", ov+"\n</Types>", 1)
			}
		}
		dst, _ := zw.Create("[Content_Types].xml")
		io.WriteString(dst, c)
		return
	}
}
