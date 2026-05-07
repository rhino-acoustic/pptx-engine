package main

import (
	"archive/zip"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/rhino-acoustic/pptx-engine/internal/compiler"
	"github.com/rhino-acoustic/pptx-engine/internal/openxml"
	"github.com/rhino-acoustic/pptx-engine/internal/reverse"
)

func main() {
	htmlFile := flag.String("html", "", "Input HTML file (from pptx2html)")
	stubFile := flag.String("stub", "", "Stub PPTX file for themes/layouts")
	outFile := flag.String("out", "roundtrip.pptx", "Output PPTX file")
	flag.Parse()

	if *htmlFile == "" || *stubFile == "" {
		log.Fatal("Usage: html2pptx -html <file.html> -stub <template.pptx> [-out output.pptx]")
	}

	log.Printf("=== HTML → PPTX Roundtrip ===")
	log.Printf("Input:  %s", *htmlFile)
	log.Printf("Stub:   %s", *stubFile)
	log.Printf("Output: %s", *outFile)

	// 1. Parse HTML → ParsedDoc
	doc, err := reverse.ParseHTMLFile(*htmlFile)
	if err != nil {
		log.Fatalf("ParseHTML Error: %v", err)
	}
	log.Printf("Parsed %d slides (%.0f×%.0f pt)", doc.SlideCount, doc.SlideWidth, doc.SlideHeight)

	// 2. Open stub PPTX
	stub, err := zip.OpenReader(*stubFile)
	if err != nil {
		log.Fatalf("Failed to open stub: %v", err)
	}
	defer stub.Close()

	// 3. Create output PPTX
	outF, err := os.Create(*outFile)
	if err != nil {
		log.Fatalf("Failed to create output: %v", err)
	}
	defer outF.Close()
	zw := zip.NewWriter(outF)
	defer zw.Close()

	// 4. Copy non-slide files from stub
	for _, f := range stub.File {
		if strings.Contains(f.Name, "ppt/slides/slide") ||
			strings.Contains(f.Name, "ppt/slides/_rels/slide") ||
			f.Name == "ppt/presentation.xml" ||
			f.Name == "ppt/_rels/presentation.xml.rels" ||
			strings.HasSuffix(f.Name, "[Content_Types].xml") {
			continue
		}
		src, err := f.Open()
		if err != nil { continue }
		dst, err := zw.Create(f.Name)
		if err != nil { src.Close(); continue }
		io.Copy(dst, src)
		src.Close()
	}

	// 5. Process each slide
	var slideRefs []string
	htmlDir := filepath.Dir(*htmlFile)

	for _, slide := range doc.Slides {
		slideNum := slide.SlideIndex
		log.Printf("Slide %d: %d shapes", slideNum, len(slide.Shapes))

		// Convert ParsedShape → HTMLShape → PptxElement
		htmlShapes := make([]compiler.HTMLShape, 0, len(slide.Shapes))
		var imgRels []string
		imgRIdStart := 10

		for _, ps := range slide.Shapes {
			hs := compiler.HTMLShape{
				Type:             ps.Type,
				X:                ps.X,
				Y:                ps.Y,
				W:                ps.W,
				H:                ps.H,
				Rotation:         ps.Rotation,
				ShapeType:        ps.ShapeType,
				FillType:         ps.FillType,
				FillColor:        ps.FillColor,
				FillTransparency: ps.FillTransparency,
				GradientAngle:    ps.GradientAngle,
				BorderColor:      ps.BorderColor,
				BorderWidth:      ps.BorderWidth,
				BorderRadius:     ps.BorderRadius,
				HasShadow:        ps.HasShadow,
				Valign:           ps.Valign,
				PadL:             ps.PadL,
				PadR:             ps.PadR,
				PadT:             ps.PadT,
				PadB:             ps.PadB,
				ImagePath:        ps.ImagePath,
				CropL:            ps.CropL,
				CropT:            ps.CropT,
				CropR:            ps.CropR,
				CropB:            ps.CropB,
				HasText:          ps.HasText,
				TableRows:        ps.TableRows,
				TableCols:        ps.TableCols,
				TableColWidths:   ps.TableColWidths,
				TableRowHeights:  ps.TableRowHeights,
				TableData:        ps.TableData,
				ChartSVG:         ps.ChartSVG,
			}

			// Gradient stops
			for _, gs := range ps.GradientStops {
				hs.GradientStops = append(hs.GradientStops, compiler.HTMLGradStop{
					Color:    gs.Color,
					Position: gs.Position,
					Alpha:    gs.Alpha,
				})
			}

			// Text runs
			for _, tr := range ps.TextRuns {
				hs.TextRuns = append(hs.TextRuns, compiler.HTMLTextRun{
					Text:      tr.Text,
					Font:      tr.Font,
					Size:      tr.Size,
					Bold:      tr.Bold,
					Italic:    tr.Italic,
					Underline: tr.Underline,
					Color:     tr.Color,
					Align:     tr.Align,
					Bullet:    tr.Bullet,
				})
			}

			// Pack images into ZIP
			if ps.ImagePath != "" && !ps.HasText {
				// Resolve relative path
				imgPath := ps.ImagePath
				if !filepath.IsAbs(imgPath) {
					imgPath = filepath.Join(htmlDir, imgPath)
				}
				imgData, err := os.ReadFile(imgPath)
				if err != nil {
					log.Printf("  WARN: image %s not found: %v", ps.ImagePath, err)
				} else {
					// Preserve original extension for Content_Types compatibility
					ext := strings.ToLower(filepath.Ext(imgPath))
					if ext == "" { ext = ".png" }
					mediaName := fmt.Sprintf("ppt/media/slide%d_img%d%s", slideNum, imgRIdStart, ext)
					dst, err := zw.Create(mediaName)
					if err == nil {
						dst.Write(imgData)
					}
					rId := fmt.Sprintf("rId%d", imgRIdStart)
					imgRels = append(imgRels, fmt.Sprintf(
						`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/slide%d_img%d%s"/>`,
						rId, slideNum, imgRIdStart, ext))
					hs.ImageRId = rId
					imgRIdStart++
				}
			}

			htmlShapes = append(htmlShapes, hs)
		}

		// Compile to PptxElements
		elements := compiler.CompileHTMLSlide(htmlShapes)

		// Write slide XML
		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		if err := openxml.BuildPPTXSlide(zw, elements, slideName); err != nil {
			log.Printf("  ERROR building slide: %v", err)
			continue
		}

		// Write slide rels
		relsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
` + strings.Join(imgRels, "\n") + `
</Relationships>`
		relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
		relsDst, _ := zw.Create(relsName)
		io.WriteString(relsDst, relsXML)

		slideRefs = append(slideRefs, fmt.Sprintf(`<p:sldId id="%d" r:id="rId%d"/>`, 256+slideNum, 100+slideNum))

		imgCount := 0
		for _, el := range elements {
			if el.Type == "image" { imgCount++ }
		}
		log.Printf("  OK: %d elements (%d images)", len(elements), imgCount)
	}

	// 6. Update presentation.xml
	updatePresentation(zw, slideRefs, stub)

	// 7. Update [Content_Types].xml
	updateTypes(zw, doc.SlideCount, stub)

	log.Printf("✅ DONE: %d slides → %s", doc.SlideCount, *outFile)
}

func updatePresentation(zw *zip.Writer, slideRefs []string, stub *zip.ReadCloser) {
	for _, f := range stub.File {
		if f.Name != "ppt/presentation.xml" { continue }
		src, _ := f.Open()
		data, _ := io.ReadAll(src)
		src.Close()

		content := string(data)
		start := strings.Index(content, "<p:sldIdLst>")
		end := strings.Index(content, "</p:sldIdLst>")
		if start != -1 && end != -1 {
			newList := "<p:sldIdLst>" + strings.Join(slideRefs, "\n") + "</p:sldIdLst>"
			content = content[:start] + newList + content[end+len("</p:sldIdLst>"):]
		}

		dst, _ := zw.Create("ppt/presentation.xml")
		io.WriteString(dst, content)

		// Update presentation.xml.rels — REMOVE old slide rels, ADD new ones
		for _, rf := range stub.File {
			if rf.Name != "ppt/_rels/presentation.xml.rels" { continue }
			rsrc, _ := rf.Open()
			rdata, _ := io.ReadAll(rsrc)
			rsrc.Close()
			rcontent := string(rdata)

			// Remove ALL existing slide relationships from stub
			slideRelRe := regexp.MustCompile(`<Relationship[^>]*relationships/slide"[^>]*/>`)
			rcontent = slideRelRe.ReplaceAllString(rcontent, "")
			// Clean up empty lines left after removal
			rcontent = regexp.MustCompile(`\n\s*\n`).ReplaceAllString(rcontent, "\n")

			// Add new slide relationships with our rIds
			for i := range slideRefs {
				rel := fmt.Sprintf(`<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>`,
					100+i+1, i+1)
				rcontent = strings.Replace(rcontent, "</Relationships>", rel+"\n</Relationships>", 1)
			}
			rdst, _ := zw.Create("ppt/_rels/presentation.xml.rels")
			io.WriteString(rdst, rcontent)
		}
		return
	}
}

func updateTypes(zw *zip.Writer, slideCount int, stub *zip.ReadCloser) {
	for _, f := range stub.File {
		if !strings.HasSuffix(f.Name, "[Content_Types].xml") { continue }
		src, _ := f.Open()
		data, _ := io.ReadAll(src)
		src.Close()
		content := string(data)
		for i := 1; i <= slideCount; i++ {
			override := fmt.Sprintf(`<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`, i)
			if !strings.Contains(content, override) {
				content = strings.Replace(content, "</Types>", override+"\n</Types>", 1)
			}
		}
		dst, _ := zw.Create("[Content_Types].xml")
		io.WriteString(dst, content)
		return
	}
}

// Unused import guard
var _ = json.Marshal
