package main

import (
	"archive/zip"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/rhino-acoustic/pptx-engine/internal/compiler"
	"github.com/rhino-acoustic/pptx-engine/internal/openxml"
)

func main() {
	jsonDir := flag.String("json", "", "Directory containing per-slide JSON files from batch_scraper.js")
	stubFile := flag.String("stub", "", "Stub PPTX file to use as template (for layouts/themes)")
	outFile := flag.String("out", "output.pptx", "Output PPTX file")
	flag.Parse()

	if *jsonDir == "" || *stubFile == "" {
		log.Fatal("Usage: -json <slides_json_dir> -stub <template.pptx> -out <output.pptx>")
	}

	// Read all JSON files sorted
	files, err := filepath.Glob(filepath.Join(*jsonDir, "slide_*.json"))
	if err != nil || len(files) == 0 {
		log.Fatalf("No slide JSON files found in %s", *jsonDir)
	}
	sort.Strings(files)
	log.Printf("Found %d slide JSON files", len(files))

	// Open stub PPTX for themes/layouts
	stub, err := zip.OpenReader(*stubFile)
	if err != nil {
		log.Fatalf("Failed to open stub: %v", err)
	}
	defer stub.Close()

	// Create output PPTX
	outF, err := os.Create(*outFile)
	if err != nil {
		log.Fatalf("Failed to create output: %v", err)
	}
	defer outF.Close()
	zipWriter := zip.NewWriter(outF)
	defer zipWriter.Close()

	// Copy non-slide files from stub (themes, layouts, etc.)
	for _, f := range stub.File {
		if strings.Contains(f.Name, "ppt/slides/slide") {
			continue // Skip all slides from stub
		}
		if strings.Contains(f.Name, "ppt/slides/_rels/slide") {
			continue
		}
		// Skip files that will be regenerated
		if f.Name == "ppt/presentation.xml" || f.Name == "ppt/_rels/presentation.xml.rels" || strings.HasSuffix(f.Name, "[Content_Types].xml") {
			continue
		}
		src, err := f.Open()
		if err != nil {
			continue
		}
		dst, err := zipWriter.Create(f.Name)
		if err != nil {
			src.Close()
			continue
		}
		io.Copy(dst, src)
		src.Close()
	}

	// Process each slide
	var slideRefs []string
	for i, jsonFile := range files {
		slideNum := i + 1
		log.Printf("Processing slide %d: %s", slideNum, filepath.Base(jsonFile))

		rawBytes, err := ioutil.ReadFile(jsonFile)
		if err != nil {
			log.Printf("  ERROR reading: %v", err)
			continue
		}

		var rawNodes []compiler.RawNode
		if err := json.Unmarshal(rawBytes, &rawNodes); err != nil {
			log.Printf("  ERROR parsing: %v", err)
			continue
		}

		// Collect images and pack into ZIP
		var imgRels []string // rId -> media path pairs for slide rels
		imgRIdStart := 10    // start image rIds at 10 to avoid collision with rId1 (slideLayout)
		jsonDir2 := filepath.Dir(jsonFile)

		for idx, raw := range rawNodes {
			if raw.Type == "image" && raw.ImageUrl != "" {
				imgPath := filepath.Join(jsonDir2, raw.ImageUrl)
				imgData, err := ioutil.ReadFile(imgPath)
				if err != nil {
					log.Printf("  WARN: image %s not found: %v", raw.ImageUrl, err)
					continue
				}
				mediaName := fmt.Sprintf("ppt/media/slide%d_img%d.jpg", slideNum, idx)
				dst, err := zipWriter.Create(mediaName)
				if err == nil {
					dst.Write(imgData)
				}
				rId := fmt.Sprintf("rId%d", imgRIdStart)
				imgRels = append(imgRels, fmt.Sprintf(
					`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/slide%d_img%d.jpg"/>`,
					rId, slideNum, idx))
				rawNodes[idx].ImageRId = rId
				imgRIdStart++
			}
		}

		// Compile all nodes with post-processing pipeline (dash2pptx.mjs parity)
		elements := compiler.CompileSlide(rawNodes)
		// Pass through image rIds
		for j := range elements {
			for _, raw := range rawNodes {
				if raw.ImageRId != "" && elements[j].Type == "image" && elements[j].ImageRId == "" {
					// Match by position proximity
					px2in := 13.333 / 1280.0
					if abs64(elements[j].X-raw.X*px2in) < 0.01 && abs64(elements[j].Y-raw.Y*px2in) < 0.01 {
						elements[j].ImageRId = raw.ImageRId
					}
				}
			}
		}

		// Sort by ZIndex
		elements = compiler.SortByZIndex(elements)

		// Apply clipping
		compiler.ApplyOverflowClipping(elements)

		// Write slide XML
		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		if err := openxml.BuildPPTXSlide(zipWriter, elements, slideName); err != nil {
			log.Printf("  ERROR building slide: %v", err)
			continue
		}

		// Write slide rels (slideLayout + images)
		relsXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
` + strings.Join(imgRels, "\n") + `
</Relationships>`
		relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
		relsDst, _ := zipWriter.Create(relsName)
		io.WriteString(relsDst, relsXML)

		slideRefs = append(slideRefs, fmt.Sprintf(`<p:sldId id="%d" r:id="rId%d"/>`, 256+slideNum, 100+slideNum))

		imgCount := 0
		for _, el := range elements {
			if el.Type == "image" {
				imgCount++
			}
		}
		log.Printf("  OK: %d elements (%d images)", len(elements), imgCount)
	}

	// Update presentation.xml with slide references
	updatePresentationXML(zipWriter, slideRefs, stub)

	// Update [Content_Types].xml
	updateContentTypes(zipWriter, len(files), stub)

	log.Printf("✅ DONE: %d slides → %s", len(files), *outFile)
}

func updatePresentationXML(zw *zip.Writer, slideRefs []string, stub *zip.ReadCloser) {
	for _, f := range stub.File {
		if f.Name == "ppt/presentation.xml" {
			src, _ := f.Open()
			data, _ := ioutil.ReadAll(src)
			src.Close()

			content := string(data)
			// Find sldIdLst and replace
			start := strings.Index(content, "<p:sldIdLst>")
			end := strings.Index(content, "</p:sldIdLst>")
			if start != -1 && end != -1 {
				newSldIdLst := "<p:sldIdLst>" + strings.Join(slideRefs, "\n") + "</p:sldIdLst>"
				content = content[:start] + newSldIdLst + content[end+len("</p:sldIdLst>"):]
			}

			// Build relationship entries
			var relsEntries []string
			for i := range slideRefs {
				relsEntries = append(relsEntries, fmt.Sprintf(
					`<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>`,
					100+i+1, i+1))
			}

			dst, _ := zw.Create("ppt/presentation.xml")
			io.WriteString(dst, content)

			// Also update presentation.xml.rels
			for _, rf := range stub.File {
				if rf.Name == "ppt/_rels/presentation.xml.rels" {
					rsrc, _ := rf.Open()
					rdata, _ := ioutil.ReadAll(rsrc)
					rsrc.Close()
					rcontent := string(rdata)
					// Append slide rels before closing tag
					for _, rel := range relsEntries {
						if !strings.Contains(rcontent, rel) {
							rcontent = strings.Replace(rcontent, "</Relationships>", rel+"\n</Relationships>", 1)
						}
					}
					rdst, _ := zw.Create("ppt/_rels/presentation.xml.rels")
					io.WriteString(rdst, rcontent)
				}
			}
			return
		}
	}
}

func updateContentTypes(zw *zip.Writer, slideCount int, stub *zip.ReadCloser) {
	for _, f := range stub.File {
		if strings.HasSuffix(f.Name, "[Content_Types].xml") {
			src, _ := f.Open()
			data, _ := ioutil.ReadAll(src)
			src.Close()
			content := string(data)

			// Add default image content types if missing
			defaults := []string{
				`<Default Extension="png" ContentType="image/png"/>`,
				`<Default Extension="jpeg" ContentType="image/jpeg"/>`,
				`<Default Extension="jpg" ContentType="image/jpeg"/>`,
				`<Default Extension="gif" ContentType="image/gif"/>`,
			}
			for _, d := range defaults {
				if !strings.Contains(content, d) {
					content = strings.Replace(content, "</Types>", d+"\n</Types>", 1)
				}
			}

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
}

func abs64(x float64) float64 {
	if x < 0 {
		return -x
	}
	return x
}

