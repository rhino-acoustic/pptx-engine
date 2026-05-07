package main

import (
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/rhino-acoustic/pptx-engine/internal/reverse"
)

func main() {
	inputPPTX := flag.String("input", "", "Input PPTX file path")
	outputHTML := flag.String("output", "testdata/output.html", "Output HTML file path")
	imageDir := flag.String("images", "slide_images", "Directory to extract slide images")
	slideNum := flag.Int("slide", 0, "Single slide number to extract (0 = all)")
	flag.Parse()

	if *inputPPTX == "" {
		log.Fatal("Usage: go run cmd/pptx2html/main.go -input <file.pptx> [-slide N]")
	}

	log.Printf("=== PPTX → HTML: %s ===", *inputPPTX)

	parsedDoc, err := reverse.ParsePPTX(*inputPPTX, *imageDir)
	if err != nil {
		log.Fatalf("ParsePPTX Error: %v", err)
	}

	log.Printf("Parsed %d slides", len(parsedDoc.Slides))

	if *slideNum > 0 {
		// Single slide mode: extract only the specified slide
		found := false
		for _, slide := range parsedDoc.Slides {
			if slide.SlideIndex == *slideNum {
				parsedDoc.Slides = []reverse.ParsedSlide{slide}
				parsedDoc.SlideCount = 1
				found = true
				break
			}
		}
		if !found {
			log.Fatalf("Slide %d not found (total: %d)", *slideNum, len(parsedDoc.Slides))
		}

		// Auto-generate output filename if default
		if *outputHTML == "testdata/output.html" {
			*outputHTML = fmt.Sprintf("testdata/slide_%02d.html", *slideNum)
		}
		log.Printf("Single slide mode: extracting slide %d", *slideNum)
	}

	html := reverse.GenerateHTMLFromParsed(parsedDoc)
	if err := os.WriteFile(*outputHTML, []byte(html), 0644); err != nil {
		log.Fatalf("WriteFile Error: %v", err)
	}

	log.Printf("Successfully generated %s (%d bytes)", *outputHTML, len(html))
}
