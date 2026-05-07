package openxml

import (
	"archive/zip"
	"encoding/xml"
	"io/ioutil"
	"regexp"
	"strings"

	"github.com/rhino-acoustic/pptx-engine/internal/reverse"
)

type Relationships struct {
	Rel []Relationship `xml:"Relationship"`
}

type Relationship struct {
	Id     string `xml:"Id,attr"`
	Target string `xml:"Target,attr"`
}

type SpTree struct {
	GrpSp  []GrpSp `xml:"grpSp"`
	Shapes []Shape `xml:"sp"`
	Pics   []Pic   `xml:"pic"`
	Frames []Frame `xml:"graphicFrame"`
}

type GrpSp struct {
	GrpSp  []GrpSp `xml:"grpSp"`
	Shapes []Shape `xml:"sp"`
	Pics   []Pic   `xml:"pic"`
	Frames []Frame `xml:"graphicFrame"`
}

type Shape struct {
	SpPr   SpPr   `xml:"spPr"`
	TxBody TxBody `xml:"txBody"`
}

type Pic struct {
	SpPr     SpPr     `xml:"spPr"`
	BlipFill BlipFill `xml:"blipFill"`
}

type BlipFill struct {
	Blip Blip `xml:"blip"`
}

type Blip struct {
	Embed string `xml:"embed,attr"`
}

type Frame struct {
	Xfrm    Xfrm    `xml:"xfrm"`
	Graphic Graphic `xml:"graphic"`
}

type Graphic struct {
	GraphicData GraphicData `xml:"graphicData"`
}

type GraphicData struct {
	Tbl Table `xml:"tbl"`
}

type Table struct {
	Tr []TableRow `xml:"tr"`
}

type TableRow struct {
	Tc []TableCell `xml:"tc"`
}

type TableCell struct {
	TxBody TxBody `xml:"txBody"`
}

type SpPr struct {
	Xfrm      Xfrm      `xml:"xfrm"`
	PrstGeom  PrstGeom  `xml:"prstGeom"`
	SolidFill SolidFill `xml:"solidFill"`
}

type Xfrm struct {
	Rot int64 `xml:"rot,attr"`
	Off Off   `xml:"off"`
	Ext Ext   `xml:"ext"`
}

type Off struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

type Ext struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

type PrstGeom struct {
	Prst string `xml:"prst,attr"`
}

type SolidFill struct {
	SrgbClr SrgbClr `xml:"srgbClr"`
}

type SrgbClr struct {
	Val   string  `xml:"val,attr"`
	Alpha []Alpha `xml:"alpha"`
}

type Alpha struct {
	Val int64 `xml:"val,attr"`
}

type TxBody struct {
	P []Paragraph `xml:"p"`
}

type Paragraph struct {
	PPr PPr   `xml:"pPr"`
	R   []Run `xml:"r"`
}

type PPr struct {
	Algn string `xml:"algn,attr"` // l, r, ctr
}

type Run struct {
	RPr RPr    `xml:"rPr"`
	T   string `xml:"t"`
}

type RPr struct {
	Sz        int       `xml:"sz,attr"` // font size (100 = 1pt)
	B         int       `xml:"b,attr"`  // bold (1=true)
	I         int       `xml:"i,attr"`  // italic (1=true)
	SolidFill SolidFill `xml:"solidFill"`
}

func emuToPt(emu int64) float64 {
	return float64(emu) / 12700.0
}

func extractTextWithStyle(txBody TxBody) (bool, []reverse.ComTextRun) {
	var runs []reverse.ComTextRun
	for _, p := range txBody.P {
		align := 0
		if p.PPr.Algn == "ctr" {
			align = 1
		} else if p.PPr.Algn == "r" {
			align = 2
		}

		for _, r := range p.R {
			if strings.TrimSpace(r.T) == "" {
				continue
			}
			size := float64(r.RPr.Sz) / 100.0
			if size == 0 {
				size = 14.0
			}
			
			// Know-how: Scale font sizes properly (Pt to Px approximation for Web)
			scaledSize := size * 1.02 // as mentioned in the VBS logs

			color := r.RPr.SolidFill.SrgbClr.Val
			if color == "" {
				color = "000000"
			}

			run := reverse.ComTextRun{
				Text:  r.T,
				Color: color,
				Size:  scaledSize,
				Bold:  r.RPr.B,
				Italic: r.RPr.I,
				Align: align,
			}
			runs = append(runs, run)
		}
		// Insert line break logic (Know-how: Prevent Slide 1 merged text)
		runs = append(runs, reverse.ComTextRun{Text: "\n"})
	}

	if len(runs) > 0 {
		return true, runs
	}
	return false, nil
}

func traverseGrpSp(grp GrpSp, shapesList *[]reverse.ComShape, relsMap map[string]string) {
	for _, s := range grp.Shapes {
		rot := float64(s.SpPr.Xfrm.Rot) / 60000.0
		shape := reverse.ComShape{
			Type:     "Shape",
			X:        emuToPt(s.SpPr.Xfrm.Off.X),
			Y:        emuToPt(s.SpPr.Xfrm.Off.Y),
			W:        emuToPt(s.SpPr.Xfrm.Ext.Cx),
			H:        emuToPt(s.SpPr.Xfrm.Ext.Cy),
			Rotation: rot,
		}
		
		hasText, textRuns := extractTextWithStyle(s.TxBody)
		if hasText {
			shape.HasText = true
			shape.TextRuns = textRuns
		}
		
		if s.SpPr.SolidFill.SrgbClr.Val != "" {
			color := s.SpPr.SolidFill.SrgbClr.Val
			transparency := 0.0
			if len(s.SpPr.SolidFill.SrgbClr.Alpha) > 0 {
				transparency = 1.0 - (float64(s.SpPr.SolidFill.SrgbClr.Alpha[0].Val) / 100000.0)
			}
			
			// Know-how: Empty white background boxes (`FFFFFF`) should be transparent to prevent Z-index obscuring
			if color == "FFFFFF" && !hasText {
				color = "transparent"
			}
			shape.FillType = "solid"
			shape.FillColor = "#" + color
			shape.FillTransparency = transparency
		}
		*shapesList = append(*shapesList, shape)
	}

	for _, pic := range grp.Pics {
		// Know-how: Image bounding box + extract base64 mapping via .rels
		imgPath := ""
		if pic.BlipFill.Blip.Embed != "" {
			imgPath = relsMap[pic.BlipFill.Blip.Embed]
		}
		
		rot := float64(pic.SpPr.Xfrm.Rot) / 60000.0
		shape := reverse.ComShape{
			Type:      "Image",
			X:         emuToPt(pic.SpPr.Xfrm.Off.X),
			Y:         emuToPt(pic.SpPr.Xfrm.Off.Y),
			W:         emuToPt(pic.SpPr.Xfrm.Ext.Cx),
			H:         emuToPt(pic.SpPr.Xfrm.Ext.Cy),
			Rotation:  rot,
			ImagePath: imgPath,
		}
		*shapesList = append(*shapesList, shape)
	}

	for _, f := range grp.Frames {
		for _, tr := range f.Graphic.GraphicData.Tbl.Tr {
			for _, tc := range tr.Tc {
				hasText, textRuns := extractTextWithStyle(tc.TxBody)
				if hasText {
					shape := reverse.ComShape{
						Type:     "TableText",
						HasText:  true,
						TextRuns: textRuns,
					}
					*shapesList = append(*shapesList, shape)
				}
			}
		}
	}

	for _, childGrp := range grp.GrpSp {
		traverseGrpSp(childGrp, shapesList, relsMap)
	}
}

func ExtractSlides(pptxPath string) (*reverse.ComDoc, error) {
	r, err := zip.OpenReader(pptxPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	doc := &reverse.ComDoc{
		SlideWidth:  1280,
		SlideHeight: 720,
	}

	slideRegex := regexp.MustCompile(`ppt/slides/slide(\d+)\.xml$`)
	relsRegex := regexp.MustCompile(`ppt/slides/_rels/slide(\d+)\.xml\.rels$`)

	// 1. Build relationship map per slide
	slideRels := make(map[string]map[string]string)
	for _, f := range r.File {
		match := relsRegex.FindStringSubmatch(f.Name)
		if len(match) > 1 {
			slideNum := match[1]
			data, _ := readZipFile(f)
			xmlStr := string(data)
			var rels Relationships
			xml.Unmarshal([]byte(xmlStr), &rels)

			relMap := make(map[string]string)
			for _, rel := range rels.Rel {
				relMap[rel.Id] = rel.Target // e.g. "../media/image1.png"
			}
			slideRels[slideNum] = relMap
		}
	}

	// 2. Parse slides
	for _, f := range r.File {
		match := slideRegex.FindStringSubmatch(f.Name)
		if len(match) > 1 {
			slideNum := match[1]
			slideData, _ := readZipFile(f)

			xmlStr := string(slideData)
			xmlStr = regexp.MustCompile(`<\w+:`).ReplaceAllString(xmlStr, "<")
			xmlStr = regexp.MustCompile(`</\w+:`).ReplaceAllString(xmlStr, "</")

			var tree SpTree
			start := strings.Index(xmlStr, "<spTree>")
			end := strings.Index(xmlStr, "</spTree>")
			if start != -1 && end != -1 {
				treeData := xmlStr[start : end+9]
				xml.Unmarshal([]byte(treeData), &tree)
			}

			relMap := slideRels[slideNum]
			slide := reverse.ComSlide{SlideIndex: 1} // Should parse index correctly later
			var allShapes []reverse.ComShape

			// Top level shapes
			for _, s := range tree.Shapes {
				rot := float64(s.SpPr.Xfrm.Rot) / 60000.0
				shape := reverse.ComShape{
					Type:     "Shape",
					X:        emuToPt(s.SpPr.Xfrm.Off.X),
					Y:        emuToPt(s.SpPr.Xfrm.Off.Y),
					W:        emuToPt(s.SpPr.Xfrm.Ext.Cx),
					H:        emuToPt(s.SpPr.Xfrm.Ext.Cy),
					Rotation: rot,
				}
				hasText, textRuns := extractTextWithStyle(s.TxBody)
				if hasText {
					shape.HasText = true
					shape.TextRuns = textRuns
				}
				if s.SpPr.SolidFill.SrgbClr.Val != "" {
					color := s.SpPr.SolidFill.SrgbClr.Val
					transparency := 0.0
					if len(s.SpPr.SolidFill.SrgbClr.Alpha) > 0 {
						transparency = 1.0 - (float64(s.SpPr.SolidFill.SrgbClr.Alpha[0].Val) / 100000.0)
					}
					if color == "FFFFFF" && !hasText {
						color = "transparent"
					}
					shape.FillType = "solid"
					if color == "transparent" {
						shape.FillColor = color
					} else {
						shape.FillColor = "#" + color
					}
					shape.FillTransparency = transparency
				}
				allShapes = append(allShapes, shape)
			}

			// Top level pictures
			for _, pic := range tree.Pics {
				imgPath := ""
				if pic.BlipFill.Blip.Embed != "" {
					imgPath = relMap[pic.BlipFill.Blip.Embed]
				}
				rot := float64(pic.SpPr.Xfrm.Rot) / 60000.0
				shape := reverse.ComShape{
					Type:      "Image",
					X:         emuToPt(pic.SpPr.Xfrm.Off.X),
					Y:         emuToPt(pic.SpPr.Xfrm.Off.Y),
					W:         emuToPt(pic.SpPr.Xfrm.Ext.Cx),
					H:         emuToPt(pic.SpPr.Xfrm.Ext.Cy),
					Rotation:  rot,
					ImagePath: imgPath,
				}
				allShapes = append(allShapes, shape)
			}

			// Top level Frames (Tables)
			for _, f := range tree.Frames {
				for _, tr := range f.Graphic.GraphicData.Tbl.Tr {
					for _, tc := range tr.Tc {
						hasText, textRuns := extractTextWithStyle(tc.TxBody)
						if hasText {
							shape := reverse.ComShape{
								Type:     "TableText",
								HasText:  true,
								TextRuns: textRuns,
							}
							allShapes = append(allShapes, shape)
						}
					}
				}
			}

			// Nested groups
			for _, grp := range tree.GrpSp {
				traverseGrpSp(grp, &allShapes, relMap)
			}

			slide.Shapes = allShapes
			doc.Slides = append(doc.Slides, slide)
		}
	}

	return doc, nil
}

func readZipFile(f *zip.File) ([]byte, error) {
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return ioutil.ReadAll(rc)
}
