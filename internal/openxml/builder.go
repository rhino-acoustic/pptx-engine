package openxml

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/rhino-acoustic/pptx-engine/internal/compiler"
	"github.com/rhino-acoustic/pptx-engine/internal/mapper"
)

// BuildPPTX takes a list of AST elements, an existing template (stub) PPTX, and writes a new PPTX.
// This completely washes away the Node.js/PptxGenJS runtime dependency.
func BuildPPTX(elements []compiler.PptxElement, stubFile string, outputFile string) error {
	// Open the stub file
	stub, err := zip.OpenReader(stubFile)
	if err != nil {
		return fmt.Errorf("failed to open stub pptx: %v", err)
	}
	defer stub.Close()

	// Create the new output file
	outF, err := os.Create(outputFile)
	if err != nil {
		return fmt.Errorf("failed to create output pptx: %v", err)
	}
	defer outF.Close()

	zipWriter := zip.NewWriter(outF)
	defer zipWriter.Close()

	for _, f := range stub.File {
		if strings.Contains(f.Name, "ppt/slides/slide") {
			// This is a slide file. We ONLY inject our own as slide1.xml, and IGNORE all other slides 
			// from the stub to prevent legacy data leakage.
			if strings.HasSuffix(f.Name, "slide1.xml") {
				if err := writeCustomSlide(zipWriter, f.Name, elements); err != nil {
					return err
				}
			}
		} else {
			// Copy all non-slide files, but fix presentation.xml slide size to 16:9
			if f.Name == "ppt/presentation.xml" {
				rc, err := f.Open()
				if err != nil { return err }
				data, _ := io.ReadAll(rc)
				rc.Close()
				// Force 16:9 widescreen (13.333" × 7.5")
				content := strings.Replace(string(data),
					`cx="9144000" cy="5143500"`,
					`cx="12192000" cy="6858000"`, 1)
				w, _ := zipWriter.Create(f.Name)
				w.Write([]byte(content))
			} else if err := copyZipFile(f, zipWriter); err != nil {
				return err
			}
		}
	}

	return nil
}

func copyZipFile(f *zip.File, zw *zip.Writer) error {
	src, err := f.Open()
	if err != nil {
		return err
	}
	defer src.Close()

	dst, err := zw.Create(f.Name)
	if err != nil {
		return err
	}

	_, err = io.Copy(dst, src)
	return err
}

func writeCustomSlide(zw *zip.Writer, filename string, elements []compiler.PptxElement) error {
	// Basic XML preamble for a slide
	slideXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
`

	// Inject elements
	idCounter := 2
	for _, el := range elements {
		emuX := int64(el.X * 914400)
		emuY := int64(el.Y * 914400)
		emuW := int64(el.W * 914400)
		emuH := int64(el.H * 914400)

		rotStr := ""
		if el.Rotation != 0 {
			rotStr = fmt.Sprintf(` rot="%d"`, int64(el.Rotation*60000))
		}

		if el.Type == "shape" || el.Type == "image" {
			// Fill XML — gradient or solid
			fillXML := ""
			if el.Gradient != nil && len(el.Gradient.Stops) >= 2 {
				gsLst := ""
				ooxAngle := ((360 - el.Gradient.Angle + 90) % 360) * 60000
				for _, stop := range el.Gradient.Stops {
					alphaStr := ""
					if stop.Alpha < 100000 {
						alphaStr = fmt.Sprintf(`<a:alpha val="%d"/>`, stop.Alpha)
					}
					gsLst += fmt.Sprintf(`<a:gs pos="%d"><a:srgbClr val="%s">%s</a:srgbClr></a:gs>`, stop.Position, stop.Color, alphaStr)
				}
				fillXML = fmt.Sprintf(`<a:gradFill><a:gsLst>%s</a:gsLst><a:lin ang="%d" scaled="1"/></a:gradFill>`, gsLst, int(ooxAngle))
			} else {
				fillColor := "EEEEEE"
				transparencyStr := ""
				if el.Fill != nil {
					if c, ok := el.Fill["color"].(string); ok && len(c) == 6 {
						fillColor = c
					}
					if t, ok := el.Fill["transparency"].(float64); ok && t > 0 {
						alpha := int64((1.0 - t) * 100000)
						transparencyStr = fmt.Sprintf(`<a:alpha val="%d"/>`, alpha)
					}
				}
				fillXML = fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s">%s</a:srgbClr></a:solidFill>`, fillColor, transparencyStr)
			}

			shapeType := "rect"
			avLstXML := "<a:avLst/>"
			if el.Shape.ShapeType == "ellipse" {
				shapeType = "ellipse"
			} else if el.Shape.ShapeType == "roundRect" {
				shapeType = "roundRect"
				// border-radius → avLst val (50000 = 50% = max round)
				// PPTX roundRect adj val: ratio of min(w,h) * 100000
				if el.Shape.RectRadius > 0 {
					minDim := el.W
					if el.H < minDim { minDim = el.H }
					if minDim > 0 {
						adjVal := int64(el.Shape.RectRadius / minDim * 100000)
						if adjVal > 50000 { adjVal = 50000 }
						avLstXML = fmt.Sprintf(`<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>`, adjVal)
					}
				}
			}

			// Border (line) XML
			lineXML := ""
			if el.Border != nil && el.Border.Width > 0 && el.Border.Color != "" {
				bwEMU := int64(el.Border.Width * 12700) // pt to EMU
				lineXML = fmt.Sprintf(`<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>`, bwEMU, el.Border.Color)
			}

			// Shadow XML
			shapeShadowXML := ""
			if el.Shadow != nil {
				blurEMU := el.Shadow.Blur * 12700
				offEMU := el.Shadow.Offset * 12700
				alphaVal := int(el.Shadow.Opacity * 100000)
				shapeShadowXML = fmt.Sprintf(`<a:effectLst><a:outerShdw blurRad="%d" dist="%d" dir="%d" algn="tl" rotWithShape="0"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>`, blurEMU, offEMU, el.Shadow.Direction, el.Shadow.Color, alphaVal)
			}

			shapeXML := fmt.Sprintf(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="Shape %d"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="%s">%s</a:prstGeom>
          %s
          %s
          %s
        </p:spPr>
      </p:sp>`, idCounter, idCounter, rotStr, emuX, emuY, emuW, emuH, shapeType, avLstXML, fillXML, lineXML, shapeShadowXML)
			slideXML += shapeXML

		} else if el.Type == "text" {
			textColor := "000000"
			fontSize := 14.0
			boldStr := `b="0"`
			italicStr := `i="0"`
			pptxAlign := "l" // default left

			if el.TextConfig != nil {
				if len(el.TextConfig.Color) == 6 {
					textColor = el.TextConfig.Color
				}
				if el.TextConfig.FontSize > 0 {
					fontSize = el.TextConfig.FontSize
				}
				if el.TextConfig.Bold {
					boldStr = `b="1"`
				}
				if el.TextConfig.Italic {
					italicStr = `i="1"`
				}
				// Map CSS align to PPTX align
				switch el.TextConfig.Align {
				case "center": pptxAlign = "ctr"
				case "right": pptxAlign = "r"
				case "justify": pptxAlign = "just"
				default: pptxAlign = "l"
				}
			}

			szVal := int(fontSize * 100)
			langAttr := `lang="en-US"`
			if containsKorean(el.Text) {
				langAttr = `lang="ko-KR"`
			}

			// Build paragraphs — split on \n for line breaks
			lines := strings.Split(el.Text, "\n")
			// Text color alpha
			textAlphaXML := ""
			if el.TextConfig != nil && el.TextConfig.ColorAlpha > 0 {
				alphaVal := int64((1.0 - float64(el.TextConfig.ColorAlpha)/100.0) * 100000)
				textAlphaXML = fmt.Sprintf(`<a:alpha val="%d"/>`, alphaVal)
			}

			fillXMLStr := ""
			if textAlphaXML != "" {
				fillXMLStr = fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s">%s</a:srgbClr></a:solidFill>`, textColor, textAlphaXML)
			} else {
				fillXMLStr = fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, textColor)
			}
			fontFace := "Arial"
			if el.TextConfig != nil && el.TextConfig.FontFace != "" {
				fontFace = el.TextConfig.FontFace
			}

			parasXML := ""
			if el.TextConfig != nil && len(el.TextConfig.TextRuns) > 0 {
				parasXML += fmt.Sprintf(`
            <a:p>
              <a:pPr algn="%s"/>`, pptxAlign)
				for _, run := range el.TextConfig.TextRuns {
					escaped := escapeXML(run.Text)
					if escaped == "" { continue }
					rSzVal := szVal
					if run.FontSize > 0 { rSzVal = int(run.FontSize * 100) }
					rBoldStr := boldStr
					if run.Bold { rBoldStr = `b="1"` }
					rItalicStr := italicStr
					if run.Italic { rItalicStr = `i="1"` }
					rFontFace := fontFace
					if run.FontFace != "" { rFontFace = run.FontFace }
					rFillXMLStr := fillXMLStr
					if run.Color != "" { rFillXMLStr = fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, run.Color) }
					
					parasXML += fmt.Sprintf(`
              <a:r>
                <a:rPr %s dirty="0" smtClean="0" sz="%d" %s %s>
                  %s
                  <a:latin typeface="%s"/>
                  <a:ea typeface="%s"/>
                </a:rPr>
                <a:t>%s</a:t>
              </a:r>`, langAttr, rSzVal, rBoldStr, rItalicStr, rFillXMLStr, rFontFace, rFontFace, escaped)
				}
				parasXML += `
            </a:p>`
			} else {
				for _, line := range lines {
					escaped := escapeXML(strings.TrimSpace(line))
					if escaped == "" { continue }
					parasXML += fmt.Sprintf(`
            <a:p>
              <a:pPr algn="%s"/>
              <a:r>
                <a:rPr %s dirty="0" smtClean="0" sz="%d" %s %s>
                  %s
                  <a:latin typeface="%s"/>
                  <a:ea typeface="%s"/>
                </a:rPr>
                <a:t>%s</a:t>
              </a:r>
            </a:p>`, pptxAlign, langAttr, szVal, boldStr, italicStr, fillXMLStr, fontFace, fontFace, escaped)
				}
			}

			// Anchor mapping from TextConfig.Valign
			anchorStr := "t" // default top
			if el.TextConfig != nil {
				switch el.TextConfig.Valign {
				case "middle": anchorStr = "ctr"
				case "bottom": anchorStr = "b"
				default: anchorStr = "t"
				}
			}

			// Shadow XML for text shapes
			shadowXML := ""
			if el.Shadow != nil {
				blurEMU := el.Shadow.Blur * 12700
				offEMU := el.Shadow.Offset * 12700
				alphaVal := int(el.Shadow.Opacity * 100000)
				shadowXML = fmt.Sprintf(`<a:effectLst><a:outerShdw blurRad="%d" dist="%d" dir="%d" algn="tl" rotWithShape="0"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>`, blurEMU, offEMU, el.Shadow.Direction, el.Shadow.Color, alphaVal)
			}

			// Wrap mapping
			wrapAttr := "square"
			if el.TextConfig != nil && !el.TextConfig.Wrap {
				wrapAttr = "none"
			}

			// Box background and border
			boxFillXML := "<a:noFill/>"
			if el.Fill != nil {
				if c, ok := el.Fill["color"].(string); ok && len(c) == 6 {
					transparencyStr := ""
					if t, ok := el.Fill["transparency"].(float64); ok && t > 0 {
						alpha := int64((1.0 - t) * 100000)
						transparencyStr = fmt.Sprintf(`<a:alpha val="%d"/>`, alpha)
					}
					boxFillXML = fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s">%s</a:srgbClr></a:solidFill>`, c, transparencyStr)
				}
			}

			shapeType := "rect"
			avLstXML := "<a:avLst/>"
			if el.Shape.ShapeType == "roundRect" {
				shapeType = "roundRect"
				if el.Shape.RectRadius > 0 {
					minDim := el.W
					if el.H < minDim { minDim = el.H }
					if minDim > 0 {
						adjVal := int64(el.Shape.RectRadius / minDim * 100000)
						if adjVal > 50000 { adjVal = 50000 }
						avLstXML = fmt.Sprintf(`<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>`, adjVal)
					}
				}
			}

			borderXML := ""
			if el.Border != nil && el.Border.Width > 0 && el.Border.Color != "" {
				bwEMU := int64(el.Border.Width * 12700)
				borderXML = fmt.Sprintf(`<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>`, bwEMU, el.Border.Color)
			}

			textXML := fmt.Sprintf(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="Text %d"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="%s">%s</a:prstGeom>
          %s
          %s
          %s
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="%s" rtlCol="0" anchor="%s" lIns="0" tIns="0" rIns="0" bIns="0">
            <a:normAutofit/>
          </a:bodyPr>
          <a:lstStyle/>%s
        </p:txBody>
      </p:sp>`, idCounter, idCounter, rotStr, emuX, emuY, emuW, emuH, shapeType, avLstXML, boxFillXML, borderXML, shadowXML, wrapAttr, anchorStr, parasXML)
			slideXML += textXML
		}
		idCounter++
	}

	slideXML += `
    </p:spTree>
    <p:extLst>
      <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
        <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="3201416410"/>
      </p:ext>
    </p:extLst>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`

	dst, err := zw.Create(filename)
	if err != nil {
		return err
	}

	_, err = io.Copy(dst, bytes.NewBufferString(slideXML))
	return err
}

// BuildPPTXSlide writes a single slide to an existing zip writer
func BuildPPTXSlide(zw *zip.Writer, elements []compiler.PptxElement, filename string) error {
	slideXML := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
`
	idCounter := 2
	for _, el := range elements {
		emuX := int64(el.X * 914400)
		emuY := int64(el.Y * 914400)
		emuW := int64(el.W * 914400)
		emuH := int64(el.H * 914400)

		rotStr := ""
		if el.Rotation != 0 {
			rotStr = fmt.Sprintf(` rot="%d"`, int64(el.Rotation*60000))
		}

		if el.Type == "image" && el.ImageRId != "" {
			// Render as embedded picture
			picXML := fmt.Sprintf(`
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="%d" name="Image %d"/>
          <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="%s"/>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>`, idCounter, idCounter, el.ImageRId, rotStr, emuX, emuY, emuW, emuH)
			slideXML += picXML

		} else if el.Type == "shape" || el.Type == "image" {
			// Fill XML — gradient or solid
			fillXML := ""
			if el.Gradient != nil && len(el.Gradient.Stops) >= 2 {
				// A등급: 네이티브 <a:gradFill>
				// CSS angle → OpenXML angle: (360 - cssAngle + 90) % 360, then × 60000
				ooxAngle := ((360 - el.Gradient.Angle + 90) % 360) * 60000
				gsLst := ""
				for _, stop := range el.Gradient.Stops {
					gsLst += fmt.Sprintf(`<a:gs pos="%d"><a:srgbClr val="%s"/></a:gs>`, stop.Position, stop.Color)
				}
				fillXML = fmt.Sprintf(`<a:gradFill><a:gsLst>%s</a:gsLst><a:lin ang="%d" scaled="1"/></a:gradFill>`, gsLst, ooxAngle)
			} else {
				// 폴백: solid
				fillColor := "EEEEEE"
				transparencyStr := ""
				if el.Fill != nil {
					if c, ok := el.Fill["color"].(string); ok && len(c) == 6 {
						fillColor = c
					}
					if t, ok := el.Fill["transparency"].(float64); ok && t > 0 {
						alpha := int64((1.0 - t) * 100000)
						transparencyStr = fmt.Sprintf(`<a:alpha val="%d"/>`, alpha)
					}
				}
				fillXML = fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s">%s</a:srgbClr></a:solidFill>`, fillColor, transparencyStr)
			}

			shapeType := el.Shape.ShapeType
			if shapeType == "" {
				shapeType = "rect"
			}

			// roundRect avLst: 반지름 값 (50000 = 50% = pill shape)
			avLstXML := "<a:avLst/>"
			if shapeType == "roundRect" && el.Shape.RectRadius > 0 {
				avVal := int(el.Shape.RectRadius * 100000) // OpenXML: 1/100000 of shape
				avLstXML = fmt.Sprintf(`<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>`, avVal)
			}

			// Border XML
			borderXML := ""
			if el.Border != nil && el.Border.Color != "" {
				bw := int(el.Border.Width * 12700) // pt → EMU
				dashAttr := ""
				if el.Border.DashType != "" {
					dashAttr = fmt.Sprintf(` prstDash="%s"`, el.Border.DashType)
				}
				borderXML = fmt.Sprintf(`<a:ln w="%d"%s><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>`, bw, dashAttr, el.Border.Color)
			}

			// Shadow XML
			shadowXML := ""
			if el.Shadow != nil {
				blurEMU := el.Shadow.Blur * 12700
				offEMU := el.Shadow.Offset * 12700
				alphaVal := int(el.Shadow.Opacity * 100000)
				shadowXML = fmt.Sprintf(`<a:effectLst><a:outerShdw blurRad="%d" dist="%d" dir="%d"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>`, blurEMU, offEMU, el.Shadow.Direction, el.Shadow.Color, alphaVal)
			}

			// Flip attributes
			flipStr := ""
			if el.Shape.FlipH {
				flipStr += ` flipH="1"`
			}
			if el.Shape.FlipV {
				flipStr += ` flipV="1"`
			}

			shapeXML := fmt.Sprintf(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="Shape %d"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="%s">%s</a:prstGeom>
          %s
          %s%s
        </p:spPr>
        <p:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef></p:style>
      </p:sp>`, idCounter, idCounter, rotStr, flipStr, emuX, emuY, emuW, emuH, shapeType, avLstXML, fillXML, borderXML, shadowXML)
			slideXML += shapeXML

		} else if el.Type == "text" {
			textColor := "000000"
			fontSize := 14.0
			boldStr := `b="0"`
			italicStr := `i="0"`
			fontFace := "Arial"
			wrapAttr := `wrap="square"`
			lineSpacingXML := ""

			if el.TextConfig != nil {
				if len(el.TextConfig.Color) == 6 {
					textColor = el.TextConfig.Color
				}
				if el.TextConfig.FontSize > 0 {
					fontSize = el.TextConfig.FontSize
				}
				if el.TextConfig.Bold {
					boldStr = `b="1"`
				}
				if el.TextConfig.Italic {
					italicStr = `i="1"`
				}
				if el.TextConfig.FontFace != "" {
					fontFace = el.TextConfig.FontFace
				}
				if !el.TextConfig.Wrap {
					wrapAttr = `wrap="none"`
				}
				// LineSpacing → <a:lnSpc><a:spcPct val="120000"/></a:lnSpc>
				if el.TextConfig.LineHeight > 0 {
					lhVal := int(el.TextConfig.LineHeight * 1000) // PPTX: thousandths of percent
					lineSpacingXML = fmt.Sprintf(`<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>`, lhVal)
				}
			}

			szVal := int(fontSize * 100)
			escaped := strings.ReplaceAll(el.Text, "&", "&amp;")
			escaped = strings.ReplaceAll(escaped, "<", "&lt;")
			escaped = strings.ReplaceAll(escaped, ">", "&gt;")

			// Detect lang for CJK
			langAttr := `lang="en-US"`
			for _, ch := range el.Text {
				if ch > 0x2E80 {
					langAttr = `lang="ko-KR"`
					break
				}
			}

			parasXML := ""
			if el.TextConfig != nil && len(el.TextConfig.TextRuns) > 0 {
				parasXML += fmt.Sprintf(`
          <a:p>
            <a:pPr>%s</a:pPr>`, lineSpacingXML)
				for _, run := range el.TextConfig.TextRuns {
					rEscaped := xmlEscape(run.Text)
					if rEscaped == "" { continue }
					rSzVal := szVal
					if run.FontSize > 0 { rSzVal = int(run.FontSize * 100) }
					rBoldStr := boldStr
					if run.Bold { rBoldStr = `b="1"` }
					rItalicStr := italicStr
					if run.Italic { rItalicStr = `i="1"` }
					rFontFace := fontFace
					if run.FontFace != "" { rFontFace = run.FontFace }
					rTextColor := textColor
					if run.Color != "" { rTextColor = run.Color }
					
					parasXML += fmt.Sprintf(`
            <a:r>
              <a:rPr %s dirty="0" sz="%d" %s %s>
                <a:solidFill><a:srgbClr val="%s"/></a:solidFill>
                <a:latin typeface="%s"/>
                <a:ea typeface="%s"/>
              </a:rPr>
              <a:t>%s</a:t>
            </a:r>`, langAttr, rSzVal, rBoldStr, rItalicStr, rTextColor, rFontFace, rFontFace, rEscaped)
				}
				parasXML += `
          </a:p>`
			} else {
				parasXML = fmt.Sprintf(`
          <a:p>
            <a:pPr>%s</a:pPr>
            <a:r>
              <a:rPr %s dirty="0" sz="%d" %s %s>
                <a:solidFill><a:srgbClr val="%s"/></a:solidFill>
                <a:latin typeface="%s"/>
                <a:ea typeface="%s"/>
              </a:rPr>
              <a:t>%s</a:t>
            </a:r>
          </a:p>`, lineSpacingXML, langAttr, szVal, boldStr, italicStr, textColor, fontFace, fontFace, escaped)
			}

			// Box background and border
			boxFillXML := "<a:noFill/>"
			if el.Fill != nil {
				if c, ok := el.Fill["color"].(string); ok && len(c) == 6 {
					transparencyStr := ""
					if t, ok := el.Fill["transparency"].(float64); ok && t > 0 {
						alpha := int64((1.0 - t) * 100000)
						transparencyStr = fmt.Sprintf(`<a:alpha val="%d"/>`, alpha)
					}
					boxFillXML = fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s">%s</a:srgbClr></a:solidFill>`, c, transparencyStr)
				}
			}

			shapeType := el.Shape.ShapeType
			if shapeType == "" { shapeType = "rect" }
			avLstXML := "<a:avLst/>"
			if shapeType == "roundRect" && el.Shape.RectRadius > 0 {
				avVal := int(el.Shape.RectRadius * 100000)
				avLstXML = fmt.Sprintf(`<a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst>`, avVal)
			}

			borderXML := ""
			if el.Border != nil && el.Border.Color != "" {
				bw := int(el.Border.Width * 12700)
				dashAttr := ""
				if el.Border.DashType != "" {
					dashAttr = fmt.Sprintf(` prstDash="%s"`, el.Border.DashType)
				}
				borderXML = fmt.Sprintf(`<a:ln w="%d"%s><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>`, bw, dashAttr, el.Border.Color)
			}

			// Shadow XML for text shapes
			shadowXML := ""
			if el.Shadow != nil {
				blurEMU := el.Shadow.Blur * 12700
				offEMU := el.Shadow.Offset * 12700
				alphaVal := int(el.Shadow.Opacity * 100000)
				shadowXML = fmt.Sprintf(`<a:effectLst><a:outerShdw blurRad="%d" dist="%d" dir="%d"><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>`, blurEMU, offEMU, el.Shadow.Direction, el.Shadow.Color, alphaVal)
			}

			textXML := fmt.Sprintf(`
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="%d" name="Text %d"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm%s>
            <a:off x="%d" y="%d"/>
            <a:ext cx="%d" cy="%d"/>
          </a:xfrm>
          <a:prstGeom prst="%s">%s</a:prstGeom>
          %s
          %s
          %s
        </p:spPr>
        <p:txBody>
          <a:bodyPr %s rtlCol="0" anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0"/>
          <a:lstStyle/>%s
        </p:txBody>
      </p:sp>`, idCounter, idCounter, rotStr, emuX, emuY, emuW, emuH, shapeType, avLstXML, boxFillXML, borderXML, shadowXML, wrapAttr, parasXML)
			slideXML += textXML

		} else if el.Type == "table" && el.Table != nil && len(el.Table.Rows) > 0 {
			// Table rendering: <p:graphicFrame> with <a:tbl>
			numCols := 0
			for _, row := range el.Table.Rows {
				if len(row) > numCols {
					numCols = len(row)
				}
			}
			if numCols == 0 {
				numCols = 1
			}

			// Column widths: equal distribution if not specified
			colWidthEMU := emuW / int64(numCols)
			gridColXML := ""
			for c := 0; c < numCols; c++ {
				w := colWidthEMU
				if c < len(el.Table.ColWidths) && el.Table.ColWidths[c] > 0 {
					w = el.Table.ColWidths[c]
				}
				gridColXML += fmt.Sprintf(`<a:gridCol w="%d"/>`, w)
			}

			// Rows
			rowsXML := ""
			numRows := len(el.Table.Rows)
			rowHeightEMU := emuH / int64(numRows)
			for ri, row := range el.Table.Rows {
				rh := rowHeightEMU
				if ri < len(el.Table.RowHeights) && el.Table.RowHeights[ri] > 0 {
					rh = el.Table.RowHeights[ri]
				}
				cellsXML := ""
				for ci := 0; ci < numCols; ci++ {
					cell := mapper.TableCell{Text: "", FillColor: "FFFFFF", Color: "000000", FontSize: 1200, Align: "l"}
					if ci < len(row) {
						cell = row[ci]
					}
					if cell.FontSize == 0 {
						cell.FontSize = 1200
					}
					if cell.Color == "" {
						cell.Color = "000000"
					}
					if cell.FillColor == "" {
						cell.FillColor = "FFFFFF"
					}
					if cell.Align == "" {
						cell.Align = "l"
					}
					// Header row: bold + dark background
					boldAttr := ""
					if cell.Bold || (el.Table.HasHeader && ri == 0) {
						boldAttr = ` b="1"`
					}
					// gridSpan for colSpan
					gridSpanAttr := ""
					if cell.ColSpan > 1 {
						gridSpanAttr = fmt.Sprintf(` gridSpan="%d"`, cell.ColSpan)
					}

					cellFontFace := "Arial"
					if cell.FontFace != "" {
						cellFontFace = cell.FontFace
					}
					escaped := xmlEscape(cell.Text)
					cellsXML += fmt.Sprintf(`<a:tc%s>
  <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="%d"%s dirty="0"><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:latin typeface="%s"/><a:ea typeface="%s"/></a:rPr><a:t>%s</a:t></a:r></a:p></a:txBody>
  <a:tcPr anchor="ctr"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:tcPr>
</a:tc>`, gridSpanAttr, cell.FontSize, boldAttr, cell.Color, cellFontFace, cellFontFace, escaped, cell.FillColor)
				}
				rowsXML += fmt.Sprintf(`<a:tr h="%d">%s</a:tr>`, rh, cellsXML)
			}

			tableXML := fmt.Sprintf(`
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="%d" name="Table %d"/>
          <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="%d" y="%d"/>
          <a:ext cx="%d" cy="%d"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr firstRow="%d" bandRow="1"/>
              <a:tblGrid>%s</a:tblGrid>
              %s
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>`, idCounter, idCounter, emuX, emuY, emuW, emuH,
				boolToInt(el.Table.HasHeader), gridColXML, rowsXML)
			slideXML += tableXML
		}
		idCounter++
	}

	slideXML += `
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`

	dst, err := zw.Create(filename)
	if err != nil {
		return err
	}
	_, err = io.Copy(dst, bytes.NewBufferString(slideXML))
	return err
}

// xmlEscape escapes XML special characters in text content.
func xmlEscape(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	return s
}

// boolToInt converts bool to 0/1 for XML attributes.
func boolToInt(b bool) int {
	if b {
		return 1
	}
	return 0
}

func containsKorean(s string) bool {
	for _, r := range s {
		if r >= 0xAC00 && r <= 0xD7AF { return true }
		if r >= 0x1100 && r <= 0x11FF { return true }
	}
	return false
}

func escapeXML(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	return s
}
