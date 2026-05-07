package reverse

import (
	"image"
	"image/color"
	_ "image/jpeg"
	"image/png"
	"os"
	"path/filepath"
	"strconv"
)

// applyClrChange preprocesses an image by replacing pixels matching fromHex color
// with the specified alpha value. Returns path to the new image, or empty string on failure.
// Cache: if output file already exists, skip processing.
func applyClrChange(imgPath string, fromHex string, toAlpha uint8) string {
	ext := filepath.Ext(imgPath)
	outPath := imgPath[:len(imgPath)-len(ext)] + "_clr" + ext

	// Cache check
	if _, err := os.Stat(outPath); err == nil {
		return outPath
	}

	// Parse target color
	if len(fromHex) != 6 {
		return ""
	}
	r, err1 := strconv.ParseUint(fromHex[0:2], 16, 8)
	g, err2 := strconv.ParseUint(fromHex[2:4], 16, 8)
	b, err3 := strconv.ParseUint(fromHex[4:6], 16, 8)
	if err1 != nil || err2 != nil || err3 != nil {
		return ""
	}
	targetR, targetG, targetB := uint8(r), uint8(g), uint8(b)

	// Open and decode image
	f, err := os.Open(imgPath)
	if err != nil {
		return ""
	}
	defer f.Close()

	img, _, err := image.Decode(f)
	if err != nil {
		return ""
	}

	bounds := img.Bounds()
	rgba := image.NewNRGBA(bounds)

	// Color distance tolerance (covers near-white variations)
	const tolerance = 30

	for y := bounds.Min.Y; y < bounds.Max.Y; y++ {
		for x := bounds.Min.X; x < bounds.Max.X; x++ {
			cr, cg, cb, ca := img.At(x, y).RGBA()
			pr := uint8(cr >> 8)
			pg := uint8(cg >> 8)
			pb := uint8(cb >> 8)
			pa := uint8(ca >> 8)

			// Check if pixel matches target color within tolerance
			dr := absUint8(pr, targetR)
			dg := absUint8(pg, targetG)
			db := absUint8(pb, targetB)
			if dr <= tolerance && dg <= tolerance && db <= tolerance {
				rgba.SetNRGBA(x, y, color.NRGBA{R: pr, G: pg, B: pb, A: toAlpha})
			} else {
				rgba.SetNRGBA(x, y, color.NRGBA{R: pr, G: pg, B: pb, A: pa})
			}
		}
	}

	// Save
	out, err := os.Create(outPath)
	if err != nil {
		return ""
	}
	defer out.Close()
	if err := png.Encode(out, rgba); err != nil {
		return ""
	}

	return outPath
}

func absUint8(a, b uint8) int {
	if a > b {
		return int(a - b)
	}
	return int(b - a)
}
