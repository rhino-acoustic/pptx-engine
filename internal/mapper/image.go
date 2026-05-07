package mapper

// ImageSizingConfig represents the scaling and cropping options for images in PPTX.
type ImageSizingConfig struct {
	Type string  `json:"type,omitempty"` // "cover", "contain", "crop"
	W    float64 `json:"w,omitempty"`
	H    float64 `json:"h,omitempty"`
}

// ImageConfig represents the mapped PPTX image attributes.
type ImageConfig struct {
	Path   string            `json:"path"`
	Sizing ImageSizingConfig `json:"sizing,omitempty"`
	Shape  ShapeConfig       `json:"shape,omitempty"` // For native picture fill with rounding
}

// MapImageFill maps CSS object-fit or background-size to PPTX image sizing.
func MapImageFill(url, objectFit string, targetW, targetH float64, shapeCfg ShapeConfig) ImageConfig {
	cfg := ImageConfig{
		Path:  url,
		Shape: shapeCfg,
	}

	// Default behavior is to stretch (no specific sizing type), but if object-fit is used:
	if objectFit == "cover" || objectFit == "contain" {
		cfg.Sizing = ImageSizingConfig{
			Type: objectFit,
			W:    targetW,
			H:    targetH,
		}
	}

	return cfg
}
