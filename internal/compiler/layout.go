package compiler

import (
	"math"
	"sort"
)

// SortByZIndex sorts nodes by their ZIndex to ensure correct stacking context mapping to PPT rendering order.
func SortByZIndex(nodes []PptxElement) []PptxElement {
	sorted := make([]PptxElement, len(nodes))
	copy(sorted, nodes)
	sort.SliceStable(sorted, func(i, j int) bool {
		return sorted[i].ZIndex < sorted[j].ZIndex
	})
	return sorted
}

// SnapRowCenters groups elements horizontally and aligns their Y-centers.
func SnapRowCenters(nodes []PptxElement, tolerance float64) []PptxElement {
	if len(nodes) == 0 {
		return nodes
	}

	// Copy slice to avoid mutating original order yet
	sorted := make([]PptxElement, len(nodes))
	copy(sorted, nodes)

	// Sort by Y first, then X
	sort.SliceStable(sorted, func(i, j int) bool {
		if math.Abs(sorted[i].Y-sorted[j].Y) > 0.01 {
			return sorted[i].Y < sorted[j].Y
		}
		return sorted[i].X < sorted[j].X
	})

	var rows [][]int
	var currentRow []int
	currentRow = append(currentRow, 0)

	for i := 1; i < len(sorted); i++ {
		if sorted[i].Y-sorted[currentRow[0]].Y < tolerance {
			currentRow = append(currentRow, i)
		} else {
			rows = append(rows, currentRow)
			currentRow = []int{i}
		}
	}
	if len(currentRow) > 0 {
		rows = append(rows, currentRow)
	}

	for _, rowIdxs := range rows {
		if len(rowIdxs) < 2 {
			continue
		}

		var sumMidY float64
		for _, idx := range rowIdxs {
			sumMidY += sorted[idx].Y + (sorted[idx].H / 2)
		}
		avgMidY := sumMidY / float64(len(rowIdxs))

		for _, idx := range rowIdxs {
			sorted[idx].Y = avgMidY - (sorted[idx].H / 2)
		}
	}

	return sorted
}

// ClipToContainer restricts the width (W) of an element if it exceeds its parent container boundaries.
func ClipToContainer(nodes []PptxElement, containerRights map[int]float64, slideWidth float64) {
	for i := range nodes {
		rightLimit, exists := containerRights[i]
		if !exists {
			rightLimit = slideWidth
		}

		maxW := rightLimit - nodes[i].X - 0.05
		if maxW > 0.15 && nodes[i].W > maxW {
			nodes[i].W = maxW
		}
	}
}

// ApplyOverflowClipping simulates CSS overflow:hidden by clipping child elements that fall outside 
// parent bounding boxes. It assumes nodes are already sorted by ZIndex.
func ApplyOverflowClipping(nodes []PptxElement) {
	for i := 0; i < len(nodes); i++ {
		parent := nodes[i]
		if parent.Overflow != "hidden" || parent.Type != "shape" {
			continue
		}

		// Check subsequent elements (which have higher or equal Z-Index)
		for j := i + 1; j < len(nodes); j++ {
			child := &nodes[j]

			// Skip if totally outside the parent horizontally or vertically (basic AABB check)
			if child.X >= parent.X+parent.W || child.X+child.W <= parent.X ||
				child.Y >= parent.Y+parent.H || child.Y+child.H <= parent.Y {
				continue
			}

			// If intersecting, clip the child dimensions to the parent's boundaries
			if child.X < parent.X {
				diff := parent.X - child.X
				child.X = parent.X
				child.W -= diff
			}
			if child.Y < parent.Y {
				diff := parent.Y - child.Y
				child.Y = parent.Y
				child.H -= diff
			}
			if child.X+child.W > parent.X+parent.W {
				child.W = (parent.X + parent.W) - child.X
			}
			if child.Y+child.H > parent.Y+parent.H {
				child.H = (parent.Y + parent.H) - child.Y
			}
		}
	}
}
