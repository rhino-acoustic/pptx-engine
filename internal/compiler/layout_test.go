package compiler

import (
	"math"
	"testing"
)

func TestSnapRowCenters(t *testing.T) {
	// Simulate elements slightly misaligned on the Y axis
	// Expected center: avg of their mid-points
	nodes := []PptxElement{
		{Type: "text", X: 1.0, Y: 2.0, W: 1.0, H: 0.2}, // Mid Y: 2.1
		{Type: "box",  X: 2.0, Y: 2.05, W: 0.5, H: 0.3}, // Mid Y: 2.2
		{Type: "text", X: 3.0, Y: 2.1, W: 1.0, H: 0.2}, // Mid Y: 2.2
	}
	// Avg Mid Y = (2.1 + 2.2 + 2.2) / 3 = 2.1666...
	// New Y for node 0: 2.166 - 0.1 = 2.066
	// New Y for node 1: 2.166 - 0.15 = 2.016
	// New Y for node 2: 2.166 - 0.1 = 2.066

	snapped := SnapRowCenters(nodes, 0.15) // Tolerance 0.15

	if len(snapped) != 3 {
		t.Fatalf("expected 3 nodes, got %d", len(snapped))
	}

	// Calculate expected average mid Y
	avgMid := (2.1 + 2.2 + 2.2) / 3.0

	for i, node := range snapped {
		expectedY := avgMid - (node.H / 2.0)
		if math.Abs(node.Y - expectedY) > 0.001 {
			t.Errorf("Node %d: expected Y %v, got %v", i, expectedY, node.Y)
		}
	}
}

func TestClipToContainer(t *testing.T) {
	nodes := []PptxElement{
		{X: 1.0, Y: 1.0, W: 5.0, H: 1.0}, // Exceeds container right (3.0) -> should clip
		{X: 1.0, Y: 2.0, W: 1.0, H: 1.0}, // Safe -> should not clip
	}
	
	limits := map[int]float64{
		0: 3.0, 
		1: 5.0,
	}

	ClipToContainer(nodes, limits, 10.0)

	// Node 0 expected W = 3.0 - 1.0 - 0.05 = 1.95
	if math.Abs(nodes[0].W - 1.95) > 0.001 {
		t.Errorf("Node 0: expected W 1.95, got %v", nodes[0].W)
	}

	// Node 1 expected W = 1.0 (unchanged)
	if nodes[1].W != 1.0 {
		t.Errorf("Node 1: expected W 1.0, got %v", nodes[1].W)
	}
}
