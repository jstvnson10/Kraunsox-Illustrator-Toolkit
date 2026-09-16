#target illustrator

// ============================================
// KRAUNSOX HALFTONE TOOLKIT
// Lesson 1: Vector Dot Grid
// ============================================

if (app.documents.length === 0) {
    alert("Open an Illustrator document first.");
} else {
    var doc = app.activeDocument;

    // ----- SETTINGS -----
    var rows = 20;
    var columns = 20;

    // Illustrator scripting uses points
    var spacing = 18;
    var dotSize = 8;

    // Starting location
    var startX = 100;
    var startY = 600;

    // Create a layer for our halftone
    var halftoneLayer = doc.layers.add();
    halftoneLayer.name = "KRAUNSOX HALFTONE";

    // Create black
    var black = new RGBColor();
    black.red = 0;
    black.green = 0;
    black.blue = 0;

    // ----- BUILD DOT GRID -----

    for (var row = 0; row < rows; row++) {

        for (var column = 0; column < columns; column++) {

            var x = startX + (column * spacing);
            var y = startY - (row * spacing);

            var dot = halftoneLayer.pathItems.ellipse(
                y,
                x,
                dotSize,
                dotSize
            );

            dot.filled = true;
            dot.fillColor = black;

            dot.stroked = false;
        }
    }

    alert("KRAUNSOX Halftone generated.");
}
