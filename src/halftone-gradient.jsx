#target illustrator

// ==========================================
// KRAUNSOX HALFTONE TOOLKIT
// Halftone Gradient v0.2
// ==========================================

if (app.documents.length === 0) {

    alert("Open an Illustrator document first.");

} else {

    var doc = app.activeDocument;

    // --------------------------------------
    // SETTINGS
    // --------------------------------------

    var rows = 20;
    var columns = 30;

    var spacing = 18;

    var minDotSize = 1;
    var maxDotSize = 16;

    // --------------------------------------
    // COLOR
    // --------------------------------------

    var black = new RGBColor();

    black.red = 0;
    black.green = 0;
    black.blue = 0;

    // --------------------------------------
    // ACTIVE ARTBOARD
    // --------------------------------------

    var artboardIndex = doc.artboards.getActiveArtboardIndex();
    var artboard = doc.artboards[artboardIndex].artboardRect;

    var left = artboard[0];
    var top = artboard[1];

    var startX = left + 50;
    var startY = top - 50;

    // --------------------------------------
    // HALFTONE LAYER
    // --------------------------------------

    var halftoneLayer = doc.layers.add();
    halftoneLayer.name = "KRAUNSOX HALFTONE GRADIENT";

    // --------------------------------------
    // GENERATE HALFTONE
    // --------------------------------------

    for (var row = 0; row < rows; row++) {

        for (var column = 0; column < columns; column++) {

            // Convert column number to 0–1
            var percentage = column / (columns - 1);

            // Calculate dot size
            var dotSize =
                minDotSize +
                percentage * (maxDotSize - minDotSize);

            // Center position of this dot
            var centerX = startX + (column * spacing);
            var centerY = startY - (row * spacing);

            // Illustrator ellipse() positions using
            // top-left corner, not center.
            var x = centerX - (dotSize / 2);
            var y = centerY + (dotSize / 2);

            // Create vector dot
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

    app.redraw();

    alert(
        "Variable halftone created!\n\n" +
        "Dots: " + (rows * columns)
    );
}
