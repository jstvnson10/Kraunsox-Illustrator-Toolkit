#target illustrator

// ==========================================
// KRAUNSOX HALFTONE TOOLKIT
// Halftone Grid v0.1
// ==========================================

if (app.documents.length === 0) {

    alert("Open an Illustrator document first.");

} else {

    var doc = app.activeDocument;

    // --------------------------------------
    // SETTINGS
    // --------------------------------------

    var rows = 30;
    var columns = 30;

    var dotSize = 5;
    var spacing = 10;

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

    // Start 50pt inside the artboard
    var startX = left + 80;
    var startY = top - 80;

    // --------------------------------------
    // HALFTONE LAYER
    // --------------------------------------

    var halftoneLayer = doc.layers.add();

    halftoneLayer.name = "KRAUNSOX HALFTONE";

    // --------------------------------------
    // GENERATE DOTS
    // --------------------------------------

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

    app.redraw();

    alert(
        "Halftone grid created!\n\n" +
        "Dots: " + (rows * columns)
    );
}
