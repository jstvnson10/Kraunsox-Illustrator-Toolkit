#target illustrator

// ==========================================
// KRAUNSOX HALFTONE TOOLKIT
// Halftone Tone Curve v0.3
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

    // TONE CONTROL
    var gamma = Number(
    prompt(
        "Enter halftone gamma:\n\n" +
        "0.25 = very heeavy\n" +
        "0.50 = heavy\n" +
        "0.75 = slightly heavy\n" +
        "1.00 = linear\n" + 
        "1.50 = lighter" +
        "2.00 = light\n" +
        "3.00 = heavy\n",
        "1.0"
    )
);

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

    halftoneLayer.name =
        "KRAUNSOX HALFTONE - GAMMA " + gamma;

    // --------------------------------------
    // GENERATE HALFTONE
    // --------------------------------------

    for (var row = 0; row < rows; row++) {

        for (var column = 0; column < columns; column++) {

            // Normalize horizontal position
            // Result: 0.0 -> 1.0
            var percentage =
                column / (columns - 1);

            // Apply tone curve
            var tone =
                Math.pow(percentage, gamma);

            // Convert tone into dot diameter
            var dotSize =
                minDotSize +
                tone * (maxDotSize - minDotSize);

            // Dot center
            var centerX =
                startX + (column * spacing);

            var centerY =
                startY - (row * spacing);

            // Compensate for changing dot diameter
            var x =
                centerX - (dotSize / 2);

            var y =
                centerY + (dotSize / 2);

            // Create vector circle
            var dot =
                halftoneLayer.pathItems.ellipse(
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
        "Tone curve halftone created!\n\n" +
        "Gamma: " + gamma + "\n" +
        "Dots: " + (rows * columns)
    );
}
