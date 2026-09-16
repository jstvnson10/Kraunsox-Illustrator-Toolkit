#target illustrator

// ==========================================
// KRAUNSOX HALFTONE TOOLKIT
// Screen Angle v0.4
// ==========================================

if (app.documents.length === 0) {

    alert("Open an Illustrator document first.");

} else {

    var doc = app.activeDocument;

    // --------------------------------------
    // SETTINGS
    // --------------------------------------

    var rows = 25;
    var columns = 35;
    var spacing = 18;

    var minDotSize = 1;
    var maxDotSize = 16;

    var gamma = Number(
        prompt(
            "Enter gamma (0.5 = heavy, 1.0 = normal, 2.0 = light):",
            "1.0"
        )
    );

    var angle = Number(
        prompt(
            "Enter screen angle:",
            "45"
        )
    );

    // --------------------------------------
    // ANGLE CONVERSION
    // --------------------------------------

    var radians = angle * Math.PI / 180;

    var cosAngle = Math.cos(radians);
    var sinAngle = Math.sin(radians);

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

    var artboardIndex =
        doc.artboards.getActiveArtboardIndex();

    var artboard =
        doc.artboards[artboardIndex].artboardRect;

    var left = artboard[0];
    var top = artboard[1];

    var startX = left + 150;
    var startY = top - 150;

    // --------------------------------------
    // LAYER
    // --------------------------------------

    var halftoneLayer = doc.layers.add();

    halftoneLayer.name =
        "HALFTONE " +
        angle +
        "deg GAMMA " +
        gamma;

    // --------------------------------------
    // GENERATE HALFTONE
    // --------------------------------------

    for (var row = 0; row < rows; row++) {

        for (var column = 0; column < columns; column++) {

            // Normalize tone from 0 to 1
            var percentage =
                column / (columns - 1);

            // Apply gamma curve
            var tone =
                Math.pow(percentage, gamma);

            // Calculate dot diameter
            var dotSize =
                minDotSize +
                tone *
                (maxDotSize - minDotSize);

            // --------------------------------
            // ORIGINAL GRID POSITION
            // --------------------------------

            var gridX =
                column * spacing;

            var gridY =
                row * spacing;

            // --------------------------------
            // ROTATE GRID POSITION
            // --------------------------------

            var rotatedX =
                gridX * cosAngle -
                gridY * sinAngle;

            var rotatedY =
                gridX * sinAngle +
                gridY * cosAngle;

            // --------------------------------
            // PLACE ON ARTBOARD
            // --------------------------------

            var centerX =
                startX + rotatedX;

            var centerY =
                startY - rotatedY;

            // ellipse() uses top-left positioning
            var x =
                centerX - (dotSize / 2);

            var y =
                centerY + (dotSize / 2);

            // --------------------------------
            // CREATE VECTOR DOT
            // --------------------------------

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
        "Halftone created!\n\n" +
        "Angle: " + angle + " degrees\n" +
        "Gamma: " + gamma + "\n" +
        "Dots: " + (rows * columns)
    );
}
