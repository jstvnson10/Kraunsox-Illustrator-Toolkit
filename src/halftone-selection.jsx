#target illustrator

// ==========================================
// KRAUNSOX HALFTONE TOOLKIT
// Selected Artwork v0.5
// ==========================================

if (app.documents.length === 0) {

    alert("Open an Illustrator document first.");

} else {

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {

        alert("Select one piece of artwork first.");

    } else {

        var artwork = doc.selection[0];

        // ----------------------------------
        // SETTINGS
        // ----------------------------------

        var spacing = 14;
        var minDotSize = 1;
        var maxDotSize = 12;

        var gamma = Number(
            prompt(
                "Enter gamma:",
                "1.0"
            )
        );

        var angle = Number(
            prompt(
                "Enter screen angle:",
                "45"
            )
        );

        // ----------------------------------
        // SELECTED ARTWORK BOUNDS
        // ----------------------------------

        var bounds = artwork.geometricBounds;

        var artLeft   = bounds[0];
        var artTop    = bounds[1];
        var artRight  = bounds[2];
        var artBottom = bounds[3];

        var artWidth =
            artRight - artLeft;

        var artHeight =
            artTop - artBottom;

        // ----------------------------------
        // CALCULATE GRID SIZE
        // ----------------------------------

        var columns =
            Math.ceil(artWidth / spacing) + 1;

        var rows =
            Math.ceil(artHeight / spacing) + 1;

        // ----------------------------------
        // ANGLE
        // ----------------------------------

        var radians =
            angle * Math.PI / 180;

        var cosAngle =
            Math.cos(radians);

        var sinAngle =
            Math.sin(radians);

        // ----------------------------------
        // COLOR
        // ----------------------------------

        var black = new RGBColor();

        black.red = 0;
        black.green = 0;
        black.blue = 0;

        // ----------------------------------
        // OUTPUT LAYER
        // ----------------------------------

        var halftoneLayer =
            doc.layers.add();

        halftoneLayer.name =
            "HALFTONE " +
            angle +
            "deg";

        // ----------------------------------
        // CENTER OF SELECTED ARTWORK
        // ----------------------------------

        var artCenterX =
            artLeft + (artWidth / 2);

        var artCenterY =
            artBottom + (artHeight / 2);

        // ----------------------------------
        // GENERATE DOTS
        // ----------------------------------

        for (
            var row = -rows;
            row <= rows;
            row++
        ) {

            for (
                var column = -columns;
                column <= columns;
                column++
            ) {

                // Normal grid position
                var gridX =
                    column * spacing;

                var gridY =
                    row * spacing;

                // Rotate coordinate
                var rotatedX =
                    gridX * cosAngle -
                    gridY * sinAngle;

                var rotatedY =
                    gridX * sinAngle +
                    gridY * cosAngle;

                // Position around artwork center
                var centerX =
                    artCenterX + rotatedX;

                var centerY =
                    artCenterY + rotatedY;

                // ----------------------------------
                // CHECK ARTWORK BOUNDS
                // ----------------------------------

                if (
                    centerX >= artLeft &&
                    centerX <= artRight &&
                    centerY <= artTop &&
                    centerY >= artBottom
                ) {

                    // TEMPORARY TONE
                    //
                    // We're still generating the
                    // tone mathematically.
                    //
                    // True artwork sampling comes next.

                    var percentage =
                        (centerX - artLeft) /
                        artWidth;

                    var tone =
                        Math.pow(
                            percentage,
                            gamma
                        );

                    var dotSize =
                        minDotSize +
                        tone *
                        (
                            maxDotSize -
                            minDotSize
                        );

                    // Center ellipse
                    var x =
                        centerX -
                        (dotSize / 2);

                    var y =
                        centerY +
                        (dotSize / 2);

                    // Create dot
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
        }

        app.redraw();

        alert(
            "Halftone field created!\n\n" +
            "Artwork width: " +
            Math.round(artWidth) +
            " pt\n" +

            "Artwork height: " +
            Math.round(artHeight) +
            " pt\n\n" +

            "Angle: " +
            angle +
            " degrees"
        );
    }
}
