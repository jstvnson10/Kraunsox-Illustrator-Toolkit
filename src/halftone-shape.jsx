#target illustrator

// ==========================================
// KRAUNSOX HALFTONE TOOLKIT
// Shape Halftone v0.6
// ==========================================

if (app.documents.length === 0) {

    alert("Open an Illustrator document first.");

} else {

    var doc = app.activeDocument;

    if (doc.selection.length !== 1) {

        alert("Select exactly one vector object.");

    } else {

        var artwork = doc.selection[0];

        // ----------------------------------
        // SETTINGS
        // ----------------------------------

        var spacing = 14;
        var minDotSize = 1;
        var maxDotSize = 12;

        var gamma = Number(
            prompt("Enter gamma:", "1.0")
        );

        var angle = Number(
            prompt("Enter screen angle:", "45")
        );

        // ----------------------------------
        // BOUNDS
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

        var artCenterX =
            artLeft + (artWidth / 2);

        var artCenterY =
            artBottom + (artHeight / 2);

        // ----------------------------------
        // GRID SIZE
        // Extra coverage because rotation
        // makes the grid extend farther.
        // ----------------------------------

        var diagonal =
            Math.sqrt(
                artWidth * artWidth +
                artHeight * artHeight
            );

        var gridRadius =
            Math.ceil(
                diagonal / spacing
            );

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
        // BLACK
        // ----------------------------------

        var black =
            new RGBColor();

        black.red = 0;
        black.green = 0;
        black.blue = 0;

        // ----------------------------------
        // OUTPUT LAYER
        // ----------------------------------

        var outputLayer =
            doc.layers.add();

        outputLayer.name =
            "KRAUNSOX HALFTONE " +
            angle +
            "deg";

        // ----------------------------------
        // CLIPPING GROUP
        // ----------------------------------

        var clippingGroup =
            outputLayer.groupItems.add();

        clippingGroup.name =
            "HALFTONE CLIPPING GROUP";

        // ----------------------------------
        // DOT GROUP
        // ----------------------------------

        var dotGroup =
            clippingGroup.groupItems.add();

        dotGroup.name =
            "VECTOR DOTS";

        // ----------------------------------
        // GENERATE DOTS
        // ----------------------------------

        for (
            var row = -gridRadius;
            row <= gridRadius;
            row++
        ) {

            for (
                var column = -gridRadius;
                column <= gridRadius;
                column++
            ) {

                var gridX =
                    column * spacing;

                var gridY =
                    row * spacing;

                // Rotate coordinate system

                var rotatedX =
                    gridX * cosAngle -
                    gridY * sinAngle;

                var rotatedY =
                    gridX * sinAngle +
                    gridY * cosAngle;

                var centerX =
                    artCenterX +
                    rotatedX;

                var centerY =
                    artCenterY +
                    rotatedY;

                // Only generate dots reasonably
                // close to artwork bounds.

                if (
                    centerX >= artLeft - maxDotSize &&
                    centerX <= artRight + maxDotSize &&
                    centerY <= artTop + maxDotSize &&
                    centerY >= artBottom - maxDotSize
                ) {

                    // ----------------------------------
                    // TEMPORARY TONE
                    // ----------------------------------

                    var percentage =
                        (centerX - artLeft) /
                        artWidth;

                    // Clamp 0 -> 1

                    if (percentage < 0) {
                        percentage = 0;
                    }

                    if (percentage > 1) {
                        percentage = 1;
                    }

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

                    // ----------------------------------
                    // CREATE DOT
                    // ----------------------------------

                    var x =
                        centerX -
                        dotSize / 2;

                    var y =
                        centerY +
                        dotSize / 2;

                    var dot =
                        dotGroup.pathItems.ellipse(
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

        // ----------------------------------
        // DUPLICATE ARTWORK FOR MASK
        // ----------------------------------

        var mask =
            artwork.duplicate(
                clippingGroup,
                ElementPlacement.PLACEATBEGINNING
            );

        mask.name =
            "HALFTONE MASK";

        // ----------------------------------
        // CREATE CLIPPING MASK
        // ----------------------------------

        mask.clipping = true;

        clippingGroup.clipped = true;

        // ----------------------------------
        // FINISH
        // ----------------------------------

        artwork.selected = false;
        clippingGroup.selected = true;

        app.redraw();

        alert(
            "Vector halftone created!\n\n" +
            "Angle: " + angle +
            " degrees\n" +
            "Gamma: " + gamma
        );
    }
}
