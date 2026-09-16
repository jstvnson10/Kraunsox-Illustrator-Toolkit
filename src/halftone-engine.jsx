#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT
// Halftone Engine v0.7
// ======================================================


// ======================================================
// CONFIGURATION
// ======================================================

var CONFIG = {
    spacing: 14,
    minDotSize: 1,
    maxDotSize: 12
};


// ======================================================
// MAIN
// ======================================================

function main() {

    // --------------------------------------
    // Check for document
    // --------------------------------------

    if (app.documents.length === 0) {
        alert("Open an Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;


    // --------------------------------------
    // Check selection
    // --------------------------------------

    if (doc.selection.length !== 1) {
        alert("Select exactly one vector object.");
        return;
    }

    var artwork = doc.selection[0];


    // --------------------------------------
    // User controls
    // --------------------------------------

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


    // --------------------------------------
    // Get artwork information
    // --------------------------------------

    var bounds =
        getArtworkBounds(artwork);

    var gridRadius =
        calculateGridRadius(
            bounds.width,
            bounds.height,
            CONFIG.spacing
        );


    // --------------------------------------
    // Create output
    // --------------------------------------

    var outputLayer =
        createOutputLayer(
            doc,
            angle,
            gamma
        );

    var clippingGroup =
        outputLayer.groupItems.add();

    clippingGroup.name =
        "HALFTONE CLIPPING GROUP";

    var dotGroup =
        clippingGroup.groupItems.add();

    dotGroup.name =
        "VECTOR DOTS";


    // --------------------------------------
    // Generate dots
    // --------------------------------------

    generateHalftone(
        dotGroup,
        bounds,
        gridRadius,
        CONFIG.spacing,
        CONFIG.minDotSize,
        CONFIG.maxDotSize,
        gamma,
        angle
    );


    // --------------------------------------
    // Create clipping mask
    // --------------------------------------

    createClippingMask(
        artwork,
        clippingGroup
    );


    // --------------------------------------
    // Finish
    // --------------------------------------

    artwork.selected = false;
    clippingGroup.selected = true;

    app.redraw();

    alert(
        "KRAUNSOX Halftone created!\n\n" +
        "Angle: " + angle + " degrees\n" +
        "Gamma: " + gamma
    );
}


// ======================================================
// GET ARTWORK BOUNDS
// ======================================================

function getArtworkBounds(artwork) {

    var b =
        artwork.geometricBounds;

    var left =
        b[0];

    var top =
        b[1];

    var right =
        b[2];

    var bottom =
        b[3];

    var width =
        right - left;

    var height =
        top - bottom;

    return {

        left: left,
        top: top,
        right: right,
        bottom: bottom,

        width: width,
        height: height,

        centerX:
            left + width / 2,

        centerY:
            bottom + height / 2
    };
}


// ======================================================
// GRID RADIUS
// ======================================================

function calculateGridRadius(
    width,
    height,
    spacing
) {

    var diagonal =
        Math.sqrt(
            width * width +
            height * height
        );

    return Math.ceil(
        diagonal / spacing
    );
}


// ======================================================
// ROTATE POINT
// ======================================================

function rotatePoint(
    x,
    y,
    angle
) {

    var radians =
        angle *
        Math.PI /
        180;

    var cosAngle =
        Math.cos(radians);

    var sinAngle =
        Math.sin(radians);

    return {

        x:
            x * cosAngle -
            y * sinAngle,

        y:
            x * sinAngle +
            y * cosAngle
    };
}


// ======================================================
// CALCULATE TONE
// ======================================================

function calculateTone(
    percentage,
    gamma
) {

    // Clamp between 0 and 1

    if (percentage < 0) {
        percentage = 0;
    }

    if (percentage > 1) {
        percentage = 1;
    }

    return Math.pow(
        percentage,
        gamma
    );
}


// ======================================================
// CALCULATE DOT SIZE
// ======================================================

function calculateDotSize(
    tone,
    minSize,
    maxSize
) {

    return (
        minSize +
        tone *
        (maxSize - minSize)
    );
}


// ======================================================
// CREATE BLACK COLOR
// ======================================================

function createBlack() {

    var black =
        new RGBColor();

    black.red = 0;
    black.green = 0;
    black.blue = 0;

    return black;
}


// ======================================================
// CREATE DOT
// ======================================================

function createDot(
    parent,
    centerX,
    centerY,
    size,
    color
) {

    var left =
        centerX -
        size / 2;

    var top =
        centerY +
        size / 2;

    var dot =
        parent.pathItems.ellipse(
            top,
            left,
            size,
            size
        );

    dot.filled = true;
    dot.fillColor = color;
    dot.stroked = false;

    return dot;
}


// ======================================================
// CREATE OUTPUT LAYER
// ======================================================

function createOutputLayer(
    doc,
    angle,
    gamma
) {

    var layer =
        doc.layers.add();

    layer.name =
        "KRAUNSOX HALFTONE " +
        angle +
        "deg GAMMA " +
        gamma;

    return layer;
}


// ======================================================
// GENERATE HALFTONE
// ======================================================

function generateHalftone(
    dotGroup,
    bounds,
    gridRadius,
    spacing,
    minDotSize,
    maxDotSize,
    gamma,
    angle
) {

    var black =
        createBlack();


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

            // ----------------------------------
            // Original grid coordinate
            // ----------------------------------

            var gridX =
                column * spacing;

            var gridY =
                row * spacing;


            // ----------------------------------
            // Rotate point
            // ----------------------------------

            var rotated =
                rotatePoint(
                    gridX,
                    gridY,
                    angle
                );


            // ----------------------------------
            // Move to artwork center
            // ----------------------------------

            var centerX =
                bounds.centerX +
                rotated.x;

            var centerY =
                bounds.centerY +
                rotated.y;


            // ----------------------------------
            // Ignore unnecessary dots
            // ----------------------------------

            if (
                centerX >=
                    bounds.left -
                    maxDotSize &&

                centerX <=
                    bounds.right +
                    maxDotSize &&

                centerY <=
                    bounds.top +
                    maxDotSize &&

                centerY >=
                    bounds.bottom -
                    maxDotSize
            ) {

                // ----------------------------------
                // TEMPORARY TONE
                // ----------------------------------

                var percentage =
                    (
                        centerX -
                        bounds.left
                    ) /
                    bounds.width;


                // ----------------------------------
                // Tone curve
                // ----------------------------------

                var tone =
                    calculateTone(
                        percentage,
                        gamma
                    );


                // ----------------------------------
                // Dot size
                // ----------------------------------

                var dotSize =
                    calculateDotSize(
                        tone,
                        minDotSize,
                        maxDotSize
                    );


                // ----------------------------------
                // Draw dot
                // ----------------------------------

                createDot(
                    dotGroup,
                    centerX,
                    centerY,
                    dotSize,
                    black
                );
            }
        }
    }
}


// ======================================================
// CREATE CLIPPING MASK
// ======================================================

function createClippingMask(
    artwork,
    clippingGroup
) {

    var mask =
        artwork.duplicate(
            clippingGroup,
            ElementPlacement.PLACEATBEGINNING
        );

    mask.name =
        "HALFTONE MASK";

    mask.clipping = true;

    clippingGroup.clipped = true;

    return mask;
}


// ======================================================
// RUN
// ======================================================

main();
