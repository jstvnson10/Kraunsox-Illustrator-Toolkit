#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT
// Spatial Vector Halftone v1.0
//
// First version where each grid position can receive
// its own tone value.
// ======================================================

var CONFIG = {
    spacing: 14,
    minDotSize: 0,
    maxDotSize: 13,
    gamma: 1.0,
    angle: 45
};


// ======================================================
// MAIN
// ======================================================

function main() {

    if (app.documents.length === 0) {
        alert("Open an Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length !== 1) {

        alert(
            "Select exactly one gradient-filled vector object."
        );

        return;
    }

    var artwork =
        doc.selection[0];


    // ==================================================
    // VALIDATION
    // ==================================================

    if (artwork.typename !== "PathItem") {

        alert(
            "v1.0 currently supports one PathItem."
        );

        return;
    }

    if (!artwork.filled) {

        alert(
            "The selected object must have a fill."
        );

        return;
    }

    if (
        artwork.fillColor.typename !==
        "GradientColor"
    ) {

        alert(
            "v1.0 expects a gradient-filled object.\n\n" +
            "Create a black-to-white linear gradient first."
        );

        return;
    }


    // ==================================================
    // USER SETTINGS
    // ==================================================

    var gammaInput =
        prompt(
            "Enter gamma:",
            "1.0"
        );

    if (gammaInput === null) {
        return;
    }

    var gamma =
        parseFloat(gammaInput);

    if (
        isNaN(gamma) ||
        gamma <= 0
    ) {

        alert(
            "Gamma must be greater than 0."
        );

        return;
    }


    var angleInput =
        prompt(
            "Enter halftone screen angle:",
            "45"
        );

    if (angleInput === null) {
        return;
    }

    var angle =
        parseFloat(angleInput);

    if (isNaN(angle)) {

        alert(
            "Angle must be a number."
        );

        return;
    }


    var spacingInput =
        prompt(
            "Enter dot spacing:",
            "14"
        );

    if (spacingInput === null) {
        return;
    }

    var spacing =
        parseFloat(spacingInput);

    if (
        isNaN(spacing) ||
        spacing <= 0
    ) {

        alert(
            "Spacing must be greater than 0."
        );

        return;
    }


    // ==================================================
    // ARTWORK INFORMATION
    // ==================================================

    var bounds =
        getArtworkBounds(
            artwork
        );

    var gradient =
        artwork.fillColor.gradient;


    // ==================================================
    // GRID COVERAGE
    // ==================================================

    var gridRadius =
        calculateGridRadius(
            bounds.width,
            bounds.height,
            spacing
        );


    // ==================================================
    // OUTPUT LAYER
    // ==================================================

    var outputLayer =
        doc.layers.add();

    outputLayer.name =
        "KRAUNSOX SPATIAL HALFTONE v1.0";


    var clippingGroup =
        outputLayer.groupItems.add();

    clippingGroup.name =
        "HALFTONE CLIPPING GROUP";


    var dotGroup =
        clippingGroup.groupItems.add();

    dotGroup.name =
        "SPATIALLY SAMPLED DOTS";


    // ==================================================
    // GENERATE
    // ==================================================

    generateSpatialHalftone(
        dotGroup,
        bounds,
        gradient,
        gridRadius,
        spacing,
        CONFIG.minDotSize,
        CONFIG.maxDotSize,
        gamma,
        angle
    );


    // ==================================================
    // CLIPPING MASK
    // ==================================================

    createClippingMask(
        artwork,
        clippingGroup
    );


    // ==================================================
    // FINISH
    // ==================================================

    artwork.selected = false;

    clippingGroup.selected = true;

    app.redraw();

    alert(
        "KRAUNSOX SPATIAL HALFTONE v1.0\n\n" +
        "Each dot was calculated from its " +
        "position across the gradient."
    );
}


// ======================================================
// ARTWORK BOUNDS
// ======================================================

function getArtworkBounds(
    artwork
) {

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
// GRID SIZE
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
// ROTATE GRID POINT
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
// SPATIAL POSITION
// ======================================================

function getHorizontalPosition(
    x,
    bounds
) {

    var percentage =
        (
            x -
            bounds.left
        ) /
        bounds.width;

    return clamp(
        percentage,
        0,
        1
    );
}


// ======================================================
// SAMPLE GRADIENT
// ======================================================

function sampleGradient(
    gradient,
    position
) {

    var stops =
        gradient.gradientStops;

    if (stops.length === 0) {
        return 1;
    }


    // Convert 0-1 position to Illustrator's
    // 0-100 gradient ramp position.

    var rampPosition =
        position * 100;


    // Before first stop

    if (
        rampPosition <=
        stops[0].rampPoint
    ) {

        return getBrightness(
            stops[0].color
        );
    }


    // Search neighboring gradient stops.

    for (
        var i = 0;
        i < stops.length - 1;
        i++
    ) {

        var stopA =
            stops[i];

        var stopB =
            stops[i + 1];


        if (
            rampPosition >=
                stopA.rampPoint &&

            rampPosition <=
                stopB.rampPoint
        ) {

            var range =
                stopB.rampPoint -
                stopA.rampPoint;


            var localPosition;

            if (range === 0) {

                localPosition = 0;

            } else {

                localPosition =
                    (
                        rampPosition -
                        stopA.rampPoint
                    ) /
                    range;
            }


            var brightnessA =
                getBrightness(
                    stopA.color
                );

            var brightnessB =
                getBrightness(
                    stopB.color
                );


            return interpolate(
                brightnessA,
                brightnessB,
                localPosition
            );
        }
    }


    // After final stop

    return getBrightness(
        stops[
            stops.length - 1
        ].color
    );
}


// ======================================================
// INTERPOLATION
// ======================================================

function interpolate(
    start,
    end,
    amount
) {

    return (
        start +
        (
            end -
            start
        ) *
        amount
    );
}


// ======================================================
// COLOR BRIGHTNESS
// ======================================================

function getBrightness(
    color
) {

    if (
        color.typename ===
        "RGBColor"
    ) {

        return rgbBrightness(
            color.red,
            color.green,
            color.blue
        );
    }


    if (
        color.typename ===
        "GrayColor"
    ) {

        return (
            1 -
            color.gray / 100
        );
    }


    if (
        color.typename ===
        "CMYKColor"
    ) {

        return cmykBrightness(
            color.cyan,
            color.magenta,
            color.yellow,
            color.black
        );
    }


    throw new Error(
        "Unsupported gradient stop color: " +
        color.typename
    );
}


// ======================================================
// RGB BRIGHTNESS
// ======================================================

function rgbBrightness(
    red,
    green,
    blue
) {

    var r =
        red / 255;

    var g =
        green / 255;

    var b =
        blue / 255;

    return (
        0.2126 * r +
        0.7152 * g +
        0.0722 * b
    );
}


// ======================================================
// CMYK BRIGHTNESS
// ======================================================

function cmykBrightness(
    cyan,
    magenta,
    yellow,
    blackValue
) {

    var c =
        cyan / 100;

    var m =
        magenta / 100;

    var y =
        yellow / 100;

    var k =
        blackValue / 100;


    var r =
        (1 - c) *
        (1 - k);

    var g =
        (1 - m) *
        (1 - k);

    var b =
        (1 - y) *
        (1 - k);


    return (
        0.2126 * r +
        0.7152 * g +
        0.0722 * b
    );
}


// ======================================================
// TONE CURVE
// ======================================================

function calculateTone(
    darkness,
    gamma
) {

    darkness =
        clamp(
            darkness,
            0,
            1
        );

    return Math.pow(
        darkness,
        gamma
    );
}


// ======================================================
// DOT SIZE
// ======================================================

function calculateDotSize(
    tone,
    minSize,
    maxSize
) {

    return (
        minSize +
        tone *
        (
            maxSize -
            minSize
        )
    );
}


// ======================================================
// CREATE BLACK
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

    if (size <= 0.01) {
        return;
    }


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

    dot.fillColor =
        color;

    dot.stroked =
        false;
}


// ======================================================
// GENERATE SPATIAL HALFTONE
// ======================================================

function generateSpatialHalftone(
    dotGroup,
    bounds,
    gradient,
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

            // ------------------------------------------
            // GRID POSITION
            // ------------------------------------------

            var gridX =
                column *
                spacing;

            var gridY =
                row *
                spacing;


            // ------------------------------------------
            // ROTATE HALFTONE SCREEN
            // ------------------------------------------

            var rotated =
                rotatePoint(
                    gridX,
                    gridY,
                    angle
                );


            var centerX =
                bounds.centerX +
                rotated.x;

            var centerY =
                bounds.centerY +
                rotated.y;


            // ------------------------------------------
            // CHECK BOUNDS
            // ------------------------------------------

            if (
                centerX >=
                    bounds.left &&

                centerX <=
                    bounds.right &&

                centerY <=
                    bounds.top &&

                centerY >=
                    bounds.bottom
            ) {

                // ======================================
                // THE IMPORTANT PART
                // ======================================

                var position =
                    getHorizontalPosition(
                        centerX,
                        bounds
                    );


                var brightness =
                    sampleGradient(
                        gradient,
                        position
                    );


                var darkness =
                    1 -
                    brightness;


                var tone =
                    calculateTone(
                        darkness,
                        gamma
                    );


                var dotSize =
                    calculateDotSize(
                        tone,
                        minDotSize,
                        maxDotSize
                    );


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
// CLIPPING MASK
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


    mask.clipping =
        true;


    clippingGroup.clipped =
        true;


    return mask;
}


// ======================================================
// CLAMP
// ======================================================

function clamp(
    value,
    minimum,
    maximum
) {

    if (value < minimum) {
        return minimum;
    }

    if (value > maximum) {
        return maximum;
    }

    return value;
}


// ======================================================
// RUN
// ======================================================

try {

    main();

} catch (error) {

    alert(
        "KRAUNSOX HALFTONE ERROR\n\n" +
        error.message +
        "\n\nLine: " +
        error.line
    );
}
