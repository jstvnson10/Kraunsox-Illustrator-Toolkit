#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT
// Gradient Direction v1.1
// ======================================================

var CONFIG = {
    spacing: 14,
    minDotSize: 0,
    maxDotSize: 13,
    gamma: 1.0,
    screenAngle: 45,
    gradientAngle: 0
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
        alert("Select exactly one gradient-filled vector object.");
        return;
    }

    var artwork = doc.selection[0];


    // ==================================================
    // VALIDATION
    // ==================================================

    if (artwork.typename !== "PathItem") {
        alert("v1.1 currently supports one PathItem.");
        return;
    }

    if (!artwork.filled) {
        alert("The selected object must have a fill.");
        return;
    }

    if (artwork.fillColor.typename !== "GradientColor") {

        alert(
            "v1.1 requires a gradient-filled object."
        );

        return;
    }


    // ==================================================
    // USER INPUT
    // ==================================================

    var gamma =
        getNumberInput(
            "Enter gamma:",
            CONFIG.gamma,
            true
        );

    if (gamma === null) {
        return;
    }


    var screenAngle =
        getNumberInput(
            "Enter HALFTONE screen angle:",
            CONFIG.screenAngle,
            false
        );

    if (screenAngle === null) {
        return;
    }


    var gradientAngle =
        getNumberInput(
            "Enter SOURCE gradient angle:",
            CONFIG.gradientAngle,
            false
        );

    if (gradientAngle === null) {
        return;
    }


    var spacing =
        getNumberInput(
            "Enter dot spacing:",
            CONFIG.spacing,
            true
        );

    if (spacing === null) {
        return;
    }


    // ==================================================
    // ARTWORK DATA
    // ==================================================

    var bounds =
        getArtworkBounds(
            artwork
        );

    var gradient =
        artwork.fillColor.gradient;


    // ==================================================
    // GRID SIZE
    // ==================================================

    var gridRadius =
        calculateGridRadius(
            bounds.width,
            bounds.height,
            spacing
        );


    // ==================================================
    // OUTPUT
    // ==================================================

    var outputLayer =
        doc.layers.add();

    outputLayer.name =
        "KRAUNSOX HALFTONE v1.1";


    var clippingGroup =
        outputLayer.groupItems.add();

    clippingGroup.name =
        "HALFTONE CLIPPING GROUP";


    var dotGroup =
        clippingGroup.groupItems.add();

    dotGroup.name =
        "SPATIAL DOTS";


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
        screenAngle,
        gradientAngle
    );


    // ==================================================
    // MASK
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
        "KRAUNSOX HALFTONE v1.1 CREATED\n\n" +

        "Screen Angle: " +
        screenAngle +
        " degrees\n" +

        "Gradient Angle: " +
        gradientAngle +
        " degrees\n" +

        "Spacing: " +
        spacing +
        " pt"
    );
}


// ======================================================
// SAFE NUMBER INPUT
// ======================================================

function getNumberInput(
    message,
    defaultValue,
    mustBePositive
) {

    var input =
        prompt(
            message,
            String(defaultValue)
        );

    if (input === null) {
        return null;
    }

    var value =
        parseFloat(input);

    if (isNaN(value)) {

        alert(
            "Please enter a valid number."
        );

        return null;
    }

    if (
        mustBePositive &&
        value <= 0
    ) {

        alert(
            "The value must be greater than 0."
        );

        return null;
    }

    return value;
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
            left +
            width / 2,

        centerY:
            bottom +
            height / 2
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
// PROJECT POINT ONTO GRADIENT AXIS
// ======================================================

function projectOntoGradient(
    x,
    y,
    bounds,
    gradientAngle
) {

    // Convert angle to radians.

    var radians =
        gradientAngle *
        Math.PI /
        180;


    // Unit vector representing the
    // direction of the gradient.

    var directionX =
        Math.cos(radians);

    var directionY =
        Math.sin(radians);


    // Position relative to artwork center.

    var relativeX =
        x -
        bounds.centerX;

    var relativeY =
        y -
        bounds.centerY;


    // Dot product.
    //
    // This tells us how far the sample
    // lies along the gradient direction.

    return (
        relativeX *
        directionX +

        relativeY *
        directionY
    );
}


// ======================================================
// GRADIENT RANGE
// ======================================================

function calculateGradientRange(
    bounds,
    gradientAngle
) {

    var radians =
        gradientAngle *
        Math.PI /
        180;

    var directionX =
        Math.cos(radians);

    var directionY =
        Math.sin(radians);


    // Half-width / half-height

    var halfWidth =
        bounds.width / 2;

    var halfHeight =
        bounds.height / 2;


    // Projection radius of rectangle
    // onto the gradient axis.

    var radius =
        Math.abs(
            directionX
        ) *
        halfWidth +

        Math.abs(
            directionY
        ) *
        halfHeight;


    return {
        minimum: -radius,
        maximum: radius
    };
}


// ======================================================
// GET POSITION ALONG GRADIENT
// ======================================================

function getGradientPosition(
    x,
    y,
    bounds,
    gradientAngle
) {

    var projection =
        projectOntoGradient(
            x,
            y,
            bounds,
            gradientAngle
        );


    var range =
        calculateGradientRange(
            bounds,
            gradientAngle
        );


    var totalRange =
        range.maximum -
        range.minimum;


    if (totalRange === 0) {
        return 0;
    }


    var position =
        (
            projection -
            range.minimum
        ) /
        totalRange;


    return clamp(
        position,
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


    var rampPosition =
        position *
        100;


    // Before first stop

    if (
        rampPosition <=
        stops[0].rampPoint
    ) {

        return getBrightness(
            stops[0].color
        );
    }


    // Find surrounding stops

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


    return getBrightness(
        stops[
            stops.length - 1
        ].color
    );
}


// ======================================================
// INTERPOLATE
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
// BRIGHTNESS
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
// TONE
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


    dot.filled =
        true;

    dot.fillColor =
        color;

    dot.stroked =
        false;
}


// ======================================================
// GENERATE HALFTONE
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
    screenAngle,
    gradientAngle
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
            // BASE GRID
            // ------------------------------------------

            var gridX =
                column *
                spacing;

            var gridY =
                row *
                spacing;


            // ------------------------------------------
            // HALFTONE SCREEN ROTATION
            // ------------------------------------------

            var rotated =
                rotatePoint(
                    gridX,
                    gridY,
                    screenAngle
                );


            var centerX =
                bounds.centerX +
                rotated.x;

            var centerY =
                bounds.centerY +
                rotated.y;


            // ------------------------------------------
            // ARTWORK BOUNDS
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
                // SPATIAL SAMPLE
                // ======================================

                var position =
                    getGradientPosition(
                        centerX,
                        centerY,
                        bounds,
                        gradientAngle
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
