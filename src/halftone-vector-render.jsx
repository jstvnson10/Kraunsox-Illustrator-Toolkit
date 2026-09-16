#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT
// Vector Tone Rendering v0.9
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
        alert("Select exactly one filled vector object.");
        return;
    }

    var artwork = doc.selection[0];


    // ==================================================
    // VALIDATE ARTWORK
    // ==================================================

    if (artwork.typename !== "PathItem") {

        alert(
            "v0.9 currently supports one PathItem.\n\n" +
            "Try a rectangle, circle, or star."
        );

        return;
    }

    if (!artwork.filled) {

        alert(
            "The selected object needs a fill color."
        );

        return;
    }


    // ==================================================
    // USER SETTINGS
    // ==================================================

    var gamma = Number(
        prompt(
            "Enter gamma:",
            CONFIG.gamma
        )
    );

    var angle = Number(
        prompt(
            "Enter screen angle:",
            CONFIG.angle
        )
    );

    var spacing = Number(
        prompt(
            "Enter dot spacing:",
            CONFIG.spacing
        )
    );


    // ==================================================
    // READ ACTUAL ARTWORK COLOR
    // ==================================================

    var brightness =
        getBrightness(
            artwork.fillColor
        );

    var darkness =
        1 - brightness;

    var tone =
        calculateTone(
            darkness,
            gamma
        );

    var dotSize =
        calculateDotSize(
            tone,
            CONFIG.minDotSize,
            CONFIG.maxDotSize
        );


    // ==================================================
    // GET ARTWORK BOUNDS
    // ==================================================

    var bounds =
        getArtworkBounds(
            artwork
        );


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
    // OUTPUT LAYER
    // ==================================================

    var outputLayer =
        doc.layers.add();

    outputLayer.name =
        "KRAUNSOX VECTOR HALFTONE";


    // ==================================================
    // CLIPPING GROUP
    // ==================================================

    var clippingGroup =
        outputLayer.groupItems.add();

    clippingGroup.name =
        "HALFTONE CLIPPING GROUP";


    // ==================================================
    // DOT GROUP
    // ==================================================

    var dotGroup =
        clippingGroup.groupItems.add();

    dotGroup.name =
        "VECTOR DOTS";


    // ==================================================
    // GENERATE HALFTONE
    // ==================================================

    generateHalftone(
        dotGroup,
        bounds,
        gridRadius,
        spacing,
        dotSize,
        angle
    );


    // ==================================================
    // CLIP TO ORIGINAL SHAPE
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
        "REAL VECTOR HALFTONE CREATED\n\n" +

        "Brightness: " +
        brightness.toFixed(3) +

        "\nDarkness: " +
        darkness.toFixed(3) +

        "\nDot Size: " +
        dotSize.toFixed(2) +
        " pt" +

        "\nAngle: " +
        angle +
        " degrees"
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

    return {

        x:
            x *
            Math.cos(radians) -
            y *
            Math.sin(radians),

        y:
            x *
            Math.sin(radians) +
            y *
            Math.cos(radians)
    };
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
        "Unsupported color type: " +
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

    if (darkness < 0) {
        darkness = 0;
    }

    if (darkness > 1) {
        darkness = 1;
    }

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
        (maxSize - minSize)
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

    // Don't create zero-size dots.

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
    dot.fillColor = color;
    dot.stroked = false;
}


// ======================================================
// GENERATE HALFTONE
// ======================================================

function generateHalftone(
    dotGroup,
    bounds,
    gridRadius,
    spacing,
    dotSize,
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

            var gridX =
                column *
                spacing;

            var gridY =
                row *
                spacing;


            // Rotate grid

            var rotated =
                rotatePoint(
                    gridX,
                    gridY,
                    angle
                );


            // Move to artwork center

            var centerX =
                bounds.centerX +
                rotated.x;

            var centerY =
                bounds.centerY +
                rotated.y;


            // Only create useful dots

            if (
                centerX >=
                    bounds.left -
                    dotSize &&

                centerX <=
                    bounds.right +
                    dotSize &&

                centerY <=
                    bounds.top +
                    dotSize &&

                centerY >=
                    bounds.bottom -
                    dotSize
            ) {

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

    mask.clipping = true;

    clippingGroup.clipped = true;

    return mask;
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
