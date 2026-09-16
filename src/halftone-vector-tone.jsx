#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT
// Vector Tone Sampling v0.8
// ======================================================

var CONFIG = {
    spacing: 14,
    minDotSize: 1,
    maxDotSize: 12,
    gamma: 1.0
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
        alert("Select exactly one filled vector shape.");
        return;
    }

    var artwork = doc.selection[0];

    if (artwork.typename !== "PathItem") {
        alert(
            "v0.8 currently supports a single PathItem.\n\n" +
            "Try a filled rectangle, circle, or star."
        );
        return;
    }

    if (!artwork.filled) {
        alert("The selected object needs a fill color.");
        return;
    }

    var gamma = Number(
        prompt(
            "Enter gamma:",
            "1.0"
        )
    );

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

    alert(
        "REAL VECTOR TONE DATA\n\n" +
        "Brightness: " +
        brightness.toFixed(3) +
        "\n" +

        "Darkness: " +
        darkness.toFixed(3) +
        "\n\n" +

        "Calculated dot size: " +
        dotSize.toFixed(2) +
        " pt"
    );
}


// ======================================================
// GET BRIGHTNESS
// ======================================================

function getBrightness(color) {

    if (color.typename === "RGBColor") {

        return rgbBrightness(
            color.red,
            color.green,
            color.blue
        );
    }


    if (color.typename === "GrayColor") {

        return (
            1 -
            color.gray / 100
        );
    }


    if (color.typename === "CMYKColor") {

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

    // Normalize RGB from 0-255 to 0-1

    var r =
        red / 255;

    var g =
        green / 255;

    var b =
        blue / 255;


    // Perceived luminance approximation

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

    // Convert percentages to 0-1

    var c =
        cyan / 100;

    var m =
        magenta / 100;

    var y =
        yellow / 100;

    var k =
        blackValue / 100;


    // Approximate CMYK -> RGB

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
// RUN
// ======================================================

try {

    main();

} catch (error) {

    alert(
        "KRAUNSOX HALFTONE ERROR\n\n" +
        error.message
    );
}
