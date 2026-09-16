#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT
// ScriptUI Edition v1.2
// ======================================================

var DEFAULTS = {
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
    // VALIDATE ARTWORK
    // ==================================================

    if (artwork.typename !== "PathItem") {
        alert("v1.2 currently supports one PathItem.");
        return;
    }

    if (!artwork.filled) {
        alert("The selected object must have a fill.");
        return;
    }

    if (artwork.fillColor.typename !== "GradientColor") {

        alert(
            "v1.2 requires a gradient-filled object."
        );

        return;
    }


    // ==================================================
    // SHOW UI
    // ==================================================

    var settings =
        showHalftoneDialog();

    if (settings === null) {
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
    // GRID
    // ==================================================

    var gridRadius =
        calculateGridRadius(
            bounds.width,
            bounds.height,
            settings.spacing
        );


    // ==================================================
    // OUTPUT
    // ==================================================

    var outputLayer =
        doc.layers.add();

    outputLayer.name =
        "KRAUNSOX HALFTONE v1.2";


    var clippingGroup =
        outputLayer.groupItems.add();

    clippingGroup.name =
        "HALFTONE CLIPPING GROUP";


    var dotGroup =
        clippingGroup.groupItems.add();

    dotGroup.name =
        "HALFTONE DOTS";


    // ==================================================
    // GENERATE
    // ==================================================

    generateSpatialHalftone(
        dotGroup,
        bounds,
        gradient,
        gridRadius,
        settings.spacing,
        settings.minDotSize,
        settings.maxDotSize,
        settings.gamma,
        settings.screenAngle,
        settings.gradientAngle
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
        "KRAUNSOX HALFTONE CREATED\n\n" +

        "Spacing: " +
        settings.spacing +
        " pt\n" +

        "Dot Range: " +
        settings.minDotSize +
        " - " +
        settings.maxDotSize +
        " pt\n" +

        "Gamma: " +
        settings.gamma +
        "\n" +

        "Screen Angle: " +
        settings.screenAngle +
        " degrees\n" +

        "Gradient Angle: " +
        settings.gradientAngle +
        " degrees"
    );
}


// ======================================================
// USER INTERFACE
// ======================================================

function showHalftoneDialog() {

    var dialog =
        new Window(
            "dialog",
            "KRAUNSOX Halftone Toolkit v1.2"
        );

    dialog.orientation =
        "column";

    dialog.alignChildren =
        ["fill", "top"];

    dialog.spacing =
        12;

    dialog.margins =
        18;


    // ==================================================
    // TITLE
    // ==================================================

    var title =
        dialog.add(
            "statictext",
            undefined,
            "KRAUNSOX HALFTONE"
        );

    title.alignment =
        "center";


    var subtitle =
        dialog.add(
            "statictext",
            undefined,
            "Vector Halftone Generator"
        );

    subtitle.alignment =
        "center";


    // ==================================================
    // DOT SETTINGS PANEL
    // ==================================================

    var dotPanel =
        dialog.add(
            "panel",
            undefined,
            "Dot Settings"
        );

    dotPanel.orientation =
        "column";

    dotPanel.alignChildren =
        ["fill", "center"];

    dotPanel.margins =
        15;


    var spacingField =
        addInputRow(
            dotPanel,
            "Spacing:",
            DEFAULTS.spacing,
            "pt"
        );


    var minDotField =
        addInputRow(
            dotPanel,
            "Min Dot:",
            DEFAULTS.minDotSize,
            "pt"
        );


    var maxDotField =
        addInputRow(
            dotPanel,
            "Max Dot:",
            DEFAULTS.maxDotSize,
            "pt"
        );


    // ==================================================
    // TONE SETTINGS
    // ==================================================

    var tonePanel =
        dialog.add(
            "panel",
            undefined,
            "Tone"
        );

    tonePanel.orientation =
        "column";

    tonePanel.alignChildren =
        ["fill", "center"];

    tonePanel.margins =
        15;


    var gammaField =
        addInputRow(
            tonePanel,
            "Gamma:",
            DEFAULTS.gamma,
            ""
        );


    // ==================================================
    // ANGLE SETTINGS
    // ==================================================

    var anglePanel =
        dialog.add(
            "panel",
            undefined,
            "Angles"
        );

    anglePanel.orientation =
        "column";

    anglePanel.alignChildren =
        ["fill", "center"];

    anglePanel.margins =
        15;


    var screenAngleField =
        addInputRow(
            anglePanel,
            "Screen:",
            DEFAULTS.screenAngle,
            "deg"
        );


    var gradientAngleField =
        addInputRow(
            anglePanel,
            "Gradient:",
            DEFAULTS.gradientAngle,
            "deg"
        );


    // ==================================================
    // BUTTONS
    // ==================================================

    var buttonGroup =
        dialog.add(
            "group"
        );

    buttonGroup.orientation =
        "row";

    buttonGroup.alignment =
        "right";


    var resetButton =
        buttonGroup.add(
            "button",
            undefined,
            "Reset"
        );


    var cancelButton =
        buttonGroup.add(
            "button",
            undefined,
            "Cancel",
            {
                name: "cancel"
            }
        );


    var generateButton =
        buttonGroup.add(
            "button",
            undefined,
            "Generate",
            {
                name: "ok"
            }
        );


    // ==================================================
    // RESET BUTTON
    // ==================================================

    resetButton.onClick =
        function() {

            spacingField.text =
                DEFAULTS.spacing;

            minDotField.text =
                DEFAULTS.minDotSize;

            maxDotField.text =
                DEFAULTS.maxDotSize;

            gammaField.text =
                DEFAULTS.gamma;

            screenAngleField.text =
                DEFAULTS.screenAngle;

            gradientAngleField.text =
                DEFAULTS.gradientAngle;
        };


    // ==================================================
    // GENERATE BUTTON
    // ==================================================

    generateButton.onClick =
        function() {

            var spacing =
                parseFloat(
                    spacingField.text
                );

            var minDotSize =
                parseFloat(
                    minDotField.text
                );

            var maxDotSize =
                parseFloat(
                    maxDotField.text
                );

            var gamma =
                parseFloat(
                    gammaField.text
                );

            var screenAngle =
                parseFloat(
                    screenAngleField.text
                );

            var gradientAngle =
                parseFloat(
                    gradientAngleField.text
                );


            // ==========================================
            // VALIDATION
            // ==========================================

            if (
                isNaN(spacing) ||
                spacing <= 0
            ) {

                alert(
                    "Spacing must be greater than 0."
                );

                return;
            }


            if (
                isNaN(minDotSize) ||
                minDotSize < 0
            ) {

                alert(
                    "Minimum dot size cannot be negative."
                );

                return;
            }


            if (
                isNaN(maxDotSize) ||
                maxDotSize <= 0
            ) {

                alert(
                    "Maximum dot size must be greater than 0."
                );

                return;
            }


            if (
                maxDotSize <
                minDotSize
            ) {

                alert(
                    "Maximum dot size must be greater than or equal to minimum dot size."
                );

                return;
            }


            if (
                isNaN(gamma) ||
                gamma <= 0
            ) {

                alert(
                    "Gamma must be greater than 0."
                );

                return;
            }


            if (
                isNaN(screenAngle)
            ) {

                alert(
                    "Screen angle must be a number."
                );

                return;
            }


            if (
                isNaN(gradientAngle)
            ) {

                alert(
                    "Gradient angle must be a number."
                );

                return;
            }


            dialog.settings = {

                spacing:
                    spacing,

                minDotSize:
                    minDotSize,

                maxDotSize:
                    maxDotSize,

                gamma:
                    gamma,

                screenAngle:
                    screenAngle,

                gradientAngle:
                    gradientAngle
            };


            dialog.close(1);
        };


    // ==================================================
    // SHOW WINDOW
    // ==================================================

    var result =
        dialog.show();


    if (result !== 1) {
        return null;
    }


    return dialog.settings;
}


// ======================================================
// UI INPUT ROW
// ======================================================

function addInputRow(
    parent,
    label,
    defaultValue,
    unit
) {

    var row =
        parent.add(
            "group"
        );

    row.orientation =
        "row";

    row.alignChildren =
        ["left", "center"];


    var labelText =
        row.add(
            "statictext",
            undefined,
            label
        );

    labelText.preferredSize.width =
        85;


    var input =
        row.add(
            "edittext",
            undefined,
            String(defaultValue)
        );

    input.characters =
        8;


    var unitText =
        row.add(
            "statictext",
            undefined,
            unit
        );

    unitText.preferredSize.width =
        30;


    return input;
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
        right -
        left;

    var height =
        top -
        bottom;


    return {

        left:
            left,

        top:
            top,

        right:
            right,

        bottom:
            bottom,

        width:
            width,

        height:
            height,

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
        diagonal /
        spacing
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
        Math.cos(
            radians
        );

    var sinAngle =
        Math.sin(
            radians
        );


    return {

        x:
            x *
            cosAngle -
            y *
            sinAngle,

        y:
            x *
            sinAngle +
            y *
            cosAngle
    };
}


// ======================================================
// PROJECT POINT ONTO GRADIENT
// ======================================================

function projectOntoGradient(
    x,
    y,
    bounds,
    gradientAngle
) {

    var radians =
        gradientAngle *
        Math.PI /
        180;


    var directionX =
        Math.cos(
            radians
        );

    var directionY =
        Math.sin(
            radians
        );


    var relativeX =
        x -
        bounds.centerX;

    var relativeY =
        y -
        bounds.centerY;


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
        Math.cos(
            radians
        );

    var directionY =
        Math.sin(
            radians
        );


    var halfWidth =
        bounds.width /
        2;

    var halfHeight =
        bounds.height /
        2;


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

        minimum:
            -radius,

        maximum:
            radius
    };
}


// ======================================================
// GRADIENT POSITION
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


    if (
        rampPosition <=
        stops[0].rampPoint
    ) {

        return getBrightness(
            stops[0].color
        );
    }


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

                localPosition =
                    0;

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
            stops.length -
            1
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
            color.gray /
            100
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
        red /
        255;

    var g =
        green /
        255;

    var b =
        blue /
        255;


    return (
        0.2126 *
        r +

        0.7152 *
        g +

        0.0722 *
        b
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
        cyan /
        100;

    var m =
        magenta /
        100;

    var y =
        yellow /
        100;

    var k =
        blackValue /
        100;


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
        0.2126 *
        r +

        0.7152 *
        g +

        0.0722 *
        b
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
//
