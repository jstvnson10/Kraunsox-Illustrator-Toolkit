#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT
// Multi-Shape Edition v1.3
// ======================================================

var DEFAULTS = {
    spacing: 14,
    minDotSize: 0,
    maxDotSize: 13,
    gamma: 1.0,
    screenAngle: 45,
    gradientAngle: 0,
    shape: "Circle"
};

var PRESETS = {

    "Fine Detail": {
        spacing: 8,
        minDotSize: 0,
        maxDotSize: 7,
        gamma: 1.1,
        screenAngle: 45,
        gradientAngle: 0,
        shape: "Circle"
    },

    "Classic Print": {
        spacing: 12,
        minDotSize: 0,
        maxDotSize: 11,
        gamma: 1.0,
        screenAngle: 45,
        gradientAngle: 0,
        shape: "Circle"
    },

    "Bold Streetwear": {
        spacing: 16,
        minDotSize: 0,
        maxDotSize: 18,
        gamma: 0.8,
        screenAngle: 45,
        gradientAngle: 0,
        shape: "Diamond"
    },

    "Grunge": {
        spacing: 11,
        minDotSize: 0,
        maxDotSize: 15,
        gamma: 0.65,
        screenAngle: 22.5,
        gradientAngle: 0,
        shape: "Line"
    }
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

    if (artwork.typename !== "PathItem") {
        alert("v1.3 currently supports one PathItem.");
        return;
    }

    if (!artwork.filled) {
        alert("The selected object must have a fill.");
        return;
    }

    if (artwork.fillColor.typename !== "GradientColor") {
        alert("v1.3 requires a gradient-filled object.");
        return;
    }


    // --------------------------------------------------
    // UI
    // --------------------------------------------------

    var settings = showHalftoneDialog();

    if (settings === null) {
        return;
    }


    // --------------------------------------------------
    // ARTWORK
    // --------------------------------------------------

    var bounds = getArtworkBounds(artwork);

    var gradient = artwork.fillColor.gradient;

    var gridRadius = calculateGridRadius(
        bounds.width,
        bounds.height,
        settings.spacing
    );


    // --------------------------------------------------
    // OUTPUT
    // --------------------------------------------------

    var outputLayer = doc.layers.add();

    outputLayer.name =
        "KRAUNSOX HALFTONE v1.3 - " +
        settings.shape;


    var clippingGroup =
        outputLayer.groupItems.add();

    clippingGroup.name =
        "HALFTONE CLIPPING GROUP";


    var markGroup =
        clippingGroup.groupItems.add();

    markGroup.name =
        settings.shape.toUpperCase() +
        " HALFTONE MARKS";


    // --------------------------------------------------
    // GENERATE
    // --------------------------------------------------

    generateSpatialHalftone(
        markGroup,
        bounds,
        gradient,
        gridRadius,
        settings
    );


    // --------------------------------------------------
    // MASK
    // --------------------------------------------------

    createClippingMask(
        artwork,
        clippingGroup
    );


    // --------------------------------------------------
    // FINISH
    // --------------------------------------------------

    artwork.selected = false;
    clippingGroup.selected = true;

    app.redraw();

    alert(
        "KRAUNSOX HALFTONE CREATED\n\n" +
        "Shape: " + settings.shape + "\n" +
        "Spacing: " + settings.spacing + " pt\n" +
        "Dot Range: " +
            settings.minDotSize +
            " - " +
            settings.maxDotSize +
            " pt\n" +
        "Gamma: " + settings.gamma + "\n" +
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

    var dialog = new Window(
        "dialog",
        "KRAUNSOX Halftone Toolkit v1.3"
    );

    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];
    dialog.spacing = 12;
    dialog.margins = 18;


    // --------------------------------------------------
    // TITLE
    // --------------------------------------------------

    var title = dialog.add(
        "statictext",
        undefined,
        "KRAUNSOX HALFTONE"
    );

    title.alignment = "center";


    var subtitle = dialog.add(
        "statictext",
        undefined,
        "Multi-Shape Vector Generator"
    );

    subtitle.alignment = "center";
// --------------------------------------------------
// PRESETS
// --------------------------------------------------

var presetPanel = dialog.add(
    "panel",
    undefined,
    "Preset"
);

presetPanel.orientation = "row";
presetPanel.alignChildren = ["left", "center"];
presetPanel.margins = 15;


var presetLabel = presetPanel.add(
    "statictext",
    undefined,
    "Style:"
);

presetLabel.preferredSize.width = 85;


var presetDropdown = presetPanel.add(
    "dropdownlist",
    undefined,
    [
        "Custom",
        "Fine Detail",
        "Classic Print",
        "Bold Streetwear",
        "Grunge"
    ]
);

presetDropdown.selection = 0;
presetDropdown.preferredSize.width = 140;

    // --------------------------------------------------
    // SHAPE
    // --------------------------------------------------

    var shapePanel = dialog.add(
        "panel",
        undefined,
        "Halftone Shape"
    );

    shapePanel.orientation = "row";
    shapePanel.alignChildren = ["left", "center"];
    shapePanel.margins = 15;


    var shapeLabel = shapePanel.add(
        "statictext",
        undefined,
        "Shape:"
    );

    shapeLabel.preferredSize.width = 85;


    var shapeDropdown = shapePanel.add(
        "dropdownlist",
        undefined,
        [
            "Circle",
            "Square",
            "Diamond",
            "Line"
        ]
    );

    shapeDropdown.selection = 0;
    shapeDropdown.preferredSize.width = 120;


    // --------------------------------------------------
    // DOT SETTINGS
    // --------------------------------------------------

    var dotPanel = dialog.add(
        "panel",
        undefined,
        "Mark Settings"
    );

    dotPanel.orientation = "column";
    dotPanel.alignChildren = ["fill", "center"];
    dotPanel.margins = 15;


    var spacingField = addInputRow(
        dotPanel,
        "Spacing:",
        DEFAULTS.spacing,
        "pt"
    );


    var minDotField = addInputRow(
        dotPanel,
        "Min Size:",
        DEFAULTS.minDotSize,
        "pt"
    );


    var maxDotField = addInputRow(
        dotPanel,
        "Max Size:",
        DEFAULTS.maxDotSize,
        "pt"
    );


    // --------------------------------------------------
    // TONE
    // --------------------------------------------------

    var tonePanel = dialog.add(
        "panel",
        undefined,
        "Tone"
    );

    tonePanel.orientation = "column";
    tonePanel.alignChildren = ["fill", "center"];
    tonePanel.margins = 15;


    var gammaField = addInputRow(
        tonePanel,
        "Gamma:",
        DEFAULTS.gamma,
        ""
    );


    // --------------------------------------------------
    // ANGLES
    // --------------------------------------------------

    var anglePanel = dialog.add(
        "panel",
        undefined,
        "Angles"
    );

    anglePanel.orientation = "column";
    anglePanel.alignChildren = ["fill", "center"];
    anglePanel.margins = 15;


    var screenAngleField = addInputRow(
        anglePanel,
        "Screen:",
        DEFAULTS.screenAngle,
        "deg"
    );


    var gradientAngleField = addInputRow(
        anglePanel,
        "Gradient:",
        DEFAULTS.gradientAngle,
        "deg"
    );

// ==================================================
// PRESET CHANGE
// ==================================================

presetDropdown.onChange = function() {

    if (
        presetDropdown.selection === null
    ) {
        return;
    }


    var presetName =
        presetDropdown.selection.text;


    // Custom doesn't overwrite anything.

    if (presetName === "Custom") {
        return;
    }


    var preset =
        PRESETS[presetName];


    if (!preset) {
        return;
    }


    spacingField.text =
        preset.spacing;

    minDotField.text =
        preset.minDotSize;

    maxDotField.text =
        preset.maxDotSize;

    gammaField.text =
        preset.gamma;

    screenAngleField.text =
        preset.screenAngle;

    gradientAngleField.text =
        preset.gradientAngle;


    // Match shape dropdown.

    for (
        var i = 0;
        i < shapeDropdown.items.length;
        i++
    ) {

        if (
            shapeDropdown.items[i].text ===
            preset.shape
        ) {

            shapeDropdown.selection =
                i;

            break;
        }
    }
};
    // --------------------------------------------------
    // BUTTONS
    // --------------------------------------------------

    var buttonGroup = dialog.add("group");

    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";


    var resetButton = buttonGroup.add(
        "button",
        undefined,
        "Reset"
    );


    var cancelButton = buttonGroup.add(
        "button",
        undefined,
        "Cancel",
        {
            name: "cancel"
        }
    );


    var generateButton = buttonGroup.add(
        "button",
        undefined,
        "Generate",
        {
            name: "ok"
        }
    );


    // --------------------------------------------------
    // RESET
    // --------------------------------------------------

    resetButton.onClick = function() {

        shapeDropdown.selection = 0;

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


    // --------------------------------------------------
    // GENERATE
    // --------------------------------------------------

    generateButton.onClick = function() {

        var spacing =
            parseFloat(spacingField.text);

        var minDotSize =
            parseFloat(minDotField.text);

        var maxDotSize =
            parseFloat(maxDotField.text);

        var gamma =
            parseFloat(gammaField.text);

        var screenAngle =
            parseFloat(screenAngleField.text);

        var gradientAngle =
            parseFloat(gradientAngleField.text);


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
                "Minimum size cannot be negative."
            );

            return;
        }


        if (
            isNaN(maxDotSize) ||
            maxDotSize <= 0
        ) {

            alert(
                "Maximum size must be greater than 0."
            );

            return;
        }


        if (
            maxDotSize < minDotSize
        ) {

            alert(
                "Maximum size must be greater than or equal to minimum size."
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


        if (isNaN(screenAngle)) {

            alert(
                "Screen angle must be a number."
            );

            return;
        }


        if (isNaN(gradientAngle)) {

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
                gradientAngle,

            shape:
                shapeDropdown.selection.text
        };


        dialog.close(1);
    };


    dialog.center();

    var result =
        dialog.show();


    if (result !== 1) {
        return null;
    }


    return dialog.settings;
}


// ======================================================
// UI ROW
// ======================================================

function addInputRow(
    parent,
    label,
    defaultValue,
    unit
) {

    var row =
        parent.add("group");

    row.orientation = "row";

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
// BOUNDS
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


    return {

        x:
            x * Math.cos(radians) -
            y * Math.sin(radians),

        y:
            x * Math.sin(radians) +
            y * Math.cos(radians)
    };
}


// ======================================================
// GRADIENT PROJECTION
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
        Math.cos(radians);

    var directionY =
        Math.sin(radians);


    var relativeX =
        x - bounds.centerX;

    var relativeY =
        y - bounds.centerY;


    return (
        relativeX * directionX +
        relativeY * directionY
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


    var halfWidth =
        bounds.width / 2;

    var halfHeight =
        bounds.height / 2;


    var radius =
        Math.abs(directionX) *
        halfWidth +

        Math.abs(directionY) *
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
        position * 100;


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
        (end - start) *
        amount
    );
}


// ======================================================
// BRIGHTNESS
// ======================================================

function getBrightness(color) {

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
// MARK SIZE
// ======================================================

function calculateMarkSize(
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
// BLACK
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
// CREATE HALFTONE MARK
// ======================================================

function createHalftoneMark(
    parent,
    centerX,
    centerY,
    size,
    color,
    shape,
    screenAngle
) {

    if (size <= 0.01) {
        return;
    }


    // --------------------------------------------------
    // CIRCLE
    // --------------------------------------------------

    if (shape === "Circle") {

        var circle =
            parent.pathItems.ellipse(
                centerY + size / 2,
                centerX - size / 2,
                size,
                size
            );

        circle.filled = true;
        circle.fillColor = color;
        circle.stroked = false;

        return;
    }


    // --------------------------------------------------
    // SQUARE
    // --------------------------------------------------

    if (shape === "Square") {

        var square =
            parent.pathItems.rectangle(
                centerY + size / 2,
                centerX - size / 2,
                size,
                size
            );

        square.filled = true;
        square.fillColor = color;
        square.stroked = false;

        return;
    }


    // --------------------------------------------------
    // DIAMOND
    // --------------------------------------------------

    if (shape === "Diamond") {

        var diamond =
            parent.pathItems.rectangle(
                centerY + size / 2,
                centerX - size / 2,
                size,
                size
            );

        diamond.filled = true;
        diamond.fillColor = color;
        diamond.stroked = false;

        diamond.rotate(
            45,
            true,
            true,
            true,
            true,
            Transformation.CENTER
        );

        return;
    }


    // --------------------------------------------------
    // LINE
    // --------------------------------------------------

    if (shape === "Line") {

        var line =
            parent.pathItems.add();

        var halfLength =
            size / 2;


        line.setEntirePath(
            [
                [
                    centerX - halfLength,
                    centerY
                ],
                [
                    centerX + halfLength,
                    centerY
                ]
            ]
        );


        line.filled =
            false;

        line.stroked =
            true;

        line.strokeColor =
            color;


        // Give darker tones thicker lines too.

        line.strokeWidth =
            Math.max(
                0.5,
                size * 0.22
            );


        line.rotate(
            screenAngle,
            true,
            true,
            true,
            true,
            Transformation.CENTER
        );

        return;
    }


    throw new Error(
        "Unknown halftone shape: " +
        shape
    );
}


// ======================================================
// GENERATE HALFTONE
// ======================================================

function generateSpatialHalftone(
    markGroup,
    bounds,
    gradient,
    gridRadius,
    settings
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
                settings.spacing;

            var gridY =
                row *
                settings.spacing;


            // Rotate the screen grid.

            var rotated =
                rotatePoint(
                    gridX,
                    gridY,
                    settings.screenAngle
                );


            var centerX =
                bounds.centerX +
                rotated.x;

            var centerY =
                bounds.centerY +
                rotated.y;


            if (
                centerX >= bounds.left &&
                centerX <= bounds.right &&
                centerY <= bounds.top &&
                centerY >= bounds.bottom
            ) {

                // --------------------------------------
                // SPATIAL TONE
                // --------------------------------------

                var position =
                    getGradientPosition(
                        centerX,
                        centerY,
                        bounds,
                        settings.gradientAngle
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
                        settings.gamma
                    );


                var markSize =
                    calculateMarkSize(
                        tone,
                        settings.minDotSize,
                        settings.maxDotSize
                    );


                // --------------------------------------
                // RENDER
                // --------------------------------------

                createHalftoneMark(
                    markGroup,
                    centerX,
                    centerY,
                    markSize,
                    black,
                    settings.shape,
                    settings.screenAngle
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
        "Message: " +
        error.message +
        "\nLine: " +
        error.line +
        "\nName: " +
        error.name
    );
}
