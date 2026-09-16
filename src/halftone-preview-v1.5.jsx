#target illustrator

// ======================================================
// KRAUNSOX HALFTONE TOOLKIT v1.5
// Vector Gradient Halftone Generator
// ======================================================


// ======================================================
// DEFAULT SETTINGS
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


// ======================================================
// PRESETS
// ======================================================

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


var PREVIEW_LAYER_NAME =
    "KRAUNSOX HALFTONE PREVIEW";


// ======================================================
// MAIN
// ======================================================

function main() {

    if (app.documents.length === 0) {
        alert("Open an Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;


    // --------------------------------------------------
    // VALIDATE SELECTION
    // --------------------------------------------------

    if (doc.selection.length !== 1) {

        alert(
            "Select exactly one gradient-filled vector object."
        );

        return;
    }


    var artwork = doc.selection[0];


    if (artwork.typename !== "PathItem") {

        alert(
            "KRAUNSOX Halftone v1.5 currently supports one PathItem."
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
            "The selected object must have a gradient fill."
        );

        return;
    }


    // --------------------------------------------------
    // OPEN UI
    // --------------------------------------------------

    var settings =
        showHalftoneDialog(
            doc,
            artwork
        );


    // --------------------------------------------------
    // CANCEL
    // --------------------------------------------------

    if (settings === null) {

        removePreviewLayer(doc);

        return;
    }


    // --------------------------------------------------
    // REMOVE TEMPORARY PREVIEW
    // --------------------------------------------------

    removePreviewLayer(doc);


    // --------------------------------------------------
    // CREATE FINAL OUTPUT
    // --------------------------------------------------

    var finalLayerName =
        "KRAUNSOX HALFTONE v1.5 - " +
        settings.preset +
        " - " +
        settings.shape;


    var outputLayer =
        renderHalftone(
            doc,
            artwork,
            settings,
            finalLayerName
        );


    // --------------------------------------------------
    // FINISH
    // --------------------------------------------------

    artwork.selected = false;

    app.redraw();


    alert(
        "KRAUNSOX HALFTONE CREATED\n\n" +

        "Preset: " +
        settings.preset +

        "\nShape: " +
        settings.shape +

        "\nSpacing: " +
        settings.spacing +
        " pt" +

        "\nSize Range: " +
        settings.minDotSize +
        " - " +
        settings.maxDotSize +
        " pt" +

        "\nGamma: " +
        settings.gamma +

        "\nScreen Angle: " +
        settings.screenAngle +
        " degrees" +

        "\nGradient Angle: " +
        settings.gradientAngle +
        " degrees"
    );
}


// ======================================================
// SCRIPT UI
// ======================================================

function showHalftoneDialog(
    doc,
    artwork
) {

    var dialog =
        new Window(
            "dialog",
            "KRAUNSOX Halftone Toolkit v1.5"
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
            "Vector Halftone Generator v1.5"
        );

    subtitle.alignment =
        "center";


    // ==================================================
    // PRESET PANEL
    // ==================================================

    var presetPanel =
        dialog.add(
            "panel",
            undefined,
            "Preset"
        );


    presetPanel.orientation =
        "row";

    presetPanel.alignChildren =
        ["left", "center"];

    presetPanel.margins =
        15;


    var presetLabel =
        presetPanel.add(
            "statictext",
            undefined,
            "Style:"
        );

    presetLabel.preferredSize.width =
        85;


    var presetDropdown =
        presetPanel.add(
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


    presetDropdown.selection =
        0;

    presetDropdown.preferredSize.width =
        150;


    // ==================================================
    // SHAPE PANEL
    // ==================================================

    var shapePanel =
        dialog.add(
            "panel",
            undefined,
            "Halftone Shape"
        );


    shapePanel.orientation =
        "row";

    shapePanel.alignChildren =
        ["left", "center"];

    shapePanel.margins =
        15;


    var shapeLabel =
        shapePanel.add(
            "statictext",
            undefined,
            "Shape:"
        );


    shapeLabel.preferredSize.width =
        85;


    var shapeDropdown =
        shapePanel.add(
            "dropdownlist",
            undefined,
            [
                "Circle",
                "Square",
                "Diamond",
                "Line"
            ]
        );


    shapeDropdown.selection =
        0;

    shapeDropdown.preferredSize.width =
        150;


    // ==================================================
    // MARK SETTINGS
    // ==================================================

    var markPanel =
        dialog.add(
            "panel",
            undefined,
            "Mark Settings"
        );


    markPanel.orientation =
        "column";

    markPanel.alignChildren =
        ["fill", "center"];

    markPanel.margins =
        15;


    var spacingField =
        addInputRow(
            markPanel,
            "Spacing:",
            DEFAULTS.spacing,
            "pt"
        );


    var minDotField =
        addInputRow(
            markPanel,
            "Min Size:",
            DEFAULTS.minDotSize,
            "pt"
        );


    var maxDotField =
        addInputRow(
            markPanel,
            "Max Size:",
            DEFAULTS.maxDotSize,
            "pt"
        );


    // ==================================================
    // TONE
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
    // ANGLES
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
        dialog.add("group");


    buttonGroup.orientation =
        "row";

    buttonGroup.alignment =
        "center";


    var resetButton =
        buttonGroup.add(
            "button",
            undefined,
            "Reset"
        );


    var previewButton =
        buttonGroup.add(
            "button",
            undefined,
            "Preview"
        );


    var cancelButton =
        buttonGroup.add(
            "button",
            undefined,
            "Cancel"
        );


    var generateButton =
        buttonGroup.add(
            "button",
            undefined,
            "Generate"
        );


    // ==================================================
    // READ SETTINGS
    // ==================================================

    function readSettingsFromUI() {

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


        // ----------------------------------------------
        // VALIDATION
        // ----------------------------------------------

        if (
            isNaN(spacing) ||
            spacing <= 0
        ) {

            alert(
                "Spacing must be greater than 0."
            );

            return null;
        }


        if (
            isNaN(minDotSize) ||
            minDotSize < 0
        ) {

            alert(
                "Minimum size cannot be negative."
            );

            return null;
        }


        if (
            isNaN(maxDotSize) ||
            maxDotSize <= 0
        ) {

            alert(
                "Maximum size must be greater than 0."
            );

            return null;
        }


        if (
            maxDotSize <
            minDotSize
        ) {

            alert(
                "Maximum size must be greater than or equal to minimum size."
            );

            return null;
        }


        if (
            isNaN(gamma) ||
            gamma <= 0
        ) {

            alert(
                "Gamma must be greater than 0."
            );

            return null;
        }


        if (
            isNaN(screenAngle)
        ) {

            alert(
                "Screen angle must be a number."
            );

            return null;
        }


        if (
            isNaN(gradientAngle)
        ) {

            alert(
                "Gradient angle must be a number."
            );

            return null;
        }


        if (
            shapeDropdown.selection ===
            null
        ) {

            alert(
                "Choose a halftone shape."
            );

            return null;
        }


        if (
            presetDropdown.selection ===
            null
        ) {

            alert(
                "Choose a preset."
            );

            return null;
        }


        return {

            preset:
                presetDropdown.selection.text,

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
    }


    // ==================================================
    // PRESET CHANGE
    // ==================================================

    presetDropdown.onChange =
        function() {

            if (
                presetDropdown.selection ===
                null
            ) {

                return;
            }


            var presetName =
                presetDropdown.selection.text;


            if (
                presetName ===
                "Custom"
            ) {

                return;
            }


            var preset =
                PRESETS[presetName];


            if (!preset) {
                return;
            }


            spacingField.text =
                String(
                    preset.spacing
                );


            minDotField.text =
                String(
                    preset.minDotSize
                );


            maxDotField.text =
                String(
                    preset.maxDotSize
                );


            gammaField.text =
                String(
                    preset.gamma
                );


            screenAngleField.text =
                String(
                    preset.screenAngle
                );


            gradientAngleField.text =
                String(
                    preset.gradientAngle
                );


            selectDropdownItem(
                shapeDropdown,
                preset.shape
            );
        };


    // ==================================================
    // RESET
    // ==================================================

    resetButton.onClick =
        function() {

            presetDropdown.selection =
                0;

            shapeDropdown.selection =
                0;


            spacingField.text =
                String(
                    DEFAULTS.spacing
                );


            minDotField.text =
                String(
                    DEFAULTS.minDotSize
                );


            maxDotField.text =
                String(
                    DEFAULTS.maxDotSize
                );


            gammaField.text =
                String(
                    DEFAULTS.gamma
                );


            screenAngleField.text =
                String(
                    DEFAULTS.screenAngle
                );


            gradientAngleField.text =
                String(
                    DEFAULTS.gradientAngle
                );


            removePreviewLayer(doc);
        };


    // ==================================================
    // PREVIEW
    // ==================================================

    previewButton.onClick =
        function() {

            var settings =
                readSettingsFromUI();


            if (
                settings === null
            ) {

                return;
            }


            try {

                // Delete old preview.

                removePreviewLayer(
                    doc
                );


                // Create replacement preview.

                renderHalftone(
                    doc,
                    artwork,
                    settings,
                    PREVIEW_LAYER_NAME
                );


                artwork.selected =
                    true;


                app.redraw();

            } catch (error) {

                alert(
                    "KRAUNSOX PREVIEW ERROR\n\n" +

                    "Message: " +
                    error.message +

                    "\nLine: " +
                    error.line +

                    "\nName: " +
                    error.name
                );
            }
        };


    // ==================================================
    // CANCEL
    // ==================================================

    cancelButton.onClick =
        function() {

            removePreviewLayer(
                doc
            );

            dialog.close(0);
        };


    // ==================================================
    // GENERATE
    // ==================================================

    generateButton.onClick =
        function() {

            var settings =
                readSettingsFromUI();


            if (
                settings === null
            ) {

                return;
            }


            dialog.settings =
                settings;


            dialog.close(1);
        };


    // ==================================================
    // SHOW WINDOW
    // ==================================================

    dialog.center();


    var result =
        dialog.show();


    if (
        result !== 1
    ) {

        removePreviewLayer(
            doc
        );

        return null;
    }


    return dialog.settings;
}


// ======================================================
// ADD UI INPUT ROW
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
// SELECT DROPDOWN ITEM
// ======================================================

function selectDropdownItem(
    dropdown,
    itemName
) {

    for (
        var i = 0;
        i < dropdown.items.length;
        i++
    ) {

        if (
            dropdown.items[i].text ===
            itemName
        ) {

            dropdown.selection =
                i;

            return;
        }
    }
}


// ======================================================
// RENDER HALFTONE
// ======================================================

function renderHalftone(
    doc,
    artwork,
    settings,
    layerName
) {

    var bounds =
        getArtworkBounds(
            artwork
        );


    var gradient =
        artwork.fillColor.gradient;


    var gridRadius =
        calculateGridRadius(
            bounds.width,
            bounds.height,
            settings.spacing
        );


    var outputLayer =
        doc.layers.add();


    outputLayer.name =
        layerName;


    var clippingGroup =
        outputLayer.groupItems.add();


    clippingGroup.name =
        "HALFTONE CLIPPING GROUP";


    var markGroup =
        clippingGroup.groupItems.add();


    markGroup.name =
        settings.shape.toUpperCase() +
        " HALFTONE MARKS";


    generateSpatialHalftone(
        markGroup,
        bounds,
        gradient,
        gridRadius,
        settings
    );


    createClippingMask(
        artwork,
        clippingGroup
    );


    app.redraw();


    return outputLayer;
}


// ======================================================
// REMOVE PREVIEW
// ======================================================

function removePreviewLayer(doc) {

    for (
        var i =
            doc.layers.length - 1;

        i >= 0;

        i--
    ) {

        if (
            doc.layers[i].name ===
            PREVIEW_LAYER_NAME
        ) {

            try {

                doc.layers[i].remove();

            } catch (error) {

                // Ignore cleanup errors.
            }
        }
    }


    app.redraw();
}


// ======================================================
// ARTWORK BOUNDS
// ======================================================

function getArtworkBounds(
    artwork
) {

    var bounds =
        artwork.geometricBounds;


    var left =
        bounds[0];

    var top =
        bounds[1];

    var right =
        bounds[2];

    var bottom =
        bounds[3];


    var width =
        right - left;

    var height =
        top - bottom;


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


    if (
        totalRange === 0
    ) {

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


    if (
        stops.length === 0
    ) {

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

        i <
        stops.length - 1;

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


            if (
                range === 0
            ) {

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
// COLOR BRIGHTNESS
// ======================================================

function getBrightness(
    color
) {

    // --------------------------------------------------
    // RGB
    // --------------------------------------------------

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


    // --------------------------------------------------
    // GRAYSCALE
    // --------------------------------------------------

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


    // --------------------------------------------------
    // CMYK
    // --------------------------------------------------

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

        0.2126 *
        r +

        0.7152 *
        g +

        0.0722 *
        b
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
// CALCULATE MARK SIZE
// ======================================================

function calculateMarkSize(
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
// CREATE BLACK COLOR
// ======================================================

function createBlack() {

    var black =
        new RGBColor();


    black.red =
        0;

    black.green =
        0;

    black.blue =
        0;


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

    if (
        size <= 0.01
    ) {

        return;
    }


    var mark;


    // ==================================================
    // CIRCLE
    // ==================================================

    if (
        shape ===
        "Circle"
    ) {

        mark =
            parent.pathItems.ellipse(
                centerY +
                    size / 2,

                centerX -
                    size / 2,

                size,
                size
            );


        mark.filled =
            true;

        mark.fillColor =
            color;

        mark.stroked =
            false;


        return;
    }


    // ==================================================
    // SQUARE
    // ==================================================

    if (
        shape ===
        "Square"
    ) {

        mark =
            parent.pathItems.rectangle(
                centerY +
                    size / 2,

                centerX -
                    size / 2,

                size,
                size
            );


        mark.filled =
            true;

        mark.fillColor =
            color;

        mark.stroked =
            false;


        return;
    }


    // ==================================================
    // DIAMOND
    // ==================================================

    if (
        shape ===
        "Diamond"
    ) {

        mark =
            parent.pathItems.rectangle(
                centerY +
                    size / 2,

                centerX -
                    size / 2,

                size,
                size
            );


        mark.filled =
            true;

        mark.fillColor =
            color;

        mark.stroked =
            false;


        mark.rotate(
            45
        );


        return;
    }


    // ==================================================
    // LINE
    // ==================================================

    if (
        shape ===
        "Line"
    ) {

        mark =
            parent.pathItems.add();


        var halfLength =
            size /
            2;


        mark.setEntirePath(
            [
                [
                    centerX -
                    halfLength,

                    centerY
                ],

                [
                    centerX +
                    halfLength,

                    centerY
                ]
            ]
        );


        mark.filled =
            false;


        mark.stroked =
            true;


        mark.strokeColor =
            color;


        mark.strokeWidth =
            Math.max(
                0.5,
                size *
                0.22
            );


        mark.rotate(
            screenAngle
        );


        return;
    }


    throw new Error(
        "Unknown halftone shape: " +
        shape
    );
}


// ======================================================
// GENERATE SPATIAL HALFTONE
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
        var row =
            -gridRadius;

        row <=
            gridRadius;

        row++
    ) {

        for (
            var column =
                -gridRadius;

            column <=
                gridRadius;

            column++
        ) {


            // ------------------------------------------
            // BASE GRID
            // ------------------------------------------

            var gridX =
                column *
                settings.spacing;


            var gridY =
                row *
                settings.spacing;


            // ------------------------------------------
            // SCREEN ROTATION
            // ------------------------------------------

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


            // ------------------------------------------
            // BOUNDS CHECK
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


                // --------------------------------------
                // GRADIENT POSITION
                // --------------------------------------

                var position =
                    getGradientPosition(
                        centerX,
                        centerY,
                        bounds,
                        settings.gradientAngle
                    );


                // --------------------------------------
                // BRIGHTNESS
                // --------------------------------------

                var brightness =
                    sampleGradient(
                        gradient,
                        position
                    );


                // --------------------------------------
                // DARKNESS
                // --------------------------------------

                var darkness =
                    1 -
                    brightness;


                // --------------------------------------
                // GAMMA
                // --------------------------------------

                var tone =
                    calculateTone(
                        darkness,
                        settings.gamma
                    );


                // --------------------------------------
                // MARK SIZE
                // --------------------------------------

                var markSize =
                    calculateMarkSize(
                        tone,
                        settings.minDotSize,
                        settings.maxDotSize
                    );


                // --------------------------------------
                // CREATE MARK
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

    if (
        value <
        minimum
    ) {

        return minimum;
    }


    if (
        value >
        maximum
    ) {

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
}        shape: "Circle"
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


var PREVIEW_LAYER_NAME =
    "KRAUNSOX HALFTONE PREVIEW";


// ======================================================
// MAIN
// ======================================================

function main() {

    if (app.documents.length === 0) {
        alert("Open an Illustrator document first.");
        return;
    }

    var doc = app.activeDocument;


    // --------------------------------------------------
    // VALIDATE SELECTION
    // --------------------------------------------------

    if (doc.selection.length !== 1) {

        alert(
            "Select exactly one gradient-filled vector object."
        );

        return;
    }


    var artwork = doc.selection[0];


    if (artwork.typename !== "PathItem") {

        alert(
            "KRAUNSOX Halftone v1.5 currently supports one PathItem."
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
            "The selected object must have a gradient fill."
        );

        return;
    }


    // --------------------------------------------------
    // OPEN UI
    // --------------------------------------------------

    var settings =
        showHalftoneDialog(
            doc,
            artwork
        );


    // --------------------------------------------------
    // CANCEL
    // --------------------------------------------------

    if (settings === null) {

        removePreviewLayer(doc);

        return;
    }


    // --------------------------------------------------
    // REMOVE TEMPORARY PREVIEW
    // --------------------------------------------------

    removePreviewLayer(doc);


    // --------------------------------------------------
    // CREATE FINAL OUTPUT
    // --------------------------------------------------

    var finalLayerName =
        "KRAUNSOX HALFTONE v1.5 - " +
        settings.preset +
        " - " +
        settings.shape;


    var outputLayer =
        renderHalftone(
            doc,
            artwork,
            settings,
            finalLayerName
        );


    // --------------------------------------------------
    // FINISH
    // --------------------------------------------------

    artwork.selected = false;

    app.redraw();


    alert(
        "KRAUNSOX HALFTONE CREATED\n\n" +

        "Preset: " +
        settings.preset +

        "\nShape: " +
        settings.shape +

        "\nSpacing: " +
        settings.spacing +
        " pt" +

        "\nSize Range: " +
        settings.minDotSize +
        " - " +
        settings.maxDotSize +
        " pt" +

        "\nGamma: " +
        settings.gamma +

        "\nScreen Angle: " +
        settings.screenAngle +
        " degrees" +

        "\nGradient Angle: " +
        settings.gradientAngle +
        " degrees"
    );
}


// ======================================================
// SCRIPT UI
// ======================================================

function showHalftoneDialog(
    doc,
    artwork
) {

    var dialog =
        new Window(
            "dialog",
            "KRAUNSOX Halftone Toolkit v1.5"
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
            "Vector Halftone Generator v1.5"
        );

    subtitle.alignment =
        "center";


    // ==================================================
    // PRESET PANEL
    // ==================================================

    var presetPanel =
        dialog.add(
            "panel",
            undefined,
            "Preset"
        );


    presetPanel.orientation =
        "row";

    presetPanel.alignChildren =
        ["left", "center"];

    presetPanel.margins =
        15;


    var presetLabel =
        presetPanel.add(
            "statictext",
            undefined,
            "Style:"
        );

    presetLabel.preferredSize.width =
        85;


    var preset
