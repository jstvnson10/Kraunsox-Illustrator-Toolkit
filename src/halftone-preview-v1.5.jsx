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


    var preset
