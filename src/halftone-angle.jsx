#target illustrator

// ==========================================
// KRAUNSOX HALFTONE TOOLKIT
// Screen Angle v0.4
// ==========================================

if (app.documents.length === 0) {

    alert("Open an Illustrator document first.");

} else {

    var doc = app.activeDocument;

    // --------------------------------------
    // SETTINGS
    // --------------------------------------

    var rows = 25;
    var columns = 35;

    var spacing = 18;

    var minDotSize = 1;
    var maxDotSize = 16;

    var gamma = Number(
        prompt(
            "Enter gamma:\n\n" +
            "0.5 = heavier\n" +
            "1.0 = normal\n" +
            "2.0 = lighter",
            "1.0"
        )
    );

    var angle = Number(
        prompt(
            "Enter screen angle:",
            "45"
        )
    );
