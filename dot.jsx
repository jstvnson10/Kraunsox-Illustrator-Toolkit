#target illustrator

if (app.documents.length === 0) {
    alert("Open an Illustrator document first.");
} else {

    var doc = app.activeDocument;

    // Create black
    var black = new RGBColor();
    black.red = 0;
    black.green = 0;
    black.blue = 0;

    // Get active artboard coordinates
    var artboardIndex = doc.artboards.getActiveArtboardIndex();
    var artboard = doc.artboards[artboardIndex].artboardRect;

    var left = artboard[0];
    var top = artboard[1];

    // Create a 50pt circle near the top-left of the artboard
    var dotSize = 50;

    var dot = doc.pathItems.ellipse(
        top - 100,
        left + 100,
        dotSize,
        dotSize
    );

    dot.filled = true;
    dot.fillColor = black;
    dot.stroked = false;

    // Select it so it's obvious where Illustrator created it
    dot.selected = true;

    app.redraw();

    alert("Dot created.");
}
