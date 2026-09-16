#target illustrator

alert("Starting UI test.");

var win = new Window(
    "dialog",
    "KRAUNSOX UI TEST"
);

win.orientation = "column";
win.alignChildren = "fill";

win.add(
    "statictext",
    undefined,
    "If you can see this, ScriptUI is working."
);

var button = win.add(
    "button",
    undefined,
    "Close",
    {
        name: "ok"
    }
);

win.center();
win.show();
