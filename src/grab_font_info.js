var document = app.activeDocument;

var fontMultiset = countCharsInFonts(document);

// For each font, display its character count.
var fonts = document.fonts.everyItem().getElements();
//Define path and file name
var path = document.filePath;
var filename = 'fontlog.csv';
//Create File object
var file = new File(path + "/" + document.name.replace(".indd","") + "_" + filename);
file.encoding = 'UTF-8';
file.open('w');
for (var i = 0; i < fonts.length; i++) {
    var fontName = fonts[i].fullName;
    var fontType = getFontType(fonts[i]);
    file.write(fontName + ", " + fontType + "," + fontMultiset[fontName] + "\n");
}
file.close();


function countCharsInFonts(document) {
    // Create the font multiset.
    var fontMultiset = {
        add: function add(fontName, number) {
            if (this.hasOwnProperty(fontName)) {
                this[fontName] += number;
            }
            else {
                this[fontName] = number;
            }
        },
    };

    // For every textStyleRange in the document, add its applied font to the multiset.
    var stories = document.stories.everyItem().getElements();
    for (var i = 0; i < stories.length; i++) {
        var story = stories[i];
        var textStyleRanges = story.textStyleRanges.everyItem().getElements();
        for (var j = 0; j < textStyleRanges.length; j++) {
            fontMultiset.add(textStyleRanges[j].appliedFont.fullName, textStyleRanges[j].length);
        }
    }

    // For any fonts that aren't applied in the document, set the character count to 0.
    var fonts = document.fonts.everyItem().getElements();
    for (var i = 0; i < fonts.length; i++) {
        var fontName = fonts[i].fullName;
        if (!fontMultiset.hasOwnProperty(fontName)) {
            fontMultiset[fontName] = 0;
        }
    }

    return fontMultiset;
}

function getFontType (font)
{
    if (font.fontType == "1718894932"){
        return "ATC";
    }
    if (font.fontType == "1718895209"){
        return "Bitmap";
    }
    if (font.fontType == "1718895433"){
        return "CID";
    }
    if (font.fontType == "1718898499"){
        return "OCF";
    }
    if (font.fontType == "1718898502"){
        return "OpenType CFF";
    }
    if (font.fontType == "1718898505"){
        return "OpenType CID";
    }
    if (font.fontType == "1718898516"){
        return "OpenType TT";
    }
    if (font.fontType == "1718899796"){
        return "True Type";
    }
    if (font.fontType == "1718899761"){
        return "Type 1";
    }
    if (font.fontType == "1433299822"){
        return "Unknown FOnt Type";
    } else {
        return "Error returning font type";
    }
}