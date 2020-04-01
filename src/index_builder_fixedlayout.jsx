var myDoc = app.activeDocument;

main();
function main (){
	
	var selectedTextFrame = checkSelection();
    var pageNumberWords = getPageNumberWords(selectedTextFrame);
    for (var i = 0; i < pageNumberWords.length; i++){
        var pageNumberWord = pageNumberWords[i];
        var hyperlinkDestination = createHyperlinkDestination(pageNumberWord);
        hyperlinkIndexPageNumber(pageNumberWord, hyperlinkDestination);
    }
}

function checkSelection(){
	if (app.selection.length > 1){
		alert("Select one thing at a time.");
		exit();
	} else if (app.selection[0] instanceof TextFrame){
		return app.selection[0];
	} else {
		alert("Select a textframe");
		exit();
	}
}

function getPageNumberWords(selectedTextFrame){
    var wordsCollection = selectedTextFrame.words.everyItem().getElements();
    var regex = "[0-9]*";
    var pageNumberWords = [];
    for (var x = 0; x < wordsCollection.length; x++){
        var word = wordsCollection[x];
        var found = word.contents.match(regex);
        if (found > 0){
            var pageNumber = [word, found];
            pageNumberWords.push(pageNumber);
        }
    }
    return pageNumberWords;
}

function createHyperlinkDestination(pageNumberWord){
    var pageNumber = pageNumberWord[1][0];
    var page = myDoc.pages.item(pageNumber);
    var desinationName = "page number " + pageNumber
    try {
        var hyperlinkDestination = myDoc.hyperlinkPageDestinations.add(page);
        //create desination name
        var increment = 1;
		var escape = false;
        var oldDestinationName = desinationName;
        while (escape === false){
            if (myDoc.hyperlinkPageDestinations.itemByName(desinationName).isValid){
                desinationName = oldDestinationName + "-" + increment;
                increment++;
            } else {
                hyperlinkDestination.name = desinationName;
                escape = true;
            }
        }
        return hyperlinkDestination;
    } catch (e) {
        alert("Failed to create hyperlink destination for " + pageNumber + "\nexiting program");
        exit();
    }
}

function hyperlinkIndexPageNumber(pageNumberWord, hyperlinkDestination){
    var wordToLink = pageNumberWord[0];
    var pageNumber = pageNumberWord[1][0];
    var hyperlinkName = "Page " + pageNumber;
    try {
		var hyperlinkSource = myDoc.hyperlinkTextSources.add(wordToLink);
		var myHyperlink = myDoc.hyperlinks.add(hyperlinkSource, hyperlinkDestination);
		var increment = 1;
		var escape = false;
        var oldHyperlinkName = hyperlinkName;
		while (escape === false){
			if (myDoc.hyperlinks.itemByName(hyperlinkName).isValid){
				hyperlinkName = oldHyperlinkName + "-" + increment;
                increment++;
			} else {
                myHyperlink.name = hyperlinkName;
				escape = true;
			}
		}
	} catch (e) {
		alert("Failed to create hyperlink for " + pageNumber + "\nexiting program");
        exit();
	}
}