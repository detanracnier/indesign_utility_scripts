if (app.documents.length == 0) ErrorExit("Please open a document and try again.");
var myDoc = app.activeDocument;

main();
function main (){
	
	var selecetedTextFrame = checkSelection();
	var words = getWords(selecetedTextFrame);

	//returns an array of the hyperlink destinations
	var glossDestinations = createGlossaryDestinations(words);
	//loop through the glossary words
	for (x = 0; x < words.length; x++){
		var name = words[x].contents;
		//return the reference word selected by the user
		var referenceWord = findReferenceDestination(words, x);
		//if user selected skip word this section won't run
		if (referenceWord != null){
			//create and return destination for reference word
			var referenceDestination = createReferenceDestination(referenceWord, name);
			//create hyperlink from glossary word to reference word
			hyperlinkGlossaryTerms(words[x],referenceDestination);
			//create hyperlink from reference word to glossary word
			hyperlinkReferenceTerms(referenceWord, glossDestinations[x]);
		}
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

function getWords(selecetedTextFrame){
	var words = selecetedTextFrame.paragraphs.everyItem().words.firstItem().getElements();
	return words;
}

function createGlossaryDestinations(words){
	var message = "";
	var hypTextDest = [];
	for(var x = 0; x < words.length; x++){
		var word = words[x];
		var name = word.contents + "_gl";

		try {
			if (!myDoc.hyperlinkTextDestinations.itemByName(name).isValid){
				hypTextDest[x] = myDoc.hyperlinkTextDestinations.add(word);
				hypTextDest[x].name = name;
			} else {
				message += "\nhyperlink destination with name: " + name + " already exists... skipping";
			}
		} catch (e) {}
	}
	if (!message==""){
		alert(message);
	}
	return hypTextDest;
}

function findReferenceDestination(words, wordNum){
	var count = 0;
	var message = "";
	var word = words[wordNum];
	var name = word.contents;

	app.findTextPreferences = null;
	app.findTextPreferences.findWhat = name;
	var found = myDoc.findText();
	if (found.length ==0 || found.length == 1){
		message +="\nCould not find word " + name;
	} else {
		for (var x = 0; x < found.length-1; x++){
			if (found[x].toString() == "[object Text]"){
				var parentTF = found[x].parentTextFrames[0];
				var wordsInFrame = parentTF.words.everyItem().getElements();
				for (var y = 0; y < wordsInFrame.length; y++){
					var checkWord = wordsInFrame[y].contents;
					if (checkWord.search(name) != -1){
						found[x] = wordsInFrame[y];
						break;
					}
				}
			}
		}
		var escape = 0;
		while (escape==0){
			found[count].select();
			app.activeWindow.zoomPercentage = app.activeWindow.zoomPercentage;
			var refWord = getInput(found, name, count);
			if (refWord == "SKIPWORD"){
				escape = 1;
				return null;
			} else if (refWord == "NEXTWORD"){
				count++;
			} else if (refWord == "PREVWORD"){
				count--;
			} else {
				escape = 1;
			}
		}
		return refWord;
	}
	if (!message==""){
		alert(message);
	}
}

function getInput (xWord, xName, num){
		var myInputWindow = new Window("dialog", "Hyperlinker");
		myInputWindow.orientation = 'column';

		var pageInfoGroup = myInputWindow.add("group");
		pageInfoGroup.orientation = 'row';
		pageInfoGroup.foundPage = pageInfoGroup.add("statictext",[0,0,500,50],xName + " found on page " + xWord[num].parentTextFrames[0].parentPage.name);

		var wordInfoGroup = myInputWindow.add("group");
		wordInfoGroup.alignment = "left";
		wordInfoGroup.foundWord = wordInfoGroup.add("statictext",undefined,xWord[num].contents);
		wordInfoGroup.foundWord.characters = 16;

		var myButtonGroup = myInputWindow.add("group");
		myButtonGroup.alignment = "right";
		var myPrevButton = myButtonGroup.add("button",undefined,"Find Previous");
		myPrevButton.onClick = prevClick;
		if (num==0){
			myButtonGroup.children[0].enabled = false;
		}
		var myNextButton = myButtonGroup.add("button",undefined,"Find Next");
		myNextButton.onClick = nextClick;
		if (num >= xWord.length-1){
			myButtonGroup.children[1].enabled = false;
		}
		var mySkippButton = myButtonGroup.add("button",undefined,"Skip word");
		mySkippButton.onClick = skippClick;
		myButtonGroup.add("button",undefined,"OK");
		myButtonGroup.add("button",undefined,"Cancel");
	
		myInputWindow.layout.layout (true);
	
		if (myInputWindow.show()==1){
			return xWord[num];
		} else {
			exit();
		}
		
		function prevClick(){
			xWord[num] = "PREVWORD";
			myInputWindow.close(1);
		}

		function nextClick(){
			xWord[num] = "NEXTWORD";
			myInputWindow.close(1);
		}

		function skippClick(){
			xWord[num] = "SKIPWORD";
			myInputWindow.close(1);
		}
}

function createReferenceDestination(referenceWord, name){
	try {
		if (referenceWord.toString() == "[object Word]"){
			if (!myDoc.hyperlinkTextDestinations.itemByName(name).isValid){
				var hypTextDest = myDoc.hyperlinkTextDestinations.add(referenceWord);
				hypTextDest.name = name;
				return hypTextDest;
			} else {
				message += "\nhyperlink destination with name: " + name + " already exists... skipping";
			}
		} else {
			alert("word did not return as word object");
		}
	} catch (e) {
		alert("create reference destination failed try");
	}
}

function hyperlinkGlossaryTerms(glossWord, refDest){
	var hyperlinkSource = myDoc.hyperlinkTextSources.add(glossWord);
	var hyperlinkDestination = refDest;
	var name = glossWord.contents + "_gl";
	var message = "";

	var myHyperlink = myDoc.hyperlinks.add(hyperlinkSource, hyperlinkDestination);
	try {
		glossWord.appliedCharacterStyle = "hyperlink_glossary";
	} catch (e) {
		alert("unable to apply character style to hyperlink");
	}
	try {
		if (!myDoc.hyperlinks.itemByName(name).isValid){
			myHyperlink.name = name;
		} else {
			message += "\nhyperlink with name: " + name + " already exists... naming hyperlink " + myHyperlink.name;
		}
	} catch (e) {}
	if (message!=""){
		alert(message);
	}	
}

function hyperlinkReferenceTerms(refWord, glosDest){
	var hyperlinkSource = myDoc.hyperlinkTextSources.add(refWord);
	var hyperlinkDestination = glosDest;
	var name = refWord.contents;
	var message = "";

	var myHyperlink = myDoc.hyperlinks.add(hyperlinkSource, hyperlinkDestination);
	try {
		refWord.appliedCharacterStyle = "hyperlink_glossary";
	} catch (e) {
		alert("unable to apply character style to hyperlink");
	}
	try {
		if (!myDoc.hyperlinks.itemByName(name).isValid){
			myHyperlink.name = name;
		} else {
			message += "\nhyperlink with name: " + name + " already exists... naming hyperlink " + myHyperlink.name;
		}
	} catch (e) {}
	if (message!=""){
		alert(message);
	}
}