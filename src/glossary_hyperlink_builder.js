if (app.documents.length == 0) ErrorExit("Please open a document and try again.");
var myDoc = app.activeDocument;

main();
function main (){
	
	var selectedTextFrame = checkSelection();
	var listOfListOfWords = getWords(selectedTextFrame);
	var overWriteStylesBool = listOfListOfWords.overWriteStyles;
	//loop through the glossary words
	for (var i = 0; i < listOfListOfWords.length; i++) {
		var wordCollection = listOfListOfWords[i] instanceof Array ? listOfListOfWords[i] : [listOfListOfWords[i]];
		// create word collection name
		var name = "";
		for (var a = 0; a < wordCollection.length; a++){
			name += wordCollection[a].contents;
			if (a < wordCollection.length-1){
				name += " ";
			}
		}
		//returns an array of the hyperlink destinations
		var glossDestination = createGlossaryDestinations(wordCollection, name);
		//return the reference word selected by the user
		var referenceWord = findReferenceDestination(name);
		//if user selected skip word this section won't run
		if (referenceWord != null){
			//create and return destination for reference word
			var referenceDestination = createReferenceDestination(referenceWord, name);
			//create hyperlink from glossary word to reference word
			for (var k = 0; k < wordCollection.length; k++){
				hyperlinkGlossaryTerms(wordCollection[k],referenceDestination,overWriteStylesBool);
			}
			//create hyperlink from reference word to glossary word
			hyperlinkReferenceTerms(referenceWord, glossDestination,overWriteStylesBool);
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
	var words = [];
	var paragraphs = selecetedTextFrame.paragraphs.everyItem().getElements();
	for(var x = 0; x < paragraphs.length; x++){
		words[x] = paragraphs[x].words.everyItem().getElements();
	}
	if (words[0][0].contents == "Glossary" || words[0][0].contents == "glossary"){
		words.splice(0,1);
	}
	var words = getInputGlossaryWords(words);
	return words;
}

function getInputGlossaryWords(words){
	var glossaryWords = [];
	var wordCollection = [];
	for (var x = 0; x < words.length; x++){
		wordCollection.push(words[x][0]);
		glossaryWords.push(wordCollection[x]);
	}
	var myInputWindow = new Window("dialog", "Glossary words:");
	myInputWindow.orientation = 'column';
	var wordGroups = myInputWindow.add("group");
	wordGroups.orientation = 'column';
	//..........................
	// Create Word groups
	//..........................
	for (var x = 0; x < words.length; x++){
		var wordGroup = wordGroups.add("group");
		wordGroup.orientation = 'row';
		wordGroup.id = x;
		wordGroup.numWords = 1;
		wordGroup.wordString = words[x][0].contents;
		wordGroup.label = wordGroup.add("statictext",undefined,wordGroup.wordString);
		wordGroup.label.characters = 35;
		//..........................
		// Remove word button
		//..........................
		wordGroup.minusBtn = wordGroup.add("button",undefined,"-");
		wordGroup.minusBtn.enabled = false;
		wordGroup.minusBtn.onClick = minus_btn;
		function minus_btn (){
			var wordGroup = this.parent;
			wordGroup.numWords--;
			wordGroup.wordString = "";
			for (var i = 0; i < wordGroup.numWords; i++){
				wordGroup.wordString += words[wordGroup.id][i].contents + " ";
			}
			glossaryWords[wordGroup.id] = glossaryWords[wordGroup.id].splice(-1,1);
			wordGroup.label.text = wordGroup.wordString;
			if (wordGroup.numWords < words[wordGroup.id].length){
				wordGroup.addBtn.enabled = true;
			} else {
				this.enabled = false;
			}
			if(wordGroup.numWords < 2){
				this.enabled = false;
			}
			myInputWindow.layout.layout (true);
		}
		//..........................
		// Add word button
		//..........................
		wordGroup.addBtn = wordGroup.add("button",undefined,"+");
		wordGroup.addBtn.onClick = add_btn;
		function add_btn (){
			var wordGroup = this.parent;
			wordGroup.numWords++
			wordGroup.wordString = "";
			for (var i = 0; i < wordGroup.numWords; i++){
				wordGroup.wordString += words[wordGroup.id][i].contents + " ";
			}
			if (glossaryWords[wordGroup.id] instanceof Word){
				wordCollection = [];
				wordCollection.push(glossaryWords[wordGroup.id]);
				wordCollection.push(words[wordGroup.id][wordGroup.numWords-1]);
				glossaryWords[wordGroup.id] = wordCollection;
			} else {
				glossaryWords[wordGroup.id].push(words[wordGroup.id][wordGroup.numWords-1]);
			}
			wordGroup.label.text = wordGroup.wordString;
			if(wordGroup.numWords > 1){
				wordGroup.minusBtn.enabled = true;
			}
			if (wordGroup.numWords >= words[wordGroup.id].length){
				this.enabled = false;
			}
			myInputWindow.layout.layout (true);
		}
	}
	var overWriteStylesCheckbox = myInputWindow.add("checkbox",undefined,"Overwrite character styles");
	overWriteStylesCheckbox.value = true;
	var myButtonGroup = myInputWindow.add("group");
	myButtonGroup.alignment = "right";
	myButtonGroup.add("button",undefined,"OK");
	myButtonGroup.add("button",undefined,"Cancel");
	myInputWindow.layout.layout (true);
	if (myInputWindow.show()==1){
		glossaryWords.overWriteStyles = overWriteStylesCheckbox.value;
		return glossaryWords;
	} else {
		exit();
	}
}

function createGlossaryDestinations(wordCollection, name){
	var glossaryName = name + "_gl";
	// create hyperlink destination
	try {
		var increment = 1;
		var escape = false;
		while (escape === false){
			if (!myDoc.hyperlinkTextDestinations.itemByName(glossaryName).isValid){
				var glossDestinations = myDoc.hyperlinkTextDestinations.add(wordCollection[0]);
				glossDestinations.name = glossaryName;
				escape = true;
			} else {
				glossaryName += increment;
				increment++;
			}
		}
	} catch (e) {
		alert("create glossary destination failed try");
	}
	return glossDestinations;
}

function findReferenceDestination(name){
	var count = 0;
	var message = "";
	var namePlus = name + "[a-zA-Z0-9]*";

	app.findGrepPreferences = null;
	app.findGrepPreferences.findWhat = namePlus;
	var found = myDoc.findGrep();
	//if (found.length ==0 || found.length == 1){
	if (found.length ==0){
		message +="\nCould not find word " + name;
	} else {
		var escape = 0;
		while (escape==0){
			found[count].select();
			app.activeWindow.zoomPercentage = app.activeWindow.zoomPercentage;
			var referenceWord = getInput(found, name, count);
			if (referenceWord == "SKIPWORD"){
				escape = 1;
				return null;
			} else if (referenceWord == "NEXTWORD"){
				count++;
			} else if (referenceWord == "PREVWORD"){
				count--;
			} else {
				escape = 1;
			}
		}
		return referenceWord;
	}
	if (!message==""){
		alert(message);
	}
}

function getInput (found, name, count){
	var buttonPressed = "";
	var myInputWindow = new Window("dialog", "Hyperlinker");
	myInputWindow.orientation = 'column';

	var pageInfoGroup = myInputWindow.add("group");
	pageInfoGroup.orientation = 'row';
	try {
		var pageNum = found[count].parentTextFrames[0].parentPage.name;
	} catch (e) {
		var pageNum = "pasteboard";
	}
	pageInfoGroup.foundPage = pageInfoGroup.add("statictext",[0,0,500,50],name + " found on page " + pageNum);

	var wordInfoGroup = myInputWindow.add("group");
	wordInfoGroup.alignment = "left";
	wordInfoGroup.foundWord = wordInfoGroup.add("statictext",undefined,found[count].contents);
	wordInfoGroup.foundWord.characters = 16;

	var myButtonGroup = myInputWindow.add("group");
	myButtonGroup.alignment = "right";
	var myPrevButton = myButtonGroup.add("button",undefined,"Find Previous");
	myPrevButton.onClick = prevClick;
	if (count==0){
		myButtonGroup.children[0].enabled = false;
	}
	var myNextButton = myButtonGroup.add("button",undefined,"Find Next");
	myNextButton.onClick = nextClick;
	if (count >= found.length-1){
		myButtonGroup.children[1].enabled = false;
	}
	var mySkippButton = myButtonGroup.add("button",undefined,"Skip word");
	mySkippButton.onClick = skippClick;
	myButtonGroup.add("button",undefined,"OK");
	myButtonGroup.add("button",undefined,"Cancel");

	myInputWindow.layout.layout (true);

	if (myInputWindow.show()==1){
		if (buttonPressed != ""){
			return buttonPressed;
		} else {
			return found[count];
		}
	} else {
		exit();
	}
	
	function prevClick(){
		buttonPressed = "PREVWORD";
		myInputWindow.close(1);
	}

	function nextClick(){
		buttonPressed = "NEXTWORD";
		myInputWindow.close(1);
	}

	function skippClick(){
		buttonPressed = "SKIPWORD";
		myInputWindow.close(1);
	}
}

function createReferenceDestination(referenceWord, name){
	var referenceName = name;
	try {
		var increment = 1;
		var escape = false;
		while (escape === false){
			if (!myDoc.hyperlinkTextDestinations.itemByName(referenceName).isValid){
				var referenceDestination = myDoc.hyperlinkTextDestinations.add(referenceWord);
				referenceDestination.name = referenceName;
				escape = true;
			} else {
				referenceName += increment;
				increment++;
			}
		}
	} catch (e) {
		alert("create reference destination failed try");
	}
	return referenceDestination;
}

function hyperlinkGlossaryTerms(glossaryWord, referenceDestination, overWriteStylesBool){
	try {
		var hyperlinkSource = myDoc.hyperlinkTextSources.add(glossaryWord);
		var hyperlinkDestination = referenceDestination;
		var name = glossaryWord.contents + "_gl";
		var myHyperlink = myDoc.hyperlinks.add(hyperlinkSource, hyperlinkDestination);
		if (overWriteStylesBool){
			try {
				if (myDoc.characterStyles.item("hyperlink_glossary").isValid){
					glossaryWord.appliedCharacterStyle = "hyperlink_glossary";
				} else {
					var hyperlinkGlossaryStyle = myDoc.characterStyles.add();
					hyperlinkGlossaryStyle.name = "hyperlink_glossary";
					glossaryWord.appliedCharacterStyle = "hyperlink_glossary";
				}
			} catch (e) {
				alert("unable to apply character style to hyperlink");
			}
		}
		var increment = 1;
		var escape = false;
		while (escape === false){
			if (!myDoc.hyperlinks.itemByName(name).isValid){
				myHyperlink.name = name;
				escape = true;
			} else {
				name += increment;
			}
		}
	} catch (e) {
		alert("This glossary word is already hyperlinked");
	}	
}

function hyperlinkReferenceTerms(referenceWord, glossDestination, overWriteStylesBool){
	try {
		var hyperlinkSource = myDoc.hyperlinkTextSources.add(referenceWord);
		var hyperlinkDestination = glossDestination;
		var name = referenceWord.contents;
		var myHyperlink = myDoc.hyperlinks.add(hyperlinkSource, hyperlinkDestination);
		if (overWriteStylesBool){
			try {
				if (myDoc.characterStyles.item("hyperlink_keyword").isValid){
					referenceWord.appliedCharacterStyle = "hyperlink_keyword";
				} else {
					var hyperlinkKeywordStyle = myDoc.characterStyles.add();
					hyperlinkKeywordStyle.name = "hyperlink_keyword";
					referenceWord.appliedCharacterStyle = "hyperlink_keyword";
				}
			} catch (e) {
				alert("unable to apply character style to hyperlink");
			}
		}
		var increment = 1;
		var escape = false;
		while (escape === false){
			if (!myDoc.hyperlinks.itemByName(name).isValid){
				myHyperlink.name = name;
				escape = true;
			} else {
				name += increment;
			}
		}
	} catch (e) {
		alert("This reference word is already hyperlinked");
	}
}