main();

function main(){
	//Make certain that user interaction (display of dialogs, etc.) is turned on.
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	if(app.documents.length != 0){
		//get ref to doc
		myDoc = app.activeDocument;
		textLayerName = "Text-fixed layout";
		backgroundLayerName = "Background-fixed layout"
		numberOfPagesInserted = 2;

		var imagePaths = inputWindow();
		var progressWindow = createProgressPalette();
		textLayerName = addAllTextToLayer(myDoc,textLayerName,progressWindow);
		addAllNonTextToLayer(myDoc,backgroundLayerName,textLayerName,progressWindow);
		if (setLayerVisibilityFalse(myDoc,textLayerName)===0){exit();}
		//insert # of pages to the beginning of page list
		insertPages(myDoc,numberOfPagesInserted,true,progressWindow);
		//insert # of pages to the end of page list
		insertPages(myDoc,numberOfPagesInserted,false,progressWindow);
		addCoverImage(myDoc,imagePaths,progressWindow);
		addBackCoverImage(myDoc,imagePaths,progressWindow);
		if(imagePaths.endsheetPath){addEndsheetImages(myDoc,imagePaths,progressWindow);}
		//exportBackgroundImages (myDoc);
		alert("Process Complete");
		progressWindow.close();		
	}else{
		alert("Please open a document and try again.");
	}
}

///..................................................
//	Progress palette & Input window
///..................................................

function inputWindow(){
	var imagePaths = {
		frontCoverPath:"",
		endsheetPath:"",
		backCoverPath:""
	}

	var myInputWindow = new Window("dialog", "Add cover file");
	//.....................................
	//......Add front cover controls.......
	//.....................................
	myInputWindow["addCoverButton"] = myInputWindow.add("button",[0,0,200,30], "Add cover image");
	myInputWindow.addCoverButton.onClick = function(){
		imagePaths.frontCoverPath = File.openDialog("Choose cover image");
		if (imagePaths.frontCoverPath){
			if (imagePaths.backCoverPath){myButtonGroup.okButton.enabled = true;}
			myInputWindow.frontCoverImagePath.text = imagePaths.frontCoverPath.fullName;
		}		
	};
	myInputWindow["frontCoverImagePath"] = myInputWindow.add("statictext",[0,0,600,50],"",{multiline:true});
	//.....................................
	//......Add endsheet controls.........
	//.....................................
	myInputWindow["addEndsheetButton"] = myInputWindow.add("button",[0,0,200,30], "Add endsheet image");
	myInputWindow.addEndsheetButton.onClick = function(){
		imagePaths.endsheetPath = File.openDialog("Choose endsheet image");
		if (imagePaths.endsheetPath){
			myInputWindow.endsheetImagePath.text = imagePaths.endsheetPath.fullName;
		}		
	};
	myInputWindow["endsheetImagePath"] = myInputWindow.add("statictext",[0,0,600,50],"",{multiline:true});
	//.....................................
	//......Add back cover controls........
	//.....................................
	myInputWindow["addBackCoverButton"] = myInputWindow.add("button",[0,0,200,30], "Add back cover image");
	myInputWindow.addBackCoverButton.onClick = function(){
		imagePaths.backCoverPath = File.openDialog("Choose endsheet image");
		if (imagePaths.backCoverPath){
			if (imagePaths.frontCoverPath){myButtonGroup.okButton.enabled = true;}
			myInputWindow.backCoverImagePath.text = imagePaths.backCoverPath.fullName;
		}		
	};
	myInputWindow["backCoverImagePath"] = myInputWindow.add("statictext",[0,0,600,50],"",{multiline:true});
	//.....................................
	//......OK and Cancel button controls..
	//.....................................
	var myButtonGroup = myInputWindow.add("group");
	myButtonGroup.alignment = "right";
	myButtonGroup["okButton"] = myButtonGroup.add("button",undefined,"OK");
	myButtonGroup.okButton.enabled = false;
	myButtonGroup.add("button",undefined,"Cancel");
	if(myInputWindow.show()===1){
		return imagePaths;
	} else {
		exit();
	}
}

function createProgressPalette(){
	var myProgressPalette = new Window("palette", "Progress", [50,50,600,400], {closeButton: false});
	with (myProgressPalette) {
		orientation = "Left";
		alignment = "Left";
		alignChildren = "Left";
	}
	myProgressPalette["processStatus"] = myProgressPalette.add("statictext",[20,20,550,380],"",{multiline:true});
	myProgressPalette.show();
	return myProgressPalette;
}

///..................................................
//	Create layers and organize objects to their layers
///..................................................

//Creates a layer and assignes all text frames to that layer
function addAllTextToLayer(myDoc,textLayerName,progressWindow){
	//create a new layer
	var myLayer = GetLayer(myDoc,textLayerName);
	//send every story to a new layer
	for(var i=0; i<myDoc.stories.length; i++){
		myStory = myDoc.stories.item(i);
		moveStoryToLayer(myStory,myLayer);
	}
	progressWindow.processStatus.text += "All found text moved to layer " + myLayer.name + "\r\n";
	return myLayer.name;
}
//Creates a layer and merges all layers into it that are not the text layer
function addAllNonTextToLayer(myDoc,backgroundLayerName,textLayerName,progressWindow){
	//create a new layer
	var myLayer = GetLayer(myDoc,backgroundLayerName);
	//send every layer besides the text layer to a new layer
	var selArr = [];
	var testit;
	var layerName;
	var layerArr = myDoc.layers.everyItem().name;
	for (i = 0; i < layerArr.length; i++) {
		if (layerArr[i] != textLayerName) {
			layerName = layerArr[i];
			selArr.push(myDoc.layers.item(layerName));
		}
	}
	myLayer.merge(selArr);
	myLayer.move(LocationOptions.AT_END);
	progressWindow.processStatus.text += "All found background objects to layer "+ myLayer.name + "\r\n";
}
//Moves text frame from a story object to a layer
function moveStoryToLayer(myStory,myLayer){
	var myTextFrame;
	for(var myCounter = myStory.textContainers.length-1; myCounter >= 0; myCounter --){
		myTextFrame = myStory.textContainers[myCounter];
		switch(myTextFrame.constructor.name){
			case "TextFrame":
				//I use a try statement here so it will continue moving text 
				//even if it finds text on a locked layer.
				try{
  					myTextFrame.move(myLayer);
				}
				catch (myError){}
				break;
			default:
				try{
					myTextFrame.parent.move(myLayer);
				}
				catch (myError){}
		}
	}
}
//Creates and returns a unique layer from provided name
function GetLayer(myDoc,name){
	var layer = myDoc.layers.item(name);
	var i = 0;
	while(layer.isValid){
		i++;
		layer = myDoc.layers.item(name+i);
	}
	if(i===0){
		layer = myDoc.layers.add({name:name});
	} else {
		layer = myDoc.layers.add({name:name+i});
	}
	return layer;

}
//Set layer visibility to false. Returns 1 = success, 0 = failure
function setLayerVisibilityFalse(myDoc,name){
	myLayer = myDoc.layers.item(name);
	if(myLayer.isValid){
		myLayer.visible = false;
		return 1;
	} else {
		alert("Set Layer Visibility FALSE Error\nNo layer named " + name + " was found.");
		return 0;
	}
}
//Set layer visibility to true. Returns 1 = success, 0 = failure
function setLayerVisibilityTrue(myDoc,name){
	myLayer = myDoc.layers.item(name);
	if(myLayer.isValid){
		myLayer.visible = true;
		return 1;
	} else {
		alert("Set Layer Visibility TRUE Error\nNo layer named " + name + " was found.");
		return 0;
	}
}

///..................................................
//	Insert pages/images and flatten background into JPG images
///..................................................

//Insert # of pages at the beginning or end of page list
function insertPages(myDoc,numberOfPagesInserted,boolInsertBefore,progressWindow){
	if (boolInsertBefore===true){
		var pageRef = myDoc.pages.item(0);
		for (var x = 0; x < numberOfPagesInserted; x++){
			var addedPage = myDoc.pages.add(LocationOptions.AT_BEGINNING, pageRef);
		}
		progressWindow.processStatus.text += "Inserted " + numberOfPagesInserted + " pages to beginning of page list\r\n";
		var mySection = addedPage.appliedSection;
		mySection.pageNumberStyle = PageNumberStyle.LOWER_ROMAN;
		myDoc.sections.add(pageRef,
			{
				continueNumbering:false,
				pageNumberStyle:PageNumberStyle.ARABIC,
				pageNumberStart:1,
				marker:"cover"
			});
		progressWindow.processStatus.text += "Renumbered page list\r\n";
	} else {
		var pageRef = myDoc.pages.item(myDoc.pages.length-1);
		var pageSpread = pageRef.parent;
		pageSpread.allowPageShuffle = true;
		for (var x = 0; x < numberOfPagesInserted; x++){
			var addedPage = myDoc.pages.add(LocationOptions.AT_END, pageRef,{});
			addedPage.appliedMaster = NothingEnum.NOTHING;
			addedPage.label = "Label " + x;
		}
		progressWindow.processStatus.text += "Inserted " + numberOfPagesInserted + " pages to end of page list\r\n";
	}
}
//Place an image on first page of pagelist
function addCoverImage(myDoc,imagePaths,progressWindow){
	var myHeight = myDoc.documentPreferences.pageHeight;
	var myWidth = myDoc.documentPreferences.pageWidth;
	var firstPage = myDoc.pages.item(0);
	var firstFrame = firstPage.rectangles.add({
		geometricBounds:[0, 0, myHeight, myWidth],
		topLeftCornerOption:CornerOptions.NONE,
		topRightCornerOption:CornerOptions.NONE,
		bottomLeftCornerOption:CornerOptions.NONE,
		bottomRightCornerOption:CornerOptions.NONE
	});
	if (!firstFrame.isValid){
		alert("Failed to create image frame for front cover");
		exit();
	} else {
		firstFrame.place(imagePaths.frontCoverPath);
		firstFrame.fit(FitOptions.FILL_PROPORTIONALLY);
		progressWindow.processStatus.text += "Created frame on first page and placed cover image\r\n";
	}
}
//Place an image on last page of pagelist
function addBackCoverImage(myDoc,imagePaths,progressWindow){
	var myHeight = myDoc.documentPreferences.pageHeight;
	var myWidth = myDoc.documentPreferences.pageWidth;
	var lastPage = myDoc.pages.item(myDoc.pages.length-1);
	var lastFrame = lastPage.rectangles.add({
		geometricBounds:[0, 0, myHeight, myWidth],
		topLeftCornerOption:CornerOptions.NONE,
		topRightCornerOption:CornerOptions.NONE,
		bottomLeftCornerOption:CornerOptions.NONE,
		bottomRightCornerOption:CornerOptions.NONE
	});
	if (!lastFrame.isValid){
		alert("Failed to create image frame for back cover");
		exit();
	} else {
		lastFrame.place(imagePaths.backCoverPath);
		lastFrame.fit(FitOptions.FILL_PROPORTIONALLY);
		progressWindow.processStatus.text += "Created frame on last page and placed back cover image\r\n";
	}
}
//Place an image on second page and second from last page of pagelist
function addEndsheetImages(myDoc,imagePaths,progressWindow){
	var myHeight = myDoc.documentPreferences.pageHeight;
	var myWidth = myDoc.documentPreferences.pageWidth;
	//front endsheet
	var secondPage = myDoc.pages.item(1);
	var secondFrame = secondPage.rectangles.add({
		geometricBounds:[0, 0, myHeight, myWidth],
		topLeftCornerOption:CornerOptions.NONE,
		topRightCornerOption:CornerOptions.NONE,
		bottomLeftCornerOption:CornerOptions.NONE,
		bottomRightCornerOption:CornerOptions.NONE
	});
	if (!secondFrame.isValid){
		alert("Failed to create image frame for front endsheet");
		exit();
	} else {
		secondFrame.place(imagePaths.endsheetPath);
		secondFrame.fit(FitOptions.FILL_PROPORTIONALLY);
		progressWindow.processStatus.text += "Created frame on second page and placed endsheet image\r\n";
	}
	//back endsheet
	var secondFromLastPage = myDoc.pages.item(myDoc.pages.length-2);
	var secondFromLastFrame = secondFromLastPage.rectangles.add({
		geometricBounds:[0, 0, myHeight, myWidth],
		topLeftCornerOption:CornerOptions.NONE,
		topRightCornerOption:CornerOptions.NONE,
		bottomLeftCornerOption:CornerOptions.NONE,
		bottomRightCornerOption:CornerOptions.NONE
	});
	secondFromLastFrame.move(undefined,[myWidth,0]);
	if (!secondFromLastFrame.isValid){
		alert("Failed to create image frame for front endsheet");
		exit();
	} else {
		secondFromLastFrame.place(imagePaths.endsheetPath);
		secondFromLastFrame.fit(FitOptions.FILL_PROPORTIONALLY);
		progressWindow.processStatus.text += "Created frame on second from last page and placed endsheet image\r\n";
	}
}
//Export each page as a JPG into a directory
function exportBackgroundImages (myDoc) {
	//Create directory for images
	var myImagesFolder = new Folder(myDoc.filePath + "/_fixed_layout_images");
	if (!myImagesFolder.exists) {
		myImagesFolder.create();
	}
	//Assign pages to array
	var pageList = myDoc.pages.everyItem().getElements();
	//Loop through pages for export
	for (var x = 0; x < pageList.length; x++){
		var imagePathName = new File(myImagesFolder + "/" +  "pageNumber_" + pageList[x].name + ".jpg");
		//Export Settings
		with (app.jpegExportPreferences) {
				exportingSpread = false;
				jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
				pageString = pageList[x].name;
				exportResolution = 300; // The export resolution expressed as a real number instead of an integer. (Range: 1.0 to 2400.0)
				antiAlias = true; //  If true, use anti-aliasing for text and vectors during export
				embedColorProfile = false; // True to embed the color profile, false otherwise
				jpegColorSpace = JpegColorSpaceEnum.RGB; // One of RGB, CMYK or GRAY
				jpegQuality = JPEGOptionsQuality.HIGH; // The compression quality: LOW / MEDIUM / HIGH / MAXIMUM
				jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; // The rendering style: BASELINE_ENCODING or PROGRESSIVE_ENCODING
				simulateOverprint = false; // If true, simulates the effects of overprinting spot and process colors in the same way they would occur when printing
				useDocumentBleeds = false; // If true, uses the document's bleed settings in the exported JPEG.
			}
		//Export Images
		myDoc.exportFile(ExportFormat.JPG, imagePathName, false);
	}
}