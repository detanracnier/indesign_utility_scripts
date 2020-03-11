main();

function main() {

	if (app.documents.length > 0) {
		var myDoc = app.activeDocument;
		var myDocumentName = myDoc.name.slice (0, -5);
		//Gather and set preferences
		/*
		userHoriz = myDoc.viewPreferences.horizontalMeasurementUnits;
		userVert = myDoc.viewPreferences.verticalMeasurementUnits;
		userOrigin = myDoc.viewPreferences.rulerOrigin;
		myDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.inches;
		myDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.inches;
		myDoc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
		*/

		var tocPage = "";
		var isbnArray = new Array ();
		var hiresPreset = "";
		var loresPreset = "";
		var exportPDFCheck = false;
		var exportHiRes = false;
		var exportLoRes = false;
		var exportImagesCheck = false;
		var isbnFormatFolderName = ["Hardcover", "Paperback", "EBook"];
		var coverPathCheck = false;
		var coverPathObject;
		var reportLog = "";

		var userSpreadPages = inputWindow()
		if (exportPDFCheck==true || exportImagesCheck==true) {
			var myExportFolder = chooseExportFolder();
			createExportFolders(myExportFolder);
			if (exportPDFCheck==true) {
				exportPDFwithPresets();
			}
			if (exportImagesCheck==true) {
				exportImages();
			}
			alert("Script complete\nReport:" + reportLog);
			//Set preferences back to normal
			/*
			myDoc.viewPreferences.horizontalMeasurementUnits = userHoriz;
			myDoc.viewPreferences.verticalMeasurementUnits = userVert;
			myDoc.viewPreferences.rulerOrigin = userOrigin;
			*/
		}

		//Input Window
		function inputWindow () {
			var myPresets = app.pdfExportPresets.everyItem().name;
			myPresets.unshift("- Select Preset -");

			//Export PDF Input
			var myInputWindow = new Window("dialog", "Export Settings");
				myInputWindow.orientation = 'column';
				myInputWindow.alignment = "Left";
				myInputWindow.alignChildren = "Left";

				//Export PDFs check
				var exportPDFCheckbox = myInputWindow.add("checkbox",undefined,"Export PDF");
				exportPDFCheckbox.value = true;
				exportPDFCheckbox.onClick = function (){
					if (this.value==true){
						myPDFExportGroup.show();
					} else {
						myPDFExportGroup.hide();
					}
				}
				
				//Export PDFs Presets
				var myPDFExportGroup = myInputWindow.add("group");
					myPDFExportGroup.orientation = 'column';
					myPDFExportGroup.alignment = "Left";
					myPDFExportGroup.alignChildren = "Left";

					var	myHiResGroup = myPDFExportGroup.add("group");
						var exportHiResPDFCheck = myHiResGroup.add("checkbox",undefined,"Export:");
						exportHiResPDFCheck.value = true;
						myHiResGroup.add("statictext",undefined,"Hi-Res PDF Preset:");
						var myPDFPresetDropdownHiRes = myHiResGroup.add('dropdownlist',undefined,undefined,{items:myPresets});
						myPDFPresetDropdownHiRes.selection=0;
					
					var myLoResGroup = myPDFExportGroup.add("group");
						var exportLoResPDFCheck = myLoResGroup.add("checkbox",undefined,"Export:");
						exportLoResPDFCheck.value = true;
						myLoResGroup.add("statictext",undefined,"Lo-Res PDF Preset:");
						var myPDFPresetDropdownLoRes = myLoResGroup.add('dropdownlist',undefined,undefined,{items:myPresets});
						myPDFPresetDropdownLoRes.selection=0;
				//Add Cover file	
					var addCoverFileGroup = myPDFExportGroup.add("group");
						addCoverFileGroup.orientation = "row";
					var coverPath = "";
						addCoverFileGroup.addButton = addCoverFileGroup.add("button",undefined,"Add Cover File to Lo-Res");
						addCoverFileGroup.addButton.onClick = addCover_btn;
						addCoverFileGroup.removeButton = addCoverFileGroup.add("button",undefined,"-");
						addCoverFileGroup.removeButton.onClick = removeCover_btn;
						addCoverFileGroup.coverpathlabel = addCoverFileGroup.add("statictext",undefined,"-No Cover File Selected-");
						addCoverFileGroup.coverpathlabel.characters = 60;

				//Export images check
				var exportImagesCheckbox = myInputWindow.add("checkbox",undefined,"Export Product Images");
				exportImagesCheckbox.value = true;
				exportImagesCheckbox.onClick = function (){
					if (this.value==true){
						myImageInputGroup.show();
						myISBNInputGroup.show();
					} else {
						myImageInputGroup.hide();
						myISBNInputGroup.hide();
					}
				}

				//Page Number Input
				var myImageInputGroup = myInputWindow.add("group");
					myImageInputGroup.orientation = 'column';

				var myTocInput = myImageInputGroup.add("group");
					myTocInput.orientation = 'row';
					myTocInput.alignment = "Left";
					myTocInput.pgnumlabel = myTocInput.add("statictext",undefined,"TOC Page #")
					myTocInput.pgnuminput = myTocInput.add("edittext",undefined,"3");
					myTocInput.pgnuminput.characters = 3;

				for (var n = 0; n < 3; n++){
					createSpreadPageGroup(myImageInputGroup);
				}

				//ISBN Input
				var myISBNInputGroup = myInputWindow.add("group");
					myISBNInputGroup.add("statictext",undefined,"Hardcover ISBN:");

					var myHCISBNInput = myISBNInputGroup.add("edittext",undefined);
						myHCISBNInput.characters = 17;

					myISBNInputGroup.add("statictext",undefined,"Paperback ISBN:");

					var myPBISBNInput = myISBNInputGroup.add("edittext",undefined);
						myPBISBNInput.characters = 17;

					myISBNInputGroup.add("statictext",undefined,"EBook ISBN:");

					var myEBISBNInput = myISBNInputGroup.add("edittext",undefined);
						myEBISBNInput.characters = 17;

				//Buttons
				var myButtonGroup = myInputWindow.add("group");
					myButtonGroup.alignment = "right";
					myButtonGroup.add("button",undefined,"OK");
					myButtonGroup.add("button",undefined,"Cancel");

				//If settings exist set input fields to last used values
				var settingsExist = app.activeDocument.extractLabel('preMediaExportInterior_SettingsTrue');
				if (settingsExist=="true"){
					//Set PDF Checkbox
					exportPDFCheckbox.value = app.activeDocument.extractLabel('preMediaExportInterior_exportPDFCheck');
					if (!exportPDFCheckbox.value){
						myPDFExportGroup.hide();
					}
					//Set Images Checkbox
					exportImagesCheckbox.value = app.activeDocument.extractLabel('preMediaExportInterior_exportImagesCheck');
					if (!exportImagesCheckbox.value){
						myImageInputGroup.hide();
						myISBNInputGroup.hide();
					}
					//Set Presets
					myPDFPresetDropdownHiRes.selection.text = app.activeDocument.extractLabel('preMediaExportInterior_hiresPreset');
					myPDFPresetDropdownLoRes.selection.text = app.activeDocument.extractLabel('preMediaExportInterior_loresPreset');
					//Set toc and spread page numbers
					myTocInput.pgnuminput.text = app.activeDocument.extractLabel('preMediaExportInterior_tocPage');
					var storedSpreadPages = app.activeDocument.extractLabel('preMediaExportInterior_spreadPages');
					storedSpreadPages = storedSpreadPages.split("*");
					storedSpreadPages.length = storedSpreadPages.length-1;
					for (var x = 3; x < storedSpreadPages.length; x++){
						createSpreadPageGroup(myImageInputGroup);
					}
					for (var x = 0; x < storedSpreadPages.length; x++){
						myImageInputGroup.children[x+1].pgnuminput.text = storedSpreadPages[x];
					}
					//Set ISBN numbers
					myHCISBNInput.text = app.activeDocument.extractLabel('preMediaExportInterior_hcISBN');
					myPBISBNInput.text = app.activeDocument.extractLabel('preMediaExportInterior_pbISBN');
					myEBISBNInput.text = app.activeDocument.extractLabel('preMediaExportInterior_ebISBN');
				}

				myInputWindow.layout.layout (true);
			
			//Input Window close
			if (myInputWindow.show()==1){

				//set variables after OK
				app.activeDocument.insertLabel('preMediaExportInterior_SettingsTrue', 'true');

				tocPage = myTocInput.pgnuminput.text;
				app.activeDocument.insertLabel('preMediaExportInterior_tocPage', tocPage);

				var spreadPages = new Array ();
				for (var x = 1; x < myImageInputGroup.children.length; x++){
					spreadPages[x-1] = myImageInputGroup.children[x].pgnuminput.text;
				}
				var spreadPagesString = "";
				for (var x = 0; x < spreadPages.length; x++){
                    spreadPagesString += spreadPages[x] + "*";
                }
                app.activeDocument.insertLabel('preMediaExportInterior_spreadPages', spreadPagesString);
				exportPDFCheck = exportPDFCheckbox.value;
				exportHiRes = exportHiResPDFCheck.value;
				exportLoRes = exportLoResPDFCheck.value;
				var exportPDFCheckString = "";
				if (exportPDFCheckbox.value){exportPDFCheckString="true";} else {exportPDFCheckString.value="false";};
				app.activeDocument.insertLabel('preMediaExportInterior_exportPDFCheck', exportPDFCheckString);

				exportImagesCheck = exportImagesCheckbox.value;
				var exportImagesCheckString = "";
				if (exportImagesCheckbox.value){exportImagesCheckString="true";} else {exportImagesCheckString.value="false";};
				app.activeDocument.insertLabel('preMediaExportInterior_exportImagesCheck', exportImagesCheckString);

				hiresPreset = myPDFPresetDropdownHiRes.selection.text;
				app.activeDocument.insertLabel('preMediaExportInterior_hiresPreset', hiresPreset);

				loresPreset = myPDFPresetDropdownLoRes.selection.text;
				app.activeDocument.insertLabel('preMediaExportInterior_loresPreset', loresPreset);

				var hcisbnHyphen = myHCISBNInput.text;
				isbnArray[0] = hcisbnHyphen.replace(/-/g,"");
				app.activeDocument.insertLabel('preMediaExportInterior_hcISBN', isbnArray[0]);

				var pbisbnHyphen = myPBISBNInput.text;
				isbnArray[1] = pbisbnHyphen.replace(/-/g,"");
				app.activeDocument.insertLabel('preMediaExportInterior_pbISBN', isbnArray[1]);

				var ebookisbnHyphen = myEBISBNInput.text;
				isbnArray[2] = ebookisbnHyphen.replace(/-/g,"");
				app.activeDocument.insertLabel('preMediaExportInterior_ebISBN', isbnArray[2]);

				coverPathObject = coverPath;

				return spreadPages;
			} else {
				exit();
			}

			//Create spread page number input field
			function createSpreadPageGroup (myImageInputGroup) {
				var mySpreadPageGroup = myImageInputGroup.add("group");
					mySpreadPageGroup.pgnumlabel = mySpreadPageGroup.add("statictext",undefined,"Start of Spread Page #");
					mySpreadPageGroup.pgnuminput = mySpreadPageGroup.add("edittext",undefined,"");
					mySpreadPageGroup.pgnuminput.characters = 3;
					mySpreadPageGroup.plus = mySpreadPageGroup.add("button",undefined, "+");
					mySpreadPageGroup.plus.onClick = add_btn;
					mySpreadPageGroup.minus = mySpreadPageGroup.add("button",undefined, "-");
					mySpreadPageGroup.minus.onClick = minus_btn;
					myInputWindow.layout.layout (true);
			}

			function add_btn() {
				createSpreadPageGroup(myImageInputGroup);
			}

			function minus_btn () {
				if (2 < myImageInputGroup.children.length) {
					myImageInputGroup.remove (this.parent);
					myInputWindow.layout.layout (true);
				}
			}

			function addCover_btn () {
				coverPath = File.openDialog("Choose cover PDF");
				if (coverPath){
					coverPathCheck = true;
				}
				addCoverFileGroup.coverpathlabel.text = coverPath.fullName;
			}

			function removeCover_btn () {
				coverPath = "";
				if (!coverPath){
					coverPathCheck = false;
				}
				addCoverFileGroup.coverpathlabel.text = "-No Cover File Selected-";
			}
		}

		//Chose Export Folder
		function chooseExportFolder () {
			var selectedExportFolder = Folder.selectDialog ("Choose series folder to export to");
			if (selectedExportFolder){
				return selectedExportFolder;
			} else {
				exit();
			}
		}

		//Create Directories
		function createExportFolders (myExportFolder) {
			var myImagesFolder = new Folder(myExportFolder + "/_Marketing_Images");
			if (!myImagesFolder.exists) {
			myImagesFolder.create();
			}
			
			if(isbnArray[0]!=""){
				var myHardCoverImagesFolder = new Folder(myExportFolder + "/_Marketing_Images/_Hardcover");
				if (!myHardCoverImagesFolder.exists) {
				myHardCoverImagesFolder.create();
				}
			}
			if(isbnArray[1]!=""){
				var myPaperBackImagesFolder = new Folder(myExportFolder + "/_Marketing_Images/_Paperback");
				if (!myPaperBackImagesFolder.exists) {
				myPaperBackImagesFolder.create();
				}
			}
			if(isbnArray[2]!=""){
				var myEbookImagesFolder = new Folder(myExportFolder + "/_Marketing_Images/_EBook");
				if (!myEbookImagesFolder.exists) {
				myEbookImagesFolder.create();
				}
			}	
			if (exportPDFCheck==true) {
				var myLoResFolder = new Folder(myExportFolder + "/_LoRes");
				if (!myLoResFolder.exists) {
					myLoResFolder.create();
				}
				if(isbnArray[0]!=""){
					var myHardCoverLoresFolder = new Folder(myExportFolder + "/_LoRes/_Hardcover");
					if (!myHardCoverLoresFolder.exists) {
					myHardCoverLoresFolder.create();
					}
				}
				if(isbnArray[1]!=""){
					var myPaperBackLoresFolder = new Folder(myExportFolder + "/_LoRes/_Paperback");
					if (!myPaperBackLoresFolder.exists) {
					myPaperBackLoresFolder.create();
					}
				}
				if(isbnArray[2]!=""){
					var myEbookLoresFolder = new Folder(myExportFolder + "/_LoRes/_EBook");
					if (!myEbookLoresFolder.exists) {
					myEbookLoresFolder.create();
					}
				}
			}
		}

		//Export PDFs
		function exportPDFwithPresets () {
			var userPagePref = app.pdfExportPreferences.pageRange;
			app.pdfExportPreferences.pageRange = "ALL_PAGES"; 
			//Export Hi-res
			if (exportHiRes){
				if (hiresPreset!="" && hiresPreset!="- Select Preset -"){
					var myFileHR = new File(myExportFolder + "/" + myDocumentName + ".pdf");
					myDoc.exportFile(ExportFormat.pdfType, myFileHR, false, hiresPreset);
				} else {
					reportLog += "\nNo Hi-res preset was selected: skipping Hi-res export";
				}
			}
			//Export Lo-res
			if (exportLoRes){
				if (loresPreset!="" && loresPreset!="- Select Preset -"){
					//Add Cover file check
					if (coverPathCheck == true){
						var docRef = app.activeDocument;
						with (docRef.documentPreferences) {
							allowPageShuffle = false;
						}
						var pageRef = docRef.pages.item(0);
						docRef.pages.add(LocationOptions.BEFORE, pageRef);
						firstPage = docRef.pages[0];
						myDoc.sections.add(pageRef,
							{
								continueNumbering:false,
								pageNumberStart:1,
								marker:"cover"
							});

						firstFrame = firstPage.rectangles.add({geometricBounds:[0, 0, docRef.documentPreferences.pageHeight, docRef.documentPreferences.pageWidth],topLeftCornerOption:CornerOptions.NONE,topRightCornerOption:CornerOptions.NONE,bottomLeftCornerOption:CornerOptions.NONE,bottomRightCornerOption:CornerOptions.NONE});
						if (!firstFrame.isValid){
							alert("Failed to create image frame");
						} else {
							firstFrame.place(coverPathObject);
							firstFrame.fit(FitOptions.FILL_PROPORTIONALLY);
						}
					}
					//Export file
					var myFileLR = new File(myExportFolder + "/_LoRes/" + myDocumentName + ".pdf");
					myDoc.exportFile(ExportFormat.pdfType, myFileLR, false, loresPreset);

					//Remove cover from indesign file
					if (coverPathCheck == true){
						firstPage.remove();
					}

					//Move Lo-res
					for (var k = 0; k < 3; k++) {
						if (isbnArray[k]!=""){
							myFileLR.copy(myExportFolder + "/_LoRes/_" + isbnFormatFolderName[k] + "/" + myDocumentName + ".pdf");
							var loResISBNFile = new File(myExportFolder + "/_LoRes/_" + isbnFormatFolderName[k] + "/" + myDocumentName + ".pdf");
							loResISBNFile.rename(isbnArray[k] + ".pdf");
						}
					}
				} else {
					reportLog += "\nNo Lo-res preset was selected: skipping Lo-res export";
				}
			}
			app.pdfExportPreferences.pageRange = userPagePref;
		}

		//Export Images
		function exportImages () {
			//image paths
			var myFileImageToc = new File(myExportFolder + "/_Marketing_Images/" + myDocumentName + "_toc.jpg");
			var myFileImageIntArray = new Array ();
			for (var z = 0; z < userSpreadPages.length; z++ ) {
				myFileImageIntArray[z] = new File(myExportFolder + "/_Marketing_Images/" + myDocumentName + "_int0" + (z+1) + ".jpg");
			}

			var interiorImageArray = new Array ();
			for (var b = 0; b < userSpreadPages.length; b++ ) {
				interiorImageArray[b] = parseInt(userSpreadPages[b]) + "-" + (1+parseInt(userSpreadPages[b]));
			}
			
			//Export Default Settings
			/*
			var myCurrentHeight = myDoc.documentPreferences.pageHeight;
			var myResizePercentage = 5.33/myCurrentHeight;
			var myExportRes = 300;*/

			with (app.jpegExportPreferences) {
					exportingSpread = false;
					jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
					pageString = tocPage;
					exportResolution = 300; // The export resolution expressed as a real number instead of an integer. (Range: 1.0 to 2400.0)
					antiAlias = true; //  If true, use anti-aliasing for text and vectors during export
					embedColorProfile = false; // True to embed the color profile, false otherwise
					jpegColorSpace = JpegColorSpaceEnum.RGB; // One of RGB, CMYK or GRAY
					jpegQuality = JPEGOptionsQuality.HIGH; // The compression quality: LOW / MEDIUM / HIGH / MAXIMUM
					jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; // The rendering style: BASELINE_ENCODING or PROGRESSIVE_ENCODING
					simulateOverprint = false; // If true, simulates the effects of overprinting spot and process colors in the same way they would occur when printing
					useDocumentBleeds = false; // If true, uses the document's bleed settings in the exported JPEG.
				}

			//Export TOC Image
			if (tocPage!=""){
				myDoc.exportFile(ExportFormat.JPG, myFileImageToc, false);
			}

			//Export Interior Images
			for (var a = 0; a < userSpreadPages.length; a++){
				if (userSpreadPages[a]!=""){
					with (app.jpegExportPreferences) {
							exportingSpread = true;
							pageString = interiorImageArray[a];
						}
					myDoc.exportFile(ExportFormat.JPG, myFileImageIntArray[a], false);
				}
			}
			alert("BEFORE CONTINUING:\nResize marketing images in photoshop\nLocated:" + myExportFolder + "/_Marketing_Images\nClick OK when ready to be copied to ISBN folders.");

			//Move images
			for (var k = 0; k < 3; k++) {
				if (isbnArray[k]!=""){
					if (tocPage!=""){
						myFileImageToc.copy(myExportFolder + "/_Marketing_Images/_" + isbnFormatFolderName[k] + "/" + myDocumentName + "_toc.jpg");
						var ImageToc = new File(myExportFolder + "/_Marketing_Images/_" + isbnFormatFolderName[k] + "/" + myDocumentName + "_toc.jpg");
						ImageToc.rename(isbnArray[k] + "_toc.jpg");
					}
					for (var a = 0; a < userSpreadPages.length; a++){
						if (userSpreadPages[a]!=""){
							myFileImageIntArray[a].copy(myExportFolder + "/_Marketing_Images/_" + isbnFormatFolderName[k] + "/" + myDocumentName + "_int0" + (a+1) + ".jpg");
							var ImageInt = new File(myExportFolder + "/_Marketing_Images/_" + isbnFormatFolderName[k] + "/" + myDocumentName + "_int0" + (a+1) + ".jpg");
							ImageInt.rename(isbnArray[k] + "_int0" + (a+1) + ".jpg");
						}
					}
				}
			}
		}
	}
}



