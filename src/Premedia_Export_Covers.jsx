main();

function main() {

	if (app.documents.length > 0) {
		var myDoc = app.activeDocument;
        var myDocumentName = myDoc.name.slice (0, -5);
        var exportHiPDFCheck = false;
        var exportLoPDFCheck = false;
        var exportImagesCheck = false;
        var reportLog = "";
        var userPagePref = app.pdfExportPreferences.pageRange;

        //Set preferences
        userHoriz = myDoc.viewPreferences.horizontalMeasurementUnits;
        userVert = myDoc.viewPreferences.verticalMeasurementUnits;
        userOrigin = myDoc.viewPreferences.rulerOrigin;
        myDoc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.pixels;
        myDoc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.pixels;
        myDoc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;

        var myCurrentWidth = myDoc.documentPreferences.pageWidth;
        var myResizePercentage = 242/myCurrentWidth;
        var myExportRes = 72 * myResizePercentage;

        //Set preferences back to normal
        myDoc.viewPreferences.horizontalMeasurementUnits = userHoriz;
        myDoc.viewPreferences.verticalMeasurementUnits = userVert;
        myDoc.viewPreferences.rulerOrigin = userOrigin;


        var userInput = inputWindow();

        if (exportHiPDFCheck==true || exportLoPDFCheck==true || exportImagesCheck==true){
            //If either check box is true prompt for export folder
            var myExportFolder = chooseExportFolder();
			createExportFolders(myExportFolder);
			if (exportHiPDFCheck==true) {
                exportHiPDFwithPresets();
            	}
			if (exportLoPDFCheck==true) {
                exportLoPDFwithPresets();
            }
            if (exportImagesCheck==true) {
                exportImages();
            }
            alert("Script complete\nReport:" + reportLog);
        }

        //Input Window
		function inputWindow () {
			var myPresets = app.pdfExportPresets.everyItem().name;
            myPresets.unshift("- Select Preset -");
            var pageNumIsBlank = true;
            var titleIsBlank = true;

            //Export PDF Input
			var myInputWindow = new Window("dialog", "Export Settings");
                myInputWindow.orientation = 'column';
                myInputWindow.alignment = "Left";
                myInputWindow.alignChildren = "Left";

            //Export Hires PDF check
            var exportHiPDFCheckbox = myInputWindow.add("checkbox",undefined,"Export Hi-res PDF");
                exportHiPDFCheckbox.value = true;
                exportHiPDFCheckbox.onClick = function (){
                if (this.value==true){
                    myHiresPDFExportGroup.show();
                } else {
                    myHiresPDFExportGroup.hide();
                }
            }
            
            //Export Hires PDF Preset
            var myHiresPDFExportGroup = myInputWindow.add("group");
                myHiresPDFExportGroup.alignment = "Left";
                myHiresPDFExportGroup.add("statictext",undefined,"Hi-Res PDF Preset:");
            var myPDFPresetDropdownHiRes = myHiresPDFExportGroup.add('dropdownlist',undefined,undefined,{items:myPresets});
                myPDFPresetDropdownHiRes.selection=0;

            //Export Lores PDF check
            var exportLoPDFCheckbox = myInputWindow.add("checkbox",undefined,"Export Lo-res PDF");
                exportLoPDFCheckbox.value = true;
                exportLoPDFCheckbox.onClick = function (){
                if (this.value==true){
                    myLoresPDFGroup.show();
                } else {
                    myLoresPDFGroup.hide();
                    myButtonGroup.children[0].enabled = true;
                }
            } 

            //Export Lores PDF Preset
            var myLoresPDFGroup = myInputWindow.add("group");
                myLoresPDFGroup.add("statictext",undefined,"Lo-Res PDF Preset:");
            var myPDFPresetDropdownLoRes = myLoresPDFGroup.add('dropdownlist',undefined,undefined,{items:myPresets});
                myPDFPresetDropdownLoRes.selection=0;

            //Export Images check
            var exportImagesCheckbox = myInputWindow.add("checkbox",undefined,"Export Marketing Images");
                exportImagesCheckbox.value = true;
                exportImagesCheckbox.onClick = function (){
                    if (this.value==true){
                        myLoresTitleGroup.show();
                        check_inputs();
                    } else {
                        myLoresTitleGroup.hide();
                        myButtonGroup.children[0].enabled = true;
                    }
                }

            //Input Titles panel
            var myLoresTitleGroup = myInputWindow.add("panel");
                myLoresTitleGroup.orientation = "column";
                myLoresTitleGroup.alignment = "Left";

            createTitleGroup(myLoresTitleGroup);

            //Buttons
			var myButtonGroup = myInputWindow.add("group");
                myButtonGroup.alignment = "right";
                myButtonGroup.add("button",undefined,"OK");
                myButtonGroup.children[0].enabled = false;
                myButtonGroup.add("button",undefined,"Cancel");
            
            //If settings exist set input fields to last used values
            var settingsExist = app.activeDocument.extractLabel('preMediaExportCover_SettingsTrue');
            if (settingsExist=="true"){
                //Set hires Checkbox
                exportHiPDFCheckbox.value = app.activeDocument.extractLabel('preMediaExportCover_exportHiPDFCheck');
                if (!exportHiPDFCheckbox.value){
                    myHiresPDFExportGroup.hide();
                }
                //Set lores Checkbox
                exportLoPDFCheckbox.value = app.activeDocument.extractLabel('preMediaExportCover_exportLoPDFCheck');
                if (!exportLoPDFCheckbox.value){
                    myLoresPDFGroup.hide();
                    myButtonGroup.children[0].enabled = true;
                }
                //Set Images Checkbox
                exportImagesCheckbox.value = app.activeDocument.extractLabel('preMediaExportCover_exportImagesCheck');
                if (!exportImagesCheckbox.value){
                    myLoresTitleGroup.hide();
                    myButtonGroup.children[0].enabled = true;
                }
                //Set titles
                var storedTitleArray = app.activeDocument.extractLabel('preMediaExportCover_titleInputArray');
                storedTitleArray = storedTitleArray.split("*");
                storedTitleArray.length = storedTitleArray.length-1;
                myPDFPresetDropdownHiRes.selection.text = storedTitleArray[0];
                myPDFPresetDropdownLoRes.selection.text = storedTitleArray[1];
                for (var x = 7; x < storedTitleArray.length; x+=5){               
                    createTitleGroup(myLoresTitleGroup);
                }
                for (var n = 0; n < myLoresTitleGroup.children.length; n++){
                    myLoresTitleGroup.children[n].pgnuminput.text = storedTitleArray[n*5+2];
                    myLoresTitleGroup.children[n].titleinput.text = storedTitleArray[n*5+3];
                    myLoresTitleGroup.children[n].hcinput.text = storedTitleArray[n*5+4];
                    myLoresTitleGroup.children[n].pbinput.text = storedTitleArray[n*5+5];
                    myLoresTitleGroup.children[n].ebinput.text = storedTitleArray[n*5+6];
                }
            }
            check_inputs();
            myInputWindow.layout.layout (true);
            
            //Input Window close
            if (myInputWindow.show()==1){

                //set variables after OK
                app.activeDocument.insertLabel('preMediaExportCover_SettingsTrue', 'true');
                
                exportLoPDFCheck = exportLoPDFCheckbox.value;
                var loresCheckString = ""
                if (exportLoPDFCheckbox.value){loresCheckString="true";} else {exportLoPDFCheckbox.value="false";};
                app.activeDocument.insertLabel('preMediaExportCover_exportLoPDFCheck', loresCheckString);

                exportHiPDFCheck = exportHiPDFCheckbox.value;
                var hiresCheckString = ""
                if (exportHiPDFCheckbox.value){hiresCheckString="true";} else {hiresCheckString.value="false";};
                app.activeDocument.insertLabel('preMediaExportCover_exportHiPDFCheck', hiresCheckString);

                exportImagesCheck = exportImagesCheckbox.value;
                var ImagesCheckString = ""
                if (exportImagesCheckbox.value){ImagesCheckString="true";} else {ImagesCheckString.value="false";};
                app.activeDocument.insertLabel('preMediaExportCover_exportImagesCheck', ImagesCheckString);

                var titleInputArray = new Array ();
                titleInputArray[0] = myPDFPresetDropdownHiRes.selection.text;
                titleInputArray[1] = myPDFPresetDropdownLoRes.selection.text;
                for (var n = 0; n < myLoresTitleGroup.children.length; n++){
                    titleInputArray[n*5+2] = myLoresTitleGroup.children[n].pgnuminput.text;
                    titleInputArray[n*5+3] = myLoresTitleGroup.children[n].titleinput.text.replace(/ /g,"_").replace(/\\/g,"_").replace(/\//g,"_");
                    titleInputArray[n*5+4] = myLoresTitleGroup.children[n].hcinput.text.replace(/-/g,"");
                    titleInputArray[n*5+5] = myLoresTitleGroup.children[n].pbinput.text.replace(/-/g,"");
                    titleInputArray[n*5+6] = myLoresTitleGroup.children[n].ebinput.text.replace(/-/g,"");
                }
                var titleArrayString = "";
                for (var x = 0; x < titleInputArray.length; x++){
                    titleArrayString += titleInputArray[x] + "*";
                }
                app.activeDocument.insertLabel('preMediaExportCover_titleInputArray', titleArrayString);

                return titleInputArray;
            } else {
                exit();
            }  
            
            //Create all the fields for a single title as a group
            function createTitleGroup (myLoresTitleGroup) {
                var titleGroup = myLoresTitleGroup.add("group");
                    titleGroup.pgnumlabel = titleGroup.add("statictext",undefined,"Page #:");
                    titleGroup.pgnuminput = titleGroup.add("edittext",undefined,"");
                    titleGroup.pgnuminput.characters = 3;
                    titleGroup.pgnuminput.onChanging = function () {
                        check_inputs();
                    }
                    pageNumIsBlank = true;
                    titleGroup.titlelabel = titleGroup.add("statictext",undefined,"Title:");
                    titleGroup.titleinput = titleGroup.add("edittext",undefined,"");
                    titleGroup.titleinput.characters = 20;
                    titleGroup.titleinput.onChanging = function () {
                        check_inputs();
                    }
                    titleIsBlank = true;
                    titleGroup.hclabel = titleGroup.add("statictext",undefined,"Hardcover ISBN:");
                    titleGroup.hcinput = titleGroup.add("edittext",undefined,"");
                    titleGroup.hcinput.characters = 17;
                    titleGroup.pblabel = titleGroup.add("statictext",undefined,"Paperback ISBN:");
                    titleGroup.pbinput = titleGroup.add("edittext",undefined,"");
                    titleGroup.pbinput.characters = 17;
                    titleGroup.eblabel = titleGroup.add("statictext",undefined,"EBook ISBN:");
                    titleGroup.ebinput = titleGroup.add("edittext",undefined,"");
                    titleGroup.ebinput.characters = 17;
                    titleGroup.plus = titleGroup.add("button", undefined, "+");
                    titleGroup.plus.onClick = add_btn;
                    titleGroup.minus = titleGroup.add("button", undefined, "-");
                    titleGroup.minus.onClick = minus_btn;
                    titleGroup.index = myLoresTitleGroup.children.length - 1;
                    if (typeof myButtonGroup != 'undefined'){
                        myButtonGroup.children[0].enabled = false;
                    }
                    myInputWindow.layout.layout (true);
            }

            function add_btn () {
                createTitleGroup (myLoresTitleGroup);
                check_inputs();
            }

            function minus_btn () {
                if (1 < myLoresTitleGroup.children.length){
                    myLoresTitleGroup.remove (this.parent);
                    myInputWindow.layout.layout (true);
                    }
                check_inputs();
            }

            function check_inputs () {
                pageNumIsBlank = false;
                titleIsBlank = false;
                for (var n = 0; n < myLoresTitleGroup.children.length; n++){
                    if(myLoresTitleGroup.children[n].pgnuminput.text == ""){
                        pageNumIsBlank = true;
                    }

                }
                for (var n = 0; n < myLoresTitleGroup.children.length; n++){
                    if(myLoresTitleGroup.children[n].titleinput.text == ""){
                        titleIsBlank = true;
                    }

                }
                if (!pageNumIsBlank && !titleIsBlank){
                    myButtonGroup.children[0].enabled = true;
                } else {
                    myButtonGroup.children[0].enabled = false;
                }
            }
        }

        //Choose Export Folder
        function chooseExportFolder () {
			var selectedExportFolder = Folder.selectDialog ("Choose a folder to export to");
			if (selectedExportFolder){
				return selectedExportFolder;
			} else {
				exit();
			}
        }
        
        //Create Directories
		function createExportFolders (myExportFolder) {
            var myLoresFolder = new Folder(myExportFolder + "/_LoRes");
            if (!myLoresFolder.exists) {
                myLoresFolder.create();
            }
                for (var n = 3; n < userInput.length ; n+=5) {
                    if (userInput[n+1]!=""){
                        var myLoresHardcoverFolder = new Folder(myLoresFolder + "/_Hardcover");
                        if (!myLoresHardcoverFolder.exists) {
                            myLoresHardcoverFolder.create();
                        }
                    }
                    if (userInput[n+2]!=""){
                        var myLoresPaperbackFolder = new Folder(myLoresFolder + "/_Paperback");
                        if (!myLoresPaperbackFolder.exists) {
                            myLoresPaperbackFolder.create();
                        }
                    }
                    if (userInput[n+3]!=""){
                        var myLoresEBookFolder = new Folder(myLoresFolder + "/_EBook");
                        if (!myLoresEBookFolder.exists) {
                            myLoresEBookFolder.create();
                        }
                    }
                }
            var myImagesFolder = new Folder(myExportFolder + "/_Marketing_Images");
            if (!myImagesFolder.exists) {
                myImagesFolder.create();
            }
                for (var n = 3; n < userInput.length ; n+=5) {
                    if (userInput[n+1]!=""){
                        var myImgHardcoverFolder = new Folder(myImagesFolder + "/_Hardcover");
                        if (!myImgHardcoverFolder.exists) {
                            myImgHardcoverFolder.create();
                        }
                    }
                    if (userInput[n+2]!=""){
                        var myImgPaperbackFolder = new Folder(myImagesFolder + "/_Paperback");
                        if (!myImgPaperbackFolder.exists) {
                            myImgPaperbackFolder.create();
                        }
                    }
                    if (userInput[n+3]!=""){
                        var myImgEBookFolder = new Folder(myImagesFolder + "/_EBook");
                        if (!myImgEBookFolder.exists) {
                            myImgEBookFolder.create();
                        }
                    }
                    if (userInput[n+1]!=""){
                        var myCIImagesFolder = new Folder(myImagesFolder + "/_CI_Cover");
                        if (!myCIImagesFolder.exists) {
                            myCIImagesFolder.create();
                        }
                    }
                }
        }

        //Export Hires PDF
		function exportHiPDFwithPresets () {
            if (userInput[0]!="" && userInput[0]!="- Select Preset -"){
                app.pdfExportPreferences.pageRange = "ALL_PAGES";
                var myFileHR = new File(myExportFolder + "/" + myDocumentName + ".pdf");
                myDoc.exportFile(ExportFormat.pdfType, myFileHR, false, userInput[0]);
                app.pdfExportPreferences.pageRange = userPagePref;
            } else{
                reportLog += "\nNo hi-res preset was selected: skipping hi-res export";
            }
        }
        
        //Export Lores PDF
        function exportLoPDFwithPresets () {
            var formats = ["Hardcover", "Paperback", "EBook"];
            if (userInput[1]!="" && userInput[1]!="- Select Preset -"){
                //Export Lo-res PDFs
                for (var n = 3; n < userInput.length ; n+=5) {
                    var myLRFolder = myExportFolder + "/_LoRes";
                    var myFileLR = new File(myLRFolder + "/" + userInput[n] + "_cover.pdf");
                    app.pdfExportPreferences.pageRange = userInput[n-1];
                    myDoc.exportFile(ExportFormat.pdfType, myFileLR, false, userInput[1]);
                    app.pdfExportPreferences.pageRange = userPagePref;
                }
            } else{
                reportLog += "\nNo Lo-res preset was selected: skipping Lo-res export\nskipping image export";
            }
        }

        //Export Images
        function exportImages () {
            //For each title...
            var formats = ["Hardcover", "Paperback", "EBook"];
            for (var n = 3; n < userInput.length ; n+=5) {
                var myImagesFolder = myExportFolder + "/_Marketing_Images";
                var myFileImage = new File(myImagesFolder + "/" + userInput[n] + ".jpg");
                with (app.jpegExportPreferences) {
                    exportingSpread = false;
                    jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
                    pageString = userInput[n-1];
                    exportResolution = 300; // The export resolution expressed as a real number instead of an integer. (Range: 1.0 to 2400.0)
                    antiAlias = true; //  If true, use anti-aliasing for text and vectors during export
                    embedColorProfile = false; // True to embed the color profile, false otherwise
                    jpegColorSpace = JpegColorSpaceEnum.RGB; // One of RGB, CMYK or GRAY
                    jpegQuality = JPEGOptionsQuality.HIGH; // The compression quality: LOW / MEDIUM / HIGH / MAXIMUM
                    jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; // The rendering style: BASELINE_ENCODING or PROGRESSIVE_ENCODING
                    simulateOverprint = false; // If true, simulates the effects of overprinting spot and process colors in the same way they would occur when printing
                    useDocumentBleeds = false; // If true, uses the document's bleed settings in the exported JPEG.
                }
                myDoc.exportFile(ExportFormat.JPG, myFileImage, false);
                
                with (app.jpegExportPreferences) {
                    exportingSpread = false;
                    jpegExportRange = ExportRangeOrAllPages.EXPORT_RANGE;
                    pageString = userInput[n-1];
                    exportResolution = myExportRes; // The export resolution expressed as a real number instead of an integer. (Range: 1.0 to 2400.0)
                    antiAlias = true; //  If true, use anti-aliasing for text and vectors during export
                    embedColorProfile = false; // True to embed the color profile, false otherwise
                    jpegColorSpace = JpegColorSpaceEnum.RGB; // One of RGB, CMYK or GRAY
                    jpegQuality = JPEGOptionsQuality.HIGH; // The compression quality: LOW / MEDIUM / HIGH / MAXIMUM
                    jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING; // The rendering style: BASELINE_ENCODING or PROGRESSIVE_ENCODING
                    simulateOverprint = false; // If true, simulates the effects of overprinting spot and process colors in the same way they would occur when printing
                    useDocumentBleeds = false; // If true, uses the document's bleed settings in the exported JPEG.
                }
                if (userInput[n+1]!=""){
                    var myFileCIImage = new File(myExportFolder + "/_Marketing_Images/_CI_Cover/" + userInput[n+1] + ".jpg");
                    myDoc.exportFile(ExportFormat.JPG, myFileCIImage, false);
                }
            }

                alert("BEFORE CONTINUING:\nResize marketing images in photoshop\nLocated:" + myExportFolder + "/_Marketing_Images/" + "\nClick OK when ready to be copied to ISBN folders.");
            for (var n = 3; n < userInput.length ; n+=5) {
                myImagesFolder = myExportFolder + "/_Marketing_Images";
                var myFileImageISBN = new File(myImagesFolder + "/" + userInput[n] + ".jpg");
                for (var x = 0; x < 3; x++) {
                    if (userInput[n+x+1]!=""){
                        myFileImageISBN.copy(myExportFolder + "/_Marketing_Images/_" + formats[x] + "/" + userInput[n+x+1] + ".jpg");
                    }
                }
            }
        }
    }
}