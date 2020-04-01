
app.generalPreferences.ungroupRemembersLayers = false;
var myDoc = app.activeDocument;
var myItemList = myDoc.groups.everyItem().ungroup();
alert("Complete");