var myDoc = app.activeDocument;
main()

function main(){
    var textFrames = getTextItems();
    var textFramesWithFill = getTextFramesWithFill(textFrames);
    for (var x = 0; x < textFramesWithFill.length; x++){
        separateIntoTextAndRectangle(textFramesWithFill[x]);
    }
}

//get all text frames
function getTextItems(){
    var allTextFrames = [];
    for (var i = myDoc.allPageItems.length - 1; i >= 0; i--) {
        var pi = myDoc.allPageItems[i];
        //if(pi instanceof TextFrame || pi instanceof TextPath) {
        if(pi instanceof TextFrame) {
            allTextFrames.push(pi);
        };
    }
    return allTextFrames;
}
//Return array of textframes with a fill color
function getTextFramesWithFill(textFrames){
    var textFramesWithFill = [];
    for(var x = textFrames.length-1; x >=0; x--){
        if(textFrames[x].fillColor.name!="None"){
            textFramesWithFill.push(textFrames[x]);
        }
    }
    return textFramesWithFill;
}
//Creates a rectangle with the same fill color as text frame behind the text from. Removes the fill color from text frame
function separateIntoTextAndRectangle(textFrame){
    var layer = textFrame.itemLayer;
    var page = textFrame.parentPage;
    var fillColor = textFrame.fillColor;
    //[y1, x1, y2, x2]
    var frameBounds = textFrame.geometricBounds;
    page.rectangles.add(layer,LocationOptions.AFTER,textFrame,{geometricBounds: [frameBounds[0],frameBounds[1],frameBounds[2], frameBounds[3]], fillColor: fillColor});
    textFrame.fillColor = "None";
}