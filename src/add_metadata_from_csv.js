main();
function main()
{
  var myDoc = app.activeDocument;
  var inputKey = "Isbn Code";
  var fileInfoFields = ["Title","Author","Descr","Copyright","Copyright"];

  var content = getContent();
  var userInput = getInput(inputKey);
  var xContent = convertComas(content);
  var xKeys = createKeys(xContent);
  var csvData = createCSVData(xContent);
  var rowNum = findRow(xKeys, csvData, inputKey, userInput)

  var fileInfo = 
  {
    title:getCellValue(xKeys, csvData, fileInfoFields[0],rowNum), 
    author:fixAuthor(getCellValue(xKeys, csvData, fileInfoFields[1],rowNum)), 
    description:getCellValue(xKeys, csvData, fileInfoFields[2],rowNum), 
    keywords:["Test word 1", "Test word 2"],
    copyrightInfoURL:"",
    copyrightNotice:getCellValue(xKeys, csvData, fileInfoFields[3],rowNum),
    copyrightStatus:"YES",
    creationDate:"",
    
  }
  setFileInfo(myDoc, fileInfo);
  
}

function getContent()
{
  var file = File.openDialog("Choose CSV file","*.csv"); // get the file
  if (file)
  {
    file.encoding = 'UTF8'; // set some encoding
    file.lineFeed = 'Macintosh'; // set the linefeeds
    file.open('r',undefined,undefined); // read the file
    var filecontents = file.read(); // get the text in it
    file.close(); // close it again
    return filecontents;
  } else
  {
    exit();
  }
}

function convertComas(contentInput)
{
  var fixedContent = "";
  var skipComma = false;
  var charAdd = "";

  for (var i = 0; i < contentInput.length; i++)
  {
    if (contentInput.charAt(i) == "\"")
    {
      i++;
      if (contentInput.charAt(i) != "\"")
      {
        if(skipComma == true)
        {
          skipComma = false;
        } else
        {
          skipComma = true;
        }
      }
    }
    if (contentInput.charAt(i) == "," && skipComma == false)
    {
      charAdd = "***";
    } else
    {
      charAdd = contentInput.charAt(i);
    }
    fixedContent += charAdd;
  }

  return fixedContent;
}

function createKeys(content)
{
  var rows = content.split('\n');// split the lines (windows should be '\r')
  var keys = rows[0].split('***'); // get the heads

  return keys;
}

function createCSVData(content)
{
  var rows = content.split('\n');// split the lines (windows should be '\r')
  var data = [];// will hold the data
  // loop the data
  for(var i = 1; i < rows.length;i++){
    var cells = rows[i].split('***');// get the cells
    data.push(cells);
    }
  return data;
}

function findRow(keys, csv, keyValue, input)
{
  var column;
  //find column of Input
  for (var i = 0; i <  keys.length; i++)
  {
    if (keys[i]==keyValue)
    {
      column = i;
      break;
    }
  }

  //find row that matches Input
  for (var row = 0; row < csv.length; row++)
  {
    if (csv[row][column]==input)
    {
      return row;
    }
  }
  alert("Input not found\nLooking in column " + keyValue);
  exit();
}

function getInput(keyValue)
{
  var message = "Enter " + keyValue;
  var myInputWindow = new Window("dialog", message);
  myInputWindow.orientation = 'row';
  myInputWindow.alignment = "Left";
  myInputWindow.alignChildren = "Left";
  var myInput = myInputWindow.add("edittext",undefined,"");
  myInput.characters = 20;
  myInputWindow.add("button",undefined,"OK");
  myInputWindow.add("button",undefined,"Cancel");
  if (myInputWindow.show()==1)
  {
    var xInput = myInput.text;
    return xInput;
  } else
  {
    exit();
  }    
}

function getCellValue(keys, csv, keyValue, row)
{
  var column = "";
  //find column of Input
  for (var i = 0; i <  keys.length; i++)
  {
    if (keys[i]==keyValue)
    {
      column = i;
      break;
    }
  }
  if (column!="")
  {
    return csv[row][column];
  } else
  {
    alert("Could not find a column for " + keyValue);
    exit();
  }

}

function setFileInfo(xMyDoc, xFileInfo)
{
  var myInputWindow = new Window("dialog", "Results");
  myInputWindow.orientation = 'column';
  myInputWindow.alignment = "Left";
  myInputWindow.alignChildren = "Left";
  myInputWindow.add("statictext",undefined,"Title: " + xFileInfo.title);
  myInputWindow.add("statictext",undefined,"Author: " + xFileInfo.author);
  var myDescLabel = myInputWindow.add("statictext",undefined,"Description: " + xFileInfo.description,{multiline:true});
  myDescLabel.preferredSize.width = 700;
  myInputWindow.add("statictext",undefined,"Keywords: " + xFileInfo.keywords);
  myInputWindow.add("statictext",undefined,"Copyright Notice: " + xFileInfo.copyrightNotice);
  var myButtonGroup = myInputWindow.add("group");
  myButtonGroup.alignment = "right";
  myButtonGroup.add("button",undefined,"OK");
  myButtonGroup.add("button",undefined,"Cancel");
  myInputWindow.layout.layout (true);

  if (myInputWindow.show()==1)
  {
    xMyDoc.metadataPreferences.documentTitle = xFileInfo.title;
    xMyDoc.metadataPreferences.author = xFileInfo.author;
    xMyDoc.metadataPreferences.description = xFileInfo.description;
    xMyDoc.metadataPreferences.keywords = xFileInfo.keywords;
    //xMyDoc.metadataPreferences.copyrightInfoURL = ;
    xMyDoc.metadataPreferences.copyrightNotice = xFileInfo.copyrightNotice;
    if (xFileInfo.copyrightStatus=="YES")
    {
        xMyDoc.metadataPreferences.copyrightStatus = CopyrightStatus.YES;
    } else if (xFileInfo.copyrightStatus=="NO")
    {
        xMyDoc.metadataPreferences.copyrightStatus = CopyrightStatus.NO;
    } else
    {
        xMyDoc.metadataPreferences.copyrightStatus = CopyrightStatus.UNKNOWN;
    }
    //xMyDoc.metadataPreferences.creationDate: ;

  } else
  {
    exit();
  }    
}

function fixAuthor(text)
{
  var fullName = "";
  var names = text.split(', ');
  for (x = names.length; x > 0; x--)
  {
    fullName += names[x-1];
    if (x > 1)
    {
      fullName += " ";
    }
  }
  return fullName;
}