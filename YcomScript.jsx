#include "json2.js";

var thisFile = new File($.fileName);
var basePath = thisFile.path;
var fileName = "/jsonFile.json"

var myFile = File(basePath+fileName);
myFile.open('r');

var content = myFile.read();
var wow = JSON.parse(content);
var pathList = wow.psdPath; 
var nameList = wow.psdName;


var num = wow.psdName.length;
var count = 0;


var i = 0;
while(i<num){
    
    var psdPath = pathList[i];
    var psdName = nameList[i];
    i++;


    var myFile = File(psdPath);
    open(myFile);


    var oldUnits = preferences.rulerUnits;
    preferences.rulerUnits = Units.PIXELS;
    var doc = app.activeDocument;

    
    var numText = wow[psdName].length;
    var count = 0;
 
    var grouplayer = doc.layerSets.add();
    grouplayer.name = ('text');

    while (count < numText) {
        
        var text = wow[psdName][count]['text'];
        
        var x = wow[psdName][count]['x'];
        var y = wow[psdName][count]['y'];
        count++;   
        
        var myLayer = doc.artLayers.add();
        myLayer.kind = LayerKind.TEXT;  
        myLayer.textItem.contents = text;
        myLayer.textItem.position = [x, y];

        myLayer.textItem.size = 10;
        

        myLayer.move(grouplayer, ElementPlacement.INSIDE);
    }
    
    doc.close(SaveOptions.SAVECHANGES);
    
}

myFile.close();
