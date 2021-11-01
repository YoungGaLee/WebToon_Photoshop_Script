/*
    YcomScript.jsx

    포토샵에서 사용하는 script 파일 
    excel2json.py을 실행시켜 생성된 jsonFile.json을 사용한다.
    json파일 내의 text, x, y 등의 정보(대사,위치값)을 불러와 레이어를 생성, 자동저장시킨다.

    마지막 수정일 : 2021-10-28
    이가영
*/

// ExtendScript는 JSON모듈을 사용할 수 없으나, json2.js 플러그인을 include하면 사용이 가능하다. 
#include "json2.js";

var SIZE = 10;
var FONT = "굴림";

// 실행되는 스크립트가 위치한 폴더 내의 jsonFile.json에 대한 절대경로를 구하여 파일을 불러와 작업한다.  
var thisFile = new File($.fileName);
var basePath = thisFile.path;
var fileName = "/jsonFile.json"

var myFile = File(basePath+fileName);
myFile.open('r');
var content = myFile.read();
var wow = JSON.parse(content);

var pathList = wow.psdPath;     // script가 위치한 폴더 내 psd파일들의 절대경로 list
var nameList = wow.psdName;     // json파일의 key(psd파일이름) list
var num = wow.psdName.length;   

var i = 0;
while(i<num){
    var psdPath = pathList[i];
    var psdName = nameList[i];
    i++;

    // 작업 psd파일 open
    var myFile = File(psdPath);
    open(myFile);

    var oldUnits = preferences.rulerUnits;
    preferences.rulerUnits = Units.PIXELS;
    var doc = app.activeDocument;
    var numText = wow[psdName].length;

    // 'text'라는 Group Layer를 만들어주고 그 안에 Text Layer를 생성한다.
    var grouplayer = doc.layerSets.add();
    grouplayer.name = ('text');

    var count = 0;
    while (count < numText) {
        var text = wow[psdName][count]['text'];
        var x = wow[psdName][count]['x'];
        var y = wow[psdName][count]['y'];
        count++;   
        
        //Text Layer 생성
        var myLayer = doc.artLayers.add();
        myLayer.kind = LayerKind.TEXT;  
        myLayer.textItem.size = SIZE;
        myLayer.textItem.font = FONT;
        myLayer.textItem.contents = text;
        myLayer.textItem.position = [x, y];

        
        myLayer.move(grouplayer, ElementPlacement.INSIDE);
    }
    
    //수정한 내용을 저장하여 닫아주기
    doc.close(SaveOptions.SAVECHANGES);
}
alert("모든 PSD파일의 레이어 생성이 완료, 저장되었습니다.")
myFile.close();

