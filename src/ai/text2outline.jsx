//　++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//　text2outline.jsx
//  all text objects makes outline objects.
//
//　2017/03/28  ver.0.0.1
//　++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//==============================================================
//  main processing
//==============================================================
var srcFolder = Folder.selectDialog("select source folder.");
var saveFolder = Folder.selectDialog("select saving folder.");

var targetExtension = "*.eps";
var saveExtension = ".eps";
var fileSeparater = "/";

if (srcFolder != null) {
    fileList = new Array;
    fileList = srcFolder.getFiles(targetExtension);

    for (f = 0; f <= fileList.length - 1; f++) {

        var srcFile = new File(fileList[f]);
        open(srcFile);
        var openFlg = srcFile.open();
        if (openFlg == true) {
            var srcDocument = app.activeDocument;
            var visibleFlag = layerUnlocked(srcDocument); // layer unlocked
            obj2outline(srcDocument.textFrames); //textframes
            obj2outline(srcDocument.stories); //stories
            layerLocked(srcDocument, visibleFlag) // layer locked
        }
        var epsSaveOptions = getEPSOptions();
        saveFile = new File(getSaveFileName (saveFolder.fsName, srcFile.name, saveExtension));
        srcDocument.saveAs(saveFile, epsSaveOptions);
        activeDocument.close(SaveOptions.DONOTSAVECHANGES);
    }
}

alert("processing complete.");

//==============================================================
//  get save file name.
//==============================================================
function getSaveFileName(saveFolderName, srcFileName, extension) {
    namecount = (srcFileName).lastIndexOf("."); //　fileObj.nameの値で"."の位置を取得
    exceptExtension = (srcFileName).substr(0, namecount);
    return (saveFolderName + fileSeparater + exceptExtension + extension);
}

//==============================================================
//  text objects makes outline ocjects.
//==============================================================
function obj2outline(objects) {

    for (n = objects.length - 1; n > -1; n--) {
        var hiddenflag = objects[n].hidden;
        objects[n].locked = false;
        if (hiddenflag == false) {
            objects[n].createOutline();
        } else {
            var newobject = objects[n].createOutline();
            newobject.hidden = true;
        }
    }
}

//==============================================================
//  all layer unlocked in document.
//==============================================================
function layerUnlocked(document){
    var visibleFlag = []

    for (i = 0; i < document.layers.length; i++) {
        var curLayer = document.layers[i];
        visibleFlag[i] = curLayer.visible;
        curLayer.visible = true;
        curLayer.locked = false;
    }
    return visibleFlag;
}

//==============================================================
// all layer locked in document.
//==============================================================
function layerLocked(document, visibleFlg){
    // layer locks and view status
    for (m = 0; m < document.layers.length; m++) {
        var curLayer = document.layers[m];
        if (visibleFlg[m] == true) {
            curLayer.locked = true;
        } else {
            curLayer.visible = false;
            curLayer.locked = true;
        }
    }
}

//==============================================================
// get saving eps options
//==============================================================
function getEPSOptions() {
    var epsSaveOptions = new EPSSaveOptions();
    epsSaveOptions.compatibility = Compatibility.ILLUSTRATOR12; //saving cs2 format
    //    epsSaveOptions.preserveEditingCapabilities = true;
    return epsSaveOptions;
}
