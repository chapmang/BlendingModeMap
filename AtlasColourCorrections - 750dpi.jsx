/**
* @@@ BUILD INFO @@@
* AtlasColourCorrection.jsx
*
* This script convers PDFs exported from ArcMap to support different colour blending modes
* to maintain specification standards established under the Mercator production process.
*
* Created October 2019
* @author Geoff Chapman
* 
* Version 1.0
*/

'use strict';  
// Suppress Illustrator Warning Dialogs
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

/*********************************************************
    Set global variables
 **********************************************************/
var logger = [];

/*********************************************************
    Polyfill to get arround lack of support for indexOf 
    method on arrays.
 **********************************************************/
if (!Array.prototype.indexOf)
  Array.prototype.indexOf = (function(Object, max, min) {
    "use strict"
    return function indexOf(member, fromIndex) {
      if (this === null || this === undefined)
        throw TypeError("Array.prototype.indexOf called on null or undefined")

      var that = Object(this), Len = that.length >>> 0, i = min(fromIndex | 0, Len)
      if (i < 0) i = max(0, Len + i)
      else if (i >= Len) return -1

      if (member === void 0) {        // undefined
        for (; i !== Len; ++i) if (that[i] === void 0 && i in that) return i
      } else if (member !== member) { // NaN
        return -1 // Since NaN !== NaN, it will never be found. Fast-path it.
      } else                          // all else
        for (; i !== Len; ++i) if (that[i] === member) return i 

      return -1 // if the value was not found, then return -1
    }
  })(Object, Math.max, Math.min)

/*********************************************************
    ShowLayers: Function to show and unlock all layers and indervidual artwork
**********************************************************/
function showLayers(activeDoc) {

    var items = activeDoc.pageItems;
    for (var i = 0; i < items.length; i += 1) {

        items[i].layer.locked = false; // Unlock whole layer
        items[i].layer.visible = true; // Unhide whole layer
        items[i].locked = false; // Unlock individual item
        items[i].hidden = false; // Unhide individual item
    }
}


/*********************************************************
    getTiffOptions: Function to define the output options of Tiff Export
**********************************************************/
function getTiffOptions(){
        
        var tiffExportOptions = new ExportOptionsTIFF();
        tiffExportOptions.antiAliasing = AntiAliasingMethod.TYPEOPTIMIZED;
        tiffExportOptions.imageColorSpace = ImageColorSpace.CMYK;
        tiffExportOptions.lZWCompression = true;
        tiffExportOptions.resolution = 750;
        tiffExportOptions.saveMultipleArtboards = true;
        tiffExportOptions.artboardRange = 1;
        return tiffExportOptions;
    }


/*********************************************************
    readAdjustmentFile: Function to translate adjustments CSV file
**********************************************************/
function readAdjustmentFile(adjustmentFilePath, headerRowStatus) {
    
    var adjustmentFile, i;
    var adjustments = []; 
        if (headerRowStatus === undefined) {
                headerRowStatus = true;
            }
        
        adjustmentFile = File(adjustmentFilePath);
        parsedAdjustmentFile = adjustmentFile.open("r");
        
        if (parsedAdjustmentFile == true) {
            var myfile = adjustmentFile.read();
            var linesArray = myfile.split("\n");
            if (headerRowStatus = true) {
                    var headers = linesArray[0].split(",");
                    i = 1;
                } else {
                    i=0;
                    }
            for(var i ; i < linesArray.length; i++) {
                if(linesArray[i].length > 0){
                    var data = linesArray[i].split(',');
                    var obj = {};
                    
                   // skip trailing lines in the csv file
                   if (data.length < 1) {
                       continue;
                       }

                    for(var j = 0; j < data.length; j++) {
                        obj[headers[j]] = data[j];
                    }
                
                     //Refelection of object to help debugging
                    
               // var props = obj.reflect.properties;
                // for (var k = 0; k < props.length; k++) {
                   //   $.writeln('this property ' + props[k].name + ' is ' + obj[props[k].name]);
                   // }
                    adjustments.push(obj);
                    }
                }
            }
        return adjustments;
        }


/*********************************************************
    verifyBlendingMode: Function to verify the blending mode values
*********************************************************/
function verifyBlendingMode(blendingMode){

     var mode,
        result = {},
        modeError = null;
     switch(blendingMode) {
            case "Darken":
                mode = BlendModes.DARKEN;
                break;
            case "Multiply":
                mode = BlendModes.MULTIPLY;
                break;
            case "Colorblend":
                mode = BlendModes.COLORBLEND;
                break;
            case "ColorBurn":
                mode = BlendModes.COLORBURN;
                break;
            case "ColorDodge":
                mode = BlendModes.COLORDODGE;
                break;
            case "Difference":
                mode = BlendModes.DIFFERENCE;
                break;
            case "Exclusion":
                mode = BlendModes.EXCLUSION;
                break;
            case "Hardlight":
                mode = BlendModes.HARDLIGHT;
                break;
            case "Hue":
                mode = BlendModes.HUE;
                break;
            case "Lighten":
                mode = BlendModes.LIGHTEN;
                break;
            case "Luminosity":
                mode = BlendModes.LUMINOSITY;
                break;
            case "Overlay":
                mode = BlendModes.OVERLAY;
                break;
            case "Saturation":
                mode = BlendModes.SATURATIONBLEND;
                break;
            case "Screen":
                mode = BlendModes.SCREEN;
                break;
            case "Softlight":
                mode = BlendModes.SOFTLIGHT;
                break;
            case "Normal":
                mode = BlendModes.NORMAL;
                break;
            default:
                modeError = 'Blending Mode Unknown: '+blendingMode;
                mode = BlendModes.NORMAL;
        }
    result.mode = mode
    result.modeError = modeError;
    return result;
    }


/*********************************************************
    adjustColours: Function to adjust the colours as per adjustment file
**********************************************************/
function adjustFills(blendingMode, adjustedSpot){
        var mode,
            selection,
            fillResults = {
                fillCount:  0,
                fillModeError: null,
                strokeCount: 0,
                strokeModeError: null
                };
            //$.writeln("Fill to be done: "+blendingMode+"|"+adjustedSpot);
             // Create a temp object to allow swatch selection    
            var temp = app.documents[0].pathItems.rectangle(-1000, -1000, 10, 10);
            temp.strokeColor = adjustedSpot.color;
            temp.fillColor = adjustedSpot.color;
            temp.selected = true;
            mode = verifyBlendingMode (blendingMode);
            //$.writeln("Fill Mode OK: "+mode.mode);
            app.executeMenuCommand('Find Fill Color menu item');
            selection = app.activeDocument.selection;
            fillResults.fillCount = selection.length;
            fillResults.fillModeError = mode.modeError;
            for (x in selection) {
                selection[x].blendingMode = mode.mode;
                }
            temp.remove();
            app.selection = null; 
        
        return fillResults;
    }

/*********************************************************
    adjustColours: Function to adjust the colours as per adjustment file
**********************************************************/
function adjustStrokes(blendingMode, adjustedSpot){
        var mode,
            selection,
            strokeResults = {
                fillCount:  0,
                fillModeError: null,
                strokeCount: 0,
                strokeModeError: null
                };

            //$.writeln("Stroke to be done: "+blendingMode+"|"+adjustedSpot);  
            // Create a temp object to allow swatch selection    
            var temp = app.documents[0].pathItems.rectangle(-1000, -1000, 10, 10);
            temp.strokeColor = adjustedSpot.color;
            temp.fillColor = adjustedSpot.color;
            temp.selected = true;
            mode = verifyBlendingMode (blendingMode);
            //$.writeln("Stroke Mode OK: "+mode.mode);
            app.executeMenuCommand('Find Stroke Color menu item');
            selection = app.activeDocument.selection;
            strokeResults.strokeCount = selection.length;
            strokeResults.strokeModeError = mode.modeError;
            for (x in selection) {
                selection[x].blendingMode = mode.mode;
                }
            temp.remove();
            app.selection = null; 

        return strokeResults;
    }
function logFileWriter(jobLogFile,fileName) {
        
        jobLogFile.encoding = "utf-8";
        jobLogFile.open('w');
        jobLogFile.writeln('*******************************');
        jobLogFile.writeln(fileName+" Log");
        jobLogFile.writeln('*******************************');
        for (var i = 0; i<logger.length; i++) {
                jobLogFile.write('\n');
                jobLogFile.writeln("Product: "+logger[i].product);
                if (logger[i].layerCount) {
                        jobLogFile.writeln("ERROR: File contained multiple layers");
                    }
                if (logger[i].extraSwatch.length > 0) {
                        jobLogFile.writeln("ERROR: File contianed unregistered spot colour");
                    }

                    for (x in logger[i].swatches) {
                        if (logger[i].swatches[x].name == 'indexOf') {continue}
                        jobLogFile.writeln("Swatch: "+logger[i].swatches[x].name);
                        if (logger[i].swatches[x].fills) {
                                jobLogFile.writeln("\tFills Adjusted: "+logger[i].swatches[x].fills.count);
                                jobLogFile.writeln("\tBlending Mode used: "+logger[i].swatches[x].fills.mode);
                                 if (logger[i].swatches[x].fills.modeError) {
                                        jobLogFile.writeln("\tFill Blending Mode not recognised");
                                    }
                                }
                        if (logger[i].swatches[x].strokes){
                                jobLogFile.writeln("\tStrokes Adjusted: "+logger[i].swatches[x].strokes.count);
                                jobLogFile.writeln("\tBlending Mode used: "+logger[i].swatches[x].strokes.mode);
                                if (logger[i].swatches[x].strokes.modeError) {
                                        jobLogFile.writeln("\tStroke Blending Mode not recognised");
                                    }
                                }
                            }
                }
        jobLogFile.close();
        return;
    }

/*********************************************************
    processFiles: Function to loop through files 
**********************************************************/
function processFiles(sourceFolderPath, adjustmentFilePath, headerRowStatus){
        
        var fileList = sourceFolderPath.getFiles(),
            passFolder = new Folder(decodeURI (sourceFolderPath)+"/Passed"),
            failFolder = new Folder(decodeURI (sourceFolderPath)+"/Failed")
        
        // Build output folders
        if(!passFolder.exists){
                passFolder.create();
            }
        if(!failFolder.exists){
                failFolder.create();
            }
        
        // Get the look up list from the adjustments file
        var adjustmentList = readAdjustmentFile (adjustmentFilePath, headerRowStatus);
        
        // Loop through files in selected folder
        var listLength = fileList.length;
        for (var i = 0; i < listLength; i++) {
            var fileLog = {
                noSwatchMatch: [], // Adjustment list entry not in file
                extraSwatch: [], // File had unexpected spot swatch
                featureCount: [],
                blendingError: [],
                swatches: []
            };
                if (fileList[i] instanceof File && fileList[i].name.match(/\.(pdf)$/i)){
                                          
                        var activeDoc = open(fileList[i]);
                        var activeDocName = fileList[i].name.slice(0,-4); // Name of document being processed
                        fileLog.product = activeDocName;
                        
                        // Logging that the file has more layers than anticipated 
                        if (activeDoc.layers.length > 1) {
                            fileLog.layerCount = activeDoc.layers.length;
                        }
                    
                        // get list of spots in document as basic array as it makes the error logging simpler
                        // also only compareing names so don't need most of the other object properties
                        var spotColourList = [];
                        for (var j=0; j<activeDoc.spots.length; j++){
                                spotColourList.push(activeDoc.spots[j].name);
                            }
                        
                        var noMatchCount = 0;
                        
                        // Log if file has more
                        // remove one from spot array as '[Registration]' always present
                        if (spotColourList.length-1 > adjustmentList.length){
                                fileLog.extraSwatch.push(activeDocName);
                            }

                        // $.writeln(adjustmentList.length);
                        for (var j=0; j<adjustmentList.length; j++){
                            if (spotColourList.indexOf(adjustmentList[j].SpotColour) === -1) {
                                continue;
                            }
                            try {
                                adjustedSpot = activeDoc.swatches.getByName(adjustmentList[j].SpotColour);
                                fileLog.swatches[j] = {};
                                if (adjustmentList[j].Type == "Fill") {
                                    // $.writeln("Fill: "+adjustmentList[j].BlendingMode);
                                    //$.writeln("Colour: "+ adjustedSpot);
                                    fillResults = adjustFills(adjustmentList[j].BlendingMode, adjustedSpot);
                                    fileLog.swatches[j].name = adjustedSpot.name;
                                    fileLog.swatches[j].fills = {};
                                    fileLog.swatches[j].fills.count = fillResults.fillCount;
                                    fileLog.swatches[j].fills.mode = adjustmentList[j].BlendingMode;
                                    fileLog.swatches[j].fills.modeError = fillResults.fillModeError;
                                    } else if (adjustmentList[j].Type == "Stroke") {
                                    //$.writeln("Stroke: "+adjustmentList[j].BlendingMode);
                                    //$.writeln("Colour: "+ adjustedSpot);
                                    strokeResults = adjustStrokes(adjustmentList[j].BlendingMode, adjustedSpot);
                                    fileLog.swatches[j].name = adjustedSpot.name;
                                    fileLog.swatches[j].strokes = {};
                                    fileLog.swatches[j].strokes.count = strokeResults.strokeCount;
                                    fileLog.swatches[j].strokes.mode = adjustmentList[j].BlendingMode;
                                    fileLog.swatches[j].strokes.modeError = strokeResults.strokeModeError;
                                    }
                                }
                            catch(e) {
                                    // The look-up colour was not in the file (not an issue just being recorded)
                                    fileLog.noSwatchMatch.push(adjustmentList[j].SpotColour);
                                }
                            
                       }
                   }
                    
                 while (app.documents.length) {
                    var tiffExportOptions = getTiffOptions();
                    var exportType = ExportType.TIFF;
                    if(fileLog.extraSwatch.length > 0 || fileLog.layerCount) {
                        exportFile = new File(decodeURI (failFolder)+"/"+activeDocName+".tif");
                    } else {
                        exportFile = new File(decodeURI (passFolder)+"/"+activeDocName+".tif");
                    }
                    activeDoc.exportFile(exportFile, exportType, tiffExportOptions);
                    app.redraw();
                    activeDoc.close(SaveOptions.DONOTSAVECHANGES);
                    logger.push(fileLog);
                }
            }
      // Write out job log file
     var jobLogFile = new File(decodeURI (sourceFolderPath)+"/"+sourceFolderPath.name+"_Log.txt");
     logFileWriter(jobLogFile,sourceFolderPath.name);
      alert(fileList.length+" files processed");
      return 
    }


/****************************
 *  RUN
 *  Launch the script and generate dialog boxes for file selection
 * 
 ****************************/
var Run = function(){

    var selectedFolder, selectedFile, headerRowStatus;
    
    // Build Main Dialog Box
    var dlg = new Window('dialog','Atlas Colour Correction');
    dlg.alignChildren = 'fill';
    // Select Folder Panel
    dlg.sourceFolderSelectPanel = dlg.add('panel', undefined, "Select Folder for Processing:"); 
    dlg.sourceFolderSelectPanel.browseFolder = dlg.sourceFolderSelectPanel.add('edittext',[0,0,275,25],'',{multiline:false,noecho:false,readonly:false});
    dlg.sourceFolderSelectPanel.browseButton = dlg.sourceFolderSelectPanel.add('button',undefined,"Browse");
    dlg.sourceFolderSelectPanel.browseFolder.enabled = false;
    dlg.sourceFolderSelectPanel.orientation='row';
    // Select Adjustment File Panel
    dlg.adjustmentFileSelectPanel = dlg.add('panel', undefined, "Select Adjustment File:"); 
    dlg.adjustmentFileSelectPanel.browser = dlg.adjustmentFileSelectPanel.add('group',undefined);
    dlg.adjustmentFileSelectPanel.browser.browseFolder = dlg.adjustmentFileSelectPanel.browser.add('edittext',[0,0,275,25],'',{multiline:false,noecho:false,readonly:false});
    dlg.adjustmentFileSelectPanel.browser.browseButton = dlg.adjustmentFileSelectPanel.browser.add('button',undefined,"Browse");
    dlg.adjustmentFileSelectPanel.browser.browseFolder.enabled = false;
    dlg.adjustmentFileSelectPanel.browser.orientation='row';
    dlg.adjustmentFileSelectPanel.options = dlg.adjustmentFileSelectPanel.add('group', undefined,'');
    dlg.adjustmentFileSelectPanel.options.headerOption = dlg.adjustmentFileSelectPanel.options.add('checkbox',[-90,0,275,25], 'Adjustment File Has Header Row');
    dlg.adjustmentFileSelectPanel.options.headerOption.value = true;
    dlg.adjustmentFileSelectPanel.options.orientation='column';
    // Run Additional Action Panel
    dlg.runActionPanel = dlg.add('panel',undefined, "Optional Action:");
    dlg.runActionPanel.actionName = dlg.runActionPanel.add('edittext',[0,0,375,25],'');
    dlg.runActionPanel.orientation = 'row';
    // Action Button Panel
    dlg.actionButtonGroup = dlg.add('group',undefined,'Action Buttons');
    dlg.actionButtonGroup.goButton = dlg.actionButtonGroup.add('button',undefined,'Process',{name:'OK'});
    dlg.actionButtonGroup.cancelButton = dlg.actionButtonGroup.add('button',undefined,'Cancel',{name:'Cancel'});
    
    // Select Folder Action
    dlg.sourceFolderSelectPanel.browseButton.onClick = function(){
            selectedFolder = Folder.selectDialog ("Please select a folder to process");
            if(selectedFolder != null) {
                    dlg.sourceFolderSelectPanel.browseFolder.text = decodeURI (selectedFolder);
                    dlg.sourceFolderSelectPanel.browseFolder.enabled = true;
                }
        }
    // Select Adjustment File Action
    dlg.adjustmentFileSelectPanel.browser.browseButton.onClick = function(){
            selectedFile = File.openDialog ("Please select a conversion reference file", "Text Files:*.txt;*.csv", false);
             if(selectedFile != null) {
                    dlg.adjustmentFileSelectPanel.browser.browseFolder.text = decodeURI (selectedFile);
                    dlg.adjustmentFileSelectPanel.browser.browseFolder.enabled = true;
                }
        }
    // Action Buttons Actions
    dlg.actionButtonGroup.goButton.onClick = function(){
            if(dlg.sourceFolderSelectPanel.browseFolder.text == '') {
                    alert("No folder has been selected!");
                    return;
                }
            else if(dlg.adjustmentFileSelectPanel.browser.browseFolder.text == '') {
                    alert("No adjustment file has been selected");
                    return
                }
            else {
                    headerRowStatus = dlg.adjustmentFileSelectPanel.options.headerOption.value;
                    processFiles (selectedFolder, selectedFile, headerRowStatus);
                }
        }
    dlg.actionButtonGroup.cancelButton.onClick = function(){
            dlg.close(1);
        }
    dlg.show();
    dlg.close (1);
    }();