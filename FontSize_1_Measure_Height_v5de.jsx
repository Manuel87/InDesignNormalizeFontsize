//Description:Mit diesem Skript ist es möglich das vertikales Ausmaß eines beliebingen Zeichens gesetzter Schriften genau auszumessen
//Dabei misst das Skript immer die die äußersten Punkte, inklusive herumfliegender loser Formen, wie beispielsweise der i-Punkt, etc.

// Gebrauch, auf eigene Gefahr.
// Wir können nicht garantieren, dass das Script je nach Schrit nicht auch unerwartete Ergebnisse erzeugen kann

// Diese Skript basiert auf dem Skript von Gerald Singelmann:
// "SetVisualCharSize.jsx"
// http://indesignsecrets.com/set-the-size-of-text-exactly-based-on-cap-or-x-height.php // http://www.indesign-faq.de/de/versal-und-andere-hohen-angleichen
//© cuppascript, 10/2009

//Autor dieser Anpassung/Weiterentwicklung: Manuel von Gebhardi


// enable ExtendScript L10N automatism
$.localize = true;
// Localization-Objects
var sNoSelection = { en: "No text selected", de: "Es ist kein Text markiert." };
var sNoDocument = { en: "No document open", de: "Es ist kein Dokument offen." };
var sDialogName = { en: "Measure Height of any Specific Glyph", de: "Messe die vertikale Größe eines beliebigen Zeichens" };
var sDefCharacter = { en: "Specific Character", de: "Zu messendes Zeichen:" }; //Maßgebendes Zeichen:
var sOvershoot_2 = { en: "Exclude Baseline Overshoot", de: "Grundlinienüberhang mit einbeziehen"};
//var sOvershoot = { en: "[**] substract Baseline Overshoot", de: "[**] substract 2times the Baseline Overshoot"};

//var sCHeight = { en: "New Height in em/1000:", de: "New Height in em/1000:" }; //Höhe des Zeichens:
var myGetRelArray = [0, 0];
var mainFuncArray = [0,0];
var resize = false;
//var mySubstractOvershoot = false;
var mySubstractOvershoot_2 = false;
var pointsizemeasuringcorrection = 1;

if (app.documents.length != 0){
	var myDocument = app.activeDocument;
	if (app.selection.length >0){
		var allSel = app.selection;
		var s = app.selection[0].constructor.name;
		switch (app.selection[0].constructor.name){
			case "Text":
			case "Character":
			case "Word":
			case "Line":
			case "TextStyleRange":
			case "TextColumn":
			case "Paragraph":
				main(false, false);
				break;
			case "TextFrame":
				var aSel = app.selection;
				var myHeight = false;
				var myEmHeight = false;
				for (var n = 0; n < aSel.length; n++) {
					if (aSel[n].constructor.name == "TextFrame") {
						if (aSel[n].parentStory.textContainers.length == 1) {
							app.select(aSel[n].texts);
							//alert("fu"+myHeight);
							//alert("fu2"+myEmHeight);
							mainFuncArray = main(myHeight, myEmHeight);
							myEmHeight = mainFuncArray[1];
							myHeight = mainFuncArray[0];
						}
					}
				}
				break;
			default:
				alert(localize(sNoSelection));
				break;
		}
		app.select(allSel);
	}
	else{
		alert(localize(sNoSelection));
	}
}
else{
	alert(localize(sNoDocument));
}

// Main Function – works, when myHeigt==false
function main(myHeight, myEmHeight) {
	var prevHUnits = app.activeDocument.viewPreferences.horizontalMeasurementUnits;
	app.activeDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
	var prevVUnits = app.activeDocument.viewPreferences.verticalMeasurementUnits;
	app.activeDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
	 
	var aStory = app.selection[0].parentStory;
	var tRanges = app.selection[0].textStyleRanges;
	if (app.extractLabel("previous_ref_character").length) {
		var mySample = app.extractLabel("previous_ref_character");
	} else {
		var mySample = "x";
	}
	
	
	
	
	
	if (myHeight == false) {
		
		//var aSize = app.selection[0].pointSize / getRel(aStory, app.selection[0].index, "H");
		//var aPointsize = app.selection[0].pointSize;
		//var aRel = getRel(aStory, app.selection[0].index, "H");
		//var aRel_result = aRel[0]/aRel[1];
		//var aSize = aPointsize/aRel_result;
		//////////////////////////////////////////////

		var myDialog = app.dialogs.add({name:localize(sDialogName) });
		with(myDialog.dialogColumns.add()){
			
			with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: localize(sDefCharacter) });
				var mySampleField = textEditboxes.add({editContents: mySample});
			}
			with(dialogRows.add()){
			var mySubstractOvershoot_2_Check  = checkboxControls.add({staticLabel: localize(sOvershoot_2), checkedState:false});
			}
			//with(dialogRows.add()){
			//var mySubstractOvershoot_Check = checkboxControls.add({staticLabel: localize(sOvershoot), checkedState:false});
			//}
			
		}
		var myResult = myDialog.show();
		if(myResult == true){
			
			var emHeight = 500;
			var mySample = mySampleField.editContents;
			//mySubstractOvershoot = mySubstractOvershoot_Check.checkedState;
			mySubstractOvershoot_2 = mySubstractOvershoot_2_Check.checkedState;
			
			// store entered Values
			if (mySample.length > 1) mySample = mySample.substr(0,1);
			app.insertLabel("previous_ref_character", mySample);
			
			
			myDialog.destroy();
		}
		else{
			
			var emHeight = false;
			myDialog.destroy();
		}
	} else {
		// Damit kann main() der gewünschte Wert übergeben werden und muss nicht mehr abgefragt werden.
		
		var emHeight = myEmHeight;
	}
	
	if (emHeight != false) {
			// look if the characters were scaled
			//if (tRanges[0].horizontalScale != 100) resize = true;

		for (var n = 0; n < tRanges.length; n++) {
			if(resize) {
				// scale to 100%
				tRanges[n].horizontalScale=100;
				tRanges[n].verticalScale=100;
			};
			var aRel = getRel(aStory, tRanges[n].index, mySample);
			
			//Calculating now
			//alert("em"+emHeight);
			tRanges[n].pointSize = aRel[0]; //Fontsize in Points
		
			var myovershoot = Math.round((aRel[2])/aRel[0]*1000)/10;
			var measured_emHeight= Math.round((aRel[1])/aRel[0]*1000)/10;
			
			if (mySubstractOvershoot_2) {
			    var myovershootstring = {en: "\n*Measurements are excluding the Overshoot ("+  myovershoot*2 + "em%" +"), which was calculated out of the Baseline Overshoot  (2 x " + myovershoot + "em%) ",
								de: "\n*Angaben sind ohne Überhang (" +  myovershoot*2 + "em%" + "), welcher aus dem Grundlinienüberhang (2 x " + myovershoot + "em%) berechnet wurde." };
			    var ex_overshoot = "*";
			    var overshoot_em = {en: "\n(Baseline Overshoot = " + myovershoot + "em%)",
								de: "\n (Grundlinienüberhang = " + myovershoot + "em%)"};
			    var overshoot_pt = Math.round(aRel[2]*100)/100 + "pt";
			    var overshoot_mm = Math.round(aRel[2]*0.35278*100)/100 + "mm";
			    var display_overshoot_pt_mm  = {en: "\n(Baseline Overshoot = " + overshoot_pt + " = " + overshoot_mm + ")\n",
										de: "\n(Grundlinienüberhang = " + overshoot_pt + " = " + overshoot_mm + ")\n"};
			    
			    var what_relative = "";
			    var what_absolute = "";
			    
			} else {
			    var myovershootstring = "";
			    var ex_overshoot = "";
			    var overshoot_em = "";
			    var display_overshoot_pt_mm = "";
			    var what_relative = { en:" (in Percent to any entered Font Size)", de: " (in Prozent zu jedweder Eingabe-Größe)"};
			    
			    var what_absolute = {en: " (at " + tRanges[n].pointSize + "pt entered Font Size)", de: " (bei " + tRanges[n].pointSize + "pt eingegebener Schriftgröße)"};
			}
			   var myResultTitle = {en: "Height of ›" + mySample + "‹ (" + tRanges[n].appliedFont.fullName + " at "+ Math.round(tRanges[n].verticalScale*100)/100 +"%)\n",
							de: "Vertikale Größe von ›" + mySample + "‹ (" + tRanges[n].appliedFont.fullName + " bei "+ Math.round(tRanges[n].verticalScale*100)/100 +"%)\n" }
			alert( myResultTitle
				+ localize({en: "Relative", de: "Relativ"}) + localize(what_relative) + "\n" 
				+ "= " + measured_emHeight + "em%" + ex_overshoot
				+ localize(overshoot_em)
				+ "\n\n"
				+ localize({en: "Absolute", de: "Absolut"}) + localize(what_absolute) + "\n" 
				+ "= "+ Math.round(aRel[1]*100)/100 + "pt" + ex_overshoot + " "
				+ "= " + Math.round(aRel[1]*0.35278*100)/100 + "mm"  + ex_overshoot
				+ localize(display_overshoot_pt_mm)
				+ "\n"
				+ localize(myovershootstring));
				
			
		}
	}
	app.activeDocument.viewPreferences.horizontalMeasurementUnits = prevHUnits;
	app.activeDocument.viewPreferences.verticalMeasurementUnits = prevVUnits;	
	mainFuncArray = [emHeight];
	return mainFuncArray;
}

function getRel(aStory, anIX, mySample) {
	aStory.characters[anIX].select();
	app.copy();
	var aPageWidth = app.activeDocument.documentPreferences.pageWidth;
	var aPageHeight = app.activeDocument.documentPreferences.pageHeight;
	var aFrame = app.layoutWindows[0].activePage.textFrames.add({geometricBounds: [0,0,aPageHeight, aPageWidth]});
	aFrame.insertionPoints[0].select();
	app.paste();
	
	var OriginalFontSize = aStory.characters[anIX].pointSize;
	var aChar = aFrame.characters[0];
	
	//reset baseline / leading
	aChar.baselineShift = 0;
	aChar.leading = aChar.pointSize //* 1.2;
	
	/////Look if there is a OVERSET/overflow text -> if so, correct it.
	/////////////////////////////////////////////
	
	if(aFrame.overflows || aChar.pointSize>(aPageHeight/4)) { //if too big
	    //set the size to half of the pages height
	    pointsizemeasuringcorrection = 1/OriginalFontSize *(aPageHeight/4);
	    aFrame.parentStory.paragraphs.item(0).pointSize = OriginalFontSize * pointsizemeasuringcorrection;
	  
	    //place the Character, which we want to measure
	    aChar.contents = mySample;
	    /////////////////////////////////////////////
	    // look if there is still no overflow, with the actual sample glyph
	     if (aFrame.overflows) {  aFrame.geometricBounds = [0,0,aPageHeight*2, aPageWidth*2];
	     } else { pointsizemeasuringcorrection = 1/OriginalFontSize *10;
			aFrame.parentStory.paragraphs.item(0).pointSize = OriginalFontSize * pointsizemeasuringcorrection;};//backup -> set the font to 10pt
	     
	  };
	  
	 
	  ///Min Size to measure
	  if (OriginalFontSize < 1) {
	    pointsizemeasuringcorrection = 1/OriginalFontSize * 1;
	    aFrame.parentStory.paragraphs.item(0).pointSize = OriginalFontSize * pointsizemeasuringcorrection;
	  }
	  
	  
	//read Character
	/////////////////////////////////////////////
	
	aChar.contents = mySample;
	var aSize = aChar.pointSize; 
	var aPath = aChar.createOutlines();
	
	//read baseline
	myRectangle = aFrame.parentStory.insertionPoints.item(-1).rectangles.add({geometricBounds:[0,0,1,1], strokeWeight:0, strokeColor:myDocument.swatches.item(0)}); // strokeAlign:2
	var baseline = myRectangle.geometricBounds[2];
	
	//read total Height
	var bSize = aPath[0].geometricBounds[2]-aPath[0].geometricBounds[0];
	
	//calc Overshoot
	var overshoot = aPath[0].geometricBounds[2]-baseline;
	
	
	//remove Overshoot
	if (mySubstractOvershoot_2) {
		var bSize = bSize - (overshoot*2);
	};
	
	
	aFrame.remove();
	myGetRelArray=[aSize/pointsizemeasuringcorrection, bSize/pointsizemeasuringcorrection, overshoot/pointsizemeasuringcorrection];
	return myGetRelArray;
}
