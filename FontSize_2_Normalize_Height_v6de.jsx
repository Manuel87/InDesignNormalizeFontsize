//Description:Mit dem Skript ist es möglich Schriften einander messtechnisch exakt anzugleichen. Nicht mehr und nicht weniger ;)
// Zur Verwendung mit markiertem Text oder Textrahmen 
// Es wird ein Buchstabe und eine Größe abgefragt. Der eingegebene Buchstabe wird vom Script testweise in Pfade umgewandelt, 
// um im Ersten Schritt zu messen, wie das Verhältnis Kegel/Buchstabenhöhe ist, um darauf aufbauend ins gewünschte Verhältnis proportional zu skalieren.

// Gebrauch, auf eigene Gefahr.
// Wir können nicht garantieren, dass das Script je nach Schrit nicht auch unerwartete Ergebnisse erzeugen kann

// Diese Skript basiert auf dem Skript von Gerald Singelmann:
// 	"SetVisualCharSize.jsx"
// 	http://indesignsecrets.com/set-the-size-of-text-exactly-based-on-cap-or-x-height.php // http://www.indesign-faq.de/de/versal-und-andere-hohen-angleichen
//	cuppascript, 10/2009
//	(he allowed me to use it and publish my modification)


//Autor dieser Weiterentwicklung: Manuel von Gebhardi (author of this script)



//-----------
// T O   D O
//-----------
//STYLES
	// Create Characterstyle Folder ( CharacterStyleGroup )
	// Name Style including: "x-Height = 40em%"

//INTERFACE / CHECKBOXES
	// [x] "within a family keep the size as intended by the typedesigner" (allready implemented, but not as a checkbox) // only works if the family name is exact the same
	// [ ] "exact match (ignore case changes to caps / small caps within Indesign)"


//~INTERFACE in general
	// maybe change to palette instead of dialog
	// add buttons normalize all / normalize selection / detailed options

//WARNINGS / Errorlists
	//if the reference character is a "x" in "simple mode" check also the "o" -> if the differences are too big -> take the size of the "o" and push a message


//FUNCTIONS
	// test for cyrillic
	// add more presets for extended/international version
		 // hebrew
		 // latin italic
		 // chinese
		 // arabic
  
//optimize the way a font is normalized within a whole document
	// (search for apllied characterstyles/paragraphstyles) ?
	// if it is a family-match or an exact match dont measure again (doesnt work yet properly over different textfeilds)



//Preset Values
var pre_Char = "x";
var pre_emHeight = 50;
var errorlist = "";


// enable ExtendScript L10N automatism
$.localize = true;
// Localization-Objects
var sNoSelection = { en: "No text selected", de: "Es ist kein Text markiert." };
var sNoDocument = { en: "No document open", de: "Es ist kein Dokument offen." };


var sDialogName = { en: "Normalize Font-Size (x-Height: ø45–50em% / Capheight: ø70em%) ", de: "Schriftgrößen angleichen (x-Höhe: ø45-50em% / Versalhöhe: ø70%)" };

var sAdvanced = { en: "Advanced Options", de: "Weitere Optionen (Für Fortgeschrittene)" };


  var sDefCharacter = { en: "Reference Character", de: "Maßgebendes Zeichen:" }; //Maßgebendes Zeichen:
  var sMittellaenge = { en: "New Height:", de: "angleichen auf:" }; //Höhe des Zeichens:
  var sCHeight = { en: "New Height:", de: "Neue vertikale Größe:" }; //Höhe des Zeichens:
  var sOvershoot_2 = { en: "[advanced] substract Baseline Overshoot x 2", de: "Doppelten Grundlinienüberhang abziehen"};
  var sCreateCharStyle = { en: "[advanced] create/update ", de: "Erstelle/Aktualisiere "};
//var sOvershoot = { en: "[advanced] substract Baseline Overshoot", de: "[advanced] substract Baseline Overshoot"};
  var sNotes = { en: "Be careful which character you choose.", de: "Achte darauf, welches Zeichen du wählst."};
  var sNotes2 = { en: "The script will take into account everything", de: "Das Skript schließt wirklich alle Formen mit ein, "};
  var sNotes3 = { en: "within, including the point of a small ›i‹, etc.", de: "einschließlich des Punktes auf dem kleinen ›i‹, etc."};
  var sNotes4= { en: "Recommendations: x,o (Mixed Case) / H,O (UPPER CASE)", de: "Empfehlungen: x,o (Mischsatz) / H,O (VERSALSATZ)"};


sDialog2_Title = { en: "Multiple Fonts of the same Family detected", de: "Schriftschnitte der gleichen Familie gefunden"};
sDialog2_Note1 = { en:  "Sure, you want to normalize the Fonts within a Family?", de: "Sicher, dass du die Schriftschnitte auch innerhalb einer Familie angleichen möchtest?"};
sDialog2_Family =  { en: "Yes, I know what I'm doing.", de: "Ja, ich weiß genau was ich tue."};
sDialog2_Note2 = { en:  "If not, the first Font of the respective Family will work as a reference.", de: "Falls nicht, wird nur der erste Schnitt angeglichen, welcher dann als Referenz für die Restlichen dient. " };


var myGetRelArray = [0, 0];
var mainFuncArray = [0,0];
var resize = false;
var myReferenceIndex = 0;
var myStyleIndexNumber = 0;

//var mySubstractOvershoot = false;
var mySubstractOvershoot_2 = false;

var FontFamiliesArray = new Array();
var FontsArray = new Array();
var FontFamiliesScaleArray = new Array();
var FontsScaleArray = new Array();
var PreserveFamily = false; // normally there are no families
var reallydontpreserveFamilies = true;
var myResult2 = false;
var skipOvershootWarning = false;
var skipDuplicateWarning = false;
var myCreateCharacterStyle = false;

//var OverallTextRangeCounter = -1;  //// so it starts with 0 for use with arrays
//Functions
function findArrayReturnIndex(key, array) {
  // The variable results needs var in this case (without 'var' a global variable is created)
  var results = undefined; //[];
  for (var i = 0; i < array.length; i++) {
    if (array[i].indexOf(key) == 0) {
      //results.push(array[i]);
      results = i;
    }
  }
  return results;
}



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
				//var myHeight = false;
				var myEmHeight = false;
				for (var n = 0; n < aSel.length; n++) {
					if (aSel[n].constructor.name == "TextFrame") {
						if (aSel[n].parentStory.textContainers.length == 1) {
							app.select(aSel[n].texts);
							//alert("fu"+myHeight);
							//alert("fu2"+myEmHeight);
							mainFuncArray = main(myEmHeight); //myHeight, 
							myEmHeight = mainFuncArray[0];
							//myHeight = mainFuncArray[0];
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
	 // allStories = app.activeDocument.stories//myDocument.findObject(Story.StoryTypes.REGULAR_STORY) //app.activeDocument.stories.);
	  //alert(allStories);
	  // app.select(allStories.everyItem().textContainers);
	  //(myDocument.layers.item(1)
	  var documentfontcount = myDocument.fonts.length;
	  //len = allStories.everyItem().texts.length; //myDocument.fonts.length// allStories.length
	// alert(myDocument.textFrames.length + " " + myDocument.textFrames.item(1).itemLayer.visible);
	 
	  var myEmHeight = false;
	  var myTextFrames = myDocument.textFrames
	  len = myTextFrames.length;
	   alert("Normalize All " + documentfontcount + " Fonts (" + len + " Textframes) withn this Document / Only a Selection \n • Continue if you want to normalize all fonts within this document.\n • or press ESC and select specific textfields/characters first")
	  
	  
	  for (var i = 0; i < len; i++) {
	    
	    if (myDocument.textFrames.item(i).itemLayer.visible) {
	      app.select(myTextFrames.item(i).texts);
	      //app.select(allStories.item(i).texts);
	  
	   
	   //app.select(aSel[n].texts);
	      mainFuncArray = main(myEmHeight); //myHeight,
	   
	      myEmHeight = mainFuncArray[0];
	    }
	    
	   //alert(myEmHeight);
	  }
	  //alert(app.PageItems.item(0));//.TextFrames);
		//alert(localize(sNoSelection));
	}
}
else{
	alert(localize(sNoDocument));
}

// Main Function – works, when myHeigt==false
function main(myEmHeight) {
	var prevHUnits = myDocument.viewPreferences.horizontalMeasurementUnits;
	myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
	var prevVUnits = myDocument.viewPreferences.verticalMeasurementUnits;
	myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
	 
	var aStory = app.selection[0].parentStory;
	var tRanges = app.selection[0].textStyleRanges;
	
	//var mySample = app.extractLabel("gs_versalhoehe");
	
	if (app.extractLabel("previous_ref_character").length) {
		var mySample = app.extractLabel("previous_ref_character");
		//myReferenceSelectedIndex
	} else {
		var mySample = pre_Char;
	}
	
	if (app.extractLabel("previous_entered_empercent").length) {
		var emHeight = parseFloat(app.extractLabel("previous_entered_empercent"));
		if (isNaN(emHeight)) { emHeight = pre_emHeight; };
	} else {
		var emHeight = pre_emHeight;
	}
	
	
	
	if (myEmHeight == false) {
		
		//var aSize = app.selection[0].pointSize / getRel(aStory, app.selection[0].index, "H");
	//	var aPointsize = app.selection[0].pointSize;
	//	var aRel = getRel(aStory, app.selection[0].index, "H");
	//	var aRel_result = aRel[0]/aRel[1];
		//var aSize = aPointsize/aRel_result;
		//////////////////////////////////////////////
		
		
		var myDialog = app.dialogs.add({name:localize(sDialogName) });
		with(myDialog.dialogColumns.add()){
			
			/*with(dialogRows.add()){
			var myMisch_Check = checkboxControls.add({staticLabel: localize({en:"x-Height (recommended)", de: "x-Höhe (empfohlen)"}), checkedState:true});
			var myVersal_Check = checkboxControls.add({staticLabel: localize({en: "Cap-Height", de: "Versalhöhe"}), checkedState:false});
			*/
			with (dialogRows.add()) {
				
				var myReferenceTypeArray = ["x-Höhe", "Versalhöhe"]; //Kleinbuchstaben / Großbuchtaben // Versalien Minuskeln // Majuskeln / Minuskeln
				
				var myReferenceTypeDropdown = dropdowns.add({minWidth: 20, stringList:myReferenceTypeArray, selectedIndex:myReferenceIndex});
				staticTexts.add({minWidth: 20, staticLabel: localize(sMittellaenge) });
				//var vHeightField = measurementEditboxes.add({editValue: (2.83465 * aSize), editUnits:MeasurementUnits.millimeters, smallNudge:0.5});
				var vHeightField = realEditboxes.add({editValue: (emHeight), smallNudge:0.5}); //measurementEditboxes
				staticTexts.add({minWidth: 20, staticLabel: localize("em%") });
				
			}
			/* with(dialogRows.add()){
			var onlySelection_Check  = checkboxControls.add({staticLabel: localize("Auf Auswahl anwenden"), checkedState:true});
			} */
			
			with(dialogRows.add()){
			var myAdvancedOptions_Check = checkboxControls.add({staticLabel: localize(sAdvanced), checkedState:false});
			}
			//with(dialogRows.add()){
			//var mySubstractOvershoot_Check = checkboxControls.add({staticLabel: localize(sOvershoot), checkedState:false});
			//}
		}
		myResult = myDialog.show();
		if(myResult == true){
		  myAdvancedOptions = myAdvancedOptions_Check.checkedState;
		  //var onlySelection = onlySelection_Check;
		  switch(myReferenceTypeDropdown.selectedIndex){
			case 0:
				var mySample = "x";
				var myReferenceSelectedIndex = 0;
				break;
			case 1:
				var mySample = "H";
				var myReferenceSelectedIndex = 1;
				break;
			};
		  var emHeight = vHeightField.editValue;
		  
		  // store entered Values
			if (mySample.length > 1) mySample = mySample.substr(0,1);
			app.insertLabel("previous_ref_character", mySample);
			app.insertLabel("previous_entered_empercent", emHeight + " ");
		  
		  if (myAdvancedOptions) {
		      skipDuplicateWarning = false;
		  
			var myDialog2 = app.dialogs.add({name:localize(sDialogName) });
			with(myDialog2.dialogColumns.add()){
			
			  with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: localize(sDefCharacter) });
				var mySampleField = textEditboxes.add({editContents: mySample});
			}
			with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: localize(sCHeight) });
				//var vHeightField = measurementEditboxes.add({editValue: (2.83465 * aSize), editUnits:MeasurementUnits.millimeters, smallNudge:0.5});
				var vHeightField = realEditboxes.add({editValue: (emHeight), smallNudge:0.5}); //measurementEditboxes
				staticTexts.add({minWidth: 20, staticLabel: localize("em%") });
			}
			with(dialogRows.add()){
			var mySubstractOvershoot_2_Check  = checkboxControls.add({staticLabel: localize(sOvershoot_2), checkedState:true});
			}
			with(dialogRows.add()){
			var myCreateCharStyle_Check  = checkboxControls.add({staticLabel: localize(sCreateCharStyle), checkedState:false});
			var myStyleArray = ["Zeichenformate", "Absatzformate"];
			//alert(myStyleArray + " " + myStyleIndexNumber) 
			var myStyleDropdown = dropdowns.add({minWidth: 20, stringList:myStyleArray, selectedIndex:myStyleIndexNumber});
				
			}
			with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: " " });	
			}
			with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: localize(sNotes) });
			}
			with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: localize(sNotes2) });
			}
			with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: localize(sNotes3) });
			}
			with (dialogRows.add()) {
				staticTexts.add({minWidth: 180, staticLabel: localize(sNotes4) });
			}
			
			
		      }
		      
		      var myResult2 = myDialog2.show();
		      if(myResult2 == true){
			  
			  //var vHeight = vHeightField.editValue * 0.35278;
			var emHeight = vHeightField.editValue;
			var mySample = mySampleField.editContents;
			//mySubstractOvershoot = mySubstractOvershoot_Check.checkedState;
			mySubstractOvershoot_2 = mySubstractOvershoot_2_Check.checkedState;
			myCreateCharacterStyle = myCreateCharStyle_Check.checkedState;
			
			switch(myStyleDropdown.selectedIndex){
			case 0:
				var myStyleType = "character";
				var myStyleIndex = 0;
				break;
			case 1:
				var myStyleType = "paragraph";
				var myStyleIndex = 1;
				break;
			};
			
			  // store entered Values
			if (mySample.length > 1) mySample = mySample.substr(0,1);
			app.insertLabel("previous_ref_character", mySample);
			app.insertLabel("previous_entered_empercent", emHeight + " ");
			  
			  myDialog2.destroy();
		      } else {
			var emHeight = false;
			myDialog2.destroy();
		      }
		  } else {
		    // Simple Path
		     
		      skipDuplicateWarning = true;
		      skipOvershootWarning = true;
		      mySubstractOvershoot_2 = true;
		      myCreateCharacterStyle = false;
		  }
		  
			
			myDialog.destroy();
		}
		else{
			//var vHeight = false;
			var emHeight = false;
			myDialog.destroy();
		}
	} else {
		// Damit kann main() der gewünschte Wert übergeben werden und muss nicht mehr abgefragt werden.
		//var vHeight = myHeight;
		var emHeight = myEmHeight;
		 
	}
	
	if (emHeight != false) {
	 
			//if (!onlySelection) {
			 //var aStory = myDocument.TextFrames.parentStory;
			//var tRanges = app.selection[0].textStyleRanges;
			//}
			
			
			// look if the characters were scaled
			// had a bug with this, where only the first character in a textfield wasn’t scaled, which resultated in wrong scalement of the following (already scaled fonts)
			//if (tRanges[0].horizontalScale != 100) {resize = true;}
			// Therefore always scale to 100% before measuring
			resize = true;
			
			
		for (var n = 0; n < tRanges.length; n++) {
			if(resize) {
				// scale to 100%
				tRanges[n].horizontalScale=100;
				tRanges[n].verticalScale=100;
			};
			
			//OverallTextRangeCounter = OverallTextRangeCounter + 1;
			var Duplicate = findArrayReturnIndex(tRanges[n].appliedFont.fontFamily, FontFamiliesArray);
			
			try{
			 var Duplicate = findArrayReturnIndex(tRanges[n].appliedFont.fontFamily, FontFamiliesArray);
			 var ExactMatch = findArrayReturnIndex(tRanges[n].appliedFont.fullName, FontsArray);
			
      
			
			  // Record all Fontfamily Names
			  FontFamiliesArray.push(tRanges[n].appliedFont.fontFamily);	
			  FontsArray.push(tRanges[n].appliedFont.fullName);	
			
			} catch (myError){
			 //font is not available!
			  return [emHeight];
			}
			
			
			
			
			
			
			if (Duplicate != undefined) {
				var isDuplicate = true;
				
			      if(skipDuplicateWarning == true) {
				
				PreserveFamily = !reallydontpreserveFamilies//true;
				
				} else {
				  if (myResult3 != true && ExactMatch == undefined) {
					
					
					
					var myDialog3 = app.dialogs.add({name:localize(sDialog2_Title) + " (" + tRanges[n].appliedFont.fontFamily + ")" });
					with(myDialog3.dialogColumns.add()){
						
						with (dialogRows.add()) {
							staticTexts.add({minWidth: 180, staticLabel:localize(sDialog2_Note1) });	
						}
						//with (dialogRows.add()) {
						//	staticTexts.add({minWidth: 180, staticLabel: tRanges[Duplicate].appliedFont.fullName + ", " + tRanges[n].appliedFont.fullName });	
						//}
						with(dialogRows.add()){
						var NormalizeDuplicate_Check  = checkboxControls.add({staticLabel: localize(sDialog2_Family), checkedState:false});
						}
						with (dialogRows.add()) {
							staticTexts.add({minWidth: 180, staticLabel: localize(sDialog2_Note2) });	
						}
						//with(dialogRows.add()){
						//var mySubstractOvershoot_Check = checkboxControls.add({staticLabel: localize(sOvershoot), checkedState:false});
						//}
					};
					var myResult3 = myDialog3.show();
				  };
				
				  if(myResult3 == true){
				  PreserveFamily = !NormalizeDuplicate_Check.checkedState;
				  var myResult3 = true;
				  };
				};
				   
			}  else {
				PreserveFamily = false; // there is no Family  
			}
			
			
				
			
			//var aRel = getRel(aStory, tRanges[n].index, mySample);
			//tRanges[n].pointSize = aRel[0]; //Fontsize in Points
			
			if (PreserveFamily && isDuplicate && FontFamiliesScaleArray[Duplicate]) { 
				
				// Apply the Family Scale from a already measured Font of the same Family
				tRanges[n].horizontalScale = FontFamiliesScaleArray[Duplicate];
				tRanges[n].verticalScale = FontFamiliesScaleArray[Duplicate];
				
			} else if (ExactMatch != undefined && FontsScaleArray[Duplicate]) { //isExactMatch == true //ONLY IF PRESERVERFAMILY == FALSE
				
				// Apply the Font Scale from the exactly same Font, already measured 
				tRanges[n].horizontalScale = FontsScaleArray[Duplicate];
				tRanges[n].verticalScale = FontsScaleArray[Duplicate];
				
			} else {
			      // MEASURE CHARACTER/GLYPH
			      ///////////////////////////////////////////////////////////////////////////////////////
				
				var aRel = getRel(aStory, tRanges[n].index, mySample);
				if (aRel === false) {
				  tRanges[n].horizontalScale = 100;
				  tRanges[n].verticalScale = 100;
				  //alert(tRanges[n].appliedFont.fullName);
				  errorlist += tRanges[n].appliedFont.fullName + ", "
				 return [emHeight];
				}
				
				var fontSize = aRel[0]; //in pt
				var measuredSize = aRel[1];  //in pt
			    
				//CALCULATE New Scale
				/////////////////////////////////////////////////////////////////////////////////////
				
				var measured_emHeight= Math.round(measuredSize/fontSize*1000)/10;
				var aRelPercentage = 100/measured_emHeight*emHeight;  // emHeight = the entered new/wanted emHeight
				
				//SCALE the Font
				//////////////////////////////////////////////////////////////////////////////////////
				tRanges[n].horizontalScale = aRelPercentage;
				tRanges[n].verticalScale = aRelPercentage;
				
				///////////////////////////////////////////////////////////////////////////////////////
				
				
			}
			
			FontFamiliesScaleArray.push(tRanges[n].horizontalScale);	
			FontsScaleArray.push(tRanges[n].horizontalScale);	
			
			//alert(myType);
			//Character Style
			if (myCreateCharacterStyle) {
				
				// Strange Bug Fix:
				if (tRanges[n] != undefined) { 
				// Define new Stylename
					if (isDuplicate &&  PreserveFamily == false) {
						var myStyleName = tRanges[n].appliedFont.fullName;
					}
					else {
						var myStyleName = tRanges[n].appliedFont.fontFamily;
					}
                //alert(myStyleName);
                
				}
				//Create the character style if it does not already exist.
				if (myStyleType == "paragraph") {
				  myStyle = myDocument.paragraphStyles.item(myStyleName);
				} else {
				  myStyle = myDocument.characterStyles.item(myStyleName); }
				
				try{
					myStyle.name;
				}
				catch (myError){
					
					if (myStyleType == "paragraph") {
					myStyle = myDocument.paragraphStyles.add({name:myStyleName}); 
					} else {
					  myStyle = myDocument.characterStyles.add({name:myStyleName}); }
				}
				// assign Values to Character Style
				
				myStyle.horizontalScale = tRanges[n].horizontalScale;
				myStyle.verticalScale = tRanges[n].horizontalScale;
				//myStyle.pointSize = tRanges[n].pointSize;
				
				if (isDuplicate &&  PreserveFamily == false) {
					myStyle.appliedFont = tRanges[n].appliedFont;
				}
				else {
					myStyle.appliedFont = tRanges[n].appliedFont.fontFamily;
				}
				
				
				// assign Style to the text
				if (myStyleType == "paragraph") {
				tRanges[n].appliedParagraphStyle = myStyle;
				} else {
				tRanges[n].appliedCharacterStyle = myStyle;
				}
			
			}
			
			
			
			
			//Calculated Before
			//aRel = aRel[0]/aRel[1];
			//tRanges[n].pointSize = vHeight * aRel;
			
		}
	}
	
	
	myDocument.viewPreferences.horizontalMeasurementUnits = prevHUnits;
	myDocument.viewPreferences.verticalMeasurementUnits = prevVUnits;	
	mainFuncArray = [emHeight];
	return mainFuncArray;
}



// THE MOST IMPORTANT FUNCTION IN THE SCRIPT // MEASURING THE FONT
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
	var pointsizemeasuringcorrection = 1;
	
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
	var aSize = aChar.pointSize; //10
    
    //read baseline
	myRectangle = aFrame.parentStory.insertionPoints.item(-1).rectangles.add({geometricBounds:[0,0,1,1], strokeWeight:0, strokeColor:myDocument.swatches.item(0)}); // strokeAlign:2
	var baseline = myRectangle.geometricBounds[2];
    
    
	try{
	  var aPath = aChar.createOutlines();
	}
	catch (myError){
	  aFrame.remove();
	  return false;
	}
	
	
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

if (errorlist !== "") {
  alert("Please check the following Fonts manually. They couldn’t get normalized\n" + errorlist);
}
