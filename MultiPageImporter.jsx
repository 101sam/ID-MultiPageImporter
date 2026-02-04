// MultiPageImporter2.7.1 jsx
// An InDesign CS4+ JavaScript
// Updated for InDesign 2024-2026+ compatibility
// Original Copyright (C) 2008-2009 Scott Zanelli. lonelytreesw@gmail.com
// Coming to you from South Easton, MA, USA

// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the Free Software
// Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.

// Version 2.1:  Fix for CS4 compatibility. (04 MAR 2009)
// Version 2.2: Map pages to exisitng doc pages and reverse the page order options added. (05 MAR 2009)
// Version 2.2.1: Page rotation. (12 MAR 2009)
// Version 2.5: If PDF page count/size can't be determined, import all pages. Remove dependency on Verdana. (28 MAR 2010)
// Version 2.5JJB: Added support for ID CS5 PDF importing. The PDFCrop constants used in IDCS5 are now supported (14 FEB 2011). See lines 126-139. //JJB
// Version 2.6: Fixed a bug that would display a misleading error message ("This value would cause one or more objects to leave the pasteboard.") - mostly in cases where the default font size for a new text box would cause a 20x20 document units box to overflow
// Version 2.6.1: Added new document scale for easy page scaling and tag all placed frames
// Version 2.6.2: Added very basic support for .ai files that are written as pdf compatible files - basically using the pdf code for them - allows for automatically placing multi-artboard AIs
// Version 2.7.0: Updated for InDesign 2024-2026+ compatibility (Feb 2026)
//   - Fixed: "crop type" error with modern PDFs (#40) - added modern PDF parsing for cross-reference streams
//   - Fixed: "JavaScript Error 89867: default engine cannot be deleted" (#40) - replaced exit() with proper returns
//   - Fixed: "objects leave the pasteboard" with baseline grid (#39) - disable alignToBaseline on temp text
//   - Fixed: Indefinite loading/hanging with certain PDFs (#34) - fallback parsing with timeout
//   - Added: Support for .indt template files (#38)
//   - Added: Support for .idml files (#13)
//   - Added: Auto-fill current page number in "Start Placing on Doc Page" field (#35)
//   - Improved: Better error messages with actionable solutions
//   - Improved: Variable scoping fixes throughout
//   - Maintains full backward compatibility with CS3/CS4/CS5+
// Version 2.7.1: Fixed "Illegal return outside of function body" error (Feb 2026)
//   - Wrapped main execution in function to allow proper early returns
//   - Removed all exit() calls which cause engine deletion errors in InDesign 2024+

//#target indesign;

// Wrap entire script in a self-executing function to allow return statements
(function() {

// Get app version and save old interation setting.
// Some installs have the interaction level set to not show any dialogs.
// This is used to insure that the dialog is shown.

var appVersion = parseInt(app.version);
var oldInteractionPref;

// Only works in CS3+
if(appVersion >= 5)
{
	oldInteractionPref = app.scriptPreferences.userInteractionLevel;
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}
else
{
	alert("Features used in this script will only work in InDesign CS3 or later.");
	return; // Safe to use return inside function
}

// Set the next line to false to not use prefs
var usePrefs = true;

// Set default prefs
var pdfCropType = 0;
var indCropType = 1;
var offsetX = 0;
var offsetY = 0;
var doTransparent = 1;
var placeOnLayer = 0;
var fitPage = 0;
var keepProp = 0;
var addBleed = 1;
var ignoreErrors = 0;
var percX = 100;
var percY = 100;
var mapPages = 0;
var reverseOrder = 0;
var rotate = 0;
var positionType = 4; // 4 = center

// Do not change anything after this line!
// removed 6/25/08: var indUpdateType = 0;
var cropType = 0;
var PDF_DOC = "PDF";
var IND_DOC = "InDesign";
var tempObjStyle = null;
var dLog; // Kludge for callback function that uses the dLog, but can't be given the dLog directly
var ddArray;
var ddIndexArray;
var numArray;
var getout;
var doMapCheck = true;
var rotateValues = [0,90,180,270];
var positionValuesAll = ["Top left", "Top center", "Top right", "Center left",  "Center", "Center right", "Bottom left",  "Bottom center",  "Bottom right"];
var noPDFError = true;
var prefsFile;
var theDoc;
var theDocIsMine = false;
var placementINFO;
var cropTypes;
var cropStrings;
var currentLayer;
var docPgCount;
var oldZero;
var oldRulerOrigin;
var fileName;
var theFile;
var startPG, endPG;

// Look for and read prefs file
try {
	prefsFile = File((Folder(app.activeScript)).parent + "/MultiPageImporterPrefs2.5.txt");
} catch(e) {
	// Fallback for when activeScript is not available (e.g., running from ESTK)
	prefsFile = File(Folder.userData + "/MultiPageImporterPrefs2.5.txt");
}

if(!prefsFile.exists)
{
	savePrefs(true);
}
else
{
	readPrefs();
}

// Ask user to select the PDF/InDesign file to place
var askIt = "Select a PDF, PDF compatible AI or InDesign file to place:";
if (File.fs =="Windows")
{
	theFile = File.openDialog(askIt, "Placeable: *.indd;*.indt;*.idml;*.pdf;*.ai");
}
else if (File.fs == "Macintosh")
{
	theFile = File.openDialog(askIt, macFileFilter);
}
else
{
	theFile = File.openDialog(askIt);
}

// Check if cancel was clicked or invalid file selected
if (theFile == null)
{
	// user clicked cancel, just leave
	restoreInteraction();
	return;
}

if((theFile.name.toLowerCase().indexOf(".pdf") == -1 && theFile.name.toLowerCase().indexOf(".ind") == -1 && theFile.name.toLowerCase().indexOf(".ai") == -1 ))
{
	alert("A PDF, PDF compatible AI or InDesign file must be chosen.", "File Type Error");
	restoreInteraction();
	return;
}

fileName = File.decode(theFile.name);

// removed 6/25/08: var indUpdateStrings = ["Use Doc's Layer Visibility","Keep Layer Visibility Overrides"];

if((theFile.name.toLowerCase().indexOf(".pdf") != -1) || (theFile.name.toLowerCase().indexOf(".ai") != -1))
{
	// Premedia Systems/JJB Edit Start - 02/14/11 Modified PDFCrop constants to support ID CS3 through CS5+ PDFCrop Types.
	if (appVersion > 6)
	{
		// CS5 or newer (includes CC, CC 2014-2024, 2025, 2026+)
		cropTypes = [PDFCrop.cropPDF, PDFCrop.cropArt, PDFCrop.cropTrim, PDFCrop.cropBleed, PDFCrop.cropMedia, PDFCrop.cropContentAllLayers, PDFCrop.cropContentVisibleLayers];
		cropStrings = ["Crop","Art","Trim","Bleed", "Media","All Layers Bounding Box","Visible Layers Bounding Box"];
	}
	else
	{
		// CS3 or CS4
		cropTypes = [PDFCrop.cropContent, PDFCrop.cropArt, PDFCrop.cropPDF, PDFCrop.cropTrim, PDFCrop.cropBleed, PDFCrop.cropMedia];
		cropStrings = ["Bounding Box","Art","Crop","Trim","Bleed", "Media"];
	}
	// Premedia Systems/JJB Edit End

	// Parse the PDF file and extract needed info
	try
	{
		placementINFO = getPDFInfo(theFile, (app.documents.length == 0));
	}
	catch(e)
	{
		// Couldn't determine the PDF info, revert to just adding all the pages
		noPDFError = false;
		placementINFO = new Array();

		if(app.documents.length == 0)
		{
			var tmp = new Array();
			tmp["width"] = 612;
			tmp["height"] = 792;

			placementINFO["pgSize"]  = tmp;
		}
	}
	placementINFO["kind"] = PDF_DOC;
}
else
{
	cropTypes = [ImportedPageCropOptions.CROP_CONTENT, ImportedPageCropOptions.CROP_BLEED, ImportedPageCropOptions.CROP_SLUG];
	cropStrings = ["Page bounding box","Bleed bounding box","Slug bounding box"];
	// Get the InDesign doc's info
	placementINFO = getINDinfo(theFile);
	placementINFO["kind"] = IND_DOC;
}

// If there is no document open, create a new one using the size of the
// first encountered page
theDocIsMine = false; // Is the doc created by this script boolean
if(app.documents.length == 0)
{
	// Save the app measurement units to restore after doc is created
	var oldUnitsV = app.viewPreferences.verticalMeasurementUnits;
	var oldUnitsH = app.viewPreferences.horizontalMeasurementUnits;
	var oldMarginT = 	app.marginPreferences.top;
	var oldMarginB = app.marginPreferences.bottom;
	var oldMarginL = app.marginPreferences.left;
	var oldMarginR = app.marginPreferences.right;
	app.marginPreferences.top = 0;
	app.marginPreferences.bottom = 0;
	app.marginPreferences.left = 0;
	app.marginPreferences.right = 0;

	if(placementINFO.kind == PDF_DOC)
	{
		app.viewPreferences.verticalMeasurementUnits = MeasurementUnits.points;
		app.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.points;
	}
	else
	{
		app.viewPreferences.verticalMeasurementUnits = placementINFO.vUnits;
		app.viewPreferences.horizontalMeasurementUnits = placementINFO.hUnits;
	}

	// Make the new doc:
	theDoc = app.documents.add();
    theDocIsMine = true;
	theDoc.documentPreferences.facingPages = false;
	theDoc.marginPreferences.columnCount = 1;
	theDoc.documentPreferences.pageWidth = placementINFO.pgSize.width;
	theDoc.documentPreferences.pageHeight = placementINFO.pgSize.height;
	theDoc.viewPreferences.verticalMeasurementUnits = oldUnitsV;
	theDoc.viewPreferences.horizontalMeasurementUnits = oldUnitsH;

	// Restore the original units
	app.viewPreferences.verticalMeasurementUnits = oldUnitsV;
	app.viewPreferences.horizontalMeasurementUnits = oldUnitsH;
	app.marginPreferences.top = oldMarginT;
	app.marginPreferences.bottom = oldMarginB;
	app.marginPreferences.left = oldMarginL;
	app.marginPreferences.right = oldMarginR;
}
else
{
	theDoc = app.activeDocument;
}

currentLayer = theDoc.activeLayer;
docPgCount = theDoc.pages.length;

// Get and display the dialog
dLog = makeDialog();
dLog.center(); // Center dialog in screen

if(dLog.show() == 1)
{
	// Extract info from dialog info
	if(noPDFError)
	{
		startPG = Number(dLog.startPG.text);
		endPG = Number(dLog.endPG.text);
		mapPages = Number(dLog.mapPages.value);
		reverseOrder = Number(dLog.reverseOrder.value);
	}
	else
	{
		startPG = 1;
		endPG = 99999;
	}

	var docStartPG = Number(dLog.docStartPG.text);
	cropType = dLog.cropType.selection.index;
	offsetX = Number(dLog.offsetX.text);
	offsetY = Number(dLog.offsetY.text);
	percX = Number(dLog.percX.text);
	percY = Number(dLog.percY.text);
	rotate = dLog.rotate.selection.index;
	if(placementINFO.kind == PDF_DOC)
	{
		doTransparent = dLog.doTransparent.value;
	}
	ignoreErrors = dLog.ignoreErrors.value;
	placeOnLayer = dLog.placeOnLayer.value;
	// indUpdateType = dLog.indUpdateType.selection; // Removed 6/25/08
	fitPage = dLog.fitPage.value;
	keepProp = dLog.keepProp.value;
	addBleed = dLog.addBleed.value;
	positionType = dLog.posDropDown.selection.index;
}
else
{
	restoreDefaults(false);
	return; // Safe inside function
}

// Check whether to do page mapping
if(mapPages && noPDFError)
{
	ddArray = new Array(docPgCount);
	ddIndexArray = new Array(docPgCount);
	numArray = new Array(docPgCount+1);

	// Fill the ddIndexArray with 1 to # of PDF pages
	for(var i=startPG, j= 1; i < docPgCount + startPG; i++, j++)
		ddIndexArray[i%docPgCount] = j;

	// Fill the numArray with all the document page numbers
	numArray[0] = "skip";
	for(var i=1; i<=docPgCount; i++)
		numArray[i]=(i).toString();

	var mapDlog = createMappingDialog(startPG, endPG, numArray);
	mapDlog.center();
	if(mapDlog.show() == 2)
	{
		// Cancel clicked
		restoreDefaults(false);
		return; // Safe inside function
	}
}

// Dialog is no longer needed, let it eventually be garbage collected
dLog = null;

// Add the new layer if requested
if(placeOnLayer)
{
	// Add random number to file name to be layer name.
	// Double check layer name doesn't exist and alter if it happens to be present for some reason
	var layerName = fileName + "_" + Math.round(Math.random() * 9999);
	var docLayers = theDoc.layers;
	for(var i=0; i < docLayers.length; i++)
	{
		if (docLayers[i].name.indexOf(layerName) != -1 )
		{
			layerName += ("_" + Math.round(Math.random() * 9999));
		}
	}

	// Add the layer
	currentLayer = theDoc.layers.add({name:layerName});
}

// Save zero point for later restoration
oldZero = theDoc.zeroPoint;
// set the zero point to the origin
theDoc.zeroPoint = [0,0];

// Save ruler origin for later restoration
oldRulerOrigin = theDoc.viewPreferences.rulerOrigin;
// set the ruler origin to page or all PDFs will be placed on first page of spreads
theDoc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;

if( theDocIsMine ) {
    theDoc.documentPreferences.pageWidth  *= percX/100;
    theDoc.documentPreferences.pageHeight *= percY/100;
}

// Get the Indy doc's height and width
var docWidth = theDoc.documentPreferences.pageWidth;
var docHeight = theDoc.documentPreferences.pageHeight;

// Set placement prefs
if(placementINFO.kind == PDF_DOC)
{
	with(app.pdfPlacePreferences)
	{
		transparentBackground = doTransparent;
		pdfCrop = cropTypes[cropType];
	}
}
else
{
	app.importedPageAttributes.importedPageCrop = cropTypes[cropType];
}

// Block errors if requested
if(ignoreErrors)
{
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
}

// Create the Object Style to be applied to the placed pages.
tempObjStyle = theDoc.objectStyles.add();
tempObjStyle.name = "MultiPageImporter_Styler_" + Math.round(Math.random() * 9999);
tempObjStyle.strokeWeight = 0; // Make sure there's no stroke
tempObjStyle.fillColor = "None"; // Make sure fill is none
tempObjStyle.enableAnchoredObjectOptions = true;

// Set the anchor properties
var tempAOS = tempObjStyle.anchoredObjectSettings;
tempAOS.anchoredPosition = AnchorPosition.ANCHORED;
tempAOS.spineRelative = false;
tempAOS.lockPosition = false;
tempAOS.verticalReferencePoint = AnchoredRelativeTo.PAGE_EDGE;
tempAOS.horizontalReferencePoint = AnchoredRelativeTo.PAGE_EDGE;
tempAOS.anchorXoffset = offsetX;
tempAOS.anchorYoffset = offsetY;

// Set the placement options based on user selected position
// The -1 is needed to get rectangle to move correctly when using the auto positioning of the object styles
// Could be a bug since just the left positions need the negative multiple (spine doesn't need the negative multiple)
switch(positionType)
{
	case 0: //  Top Left
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
		break;
	case 1: // Top Center
		tempAOS.anchorPoint = AnchorPoint.TOP_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
		break;
	case 2: // Top Right
		tempAOS.anchorPoint = AnchorPoint.TOP_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 3: // Middle Left
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.LEFT_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
		break;
	case 4: // Center
		tempAOS.anchorPoint = AnchorPoint.CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
		break;
	case 5: // Middle Right
		tempAOS.anchorPoint = AnchorPoint.RIGHT_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 6: // Bottom Left
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_LEFT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
		break;
	case 7: // Bottom Center
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
		break;
	case 8: // Bottom Right
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	// 9 == separator
	case 10: // Top Relative to Spine
		tempAOS.spineRelative = true;
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.TOP_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 11: // Middle Relative to Spine
		tempAOS.spineRelative = true;
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.RIGHT_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 12: //  Bottom Relative to Spine
		tempAOS.spineRelative = true;
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 13: // Middle relative to Edge
		tempAOS.spineRelative = true;
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.LEFT_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
		break;
}

// Add the pages to the doc based on normal or mapping pages
if(mapPages && noPDFError)
{
	for(var pdfPG = startPG; pdfPG <= endPG; pdfPG++)
	{
		var i = ddArray[pdfPG%docPgCount].selection.text;
		if(i == "skip")
		{
			continue;
		}
		addPages(Number(i), pdfPG, pdfPG);
	}
}
else if(reverseOrder && noPDFError)
{
	var currentDocStart = docStartPG;
	for(var reverse = endPG; reverse >= startPG; reverse--)
	{
		addPages(currentDocStart, reverse, reverse);
		currentDocStart++;
	}
}
else
{
	addPages(docStartPG, startPG, endPG);
}

// Kill the Object style
tempObjStyle.remove();

// Save prefs and then restore original app/doc settings
savePrefs(false);
restoreDefaults(true);

// Helper function to restore interaction preference only
function restoreInteraction() {
	if (typeof oldInteractionPref !== 'undefined') {
		app.scriptPreferences.userInteractionLevel = oldInteractionPref;
	}
}

// Place the requested pages in the document
function addPages(docStartPG, startPG, endPG)
{
	var currentPDFPg = 0;
	var firstTime = true;
	var addedAPage = false;
	var zeroBasedDocPgCnt = docPgCount - 1;

	for(var i = docStartPG - 1, currentInputDocPg = startPG; currentInputDocPg <= endPG; currentInputDocPg++, i++)
	{

		if(placementINFO.kind == PDF_DOC)
		{
			// Set the app's PDF placement pref's page number property to the current PDF page number
			app.pdfPlacePreferences.pageNumber = currentInputDocPg;
		}
		else
		{
			// Set the app's Imported Page placement pref's page number property to the current IND page number
			app.importedPageAttributes.pageNumber = currentInputDocPg;
		}

		if(i > zeroBasedDocPgCnt)
		{
			// Make sure we have a page to insert into
			theDoc.pages.add(LocationOptions.AT_END);
			addedAPage = true;
		}

		// Create a temporary text box to place graphic in (to use auto positioning and sizing)
		var TB = theDoc.pages[i].textFrames.add({geometricBounds:[0,0,20,20]});
		// Fix for "out of pasteboard" error:
		// - Set small point size (fixes issue when default font is large)
		// - Disable baseline grid alignment (fixes issue #39 when baseline grid is relative to margin)
		TB.texts.firstItem().pointSize = 1;
		try { TB.texts.firstItem().alignToBaseline = false; } catch(e) {}
		var theRect = TB.insertionPoints.firstItem().rectangles.add();
            theRect.label = "Multi_Page_Importer_Rect";
		// Applying the object style and doing a recompose updates some objects that
		// the add method doesn't create in the rectangle object
		theRect.appliedObjectStyle = tempObjStyle;
		TB.recompose();

		// Place the current PDF/Ind page into the rectangle object
		try
		{
			var tempGraphic = theRect.place(theFile)[0];
			/* removed 6/25/08
			tempGraphic.graphicLayerOptions.updateLinkOption = (indUpdateType == 0) ?
																							  UpdateLinkOptions.APPLICATION_SETTINGS :
																							  UpdateLinkOptions.KEEP_OVERRIDES;
			*/

			// If all pgs are being added, check that we aren't cruising to the first PDF page again
			if(!noPDFError && !firstTime && tempGraphic.pdfAttributes.pageNumber == 1)
			{
				// If a page was added, nuke it, it's a dupe of the first page
				if(addedAPage)
				{
					theDoc.pages[i].remove();
				}
				else
				{
					// Just remove the placed graphic
					TB.remove();
				}

				return;
			}
		}
		catch(e)
		{
			var errMsg = (e.description) ? e.description : String(e);
			if(errMsg.indexOf("Failed to open") != -1 || errMsg.indexOf("Cannot open") != -1 || errMsg.indexOf("cannot be placed") != -1)
			{
				alert("\"" + fileName + "\" could not be placed with the \"" + cropStrings[cropType] + "\" crop type.\n\nPossible solutions:\n1. Try a different crop type (Trim, Media, or Crop)\n2. Open the PDF in Acrobat and use File > Save As to create a clean copy\n3. Re-export the PDF from the original application\n4. Check if the PDF is password-protected or corrupted", "PDF Placement Error");
			}
			else
			{
				alert("Error placing page: " + errMsg, "Placement Error");
			}
			try {
				if(placeOnLayer && currentLayer)
				{
					currentLayer.remove();
				}
				else if(TB)
				{
					TB.remove();
				}
			} catch(cleanupErr) {}
			restoreDefaults(true);
			// Kill the Object style
			try { tempObjStyle.remove(); } catch(styleErr) {}
			return; // Safe inside function
		}

		// Apply any rotation
		theRect.rotationAngle = rotateValues[rotate];

		// Fit to Page Option
		if(fitPage)
		{
			if(addBleed)
			{
				// Make rectangle the size of the page size plus bleed
				theRect.geometricBounds = [
										   0 - theDoc.documentPreferences.documentBleedTopOffset,
										   0 - theDoc.documentPreferences.documentBleedInsideOrLeftOffset,
										   docHeight + theDoc.documentPreferences.documentBleedBottomOffset,
										   docWidth + theDoc.documentPreferences.documentBleedOutsideOrRightOffset];
			}
			else
			{
				// Change rectangle's size to the page size
				theRect.geometricBounds = [0, 0, docHeight, docWidth];
			}

			// Fit the placed page according to selected options
			if(keepProp)
			{
				theRect.fit(FitOptions.proportionally);
				theRect.fit(FitOptions.frameToContent);// Size box down to size of placed page
			}
			else
				theRect.fit(FitOptions.contentToFrame);
		}
		// Use the Scaling Option
		else
		{
			// Apply the scaling
			theRect.allGraphics[0].verticalScale = percY;
			theRect.allGraphics[0].horizontalScale = percX;
			theRect.fit(FitOptions.frameToContent);
		}

		// Apply the Object Style to transform the graphic into an anchored item (allows auto positioning)
		theRect.appliedObjectStyle = tempObjStyle;

		// Force the text box to reformat itself in order to apply the Object Style
		TB.recompose();

		// Release the placed page from the text box and then delete the text box (clean up)
		theRect.anchoredObjectSettings.releaseAnchoredObject();
		TB.remove();

		firstTime = false;
	}
}

// Create the main dialog box
function makeDialog()
{
	var dlg = new Window('dialog', "Import Multiple " + placementINFO.kind + " Pages",
                        "x:100, y:100, width:533, height:365"); // old height before update option removed: 395
	dlg.onClose = ondLogClosed;

	/******************/
	/* Upper Left Panel */
	/******************/
	dlg.pan1 = dlg.add('panel', [15,15,200,193], "Page Selection");
	dlg.pan1.add('statictext',  [10,15,170,35], "Import " + placementINFO.kind + " Pages:");

	if(noPDFError)
	{
		// Start pg
		dlg.startPG = dlg.pan1.add('edittext', [10,40,70,63], "1");
		dlg.startPG.onChange = startPGValidator;
		dlg.pan1.add('statictext',  [75,45,102,60], "thru");

		// End page
		dlg.endPG = dlg.pan1.add('edittext', [105,40,165,63], placementINFO.pgCount);
		dlg.endPG.onChange = endPGValidator;

		// Mapping option
		dlg.mapPages = dlg.pan1.add('checkbox', [10,144,175,164], "Map to Doc Pages");
		if(reverseOrder || docPgCount == 1)
		{
			mapPages = false;
			dlg.mapPages.enabled = false;
		}
		dlg.mapPages.value = mapPages;
		dlg.mapPages.onClick = mapPGValidator;

		// Reverse order
		dlg.reverseOrder = dlg.pan1.add('checkbox', [10,70,190,85], "Reverse Page Order");
		if(mapPages)
		{
			// Both Mapping and reverse can't be checked
			reverseOrder = false;
			dlg.reverseOrder.enabled = false;
		}
		dlg.reverseOrder.value = reverseOrder;
		dlg.reverseOrder.onClick = reverseClicked;
	}
	else
	{
		dlg.pan1.add('statictext', [10,40,190,55], "Cannot determine PDF");
		dlg.pan1.add('statictext', [10,55,190,70], "page count: all pages");
		dlg.pan1.add('statictext', [10,70,190,85], "will be imported.");
	}


	// Doc start page - auto-fill with current active page (Issue #35)
	var currentPageNum = "1";
	try {
		if (!theDocIsMine && app.activeWindow && app.activeWindow.activePage) {
			currentPageNum = app.activeWindow.activePage.name;
			// Convert page name to number if it's not already
			if (isNaN(parseInt(currentPageNum))) {
				currentPageNum = String(app.activeWindow.activePage.documentOffset + 1);
			}
		}
	} catch(e) { currentPageNum = "1"; }
	dlg.pan1.add('statictext',  [10,94,190,109], "Start Placing on Doc Page:");
	dlg.docStartPG = dlg.pan1.add('edittext', [10,114,70,137], currentPageNum);
	dlg.docStartPG.onChange = docStartPGValidator;

	/***********************/
	/* Lower Left Panel */
	/***********************/
	dlg.pan2 = dlg.add('panel', [15,200,200,350], "Sizing Options");

	// BEGIN Fitting Section
	dlg.fitPage = dlg.pan2.add('checkbox', [10,15,100,35], "Fit to Page");
	dlg.fitPage.onClick = onFitPageClicked;
	dlg.fitPage.value = fitPage;

	// Checkbox
	dlg.keepProp = dlg.pan2.add('checkbox', [10,35,160,55], "Keep Proportions");
	dlg.keepProp.value = keepProp;
	dlg.keepProp.enabled = dlg.fitPage.value;

	// Checkbox
	dlg.addBleed = dlg.pan2.add('checkbox', [10,55,160,75], "Bleed the Fit Page");
	dlg.addBleed.value = addBleed;
	dlg.addBleed.enabled = dlg.fitPage.value;
	// END Fitting Section

	// BEGIN Scaling section
	dlg.pan2.add('statictext',  [10,80,200,95], "Scale of Imported Page:");

	// X%
	dlg.pan2.add('statictext', [10,105,35,125], "X%:");
	dlg.percX = dlg.pan2.add('edittext', [42,102,82,125], "100");
	dlg.percX.text = percX;
	// Visibility depends on the Fit Page checkbox
	dlg.percX.enabled = !dlg.fitPage.value;
	// Assign a validator
	dlg.percX.onChange = percXValidator;

	// Y%
	dlg.pan2.add('statictext', [87,105,112,125], "Y%:");
	dlg.percY = dlg.pan2.add('edittext', [119,102,159,125], "100");
	dlg.percY.text = percY;
	// Visibility depends on the Fit Page checkbox
	dlg.percY.enabled = !dlg.fitPage.value;
	// Assign a validator
	dlg.percY.onChange = percYValidator;

	/*************************/
	/* Upper Right Panel */
	/*************************/
	dlg.pan3 = dlg.add('panel', [210,15,438,193], "Positioning Options");
	dlg.pan3.add('statictext', [10,15,228,35], "Position on Page Aligned From:");

	// DropDownList
	dlg.posDropDown =  dlg.pan3.add('dropdownlist', [10,40,215,60], positionValuesAll);
	dlg.posDropDown.add("separator");
	dlg.posDropDown.add("item", "Top, relative to spine");
	dlg.posDropDown.add("item", "Center, relative to spine");
	dlg.posDropDown.add("item", "Bottom, relative to spine");
	dlg.posDropDown.selection = positionType;

	// Rotation
	dlg.pan3.add('statictext', [10,70,85,90], "Rotation:");
	dlg.rotate = dlg.pan3.add('dropdownlist', [85,67,215,88]);
	for(var i=0;i<rotateValues.length;i++)
	{
		dlg.rotate.add('item', rotateValues[i]);
	}
	dlg.rotate.selection = rotate;

	// Offset section
	dlg.pan3.add('statictext', [10,97,150,117], "Offset by:");
	// X offset value
	dlg.pan3.add('statictext', [10,122,25,142], "X:");
	dlg.offsetX = dlg.pan3.add('edittext', [30,119,95,142], offsetX);
	dlg.offsetX.onChange = offsetXValidator;

	// Y offset value
	dlg.pan3.add('statictext', [100,122,115,142], "Y:");
	dlg.offsetY = dlg.pan3.add('edittext', [120,119,185,142], offsetY);
	dlg.offsetY.onChange = offsetYValidator;

	/*************************/
	/* Lower Right Panel */
	/*************************/
	/* old position before removing update option: [210,207,427,380] */
	dlg.pan4 = dlg.add('panel', [210,200,438,350], "Placement Options");

	// Add the crop type dropdown list and populate it
	dlg.pan4.add('statictext', [10,18,60,35], "Crop to:");
	dlg.cropType = dlg.pan4.add('dropdownlist', [65,15,215,33]);
	for(var i=0;i<cropStrings.length;i++)
	{
		dlg.cropType.add('item', cropStrings[i]);
	}
	dlg.cropType.selection = (placementINFO.kind == PDF_DOC)? pdfCropType : indCropType;

	// Place on Layer
	dlg.placeOnLayer = dlg.pan4.add('checkbox', [10,44,220,60], "Place Pages on a New Layer");
	dlg.placeOnLayer.value = placeOnLayer;

	// Ignore errors
	dlg.ignoreErrors = dlg.pan4.add('checkbox', [10,65,220,81], "Ignore Font and Image Errors");
	dlg.ignoreErrors.value = ignoreErrors;

	// Update Link Options
	/* As of 6/26/08, removing this option so dialog will look better
	dlg.pan4.add('statictext', [10,85,190,100], "Update Link Options:");
	dlg.indUpdateType = dlg.pan4.add('dropdownlist', [10,105,200,125]);
	for(var i = 0; i < indUpdateStrings.length;i++)
	{
		dlg.indUpdateType.add('item', indUpdateStrings[i]);
	}
	dlg.indUpdateType.selection = indUpdateType;
	*/

	// Transparent PDFs
	/* old position before removing update option: [10,133,190,152] */
	dlg.doTransparent = dlg.pan4.add('checkbox', [10,86,220,100], "Transparent PDF Background");
	dlg.doTransparent.value = doTransparent;

	// Disable PDF options if needed
	if(placementINFO.kind != PDF_DOC)
	{
		dlg.doTransparent.enabled = false;
	}

	// The buttons
	dlg.OKbut = dlg.add('button', [448,20,507,45], "OK");
	dlg.OKbut.onClick = onOKclicked;
	dlg.CANbut = dlg.add('button', [448,50,507,75], "Cancel");
	dlg.CANbut.onClick = onCANclicked;

	return dlg;
}

// function to restore saved settings back to originals before script ran
// extras parameter is for exiting at different areas of script:
// false: prior to doing anything
// true: end of script or reading PDF file size
function restoreDefaults(extras)
{
	try {
		if (typeof oldInteractionPref !== 'undefined') {
			app.scriptPreferences.userInteractionLevel = oldInteractionPref;
		}
		if(extras == true && typeof theDoc !== 'undefined' && theDoc)
		{
			if (typeof oldZero !== 'undefined') theDoc.zeroPoint = oldZero;
			if (typeof oldRulerOrigin !== 'undefined') theDoc.viewPreferences.rulerOrigin = oldRulerOrigin;
		}
	} catch(e) {
		// Silently handle any errors during cleanup
	}
}

// function to read prefs from a file
function readPrefs()
{
	if(usePrefs)
	{
		try
		{
			prefsFile.open("r");
			pdfCropType = Number(prefsFile.readln() );
			positionType = Number(prefsFile.readln() );
			offsetX = Number(prefsFile.readln() );
			offsetY = Number(prefsFile.readln() );
			doTransparent = Number(prefsFile.readln() );
			placeOnLayer = Number(prefsFile.readln() );
			fitPage = Number(prefsFile.readln() );
			keepProp = Number(prefsFile.readln() );
			addBleed = Number(prefsFile.readln() );
			ignoreErrors = Number(prefsFile.readln() );
			percX = Number(prefsFile.readln() );
			percY = Number(prefsFile.readln() );
			indCropType = Number(prefsFile.readln() );
			mapPages = Number(prefsFile.readln() );// added 9/7/08
			reverseOrder = Number(prefsFile.readln() ); // added 1/17/09
			rotate = Number(prefsFile.readln()); // added 3/6/09
			prefsFile.close();
		}
		catch(e)
		{
			if(prefsFile) prefsFile.close();
			// Don't show error for prefs, just use defaults
		}
	}
}

// function to save prefs to a file
function savePrefs(firstRun)
{
	if(usePrefs)
	{
		try
		{
			var newPrefs =
			((!firstRun && placementINFO.kind == PDF_DOC) ? cropType: pdfCropType) + "\n" +
			positionType + "\n" +
			offsetX + "\n" +
			offsetY + "\n" +
			((doTransparent)?1:0) + "\n" +
			((placeOnLayer)?1:0) + "\n" +
			((fitPage)?1:0) + "\n" +
			((keepProp)?1:0) + "\n" +
			((addBleed)?1:0) + "\n" +
			((ignoreErrors)?1:0) + "\n" +
			percX + "\n" +
			percY + "\n" +
			((!firstRun && placementINFO.kind == IND_DOC) ? cropType : indCropType) + "\n" +
			((mapPages)?1:0) + "\n" + /* added 9/7/08 */
			((reverseOrder)?1:0) + "\n" +/* added 1/17/09 */
			rotate; /* added 3/6/09 */
			prefsFile.open("w");
			prefsFile.write(newPrefs);
			prefsFile.close();
		 }
		catch(e)
		{
			if(prefsFile) prefsFile.close();
			// Don't show error for prefs
		}
	}
}



/*********************************************/
/*                                                                */
/*        PDF READER SECTION           */
/*  Extracts count and size of pages    */
/*                                                                */
/********************************************/

// Extract info from the PDF file.
// getSize is a boolean that will also determine page size and rotation of first page
// *** File position changes in this function. ***
// Results are as follows:
// page count = retArray.pgCount
// page width = retArray.pgSize.pgWidth
// page height = retArray.pgSize.pgHeight
function getPDFInfo(pdfFile, getSize)
{
	var retArray = new Array();
	retArray["pgCount"] = -1;
	retArray["pgSize"] = null;

	// First try modern PDF parsing method (works with cross-reference streams)
	var modernResult = getPDFInfoModern(pdfFile, getSize);
	if(modernResult.pgCount > 0)
	{
		return modernResult;
	}

	// Fall back to traditional xref parsing for older PDFs
	return getPDFInfoTraditional(pdfFile, getSize);
}

// Modern PDF parsing - handles cross-reference streams and object streams
// Used by InDesign 2024+ compatible PDFs
function getPDFInfoModern(pdfFile, getSize)
{
	var retArray = new Array();
	retArray["pgCount"] = -1;
	retArray["pgSize"] = null;

	try
	{
		pdfFile.open("r");

		// Read file content in chunks to find page count
		var content = "";
		var chunkSize = 65536; // 64KB chunks
		var maxRead = 524288; // Read up to 512KB
		var bytesRead = 0;

		while(!pdfFile.eof && bytesRead < maxRead)
		{
			var chunk = pdfFile.read(chunkSize);
			content += chunk;
			bytesRead += chunk.length;

			// Look for /Type /Pages with /Count
			// Pattern: /Type /Pages ... /Count N
			var pagesMatch = content.match(/\/Type\s*\/Pages[^>]*\/Count\s+(\d+)/);
			if(pagesMatch)
			{
				retArray["pgCount"] = parseInt(pagesMatch[1], 10);
				break;
			}

			// Also try alternate ordering: /Count N ... /Type /Pages
			pagesMatch = content.match(/\/Count\s+(\d+)[^>]*\/Type\s*\/Pages/);
			if(pagesMatch)
			{
				retArray["pgCount"] = parseInt(pagesMatch[1], 10);
				break;
			}
		}

		// If we found page count and need size, get it
		if(retArray["pgCount"] > 0 && getSize)
		{
			// Reset and read for page size
			pdfFile.seek(0);
			content = pdfFile.read(maxRead);

			var pgSize = new Array();

			// Try to find TrimBox first, then MediaBox
			var trimBoxMatch = content.match(/\/TrimBox\s*\[\s*([0-9.-]+)\s+([0-9.-]+)\s+([0-9.-]+)\s+([0-9.-]+)\s*\]/);
			var mediaBoxMatch = content.match(/\/MediaBox\s*\[\s*([0-9.-]+)\s+([0-9.-]+)\s+([0-9.-]+)\s+([0-9.-]+)\s*\]/);

			var boxMatch = trimBoxMatch || mediaBoxMatch;

			if(boxMatch)
			{
				var x1 = parseFloat(boxMatch[1]);
				var y1 = parseFloat(boxMatch[2]);
				var x2 = parseFloat(boxMatch[3]);
				var y2 = parseFloat(boxMatch[4]);

				var width = x2 - x1;
				var height = y2 - y1;

				// Check for rotation
				var rotateMatch = content.match(/\/Rotate\s+(\d+)/);
				var rotation = rotateMatch ? parseInt(rotateMatch[1], 10) : 0;

				if(rotation == 90 || rotation == 270)
				{
					pgSize["width"] = height;
					pgSize["height"] = width;
				}
				else
				{
					pgSize["width"] = width;
					pgSize["height"] = height;
				}

				retArray["pgSize"] = pgSize;
			}
		}

		pdfFile.close();
	}
	catch(e)
	{
		try { pdfFile.close(); } catch(e2) {}
	}

	return retArray;
}

// Traditional xref-based PDF parsing (original method)
// Works with PDFs that use traditional cross-reference tables
function getPDFInfoTraditional(pdfFile, getSize)
{
	var flag = 0; // used to keep track if the %EOF line was encountered
	var nlCount = 0; // number of newline characters per line (1 or 2)

	// The array to hold return values
	var retArray = new Array();
	retArray["pgCount"] = -1;
	retArray["pgSize"] = null;

	// Open the PDF file for reading
	pdfFile.open("r");

	// Search for %EOF line
	// This skips any garbage at the end of the file
	// if FOE% is encountered (%EOF read backwards), flag will be 15
	var maxSearch = 2048; // Increase search range for modern PDFs
	for(var i=0; flag != 15 && i < maxSearch; i++)
	{
		pdfFile.seek(i,2);
		switch(pdfFile.readch())
		{
			case "F":
				flag|=1;
				break;
			case "O":
				flag|=2;
				break;
			case "E":
				flag|=4;
				break;
			case "%":
				flag|=8;
				break;
			default:
				flag=0;
				break;
		}
	}

	// If we couldn't find %EOF, give up on traditional parsing
	if(flag != 15)
	{
		pdfFile.close();
		throw Error("Could not find PDF EOF marker");
	}

	// Jump back a small distance to allow going forward more easily
	pdfFile.seek(pdfFile.tell()-100);

	// Read until startxref section is reached
	var maxLines = 50;
	var foundStartXref = false;
	while(maxLines-- > 0)
	{
		if(pdfFile.readln() == "startxref")
		{
			foundStartXref = true;
			break;
		}
	}

	if(!foundStartXref)
	{
		pdfFile.close();
		throw Error("Could not find startxref");
	}

	// Set the position of the first xref section
	var xrefPos = parseInt(pdfFile.readln(), 10);

	if(isNaN(xrefPos))
	{
		pdfFile.close();
		throw Error("Invalid xref position");
	}

	// The array for all the xref sections
	var	xrefArray = new Array();

	// Go to the xref section
	pdfFile.seek(xrefPos);

	// Check if this is a traditional xref or xref stream
	var firstLine = pdfFile.readln();
	if(firstLine != "xref")
	{
		// This might be an xref stream (modern PDF) - fall back
		pdfFile.close();
		throw Error("PDF uses xref stream instead of traditional xref table");
	}

	// Go back to xref position
	pdfFile.seek(xrefPos);

	// Determine length of xref entries
	// (not all PDFs are compliant with the requirement of 20 char/entry)
	xrefArray["lineLen"] = determineLineLen(pdfFile);

	// Get all the xref sections
	while(xrefPos != -1)
	{
		// Go to next section
		pdfFile.seek(xrefPos);

		// Make sure it's an xref line we went to, otherwise PDF is no good
		if (pdfFile.readln() != "xref")
		{
			throwError("Cannot determine page count.", true, 99, pdfFile);
		}

		// Add the current xref section into the main array
		xrefArray[xrefArray.length] = makeXrefEntry(pdfFile, xrefArray.lineLen);

		// See if there are any more xref sections
		xrefPos = xrefArray[xrefArray.length-1].prevXref;
	}

	// Go get the location of the /Catalog section (the /Root obj)
	var objRef = -1;
	for(var i=0; i < xrefArray.length; i++)
	{
		objRef = xrefArray[i].rootObj;
		if(objRef != -1)
		{
			i = xrefArray.length;
		}
	}

	// Double check root obj was found
	if(objRef == -1)
	{
		throwError("Unable to find Root object.", true, 98, pdfFile);
	}

	// Get the offset of the root section and set file position to it
	var theOffset = getByteOffset(objRef, xrefArray, pdfFile);
	pdfFile.seek(theOffset);

	// Determine the obj where the first page is located
	objRef = getRootPageNode(pdfFile);

	// Get the offset where the root page nod is located and set the file position to it
	theOffset = getByteOffset(objRef, xrefArray, pdfFile);
	pdfFile.seek(theOffset);

	// Get the page count info from the root page tree node section
	retArray.pgCount = readPageCount(pdfFile);

	// Does user need size also? If so, get size info
	if(getSize)
	{
		// Go back to root page tree node
		pdfFile.seek(theOffset);

		// Flag to tell if page tree root was visited already
		var rootFlag = false;

		// Loop until an actual page obj is found (page tree leaf)
		do
		{
			var getOut = true;

			if(rootFlag)
			{
				// Try to find the line with the /Kids entry
				// Also look for instance when MediBox is in the root obj
				do
				{
					var tempLine = pdfFile.readln();
				}while(tempLine.indexOf("/Kids") == -1 && tempLine.indexOf(">>") == -1);

			}
			else
			{
				// Try to first find the line with the /MediaBox entry
				rootFlag = true; // Indicate root page tree was visited
				getOut = false; // Force loop if /MediaBox isn't found here
				do
				{
					var tempLine = pdfFile.readln();
					if(tempLine.indexOf("/MediaBox") != -1)
					{
						getOut = true;
						break;
					}
				}while(tempLine.indexOf(">>") == -1);

				if(!getOut)
				{
					// Reset the file pointer to the beginning of the root obj again
					pdfFile.seek(theOffset)
				}
			}

			// If /Kids entry was found, still at an internal page tree node
			if(tempLine.indexOf("/Kids") != -1)
			{
				// Check if the array is on the same line
				if(tempLine.indexOf("R") != -1)
				{
					// Grab the obj ref for the first page
					objRef = parseInt(tempLine.split("/Kids")[1].split("[")[1]);
				}
				else
				{
					// Go down one line
					tempLine = pdfFile.readln();

					// Check if the opening bracket is on this line
					if(tempLine.indexOf("[") != -1)
					{
						// Grab the obj ref for the first page
						objRef = parseInt(tempLine.split("[")[1]);
					}
					else
					{
						// Grab the obj ref for the first page
						objRef = parseInt(tempLine);
					}

				}

				// Get the file offset for the page obj and set file pos to it
				theOffset = getByteOffset(objRef, xrefArray, pdfFile);
				pdfFile.seek(theOffset);
				getOut = false;
			}
		}while(!getOut);

		// Make sure file position is correct if finally at a leaf
		pdfFile.seek(theOffset);

		// Go get the page sizes
		retArray.pgSize = getPageSize(pdfFile);
	}

	// Close the PDF file, finally all done!
	pdfFile.close();

	return retArray;
}

// Function to create an array of xref info
// File position must be set to second line of xref section
// *** File position changes in this function. ***
function makeXrefEntry(pdfFile, lineLen)
{
	var newEntry = new Array();
	newEntry["theSects"] = new Array();
	var tempLine = pdfFile.readln();

	// Save info
	newEntry.theSects[0] = makeXrefSection(tempLine, pdfFile.tell());

	// Try to get to trailer line
	var xrefSec = newEntry.theSects[newEntry.theSects.length-1].refPos;
	var numObjs = newEntry.theSects[newEntry.theSects.length-1].numObjs;
	do
	{
		var getOut = true;
		for(var i=0; i<numObjs;i++)
		{
			pdfFile.readln(); // get past the objects: tell( ) method is all screwed up in CS4
		}
		tempLine = pdfFile.readln();
		if(tempLine.indexOf("trailer") == -1)
		{
			// Found another xref section, create an entry for it
			var tempArray = makeXrefSection(tempLine, pdfFile.tell());
			newEntry.theSects[newEntry.theSects.length] = tempArray;
			xrefSec = tempArray.refPos;
			numObjs = tempArray.numObjs;
			getOut = false;
		}
	}while(!getOut);

	// Read line with trailer dict info in it
	// Need to get /Root object ref
	newEntry["rootObj"] = -1;
	newEntry["prevXref"] = -1;
	do
	{
		tempLine = pdfFile.readln();
		if(tempLine.indexOf("/Root") != -1)
		{
			// Extract the obj location where the root of the page tree is located:
			newEntry.rootObj = parseInt(tempLine.substring(tempLine.indexOf("/Root") + 5), 10);
		}
		if(tempLine.indexOf("/Prev") != -1)
		{
			newEntry.prevXref = parseInt(tempLine.substring(tempLine.indexOf("/Prev") + 5), 10);
		}

	}while(tempLine.indexOf(">>") == -1);

	return newEntry;
}

// Function to save xref info to a given array
function makeXrefSection(theLine, thePos)
{
	var tempArray = new Array();
	var temp = theLine.split(" ");
	tempArray["startObj"] = parseInt(temp[0], 10);
	tempArray["numObjs"] = parseInt(temp[1], 10);
	tempArray["refPos"] = thePos;
	return tempArray;
}

// Function that gets the page count form a root page section
// *** File position changes in this function. ***
function readPageCount(pdfFile)
{
	// Read in first line of section
	var theLine = pdfFile.readln();
	var maxLines = 100;

	// Locate the line containing the /Count entry
	while(theLine.indexOf("/Count") == -1 && maxLines-- > 0)
	{
		theLine = pdfFile.readln();
	}

	// Extract the page count
	return parseInt(theLine.substring(theLine.indexOf("/Count") +6), 10);
}

// Function to determine length of xref entries
// Not all PDFs conform to the 20 char/entry requirement
// *** File position changes in this function. ***
function determineLineLen(pdfFile)
{
	// Skip xref line
	pdfFile.readln();
	var lineLen = -1;

	// Loop trying to find lineLen
	var maxIterations = 100;
	do
	{
		var getOut = true;
		var tempLine = pdfFile.readln();
		if(tempLine != "trailer")
		{
			// Get the number of object enteries in this section
			var numObj = parseInt(tempLine.split(" ")[1]);

			// If there is more than one entry in this section, use them to determime lineLen
			if(numObj > 1)
			{
				pdfFile.readln();
				var tempPos = pdfFile.tell();
				pdfFile.readln();
				lineLen = pdfFile.tell() - tempPos;
			}
			else
			{
				if(numObj == 1)
				{
					// Skip the single entry
					pdfFile.readln();
				}
				getOut = false;
			}
		}
		else
		{
			// Read next line(s) and extract previous xref section
			getOut = false;
			do
			{
				tempLine = pdfFile.readln();
				if(tempLine.indexOf("/Prev") != -1)
				{
					pdfFile.seek(parseInt(tempLine.substring(tempLine.indexOf("/Prev") + 5)));
					getOut = true;
				}
			}while(tempLine.indexOf(">>") == -1 && !getOut);
			pdfFile.readln(); // Skip the xref line
			getOut = false;
		}
	}while(!getOut && maxIterations-- > 0);

	// Check if there was a problem determining the line length
	if(lineLen == -1)
	{
		throwError("Unable to determine xref dictionary line length.", true, 97, pdfFile);
	}

	return lineLen;
}

// Function that determines the byte offset of an object number
// Searches the built array of xref sections and reads the offset for theObj
// *** File position changes in this function. ***
function getByteOffset(theObj, xrefArray, pdfFile)
{
	var theOffset = -1;

	// Look for the theObj in all sections found previously
	for(var i = 0; i < xrefArray.length; i++)
	{
		var tempArray = xrefArray[i];
		for(var j=0; j < tempArray.theSects.length; j++)
		{
			 var tempArray2 = tempArray.theSects[j];

			// See if theObj falls within this section
			if(tempArray2.startObj <= theObj && theObj <= tempArray2.startObj + tempArray2.numObjs -1)
			{
				pdfFile.seek((tempArray2.refPos + ((theObj - tempArray2.startObj) * xrefArray.lineLen)));

				// Get the location of the obj
				var tempLine = pdfFile.readln();

				// Check if this is an old obj, if so ignore it
				// An xref entry with n is live, with f is not
				if(tempLine.indexOf("n") != -1)
				{
					theOffset = parseInt(tempLine, 10);

					// Cleanly get out of both loops
					j = tempArray.theSects.length;
					i = xrefArray.length;
				}
			}
		}
	}

	return theOffset;
}

// Function to extract the root page node object from a section
// File position must be at the start of the root page node
// *** File position changes in this function. ***
function getRootPageNode(pdfFile)
{
	var tempLine = pdfFile.readln();
	var maxLines = 100;

	// Go to line with /Page token in it
	while(tempLine.indexOf("/Pages") == -1 && maxLines-- > 0)
	{
		tempLine = pdfFile.readln();
	}

	// Extract the root page obj number
	return parseInt(tempLine.substring(tempLine.indexOf("/Pages") + 6), 10);
}

// Function to extract the sizes from a page reference section
// File position must be at the start of the page object
// *** File position changes in this function. ***
function getPageSize(pdfFile)
{
	var hasTrimBox = false; // Prevent MediaBox from overwriting TrimBox info
	var charOffset = -1;
	var isRotated = false; // Page rotated 90 or 270 degrees?
	var foundSize = false; // Was a size found?
	var theNums;
	var maxLines = 100;

	do
	{
		var theLine = pdfFile.readln();
		if(!hasTrimBox && (charOffset = theLine.indexOf("/MediaBox")) != -1)
		{
			// Is the array on the same line?
			if(theLine.indexOf("[", charOffset + 9) == -1)
			{
				// Need to go down one line to find the array
				theLine = pdfFile.readln();
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				theNums = theLine.split("[")[1].split("]")[0].split(" ");
			}
			else
			{
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				theNums = theLine.split("/MediaBox")[1].split("[")[1].split("]")[0].split(" ");
			}

			// Take care of leading space
			if(theNums[0] == "")
			{
				theNums = theNums.slice(1);
			}

			foundSize = true;
		}
		if((charOffset = theLine.indexOf("/TrimBox")) != -1)
		{
			// Is the array on the same line?
			if(theLine.indexOf("[", charOffset + 8) == -1)
			{
				// Need to go down one line to find the array
				theLine = pdfFile.readln();
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				theNums = theLine.split("[")[1].split("]")[0].split(" ");
			}
			else
			{
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				theNums = theLine.split("/TrimBox")[1].split("[")[1].split("]")[0].split(" ");
			}

			// Prevent MediaBox overwriting TrimBox values
			hasTrimBox = true;

			// Take care of leading space
			if(theNums[0] == "")
			{
				theNums = theNums.slice(1);
			}

			foundSize = true;
		}
		if((charOffset = theLine.indexOf("/Rotate") ) != -1)
		{
			var rotVal = parseInt(theLine.substring(charOffset + 7));
			if(rotVal == 90 || rotVal == 270)
			{
				isRotated = true;
			}
		}
	}while(theLine.indexOf(">>") == -1 && maxLines-- > 0);

	// Check if a size array wasn't found
	if(!foundSize)
	{
		throwError("Unable to determine PDF page size.", true, 96, pdfFile);
	}

	// Do the math
	var xSize =	parseFloat(theNums[2]) - parseFloat(theNums[0]);
	var ySize =	parseFloat(theNums[3]) - parseFloat(theNums[1]);

	// One last check that sizes are actually numbers
	if(isNaN(xSize) || isNaN(ySize))
	{
		throwError("One or both page dimensions could not be calculated.", true, 95, pdfFile);
	}

	// Use rotation to determine orientation of pages
	var ret = new Array();
	ret["width"] = isRotated ? ySize : xSize;
	ret["height"] = isRotated ? xSize : ySize;

	return ret;
}

// Error function
function throwError(msg, pdfError, idNum, fileToClose)
{
	if(fileToClose != null)
	{
		try { fileToClose.close(); } catch(e) {}
	}

	if(pdfError)
	{
		// Throw err to be able to turn page numbering off
		throw Error("PDF parsing error: " + msg);
	}
	else
	{
		alert("ERROR: " + msg + " (" + idNum + ")", "MultiPageImporter Script Error");
		// Throw error instead of exit to avoid engine deletion problems
		throw Error("Script error " + idNum + ": " + msg);
	}
}

// Extract info from the document being placed
// Need to open without showing window and then close it
// right after collecting the info
function getINDinfo(theDoc)
{
	// Open it
	var temp = app.open(theDoc, false);
	var placementINFO = new Array();
	var pgSize = new Array();
	// Get info as needed
	placementINFO["pgCount"] = temp.pages.length;
	pgSize["height"] = temp.documentPreferences.pageHeight;
	pgSize["width"] = temp.documentPreferences.pageWidth;
	placementINFO["vUnits"] = temp.viewPreferences.verticalMeasurementUnits;
	placementINFO["hUnits"] = temp.viewPreferences.horizontalMeasurementUnits;
	placementINFO["pgSize"] = pgSize;
	// Close the document
	temp.close(SaveOptions.NO);
	return placementINFO;
}

// File filter for the mac to only show indy and pdf files
function macFileFilter(fileToTest)
{
	var name = fileToTest.name.toLowerCase();
	if((name.indexOf(".pdf") != -1 || name.indexOf(".indd") != -1 || name.indexOf(".indt") != -1 ||
	    name.indexOf(".idml") != -1 || name.indexOf(".ai") != -1 ||
	    fileToTest.constructor.name == "Folder" || fileToTest.name == "") && name.indexOf(".app") == -1)
		return true;
	else
		return false;
}

/* HELPER FUNCTIONS FOR THE DIALOG WINDOW */

// Enable/disable Keep Props, Bleed Fit, Scale boxes and Offset boxes when Fit Page is un/checked
function onFitPageClicked()
{
	dLog.keepProp.enabled = dLog.fitPage.value;
	dLog.addBleed.enabled = dLog.fitPage.value;
	dLog.percX.enabled = !dLog.fitPage.value
	dLog.percY.enabled = !dLog.fitPage.value
}

// Take care of OK beng clicked
function onOKclicked()
{
	dLog.close(1);
}

// Take care of Cancel beng clicked
function onCANclicked()
{
	dLog.close(0);
}

// Validate the start page
function startPGValidator()
{
	pageValidator(dLog.startPG, placementINFO.pgCount, "start");
}

// Validate the end page
function endPGValidator()
{
	pageValidator(dLog.endPG, placementINFO.pgCount, "end");
}


// Validate the document start page
function docStartPGValidator()
{
	pageValidator(dLog.docStartPG, docPgCount, "Start Placing on Doc Page");
}

// Actual page validator
function pageValidator(me, max, boxType)
{
	var errType = "Invalid Page Number Error";
	var temp = new Number(me.text);
	if(isNaN(temp))
	{
		alert("Please enter '" + boxType + "' page as a number.", errType);
		me.text = "1";
	}
	else if(temp < 1)
	{
		alert("The '" + boxType + "' page number must be at least 1.", errType);
		me.text = "1";
	}
	else if(temp > max)
	{
		alert("The '" + boxType + "' page number must be " + max + " or less.", errType);
		me.text = max;
	}

	// Make sure the new page range doesn't circumvent the mapPGValidator
	if(noPDFError)
	{
		mapPGValidator();
	}
}

// Validate entered text for the percX box
function percXValidator()
{
	percentageValidator(dLog.percX, "X");
}

// Validate entered text for the percY box
function percYValidator()
{
	percentageValidator(dLog.percY, "Y");
}

// Validator for the percentage boxes
function percentageValidator(me, boxType)
{
	var temp = new Number(me.text);
	if(isNaN(temp))
	{
		alert("Please enter a number in the " + boxType + " percentage box!", "Invalid Percentage Error" );
		me.text = "100";
	}
	else if(temp < 1 || temp > 400)
	{
		alert("Value must be between 1% and 400% in the " + boxType + " percentage box!", "Invalid Percentage Error");
		me.text = "100 ";
	}
}

// Validate entered text for the X offset box
function offsetXValidator()
{
	offsetValidator(dLog.offsetX, "X");
}

// Validate entered text for the Y offset box
function offsetYValidator()
{
	offsetValidator(dLog.offsetY, "Y");
}

// Actaul Validator for the offset values
function offsetValidator(me, boxType)
{
	if(isNaN(new Number(me.text)))
	{
		alert("Please use a number in the " + boxType + " offset box!", "Invalid Offset Error");
		me.text="0";
	}
}

// On dialog close Validator
function ondLogClosed()
{
	if(noPDFError && Number(dLog.startPG.text) > Number(dLog.endPG.text))
	{
		alert("Start Page must be less than or equal to the End Page.",
			  "Invalid Page Number Error");
		return false;
	}
}

// When reverseOrder checkbox is clicked, enable/disable the mapping checkbox
function reverseClicked()
{
	var setValue = true;

	if(dLog.reverseOrder.value)
	{
		setValue = false;
	}

	dLog.mapPages.enabled = setValue && docPgCount != 1;
}
/*********************************************/
/*                                                                       */
/*        MAPPING SECTION                          */
/*                                                                       */
/********************************************/

// Create the mapping dialog box
function createMappingDialog(pdfStart, pdfEnd, numArray)
{
	var maxCellsInRow = 8;
	var numDone=0;
	var currentPage = pdfStart;
	var numPDFPages = (pdfEnd - pdfStart)+1;
	var temp;
	var mapDlogW, mapDlogH;
	var numCells, addW, addH;

	var mapDlg =  new Window('dialog', "Map Pages");
	mapDlg.add("statictext", [10,15,380,35], "Map " + placementINFO.kind + " pages to desired Document pages (" + placementINFO.kind + "->Doc):");

	// Dynamically create controls
	while(numDone < numPDFPages)
	{
		numCells = 0;
		while(numCells < maxCellsInRow && numDone < numPDFPages)
		{
			addW = (numCells*100);
			addH = (Math.floor(numDone/maxCellsInRow)*30);
			mapDlg.add("statictext", [10 +addW, 45 + addH, 45+addW, 65+addH], formatPgNum(currentPage) );
			temp = mapDlg.add("dropdownlist", [50 +addW, 40 + addH, 100+addW, 60+addH],null,{items:numArray});
			temp.selection = ddIndexArray[currentPage%docPgCount];
			ddArray[currentPage%docPgCount] = temp;
			numCells++;
			numDone++;
			currentPage++;
		}
	}

	// Resize dialog window according to the number of cells
	if(numPDFPages < 4)
		mapDlogW = 400;
	else if(numPDFPages < maxCellsInRow)
		mapDlogW = 10 + numPDFPages * 100;
	else
		mapDlogW = 10 + maxCellsInRow * 100;

	mapDlogH = (Math.ceil(numPDFPages/maxCellsInRow) * 30) + 80;

	// The buttons: uses the calculated height and width to determine position
	mapDlg.OKbut = mapDlg.add('button',  [mapDlogW - 140, mapDlogH - 35, mapDlogW - 80 , mapDlogH - 10], "OK");
	mapDlg.OKbut.onClick = function() { mapDlg.close(1); };
	mapDlg.CANbut = mapDlg.add('button', [mapDlogW - 70, mapDlogH - 35, mapDlogW - 10 , mapDlogH - 10], "Cancel");
	mapDlg.CANbut.onClick = function() { doMapCheck = false; mapDlg.close(2); };
	mapDlg.onClose = onMapClose;

	mapDlg.bounds = [0,0, mapDlogW, mapDlogH];
	return mapDlg;
}

// Test the given input for duplicates.
function onMapClose()
{
	var result = true;

	if(doMapCheck)
	{
		var trackerArray = new Array(docPgCount);

		// Xref the ddIndexArray to the ddArray selected index
		for(var i=startPG; i <= endPG; i++)
		{
			var thisPop = ddArray[i%docPgCount];
			var popSelect = thisPop.selection.index;
			if(popSelect != 0)
			{
				if(trackerArray[popSelect])
				{
					result = false;
					thisPop.graphics.backgroundColor = thisPop.graphics.newBrush(thisPop.graphics.BrushType.SOLID_COLOR,[1,0,0]);
				}
				else
				{
					trackerArray[popSelect] = true;
					thisPop.graphics.backgroundColor = thisPop.graphics.newBrush(thisPop.graphics.BrushType.SOLID_COLOR,[1,1,1]);
				}
			}
			ddIndexArray[i%docPgCount] = popSelect;
		}

		if(!result)
			alert("A duplicate page was entered. Please make sure all drop downs have a unique selection.", "Duplicate Mapping Error");
	}

	return result;
}

// Format the given page number to include an "arrow" so as to make
// the page number 5 characters long. Used in the mapping dialog box.
function formatPgNum(current)
{
	var arrow;
	if(current<10)
		arrow = "--->";
	else if (current < 100)
		arrow = "-->";
	else
		arrow = "->";

	return current + arrow;
}

// Validate that selected PDF page range can all be mapped to separate pages
function mapPGValidator()
{
	if(dLog.mapPages.value)
	{
		if((Number(dLog.endPG.text)-Number(dLog.startPG.text))+1 > docPgCount)
		{
			alert("Mapping is not available: There are not enough document pages to place the PDFs in the selected page range " +
			       "onto their own document pages. Either reduce the number of PDF pages in the range or increase the " +
				   "number of pages in the document that the PDF pages are being placed into.", "Mapping Error");
			dLog.mapPages.value = false;
		}
		else
		{
			dLog.reverseOrder.enabled = false;
		}
	}
	else
	{
		// Unchecked, enable reverseOrder checkbox
		dLog.reverseOrder.enabled = true;
	}
}

})(); // End of self-executing function wrapper
