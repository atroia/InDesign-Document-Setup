/* --------------------------------------
File Setup
by Aaron Troia (@atroia)
Modified Date: 1/19/23

Description: 
1. Imports Text, Table/Cell, and Object Styles from stylesheet
2. Imports and overrite custom variables (Running Header & Running Header 2)
3. Sets document and footnote options

-------------------------------------- */

// https://yourscriptdoctor.com/indesign-scripts-for-paragraph-styles/

//assumes target document is document 1
var d = app.activeDocument;
var txtVariable = "Running Header";


main();

function main() {
  if (app.documents.length == 0) {
    alert("No documents are open.");
  } else {
    docPrefs();
    loadAllStyles();
    setMatchParagraphTextVariable(d, txtVariable, "title");
    setMatchParagraphTextVariable(d, txtVariable + " 2", "title2");
    footnoteOptions();
  }
}

function docPrefs(){
  // set ruler mesurements to Points for bleed
  d.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.INCHES; // POINTS, PICAS, INCHES, INCHES_DECIMAL, MILLIMETERS, CENTIMETERS, CICEROS, CUSTOM, AGATES, U, BAI, MILS, PIXELS, Q, HA, AMERICAN_POINTS
  d.viewPreferences.verticalMeasurementUnits = MeasurementUnits.INCHES; // POINTS, PICAS, INCHES, INCHES_DECIMAL, MILLIMETERS, CENTIMETERS, CICEROS, CUSTOM, AGATES, U, BAI, MILS, PIXELS, Q, HA, AMERICAN_POINTS
  // Document Setup : Bleeds
  d.documentPreferences.documentBleedBottomOffset = .125;
  d.documentPreferences.documentBleedInsideOrLeftOffset = .125;
  d.documentPreferences.documentBleedOutsideOrRightOffset = .125;
  d.documentPreferences.documentBleedTopOffset = .125;
  // Preferences : Type
  d.textPreferences.typographersQuotes = true;
  d.textPreferences.deleteEmptyPages = false;
  // Preferences : Guides & Pasteboard
  // d.documentPreferences.marginGuideColor = UIColors.MAGENTA;
  // d.documentPreferences.columnGuideColor = UIColors.VIOLET;
  // d.pasteboardPreferences.bleedGuideColor = UIColors.RED;
  // d.pasteboardPreferences.slugGuideColor = UIColors.GRID_BLUE;
  // d.pasteboardPreferences.matchPreviewBackgroundToThemeColor = false;
  // d.pasteboardPreferences.previewBackgroundColor = UIColors.LIGHT_GRAY;
  // d.guidePreferences.rulerGuidesColor = UIColors.LIGHT_GRAY; // This setting isn't in the user interface
  // app.smartGuidePreferences.guideColor = UIColors.GRID_GREEN;
  d.pasteboardPreferences.pasteboardMargins = ["864 pt", "72 pt"]; // Horizontal, vertical. A horizontal margin of -1 means one document page width.
  // Preferences : Units & Increments
  d.textPreferences.kerningKeyIncrement = 1; // 1-100
  // Preferences : File Handling
  app.fontSyncPreferences.autoActivateFont = false;
  // Type > Show hidden characters
  d.textPreferences.showInvisibles = true;
  // View > Extras
  d.viewPreferences.showFrameEdges = true;
  // View > Grids & Guides
  d.guidePreferences.guidesShown = true;
  d.guidePreferences.guidesLocked = false;
  d.documentPreferences.columnGuideLocked = true;
  d.guidePreferences.guidesSnapto = true;
  app.smartGuidePreferences.enabled = true;
  d.gridPreferences.baselineGridShown = false;
  d.gridPreferences.documentGridShown = false;
  d.gridPreferences.documentGridSnapto = false;
}

function loadAllStyles(){
  var docRef = app.documents.item(0);
  var appPath = app.filePath.absoluteURI;
  var folderPath = "/Volumes/Active/Print/__TEMPLATES__/BookStylesheet/";
  var fileNames = "BookStylesheet_CC2022.indd";
  // var folderPath = appPath + "/Templates";
  try {
    var fileRef = File(folderPath + "/" + fileNames);
    // by default, globalClashResolutionStrategy is LOAD_ALL_WITH_OVERWRITE
    // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#ImportFormat.html
    docRef.importStyles(
      ImportFormat.TEXT_STYLES_FORMAT,
      fileRef,
      GlobalClashResolutionStrategy.LOAD_ALL_WITH_OVERWRITE
    );
    docRef.importStyles(
      ImportFormat.TABLE_AND_CELL_STYLES_FORMAT,
      fileRef,
      GlobalClashResolutionStrategy.LOAD_ALL_WITH_OVERWRITE
    );
    docRef.importStyles(
      ImportFormat.OBJECT_STYLES_FORMAT,
      fileRef,
      GlobalClashResolutionStrategy.LOAD_ALL_WITH_OVERWRITE
    );
    // https://www.indesignjs.de/extendscriptAPI/indesign-latest/#Document.html#d1e49413__d1e54569
    // docRef.loadMasters (masterRef, GlobalClashResolutionStrategyForMasterPage.LOAD_ALL_WITH_OVERWRITE);
  } catch (e) {
    alert(e);
  }
}

function setMatchParagraphTextVariable(docRef, name, style) {
  try {
    try {
      var stylePresent = docRef.paragraphStyles.item(style).name;
      if (stylePresent != style) {
        throw new Exception("StyleDoesntMatch");
      }
    } catch (noStylePresent) {
      alert(
        "Missing Paragraph Style: " + style + ". Please create or import it.",
        "Paragraph Style Missing",
        true
      );
      return false;
    }

    try {
      var update = 0;
      var newTV = docRef.textVariables.add({
        name: name,
        variableType: VariableTypes.MATCH_PARAGRAPH_STYLE_TYPE,
      });
    } catch (variablePresent) {
      var newTV = docRef.textVariables.item(name);
      update = 1;
    } finally {
      newTV.variableOptions.appliedParagraphStyle = style;
      // if (update == 1) {
      //   alert(
      //     "Success!!! Style updated: " +
      //       name +
      //       ". Paragraph style applied: " +
      //       style,
      //     "Paragraph Style Text Variable Applied",
      //     false
      //   );
      // } else {
      //   alert(
      //     "Success!!! Style added: " +
      //       name +
      //       ". Paragraph style applied: " +
      //       style,
      //     "Paragraph Style Text Variable Added",
      //     false
      //   );
      // }
      return true;
    }
  } catch (failSilently) {
    alert(
      "Failed to add " +
        name +
        " style. It may already be present, or you may need to add it manually.",
      "General Error",
      true
    ); //Don't fail as silently...
    //		var localError = failSilently;
    return false;
  }
}

function footnoteOptions() {
  /* ==================================== */
  /* ====  Numbering and Formatting  ==== */
  /* ==================================== */

  // ----- NUMBERING ----- //

  // Style:
  d.footnoteOptions.footnoteNumberingStyle = FootnoteNumberingStyle.KANJI; // 1, 2, 3, 4...
  // d.footnoteOptions.footnoteNumberingStyle = FootnoteNumberingStyle.SYMBOLS;
  // Start At:
  d.footnoteOptions.startAt = 1;
  // Restart Numbering Every:
  d.footnoteOptions.restartNumbering = FootnoteRestarting.PAGE_RESTART;
  d.footnoteOptions.showPrefixSuffix = FootnotePrefixSuffix.NO_PREFIX_SUFFIX;
  d.footnoteOptions.prefix;
  d.footnoteOptions.suffix;

  // ----- FORMATTING ----- //

  // Footnote Reference Number in Text
  // Position:
  d.footnoteOptions.markerPositioning =
    FootnoteMarkerPositioning.SUPERSCRIPT_MARKER;
  // Character Style:
  d.footnoteOptions.footnoteMarkerStyle = d.characterStyleGroups
    .item("Superscript")
    .characterStyles.item("superscript");

  // Footnote Formatting
  // Paragraph Style
  d.footnoteOptions.footnoteTextStyle = d.paragraphStyleGroups
    .item("Footnotes")
    .paragraphStyles.item("footnote (symbols)");
  // Seperator:
  d.footnoteOptions.separatorText = "";

  /* ==================================== */
  /* =============  Layout  ============= */
  /* ==================================== */

  // Span footnotes across columns
  d.footnoteOptions.enableStraddling = true;

  // ----- SPACING OPTIONS ----- //

  // Minimum Space Before First Footnote:
  d.footnoteOptions.spacer = 9;

  // Space Between Footnotes:
  d.footnoteOptions.spaceBetween = 0;

  // ----- FIRST BASELINE ----- //

  // Offset
  d.footnoteOptions.footnoteFirstBaselineOffset =
    FootnoteFirstBaseline.LEADING_OFFSET;
  // Min
  d.footnoteOptions.footnoteMinimumFirstBaselineOffset = 0;

  // ----- PLACEMENT OPTIONS ----- //

  // Place End of Story Footnotes at Bottom of Text
  d.footnoteOptions.eosPlacement = false;
  // Allow Split Footnotes
  d.footnoteOptions.noSplitting = true;

  // ----- RULE ABOVE OPTIONS ----- //

  // Rule Above:
  d.footnoteOptions.continuingRuleOn = false;
  // Rule on:
  d.footnoteOptions.ruleOn = true;
  // Weight:
  d.footnoteOptions.ruleLineWeight = 0.3;
  // Type:
  d.footnoteOptions.ruleType = "Solid";
  // Color:
  d.footnoteOptions.ruleColor = "Black";
  // Tint:
  d.footnoteOptions.ruleTint = 100;
  // Overprint Stroke
  d.footnoteOptions.ruleOverprint = false;
  // Gap Color:
  d.footnoteOptions.ruleGapColor;
  // Gap Tint:
  d.footnoteOptions.ruleGapTint;
  // Overprint Gap
  d.footnoteOptions.ruleGapOverprint;
  // Left Indent:
  d.footnoteOptions.ruleLeftIndent = 0;
  // Width:
  d.footnoteOptions.ruleWidth = 72;
  // Offset:
  d.footnoteOptions.ruleOffset = 0;
}
