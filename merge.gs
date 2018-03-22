// From: https://github.com/hadaf/SheetsToDocsMerge


/*  This is the main method that should be invoked. 
 *  Copy and paste the ID of your template Doc in the first line of this method.
 *
 *  Make sure the first row of the data Sheet is column headers.
 *
 *  Reference the column headers in the template by enclosing the header in square brackets.
 *  Example: "This is [header1] that corresponds to a value of [header2]."
 */
 
/*  Modified to:
 *    receive parameters
 *    allow for an output destination folder
 *    allow for splitting large datasets into several output docs, to avoid exceeding Google Apps script maximum run time limits
 *    optionally save output in PDF format
 */

function doMerge(
  mergeTemplateGdocId, // ID of the template document
  dataSS, // spreadsheet object, eg SpreadsheetApp.getActiveSpreadsheet()
  dataSheetId,  // Sheet Id
  outputFilenameBase,  // prefix for filenames generated
  outputFolderId,  // Google Drive folder ID
  // Limit the size of the created document to prevent: "Too many changes applied before saving document. Please save changes in smaller batches using Document.saveAndClose()"
  maxDataRowsPerDocument, // 100 is a good maximum if the file needs to be saved as a PDF. Saving 200 or 300 pages to PDF often gave: We're sorry, a server error occurred. Please wait a bit and try again.
  saveAsPDF, // Sometimes fails for large documents, or exceed maximum run time if there are more than 400 rows.
  deleteGdoc, // If saveAsPDF, then optionaly delete Google Doc
  // These next two to allow the merge to be split into multiple instances so that the Google scripts maximum execution time is not exceeded
  skipRows, // number of rows to skip in the dataSheet
  exitAfterOneDocument, // Exit after the maxDataRowsPerDocument have been written to the first document
  addBlankPageAtEnd // Hack for documents where drawings repeat on the last page of the document
) {
  
  if (!maxDataRowsPerDocument) { maxDataRowsPerDocument = 100};
  if (!skipRows) {skipRows=0}; // Default
  
  // Calculated variables
  var templateFile = DriveApp.getFileById(mergeTemplateGdocId);
  var outputFolder = DriveApp.getFolderById(outputFolderId);
  var templateDoc = DocumentApp.openById(mergeTemplateGdocId);
  
  var testing = false;
  if (testing) {
    maxDataRowsPerDocument = 2;
    saveAsPDF=false;
  }


  // var sheet = SpreadsheetApp.getActiveSheet();//current sheet
  var sheet = getSheetById(dataSS, dataSheetId);
  
  var rows = sheet.getDataRange(); // Just rows containing data
  var numRows = rows.getNumRows() -1 ; // -1 for header row
  if (testing) {numRows= maxDataRowsPerDocument }//* 2.5};
  Logger.log("numRows=%s",numRows);
  // Exit if skipping more rows than there are.
  if (skipRows >= numRows) {
    return sprintf('skipRows %s is >= numRows %s', skipRows, numRows);
  }
  
  var maxRowNumberLen= (""+numRows).length; // For left 0-padding the row counter
  Logger.log("maxRowNumberLen=%s",maxRowNumberLen);
  var values = rows.getValues();
  var fieldNames = values[0]; //First row of the sheet must be the the field names
  
  var mergedFile;
  var mergedDoc;
  var numMergedDocs = 0;
  var numMergedRows = 0;
  
  var row = skipRows; // Row counter. Skip some rows if specified
  
  // Create documents with maxDataRowsPerDocument copies of the mergeTemplateGdoc
  do {
    
    // Create next merged document
    // Make a copy of the template file to use for the merged File. Note: It is necessary to make a copy upfront, and do the rest of the content manipulation inside this single copied file, 
    // otherwise, if the destination file and the template file are separate, a Google bug will prevent copying of images from the template to the destination. 
    // See the description of the bug here: https://code.google.com/p/google-apps-script-issues/issues/detail?id=1612#c14
    mergedFile = templateFile.makeCopy(outputFolder);
    // mergedFile.setName("filled_"+templateFile.getName());//give a custom name to the new file (otherwise it is called "copy of ...")
    // var rowRange = pad(row,maxRowNumberLen) +'-'+ pad(Math.min(numRows-1,row-1+maxDataRowsPerDocument),maxRowNumberLen); // eg 001-100
    // At this point, "row" is one less than next row to be processed
    var rowRange = pad(row+1,maxRowNumberLen) +'-'+ pad(Math.min(numRows,row+maxDataRowsPerDocument),maxRowNumberLen); // eg 001-100
    var outputFilename = sprintf('%s%s %s %s', outputFilenameBase, (testing?' testing':''),rowRange, getTimestampStr());
    mergedFile.setName(outputFilename);//give a custom name to the new file (otherwise it is called "copy of ...")
    
    mergedDoc = DocumentApp.openById(mergedFile.getId()); // At this point is the same as the template doc.
    mergedDoc.getBody().clear(); // Clear the body of the mergedDoc so that we can write the new data in it.

    // Copy template file for up to maxDataRowsPerDocument rows
    for (var i=0; i<maxDataRowsPerDocument; i++) {
      if (++row > numRows) { break } // Stop if reach end of the data rows
      numMergedRows++;
      
      var rowVals = values[row];
      var body = templateDoc.getBody().copy();
      
      // Replace document [fieldnames] with actual values
      for (var f = 0; f < fieldNames.length; f++) {
        body.replaceText("(?i)\\[" + fieldNames[f] + "\\]", rowVals[f]);// replace [fieldName] with the respective data value
      }
      // Replace some standard fieldnames
      // @year -> current year
      var now = new Date();
      var year = now.getFullYear();
      body.replaceText("(?i)\\[@year\\]", year.toString());
      
      // Go over all the content of the modified template doc, and copy it into the merged document
      var numChildren = body.getNumChildren();//number of the contents in the template doc
      for (var c = 0; c < numChildren; c++) {
        var child = body.getChild(c);
        child = child.copy();
        if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
          mergedDoc.appendHorizontalRule(child);
        } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
          mergedDoc.appendImage(child.getBlob());
        } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
          mergedDoc.appendParagraph(child);
        } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
          mergedDoc.appendListItem(child);
        } else if (child.getType() == DocumentApp.ElementType.TABLE) {
          mergedDoc.appendTable(child);
        } else {
          sendAlert(sprintf("%s: unknown element type: " + child, arguments.callee.name));
        }
      }
      mergedDoc.appendPageBreak();//Appending page break. Each row will be merged into a new page.
    } // Copy template file for each row i=0; i<maxDataRowsPerDocument;
    if (addBlankPageAtEnd) {mergedDoc.appendPageBreak()}
    
    mergedDoc.saveAndClose();
    numMergedDocs++;
    
    // Save as PDF
    if (saveAsPDF) {
      var docblob = mergedDoc.getAs('application/pdf');
      docblob.setName(mergedDoc.getName() +'.pdf'); // +'.'+ getTimestampStr()+  '.pdf'); // Add the PDF extension
      // When was closing mergedDoc before creating pdf, sometimes get: We're sorry, a server error occurred. Please wait a bit and try again.
      for (i=0; i<3; i++){ // Try 3 times
        try {
          var file = outputFolder.createFile(docblob);
          break;
        }
        catch(e) {};
      }
      
      if (deleteGdoc) mergedFile.setTrashed(true);
      
    } // if saveAsPDF
    
    var m = sprintf('Row %s written to %s', row, mergedDoc.getName());
    SpreadsheetApp.getActiveSpreadsheet().toast(m, "",120);
    Logger.log(m);
    
    if (exitAfterOneDocument) { break }

  } while (row<numRows);

  return sprintf('rows %s-%s merged into %s documents', skipRows+1, skipRows+numMergedRows, numMergedDocs);
}



////////////////////////////////////////////////////////////////////////////////////////////////////
// Supporting functions

function sprintf( format ) {
  // http://www.harryonline.net/scripts/sprintf-javascript/385
  for( var i=1; i < arguments.length; i++ ) {
    format = format.replace( /%s/, arguments[i] );
  }
  return format;
}

////////

function sendAlert(msg){
  Logger.log(msg);
  sendEmail(alertEmailAddress, 'Alert:'+msg);
}

/////////////////////////////////////////////////////////
function sendEmail(recipient, subject, body, options){
  // Change null or undefined parameters to empty strings
  if (!options) options = {}; // Must be a hash
  if (!body) body = '';
  // applog('sending email', recipient, subject, body, options);
  MailApp.sendEmail(recipient, subject, body, options);
  // applog('sent email', recipient, subject);
}

/////////////////////////////////////////////////////////

function applog(event){
  // Optional parameters follow event
  //  Logger.log("arguments.length=%s",arguments.length);
  var numNamedArgs=1;
  var args = Array.prototype.slice.call(arguments).slice(numNamedArgs);
  //  Logger.log("args=%s",args);
  var now = new Date();
  var values = [];
  values.push(now,event);
  // Expand array arguments
  for (var i=0; i<args.length; i++) {
    var v = args[i];
    //    Logger.log("i=%s,v=%s",i.toString(),v);
    if (typeof(v) != 'object'){
      values.push(v);
      continue;
    } 
    // Collapse arrays to 1-dimensional and append to values
    if (Array.isArray(v)){
      values.concat(get1DArray(v)); // http://www.jstips.co/en/flattening-multidimensional-arrays-in-javascript/
      continue;
    }
    // Hash
    values.push(v);
  }
  //  Logger.log("values=%s",values);

  // Append values to log sheet
  var applogSS = SpreadsheetApp.openById(applogSSID);
  var applogSheet = getSheetById(applogSS, applogSheetID);
  applogSheet.appendRow(values);
  Logger.log('applog: '+values);
}

/////////////////////////////////////////////////////////
// https://code.google.com/p/google-apps-script-issues/issues/detail?id=3066

function getSheetById(ss, id) {
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getSheetId() == id) {
      return sheets[i];
    }
  }
  throw new Error(sprintf('%s: no sheet with id %s found in spreadsheet (id:%s) %s', arguments.callee.name,  id, ss.getId(), ss.getName()));
}
