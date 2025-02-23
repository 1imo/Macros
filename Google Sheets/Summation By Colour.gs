
function SUM_BY_FONT_COLOUR(input) {
  if (!input || !input.map) {
    return [["Error", "Invalid input. Please select a valid range."]];
  }

  var colourSums = {};

  // Sum values by font color
  input.forEach(function(row, rowIndex) {
    row.forEach(function(cell, colIndex) {
      if (cell !== "") {
        var range = SpreadsheetApp.getActiveSheet().getRange(rowIndex + 1, colIndex + 1);
        var colour = range.getFontColor();
        var value = parseFloat(cell);
        
        if (!isNaN(value)) {
          if (!colourSums[colour]) {
            colourSums[colour] = 0;
          }
          colourSums[colour] += value;
        }
      }
    });
  });

  // Prepare output
  var output = [["Color", "Sum"]];
  for (var color in colourSums) {
    output.push([color, colourSums[color]]);
  }

  return output;
}

function process() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  
  // Check if a range is selected
  if (!range || !range.getValues().length) {
    SpreadsheetApp.getUi().alert("Please select a valid range of cells to sum by colour.");
    return;
  }

  var result = SUM_BY_FONT_COLOUR(range.getValues());
  
  // Output the result
  var outputRange = sheet.getRange(1, sheet.getLastColumn() + 2, result.length, result[0].length);
  outputRange.setValues(result);
  
  // Color the cells in the Color column
  var colourColumn = outputRange.offset(1, 0, result.length - 1, 1);
  colourColumn.setFontColors(result.slice(1).map(row => [row[0]]));
  
  // Set left alignment for the result cells
  outputRange.setHorizontalAlignment("left");
  
  SpreadsheetApp.getUi().alert("Summation by colour complete. Results are in the last two columns.");
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Colour Summation')
      .addItem('Sum by Font Colour (Selected Range)', 'process')
      .addToUi();
}
