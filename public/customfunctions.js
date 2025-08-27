/* global Office, CustomFunctions */

/**
 * PC.IRR - Takes five cells as input and adds them together
 * @customfunction
 * @param {number} cell1 First cell value
 * @param {number} cell2 Second cell value  
 * @param {number} cell3 Third cell value
 * @param {number} cell4 Fourth cell value
 * @param {number} cell5 Fifth cell value
 * @returns {number} The sum of all five cell values
 */
function IRR(cell1, cell2, cell3, cell4, cell5) {
  // Validate inputs are numbers
  var inputs = [cell1, cell2, cell3, cell4, cell5];
  
  for (var i = 0; i < inputs.length; i++) {
    if (typeof inputs[i] !== 'number' || isNaN(inputs[i])) {
      throw new Error('Cell ' + (i + 1) + ' must be a valid number');
    }
  }
  
  // Add all five values together
  var result = cell1 + cell2 + cell3 + cell4 + cell5;
  console.log('IRR calculation: ' + cell1 + ' + ' + cell2 + ' + ' + cell3 + ' + ' + cell4 + ' + ' + cell5 + ' = ' + result);
  return result;
}

/**
 * PC.JOINCELLS - Joins cells from a range into a single string with specified delimiter
 * @customfunction
 * @param {any[][]} range The range of cells to join
 * @param {string} [delimiter=", "] The delimiter to use (default comma with space)
 * @returns {string} The joined string
 */
function JOINCELLS(range, delimiter) {
  if (typeof delimiter === 'undefined') {
    delimiter = ', ';
  }
  
  console.log('JOINCELLS called with range:', range, 'delimiter:', delimiter);
  
  if (!range || !Array.isArray(range)) {
    throw new Error("Invalid range provided");
  }
  
  var values = [];
  
  // Flatten the range and collect non-empty values
  for (var i = 0; i < range.length; i++) {
    if (Array.isArray(range[i])) {
      for (var j = 0; j < range[i].length; j++) {
        var value = range[i][j];
        if (value !== null && value !== undefined && String(value).trim() !== "") {
          values.push(String(value).trim());
        }
      }
    } else {
      var value = range[i];
      if (value !== null && value !== undefined && String(value).trim() !== "") {
        values.push(String(value).trim());
      }
    }
  }
  
  // Join with delimiter 
  var result = values.join(delimiter);
  console.log('JOINCELLS result: ' + result);
  return result;
}

// Make functions available globally for Excel's Custom Functions runtime
window.IRR = IRR;
window.JOINCELLS = JOINCELLS;

// Initialize when Office is ready
Office.onReady(function() {
  console.log('Prime Capital Custom Functions runtime ready');
  
  // Register functions with Excel's Custom Functions runtime
  if (typeof CustomFunctions !== 'undefined') {
    try {
      CustomFunctions.associate('PC.IRR', IRR);
      CustomFunctions.associate('PC.JOINCELLS', JOINCELLS);
      console.log('Custom Functions successfully registered: PC.IRR and PC.JOINCELLS');
    } catch (error) {
      console.error('Error registering custom functions:', error);
    }
  } else {
    console.warn('CustomFunctions API not available');
  }
});