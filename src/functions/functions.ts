/* global Office, CustomFunctions */

/**
 * PC.IRR - Takes five cells as input and adds them together
 * @customfunction PC.IRR
 * @param cell1 First cell value
 * @param cell2 Second cell value  
 * @param cell3 Third cell value
 * @param cell4 Fourth cell value
 * @param cell5 Fifth cell value
 * @returns The sum of all five cell values
 */
function IRR(cell1: number, cell2: number, cell3: number, cell4: number, cell5: number): number {
  // Validate inputs are numbers
  const inputs = [cell1, cell2, cell3, cell4, cell5];
  
  for (let i = 0; i < inputs.length; i++) {
    if (typeof inputs[i] !== 'number' || isNaN(inputs[i])) {
      throw new Error(`Cell ${i + 1} must be a valid number`);
    }
  }
  
  // Add all five values together
  const result = cell1 + cell2 + cell3 + cell4 + cell5;
  console.log(`IRR calculation: ${cell1} + ${cell2} + ${cell3} + ${cell4} + ${cell5} = ${result}`);
  return result;
}

/**
 * PC.JOINCELLS - Joins cells from a range into a single string with specified delimiter
 * @customfunction PC.JOINCELLS
 * @param range The range of cells to join
 * @param delimiter The delimiter to use (default comma with space)
 * @returns The joined string
 */
function JOINCELLS(range: any[][], delimiter: string = ", "): string {
  console.log('JOINCELLS called with range:', range, 'delimiter:', delimiter);
  
  if (!range || !Array.isArray(range)) {
    throw new Error("Invalid range provided");
  }
  
  const values: string[] = [];
  
  // Flatten the range and collect non-empty values
  for (let i = 0; i < range.length; i++) {
    if (Array.isArray(range[i])) {
      for (let j = 0; j < range[i].length; j++) {
        const value = range[i][j];
        if (value !== null && value !== undefined && String(value).trim() !== "") {
          values.push(String(value).trim());
        }
      }
    } else {
      const value = range[i];
      if (value !== null && value !== undefined && String(value).trim() !== "") {
        values.push(String(value).trim());
      }
    }
  }
  
  // Join with delimiter 
  const result = values.join(delimiter);
  console.log('JOINCELLS result:', result);
  return result;
}

// Register the functions with Office
Office.onReady(() => {
  console.log('Prime Capital Custom Functions ready');
});