/* global Office, CustomFunctions */

/**
 * Calculates aggregate IRR by dividing expected future value by original beginning value
 * @customfunction
 * @param expectedFutureValue The expected future value
 * @param originalBeginningValue The original beginning value
 * @returns The aggregate IRR (future value / beginning value)
 */
function aggirr(expectedFutureValue: number, originalBeginningValue: number): number {
  if (originalBeginningValue === 0) {
    throw new Error("Division by zero: original beginning value cannot be zero");
  }
  return expectedFutureValue / originalBeginningValue;
}

/**
 * Joins cells from a range into a single string with specified delimiter
 * @customfunction
 * @param range The range of cells to join
 * @param delimiter The delimiter to use (default comma)
 * @returns The joined string
 */
function joincells(range: any[][], delimiter: string = ","): string {
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
          values.push(String(value));
        }
      }
    } else {
      const value = range[i];
      if (value !== null && value !== undefined && String(value).trim() !== "") {
        values.push(String(value));
      }
    }
  }
  
  return values.join(delimiter);
}

// Register functions in the global namespace for Excel to find them
(self as any).aggirr = aggirr;
(self as any).joincells = joincells;