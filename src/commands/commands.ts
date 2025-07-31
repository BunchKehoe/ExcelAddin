/* global Office, CustomFunctions */

/**
 * Calculates aggregate IRR by dividing expected future value by original beginning value
 * @customfunction AGGIRR
 * @param expectedFutureValue The expected future value
 * @param originalBeginningValue The original beginning value
 * @returns The aggregate IRR (future value / beginning value)
 */
function AGGIRR(expectedFutureValue: number, originalBeginningValue: number): number {
  if (originalBeginningValue === 0) {
    throw new Error("Division by zero: original beginning value cannot be zero");
  }
  return expectedFutureValue / originalBeginningValue;
}

/**
 * Joins cells from a range into a single string with specified delimiter
 * @customfunction JOINCELLS
 * @param range The range of cells to join
 * @param delimiter The delimiter to use (default comma)
 * @returns The joined string
 */
function JOINCELLS(range: any[][], delimiter: string = ","): string {
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

// Make functions available globally for Excel
(globalThis as any).AGGIRR = AGGIRR;
(globalThis as any).JOINCELLS = JOINCELLS;