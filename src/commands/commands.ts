/* global Office, CustomFunctions */

Office.onReady(() => {
  // The initialize function must be run each time a new page is loaded
  console.log('Office.js commands loaded');
});

/**
 * Calculates aggregate IRR by dividing expected future value by original beginning value
 * @param {number} expectedFutureValue - The expected future value
 * @param {number} originalBeginningValue - The original beginning value
 * @returns {number} The aggregate IRR (future value / beginning value)
 * @customfunction
 * @helpurl https://www.primeexcelence.com/functions/aggirr
 */
function AGGIRR(expectedFutureValue: number, originalBeginningValue: number): number {
  if (originalBeginningValue === 0) {
    throw new Error("Division by zero: original beginning value cannot be zero");
  }
  return expectedFutureValue / originalBeginningValue;
}

/**
 * Joins cells from a range into a single string with specified delimiter
 * @param {any[][]} range - The range of cells to join
 * @param {string} [delimiter=","] - The delimiter to use (default comma)
 * @returns {string} The joined string
 * @customfunction
 * @helpurl https://www.primeexcelence.com/functions/joincells
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

// For Excel custom functions, we need to register them globally
(window as any).AGGIRR = AGGIRR;
(window as any).JOINCELLS = JOINCELLS;