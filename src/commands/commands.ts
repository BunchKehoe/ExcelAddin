/* global Office, CustomFunctions */

// Interface for CustomFunctions
interface CustomFunctionsType {
  associate(id: string, fn: Function): void;
}

// Declare CustomFunctions as a global variable
declare const CustomFunctions: CustomFunctionsType;

/**
 * Adds two numbers together.
 * @param {number} number1 First number to add
 * @param {number} number2 Second number to add
 * @returns {number} The sum of the two numbers
 * @customfunction PC.EXAMPLEFUNCTION
 */
function exampleFunction(number1: number, number2: number): number {
  return number1 + number2;
}

// Register the custom function
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate('EXAMPLEFUNCTION', exampleFunction);
}

// Initialize when Office is ready
Office.onReady(() => {
  console.log('Prime Capital Commands ready');
});