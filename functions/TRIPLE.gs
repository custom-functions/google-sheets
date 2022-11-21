/**
 * Multiplies the input value by 3.
 *
 * @param {number} input The value or range of cells to multiply.
 * @return The input multiplied by 3.
 * @customfunction
 */
function TRIPLE(input) {
  return Array.isArray(input) ?
    input.map(row => row.map(cell => cell * 3)) :
    input * 3;
}
