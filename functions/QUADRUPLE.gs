/**
 * Multiplies the input value by 4.
 *
 * @param {number|Array<Array<number>>} input The value or range of cells to multiply.
 * @return The input multiplied by 4.
 * @customfunction
 */
function QUADRUPLE(input) {
  return Array.isArray(input) ?
    input.map(row => row.map(cell => cell * 4)) :
    input * 4;
}
