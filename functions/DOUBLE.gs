/**
 *
 * @param {number} input The value or range of cells to multiply.
 * @return {number} The input multiplied by 2.
 * @customfunction
 */
function DOUBLE(input) {
  return Array.isArray(input) ?
    input.map(row => row.map(cell => cell * 2)) :
    input * 2;
}
