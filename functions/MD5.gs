/**
 * Generates an MD5 hash of the input.
 *
 * @author KEINOS
 *
 * @param {string} input The value to be hashed.
 * @return {string} The MD5 hash of the input.
 * @customfunction
 */
function MD5(input) {
  return Array.isArray(input) ?
    input.map(row => row.map(cell => helper_(cell))) :
    helper_(input);
}

const helper_ = (input) => {
  let output = new String();
  Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, input)
    .forEach(val => {
      val < 0 ? val += 256 : null;
      val.toString(16).length == 1 ? output += '0' : null;
      output += val.toString(16);
    });
  return output;
}
