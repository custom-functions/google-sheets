// @author Nwachukwu Ujubuonu https://github.com/Hemephelus/google-sheets-custom-function
/**
 * Returns the values of the Roman Text as an Integer.
 * 
 * @param {A2:A9 or "XXIII"} range The roman text to convert.
 *
 * @return the values of the Roman Text as an Integer.
 * @customfunction
 */  
function ROMAN_TO_NUMBER(range) {

  let array = []

  if(!Array.isArray(range))return ROM_TO_NUM_(range)


  for (let i = 0; i < range.length; i++) {

    array.push([ROM_TO_NUM_(range[i][0])] )
  
  }
  return array
};

function ROM_TO_NUM_(roman_text) {

    let singleLetters = {
    "I": 1,
    "V": 5,
    "X": 10,
    "L": 50,
    "C": 100,
    "D": 500,
    "M": 1000,
  }


  let total = 0
  let val = 0

  for (let i = 0; i < roman_text.length; i++) {

    if (singleLetters[roman_text[i]] < singleLetters[roman_text[i + 1]]) {

      val = singleLetters[roman_text[i + 1]] - singleLetters[roman_text[i]]
      total += val
      i++

    } else {

      total += singleLetters[roman_text[i]]

    }
  
  }
  return total
};