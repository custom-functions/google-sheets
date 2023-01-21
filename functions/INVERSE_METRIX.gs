// @author Nwachukwu Ujubuonu https://github.com/Hemephelus/google-sheets-custom-function
/**
 * Returns the inverse of a square metrix.
 * 
 * @param {C1:D2} range The square metrix to invert.
 *
 * @return the inverse of a square metrix.
 * @customfunction
 */  
function INVERSE_METRIX(range) {

  if(range[i].length !== range.length)return 'this is not a square metrix'

  var temp,
    N = range.length,
    E = [];

  for (var i = 0; i < N; i++)

    E[i] = [];

  for (i = 0; i < N; i++)

    for (var j = 0; j < N; j++) {

      E[i][j] = 0;
      if (i == j)
        E[i][j] = 1;

    }

  for (var k = 0; k < N; k++) {

    temp = range[k][k];

    for (var j = 0; j < N; j++) {

      range[k][j] /= temp;
      E[k][j] /= temp;
    }

    for (var i = k + 1; i < N; i++) {

      temp = range[i][k];

      for (var j = 0; j < N; j++) {

        range[i][j] -= range[k][j] * temp;
        E[i][j] -= E[k][j] * temp;

      }

    }

  }

  for (var k = N - 1; k > 0; k--) {

    for (var i = k - 1; i >= 0; i--) {

      temp = range[i][k];

      for (var j = 0; j < N; j++) {

        range[i][j] -= range[k][j] * temp;
        E[i][j] -= E[k][j] * temp;

      }

    }

  }

  for (var i = 0; i < N; i++)

    for (var j = 0; j < N; j++)

      range[i][j] = E[i][j];
  return range;

}
