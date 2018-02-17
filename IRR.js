/*
Copyright (C) 2018 N. Malkin

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.c
*/

var CODES_SEPARATOR = ',';

function filterEmpty(array) {
  return array.filter(function(value) {
    return value != '';
  });
}

function getCodeList(cell) {
  return filterEmpty(cell.split(CODES_SEPARATOR));
}

function CONCORDANCE(cellA, cellB, questionId) {
  var codeListA = getCodeList(cellA);
  var codeListB = getCodeList(cellB);

  // Let the variables a_i, and b_i, denote the numbers of attributes for the i-th unit chosen by raters A and B, respectively
  var a_i = codeListA.length;
  var b_i = codeListB.length;

  // Get the flags from the codebook, so we know which entries to ignore
  var codesAndFlags = getCodesAndFlags(questionId);

  // Let the random variable Xi denote the number of elements common to the sets A_i and B_i
  var x_i = 0;
  for (var i = 0; i < codeListA.length; i++) {
    var el = codeListA[i];
    if (codesAndFlags.codes.includes(el)) {
      if (codeListB.includes(el)) {
        x_i++;
      }
    } else if (!codesAndFlags.flags.includes(el)) {
      throw 'Not recognized as either code or flag: ' + el;
    }
  }

  // The observed proportion of concordance is pi_hat_i = x_i / max(a_i, b_i)
  var pi_hat_i = x_i / Math.max(a_i, b_i);
  return pi_hat_i;
}
