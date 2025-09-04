import { describe, it, expect, beforeEach } from 'vitest';
import { installAppScriptMock } from './mocks/appscript.js';

/** Example of testing code that touches SpreadsheetApp via a tiny adapter */
function writeMatrix(sheetName, startR, startC, matrix, facade) {
  const ss = facade?.getActive?.() || SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName) || ss._addSheet(sheetName);
  sh.getRange(startR, startC, matrix.length, matrix[0].length).setValues(matrix);
  return sh.getDataRange().getValues();
}

let env;
beforeEach(() => { env = installAppScriptMock(); });

describe('SpreadsheetApp mock adapter', () => {
  it('writes and reads matrix', () => {
    const data = [[1,2],[3,4]];
    const res = writeMatrix('Sheet1', 1, 1, data);
    expect(res.slice(0,2)).toEqual(data);
  });
});
