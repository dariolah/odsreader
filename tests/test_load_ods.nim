# To run these tests, simply execute `nimble test`.

import unittest

import odsreader
import strformat

suite "Load first sheet from ODS as seq":
  let ods_filename = "tests/test.ods"
  echo &"Loading {ods_filename}"
  let odsTable = loadOdsAsSeq(ods_filename)

  test "dimension is 3x4":
    check len(odsTable) == 4
    check len(odsTable[0]) == 3

  test "repeated empty cells in row":
    check odsTable[0][1] == ""
    check odsTable[0][2] == ""
    check odsTable[2][0] == ""
    check odsTable[2][1] == ""

  test "repeated values in row":
    check odsTable[3][0] == "last"
    check odsTable[3][1] == "last"

  test "values in row":
    check odsTable[1][0] == ""
    check odsTable[1][1] == "second"
    check odsTable[1][2] == ""
