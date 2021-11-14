# To run these tests, simply execute `nimble test`.

import unittest

import odsreader
import strformat
import std/tables
import strutils

suite "Load first sheet from ODS as seq":
  let ods_filename = "tests/test.ods"
  echo &"Loading first sheet from {ods_filename}"
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

suite "Load sheet by name":
  let ods_filename = "tests/test.ods"
  echo &"Loading sheet Sheet2 from {ods_filename}"
  let odsTable = loadOdsAsSeq(ods_filename, "Sheet2")

  test "dimension is 3x4":
    check len(odsTable) == 4
    check len(odsTable[0]) == 3

  test "repeated empty cells in row":
    check odsTable[0][1] == ""
    check odsTable[0][2] == ""
    check odsTable[2][0] == ""
    check odsTable[2][1] == ""

  test "repeated values in row":
    check odsTable[3][0] == "last2"
    check odsTable[3][1] == "last2"

  test "values in row":
    check odsTable[1][0] == ""
    check odsTable[1][1] == "second2"
    check odsTable[1][2] == ""

suite "Load sheet by name exception":
  let ods_filename = "tests/test.ods"

  test "exception when sheet name is not found":
    expect(SheetNotFoundException):
      echo &"Loading {ods_filename} with invalid sheet name: Sheet3"
      discard loadOdsAsSeq(ods_filename, "Sheet3")

suite "Load all sheets as table":
  let ods_filename = "tests/test.ods"
  echo &"Loading all sheets from {ods_filename} as table"
  let odsTable = loadOdsAsTable(ods_filename)

  test "all sheets loaded":
    doAssert len(odsTable) == 2
    check odsTable.hasKey("Sheet1")
    check odsTable.hasKey("Sheet2")

  test "sheet: Sheet1":
    check len(odsTable["Sheet1"]) == 4
    check len(odsTable["Sheet1"][0]) == 3
    check odsTable["Sheet1"][0][1] == ""
    check odsTable["Sheet1"][0][2] == ""
    check odsTable["Sheet1"][2][0] == ""
    check odsTable["Sheet1"][2][1] == ""
    check odsTable["Sheet1"][3][0] == "last"
    check odsTable["Sheet1"][3][1] == "last"
    check odsTable["Sheet1"][1][0] == ""
    check odsTable["Sheet1"][1][1] == "second"
    check odsTable["Sheet1"][1][2] == ""

  test "sheet: Sheet2":
    check len(odsTable["Sheet2"]) == 4
    check len(odsTable["Sheet2"][0]) == 3
    check odsTable["Sheet2"][0][1] == ""
    check odsTable["Sheet2"][0][2] == ""
    check odsTable["Sheet2"][2][0] == ""
    check odsTable["Sheet2"][2][1] == ""
    check odsTable["Sheet2"][3][0] == "last2"
    check odsTable["Sheet2"][3][1] == "last2"
    check odsTable["Sheet2"][1][0] == ""
    check odsTable["Sheet2"][1][1] == "second2"
    check odsTable["Sheet2"][1][2] == ""
    check odsTable["Sheet2"][3][2].startsWith("Element")
    check odsTable["Sheet2"][3][2].endsWith("menus.")


suite "Load ODS Document":
  let ods_filename = "tests/test.ods"
  echo &"Loading all sheets from {ods_filename} as Document"
  let doc = loadOds(ods_filename)

  test "Sheets in Document":
    check doc.getSheetNames() == @["Sheet1", "Sheet2"]

  test "Column names from first row":
    let colNamesTable = doc["Sheet1"].getColumnNames()
    check colNamesTable["first"] == 0
    check colNamesTable["Column1"] == 1

  test "Row column name subscript access":
    check doc["Sheet1"][0]["Column1"] == "second"

  test "Sheet rows iterator":
    for index, row in doc["Sheet1"]:
      case index:
        of 0:
          check row["Column1"] == "second"
        of 1:
          check row["Column2"] == "third"
        of 2:
          check row["first"] == "last"
        else:
          raiseAssert("Unexpected row!")

  test "Generated column names":
    let doc2 = loadOds(ods_filename, firstRowAreHeaders=false)
    let colNamesTable2 = doc2["Sheet1"].getColumnNames()
    check colNamesTable2["Column0"] == 0
    check colNamesTable2["Column1"] == 1

