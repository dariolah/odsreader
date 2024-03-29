# odsreader - library for reading OpenDocument Spreadsheet (.ods) files


## QuickStart

demo.nim
```Nim
import odsreader
import strformat
import std/tables

proc mainProc =
  let ods_filename = "tests/test.ods"
  echo &"Loading first sheet from {ods_filename}"
  let doc = loadOds(filename)

  echo doc.getSheetNames()
  echo doc["Sheet1"].getColumnNames()

  # Sheet iterator
  for row in doc.["Sheet1"]:
    # row column access
    echo row["first"]

if isMainModule:
  mainProc()
```


Or just use nested sequences and tables...

demoseq.nim
```Nim
import odsreader
import strformat
import std/tables

proc mainProc =
  let ods_filename = "tests/test.ods"
  echo &"Loading first sheet from {ods_filename}"
  let odsFirstSheet = loadOdsAsSeq(ods_filename)
  echo odsFirstSheet

  # load sheet by name
  echo &"Loading sheet Sheet2 from {ods_filename}"
  let odsSheet2 = loadOdsAsSeq(ods_filename, "Sheet2")
  echo odsSheet2

  # load all sheets as table
  echo &"Loading all sheets from {ods_filename} as table"
  let odsTable = loadOdsAsTable(ods_filename)
  echo odsTable

if isMainModule:
  mainProc()
```
