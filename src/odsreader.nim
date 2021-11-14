## This module loads OpenDocument Spreadsheet

import zip/zipfiles
import streams
import strutils
import std/parsexml
import std/tables
import std/options
import strformat

type
  SheetNotFoundException* = object of CatchableError
  InvalidColumnNameException* = object of CatchableError

proc loadOdsAsTable*(filename: string, sheetName: Option[string]=none(string), onlyFirstSheet: bool = false):
  OrderedTableRef[string, seq[seq[string]]] =
  ## Load ODS file as ordered table:
  ##    sheet name -> sequence (rows) of sequence (cols) of strings
  ## If *sheetName* is specified, load only sheet with that name or raise error
  ## If *onlyFirstSheet* is set to true, only first sheet wil be loaded. *sheetName* arg has precedence.
  ## Default: load all sheets

  var z: ZipArchive
  if not z.open(filename, fmRead):
    echo &"loadOdsAsSeq: failed to open {filename}"
    quit(1)

  let outStream = newStringStream("")
  z.extractFile("content.xml", outStream)
  outStream.setPosition(0)

  var x: XmlParser
  open(x, outStream, filename)

  var
    cellValue: string = ""
    odsSheet: seq[seq[string]] = newSeq[seq[string]]()
    odsRow: seq[string] = newSeq[string]()
    repeat: int = 1
    inCell, inTextP, inTable: bool = false
    currentSheetName: string = ""
    odsTable = newOrderedTable[string, seq[seq[string]]]()

  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      case x.elementName
      of "table:table-cell":
        inCell = true
      else: discard
    of xmlElementStart:
      case x.elementName
      of "text:p":
        inTextP = true
      else: discard
    of xmlElementEnd:
      case x.elementName
      of "table:table-row":
        odsSheet.add(odsRow)
        odsRow = newSeq[string]()
      of "table:table-cell":
        for i in countup(1, repeat):
          odsRow.add(cellValue.strip())
        repeat = 1
        cellValue = ""
        inCell = false
        inTextP = false
        inTable = false
      of "table:table":
        # end of sheet
        odsTable[currentSheetName] = odsSheet
        odsSheet = newSeq[seq[string]]()
        if onlyFirstSheet:
          # load only first sheet
            break
      of "text:p":
        if inTextP == true:
          cellValue = cellValue & "\n"
      else: discard
    of xmlAttribute:
      if inCell:
        if cmpIgnoreCase(x.attrKey, "table:number-columns-repeated") == 0:
          repeat = parseInt(x.attrValue)
      else:
        if x.attrKey == "table:name":
          if sheetName.isSome:
            # loading specified sheet only
            if sheetName.get() == x.attrValue:
              currentSheetName = x.attrValue
            else:
              # not the one we are looking for, skip to the end of sheet
              while true:
                x.next()
                if x.kind == xmlElementEnd and x.elementName == "table:table": break
          elif onlyFirstSheet:
            # load only first sheet
              currentSheetName = x.attrValue
          else:
            # loading all sheets
            currentSheetName = x.attrValue
    of xmlCharData:
      if inTextP:
        cellValue = cellValue & x.charData
    of xmlEof: break # end of file reached
    else:
      discard # ignore other events

  x.close()

  if sheetName.isSome and not odsTable.hasKey(sheetName.get()):
    raise SheetNotFoundException.newException(&"Sheet named >{sheetName.get()}< not found!")
  return odsTable

proc loadOdsAsSeq*(filename: string, sheetName: string): seq[seq[string]] =
  ## Load sheet name *sheetName* from ODS file as:
  ##    sequence (rows) of sequence (cols) of strings
  ## Raise SheetNotFoundException if sheet not found
  let odsTable = loadOdsAsTable(filename, some[string](sheetName))
  for value in odsTable.values:
    # return first value from ordered table
    return value

proc loadOdsAsSeq*(filename: string): seq[seq[string]] =
  ## Load first sheet from ODS file as:
  ##    sequence (rows) of sequence (cols) of strings
  let odsTable = loadOdsAsTable(filename, none[string](), true)
  for value in odsTable.values:
    # return first value from ordered table
    return value

type
  Row* = ref object
    data: seq[string]
    columnNames: OrderedTableRef[string, int]

  Sheet* = ref object
    data: seq[seq[string]]
    columnNames: OrderedTableRef[string, int]

  Document* = ref object
    data: OrderedTableRef[string, seq[seq[string]]]
    sheet: OrderedTableRef[string, Sheet]

proc `[]`*(doc: Document, sheetName: string): Sheet =
  return doc.sheet[sheetName]

proc getSheetNames*(doc: Document): seq[string] =
  var names: seq[string] = @[]
  for name in doc.sheet.keys:
    names.add(name)
  return names

proc `$`*(sheet: Sheet): string =
  let outStream = newStringStream("")
  for col in sheet.columnNames.keys:
    outStream.write(col)
    outStream.write(" | ")
  outStream.write("\n")

  for row in sheet.data:
    outStream.write($row)
    outStream.write("\n")

  return outStream.data

proc `$`*(row: Row): string =
  let outStream = newStringStream("")

  for item in row.data:
    outStream.write(item)
    outStream.write(" | ")

  return outStream.data

proc getColumnNames*(sheet: Sheet): OrderedTableRef[string, int] =
  return sheet.columnNames

proc `[]`*(sheet: Sheet, rowIndex: int): Row =
  let row = Row(data: sheet.data[rowIndex], columnNames: sheet.columnNames)
  return row

proc `[]`*(row: Row, colName: string): string =
  if not row.columnNames.hasKey(colName):
    raise InvalidColumnNameException.newException(&"Column name >{colName}< not found!")

  return row.data[row.columnNames[colName]]

iterator items*(sheet: Sheet): Row =
  for index in countup(0, len(sheet.data) - 1):
    let row = Row(data: sheet.data[index], columnNames: sheet.columnNames)
    yield row

iterator pairs*(sheet: Sheet): tuple[index:int, row:Row] =
  for index in countup(0, len(sheet.data) - 1):
    let row = Row(data: sheet.data[index], columnNames: sheet.columnNames)
    yield (index, row)

proc loadOds*(filename: string, firstRowAreHeaders=true): Document =
  let ods = loadOdsAsTable(filename)
  var sheets = newOrderedTable[string, Sheet]()

  for sheetName in ods.keys:
    var columnNames = newOrderedTable[string, int]()
    if firstRowAreHeaders:
      var index = 0
      for columnValue in ods[sheetName][0]:
        var columnName = columnValue
        if columnName == "":
          # missing column name
          columnName = &"Column{index}"
        elif columnName in columnNames:
          # duplicate
          columnName = &"{columnValue}_{index}"

        columnNames[columnName] = index
        index += 1
      ods[sheetName].delete(0)
    else:
      # firstRowAreHeaders=false
      var columnName = ""
      for colNumber in countup(0, len(ods[sheetName][0]) - 1):
        columnName = &"Column{colNumber}"
        columnNames[columnName] = colNumber
    var sheet = Sheet(data: ods[sheetName], columnNames: columnNames)
    sheets[sheetName] = sheet

  return Document(data: ods, sheet: sheets)

if isMainModule:
  let filename = "../tests/test.ods"
  let doc = loadOds(filename)
  echo doc.getSheetNames()
  echo doc["Sheet2"].getColumnNames()
  echo doc["Sheet2"]
  #for row in doc["Sheet1"]:
  #  echo row["first"]
