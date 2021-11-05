## This module loads OpenDocument Spreadsheet

import zip/zipfiles
import streams
import strutils
import std/parsexml
import std/tables
import std/options
import strformat

type SheetNotFoundException* = object of CatchableError

proc loadOdsAsTable*(filename: string, sheetName: Option[string]=none(string), onlyFirstSheet: bool = false):
  OrderedTableRef[string, seq[seq[string]]] =
  ## Load ODS file as ordered table:
  ##    sheet name -> sequence (rows) of sequence (cols) of strings
  ## If *sheetName* is specified, load only sheet with that name or raise error
  ## If *onlyFirstSheet* is set to true, only first sheet wil be loaded. *sheetName* arg has precedence.
  ## Default: load all sheets

  var z: ZipArchive
  if not z.open(filename, fmRead):
    echo "loadOdsAsSeq: open failed"
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
    of xmlElementStart, xmlElementOpen:
      case x.elementName
      of "table:table-cell":
        inCell = true
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
          odsRow.add(cellValue)
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
        cellValue = x.charData
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

if isMainModule:
  let filename = "../tests/test.ods"
  let odsTable = loadOdsAsTable(filename)
  echo odsTable
  let odsTable2 = loadOdsAsSeq(filename, "Sheet2")
  echo odsTable2
  let odsTable3 = loadOdsAsSeq(filename)
  echo odsTable3
