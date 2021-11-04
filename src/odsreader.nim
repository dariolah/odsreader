## This module loads OpenDocument Spreadsheet

import zip/zipfiles
import streams
import strutils
import std/parsexml

proc loadOdsAsSeq*(filename: string): seq[seq[string]] =
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
    odsTable: seq[seq[string]] = newSeq[seq[string]]()
    odsRow: seq[string] = newSeq[string]()
    repeat: int = 1
    inCell, inTextP, inTable: bool = false

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
        odsTable.add(odsRow)
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
        # end of first sheet
        break
      else: discard
    of xmlAttribute:
      if inCell:
        if cmpIgnoreCase(x.attrKey, "table:number-columns-repeated") == 0:
          repeat = parseInt(x.attrValue)
    of xmlCharData:
      if inTextP:
        cellValue = x.charData
    of xmlEof: break # end of file reached
    else:
      discard # ignore other events

  x.close()
  return odsTable

if isMainModule:
  let
    filename = "/home/dlah/Documents/test.ods"
    odsTable = loadOdsAsSeq(filename)
  echo odsTable
