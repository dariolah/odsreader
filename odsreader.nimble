# Package

backend       = "c"
version       = "0.4.0"
author        = "Dario Lah"
description   = "OpenDocument Spreadhseet reader"
license       = "MIT"
srcDir        = "src"


# Dependencies

requires "nim >= 1.6.0"
requires "zippy >= 0.10.0"

task test, "Runs the test suite":
  exec "nim c -r tests/test_load_ods"
