VBScript Version
================

VBScript version consists of CSVUtils.vbs and CSVUtils_Test.vbs.
These script files are automatically converted from VBA version of v1.7 by using convert2vbs.vbs.
See README.md of v1.7 (README_v17.md) for the specification of the functions.

VBScript version is different from VBA version in the following points.
* `SetCSVUtilsAnyErrorIsFatal False` causes no effect. Any Error is always fatal.
* All the arugments of the functions are mandatory (not optional).
* All the arrays start with index 0. Please mind that the array returned by `ParseCSVToArray()` starts with 0.
* `CSVUtils_Test.vbs` excludes test cases that cause error. It tests only the successfull cases.
* VBScript version is much slower.
