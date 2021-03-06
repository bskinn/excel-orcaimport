Attribute VB_Name = "Util_String"
'-------------------------------------------------------------------------------
' Name:        Util_String (Module)
' Purpose:     String helper functions for ORCA Importer
'
' Author:      Brian Skinn
'                bskinn@alum.mit.edu
'
' Created:     10 May 2016
' Copyright:   (c) Brian Skinn 2016
' License:     The MIT License; see "license.txt" for full license terms
'                   and contributor agreement.
'
'       This file is part of ORCA Importer, an Excel VBA add-in providing
'       memoized import functionality for output generated by the ORCA
'       quantum chemistry software program package
'       (http://orcaforum.cec.mpg.de).
'
'       http://www.github.com/bskinn/excel-orcaimport
'
'-------------------------------------------------------------------------------

Option Explicit
Option Base 1

Public Function strStartsWith(ByVal str As String, ByVal subStr As String) As Boolean
    strStartsWith = (subStr = Left(str, Len(subStr)))
End Function

Public Function strEndsWith(ByVal str As String, ByVal subStr As String) As Boolean
    strEndsWith = (subStr = Right(str, Len(subStr)))
End Function

