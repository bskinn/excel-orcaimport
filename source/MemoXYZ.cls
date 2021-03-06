VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemoXYZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Name:        MemoXYZ (Class Module)
' Purpose:     Memoized retrieval of XYZ data from disk
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

Implements IMemoized

' Special string indicating error
Private Const ErrStr As String = "!!ERROR"

' Repository for memoized data
Private dict As New Scripting.Dictionary

' Internal filesystem object for faster access
Private fso As New Scripting.FileSystemObject

Public Function IMemoized_memoItem(ParamArray args() As Variant) As Variant
    ' Calls hash generator using the input parameters, checks
    '  if hash has been used already, makes the function call
    '  if not, and then returns either the stored or the newly
    '  calculated item.
    '
    ' On VBA RTEs, a StrObj is returned with the error information.
    ' Otherwise, an ORCA_XYZ is returned. If problems occurred during
    '  parsing, that information is output when data is requested from
    '  the object.
    '
    ' Arguments:
    '
    '       1) String path to the xyz file to be parsed
    '
    
    Dim workVt() As Variant, hashStr As String, SO As New StrObj
    
    ' Notify when called if debug flag set
    If ThisWorkbook.DEBUG_FLAG Then MsgBox "Called"
    
    ' Have to work with a copy of the ParamArray for some reason
    workVt = args
    
    ' Hash the input
    hashStr = hash(workVt)
    
    ' Check for leading error string
    If strStartsWith(hashStr, ErrStr) Then
        ' Error occurred; report it and return
        SO.StrValue = hashStr
        Set IMemoized_memoItem = SO
        Exit Function
    End If
    
    ' If not already in memo dict, retrieve and add
    If Not dict.Exists(hashStr) Then
        If ThisWorkbook.DEBUG_FLAG Then MsgBox "Loaded XYZ from disk"
        dict.Add hashStr, func(workVt)
    End If
    
    ' Retrieve dict copy of item for return
    Set IMemoized_memoItem = dict.item(hashStr)
    
End Function

Private Function func(argArray() As Variant) As Variant
    ' Performs the (presumably) costly item generation
    '  action for the memoized entity.
    ' In this case, it's just creating and binding an XYZ object
    
    Dim OX As New ORCA_XYZ
    
    ' Perform the import
    OX.readXYZ argArray(0), fso, False
    
    ' Return the resulting object
    Set func = OX
    
End Function

Private Function hash(argArray() As Variant) As String
    ' Hash generator for the arguments array passed in.
    ' Must ensure unique hashes.
    '
    ' Catches errors on attempt to bind the file and
    '  adjusts hash accordingly
    
    Dim fl As File, errNum As Long, errDesc As String
    
    ' Enable local error handling
    On Error Resume Next
    
    ' Try to bind the file
    Set fl = fso.GetFile(argArray(0))
    
    ' Store the error information
    errNum = Err.Number
    errDesc = Err.Description
    
    ' Reset error state
    Err.Clear
    On Error GoTo 0
    
    ' Generate hash string
    If errNum = 0 Then
        ' All is well. Hash string is the full file path joined to the last
        '  modified date/time by two underscores
        hash = fl.path & "__" & Format(fl.DateLastModified, "yyyyMMddhhmmss")
    Else
        ' Some error occurred.
        hash = ErrStr & " " & errNum & ": " & errDesc
    End If
    
End Function

Public Sub IMemoized_flushMemo()
    ' Flush the memoization dictionary
    dict.RemoveAll
End Sub


