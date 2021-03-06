Attribute VB_Name = "ENGRAD"
'-------------------------------------------------------------------------------
' Name:        ENGRAD (Module)
' Purpose:     UDFs &c. providing access to memoized ORCA_ENGRAD properties
'
' Author:      Brian Skinn
'                bskinn@alum.mit.edu
'
' Created:     30 Apr 2016
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

Dim mENGRAD As New MemoENGRAD

Public Function ENGRADGradient(ByVal path As Variant) As Variant
    ' Retrieve memoized .engrad and return the array or
    '  string error value.
    
    ' Set volatile status per global flag
    Application.Volatile ThisWorkbook.VOLATILE_FLAG
    
    Dim retObj As Object
    
    ' Bind the object
    Set retObj = mENGRAD.IMemoized_memoItem(dearrayify(path))
    If TypeOf retObj Is ORCA_ENGRAD Then
        ENGRADGradient = retObj.gradientArray
    Else
        ENGRADGradient = retObj.StrValue
    End If
    
End Function

Public Function ENGRADEnergy(ByVal path As Variant) As Variant
    ' Retrieve memoized .engrad and return the energy value or
    '  string error value.
    
    ' Set volatile status per global flag
    Application.Volatile ThisWorkbook.VOLATILE_FLAG
    
    Dim retObj As Object
    
    ' Bind the object
    Set retObj = mENGRAD.IMemoized_memoItem(dearrayify(path))
    If TypeOf retObj Is ORCA_ENGRAD Then
        ENGRADEnergy = retObj.totalEnergy
    Else
        ENGRADEnergy = retObj.StrValue
    End If
    
End Function

Public Function ENGRADNumAtoms(ByVal path As Variant) As Variant
    ' Retrieve memoized .engrad and return the number of atoms or
    '  string error value.
    
    ' Set volatile status per global flag
    Application.Volatile ThisWorkbook.VOLATILE_FLAG
    
    Dim retObj As Object
    
    ' Bind the object
    Set retObj = mENGRAD.IMemoized_memoItem(dearrayify(path))
    If TypeOf retObj Is ORCA_ENGRAD Then
        ENGRADNumAtoms = retObj.numOfAtoms
    Else
        ENGRADNumAtoms = retObj.StrValue
    End If
    
End Function

Public Function ENGRADCoords(ByVal path As Variant) As Variant
    ' Retrieve memoized .engrad and return the 3Nx1 coordinates vector or
    '  string error value.
    
    ' Set volatile status per global flag
    Application.Volatile ThisWorkbook.VOLATILE_FLAG
    
    Dim retObj As Object
    
    ' Bind the object
    Set retObj = mENGRAD.IMemoized_memoItem(dearrayify(path))
    If TypeOf retObj Is ORCA_ENGRAD Then
        ENGRADCoords = retObj.atomCoordsArray
    Else
        ENGRADCoords = retObj.StrValue
    End If
    
End Function

Public Function ENGRADAtomicNums(ByVal path As Variant) As Variant
    ' Retrieve memoized .engrad and return the Nx1 atomic numbers vector or
    '  string error value.
    
    ' Set volatile status per global flag
    Application.Volatile ThisWorkbook.VOLATILE_FLAG
    
    Dim retObj As Object
    
    ' Bind the object
    Set retObj = mENGRAD.IMemoized_memoItem(dearrayify(path))
    If TypeOf retObj Is ORCA_ENGRAD Then
        ENGRADAtomicNums = retObj.atomicNumsArray
    Else
        ENGRADAtomicNums = retObj.StrValue
    End If

End Function

Public Function ENGRADAtomicSyms(ByVal path As Variant) As Variant
    ' Retrieve memoized .engrad and return the Nx1 atomic symbols vector or
    '  string error value.
    
    ' Set volatile status per global flag
    Application.Volatile ThisWorkbook.VOLATILE_FLAG
    
    Dim retObj As Object
    
    ' Bind the object
    Set retObj = mENGRAD.IMemoized_memoItem(dearrayify(path))
    If TypeOf retObj Is ORCA_ENGRAD Then
        ENGRADAtomicSyms = retObj.atomicSymsArray
    Else
        ENGRADAtomicSyms = retObj.StrValue
    End If

End Function

Function flushMemoENGRAD() As Variant
    ' Returns string describing results of flush attempt
    '  on the ENGRAD memoizer object
    
    ' Enable local error handling
    On Error Resume Next
    
    ' Try to flush the memo dict
    mENGRAD.IMemoized_flushMemo
    
    ' Check the error status and define the return string accordingly
    If Err.Number = 0 Then
        flushMemoENGRAD = "SUCCESS"
    Else
        flushMemoENGRAD = "FAILED: RTE #" & Err.Number & "(" & _
                            Err.Description & ")"
    End If
    
    ' Clear the error state and resume normal exception handling
    Err.Clear
    On Error GoTo 0
    
End Function
