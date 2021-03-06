VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ORCA_XYZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Name:        ORCA_XYZ (Class Module)
' Purpose:     Handles imports for ORCA .xyz and .trj (OpenBabel) files
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

' Container for data imported from an ORCA XYZ or TRJ file
'
' Should actually work for any file in the standard XYZ format where
'  the number of atoms remains constant through the whole file
'
' Object not to be modified after creation; constructor routine to populate all internal
'  variables from HESS file at indicated path, and input parameters
'
' INTERNAL UNITS ARE IN BOHRS
'

Private NumAtoms As Long, NumGeoms As Long, Comments As Variant
Private AtomicNums As Variant, AtomCoords As Variant
Private wsf As WorksheetFunction
Private ErrorStr As String
Private Initialized As Boolean



'  ===  CONSTRUCTOR  ===
Public Sub readXYZ( _
            ByVal XYZ_Path As String, _
            ByRef fso As FileSystemObject, _
            Optional ByVal bohrs As Boolean = False _
            )
    
    Dim rx As New RegExp, ts As TextStream
    Dim str As String, str2 As String
    Dim iter As Long, iter2 As Long, iter3 As Long
    Dim mch As Match, mchs As MatchCollection, mchs2 As MatchCollection
    Dim val As Long, val2 As Long
    Dim valStr As String
    
    ' Only proceed if object hasn't already been configured
    If Initialized Then
        Err.Raise Number:=425, Description:="ORCA_ENGRAD object already initialized"
    End If
    
    ' Bind the filesystem
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open the XYZ file for reading
    If fso.FileExists(XYZ_Path) Then
        Set ts = fso.OpenTextFile(XYZ_Path, ForReading)
    Else
        ErrorStr = "Specified file not found"
        Exit Sub
    End If
    
    ' Pull entire file contents
    str = ts.ReadAll
    
    With rx
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        
        ' == GEOMETRIES ==
        ' This does not check for badly malformed OpenBabel XYZ files, where content exists that is
        '  improperly structured to the point that the geometry block RegExp does not retrieve it.
        '  It will find all properly formatted geometry blocks, and will return an error: (1) if the wrong
        '  number of coordinates is provided within a properly formatted block; (2) if not all of the
        '  properly formatted blocks have the same number of atoms; (3) if no properly formatted blocks
        '  are found; (4) if atoms collide in a given geometry; (5)...
        .Pattern = "([0-9]+)\s*\n(.*)\n((\s+([a-zA-Z]+|[0-9]+)(\s+[0-9.-]+){3}\s*\n)+)"
        
        If Not .Test(str) Then
            ' No geometries found
            ErrorStr = "XYZ data not found in indicated file"
            Exit Sub
        End If
        
        ' Store the matches
        Set mchs = .Execute(str)
        
        ' Store relevant values and dimension the various arrays
        NumAtoms = CLng(mchs(0).SubMatches(0))
        NumGeoms = mchs.Count
        ReDim Comments(1 To NumGeoms, 1 To 1)
        Comments(1, 1) = mchs(0).SubMatches(1)
        ReDim AtomicNums(1 To NumAtoms, 1 To 1)
        ReDim AtomCoords(1 To NumAtoms * 3, 1 To NumGeoms) ' column vectors of atom positions
        
        ' Loop to check that the number of atoms is the same as the first geom, and to store
        '  the comments
        If NumGeoms > 1 Then
            For iter = 2 To NumGeoms
                Comments(iter, 1) = mchs(iter - 1).SubMatches(1)
                If Not CLng(mchs(iter - 1).SubMatches(0)) = NumAtoms Then
                    ErrorStr = "Mismatch in 'number of atoms' specification at geometry #" & iter
                    Exit Sub
                End If
            Next iter
        End If
        
        ' Reprocess the coordinates of each match to retrieve the atom and coordinate information
        '  The pass on the first geometry (defining values) is different than the rest (proofing values)
        .Pattern = "\s+([a-zA-Z]+|[0-9]+)\s+([0-9.-]+)\s+([0-9.-]+)\s+([0-9.-]+)\s*\n"
        
        ' Loop and retrieve geometries
        For iter = 1 To mchs.Count
            ' Check to ensure coordinate data actually present
            If Not .Test(mchs(iter - 1).SubMatches(2)) Then
                ' This should never occur, as the pattern is chosen to succeed on any of mchs
                ErrorStr = "Coordinate data not found when parsing geometry block #" & iter
                Exit Sub
            End If
            
            ' Pull atom symbols and coordinates as Matches
            Set mchs2 = .Execute(mchs(iter - 1).SubMatches(2))
            
            ' Confirm number of atoms actually found is correct
            If Not NumAtoms = mchs2.Count Then
                ErrorStr = "Incorrect number of atoms found when parsing geometry block #" & iter
                Exit Sub
            End If
            
            For iter2 = 1 To mchs2.Count  ' Loop over each atom
                ' For the first geometry, proof and store the atomic number/symbol.
                '  For all later geometries, confirm consistency
                ' Store the submatch string
                valStr = mchs2.item(iter2 - 1).SubMatches(0)
                
                ' If it's a number, convert to an atomic symbol
                If IsNumeric(valStr) Then valStr = atomSym(valStr)
                
                If iter = 1 Then
                    ' Check validity in first block (element is simulable) and store
                    If atomNum(valStr) = Invalid_AtomNum_Arg Then
                        ErrorStr = "Invalid atom """ & valStr & """ at position " & iter
                        Exit Sub
                    Else
                        AtomicNums(iter2, 1) = CLng(atomNum(valStr))
                    End If
                Else
                    ' Check consistency with value stored from first block
                    If Not CLng(atomNum(valStr)) = AtomicNums(iter2, 1) Then
                        ErrorStr = "Atomic symbol mismatch in block " & iter & ", atom " & iter2
                        Exit Sub
                    End If
                End If
                
                ' Loop and store the coordinates, CONVERTING TO BOHRS IF NEEDED!
                For iter3 = 1 To 3
                    If bohrs Then
                        AtomCoords(3 * (iter2 - 1) + iter3, iter) = _
                                CDbl(mchs2(iter2 - 1).SubMatches(iter3))
                    Else
                        AtomCoords(3 * (iter2 - 1) + iter3, iter) = _
                                CDbl(mchs2(iter2 - 1).SubMatches(iter3)) / Angs_per_Bohr
                    End If
                Next iter3
            Next iter2
        Next iter
    End With
    
    ' Set the initialized flag, and done.
    Initialized = True
    
End Sub


'  ===  PROPERTY GET  ===
Property Get numOfAtoms() As Variant
    If ErrorStr <> "" Then
        numOfAtoms = ErrorStr
    Else
        If Initialized Then
            numOfAtoms = NumAtoms
        Else
            RaiseUninitError
        End If
    End If
End Property

Property Get numOfGeoms() As Variant
    If ErrorStr <> "" Then
        numOfGeoms = ErrorStr
    Else
        If Initialized Then
            numOfGeoms = NumGeoms
        Else
            RaiseUninitError
        End If
    End If
End Property

Property Get isInitialized() As Boolean
    isInitialized = Initialized
End Property

Property Get commentsArray() As Variant
    If ErrorStr <> "" Then
        commentsArray = ErrorStr
    Else
        If Initialized Then
            commentsArray = Comments
        Else
            RaiseUninitError
        End If
    End If
End Property

Property Get atomicNumsArray() As Variant
    If ErrorStr <> "" Then
        atomicNumsArray = ErrorStr
    Else
        If Initialized Then
            atomicNumsArray = AtomicNums
        Else
            RaiseUninitError
        End If
    End If
End Property

Property Get atomicSymsArray() As Variant
    
    Dim iter As Long, workVnt As Variant
    
    If ErrorStr <> "" Then
        atomicSymsArray = ErrorStr
    Else
        If Initialized Then
            ReDim workVnt(1 To NumAtoms, 1 To 1)
            For iter = 1 To NumAtoms
                workVnt(iter, 1) = atomSym(CLng(AtomicNums(iter, 1)))
            Next iter
            atomicSymsArray = workVnt
        Else
            RaiseUninitError
        End If
    End If
    
End Property

Property Get atomCoordsArray() As Variant
    If ErrorStr <> "" Then
        atomCoordsArray = ErrorStr
    Else
        If Initialized Then
            atomCoordsArray = AtomCoords
        Else
            RaiseUninitError
        End If
    End If
End Property

Property Get singleGeometry( _
            geomNum As Long _
            ) As Variant
    
    Dim iter As Long, workVnt As Variant
    
    ' Return string if error occurred during parsing
    If ErrorStr <> "" Then
        singleGeometry = ErrorStr
    Else
        If Initialized Then
            ' Size working array
            ReDim workVnt(1 To NumAtoms * 3, 1 To 1)
            
            ' Iterate and fill working array
            For iter = 1 To 3 * NumAtoms
                ' If a valid geomNum, fill the geometry
                If geomNum >= 1 And geomNum <= NumGeoms Then
                    workVnt(iter, 1) = CDbl(AtomCoords(iter, geomNum))
                Else
                    singleGeometry = "Invalid geometry number"
                    Exit Property
                End If
            Next iter
            
            ' Store the return value
            singleGeometry = workVnt
            
        Else
            RaiseUninitError
        End If
    End If
    
End Property

Property Get atomDistance( _
            ByVal geomNum As Long, _
            ByVal atomIdx1 As Long, _
            ByVal atomIdx2 As Long _
            ) As Variant
    
    ' Distance in Bohrs between the two indicated atoms
    
    Dim workVec As Variant, workDbl As Double, iter As Long
    
    ' If a prior error, return it
    If ErrorStr <> "" Then
        atomDistance = ErrorStr
        Exit Property
    End If
    
    ' If not initialized, dump with error
    If Not Initialized Then
        RaiseUninitError
    End If
    
    ' Retrieve the relevant geometry; passing out any error behavior
    workVec = Me.singleGeometry(geomNum)
    If Not IsArray(workVec) Then
        atomDistance = workVec
        Exit Property
    End If
    
    ' Reshape to a more convenient form
    workVec = mReshape(workVec, Me.numOfAtoms, 3)
    
    ' Confirm provided atom numbers are valid
    If atomIdx1 < 1 Or atomIdx1 > Me.numOfAtoms Then
        ' Not within range of atoms present for atom 1
        atomDistance = "Invalid atom index #1: '" & atomIdx1 & "'"
        Exit Property
    ElseIf atomIdx2 < 1 Or atomIdx2 > Me.numOfAtoms Then
        atomDistance = "Invalid atom index #2: '" & atomIdx2 & "'"
        Exit Property
    Else
        ' Both are within range -- calculate the distance
        workDbl = 0#
        For iter = 1 To 3
            workDbl = workDbl + ( _
                    workVec(atomIdx1, iter) - workVec(atomIdx2, iter) _
                                ) ^ 2
        Next iter
        atomDistance = Sqr(workDbl)
    End If
    
End Property

Property Get atomDisplacement( _
            ByVal geomNum As Long, _
            ByVal atomIdxFrom As Long, _
            ByVal atomIdxTo As Long _
            ) As Variant
    
    ' Vector displacement between the two indicated atoms, from the
    '  indicated geometry, as a 3x1 2-D array. Bohrs units.
    
    Dim workVec As Variant, iter As Long
    Dim outVt As Variant
    
    ' If a prior error, return it
    If ErrorStr <> "" Then
        atomDisplacement = ErrorStr
        Exit Property
    End If
    
    ' If not initialized, dump with error
    If Not Initialized Then
        RaiseUninitError
    End If
    
    ' Retrieve the relevant geometry; passing out any error behavior
    workVec = Me.singleGeometry(geomNum)
    If Not IsArray(workVec) Then
        atomDisplacement = workVec
        Exit Property
    End If
    
    ' Reshape to a more convenient form
    workVec = mReshape(workVec, Me.numOfAtoms, 3)
    
    ' Confirm provided atom numbers are valid
    If atomIdxFrom < 1 Or atomIdxFrom > Me.numOfAtoms Then
        ' Not within range of atoms present for 'from' atom
        atomDisplacement = "Invalid 'from' atom index: '" & atomIdxFrom & "'"
        Exit Property
    ElseIf atomIdxTo < 1 Or atomIdxTo > Me.numOfAtoms Then
        atomDisplacement = "Invalid 'to' atom index: '" & atomIdxTo & "'"
        Exit Property
    Else
        ' Both are within range -- calculate the displacement
        ReDim outVt(1 To 3, 1 To 1)
        For iter = 1 To 3
            outVt(iter, 1) = _
                    workVec(atomIdxTo, iter) - workVec(atomIdxFrom, iter)
        Next iter
        atomDisplacement = outVt
    End If
    
End Property

'  ===  AUX METHODS  ===
Private Sub RaiseUninitError()
    Err.Raise Number:=425, Description:="Cannot use improperly " & _
                                    "initialized ORCA_XYZ object"
End Sub

'  ===  CLASS EVENTS  ===
Private Sub Class_Initialize()
    Initialized = False
    ErrorStr = ""
    
    Set wsf = Application.WorksheetFunction
End Sub
