VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ORCA_HESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Name:        ORCA_HESS (Class Module)
' Purpose:     Handles imports for ORCA .hess files
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

' Container for data imported from an ORCA HESS file
'
' Includes variables for the displaced mode, direction, and magnitude using
'  common Enum types
'
' Object not to be modified after creation; constructor routine to populate all internal
'  variables from HESS file at indicated path, and input parameters
'
' THE DISPLACEMENT MAGNITUDE IS ALWAYS POSITIVE, if DispDir is not dispNone. The
'  DispDirection Enum is defined such that the signed displacement can be obtained
'  by: DispDir * DispMag

' Want to pull Hessian, frequencies, normal modes, atoms/masses, geometry from .HESS
'  If present, also retrieve the eigenvalues/-vectors?  Mainly only useful as a cross-check,
'  probably?


Private NumAtoms As Long, NumFreqs As Long
Private Hessian As Variant, Freqs As Variant, Modes As Variant
Private AtomicNums As Variant, AtomicWeights As Variant, AtomCoords As Variant
Private Initialized As Boolean



'  ===  CONSTRUCTOR  ===
Public Function readHESS(ByVal HESS_Path As String, ByRef fso As FileSystemObject) As String
    ' Null string return means file indicated by pathStr exists and conforms to expected pattern,
    '  that the data was retrieved successfully, and that the other entered parameters
    '  were sensible.  (Hessian object cannot internally proofread things like whether the
    '  indicated mode number exists for the problem at hand.)
    ' Non-null string contains description of error encountered.
    
    Dim rx As New RegExp, ts As TextStream
    Dim str As String, str2 As String
    Dim iter As Long, mch As Match, mchs As MatchCollection, val As Long
    Dim iter2 As Long, mchs2 As MatchCollection, val2 As Long
    Dim iter3 As Long, valStr As String
    Dim rowMatches As Long
    
    ' Initialize return value
    readHESS = "Constructor exited without specifying return value"
    
    ' Only proceed if object hasn't already been configured
    If Initialized Then
        readHESS = "Object has already been initialized"
        GoTo Error_Exit
    End If
    
    ' Bind the filesystem
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Open the HESS file for reading
    If fso.FileExists(HESS_Path) Then
        Set ts = fso.OpenTextFile(HESS_Path, ForReading)
    Else
        readHESS = "Specified file not found"
        GoTo Error_Exit
    End If
    
    ' Pull entire file contents
    str = ts.ReadAll
    
    With rx
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        
        ' == NUMBER OF ATOMS, IDENTITIES, MASSES, COORDINATES ==
        ' Might as well retrieve all of these together, since the "$act_atom" field
        '  is unreliable in analytical Hessians
        .Pattern = "#.*\n\$atoms.*\n([0-9]+).*(\n([^a-zA-Z$#]+[a-zA-Z]+" & _
                "[^0-9\n]+[0-9.]+([^0-9-]+[0-9.-]+){3}.*\n)+)"
        If Not .Test(str) Then
            ' Block not found
            readHESS = "Atom coordinate and mass block could not be found in the " & _
                    "indicated file"
            GoTo Error_Exit
        End If
        
        ' Store the substring for subsequent parsing and the number of atoms
        Set mch = .Execute(str).item(0)
        str2 = mch.SubMatches(1)
        NumAtoms = CLng(mch.SubMatches(0))
        
        ' Reset pattern to pull each line of atom data and execute
        .Pattern = "\n[^a-zA-Z0-9$]+([a-zA-Z]+)[^0-9]+([0-9.]+)[^0-9-]+([0-9.-]+)[^0-9-]+([0-9.-]+)[^0-9-]+([0-9.-]+)"
        Set mchs = .Execute(str2)
        
        ' Consistency check
        If Not mchs.Count = NumAtoms Then
            readHESS = "Number of atoms found in coordinates block (" & mchs.Count & _
                    ") does not match count indicated at start of block (" & NumAtoms & ")"
            GoTo Error_Exit
        End If
        
        ' Resize the arrays
        ReDim AtomicNums(1 To mchs.Count, 1 To 1)
        ReDim AtomicWeights(1 To mchs.Count, 1 To 1)
        ReDim AtomCoords(1 To (3 * mchs.Count), 1 To 1)
        
        ' Loop and store values
        For iter = 1 To mchs.Count
            Set mch = mchs(iter - 1)
            If atomNum(mch.SubMatches(0)) = Invalid_AtomNum_Arg Then
                readHESS = "Invalid atomic symbol """ & mch.SubMatches(0) & """"
                GoTo Error_Exit
            Else
                AtomicNums(iter, 1) = CLng(atomNum(mch.SubMatches(0)))
            End If
            AtomicWeights(iter, 1) = CDbl(mch.SubMatches(1))
            For iter2 = 0 To 2
                AtomCoords(3 * iter + iter2 - 2, 1) = CDbl(mch.SubMatches(iter2 + 2))
            Next iter2
        Next iter
        
        
        ' == HESSIAN ==
        ' Set the search pattern for the entire Hessian block
        .Pattern = "\$hessian.*\n([0-9]+).*\n(([^0-9$\n-]+[0-9.-]+)+.*\n)+"
        If .Test(str) Then
            ' Pull the block, consistency check of the dimensionality annotation
            str2 = .Execute(str).item(0)
            val = .Execute(str).item(0).SubMatches(0)
            If Not val = NumAtoms * 3 Then
                readHESS = "Hessian block dimensionality specification (" & val & ") does not " & _
                        "match 3*number of atoms (" & 3 * NumAtoms & ")"
                GoTo Error_Exit
            End If
            
            ' Identify the subsections of the Hessian; set the pattern
            .Pattern = "(\s+[0-9]+)+\s*\n(\s+[0-9]+(\s*[-]?[0-9]+\.[0-9]{6})+\s*\n)+"
            
            If .Test(str2) Then
                ' If subsection parsing worked, store the Matches and redim the Hessian
                Set mchs = .Execute(str2)
                ReDim Hessian(1 To val, 1 To val)
                
                ' Redefine the pattern to find just the data points
                .Pattern = "\s+([0-9]+)\s*(\s*[-]?[0-9]+\.[0-9]{6})(\s*[-]?[0-9]+\.[0-9]{6})?" & _
                        "(\s*[-]?[0-9]+\.[0-9]{6})?(\s*[-]?[0-9]+\.[0-9]{6})?" & _
                        "(\s*[-]?[0-9]+\.[0-9]{6})?(\s*[-]?[0-9]+\.[0-9]{6})?\s*\n"
                
                ' Loop through the Hessian blocks that were found
                For iter = 0 To mchs.Count - 1
                    ' (Re-)initialize rowMatches counter
                    rowMatches = 0
                    
                    ' Within each block, pull the matches from the RegEx...
                    Set mchs2 = .Execute(mchs.item(iter))
                    
                    ' Throw error if number of matches is not consistent
                    If mchs2.Count <> NumAtoms * 3 Then
                        readHESS = "Insufficient data found in Hessian sub-block " & iter + 1
                        GoTo Error_Exit
                    End If
                    
                    ' And loop through all of the internal matches ...
                    For iter2 = 0 To mchs2.Count - 1
                        ' ... to find and store the values of the Hessian
                        ' Each block of the Hessian is six values wide, as currently formulated
                        ' The first SubMatch should correspond to the row of the Hessian
                            '  being populated
                        If Not CLng(mchs2.item(iter2).SubMatches(0)) = iter2 Then
                            readHESS = "Row mismatch error while parsing Hessian, data block " & iter + 1 & _
                                    ", row " & iter2 + 1
                            GoTo Error_Exit
                        End If
                        For iter3 = 1 To 6
                            ' Check whether the relevant submatch exists. This should only come up
                            '  in the final block of the Hessian (if properly formatted)
                            valStr = mchs2.item(iter2).SubMatches(iter3)
                            If Len(valStr) >= 1 Then
                                ' Go ahead and tag the value in ...
                                Hessian(iter2 + 1, iter3 + 6 * iter) = CDbl(valStr)
                                
                                ' ... but either increment the counter (first row) or perform consistency check (others)
                                If iter2 = 0 Then
                                    rowMatches = rowMatches + 1
                                Else
                                    If iter3 > rowMatches Then
                                        readHESS = "Inconsistent number of data points found in Hessian sub-block " & _
                                                iter + 1 & ", row " & iter2 + 1
                                        GoTo Error_Exit
                                    End If
                                End If
                            Else
                                ' This will catch *any* row that's supposed to have six values but doesn't
                                If iter < mchs.Count - 1 Then
                                    readHESS = "Insufficient data found while parsing Hessian: " & _
                                            "data block " & iter + 1 & ", row " & iter2 + 1
                                    GoTo Error_Exit
                                End If
                            End If
                        Next iter3
                    Next iter2
                Next iter
            Else
                readHESS = "Error parsing Hessian data block -- data not in expected format"
                GoTo Error_Exit
            End If
        Else
            readHESS = "Hessian data block could not be found in the indicated file"
            GoTo Error_Exit
        End If
        
        
        ' == FREQUENCIES ==
        ' Reset the rx pattern to find the frequencies block
        .Pattern = "\$vibrational_frequencies\s*\n([0-9]+)\s*\n" & _
                "(\s*([0-9]+)\s*[-]?[0-9]+\.[0-9]+\s*\n)+\s*\n"
        If .Test(str) Then
            ' Store the block and the number of frequencies
            Set mch = .Execute(str).item(0)
            NumFreqs = CLng(mch.SubMatches(0))
            str2 = mch.Value
            
            ' Reset the pattern to pull the individual frequencies
            .Pattern = "\s*[0-9]+\s*([-]?[0-9]+\.[0-9]+)\s*\n"
            Set mchs = .Execute(str2)
            ReDim Freqs(1 To NumFreqs, 1 To 1)
            For iter = 1 To NumFreqs
                Freqs(iter, 1) = CDbl(mchs.item(iter - 1).SubMatches(0))
            Next iter
        Else
            readHESS = "Vibrational frequencies block could not be found in the indicated file"
            GoTo Error_Exit
        End If
        
        ' == MODES ==
        ' Reset the pattern to find the modes block
        .Pattern = "\$normal_modes\s*\n([0-9]+)\s+([0-9]+)\s*\n" & _
                "(([^0-9$\n-]+[0-9.-]+)+\s*\n)+"
        If .Test(str) Then
            ' Store the block and the indicated size of the modes array; perform
            '  consistency check on the size specs
            Set mch = .Execute(str).item(0)
            str2 = mch.Value
            val = mch.SubMatches(0)
            val2 = mch.SubMatches(1)
            If Not (val = NumAtoms * 3 And val = val2) Then
                readHESS = "Normal mode block dimensionality specification (" & val & _
                        "x" & val2 & ") does not " & _
                        "match 3*number of atoms (" & 3 * NumAtoms & ")"
                GoTo Error_Exit
            End If
            
            ' Identify the subsections of the normal modes; set the pattern
            .Pattern = "(\s+[0-9]+)+\s*\n(\s+[0-9]+(\s*[-]?[0-9]+\.[0-9]{6})+\s*\n)+"
            
            If .Test(str2) Then
                ' If subsection parsing worked, store the Matches and redim the modes array
                Set mchs = .Execute(str2)
                ReDim Modes(1 To val, 1 To val2)  ' val should equal val2...
                
                ' Redefine the pattern to find just the data points
                .Pattern = "\s+([0-9]+)\s*(\s*[-]?[0-9]+\.[0-9]{6})(\s*[-]?[0-9]+\.[0-9]{6})?" & _
                        "(\s*[-]?[0-9]+\.[0-9]{6})?(\s*[-]?[0-9]+\.[0-9]{6})?" & _
                        "(\s*[-]?[0-9]+\.[0-9]{6})?(\s*[-]?[0-9]+\.[0-9]{6})?\s*\n"
                
                ' Loop through the mode blocks that were found
                For iter = 0 To mchs.Count - 1
                    ' (Re-)initialize rowMatches counter
                    rowMatches = 0
                    
                    ' Within each block, pull the matches from the RegEx...
                    Set mchs2 = .Execute(mchs.item(iter))
                    
                    ' Throw error if number of matches is not consistent
                    If mchs2.Count <> NumAtoms * 3 Then
                        readHESS = "Insufficient data found in normal modes sub-block " & iter + 1
                        GoTo Error_Exit
                    End If
                    
                    ' And loop through all of the internal matches ...
                    For iter2 = 0 To mchs2.Count - 1
                        ' ... to find and store the values of the modes
                        ' Each sub-block of the modes content is six values wide, as currently formulated
                        ' The first SubMatch should correspond to the row of the modes list
                            '  being populated
                        If Not CLng(mchs2.item(iter2).SubMatches(0)) = iter2 Then
                            readHESS = "Row mismatch error while parsing normal modes, data block " & iter + 1 & _
                                    ", row " & iter2 + 1
                            GoTo Error_Exit
                        End If
                        For iter3 = 1 To 6
                            ' Check whether the relevant submatch exists. This should only come up
                            '  in the final sub-block of the modes content
                            valStr = mchs2.item(iter2).SubMatches(iter3)
                            If Len(valStr) >= 1 Then
                                ' Go ahead and tag the value in ...
                                Modes(iter2 + 1, iter3 + 6 * iter) = CDbl(valStr)
                                
                                ' ... but either increment the counter (first row) or perform consistency check (others)
                                If iter2 = 0 Then
                                    rowMatches = rowMatches + 1
                                Else
                                    If iter3 > rowMatches Then
                                        readHESS = "Inconsistent number of data points found in normal modes sub-block " & _
                                                iter + 1 & ", row " & iter2 + 1
                                        GoTo Error_Exit
                                    End If
                                End If
                            Else
                                ' This will catch *any* row that's supposed to have six values but doesn't
                                If iter < mchs.Count - 1 Then
                                    readHESS = "Insufficient data found while parsing normal modes: " & _
                                            "data block " & iter + 1 & ", row " & iter2 + 1
                                    GoTo Error_Exit
                                End If
                            End If
                        Next iter3
                    Next iter2
                Next iter
            Else
                readHESS = "Error parsing norma modes data block -- data not in expected format"
                GoTo Error_Exit
            End If
        Else
            readHESS = "Normal modes block could not be found in the indicated file"
            GoTo Error_Exit
        End If
        
    End With
    
    ' Set the success (null string) return value
    readHESS = ""
    
    ' Set the initialized flag
    Initialized = True
    Exit Function
    
Error_Exit:
    readHESS = "ORCA_HESS: " & readHESS
    Exit Function
    
End Function



' ===  PROPERTY GET  ===
Property Get hessianArray() As Variant
    If Initialized Then
        hessianArray = Hessian
    Else
        hessianArray = Empty
    End If
End Property

Property Get freqsArray() As Variant
    If Initialized Then
        freqsArray = Freqs
    Else
        freqsArray = Empty
    End If
End Property

Property Get modesArray() As Variant
    If Initialized Then
        modesArray = Modes
    Else
        modesArray = Empty
    End If
End Property

Property Get atomicNumsArray() As Variant
    If Initialized Then
        atomicNumsArray = AtomicNums
    Else
        atomicNumsArray = Empty
    End If
End Property


Property Get atomicSymsArray() As Variant
    
    Dim iter As Long, workVnt As Variant
    
    If Initialized Then
        ReDim workVnt(1 To NumAtoms, 1 To 1)
        For iter = 1 To NumAtoms
            workVnt(iter, 1) = atomSym(CLng(AtomicNums(iter, 1)))
        Next iter
        atomicSymsArray = workVnt
    Else
        atomicSymsArray = Empty
    End If
    
End Property

Property Get expandedAtomicNumsArray() As Variant
    
    Dim iter As Long, workVnt As Variant
    
    If Initialized Then
        ReDim workVnt(1 To NumAtoms * 3, 1 To 1)
        For iter = 1 To NumAtoms
            workVnt(3 * (iter - 1) + 1, 1) = AtomicNums(iter, 1)
            workVnt(3 * (iter - 1) + 2, 1) = AtomicNums(iter, 1)
            workVnt(3 * (iter - 1) + 3, 1) = AtomicNums(iter, 1)
        Next iter
        expandedAtomicNumsArray = workVnt
    Else
        expandedAtomicNumsArray = Empty
    End If
End Property

Property Get expandedAtomicSymsArray() As Variant
    
    Dim iter As Long, workVnt As Variant
    
    If Initialized Then
        ReDim workVnt(1 To NumAtoms * 3, 1 To 1)
        For iter = 1 To NumAtoms
            workVnt(3 * (iter - 1) + 1, 1) = atomSym(CLng(AtomicNums(iter, 1)))
            workVnt(3 * (iter - 1) + 2, 1) = atomSym(CLng(AtomicNums(iter, 1)))
            workVnt(3 * (iter - 1) + 3, 1) = atomSym(CLng(AtomicNums(iter, 1)))
        Next iter
        expandedAtomicSymsArray = workVnt
    Else
        expandedAtomicSymsArray = Empty
    End If
End Property

Property Get atomCoordsArray() As Variant
    If Initialized Then
        atomCoordsArray = AtomCoords
    Else
        atomCoordsArray = Empty
    End If
End Property

Property Get atomicWeightsArray() As Variant
    If Initialized Then
        atomicWeightsArray = AtomicWeights
    Else
        atomicWeightsArray = Empty
    End If
End Property

Property Get atomicWeightsVector() As Variant
    
    Dim iter As Long, workVnt As Variant
    
    If Initialized Then
        ReDim workVnt(1 To NumAtoms * 3, 1 To 1)
        For iter = 1 To NumAtoms
            workVnt(3 * (iter - 1) + 1, 1) = AtomicWeights(iter, 1)
            workVnt(3 * (iter - 1) + 2, 1) = AtomicWeights(iter, 1)
            workVnt(3 * (iter - 1) + 3, 1) = AtomicWeights(iter, 1)
        Next iter
        atomicWeightsVector = workVnt
    Else
        atomicWeightsVector = Empty
    End If
End Property

Property Get atomicWeightsMatrix() As Variant
    If Initialized Then
        atomicWeightsMatrix = vToDiag(Me.atomicWeightsVector)
    Else
        atomicWeightsMatrix = Empty
    End If
End Property

Property Get rootAtWtsMatrix() As Variant
    Dim vt As Variant, iter As Long
    
    If Initialized Then
        ReDim vt(1 To NumAtoms * 3, 1 To 1)
        For iter = 1 To NumAtoms * 3
            vt(iter, 1) = Me.atomicWeightsVector(iter, 1) ^ 0.5
        Next iter
        
        rootAtWtsMatrix = vToDiag(vt)
    Else
        rootAtWtsMatrix = Empty
    End If
End Property

Property Get rootInvAtWtsMatrix() As Variant
    If Initialized Then
        rootInvAtWtsMatrix = Application.WorksheetFunction.MInverse(Me.rootAtWtsMatrix)
    Else
        rootInvAtWtsMatrix = Empty
    End If
End Property

Property Get numOfAtoms() As Double
    If Initialized Then
        numOfAtoms = NumAtoms
    Else
        numOfAtoms = Uninit_Unsigned_Long
    End If
End Property

Property Get isInitialized() As Boolean
    isInitialized = Initialized
End Property

' ===  CLASS EVENTS  ===
Private Sub Class_Initialize()
    ' Probably not necessary to include explicitly, but can't hurt
    Initialized = False
End Sub