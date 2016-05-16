Attribute VB_Name = "Util_Vector"
'-------------------------------------------------------------------------------
' Name:        Util_Vector (Module)
' Purpose:     Helper functions for ORCA Importer taking 1-D array or
'              Nx1/1xN 2-D array inputs
'
' Author:      Brian Skinn
'                bskinn@alum.mit.edu
'
' Created:     11 May 2016
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

' Vector functions can take either 1-D arrays, or 2-D arrays with
'  dimensions 1xN or Nx1.

' Unless otherwise specified, all functions returning vector-like arrays will
'  return ******Nx1 2-D arrays******

' Most (all?) of these functions will handle Arrays of any dimension base,
'  but all are configured to return Base 1 Arrays.  This is, in part,
'  in order to ensure that the Excel worksheet interface plays nicely
'  with them.



Public Function vToColMat(ByVal vec As Variant) As Variant
    ' Convert/coerce a vector to an Nx1 column matrix
    Dim mSpecs As Scripting.Dictionary
    Dim workVt As Variant
    Dim iter As Long
    Dim wsf As WorksheetFunction
    
    ' Bind wsf
    Set wsf = Application.WorksheetFunction
    
    ' Insist on an array
    vec = arrayify(vec)
    assertIsArray vec
    assertIsVector vec
    
    ' Store specs
    Set mSpecs = mtxSpecs(vec)
    
    If arrRank(vec) = 2 Then
        ' Already a 2-D array
        If mSpecs(msLB1) = 1 And mSpecs(msUB1) = 1 Then
            ' Already Base 1; just transpose if needed
            If mSpecs(msDim2) > mSpecs(msDim1) Then
                vToColMat = mTranspose(vec)
            Else
                vToColMat = vec
            End If
        Else
            ' Not Base 1. Need to re-make
            ReDim workVt(1 To wsf.Max(mSpecs(msDim1), mSpecs(msDim2)), 1 To 1)
            For iter = 1 To UBound(workVt, 1)
                If mSpecs(msDim2) > mSpecs(msDim1) Then
                    ' Row vector input
                    workVt(iter, 1) = vec(mSpecs(msLB1), mSpecs(msLB2) + iter - 1)
                Else
                    workVt(iter, 1) = vec(mSpecs(msLB1) + iter - 1, mSpecs(msLB2))
                End If
            Next iter
            
            ' Assign the return
            vToColMat = workVt
        End If
    Else
        ' 1-D array; will need to iterate and fill
        ReDim workVt(1 To mSpecs(msDim1), 1 To 1)
        For iter = 1 To mSpecs(msDim1)
            workVt(iter, 1) = vec(mSpecs(msLB1) + iter - 1)
        Next iter
        
        ' Assign the return
        vToColMat = workVt
    End If
    
End Function

Public Function vToRowMat(ByVal vec As Variant) As Variant
    ' Convert/coerce a vector to an Nx1 column matrix
    Dim mSpecs As Scripting.Dictionary
    Dim workVt As Variant
    Dim iter As Long
    Dim wsf As WorksheetFunction
    
    ' Bind wsf
    Set wsf = Application.WorksheetFunction
    
    ' Insist on an array
    vec = arrayify(vec)
    assertIsArray vec
    assertIsVector vec
    
    ' Store specs
    Set mSpecs = mtxSpecs(vec)
    
    If arrRank(vec) = 2 Then
        ' Already a 2-D array
        If mSpecs(msLB1) = 1 And mSpecs(msUB1) = 1 Then
            ' Already Base 1; just transpose if needed
            If mSpecs(msDim1) > mSpecs(msDim2) Then
                vToRowMat = mTranspose(vec)
            Else
                vToRowMat = vec
            End If
        Else
            ' Not Base 1. Need to re-make
            ReDim workVt(1 To 1, 1 To wsf.Max(mSpecs(msDim1), mSpecs(msDim2)))
            For iter = 1 To UBound(workVt, 2)
                If mSpecs(msDim2) > mSpecs(msDim1) Then
                    ' Row vector input
                    workVt(1, iter) = vec(mSpecs(msLB1), mSpecs(msLB2) + iter - 1)
                Else
                    workVt(1, iter) = vec(mSpecs(msLB1) + iter - 1, mSpecs(msLB2))
                End If
            Next iter
            
            ' Assign the return
            vToRowMat = workVt
        End If
    Else
        ' 1-D array; will need to iterate and fill
        ReDim workVt(1 To 1, 1 To mSpecs(msDim1))
        For iter = 1 To mSpecs(msDim1)
            workVt(1, iter) = vec(mSpecs(msLB1) + iter - 1)
        Next iter
        
        ' Assign the return
        vToRowMat = workVt
    End If
    
End Function

Public Function vToDiag(ByVal vec As Variant) As Variant
    ' Convert a 1-D or an Nx1 or 1xN 2-D array to a square 2-D diagonal array
    '
    ' Raises RTE #13 (type mismatch) if vec has a bad shape
    
    Dim workVt() As Variant, mDim As Long, iter As Long
    Dim mSpecs As Scripting.Dictionary
    Dim wsf As WorksheetFunction
    
    ' Attach the wsf
    Set wsf = Application.WorksheetFunction
    
    ' arrayify the input
    vec = arrayify(vec)
    
    ' Complain if it's not now an array
    assertIsArray vec
    
    ' Raise error if not 1-D, or 2-D Nx1 or 1xN
    assertIsVector vec
    
    ' Store the dimensions for convenience
    Set mSpecs = mtxSpecs(vec)
    
    ' Store the length of the long dimension for convenience
    mDim = wsf.Max(mSpecs(msDim1), mSpecs(msDim2))
    
    ' Redimension the working Variant
    ReDim workVt(1 To mDim, 1 To mDim)
    
    ' Iterate to fill the work variable
    For iter = 1 To mDim
        If arrRank(vec) = 2 Then
            If mSpecs(msDim1) = 1 Then
                ' Row vector; fill accordingly
                workVt(iter, iter) = vec(mSpecs(msLB1), mSpecs(msLB2) - 1 + iter)
            Else
                ' Column vector; fill accordingly
                workVt(iter, iter) = vec(mSpecs(msLB1) - 1 + iter, mSpecs(msLB2))
            End If
        Else
            ' 1-D vector; fill accordingly
            workVt(iter, iter) = vec(mSpecs(msLB1) - 1 + iter)
        End If
    Next iter
    
    ' Return the result
    vToDiag = workVt
    
End Function

Public Function vDot(ByVal vec1 As Variant, ByVal vec2 As Variant) As Variant
    ' Dot product of two 1-D or Nx1 or 1xN 2-D arrays, with freely mixed
    '  row/column orientation.
    
    Dim mSpecs1 As Scripting.Dictionary, mSpecs2 As Scripting.Dictionary
    Dim iter As Long, accum As Double
    Dim rowScan1 As Long, rowScan2 As Long
    Dim val1 As Double, val2 As Double
    Dim wsf As WorksheetFunction
    
    ' Bind wsf
    Set wsf = Application.WorksheetFunction
    
    ' Ensure inputs are vectors
    vec1 = arrayify(vec1)
    vec2 = arrayify(vec2)
    assertIsArray vec1
    assertIsArray vec2
    assertIsVector vec1
    assertIsVector vec2
    
    ' Pull specs
    Set mSpecs1 = mtxSpecs(vec1)
    Set mSpecs2 = mtxSpecs(vec2)
    
    ' Complain if dimensions mismatch
    assertEqual wsf.Max(mSpecs1(msDim1), mSpecs1(msDim2)), _
                wsf.Max(mSpecs2(msDim1), mSpecs2(msDim2))
    
    ' Define the rowScan variables: 1 = scan down row; 0 = scan across column
    If arrRank(vec1) = 2 Then
        If mSpecs1(msDim1) > 1 Then
            rowScan1 = 1
        End If
    End If
    If arrRank(vec2) = 2 Then
        If mSpecs2(msDim1) > 1 Then
            rowScan2 = 1
        End If
    End If
    
    ' Iterate, accumulate, and return
    accum = 0#
    For iter = 1 To wsf.Max(mSpecs1(msDim1), mSpecs1(msDim2))
        If mSpecs1(msDim2) = 0 Then
            val1 = vec1(mSpecs1(msLB1) + iter - 1)
        Else
            val1 = vec1(mSpecs1(msLB1) + rowScan1 * (iter - 1), _
                        mSpecs1(msLB2) + (1 - rowScan1) * (iter - 1))
        End If
        
        If mSpecs2(msDim2) = 0 Then
            val2 = vec2(mSpecs2(msLB1) + iter - 1)
        Else
            val2 = vec2(mSpecs2(msLB1) + rowScan2 * (iter - 1), _
                        mSpecs2(msLB2) + (1 - rowScan2) * (iter - 1))
        End If
        
        accum = accum + (val1 * val2)
    Next iter
    
    vDot = accum
    
End Function

' Function vProj TODO

' Function vRej TODO

' Function vCross TODO

Public Function vNorm(ByVal vec As Variant) As Variant
    ' Return the 2-norm of the input vector. Accepts both 1-D and
    '  Nx1 or 1xN 2-D vectors.
    Dim iter As Long, accum As Double
    Dim mSpecs As Scripting.Dictionary
    
    ' Ensure input is a vector
    vec = arrayify(vec)
    assertIsArray vec
    assertIsVector vec
    
    ' Pull specs
    Set mSpecs = mtxSpecs(vec)
    
    ' If 2-D, just cheat and use the matrix function
    If arrRank(vec) = 2 Then
        vNorm = mNorm(vec)
    Else
        ' 1-D, must calculate here
        accum = 0#
        For iter = mSpecs(msLB1) To mSpecs(msUB1)
            accum = accum + vec(iter) ^ 2#
        Next iter
        
        vNorm = accum ^ 0.5
    End If
    
End Function

Public Function vNormalize(ByVal vec As Variant) As Variant
    ' Return the normalized input vector. ALWAYS RETURNS AN NX1 COLUMN
    '  VECTOR. Both 1-D and Nx1/1xN 2-D vectors are accepted.
    
    Dim iter As Long, normVal As Double
    Dim workVt As Variant
    Dim mSpecs As Scripting.Dictionary
    
    ' Arrayify and confirm is vector
    vec = arrayify(vec)
    assertIsArray vec
    assertIsVector vec
    
    ' Pull specs
    Set mSpecs = mtxSpecs(vec)
    
    If arrRank(vec) = 2 Then
        ' If 2-D, cheat and use the matrix function
        workVt = mNormalize(vec)
        
        ' Transpose if needed to get the column vector output
        If mSpecs(msDim2) > mSpecs(msDim1) Then
            workVt = mTranspose(workVt)
        End If
        
        ' Set the return value
        vNormalize = workVt
    Else
        ' 1-D; must handle here
        ' Resize the working array
        ReDim workVt(1 To mSpecs(msDim1), 1 To 1)
        
        ' Retrieve the vector norm
        normVal = vNorm(vec)
        
        ' Loop and calculate the normalized values
        For iter = 1 To mSpecs(msDim1)
            workVt(iter, 1) = (1# / normVal) * vec(mSpecs(msLB1) + iter - 1)
        Next iter
        
        ' Store the return
        vNormalize = workVt
        
    End If
    
End Function

' Function vOrthoBasis TODO
' Function vLength TODO
' Function vAngle TODO

'Public Function vecsOrthonormCheck(ParamArray vecs()) As VbTriState
'    ' ##TODO## CONVERT TO STRING RETURN TYPE
'    Dim maxVecIdx As Long, vecLength As Long
'    Dim iter As Long, iter2 As Long
'    Dim orthonormTol As Double
'
'    ' ParamArray is always Base 0, regardless of Option Base setting
'    ' A single double value can be passed as the last argument to adjust
'    '  the orthonormality tolerance. Otherwise, the default value is used.
'
'    ' No-good return if no argument passed
'    If IsMissing(vecs) Then
'        vecsOrthonormCheck = vbUseDefault
'        Exit Function
'    End If
'
'    ' Store the index of the last element in the vector array
'    maxVecIdx = UBound(vecs)
'
'    ' Check if the last element is a non-array single value. If so, treat as the tolerance
'    '  and decrement the max vector index (crashing out if this leaves no vectors).
'    '  Otherwise, set the default tolerance.
'    If IsNumeric(vecs(maxVecIdx)) Then
'        orthonormTol = CDbl(vecs(maxVecIdx))
'        If maxVecIdx > 0 Then
'            maxVecIdx = maxVecIdx - 1
'        Else
'            vecsOrthonormCheck = vbUseDefault
'            Exit Function
'        End If
'    Else
'        orthonormTol = DEF_Orthonorm_Tol
'    End If
'
'    ' arrayify all of the vectors, dumping out if any arguments aren't workable
'    For iter = 0 To maxVecIdx
'        vecs(iter) = arrayify(vecs(iter))
'        If IsEmpty(vecs(iter)) Then
'            vecsOrthonormCheck = vbUseDefault
'            Exit Function
'        End If
'    Next iter
'
'    ' Store the length of the first vector
'    vecLength = UBound(vecs(0), 1)
'
'    ' Check to ensure all vectors are this length, crashing out if not
'    For iter = 0 To maxVecIdx
'        If UBound(vecs(iter), 1) <> vecLength Then
'            vecsOrthonormCheck = vbUseDefault
'            Exit Function
'        End If
'    Next iter
'
'    ' Initialize the success return
'    vecsOrthonormCheck = vbTrue
'
'    ' Loop through the vectors, confirming orthonormality.
'    ' If any fail, set the fail return and dump from function.
'    For iter = 0 To maxVecIdx
'        For iter2 = iter To maxVecIdx
'            If (Abs(ProdScal(vecs(iter), vecs(iter2))) - deltaFxn(iter, iter2)) _
'                    > CDbl(orthonormTol) Then
'                vecsOrthonormCheck = vbFalse
'                Exit Function
'            End If
'        Next iter2
'    Next iter
'
'End Function


