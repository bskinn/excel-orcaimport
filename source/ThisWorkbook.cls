VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public DEBUG_FLAG As Boolean
Public VOLATILE_FLAG As Boolean

Private Sub Workbook_Open()
    DEBUG_FLAG = False
    VOLATILE_FLAG = False
End Sub
