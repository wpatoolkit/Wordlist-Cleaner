Attribute VB_Name = "Module1"
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Type tagInitCommonControlsEx
 lngSize As Long
 lngICC As Long
End Type

Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
 On Error Resume Next
 Dim iccex As tagInitCommonControlsEx
 iccex.lngSize = LenB(iccex)
 iccex.lngICC = ICC_USEREX_CLASSES
 InitCommonControlsEx iccex
 Load Form1
 Form1.Show
End Sub
