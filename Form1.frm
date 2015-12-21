VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Wordlist Cleaner"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6870
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNewline 
      Caption         =   "Convert all newlines to Unix format"
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   6330
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CheckBox chkTrimWhitespace 
      Caption         =   "Trim leading and trailing whitespace"
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   5940
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CheckBox chkUpper 
      Caption         =   "Convert all words to uppercase"
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   5550
      Width           =   3855
   End
   Begin VB.CheckBox chkLower 
      Caption         =   "Convert all words to lowercase"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   5175
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CheckBox chkRepeat 
      Caption         =   "Repeat word until it reaches min length"
      Height          =   375
      Left            =   3345
      TabIndex        =   23
      Top             =   4395
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.TextBox txtNumericMax 
      Height          =   360
      Left            =   3795
      TabIndex        =   13
      Text            =   "11"
      Top             =   3630
      Width           =   615
   End
   Begin VB.TextBox txtNumericMin 
      Height          =   360
      Left            =   2715
      TabIndex        =   11
      Text            =   "1"
      Top             =   3630
      Width           =   615
   End
   Begin VB.CheckBox chkRemoveNumeric 
      Caption         =   "Remove numerics between"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3615
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox txtMax 
      Height          =   360
      Left            =   2520
      TabIndex        =   19
      Text            =   "24"
      Top             =   4785
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Height          =   360
      Left            =   2520
      TabIndex        =   17
      Text            =   "8"
      Top             =   4395
      Width           =   615
   End
   Begin VB.CheckBox chkMax 
      Caption         =   "Maximum word length:"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   4785
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.OptionButton OptThrow 
      Caption         =   "Throw away the line"
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   9
      Top             =   3285
      Width           =   2175
   End
   Begin VB.OptionButton OptThrow 
      Caption         =   "Throw away the character"
      Height          =   255
      Index           =   0
      Left            =   510
      TabIndex        =   8
      Top             =   3030
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox txtOnlyAllow 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Text            =   " !""#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]^_`abcdefghijklmnopqrstuvwxyz{|}~"
      Top             =   2280
      Width           =   5655
   End
   Begin VB.CheckBox chkOnlyAllow 
      Caption         =   "Only allow these characters"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Value           =   1  'Checked
      Width           =   6255
   End
   Begin VB.CheckBox chkReplaceAccent 
      Caption         =   "Replace accented characters with non-accented versions (á --> a)"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   4005
      Value           =   1  'Checked
      Width           =   6255
   End
   Begin VB.CommandButton btnOutputFile 
      Caption         =   "..."
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1410
      Width           =   495
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   5655
   End
   Begin VB.CommandButton btnInputFile 
      Caption         =   "..."
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox txtInputFile 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   5655
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   735
      Left            =   3720
      TabIndex        =   22
      Top             =   6990
      Width           =   2295
   End
   Begin VB.CommandButton btnProcess 
      Caption         =   "&Process"
      Default         =   -1  'True
      Height          =   735
      Left            =   960
      TabIndex        =   20
      Top             =   6990
      Width           =   2295
   End
   Begin VB.CheckBox chkMin 
      Caption         =   "Minimum word length:"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   4395
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Label lblSpace 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " | "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2985
      TabIndex        =   29
      Top             =   210
      Width           =   255
   End
   Begin VB.Label lblFileMode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "File Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1935
      TabIndex        =   28
      Top             =   210
      Width           =   975
   End
   Begin VB.Label lblDirectoryMode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Directory Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3405
      TabIndex        =   27
      Top             =   210
      Width           =   1455
   End
   Begin VB.Label lblNumericChar 
      AutoSize        =   -1  'True
      Caption         =   "characters"
      Height          =   195
      Left            =   4530
      TabIndex        =   14
      Top             =   3705
      Width           =   750
   End
   Begin VB.Label lblNumericAnd 
      Caption         =   "and"
      Height          =   375
      Left            =   3405
      TabIndex        =   12
      Top             =   3705
      Width           =   375
   End
   Begin VB.Label lblOptThrow 
      Caption         =   "When a non-allowed character is found:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   5850
   End
   Begin VB.Label lblOutputFile 
      Caption         =   "Output File:"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   1200
      Width           =   2730
   End
   Begin VB.Label lblInputFile 
      Caption         =   "Input File:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2370
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnInputFile_Click()
If lblFileMode.ForeColor = &H80000012 Then 'Browse for File
 Dim file_to_open As String
 file_to_open = ShowOpenFileDialog(Me.hwnd, "Choose Input File")
 If (file_to_open <> "") Then
  txtInputFile.Text = file_to_open
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  txtOutputFile.Text = fso.GetParentFolderName(txtInputFile.Text) & "\" & fso.GetBaseName(txtInputFile.Text) & "_CLEANED." & fso.GetExtensionName(txtInputFile.Text)
  Set fso = Nothing
 End If
Else 'Browse for Directory
 Dim tBrowseInfo As BrowseInfo
 tBrowseInfo.hwndOwner = Me.hwnd
 tBrowseInfo.lpszTitle = lstrcat("Choose Input Directory", "")
 tBrowseInfo.ulFlags = 1 + 2 + &H4&
 Dim tmpLong As Long
 tmpLong = SHBrowseForFolder(tBrowseInfo)
 If (tmpLong) Then
  Dim sBuffer As String
  sBuffer = Space(260)
  SHGetPathFromIDList tmpLong, sBuffer
  sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  txtInputFile.Text = sBuffer
  txtOutputFile.Text = txtInputFile.Text & "_CLEANED"
 End If
End If
End Sub

Private Sub btnOutputFile_Click()
If lblFileMode.ForeColor = &H80000012 Then 'Browse for File
 Dim file_to_save As String
 file_to_save = ShowSaveFileDialog(Me.hwnd, "Choose Output File")
 If (file_to_save <> "") Then
  txtOutputFile.Text = file_to_save
 End If
Else 'Browse for Directory
 Dim tBrowseInfo As BrowseInfo
 tBrowseInfo.hwndOwner = Me.hwnd
 tBrowseInfo.lpszTitle = lstrcat("Choose Output Directory", "")
 tBrowseInfo.ulFlags = 1 + 2 + &H4&
 Dim tmpLong As Long
 tmpLong = SHBrowseForFolder(tBrowseInfo)
 If (tmpLong) Then
  Dim sBuffer As String
  sBuffer = Space(260)
  SHGetPathFromIDList tmpLong, sBuffer
  sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  txtOutputFile.Text = sBuffer
 End If
End If
End Sub

Private Sub process_one_file(input_filename As String, output_filename As String)
Dim fso As Scripting.FileSystemObject
Dim input_text_stream As Scripting.TextStream
Dim output_text_stream As Scripting.TextStream
Set fso = New Scripting.FileSystemObject
Dim one_line_in_orig As String
Dim one_line_in As String
Dim one_line_out As String
Dim one_char As String
Dim one_length As Integer
Dim i As Integer
Dim min_length As Long
Dim max_length As Long
Dim is_numeric As Boolean
Dim min_numeric_length As Long
Dim max_numeric_length As Long

min_numeric_length = 0
If (chkRemoveNumeric.Value = 1) And (Trim$(txtNumericMin.Text) <> "") And (IsNumeric(Trim$(txtNumericMin.Text))) Then
 min_numeric_length = CLng(Trim$(txtNumericMin.Text))
End If

max_numeric_length = 9999999
If (chkRemoveNumeric.Value = 1) And (Trim$(txtNumericMax.Text) <> "") And (IsNumeric(Trim$(txtNumericMax.Text))) Then
 max_numeric_length = CLng(Trim$(txtNumericMax.Text))
End If

min_length = 0
If (chkMin.Value = 1) And (Trim$(txtMin.Text) <> "") And (IsNumeric(Trim$(txtMin.Text))) Then
 min_length = CLng(Trim$(txtMin.Text))
End If

max_length = 9999999
If (chkMax.Value = 1) And (Trim$(txtMax.Text) <> "") And (IsNumeric(Trim$(txtMax.Text))) Then
 max_length = CLng(Trim$(txtMax.Text))
End If

If fso.FileExists(input_filename) Then
Set input_text_stream = fso.OpenTextFile(input_filename, ForReading, False)
Set output_text_stream = fso.OpenTextFile(output_filename, ForWriting, True)
Do While Not (input_text_stream.AtEndOfStream)
 one_line_in = input_text_stream.ReadLine
 one_line_out = ""
 
 'replace special characters
 If (chkReplaceAccent.Value = 1) Then
  Call replace_diacritics(one_line_in)
 End If
 
 'trim whitespace
  If (chkTrimWhitespace.Value = 1) Then
   one_line_in = Replace$(Trim$(one_line_in), Chr$(9), "")
  End If
 
 'repeat word if too short
 one_line_in_orig = one_line_in
 one_length = Len(one_line_in)
 If (chkMin.Value = 1) And (chkRepeat.Value = 1) And (one_length > 0) And (min_length > 0) And (one_length < min_length) Then
 Do While (one_length < min_length)
  one_line_in = one_line_in & one_line_in_orig
  one_length = Len(one_line_in)
 Loop
 End If
 
 is_numeric = True
 If ((chkMin.Value = 0) Or ((chkMin.Value = 1) And (one_length >= min_length))) And ((chkMax.Value = 0) Or ((chkMax.Value = 1) And (one_length <= max_length))) Then
  For i = 1 To one_length
   one_char = Mid$(one_line_in, i, 1)
   If (chkRemoveNumeric.Value = 1) And (is_numeric = True) Then
    If ((one_char <> "0") And (one_char <> "1") And (one_char <> "2") And (one_char <> "3") And (one_char <> "4") And (one_char <> "5") And (one_char <> "6") And (one_char <> "7") And (one_char <> "8") And (one_char <> "9")) Then
     is_numeric = False
    End If
   End If
   If (chkOnlyAllow.Value = 0) Or ((chkOnlyAllow.Value = 1) And (OptThrow(0).Value = True) And (InStr(txtOnlyAllow.Text, one_char) > 0)) Then
    one_line_out = one_line_out & one_char
   End If
  Next
  
  'lower case
  If (chkLower.Value = 1) Then
   one_line_out = LCase$(one_line_out)
  End If
 
  'upper case
  If (chkUpper.Value = 1) Then
   one_line_out = UCase$(one_line_out)
  End If
  
  If (one_line_out <> "") Then
   If (chkRemoveNumeric.Value = 1) And (is_numeric = True) Then
    If (Len(one_line_out) >= min_numeric_length) And (Len(one_line_out) <= max_numeric_length) Then
   
    Else
     output_text_stream.Write one_line_out
     output_text_stream.Write IIf(chkNewline.Value = 0, Chr$(13) & Chr$(10), Chr$(10))
    End If
   Else
    output_text_stream.Write one_line_out
    output_text_stream.Write IIf(chkNewline.Value = 0, Chr$(13) & Chr$(10), Chr$(10))
   End If
  End If
  
 End If
Loop
End If
If Not input_text_stream Is Nothing Then
 Call input_text_stream.Close
End If
If Not output_text_stream Is Nothing Then
 Call output_text_stream.Close
End If
Set input_text_stream = Nothing
Set output_text_stream = Nothing
Set fso = Nothing
End Sub

Private Sub btnProcess_Click()
If (Trim$(txtInputFile.Text) = "") Then
 MsgBox "No input " & IIf(lblFileMode.ForeColor = &H80000012, "file", "directory") & " specified!", vbExclamation + vbOKOnly, "Wordlist Cleaner"
 txtInputFile.SetFocus
 Exit Sub
ElseIf (Trim$(txtOutputFile.Text) = "") Then
 MsgBox "No output " & IIf(lblFileMode.ForeColor = &H80000012, "file", "directory") & " specified!", vbExclamation + vbOKOnly, "Wordlist Cleaner"
 txtOutputFile.SetFocus
 Exit Sub
ElseIf (chkMin.Value = 1) And (Trim$(txtMin.Text) = "") Then
 MsgBox "Minimum word length not specified!", vbExclamation + vbOKOnly, "Wordlist Cleaner"
 txtMin.SetFocus
 Exit Sub
ElseIf (chkMax.Value = 1) And (Trim$(txtMax.Text) = "") Then
 MsgBox "Maximum word length not specified!", vbExclamation + vbOKOnly, "Wordlist Cleaner"
 txtMax.SetFocus
 Exit Sub
ElseIf (lblFileMode.ForeColor = &H80000012) And (is_file(txtInputFile.Text) = False) Then
 MsgBox "Input File does not exist.", vbExclamation + vbOKOnly, "Wordlist Cleaner"
 txtInputFile.SetFocus
 Exit Sub
ElseIf (lblDirectoryMode.ForeColor = &H80000012) And (is_folder(txtInputFile.Text) = False) Then
 MsgBox "Input Directory does not exist.", vbExclamation + vbOKOnly, "Wordlist Cleaner"
 txtInputFile.SetFocus
 Exit Sub
ElseIf (lblDirectoryMode.ForeColor = &H80000012) And (is_folder(txtOutputFile.Text) = False) Then
 If (MsgBox("Output directory does not exist." & vbNewLine & "Would you like to create it?", vbQuestion + vbYesNo, "Wordlist Cleaner") = vbYes) Then
  Dim fso2 As Scripting.FileSystemObject
  Set fso2 = New Scripting.FileSystemObject
  fso2.CreateFolder txtOutputFile.Text
  Set fso2 = Nothing
  If (is_folder(txtOutputFile.Text) = False) Then
   txtOutputFile.SetFocus
   Exit Sub
  End If
 Else
  txtOutputFile.SetFocus
  Exit Sub
 End If
End If

Dim msgbox_result As VbMsgBoxResult
If (txtOutputFile.Text <> "") Then
 If is_file(txtOutputFile.Text) = True Then
  msgbox_result = MsgBox("A file named """ & txtOutputFile.Text & """ already exists. Replace?", vbExclamation + vbYesNo, "Warning")
  If (msgbox_result = vbNo) Then
   Call btnOutputFile_Click
   Exit Sub
  End If
 End If
End If

btnProcess.Enabled = False
If lblFileMode.ForeColor = &H80000012 Then 'Process File
 Call process_one_file(txtInputFile.Text, txtOutputFile.Text)
Else 'Process Folder
 Dim fso As Scripting.FileSystemObject
 Set fso = New Scripting.FileSystemObject
 Dim input_directory As String
 Dim output_directory As String
 input_directory = txtInputFile.Text & IIf(Right$(txtInputFile.Text, 1) <> "\", "\", "")
 output_directory = txtOutputFile.Text & IIf(Right$(txtOutputFile.Text, 1) <> "\", "\", "")
 Dim sFilename As String
 sFilename = Dir(input_directory)
 Do While sFilename > ""
  Call process_one_file(input_directory & sFilename, output_directory & fso.GetBaseName(output_directory & sFilename) & "_CLEANED." & fso.GetExtensionName(output_directory & sFilename))
  sFilename = Dir()
 Loop
 Set fso = Nothing
End If
btnProcess.Enabled = True
MsgBox "Done!", vbInformation + vbOKOnly, "Wordlist Cleaner"
End Sub

Private Sub btnExit_Click()
 Unload Me
End Sub

Private Sub ensure_numbers_only(ByRef box As TextBox)
 Dim i As Integer
 Dim str As String
 If Len(box.Text) > 0 Then
  For i = 1 To Len(box.Text)
   If (Asc(Mid$(box.Text, i, 1)) >= 48) And (Asc(Mid$(box.Text, i, 1)) <= 57) Then
    str = str & Mid$(box.Text, i, 1)
   End If
  Next
  box.Text = str
 End If
End Sub

Private Sub chkRemoveNumeric_Click()
 txtNumericMin.Enabled = chkRemoveNumeric.Value
 txtNumericMax.Enabled = chkRemoveNumeric.Value
 lblNumericAnd.Enabled = chkRemoveNumeric.Value
 lblNumericChar.Enabled = chkRemoveNumeric.Value
End Sub

Private Sub Form_Load()
 RemoveMenu GetSystemMenu(Me.hwnd, 0), 2, &H400& 'prevent resizing
End Sub

Private Sub lblDirectoryMode_Click()
 If lblDirectoryMode.ForeColor = &HFF0000 Then 'blue
  lblDirectoryMode.ForeColor = &H80000012 'black
  lblDirectoryMode.FontUnderline = False
  lblFileMode.ForeColor = &HFF0000 'blue
  lblFileMode.FontUnderline = True
  lblInputFile.caption = "Input Directory:"
  lblOutputFile.caption = "Output Directory:"
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  If (is_file(txtInputFile.Text) = True) Then
   txtInputFile.Text = fso.GetParentFolderName(txtInputFile.Text)
  End If
  If (is_folder(txtInputFile.Text) = True) Then
   txtOutputFile.Text = txtInputFile.Text & "_CLEANED"
  End If
  Set fso = Nothing
 End If
End Sub

Private Sub lblFileMode_Click()
If lblFileMode.ForeColor = &HFF0000 Then 'blue
  lblFileMode.ForeColor = &H80000012 'black
  lblFileMode.FontUnderline = False
  lblDirectoryMode.ForeColor = &HFF0000 'blue
  lblDirectoryMode.FontUnderline = True
  lblInputFile.caption = "Input File:"
  lblOutputFile.caption = "Output File:"
  txtInputFile.Text = ""
  txtOutputFile.Text = ""
 End If
End Sub

Private Sub lblDirectoryMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If lblDirectoryMode.ForeColor = &HFF0000 Then
  SetCursor LoadCursor(0, 32649&)
 End If
End Sub

Private Sub lblFileMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblFileMode.ForeColor = &HFF0000 Then
  SetCursor LoadCursor(0, 32649&)
 End If
End Sub

Private Sub txtMax_LostFocus()
 Call ensure_numbers_only(txtMax)
End Sub

Private Sub txtMin_LostFocus()
 Call ensure_numbers_only(txtMin)
End Sub

Private Sub chkMax_Click()
 txtMax.Enabled = chkMax.Value
End Sub

Private Sub chkMin_Click()
 txtMin.Enabled = chkMin.Value
 chkRepeat.Enabled = chkMin.Value
End Sub

Private Sub chkOnlyAllow_Click()
 If (chkOnlyAllow.Value = 0) Then
  OptThrow(0).Enabled = False
  OptThrow(1).Enabled = False
  lblOptThrow.Enabled = False
  txtOnlyAllow.Enabled = False
 Else
  OptThrow(0).Enabled = True
  OptThrow(1).Enabled = True
  lblOptThrow.Enabled = True
  txtOnlyAllow.Enabled = True
 End If
End Sub

Private Sub txtNumericMin_LostFocus()
 Call ensure_numbers_only(txtNumericMin)
End Sub

Private Sub txtNumericMax_LostFocus()
 Call ensure_numbers_only(txtNumericMax)
End Sub
