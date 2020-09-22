VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Ultimate Custom Serial Number Generator"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "customserial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "How to open serial numbers"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "DL"
      Top             =   480
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar PG3 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Text            =   "5000"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "installinfo"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3120
      MaxLength       =   5
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "-"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Numbers:"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Real Serial Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Serial Number Prefix:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "# of Lines:"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Command1.Caption = "Generate " & Text1.Text & Combo1.Text

End Sub

Private Sub Combo1_Click()
Command1.Caption = "Generate " & Text1.Text & Combo1.Text

End Sub

Private Sub Combo2_Change()
Text4.MaxLength = Combo2.Text
Text4.Text = Left(Text4.Text, Text4.MaxLength)

End Sub

Private Sub Combo2_Click()
Text4.MaxLength = Combo2.Text
Text4.Text = Left(Text4.Text, Text4.MaxLength)

End Sub

Private Sub Combo3_Change()
Text6.MaxLength = Combo3.Text
Text6.Text = Left(Text6.Text, Text6.MaxLength)

End Sub

Private Sub Combo3_Click()
Text6.MaxLength = Combo3.Text
Text6.Text = Left(Text6.Text, Text6.MaxLength)

End Sub

Private Sub Command1_Click()
Call GEN_SERIAL(App.Path & "\" & Text1.Text & Combo1.Text, Text5.Text, Combo2.Text, Combo3.Text)
PG3.Value = 0

End Sub

Private Sub Command2_Click()
MsgBox "To use the Custom Serial Number Generator, all you have to do is write a module or whatever you want to use to read the serial number between the two periods in " & Text1.Text & Combo1.Text & "  I will probably release an application that does this or generates code for you to do this soon!  Thanks for using my program!", vbInformation, "How to use"

End Sub

Private Sub Form_Load()
Combo1.AddItem ".DLL"
Combo1.AddItem ".Exe"
Combo1.AddItem ".Mp3"
Combo1.Text = ".DLL"
Command1.Caption = "Generate " & Text1.Text & Combo1.Text
num% = 2
Do Until num% = 6
Combo2.AddItem num%
Combo3.AddItem num%
num% = num% + 1
Loop
Combo2.Text = "3"
Combo3.Text = "4"
Text5.Text = Text3.Text & "-"
Text6.MaxLength = Combo3.Text
Text4.MaxLength = Combo2.Text
Text7.Text = Text5.Text & Text4.Text & "-" & Text6.Text

End Sub

Function GEN_SERIAL(FILE_NAME As String, Prefix As String, LEN1 As Integer, LEN2 As Integer)
'**********************************************************
'* This function was created by Daniel Errante.           *
'* Copyright © 1999 Daniel Errante.  All rights reserved. *
'**********************************************************
Dim THEINT, THEINT2 As Integer
If LEN1 = 2 Then
On Error GoTo err:
THEINT = 100
ElseIf LEN1 = 3 Then
On Error GoTo err:
THEINT = 1000
ElseIf LEN1 = 4 Then
On Error GoTo err:
THEINT = 10000
ElseIf LEN1 = 5 Then
On Error GoTo err:
THEINT = 100000
End If
If LEN2 = 2 Then
On Error GoTo err:
THEINT2 = 100
ElseIf LEN2 = 3 Then
On Error GoTo err:
THEINT2 = 1000
ElseIf LEN2 = 4 Then
On Error GoTo err:
THEINT2 = 10000
ElseIf LEN2 = 5 Then
On Error GoTo err:
THEINT2 = 100000
End If
Screen.MousePointer = vbHourglass
On Error GoTo err:
CHARS_IN_FILE% = Text2.Text / 2
With Form1
.PG3.Max = CHARS_IN_FILE%
Open FILE_NAME For Output As #1
For i% = 1 To CHARS_IN_FILE%
'PART TO CHANGE:
Print #1, Int(Rnd * THEINT) & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Int(Rnd * THEINT) & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2)
'END
.PG3.Value = i%
val1% = CHARS_IN_FILE% - i%
equation% = val1% / i%
.Caption = "Writing..." & i% & " of " & CHARS_IN_FILE% & " bytes (" & val1% & " bytes left)."
Next i%
Print #1, Int(Rnd * THEINT) & Int(Rnd * THEINT2) & "." & Text7.Text & "." & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Int(Rnd * THEINT) & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2)
For i% = 1 To CHARS_IN_FILE%
'PART TO CHANGE:
Print #1, Int(Rnd * THEINT) & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Int(Rnd * THEINT) & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2) & Prefix & Int(Rnd * THEINT) & "-" & Int(Rnd * THEINT2)
'END
.PG3.Value = i%
val1% = CHARS_IN_FILE% - i%
equation% = val1% / i%
.Caption = "Writing..." & i% & " of " & CHARS_IN_FILE% & " bytes (" & val1% & " bytes left)."
Next i%
Close #1
Screen.MousePointer = vbDefault
.Caption = "The Ultimate Custom Serial Number Generator"
End With
Exit Function
err:
If err.Number = 6 Then
MsgBox "There are too many serial numbers or they are too long.  Please shorten them.", vbExclamation
Combo2.SetFocus
Screen.MousePointer = vbDefault

Exit Function
End If

'If err.Number = 13 Then
'MsgBox "Please enter the number of lines that the generated file will be!", vbExclamation
'Text3.SelStart = 0
'Text3.SelLength = Len(Text3.Text)
'Text3.SetFocus
'Screen.MousePointer = 0
'Exit Function
'End If
MsgBox err.Description, vbExclamation
Screen.MousePointer = vbDefault

End Function

Private Sub Text1_Change()
Command1.Caption = "Generate " & Text1.Text & Combo1.Text

End Sub

Private Sub Text3_Change()
Text5.Text = Text3.Text & "-"

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 45 Then
KeyAscii = 0
Beep
ElseIf KeyAscii >= 48 And KeyAscii <= 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub Text4_Change()
Text7.Text = Text5.Text & Text4.Text & "-" & Text6.Text

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 45 Then
KeyAscii = 0
Beep
ElseIf KeyAscii < 48 And KeyAscii > 32 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub Text5_Change()
Text7.Text = Text5.Text & Text4.Text & "-" & Text6.Text

End Sub

Private Sub Text6_Change()
Text7.Text = Text5.Text & Text4.Text & "-" & Text6.Text

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 45 Then
KeyAscii = 0
Beep
ElseIf KeyAscii < 48 And KeyAscii > 32 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub
