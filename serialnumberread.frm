VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form scanserial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial Number Scanner with improved read speed"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "serialnumberread.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan Serial Number"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Length:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Serial Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "scanserial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
Call READ_FILE(Text1.Text)
Screen.MousePointer = vbDefault
If Right(Text2.Text, 1) = "." Then Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)

End Sub

Function READ_FILE(FILE_NAME As String)
Call GET_LEN
Dim cap As String
Dim w, h As Long
cap = Me.Caption
Me.Caption = "Opening File Properties...Please wait..."
ProgressBar1.Max = Text3.Text
Me.Caption = "Preparing to load Serial Numbers..."
Open FILE_NAME For Input As #1
Dim i As Long
Dim d As Long
i = 1
h = 0
Me.Caption = "Scanning Serial Number File..."
Do Until EOF(1)
Line Input #1, lineoftext$
d = Len(lineoftext$) + 1
    pg1.Max = d
    i = 1
    Do Until i = d
    letter$ = Mid$(lineoftext$, i, 1)
    If letter$ = "." Then
    Text2.Text = Mid(lineoftext$, i + 1, Len(lineoftext$))
    Call GET_REST
    GoTo clean:
    End If
    i = i + 1
    pg1.Value = i
    Me.Caption = cap & " - Location of scanner: " & "Letter " & i & " of line " & h
    Loop
h = h + 1
ProgressBar1.Value = h
Loop
clean:
ProgressBar1.Value = 0
pg1.Value = 0
Me.Caption = cap
End Function

Function GET_REST()
Me.Caption = "Scanning rest of Serial Number File..."
Dim d As String
Dim i As Integer
Do Until i = Len(Text2.Text)
d = Right(Text2.Text, 1)
txt$ = Left(Text2.Text, Len(Text2.Text) - 1)
If d = "." Then
Exit Function
End If
Text2.Text = txt$
Loop

End Function

Private Sub Command2_Click()
End

End Sub

Function GET_LEN()
Me.Caption = "Opening File Properties...This may take a few minutes..."
Dim alltext As String
Dim lineoftext As String
Dim b As Long
Close #1
On Error GoTo err:
Open Text1.Text For Input As #1
Do Until EOF(1)
Line Input #1, lineoftext
b = b + 1
Loop
Text3.Text = b / 2
Close #1
Me.Caption = "Serial Number Scanner"
Exit Function
err:
MsgBox err.Description
'MsgBox "Error!  Setup file 'installinfo.exe' not found!  Please ask me for another copy of this file.  E-mail me at danoph@hotmail.com.  Please include your name, e-mail, and serial number in the e-mail.", vbExclamation
'End
End Function

