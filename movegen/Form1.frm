VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Control-Move Code Generator"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   5865
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1350
      Width           =   10005
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Text            =   "Select One"
      Top             =   540
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1170
      TabIndex        =   1
      Text            =   "Command1"
      Top             =   90
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate Code!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5580
      TabIndex        =   0
      Top             =   180
      Width           =   3525
   End
   Begin VB.Label Label2 
      Caption         =   "Mouse Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Control Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TrueBoolean As String
Dim Dims As String
Dim Line1 As String, Line2 As String, Line3 As String, Line4 As String, Line5 As String, Line6 As String
Dim RightorLeft As String
Dim FirstLine As String, LastLine As String
Dim EnterLine As String
Dim ControlName As String
Dim SFirstLine As String, SLastLine As String
Dim SLine1 As String, SLine2 As String, SLine3 As String, SLine4 As String
Dim TFirstLine As String, TLastLine As String
Dim TLine1 As String, TLine As String, TLine3 As String

Private Sub Command1_Click()
ControlName = Text1.Text

If Combo1.Text = "Right Button Click" Then
   
    RightorLeft = "vbRightButton"
ElseIf Combo1.Text = "Left Button Click" Then
    RightorLeft = "vbLeftButton"
Else
    MsgBox "Error, no mouse button mode selected.", vbOKOnly, "Move Gen"
    Exit Sub
End If






    FirstLine = "Private Sub " & ControlName & "_MouseDown{(}Button As Integer, Shift As Integer, X As Single, Y As Single{)}"
Dims = "Dim OldX as Integer, OldY as Integer, MoveIt as Boolean"
Line1 = "    If Button = " & RightorLeft & " Then"
Line2 = "        OldX = X"
Line3 = "        OldY = Y"
Line4 = "        MoveIt = True"
Line5 = "    End If"
    LastLine = "End Sub"

    SFirstLine = "Private Sub " & ControlName & "_MouseMove{(}Button As Integer, Shift As Integer, X As Single, Y As Single{)}"
SLine1 = "If MoveIt = True Then"
SLine2 = "    " & ControlName & ".Top = " & ControlName & ".Top {+} Y - OldY"
SLine3 = "    " & ControlName & ".Left = " & ControlName & ".Left {+} X - OldX"
SLine4 = "End If"
    SLastLine = "End Sub"

    TFirstLine = "Private Sub " & ControlName & "_MouseUp{(}Button As Integer, Shift As Integer, X As Single, Y As Single{)}"
TLine1 = "    MoveIt = False"
    TLastLine = "End Sub"

'DONE SETTING VARIABLES

Text2.SetFocus
Text2.Text = ""

SendKeys "Option Explicit"
SendKeys "{Enter}"
SendKeys Dims
SendKeys "{Enter}"
SendKeys "{Enter}"
SendKeys FirstLine
SendKeys "{Enter}"
SendKeys Line1
SendKeys "{Enter}"
SendKeys Line2
SendKeys "{Enter}"
SendKeys Line3
SendKeys "{Enter}"
SendKeys Line4
SendKeys "{Enter}"
SendKeys Line5
SendKeys "{Enter}"
SendKeys LastLine
SendKeys "{Enter}"
SendKeys "{Enter}"
SendKeys SFirstLine
SendKeys "{Enter}"
SendKeys SLine1
SendKeys "{Enter}"
SendKeys SLine2
SendKeys "{Enter}"
SendKeys SLine3
SendKeys "{Enter}"
SendKeys SLine4
SendKeys "{Enter}"
SendKeys SLastLine
SendKeys "{Enter}"
SendKeys "{Enter}"
SendKeys TFirstLine
SendKeys "{Enter}"
SendKeys TLine1
SendKeys "{Enter}"
SendKeys TLastLine
SendKeys "{Enter}"


End Sub


Private Sub Form_Load()
RightorLeft = "vbRightButton"
ControlName = "Command1"
Combo1.AddItem "Right Button Click"
Combo1.AddItem "Left Button Click"
End Sub
