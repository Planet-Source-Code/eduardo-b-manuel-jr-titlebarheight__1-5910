VERSION 5.00
Begin VB.Form frmTitBarHeightSamp 
   Caption         =   "Find Title Bar Height Sample App"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2.219
   ScaleMode       =   5  'Inch
   ScaleWidth      =   3.25
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2985
      TabIndex        =   3
      Top             =   630
      Width           =   1170
   End
   Begin VB.TextBox txtTitBarHeight 
      Height          =   315
      Left            =   3285
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   855
   End
   Begin VB.OptionButton OptScaleMode 
      Caption         =   "User"
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   195
      Width           =   2205
   End
   Begin VB.Label lblTitBarHeight 
      Caption         =   "Title Bar Height"
      Height          =   195
      Left            =   2010
      TabIndex        =   2
      Top             =   270
      Width           =   1365
   End
End
Attribute VB_Name = "frmTitBarHeightSamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is the sample application for the function.
'Experiment with the form's ScaleMode property and you'll get the hang of things
Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Load()
Dim i As Integer
'Load remaining Options by code
For i = 1 To 8
    Load OptScaleMode(i)
    'Set the position(the next one always 300 twips down from the last).
    OptScaleMode(i).Top = OptScaleMode(i - 1).Top + Me.ScaleY(300, 1, Me.ScaleMode)
    'All are visible
    OptScaleMode(i).Visible = True
Next i
'set each OptionBox caption
OptScaleMode(1).Caption = "Twips"
OptScaleMode(2).Caption = "Points"
OptScaleMode(3).Caption = "Pixels"
OptScaleMode(4).Caption = "Characters"
OptScaleMode(5).Caption = "Inches"
OptScaleMode(6).Caption = "Millimeters"
OptScaleMode(7).Caption = "Centimeters"
OptScaleMode(8).Caption = "This form's ScaleMode"
'Default Measurement
OptScaleMode(1).Value = True
End Sub

Private Sub OptScaleMode_Click(Index As Integer)
'If index=8, do not pass the second parameter
txtTitBarHeight.Text = IIf(Index = 8, TitleBarHeight(Me), TitleBarHeight(Me, Index))
End Sub

Private Sub txtTitBarHeight_GotFocus()
'Don't allow this textbox to get focus
cmdExit.SetFocus
End Sub
