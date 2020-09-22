VERSION 5.00
Begin VB.Form frmCalculate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calculate"
   ClientHeight    =   3735
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   295
      Left            =   1800
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1572
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3012
      Begin VB.TextBox txtLogBase 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   1680
         TabIndex        =   7
         Text            =   "10"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox txtDecimals 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   1680
         TabIndex        =   6
         Text            =   "14"
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox txtBaseMode 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   240
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox txtAngleMode 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   240
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label lblLogBase 
         AutoSize        =   -1  'True
         Caption         =   "Log Base:"
         Height          =   192
         Left            =   1560
         TabIndex        =   11
         Top             =   840
         Width           =   732
      End
      Begin VB.Label lblDecimals 
         AutoSize        =   -1  'True
         Caption         =   "Decimals:"
         Height          =   192
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblBaseMode 
         AutoSize        =   -1  'True
         Caption         =   "Base Mode:"
         Height          =   192
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   876
      End
      Begin VB.Label lblAngleMode 
         AutoSize        =   -1  'True
         Caption         =   "Angle Mode:"
         Height          =   192
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   912
      End
   End
   Begin VB.TextBox txtOutputString 
      Height          =   1092
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   3012
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   295
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtInputString 
      Height          =   288
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3012
   End
End
Attribute VB_Name = "frmCalculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********
'
'Syntax:
'
'{OutputString} = CalculateString({InputString}, {AngleMode}, {BaseMode}, {Decimals}, {LogBase})
'
'**********
'
'Explanation:
'
'{OutputString} - valid string or text box that can hold the returned string
'
'{InputString} - equation in string form
'
'{AngleMode} - 0 or 1.
'0 = Degrees
'1 = Radians
'
'{BaseMode} - integer between 0 and 3.
'0 = Decimal
'1 = Binary
'2 = Hexadecimal
'3 = Octal
'
'{Decimals} - integer greater than or equal to 0.
'0 - 13 = number of decimal places
' > 13  = floating
'
'{LogBase} - any numeric value
'
'**********

Private Sub cmdCalculate_Click()

    txtOutputString.Text = CalculateString(txtInputString.Text, Val(txtAngleMode.Text), Val(txtBaseMode.Text), Val(txtDecimals.Text), Val(txtLogBase.Text)) + vbNewLine + txtOutputString.Text

End Sub

Private Sub cmdClear_Click()

    txtOutputString.Text = vbNullString

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then cmdCalculate_Click

End Sub
