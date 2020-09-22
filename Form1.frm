VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ColorPicker Test"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ctlFarver ctlFarver1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "HEXColor:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "VBColor:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ctlFarver1_SelectColor()
  Label2.Caption = "VBColor: " & ctlFarver1.VBColor
  Label3.Caption = "HexColor: " & ctlFarver1.HexColor
End Sub
