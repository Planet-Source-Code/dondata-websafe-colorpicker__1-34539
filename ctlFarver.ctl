VERSION 5.00
Begin VB.UserControl ctlFarver 
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ScaleHeight     =   2415
   ScaleWidth      =   2715
   ToolboxBitmap   =   "ctlFarver.ctx":0000
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   60
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   2140
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   130
      Index           =   0
      Left            =   10
      ScaleHeight     =   68.824
      ScaleMode       =   0  'User
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   330
      Width           =   130
   End
   Begin VB.CommandButton cmdPopup 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2460
      TabIndex        =   2
      Top             =   30
      Width           =   225
   End
   Begin VB.Shape Shape3 
      Height          =   1095
      Left            =   0
      Top             =   315
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   360
      Top             =   705
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   225
      Left            =   0
      Top             =   705
      Width           =   225
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#FFFFFF"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1830
      TabIndex        =   3
      Top             =   2170
      Width           =   765
   End
   Begin VB.Label lblFarver 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "ctlFarver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public initialColor As Long
Public ResultColor As String
Public HexColor As String
Public VBColor As String
Dim ilastIndex As Long
Dim Tilstand As Boolean

Public Event SelectColor()

Function Vis()
  UserControl.Height = 2430
  Tilstand = True
End Function

Function Skjul()
  UserControl.Height = 315
  Tilstand = False
End Function

Private Sub cmdPopup_Click()
  If Tilstand = True Then
     Skjul
  Else
     Vis
  End If
  Picture1(0).SetFocus
End Sub

Private Sub lblFarver_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then cmdPopup_Click
End Sub

Private Sub Picture1_Click(Index As Integer)
  lblFarver.BackColor = Picture1(Index).BackColor
  VBColor = Picture1(Index).BackColor
  HexColor = Label1.Caption
  RaiseEvent SelectColor
  Skjul
End Sub

Private Sub UserControl_Initialize()
    Dim iRed As Long, iGreen As Long, iBlue As Long
    Dim i As Long, iColor As Long
    Dim iIndex As Long
    
    'Position color box 18 per row
    For i = 1 To 215
        iIndex = Picture1.Count
        Load Picture1(iIndex)
        Picture1(iIndex).Left = Picture1(0).Left + 150 * (i Mod 18)
        Picture1(iIndex).Top = Picture1(0).Top + 150 * (i \ 18)
        Picture1(iIndex).Visible = True
    Next i
    
    i = 0
    For iRed = 0 To 255 Step 51
        For iGreen = 0 To 255 Step 51
            For iBlue = 0 To 255 Step 51
                iColor = iBlue * 256 * 256 + iGreen * 256 + iRed
                Picture1(i).BackColor = iColor
                Picture1(i).Tag = "#" & fillZero(Hex(iRed)) & fillZero(Hex(iGreen)) & fillZero(Hex(iBlue))
                i = i + 1
            Next iBlue
        Next iGreen
    Next iRed
    
    ilastIndex = -1
    If initialColor > 0 Then
        iRed = (initialColor \ 256) \ 256
        iGreen = (initialColor - iRed * 256 * 256) \ 256
        iBlue = initialColor - iRed * 256 * 256 - iGreen * 256
        If (iRed Mod 51) <> 0 Or (iGreen Mod 51) <> 0 Or (iBlue Mod 51) <> 0 Then
            Picture2.BackColor = initialColor
            Label1.Caption = "#" & fillZero(Hex(iRed)) & fillZero(Hex(iGreen)) & fillZero(Hex(iBlue))
        Else
            Picture1_Click (iRed \ 51) * 6 * 6 + (iGreen \ 51) * 6 + (iBlue \ 51)
        End If
    Else
        Picture1_Click 0
    End If
    
    Skjul
End Sub

Private Function fillZero(ByVal iColor As String) As String
  fillZero = iColor
  If Len(iColor) = 1 Then fillZero = "0" & iColor
End Function

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Shape1.Left = Picture1(Index).Left - 1 * Screen.TwipsPerPixelX
  Shape1.Top = Picture1(Index).Top - 1 * Screen.TwipsPerPixelY
  Shape1.Height = Picture1(Index).Height + 1 * Screen.TwipsPerPixelY
  Shape1.Width = Picture1(Index).Width + 1 * Screen.TwipsPerPixelX

  Shape2.Left = Picture1(Index).Left - 1 * Screen.TwipsPerPixelX
  Shape2.Top = Picture1(Index).Top - 1 * Screen.TwipsPerPixelY
  Shape2.Height = Picture1(Index).Height + 1 * Screen.TwipsPerPixelY
  Shape2.Width = Picture1(Index).Width + 1 * Screen.TwipsPerPixelX

  Picture2.BackColor = Picture1(Index).BackColor
  Label1.Caption = Picture1(Index).Tag
  ilastIndex = Index
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = 2715
  Shape3.Width = UserControl.Width
  Shape3.Height = ScaleHeight - 315
End Sub


Public Property Get SelColor() As Variant
  SelColor = lblFarver.BackColor
  RaiseEvent SelectColor
End Property

Public Property Let SelColor(ByVal vNewValue As Variant)
  lblFarver.BackColor = vNewValue
  RaiseEvent SelectColor
End Property
