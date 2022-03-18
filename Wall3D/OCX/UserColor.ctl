VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl UserColor 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ScaleHeight     =   1485
   ScaleWidth      =   2310
   ToolboxBitmap   =   "UserColor.ctx":0000
   Begin VB.HScrollBar SColor 
      Height          =   220
      Index           =   3
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin MSComDlg.CommonDialog ColorBox 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox ColorPicture 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.HScrollBar SColor 
      Height          =   220
      Index           =   2
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.HScrollBar SColor 
      Height          =   220
      Index           =   1
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.HScrollBar SColor 
      Height          =   220
      Index           =   0
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label LColor 
      Caption         =   "A"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   135
   End
   Begin VB.Label LColor 
      Caption         =   "B"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   135
   End
   Begin VB.Label LColor 
      Caption         =   "G"
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   135
   End
   Begin VB.Label LColor 
      Caption         =   "R"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   135
   End
   Begin VB.Label LColor 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.Label LColor 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.Label LColor 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.Label LColor 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "UserColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' **********************************
' Select user color
' © Kerdal Jean-Michel 1998
' Jean-Michel.Kerdal@gazdefrance.com
' http://members.aol.com/DemiCitron/
' **********************************
'
Option Explicit
' Raise the event when change some value
Event ChangeRed()
Event ChangeGreen()
Event ChangeBlue()
Event ChangeAlpha()

'
' *******************
' Link to alpha color
' *******************
'
Public Function Alpha() As Object
    Set Alpha = SColor(3)
End Function

'
' ******************
' Link to blue color
' ******************
'
Public Function Blue() As Object
    Set Blue = SColor(2)
End Function

'
' *******************
' Link to green color
' *******************
'
Public Function Green() As Object
    Set Green = SColor(1)
End Function

'
' ************************
' Chose a color in palette
' ************************
'
Private Sub ColorPicture_DblClick()
    Dim ReturnColor&
    Call ColorBox.ShowColor
    If ColorBox.CancelError = False Then
        ReturnColor& = ColorBox.Color
        SColor(2) = ReturnColor& \ 2 ^ 16
        ReturnColor& = ReturnColor& - SColor(2) * 2 ^ 16
        SColor(1) = ReturnColor& \ 2 ^ 8
        SColor(0) = ReturnColor& Mod 2 ^ 8
    End If
End Sub

'
' ****************************
' Update color value to screen
' ****************************
'
Private Sub SColor_Change(Index As Integer)
    LColor(Index).Caption = SColor(Index).Value
    ColorPicture.BackColor = SColor(0).Value + SColor(1).Value * 2 ^ 8 + SColor(2).Value * 2 ^ 16
    Select Case Index
    Case 0
        RaiseEvent ChangeRed
    Case 1
        RaiseEvent ChangeGreen
    Case 2
        RaiseEvent ChangeBlue
    Case 3
        RaiseEvent ChangeAlpha
    End Select
End Sub

'
' *****************
' Link to red color
' *****************
'
Public Function Red() As Object
    Set Red = SColor(0)
End Function

'
' *********************
' Hide/Show Alpha color
' *********************
'
Public Property Let ShowAlpha(ByVal vNewValue As Boolean)
    If vNewValue = True Then
        SColor(3).Visible = True
        LColor(3).Visible = True
        LColor(7).Visible = True
    End If
    If vNewValue = False Then
        SColor(3).Visible = False
        LColor(3).Visible = False
        LColor(7).Visible = False
    End If
End Property

