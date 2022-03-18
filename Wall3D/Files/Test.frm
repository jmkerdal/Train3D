VERSION 5.00
Object = "{27D094C0-7329-11D2-92D1-BA7D894B6D3D}#6.0#0"; "SelectColor.ocx"
Begin VB.Form Test 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   Begin VB.VScrollBar SAngle2 
      Height          =   1455
      LargeChange     =   10
      Left            =   5400
      Max             =   360
      Min             =   -100
      TabIndex        =   10
      Top             =   1440
      Width           =   255
   End
   Begin VB.VScrollBar SAngle1 
      Height          =   1455
      LargeChange     =   10
      Left            =   4920
      Max             =   460
      Min             =   -100
      TabIndex        =   9
      Top             =   1440
      Width           =   255
   End
   Begin VB.VScrollBar Ray 
      Height          =   975
      LargeChange     =   10
      Left            =   3120
      Max             =   100
      Min             =   -100
      TabIndex        =   4
      Top             =   3240
      Width           =   255
   End
   Begin VB.VScrollBar nVertex 
      Height          =   975
      LargeChange     =   10
      Left            =   2760
      Max             =   100
      Min             =   1
      TabIndex        =   3
      Top             =   3240
      Value           =   1
      Width           =   255
   End
   Begin VB.HScrollBar dY 
      Height          =   255
      LargeChange     =   10
      Left            =   600
      Max             =   100
      Min             =   -100
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.HScrollBar dX 
      Height          =   255
      LargeChange     =   10
      Left            =   600
      Max             =   100
      Min             =   -100
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin SelectColor.UserColor UserColor1 
      Height          =   975
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1720
   End
   Begin VB.Label LAngle2 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label LAngle1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   3000
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   8
      X2              =   200
      Y1              =   100
      Y2              =   100
   End
   Begin VB.Line Line1 
      X1              =   100
      X2              =   100
      Y1              =   8
      Y2              =   200
   End
   Begin VB.Label LRay 
      Caption         =   "Radius"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label LVertex 
      Caption         =   "Nb Vertex"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label LY 
      Caption         =   "Y"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label LX 
      Caption         =   "X"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   375
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' *******************************************
' This project is used to test OCX components
' © Kerdal Jean-Michel 1998-99
' Jean-Michel.Kerdal@gazdefrance.com
' http://members.aol.com/DemiCitron/
' *******************************************
'
Option Explicit
Dim Curve1 As New Curve
Dim Arc1 As New Arc

Private Sub dX_Change()
    Curve1.X2 = 100 + dX
    LX = dX
    Call DrawCurve
End Sub

Private Sub dY_Change()
    Curve1.Y2 = 100 + dY
    LY = dY
    Call DrawCurve
End Sub

Private Sub Form_Load()
    Show
    UserColor1.Red = &HFF&
    UserColor1.Green = 0
    UserColor1.Blue = 0
    Curve1.X1 = 100
    Curve1.Y1 = 100
    dX = 50
    dY = 50
    nVertex = 1
    Ray = 10
    Call SplineCurve.main(0, "")
End Sub

'
' ***********************************************
' Test if I hit the curve with mouse when I click
' ***********************************************
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "Click "; Curve1.TestHit(X, Y)
End Sub

Private Sub nVertex_Change()
    Curve1.Vertex = nVertex
    LVertex = nVertex
    Call DrawCurve
End Sub

Private Sub Ray_Change()
    Curve1.Radius = Ray
    LRay = Ray
    Call DrawCurve
End Sub

'
' ************************
' Draw the curve on screen
' ************************
'
Public Sub DrawCurve()
'Debug.Print "*"; Curve21.GetX(0); " "; Curve21.GetY(0)
    'Curve1.Top = 100 - Curve1.GetY(0)
    'Curve1.Left = 100 - Curve1.GetX(0)
    'Curve1.Height = Curve1.GetHeight
    'Curve1.Width = Curve1.GetWidth
    Cls
    Curve1.Red = UserColor1.Red
    Curve1.Green = UserColor1.Green
    Curve1.Blue = UserColor1.Blue
    Curve1.Alpha = UserColor1.Alpha
    Call Curve1.Draw(Me)
    '
    Arc1.X = 100
    Arc1.Y = 100
    Arc1.Red = UserColor1.Red
    Arc1.Green = UserColor1.Green
    Arc1.Blue = UserColor1.Blue
    Arc1.Alpha = UserColor1.Alpha
    Arc1.Angle1 = SAngle1
    Arc1.Angle2 = SAngle2
    Arc1.Vertex = 8
    Arc1.Radius = Ray
    Call Arc1.Draw(Me)
End Sub

Private Sub SAngle1_Change()
    LAngle1 = SAngle1
    Call DrawCurve
End Sub

Private Sub SAngle2_Change()
    LAngle2 = SAngle2
    Call DrawCurve
End Sub

