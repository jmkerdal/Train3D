VERSION 5.00
Object = "{EC2CC72E-13BA-11D5-BB31-400001686160}#1.0#0"; "SELECTCOLOR.OCX"
Begin VB.Form Shape 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a new shape (x100)"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "Shape.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Cursor"
      Height          =   495
      Left            =   3480
      TabIndex        =   39
      Top             =   0
      Width           =   2535
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LY 
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         Height          =   255
         Left            =   1680
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LX 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parameters"
      Height          =   2295
      Left            =   3480
      TabIndex        =   17
      Top             =   3720
      Width           =   2535
      Begin VB.CommandButton Reset 
         Caption         =   "Reset"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox CView 
         Caption         =   "Close side"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox CView 
         Caption         =   "Out-Sides"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox CView 
         Caption         =   "In-Sides"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton CLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton CreateShape 
         Caption         =   "Create Shape"
         Height          =   375
         Left            =   1320
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.HScrollBar SFace 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   360
         Min             =   1
         TabIndex        =   20
         Top             =   240
         Value           =   3
         Width           =   1455
      End
      Begin VB.HScrollBar SAngle 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   360
         TabIndex        =   18
         Top             =   600
         Value           =   360
         Width           =   1455
      End
      Begin VB.Label LFace 
         Caption         =   "N Faces"
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LAngle 
         Alignment       =   2  'Center
         Caption         =   "360°"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Curves"
      Height          =   3255
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   2535
      Begin VB.CommandButton CPlus 
         Caption         =   ">"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   38
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton CPlus 
         Caption         =   ">"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   37
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton CPlus 
         Caption         =   ">"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   36
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton CPlus 
         Caption         =   ">"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   35
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton CPlus 
         Caption         =   ">"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   34
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton CPlus 
         Caption         =   ">"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton CMinus 
         Caption         =   "<"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   32
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton CMinus 
         Caption         =   "<"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   31
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton CMinus 
         Caption         =   "<"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   30
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton CMinus 
         Caption         =   "<"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   29
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton CMinus 
         Caption         =   "<"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   28
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton CMinus 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox TRay 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TVertex 
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Text            =   "1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TdY 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TdX 
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox TStart 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   12
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar SSegment 
         Height          =   255
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   5
         Top             =   240
         Value           =   1
         Width           =   1695
      End
      Begin VB.TextBox TStart 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin SelectColor.UserColor FaceColor 
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1720
      End
      Begin VB.Label LSegment 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Radius"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Vertex"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "End Y"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "End X"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Start Y"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Start X"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Label Axe_X 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Axe_Y 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Origin 
      BackStyle       =   0  'Transparent
      Caption         =   "0,0"
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   255
   End
End
Attribute VB_Name = "Shape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim SelectX1%
Dim SelectY1%
Dim SelectX2%
Dim SelectY2%
Dim SelectN%
Dim Apply As D3DVECTOR
Dim Normal As D3DVECTOR

'
' ********************
' Load a profile shape
' ********************
'
Private Sub CLoad_Click()
    Dim File$, i%, a!, f%
    File$ = Open_Box$("Load shape", "", "All files *.shape|*.shape", BOX_LOAD, Wall3D.BoiteMur3D)
    If Exist(File$) = False Then Exit Sub
    f% = FreeFile()
    Open File$ For Input As #f%
        For i% = 1 To NbSegment%
            Input #f%, a!: Segment(i%).X1 = a! + StartX%
            Input #f%, a!: Segment(i%).Y1 = a! + StartY%
            Input #f%, a!: Segment(i%).X2 = a! + StartX%
            Input #f%, a!: Segment(i%).Y2 = a! + StartY%
            Input #f%, a!: Segment(i%).Vertex = a!
            Input #f%, a!: Segment(i%).Radius = a!
            Input #f%, a!: Segment(i%).Red = a!
            Input #f%, a!: Segment(i%).Green = a!
            Input #f%, a!: Segment(i%).Blue = a!
            Input #f%, a!: Segment(i%).Alpha = a!
        Next i%
    Close #f%
    Call UpdateValue
    Call DrawShape
End Sub

'
' **********************
' -1 for each parameters
' **********************
'
Private Sub CMinus_Click(Index As Integer)
    Select Case Index
    Case 0
        Segment(SSegment).X1 = Segment(SSegment).X1 - 1
    Case 1
        Segment(SSegment).Y1 = Segment(SSegment).Y1 - 1
    Case 2
        Segment(SSegment).X2 = Segment(SSegment).X2 - 1
    Case 3
        Segment(SSegment).Y2 = Segment(SSegment).Y2 - 1
    Case 4
        Segment(SSegment).Vertex = Segment(SSegment).Vertex - 1
    Case 5
        Segment(SSegment).Radius = Segment(SSegment).Radius - 1
    End Select
    Call UpdateValue
    Call DrawShape
End Sub

'
' **********************
' -1 for each parameters
' **********************
'
Private Sub CPlus_Click(Index As Integer)
    Select Case Index
    Case 0
        Segment(SSegment).X1 = Segment(SSegment).X1 + 1
    Case 1
        Segment(SSegment).Y1 = Segment(SSegment).Y1 + 1
    Case 2
        Segment(SSegment).X2 = Segment(SSegment).X2 + 1
    Case 3
        Segment(SSegment).Y2 = Segment(SSegment).Y2 + 1
    Case 4
        Segment(SSegment).Vertex = Segment(SSegment).Vertex + 1
    Case 5
        Segment(SSegment).Radius = Segment(SSegment).Radius + 1
    End Select
    Call UpdateValue
    Call DrawShape
End Sub

'
' ******************
' Create a new Shape
' ******************
'
Private Sub CreateShape_Click()
    Dim n%, j%
    Dim X1!, Y1!, X2!, Y2!
    Dim pp%
    For n% = 1 To 100
        If Segment(n%).X1 <> Segment(n%).X2 Or Segment(n%).Y1 <> Segment(n%).Y2 Then
            For j% = 1 To Segment(n%).Vertex
                ReDim Preserve ListPoint(pp%) As TypeShape
                ListPoint(pp%).X1! = (Segment(n%).GetX(j% - 1) - StartX%) / 100
                ListPoint(pp%).Y1! = (Segment(n%).GetY(j% - 1) - StartY%) / -100
                ListPoint(pp%).X2! = (Segment(n%).GetX(j%) - StartX%) / 100
                ListPoint(pp%).Y2! = (Segment(n%).GetY(j%) - StartY%) / -100
                ListPoint(pp%).Color.Red = Segment(n%).Red
                ListPoint(pp%).Color.Green = Segment(n%).Green
                ListPoint(pp%).Color.Blue = Segment(n%).Blue
                ListPoint(pp%).Color.flags = Segment(n%).Alpha
                pp% = pp% + 1
            Next j%
        End If
    Next n%
    If CView(0) = vbChecked Then
        Call dxShape.Add_Shape(Wall3D.No_Mur.Value, ListPoint(), SAngle.Value, SFace.Value, False, CView(2).Value)
    End If
    If CView(1) = vbChecked Then
        Call dxShape.Add_Shape(Wall3D.No_Mur.Value, ListPoint(), SAngle.Value, SFace.Value, True, CView(2).Value)
    End If
    Call Wall3D.Update
    Unload Me
End Sub

'
' ********************
' Save a profile shape
' ********************
'
Private Sub CSave_Click()
    Dim File$, i%, f%
    File$ = Open_Box$("Save shape", "", "All files *.shape|*.shape", BOX_SAVE, Wall3D.BoiteMur3D)
    f% = FreeFile()
    Open File$ For Output As #f%
        For i% = 1 To NbSegment%
            Call Save_Number(f%, Segment(i%).X1 - StartX%, False)
            Call Save_Number(f%, Segment(i%).Y1 - StartY%, False)
            Call Save_Number(f%, Segment(i%).X2 - StartX%, False)
            Call Save_Number(f%, Segment(i%).Y2 - StartY%, False)
            Call Save_Number(f%, Segment(i%).Vertex, False)
            Call Save_Number(f%, Segment(i%).Radius, False)
            Call Save_Number(f%, Segment(i%).Red, False)
            Call Save_Number(f%, Segment(i%).Green, False)
            Call Save_Number(f%, Segment(i%).Blue, False)
            Call Save_Number(f%, Segment(i%).Alpha, True)
        Next i%
    Close #f%
End Sub

'
' *********
' Enable OK
' *********
'
Private Sub CView_Click(Index As Integer)
    If CView(0) = vbUnchecked And CView(1) = vbUnchecked Then
        CreateShape.Enabled = False
    Else
        CreateShape.Enabled = True
    End If
End Sub

Private Sub FaceColor_ChangeAlpha()
    Segment(SSegment).Alpha = FaceColor.Alpha
    Call DrawShape
End Sub

Private Sub FaceColor_ChangeBlue()
    Segment(SSegment).Blue = FaceColor.Blue
    Call DrawShape
End Sub

Private Sub FaceColor_ChangeGreen()
    Segment(SSegment).Green = FaceColor.Green
    Call DrawShape
End Sub

Private Sub FaceColor_ChangeRed()
    Segment(SSegment).Red = FaceColor.Red
    Call DrawShape
End Sub

'
' **************
' Initialisation
' **************
'
Private Sub Form_Load()
    Call UpdateValue
    Call DrawShape
End Sub

'
' *****************
' Add a new Segment
' *****************
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i%, n%
    If X <= SelectX1% + 2 And X >= SelectX1% - 2 And Y <= SelectY1% + 2 And Y >= SelectY1% - 2 Then
        If SelectN% <> 1 Then
            SelectN% = 1
            Call DrawSelect
            Exit Sub
        End If
    End If
    If X <= SelectX2% + 2 And X >= SelectX2% - 2 And Y <= SelectY2% + 2 And Y >= SelectY2% - 2 Then
        If SelectN% <> 2 Then
            SelectN% = 2
            Call DrawSelect
            Exit Sub
        End If
    End If
    If SelectN% <> 0 Then
        If SelectN% = 1 Then
            Segment(SSegment).X1 = X
            Segment(SSegment).Y1 = Y
        End If
        If SelectN% = 2 Then
            Segment(SSegment).X2 = X
            Segment(SSegment).Y2 = Y
        End If
        SelectN% = 0
        Call UpdateValue
    Else
        For i% = 1 To NbSegment%
            If n% = 0 Then
                If Segment(i%).X1 <> Segment(i%).X2 Or Segment(i%).Y1 <> Segment(i%).Y2 Then
                    If Segment(i%).TestHit(X, Y) = True Then n% = i%
                End If
            End If
        Next i%
        If n% <> 0 Then
            SSegment.Value = n%
            SelectN% = 0
        End If
    End If
    Call DrawShape
End Sub

'
' ******************************
' Show X Y position in real time
' ******************************
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LX = X - StartX
    LY = Y - StartY
End Sub

'
' ****************
' Reset all values
' ****************
'
Private Sub Reset_Click()
    If MsgBox("Reset all segments", vbQuestion + vbYesNo + vbDefaultButton1, "Shape") = vbNo Then Exit Sub
    Call Init.Shape_Init
    Call UpdateValue
    SelectN% = 0
    Call DrawShape
End Sub

Private Sub SAngle_Change()
    LAngle = SAngle & "°"
End Sub

Private Sub SFace_Change()
    NbFaceSegment% = SFace
    LFace = SFace
End Sub

Private Sub SSegment_Change()
    Call UpdateValue
    SelectN% = 0
    Call DrawShape
End Sub

Private Sub TdX_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chr$(KeyCode) <> vbCr Then Exit Sub
    Segment(SSegment).X2 = Val(TdX) + StartX%
    Call DrawShape
End Sub

Private Sub TdY_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chr$(KeyCode) <> vbCr Then Exit Sub
    Segment(SSegment).Y2 = Val(TdY) + StartY%
    Call DrawShape
End Sub

'
' ***********************************
' Draw collection of curves to screen
' ***********************************
'
Public Sub DrawShape()
    Dim i%
    Cls
    Me.Line (StartX%, 0)-(StartX%, StartY% * 2), vbBlack
    Me.Line (0, StartY%)-(200 + StartX%, StartY%), vbBlack
    For i% = 1 To NbSegment%
        Call Segment(i%).Draw(Me)
    Next i%
    If Segment(SSegment).X1 = Segment(SSegment).X2 And Segment(SSegment).Y1 = Segment(SSegment).Y2 Then Exit Sub
    Call DrawSelect
End Sub

Private Sub TRay_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chr$(KeyCode) <> vbCr Then Exit Sub
    Segment(SSegment).Radius = Val(TRay)
    Call DrawShape
End Sub

Private Sub TStart_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Chr$(KeyCode) <> vbCr Then Exit Sub
    If Index = 0 Then
        Segment(SSegment).X1 = Val(TStart(Index)) + StartX%
    Else
        Segment(SSegment).Y1 = Val(TStart(Index)) + StartY%
    End If
    Call DrawShape
End Sub

Private Sub TVertex_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chr$(KeyCode) <> vbCr Then Exit Sub
    Segment(SSegment).Vertex = Val(TVertex)
    Call DrawShape
End Sub

'
' **********************************
' Put mark to screen on curve select
' **********************************
'
Public Sub DrawSelect()
    Dim i%
    SelectX1% = Segment(SSegment).GetX(0)
    SelectY1% = Segment(SSegment).GetY(0)
    If SelectN% = 1 Then
        Me.Line (SelectX1% - 2, SelectY1% - 2)-(SelectX1% + 2, SelectY1% + 2), vbBlue, B
    Else
        Me.Line (SelectX1% - 2, SelectY1% - 2)-(SelectX1% + 2, SelectY1% + 2), vbRed, B
    End If
    SelectX2% = Segment(SSegment).GetX(Segment(SSegment).Vertex)
    SelectY2% = Segment(SSegment).GetY(Segment(SSegment).Vertex)
    If SelectN% = 2 Then
        Me.Line (SelectX2% - 2, SelectY2% - 2)-(SelectX2% + 2, SelectY2% + 2), vbBlue, B
    Else
        Me.Line (SelectX2% - 2, SelectY2% - 2)-(SelectX2% + 2, SelectY2% + 2), vbRed, B
    End If
    For i% = 1 To Segment(SSegment).Vertex
        Call Segment(SSegment).GetNormal(i%, Apply, Normal)
        Me.Line (Apply.X, Apply.Y)-(Apply.X + Normal.X * 5, Apply.Y + Normal.Y * 5), vbRed
    Next i%
End Sub

'
' *****************************
' Update new value of the curve
' *****************************
'
Public Sub UpdateValue()
    TStart(0) = Segment(SSegment).X1 - StartX%
    TStart(1) = Segment(SSegment).Y1 - StartY%
    TdX = Segment(SSegment).X2 - StartX%
    TdY = Segment(SSegment).Y2 - StartY%
    TVertex = Segment(SSegment).Vertex
    TRay = Segment(SSegment).Radius
    Dim TheRed%, TheBlue%, TheGreen%, TheAlpha%
    TheRed% = Segment(SSegment).Red
    TheGreen% = Segment(SSegment).Green
    TheBlue% = Segment(SSegment).Blue
    TheAlpha% = Segment(SSegment).Alpha
    FaceColor.Red = TheRed%
    FaceColor.Green = TheGreen%
    FaceColor.Blue = TheBlue%
    FaceColor.Alpha = TheAlpha%
    LSegment.Caption = SSegment.Value
    SFace = NbFaceSegment%
    LFace = SFace
End Sub

