VERSION 5.00
Begin VB.Form Faces 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Faces (x100)"
   ClientHeight    =   4470
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6915
   Icon            =   "Faces.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Faces"
      Height          =   4335
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.TextBox TY 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TY 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TY 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TY 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TX 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TX 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   3720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TX 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TX 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox CView 
         Caption         =   "Inverse faces"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox CView 
         Caption         =   "Close sides"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox CView 
         Caption         =   "3D"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox CView 
         Caption         =   "In-Sides"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox CView 
         Caption         =   "Out-Sides"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton CSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton CLoad 
         Caption         =   "Load"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton CreateFace 
         Caption         =   "Create Face"
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label LXY 
         Caption         =   "Y4"
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   20
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label LXY 
         Caption         =   "Y3"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   19
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label LXY 
         Caption         =   "Y2"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   18
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label LXY 
         Caption         =   "Y1"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   17
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label LXY 
         Caption         =   "X4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label LXY 
         Caption         =   "X3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label LXY 
         Caption         =   "X2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label LXY 
         Caption         =   "X1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label LY 
         Caption         =   "Label1"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Y:"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label LX 
         Caption         =   "Label1"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Menu FACES_Create 
      Caption         =   "Create"
      Begin VB.Menu FACES_Triangle 
         Caption         =   "Create triangle"
      End
      Begin VB.Menu FACES_Rectangle 
         Caption         =   "Create rectangle"
      End
      Begin VB.Menu FACES_Curve 
         Caption         =   "Create curve"
      End
      Begin VB.Menu FACES_Arc 
         Caption         =   "Create arc"
      End
   End
   Begin VB.Menu FACES_Curves 
      Caption         =   "Curve"
      Begin VB.Menu FACES_Vertex 
         Caption         =   "Vertex"
      End
      Begin VB.Menu FACES_Radius 
         Caption         =   "Radius"
      End
      Begin VB.Menu FACES_Color 
         Caption         =   "Color"
      End
   End
   Begin VB.Menu FACES_Arcs 
      Caption         =   "Arc"
      Begin VB.Menu FACES_AVertex 
         Caption         =   "Vertex"
      End
      Begin VB.Menu FACES_ARadius 
         Caption         =   "Radius"
      End
      Begin VB.Menu FACES_End 
         Caption         =   "End"
      End
      Begin VB.Menu FACES_Start 
         Caption         =   "Start"
      End
      Begin VB.Menu FACES_AColor 
         Caption         =   "Color"
      End
   End
End
Attribute VB_Name = "Faces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Enum EnumForm
    FormTriangle = 1
    FormRectangle
    FormCurve
    FormArc
End Enum
Private Type TypeForm
    TheType As EnumForm
    Parameter(3) As D3DVECTOR
    OneCurve As New Curve
    OneArc As New Arc
End Type
Dim ListForm() As TypeForm
Dim StartX%
Dim StartY%
Dim Selected%
Dim SelectN%

'
' *********************
' Load faces definition
' *********************
'
Private Sub CLoad_Click()
    Dim File$, i%, j%, a!, f%
    File$ = Open_Box$("Load face", "", "All files *.face|*.face", BOX_LOAD, Wall3D.BoiteMur3D)
    If Exist(File$) = False Then Exit Sub
    f% = FreeFile()
    Open File$ For Input As #f%
        Input #f%, a!
        ReDim ListForm(a!) As TypeForm
        For i% = 1 To UBound(ListForm())
            Input #f%, a!: ListForm(i%).TheType = a!
            If ListForm(i%).TheType = FormCurve Then
                Input #f%, a!:  ListForm(i%).OneCurve.X1 = a! + StartX%
                Input #f%, a!:  ListForm(i%).OneCurve.Y1 = a! + StartY%
                Input #f%, a!:  ListForm(i%).OneCurve.X2 = a! + StartX%
                Input #f%, a!:  ListForm(i%).OneCurve.Y2 = a! + StartY%
                Input #f%, a!:  ListForm(i%).OneCurve.Radius = a!
                Input #f%, a!:  ListForm(i%).OneCurve.Vertex = a!
                Input #f%, a!:  ListForm(i%).OneCurve.Red = a!
                Input #f%, a!:  ListForm(i%).OneCurve.Green = a!
                Input #f%, a!:  ListForm(i%).OneCurve.Blue = a!
                Input #f%, a!:  ListForm(i%).OneCurve.Alpha = a!
            End If
            If ListForm(i%).TheType = FormArc Then
                Input #f%, a!:  ListForm(i%).OneArc.X = a! + StartX%
                Input #f%, a!:  ListForm(i%).OneArc.Y = a! + StartY%
                Input #f%, a!:  ListForm(i%).OneArc.Radius = a!
                Input #f%, a!:  ListForm(i%).OneArc.Vertex = a!
                Input #f%, a!:  ListForm(i%).OneArc.Angle1 = a!
                Input #f%, a!:  ListForm(i%).OneArc.Angle2 = a!
                Input #f%, a!:  ListForm(i%).OneArc.Red = a!
                Input #f%, a!:  ListForm(i%).OneArc.Green = a!
                Input #f%, a!:  ListForm(i%).OneArc.Blue = a!
                Input #f%, a!:  ListForm(i%).OneArc.Alpha = a!
            End If
            For j% = 0 To 3
                Input #f%, a!:  ListForm(i%).Parameter(j%).X = a!
                Input #f%, a!: ListForm(i%).Parameter(j%).Y = a!
            Next j%
        Next i%
    Close #f%
    Selected% = 0
    SelectN% = -1
    Call DrawForm
End Sub

'
' *********************
' Create a complex face
' *********************
'
Private Sub CreateFace_Click()
    Dim i%, j%, t%, Volume As Boolean
    Dim Inversion1 As Boolean
    Dim Inversion2 As Boolean
    Dim Color As PALETTEENTRY
    Dim Triangle(2) As D3DVECTOR
    t% = UBound(ListForm())
    If t% = 0 Then Exit Sub
    If CView(0).Value = vbChecked And CView(1).Value = vbChecked And CView(2).Value = vbChecked Then
        Volume = True
    Else
        Volume = False
    End If
    If CView(4).Value = vbChecked Then
        Inversion1 = True
        Inversion2 = False
    Else
        Inversion1 = False
        Inversion2 = True
    End If
    For i% = 1 To t%
        ' ***** Scale /100 and change coordinates
        For j% = 0 To 3
            ListForm(i%).Parameter(j%).X = ListForm(i%).Parameter(j%).X / 100
            ListForm(i%).Parameter(j%).Y = ListForm(i%).Parameter(j%).Y / -100
            If Volume = True Then
                ListForm(i%).Parameter(j%).z = -0.5
            End If
        Next j%
        Select Case ListForm(i%).TheType
        Case EnumForm.FormTriangle
            Color.Red = 255
            Color.Green = 255
            Color.Blue = 255
            Color.flags = 255
            If CView(0).Value = vbChecked Then
                Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 3, ListForm(i%).Parameter(), Inversion1)
            End If
            If CView(1).Value = vbChecked Then
                If Volume = True Then
                    For j% = 0 To 2
                        ListForm(i%).Parameter(j%).z = 0.5
                    Next j%
                End If
                Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 3, ListForm(i%).Parameter(), Inversion2)
            End If
            If Volume = True Then
                Call CreateSide(Color, ListForm(i%).Parameter(0), ListForm(i%).Parameter(1), Inversion1)
                Call CreateSide(Color, ListForm(i%).Parameter(1), ListForm(i%).Parameter(2), Inversion1)
                Call CreateSide(Color, ListForm(i%).Parameter(2), ListForm(i%).Parameter(0), Inversion1)
            End If
        Case EnumForm.FormRectangle
            Color.Red = 255
            Color.Green = 255
            Color.Blue = 255
            Color.flags = 255
            If CView(0).Value = vbChecked Then
                Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 4, ListForm(i%).Parameter(), Inversion1)
            End If
            If CView(1).Value = vbChecked Then
                If Volume = True Then
                    For j% = 0 To 3
                        ListForm(i%).Parameter(j%).z = 0.5
                    Next j%
                End If
                Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 4, ListForm(i%).Parameter(), Inversion2)
            End If
            If Volume = True Then
                Call CreateSide(Color, ListForm(i%).Parameter(0), ListForm(i%).Parameter(1), Inversion1)
                Call CreateSide(Color, ListForm(i%).Parameter(1), ListForm(i%).Parameter(2), Inversion1)
                Call CreateSide(Color, ListForm(i%).Parameter(2), ListForm(i%).Parameter(3), Inversion1)
                Call CreateSide(Color, ListForm(i%).Parameter(3), ListForm(i%).Parameter(0), Inversion1)
            End If
        Case EnumForm.FormCurve
            Color.Red = ListForm(i%).OneCurve.Red
            Color.Green = ListForm(i%).OneCurve.Green
            Color.Blue = ListForm(i%).OneCurve.Blue
            Color.flags = ListForm(i%).OneCurve.Alpha
            Triangle(0) = ListForm(i%).Parameter(0)
            For j% = 1 To ListForm(i%).OneCurve.Vertex
                Triangle(1).X = (ListForm(i%).OneCurve.GetX(j% - 1) - StartX%) / 100
                Triangle(1).Y = (ListForm(i%).OneCurve.GetY(j% - 1) - StartY%) / -100
                Triangle(2).X = (ListForm(i%).OneCurve.GetX(j%) - StartX%) / 100
                Triangle(2).Y = (ListForm(i%).OneCurve.GetY(j%) - StartY%) / -100
                If CView(0).Value = vbChecked Then
                    If Volume = True Then
                        Triangle(0).z = -0.5
                        Triangle(1).z = -0.5
                        Triangle(2).z = -0.5
                    End If
                    Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 3, Triangle(), Inversion1)
                End If
                If CView(1).Value = vbChecked Then
                    If Volume = True Then
                        Triangle(0).z = 0.5
                        Triangle(1).z = 0.5
                        Triangle(2).z = 0.5
                    End If
                    Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 3, Triangle(), Inversion2)
                End If
                If Volume = True Then
                    If CView(3).Value = vbChecked And j% = 1 Then
                        Call CreateSide(Color, Triangle(0), Triangle(1), Inversion1)
                    End If
                    Call CreateSide(Color, Triangle(1), Triangle(2), Inversion1)
                    If CView(3).Value = vbChecked And j% = ListForm(i%).OneCurve.Vertex Then
                        Call CreateSide(Color, Triangle(2), Triangle(0), Inversion1)
                    End If
                End If
            Next j%
        Case EnumForm.FormArc
            Color.Red = ListForm(i%).OneArc.Red
            Color.Green = ListForm(i%).OneArc.Green
            Color.Blue = ListForm(i%).OneArc.Blue
            Color.flags = ListForm(i%).OneArc.Alpha
            Triangle(0) = ListForm(i%).Parameter(0)
            For j% = 1 To ListForm(i%).OneArc.Vertex
                Triangle(1).X = (ListForm(i%).OneArc.GetX(j% - 1) - StartX%) / 100
                Triangle(1).Y = (ListForm(i%).OneArc.GetY(j% - 1) - StartY%) / -100
                Triangle(2).X = (ListForm(i%).OneArc.GetX(j%) - StartX%) / 100
                Triangle(2).Y = (ListForm(i%).OneArc.GetY(j%) - StartY%) / -100
                If CView(0).Value = vbChecked Then
                    If Volume = True Then
                        Triangle(0).z = -0.5
                        Triangle(1).z = -0.5
                        Triangle(2).z = -0.5
                    End If
                    Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 3, Triangle(), Inversion1)
                End If
                If CView(1).Value = vbChecked Then
                    If Volume = True Then
                        Triangle(0).z = 0.5
                        Triangle(1).z = 0.5
                        Triangle(2).z = 0.5
                    End If
                    Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 3, Triangle(), Inversion2)
                End If
                If Volume = True Then
                    If CView(3).Value = vbChecked And j% = 1 Then
                        Call CreateSide(Color, Triangle(0), Triangle(1), Inversion1)
                    End If
                    Call CreateSide(Color, Triangle(1), Triangle(2), Inversion1)
                    If CView(3).Value = vbChecked And j% = ListForm(i%).OneArc.Vertex Then
                        Call CreateSide(Color, Triangle(2), Triangle(0), Inversion1)
                    End If
                End If
            Next j%
        End Select
    Next i%
    Call Wall3D.Update
    Unload Me
End Sub

'
' *********************
' Save faces definition
' *********************
'
Private Sub CSave_Click()
    Dim File$, i%, j%, f%
    File$ = Open_Box$("Save face", "", "All files *.face|*.face", BOX_SAVE, Wall3D.BoiteMur3D)
    f% = FreeFile()
    Open File$ For Output As #f%
        Call Save_Number(f%, UBound(ListForm()), True)
        For i% = 1 To UBound(ListForm())
            Call Save_Number(f%, ListForm(i%).TheType, False)
            If ListForm(i%).TheType = FormCurve Then
                Call Save_Number(f%, ListForm(i%).OneCurve.X1 - StartX%, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Y1 - StartY%, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.X2 - StartX%, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Y2 - StartY%, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Radius, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Vertex, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Red, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Green, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Blue, False)
                Call Save_Number(f%, ListForm(i%).OneCurve.Alpha, True)
            End If
            If ListForm(i%).TheType = FormArc Then
                Call Save_Number(f%, ListForm(i%).OneArc.X - StartX%, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Y - StartY%, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Radius, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Vertex, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Angle1, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Angle2, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Red, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Green, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Blue, False)
                Call Save_Number(f%, ListForm(i%).OneArc.Alpha, False)
            End If
            For j% = 0 To 3
                Call Save_Number(f%, ListForm(i%).Parameter(j%).X, False)
                If j% = 3 Then
                    Call Save_Number(f%, ListForm(i%).Parameter(j%).Y, True)
                Else
                    Call Save_Number(f%, ListForm(i%).Parameter(j%).Y, False)
                End If
            Next j%
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
        CreateFace.Enabled = False
    Else
        CreateFace.Enabled = True
    End If
    If CView(0) = vbChecked And CView(1) = vbChecked Then
        CView(2).Enabled = True
        CView(3).Enabled = True
    Else
        CView(2).Enabled = False
        CView(3).Enabled = False
    End If
End Sub

'
' ********************
' Change the arc color
' ********************
'
Private Sub FACES_AColor_Click()
    Load ColorSelection
    ColorSelection.TheColor.Red = ListForm(Selected%).OneArc.Red
    ColorSelection.TheColor.Green = ListForm(Selected%).OneArc.Green
    ColorSelection.TheColor.Blue = ListForm(Selected%).OneArc.Blue
    ColorSelection.TheColor.Alpha = ListForm(Selected%).OneArc.Alpha
    ColorSelection.Show vbModal
    ListForm(Selected%).OneArc.Red = TheRed
    ListForm(Selected%).OneArc.Green = TheGreen
    ListForm(Selected%).OneArc.Blue = TheBlue
    ListForm(Selected%).OneArc.Alpha = TheAlpha
    Call DrawForm
End Sub

'
' ****************************
' Change the radius of the arc
' ****************************
'
Private Sub FACES_ARadius_Click()
    Dim r!
    r! = Val(InputBox("Arc radius", "Faces", ListForm(Selected%).OneArc.Radius))
    ListForm(Selected%).OneArc.Radius = r!
    Call DrawForm
End Sub

'
' ****************
' Create a new arc
' ****************
'
Private Sub FACES_Arc_Click()
    Dim n%
    n% = UBound(ListForm()) + 1
    ReDim Preserve ListForm(n%) As TypeForm
    ListForm(n%).TheType = EnumForm.FormArc
    ListForm(n%).OneArc.Vertex = 8
    ListForm(n%).OneArc.Radius = 20
    ListForm(n%).OneArc.Angle1 = 180
    ListForm(n%).OneArc.Angle2 = 0
    ListForm(n%).OneArc.Red = 255
    ListForm(n%).OneArc.Green = 255
    ListForm(n%).OneArc.Blue = 255
    ListForm(n%).OneArc.Alpha = 255
    ListForm(n%).OneArc.X = Faces.LX + StartX%
    ListForm(n%).OneArc.Y = Faces.LY + StartY%
    ListForm(n%).Parameter(0).X = Faces.LX - 5
    ListForm(n%).Parameter(0).Y = Faces.LY + 5
    Call DrawForm
End Sub

'
' ***************************
' Change the number of vertex
' ***************************
'
Private Sub FACES_AVertex_Click()
    Dim v%
    v% = Val(InputBox("Number of vertex", "Faces", ListForm(Selected%).OneArc.Vertex))
    If v% <= 0 Then Exit Sub
    ListForm(Selected%).OneArc.Vertex = v%
    Call DrawForm
End Sub

'
' **********************
' Change the curve color
' **********************
'
Private Sub FACES_Color_Click()
    Load ColorSelection
    ColorSelection.TheColor.Red = ListForm(Selected%).OneCurve.Red
    ColorSelection.TheColor.Green = ListForm(Selected%).OneCurve.Green
    ColorSelection.TheColor.Blue = ListForm(Selected%).OneCurve.Blue
    ColorSelection.TheColor.Alpha = ListForm(Selected%).OneCurve.Alpha
    ColorSelection.Show vbModal
    ListForm(Selected%).OneCurve.Red = TheRed
    ListForm(Selected%).OneCurve.Green = TheGreen
    ListForm(Selected%).OneCurve.Blue = TheBlue
    ListForm(Selected%).OneCurve.Alpha = TheAlpha
    Call DrawForm
End Sub

'
' ******************
' Create a new curve
' ******************
'
Private Sub FACES_Curve_Click()
    Dim n%
    n% = UBound(ListForm()) + 1
    ReDim Preserve ListForm(n%) As TypeForm
    ListForm(n%).TheType = EnumForm.FormCurve
    ListForm(n%).OneCurve.Vertex = 1
    ListForm(n%).OneCurve.Radius = 0
    ListForm(n%).OneCurve.Red = 255
    ListForm(n%).OneCurve.Green = 255
    ListForm(n%).OneCurve.Blue = 255
    ListForm(n%).OneCurve.Alpha = 255
    ListForm(n%).OneCurve.X1 = Faces.LX - 5 + StartX%
    ListForm(n%).OneCurve.Y1 = Faces.LY + StartY%
    ListForm(n%).OneCurve.X2 = Faces.LX + 5 + StartX%
    ListForm(n%).OneCurve.Y2 = Faces.LY + StartY%
    ListForm(n%).Parameter(0).X = Faces.LX
    ListForm(n%).Parameter(0).Y = Faces.LY + 5
    Call DrawForm
End Sub

'
' *********************
' Change the end of arc
' *********************
'
Private Sub FACES_End_Click()
    Dim v%
    v% = Val(InputBox("Arc end", "Faces", ListForm(Selected%).OneArc.Angle1))
    ListForm(Selected%).OneArc.Angle1 = v%
    Call DrawForm
End Sub

'
' ******************************
' Change the radius of the curve
' ******************************
'
Private Sub FACES_Radius_Click()
    Dim r!
    r! = Val(InputBox("Curve radius", "Faces", ListForm(Selected%).OneCurve.Radius))
    ListForm(Selected%).OneCurve.Radius = r!
    Call DrawForm
End Sub

'
' **********************
' Create a new rectangle
' **********************
'
Private Sub FACES_Rectangle_Click()
    Dim n%
    n% = UBound(ListForm()) + 1
    ReDim Preserve ListForm(n%) As TypeForm
    ListForm(n%).TheType = FormRectangle
    ListForm(n%).Parameter(0).X = Faces.LX - 5
    ListForm(n%).Parameter(0).Y = Faces.LY - 5
    ListForm(n%).Parameter(1).X = Faces.LX + 5
    ListForm(n%).Parameter(1).Y = Faces.LY - 5
    ListForm(n%).Parameter(2).X = Faces.LX + 5
    ListForm(n%).Parameter(2).Y = Faces.LY + 5
    ListForm(n%).Parameter(3).X = Faces.LX - 5
    ListForm(n%).Parameter(3).Y = Faces.LY + 5
    Call DrawForm
End Sub

'
' ***********************
' Change the start of arc
' ***********************
'
Private Sub FACES_Start_Click()
    Dim v%
    v% = Val(InputBox("Arc start", "Faces", ListForm(Selected%).OneArc.Angle2))
    ListForm(Selected%).OneArc.Angle2 = v%
    Call DrawForm
End Sub

'
' *********************
' Create a new triangle
' *********************
'
Private Sub FACES_Triangle_Click()
    Dim n%
    n% = UBound(ListForm()) + 1
    ReDim Preserve ListForm(n%) As TypeForm
    ListForm(n%).TheType = FormTriangle
    ListForm(n%).Parameter(0).X = LX
    ListForm(n%).Parameter(0).Y = LY - 5
    ListForm(n%).Parameter(1).X = LX + 5
    ListForm(n%).Parameter(1).Y = LY + 5
    ListForm(n%).Parameter(2).X = LX - 5
    ListForm(n%).Parameter(2).Y = LY + 5
    Call DrawForm
End Sub

'
' ***************************
' Change the number of vertex
' ***************************
'
Private Sub FACES_Vertex_Click()
    Dim v%
    v% = Val(InputBox("Number of vertex", "Faces", ListForm(Selected%).OneCurve.Vertex))
    If v% <= 0 Then Exit Sub
    ListForm(Selected%).OneCurve.Vertex = v%
    Call DrawForm
End Sub

'
' *************
' Delete a form
' *************
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i%, t%
    If KeyCode = vbKeyDelete Then
        If Selected% <> 0 Then
            t% = UBound(ListForm())
            If Selected% <> t% Then
                For i% = Selected% To t% - 1
                    ListForm(i%) = ListForm(i% + 1)
                Next i%
            End If
            ReDim Preserve ListForm(t% - 1) As TypeForm
            Selected% = 0
            SelectN% = -1
            Call DrawForm
        End If
    End If
End Sub

'
' ************************
' Initialize default value
' ************************
'
Private Sub Form_Load()
    FACES_Create.Visible = False ' Hide menu
    FACES_Curves.Visible = False
    FACES_Arcs.Visible = False
    StartX% = 120
    StartY% = 120
    ReDim ListForm(0) As TypeForm
    Selected% = 0
    Call DrawForm
End Sub

'
' ***********************
' Select a form, if exist
' ***********************
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n%, i%, t%
    Dim Xm%, Ym%
    If Button <> vbLeftButton Then Exit Sub
    If Selected% <> 0 Then
        Select Case ListForm(Selected%).TheType
        Case EnumForm.FormTriangle
            For i% = 0 To 2
                If X - StartX% <= ListForm(Selected%).Parameter(i%).X + 2 _
                And X - StartX% >= ListForm(Selected%).Parameter(i%).X - 2 _
                And Y - StartY% <= ListForm(Selected%).Parameter(i%).Y + 2 _
                And Y - StartY% >= ListForm(Selected%).Parameter(i%).Y - 2 Then
                    If SelectN% <> i% Then
                        SelectN% = i%
                        Call DrawForm
                        Exit Sub
                    End If
                End If
            Next i%
        Case EnumForm.FormRectangle
            For i% = 0 To 3
                If X - StartX% <= ListForm(Selected%).Parameter(i%).X + 2 _
                And X - StartX% >= ListForm(Selected%).Parameter(i%).X - 2 _
                And Y - StartY% <= ListForm(Selected%).Parameter(i%).Y + 2 _
                And Y - StartY% >= ListForm(Selected%).Parameter(i%).Y - 2 Then
                    If SelectN% <> i% Then
                        SelectN% = i%
                        Call DrawForm
                        Exit Sub
                    End If
                End If
            Next i%
        Case EnumForm.FormCurve
            If X <= ListForm(Selected%).OneCurve.X1 + 2 _
            And X >= ListForm(Selected%).OneCurve.X1 - 2 _
            And Y <= ListForm(Selected%).OneCurve.Y1 + 2 _
            And Y >= ListForm(Selected%).OneCurve.Y1 - 2 Then
                If SelectN% <> 0 Then
                    SelectN% = 0
                    Call DrawForm
                    Exit Sub
                End If
            End If
            If X <= ListForm(Selected%).OneCurve.X2 + 2 _
            And X >= ListForm(Selected%).OneCurve.X2 - 2 _
            And Y <= ListForm(Selected%).OneCurve.Y2 + 2 _
            And Y >= ListForm(Selected%).OneCurve.Y2 - 2 Then
                If SelectN% <> 1 Then
                    SelectN% = 1
                    Call DrawForm
                    Exit Sub
                End If
            End If
            If X - StartX% <= ListForm(Selected%).Parameter(0).X + 2 _
            And X - StartX% >= ListForm(Selected%).Parameter(0).X - 2 _
            And Y - StartY% <= ListForm(Selected%).Parameter(0).Y + 2 _
            And Y - StartY% >= ListForm(Selected%).Parameter(0).Y - 2 Then
                If SelectN% <> 2 Then
                    SelectN% = 2
                    Call DrawForm
                    Exit Sub
                End If
            End If
        Case EnumForm.FormArc
            If X <= ListForm(Selected%).OneArc.X + 2 _
            And X >= ListForm(Selected%).OneArc.X - 2 _
            And Y <= ListForm(Selected%).OneArc.Y + 2 _
            And Y >= ListForm(Selected%).OneArc.Y - 2 Then
                If SelectN% <> 0 Then
                    SelectN% = 0
                    Call DrawForm
                    Exit Sub
                End If
            End If
            If X - StartX% <= ListForm(Selected%).Parameter(0).X + 2 _
            And X - StartX% >= ListForm(Selected%).Parameter(0).X - 2 _
            And Y - StartY% <= ListForm(Selected%).Parameter(0).Y + 2 _
            And Y - StartY% >= ListForm(Selected%).Parameter(0).Y - 2 Then
                If SelectN% <> 2 Then
                    SelectN% = 2
                    Call DrawForm
                    Exit Sub
                End If
            End If
        End Select
    End If
    If SelectN% <> -1 Then
        Select Case ListForm(Selected%).TheType
        Case EnumForm.FormTriangle, EnumForm.FormRectangle
            ListForm(Selected%).Parameter(SelectN%).X = X - StartX%
            ListForm(Selected%).Parameter(SelectN%).Y = Y - StartY%
        Case EnumForm.FormCurve
            If SelectN% = 2 Then
                ListForm(Selected%).Parameter(0).X = X - StartX%
                ListForm(Selected%).Parameter(0).Y = Y - StartY%
            Else
                If SelectN% = 0 Then
                    ListForm(Selected%).OneCurve.X1 = X
                    ListForm(Selected%).OneCurve.Y1 = Y
                Else
                    ListForm(Selected%).OneCurve.X2 = X
                    ListForm(Selected%).OneCurve.Y2 = Y
                End If
            End If
        Case EnumForm.FormArc
            If SelectN% = 2 Then
                ListForm(Selected%).Parameter(0).X = X - StartX%
                ListForm(Selected%).Parameter(0).Y = Y - StartY%
            Else
                ListForm(Selected%).OneArc.X = X
                ListForm(Selected%).OneArc.Y = Y
            End If
        End Select
        SelectN% = -1
    Else
        Xm% = X - StartX%
        Ym% = Y - StartY%
        Selected% = 0
        SelectN% = -1
        If UBound(ListForm()) <> 0 Then
            For n% = 1 To UBound(ListForm())
                Select Case ListForm(n%).TheType
                Case EnumForm.FormTriangle
                    For i% = 0 To 2
                        If MouseTestHit(Xm%, Ym%, _
                            ListForm(n%).Parameter(i%).X, _
                            ListForm(n%).Parameter(i%).Y, _
                            ListForm(n%).Parameter((i% + 1) Mod 3).X, _
                            ListForm(n%).Parameter((i% + 1) Mod 3).Y) = True Then Selected% = n%
                    Next i%
                Case EnumForm.FormRectangle
                    For i% = 0 To 3
                        If MouseTestHit(Xm%, Ym%, _
                            ListForm(n%).Parameter(i%).X, _
                            ListForm(n%).Parameter(i%).Y, _
                            ListForm(n%).Parameter((i% + 1) Mod 4).X, _
                            ListForm(n%).Parameter((i% + 1) Mod 4).Y) = True Then Selected% = n%
                    Next i%
                Case EnumForm.FormCurve
                    If ListForm(n%).OneCurve.TestHit(X, Y) = True Then
                        Selected% = n%
                    End If
                Case EnumForm.FormArc
                    If ListForm(n%).OneArc.TestHit(X, Y) = True Then
                        Selected% = n%
                    End If
                End Select
                If Selected% <> 0 Then Exit For
            Next n%
        End If
    End If
    Call DrawForm
End Sub

'
' *******************
' Show mouse position
' *******************
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LX = X - StartX%
    LY = Y - StartY%
End Sub

'
' ******************
' Activate PopupMenu
' ******************
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    If Selected% <> 0 Then
        If ListForm(Selected%).TheType = FormCurve Then
        PopupMenu FACES_Curves
        ElseIf ListForm(Selected%).TheType = FormArc Then
            PopupMenu FACES_Arcs
        Else
            PopupMenu FACES_Create
        End If
    Else
        PopupMenu FACES_Create
    End If
End Sub

'
' ****************************************
' Draw axes and form for the faces builder
' ****************************************
'
Public Sub DrawForm()
    Dim i%, j%
    Dim Apply As D3DVECTOR
    Dim Normal As D3DVECTOR
    If UBound(ListForm()) = 0 Then
        CreateFace.Enabled = False
    Else
        CreateFace.Enabled = True
    End If
    Cls
    Me.Line (0, StartY%)-(StartX% * 2, StartY%), vbBlack
    Me.Line (StartX%, 0)-(StartX%, StartY% * 2), vbBlack
    If UBound(ListForm()) = 0 Then Exit Sub
    For i% = 1 To UBound(ListForm())
        With ListForm(i%)
            Select Case .TheType
            Case EnumForm.FormTriangle
                Me.Line (.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%)-(.Parameter(1).X + StartX%, .Parameter(1).Y + StartY%), vbBlack
                Me.Line (.Parameter(1).X + StartX%, .Parameter(1).Y + StartY%)-(.Parameter(2).X + StartX%, .Parameter(2).Y + StartY%), vbBlack
                Me.Line (.Parameter(2).X + StartX%, .Parameter(2).Y + StartY%)-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), vbBlack
            Case EnumForm.FormRectangle
                Me.Line (.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%)-(.Parameter(1).X + StartX%, .Parameter(1).Y + StartY%), vbBlack
                Me.Line (.Parameter(1).X + StartX%, .Parameter(1).Y + StartY%)-(.Parameter(2).X + StartX%, .Parameter(2).Y + StartY%), vbBlack
                Me.Line (.Parameter(2).X + StartX%, .Parameter(2).Y + StartY%)-(.Parameter(3).X + StartX%, .Parameter(3).Y + StartY%), vbBlack
                Me.Line (.Parameter(3).X + StartX%, .Parameter(3).Y + StartY%)-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), vbBlack
            Case EnumForm.FormCurve
                .OneCurve.Draw Me
                For j% = 1 To .OneCurve.Vertex
                    .OneCurve.GetNormal j%, Apply, Normal
                    If j% <> .OneCurve.Vertex Then
                        Me.Line (.OneCurve.GetX(j%), .OneCurve.GetY(j%))-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), RGB(150, 150, 150)
                    End If
                    Me.Line (Apply.X, Apply.Y)-(Apply.X + Normal.X * 5, Apply.Y + Normal.Y * 5), vbRed
                Next j%
                Me.Line (.OneCurve.X1, .OneCurve.Y1)-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), vbMagenta
                Me.Line (.OneCurve.X2, .OneCurve.Y2)-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), vbMagenta
            Case EnumForm.FormArc
                .OneArc.Draw Me
                For j% = 1 To .OneArc.Vertex
                    .OneArc.GetNormal j%, Apply, Normal
                    If j% <> .OneArc.Vertex Then
                        Me.Line (.OneArc.GetX(j%), .OneArc.GetY(j%))-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), RGB(150, 150, 150)
                    End If
                    Me.Line (Apply.X, Apply.Y)-(Apply.X + Normal.X * 5, Apply.Y + Normal.Y * 5), vbRed
                Next j%
                Me.Line (.OneArc.GetX(0), .OneArc.GetY(0))-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), vbMagenta
                Me.Line (.OneArc.GetX(.OneArc.Vertex), .OneArc.GetY(.OneArc.Vertex))-(.Parameter(0).X + StartX%, .Parameter(0).Y + StartY%), vbMagenta
            End Select
        End With
    Next i%
    If Selected% <> 0 Then
        With ListForm(Selected%)
            Select Case .TheType
            Case EnumForm.FormTriangle
                TX(3).Visible = False
                TY(3).Visible = False
                For i% = 0 To 2
                    TX(i%).Visible = True
                    TY(i%).Visible = True
                    TX(i%).Text = .Parameter(i%).X
                    TY(i%).Text = .Parameter(i%).Y
                    If SelectN% = i% Then
                        Me.Line ( _
                        .Parameter(i%).X - 2 + StartX%, _
                        .Parameter(i%).Y - 2 + StartY%)-( _
                        .Parameter(i%).X + 2 + StartX%, _
                        .Parameter(i%).Y + 2 + StartY%), vbBlue, B
                    Else
                        Me.Line ( _
                        .Parameter(i%).X - 2 + StartX%, _
                        .Parameter(i%).Y - 2 + StartY%)-( _
                        .Parameter(i%).X + 2 + StartX%, _
                        .Parameter(i%).Y + 2 + StartY%), vbRed, B
                    End If
                Next i%
            Case EnumForm.FormRectangle
                For i% = 0 To 3
                    TX(i%).Visible = True
                    TY(i%).Visible = True
                    TX(i%).Text = .Parameter(i%).X
                    TY(i%).Text = .Parameter(i%).Y
                    If SelectN% = i% Then
                        Me.Line ( _
                        .Parameter(i%).X - 2 + StartX%, _
                        .Parameter(i%).Y - 2 + StartY%)-( _
                        .Parameter(i%).X + 2 + StartX%, _
                        .Parameter(i%).Y + 2 + StartY%), vbBlue, B
                    Else
                        Me.Line ( _
                        .Parameter(i%).X - 2 + StartX%, _
                        .Parameter(i%).Y - 2 + StartY%)-( _
                        .Parameter(i%).X + 2 + StartX%, _
                        .Parameter(i%).Y + 2 + StartY%), vbRed, B
                    End If
                Next i%
            Case EnumForm.FormCurve
                TX(0).Visible = True
                TY(0).Visible = True
                TX(0).Text = .Parameter(i%).X
                TY(0).Text = .Parameter(i%).Y
                For i% = 1 To 3
                    TX(i%).Visible = False
                    TY(i%).Visible = False
                Next i%
                If SelectN% = 0 Then
                    Me.Line (.OneCurve.X1 - 2, .OneCurve.Y1 - 2)-(.OneCurve.X1 + 2, .OneCurve.Y1 + 2), vbBlue, B
                Else
                    Me.Line (.OneCurve.X1 - 2, .OneCurve.Y1 - 2)-(.OneCurve.X1 + 2, .OneCurve.Y1 + 2), vbRed, B
                End If
                If SelectN% = 1 Then
                    Me.Line (.OneCurve.X2 - 2, .OneCurve.Y2 - 2)-(.OneCurve.X2 + 2, .OneCurve.Y2 + 2), vbBlue, B
                Else
                    Me.Line (.OneCurve.X2 - 2, .OneCurve.Y2 - 2)-(.OneCurve.X2 + 2, .OneCurve.Y2 + 2), vbRed, B
                End If
                If SelectN% = 2 Then
                    Me.Line ( _
                    .Parameter(0).X - 2 + StartX%, _
                    .Parameter(0).Y - 2 + StartY%)-( _
                    .Parameter(0).X + 2 + StartX%, _
                    .Parameter(0).Y + 2 + StartY%), vbBlue, B
                Else
                    Me.Line ( _
                    .Parameter(0).X - 2 + StartX%, _
                    .Parameter(0).Y - 2 + StartY%)-( _
                    .Parameter(0).X + 2 + StartX%, _
                    .Parameter(0).Y + 2 + StartY%), vbRed, B
                End If
            Case EnumForm.FormArc
                TX(0).Visible = True
                TY(0).Visible = True
                TX(0).Text = .Parameter(i%).X
                TY(0).Text = .Parameter(i%).Y
                For i% = 1 To 3
                    TX(i%).Visible = False
                    TY(i%).Visible = False
                Next i%
                If SelectN% = 0 Then
                    Me.Line (.OneArc.X - 2, .OneArc.Y - 2)-(.OneArc.X + 2, .OneArc.Y + 2), vbBlue, B
                Else
                    Me.Line (.OneArc.X - 2, .OneArc.Y - 2)-(.OneArc.X + 2, .OneArc.Y + 2), vbRed, B
                End If
                If SelectN% = 2 Then
                    Me.Line ( _
                    .Parameter(0).X - 2 + StartX%, _
                    .Parameter(0).Y - 2 + StartY%)-( _
                    .Parameter(0).X + 2 + StartX%, _
                    .Parameter(0).Y + 2 + StartY%), vbBlue, B
                Else
                    Me.Line ( _
                    .Parameter(0).X - 2 + StartX%, _
                    .Parameter(0).Y - 2 + StartY%)-( _
                    .Parameter(0).X + 2 + StartX%, _
                    .Parameter(0).Y + 2 + StartY%), vbRed, B
                End If
            End Select
        End With
    Else
        For i% = 0 To 3
            TX(i%).Visible = False
            TY(i%).Visible = False
        Next i%
    End If
End Sub

'
' *******************************
' Test if the mouse hit a segment
' *******************************
'
Public Function MouseTestHit(ByVal X%, ByVal Y%, ByVal X1%, ByVal Y1%, ByVal X2%, ByVal Y2%) As Boolean
    Dim a!, b!
    Dim Xa%, Xb%, Ya%, Yb%
    MouseTestHit = False
    If X1% < X2% Then
        Xa% = X1%
        Xb% = X2%
    Else
        Xb% = X1%
        Xa% = X2%
    End If
    If Y1% < Y2% Then
        Ya% = Y1%
        Yb% = Y2%
    Else
        Yb% = Y1%
        Ya% = Y2%
    End If
    If Xa% - 1 <= X% And X% <= Xb% + 1 Then
        If Ya% - 1 <= Y% And Y% <= Yb% + 1 Then
            If X1% = X2% Then
                MouseTestHit = True
            Else
                a! = (Y1% - Y2%) / (X1% - X2%)
                b! = Y2% - a! * X2%
                If Abs(a!) < 1 Then
                    If Abs((a! * X% + b! - Y%) * a!) < 1 Then MouseTestHit = True
                Else
                    If Abs((a! * X% + b! - Y%) / a!) < 1 Then MouseTestHit = True
                End If
            End If
        End If
    End If
End Function

'
' ************************************
' Create side for object with a volume
' ************************************
'
Public Sub CreateSide(Color As PALETTEENTRY, S1 As D3DVECTOR, S2 As D3DVECTOR, Inversion As Boolean)
    Dim p(3) As D3DVECTOR
    p(0).X = S1.X: p(0).Y = S1.Y: p(0).z = -0.5
    p(1).X = S1.X: p(1).Y = S1.Y: p(1).z = 0.5
    p(2).X = S2.X: p(2).Y = S2.Y: p(2).z = 0.5
    p(3).X = S2.X: p(3).Y = S2.Y: p(3).z = -0.5
    Call Tools3D.Face_Add(Wall3D.No_Mur, Color, 0, 4, p(), Inversion)
End Sub

'
' ****************
' Input X position
' ****************
'
Private Sub TX_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    ListForm(Selected%).Parameter(Index).X = Val(TX(Index).Text)
    Call DrawForm
End Sub

'
' ****************
' Input Y position
' ****************
'
Private Sub TY_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    ListForm(Selected%).Parameter(Index).Y = Val(TY(Index).Text)
    Call DrawForm
End Sub

