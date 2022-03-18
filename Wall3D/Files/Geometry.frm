VERSION 5.00
Object = "{EC2CC72E-13BA-11D5-BB31-400001686160}#1.0#0"; "SELECTCOLOR.OCX"
Begin VB.Form Geometry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add a new geometry"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   Icon            =   "Geometry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CView 
      Caption         =   "No top"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox CView 
      Caption         =   "Close side"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox CView 
      Caption         =   "Out-Sides"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox CView 
      Caption         =   "In-Sides"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.HScrollBar NbFaces 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   360
      Min             =   3
      TabIndex        =   10
      Top             =   600
      Value           =   5
      Width           =   1455
   End
   Begin VB.OptionButton Choix_Geometry 
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   1920
      MaskColor       =   &H0000FFFF&
      Picture         =   "Geometry.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Tore"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Choix_Geometry 
      Height          =   375
      Index           =   4
      Left            =   1560
      MaskColor       =   &H0000FFFF&
      Picture         =   "Geometry.frx":064C
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Cone"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin SelectColor.UserColor GeoColor 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1720
   End
   Begin VB.HScrollBar Ouverture 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   360
      TabIndex        =   5
      Top             =   960
      Value           =   360
      Width           =   1455
   End
   Begin VB.CommandButton Geometry_OK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton Choix_Geometry 
      Height          =   375
      Index           =   3
      Left            =   1200
      MaskColor       =   &H0000FFFF&
      Picture         =   "Geometry.frx":098E
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Sphere"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Choix_Geometry 
      Height          =   375
      Index           =   2
      Left            =   840
      MaskColor       =   &H0000FFFF&
      Picture         =   "Geometry.frx":0CD0
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Cylinder"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Choix_Geometry 
      Height          =   375
      Index           =   1
      Left            =   480
      MaskColor       =   &H0000FFFF&
      Picture         =   "Geometry.frx":1012
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Cube"
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton Choix_Geometry 
      Height          =   375
      Index           =   0
      Left            =   120
      MaskColor       =   &H0000FFFF&
      Picture         =   "Geometry.frx":1354
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Face"
      Top             =   120
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.Label LFaces 
      Caption         =   "N faces"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LOuverture 
      Caption         =   "360°"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Geometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ****
' Exit
' ****
'
Private Sub Cancel_Click()
    Unload Me
End Sub

'
' *********************
' Select enabled option
' *********************
'
Private Sub Choix_Geometry_Click(Index As Integer)
    Select Case Index
    Case 0
        CView(2).Visible = False
        CView(3).Visible = False
    Case 1
        CView(2).Visible = False
        CView(3).Visible = False
    Case 2
        CView(2).Visible = True
        CView(3).Visible = True
    Case 3
        CView(2).Visible = True
        CView(3).Visible = False
    Case 4
        CView(2).Visible = True
        CView(3).Visible = True
    Case 5
    End Select
End Sub

'
' *********
' Enable OK
' *********
'
Private Sub CView_Click(Index As Integer)
    If CView(0) = vbUnchecked And CView(1) = vbUnchecked Then
        Geometry_OK.Enabled = False
    Else
        Geometry_OK.Enabled = True
    End If
End Sub

'
' ***********************
' Default value for color
' ***********************
'
Private Sub Form_Load()
    GeoColor.Red = 255
    GeoColor.Green = 255
    GeoColor.Blue = 255
    GeoColor.Alpha = 255
    NbFaces = 10
End Sub

'
' *********************
' Add faces to the mesh
' *********************
'
Private Sub Geometry_OK_Click()
    Dim i%
    Dim p(3) As D3DVECTOR
    Dim Color As PALETTEENTRY
    Color.Red = GeoColor.Red
    Color.Green = GeoColor.Green
    Color.Blue = GeoColor.Blue
    Color.flags = GeoColor.Alpha
    If Choix_Geometry(0).Value = True Then
        p(0).X = -0.5: p(0).Y = -0.5: p(0).z = 0
        p(1).X = -0.5: p(1).Y = 0.5: p(1).z = 0
        p(2).X = 0.5: p(2).Y = 0.5: p(2).z = 0
        p(3).X = 0.5: p(3).Y = -0.5: p(3).z = 0
        If CView(0) = vbChecked Then
            Call Tools3D.Face_Add(Wall3D.No_Mur.Value, Color, 0, 4, p(), False)
        End If
        If CView(1) = vbChecked Then
            Call Tools3D.Face_Add(Wall3D.No_Mur.Value, Color, 0, 4, p(), True)
        End If
    End If
    If Choix_Geometry(1).Value = True Then
        If CView(0) = vbChecked Then
            Call dxShape.Add_Cube(Wall3D.No_Mur.Value, Color, False)
        End If
        If CView(1) = vbChecked Then
            Call dxShape.Add_Cube(Wall3D.No_Mur.Value, Color, True)
        End If
    End If
    If Choix_Geometry(2).Value = True Then
        If CView(3) = vbChecked Then
            ReDim ListPoint(0) As TypeShape ' Cylinder with no top
            ListPoint(0).X1 = 0.5: ListPoint(0).Y1 = 0.5
            ListPoint(0).X2 = 0.5: ListPoint(0).Y2 = -0.5
            ListPoint(0).Color = Color
        Else
            ReDim ListPoint(2) As TypeShape ' Cylinder
            ListPoint(0).X1 = 0: ListPoint(0).Y1 = 0.5
            ListPoint(0).X2 = 0.5: ListPoint(0).Y2 = 0.5
            ListPoint(0).Color = Color
            ListPoint(1).X1 = 0.5: ListPoint(1).Y1 = 0.5
            ListPoint(1).X2 = 0.5: ListPoint(1).Y2 = -0.5
            ListPoint(1).Color = Color
            ListPoint(2).X1 = 0.5: ListPoint(2).Y1 = -0.5
            ListPoint(2).X2 = 0: ListPoint(2).Y2 = -0.5
            ListPoint(2).Color = Color
        End If
        If CView(0) = vbChecked Then
            Call Add_Shape(Wall3D.No_Mur.Value, ListPoint(), Ouverture.Value, NbFaces.Value, False, CView(2).Value)
        End If
        If CView(1) = vbChecked Then
            Call Add_Shape(Wall3D.No_Mur.Value, ListPoint(), Ouverture.Value, NbFaces.Value, True, CView(2).Value)
        End If
    End If
    If Choix_Geometry(3).Value = True Then
        If CView(0) = vbChecked Then
            Call dxShape.Add_Sphere(Wall3D.No_Mur.Value, NbFaces.Value, Ouverture.Value, Color, False, CView(2).Value)
        End If
        If CView(1) = vbChecked Then
            Call dxShape.Add_Sphere(Wall3D.No_Mur.Value, NbFaces.Value, Ouverture.Value, Color, True, CView(2).Value)
        End If
    End If
    If Choix_Geometry(4).Value = True Then ' Cone
        If CView(3) = vbChecked Then
            ReDim ListPoint(0) As TypeShape
            ListPoint(0).X1 = 0: ListPoint(0).Y1 = 1
            ListPoint(0).X2 = 1: ListPoint(0).Y2 = 0
            ListPoint(0).Color = Color
        Else
            ReDim ListPoint(1) As TypeShape
            ListPoint(0).X1 = 0: ListPoint(0).Y1 = 1
            ListPoint(0).X2 = 1: ListPoint(0).Y2 = 0
            ListPoint(0).Color = Color
            ListPoint(1).X1 = 1: ListPoint(1).Y1 = 0
            ListPoint(1).X2 = 0: ListPoint(1).Y2 = 0
            ListPoint(1).Color = Color
        End If
        If CView(0) = vbChecked Then
            Call Add_Shape(Wall3D.No_Mur.Value, ListPoint(), Ouverture.Value, NbFaces.Value, False, CView(2).Value)
        End If
        If CView(1) = vbChecked Then
            Call Add_Shape(Wall3D.No_Mur.Value, ListPoint(), Ouverture.Value, NbFaces.Value, True, CView(2).Value)
        End If
    End If
    Wall3D.Update
    Unload Me
End Sub

Private Sub NbFaces_Change()
    LFaces.Caption = Format$(NbFaces) & " faces"
End Sub

'
' ****************************************
' Change the open angle for spherique mesh
' ****************************************
'
Private Sub Ouverture_Change()
    LOuverture.Caption = Format$(Ouverture.Value) & "°"
End Sub

