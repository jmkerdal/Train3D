VERSION 5.00
Object = "{EC2CC72E-13BA-11D5-BB31-400001686160}#1.0#0"; "SELECTCOLOR.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Wall3D 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mesh & Faces informations"
   ClientHeight    =   3750
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5415
   ControlBox      =   0   'False
   Icon            =   "Mur 3D.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6165
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "Mesh"
      TabPicture(0)   =   "Mur 3D.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FMesh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FPosition"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FWrap"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Faces"
      TabPicture(1)   =   "Mur 3D.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FFace"
      Tab(1).Control(1)=   "FTexture"
      Tab(1).Control(2)=   "ColorAll"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Material"
      TabPicture(2)   =   "Mur 3D.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MaterialColor(0)"
      Tab(2).Control(1)=   "MaterialColor(1)"
      Tab(2).Control(2)=   "MaterialColor(2)"
      Tab(2).Control(3)=   "Label1(2)"
      Tab(2).Control(4)=   "Label1(1)"
      Tab(2).Control(5)=   "Label1(0)"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton ColorAll 
         Caption         =   "Color all faces"
         Height          =   495
         Left            =   -71520
         TabIndex        =   80
         Top             =   2280
         Width           =   1215
      End
      Begin SelectColor.UserColor MaterialColor 
         Height          =   975
         Index           =   0
         Left            =   -74760
         TabIndex        =   74
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1720
      End
      Begin VB.Frame FWrap 
         Caption         =   "Wrapping"
         Height          =   3015
         Left            =   3960
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         Begin VB.OptionButton Choix_Wrap 
            Height          =   375
            Index           =   5
            Left            =   600
            MaskColor       =   &H0000FFFF&
            Picture         =   "Mur 3D.frx":035E
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Box"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Choix_Wrap 
            Height          =   375
            Index           =   4
            Left            =   240
            MaskColor       =   &H0000FFFF&
            Picture         =   "Mur 3D.frx":06A0
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Sheet"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.ComboBox CTexture 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   68
            Text            =   "Combo1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Wrap_Ny 
            Height          =   285
            Left            =   360
            TabIndex        =   57
            Text            =   "Ny"
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox Wrap_Nx 
            Height          =   285
            Left            =   360
            TabIndex        =   56
            Text            =   "Nx"
            Top             =   1800
            Width           =   735
         End
         Begin VB.OptionButton Choix_Wrap 
            Height          =   375
            Index           =   3
            Left            =   600
            MaskColor       =   &H0000FFFF&
            Picture         =   "Mur 3D.frx":09E2
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Chrome"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Choix_Wrap 
            Height          =   375
            Index           =   2
            Left            =   240
            MaskColor       =   &H0000FFFF&
            Picture         =   "Mur 3D.frx":0D24
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Spherical"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Choix_Wrap 
            Height          =   375
            Index           =   1
            Left            =   600
            MaskColor       =   &H0000FFFF&
            Picture         =   "Mur 3D.frx":1066
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Cylindrical"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.OptionButton Choix_Wrap 
            Height          =   375
            Index           =   0
            Left            =   240
            MaskColor       =   &H0000FFFF&
            Picture         =   "Mur 3D.frx":13A8
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Flat"
            Top             =   600
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   375
         End
         Begin VB.Label LWrap 
            Caption         =   "Ny"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   59
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label LWrap 
            Caption         =   "Nx"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   1800
            Width           =   255
         End
      End
      Begin VB.Frame FPosition 
         Caption         =   "Position, object rotation, scaling and axe rotation"
         Height          =   1935
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox No_Parent 
            Height          =   285
            Left            =   2880
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   480
            Width           =   735
         End
         Begin VB.Label LPosition 
            Caption         =   "Axe Roll"
            Height          =   255
            Index           =   11
            Left            =   1680
            TabIndex        =   67
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   11
            Left            =   2520
            TabIndex        =   66
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label LPosition 
            Caption         =   "Axe Phi"
            Height          =   255
            Index           =   10
            Left            =   1680
            TabIndex        =   65
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   10
            Left            =   2520
            TabIndex        =   64
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label LPosition 
            Caption         =   "Axe Theta"
            Height          =   255
            Index           =   9
            Left            =   1680
            TabIndex        =   63
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   9
            Left            =   2520
            TabIndex        =   62
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label LParent 
            Caption         =   "Parent Frame"
            Height          =   255
            Left            =   2760
            TabIndex        =   61
            Top             =   240
            Width           =   975
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   8
            Left            =   1920
            TabIndex        =   51
            Top             =   720
            Width           =   735
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   7
            Left            =   1920
            TabIndex        =   50
            Top             =   480
            Width           =   735
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   6
            Left            =   1920
            TabIndex        =   49
            Top             =   240
            Width           =   735
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   48
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   47
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   46
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   45
            Top             =   720
            Width           =   735
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   44
            Top             =   480
            Width           =   735
         End
         Begin VB.Label TPosition 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   43
            Top             =   240
            Width           =   735
         End
         Begin VB.Label LPosition 
            Caption         =   "Scale Z"
            Height          =   255
            Index           =   8
            Left            =   1320
            TabIndex        =   41
            Top             =   720
            Width           =   615
         End
         Begin VB.Label LPosition 
            Caption         =   "Scale Y"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   40
            Top             =   480
            Width           =   615
         End
         Begin VB.Label LPosition 
            Caption         =   "Roll"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label LPosition 
            Caption         =   "Phi"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   37
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label LPosition 
            Caption         =   "Theta"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label LPosition 
            Caption         =   "Dz"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   255
         End
         Begin VB.Label LPosition 
            Caption         =   "Dy"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   255
         End
         Begin VB.Label LPosition 
            Caption         =   "Dx"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LPosition 
            Caption         =   "Scale X"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame FTexture 
         Caption         =   "Texture && Color"
         Height          =   1695
         Left            =   -72240
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
         Begin VB.ComboBox CTexture 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Text            =   "Combo1"
            Top             =   240
            Width           =   2295
         End
         Begin SelectColor.UserColor FaceColor 
            Height          =   975
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1720
         End
      End
      Begin VB.Frame FMesh 
         Caption         =   "X mesh"
         Height          =   1095
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton LoadAnimation 
            Caption         =   "Load animation"
            Height          =   375
            Left            =   2520
            TabIndex        =   71
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton LoadX 
            Caption         =   "Load X mesh"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton SaveX 
            Caption         =   "Save X mesh"
            Height          =   375
            Left            =   1320
            TabIndex        =   27
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Fichier_X 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X file name"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame FFace 
         Caption         =   "Face"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton Offset 
            Caption         =   "-"
            Height          =   195
            Index           =   5
            Left            =   2280
            TabIndex        =   91
            Top             =   2640
            Width           =   255
         End
         Begin VB.CommandButton Offset 
            Caption         =   "+"
            Height          =   195
            Index           =   4
            Left            =   2040
            TabIndex        =   90
            Top             =   2640
            Width           =   255
         End
         Begin VB.CommandButton Offset 
            Caption         =   "-"
            Height          =   195
            Index           =   3
            Left            =   1440
            TabIndex        =   89
            Top             =   2640
            Width           =   255
         End
         Begin VB.CommandButton Offset 
            Caption         =   "+"
            Height          =   195
            Index           =   2
            Left            =   1200
            TabIndex        =   88
            Top             =   2640
            Width           =   255
         End
         Begin VB.CommandButton Offset 
            Caption         =   "-"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   87
            Top             =   2640
            Width           =   255
         End
         Begin VB.CommandButton Offset 
            Caption         =   "+"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   86
            Top             =   2640
            Width           =   255
         End
         Begin VB.CommandButton FaceInvertion 
            Caption         =   "Invertion"
            Height          =   375
            Left            =   1560
            TabIndex        =   82
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton FaceRotate 
            Caption         =   "Rotate"
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox Point_X 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   17
            Text            =   "X"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_X 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   16
            Text            =   "X"
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_X 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   15
            Text            =   "X"
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_X 
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   14
            Text            =   "X"
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Y 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   13
            Text            =   "Y"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Y 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Text            =   "Y"
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Y 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   11
            Text            =   "Y"
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Y 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   10
            Text            =   "Y"
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Z 
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   9
            Text            =   "Z"
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Z 
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   8
            Text            =   "Z"
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Z 
            Height          =   285
            Index           =   2
            Left            =   1920
            TabIndex        =   7
            Text            =   "Z"
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Point_Z 
            Height          =   285
            Index           =   3
            Left            =   1920
            TabIndex        =   6
            Text            =   "Z"
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.HScrollBar No_Face 
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox Nb_Cote 
            Caption         =   "Triangle/Rectangle"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1560
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label LabelDecale 
            Caption         =   "Z"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   85
            Top             =   2640
            Width           =   135
         End
         Begin VB.Label LabelDecale 
            Caption         =   "Y"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   84
            Top             =   2640
            Width           =   135
         End
         Begin VB.Label LabelDecale 
            Caption         =   "X"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   83
            Top             =   2640
            Width           =   135
         End
         Begin VB.Label LPoint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Point 1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label LPoint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Point 2"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label LPoint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Point 3"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label LPoint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Point 4"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label LNo_Face 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "No Face"
            Height          =   255
            Left            =   1920
            TabIndex        =   21
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label LPoint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X"
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   20
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LPoint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Y"
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LPoint 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Z"
            Height          =   255
            Index           =   6
            Left            =   1920
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
      End
      Begin SelectColor.UserColor MaterialColor 
         Height          =   975
         Index           =   1
         Left            =   -74760
         TabIndex        =   75
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1720
      End
      Begin SelectColor.UserColor MaterialColor 
         Height          =   975
         Index           =   2
         Left            =   -72120
         TabIndex        =   76
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1720
      End
      Begin VB.Label Label1 
         Caption         =   "Specular light"
         Height          =   255
         Index           =   2
         Left            =   -72120
         TabIndex        =   79
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Emissive light"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   78
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Ambient light"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   77
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.HScrollBar No_Mur 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog BoiteMur3D 
      Left            =   360
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   70
      Top             =   3480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label LNo_Mur 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   375
   End
End
Attribute VB_Name = "Wall3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' *******************
' Chose wrapping mode
' *******************
'
Private Sub Choix_Wrap_Click(Index As Integer)
    DefWall(No_Mur).WrapType = Index
    Select Case Index
    Case 0, 4, 5
        Wrap_Nx.Visible = True
        LWrap(1).Visible = True
        Wrap_Ny.Visible = True
        LWrap(2).Visible = True
    Case 1, 2, 3
        Wrap_Nx.Visible = False
        LWrap(1).Visible = False
        Wrap_Ny.Visible = False
        LWrap(2).Visible = False
    End Select
    Call Build_Wall(dxEngine, No_Mur.Value, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
    Call View.SetFocus
End Sub

'
' ****************************
' Change the color of all face
' with current color
' ****************************
'
Private Sub ColorAll_Click()
    Dim i%
    Dim NewColor As PALETTEENTRY
    NewColor = DefFace(No_Face).Color
    For i% = No_Face.MIN To No_Face.MAX
        DefFace(i%).Color = NewColor
    Next i%
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' *************************
' Chose the texture to wrap
' *************************
'
Private Sub CTexture_Click(Index As Integer)
    If Index = 0 Then
        DefWall(No_Mur.Value).WrapTexture% = CTexture(0).ListIndex
    Else
        DefFace(No_Face).Texture = CTexture(1).ListIndex
    End If
    Call Build_Wall(dxEngine, No_Mur.Value, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' ******************
' Modify Alpha color
' ******************
'
Private Sub FaceColor_ChangeAlpha()
    DefFace(No_Face).Color.flags = FaceColor.Alpha
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' *****************
' Modify Blue color
' *****************
'
Private Sub FaceColor_ChangeBlue()
    DefFace(No_Face).Color.Blue = FaceColor.Blue
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' ******************
' Modify Green color
' ******************
'
Private Sub FaceColor_ChangeGreen()
    DefFace(No_Face).Color.Green = FaceColor.Green
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub


'
' ****************
' Modify Red color
' ****************
'
Private Sub FaceColor_ChangeRed()
    DefFace(No_Face).Color.Red = FaceColor.Red
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' **************
' Face inversion
' **************
'
Private Sub FaceInvertion_Click()
    Dim Temp As D3DVECTOR
    If DefFace(No_Face).NbPoint = 4 Then
        Temp = DefFace(No_Face).Point(0)
        DefFace(No_Face).Point(0) = DefFace(No_Face).Point(1)
        DefFace(No_Face).Point(1) = Temp
        Temp = DefFace(No_Face).Point(2)
        DefFace(No_Face).Point(2) = DefFace(No_Face).Point(3)
        DefFace(No_Face).Point(3) = Temp
    Else
        Temp = DefFace(No_Face).Point(0)
        DefFace(No_Face).Point(0) = DefFace(No_Face).Point(2)
        DefFace(No_Face).Point(2) = Temp
    End If
    Call Show_Face
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' *************
' Rotate a face
' *************
'
Private Sub FaceRotate_Click()
    Dim Temp As D3DVECTOR
    Temp = DefFace(No_Face).Point(0)
    If DefFace(No_Face).NbPoint = 4 Then
        DefFace(No_Face).Point(0) = DefFace(No_Face).Point(1)
        DefFace(No_Face).Point(1) = DefFace(No_Face).Point(2)
        DefFace(No_Face).Point(2) = DefFace(No_Face).Point(3)
        DefFace(No_Face).Point(3) = Temp
    Else
        DefFace(No_Face).Point(0) = DefFace(No_Face).Point(1)
        DefFace(No_Face).Point(1) = DefFace(No_Face).Point(2)
        DefFace(No_Face).Point(2) = Temp
    End If
    Call Show_Face
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' **********************
' Query a new frame name
' **********************
'
Private Sub LName_DblClick()
    Dim a$, i%
    a$ = InputBox("Enter frame name", , DefWall(No_Mur).FrameName)
    If a$ = "" Then Exit Sub
    If a$ = DefWall(No_Mur).FrameName$ Then Exit Sub
    For i% = 1 To NbWall%
        If DefWall(i%).FrameName$ = a$ Then
            Call MsgBox("Another frame as this name", vbExclamation + vbOKOnly)
            Exit Sub
        End If
    Next i%
    DefWall(No_Mur).FrameName$ = a$
    LName = a$
End Sub

'
' *****************
' Load an animation
' *****************
'
Private Sub LoadAnimation_Click()
    Dim File$
    Dim Folder$
    File$ = Open_Box$("Load X Mesh", "", "DirectX X file|*.x", BOX_LOAD, Wall3D.BoiteMur3D, Folder$)
    If File$ <> "" Then
        Call dxEngine.Load_Animation(File$, DxWall(No_Mur).WallFrame)
        DefWall(No_Mur.Value).Faces = True
    Else
        DefWall(No_Mur.Value).TheFile$ = ""
        Set DxWall(No_Mur).WallMeshBuilder = dxEngine.dxD3Drm.CreateMeshBuilder
        DefWall(No_Mur.Value).Faces = False
    End If
    DefWall(No_Mur.Value).AnimationFile$ = File$
    Wall3D.Fichier_X.Caption = DefWall(No_Mur).AnimationFile$
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

Private Sub MaterialColor_ChangeAlpha(Index As Integer)
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

Private Sub MaterialColor_ChangeBlue(Index As Integer)
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

Private Sub MaterialColor_ChangeGreen(Index As Integer)
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

Private Sub MaterialColor_ChangeRed(Index As Integer)
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' ***********************
' Change the parent frame
' ***********************
'
Private Sub No_Parent_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    DefWall(No_Mur).ParentNo% = Val(No_Parent)
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' **************************************
' Modify the offset of a standard volume
' step is 0.5
' **************************************
'
Private Sub Offset_Click(Index As Integer)
    Dim i%, j%, n%
    For i% = No_Face.MIN To No_Face.MAX
        Select Case Index
        Case 0
            For j% = 0 To DefFace(i%).NbPoint% - 1
                DefFace(i%).Point(j%).X = DefFace(i%).Point(j%).X + 0.5
            Next j%
        Case 1
            For j% = 0 To DefFace(i%).NbPoint% - 1
                DefFace(i%).Point(j%).X = DefFace(i%).Point(j%).X - 0.5
            Next j%
        Case 2
            For j% = 0 To DefFace(i%).NbPoint% - 1
                DefFace(i%).Point(j%).Y = DefFace(i%).Point(j%).Y + 0.5
            Next j%
        Case 3
            For j% = 0 To DefFace(i%).NbPoint% - 1
                DefFace(i%).Point(j%).Y = DefFace(i%).Point(j%).Y - 0.5
            Next j%
        Case 4
            For j% = 0 To DefFace(i%).NbPoint% - 1
                DefFace(i%).Point(j%).z = DefFace(i%).Point(j%).z + 0.5
            Next j%
        Case 5
            For j% = 0 To DefFace(i%).NbPoint% - 1
                DefFace(i%).Point(j%).z = DefFace(i%).Point(j%).z - 0.5
            Next j%
        End Select
    Next i%
    Call Show_Face
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' ***********************
' Save actual meshbuilder
' ***********************
'
Private Sub SaveX_Click()
    Dim File$
    File$ = Open_Box$("Save X Mesh", "", "DirectX X file|*.x", BOX_SAVE, Wall3D.BoiteMur3D)
    If File$ <> "" Then
        Call DxWall(No_Mur).WallMeshBuilder.Save(File$, D3DRMXOF_TEXT, D3DRMXOFSAVE_ALL + D3DRMXOFSAVE_TEXTURETOPOLOGY)
    End If
End Sub

'
' ****************************
' Load a meshbuilder from disk
' ****************************
'
Private Sub LoadX_Click()
    Dim File$
    Dim Folder$
    File$ = Open_Box$("Load X Mesh", "", "DirectX X file|*.x", BOX_LOAD, Wall3D.BoiteMur3D, Folder$)
    If File$ <> "" Then
        Set DxWall(No_Mur.Value).WallMeshBuilder = dxEngine.Load_MeshBuilder(File$)
        DefWall(No_Mur.Value).Faces = True
    Else
        DefWall(No_Mur.Value).TheFile$ = ""
        Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    End If
    DefWall(No_Mur.Value).TheFile$ = File$
    Wall3D.Fichier_X.Caption = DefWall(No_Mur).TheFile$
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' *********************
' Initialise parameters
' *********************
'
Private Sub Form_Load()
    '
    ' ***** Create the Direct3D Retained Mode object
    '
    Call dxEngine.Create_3DRM(View.shot, View.shot.ScaleWidth, View.shot.ScaleHeight, Mode3DSurface)
    Call InitScene
    Call Axe_Camera.Show
    Call Axe_Camera.Init
    Call Axe_Camera.Position
    Call Transform.Show
    Call dxEngine.dxScene.SetSceneBackgroundRGB(ColorBack.Red / 255, ColorBack.Green / 255, ColorBack.Blue / 255)
    MaterialColor(0).ShowAlpha = False
    MaterialColor(1).ShowAlpha = False
    MaterialColor(2).ShowAlpha = True
End Sub

'
' ********************************
' Select the number of side 3 to 4
' ********************************
'
Private Sub Nb_Cote_Click()
    If Nb_Cote.Value = vbChecked Then
        DefFace(No_Face).NbPoint = 4
    Else
        DefFace(No_Face).NbPoint = 3
    End If
    Call Show_Face
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' ********************
' Change the edit face
' ********************
'
Private Sub No_Face_Change()
    Call Show_Face
End Sub

'
' ***********************
' Change edit meshbuilder
' ***********************
'
Public Sub No_Mur_Change()
    Call ReDraw(1, True)
    LNo_Mur.Caption = No_Mur.Value
    Call Find_Faces
    No_Face.Value = No_Face.MIN
    Call Show_Face
    Call UpdateValue
    Call ReDraw(1, True)
End Sub

'
' **********************************
' Show all wall from meshbuilders
' n = 0: Load mesh, frame and render
' n = 1: Unload mesh and frame
' FlipLock : Lock render update
' **********************************
'
Public Sub ReDraw(n%, FlipLock As Boolean)
    Dim i%
    Static StartTime& ' Time render calculate
    Static NoRefresh As Boolean ' Lock the render when selecting by mouse
    Static LevelRefresh As Integer ' Level to unlock
    '
    If FlipLock = True Then
        If n% >= LevelRefresh Then
            NoRefresh = Not NoRefresh
            If NoRefresh = False Then
                LevelRefresh = 0
            Else
                LevelRefresh = n%
            End If
        End If
        Exit Sub
    End If
    If NoRefresh = True Then
        Exit Sub
    End If
    '
    If n% = 0 Then
        StartTime& = timeGetTime
        '
        If Axe_Camera.ShowAxes.Value = vbChecked Then
            Call DxWall(0).WallFrame.AddVisual(DxWall(0).WallMeshBuilder)
            Call dxEngine.dxScene.AddChild(DxWall(0).WallFrame)
            Call DxWall(0).WallFrame.AddScale(D3DRMCOMBINE_REPLACE, _
            TheCamera.range / 200, TheCamera.range / 200, TheCamera.range / 200)
        End If
        For i% = 1 To NbWall%
            If DefWall(i%).Faces = True Then
                '
                Set DxWall(i%).WallShadow = dxEngine.dxD3Drm.CreateShadow(DxWall(i%).WallMeshBuilder, dxShadowLight, 0, 800, 0, 0, 1, 0)
                Call DxWall(i%).WallShadow.SetOptions(D3DRMSHADOW_TRUEALPHA)
                '
                Call DxWall(i%).WallFrame.AddVisual(DxWall(i%).WallMeshBuilder)
                Call DxWall(i%).WallFrame.AddVisual(DxWall(i%).WallShadow)
                '
                ' ***** Replace mesh
                If DefWall(i%).ParentNo% = 0 Then
                    Call dxEngine.dxScene.AddChild(DxWall(i%).WallFrame)
                    Call DxWall(i%).WallFrame.SetPosition(dxEngine.dxScene, 0, 0, 0)
                    Call DxWall(i%).WallFrame.SetOrientation(dxEngine.dxScene, 0, 0, 1, 0, 1, 0)
                Else
                    Call DxWall(DefWall(i%).ParentNo%).WallFrame.AddChild(DxWall(i%).WallFrame)
                    Call DxWall(i%).WallFrame.SetPosition(DxWall(DefWall(i%).ParentNo%).WallFrame, 0, 0, 0)
                    Call DxWall(i%).WallFrame.SetOrientation(DxWall(DefWall(i%).ParentNo%).WallFrame, 0, 0, 1, 0, 1, 0)
                End If
                ' ***** Transform mesh
                Call DxWall(i%).WallFrame.AddScale(D3DRMCOMBINE_AFTER, Animation&(i%, 6, NoAnimation%) / 10 ^ 5, Animation&(i%, 7, NoAnimation%) / 10 ^ 5, Animation&(i%, 8, NoAnimation%) / 10 ^ 5)
                Call DxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 1, 0, Animation&(i%, 3, NoAnimation%) / 10 ^ 5 / 180 * PI!)
                Call DxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 1, 0, 0, Animation&(i%, 4, NoAnimation%) / 10 ^ 5 / 180 * PI!)
                Call DxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 0, 1, Animation&(i%, 5, NoAnimation%) / 10 ^ 5 / 180 * PI!)
                Call DxWall(i%).WallFrame.AddTranslation(D3DRMCOMBINE_AFTER, Animation&(i%, 0, NoAnimation%) / 10 ^ 5, Animation&(i%, 1, NoAnimation%) / 10 ^ 5, Animation&(i%, 2, NoAnimation%) / 10 ^ 5)
                Call DxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 1, 0, Animation&(i%, 9, NoAnimation%) / 10 ^ 5 / 180 * PI!)
                Call DxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 1, 0, 0, Animation&(i%, 10, NoAnimation%) / 10 ^ 5 / 180 * PI!)
                Call DxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 0, 1, Animation&(i%, 11, NoAnimation%) / 10 ^ 5 / 180 * PI!)
            End If
        Next i%
        Call dxEngine.Render(False)
        Call dxEngine.Render(True)
    Else
        If Axe_Camera.ShowAxes.Value = vbChecked Then
            Call dxEngine.dxScene.DeleteChild(DxWall(0).WallFrame)
            Call DxWall(0).WallFrame.DeleteVisual(DxWall(0).WallMeshBuilder)
        End If
        For i% = 1 To NbWall%
            If DefWall(i%).Faces = True Then
                If DefWall(i%).ParentNo% = 0 Then
                    Call dxEngine.dxScene.DeleteChild(DxWall(i%).WallFrame)
                Else
                    Call DxWall(DefWall(i%).ParentNo%).WallFrame.DeleteChild(DxWall(i%).WallFrame)
                End If
                '
                Call DxWall(i%).WallFrame.DeleteVisual(DxWall(i%).WallShadow)
                Set DxWall(i%).WallShadow = Nothing
                '
                Call DxWall(i%).WallFrame.DeleteVisual(DxWall(i%).WallMeshBuilder)
            End If
        Next i%
        '
        View.Caption = "View:" + Format$(timeGetTime - StartTime&) + " ms" + "[" + Format$(View.ScaleWidth) + "/" + Format$(View.ScaleHeight) + "]"
    End If
End Sub

'
' *********************************
' Show the face and update controls
' *********************************
'
Public Sub Show_Face()
    Dim i%
    If No_Face.MIN = 0 Then
        No_Face.Visible = False
        LNo_Face.Visible = False
        FaceColor.Visible = False
        Nb_Cote.Visible = False
        CTexture(1).Visible = False
        For i% = 0 To 3
            LPoint(i%).Visible = False
            Point_X(i%).Visible = False
            Point_Y(i%).Visible = False
            Point_Z(i%).Visible = False
        Next i%
        ColorAll.Visible = False
        FaceRotate.Visible = False
        FaceInvertion.Visible = False
    Else
        No_Face.Visible = True
        LNo_Face.Visible = True
        LNo_Face = No_Face
        CTexture(1).Visible = True
        CTexture(1).ListIndex = DefFace(No_Face).Texture
        FaceColor.Visible = True
        FaceColor.Red = DefFace(No_Face).Color.Red
        FaceColor.Green = DefFace(No_Face).Color.Green
        FaceColor.Blue = DefFace(No_Face).Color.Blue
        FaceColor.Alpha = DefFace(No_Face).Color.flags
        Nb_Cote.Visible = True
        ColorAll.Visible = True
        For i% = 0 To 3
            If i% = 3 And DefFace(No_Face).NbPoint = 3 Then
                LPoint(i%).Visible = False
                Point_X(i%).Visible = False
                Point_Y(i%).Visible = False
                Point_Z(i%).Visible = False
            Else
                LPoint(i%).Visible = True
                Point_X(i%).Visible = True
                Point_Y(i%).Visible = True
                Point_Z(i%).Visible = True
                Point_X(i%).Text = DefFace(No_Face).Point(i%).X
                Point_Y(i%).Text = DefFace(No_Face).Point(i%).Y
                Point_Z(i%).Text = DefFace(No_Face).Point(i%).z
            End If
        Next i%
        If DefFace(No_Face).NbPoint = 3 Then
            Nb_Cote.Value = vbUnchecked
        Else
            Nb_Cote.Value = vbChecked
        End If
        FaceRotate.Visible = True
        FaceInvertion.Visible = True
    End If
End Sub

'
' **************************
' Change X value for a point
' **************************
'
Private Sub Point_X_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    DefFace(No_Face).Point(Index).X = Val(Point_X(Index).Text)
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' **************************
' Change Y value for a point
' **************************
'
Private Sub Point_Y_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    DefFace(No_Face).Point(Index).Y = Val(Point_Y(Index).Text)
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' **************************
' Change Z value for a point
' **************************
'
Private Sub Point_Z_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    DefFace(No_Face).Point(Index).z = Val(Point_Z(Index).Text)
    Call Build_Wall(dxEngine, No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' ******************************
' Looking for the first and last
' face for a wall
' ******************************
'
Public Sub Find_Faces()
    Dim a%, b%
    Call Init.Find_List(FaceIndex%(), No_Mur, 1, UBound(DefFace), a%, b%)
    No_Face.MIN = a%
    No_Face.MAX = b%
End Sub

'
' **********************************************
' Change the position or dimension of the object
' **********************************************
'
Private Sub TPosition_Click(Index As Integer)
    Dim r$
    r$ = InputBox("New value", , Animation&(Wall3D.No_Mur, Index, NoAnimation%) / 10 ^ 5)
    If r$ = "" Then Exit Sub
    Call Modify_Attibute(Index, Val(r$) * 10 ^ 5)
End Sub

Private Sub Wrap_Nx_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    DefWall(No_Mur.Value).WrapNx% = Val(Wrap_Nx)
    Call Build_Wall(dxEngine, No_Mur.Value, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

Private Sub Wrap_Ny_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    DefWall(No_Mur.Value).WrapNy% = Val(Wrap_Ny)
    Call Build_Wall(dxEngine, No_Mur.Value, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call ReDraw(0, False)
    Call ReDraw(1, False)
End Sub

'
' ***********************************
' Update the view after modify a mesh
' ***********************************
'
Public Sub Update()
    Call Wall3D.Find_Faces
    Wall3D.No_Face.Value = Wall3D.No_Face.MIN
    Call Wall3D.Show_Face
    Call Build_Wall(dxEngine, Wall3D.No_Mur, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' ***********************************
' Initialize default value at startup
' ***********************************
'
Public Sub InitScene()
    Dim i%
    '
    ' First dimension the animation array
    '
    NbAnimation% = 0
    ReDim Animation&(NbWall%, 11, NbAnimation%) ' No Frame
    '
    For i% = 0 To NbWall%
        Set DxWall(i%).WallFrame = dxEngine.dxD3Drm.CreateFrame(Nothing)
        Call DxWall(i%).WallFrame.SetAppData(i%) ' Set the number for this frame (used in frame detection)
        DefWall(i%).FrameName$ = "Frame " + Format$(i%)
        Set DxWall(i%).WallMeshBuilder = dxEngine.dxD3Drm.CreateMeshBuilder
        Animation&(i%, 0, 0) = 0
        Animation&(i%, 1, 0) = 0
        Animation&(i%, 2, 0) = 0
        Animation&(i%, 3, 0) = 0
        Animation&(i%, 4, 0) = 0
        Animation&(i%, 5, 0) = 0
        Animation&(i%, 6, 0) = 10 ^ 5 ' Scale factor = 1
        Animation&(i%, 7, 0) = 10 ^ 5 ' Scale factor = 1
        Animation&(i%, 8, 0) = 10 ^ 5 ' Scale factor = 1
        Animation&(i%, 9, 0) = 0
        Animation&(i%, 10, 0) = 0
        Animation&(i%, 11, 0) = 0
        DefWall(i%).WrapType = D3DRMWRAP_FLAT
        DefWall(i%).WrapNx% = 1
        DefWall(i%).WrapNy% = 1
        DefWall(i%).WrapTexture% = 0
        DefWall(i%).ParentNo% = 0
        DefWall(i%).TheFile$ = ""
        DefWall(i%).Faces = False
        DefWall(i%).FrameName$ = "Frame" + Str$(i%)
    Next i%
    '
    ' ***** Default value
    ' ***** at load time
    '
    No_Mur.MAX = NbWall%
    No_Mur.MIN = 0
    No_Mur.Value = 0
    ReDim FaceIndex%(0)
    ReDim DefFace(0) As TypeMur3D
    Call Create_Axes
End Sub

'
' **************************
' Add a texture in the combo
' **************************
'
Public Sub Add_Texture(FileName$, Position%)
    Dim ShortName$, i%
    Dim X%, Y%
    ShortName$ = Right$(FileName$, Len(FileName$) - InStrRev(FileName$, "\"))
    i% = InStrRev(ShortName$, ".")
    If i% <> 0 Then ShortName$ = Left$(ShortName$, i% - 1)
    For i% = 0 To 1
        If Position% + 1 > CTexture(i%).ListCount Then
            CTexture(i%).AddItem ShortName$
        Else
            CTexture(i%).RemoveItem Position%
            CTexture(i%).AddItem ShortName$, Position%
        End If
    Next i%
    '
    If Position% <> 0 Then
        X% = ((Position% - 1) Mod 16) * 16
        Y% = ((Position% - 1) \ 16) * 16
        If FileName$ <> "" Then
            If Exist(FileName$) = True Then
                Textures.Image1.Picture = LoadPicture(FileName$)
                Call Textures.MiniTextures.PaintPicture(Textures.Image1, X%, Y%, 16, 16)
            Else
                Textures.MiniTextures.Line (X%, Y%)-(X% + 15, Y% + 15), RGB(255, 255, 255), BF
            End If
        Else
            Textures.MiniTextures.Line (X%, Y%)-(X% + 15, Y% + 15), RGB(255, 255, 255), BF
        End If
    End If
End Sub

'
' ****************************
' Update mesh value on the box
' ****************************
'
Public Sub UpdateValue()
    Dim i%
    If No_Mur = 0 Then
        FFace.Visible = False
        FMesh.Visible = False
        FPosition.Visible = False
        FWrap.Visible = False
        FTexture.Visible = False
        MainWindow.MENU_Standard.Enabled = False
        MainWindow.MENU_Shape.Enabled = False
        MainWindow.MENU_Faces.Enabled = False
        MainWindow.MENU_Copy.Enabled = False
        MainWindow.MENU_Cut.Enabled = False
        LName.Visible = False
    Else
        FFace.Visible = True
        FMesh.Visible = True
        FPosition.Visible = True
        FWrap.Visible = True
        FTexture.Visible = True
        MainWindow.MENU_Standard.Enabled = True
        MainWindow.MENU_Shape.Enabled = True
        MainWindow.MENU_Faces.Enabled = True
        MainWindow.MENU_Copy.Enabled = True
        MainWindow.MENU_Cut.Enabled = True
        LName.Visible = True
        '
        Wall3D.Fichier_X.Caption = DefWall(No_Mur).TheFile$
        For i% = 0 To 11
            Wall3D.TPosition(i%).Caption = Animation&(No_Mur, i%, NoAnimation%) / 10 ^ 5
        Next i%
        CTexture(0).ListIndex = DefWall(No_Mur).WrapTexture%
        Wrap_Nx = DefWall(No_Mur).WrapNx%
        Wrap_Ny = DefWall(No_Mur).WrapNy%
        Choix_Wrap(DefWall(No_Mur).WrapType) = True
        No_Parent = DefWall(No_Mur).ParentNo%
        LName = DefWall(No_Mur).FrameName$
    End If
End Sub

'
' ***************************************
' Modify attribut dimension for the frame
' ***************************************
'
Public Sub Modify_Attibute(n%, v&)
    Animation&(Wall3D.No_Mur, n%, NoAnimation%) = v&
    If (n% >= 3 And n% <= 5) Or n% >= 9 Then ' For rotation only
        Animation&(Wall3D.No_Mur, n%, NoAnimation%) = Animation&(Wall3D.No_Mur, n%, NoAnimation%) Mod 360 * 10 ^ 5
    End If
    Wall3D.TPosition(n%) = Animation&(Wall3D.No_Mur, n%, NoAnimation%) / 10 ^ 5
    DoEvents
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

