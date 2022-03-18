VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SaisieRail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saisie d'une voie"
   ClientHeight    =   6255
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7035
   Icon            =   "Saisie rail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "Param"
      TabPicture(0)   =   "Saisie rail.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FPoint"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FSegment"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TLibelle(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CLibelle(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CLibelle(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TLibelle(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TLibelle(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Connect"
      TabPicture(1)   =   "Saisie rail.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FConnecte"
      Tab(1).Control(1)=   "FAiguille"
      Tab(1).Control(2)=   "FCatenaire"
      Tab(1).Control(3)=   "FMobile"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Vue"
      TabPicture(2)   =   "Saisie rail.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Affiche"
      Tab(2).ControlCount=   1
      Begin VB.TextBox TLibelle 
         Height          =   285
         Index           =   2
         Left            =   6000
         TabIndex        =   202
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TLibelle 
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   201
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox CLibelle 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   200
         Text            =   "Combo1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox CLibelle 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   199
         Text            =   "Combo1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TLibelle 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   198
         Top             =   360
         Width           =   5775
      End
      Begin VB.Frame FMobile 
         Caption         =   "Test mobile"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   172
         Top             =   4200
         Width           =   4335
         Begin VB.TextBox TDeplace 
            Height          =   285
            Left            =   1200
            TabIndex        =   176
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox TEntree 
            Height          =   285
            Left            =   1200
            TabIndex        =   175
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   375
            Left            =   1920
            TabIndex        =   174
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Position 
            Caption         =   "Bascule aiguillage"
            Height          =   255
            Left            =   120
            TabIndex        =   173
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Avancement"
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Point d'entrée"
            Height          =   255
            Left            =   120
            TabIndex        =   177
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame FCatenaire 
         Caption         =   "Positions catenaires"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   152
         Top             =   2880
         Width           =   6855
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   7
            Left            =   6120
            TabIndex        =   194
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   6
            Left            =   5400
            TabIndex        =   193
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   5
            Left            =   4680
            TabIndex        =   192
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   4
            Left            =   3960
            TabIndex        =   191
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   7
            Left            =   6120
            TabIndex        =   190
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   6
            Left            =   5400
            TabIndex        =   189
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   5
            Left            =   4680
            TabIndex        =   188
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   4
            Left            =   3960
            TabIndex        =   187
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   7
            Left            =   6120
            TabIndex        =   183
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   6
            Left            =   5400
            TabIndex        =   182
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   5
            Left            =   4680
            TabIndex        =   181
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   4
            Left            =   3960
            TabIndex        =   180
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   3
            Left            =   3240
            TabIndex        =   164
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   163
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   162
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TSCatenaire 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   161
            Text            =   "Text1"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   3
            Left            =   3240
            TabIndex        =   160
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   159
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   158
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TPCatenaire 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   157
            Text            =   "Text1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   3
            Left            =   3240
            TabIndex        =   156
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   155
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   154
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TNCatenaire 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   153
            Text            =   "Text1"
            Top             =   480
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°8"
            Height          =   255
            Index           =   10
            Left            =   6120
            TabIndex        =   186
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°7"
            Height          =   255
            Index           =   9
            Left            =   5400
            TabIndex        =   185
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°6"
            Height          =   255
            Index           =   8
            Left            =   4680
            TabIndex        =   184
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°5"
            Height          =   255
            Index           =   7
            Left            =   3960
            TabIndex        =   179
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Caption         =   "Sens"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   171
            Top             =   960
            Width           =   855
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°4"
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   170
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°3"
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   169
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°2"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   168
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Alignment       =   2  'Center
            Caption         =   "N°1"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   167
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LCatenaire 
            Caption         =   "Position"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   166
            Top             =   720
            Width           =   855
         End
         Begin VB.Label LCatenaire 
            Caption         =   "N°Segment"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   165
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame FAiguille 
         Caption         =   "Aiguillages"
         Height          =   2535
         Left            =   -72480
         TabIndex        =   138
         Top             =   360
         Width           =   2415
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   63
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   336
            Text            =   "X"
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   62
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   335
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   61
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   334
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   60
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   333
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   59
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   332
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   58
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   331
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   57
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   330
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   56
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   329
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   55
            Left            =   2040
            TabIndex        =   328
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   54
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   327
            Text            =   "X"
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   53
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   326
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   52
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   325
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   51
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   324
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   50
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   323
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   49
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   322
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   48
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   321
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   47
            Left            =   2040
            TabIndex        =   320
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   46
            Left            =   1800
            TabIndex        =   319
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   45
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   318
            Text            =   "X"
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   44
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   317
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   43
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   316
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   42
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   315
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   41
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   314
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   40
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   313
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   39
            Left            =   2040
            TabIndex        =   312
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   38
            Left            =   1800
            TabIndex        =   311
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   37
            Left            =   1560
            TabIndex        =   310
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   36
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   309
            Text            =   "X"
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   35
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   308
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   34
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   307
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   33
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   306
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   32
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   305
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   31
            Left            =   2040
            TabIndex        =   304
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   30
            Left            =   1800
            TabIndex        =   303
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   29
            Left            =   1560
            TabIndex        =   302
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   28
            Left            =   1320
            TabIndex        =   301
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   27
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   300
            Text            =   "X"
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   26
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   299
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   25
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   298
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   24
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   297
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   23
            Left            =   2040
            TabIndex        =   296
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   22
            Left            =   1800
            TabIndex        =   295
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   21
            Left            =   1560
            TabIndex        =   294
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   20
            Left            =   1320
            TabIndex        =   293
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   19
            Left            =   1080
            TabIndex        =   292
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   291
            Text            =   "X"
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   17
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   290
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   16
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   289
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   15
            Left            =   2040
            TabIndex        =   288
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   1800
            TabIndex        =   287
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   1560
            TabIndex        =   286
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   1320
            TabIndex        =   285
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   1080
            TabIndex        =   284
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   840
            TabIndex        =   283
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   282
            Text            =   "X"
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   2040
            TabIndex        =   281
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   280
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   279
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   278
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   8
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   143
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   142
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   141
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   600
            TabIndex        =   140
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatAiguille 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   139
            Text            =   "X"
            Top             =   480
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "H"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   344
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "G"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   343
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "F"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   342
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "E"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   341
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "H"
            Height          =   255
            Index           =   11
            Left            =   2040
            TabIndex        =   340
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "G"
            Height          =   255
            Index           =   10
            Left            =   1800
            TabIndex        =   339
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "F"
            Height          =   255
            Index           =   9
            Left            =   1560
            TabIndex        =   338
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "E"
            Height          =   255
            Index           =   8
            Left            =   1320
            TabIndex        =   337
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   151
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "C"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   150
            Top             =   960
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "B"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   149
            Top             =   720
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "A"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   148
            Top             =   480
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   147
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "C"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   146
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "B"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   145
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LAiguille 
            Alignment       =   2  'Center
            Caption         =   "A"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   144
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame FConnecte 
         Caption         =   "Connections"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   124
         Top             =   360
         Width           =   2415
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   63
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   273
            Text            =   "X"
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   62
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   272
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   61
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   271
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   60
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   270
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   59
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   269
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   58
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   268
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   57
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   267
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   56
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   266
            Top             =   2160
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   55
            Left            =   2040
            TabIndex        =   265
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   54
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   264
            Text            =   "X"
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   53
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   263
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   52
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   262
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   51
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   261
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   50
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   260
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   49
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   259
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   48
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   258
            Top             =   1920
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   47
            Left            =   2040
            TabIndex        =   257
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   46
            Left            =   1800
            TabIndex        =   256
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   45
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   255
            Text            =   "X"
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   44
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   254
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   43
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   253
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   42
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   252
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   41
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   251
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   40
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   250
            Top             =   1680
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   39
            Left            =   2040
            TabIndex        =   249
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   38
            Left            =   1800
            TabIndex        =   248
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   37
            Left            =   1560
            TabIndex        =   247
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   36
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   246
            Text            =   "X"
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   35
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   245
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   34
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   244
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   33
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   243
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   32
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   242
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   31
            Left            =   2040
            TabIndex        =   241
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   30
            Left            =   1800
            TabIndex        =   240
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   29
            Left            =   1560
            TabIndex        =   239
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   28
            Left            =   1320
            TabIndex        =   238
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   27
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   237
            Text            =   "X"
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   26
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   236
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   25
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   235
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   24
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   234
            Top             =   1200
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   23
            Left            =   2040
            TabIndex        =   233
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   22
            Left            =   1800
            TabIndex        =   232
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   21
            Left            =   1560
            TabIndex        =   231
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   20
            Left            =   1320
            TabIndex        =   230
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   19
            Left            =   1080
            TabIndex        =   229
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   228
            Text            =   "X"
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   17
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   227
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   16
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   226
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   15
            Left            =   2040
            TabIndex        =   225
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   14
            Left            =   1800
            TabIndex        =   224
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   13
            Left            =   1560
            TabIndex        =   223
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   12
            Left            =   1320
            TabIndex        =   222
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   11
            Left            =   1080
            TabIndex        =   221
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   10
            Left            =   840
            TabIndex        =   220
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   219
            Text            =   "X"
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   7
            Left            =   2040
            TabIndex        =   218
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   217
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   216
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   215
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            Height          =   285
            Index           =   8
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   129
            Top             =   720
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   128
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   127
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   600
            TabIndex        =   126
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox MatConnecte 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "X"
            Top             =   480
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "H"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   277
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "G"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   276
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "F"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   275
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "E"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   274
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "H"
            Height          =   255
            Index           =   11
            Left            =   2040
            TabIndex        =   214
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "G"
            Height          =   255
            Index           =   10
            Left            =   1800
            TabIndex        =   213
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "F"
            Height          =   255
            Index           =   9
            Left            =   1560
            TabIndex        =   212
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "E"
            Height          =   255
            Index           =   8
            Left            =   1320
            TabIndex        =   211
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   137
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "C"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   136
            Top             =   960
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "B"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   135
            Top             =   720
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "A"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   134
            Top             =   480
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   133
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "C"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   132
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "B"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   131
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LMatrice 
            Alignment       =   2  'Center
            Caption         =   "A"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   130
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.PictureBox Affiche 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   -74880
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   123
         Top             =   360
         Width           =   6855
      End
      Begin VB.Frame FSegment 
         Caption         =   "Segments de voie"
         Height          =   1935
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   6735
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   4560
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   95
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   7
            Left            =   6000
            TabIndex        =   94
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   6
            Left            =   5280
            TabIndex        =   93
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   5
            Left            =   4560
            TabIndex        =   92
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   4
            Left            =   3840
            TabIndex        =   91
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   7
            Left            =   6000
            TabIndex        =   90
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   6
            Left            =   5280
            TabIndex        =   89
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   5
            Left            =   4560
            TabIndex        =   88
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   4
            Left            =   3840
            TabIndex        =   87
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   7
            Left            =   6000
            TabIndex        =   86
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   6
            Left            =   5280
            TabIndex        =   85
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   5
            Left            =   4560
            TabIndex        =   84
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   4
            Left            =   3840
            TabIndex        =   83
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   7
            Left            =   6000
            TabIndex        =   78
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   6
            Left            =   5280
            TabIndex        =   77
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   5
            Left            =   4560
            TabIndex        =   76
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   4
            Left            =   3840
            TabIndex        =   75
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   3
            Left            =   3120
            TabIndex        =   65
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   3
            Left            =   3120
            TabIndex        =   64
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   3
            Left            =   3120
            TabIndex        =   63
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   3
            Left            =   3120
            TabIndex        =   62
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   61
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   60
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   59
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   58
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   57
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   56
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   55
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   54
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TRotation 
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   53
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox TAngle 
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   52
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TRayon 
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   51
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TLongueur 
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   50
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox TTaille 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°8"
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   82
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°7"
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   81
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°6"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   80
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°5"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   79
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Longueur"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Rayon"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   73
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Angle"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Taille"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   71
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Rotation"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   70
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°1"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   69
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°2"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   68
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°3"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   67
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LSegment 
            Alignment       =   2  'Center
            Caption         =   "N°4"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame FPoint 
         Caption         =   "Définition des points"
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   6735
         Begin VB.CheckBox MatJonction 
            Caption         =   "JH"
            Height          =   195
            Index           =   7
            Left            =   6000
            TabIndex        =   122
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox MatJonction 
            Caption         =   "JG"
            Height          =   195
            Index           =   6
            Left            =   5280
            TabIndex        =   121
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox MatJonction 
            Caption         =   "JF"
            Height          =   195
            Index           =   5
            Left            =   4560
            TabIndex        =   120
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox MatJonction 
            Caption         =   "JE"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   119
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TH"
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   118
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TG"
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   117
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TF"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   116
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TE"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   115
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   6000
            TabIndex        =   114
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   5280
            TabIndex        =   113
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   4560
            TabIndex        =   112
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3840
            TabIndex        =   111
            Top             =   960
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PH"
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   108
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PG"
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   107
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PF"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   106
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   6000
            TabIndex        =   105
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   5280
            TabIndex        =   104
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   4560
            TabIndex        =   103
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3840
            TabIndex        =   102
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PE"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   101
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox MatJonction 
            Caption         =   "JD"
            Height          =   195
            Index           =   3
            Left            =   3120
            TabIndex        =   27
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox MatJonction 
            Caption         =   "JC"
            Height          =   195
            Index           =   2
            Left            =   2400
            TabIndex        =   26
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox MatJonction 
            Caption         =   "JB"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   25
            Top             =   2160
            Width           =   615
         End
         Begin VB.CheckBox MatJonction 
            Caption         =   "JA"
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   24
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TD"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   23
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TC"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   22
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TB"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   21
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox MatTerminaison 
            Caption         =   "TA"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   20
            Top             =   1800
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   18
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   17
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TY 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3120
            TabIndex        =   16
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   15
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   14
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   13
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox TX 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3120
            TabIndex        =   12
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PD"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   11
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PC"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   10
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PB"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   9
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox CPoint 
            Caption         =   "PA"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   8
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   210
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   209
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   208
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   207
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   206
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   205
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   204
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   203
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "H"
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   110
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "G"
            Height          =   255
            Index           =   6
            Left            =   5280
            TabIndex        =   109
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "F"
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   100
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "E"
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   99
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Jonction"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   44
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Terminaison"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   43
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   42
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   41
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   40
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_Y 
            Caption         =   "Y"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   39
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   38
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   37
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   36
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label D_X 
            Caption         =   "X"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   35
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Offset"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   34
            Top             =   480
            Width           =   495
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "D"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   33
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "C"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "B"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LTerminaison 
            Alignment       =   2  'Center
            Caption         =   "A"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   30
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "DX"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   29
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "DY"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   28
            Top             =   960
            Width           =   255
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "mm"
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   197
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "°"
         Height          =   255
         Index           =   3
         Left            =   6720
         TabIndex        =   196
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Libellé:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   195
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.TextBox TRef 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Efface 
      Caption         =   "&Efface"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Ajoute 
      Caption         =   "&Ajoute"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
   End
   Begin VB.HScrollBar Choix 
      Height          =   255
      Left            =   0
      Max             =   100
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Référence:"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label NoVoie 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   5760
      Width           =   375
   End
   Begin VB.Menu MENU_Fichier 
      Caption         =   "Fichier"
      Begin VB.Menu MENU_Charger 
         Caption         =   "&Charger"
      End
      Begin VB.Menu MENU_Enregistrer 
         Caption         =   "&Enregistrer"
      End
      Begin VB.Menu MENU_MoinsFichier 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Quitter 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "SaisieRail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Frezze As Boolean

'
' ********************************
' Dessine le rail avec les valeurs
' saisies par l'opérateur
' ********************************
'
Public Sub Dessine()
    Dim i%, j%, n%, sX%, sY%
    Dim Couleur As Long
    If Frezze = True Then Exit Sub
    '
    ' ***** Mise à jour des indicateurs pour information
    '
    For i% = 0 To NbSegment%
        TTaille(i%) = Voie(Choix.Value).Segment(i%).Segment_Taille()
    Next i%
    Call Initialisation.Recalcule_Point(Choix.Value)
    '
    ' ***** Affiche grille
    '
    sX% = 50
    sY% = 150
    Affiche.Cls
    Affiche.Line (sX%, 0)-(sX%, 300), vbCyan
    Affiche.Line (0, sY%)-(300, sY%), vbCyan
    '
    ' ***** Calcul des segments
    '
    For i% = 0 To NbSegment% - 1
        For j% = i% + 1 To NbSegment%
            If Voie(Choix.Value).MatConnecte(i%, j%) <> 0 Then
                Call Calcule_Point(Choix.Value, i%, j%)
                If Voie(Choix.Value).MatAiguille(i%, j%) = 1 Then
                    If Position.Value = 0 Then
                        Couleur& = vbWhite
                    Else
                        Couleur& = vbBlue
                    End If
                ElseIf Voie(Choix.Value).MatAiguille(i%, j%) = 2 Then
                    If Position.Value = 0 Then
                        Couleur& = vbBlue
                    Else
                        Couleur& = vbWhite
                    End If
                Else
                    Couleur& = vbWhite
                End If
                If Voie(Choix.Value).Offset(i%) = 0 Then
                    Call Dessine_Segment(Voie(Choix.Value).MatConnecte(i%, j%) - 1, _
                    sX% + Voie(Choix.Value).pX!(i%), _
                    sY% - Voie(Choix.Value).pZ!(i%), Couleur&)
                Else
                    Call Dessine_Segment(Voie(Choix.Value).MatConnecte(i%, j%) - 1, _
                    sX% + Voie(Choix.Value).dX!(i%), _
                    sY% - Voie(Choix.Value).dz!(i%), Couleur&)
                End If
            End If
        Next j%
        D_X(i%) = Format$(Voie(Choix.Value).pX!(i%), "#.##")
        D_Y(i%) = Format$(Voie(Choix.Value).pZ!(i%), "#.##")
    Next i%
    '
    ' ***** Affiche les points calculés
    '
    Dim c As Long
    For i% = 1 To NbSegment%
        If Voie(Choix.Value).Terminaison(i%) <> 0 Then
            c = vbRed
        Else
            c = vbGreen
        End If
        If Voie(Choix.Value).Offset(i%) = 0 Then
            If Voie(Choix.Value).pX!(i%) <> 0 Or Voie(Choix.Value).pZ!(i%) <> 0 Then
                Affiche.Line (sX% + Voie(Choix.Value).pX!(i%) - 3, sY% - Voie(Choix.Value).pZ!(i%) - 3)-Step(7, 7), c
                Affiche.Line (sX% + Voie(Choix.Value).pX!(i%) - 3, sY% - Voie(Choix.Value).pZ!(i%) + 3)-Step(7, -7), c
            End If
        Else
            If Voie(Choix.Value).dX!(i%) <> 0 Or Voie(Choix.Value).dz!(i%) <> 0 Then
                Affiche.Line (sX% + Voie(Choix.Value).dX!(i%) - 3, sY% - Voie(Choix.Value).dz!(i%) - 3)-Step(7, 7), c
                Affiche.Line (sX% + Voie(Choix.Value).dX!(i%) - 3, sY% - Voie(Choix.Value).dz!(i%) + 3)-Step(7, -7), c
            End If
        End If
    Next i%
    '
    ' ***** Fin calcul
    '
    Affiche.Refresh
End Sub

Public Sub Dessine_Vecteur(P As D3DVECTOR, v As D3DVECTOR)
    Affiche.Line (P.X, P.Y)-Step(v.X, v.Y), vbRed
    Affiche.Line (P.X, P.Y)-Step(1, 1), vbGreen
End Sub

'
' *****************************
' Ajoute une voie dans la liste
' *****************************
'
Private Sub Ajoute_Click()
    Dim n%
    n% = UBound(Voie()) + 1
    ReDim Preserve Voie(n%) As TypeVoie
    ReDim Preserve FormeVoie(n%) As New ClassCatenaire
    Choix.Max = n%
    Choix.Value = n%
End Sub

'
' *************************
' Séléctionne un autre rail
' *************************
'
Private Sub Choix_Change()
    Call MAJ
End Sub

'
' *****************
' Change le libellé
' *****************
'
Private Sub CLibelle_Click(Index As Integer)
    Voie(Choix.Value).Zone(Index) = CLibelle(Index).ListIndex
    Call MAJ
End Sub

'
' ******************************************
' Test de déplacement d'un point sur la voie
' ******************************************
'
Private Sub Command1_Click()
    'Dim Entree%, Dep!
    'Entree% = Val(TEntree)
    'Dep! = Val(TDeplace)
    'TestBogie.BogieReseau% = Choix.Value
    'If Entree% = 0 Then
    '    Call Voies.Deplace(TestBogie, Dep!)
    'Else
    '    TrainSegment% = Voies.Selectionne_Segment(Choix.Value, Entree%, RailPosition%, 0)
'Debug.Print "depart sur le segment"; TrainSegment%
    '    If TrainSegment% = 0 Then
    '        MsgBox "Entrée sur ce point invalide"
    '    Else
    '        If TrainSens% = 1 Then
    '            TrainPosition! = 0
    '        Else
    '            TrainPosition! = Voie(Choix.Value).Segment(TrainSegment% - 1).Segment_Taille()
    '        End If
    '        Call Voies.Deplace(TestBogie, Dep!)
    '    End If
    'End If
    'Call Dessine
End Sub

'
' ******************************
' Impose un offset pour un point
' ******************************
'
Private Sub CPoint_Click(Index As Integer)
    If CPoint(Index).Value = vbUnchecked Then
        TX(Index).Enabled = False
        TY(Index).Enabled = False
        Voie(Choix.Value).Offset(Index) = 0
    Else
        TX(Index).Enabled = True
        TY(Index).Enabled = True
        Voie(Choix.Value).Offset(Index) = 1
    End If
    Call Dessine
End Sub

'
' ***************************
' Efface une voie de la liste
' ***************************
'
Private Sub Efface_Click()
    Dim n%, i%
    If MsgBox(Localisation$(CleVOIE% + 19), vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Exit Sub
    n% = UBound(Voie()) - 1
    For i% = Choix.Value To n%
        Voie(i%) = Voie(i% + 1)
        Set FormeVoie(i%) = FormeVoie(i% + 1)
    Next i%
    ReDim Preserve Voie(n%) As TypeVoie
    ReDim Preserve FormeVoie(n%) As New ClassCatenaire
    Choix.Max = n%
    Choix.Value = 0
End Sub

'
' ***************************
' Initialisation de l'édition
' ***************************
'
Private Sub Form_Load()
    Dim i%, j%
    For j% = 0 To 1
        CLibelle(j%).Clear
        Call CLibelle(j%).AddItem("")
        For i% = CleRAIL% To UBound(Localisation$())
            Call CLibelle(j%).AddItem(Localisation$(i%))
        Next i%
    Next j%
    If Command$ <> Super$ Then FMobile.Visible = False
    Me.Caption = Localisation$(CleVOIE%)
    MENU_Fichier.Caption = Localisation$(1)
    MENU_Charger.Caption = Localisation$(3)
    MENU_Enregistrer.Caption = Localisation$(4)
    MENU_Quitter.Caption = Localisation$(7)
    FSegment.Caption = Localisation$(CleVOIE% + 1)
    Label1(0).Caption = Localisation$(CleVOIE% + 2)
    Label1(1).Caption = Localisation$(CleVOIE% + 3)
    Label1(2).Caption = Localisation$(CleVOIE% + 4)
    Label1(4).Caption = Localisation$(CleVOIE% + 5)
    Label1(3).Caption = Localisation$(CleVOIE% + 6)
    FPoint.Caption = Localisation$(CleVOIE% + 7)
    Label1(7).Caption = Localisation$(CleVOIE% + 8)
    Label1(8).Caption = Localisation$(CleVOIE% + 9)
    Label1(9).Caption = Localisation$(CleVOIE% + 10)
    FConnecte.Caption = Localisation$(CleVOIE% + 11)
    FAiguille.Caption = Localisation$(CleVOIE% + 12)
    FCatenaire.Caption = Localisation$(CleVOIE% + 13)
    LCatenaire(0).Caption = Localisation$(CleVOIE% + 14)
    LCatenaire(1).Caption = Localisation$(CleVOIE% + 15)
    LCatenaire(6).Caption = Localisation$(CleVOIE% + 16)
    Label4(0).Caption = Localisation$(CleVOIE% + 17)
    Label4(1).Caption = Localisation$(CleVOIE% + 18)
    Ajoute.Caption = Localisation$(17)
    Efface.Caption = Localisation$(18)
    '
    Choix.Max = UBound(Voie())
    Call MAJ
End Sub

'
' ***********************************************
' Change le lien suivant position de l'aiguillage
' 0: sans effet
' 1 ou 2: bascule suivant position
' ***********************************************
'
Private Sub MatAiguille_Change(Index As Integer)
    Dim l%, c%
    l% = Index \ (NbSegment% + 1)
    c% = Index Mod (NbSegment% + 1)
    Voie(Choix.Value).MatAiguille(l%, c%) = Val(MatAiguille(Index).Text)
    Voie(Choix.Value).MatAiguille(c%, l%) = Voie(Choix.Value).MatAiguille(l%, c%)
    MatAiguille(c% * (NbSegment% + 1) + l%) = Voie(Choix.Value).MatAiguille(l%, c%)
    Call Dessine
End Sub

'
' *****************************************
' Modifie la connection entre deux segments
' *****************************************
'
Private Sub MatConnecte_Change(Index As Integer)
    Dim l%, c%, v%
    l% = Index \ (NbSegment% + 1)
    c% = Index Mod (NbSegment% + 1)
    v% = Val(MatConnecte(Index).Text)
    If v% < 0 Or v% > (NbSegment% + 1) Then Exit Sub
    Voie(Choix.Value).MatConnecte(l%, c%) = v%
    Voie(Choix.Value).MatConnecte(c%, l%) = v%
    MatConnecte(c% * (NbSegment% + 1) + l%) = v%
    Call Dessine
End Sub

'
' *******************************************
' Ce point est une jonction, donc pas un bout
' *******************************************
'
Private Sub MatJonction_Click(Index As Integer)
    Voie(Choix.Value).Jonction(Index) = MatJonction(Index).Value
    Call Dessine
End Sub

'
' **********************************
' Indique une terminaison de segment
' **********************************
'
Private Sub MatTerminaison_Click(Index As Integer)
    Voie(Choix.Value).Terminaison(Index) = MatTerminaison(Index).Value
    Call Dessine
End Sub

'
' ******************************
' Charge la définition des rails
' ******************************
'
Private Sub MENU_Charger_Click()
    Dim Fichier$
    Fichier$ = Tools.Open_Box$(Localisation$(CleEDITION% + 8), CheminRail$, "All files *.rail|*.rail", BOX_LOAD, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    CheminRail$ = Fichier$
    Call Voies.Charger(CheminRail$)
End Sub

'
' *****************************
' Sauve la définition des rails
' *****************************
'
Private Sub MENU_Enregistrer_Click()
    Dim Fichier$
    Fichier$ = Tools.Open_Box$(Localisation$(CleEDITION% + 9), CheminRail$, "All files *.rail|*.rail", BOX_SAVE, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    CheminRail$ = Fichier$
    Call Voies.Sauver(CheminRail$)
End Sub

'
' **********************
' Sauve avant de quitter
' **********************
'
Private Sub MENU_Quitter_Click()
    Unload Me
End Sub

Private Sub Position_Click()
    Call Dessine
End Sub

Private Sub TAngle_Change(Index As Integer)
    Voie(Choix.Value).Segment(Index).Angle = Val(TAngle(Index))
    Call Dessine
End Sub

Private Sub TLongueur_Change(Index As Integer)
    Voie(Choix.Value).Segment(Index).Longueur = Val(TLongueur(Index))
    Call Dessine
End Sub

Private Sub TNCatenaire_Change(Index As Integer)
    Voie(Choix.Value).CatenaireSegment(Index) = Val(TNCatenaire(Index))
    Call Dessine
End Sub

Private Sub TPCatenaire_Change(Index As Integer)
    Voie(Choix.Value).CatenairePosition(Index) = Val(TPCatenaire(Index))
    Call Dessine
End Sub

Private Sub TRayon_Change(Index As Integer)
    Voie(Choix.Value).Segment(Index).Rayon = Val(TRayon(Index))
    Call Dessine
End Sub

Private Sub Tlibelle_Change(Index As Integer)
    Voie(Choix.Value).Libelle(Index) = TLibelle(Index)
    Call MAJ
End Sub

Private Sub TRef_Change()
    Voie(Choix.Value).Ref = TRef
End Sub

Private Sub TRotation_Change(Index As Integer)
    Voie(Choix.Value).Segment(Index).Rotation = Val(TRotation(Index))
    Call Dessine
End Sub

'
' ******************
' Dessine un segment
' ******************
'
Public Sub Dessine_Segment(n%, dX%, dY%, c&)
    Dim i%
    '
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    For i% = 1 To 10
        v2 = Voie(Choix.Value).Segment(n%).Point(i% / 10 * Voie(Choix.Value).Segment(n%).Segment_Taille())
        Affiche.Line (dX% + v1.X, dY% - v1.z)-(dX% + v2.X, dY% - v2.z), c&
        v1 = v2
    Next i%
    '
    ' ***** Pose la tangente
    '
    'If n% = TrainSegment% - 1 Then
    '    v1 = Voie(Choix.Value).Segment(n%).Tangente(TrainPosition!) ' / 100 * Voie(Choix.Value).Segment(n%).Segment_Taille())
    '    v1.x = v1.x * 10
    '    v1.y = -v1.z * 10
    '    v2 = Voie(Choix.Value).Segment(n%).Point(TrainPosition!) ' / 100 * Voie(Choix.Value).Segment(n%).Segment_Taille())
    '    v2.x = dX% + v2.x
    '    v2.y = dY% - v2.z
    '    Call Dessine_Vecteur(v2, v1)
    'End If
End Sub

Private Sub TSCatenaire_Change(Index As Integer)
    Voie(Choix.Value).CatenaireSens(Index) = Val(TSCatenaire(Index))
    Call Dessine
End Sub

Private Sub TX_Change(Index As Integer)
    Voie(Choix.Value).dX!(Index) = Val(TX(Index))
    Call Dessine
End Sub

Private Sub TY_Change(Index As Integer)
    Voie(Choix.Value).dz!(Index) = Val(TY(Index))
    Call Dessine
End Sub

'
' **********************************
' Mise à jour des valeurs d'une voie
' **********************************
'
Public Sub MAJ()
    Dim i%, j%
    Frezze = True
    NoVoie = Choix.Value
    If Choix.Value = 0 Then
        FSegment.Visible = False
        FPoint.Visible = False
        FConnecte.Visible = False
        FAiguille.Visible = False
        TLibelle(0).Visible = False
        TLibelle(1).Visible = False
        TLibelle(2).Visible = False
        CLibelle(0).Visible = False
        CLibelle(1).Visible = False
        TRef.Visible = False
        Efface.Visible = False
        FCatenaire.Visible = False
    Else
        FSegment.Visible = True
        FPoint.Visible = True
        FConnecte.Visible = True
        FAiguille.Visible = True
        TLibelle(0).Visible = True
        TLibelle(1).Visible = True
        TLibelle(2).Visible = True
        CLibelle(0).Visible = True
        CLibelle(1).Visible = True
        TRef.Visible = True
        Efface.Visible = True
        FCatenaire.Visible = True
        Call Voies.Libelle_Cree(Choix.Value)
        With Voie(Choix.Value)
            CLibelle(0).ListIndex = .Zone(0)
            CLibelle(1).ListIndex = .Zone(1)
            For i% = 0 To 2
                TLibelle(i%) = .Libelle$(i%)
            Next i%
            TRef = .Ref
            For i% = 0 To NbSegment%
                TLongueur(i%) = .Segment(i%).Longueur
                TRayon(i%) = .Segment(i%).Rayon
                TAngle(i%) = .Segment(i%).Angle
                TRotation(i%) = .Segment(i%).Rotation
                If .Offset(i%) = 0 Then
                    CPoint(i%) = vbUnchecked
                Else
                    CPoint(i%) = vbChecked
                End If
                TX(i%) = .dX!(i%)
                TY(i%) = .dz!(i%)
                D_X(i%) = Format$(.pX!(i%), "#.##")
                D_Y(i%) = Format$(.pZ!(i%), "#.##")
                MatTerminaison(i%) = .Terminaison(i%)
                MatJonction(i%) = .Jonction(i%)
                For j% = 0 To NbSegment%
                    If i% <> j% Then
                        MatConnecte(i% * (NbSegment% + 1) + j%) = .MatConnecte(i%, j%)
                        MatAiguille(i% * (NbSegment% + 1) + j%) = .MatAiguille(i%, j%)
                    End If
                Next j%
                TNCatenaire(i%) = .CatenaireSegment%(i%)
                TPCatenaire(i%) = .CatenairePosition!(i%)
                TSCatenaire(i%) = .CatenaireSens%(i%)
            Next i%
        End With
    End If
    Frezze = False
    Call Dessine
End Sub

