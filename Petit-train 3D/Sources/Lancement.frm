VERSION 5.00
Begin VB.Form Lancement 
   BackColor       =   &H00408000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Lancement.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00408000&
      Height          =   4875
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7905
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Compagnie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   3360
         Width           =   2535
      End
      Begin VB.Label lblWarning 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "This's a Beta testing version and can't be sold."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   4560
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows 95"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   6
         Top             =   2280
         Width           =   1830
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4680
         TabIndex        =   8
         Top             =   1320
         Width           =   2430
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Licence accordée à"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7695
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J-Michel Kerdal present"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   4065
      End
      Begin VB.Image imgLogo 
         Height          =   4560
         Left            =   120
         Top             =   240
         Width           =   3840
      End
   End
End
Attribute VB_Name = "Lancement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' *********************
' Fenêtre de chargement
' *********************
'
Option Explicit

'
' ******************
' Fin de l'affichage
' ******************
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

'
' ***************
' Maj de la fiche
' ***************
'
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCompany.Caption = App.CompanyName
    lblCopyright.Caption = App.LegalCopyright
    Me.imgLogo.Picture = LoadPicture(".\Petit-train 3D\Textures\BBcote.jpg")
End Sub

'
' ******************
' Fin de l'affichage
' ******************
'
Private Sub Frame1_Click()
    Unload Me
End Sub

