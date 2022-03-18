VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos de MonApplication"
   ClientHeight    =   3030
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5160
   ClipControls    =   0   'False
   Icon            =   "A propos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   120
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   0
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   1950
      Left            =   3720
      Top             =   120
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   6
      X2              =   336
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Kerdal Jean-Michel"
      ForeColor       =   &H00000000&
      Height          =   1275
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Titre de l'application"
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   336
      Y1              =   145
      Y2              =   145
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissement: ..."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ***********
' Window exit
' ***********
'
Private Sub cmdOK_Click()
    Unload Me
End Sub

'
' ************************
' Show an "About form"
' with Microsoft look like
' ************************
'
Public Sub About(TheForm As Form, Name$, Description$, Version$, Disclaim$, Optional Face$)
    Me.picIcon.Picture = TheForm.Icon
    Me.Caption = Name$
    Me.Icon = TheForm.Icon
    Me.lblTitle.Caption = App.Title
    Me.lblDescription = Description$
    Me.lblDisclaimer = Disclaim$
    Me.lblVersion.Caption = Version$ & " " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Icon = picIcon.Picture
    Me.Image1.Picture = LoadPicture(Face$)
    Call Me.Show(vbModal)
End Sub

