VERSION 5.00
Begin VB.Form Driver 
   BackColor       =   &H00408000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Driver"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   ControlBox      =   0   'False
   Icon            =   "Driver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListeRail 
      Height          =   1035
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox ModeTexture 
      BackColor       =   &H00408000&
      Caption         =   "Pas de texture dynamique"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      Top             =   3840
      Width           =   6615
   End
   Begin VB.ListBox ListeLangue 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.ListBox ListeDriver 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   6615
   End
   Begin VB.Image Drapeau 
      Height          =   495
      Left            =   4440
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Driver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ***********************************
' Liste des drivers vidéo disponibles
' ***********************************
'
Private Sub Form_Load()
    Dim Nom$, i%
    Me.Picture = LoadPicture(".\Petit-train 3D\Textures\MTS.jpg")
    If SansTextureDynamique = False Then
        ModeTexture.Value = vbUnchecked
    Else
        ModeTexture.Value = vbChecked
    End If
    Nom$ = Dir$(".\Petit-train 3D\Local\*.txt")
    While Nom$ <> ""
        Nom$ = Mid$(Nom$, 1, Len(Nom$) - 4)
        Call ListeLangue.AddItem(Nom$)
        If Nom$ = NomFichier$(4) Then
            ListeLangue.ListIndex = ListeLangue.ListCount - 1
        End If
        Nom$ = Dir$
    Wend
    If ListeLangue.ListCount = 0 Then
        Call MsgBox("No localisation language available", vbCritical + vbOKOnly)
        End
    End If
    If ListeLangue.ListIndex = -1 Then
        For i% = 0 To ListeLangue.ListCount - 1
            If LCase(ListeLangue.List(i%)) = "english" Then ListeLangue.ListIndex = i%
        Next i%
    End If
    If ListeLangue.ListIndex = -1 Then ListeLangue.ListIndex = 0
    '
    ' ***** Recherche les modèle de rail disponible
    '
    Nom$ = Dir$("Petit-train 3D\*.rail")
    Do While Nom$ <> ""
        Nom$ = Mid$(Nom$, 1, Len(Nom$) - 5)
        Call ListeRail.AddItem(Nom$)
        If Nom$ = NomFichier$(6) Then
            ListeRail.ListIndex = ListeRail.ListCount - 1
        End If
        Nom$ = Dir
    Loop
    If ListeRail.ListCount = 0 Then
        Call MsgBox(Localisation$(CleDRIVER% + 4), vbCritical + vbOKOnly)
        End
    End If
    '
    ' ***** Recherche les drivers DirectX
    '
    Call ListeDriver.Clear
    Call ListeDriver.AddItem("<" + Localisation$(CleDRIVER%) + ">")
    For i% = 1 To dxVue.dxEnum.GetCount()
        Call ListeDriver.AddItem(dxVue.dxEnum.GetDescription(i%))
        If dxVue.dxEnum.GetDescription(i%) = NomFichier$(5) Then
            ListeDriver.ListIndex = ListeDriver.ListCount - 1
        End If
    Next i%
    '
    Call Test_OK
End Sub

'
' *************************************
' Déplacement dans la liste des drivers
' *************************************
'
Private Sub ListeDriver_Click()
    Call Test_OK
End Sub

'
' ******************
' Choix d'une langue
' ******************
'
Private Sub ListeLangue_Click()
    Const PathLocal$ = ".\Petit-train 3D\Local\"
    If ListeLangue.ListIndex < 0 Then Exit Sub
    Call Initialisation.Local_Charge(PathLocal$ + ListeLangue.List(ListeLangue.ListIndex) + ".txt")
    Me.Caption = Localisation$(CleDRIVER% + 2)
    Me.OK.Caption = Localisation$(CleDRIVER% + 1)
    Me.ModeTexture.Caption = Localisation$(CleDRIVER% + 3)
    ListeDriver.List(0) = "<" + Localisation$(CleDRIVER%) + ">"
    If Tools.Exist(PathLocal$ + ListeLangue.List(ListeLangue.ListIndex) + ".gif") = True Then
        Drapeau.Picture = LoadPicture(PathLocal$ + ListeLangue.List(ListeLangue.ListIndex) + ".gif")
    Else
        Drapeau.Picture = LoadPicture()
    End If
    Call Test_OK
End Sub

'
' ***************************
' Séléction du modèle de voie
' ***************************
'
Private Sub ListeRail_Click()
    Call Test_OK
End Sub

'
' **************************************
' Valide l'option sans texture dynamique
' pour gérer les drivers défaillants
' **************************************
'
Private Sub ModeTexture_Click()
    If ModeTexture.Value = vbChecked Then
        SansTextureDynamique = True
    Else
        SansTextureDynamique = False
    End If
End Sub

'
' **********************
' Choix du driver validé
' **********************
'
Private Sub OK_Click()
    dxVue.VideoDriver% = ListeDriver.ListIndex
    Call ProgramLog.Write_File(dxVue.VideoDriver%, dxVue.dxEnum.GetName(dxVue.VideoDriver%))
    Call ProgramLog.Write_File(dxVue.VideoDriver%, dxVue.dxEnum.GetGuid(dxVue.VideoDriver%))
    Call ProgramLog.Write_File(dxVue.VideoDriver%, dxVue.dxEnum.GetDescription(dxVue.VideoDriver%))
    CheminRail$ = "Petit-train 3D\" + ListeRail.List(ListeRail.ListIndex) + ".rail"
    NomFichier$(4) = ListeLangue.List(Me.ListeLangue.ListIndex)
    NomFichier$(5) = dxVue.dxEnum.GetDescription(dxVue.VideoDriver%)
    NomFichier$(6) = ListeRail.List(ListeRail.ListIndex)
    Call NomFichier_Sauve
    Unload Me
    DoEvents
End Sub

'
' *************************************
' Test si la configuration est correcte
' *************************************
'
Public Sub Test_OK()
    If ListeLangue.ListIndex <> -1 And ListeDriver.ListIndex > 0 And ListeRail.ListIndex <> -1 Then
        OK.Visible = True
    Else
        OK.Visible = False
    End If
End Sub

