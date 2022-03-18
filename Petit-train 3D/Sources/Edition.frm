VERSION 5.00
Begin VB.Form Edition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edition des éléments"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "Edition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Selection 
      Caption         =   "Décor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton Selection 
      Caption         =   "Voie"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox ChoixDecor 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   4095
   End
   Begin VB.ComboBox ChoixVoie 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox VueVoie 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   1080
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton Valide 
      Caption         =   "Valide"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "Edition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dxEdite As New ClassDirectX

'
' *******************
' Change la selection
' *******************
'
Private Sub ChoixDecor_Click()
    Call Affiche_Decor
End Sub

'
' *******************
' Change la selection
' *******************
'
Private Sub ChoixVoie_Click()
    Call Affiche_Voie
End Sub

'
' **************************
' Charge les valeurs saisies
' **************************
'
Private Sub Form_Load()
    Dim i%
    Me.Caption = Localisation$(CleEDITION% + 3)
    Me.Selection(0).Caption = Localisation$(CleEDITION%)
    Me.Selection(1).Caption = Localisation$(CleEDITION% + 1)
    Me.Valide.Caption = Localisation$(CleEDITION% + 2)
    ChoixVoie.Clear
    Call ChoixVoie.AddItem("<" + Localisation$(CleEDITION% + 4) + ">")
    For i% = 1 To UBound(Voie())
        Call ChoixVoie.AddItem("[" & Voie(i%).Ref$ & "] " & Voie(i%).Libelle$(0))
    Next i%
    ChoixDecor.Clear
    Call ChoixDecor.AddItem("<" + Localisation$(CleEDITION% + 4) + ">")
    For i% = 1 To UBound(ListeDecor())
        Call ChoixDecor.AddItem("[" & ListeDecor(i%).Ref$ & "] " & ListeDecor(i%).Nom$)
    Next i%
End Sub

'
' ******************
' Lance la vue en 3D
' ******************
'
Private Sub Form_Resize()
    Call dxEdite.Create_3DRM(Me.VueVoie, Me.VueVoie.ScaleWidth, Me.VueVoie.ScaleHeight, Mode3DClipper)
    With dxEdite.dxViewport
        .SetFront 1 / REDUCTION%
        .SetBack 10000! / REDUCTION%
        .SetField 0.5 / REDUCTION%
    End With
    ChoixVoie.ListIndex = 0
    ChoixDecor.ListIndex = 0
    Selection(0).Value = True
End Sub

'
' ******************
' Décharge la vue 3D
' ******************
'
Private Sub Form_Unload(Cancel As Integer)
    Set dxEdite = Nothing
End Sub

'
' **************************
' Séléction du choix visible
' **************************
'
Private Sub Selection_Click(Index As Integer)
    If Index = 0 Then
        ChoixVoie.Visible = True
        Call Affiche_Voie
    Else
        ChoixVoie.Visible = False
    End If
    If Index = 1 Then
        ChoixDecor.Visible = True
        Call Affiche_Decor
    Else
        ChoixDecor.Visible = False
    End If
End Sub

'
' *************************
' Ajoute une voie au réseau
' *************************
'
Private Sub Valide_Click()
    If Selection(0).Value = True Then Call Voies.Ajoute(ChoixVoie.ListIndex, 0)
    If Selection(1).Value = True Then Call Initialisation.Décor_Ajoute(ChoixDecor.ListIndex, 0)
    Unload Me
End Sub

'
' *****************
' Reaffiche la voie
' *****************
'
Private Sub VueVoie_Paint()
    If Selection(0).Value = True Then Call ChoixVoie_Click
    If Selection(1).Value = True Then Call ChoixDecor_Click
End Sub

'
' ************************
' Affiche la nouvelle voie
' ************************
'
Public Sub Affiche_Voie()
    Dim box As D3DRMBOX
    Dim Recul!, dX!, dz!
    If ChoixVoie.ListIndex > 0 Then
        Valide.Enabled = True
        Call dxEdite.dxScene.AddVisual(Voie(ChoixVoie.ListIndex).VoieMeshBuilder)
        Call Voie(ChoixVoie.ListIndex).VoieMeshBuilder.GetBox(box)
        dX! = box.Max.X - box.Min.X
        dz! = box.Max.z - box.Min.z
        If dX! > dz! Then
            Recul! = dX! * 1.1
        Else
            Recul! = dz! * 1.1
        End If
    Else
        Valide.Enabled = False
        Recul! = 2
    End If
    If Recul! < 2 Then Recul! = 2
    Call dxEdite.dxCamera.SetPosition(dxEdite.dxScene, dX! / 2 + box.Min.X, Recul!, dz! / 2 + box.Min.z)
    Call dxEdite.dxCamera.SetOrientation(dxEdite.dxScene, 0, -1, 0, 0, 0, 1)
    '
    Call dxEdite.Render(False)
    '
    If ChoixVoie.ListIndex > 0 Then
        Call dxEdite.dxScene.DeleteVisual(Voie(ChoixVoie.ListIndex).VoieMeshBuilder)
    End If
End Sub

'
' ************************
' Affiche le nouveau décor
' ************************
'
Public Sub Affiche_Decor()
    Dim box As D3DRMBOX
    Dim Recul!, dX!, dY!, dz!
    If ChoixDecor.ListIndex > 0 Then
        If Not (ListeDecor(ChoixDecor.ListIndex).dxDecor Is Nothing) Then
            Valide.Enabled = True
            Call dxEdite.dxScene.AddVisual(ListeDecor(ChoixDecor.ListIndex).dxDecor)
            Call ListeDecor(ChoixDecor.ListIndex).dxDecor.GetBox(box)
            dX! = box.Max.X - box.Min.X
            dY! = box.Max.Y - box.Min.Y
            dz! = box.Max.z - box.Min.z
            Recul! = dX!
            If dY! > Recul! Then Recul! = dY!
            If dz! > Recul! Then Recul! = dz!
            Recul! = Recul! * 2
        Else
            Valide.Enabled = False
            Recul! = 2
        End If
    Else
        Valide.Enabled = False
        Recul! = 2
    End If
    If Recul! < 2 Then Recul! = 2
    Call dxEdite.dxCamera.SetPosition(dxEdite.dxScene, dX! / 2 + box.Min.X, dY! / 2 + box.Min.Y, dz! / 2 + box.Min.z)
    Call dxEdite.dxCamera.SetOrientation(dxEdite.dxScene, 0, 0, 1, 0, 1, 0)
    Call dxEdite.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 1, 0, 0, 15 * DegRad!)
    Call dxEdite.dxCamera.AddTranslation(D3DRMCOMBINE_BEFORE, 0, 0, -Recul!)
    '
    Call dxEdite.Render(False)
    '
    If ChoixDecor.ListIndex > 0 Then
        If Not (ListeDecor(ChoixDecor.ListIndex).dxDecor Is Nothing) Then
            Call dxEdite.dxScene.DeleteVisual(ListeDecor(ChoixDecor.ListIndex).dxDecor)
        End If
    End If
End Sub

