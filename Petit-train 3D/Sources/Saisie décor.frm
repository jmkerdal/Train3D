VERSION 5.00
Begin VB.Form SaisieDécor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Décors"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "Saisie décor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox RefDecor 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox NomDecor 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Frame FCamera 
      Caption         =   "Caméra"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
      Begin VB.TextBox PosCamera 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox PosCamera 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox PosCamera 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Z"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.ComboBox ChoixDecor 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   3975
   End
   Begin VB.PictureBox VueDecor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton Ajoute 
      Caption         =   "Ajoute"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Timer Rafraichir 
      Interval        =   100
      Left            =   1560
      Top             =   3720
   End
   Begin VB.Label Label1 
      Caption         =   "Réf:"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   11
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "SaisieDécor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dxEdite As New ClassDirectX
Dim FlagSauver As Boolean
Dim Angle!

'
' ***************
' Ajoute un wagon
' ***************
'
Private Sub Ajoute_Click()
    Dim Fichier$
    Dim n%
    '
    Fichier$ = Tools.Open_Box$(Localisation$(CleEDITION% + 5), "", "All files (*.x;*.wall)|*.x;*.wall", BOX_LOAD, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    n% = UBound(ListeDecor()) + 1
    ReDim Preserve ListeDecor(n%) As TypeDecor
    ListeDecor(n%).Nom$ = Fichier$
    ListeDecor(n%).Fichier$ = Fichier$
    Call Initialisation.Décor_Charge_Mesh(n%)
    Call SaisieDécor.Charge_Liste
    SaisieDécor.ChoixDecor.ListIndex = -1
    SaisieDécor.ChoixDecor.ListIndex = SaisieDécor.ChoixDecor.ListCount - 1
    FlagSauver = True
End Sub

'
' **************************
' Mise à jour des paramêtres
' **************************
'
Private Sub ChoixDecor_Click()
    If ChoixDecor.ListIndex <= 0 Then
        FCamera.Visible = False
        NomDecor.Visible = False
        RefDecor.Visible = False
    Else
        FCamera.Visible = True
        NomDecor.Visible = True
        RefDecor.Visible = True
        PosCamera(0).Text = ListeDecor(ChoixDecor.ListIndex).PositionCamera.X
        PosCamera(1).Text = ListeDecor(ChoixDecor.ListIndex).PositionCamera.Y
        PosCamera(2).Text = ListeDecor(ChoixDecor.ListIndex).PositionCamera.z
        NomDecor.Text = ListeDecor(ChoixDecor.ListIndex).Nom$
        RefDecor.Text = ListeDecor(ChoixDecor.ListIndex).Ref$
    End If
End Sub

'
' *********************************
' Lancement de la saisie des décors
' *********************************
'
Private Sub Form_Load()
    Me.Caption = Localisation$(CleEDITION% + 1)
    FCamera.Caption = Localisation$(CleEDITION% + 20)
    Label1(3).Caption = Localisation$(CleEDITION% + 17)
    Ajoute.Caption = Localisation$(17)
    '
    Call SaisieDécor.Charge_Liste
    ChoixDecor.ListIndex = 0
End Sub

'
' **************************
' Charge la liste des décors
' **************************
'
Public Sub Charge_Liste()
    Dim i%
    ChoixDecor.Clear
    Call ChoixDecor.AddItem("<Vide>")
    For i% = 1 To UBound(ListeDecor())
        Call ChoixDecor.AddItem(ListeDecor(i%).Fichier$)
    Next i%
End Sub

'
' ********************
' Initialise la vue 3D
' ********************
'
Private Sub Form_Resize()
    Call dxEdite.Create_3DRM(Me.VueDecor, Me.VueDecor.ScaleWidth, Me.VueDecor.ScaleHeight, Mode3DClipper)
    With dxEdite.dxViewport
        .SetFront 1 / REDUCTION%
        .SetBack 10000! / REDUCTION%
        .SetField 0.5 / REDUCTION%
    End With
End Sub

'
' ******************
' Décharge la vue 3D
' ******************
'
Private Sub Form_Unload(Cancel As Integer)
    If FlagSauver = True Then
        If MsgBox(Localisation$(ClePRINCIPALE% + 16), vbQuestion + vbYesNo, ".\Petit-train 3D\Base.deco") = vbYes Then
            Call Initialisation.Décor_Sauver
        End If
    End If
    Set dxEdite = Nothing
End Sub

'
' ***********************
' Modifie le nom du décor
' ***********************
'
Private Sub NomDecor_KeyUp(KeyCode As Integer, Shift As Integer)
    If ChoixDecor.ListIndex <= 0 Then Exit Sub
    ListeDecor(ChoixDecor.ListIndex).Nom$ = NomDecor.Text
    FlagSauver = True
End Sub

'
' *********************************************
' Position par défaut de la caméra sur le décor
' *********************************************
'
Private Sub PosCamera_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If ChoixDecor.ListIndex <= 0 Then Exit Sub
    Select Case Index
    Case 0
        ListeDecor(ChoixDecor.ListIndex).PositionCamera.X = Val(PosCamera(Index).Text)
    Case 1
        ListeDecor(ChoixDecor.ListIndex).PositionCamera.Y = Val(PosCamera(Index).Text)
    Case 2
        ListeDecor(ChoixDecor.ListIndex).PositionCamera.z = Val(PosCamera(Index).Text)
    End Select
    FlagSauver = True
End Sub

'
' ****************************
' Affiche le wagon séléctionné
' ****************************
'
Private Sub Rafraichir_Timer()
    Dim box As D3DRMBOX
    Dim Recul!, dX!, dY!, dz!
    Dim n%
    n% = ChoixDecor.ListIndex
    If n% > 0 Then
        If Not (ListeDecor(n%).dxDecor Is Nothing) Then
            Call dxEdite.dxScene.AddVisual(ListeDecor(n%).dxDecor)
            Call ListeDecor(n%).dxDecor.GetBox(box)
            dX! = box.Max.X - box.Min.X
            dY! = box.Max.Y - box.Min.Y
            dz! = box.Max.z - box.Min.z
            Recul! = dX!
            If dY! > Recul! Then Recul! = dY!
            If dz! > Recul! Then Recul! = dz!
            Recul! = Recul! * 2
        Else
            Recul! = 100
        End If
    Else
        Recul! = 100
    End If
    Call dxEdite.dxCamera.SetPosition(dxEdite.dxScene, dX! / 2 + box.Min.X, dY! / 2 + box.Min.Y, dz! / 2 + box.Min.z)
    Call dxEdite.dxCamera.SetOrientation(dxEdite.dxScene, 0, 0, 1, 0, 1, 0)
    Call dxEdite.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, Angle!)
    Call dxEdite.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 1, 0, 0, 15 * DegRad!)
    Call dxEdite.dxCamera.AddTranslation(D3DRMCOMBINE_BEFORE, 0, 0, -Recul!)
    Angle! = Angle! + PI! / 16
    If Angle! > 2 * PI! Then
        Angle! = Angle! - 2 * PI!
    End If
    '
    Call dxEdite.Render(False)
    '
    If n% > 0 Then
        If Not (ListeDecor(n%).dxDecor Is Nothing) Then
            Call dxEdite.dxScene.DeleteVisual(ListeDecor(n%).dxDecor)
        End If
    End If
End Sub

'
' ********************
' Modifie la référence
' ********************
'
Private Sub RefDecor_KeyUp(KeyCode As Integer, Shift As Integer)
    If ChoixDecor.ListIndex <= 0 Then Exit Sub
    ListeDecor(ChoixDecor.ListIndex).Ref$ = RefDecor.Text
    FlagSauver = True
End Sub

