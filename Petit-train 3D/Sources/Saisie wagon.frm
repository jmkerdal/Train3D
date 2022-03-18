VERSION 5.00
Begin VB.Form SaisieWagon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liste des wagons"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "Saisie wagon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox InterieurWagon 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   840
      Width           =   4215
   End
   Begin VB.CheckBox ChoixElectrique 
      Caption         =   "Electrique"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox BogieWagon 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox RefWagon 
      Height          =   285
      Left            =   2640
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox NomWagon 
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Frame FCamera 
      Caption         =   "Caméra"
      Height          =   1095
      Left            =   2160
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
      Begin VB.TextBox PosCamera 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox PosCamera 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox PosCamera 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Z"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "X"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.CheckBox ChoixMotrice 
      Caption         =   "Motrice"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.Frame FEssieu 
      Caption         =   "Position des essieux"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   2055
      Begin VB.TextBox Longueur_Wagon 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Eccart_Essieu 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Longueur"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Eccart Essieux"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FBogie 
      Caption         =   "Eccart entre les bogies"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
      Begin VB.TextBox Eccart_Bogie 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Eccart_Bogie 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Groupe B"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Groupe A"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer Rafraichir 
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin VB.CommandButton Ajoute 
      Caption         =   "Ajoute"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   4920
      Width           =   975
   End
   Begin VB.PictureBox VueWagon 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.ComboBox ChoixWagon 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Réf:"
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   23
      Top             =   3840
      Width           =   375
   End
End
Attribute VB_Name = "SaisieWagon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FlagSauver As Boolean
Dim dxSaisieWagon As New ClassWagon

'
' ***************
' Ajoute un wagon
' ***************
'
Private Sub Ajoute_Click()
    Dim Fichier$
    Dim n%
    Fichier$ = Tools.Open_Box$(Localisation$(CleEDITION% + 6), "", "All files (*.wall;*.x)|*.wall;*.x", BOX_LOAD, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    n% = UBound(ListeWagon()) + 1
    ReDim Preserve ListeWagon(n%) As TypeWagon
    ListeWagon(n%).Fichier$ = Fichier$
    ListeWagon(n%).Nom$ = Fichier$
    Call Initialisation.Wagon_Charger_Mesh(n%)
    Call SaisieWagon.Charge_Liste
    SaisieWagon.ChoixWagon.ListIndex = -1
    SaisieWagon.ChoixWagon.ListIndex = SaisieWagon.ChoixWagon.ListCount - 1
    FlagSauver = True
End Sub

'
' *******************************
' Charge la référence de la bogie
' *******************************
'
Private Sub BogieWagon_DblClick()
    Dim Fichier$
    Dim n%
    Fichier$ = Tools.Open_Box$(Localisation$(CleEDITION% + 7), "", "All files (*.wall;*.x)|*.wall;*.x", BOX_LOAD, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    n% = ChoixWagon.ListIndex
    ListeWagon(n%).FichierBogie$ = Fichier$
    BogieWagon.Text = Fichier$
    If Right$(ListeWagon(n%).FichierBogie$, 2) = ".x" Then
        Set ListeWagon(n%).dxBogie = dxVue.Load_MeshBuilder(ListeWagon(n%).FichierBogie$)
    Else
        Set ListeWagon(n%).dxBogie = Charge_Wall(ListeWagon(n%).FichierBogie$)
    End If
    Call ListeWagon(n%).dxBogie.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    FlagSauver = True
End Sub

'
' *************************************
' Indicateur de motrice électrique
' génére automatiquement les catenaires
' *************************************
'
Private Sub ChoixElectrique_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    ListeWagon(ChoixWagon.ListIndex).Electrique% = ChoixElectrique.Value
Debug.Print "Click2"
    FlagSauver = True
End Sub

'
' ************************************
' Indique que le wagon est une motrice
' ************************************
'
Private Sub ChoixMotrice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    ListeWagon(ChoixWagon.ListIndex).Motrice% = ChoixMotrice.Value
    FlagSauver = True
End Sub

'
' *****************************
' Séléctionne le wagon à éditer
' *****************************
'
Private Sub ChoixWagon_Click()
    Dim i%, n%
    n% = ChoixWagon.ListIndex
    If n% <= 0 Then
        FBogie.Visible = False
        FEssieu.Visible = False
        FCamera.Visible = False
        NomWagon.Visible = False
        RefWagon.Visible = False
        ChoixMotrice.Visible = False
        ChoixElectrique.Visible = False
        BogieWagon.Visible = False
        InterieurWagon.Visible = False
    Else
        FBogie.Visible = True
        FEssieu.Visible = True
        FCamera.Visible = True
        NomWagon.Visible = True
        RefWagon.Visible = True
        ChoixMotrice.Visible = True
        ChoixElectrique.Visible = True
        BogieWagon.Visible = True
        InterieurWagon.Visible = True
        For i% = 0 To 1
            Eccart_Bogie(i%).Text = ListeWagon(n%).EccartBogie!(i%)
        Next i%
        Eccart_Essieu.Text = ListeWagon(n%).EccartEssieu!
        Longueur_Wagon.Text = ListeWagon(n%).Longueur!
        ChoixMotrice.Value = ListeWagon(n%).Motrice%
        ChoixElectrique.Value = ListeWagon(n%).Electrique%
        PosCamera(0).Text = ListeWagon(n%).PositionCamera.X
        PosCamera(1).Text = ListeWagon(n%).PositionCamera.Y
        PosCamera(2).Text = ListeWagon(n%).PositionCamera.z
        NomWagon.Text = ListeWagon(n%).Nom$
        RefWagon.Text = ListeWagon(n%).Ref$
        BogieWagon.Text = ListeWagon(n%).FichierBogie$
        InterieurWagon.Text = ListeWagon(n%).FichierInterieur$
    End If
End Sub

'
' ***********************************************
' Change l'eccart entre deux bogies sur un essieu
' ***********************************************
'
Private Sub Eccart_Bogie_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    ListeWagon(ChoixWagon.ListIndex).EccartBogie!(Index) = Val(Eccart_Bogie(Index).Text)
    FlagSauver = True
End Sub

'
' *********************************
' Change l'eccart entre les essieux
' *********************************
'
Private Sub Eccart_Essieu_KeyUp(KeyCode As Integer, Shift As Integer)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    ListeWagon(ChoixWagon.ListIndex).EccartEssieu! = Val(Eccart_Essieu.Text)
    FlagSauver = True
End Sub

'
' *********************************
' Lancement de la saisie des wagons
' *********************************
'
Private Sub Form_Load()
    Me.Caption = Localisation$(CleEDITION% + 10)
    FBogie.Caption = Localisation$(CleEDITION% + 11)
    Label1(0).Caption = Localisation$(CleEDITION% + 12)
    Label1(1).Caption = Localisation$(CleEDITION% + 13)
    FEssieu.Caption = Localisation$(CleEDITION% + 14)
    Label1(2).Caption = Localisation$(CleEDITION% + 15)
    Label1(3).Caption = Localisation$(CleEDITION% + 16)
    Label1(7).Caption = Localisation$(CleEDITION% + 17)
    ChoixMotrice.Caption = Localisation$(CleEDITION% + 18)
    ChoixElectrique.Caption = Localisation$(CleEDITION% + 19)
    FCamera.Caption = Localisation$(CleEDITION% + 20)
    Ajoute.Caption = Localisation$(17)
    '
    Call SaisieWagon.Charge_Liste
    ChoixWagon.ListIndex = 0
End Sub

'
' **************************
' Charge la liste des wagons
' **************************
'
Public Sub Charge_Liste()
    Dim i%
    ChoixWagon.Clear
    Call ChoixWagon.AddItem("<Vide>")
    For i% = 1 To UBound(ListeWagon())
        Call ChoixWagon.AddItem(ListeWagon(i%).Fichier$)
    Next i%
End Sub

'
' ********************
' Initialise la vue 3D
' ********************
'
Private Sub Form_Resize()
    Call dxSaisieWagon.Charge(Me.VueWagon)
End Sub

'
' ******************
' Décharge la vue 3D
' ******************
'
Private Sub Form_Unload(Cancel As Integer)
    If FlagSauver = True Then
        If MsgBox(Localisation$(ClePRINCIPALE% + 16), vbQuestion + vbYesNo, ".\Petit-train 3D\Base.wag") = vbYes Then
            Call Initialisation.Wagon_Sauver
        End If
    End If
    Set dxSaisieWagon = Nothing
End Sub

'
' **********************************
' Charge la référence de l'interieur
' **********************************
'
Private Sub InterieurWagon_DblClick()
    Dim Fichier$
    Dim n%
    Fichier$ = Tools.Open_Box$(Localisation$(CleEDITION% + 23), "", "All files (*.wall;*.x)|*.wall;*.x", BOX_LOAD, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    n% = ChoixWagon.ListIndex
    ListeWagon(n%).FichierInterieur$ = Fichier$
    InterieurWagon.Text = Fichier$
    If Right$(ListeWagon(n%).FichierInterieur$, 2) = ".x" Then
        Set ListeWagon(n%).dxInterieur = dxVue.Load_MeshBuilder(ListeWagon(n%).FichierInterieur$)
    Else
        Set ListeWagon(n%).dxInterieur = Charge_Wall(ListeWagon(n%).FichierInterieur$)
    End If
    Call ListeWagon(n%).dxInterieur.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    FlagSauver = True
End Sub

'
' ***************************
' Change la longueur du wagon
' ***************************
'
Private Sub Longueur_Wagon_KeyUp(KeyCode As Integer, Shift As Integer)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    ListeWagon(ChoixWagon.ListIndex).Longueur! = Val(Longueur_Wagon.Text)
    FlagSauver = True
End Sub

'
' ***********************
' Modifie le nom du wagon
' ***********************
'
Private Sub NomWagon_KeyUp(KeyCode As Integer, Shift As Integer)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    ListeWagon(ChoixWagon.ListIndex).Nom$ = NomWagon.Text
    FlagSauver = True
End Sub

'
' **********************************************
' Position par défaut de la caméra dans le wagon
' **********************************************
'
Private Sub PosCamera_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    Select Case Index
    Case 0
        ListeWagon(ChoixWagon.ListIndex).PositionCamera.X = Val(PosCamera(Index).Text)
    Case 1
        ListeWagon(ChoixWagon.ListIndex).PositionCamera.Y = Val(PosCamera(Index).Text)
    Case 2
        ListeWagon(ChoixWagon.ListIndex).PositionCamera.z = Val(PosCamera(Index).Text)
    End Select
    FlagSauver = True
End Sub

'
' ****************************
' Affiche le wagon séléctionné
' ****************************
'
Private Sub Rafraichir_Timer()
    Call dxSaisieWagon.Rafraichir(ChoixWagon.ListIndex)
End Sub

'
' ********************
' Modifie la référence
' ********************
'
Private Sub RefWagon_KeyUp(KeyCode As Integer, Shift As Integer)
    If ChoixWagon.ListIndex <= 0 Then Exit Sub
    ListeWagon(ChoixWagon.ListIndex).Ref$ = RefWagon.Text
    FlagSauver = True
End Sub

