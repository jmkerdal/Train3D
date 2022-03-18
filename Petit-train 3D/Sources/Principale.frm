VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Principale 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Petit train 3D"
   ClientHeight    =   3780
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9195
   Icon            =   "Principale.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Boite 
      Left            =   120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.CheckBox CTunnel 
         Height          =   615
         Left            =   1560
         Picture         =   "Principale.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Tunnel"
         Top             =   120
         Width           =   615
      End
      Begin MSComctlLib.Slider RegleVitesse 
         Height          =   495
         Left            =   1560
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -40
         Max             =   40
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider RegleRecul 
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   50
         Min             =   300
         Max             =   2000
         SelStart        =   500
         TickFrequency   =   100
         Value           =   500
      End
      Begin VB.ComboBox ChoixElement 
         Height          =   315
         Index           =   1
         Left            =   4800
         TabIndex        =   10
         Text            =   "Choix élément"
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox ChoixElement 
         Height          =   315
         Index           =   0
         Left            =   4800
         TabIndex        =   9
         Text            =   "Choix élément"
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame FMode 
         Caption         =   "Mode actuelle"
         Height          =   660
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1455
         Begin VB.OptionButton ModeSaisie 
            Caption         =   "Visu."
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton ModeSaisie 
            Caption         =   "Edition"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.OptionButton ModeVue 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   4320
         Picture         =   "Principale.frx":0D8C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Poursuite"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton ModeVue 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3960
         Picture         =   "Principale.frx":127E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "A bord"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton ModeVue 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3600
         Picture         =   "Principale.frx":1770
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Vue dessus"
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton ModeVue 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   3240
         Picture         =   "Principale.frx":1C62
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Plein écran"
         Top             =   0
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.Label VitesseRotatif 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9195
      TabIndex        =   2
      Top             =   735
      Width           =   9195
   End
   Begin VB.Menu MENU_Fichier 
      Caption         =   "Fichier"
      Begin VB.Menu MENU_Nouveau 
         Caption         =   "Nouveau"
      End
      Begin VB.Menu MENU_Charger 
         Caption         =   "Charger"
      End
      Begin VB.Menu MENU_Enregistrer 
         Caption         =   "Enregistrer"
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_Parametre 
         Caption         =   "Paramêtres"
      End
      Begin VB.Menu MENU_Imprimantes 
         Caption         =   "Imprimantes"
      End
      Begin VB.Menu MENU_Fichier_Moins 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_NomFichier 
         Caption         =   "<>"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_NomFichier 
         Caption         =   "<>"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_NomFichier 
         Caption         =   "<>"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_NomFichier 
         Caption         =   "<>"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Fichier_Moins1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Quitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu MENU_Edition 
      Caption         =   "Edition"
      Begin VB.Menu MENU_Train 
         Caption         =   "Train"
      End
      Begin VB.Menu MENU_Couper 
         Caption         =   "Couper"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu MENU_Copier 
         Caption         =   "Copier"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu MENU_Coller 
         Caption         =   "Coller"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu MENU_Inventaire 
         Caption         =   "Inventaire"
      End
      Begin VB.Menu MENU_Element 
         Caption         =   "Elements chargés"
      End
      Begin VB.Menu MENU_Edition_Moins 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Wagon 
         Caption         =   "Wagons"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Décor 
         Caption         =   "Décor"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Voie 
         Caption         =   "Voie"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Fil_De_Fer 
         Caption         =   "Fil de fer"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MENU_Vue 
      Caption         =   "Menu Vue"
      Begin VB.Menu VUE_Ajoute 
         Caption         =   "Ajoute"
      End
      Begin VB.Menu VUE_Supprime 
         Caption         =   "Supprime"
      End
      Begin VB.Menu VUE_Origine 
         Caption         =   "Origine"
      End
      Begin VB.Menu VUE_Rotation 
         Caption         =   "Rotation"
      End
   End
   Begin VB.Menu MENU_Tunnel 
      Caption         =   "Menu Tunnel"
      Begin VB.Menu TUNNEL_Insere 
         Caption         =   "Insere"
      End
      Begin VB.Menu TUNNEL_Supprime 
         Caption         =   "Supprime"
      End
      Begin VB.Menu TUNNEL_Inverse 
         Caption         =   "Inverse"
      End
      Begin VB.Menu TUNNEL_Moins 
         Caption         =   "-"
      End
      Begin VB.Menu TUNNEL_Ajoute 
         Caption         =   "Ajoute"
      End
      Begin VB.Menu TUNNEL_Efface 
         Caption         =   "Efface"
      End
   End
   Begin VB.Menu MENU_Interrogation 
      Caption         =   "?"
      Begin VB.Menu MENU_A_Propos 
         Caption         =   "A propos"
      End
   End
End
Attribute VB_Name = "Principale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' ****************************
' Efface un élément voie/décor
' ****************************
'
Private Sub Efface_Element()
    Dim i%
    Call Initialisation.Vue_Decharge
    If Pointe(0).Nom$ = "Voie" Then
        For i% = 0 To NbSegment%
            Call Initialisation.Voie_Supprime(Pointe(0).No%, i%, Reseau(Pointe(0).No%).Connecte(i%), Reseau(Pointe(0).No%).Entree(i%))
        Next i%
        TypeCopie% = 1
        NoCopie% = Reseau(Pointe(0).No%).NoVoie%
        AngleCopie! = Reseau(Pointe(0).No%).Angle!
        Reseau(Pointe(0).No%).NoVoie% = 0
    End If
    If Pointe(0).Nom$ = "Decor" Then
        TypeCopie% = 2
        NoCopie% = ElementDecor(Pointe(0).No%).NoDecor%
        AngleCopie! = ElementDecor(Pointe(0).No%).Angle%
        ElementDecor(Pointe(0).No%).NoDecor% = 0
    End If
    Principale.MENU_Coller.Enabled = True
    Call Vue.Calcule_Reseau
    Call Initialisation.Vue_Charge
End Sub

'
' *************************************
' Modification de l'élément séléctionné
' *************************************
'
Private Sub ChoixElement_Click(Index As Integer)
    IndexElement%(Index) = IndexChoixElement%(0, Principale.ChoixElement(Index).ListIndex)
    IndexChoix%(Index) = IndexChoixElement%(1, Principale.ChoixElement(Index).ListIndex)
    CameraTourne%(2) = 0
End Sub

'
' *********************************
' Charge les informations initiales
' *********************************
'
Private Sub MDIForm_Load()
    Dim ta!, tb!
    '
    Call ProgramLog.Select_Path("Train3D.log", LogMode_Overwrite)
    Call ProgramLog.Write_File(0, "Start application")
    '
    DegRad! = PI! / 180
    ParamElevation = True
    ParamHauteur% = 5
    ParamCatenaire = True
    ParamCiel = True
    Super$ = Chr$(115) + Chr$(117) + Chr$(112) + Chr$(101) + Chr$(114)
    Super$ = Super$ + Chr$(99) + Chr$(105) + Chr$(116) + Chr$(114) + Chr$(111) + Chr$(110)
    '
    Me.MENU_Vue.Visible = False
    Me.MENU_Tunnel.Visible = False
    '
    Call Initialisation.NomFichier_Charge
    Driver.Show vbModal
    Call Initialisation.Valide_Menu
    If Command$ <> Super$ Then Call Lancement.Show
    DoEvents
    Me.FMode.Caption = Localisation$(ClePRINCIPALE%)
    Me.ModeSaisie(0).Caption = Localisation$(ClePRINCIPALE% + 1)
    Me.ModeSaisie(1).Caption = Localisation$(ClePRINCIPALE% + 2)
    '
    ta! = Timer
    '
    Call ProgramLog.Write_File(0, "Create DSound")
    Call DSound.MainForm(Me.hWnd&, 2, 16)
    Call DSound.Add(".\Petit-train 3D\Roule.wav")
    '
    Call ProgramLog.Write_File(0, "Load base wall definition")
    Set dxTraverse(0) = Charge_Wall(".\Petit-train 3D\Textures\Traverse.wall")
    Set dxTraverse(1) = Charge_Wall(".\Petit-train 3D\Textures\Traverse light.wall")
    Set dxTraverse(2) = Charge_Wall(".\Petit-train 3D\Textures\Traverse couche 0.wall")
    Set dxTraverse(3) = Charge_Wall(".\Petit-train 3D\Textures\Traverse couche 1.wall")
    Set dxTraverse(4) = Charge_Wall(".\Petit-train 3D\Textures\Traverse couche 2.wall")
    Set dxHeurtoir = Charge_Wall(".\Petit-train 3D\Textures\Heurtoir.wall")
    Set dxCatenaire = Charge_Wall(".\Petit-train 3D\Textures\Catenaire.wall")
    Call dxCatenaire.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    Set dxCiel = Charge_Wall(".\Petit-train 3D\Textures\Ciel.wall")
    '
    Set dxRoue = Charge_Wall(".\Petit-train 3D\Textures\Roue2.wall")
    Call dxRoue.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    Set dxSelection(0) = Charge_Wall(".\Petit-train 3D\Textures\Selection rouge.wall")
    Set dxSelection(1) = Charge_Wall(".\Petit-train 3D\Textures\Selection vert.wall")
    Set dxSelection(2) = Charge_Wall(".\Petit-train 3D\Textures\Selection bleu.wall")
    Call dxSelection(0).ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    Call dxSelection(1).ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    Call dxSelection(2).ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    Set dxAiguille = Charge_Wall(".\Petit-train 3D\Textures\Aiguille.wall")
    Call dxAiguille.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    Set dxAttele = Charge_Wall(".\Petit-train 3D\Textures\Attele.wall")
    Call dxAttele.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    '
    Set dxSolTexture(0) = dxVue.Load_Texture(".\Petit-train 3D\Textures\Verdure.bmp")
    Set dxSolTexture(1) = dxVue.Load_Texture(".\Petit-train 3D\Textures\Ballast.bmp")
    '
    Call Initialisation.Raz
    Call ProgramLog.Write_File(0, "Load tracks")
    Call Voies.Charger(CheminRail$)
    Call Voies.Cree_Meshbuilders
    '
    Call ProgramLog.Write_File(0, "Load wagons")
    Call Initialisation.Wagon_Charger
    Call ProgramLog.Write_File(0, "Load decors")
    Call Initialisation.Décor_Charger
    Call ProgramLog.Write_File(0, "Create DirectX objects")
    Call Initialisation.ObjetDX_Cree
    '
    CameraRecul% = 250
    PosCamera.Y = 500
    PosCamera.X = 0
    PosCamera.z = 0
    '
    If Command$ <> Super$ Then Unload Lancement
    DoEvents
    '
    tb! = Timer
    Call ProgramLog.Write_File(tb! - ta!, "Run application")
    Call Aide.Move(0, Screen.Height / 2)
    Call Aide.Ecrit(Localisation$(CleEDITION% + 22) + ":" + Str$(tb! - ta!))
    '
    Call Vue.Move(0, 0, Screen.Width / 2, Screen.Height / 2)
    Call Vue.Show
End Sub

'
' ****************
' Fin du programme
' ****************
'
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i%, j%
    Tourne = False
    If FlagNomFichier = True Then
        Call Initialisation.NomFichier_Sauve
    End If
    Call Initialisation.ObjetDX_Detruit
    For i% = 0 To UBound(Voie())
        Set Voie(i%).VoieMeshBuilder = Nothing
        Set Voie(i%).VoieLightMeshBuilder = Nothing
        Set Voie(i%).VoieCatenaireMeshBuilder = Nothing
        For j% = 0 To 2
            Set Voie(i%).VoieTexture(0) = Nothing
        Next j%
        Set FormeVoie(i%) = Nothing
    Next i%
    For i% = 0 To 2
        Set dxSelection(i%) = Nothing
    Next i%
    For i% = 1 To UBound(ListeWagon())
        Set ListeWagon(i%).dxWagon = Nothing
        Set ListeWagon(i%).dxBogie = Nothing
        Set ListeWagon(i%).dxInterieur = Nothing
    Next i%
    For i% = 1 To UBound(ListeDecor())
        Set ListeDecor(i%).dxDecor = Nothing
    Next i%
    '
    For i% = 0 To UBound(dxTraverse())
        Set dxTraverse(i%) = Nothing
    Next i%
    Set dxRoue = Nothing
    Set dxHeurtoir = Nothing
    Set dxAiguille = Nothing
    Set dxAttele = Nothing
    Set dxCatenaire = Nothing
    Set dxCiel = Nothing
    Set dxSolTexture(0) = Nothing
    Set dxSolTexture(1) = Nothing
    '
    Set dxVue = Nothing
    Set DSound = Nothing
End Sub

'
' ******************************
' Appel de la fenêtre d'à propos
' ******************************
'
Private Sub MENU_A_Propos_Click()
    Dim a$, r$
    a$ = Localisation$(CleAPROPOS% + 2) + vbCr + _
        "2, place des Martyrs" + vbCr + _
        "92110 Clichy, France" + vbCr + _
        Localisation$(CleAPROPOS% + 3) + ": +33 01 49 68 83 67"
    Call frmAbout.About(Me, Localisation$(CleINVENTAIRE% + 3), a$, Localisation$(CleAPROPOS% + 1), Localisation$(ClePRINCIPALE% + 3), ".\Petit-train 3D\Textures\Kerdal.gif")
End Sub

'
' **********************
' Chargement d'un réseau
' **********************
'
Private Sub MENU_Charger_Click()
    Dim Fichier$
    Fichier$ = Tools.Open_Box$(Localisation$(ClePRINCIPALE% + 4), CheminReseau$, "All files *.res|*.res", BOX_LOAD, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    Call Me.MAJ_NomFichier(Fichier$)
    Call Réseau_Charge(Fichier$)
End Sub

'
' ****************************
' Ajoute la même dernière voie
' ****************************
'
Private Sub MENU_Coller_Click()
    If ModeActuelle = ModeEdition Then
        If TypeCopie% = 1 Then
            Call Voies.Ajoute(NoCopie%, AngleCopie!)
        End If
        If TypeCopie% = 2 Then
            Call Initialisation.Décor_Ajoute(NoCopie%, AngleCopie!)
        End If
    End If
End Sub

'
' *************
' Copie la voie
' *************
'
Private Sub MENU_Copier_Click()
    If Pointe(0).Nom$ = "Voie" Then
        TypeCopie% = 1
        NoCopie% = Reseau(Pointe(0).No%).NoVoie%
        AngleCopie! = Reseau(Pointe(0).No%).Angle!
    Else
        TypeCopie% = 2
        NoCopie% = ElementDecor(Pointe(0).No%).NoDecor%
        AngleCopie! = ElementDecor(Pointe(0).No%).Angle%
    End If
    Principale.MENU_Coller.Enabled = True
End Sub

'
' *****************
' Efface un élément
' *****************
'
Private Sub MENU_Couper_Click()
    Call Efface_Element
End Sub

'
' *****************************
' Saisie de la liste des décors
' *****************************
'
Private Sub MENU_Décor_Click()
    Call SaisieDécor.Show
End Sub

'
' ****************************************
' Affiche la feneêtre des éléments chargés
' ****************************************
'
Private Sub MENU_Element_Click()
    SaisieCharge.Show vbModal
End Sub

'
' ********************
' Enregistre le réseau
' ********************
'
Private Sub MENU_Enregistrer_Click()
    Dim Fichier$, f%
    Dim i%, j%, n%, a%
    Fichier$ = Tools.Open_Box$(Localisation$(ClePRINCIPALE% + 5), CheminReseau$, "All files *.res|*.res", BOX_SAVE, Principale.Boite)
    If Fichier$ = "" Then Exit Sub
    If Tools.Exist(Fichier$) = True Then
        If MsgBox(Localisation$(ClePRINCIPALE% + 16), vbQuestion + vbYesNo, Fichier$) = vbNo Then Exit Sub
    End If
    Call Me.MAJ_NomFichier(Fichier$)
    CheminReseau$ = Fichier$
    Me.Caption = CheminReseau$ + " :" + Localisation$(CleINVENTAIRE% + 3)
    f% = FreeFile()
    Open CheminReseau$ For Output As #f%
    ' ***** Composition du train
    n% = UBound(ListeTrain())
    Print #f%, n%
    If n% <> 0 Then
        For i% = 1 To n%
            Print #f%, ListeWagon(ListeTrain(i%).NoWagon%).Ref$
        Next i%
    End If
    ' ***** Liste des rails
    ' ***** Info de départ
    n% = UBound(Reseau())
    Print #f%, n%
    Print #f%, Origine%
    Print #f%, Reseau(Origine%).Angle
    Print #f%, Reseau(Origine%).Position.X;
    Print #f%, Reseau(Origine%).Position.Y;
    Print #f%, Reseau(Origine%).Position.z
    For i% = 1 To n%
        Print #f%, Voie(Reseau(i%).NoVoie%).Ref$
        For j% = 0 To NbSegment%
            Print #f%, Reseau(i%).Connecte%(j%); Reseau(i%).Entree%(j%);
        Next j%
        Print #f%,
    Next i%
    ' ***** Liste des décors
    n% = UBound(ElementDecor())
    Print #f%, n%
    If n% <> 0 Then
        For i% = 1 To n%
            Print #f%, ListeDecor(ElementDecor(i%).NoDecor%).Ref$
            Print #f%, (ElementDecor(i%).Position.X \ 10) * 10;
            Print #f%, (ElementDecor(i%).Position.Y \ 10) * 10;
            Print #f%, (ElementDecor(i%).Position.z \ 10) * 10;
            Print #f%, ElementDecor(i%).Angle%
        Next i%
    End If
    ' ***** Tunnels
    n% = UBound(ListeTunnel())
    Print #f%, n%
    For i% = 0 To n% - 1
        Print #f%, ListeTunnel(i%).Nb_Point%
        For j% = 0 To ListeTunnel(i%).Nb_Point% - 1
            Print #f%, ListeTunnel(i%).PositionX(j%);
            Print #f%, ListeTunnel(i%).PositionZ(j%);
            If ListeTunnel(i%).Face(j%) = False Then
                Print #f%, 0
            Else
                Print #f%, 1
            End If
        Next j%
    Next i%
    ' Sauve les paramêtres
    Print #f%, IIf(ParamElevation = True, " 1 ", " 0 ");
    Print #f%, ParamHauteur%;
    Print #f%, IIf(ParamCatenaire = True, " 1 ", " 0 ");
    Print #f%, IIf(ParamSens = True, " 1 ", " 0 ");
    Print #f%, IIf(ParamCiel = True, " 1 ", " 0 ")
    Close #f%
End Sub

'
' ************************
' Passe en mode fil de fer
' ************************
'
Private Sub MENU_Fil_De_Fer_Click()
    If MENU_Fil_De_Fer.Checked = False Then
        Call dxVue.dxDevice.SetQuality(QUALITE_VISION)
        MENU_Fil_De_Fer.Checked = True
    Else
        Call dxVue.dxDevice.SetQuality(QUALITE_NORMAL)
        MENU_Fil_De_Fer.Checked = False
    End If
End Sub

'
' *************************
' Séléction de l'imprimante
' *************************
'
Private Sub MENU_Imprimantes_Click()
    On Error GoTo Fin
    Principale.Boite.ShowPrinter
Fin:
End Sub

'
' ***********************************
' Appel de la fenêtre des inventaires
' ***********************************
'
Private Sub MENU_Inventaire_Click()
    Inventaire.Show
End Sub

'
' *************************************
' Charge un réseau à partir de la liste
' *************************************
'
Private Sub MENU_NomFichier_Click(Index As Integer)
    Pointe(0).No% = 0
    Pointe(0).Nom$ = ""
    Pointe(1).No% = 0
    Pointe(1).Nom$ = ""
    CameraRecul% = 250
    PosCamera.Y = 500
    PosCamera.X = 0
    PosCamera.z = 0
    Call Réseau_Charge(NomFichier$(Index))
End Sub

'
' ****************************
' Création d'un nouveau réseau
' ****************************
'
Private Sub MENU_Nouveau_Click()
    If MsgBox(Localisation$(ClePRINCIPALE% + 6), vbQuestion + vbOKCancel) = vbCancel Then Exit Sub
    ModeSaisie(0).Value = True
    Call Initialisation.ObjetDX_Detruit
    Pointe(0).No% = 0
    Pointe(0).Nom$ = ""
    Pointe(1).No% = 0
    Pointe(1).Nom$ = ""
    CameraRecul% = 250
    PosCamera.Y = 500
    PosCamera.X = 0
    PosCamera.z = 0
    Origine% = 0
    CheminReseau$ = ""
    Call Initialisation.Raz
    Call Initialisation.ObjetDX_Cree
End Sub

'
' *********************************
' Affiche la fenêtre des paramètres
' *********************************
'
Private Sub MENU_Parametre_Click()
    Parametres.Show
End Sub

'
' *******************
' Quitte le programme
' *******************
'
Private Sub MENU_Quitter_Click()
    Unload Me
End Sub

'
' ***********************************
' Appel la fenêtre de saisie du train
' ***********************************
'
Private Sub MENU_Train_Click()
    ModeSaisie(0).Value = True
    Call Initialisation.ObjetDX_Detruit
    Call SaisieTrain.Show(vbModal)
    Call Initialisation.ObjetDX_Cree
End Sub

'
' *******************************************
' Appel de la fenêtre de définition des voies
' *******************************************
'
Private Sub MENU_Voie_Click()
    Call SaisieRail.Show
End Sub

'
' *************************************
' Appel la fenêtre de saisie des wagons
' *************************************
'
Private Sub MENU_Wagon_Click()
    Call SaisieWagon.Show
End Sub

'
' *************************
' Indique le mode de la vue
' *************************
'
Private Sub ModeSaisie_Click(Index As Integer)
    Dim i%, t As Boolean
    '
    If Index = ModeVisualisation Then
        If UBound(ListeTrain()) = 0 Then
            Call MsgBox(Localisation$(ClePRINCIPALE% + 7), vbExclamation + vbOKOnly + vbDefaultButton1)
            ModeSaisie(0).Value = True
            Exit Sub
        End If
        t = False
        For i% = 1 To UBound(ListeTrain())
            If ListeWagon(ListeTrain(i%).NoWagon%).Motrice% <> 0 Then
                t = True
            End If
        Next i%
        If t = False Then
            Call MsgBox(Localisation$(ClePRINCIPALE% + 8), vbExclamation + vbOKOnly + vbDefaultButton1)
            ModeSaisie(0).Value = True
            Exit Sub
        End If
    End If
    '
    ' ***** Change de mode
    '
    Call Initialisation.Vue_Decharge
    ModeActuelle = Index
    '
    If ModeActuelle = ModeVisualisation Then
        '
        ' ***** Test si on est électrifié
        '
        ReseauElectrique = False
        If ParamCatenaire = True Then
            For i% = 1 To UBound(ListeTrain())
                If ListeWagon(ListeTrain(i%).NoWagon%).Motrice% <> 0 Then
                    If ListeWagon(ListeTrain(i%).NoWagon%).Electrique% <> 0 Then
                        ReseauElectrique = True
                    End If
                End If
            Next i%
        End If
        '
        For i% = 1 To UBound(Reseau())
            Reseau(i%).AiguilleForce% = Reseau(i%).Aiguille%
        Next i%
        ListeBogie(0).BogieReseau% = Origine%
        ListeBogie(0).BogiePosition = 0
        ListeBogie(0).BogieSegment% = 1
        If ParamSens = False Then
            ListeBogie(0).BogieSens% = 1
        Else
            ListeBogie(0).BogieSens% = -1
        End If
        For i% = 0 To 3
            ModeVue(i%).Enabled = True
        Next i%
        RegleVitesse = 0
        RegleVitesse.Visible = True
        CTunnel.Visible = False
        VitesseRotatif.Caption = RegleVitesse
        RegleRecul.Enabled = True
        For i% = 0 To 1
            CameraAngle%(i%) = 0
            CameraTourne%(i%) = 0
        Next i%
        CameraTourne%(2) = 0
        Vue.Affiche.ToolTipText = ""
        ChoixElement(0).Enabled = True
        ChoixElement(1).Enabled = True
        ChoixElement(0).ListIndex = 0
        ChoixElement(1).ListIndex = 0
        Call dxVue.dxScene.SetSceneBackgroundRGB(0, 0, 0.5)
        Call dxVue.dxBack.SetForeColor(vbWhite)
        Me.MENU_Element.Enabled = False
        Call Tunnel.Cree_Mesh
        Me.MENU_Parametre.Enabled = False
        Unload Parametres
    Else
        For i% = 0 To 3
            ModeVue(i%).Enabled = False
        Next i%
        RegleVitesse.Visible = False
        CTunnel.Visible = True
        RegleRecul.Enabled = False
        ChoixElement(0).Enabled = False
        ChoixElement(1).Enabled = False
        Call dxVue.dxScene.SetSceneBackgroundRGB(0, 0, 0)
        Call Tunnel.Detruit_Mesh
        Me.MENU_Element.Enabled = True
        Me.MENU_Parametre.Enabled = True
    End If
    '
    ' ***** Recharge le sol et la vue quand on change de mode
    '
    Call Initialisation.Vue_Charge
End Sub

'
' ************************
' Séléction du mode de vue
' ************************
'
Private Sub ModeVue_Click(Index As Integer)
    VueActuelle = Index
    If VueActuelle <> VueTable Then
        ChoixElement(0).Visible = True
    Else
        ChoixElement(0).Visible = False
    End If
    If VueActuelle = VuePoursuite Then
        ChoixElement(1).Visible = True
    Else
        ChoixElement(1).Visible = False
    End If
    If VueActuelle = VueSurvol Then
        RegleRecul.Visible = True
    Else
        RegleRecul.Visible = False
    End If
End Sub

'
' *************************
' Mise à jour de la vitesse
' *************************
'
Private Sub RegleVitesse_Change()
    VitesseRotatif.Caption = RegleVitesse.Value
End Sub

'
' ************************
' Ajoute un nouveau tunnel
' ************************
'
Private Sub TUNNEL_Ajoute_Click()
    Call Tunnel.Cree
    Call Tunnel.Tunnel_Pointe
End Sub

'
' *****************************
' Efface complétement un tunnel
' *****************************
'
Private Sub TUNNEL_Efface_Click()
    If NoTunnel% <> -1 Then
        Set ListeTunnel(NoTunnel%) = Nothing
        Call Tunnel.Tunnel_Pointe
        Call Initialisation.Vue_Decharge
        Call Vue.Calcule_Reseau
        Call Initialisation.Vue_Charge
    End If
End Sub

'
' ******************************
' Insere un point dans le tunnel
' ******************************
'
Private Sub TUNNEL_Insere_Click()
    If NoTunnel% <> -1 Then
        ListeTunnel(NoTunnel%).Point_Insere
        Call Tunnel.Tunnel_Pointe
    End If
End Sub

'
' **********************************
' Inverse le type de face vide/plein
' **********************************
'
Private Sub TUNNEL_Inverse_Click()
    If NoTunnel% <> -1 Then
        ListeTunnel(NoTunnel%).Inverse
    End If
End Sub

'
' **************************
' Supprime le segment pointé
' **************************
'
Private Sub TUNNEL_Supprime_Click()
    If NoTunnel% <> -1 Then
        ListeTunnel(NoTunnel%).Point_Supprime
        Call Tunnel.Tunnel_Pointe
        Call Initialisation.Vue_Decharge
        Call Vue.Calcule_Reseau
        Call Initialisation.Vue_Charge
    End If
End Sub

'
' ************************
' Ajoute une nouvelle voie
' ************************
'
Private Sub VUE_Ajoute_Click()
    Edition.Show vbModal
End Sub

'
' **************************************************
' Séléctionne la voie en cours comme celle d'origine
' **************************************************
'
Private Sub VUE_Origine_Click()
    Origine% = Pointe(0).No%
End Sub

'
' *****************
' Rotation du décor
' ou de la voie
' *****************
'
Private Sub VUE_Rotation_Click()
    Dim r$, n%
    If Pointe(0).Nom$ = "Decor" Then
        r$ = InputBox(Localisation$(ClePRINCIPALE% + 9), Localisation$(ClePRINCIPALE% + 10), ElementDecor(Pointe(0).No%).Angle%)
        If r$ = "" Then Exit Sub
        ElementDecor(Pointe(0).No%).Angle% = Val(r$) Mod 360
        Call Vue.Calcule_Reseau
        Call Initialisation.Vue_Decharge
        Call Initialisation.Vue_Charge
    Else
        r$ = InputBox(Localisation$(ClePRINCIPALE% + 15), Localisation$(ClePRINCIPALE% + 10), Reseau(Pointe(0).No%).Angle!)
        If r$ = "" Then Exit Sub
        n% = Origine%
        Origine% = Pointe(0).No%
        Reseau(Pointe(0).No%).Angle! = Val(r$) Mod 360
        Call Vue.Calcule_Reseau
        Call Initialisation.Vue_Decharge
        Call Initialisation.Vue_Charge
        Origine% = n%
    End If
End Sub

'
' ************************************
' Supprime une voie et ses connections
' ************************************
'
Private Sub VUE_Supprime_Click()
    Call Efface_Element
End Sub

'
' ********************************
' Met à jour la liste des fichiers
' ********************************
'
Public Sub MAJ_NomFichier(Nom$)
    Dim i%
    For i% = 0 To 3
        If NomFichier$(i%) = Nom$ Then Exit Sub ' Existe déjà
    Next i%
    For i% = 3 To 1 Step -1
        NomFichier$(i%) = NomFichier$(i% - 1)
    Next i%
    NomFichier$(0) = Nom$
    Call MAJ_MenuNomFichier
    FlagNomFichier = True
End Sub

'
' ********************************
' Met à jour la liste des fichiers
' ********************************
'
Public Sub MAJ_MenuNomFichier()
    Dim i%, t As Boolean
    For i% = 0 To 3
        If NomFichier$(i%) = "" Then
            MENU_NomFichier(i%).Visible = False
        Else
            t = True
            MENU_NomFichier(i%).Visible = True
            MENU_NomFichier(i%).Caption = "&" + Format$(i% + 1) + " " + NomFichier$(i%)
        End If
    Next i%
    MENU_Fichier_Moins1.Visible = t
End Sub

'
' ************************
' Charge le nouveau réseau
' ************************
'
Public Sub Réseau_Charge(Fichier$)
    Dim f%, n%, i%, j%
    Dim dX%, dz%
    Dim r%, Reference$, taille%
    ModeVue(0).Value = True
    ModeSaisie(0).Value = True
    Call Initialisation.ObjetDX_Detruit
    CheminReseau$ = Fichier$
    Me.Caption = CheminReseau$ + " :" + Localisation$(CleINVENTAIRE% + 3)
    f% = FreeFile()
    Call Tools.Open_File(Fichier$, f%, OPEN_NORMAL)
    '
    Input #f%, n%
    ReDim ListeTrain(n%) As TypeTrain
    If n% <> 0 Then
        taille% = UBound(ListeWagon())
        For i% = 1 To n%
            Input #f%, Reference$
            For r% = 1 To taille%
                If ListeWagon(r%).Ref$ = Reference$ Then
                    ListeTrain(i%).NoWagon% = r%
                    If ListeWagon(r%).dxWagon Is Nothing Then
                        Call Initialisation.Wagon_Charger_Mesh(r%)
                    End If
                End If
            Next r%
        Next i%
    End If
    '
    Input #f%, n%
    ReDim Reseau(n%) As TypeReseau
    Input #f%, Origine%
    Line Input #f%, Reference$
    Reseau(Origine%).Angle! = Val(Reference$)
    Input #f%, Reseau(Origine%).Position.X
    Input #f%, Reseau(Origine%).Position.Y
    Input #f%, Reseau(Origine%).Position.z
    '
    taille% = UBound(Voie())
    For i% = 1 To n%
        Input #f%, Reference$
        For r% = 1 To taille%
            If Voie(r%).Ref$ = Reference$ Then Reseau(i%).NoVoie% = r%
        Next r%
        For j% = 0 To NbSegment%
            Input #f%, Reseau(i%).Connecte%(j%), Reseau(i%).Entree%(j%)
        Next j%
        Reseau(i%).Aiguille% = 1
    Next i%
    '
    Input #f%, n%
    ReDim ElementDecor(n%) As TypeElementDecor
    If n% <> 0 Then
        taille% = UBound(ListeDecor())
        For i% = 1 To n%
            Input #f%, Reference$
            For r% = 1 To taille%
                If ListeDecor(r%).Ref$ = Reference$ Then
                    ElementDecor(i%).NoDecor% = r%
                    If ListeDecor(r%).dxDecor Is Nothing Then
                        Call Initialisation.Décor_Charge_Mesh(r%)
                    End If
                End If
            Next r%
            Input #f%, ElementDecor(i%).Position.X
            Input #f%, ElementDecor(i%).Position.Y
            Input #f%, ElementDecor(i%).Position.z
            Input #f%, ElementDecor(i%).Angle%
        Next i%
    End If
    '
    ' Charge les tunnels
    '
    Dim X!, z!, P%
    Input #f%, n%
    ReDim ListeTunnel(n%) As New ClassTunnel
    For i% = 0 To n% - 1
        Input #f%, P%
        For j% = 0 To P% - 1
            Input #f%, X!
            Input #f%, z!
            Call ListeTunnel(i%).Point_Ajoute(X!, z!)
            Input #f%, X!
            If X! = 0 Then
                ListeTunnel(i%).Face(j%) = False
            Else
                ListeTunnel(i%).Face(j%) = True
            End If
        Next j%
    Next i%
    '
    Input #f%, X!: ParamElevation = IIf(X! = 0, False, True)
    Input #f%, ParamHauteur%
    Input #f%, X!: ParamCatenaire = IIf(X! = 0, False, True)
    Input #f%, X!: ParamSens = IIf(X! = 0, False, True)
    Input #f%, X!: ParamCiel = IIf(X! = 0, False, True)
    '
    Close #f%
    '
    Call Initialisation.ObjetDX_Cree
    dX% = xPlateauMax% - xPlateauMin% + 2 * BORD%
    dz% = zPlateauMax% - zPlateauMin% + 2 * BORD%
    If dX% > dz% Then
        CameraRecul% = dX% / 2
    Else
        CameraRecul% = dz% / 2
    End If
    If CameraRecul% > 1500 Then CameraRecul% = 1500
    PosCamera.X = dX% / 2 + xPlateauMin% - BORD%
    PosCamera.z = dz% / 2 + zPlateauMin% - BORD
End Sub

