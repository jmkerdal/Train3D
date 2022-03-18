VERSION 5.00
Begin VB.Form Vue 
   Caption         =   "Vue du réseau"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   ControlBox      =   0   'False
   Icon            =   "Vue.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   Begin VB.PictureBox Affiche 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   0
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "Vue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' *************************
' Ajoute ou efface une voie
' *************************
'
Private Sub Affiche_DblClick()
    If ModeActuelle = ModeEdition Then
        If Principale.CTunnel.Value = vbUnchecked Then
            If Pointe(0).Nom$ = "Voie" Then
                Principale.VUE_Origine.Enabled = True
            Else
                Principale.VUE_Origine.Enabled = False
            End If
            If Pointe(0).Nom$ = "Voie" Or Pointe(0).Nom$ = "Decor" Then
                Principale.VUE_Supprime.Enabled = True
                Principale.VUE_Rotation.Enabled = True
            Else
                Principale.VUE_Supprime.Enabled = False
                Principale.VUE_Rotation.Enabled = False
            End If
            Call PopupMenu(Principale.MENU_Vue)
        Else
            If SelectionTunnel% = -1 Then
                Principale.TUNNEL_Insere.Enabled = False
                Principale.TUNNEL_Supprime.Enabled = False
                Principale.TUNNEL_Inverse.Enabled = False
                '
                Principale.TUNNEL_Efface.Enabled = False
                If NoTunnel% <> -1 Then
                    If ListeTunnel(NoTunnel%).SegmentPointe% <> -1 Then
                        Principale.TUNNEL_Insere.Enabled = True
                        If ListeTunnel(NoTunnel).Nb_Point% > 3 Then
                            Principale.TUNNEL_Supprime.Enabled = True
                        Else
                            Principale.TUNNEL_Supprime.Enabled = False
                        End If
                        Principale.TUNNEL_Inverse.Enabled = True
                        '
                        Principale.TUNNEL_Efface.Enabled = True
                    End If
                End If
                Call PopupMenu(Principale.MENU_Tunnel)
            End If
        End If
    End If
End Sub

'
' ****************
' Tourne la caméra
' ****************
'
Private Sub Affiche_KeyDown(KeyCode As Integer, Shift As Integer)
    If VueActuelle = VuePoursuite Then Exit Sub
    Select Case KeyCode
    Case vbKeyLeft
        If Shift And 1 = 1 Then
            CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) - 20
        Else
            CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) - 5
        End If
        If CameraTourne%(VueActuelle) < 0 Then CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) + 360
    Case vbKeyUp
        If VueActuelle <= VueSurvol Then
            If CameraAngle%(VueActuelle) < 80 Then CameraAngle%(VueActuelle) = CameraAngle%(VueActuelle) + 5
        End If
    Case vbKeyRight
        If Shift And 1 = 1 Then
            CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) + 20
        Else
            CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) + 5
        End If
        If CameraTourne%(VueActuelle) >= 360 Then CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) - 360
    Case vbKeyDown
        If VueActuelle <= VueSurvol Then
            If CameraAngle%(VueActuelle) > 0 Then CameraAngle%(VueActuelle) = CameraAngle%(VueActuelle) - 5
        End If
    Case vbKeyF
        AfficheFps = Not AfficheFps
    End Select
End Sub

'
' **************************************
' Modifie la vitesse du train au clavier
' **************************************
'
Private Sub Affiche_KeyPress(KeyAscii As Integer)
    If ModeActuelle = ModeEdition Then Exit Sub
    Select Case KeyAscii
    Case 43 ' +
        If Principale.RegleVitesse.Value <> Principale.RegleVitesse.Max Then
            Principale.RegleVitesse.Value = Principale.RegleVitesse.Value + 1
        End If
    Case 45 ' -
        If Principale.RegleVitesse.Value <> Principale.RegleVitesse.Min Then
            Principale.RegleVitesse.Value = Principale.RegleVitesse.Value - 1
        End If
    End Select
End Sub

'
' ****************************
' Edition du réseau
' ou séléction de l'aiguillage
' ****************************
'
Private Sub Affiche_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Vue.Selection(Shift)
    If Button = vbLeftButton And ModeActuelle = ModeVisualisation Then
        '
        ' ***** Bascule l'aiguillage
        '
        If Pointe(0).Nom$ = "Voie" Then
            If Reseau(Pointe(0).No%).Dessus = False Then
                If Reseau(Pointe(0).No%).Aiguille% >= Voie(Reseau(Pointe(0).No%).NoVoie%).AiguillePosition% Then
                    Reseau(Pointe(0).No%).Aiguille% = 1
                Else
                    Reseau(Pointe(0).No%).Aiguille% = Reseau(Pointe(0).No%).Aiguille% + 1
                End If
            End If
        End If
    End If
    If Button = vbLeftButton And ModeActuelle = ModeEdition Then
        '
        ' ***** Edition
        '
        If Principale.CTunnel.Value = vbUnchecked Then
            If Pointe(1).Nom$ = "Voie" Then
                If Pointe(0).Nom$ = "Voie" And Pointe(0).Couleur = SelectionVert Then ' C'est vert
                    If Pointe(1).No% = Pointe(0).No% Then
                        '
                        ' ***** On change de bout
                        '
                        Pointe(1) = Pointe(0)
                    Else
                        '
                        ' ***** Connecte deux voies
                        '
                        Call Initialisation.Voie_Ajoute(Pointe(1).No%, Pointe(1).Ou%, Pointe(0).No%, Pointe(0).Ou%)
                        Pointe(1).Nom$ = ""
                        Pointe(1).No% = 0
                        Pointe(1).Ou% = -1
                        Pointe(1).Couleur = SelectionSans
                    End If
                    Call Vue.Calcule_Reseau
                    Call Initialisation.Vue_Decharge
                    Call Initialisation.Vue_Charge
                Else
                    Pointe(1).Nom$ = ""
                    Pointe(1).No% = 0
                    Pointe(1).Ou% = -1
                    Pointe(1).Couleur = SelectionSans
                End If
            Else
                '
                ' ***** Sélectionne l'objet pointé
                '
                If Pointe(0).Couleur = SelectionVert Then
                    Pointe(1) = Pointe(0)
                End If
            End If
        Else
            If SelectionTunnel% <> -1 Then
                Call ListeTunnel(SelectionTunnel%).Point_Deplace(ListeTunnel(SelectionTunnel%).PointSelection%, PosSouris.X * REDUCTION%, PosSouris.z * REDUCTION%)
                ListeTunnel(SelectionTunnel%).PointSelection% = -1
                SelectionTunnel% = -1
                Call Initialisation.Vue_Decharge
                Call Vue.Calcule_Reseau
                Call Initialisation.Vue_Charge
            ElseIf NoTunnel% <> -1 Then
                If ListeTunnel(NoTunnel%).PointSelection% = -1 Then
                    If ListeTunnel(NoTunnel).PointPointe% <> -1 Then
                        ListeTunnel(NoTunnel%).PointSelection% = ListeTunnel(NoTunnel).PointPointe%
                        SelectionTunnel% = NoTunnel%
                    End If
                End If
            End If
        End If
    End If
End Sub

'
' *********************************
' Mémorise la position de la souris
' Déplace la camera
' *********************************
'
Private Sub Affiche_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim P1 As D3DRMVECTOR4D
    Dim i%, n%
    '
    OldX% = SourisX%
    OldY% = SourisY%
    SourisX% = X
    SourisY% = Y
    SourisClick% = Button
    '
    P1.X = SourisX%
    P1.Y = SourisY%
    P1.z = 0: P1.w = 1
    Call dxVue.dxViewport.InverseTransform(PosSouris, P1)
    '
    P1.X = OldX%
    P1.Y = OldY%
    P1.z = 0: P1.w = 1
    Call dxVue.dxViewport.InverseTransform(OldPosSouris, P1)
    '
    If ModeActuelle = ModeEdition Then
        If Principale.CTunnel.Value = vbUnchecked Then
            Call Vue.Selection(Shift)
        Else
            Call Tunnel.Tunnel_Pointe
        End If
    End If
    '
    ' ***** Déplace la vue/objet
    '
    If ModeActuelle = ModeVisualisation Then
        If SourisClick% = 0 Then
            Me.MousePointer = vbDefault
        ElseIf SourisClick% = vbLeftButton + vbRightButton Then
            If VueActuelle = VueSurvol Then
                Me.MousePointer = vbSizeNS
                Principale.RegleRecul = Principale.RegleRecul + (SourisY% - OldY%) * 2
            End If
        ElseIf SourisClick% = vbRightButton Then
            If VueActuelle <> VuePoursuite Then
                Me.MousePointer = vbSizeAll
                If VueActuelle <> VueSurBogie Then
                    CameraAngle%(VueActuelle) = CameraAngle%(VueActuelle) - (SourisY% - OldY%)
                    If CameraAngle%(VueActuelle) < 0 Then CameraAngle%(VueActuelle) = 0
                    If CameraAngle%(VueActuelle) > 80 Then CameraAngle%(VueActuelle) = 80
                End If
                CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) + (SourisX% - OldX%)
                If CameraTourne%(VueActuelle) < 0 Then CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) + 360
                If CameraTourne%(VueActuelle) > 360 Then CameraTourne%(VueActuelle) = CameraTourne%(VueActuelle) - 360
            End If
        End If
    Else
        If SourisClick% = 0 Then
            Me.MousePointer = vbDefault
        ElseIf SourisClick% = vbLeftButton Then
            '
            ' ***** Déplace voie/décor
            '
            If Principale.CTunnel.Value = vbUnchecked Then
                Me.MousePointer = vbSizeAll
                If Pointe(0).Nom$ = "Voie" Then
                    n% = Origine%
                    Origine% = Pointe(0).No%
                    Reseau(Origine%).Position.X = Reseau(Origine%).Position.X - OldPosSouris.X * REDUCTION% + PosSouris.X * REDUCTION%
                    Reseau(Origine%).Position.z = Reseau(Origine%).Position.z - OldPosSouris.z * REDUCTION% + PosSouris.z * REDUCTION%
                    If Shift = 2 Then
                        '
                        ' ***** Détache la voie
                        '
                        For i% = 0 To NbSegment%
                            Call Initialisation.Voie_Supprime(Pointe(0).No%, i%, Reseau(Pointe(0).No%).Connecte(i%), Reseau(Pointe(0).No%).Entree(i%))
                        Next i%
                    End If
                    Call Vue.Calcule_Reseau
                    Call Initialisation.Vue_Decharge
                    Call Initialisation.Vue_Charge
                    Origine% = n%
                ElseIf Pointe(0).Nom$ = "Decor" Then
                    ElementDecor(Pointe(0).No%).Position.X = ElementDecor(Pointe(0).No%).Position.X - OldPosSouris.X * REDUCTION% + PosSouris.X * REDUCTION%
                    ElementDecor(Pointe(0).No%).Position.z = ElementDecor(Pointe(0).No%).Position.z - OldPosSouris.z * REDUCTION% + PosSouris.z * REDUCTION%
                    Call Vue.Calcule_Reseau
                    Call Initialisation.Vue_Decharge
                    Call Initialisation.Vue_Charge
                End If
            End If
        ElseIf SourisClick% = vbRightButton Then
            '
            ' ***** Déplace la caméra
            '
            PosCamera.X = PosCamera.X + (OldPosSouris.X - PosSouris.X) * REDUCTION%
            PosCamera.z = PosCamera.z + (OldPosSouris.z - PosSouris.z) * REDUCTION%
            If PosCamera.X > xPlateauMax% + BORD% Then PosCamera.X = xPlateauMax% + BORD%
            If PosCamera.X < xPlateauMin% - BORD% Then PosCamera.X = xPlateauMin% - BORD%
            If PosCamera.z > zPlateauMax% + BORD% Then PosCamera.z = zPlateauMax% + BORD%
            If PosCamera.z < zPlateauMin% - BORD% Then PosCamera.z = zPlateauMin% - BORD%
        ElseIf SourisClick% = vbLeftButton + vbRightButton Then
            '
            ' ***** Recul la caméra
            '
            Me.MousePointer = vbSizeNS
            CameraRecul% = CameraRecul% + SourisY% - OldY%
            If CameraRecul% > 1500 Then CameraRecul% = 1500
            If CameraRecul% < 50 Then CameraRecul% = 50
        End If
    End If
End Sub

'
' *****************
' Lance l'affichage
' en boucle infinie
' *****************
'
Private Sub Form_Activate()
    If Tourne = False Then Call Boucle
End Sub

'
' ***************
' Retaille la vue
' ***************
'
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    CompteLancement% = CompteLancement% + 1
    If CompteLancement% = 1 Then Exit Sub
    TailleX% = Vue.ScaleWidth
    TailleY% = Vue.ScaleHeight
    Affiche.Width = TailleX%
    Affiche.Height = TailleY%
    Call ProgramLog.Write_File(CompteLancement%, "Size:" + Str$(TailleX%) + Str$(TailleY%))
    Call dxVue.Create_3DRM(Affiche, TailleX%, TailleY%, Mode3DSurface)
End Sub

'
' *****************
' Calcule du réseau
' *****************
'
Public Sub Calcule_Reseau()
    Dim i%, j%
    Dim ii%, jj%
    Dim start%, t As Boolean
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim d!
    Dim PremierElement As Boolean
    Dim dxBox As D3DRMBOX
    '
    If UBound(Reseau()) <> 0 Then
        If Reseau(Origine%).NoVoie% <> 0 Then
            start% = Origine%
        Else
            start% = 0
        End If
        For i% = 1 To UBound(Reseau())
            Reseau(i%).Calcule = False ' Par défaut pas encore calculé
            If start% = 0 Then
                If Reseau(i%).NoVoie% <> 0 Then start% = i%
            End If
        Next i%
        Origine% = start% ' Mémorise la position de départ si elle change
        If start% <> 0 Then
            v1 = Reseau(Origine%).Position
            d! = Reseau(Origine%).Angle! + Voie(Reseau(Origine%).NoVoie%).AngleNormal(0) + 180
            Call Calcule_Suivant(Origine%, 0, v1, d!)
        End If
        '
        ' ***** Tente de calculer les bouts non connectés à l'origine
        '
        For i% = 1 To UBound(Reseau())
            If Reseau(i%).Calcule = False Then
                v1 = Reseau(i%).Position
                d! = Reseau(i%).Angle! + Voie(Reseau(i%).NoVoie%).AngleNormal(0) + 180
                Call Calcule_Suivant(i%, 0, v1, d!)
            End If
        Next i%
        '
        ' ***** Cherche les connections automatiques
        '
'Debug.Print "Liste des points"
        ReseauFermée = True ' Par défaut c'est bon
        For i% = 1 To UBound(Reseau())
'Debug.Print Format$(i%);
            If Reseau(i%).NoVoie% <> 0 Then
                For j% = 0 To NbSegment%
                    v1 = Voies.Position_Point(i%, j%)
                    If (i% = start% And j% = 0) Or v1.X <> 0 Or v1.z <> 0 Then
'Debug.Print " ["; Format$(v1.x, "0.0"); ","; Format$(v1.z, "0.0"); "]";
                        If Reseau(i%).Connecte(j%) = 0 And Voie(Reseau(i%).NoVoie%).Terminaison(j%) = 0 And Voie(Reseau(i%).NoVoie%).Jonction(j%) = 0 Then
'Debug.Print " Non connecté";
                            t = False ' Cherche si on trouve la connection
                            For ii% = i% + 1 To UBound(Reseau())
                                If Reseau(ii%).NoVoie% <> 0 Then
                                    For jj% = 0 To NbSegment%
                                        v2 = Voies.Position_Point(ii%, jj%)
                                        If (ii% = start% And jj% = 0) Or v2.X <> 0 Or v2.z <> 0 Then
                                            If Reseau(ii%).Connecte(jj%) = 0 And Voie(Reseau(ii%).NoVoie%).Terminaison(jj%) = 0 And Voie(Reseau(ii%).NoVoie%).Jonction(jj%) = 0 Then
                                                d! = Sqr((v1.X - v2.X) ^ 2 + (v1.z - v2.z) ^ 2)
'If d! < 10 Then Debug.Print d!
                                                If d! < 1 Then
'Debug.Print i%; "se connecte sur "; ii%;
                                                    'Call MsgBox(Localisation$(CleVUE% + 1) + Str$(i%) + " " + Localisation$(CleVUE% + 2) + Str$(ii%) + _
                                                    vbCr + Localisation$(CleVUE% + 3) + "=" + Format$(d!, "0.000"), vbInformation + vbOKOnly)
                                                    Call Aide.Ecrit(Localisation$(CleVUE% + 1) + Str$(i%) + " " + Localisation$(CleVUE% + 2) + Str$(ii%) + _
                                                    " " + Localisation$(CleVUE% + 3) + "=" + Format$(d!, "0.000"))
                                                    t = True
                                                    Call Initialisation.Voie_Ajoute(i%, j%, ii%, jj%)
                                                End If
                                            End If
                                        End If
                                    Next jj%
                                End If
                            Next ii%
                            If t = False Then
                                '
                                ' Le réseau est ouvert
                                '
                                ReseauFermée = False
                            End If
                        End If
                        If PremierElement = False Then
                            PremierElement = True
                            xPlateauMin% = v1.X
                            xPlateauMax% = v1.X
                            zPlateauMin% = v1.z
                            zPlateauMax% = v1.z
                        Else
                            If v1.X < xPlateauMin% Then xPlateauMin% = v1.X
                            If v1.X > xPlateauMax% Then xPlateauMax% = v1.X
                            If v1.z < zPlateauMin% Then zPlateauMin% = v1.z
                            If v1.z > zPlateauMax% Then zPlateauMax% = v1.z
                        End If
                    End If
                Next j%
            End If
        Next i%
        '
        ' ***** Valide ou non la possibilité de passer en mode visualisation
        '
        Principale.ModeSaisie(1).Enabled = ReseauFermée
    End If
    '
    ' ***** Ajoute l'offset des décors
    '
    Dim v(3) As D3DVECTOR
    Dim AxeY As D3DVECTOR
    Dim x1%, z1%, x2%, z2%
    '
    AxeY.X = 0: AxeY.Y = 1: AxeY.z = 0
    '
    For i% = 1 To UBound(ElementDecor())
        If ElementDecor(i%).NoDecor% <> 0 Then
            Call ListeDecor(ElementDecor(i%).NoDecor%).dxDecor.GetBox(dxBox)
            '
            v(0) = dxBox.Min
            v(1) = dxBox.Min: v(1).X = dxBox.Max.X
            v(2) = dxBox.Max
            v(3) = dxBox.Max: v(3).X = dxBox.Min.X
            '
            For j% = 0 To 3
                Call dxVue.dX7.VectorScale(v(j%), v(j%), REDUCTION%) ' Remet a la bonne taille
                d! = dxVue.dX7.VectorModulus(v(j%)) ' Tourne le rectangle
                Call dxVue.dX7.VectorRotate(v(j%), v(j%), AxeY, ElementDecor(i%).Angle% * DegRad!)
                Call dxVue.dX7.VectorScale(v(j%), v(j%), d!)
                v(j%).X = v(j%).X + ElementDecor(i%).Position.X ' - dX% ' positionne le rectangle
                v(j%).z = v(j%).z + ElementDecor(i%).Position.z ' - dz%
                If j% = 0 Then
                    x1% = v(0).X
                    z1% = v(0).z
                    x2% = v(0).X
                    z2% = v(0).z
                Else
                    If v(j%).X < x1% Then x1% = v(j%).X
                    If v(j%).z < z1% Then z1% = v(j%).z
                    If v(j%).X > x2% Then x2% = v(j%).X
                    If v(j%).z > z2% Then z2% = v(j%).z
                End If
            Next j%
            '
            If PremierElement = False Then
                PremierElement = True
                xPlateauMin% = x1%
                xPlateauMax% = x2%
                zPlateauMin% = z1%
                zPlateauMax% = z2%
            Else
                If x1% < xPlateauMin% Then xPlateauMin% = x1%
                If x2% > xPlateauMax% Then xPlateauMax% = x2%
                If z1% < zPlateauMin% Then zPlateauMin% = z1%
                If z2% > zPlateauMax% Then zPlateauMax% = z2%
            End If
        End If
    Next i%
    '
    ' ***** Ajoute les tunnels
    '
    For i% = 0 To UBound(ListeTunnel()) - 1
        If ListeTunnel(i%).Nb_Point% <> 0 Then
            For j% = 0 To ListeTunnel(i%).Nb_Point% - 1
                If PremierElement = False Then
                    PremierElement = True
                    xPlateauMin% = ListeTunnel(i%).PositionX(j%)
                    zPlateauMin% = ListeTunnel(i%).PositionZ(j%)
                    xPlateauMax% = xPlateauMin%
                    zPlateauMax% = zPlateauMin%
                Else
                    If ListeTunnel(i%).PositionX(j%) < xPlateauMin% Then xPlateauMin% = ListeTunnel(i%).PositionX(j%)
                    If ListeTunnel(i%).PositionX(j%) > xPlateauMax% Then xPlateauMax% = ListeTunnel(i%).PositionX(j%)
                    If ListeTunnel(i%).PositionZ(j%) < zPlateauMin% Then zPlateauMin% = ListeTunnel(i%).PositionZ(j%)
                    If ListeTunnel(i%).PositionZ(j%) > zPlateauMax% Then zPlateauMax% = ListeTunnel(i%).PositionZ(j%)
                End If
            Next j%
        End If
    Next i%
    If PremierElement = False Then
        xPlateauMin% = 0
        xPlateauMax% = 0
        zPlateauMin% = 0
        zPlateauMax% = 0
    End If
End Sub

'
' ***************************
' Positionne la voie suivante
' ***************************
'
Public Sub Calcule_Suivant(n%, PtEntree%, vPoint As D3DVECTOR, Dir!)
    Dim i%, Suivant%
    Dim vNPoint As D3DVECTOR
    Dim VRotation As D3DVECTOR
    Dim Vy As D3DVECTOR
    Dim NDir!
    Dim l!
    '
    If Reseau(n%).Calcule = True Then Exit Sub ' Déjà fait
'Debug.Print n%; PtEntree%; " Pos="; vPoint.X; " "; vPoint.z; " Dir="; Dir!
'DoEvents
    '
    Reseau(n%).Angle! = Dir! - Voie(Reseau(n%).NoVoie%).AngleNormal(PtEntree%) + 180
    If Reseau(n%).Angle! >= 360 Then Reseau(n%).Angle! = Reseau(n%).Angle! - 360
    If PtEntree% <> 0 Then ' Entrée à rebours
'Debug.Print "Entre à rebours"
'Debug.Print "AngleInitial="; Dir!;
'Debug.Print "AngleEntree="; Voie(Reseau(n%).NoVoie%).AngleNormal(PtEntree%)
        '
        If Voie(Reseau(n%).NoVoie%).Offset(PtEntree%) = 0 Then
            VRotation.X = Voie(Reseau(n%).NoVoie%).pX!(PtEntree%)
            VRotation.z = Voie(Reseau(n%).NoVoie%).pZ!(PtEntree%)
        Else
            VRotation.X = Voie(Reseau(n%).NoVoie%).dX!(PtEntree%)
            VRotation.z = Voie(Reseau(n%).NoVoie%).dz!(PtEntree%)
        End If
        l! = dxVue.dX7.VectorModulus(VRotation)
        Vy.Y = 1
        Call dxVue.dX7.VectorRotate(vNPoint, VRotation, Vy, _
        Reseau(n%).Angle! * DegRad!)
        Call dxVue.dX7.VectorScale(vNPoint, vNPoint, l!)
        Call dxVue.dX7.VectorSubtract(vPoint, vPoint, vNPoint)
        Reseau(n%).Position = vPoint
'Debug.Print "Angle="; Reseau(n%).Angle!
'Debug.Print "Offset="; vNPoint.X; vNPoint.z
'Debug.Print "Origine="; vPoint.X; vPoint.z
    End If
    Reseau(n%).Position = vPoint
    Reseau(n%).Calcule = True
    '
    For i% = 0 To NbSegment%
        Suivant% = Reseau(n%).Connecte(i%)
        If Suivant% <> 0 Then 'And i% <> PtEntree% Then
            If i% <> 0 Then
                If Voie(Reseau(n%).NoVoie%).Offset(i%) = 0 Then
                    VRotation.X = Voie(Reseau(n%).NoVoie%).pX!(i%)
                    VRotation.z = Voie(Reseau(n%).NoVoie%).pZ!(i%)
                Else
                    VRotation.X = Voie(Reseau(n%).NoVoie%).dX!(i%)
                    VRotation.z = Voie(Reseau(n%).NoVoie%).dz!(i%)
                End If
'Debug.Print "a"; VRotation.X; VRotation.z; Dir!
                l! = dxVue.dX7.VectorModulus(VRotation)
                Vy.Y = 1
                Call dxVue.dX7.VectorRotate(vNPoint, VRotation, Vy, _
                Reseau(n%).Angle! * DegRad!)
                Call dxVue.dX7.VectorScale(vNPoint, vNPoint, l!)
'Debug.Print "b"; vNPoint.X; vNPoint.z
                Call dxVue.dX7.VectorAdd(vNPoint, vNPoint, vPoint)
'Debug.Print "c"; vNPoint.X; vNPoint.z
            Else
                vNPoint = vPoint
            End If
            NDir! = Reseau(n%).Angle! + Voie(Reseau(n%).NoVoie%).AngleNormal(i%)
            If NDir! > 180 Then NDir! = NDir! - 360
            If NDir! < -180 Then NDir! = NDir! + 360
'Debug.Print n%; "Suivant="; suivant%; Reseau(n%).Entree(i%); vNPoint.X; vNPoint.z; NDir!
            Call Calcule_Suivant(Suivant%, Reseau(n%).Entree(i%), vNPoint, NDir!)
        End If
    Next i%
End Sub

'
' ****************************
' Positionne la bogie suivante
' ****************************
'
Public Sub Positionne_Bogie_Suivante(s%, ByVal d!)
    Dim PBogie(1) As D3DVECTOR
    Dim DBogie!
    Dim Longueur!
    Dim Inverse As Boolean
    '
    ListeBogie(s%).BogieReseau% = ListeBogie(s% - 1).BogieReseau%
    ListeBogie(s%).BogieSegment% = ListeBogie(s% - 1).BogieSegment%
    ListeBogie(s%).BogiePosition! = ListeBogie(s% - 1).BogiePosition!
    ListeBogie(s%).BogieSens% = ListeBogie(s% - 1).BogieSens%
    If d! < 0 Then
        ListeBogie(s%).BogieSens% = -ListeBogie(s%).BogieSens%
        d! = -d!
        Inverse = True
    End If
    Call Voies.Deplace_Bogie(ListeBogie(s%), d!)
    '
    ' ***** Ajuste 1 fois
    ' ***** ça suffit pour être relativement précis
    '
    Call Voies.Position_Bogie(ListeBogie(s% - 1), PBogie(0), DBogie!)
    Call Voies.Position_Bogie(ListeBogie(s%), PBogie(1), DBogie!)
    Longueur! = Sqr((PBogie(1).X - PBogie(0).X) ^ 2 + (PBogie(1).z - PBogie(0).z) ^ 2)
    If Abs(d! - Longueur!) > 1 Then
'Debug.Print i%; "Distance:"; Longueur!; " reste:"; d! - Longueur!
        Call Voies.Deplace_Bogie(ListeBogie(s%), d! - Longueur!)
'Call Voies.Position_Bogie(listebogie(s%), PBogie(1), DBogie!)
'Longueur! = Sqr((PBogie(1).x - PBogie(0).x) ^ 2 + (PBogie(1).z - PBogie(0).z) ^ 2)
'Debug.Print "Distance:"; Longueur!; " reste:"; d! - Longueur!
    End If
    If Inverse = True Then
        ListeBogie(s%).BogieSens% = -ListeBogie(s%).BogieSens%
    End If
End Sub

'
' ********************************
' Calcul de la position des objets
' Affichage de la vue
' ********************************
'
Public Sub Vue_MAJ()
    Dim i%, j%
    Dim n%
    Dim p%, Trouve As Boolean
    Dim DBogie!
    Dim pWagon(1) As D3DVECTOR
    Dim dWagon As D3DVECTOR
    Dim P1!, s%
    Dim dxBox As D3DRMBOX
    Dim pv(3) As D3DVECTOR
    Dim Dist!, h As D3DVECTOR
    '
    'Dim t1&, t2&, t3&, t4&
    'Dim ta&, tb&, tc&
    '
    't1& = timeGetTime
    If ModeActuelle = ModeVisualisation Then
        '
        ' ***** Positionne les bogies
        '
        For i% = 1 To UBound(ListeBogie()) - 1
            Call Voies.Position_Bogie(ListeBogie(i%), ListeBogie(i%).BogieVecteur, DBogie!)
            Call dxFrameRoue(i%).SetPosition(dxVue.dxScene, ListeBogie(i%).BogieVecteur.X / REDUCTION%, ListeBogie(i%).BogieVecteur.Y / REDUCTION%, ListeBogie(i%).BogieVecteur.z / REDUCTION%)
            Call dxFrameRoue(i%).SetOrientation(dxVue.dxScene, 0, 0, 1, 0, 1, 0)
            Call dxFrameRoue(i%).AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, DBogie! * DegRad!)
        Next i%
        '
        ' ***** Positionne les wagons
        '
        For i% = 1 To UBound(ListeTrain())
            '
            ' ***** Pose le wagon et les bogies
            '
            For j% = 0 To 1
                If ListeTrain(i%).TrainBogie(j% * 2 + 1) <> -1 Then
                    pWagon(j%).X = (ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2)).BogieVecteur.X + ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2 + 1)).BogieVecteur.X) / 2
                    pWagon(j%).Y = (ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2)).BogieVecteur.Y + ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2 + 1)).BogieVecteur.Y) / 2
                    pWagon(j%).z = (ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2)).BogieVecteur.z + ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2 + 1)).BogieVecteur.z) / 2
                    If ListeWagon(ListeTrain(i%).NoWagon%).FichierBogie$ <> "" Then
                        Call dxFrameWagon(i%, j% + 3).SetPosition(dxVue.dxScene, pWagon(j%).X / REDUCTION%, (pWagon(j%).Y + 10) / REDUCTION%, pWagon(j%).z / REDUCTION%)
                        Call dxFrameWagon(i%, j% + 3).SetOrientation(dxVue.dxScene, _
                        ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2)).BogieVecteur.X - ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2 + 1)).BogieVecteur.X, _
                        ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2)).BogieVecteur.Y - ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2 + 1)).BogieVecteur.Y, _
                        ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2)).BogieVecteur.z - ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2 + 1)).BogieVecteur.z, _
                        0, 1, 0)
                        Call dxFrameWagon(i%, j% + 3).AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * DegRad!)
                    End If
                Else
                    pWagon(j%) = ListeBogie(ListeTrain(i%).TrainBogie%(j% * 2)).BogieVecteur
                End If
            Next j%
            dWagon.X = pWagon(0).X - pWagon(1).X
            dWagon.Y = pWagon(0).Y - pWagon(1).Y
            dWagon.z = pWagon(0).z - pWagon(1).z
            Call dxFrameWagon(i%, 0).SetPosition(dxVue.dxScene, pWagon(0).X / REDUCTION%, (pWagon(0).Y + 20) / REDUCTION%, pWagon(0).z / REDUCTION%)
            Call dxFrameWagon(i%, 0).SetOrientation(dxVue.dxScene, dWagon.X, dWagon.Y, dWagon.z, 0, 1, 0)
            '
            ' ***** Positionne les attelages
            '
            P1! = ListeWagon(ListeTrain(i%).NoWagon%).Longueur! - ListeWagon(ListeTrain(i%).NoWagon%).EccartEssieu! - (ListeWagon(ListeTrain(i%).NoWagon%).EccartBogie!(0) + ListeWagon(ListeTrain(i%).NoWagon%).EccartBogie!(1)) / 2
            s% = ListeBogie(ListeTrain(i%).TrainBogie(0)).BogieSens%
            Call dxFrameWagon(i%, 1).SetPosition(dxFrameRoue(ListeTrain(i%).TrainBogie(0)), (P1! / 2 * s%) / REDUCTION%, 10 / REDUCTION%, 0)
            Call dxFrameWagon(i%, 1).SetOrientation(dxFrameRoue(ListeTrain(i%).TrainBogie(0)), 0, 0, -s%, 0, 1, 0)
            If ListeTrain(i%).TrainBogie(3) <> -1 Then
                n% = ListeTrain(i%).TrainBogie(3)
            Else
                n% = ListeTrain(i%).TrainBogie(2)
            End If
            s% = ListeBogie(n%).BogieSens%
            Call dxFrameWagon(i%, 2).SetPosition(dxFrameRoue(n%), (-P1! / 2 * s%) / REDUCTION%, 10 / REDUCTION%, 0)
            Call dxFrameWagon(i%, 2).SetOrientation(dxFrameRoue(n%), 0, 0, s%, 0, 1, 0)
        Next i%
        '
        ' ***** Attache l'intérieur
        '
        If VueActuelle = VueSurBogie And IndexElement%(0) = 3 Then
            If Not (ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).dxInterieur Is Nothing) Then
                Call dxFrameInterieur.AddVisual(ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).dxInterieur)
                Call dxVue.dxScene.AddChild(dxFrameInterieur)
                Call dxFrameInterieur.SetPosition(dxFrameWagon(IndexChoix%(0), 0), 0, 0, 0)
                Call dxFrameInterieur.SetOrientation(dxFrameWagon(IndexChoix%(0), 0), 0, 0, 1, 0, 1, 0)
            End If
        End If
        If VueActuelle = VuePoursuite And IndexElement%(1) = 3 Then
            If Not (ListeWagon(ListeTrain(IndexChoix%(1)).NoWagon%).dxInterieur Is Nothing) Then
                Call dxFrameInterieur.AddVisual(ListeWagon(ListeTrain(IndexChoix%(1)).NoWagon%).dxInterieur)
                Call dxVue.dxScene.AddChild(dxFrameInterieur)
                Call dxFrameInterieur.SetPosition(dxFrameWagon(IndexChoix%(1), 0), 0, 0, 0)
                Call dxFrameInterieur.SetOrientation(dxFrameWagon(IndexChoix%(1), 0), 0, 0, 1, 0, 1, 0)
            End If
        End If
        '
        ' ***** Positionne l'indicateur d'aiguillage
        '
        For i% = 1 To UBound(dxFrameVoie(), 2)
            n% = Reseau(i%).NoVoie%
            If n% <> 0 Then
                Trouve = False
                For j% = 0 To NbSegment%
                    If Trouve = False And Voie(n%).SegmentPoint(j%, 0) <> 0 And Voie(n%).SegmentPoint(j%, 1) <> 0 Then
                        p% = Voie(n%).SegmentPoint(j%, 1) - 1
                        If Voie(n%).MatAiguille(Voie(n%).SegmentPoint(j%, 0) - 1, p%) = Reseau(i%).AiguilleForce% Then
                            dWagon = Voies.Position_Point(i%, p%)
                            Call dxFrameVoie(1, i%).SetPosition(dxVue.dxScene, dWagon.X / REDUCTION%, 0, dWagon.z / REDUCTION%)
                            Call dxFrameVoie(1, i%).SetOrientation(dxFrameVoie(0, i%), Voie(n%).Normal(p%).X, Voie(n%).Normal(p%).Y, Voie(n%).Normal(p%).z, 0, 1, 0)
                            Trouve = True
                        End If
                    End If
                Next j%
            End If
        Next i%
    End If
    '
    ' ***** Positionne la caméra
    '
    'ta& = timeGetTime
    Call Me.Vue_Camera
    'tb& = timeGetTime
    '
    ' ***** Affiche la sélection
    '
    If ModeActuelle = ModeEdition Then
        If Principale.CTunnel.Value = vbUnchecked Then
            For i% = 0 To 1
                If Pointe(i%).Nom$ = "Voie" And Pointe(i%).Couleur <> SelectionSans Then
                    Call dxFramePointe(i%).AddVisual(dxSelection(Pointe(i%).Couleur - 1))
                    Call dxFrameVoie(0, Pointe(i%).No%).AddChild(dxFramePointe(i%))
                    n% = Reseau(Pointe(i%).No%).NoVoie%
                    If Voie(n%).Offset(Pointe(i%).Ou%) = 0 Then
                        Call dxFramePointe(i%).SetPosition(dxFrameVoie(0, Pointe(i%).No%), Voie(n%).pX!(Pointe(i%).Ou%) / REDUCTION%, 0, Voie(n%).pZ!(Pointe(i%).Ou%) / REDUCTION%)
                        Call dxFramePointe(i%).SetOrientation(dxFrameVoie(0, Pointe(i%).No%), Voie(n%).Normal(Pointe(i%).Ou%).X, Voie(n%).Normal(Pointe(i%).Ou%).Y, Voie(n%).Normal(Pointe(i%).Ou%).z, 0, 1, 0)
                    Else
                        Call dxFramePointe(i%).SetPosition(dxFrameVoie(0, Pointe(i%).No%), Voie(n%).dX!(Pointe(i%).Ou%) / REDUCTION%, 0, Voie(n%).dz!(Pointe(i%).Ou%) / REDUCTION%)
                        Call dxFramePointe(i%).SetOrientation(dxFrameVoie(0, Pointe(i%).No%), Voie(n%).Normal(Pointe(i%).Ou%).X, Voie(n%).Normal(Pointe(i%).Ou%).Y, Voie(n%).Normal(Pointe(i%).Ou%).z, 0, 1, 0)
                    End If
                End If
            Next i%
            If Pointe(0).Nom$ = "Voie" And Pointe(0).Couleur > SelectionRouge And _
            Pointe(1).Nom$ = "Voie" And Pointe(1).Couleur > SelectionRouge Then
                Call dxFramePointe(0).GetPosition(dxVue.dxScene, pv(0))
                Call dxFramePointe(1).GetPosition(dxVue.dxScene, pv(1))
                Vue.Caption = "d= " + Format$(Sqr((pv(0).X - pv(1).X) ^ 2 + (pv(0).z - pv(1).z) ^ 2) * REDUCTION%, "#.## mm")
            Else
                Vue.Caption = Localisation$(CleVUE%)
            End If
            If Pointe(0).Nom$ = "Voie" Then
                Call Voie(Reseau(Pointe(0).No%).NoVoie%).VoieMeshBuilder.GetBox(dxBox)
                Set dxPointe = dxVue.dxD3Drm.CreateMeshBuilder()
                Call dxConstruitMeshBuilder.Init
                pv(0).X = dxBox.Min.X: pv(0).Y = 2 / REDUCTION%: pv(0).z = dxBox.Max.z
                pv(1).X = dxBox.Max.X: pv(1).Y = 2 / REDUCTION%: pv(1).z = dxBox.Max.z
                pv(2).X = dxBox.Max.X: pv(2).Y = 2 / REDUCTION%: pv(2).z = dxBox.Min.z
                pv(3).X = dxBox.Min.X: pv(3).Y = 2 / REDUCTION%: pv(3).z = dxBox.Min.z
                Call dxConstruitMeshBuilder.Add_Rectangle(pv(0), pv(1), pv(2), pv(3), -1, &H80FFFFFF)
                Call dxConstruitMeshBuilder.Build(dxPointe, dxSolTexture())
                Call dxFramePointe(2).AddVisual(dxPointe)
                Call dxFrameVoie(0, Pointe(0).No%).AddChild(dxFramePointe(2))
                Call dxFramePointe(2).SetPosition(dxFrameVoie(0, Pointe(0).No%), 0, 0, 0)
                Call dxFramePointe(2).SetOrientation(dxFrameVoie(0, Pointe(0).No%), 0, 0, 1, 0, 1, 0)
            End If
            If Pointe(0).Nom$ = "Decor" Then
                Call ListeDecor(ElementDecor(Pointe(0).No%).NoDecor%).dxDecor.GetBox(dxBox)
                Set dxPointe = dxVue.dxD3Drm.CreateMeshBuilder()
                Call dxConstruitMeshBuilder.Init
                pv(0).X = dxBox.Min.X: pv(0).Y = dxBox.Max.Y: pv(0).z = dxBox.Max.z
                pv(1).X = dxBox.Max.X: pv(1).Y = dxBox.Max.Y: pv(1).z = dxBox.Max.z
                pv(2).X = dxBox.Max.X: pv(2).Y = dxBox.Max.Y: pv(2).z = dxBox.Min.z
                pv(3).X = dxBox.Min.X: pv(3).Y = dxBox.Max.Y: pv(3).z = dxBox.Min.z
                Call dxConstruitMeshBuilder.Add_Rectangle(pv(0), pv(1), pv(2), pv(3), -1, &H8000FF00)
                Call dxConstruitMeshBuilder.Build(dxPointe, dxSolTexture())
                Call dxFramePointe(2).AddVisual(dxPointe)
                Call dxFrameDecor(Pointe(0).No%).AddChild(dxFramePointe(2))
                Call dxFramePointe(2).SetPosition(dxFrameDecor(Pointe(0).No%), 0, 0, 0)
                Call dxFramePointe(2).SetOrientation(dxFrameDecor(Pointe(0).No%), 0, 0, 1, 0, 1, 0)
            End If
        End If
    Else
        '
        ' ***** Positionne le ciel en fonction de la caméra
        '
        If ParamCiel = True Then
            Call dxVue.dxCamera.GetPosition(dxVue.dxScene, h)
            Dist! = dxVue.dxViewport.GetBack * Sqr(2)
            Call dxFrameCiel.AddTranslation(D3DRMCOMBINE_REPLACE, 0, 0, 0)
            Call dxFrameCiel.AddScale(D3DRMCOMBINE_REPLACE, Dist!, Dist!, Dist!)
            Call dxFrameCiel.SetPosition(dxVue.dxScene, 0, -50 / REDUCTION%, 0)
            Call dxFrameCiel.AddTranslation(D3DRMCOMBINE_BEFORE, h.X / Dist!, 0, h.z / Dist!)
        End If
    End If
    '
    't2& = timeGetTime
    Call dxVue.dxScene.AddChild(dxScene2)
    Call dxVue.Render(False)
    Call dxVue.dxScene.DeleteChild(dxScene2)
    '
    ' ***** Affiche les tunnels
    '
    If ModeActuelle = ModeEdition Then
        For i% = 0 To UBound(ListeTunnel()) - 1
            If ListeTunnel(i%).Nb_Point% <> 0 Then
                Call Tunnel.Affiche(i%)
            End If
        Next i%
    End If
    '
    ' ***** FPS
    '
    If AfficheFps = True Then
        Dim t$
        If TempsAffichage& > 200 Then
            t$ = Format$(TempsAffichage& / 1000) + " s"
        Else
            t$ = Format$(1000 \ TempsAffichage&) + " fps"
        End If
        Call dxVue.dxBack.SetForeColor(vbBlack)
        Call dxVue.dxBack.DrawText(1, 1, t$, False)
        Call dxVue.dxBack.SetForeColor(vbWhite)
        Call dxVue.dxBack.DrawText(0, 0, t$, False)
    End If
    '
    Call dxVue.Render(True)
    '
    't3& = timeGetTime
    '
    ' ***** Supprime la sélection
    '
    If ModeActuelle = ModeEdition Then
        If Principale.CTunnel.Value = vbUnchecked Then
            If Pointe(0).Nom$ = "Voie" Then
                Call dxFrameVoie(0, Pointe(0).No%).DeleteChild(dxFramePointe(2))
                Call dxFramePointe(2).DeleteVisual(dxPointe)
                Set dxPointe = Nothing
            End If
            If Pointe(0).Nom$ = "Decor" Then
                Call dxFrameDecor(Pointe(0).No%).DeleteChild(dxFramePointe(2))
                Call dxFramePointe(2).DeleteVisual(dxPointe)
                Set dxPointe = Nothing
            End If
            For i% = 0 To 1
                If Pointe(i%).Nom$ = "Voie" And Pointe(i%).Couleur <> SelectionSans Then
                    Call dxFrameVoie(0, Pointe(i%).No%).DeleteChild(dxFramePointe(i%))
                    Call dxFramePointe(i%).DeleteVisual(dxSelection(Pointe(i%).Couleur - 1))
                End If
            Next i%
        End If
    Else
        If VueActuelle = VueSurBogie And IndexElement%(0) = 3 Then
            If Not (ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).dxInterieur Is Nothing) Then
                Call dxVue.dxScene.DeleteChild(dxFrameInterieur)
                Call dxFrameInterieur.DeleteVisual(ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).dxInterieur)
            End If
        End If
        If VueActuelle = VuePoursuite And IndexElement%(1) = 3 Then
            If Not (ListeWagon(ListeTrain(IndexChoix%(1)).NoWagon%).dxInterieur Is Nothing) Then
                Call dxVue.dxScene.DeleteChild(dxFrameInterieur)
                Call dxFrameInterieur.DeleteVisual(ListeWagon(ListeTrain(IndexChoix%(1)).NoWagon%).dxInterieur)
            End If
        End If
    End If
    't4& = timeGetTime
'Debug.Print "MAJ:"; t2& - t1&; ":"; ta& - t1&; tb& - ta&; t2& - tb&
'Debug.Print t3& - t2&; t4& - t3&
End Sub

'
' *************************************
' Cherche la frame sur lequel la souris
' est positionnée
' Code based from Nigel Thompson book
' "3D graphics programming for W95"
' translate to VB from C++
' *************************************
'
Public Sub Selection(Shift As Integer)
    '
    ' ***** Sélection d'un élément
    '
    Dim dxPickArray As Direct3DRMPickArray
    Dim dxFrameArray As Direct3DRMFrameArray
    Dim dxFrame As Direct3DRMFrame3
    Dim dxD3DRMPickDesc As D3DRMPICKDESC
    Dim i%, v%
    Dim Vec As D3DVECTOR
    '
    Set dxPickArray = dxVue.dxViewport.Pick(SourisX%, SourisY%)
    If dxPickArray.GetSize <> 0 Then
        Set dxFrameArray = dxPickArray.GetPickFrame(0, dxD3DRMPickDesc)
        Set dxFrame = dxFrameArray.GetElement(dxFrameArray.GetSize - 1)
        '
        Pointe(0).No% = dxFrame.GetAppData
        Pointe(0).Nom$ = dxFrame.GetName
        Pointe(0).Couleur = SelectionSans
        '
        If Pointe(0).Nom$ = "Voie" Then
            v% = Reseau(Pointe(0).No%).NoVoie
            Pointe(0).Ou% = -1 ' Par défaut aucun
            For i% = 0 To NbSegment%
                If Voie(v%).Jonction(i%) = 0 Then ' C'est pas une jonction
                    Vec = Voies.Position_Point(Pointe(0).No%, i%)
                    If i% = 0 Or Vec.X <> 0 Or Vec.z <> 0 Then
                        If Sqr((PosSouris.X - Vec.X / REDUCTION%) ^ 2 + (PosSouris.z - Vec.z / REDUCTION%) ^ 2) < 10 / REDUCTION% Then
                            Pointe(0).Ou% = i%
                        End If
                    End If
                End If
            Next i%
        End If
        Set dxFrame = Nothing
        Set dxFrameArray = Nothing
    Else
        Pointe(0).No% = 0
        Pointe(0).Nom$ = ""
        Pointe(0).Ou% = -1
        Pointe(0).Couleur = SelectionSans
    End If
    '
    Set dxPickArray = Nothing
    '
    If ModeActuelle = ModeEdition Then
        '
        ' ***** Cherche la couleur de la sélection
        '
        If Pointe(1).Nom$ = "Voie" And Pointe(1).Ou% <> -1 Then
            Pointe(1).Couleur = SelectionBleu
        Else
            Pointe(1).Couleur = SelectionSans
        End If
        If Pointe(0).Nom$ = "Voie" And Pointe(0).Ou% <> -1 Then
            If Reseau(Pointe(0).No%).Connecte(Pointe(0).Ou%) = 0 And _
            Voie(Reseau(Pointe(0).No%).NoVoie%).Terminaison(Pointe(0).Ou%) = 0 Then
                If Pointe(0).No% <> Pointe(1).No% Then
                    Pointe(0).Couleur = SelectionVert
                Else
                    Pointe(0).Couleur = SelectionSans
                End If
            Else
                Pointe(0).Couleur = SelectionRouge
            End If
        Else
            Pointe(0).Couleur = SelectionSans
        End If
    End If
    '
    ' ***** Valide le menu
    '
    If (Pointe(0).Nom$ = "Voie" Or Pointe(0).Nom$ = "Decor") And ModeActuelle = ModeEdition Then
        Principale.MENU_Copier.Enabled = True
        Principale.MENU_Couper.Enabled = True
        If Pointe(0).Nom$ = "Voie" Then
            Vue.Affiche.ToolTipText = "[" + Voie(Reseau(Pointe(0).No%).NoVoie%).Ref + "] " + Voie(Reseau(Pointe(0).No%).NoVoie%).Libelle$(0)
        Else
            Vue.Affiche.ToolTipText = ListeDecor(ElementDecor(Pointe(0).No%).NoDecor%).Nom$
        End If
    Else
        Principale.MENU_Copier.Enabled = False
        Principale.MENU_Couper.Enabled = False
        Vue.Affiche.ToolTipText = ""
    End If
End Sub

'
' ******************************
' Positionne la bogie précedente
' ******************************
'
Public Sub Positionne_Bogie_Precedente(s%, ByVal d!)
    Dim PBogie(1) As D3DVECTOR
    Dim DBogie!
    Dim Longueur!
    Dim Inverse As Boolean
    '
    ListeBogie(s%).BogieReseau% = ListeBogie(s% + 1).BogieReseau%
    ListeBogie(s%).BogieSegment% = ListeBogie(s% + 1).BogieSegment%
    ListeBogie(s%).BogiePosition! = ListeBogie(s% + 1).BogiePosition!
    ListeBogie(s%).BogieSens% = ListeBogie(s% + 1).BogieSens%
    If d! < 0 Then
        ListeBogie(s%).BogieSens% = -ListeBogie(s%).BogieSens%
        d! = -d!
        Inverse = True
    End If
    Call Voies.Deplace_Bogie(ListeBogie(s%), d!)
    '
    ' ***** Ajuste 1 fois
    ' ***** ça suffit pour être relativement précis
    '
    Call Voies.Position_Bogie(ListeBogie(s% + 1), PBogie(0), DBogie!)
    Call Voies.Position_Bogie(ListeBogie(s%), PBogie(1), DBogie!)
    Longueur! = Sqr((PBogie(1).X - PBogie(0).X) ^ 2 + (PBogie(1).z - PBogie(0).z) ^ 2)
    If Abs(d! - Longueur!) > 1 Then
'Debug.Print i%; "Distance:"; Longueur!; " reste:"; d! - Longueur!
        Call Voies.Deplace_Bogie(ListeBogie(s%), d! - Longueur!)
'Call Voies.Position_Bogie(listebogie(s%), PBogie(1), DBogie!)
'Longueur! = Sqr((PBogie(1).x - PBogie(0).x) ^ 2 + (PBogie(1).z - PBogie(0).z) ^ 2)
'Debug.Print "Distance:"; Longueur!; " reste:"; d! - Longueur!
    End If
    If Inverse = True Then
        ListeBogie(s%).BogieSens% = -ListeBogie(s%).BogieSens%
    End If
End Sub

'
' ********************
' Positionne la caméra
' ********************
'
Public Sub Vue_Camera()
    Dim dX%, dz%, PosY%
    Dim HauteurCameraCentre%
    '
    HauteurCameraCentre% = PAS_CASE% / 2.5 * HMax% / REDUCTION%
    '
    If ModeActuelle = ModeEdition Then
        Call dxVue.dxCamera.SetPosition(dxVue.dxScene, PosCamera.X / REDUCTION%, PosCamera.Y / REDUCTION%, PosCamera.z / REDUCTION%)
        Call dxVue.dxCamera.SetOrientation(dxVue.dxScene, 0, -1, 0, 0, 0, 1)
        Call dxVue.dxViewport.SetProjection(D3DRMPROJECT_ORTHOGRAPHIC)
        With dxVue.dxViewport
            .SetFront 1 / REDUCTION%
            .SetBack 10000! / REDUCTION%
            .SetField CameraRecul% / REDUCTION%
        End With
    Else
        dX% = xPlateauMax% - xPlateauMin% + 2 * BORD%
        dz% = zPlateauMax% - zPlateauMin% + 2 * BORD%
        If dX% > dz% Then
            PosY% = dX%
        Else
            PosY% = dz%
        End If
        Select Case VueActuelle
        Case EnumVue.VueSurvol
            '
            ' ***** Caméra suit un élément sur le dessus
            '
            Select Case IndexElement%(0)
            Case 0
                Call dxVue.dxCamera.SetPosition(dxVue.dxScene, (dX% / 2 + xPlateauMin% - BORD%) / REDUCTION%, HauteurCameraCentre% + Principale.RegleRecul.Value / REDUCTION%, (dz% / 2 + zPlateauMin% - BORD%) / REDUCTION%)
                Call dxVue.dxCamera.SetOrientation(dxVue.dxScene, 0, -1, 0, 0, 0, 1)
            Case 1
                Call dxVue.dxCamera.SetPosition(dxFrameVoie(0, IndexChoix%(0)), 0, Principale.RegleRecul.Value / REDUCTION%, 0)
                Call dxVue.dxCamera.SetOrientation(dxVue.dxScene, 0, -1, 0, 0, 0, 1)
            Case 2
                Call dxVue.dxCamera.SetPosition(dxFrameDecor(IndexChoix%(0)), 0, Principale.RegleRecul.Value / REDUCTION%, 0)
                Call dxVue.dxCamera.SetOrientation(dxVue.dxScene, 0, -1, 0, 0, 0, 1)
            Case 3
                Call dxVue.dxCamera.SetPosition(dxFrameWagon(IndexChoix%(0), 0), _
                    ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).PositionCamera.X / REDUCTION%, _
                    ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).PositionCamera.Y / REDUCTION%, _
                    ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).PositionCamera.z / REDUCTION%)
                Call dxVue.dxCamera.SetOrientation(dxFrameWagon(IndexChoix%(0), 0), 0, 0, 1, 0, 1, 0)
                Call dxVue.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, CameraTourne%(1) * DegRad!)
                'Call dxVue.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 0, 0, 1, 0 / 180 * PI!)
                Call dxVue.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 1, 0, 0, (90 - CameraAngle%(1)) * DegRad!)
                Call dxVue.dxCamera.AddTranslation(D3DRMCOMBINE_BEFORE, 0, 0, -Principale.RegleRecul.Value / REDUCTION%)
            End Select
            If IndexElement%(0) = 3 Then
                With dxVue.dxViewport
                    .SetFront 1 / REDUCTION%
                    .SetBack (PosY% + 40) / REDUCTION% * 2
                    .SetField 0.5 / REDUCTION%
                End With
            Else
                With dxVue.dxViewport
                    .SetFront 1
                    If IndexElement%(0) = 0 Then
                        .SetBack (Principale.RegleRecul.Value + 40) / REDUCTION% + HauteurCameraCentre%
                    Else
                        .SetBack (Principale.RegleRecul.Value + 40) / REDUCTION%
                    End If
                    .SetField 0.5
                End With
            End If
        Case EnumVue.VueSurBogie
            '
            ' ***** Caméra dans un élément
            '
            Select Case IndexElement%(0)
            Case 0
                Call dxVue.dxCamera.SetPosition(dxVue.dxScene, (dX% / 2 + xPlateauMin% - BORD%) / REDUCTION%, HauteurCameraCentre% + 100 / REDUCTION%, (dz% / 2 + zPlateauMin% - BORD%) / REDUCTION%)
                Call dxVue.dxCamera.SetOrientation(dxVue.dxScene, 0, 0, 1, 0, 1, 0)
            Case 1
                Call dxVue.dxCamera.SetPosition(dxFrameVoie(0, IndexChoix%(0)), 0, 20 / REDUCTION%, 0)
                Call dxVue.dxCamera.SetOrientation(dxFrameVoie(0, IndexChoix%(0)), 0, 0, 1, 0, 1, 0)
            Case 2
                Call dxVue.dxCamera.SetPosition(dxFrameDecor(IndexChoix%(0)), _
                    ListeDecor(ElementDecor(IndexChoix%(0)).NoDecor%).PositionCamera.X / REDUCTION%, _
                    ListeDecor(ElementDecor(IndexChoix%(0)).NoDecor%).PositionCamera.Y / REDUCTION%, _
                    ListeDecor(ElementDecor(IndexChoix%(0)).NoDecor%).PositionCamera.z / REDUCTION%)
                Call dxVue.dxCamera.SetOrientation(dxFrameDecor(IndexChoix%(0)), 0, 0, 1, 0, 1, 0)
            Case 3
                Call dxVue.dxCamera.SetPosition(dxFrameWagon(IndexChoix%(0), 0), _
                    ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).PositionCamera.X / REDUCTION%, _
                    ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).PositionCamera.Y / REDUCTION%, _
                    ListeWagon(ListeTrain(IndexChoix%(0)).NoWagon%).PositionCamera.z / REDUCTION%)
                Call dxVue.dxCamera.SetOrientation(dxFrameWagon(IndexChoix%(0), 0), 0, 0, 1, 0, 1, 0)
            End Select
            Call dxVue.dxCamera.AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, CameraTourne%(2) * DegRad!)
            With dxVue.dxViewport
                .SetFront 1 / REDUCTION%
                .SetBack 10000! / REDUCTION%
                .SetField 0.5 / REDUCTION%
            End With
        Case EnumVue.VuePoursuite
            '
            ' ***** Caméra dans un élément suit un autre élément
            '
            Select Case IndexElement%(1)
            Case 0
                Call dxVue.dxCamera.SetPosition(dxVue.dxScene, (dX% / 2 + xPlateauMin% - BORD%) / REDUCTION%, HauteurCameraCentre% + 100 / REDUCTION%, (dz% / 2 + zPlateauMin% - BORD%) / REDUCTION%)
                'Call dxVue.dxCamera.SetPosition(dxVue.dxScene, (dX% / 2 + Xmin% - BORD%) / REDUCTION%, 100 / REDUCTION%, (dz% / 2 + zMin% - BORD%) / REDUCTION%)
            Case 1
                Call dxVue.dxCamera.SetPosition(dxFrameVoie(0, IndexChoix%(1)), 0, 20 / REDUCTION%, 0)
            Case 2
                Call dxVue.dxCamera.SetPosition(dxFrameDecor(IndexChoix%(1)), _
                    ListeDecor(ElementDecor(IndexChoix%(1)).NoDecor%).PositionCamera.X / REDUCTION%, _
                    ListeDecor(ElementDecor(IndexChoix%(1)).NoDecor%).PositionCamera.Y / REDUCTION%, _
                    ListeDecor(ElementDecor(IndexChoix%(1)).NoDecor%).PositionCamera.z / REDUCTION%)
            Case 3
                Call dxVue.dxCamera.SetPosition(dxFrameWagon(IndexChoix%(1), 0), _
                    ListeWagon(ListeTrain(IndexChoix%(1)).NoWagon%).PositionCamera.X / REDUCTION%, _
                    ListeWagon(ListeTrain(IndexChoix%(1)).NoWagon%).PositionCamera.Y / REDUCTION%, _
                    ListeWagon(ListeTrain(IndexChoix%(1)).NoWagon%).PositionCamera.z / REDUCTION%)
            End Select
            '
            Select Case IndexElement%(0)
            Case 0
                Call dxVue.dxCamera.LookAt(dxVue.dxCamera, dxVue.dxScene, D3DRMCONSTRAIN_Z)
            Case 1
                Call dxVue.dxCamera.LookAt(dxFrameVoie(0, IndexChoix%(0)), dxVue.dxScene, D3DRMCONSTRAIN_Z)
            Case 2
                Call dxVue.dxCamera.LookAt(dxFrameDecor(IndexChoix%(0)), dxVue.dxScene, D3DRMCONSTRAIN_Z)
            Case 3
                Call dxVue.dxCamera.LookAt(dxFrameWagon(IndexChoix%(0), 0), dxVue.dxScene, D3DRMCONSTRAIN_Z)
            End Select
            '
            With dxVue.dxViewport
                If IndexElement%(1) = 0 Then ' Caméra au centre
                    .SetFront 1
                    .SetBack 10000!
                    .SetField 0.5
                ElseIf IndexElement%(1) = 3 Then ' Caméra dans un wagon
                    .SetFront 1 / REDUCTION%
                    .SetBack 10000! / REDUCTION%
                    .SetField 0.5 / REDUCTION%
                Else
                    .SetFront 10 / REDUCTION%
                    .SetBack 100000! / REDUCTION%
                    .SetField 5 / REDUCTION%
                End If
            End With
        Case Else
            '
            ' ***** Caméra vue de dessus
            '
            Call dxVue.dxCamera.SetPosition(dxVue.dxScene, 0, 0, 0)
            Call dxVue.dxCamera.SetOrientation(dxVue.dxScene, 0, -1, 0, 0, 0, 1)
            Call dxVue.dxCamera.AddTranslation(D3DRMCOMBINE_AFTER, 0, PosY% / REDUCTION%, 0)
            Call dxVue.dxCamera.AddRotation(D3DRMCOMBINE_AFTER, -1, 0, 0, CameraAngle%(0) * DegRad!)
            Call dxVue.dxCamera.AddRotation(D3DRMCOMBINE_AFTER, 0, 1, 0, CameraTourne%(0) * DegRad!)
            Call dxVue.dxCamera.AddTranslation(D3DRMCOMBINE_AFTER, (dX% / 2 + xPlateauMin% - BORD%) / REDUCTION%, 0, (dz% / 2 + zPlateauMin% - BORD%) / REDUCTION%)
            With dxVue.dxViewport
                .SetFront 1
                .SetBack (PosY% + 40) / REDUCTION% * 2
                .SetField 0.5
            End With
        End Select
        Call dxVue.dxViewport.SetProjection(D3DRMPROJECT_PERSPECTIVE)
    End If
End Sub

'
' **************************
' Boucle infinie d'affichage
' **************************
'
Private Sub Boucle()
    Dim l1!
    Dim i%, n%
    ' Temps/Vitesse
    Dim t1&, t2&
    Dim Vitesse!
    Dim TempsPrecedent&
    Dim TempsActuelle&
    '
    Dim Heurtoir As Boolean
    ' Son
    Dim p As D3DVECTOR
    Dim v As D3DVECTOR
    Dim w As D3DVECTOR
    Dim Facteur!
    '
    Tourne = True
    Vue.Caption = Localisation$(CleVUE%)
    Do
        If ModeActuelle = ModeVisualisation Then
            '
            ' ***** Pondère la vitesse en fonction de la rapidité d'affichage
            '
            Vitesse! = Principale.RegleVitesse.Value * TempsAffichage& / 100
            '
            ' ***** Remet à zéro l'indicateur de passage
            '
            For i% = 1 To UBound(Reseau())
                Reseau(i%).Dessus = False
            Next i%
            '
            ' ***** Place la première bogie
            ' ***** Puis les bogies suivantes
            '
            n% = UBound(ListeBogie())
            If Vitesse! < 0 Then
                ListeBogie(n%).Jonction = False
                ListeBogie(n%).BogieSens = -ListeBogie(n%).BogieSens
                Heurtoir = Voies.Deplace_Bogie(ListeBogie(n%), -Vitesse!)
                ListeBogie(n%).BogieSens = -ListeBogie(n%).BogieSens
                '
                For i% = n% - 1 To 0 Step -1
                    ListeBogie(i% + 1).Jonction = False
                    l1! = ListeBogie(i% + 1).BogieDecale!
                    Call Vue.Positionne_Bogie_Precedente(i%, l1!)
                Next i%
            Else
                ListeBogie(0).Jonction = False
                Heurtoir = Voies.Deplace_Bogie(ListeBogie(0), Vitesse!)
                '
                For i% = 1 To n%
                    ListeBogie(i%).Jonction = False
                    l1! = -ListeBogie(i%).BogieDecale!
                    Call Vue.Positionne_Bogie_Suivante(i%, l1!)
                Next i%
            End If
        End If
        t1& = timeGetTime
        Call Vue.Vue_MAJ
        t2& = timeGetTime
        If ModeActuelle = ModeVisualisation Then
            '
            ' ***** réactive la mémoire de la position de l'aiguillage
            '
            For i% = 1 To UBound(Reseau())
                If Reseau(i%).Dessus = False Then
                    Reseau(i%).AiguilleForce% = Reseau(i%).Aiguille%
                End If
            Next i%
            '
            ' ***** Gestion du son
            '
            Call dxVue.dxCamera.GetPosition(dxVue.dxScene, v)
            Call dxVue.dX7.VectorScale(p, v, REDUCTION%)
            Call dxVue.dxCamera.GetOrientation(dxVue.dxScene, v, w)
            If DSound.FoundCard = True Then
                Call DSound.dsListener.SetOrientation(v.X, v.Y, v.z, w.X, w.Y, w.z, DS3D_IMMEDIATE)
            Facteur! = 200 * Abs(Principale.RegleVitesse.Value) / Principale.RegleVitesse.Max
            End If
'Debug.Print "*****"
            For i% = 0 To UBound(ListeBogie())
                If Facteur! <> 0 And Heurtoir = True Then
                    Call dxVue.dX7.VectorSubtract(w, ListeBogie(i%).BogieVecteur, p)
                    Call dxVue.dX7.VectorScale(v, w, 1 / Facteur!)
                    Call DSound.SetPosition(ListeBogie(i%).Son%, v.X, v.Y, v.z)
                    Call DSound.SetVolume(ListeBogie(i%).Son%, 0)
'                    If ListeBogie(i%).Jonction = True Then
'Debug.Print i%; v.x; v.y; v.z
'                        Call DSound.Play3D(2, v.x, v.y, v.z, DSBPLAY_DEFAULT)
'                    End If
                Else
                    Call DSound.SetVolume(ListeBogie(i%).Son%, -10000)
                End If
            Next i%
        Else
            '
            ' ***** Coupe le son
            '
            For i% = 0 To UBound(ListeBogie())
                Call DSound.SetVolume(ListeBogie(i%).Son%, -10000)
            Next i%
        End If
        TempsActuelle& = timeGetTime
        TempsAffichage& = TempsActuelle& - TempsPrecedent&
'Debug.Print "Total:"; TempsAffichage&; t1& - TempsPrecedent&; t2& - t1&; TempsActuelle& - t2&
        TempsPrecedent& = TempsActuelle&
        If TempsAffichage& = 0 Then TempsAffichage& = 1
        DoEvents
    Loop While Tourne = True
End Sub

