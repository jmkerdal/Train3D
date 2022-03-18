Attribute VB_Name = "Initialisation"
Option Explicit
'
Global Const REDUCTION% = 100
Global Super$
'
Global ParamElevation As Boolean
Global ParamHauteur%
Global ParamCatenaire As Boolean
Global ParamSens As Boolean
Global ParamCiel As Boolean
'
Global Localisation$()
Global CleDRIVER%
Global CleEDITION%
Global CleAPROPOS%
Global CleINVENTAIRE%
Global ClePRINCIPALE%
Global CleSAISIETRAIN%
Global CleVUE%
Global CleVOIE%
Global CleRAIL%
'
Enum EnumVue ' Mode de la vue
    VueTable
    VueSurvol
    VueSurBogie
    VuePoursuite
End Enum
Global VueActuelle As EnumVue
'
Global dxVue As New ClassDirectX
Global DSound As New ClassDirectSound
Global Tourne As Boolean
'
Global dxTraverse(4) As Direct3DRMMeshBuilder3
Global dxRoue As Direct3DRMMeshBuilder3
Global dxHeurtoir As Direct3DRMMeshBuilder3
Global dxAiguille As Direct3DRMMeshBuilder3
Global dxSelection(2) As Direct3DRMMeshBuilder3 ' Séléction de couleur
Global dxPointe As Direct3DRMMeshBuilder3 ' Entoure l'objet pointé
Global dxFond As Direct3DRMMeshBuilder3
Global dxAttele As Direct3DRMMeshBuilder3
Global dxCatenaire As Direct3DRMMeshBuilder3
Global dxCiel As Direct3DRMMeshBuilder3
'
Global dxScene2 As Direct3DRMFrame3
Global dxFrameInterieur As Direct3DRMFrame3
Global dxFrameVoie() As Direct3DRMFrame3
Global dxFrameRoue() As Direct3DRMFrame3
Global dxFrameWagon() As Direct3DRMFrame3
Global dxFrameDecor() As Direct3DRMFrame3
Global dxFramePointe(2) As Direct3DRMFrame3
Global dxFrameCiel As Direct3DRMFrame3
Global dxSolTexture(1) As Direct3DRMTexture3
Global DegRad! ' Conversion Degré vers Radian
Global xPlateauMin%, xPlateauMax%
Global zPlateauMin%, zPlateauMax%
Global PosCamera As D3DVECTOR
Global PosSouris As D3DVECTOR
Global OldPosSouris As D3DVECTOR
'
Enum EnumMode
    ModeEdition
    ModeVisualisation
End Enum
Global ModeActuelle As EnumMode
'
Type TypeSelection
    Nom As String
    No As Integer
    Ou As Integer
    Couleur As EnumCouleurSelection
End Type
Enum EnumCouleurSelection
    SelectionSans
    SelectionRouge
    SelectionVert
    SelectionBleu
End Enum
Global Pointe(1) As TypeSelection
Global Origine% ' Voie d'origine
'
Global CheminReseau$ ' Chemin du réseau à sauver
Global CheminRail$ ' Chemin du fichier de rail
Global ReseauFermée As Boolean
Global NomFichier$(6) ' Nom des 4 derniers fichiers + Langue + Driver + Reseau
Global FlagNomFichier As Boolean
Global ReseauElectrique As Boolean
'
Global OldX%, OldY%
Global SourisX%, SourisY%
Global SourisClick%
'
' Position de la caméra
Global CameraRecul%
Global TailleX%, TailleY%
Global CameraAngle%(1)
Global CameraTourne%(2)
'
Global Const BORD% = 80 ' Largeur du bord
'
Type TypeWagon
    Nom As String
    Ref As String
    Fichier As String
    FichierBogie As String
    FichierInterieur As String
    EccartBogie(1) As Single ' Eccart bogie sur un essieu
    EccartEssieu As Single ' Eccart entre les 2 essieux
    Longueur As Single ' Longueur totale
    Motrice As Integer
    Electrique As Integer
    PositionCamera As D3DVECTOR
    PreCharge As Integer
    '
    dxWagon As Direct3DRMMeshBuilder3
    dxBogie As Direct3DRMMeshBuilder3
    dxInterieur As Direct3DRMMeshBuilder3
End Type
Global ListeWagon() As TypeWagon
'
Public Type TypeBogie
    BogieReseau As Integer ' Pointe sur l'indice du réseau
    BogieSegment As Integer ' No du Segment
    BogiePosition As Single ' Position sur le segment
    BogieSens As Integer ' Sens d'évolution
    '
    BogieWagon As Integer ' N° du wagon de la bogie
    BogieNo As Integer ' N° de bogie sur le wagon
    BogieDecale As Single ' Déplacement par rapport à la bogie précédente
    '
    BogieVecteur As D3DVECTOR ' Position vectoriel de la bogie
    Visible As Boolean
    Son As Integer ' N° du son associé
    Jonction As Boolean ' La bogie passe une jonction
End Type
Global ListeBogie() As TypeBogie
'
Public Type TypeTrain
    NoWagon As Integer
    TrainBogie(3) As Integer ' liste des bogies sur le train
    Visible As Boolean
    BogieVisible(1) As Boolean
    AttacheVisible(1) As Boolean
End Type
Global ListeTrain() As TypeTrain
'
Public Type TypeDecor
    Nom As String
    Ref As String
    Fichier As String
    PreCharge As Integer
    PositionCamera As D3DVECTOR
    dxDecor As Direct3DRMMeshBuilder3
    '
    Inventaire As Integer
End Type
Global ListeDecor() As TypeDecor
'
Public Type TypeElementDecor
    NoDecor As Integer
    Position As D3DVECTOR
    Angle As Integer
    '
    Visible As Boolean
End Type
Global ElementDecor() As TypeElementDecor
'
Global TypeCopie% ' Type d'élément à copier
Global NoCopie% ' N° de voie/décor à copier
Global AngleCopie! ' Angle de la copie
Global TempsAffichage& ' Mémorise le temps d'affichage
Global IndexChoixElement%() ' Index de l'élément choisi
Global IndexElement%(1), IndexChoix%(1)
Global AfficheFps As Boolean
'
Global CompteLancement% ' compteur de lancement de vue
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'
' *******************************
' Test si un OctObjet est visible
' *******************************
'
'Public Function Test_Visible(Frame As Direct3DRMFrame3, Mesh As Direct3DRMMeshBuilder3) As Boolean
'Test_Visible = True
'Exit Function
    'Dim dxBox As D3DRMBOX ' Boite du mesh
    'Dim VueBox As D3DRMBOX ' Coordonnées projetés
    'Dim OctObjet As D3DVECTOR
    'Dim OctObjet4D As D3DRMVECTOR4D
    'Dim Tv As D3DVECTOR
    'Dim i%
    '
    ' ***** Transforme un D3DRMBOX en OctObjet
    '
    'Call Mesh.GetBox(dxBox)
    'VueBox.Max.X = -1: VueBox.Max.Y = -1
    'VueBox.Min.X = TailleX% + 1: VueBox.Min.Y = TailleY% + 1
    'For i% = 0 To 7
    '    If i% = 0 Then
    '        OctObjet.X = dxBox.Min.X: OctObjet.Y = dxBox.Max.Y: OctObjet.z = dxBox.Max.z
    '    ElseIf i% = 1 Then
    '        OctObjet.X = dxBox.Max.X: OctObjet.Y = dxBox.Max.Y: OctObjet.z = dxBox.Max.z
    '    ElseIf i% = 2 Then
    '        OctObjet.X = dxBox.Max.X: OctObjet.Y = dxBox.Max.Y: OctObjet.z = dxBox.Min.z
    '    ElseIf i% = 3 Then
    '        OctObjet.X = dxBox.Min.X: OctObjet.Y = dxBox.Max.Y: OctObjet.z = dxBox.Min.z
    '    ElseIf i% = 4 Then
    '        OctObjet.X = dxBox.Min.X: OctObjet.Y = dxBox.Min.Y: OctObjet.z = dxBox.Max.z
    '    ElseIf i% = 5 Then
    '        OctObjet.X = dxBox.Max.X: OctObjet.Y = dxBox.Min.Y: OctObjet.z = dxBox.Max.z
    '    ElseIf i% = 6 Then
    '        OctObjet.X = dxBox.Max.X: OctObjet.Y = dxBox.Min.Y: OctObjet.z = dxBox.Min.z
    '    Else
    '        OctObjet.X = dxBox.Min.X: OctObjet.Y = dxBox.Min.Y: OctObjet.z = dxBox.Min.z
    '    End If
    '    '
    '    ' ***** Applique les transformations
    '    '
    '    Call Frame.Transform(Tv, OctObjet)
    '    Call dxVue.dxViewport.Transform(OctObjet4D, Tv)
    '    If OctObjet4D.w > 0 Then
    '        OctObjet4D.X = OctObjet4D.X / OctObjet4D.w
    '        OctObjet4D.Y = OctObjet4D.Y / OctObjet4D.w
    '        If OctObjet4D.X > VueBox.Max.X Then VueBox.Max.X = OctObjet4D.X
    '        If OctObjet4D.X < VueBox.Min.X Then VueBox.Min.X = OctObjet4D.X
    '        If OctObjet4D.Y > VueBox.Max.Y Then VueBox.Max.Y = OctObjet4D.Y
    '        If OctObjet4D.Y < VueBox.Min.Y Then VueBox.Min.Y = OctObjet4D.Y
    '    End If
    'Next i%
    'If VueBox.Max.X >= 0 Then
    '    If VueBox.Min.X <= TailleX% Then
    '        If VueBox.Max.Y >= 0 Then
    '            If VueBox.Min.Y <= TailleY% Then
    '                Test_Visible = True
    '            End If
    '        End If
    '    End If
    'End If
'End Function

'
' ***********************************
' Décharge la vue des voies et trains
' ***********************************
'
Public Sub Vue_Decharge()
    Dim i%, j%
    Dim ta&, tb&
    Dim P%, Trouve As Boolean
'Debug.Print "Decharge"
    ta& = timeGetTime
    Call dxFrameCiel.DeleteVisual(dxCiel)
    Call dxFrameVoie(0, 0).DeleteVisual(dxFond)
    '
    For i% = 1 To UBound(Reseau())
        If Reseau(i%).NoVoie% <> 0 Then
            If ModeActuelle = ModeVisualisation Then
            'If VueActuelle < VueSurBogie Then
                Call dxFrameVoie(0, i%).DeleteVisual(Voie(Reseau(i%).NoVoie%).VoieMeshBuilder)
                'Call dxFrameVoie(1, i%).DeleteVisual(dxAiguille)
                '
                Trouve = False
                For j% = 0 To NbSegment%
                    ' ***** Pose l'indicateur d'aiguillage
                    If Trouve = False And Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 0) <> 0 And Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 1) <> 0 Then
                        P% = Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 1) - 1
                        If Voie(Reseau(i%).NoVoie%).MatAiguille(Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 0) - 1, P%) = Reseau(i%).AiguilleForce% Then
                            Call dxFrameVoie(1, i%).DeleteVisual(dxAiguille)
                            Call dxScene2.DeleteChild(dxFrameVoie(1, i%))
                            Trouve = True
                        End If
                    End If
                Next j%
                '
                If ReseauElectrique = True Then
                    Call dxScene2.DeleteChild(dxFrameVoie(2, i%))
                    Call dxFrameVoie(2, i%).DeleteVisual(Voie(Reseau(i%).NoVoie%).VoieCatenaireMeshBuilder)
                    For j% = 0 To NbSegment%
                        'If Voie(Reseau(i%).NoVoie%).CatenaireSegment%(j%) <> 0 And Reseau(i%).MetCatenaire(j%) = True Then
                        If Reseau(i%).MetCatenaire(j%) = True Then
                            Call dxScene2.DeleteChild(dxFrameVoie(j% + 3, i%))
                            Call dxFrameVoie(j% + 3, i%).DeleteVisual(dxCatenaire)
                        End If
                    Next j%
                End If
            Else
                Call dxFrameVoie(0, i%).DeleteVisual(Voie(Reseau(i%).NoVoie%).VoieLightMeshBuilder)
            End If
            '
        End If
    Next i%
    '
    For i% = 1 To UBound(ElementDecor())
        If ElementDecor(i%).NoDecor% <> 0 Then
            Call dxFrameDecor(i%).DeleteVisual(ListeDecor(ElementDecor(i%).NoDecor%).dxDecor)
        End If
    Next i%
    '
    If ModeActuelle = ModeVisualisation Then
        For i% = 1 To UBound(ListeBogie()) - 1
            Call dxVue.dxScene.DeleteChild(dxFrameRoue(i%))
            Call dxFrameRoue(i%).DeleteVisual(dxRoue)
        Next i%
        '
        For i% = 1 To UBound(ListeTrain())
            Call dxVue.dxScene.DeleteChild(dxFrameWagon(i%, 0))
            Call dxFrameWagon(i%, 0).DeleteVisual(ListeWagon(ListeTrain(i%).NoWagon%).dxWagon)
            Call dxFrameWagon(i%, 1).DeleteVisual(dxAttele)
            Call dxFrameWagon(i%, 2).DeleteVisual(dxAttele)
            For j% = 0 To 1
                Call dxVue.dxScene.DeleteChild(dxFrameWagon(i%, j% + 1))
                If ListeTrain(i%).TrainBogie(j% * 2 + 1) <> -1 Then
                    If ListeWagon(ListeTrain(i%).NoWagon%).FichierBogie$ <> "" Then
                        Call dxVue.dxScene.DeleteChild(dxFrameWagon(i%, j% + 3))
                        Call dxFrameWagon(i%, j% + 3).DeleteVisual(ListeWagon(ListeTrain(i%).NoWagon%).dxBogie)
                    End If
                End If
            Next j%
        Next i%
        '
        If ParamCiel = True Then
            Call dxScene2.DeleteChild(dxFrameCiel)
        End If
        '
        For i% = 0 To UBound(ListeTunnel()) - 1
            Call dxScene2.DeleteChild(FrameTunnel(i%))
        Next i%
    End If
    tb& = timeGetTime
'Debug.Print tb& - ta&
End Sub

'
' ********************************
' Création de la vue voie et train
' ********************************
'
Public Sub Vue_Charge()
    Dim i%, j%, X%, Y%, z%
    Dim ta&, tb&
    Dim Index%
    Dim PosCatenaire As TypeBogie
    Dim DirCatenaire!
    Dim P%, Trouve As Boolean
    '
'Debug.Print "Charge"
    ta& = timeGetTime
    Call dxFrameCiel.AddVisual(dxCiel)
    '
    ' ***** Attache les wagons
    '
    If ModeActuelle = ModeVisualisation Then
        '
        ' ***** Pose les bogies
        ' ***** sauf la première et la dernière
        '
        For i% = 1 To UBound(ListeBogie()) - 1
            Call dxFrameRoue(i%).AddVisual(dxRoue)
            Call dxVue.dxScene.AddChild(dxFrameRoue(i%))
        Next i%
        '
        ' ***** Affiche les wagons
        '
        For i% = 1 To UBound(ListeTrain())
            Call dxFrameWagon(i%, 0).AddVisual(ListeWagon(ListeTrain(i%).NoWagon%).dxWagon)
            Call dxFrameWagon(i%, 1).AddVisual(dxAttele)
            Call dxFrameWagon(i%, 2).AddVisual(dxAttele)
            Call dxVue.dxScene.AddChild(dxFrameWagon(i%, 0))
            For j% = 0 To 1
                Call dxVue.dxScene.AddChild(dxFrameWagon(i%, j% + 1))
                If ListeTrain(i%).TrainBogie(j% * 2 + 1) <> -1 Then
                    If ListeWagon(ListeTrain(i%).NoWagon%).FichierBogie$ <> "" Then
                        Call dxFrameWagon(i%, j% + 3).AddVisual(ListeWagon(ListeTrain(i%).NoWagon%).dxBogie)
                        Call dxVue.dxScene.AddChild(dxFrameWagon(i%, j% + 3))
                    End If
                End If
            Next j%
        Next i%
        '
        ' ***** Attache le ciel
        '
        If ParamCiel = True Then
            Call dxScene2.AddChild(dxFrameCiel)
        End If
        '
        ' ***** Attache les tunnels
        '
        For i% = 0 To UBound(ListeTunnel()) - 1
            Call dxScene2.AddChild(FrameTunnel(i%))
        Next i%
    End If
    '
    ' ***** Attache le décor
    '
    For i% = 1 To UBound(ElementDecor())
        If ElementDecor(i%).NoDecor% <> 0 Then
            Call dxFrameDecor(i%).AddVisual(ListeDecor(ElementDecor(i%).NoDecor%).dxDecor)
            X% = ElementDecor(i%).Position.X
            X% = (X% \ 10) * 10
            Y% = ElementDecor(i%).Position.Y
            Y% = (Y% \ 10) * 10
            z% = ElementDecor(i%).Position.z
            z% = (z% \ 10) * 10
            Call dxFrameDecor(i%).SetPosition(dxVue.dxScene, X% / REDUCTION%, Y% / REDUCTION%, z% / REDUCTION%)
            Call dxFrameDecor(i%).SetOrientation(dxVue.dxScene, 0, 0, 1, 0, 1, 0)
            Call dxFrameDecor(i%).AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, ElementDecor(i%).Angle% * DegRad!)
        End If
    Next i%
    '
    ' ***** Attache le réseau
    '
    For i% = 1 To UBound(Reseau())
        If Reseau(i%).NoVoie% <> 0 Then
            Call dxFrameVoie(0, i%).SetPosition(dxVue.dxScene, Reseau(i%).Position.X / REDUCTION%, Reseau(i%).Position.Y / REDUCTION%, Reseau(i%).Position.z / REDUCTION%)
            Call dxFrameVoie(0, i%).SetOrientation(dxVue.dxScene, 0, 0, 1, 0, 1, 0)
            Call dxFrameVoie(0, i%).AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, Reseau(i%).Angle! * DegRad!)
            If ModeActuelle = ModeVisualisation Then
                Call dxFrameVoie(0, i%).AddVisual(Voie(Reseau(i%).NoVoie%).VoieMeshBuilder)
                Trouve = False
                For j% = 0 To NbSegment%
                    ' ***** Pose l'indicateur d'aiguillage
                    If Trouve = False And Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 0) <> 0 And Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 1) <> 0 Then
                        P% = Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 1) - 1
                        If Voie(Reseau(i%).NoVoie%).MatAiguille(Voie(Reseau(i%).NoVoie%).SegmentPoint(j%, 0) - 1, P%) = Reseau(i%).AiguilleForce% Then
                            Call dxFrameVoie(1, i%).AddVisual(dxAiguille)
                            Call dxScene2.AddChild(dxFrameVoie(1, i%))
                            Trouve = True
                        End If
                    End If
                    '
                    If Voie(Reseau(i%).NoVoie%).CatenaireSegment%(j%) <> 0 Then
                        Call dxFrameVoie(j% + 3, i%).SetPosition(dxFrameVoie(0, i%), 0, 0, 0)
                        PosCatenaire.BogieReseau% = i%
                        PosCatenaire.BogieSegment% = Voie(Reseau(i%).NoVoie%).CatenaireSegment%(j%)
                        PosCatenaire.BogiePosition! = Voie(Reseau(i%).NoVoie%).CatenairePosition!(j%)
                        Call Voies.Position_Bogie(PosCatenaire, PosCatenaire.BogieVecteur, DirCatenaire!)
                        Call dxFrameVoie(j% + 3, i%).SetPosition(dxVue.dxScene, PosCatenaire.BogieVecteur.X / REDUCTION%, 0, PosCatenaire.BogieVecteur.z / REDUCTION%)
                        Call dxFrameVoie(j% + 3, i%).SetOrientation(dxVue.dxScene, 0, 0, 1, 0, 1, 0)
                        Call dxFrameVoie(j% + 3, i%).AddRotation(D3DRMCOMBINE_BEFORE, 0, 1, 0, (DirCatenaire! + Voie(Reseau(i%).NoVoie%).CatenaireSens%(j%)) * DegRad!)
                    End If
                    '
                Next j%
            Else
                Call dxFrameVoie(0, i%).AddVisual(Voie(Reseau(i%).NoVoie%).VoieLightMeshBuilder)
            End If
            '
        End If
    Next i%
    '
    Set dxFond = Nothing
    Call Sol.Cree_Sol
    Call dxFrameVoie(0, 0).AddVisual(dxFond)
    '
    ' ***** Attache les caténaires
    '
    If ModeActuelle = ModeVisualisation And ReseauElectrique = True Then
        For i% = 1 To UBound(Reseau())
            If Reseau(i%).NoVoie% <> 0 Then
                Call dxFrameVoie(2, i%).AddVisual(Voie(Reseau(i%).NoVoie%).VoieCatenaireMeshBuilder)
                Call dxFrameVoie(2, i%).SetPosition(dxFrameVoie(0, i%), 0, 0, 0)
                Call dxFrameVoie(2, i%).SetOrientation(dxFrameVoie(0, i%), 0, 0, 1, 0, 1, 0)
                Call dxScene2.AddChild(dxFrameVoie(2, i%))
                '
                For j% = 0 To NbSegment%
                    If Reseau(i%).MetCatenaire(j%) = True Then
                        Call dxFrameVoie(j% + 3, i%).AddVisual(dxCatenaire)
                        Call dxScene2.AddChild(dxFrameVoie(j% + 3, i%))
                    End If
                Next j%
            End If
        Next i%
    End If
    '
    tb& = timeGetTime
'Debug.Print tb& - ta&
End Sub

'
' ***************************************************
' Ajoute une nouvelle voie sur un point de connection
' ***************************************************
'
Public Sub Voie_Ajoute(Voie%, EntreeVoie%, Precedent%, EntreePrecedent%)
    Reseau(Voie%).Connecte(EntreeVoie%) = Precedent%
    Reseau(Voie%).Entree(EntreeVoie%) = EntreePrecedent%
    Reseau(Precedent%).Connecte(EntreePrecedent%) = Voie%
    Reseau(Precedent%).Entree(EntreePrecedent%) = EntreeVoie%
End Sub

'
' ******************************************
' Recalcule les points associés aux segments
' ******************************************
'
Public Sub Recalcule_Point(No%)
    Dim i%, j%, n%
    Voie(No%).AiguillePosition% = 0
    For i% = 0 To NbSegment%
        Voie(No%).pX!(i%) = 0 ' Remise à zéro des points
        Voie(No%).pZ!(i%) = 0
        For j% = 0 To NbSegment%
            If i% <> j% Then
                n% = Voie(No%).MatConnecte(i%, j%)
                If n% <> 0 Then
                    Voie(No%).SegmentPoint(n% - 1, 0) = j% + 1
                    Voie(No%).SegmentPoint(n% - 1, 1) = i% + 1
                End If
                n% = Voie(No%).MatAiguille(i%, j%)
                If n% > Voie(No%).AiguillePosition% Then
                    Voie(No%).AiguillePosition% = n%
                End If
            End If
        Next j%
    Next i%
End Sub

'
' ***************************************
' Supprime la connection entre deux voies
' ***************************************
'
Public Sub Voie_Supprime(Voie%, EntreeVoie%, Precedent%, EntreePrecedent%)
    Reseau(Precedent%).Connecte(EntreePrecedent%) = 0
    Reseau(Precedent%).Entree(EntreePrecedent%) = 0
    Reseau(Voie%).Connecte(EntreeVoie%) = 0
    Reseau(Voie%).Entree(EntreeVoie%) = 0
End Sub

'
' **************************
' Charge la liste des wagons
' **************************
'
Public Sub Wagon_Charger()
    Dim f%, n%, i%
    f% = FreeFile()
    Call Tools.Open_File(".\Petit-train 3D\Base.wag", f%, OPEN_NORMAL)
    Input #f%, n%
    ReDim ListeWagon(n%) As TypeWagon
    If n% <> 0 Then
        For i% = 1 To n%
            DoEvents
            Input #f%, ListeWagon(i%).Nom$
            Input #f%, ListeWagon(i%).Ref$
            Input #f%, ListeWagon(i%).Fichier$
            Input #f%, ListeWagon(i%).FichierBogie$
            Input #f%, ListeWagon(i%).FichierInterieur$
            Input #f%, ListeWagon(i%).PreCharge%
            Input #f%, ListeWagon(i%).EccartBogie!(0)
            Input #f%, ListeWagon(i%).EccartBogie!(1)
            Input #f%, ListeWagon(i%).EccartEssieu!
            Input #f%, ListeWagon(i%).Longueur!
            Input #f%, ListeWagon(i%).Motrice%
            Input #f%, ListeWagon(i%).Electrique%
            Input #f%, ListeWagon(i%).PositionCamera.X
            Input #f%, ListeWagon(i%).PositionCamera.Y
            Input #f%, ListeWagon(i%).PositionCamera.z
            If ListeWagon(i%).PreCharge% = 1 Then
                Call Wagon_Charger_Mesh(i%)
            End If
        Next i%
    End If
    Close #f%
End Sub

'
' *************************
' Sauve la liste des wagons
' *************************
'
Public Sub Wagon_Sauver()
    Dim f%, n%, i%
    f% = FreeFile()
    Open ".\Petit-train 3D\Base.wag" For Output As #f%
    n% = UBound(ListeWagon())
    Print #f%, n%
    If n% <> 0 Then
        For i% = 1 To n%
            Print #f%, ListeWagon(i%).Nom$
            Print #f%, ListeWagon(i%).Ref$
            Print #f%, ListeWagon(i%).Fichier$
            Print #f%, ListeWagon(i%).FichierBogie$
            Print #f%, ListeWagon(i%).FichierInterieur$
            Print #f%, ListeWagon(i%).PreCharge%;
            Print #f%, ListeWagon(i%).EccartBogie!(0);
            Print #f%, ListeWagon(i%).EccartBogie!(1);
            Print #f%, ListeWagon(i%).EccartEssieu!;
            Print #f%, ListeWagon(i%).Longueur!;
            Print #f%, ListeWagon(i%).Motrice%;
            Print #f%, ListeWagon(i%).Electrique%;
            Print #f%, ListeWagon(i%).PositionCamera.X;
            Print #f%, ListeWagon(i%).PositionCamera.Y;
            Print #f%, ListeWagon(i%).PositionCamera.z
        Next i%
    End If
    Close #f%
End Sub

'
' ************************
' Charge la liste du décor
' ************************
'
Public Sub Décor_Charger()
    Dim f%, n%, i%
    f% = FreeFile()
    Call Tools.Open_File(".\Petit-train 3D\Base.deco", f%, OPEN_NORMAL)
    Input #f%, n%
    ReDim ListeDecor(n%) As TypeDecor
    Set ListeDecor(0).dxDecor = dxVue.dxD3Drm.CreateMeshBuilder
    If n% <> 0 Then
        For i% = 1 To n%
            DoEvents
            Input #f%, ListeDecor(i%).Nom$
            Input #f%, ListeDecor(i%).Ref$
            Input #f%, ListeDecor(i%).Fichier$
            Input #f%, ListeDecor(i%).PreCharge%
            Input #f%, ListeDecor(i%).PositionCamera.X
            Input #f%, ListeDecor(i%).PositionCamera.Y
            Input #f%, ListeDecor(i%).PositionCamera.z
            If ListeDecor(i%).PreCharge% = 1 Then
                Call Décor_Charge_Mesh(i%)
            End If
        Next i%
    End If
    Close #f%
End Sub

'
' ***********************
' Sauve la liste du décor
' ***********************
'
Public Sub Décor_Sauver()
    Dim f%, n%, i%
    f% = FreeFile()
    Open ".\Petit-train 3D\Base.deco" For Output As #f%
    n% = UBound(ListeDecor())
    Print #f%, n%
    If n% <> 0 Then
        For i% = 1 To n%
            Print #f%, ListeDecor(i%).Nom$
            Print #f%, ListeDecor(i%).Ref$
            Print #f%, ListeDecor(i%).Fichier$
            Print #f%, ListeDecor(i%).PreCharge%
            Print #f%, ListeDecor(i%).PositionCamera.X;
            Print #f%, ListeDecor(i%).PositionCamera.Y;
            Print #f%, ListeDecor(i%).PositionCamera.z
        Next i%
    End If
    Close #f%
End Sub

'
' *************************
' Ajoute un décor au réseau
' *************************
'
Public Sub Décor_Ajoute(No_Decor%, Angle_Decor!)
    Dim i%, t%, n%
    Call Initialisation.Vue_Decharge
    n% = UBound(ElementDecor())
    For i% = 1 To n%
        If ElementDecor(i%).NoDecor% = 0 Then
            t% = i%
        End If
    Next i%
    If t% = 0 Then
        t% = n% + 1
        ReDim Preserve ElementDecor(t%) As TypeElementDecor
        ReDim Preserve dxFrameDecor(t%) As Direct3DRMFrame3
        Set dxFrameDecor(t%) = dxVue.dxD3Drm.CreateFrame(dxVue.dxScene)
        Call dxFrameDecor(t%).SetAppData(t%)
        Call dxFrameDecor(t%).SetName("Decor")
    End If
    ElementDecor(t%).NoDecor% = No_Decor%
    ElementDecor(t%).Position.X = PosSouris.X * REDUCTION%
    ElementDecor(t%).Position.z = PosSouris.z * REDUCTION%
    ElementDecor(t%).Angle% = Angle_Decor!
    TypeCopie% = 2
    NoCopie% = No_Decor%
    AngleCopie! = Angle_Decor!
    Call Vue.Calcule_Reseau
    Call Initialisation.Vue_Charge
    Principale.MENU_Coller.Enabled = True
    Call Initialisation.Liste_Charge
End Sub

'
' ********************
' Charge les objets DX
' met à jour les index
' ********************
'
Public Sub ObjetDX_Cree()
    Dim i%, j%, n%
'Debug.Print "ObjetDX_Cree"
    Set dxScene2 = dxVue.dxD3Drm.CreateFrame(Nothing)
    Set dxFrameCiel = dxVue.dxD3Drm.CreateFrame(Nothing)
    Set dxFrameInterieur = dxVue.dxD3Drm.CreateFrame(Nothing)
    '
    ' ***** Système de pointage
    '
    For i% = 0 To 2
        Set dxFramePointe(i%) = dxVue.dxD3Drm.CreateFrame(Nothing)
    Next i%
    '
    ' ***** Calcul le train
    '
    n% = 0
    For i% = 1 To UBound(ListeTrain())
        For j% = 0 To 1
            n% = n% + 1
            ListeTrain(i%).TrainBogie%(j% * 2) = n%
            If ListeWagon(ListeTrain(i%).NoWagon%).EccartBogie!(j%) = 0 Then
                ListeTrain(i%).TrainBogie%(j% * 2 + 1) = -1
            Else
                n% = n% + 1
                ListeTrain(i%).TrainBogie%(j% * 2 + 1) = n%
            End If
        Next j%
    Next i%
    n% = n% + 1
    '
    ' ***** Crée les bogies
    '
    ReDim ListeBogie(n%) As TypeBogie
    ReDim dxFrameRoue(UBound(ListeBogie()))
    For i% = 0 To UBound(ListeBogie())
        Set dxFrameRoue(i%) = dxVue.dxD3Drm.CreateFrame(Nothing)
        Call dxFrameRoue(i%).SetAppData(i%)
        Call dxFrameRoue(i%).SetName("Bogie")
    Next i%
    '
    ' ***** Recherche le lien entre bogie et wagon
    '
    Dim l1!, l2!
    Dim w1%, w2%
    'ListeBogie(0).BogieDecale! = ListeTrain(1).NoWagon
    For i% = 1 To UBound(ListeTrain())
        For j% = 0 To 3
            If ListeTrain(i%).TrainBogie(j%) <> -1 Then
                ListeBogie(ListeTrain(i%).TrainBogie(j%)).BogieWagon% = i%
                ListeBogie(ListeTrain(i%).TrainBogie(j%)).BogieNo% = j%
                w1% = ListeTrain(i% - 1).NoWagon%
                w2% = ListeTrain(i%).NoWagon%
                If j% = 0 Then
                    l2! = ListeWagon(w2%).Longueur! - ListeWagon(w2%).EccartEssieu! - ListeWagon(w2%).EccartBogie(0) / 2 - ListeWagon(w2%).EccartBogie(1) / 2
                    If i% = 1 Then
                        ' Bogie virtuelle
                        ListeBogie(ListeTrain(i%).TrainBogie(j%)).BogieDecale! = l2! / 2
                    Else
                        l1! = ListeWagon(w1%).Longueur! - ListeWagon(w1%).EccartEssieu! - ListeWagon(w1%).EccartBogie(0) / 2 - ListeWagon(w1%).EccartBogie(1) / 2
                        ListeBogie(ListeTrain(i%).TrainBogie(j%)).BogieDecale! = (l1! + l2!) / 2
                    End If
                End If
                If j% = 1 Then
                    ListeBogie(ListeTrain(i%).TrainBogie(j%)).BogieDecale! = ListeWagon(w2%).EccartBogie!(0)
                End If
                If j% = 2 Then
                    ListeBogie(ListeTrain(i%).TrainBogie(j%)).BogieDecale! = ListeWagon(w2%).EccartEssieu! - ListeWagon(w2%).EccartBogie!(0)
                End If
                If j% = 3 Then
                    ListeBogie(ListeTrain(i%).TrainBogie(j%)).BogieDecale! = ListeWagon(w2%).EccartBogie!(1)
                End If
                If i% = UBound(ListeTrain()) Then
                    ListeBogie(n%).BogieDecale! = l2! / 2
                End If
            End If
        Next j%
    Next i%
    '
    ' ***** Crée les wagons
    '
    ReDim dxFrameWagon(UBound(ListeTrain()), 4)
    For i% = 1 To UBound(dxFrameWagon())
        For j% = 0 To 4
            If j% = 0 Then
                Set dxFrameWagon(i%, j%) = dxVue.dxD3Drm.CreateFrame(dxVue.dxScene)
            Else
                Set dxFrameWagon(i%, j%) = dxVue.dxD3Drm.CreateFrame(Nothing)
            End If
        Next j%
        Call dxFrameWagon(i%, 0).SetAppData(i%)
        Call dxFrameWagon(i%, 0).SetName("Wagon")
        Call dxFrameWagon(i%, 1).SetAppData(i%)
        Call dxFrameWagon(i%, 1).SetName("Attele A")
        Call dxFrameWagon(i%, 2).SetAppData(i%)
        Call dxFrameWagon(i%, 2).SetName("Attele B")
        Call dxFrameWagon(i%, 3).SetAppData(i%)
        Call dxFrameWagon(i%, 3).SetName("Bogie A")
        Call dxFrameWagon(i%, 4).SetAppData(i%)
        Call dxFrameWagon(i%, 4).SetName("Bogie B")
    Next i%
    '
    ' ***** Mis en place du décor
    '
    n% = UBound(ElementDecor())
    ReDim dxFrameDecor(n%) As Direct3DRMFrame3
    For i% = 1 To n%
        Set dxFrameDecor(i%) = dxVue.dxD3Drm.CreateFrame(dxVue.dxScene)
        'Set dxFrameDecor(i%) = dxVue.dxD3Drm.CreateFrame(Nothing)
        Call dxFrameDecor(i%).SetAppData(i%)
        Call dxFrameDecor(i%).SetName("Decor")
    Next i%
    '
    ' ***** Création des voies
    '
    ReDim dxFrameVoie(3 + NbSegment%, UBound(Reseau())) As Direct3DRMFrame3
    For i% = 0 To UBound(Reseau())
        Set dxFrameVoie(0, i%) = dxVue.dxD3Drm.CreateFrame(dxVue.dxScene)
        For j% = 1 To 3 + NbSegment%
            Set dxFrameVoie(j%, i%) = dxVue.dxD3Drm.CreateFrame(Nothing)
        Next j%
        If i% = 0 Then
            'For j% = 0 To 3 + NbSegment%
            '    Set dxFrameVoie(j%, i%) = dxVue.dxD3Drm.CreateFrame(dxVue.dxScene)
            'Next j%
            Call dxFrameVoie(0, i%).SetName("Fond")
            Call dxFrameVoie(0, i%).SetOrientation(dxVue.dxScene, 0, -1, 0, 0, 0, 1)
        Else
            'Set dxFrameVoie(0, i%) = dxVue.dxD3Drm.CreateFrame(dxVue.dxScene)
            'For j% = 1 To 3 + NbSegment%
            '    Set dxFrameVoie(j%, i%) = dxVue.dxD3Drm.CreateFrame(Nothing)
            'Next j%
            Call dxFrameVoie(0, i%).SetAppData(i%)
            Call dxFrameVoie(0, i%).SetName("Voie")
        End If
    Next i%
    '
    ' ***** Lance le son
    '
    For i% = 0 To UBound(ListeBogie())
        ListeBogie(i%).Son% = DSound.Play3D%(1, 0, 0, 0, DSBPLAY_LOOPING)
        Call DSound.SetVolume(ListeBogie(i%).Son%, -10000) ' Volume à zéro
    Next i%
    Call Vue.Calcule_Reseau
    Call Initialisation.Liste_Charge
    Call Initialisation.Vue_Charge
End Sub

'
' *************************
' Supprime les frames crées
' *************************
'
Public Sub ObjetDX_Detruit()
    Dim i%, j%
    Call Initialisation.Vue_Decharge
'Debug.Print "ObjetDX_Detruit"
    For i% = 0 To UBound(dxFrameVoie(), 2)
        For j% = 0 To 3 + NbSegment%
            Set dxFrameVoie(j%, i%) = Nothing
        Next j%
    Next i%
    For i% = 1 To UBound(dxFrameDecor())
        Set dxFrameDecor(i%) = Nothing
    Next i%
    For i% = 0 To UBound(dxFrameRoue())
        Set dxFrameRoue(i%) = Nothing
    Next i%
    For i% = 1 To UBound(dxFrameWagon())
        For j% = 0 To 4
            Set dxFrameWagon(i%, j%) = Nothing
        Next j%
    Next i%
    For i% = 0 To 2
        Set dxFramePointe(i%) = Nothing
    Next i%
    Set dxFrameInterieur = Nothing
    Set dxFrameCiel = Nothing
    Set dxScene2 = Nothing
    '
    ' ***** Arrête le son
    '
    For i% = 0 To UBound(ListeBogie())
        Call DSound.StopPlaying(ListeBogie(i%).Son%)
    Next i%
End Sub

'
' **************************
' Remise à zéro des tableaux
' **************************
'
Public Sub Raz()
    ReDim Reseau(0) As TypeReseau
    ReDim ElementDecor(0) As TypeElementDecor
    ReDim ListeTrain(0) As TypeTrain
    ReDim ListeTunnel(0) As New ClassTunnel
    SelectionTunnel% = -1
    TempsAffichage& = 1
End Sub

'
' ****************************
' Charge le nom des 4 fichiers
' ****************************
'
Public Sub NomFichier_Charge()
    Dim i%, f%, a%
    If Tools.Exist(".\Petit-train 3D\Train.ini") = True Then
        f% = FreeFile()
        Call Tools.Open_File(".\Petit-train 3D\Train.ini", f%, OPEN_NORMAL)
        Input #f%, a%
        If a% = 0 Then SansTextureDynamique = False Else SansTextureDynamique = True
        For i% = 0 To 6
            Line Input #f%, NomFichier$(i%)
        Next i%
        Close #f%
    Else
        SansTextureDynamique = False
        For i% = 0 To 6
            NomFichier$(i%) = ""
        Next i%
    End If
    Call Principale.MAJ_MenuNomFichier
End Sub

'
' ***************************
' Sauve le nom des 4 fichiers
' ***************************
'
Public Sub NomFichier_Sauve()
    Dim i%, f%
    f% = FreeFile()
    Open ".\Petit-train 3D\Train.ini" For Output As #f%
    If SansTextureDynamique = False Then Print #f%, 0 Else Print #f%, 1
    For i% = 0 To 6
        Print #f%, NomFichier$(i%)
    Next i%
    Close #f%
End Sub

'
' *************************************
' Chargement du fichier de localisation
' *************************************
'
Public Sub Local_Charge(Fichier$)
    Dim f%, n%
    Dim Ligne$, Cle$
    '
    f% = FreeFile()
    Call Tools.Open_File(Fichier$, f%, OPEN_NORMAL)
    ReDim Localisation$(0)
    ReDim CleLocal$(0)
    ReDim IndiceCleLocal%(0)
    n% = n% + 1
    While Not EOF(f%)
        Line Input #f%, Ligne$
        If Left$(Ligne$, 1) = "[" Then
            Cle$ = Mid$(Ligne$, 2, Len(Ligne$) - 2)
            If Cle$ = "DRIVER" Then CleDRIVER% = n%
            If Cle$ = "EDITION" Then CleEDITION% = n%
            If Cle$ = "A PROPOS" Then CleAPROPOS% = n%
            If Cle$ = "INVENTAIRE" Then CleINVENTAIRE% = n%
            If Cle$ = "PRINCIPALE" Then ClePRINCIPALE% = n%
            If Cle$ = "SAISIE TRAIN" Then CleSAISIETRAIN% = n%
            If Cle$ = "VUE" Then CleVUE% = n%
            If Cle$ = "VOIE" Then CleVOIE% = n%
            If Cle$ = "RAIL" Then CleRAIL% = n%
        Else
            ReDim Preserve Localisation$(n%)
            Localisation(n%) = Ligne$
            n% = n% + 1
        End If
    Wend
    Close #f%
    '
    Principale.MENU_Fichier.Caption = Localisation$(1)
    Principale.MENU_Nouveau.Caption = Localisation$(2)
    Principale.MENU_Charger.Caption = Localisation$(3)
    Principale.MENU_Enregistrer.Caption = Localisation$(4)
    'Principale.MENU_Visualisation.Caption = Localisation$(5)
    Principale.MENU_Imprimantes.Caption = Localisation$(6)
    Principale.MENU_Quitter.Caption = Localisation$(7)
    Principale.MENU_Edition.Caption = Localisation$(8)
    Principale.MENU_Train.Caption = Localisation$(9)
    Principale.MENU_Couper.Caption = Localisation$(10)
    Principale.MENU_Copier.Caption = Localisation$(11)
    Principale.MENU_Coller.Caption = Localisation$(12)
    Principale.MENU_Inventaire.Caption = Localisation$(13)
    Principale.MENU_Wagon.Caption = Localisation$(14)
    Principale.MENU_Décor.Caption = Localisation$(15)
    Principale.MENU_Voie.Caption = Localisation$(16)
    Principale.VUE_Ajoute.Caption = Localisation$(17)
    Principale.VUE_Supprime.Caption = Localisation$(18)
    Principale.VUE_Origine.Caption = Localisation$(19)
    Principale.VUE_Rotation.Caption = Localisation$(20)
    Principale.MENU_A_Propos.Caption = Localisation$(21)
    Principale.MENU_Parametre.Caption = Localisation$(22)
    Principale.MENU_Element.Caption = Localisation$(23)
    Principale.TUNNEL_Insere.Caption = Localisation$(24)
    Principale.TUNNEL_Supprime.Caption = Localisation$(18)
    Principale.TUNNEL_Inverse.Caption = Localisation$(25)
    Principale.TUNNEL_Ajoute.Caption = Localisation$(17)
    Principale.TUNNEL_Efface.Caption = Localisation$(26)
    '
    For n% = 0 To 3
        Principale.ModeVue(n%).ToolTipText = Localisation$(ClePRINCIPALE% + 11 + n%)
    Next n%
End Sub

'
' ********************************************
' Valide les menus en fonction des protections
' ********************************************
'
Public Sub Valide_Menu()
    Dim Fso As New FileSystemObject
    Dim i%, Lettre$
    Dim Retour As VbMsgBoxResult
    '
    If Command$ = Super$ Then
        '
        ' ***** Force la sauvegarde en mode superuser
        ' ***** Plus les modes accéssibles
        '
        Principale.MENU_Enregistrer.Enabled = True
        Principale.MENU_Wagon.Visible = True
        Principale.MENU_Décor.Visible = True
        Principale.MENU_Voie.Visible = True
        Principale.MENU_Edition_Moins.Visible = True
        Principale.MENU_Fil_De_Fer.Visible = True
    Else
        Do
            For i% = 2 To 25
                Lettre$ = Chr$(65 + i%)
                If Fso.DriveExists(Lettre$) = True Then
                    If Fso.Drives(Lettre$).DriveType = CDRom Then
                        If Fso.Drives(Lettre$).IsReady = True Then
                            If Fso.Drives(Lettre$).VolumeName = "TRAIN 3D" Then
                                Principale.MENU_Enregistrer.Enabled = True
                                Principale.MENU_Wagon.Visible = True
                                Principale.MENU_Décor.Visible = True
                                Principale.MENU_Voie.Visible = True
                                Principale.MENU_Edition_Moins.Visible = True
                                Exit Do
                            End If
                        End If
                    End If
                End If
            Next i%
            Retour = MsgBox(Localisation$(CleDRIVER% + 5), vbAbortRetryIgnore + vbCritical + vbDefaultButton1)
            If Retour = vbAbort Then End
            If Retour = vbIgnore Then Exit Do
        Loop
        '
        Set Fso = Nothing
    End If
    '
    'Else
    '    Principale.MENU_Wagon.Visible = False
    '    Principale.MENU_Décor.Visible = False
    '    Principale.MENU_Voie.Visible = False
    '    Principale.MENU_Edition_Moins.Visible = False
    'End If
End Sub

'
' *******************************************
' Charge la liste des éléments sélectionables
' *******************************************
'
Public Sub Liste_Charge()
    Dim i%, Index%
    '
    ' ***** Création de l'index
    '
    Call Principale.ChoixElement(0).Clear
    Call Principale.ChoixElement(0).AddItem("Centre")
    Call Principale.ChoixElement(1).Clear
    Call Principale.ChoixElement(1).AddItem("Centre")
    Index% = 0
    ReDim IndexChoixElement%(1, Index%)
    IndexChoixElement%(0, Index%) = 0
    IndexChoixElement%(1, Index%) = 0
    '
    ' ***** Les wagons
    '
    For i% = 1 To UBound(ListeTrain())
        '
        Call Principale.ChoixElement(0).AddItem(ListeWagon(ListeTrain(i%).NoWagon%).Nom$)
        Call Principale.ChoixElement(1).AddItem(ListeWagon(ListeTrain(i%).NoWagon%).Nom$)
        Index% = Index% + 1
        ReDim Preserve IndexChoixElement%(1, Index%)
        IndexChoixElement%(0, Index%) = 3
        IndexChoixElement%(1, Index%) = i%
    Next i%
    '
    ' ***** Le décor
    '
    For i% = 1 To UBound(ElementDecor())
        If ElementDecor(i%).NoDecor% <> 0 Then
            Call Principale.ChoixElement(0).AddItem(ListeDecor(ElementDecor(i%).NoDecor%).Nom$)
            Call Principale.ChoixElement(1).AddItem(ListeDecor(ElementDecor(i%).NoDecor%).Nom$)
            Index% = Index% + 1
            ReDim Preserve IndexChoixElement%(1, Index%)
            IndexChoixElement%(0, Index%) = 2
            IndexChoixElement%(1, Index%) = i%
        End If
    Next i%
    '
    ' ***** Le réseau
    '
    For i% = 1 To UBound(Reseau())
        If Reseau(i%).NoVoie% <> 0 Then
            Call Principale.ChoixElement(0).AddItem(Voie(Reseau(i%).NoVoie%).Libelle$(0))
            Call Principale.ChoixElement(1).AddItem(Voie(Reseau(i%).NoVoie%).Libelle$(0))
            Index% = Index% + 1
            ReDim Preserve IndexChoixElement%(1, Index%)
            IndexChoixElement%(0, Index%) = 1
            IndexChoixElement%(1, Index%) = i%
        End If
    Next i%
    '
    Principale.ChoixElement(0).ListIndex = 0
    Principale.ChoixElement(1).ListIndex = 0
End Sub

'
' ***********************
' Charge un fichier .wall
' ***********************
'
Public Function Charge_Wall(File$) As Direct3DRMMeshBuilder3
    Dim i%
    Dim TMeshBuilder() As Direct3DRMMeshBuilder3
    Dim TempDxTexture(NbTexture%) As Direct3DRMTexture3
    Dim TempTexture(NbTexture%) As TypeTexture
    '
    Call Tools3D.Build_MeshBuilder(dxVue, File$, "", TMeshBuilder(), TempDxTexture(), TempTexture())
    '
    Set Charge_Wall = dxVue.dxD3Drm.CreateMeshBuilder
    Call Charge_Wall.AddMesh(TMeshBuilder(0).CreateMesh)
    '
    For i% = 0 To UBound(TMeshBuilder())
        Set TMeshBuilder(i%) = Nothing
    Next i%
    For i% = 0 To NbTexture%
        Set TempDxTexture(i%) = Nothing
    Next i%
End Function

'
' ***********************
' Charge un mesh du décor
' ***********************
'
Public Sub Décor_Charge_Mesh(n%)
    If Right$(ListeDecor(n%).Fichier$, 2) = ".x" Then
        Set ListeDecor(n%).dxDecor = dxVue.Load_MeshBuilder(ListeDecor(n%).Fichier$)
    Else
        Set ListeDecor(n%).dxDecor = Charge_Wall(ListeDecor(n%).Fichier$)
    End If
    Call ListeDecor(n%).dxDecor.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
End Sub

'
' ***********************
' Charge un mesh de wagon
' ***********************
'
Public Sub Wagon_Charger_Mesh(n%)
    If Right$(ListeWagon(n%).Fichier$, 2) = ".x" Then
        Set ListeWagon(n%).dxWagon = dxVue.Load_MeshBuilder(ListeWagon(n%).Fichier$)
    Else
        Set ListeWagon(n%).dxWagon = Charge_Wall(ListeWagon(n%).Fichier$)
    End If
    Call ListeWagon(n%).dxWagon.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    If ListeWagon(n%).FichierBogie$ <> "" Then
        If Right$(ListeWagon(n%).FichierBogie$, 2) = ".x" Then
            Set ListeWagon(n%).dxBogie = dxVue.Load_MeshBuilder(ListeWagon(n%).FichierBogie$)
        Else
            Set ListeWagon(n%).dxBogie = Charge_Wall(ListeWagon(n%).FichierBogie$)
        End If
        Call ListeWagon(n%).dxBogie.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    End If
    If ListeWagon(n%).FichierInterieur$ <> "" Then
        If Right$(ListeWagon(n%).FichierInterieur$, 2) = ".x" Then
            Set ListeWagon(n%).dxInterieur = dxVue.Load_MeshBuilder(ListeWagon(n%).FichierInterieur$)
        Else
            Set ListeWagon(n%).dxInterieur = Charge_Wall(ListeWagon(n%).FichierInterieur$)
        End If
        Call ListeWagon(n%).dxInterieur.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    End If
End Sub

