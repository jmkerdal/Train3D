Attribute VB_Name = "Voies"
Option Explicit
'
Public Const NbSegment% = 7
Public Type TypeVoie
    ' Référence de la voie
    Ref As String
    Libelle(2) As String
    Zone(1) As Integer
    ' Segments
    Segment(NbSegment%) As New ClassSegment ' Equation des segments
    ' Points
    Terminaison(NbSegment%) As Integer ' Terminaison des voies
    Jonction(NbSegment%) As Integer ' Point jonction non connectable
    Offset(NbSegment%) As Integer ' Indique si on impose l'offset ou si on calcule le point
    dX(NbSegment%) As Single ' Point de départ X imposé
    dz(NbSegment%) As Single ' Point de départ Y
    ' Matrice points/segments
    MatConnecte(NbSegment%, NbSegment%) As Integer ' Matrice des connections
    MatAiguille(NbSegment%, NbSegment%) As Integer ' Matrice des aiguillages
    ' Position Catenaire
    CatenaireSegment(NbSegment%) As Integer
    CatenairePosition(NbSegment%) As Single
    CatenaireSens(NbSegment%) As Integer
    '
    ' ***** Valeurs calculés
    '
    pX(NbSegment%) As Single ' Position X d'un point calculé de façon dynamique
    pZ(NbSegment%) As Single ' Position Y d'un point
    SegmentPoint(NbSegment%, 1) As Integer ' Points associés au segment
    Normal(NbSegment%) As D3DVECTOR ' Vecteur normal aux terminaisons
    AngleNormal(NbSegment%) As Single ' Angle aux terminaisons
    AiguillePosition As Integer ' Nombre de positions d'aiguillage
    Inventaire As Integer ' Compteur pour les inventaires
    '
    VoieMeshBuilder As Direct3DRMMeshBuilder3 ' Mesh généré à partir des infos
    VoieLightMeshBuilder As Direct3DRMMeshBuilder3 ' Mesh généré à partir des infos
    VoieCatenaireMeshBuilder As Direct3DRMMeshBuilder3 ' Mesh des cables des catenaires
    VoieTexture(2) As Direct3DRMTexture3 ' Sauvegarde des textures des voies
End Type
Public Voie() As TypeVoie
Public FormeVoie() As New ClassCatenaire
Public SansTextureDynamique As Boolean
'
Public Type TypeReseau
    NoVoie As Integer ' N° de la voie
    Connecte(NbSegment%) As Integer ' Pointe sur la voie suivante
    Entree(NbSegment%) As Integer ' Pointe sur l'entrée de la voie suivante
    ' Valeurs calculées
    Calcule As Boolean ' Indique si la position a été calculée
    Position As D3DVECTOR ' Position de la voie
    Angle As Single ' Angle de la voie
    ' Valeurs pour l'exploitation
    Aiguille As Integer ' Position de l'aiguillage
    AiguilleForce As Integer ' Position forcée de l'aiguillage
    Dessus As Boolean ' Indique si un train est passé dessus
    MetCatenaire(NbSegment%) As Boolean
End Type
Public Reseau() As TypeReseau
'
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'
' ******************************
' Charge la définition des voies
' ******************************
'
Public Sub Charger(Fichier$)
    Dim n%, i%, j%, a!
    Dim nbs%, f%
    '
    f% = FreeFile()
    Call Tools.Open_File(Fichier$, f%, OPEN_NORMAL)
    Input #f%, nbs%
    ReDim Voie(nbs%) As TypeVoie
    ReDim FormeVoie(nbs%) As New ClassCatenaire
    For n% = 1 To nbs%
        With Voie(n%)
            Input #f%, .Libelle$(0)
            Input #f%, .Zone(0)
            Input #f%, .Zone(1)
            Input #f%, .Libelle$(1)
            Input #f%, .Libelle$(2)
            Input #f%, .Ref
            For i% = 0 To NbSegment%
                Input #f%, a!: .Segment(i%).Longueur = a!
                Input #f%, a!: .Segment(i%).Rayon = a!
                Input #f%, a!: .Segment(i%).Angle = a!
                Input #f%, a!: .Segment(i%).Rotation = a!
                '
                Input #f%, a!: .Offset(i%) = a!
                Input #f%, a!: .dX!(i%) = a!
                Input #f%, a!: .dz!(i%) = a!
                Input #f%, .Terminaison(i%)
                Input #f%, .Jonction(i%)
                Input #f%, .CatenaireSegment%(i%)
                Input #f%, .CatenairePosition!(i%)
                Input #f%, .CatenaireSens%(i%)
                '
                For j% = 0 To NbSegment%
                    If i% <> j% Then
                        Input #f%, a!: .MatConnecte(i%, j%) = a!
                        Input #f%, a!: .MatAiguille(i%, j%) = a!
                    End If
                Next j%
            Next i%
        End With
        Call Voies.Libelle_Cree(n%)
        Call FormeVoie(n%).Cree_Couche(n%, 16) ' Génére la forme du catenaire
    Next n%
    Close #f%
End Sub

'
' ************************
' Sauve la liste des voies
' ************************
'
Public Sub Sauver(Fichier$)
    Dim n%, i%, j%, a%
    Dim f%
    '
    f% = FreeFile()
    Open Fichier$ For Output As #f%
    Print #f%, UBound(Voie())
    For n% = 1 To UBound(Voie())
        With Voie(n%)
        Print #f%, .Libelle$(0)
        Print #f%, .Zone(0);
        Print #f%, .Zone(1)
        Print #f%, .Libelle$(1)
        Print #f%, .Libelle$(2)
        Print #f%, .Ref
            For i% = 0 To NbSegment%
                Print #f%, .Segment(i%).Longueur;
                Print #f%, .Segment(i%).Rayon;
                Print #f%, .Segment(i%).Angle;
                Print #f%, .Segment(i%).Rotation;
                '
                Print #f%, .Offset(i%);
                Print #f%, .dX!(i%);
                Print #f%, .dz!(i%);
                Print #f%, .Terminaison(i%)
                Print #f%, .Jonction(i%)
                Print #f%, .CatenaireSegment%(i%);
                Print #f%, .CatenairePosition!(i%);
                Print #f%, .CatenaireSens%(i%)
                '
                For j% = 0 To NbSegment%
                    If i% <> j% Then
                        Print #f%, .MatConnecte(i%, j%);
                        Print #f%, .MatAiguille(i%, j%);
                    End If
                Next j%
                Print #f%,
            Next i%
        End With
    Next n%
    Close #f%
End Sub

'
' ****************************
' Déplace un point sur un rail
' ****************************
'
Public Sub Deplace(Bogie As TypeBogie, d!, ptSortie%)
    Dim s%, P%
    Dim l!
    Dim n%
    '
    n% = Reseau(Bogie.BogieReseau%).NoVoie%
    l! = Voie(n%).Segment(Bogie.BogieSegment% - 1).Segment_Taille()
    If Bogie.BogiePosition! + d! * Bogie.BogieSens% > l! Or Bogie.BogiePosition! + d! * Bogie.BogieSens% < 0 Then
'Debug.Print "d! avant"; d!
        Bogie.Jonction = True
        If Bogie.BogieSens% = 1 Then
            P% = Voie(n%).SegmentPoint(Bogie.BogieSegment% - 1, 1)
            d! = d! - (l! - Bogie.BogiePosition!)
        Else
            P% = Voie(n%).SegmentPoint(Bogie.BogieSegment% - 1, 0)
            d! = d! - Bogie.BogiePosition!
        End If
'Debug.Print "d! après"; d!
        s% = Selectionne_Segment(n%, P%, Bogie)
'Debug.Print "s="; s%; " d="; d!
        If s% = 0 Then
            '
            ' ***** Dépassement
            '
'Debug.Print "On sort par le point"; p%; " reste="; d!
            ptSortie% = P%
        Else
            Bogie.BogieSegment% = s%
            If Bogie.BogieSens% = 1 Then
                Bogie.BogiePosition! = 0
            Else
                Bogie.BogiePosition! = Voie(n%).Segment(Bogie.BogieSegment% - 1).Segment_Taille()
            End If
'Debug.Print "On se connecte sur le segment"; Bogie.BogieSegment%
            Call Deplace(Bogie, d!, P%)
        End If
'Debug.Print "Reste"; d!
    Else
        Bogie.BogiePosition! = Bogie.BogiePosition! + d! * Bogie.BogieSens%
        d! = 0
'Debug.Print "Avance jusqu'à"; Bogie.BogiePosition!
    End If
End Sub

'
' *********************************************************
' Séléctionne le segment en fonction du point d'application
' *********************************************************
'
Public Function Selectionne_Segment%(n%, LePoint%, Bogie As TypeBogie) ' Position%, Sens%)
    Dim i%, c%
'Debug.Print "Point d'arrivée"; LePoint%; " Sens"; Bogie.BogieSens%
    If Bogie.BogieSens% <> -1 Then
        For i% = LePoint% To NbSegment%
            If Selectionne_Segment% = 0 Then
                ' ***** Point entrant
                If Voie(n%).MatConnecte(LePoint% - 1, i%) <> 0 Then
                    c% = Voie(n%).MatAiguille(LePoint% - 1, i%)
                    If c% = 0 Or (c% = Reseau(Bogie.BogieReseau%).AiguilleForce%) Then
'Debug.Print "Point entrant"; c%
                        Selectionne_Segment% = Voie(n%).MatConnecte(LePoint% - 1, i%)
                        Bogie.BogieSens% = 1
                    End If
                End If
            End If
        Next i%
    End If
    If Bogie.BogieSens% <> 1 Then
        For i% = 0 To LePoint% - 1
            If Selectionne_Segment% = 0 Then
                ' ***** Point sortant
                If Voie(n%).MatConnecte(i%, LePoint% - 1) <> 0 Then
                    c% = Voie(n%).MatAiguille(i%, LePoint% - 1)
                    If c% = 0 Or (c% = Reseau(Bogie.BogieReseau%).AiguilleForce%) Then
'Debug.Print "Point sortant"; c%
                        Selectionne_Segment% = Voie(n%).MatConnecte(i%, LePoint% - 1)
                        Bogie.BogieSens% = -1
                    End If
                End If
            End If
        Next i%
    End If
End Function

'
' ********************************************************
' Crée un MeshBuilder à partir des informations de la voie
' ********************************************************
'
Public Function Cree_Couche_MeshBuilder(No%, Couche%) As Direct3DRMMeshBuilder3
    Dim i%, j%, k%
    Dim n%, t!, X!, z!
    Dim nTraverse%, dTraverse!
    Dim v As D3DVECTOR
    Dim vt As D3DVECTOR
    Dim Vx!, Vz!
    '
    Dim TempFrame() As Direct3DRMFrame3
    Dim TempMesh() As Direct3DRMMesh
    Dim TempHeurtoir(NbSegment%) As Direct3DRMFrame3
    Dim nbFrame%
    '
    ' ***** Calcul des segments
    '
    Call Initialisation.Recalcule_Point(No%)
    nbFrame% = 0
    ReDim TempFrame(nbFrame%) As Direct3DRMFrame3
    ReDim TempMesh(nbFrame%) As Direct3DRMMesh ' Copie des MeshBuilders pour mise à l'échelle
    Set TempFrame(nbFrame%) = dxVue.dxD3Drm.CreateFrame(Nothing)
    For i% = 1 To NbSegment%
        Set TempHeurtoir(i%) = dxVue.dxD3Drm.CreateFrame(TempFrame(0))
    Next i%
    '
    For i% = 0 To NbSegment% - 1
        For j% = i% + 1 To NbSegment%
            If Voie(No%).MatConnecte(i%, j%) <> 0 Then
                Call Calcule_Point(No%, i%, j%)
                n% = Voie(No%).MatConnecte(i%, j%) - 1 ' N° segment
                t! = Voie(No%).Segment(n%).Segment_Taille() ' Taille segment
                If Voie(No%).Offset(i%) = 0 Then
                    X! = Voie(No%).pX!(i%)
                    z! = Voie(No%).pZ!(i%)
                Else
                    X! = Voie(No%).dX!(i%)
                    z! = Voie(No%).dz!(i%)
                End If
                If Couche% = 3 And Voie(No%).Segment(n%).Segment_Droit = True Then
                    v = Voie(No%).Segment(n%).Point(0)
                    vt = Voie(No%).Segment(n%).Tangente(0)
                    nbFrame% = nbFrame% + 1
                    ReDim Preserve TempFrame(nbFrame%) As Direct3DRMFrame3
                    ReDim Preserve TempMesh(nbFrame%) As Direct3DRMMesh
                    Set TempFrame(nbFrame%) = dxVue.dxD3Drm.CreateFrame(TempFrame(0))
                    Set TempMesh(nbFrame%) = dxTraverse(Couche%).CreateMesh
                    Call TempMesh(nbFrame).Translate(0, 0, 3.5)
                    Call TempMesh(nbFrame).ScaleMesh(1, 1, t! / 7)
                    Call TempFrame(nbFrame%).AddVisual(TempMesh(nbFrame%))
                    Call TempFrame(nbFrame%).SetPosition(TempFrame(0), v.X + X!, v.Y, v.z + z!)
                    Call TempFrame(nbFrame%).SetOrientation(TempFrame(0), vt.X, vt.Y, vt.z, 0, 1, 0)
                    If Voie(No%).Terminaison(j%) <> 0 Then
                        Call TempHeurtoir(j%).AddVisual(dxHeurtoir)
                        v = Voie(No%).Segment(n%).Point(t! - 3.5)
                        vt = Voie(No%).Segment(n%).Tangente(t - 3.5)
                        Call TempHeurtoir(j%).SetPosition(TempFrame(0), v.X + X!, v.Y, v.z + z!)
                        Call TempHeurtoir(j%).SetOrientation(TempFrame(0), vt.X, vt.Y, vt.z, 0, 1, 0)
                    End If
                Else
                    nTraverse% = t! / 7
                    dTraverse! = t! / (nTraverse% + 1)
                    For k% = 0 To nTraverse%
                        v = Voie(No%).Segment(n%).Point(k% * dTraverse! + dTraverse! / 2)
                        vt = Voie(No%).Segment(n%).Tangente(k% * dTraverse! + dTraverse! / 2)
                        nbFrame% = nbFrame% + 1
                        ReDim Preserve TempFrame(nbFrame%) As Direct3DRMFrame3
                        ReDim Preserve TempMesh(nbFrame%) As Direct3DRMMesh
                        Set TempFrame(nbFrame%) = dxVue.dxD3Drm.CreateFrame(TempFrame(0))
                        Set TempMesh(nbFrame%) = dxTraverse(Couche%).CreateMesh
                        Call TempFrame(nbFrame%).AddVisual(TempMesh(nbFrame%))
                        Call TempFrame(nbFrame%).SetPosition(TempFrame(0), v.X + X!, v.Y, v.z + z!)
                        Call TempFrame(nbFrame%).SetOrientation(TempFrame(0), vt.X, vt.Y, vt.z, 0, 1, 0)
                    Next k%
                    If Couche% <= 1 And Voie(No%).Terminaison(j%) <> 0 Then
                        Call TempHeurtoir(j%).AddVisual(dxHeurtoir)
                        Call TempHeurtoir(j%).SetPosition(TempFrame(nbFrame%), 0, 0, 0)
                        Call TempHeurtoir(j%).SetOrientation(TempFrame(nbFrame%), 0, 0, 1, 0, 1, 0)
                    End If
                End If
            End If
        Next j%
    Next i%
    '
    Set Cree_Couche_MeshBuilder = dxVue.dxD3Drm.CreateMeshBuilder
    Call Cree_Couche_MeshBuilder.AddFrame(TempFrame(0))
    For i% = nbFrame% To 1
        Call TempFrame(i%).DeleteVisual(TempMesh(i%))
        Set TempFrame(i%) = Nothing
        Set TempMesh(i%) = Nothing
    Next i%
    Set TempFrame(0) = Nothing
    For i% = 1 To NbSegment%
        If (Couche% = 3 Or Couche% = 1) And Voie(No%).Terminaison(i%) <> 0 Then
            Call TempHeurtoir(i%).DeleteVisual(dxHeurtoir)
        End If
        Set TempHeurtoir(i%) = Nothing
    Next i%
End Function

'
' *******************************
' Retourne le point d'application
' de la bogie sur le réseau
' *******************************
'
Public Sub Position_Bogie(Bogie As TypeBogie, Pos As D3DVECTOR, Dir!)
    Dim v As D3DVECTOR, l!
    Dim n%, pt As D3DVECTOR
    Dim Vy As D3DVECTOR
    Dim P%
    Vy.Y = 1
    n% = Reseau(Bogie.BogieReseau%).NoVoie%
    P% = Voie(n%).SegmentPoint(Bogie.BogieSegment% - 1, 0) - 1
    '
    ' ***** Message sur erreur positionnement bogie: ex pb caténaire
    '
    If P% = -1 Then
        Call MsgBox("Possible problem in rail definition", vbExclamation)
        Exit Sub
    End If
    If Voie(n%).Offset(P%) = 0 Then
        v.X = Voie(n%).pX!(P%)
        v.z = Voie(n%).pZ!(P%)
    Else
        v.X = Voie(n%).dX!(P%)
        v.z = Voie(n%).dz!(P%)
    End If
    pt = Voie(n%).Segment(Bogie.BogieSegment% - 1).Point(Bogie.BogiePosition!)
    Call dxVue.dX7.VectorAdd(v, v, pt)
    '
    l! = dxVue.dX7.VectorModulus(v)
    Dir! = Reseau(Bogie.BogieReseau%).Angle
    Call dxVue.dX7.VectorRotate(v, v, Vy, Dir! * DegRad!)
    Call dxVue.dX7.VectorScale(v, v, l!)
    Call dxVue.dX7.VectorAdd(Pos, v, Reseau(Bogie.BogieReseau%).Position)
    Dir! = Dir! + Voie(n%).Segment(Bogie.BogieSegment% - 1).Theta(Bogie.BogiePosition!)
End Sub

'
' ***********************
' Déplacement d'une bogie
' ***********************
'
Public Function Deplace_Bogie(Bogie As TypeBogie, ByVal d!) As Boolean
    Dim P%
    Dim VSuivant%, VEntree%
    Dim a%
    Do
        Call Deplace(Bogie, d!, P%)
'Debug.Print "PtRetour="; p%; " reste="; d!
        If P% = 0 Then
            Reseau(Bogie.BogieReseau%).Dessus = True ' Un train est dessus
            Exit Do
        End If
'Debug.Print "Sort par le point"; p%
        VSuivant% = Reseau(Bogie.BogieReseau).Connecte(P% - 1)
        If VSuivant% = 0 Then
'Debug.Print "PtRetour="; p%; " reste="; d!
            Bogie.BogiePosition = Voie(Reseau(Bogie.BogieReseau%).NoVoie%).Segment(Bogie.BogieSegment% - 1).Segment_Taille()
            Exit Function ' On est sur un heurtoir
        End If
        VEntree% = Reseau(Bogie.BogieReseau%).Entree(P% - 1) + 1
        Bogie.BogieReseau% = VSuivant%
        'if reseau(vsuivant%).Aiguille
'Debug.Print "Suivant"; VSuivant; " Entrée"; VEntree%
        Bogie.BogieSegment% = Entre_Segment(Reseau(VSuivant%).NoVoie%, VEntree%, Bogie)
        '
        Reseau(VSuivant%).Dessus = True ' Un train passe dessus
        '
        ' ***** Bascule l'aiguillage si on arrive à l'envers
        '
        If Voie(Reseau(VSuivant%).NoVoie%).Jonction(VEntree% - 1) = 0 Then
            a% = Voie(Reseau(VSuivant%).NoVoie%).MatAiguille( _
            Voie(Reseau(VSuivant%).NoVoie%).SegmentPoint(Bogie.BogieSegment% - 1, 0) - 1, _
            Voie(Reseau(VSuivant%).NoVoie%).SegmentPoint(Bogie.BogieSegment% - 1, 1) - 1)
            If a% <> 0 Then
                If Reseau(VSuivant%).AiguilleForce% <> a% Then
                    Reseau(VSuivant%).AiguilleForce% = a%
                End If
            End If
        End If
        '
'Debug.Print "Segment"; Bogie.BogieSegment%; " d!="; d!
        If Bogie.BogieSens% = 1 Then
            'Bogie.BogiePosition! = d!
            Bogie.BogiePosition! = 0
        Else
            'Bogie.BogiePosition! = Voie(Reseau(VSuivant%).NoVoie%).Segment(Bogie.BogieSegment% - 1).Segment_Taille() - d!
            Bogie.BogiePosition! = Voie(Reseau(VSuivant%).NoVoie%).Segment(Bogie.BogieSegment% - 1).Segment_Taille()
'Debug.Print "No"; Reseau(VSuivant%).NoVoie%; " Taille="; Voie(Reseau(VSuivant%).NoVoie%).Segment(Bogie.BogieSegment% - 1).Segment_Taille()
        End If
'Debug.Print "Position="; Bogie.BogiePosition
        P% = 0
    Loop
    Deplace_Bogie = True
End Function

Public Function Entre_Segment%(n%, LePoint%, Bogie As TypeBogie)
    Dim i%, c%
'Debug.Print "Point d'entrée"; LePoint%; " Sens"; Bogie.BogieSens%
    For i% = 0 To NbSegment%
        If i% <> LePoint% - 1 And Entre_Segment% = 0 Then
            ' ***** Point entrant
'Debug.Print "i="; i%; " connecte"; Voie(n%).Connecte(LePoint% - 1, i%)
            If Voie(n%).MatConnecte(LePoint% - 1, i%) <> 0 Then
                c% = Voie(n%).MatAiguille(LePoint% - 1, i%)
'Debug.Print "c="; c%; " aiguille"; Reseau(Bogie.BogieReseau%).Aiguille%
                If c% = 0 Or (c% = Reseau(Bogie.BogieReseau%).AiguilleForce%) Then
'Debug.Print "Point entrée"; c%
                    Entre_Segment% = Voie(n%).MatConnecte(LePoint% - 1, i%)
                    If i% < LePoint% - 1 Then
                        Bogie.BogieSens% = -1
                    Else
                        Bogie.BogieSens% = 1
                    End If
                End If
            End If
        End If
    Next i%
    '
    If Entre_Segment% <> 0 Then Exit Function
    '
    ' Relance le test si on a échoué
    ' Recherche le segment que l'on prend à rebrousse poil
    '
    For i% = 0 To NbSegment%
        If i% <> LePoint% - 1 And Entre_Segment% = 0 Then
            ' ***** Point entrant
            If Voie(n%).MatConnecte(LePoint% - 1, i%) <> 0 Then
                If i% < LePoint% - 1 Then ' Toujours vrai quand on rentre à l'envert
                    Entre_Segment% = Voie(n%).MatConnecte(LePoint% - 1, i%)
                    If i% < LePoint% - 1 Then
                        Bogie.BogieSens% = -1
                    Else
                        Bogie.BogieSens% = 1
                    End If
                End If
            End If
        End If
    Next i%
End Function

'
' *******************
' Position d'un point
' *******************
'
Public Function Position_Point(n%, P%) As D3DVECTOR
    Dim l!
    Dim Vy As D3DVECTOR
    Dim v%
    v% = Reseau(n%).NoVoie%
    Vy.Y = 1
    If Voie(v%).Offset(P%) = 0 Then
        Position_Point.X = Voie(v%).pX!(P%)
        Position_Point.z = Voie(v%).pZ!(P%)
    Else
        Position_Point.X = Voie(v%).dX!(P%)
        Position_Point.z = Voie(v%).dz!(P%)
    End If
    If P% = 0 Or Position_Point.X <> 0 Or Position_Point.z <> 0 Then
        l! = dxVue.dX7.VectorModulus(Position_Point)
        Call dxVue.dX7.VectorRotate(Position_Point, Position_Point, Vy, Reseau(n%).Angle! * DegRad!)
        Call dxVue.dX7.VectorScale(Position_Point, Position_Point, l!)
        Call dxVue.dX7.VectorAdd(Position_Point, Position_Point, Reseau(n%).Position)
    End If
End Function

'
' *************************
' Ajoute une voie au réseau
' *************************
'
Public Sub Ajoute(No_Voie%, Angle_Voie!)
    Dim i%, j%, t%, n%
    Call Initialisation.Vue_Decharge
    n% = UBound(Reseau())
    For i% = 1 To n%
        If Reseau(i%).NoVoie% = 0 Then
            t% = i%
        End If
    Next i%
    If t% = 0 Then
        t% = n% + 1
        ReDim Preserve Reseau(t%) As TypeReseau
        ReDim Preserve dxFrameVoie(3 + NbSegment%, t%) As Direct3DRMFrame3
        Set dxFrameVoie(0, t%) = dxVue.dxD3Drm.CreateFrame(dxVue.dxScene)
        Call dxFrameVoie(0, t%).SetAppData(t%)
        Call dxFrameVoie(0, t%).SetName("Voie")
        For j% = 1 To 3 + NbSegment%
            Set dxFrameVoie(j%, t%) = dxVue.dxD3Drm.CreateFrame(Nothing)
        Next j%
    End If
    Reseau(t%).NoVoie% = No_Voie%
    Reseau(t%).Position.X = PosSouris.X * REDUCTION%
    Reseau(t%).Position.z = PosSouris.z * REDUCTION%
    Reseau(t%).Angle! = Angle_Voie!
    Reseau(t%).Aiguille% = 1
    TypeCopie% = 1
    NoCopie% = No_Voie%
    AngleCopie! = Angle_Voie!
    Call Vue.Calcule_Reseau
    Call Initialisation.Vue_Charge
    Principale.MENU_Coller.Enabled = True
    Call Initialisation.Liste_Charge
End Sub

'
' ******************************************
' Crée une texture à partir d'un meshbuilder
' ******************************************
'
Public Function CreateTextureFromMesh(dxEngine As ClassDirectX, LeMesh As Direct3DRMMeshBuilder3, LeBox As D3DRMBOX, MaxX%, MaxY%) As Direct3DRMTexture3
    Dim dv As D3DVECTOR, sv As D3DVECTOR
    Dim dxv4D As D3DRMVECTOR4D
    Dim dxSurfaceBox As D3DRMBOX, DDSD As DDSURFACEDESC2
    Dim dxTempSurface As DirectDrawSurface4
    Dim TX%, TY%
    Dim RectSource As RECT
    Dim RectDest As RECT
    Dim R1%, R2%, Recul%
    Dim DestDC As Long
    '
    Call LeMesh.GetBox(LeBox)
    PosCamera.X = (LeBox.Min.X + LeBox.Max.X) / 2
    PosCamera.z = (LeBox.Min.z + LeBox.Max.z) / 2
    Call dxEngine.dxCamera.SetPosition(dxEngine.dxScene, PosCamera.X, PosCamera.Y, PosCamera.z)
    Call dxEngine.dxCamera.SetOrientation(dxEngine.dxScene, 0, -1, 0, 0, 0, 1)
    Call dxEngine.dxViewport.SetProjection(D3DRMPROJECT_ORTHOGRAPHIC)
    '
    R1% = LeBox.Max.X - LeBox.Min.X
    R2% = LeBox.Max.z - LeBox.Min.z
    If R1% > R2% Then
        Recul% = R1% / 2 * 1.01
    Else
        Recul% = R2% / 2 * 1.01
    End If
    If Recul% < 150 Then Recul% = 150
    Call dxEngine.dxViewport.SetField(Recul%)
    '
    Call dxEngine.dxScene.AddVisual(LeMesh)
    Call dxEngine.Render(False)
    '
    ' ***** Récupère texture
    '
    sv = LeBox.Min
    Call dxEngine.dxScene.Transform(dv, sv)
    Call dxEngine.dxViewport.Transform(dxv4D, dv)
    dxSurfaceBox.Min.X = dxv4D.X / dxv4D.w
    dxSurfaceBox.Min.Y = dxv4D.Y / dxv4D.w
    sv = LeBox.Max
    Call dxEngine.dxScene.Transform(dv, sv)
    Call dxEngine.dxViewport.Transform(dxv4D, dv)
    dxSurfaceBox.Max.X = dxv4D.X / dxv4D.w
    dxSurfaceBox.Max.Y = dxv4D.Y / dxv4D.w
    TX% = -dxSurfaceBox.Min.X + dxSurfaceBox.Max.X
    TY% = dxSurfaceBox.Min.Y - dxSurfaceBox.Max.Y
    '
    If TX% > 512 Then
        TX% = 1024
    ElseIf TX% > 256 Then
        TX% = 512
    ElseIf TX% > 128 Then
        TX% = 256
    ElseIf TX% > 64 Then
        TX% = 128
    ElseIf TX% > 32 Then
        TX% = 64
    ElseIf TX% > 16 Then
        TX% = 32
    Else
        TX% = 16
    End If
    If TY% > 512 Then
        TY% = 1024
    ElseIf TY% > 256 Then
        TY% = 512
    ElseIf TY% > 128 Then
        TY% = 256
    ElseIf TY% > 64 Then
        TY% = 128
    ElseIf TY% > 32 Then
        TY% = 64
    ElseIf TY% > 16 Then
        TY% = 32
    Else
        TY% = 16
    End If
    If TX% > MaxX% Then TX% = MaxX%
    If TY% > MaxY% Then TY% = MaxY%
    Call ProgramLog.Write_File(0, "Texture size:" + Str$(TX%) + Str$(TY%))
    '
    With DDSD
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .lWidth = TX%
        .lHeight = TY%
        .ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY Or DDSCAPS_TEXTURE
    End With
    Set dxTempSurface = dxEngine.dxDirectDraw.CreateSurface(DDSD)
    RectSource.Left = dxSurfaceBox.Min.X
    RectSource.Top = dxSurfaceBox.Max.Y
    RectSource.Right = dxSurfaceBox.Max.X
    RectSource.Bottom = dxSurfaceBox.Min.Y
    '
    RectDest.Top = 0
    RectDest.Left = 0
    RectDest.Right = -dxSurfaceBox.Min.X + dxSurfaceBox.Max.X
    RectDest.Bottom = dxSurfaceBox.Min.Y - dxSurfaceBox.Max.Y
    '
    Call ProgramLog.Write_File(0, "Surface StretchBlt")
    Call dxTempSurface.Blt(RectDest, dxEngine.dxBack, RectSource, DDBLT_WAIT)
    DestDC = dxTempSurface.GetDC
    Call StretchBlt(DestDC, 0, 0, TX%, TY%, DestDC, 0, 0, RectDest.Right, RectDest.Bottom, vbSrcCopy)
    Call dxTempSurface.ReleaseDC(DestDC)
    '
    Call ProgramLog.Write_File(0, "CreateTextureFromSurface")
    Set CreateTextureFromMesh = dxEngine.dxD3Drm.CreateTextureFromSurface(dxTempSurface)
    Call CreateTextureFromMesh.SetDecalTransparency(D_TRUE)
    '
    ' ***** Efface les objets
    '
    Call dxEngine.dxScene.DeleteVisual(LeMesh)
    Set dxTempSurface = Nothing
End Function

'
' *********************************
' Fabrique les meshbuilder à partir
' des données de la voie
' *********************************
'
Public Sub Cree_Meshbuilders()
    Dim n%, i%, j%
    '
    Dim TempMesh As Direct3DRMMeshBuilder3
    Dim dxBox As D3DRMBOX
    Dim dv As D3DVECTOR, sv As D3DVECTOR
    Dim dxEngine As New ClassDirectX
    Dim dxFrame As Direct3DRMFrame3
    Dim TX%, TY%
    Dim LeCatenaire As New ClassCatenaire
    Dim Pa(3) As D3DVECTOR, P1(3) As D3DVECTOR
    '
    CameraRecul% = 250
    PosCamera.Y = 500
    '
    If SansTextureDynamique = True Then
        Call ProgramLog.Write_File(0, "No dynamic texture")
    Else
        If dxEngine.TextureMaxWidth < 1024 Then
            TX% = dxEngine.TextureMaxWidth
        Else
            TX% = 1024
        End If
        If dxEngine.TextureMaxHeight < 1024 Then
            TY% = dxEngine.TextureMaxHeight
        Else
            TY% = 1024
        End If
        If Command$ = Super$ Then
            ' ***** Pour les tests
            TX% = 256
            TY% = 256
        End If
        Call ProgramLog.Write_File(0, "Create dynamic render engine")
        Call ProgramLog.Write_File(TX%, "Width")
        Call ProgramLog.Write_File(TY%, "Height")
        Call dxEngine.Create_3DRM(Vue.Affiche, TX%, TY%, Mode3DSurface)
    End If
    '
    Set Voie(0).VoieMeshBuilder = Nothing
    Set Voie(0).VoieLightMeshBuilder = Nothing
    Set Voie(0).VoieCatenaireMeshBuilder = Nothing
    Set Voie(0).VoieMeshBuilder = dxVue.dxD3Drm.CreateMeshBuilder
    Set Voie(0).VoieLightMeshBuilder = dxVue.dxD3Drm.CreateMeshBuilder
    Set Voie(0).VoieCatenaireMeshBuilder = dxVue.dxD3Drm.CreateMeshBuilder
    '
    For n% = 1 To UBound(Voie())
        Call ProgramLog.Write_File(n%, "Create track")
        DoEvents
        With Voie(n%)
            Set .VoieMeshBuilder = Nothing
            Set .VoieLightMeshBuilder = Nothing
            Set .VoieCatenaireMeshBuilder = Nothing
            '
            ' ***** Création du catenaire
            '
            Call dxConstruitMeshBuilder.Init
            Call LeCatenaire.Cree_Couche(n%, 0.5)
            For i% = 1 To LeCatenaire.NbLRectangle%
                Call LeCatenaire.Liste_Rectangle(i%, Pa(0), Pa(1), Pa(2), Pa(3))
                For j% = 0 To 3
                    Pa(j%).Y = 60
                    P1(j%) = Pa(j%)
                    P1(j%).Y = 61
                Next j%
                Call dxConstruitMeshBuilder.Add_Rectangle(P1(0), P1(1), P1(2), P1(3), -1, &HFF000000)
                Call dxConstruitMeshBuilder.Add_Rectangle(Pa(0), Pa(3), Pa(2), Pa(1), -1, &HFF000000)
                Call dxConstruitMeshBuilder.Add_Rectangle(P1(3), P1(2), Pa(2), Pa(3), -1, &HFF000000)
                Call dxConstruitMeshBuilder.Add_Rectangle(P1(1), P1(0), Pa(0), Pa(1), -1, &HFF000000)
                '
                For j% = 0 To 3
                    Pa(j%).Y = 75
                    P1(j%).Y = 76
                Next j%
                Call dxConstruitMeshBuilder.Add_Rectangle(P1(0), P1(1), P1(2), P1(3), -1, &HFF000000)
                Call dxConstruitMeshBuilder.Add_Rectangle(Pa(0), Pa(3), Pa(2), Pa(1), -1, &HFF000000)
                Call dxConstruitMeshBuilder.Add_Rectangle(P1(3), P1(2), Pa(2), Pa(3), -1, &HFF000000)
                Call dxConstruitMeshBuilder.Add_Rectangle(P1(1), P1(0), Pa(0), Pa(1), -1, &HFF000000)
                '
                Pa(0).Y = 61
                Pa(3).Y = 61
                P1(0).Y = 75
                P1(3).Y = 75
                Call dxConstruitMeshBuilder.Add_Rectangle(P1(0), Pa(0), Pa(3), P1(3), -1, &HFF000000)
                Call dxConstruitMeshBuilder.Add_Rectangle(Pa(0), P1(0), P1(3), Pa(3), -1, &HFF000000)
            Next i%
            Set .VoieCatenaireMeshBuilder = dxVue.dxD3Drm.CreateMeshBuilder
            Call dxConstruitMeshBuilder.Build(.VoieCatenaireMeshBuilder, dxSolTexture())
            Call .VoieCatenaireMeshBuilder.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
            '
            ' ***** Création du nouveau mesh light
            '
            If SansTextureDynamique = True Then
                Set .VoieLightMeshBuilder = Cree_Couche_MeshBuilder(n%, 1)
            Else
                Set TempMesh = Cree_Couche_MeshBuilder(n%, 1)
                Set .VoieTexture(0) = Voies.CreateTextureFromMesh(dxEngine, TempMesh, dxBox, TX%, TY%)
                Set TempMesh = Nothing
                Call dxConstruitMeshBuilder.Init
                dxBox.Min.Y = 1
                dxBox.Max.Y = 1
                dv.X = dxBox.Max.X
                dv.Y = 1
                dv.z = dxBox.Min.z
                sv.X = dxBox.Min.X
                sv.Y = 1
                sv.z = dxBox.Max.z
                Call dxConstruitMeshBuilder.Add_Rectangle(dxBox.Min, sv, dxBox.Max, dv, 0, &HFFFFFFFF)
                Set .VoieLightMeshBuilder = dxVue.dxD3Drm.CreateMeshBuilder
                Call dxConstruitMeshBuilder.Build(.VoieLightMeshBuilder, .VoieTexture())
            End If
            Call .VoieLightMeshBuilder.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
            Call .VoieLightMeshBuilder.Optimize
            '
            ' ***** Création du mesh vue intérieur/poursuite
            '
            If SansTextureDynamique = True Then
                Set .VoieMeshBuilder = Cree_Couche_MeshBuilder(n%, 0)
            Else
                Call dxConstruitMeshBuilder.Init
                Set dxFrame = dxEngine.dxD3Drm.CreateFrame(Nothing)
                '
                Set TempMesh = Cree_Couche_MeshBuilder(n%, 3)
                Call dxFrame.AddVisual(TempMesh)
                Set TempMesh = Nothing
                '
                Set TempMesh = Cree_Couche_MeshBuilder(n%, 2)
                Set .VoieTexture(1) = Voies.CreateTextureFromMesh(dxEngine, TempMesh, dxBox, TX%, TY%)
                Set TempMesh = Nothing
                '
                dxBox.Min.Y = 1
                dxBox.Max.Y = dxBox.Min.Y
                dv.X = dxBox.Max.X
                dv.Y = dxBox.Min.Y
                dv.z = dxBox.Min.z
                sv.X = dxBox.Min.X
                sv.Y = dxBox.Min.Y
                sv.z = dxBox.Max.z
                Call dxConstruitMeshBuilder.Add_Rectangle(dxBox.Min, sv, dxBox.Max, dv, 1, &HFFFFFFFF)
                '
                Set TempMesh = Cree_Couche_MeshBuilder(n%, 4)
                Set .VoieTexture(2) = Voies.CreateTextureFromMesh(dxEngine, TempMesh, dxBox, TX%, TY%)
                Set TempMesh = Nothing
                dxBox.Min.Y = 3.5
                dxBox.Max.Y = dxBox.Min.Y
                dv.X = dxBox.Max.X
                dv.Y = dxBox.Min.Y
                dv.z = dxBox.Min.z
                sv.X = dxBox.Min.X
                sv.Y = dxBox.Min.Y
                sv.z = dxBox.Max.z
                Call dxConstruitMeshBuilder.Add_Rectangle(dxBox.Min, sv, dxBox.Max, dv, 2, &HFFFFFFFF)
                '
                Set TempMesh = dxVue.dxD3Drm.CreateMeshBuilder
                Call dxConstruitMeshBuilder.Build(TempMesh, .VoieTexture())
                Call dxFrame.AddVisual(TempMesh)
                Set TempMesh = Nothing
                Set .VoieMeshBuilder = dxVue.dxD3Drm.CreateMeshBuilder
                Call .VoieMeshBuilder.AddFrame(dxFrame)
                Set dxFrame = Nothing
            End If
            Call .VoieMeshBuilder.ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
            Call .VoieMeshBuilder.Optimize
        End With
    Next n%
    '
    Set LeCatenaire = Nothing
    If SansTextureDynamique = False Then
        Call ProgramLog.Write_File(0, "Destroy dynamic render engine")
        Set dxEngine = Nothing
    End If
End Sub

'
' **********************************
' Calcul le point d'arrivée B,C ou D
' et son angle à l'extremité
' **********************************
'
Public Sub Calcule_Point(No%, o%, P%)
    Dim n%
    n% = Voie(No%).MatConnecte(o%, P%) - 1
    If Voie(No%).pX!(P%) = 0 And Voie(No%).pZ!(P%) = 0 Then
        If Voie(No%).Offset(o%) = 0 And o% <> 0 Then
            Voie(No%).pX!(P%) = Voie(No%).pX!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).X
            Voie(No%).pZ!(P%) = Voie(No%).pZ!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).z
        Else
            Voie(No%).pX!(P%) = Voie(No%).dX!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).X
            Voie(No%).pZ!(P%) = Voie(No%).dz!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).z
            '
            ' ***** Tourne le segment pour avoir la normale dans le sens inverse de l'entrée
            '
            Voie(No%).Segment(n%).Rotation! = Voie(No%).Segment(n%).Rotation! + 180
            Voie(No%).Normal(o%) = Voie(No%).Segment(n%).Tangente(0)
            Voie(No%).AngleNormal(o%) = Voie(No%).Segment(n%).Theta(0)
            Voie(No%).Segment(n%).Rotation! = Voie(No%).Segment(n%).Rotation! - 180
        End If
        Voie(No%).Normal(P%) = Voie(No%).Segment(n%).Tangente(Voie(No%).Segment(n%).Segment_Taille())
        Voie(No%).AngleNormal(P%) = Voie(No%).Segment(n%).Theta(Voie(No%).Segment(n%).Segment_Taille())
    End If
End Sub

'
' *******************
' Création du libellé
' *******************
'
Public Sub Libelle_Cree(n%)
    Dim i%
    Voie(n%).Libelle$(0) = ""
    For i% = 0 To 1
        If Voie(n%).Zone(i%) <> 0 Then
            Voie(n%).Libelle$(0) = Voie(n%).Libelle$(0) + Localisation$(CleRAIL% + Voie(n%).Zone(i%) - 1) + " "
        End If
    Next i%
    If Voie(n%).Libelle$(1) <> "" Then
        Voie(n%).Libelle$(0) = Voie(n%).Libelle$(0) + Voie(n%).Libelle$(1) + " mm "
    End If
    If Voie(n%).Libelle$(2) <> "" Then
        Voie(n%).Libelle$(0) = Voie(n%).Libelle$(0) + Voie(n%).Libelle$(2) + " °"
    End If
    Voie(n%).Libelle$(0) = RTrim$(Voie(n%).Libelle$(0))
End Sub

