Attribute VB_Name = "Module1"
Option Explicit
Public Type TypeVoie
    ' Référence de la voie
    Ref As String
    Libelle As String
    ' Segments
    Segment(3) As New ClassSegment ' Equation des segments
    ' Points
    Terminaison(3) As Integer ' Terminaison des voies
    Jonction(3) As Integer ' Point jonction non connectable
    Offset(3) As Integer ' Indique si on impose l'offset ou si on calcule le point
    dX(3) As Single ' Point de départ X imposé
    dz(3) As Single ' Point de départ Y
    ' Matrice points/segments
    MatConnecte(3, 3) As Integer ' Matrice des connections
    MatAiguille(3, 3) As Integer ' Matrice des aiguillages
    ' Position Catenaire
    CatenaireSegment(3) As Integer
    CatenairePosition(3) As Single
    CatenaireSens(3) As Integer
    '
    ' ***** Valeurs calculés
    '
    pX(3) As Single ' Position X d'un point calculé de façon dynamique
    pZ(3) As Single ' Position Y d'un point
    SegmentPoint(3, 1) As Integer ' Points associés au segment
    Normal(3) As D3DVECTOR ' Vecteur normal aux terminaisons
    AngleNormal(3) As Single ' Angle aux terminaisons
    AiguillePosition As Integer ' Nombre de positions d'aiguillage
    Inventaire As Integer ' Compteur pour les inventaires
    '
    VoieMeshBuilder As Direct3DRMMeshBuilder3 ' Mesh généré à partir des infos
    VoieLightMeshBuilder As Direct3DRMMeshBuilder3 ' Mesh généré à partir des infos
    VoieCatenaireMeshBuilder As Direct3DRMMeshBuilder3 ' Mesh des cables des catenaires
    VoieTexture(2) As Direct3DRMTexture3 ' Sauvegarde des textures des voies
End Type
Public Voie() As TypeVoie
Public FormeVoie As New ClassTunnel

'
' ******************************************
' Recalcule les points associés aux segments
' ******************************************
'
Public Sub Recalcule_Point(No%)
    Dim i%, j%, n%
    Voie(No%).AiguillePosition% = 0
    For i% = 0 To 3
        Voie(No%).pX!(i%) = 0 ' Remise à zéro des points
        Voie(No%).pZ!(i%) = 0
        For j% = 0 To 3
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
' **********************************
' Calcul le point d'arrivée B,C ou D
' et son angle à l'extremité
' **********************************
'
Public Sub Calcule_Point(No%, o%, p%)
    Dim n%
    n% = Voie(No%).MatConnecte(o%, p%) - 1
    If Voie(No%).pX!(p%) = 0 And Voie(No%).pZ!(p%) = 0 Then
        If Voie(No%).Offset(o%) = 0 And o% <> 0 Then
            Voie(No%).pX!(p%) = Voie(No%).pX!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).X
            Voie(No%).pZ!(p%) = Voie(No%).pZ!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).z
        Else
            Voie(No%).pX!(p%) = Voie(No%).dX!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).X
            Voie(No%).pZ!(p%) = Voie(No%).dz!(o%) + Voie(No%).Segment(n%).Point(Voie(No%).Segment(n%).Segment_Taille()).z
            '
            ' ***** Tourne le segment pour avoir la normale dans le sens inverse de l'entrée
            '
            Voie(No%).Segment(n%).Rotation! = Voie(No%).Segment(n%).Rotation! + 180
            Voie(No%).Normal(o%) = Voie(No%).Segment(n%).Tangente(0)
            Voie(No%).AngleNormal(o%) = Voie(No%).Segment(n%).Theta(0)
            Voie(No%).Segment(n%).Rotation! = Voie(No%).Segment(n%).Rotation! - 180
        End If
        Voie(No%).Normal(p%) = Voie(No%).Segment(n%).Tangente(Voie(No%).Segment(n%).Segment_Taille())
        Voie(No%).AngleNormal(p%) = Voie(No%).Segment(n%).Theta(Voie(No%).Segment(n%).Segment_Taille())
    End If
End Sub

'
' ******************************
' Charge la définition des voies
' ******************************
'
Public Sub Charger(Fichier$)
    Dim n%, i%, j%, a!
    Dim nbs%, f%
    
    f% = FreeFile()
    Call Tools.Open_File(Fichier$, f%, OPEN_NORMAL)
    Input #f%, nbs%
    ReDim Voie(nbs%) As TypeVoie
    'ReDim FormeVoie As New ClassTunnel
    For n% = 1 To nbs%
        With Voie(n%)
            Input #f%, .Libelle
            Input #f%, .Ref
            For i% = 0 To 3
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
                
                For j% = 0 To 3
                    If i% <> j% Then
                        Input #f%, a!: .MatConnecte(i%, j%) = a!
                        Input #f%, a!: .MatAiguille(i%, j%) = a!
                    End If
                Next j%
            Next i%
        End With
    Next n%
    Close #f%
End Sub

'
' **********************************
' Test si deux segments sont secants
' Module Horizontal
' **********************************
'
Public Function SecanteH(Ax1!, Ay1!, Ax2!, Ay2!, bX1%, bX2%, By%) As Boolean
    Dim X!
    Dim a!
    Dim b!
    
    If Ay1! = Ay2! Then
        If By% = Ay1! Then
            If Ax1! <= bX1% And Ax2! >= bX2% Then
                SecanteH = True
            Else
                SecanteH = False
            End If
        Else
            SecanteH = False
        End If
    Else
        '
        ' Cas général
        '
        If Ax1! = Ax2! Then ' Segment vertical
            If bX1% < Ax1! And Ax1! < bX2% Then
                If Ay1! < Ay2! Then
                    If Ay1! < By% And By% < Ay2! Then
                        SecanteH = True
                    Else
                        SecanteH = False
                    End If
                Else
                    If Ay2! < By% And By% < Ay1! Then
                        SecanteH = True
                    Else
                        SecanteH = False
                    End If
                End If
            Else
                SecanteH = False
            End If
        Else
            a! = (Ay2! - Ay1!) / (Ax2! - Ax1!)
            b! = Ay1! - a! * Ax1!
            X! = (By% - b!) / a!
            If bX1% < X! And X! < bX2% Then
                SecanteH = True
            Else
                SecanteH = False
            End If
        End If
    End If
End Function

'
' **********************************
' Test si deux segments sont secants
' Module Vertical
' **********************************
'
Public Function SecanteV(Ax1!, Ay1!, Ax2!, Ay2!, By1%, By2%, Bx%) As Boolean
    Dim Y!
    Dim a!
    Dim b!
    
    If Ax1! = Ax2! Then
        If Bx% = Ax1! Then
            If Ay1! <= By1% And Ay2! >= By2% Then
                SecanteV = True
            Else
                SecanteV = False
            End If
        Else
            SecanteV = False
        End If
    Else
        '
        ' Cas général
        '
        If Ay1! = Ay2! Then ' Segment horizontal
            If By1% < Ay1! And Ay1! < By2% Then
                If Ax1! < Ax2! Then
                    If Ax1! < Bx% And Bx% < Ax2! Then
                        SecanteV = True
                    Else
                        SecanteV = False
                    End If
                Else
                    If Ax2! < Bx% And Bx% < Ax1! Then
                        SecanteV = True
                    Else
                        SecanteV = False
                    End If
                End If
            Else
                SecanteV = False
            End If
        Else
            a! = (Ay2! - Ay1!) / (Ax2! - Ax1!)
            b! = Ay1! - a! * Ax1!
            Y! = a! * Bx% + b!
            If By1% < Y! And Y! < By2% Then
                SecanteV = True
            Else
                SecanteV = False
            End If
        End If
    End If
End Function

