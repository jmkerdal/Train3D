VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTunnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim LRectangle() As D3DVECTOR
Public NbLRectangle%
Dim LPoints() As D3DVECTOR
Dim LSegments%()
Public NbLSegment%

Dim NPoints() As D3DVECTOR
Dim NSegments%()
Public NbNSegment%
Public NbContour% ' Nombre de segment du contour

Dim PI#
Dim DegRad As Single
Const PAS% = 50 ' Pas pour la création des cotés

'
' ********************************************
' Génération du Mesh à partir des informations
' ********************************************
'
Public Function Cree_Mesh() '(dxEngine As ClassDirectX) As Direct3DRMMeshBuilder3
    'Dim i%, c&
    'Dim Texture(0) As Direct3DRMTexture3
    'Dim p() As D3DVECTOR
    'Dim n%
    'Dim pt(3) As D3DVECTOR
    
    'ReDim p(NbContour%) As D3DVECTOR
    'Call dxConstruitMeshBuilder.Init
    'n% = 0
    'For i% = 1 To NbNSegment%
    '    If NSegments%(2, i%) <> -1 Then
    '        n% = n% + 1
    '        p(n%) = NPoints(NSegments%(0, i%))
    '        p(n%).y = 30
    '        '
    '        pt(0) = NPoints(NSegments%(1, i%))
    '        pt(3) = pt(0)
    '        pt(0).y = 30
    '        '
    '        pt(1) = NPoints(NSegments%(0, i%))
    '        pt(1).y = 30
    '        pt(2) = pt(1): pt(2).y = 0
    '        If NSegments%(2, i%) = 0 Then
    '            c& = &HFF0000FF
    '        Else
    '            c& = &HFFFF00FF
    '            Call dxConstruitMeshBuilder.Add_Rectangle(pt(0), pt(1), pt(2), pt(3), 0, c&)
    '        End If
    '        'Call dxConstruitMeshBuilder.Add_Rectangle(pt(0), pt(1), pt(2), pt(3), 0, c&)
    '    End If
    'Next i%
    ''Call dxConstruitMeshBuilder.Add_Polygon(p(), 0, &HFFFFFFFF)
    
    'For i% = 1 To NbLRectangle%
    '    c& = dxEngine.dX7.CreateColorRGBA(Rnd(), Rnd(), Rnd(), 1)
    '    Call dxConstruitMeshBuilder.Add_Rectangle(LRectangle(0, i%), LRectangle(1, i%), LRectangle(2, i%), LRectangle(3, i%), 0, c&)
    'Next i%
    
    'Set Cree_Mesh = dxEngine.dxD3Drm.CreateMeshBuilder
    'Call dxConstruitMeshBuilder.Build(Cree_Mesh, Texture())
    ''Call Sauve.Save("Text" + Format$(No%) + ".x", D3DRMXOF_TEXT, D3DRMXOFSAVE_ALL)
    ''Set Sauve = Nothing
End Function

'
' ************************************************
' Ajoute un point dans la liste et rend son numero
' , le test se fait juste sur le x et le z
' ************************************************
'
Private Function Ajoute_Point%(TempPoint() As D3DVECTOR, Valeur As D3DVECTOR)
    Dim i%
    If UBound(TempPoint()) <> 0 Then
        For i% = 1 To UBound(TempPoint())
            If (TempPoint(i%).X - Valeur.X) ^ 2 + (TempPoint(i%).z - Valeur.z) ^ 2 < 1 Then
                Ajoute_Point% = i%
                Exit Function
            End If
            'If TempPoint(i%).X = Valeur.X Then
            '    If TempPoint(i%).z = Valeur.z Then
            '        Ajoute_Point% = i%
            '        Exit Function
            '    End If
            'End If
        Next i%
    End If
    Ajoute_Point% = UBound(TempPoint()) + 1
    ReDim Preserve TempPoint(Ajoute_Point%)
    TempPoint(Ajoute_Point%) = Valeur
End Function

'
' ************************************
' Retourne l'angle fait par le segment
' ************************************
'
Private Function Calcul_Angle%(x1!, z1!, x2!, z2!)
    Dim Tangente#
    If x1! = x2! Then
        If z2! > z1! Then
            Calcul_Angle% = 90
        Else
            Calcul_Angle% = 270
        End If
    Else
        Tangente# = (z2! - z1!) / (x2! - x1!)
        Calcul_Angle% = Atn(Tangente#) * 180 / PI#
        If (x2! - x1!) > 0 Then
            If (z2! - z1!) < 0 Then
                Calcul_Angle% = 360 + Calcul_Angle%
            End If
        Else
            Calcul_Angle% = 180 + Calcul_Angle%
        End If
    End If
End Function

'
' ************************************************
' Crée le contour d'une voie à partir des segments
' ************************************************
'
Public Function Cree_Contour() As Boolean
    Dim i%, j%
    Dim Trouve As Boolean
    Dim P1%, P2%

    Cree_Contour = True ' Par défaut
    '
    ' ***** Elimine les doubles segments
    '
    For i% = 1 To NbLSegment%
        If LSegments%(2, i%) <> -1 Then
            For j% = i% + 1 To NbLSegment%
                If LSegments%(0, i%) = LSegments%(0, j%) And LSegments%(1, i%) = LSegments%(1, j%) Then
                    LSegments%(2, j%) = -1
                End If
                If LSegments%(0, i%) = LSegments%(1, j%) And LSegments%(1, i%) = LSegments%(0, j%) Then
                    LSegments%(2, j%) = -1
                End If
            Next j%
        End If
    Next i%
    '
    ' ***** Calcul des équations de droite
    '
    Dim dX#, dz#
    Dim Coef#()
    ReDim Coef#(2, NbLSegment%)
    For i% = 1 To NbLSegment%
        If LSegments%(2, i%) = 1 Then
            dX# = LPoints(LSegments%(0, i%)).X - LPoints(LSegments%(1, i%)).X
            dz# = LPoints(LSegments%(0, i%)).z - LPoints(LSegments%(1, i%)).z
            If dX# = 0 Then
                Coef#(0, i%) = LPoints(LSegments%(0, i%)).X
                Coef#(1, i%) = dz#
                Coef#(2, i%) = 1
            Else
                Coef#(0, i%) = dz# / dX#
                Coef#(1, i%) = LPoints(LSegments%(0, i%)).z - Coef#(0, i%) * LPoints(LSegments%(0, i%)).X
                Coef#(2, i%) = 0
            End If
        End If
    Next i%
    '
    ' ***** Recherche des segments sécants
    '
    Dim NouveauPoint() As D3DVECTOR
    Dim NbNouveauPoint%
    Dim LePoint As D3DVECTOR
    Dim BoutCoupe%() ' Indique qu'on est coupé par un bout
    For i% = 1 To NbLSegment%
        If LSegments%(2, i%) = 1 Then
            NbNouveauPoint% = 0
            ReDim NouveauPoint(NbNouveauPoint%)
            ReDim BoutCoupe%(NbNouveauPoint%)
            Trouve = False
            For j% = 1 To NbLSegment%
                If i% <> j% And LSegments%(2, j%) = 1 Then
                    If i% < j% Then
                        P1% = i%
                        P2% = j%
                    Else
                        P1% = j%
                        P2% = i%
                    End If
                    If Secante(Coef#(), P1%, P2%, LPoints(), LSegments%(), LePoint) = True Then
                        'Debug.Print "Coupe"; i%; j%; LePoint.x; LePoint.z
                        NbNouveauPoint% = NbNouveauPoint% + 1
                        ReDim Preserve NouveauPoint(NbNouveauPoint%)
                        ReDim Preserve BoutCoupe%(NbNouveauPoint%)
                        NouveauPoint(NbNouveauPoint%) = LePoint
                        If LSegments%(3, j%) = 1 Then
                            BoutCoupe%(NbNouveauPoint%) = 1
                        End If
                    End If
                End If
            Next j%
            '
            ' ***** Trie à bulle
            '
            For P1% = 1 To NbNouveauPoint% - 1
                For P2% = P1% + 1 To NbNouveauPoint%
                    If LPoints(LSegments%(0, i%)).X = LPoints(LSegments%(1, i%)).X Then
                        If LPoints(LSegments%(0, i%)).z < LPoints(LSegments%(1, i%)).z Then
                            If NouveauPoint(P1%).z > NouveauPoint(P2%).z Then
                                LePoint = NouveauPoint(P1%)
                                NouveauPoint(P1%) = NouveauPoint(P2%)
                                NouveauPoint(P2%) = LePoint
                                j% = BoutCoupe%(P1%)
                                BoutCoupe%(P1%) = BoutCoupe%(P2%)
                                BoutCoupe%(P2%) = j%
                            End If
                        Else
                            If NouveauPoint(P1%).z < NouveauPoint(P2%).z Then
                                LePoint = NouveauPoint(P1%)
                                NouveauPoint(P1%) = NouveauPoint(P2%)
                                NouveauPoint(P2%) = LePoint
                                j% = BoutCoupe%(P1%)
                                BoutCoupe%(P1%) = BoutCoupe%(P2%)
                                BoutCoupe%(P2%) = j%
                            End If
                        End If
                    ElseIf LPoints(LSegments%(0, i%)).X < LPoints(LSegments%(1, i%)).X Then
                        If NouveauPoint(P1%).X > NouveauPoint(P2%).X Then
                            LePoint = NouveauPoint(P1%)
                            NouveauPoint(P1%) = NouveauPoint(P2%)
                            NouveauPoint(P2%) = LePoint
                            j% = BoutCoupe%(P1%)
                            BoutCoupe%(P1%) = BoutCoupe%(P2%)
                            BoutCoupe%(P2%) = j%
                        End If
                    Else
                        If NouveauPoint(P1%).X < NouveauPoint(P2%).X Then
                            LePoint = NouveauPoint(P1%)
                            NouveauPoint(P1%) = NouveauPoint(P2%)
                            NouveauPoint(P2%) = LePoint
                            j% = BoutCoupe%(P1%)
                            BoutCoupe%(P1%) = BoutCoupe%(P2%)
                            BoutCoupe%(P2%) = j%
                        End If
                    End If
                Next P2%
            Next P1%
            '
            ' ***** Ajoute entre 1 et N segments
            '
            For j% = 0 To NbNouveauPoint%
                If j% = 0 Then
                    P1% = Ajoute_Point(NPoints(), LPoints(LSegments%(0, i%)))
                Else
                    P1% = Ajoute_Point(NPoints(), NouveauPoint(j%))
                End If
                If j% = NbNouveauPoint% Then
                    P2% = Ajoute_Point(NPoints(), LPoints(LSegments%(1, i%)))
                Else
                    P2% = Ajoute_Point(NPoints(), NouveauPoint(j% + 1))
                End If
                NbNSegment% = NbNSegment% + 1
                ReDim Preserve NSegments%(4, NbNSegment%)
                NSegments%(0, NbNSegment%) = P1%
                NSegments%(1, NbNSegment%) = P2%
                NSegments%(2, NbNSegment%) = 1 ' A Afficher
                NSegments%(3, NbNSegment%) = LSegments%(3, i%) ' Mémorise le bout ou coté
                If NSegments%(3, NbNSegment%) = 0 Then
                    'If j% <> 0 Then
                        NSegments%(4, NbNSegment%) = BoutCoupe%(j%) 'LSegments%(3, i%) ' Mémorise le bout ou coté
                    'End If
                End If
'Debug.Print "Segment"; i%; j%; Format$(Coef#(0, i%), "##.## ");
'Debug.Print p1%; "=["; NPoints(p1%).x; NPoints(p1%).z; "]";
'Debug.Print p2%; "=["; NPoints(p2%).x; NPoints(p2%).z; "]"
            Next j%
        End If
    Next i%
    '
    ' ***** Génération du contour (enfin!)
    '
    Dim n%, s%
    
    For j% = 1 To UBound(NPoints())
        n% = 0
        s% = 0
        For i% = 1 To NbNSegment%
            If NSegments%(0, i%) = j% Or NSegments%(1, i%) = j% Then
                n% = n% + 1
                s% = i%
            End If
        Next i%
        If n% = 2 Then
            Cree_Contour = Genere_Contour(s%)
'Debug.Print j%; Cree_Contour
            If Cree_Contour = True Then Exit Function
        End If
    Next j%
    
    'For i% = 1 To NbNSegment%
    '    Cree_Contour = Genere_Contour(i%)
    '    If Cree_Contour = True Then Exit Function
    'Next i%
    'Cree_Contour = Genere_Contour(1)
End Function

'
' **********************************
' Crée les segments de départ à
' partir des informations de la voie
' **********************************
'
Public Sub Cree_Couche(No%, Largeur!, Pos As D3DVECTOR, Dir!)
    Dim i%, j%, k%, l%
    Dim n%, t!, X!, z!
    Dim nTraverse%, dTraverse!
    Dim v As D3DVECTOR
    Dim vt As Single
    Dim AxeY As D3DVECTOR
    Dim dX7 As New DirectX7
    Dim module!
    Dim NbLPoint%
    
    Dim p(1) As D3DVECTOR
    Dim NoP%(3)
    
    AxeY.X = 0
    AxeY.Y = 1
    AxeY.z = 0
    '
    ' ***** Calcul des segments
    '
    Call Recalcule_Point(No%)
    NbLPoint% = 0
    'NbLSegment% = 0
    'NbLRectangle% = 0
    '
    For i% = 0 To 2
        For j% = i% + 1 To 3
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
                If Voie(No%).Segment(n%).Segment_Droit = True Then
                    nTraverse% = 0
                Else
                    nTraverse% = t! / PAS%
                End If
                dTraverse! = t! / (nTraverse% + 1)
                For k% = 0 To nTraverse% + 1
                    v = Voie(No%).Segment(n%).Point(k% * dTraverse!)
                    v.X = v.X + X!
                    v.z = v.z + z!
                    vt! = Voie(No%).Segment(n%).Theta(k% * dTraverse!) * DegRad!
                    
                    p(0).X = 0
                    p(0).z = Largeur!
                    p(1).X = 0
                    p(1).z = -Largeur!
                    
                    For l% = 0 To 1
                        If vt! <> 0 Then
                            module! = dX7.VectorModulus(p(l%))
                            Call dX7.VectorRotate(p(l%), p(l%), AxeY, vt!)
                            Call dX7.VectorScale(p(l%), p(l%), module!)
                        End If
                        Call dX7.VectorAdd(p(l%), p(l%), v)
                        '
                        ' ****** Ajoute l'offset, utilisé pour la génération des tunnels
                        ' ****** avec plusieurs voies mises bout a bout
                        '
                        module! = dX7.VectorModulus(p(l%))
                        Call dX7.VectorRotate(p(l%), p(l%), AxeY, Dir! * DegRad!)
                        Call dX7.VectorScale(p(l%), p(l%), module!)
                        Call dX7.VectorAdd(p(l%), p(l%), Pos)
                    Next l%
                    
                    NoP%(0) = NoP%(2) ' Mémorisation de la valeur précédente
                    NoP%(1) = NoP%(3)
                    NoP%(2) = Ajoute_Point(LPoints(), p(0))
                    NoP%(3) = Ajoute_Point(LPoints(), p(1))
                    
                    If k% = 0 Then
                        NbLSegment% = NbLSegment% + 1
                        ReDim Preserve LSegments(3, NbLSegment%) As Integer
                        LSegments%(0, NbLSegment%) = NoP%(3)
                        LSegments%(1, NbLSegment%) = NoP%(2)
                        LSegments%(2, NbLSegment%) = 1 ' Bout a tracer
                        LSegments%(3, NbLSegment%) = 1 ' Bout
                    Else
                        NbLSegment% = NbLSegment% + 2
                        ReDim Preserve LSegments(3, NbLSegment%) As Integer
                        LSegments%(0, NbLSegment% - 1) = NoP%(0)
                        LSegments%(1, NbLSegment% - 1) = NoP%(2)
                        LSegments%(0, NbLSegment%) = NoP%(3)
                        LSegments%(1, NbLSegment%) = NoP%(1)
                        LSegments%(2, NbLSegment% - 1) = 1 ' A tracer
                        LSegments%(2, NbLSegment%) = 1 ' A tracer
                        LSegments%(3, NbLSegment% - 1) = 0 ' coté
                        LSegments%(3, NbLSegment%) = 0 ' coté

                        NbLRectangle% = NbLRectangle% + 1
                        ReDim Preserve LRectangle(3, NbLRectangle%)
                        LRectangle(0, NbLRectangle%) = LPoints(NoP%(0))
                        LRectangle(1, NbLRectangle%) = LPoints(NoP%(2))
                        LRectangle(2, NbLRectangle%) = LPoints(NoP%(3))
                        LRectangle(3, NbLRectangle%) = LPoints(NoP%(1))
                        LRectangle(0, NbLRectangle%).Y = 30
                        LRectangle(1, NbLRectangle%).Y = 30
                        LRectangle(2, NbLRectangle%).Y = 30
                        LRectangle(3, NbLRectangle%).Y = 30

                    End If
                Next k%
                
                NbLSegment% = NbLSegment% + 1
                ReDim Preserve LSegments(3, NbLSegment%) As Integer
                LSegments%(0, NbLSegment%) = NoP%(2)
                LSegments%(1, NbLSegment%) = NoP%(3)
                LSegments%(2, NbLSegment%) = 1 ' Bout a tracer
                LSegments%(3, NbLSegment%) = 1 ' Bout

            End If
        Next j%
    Next i%
    Set dX7 = Nothing

    NbLPoint% = UBound(LPoints())
End Sub

'
' **************************************
' Test si les deux segments sont sécants
' **************************************
'
Private Function Secante(c#(), S1%, S2%, TempPoint() As D3DVECTOR, TempSegment%(), LePoint As D3DVECTOR) As Boolean

    Dim bX1!, bZ1!
    Dim bX2!, bZ2!
    Dim bX3!, bZ3!
    Dim bX4!, bZ4!

    '
    ' ***** Test les extremités
    '
    If Abs(TempPoint(TempSegment(0, S1%)).X - TempPoint(TempSegment(0, S2%)).X) < 0.5 _
    And Abs(TempPoint(TempSegment(0, S1%)).z - TempPoint(TempSegment(0, S2%)).z) < 0.5 _
    Then
        Exit Function
    End If
    If Abs(TempPoint(TempSegment(0, S1%)).X - TempPoint(TempSegment(1, S2%)).X) < 0.5 _
    And Abs(TempPoint(TempSegment(0, S1%)).z - TempPoint(TempSegment(1, S2%)).z) < 0.5 _
    Then
        Exit Function
    End If
    If Abs(TempPoint(TempSegment(1, S1%)).X - TempPoint(TempSegment(0, S2%)).X) < 0.5 _
    And Abs(TempPoint(TempSegment(1, S1%)).z - TempPoint(TempSegment(0, S2%)).z) < 0.5 _
    Then
        Exit Function
    End If
    If Abs(TempPoint(TempSegment(1, S1%)).X - TempPoint(TempSegment(1, S2%)).X) < 0.5 _
    And Abs(TempPoint(TempSegment(1, S1%)).z - TempPoint(TempSegment(1, S2%)).z) < 0.5 _
    Then
        Exit Function
    End If
    '
    If TempPoint(TempSegment(0, S1%)).X < TempPoint(TempSegment(1, S1%)).X Then
        bX1! = TempPoint(TempSegment(0, S1%)).X
        bX2! = TempPoint(TempSegment(1, S1%)).X
    Else
        bX1! = TempPoint(TempSegment(1, S1%)).X
        bX2! = TempPoint(TempSegment(0, S1%)).X
    End If
    If TempPoint(TempSegment(0, S1%)).z < TempPoint(TempSegment(1, S1%)).z Then
        bZ1! = TempPoint(TempSegment(0, S1%)).z
        bZ2! = TempPoint(TempSegment(1, S1%)).z
    Else
        bZ1! = TempPoint(TempSegment(1, S1%)).z
        bZ2! = TempPoint(TempSegment(0, S1%)).z
    End If
    If TempPoint(TempSegment(0, S2%)).X < TempPoint(TempSegment(1, S2%)).X Then
        bX3! = TempPoint(TempSegment(0, S2%)).X
        bX4! = TempPoint(TempSegment(1, S2%)).X
    Else
        bX3! = TempPoint(TempSegment(1, S2%)).X
        bX4! = TempPoint(TempSegment(0, S2%)).X
    End If
    If TempPoint(TempSegment(0, S2%)).z < TempPoint(TempSegment(1, S2%)).z Then
        bZ3! = TempPoint(TempSegment(0, S2%)).z
        bZ4! = TempPoint(TempSegment(1, S2%)).z
    Else
        bZ3! = TempPoint(TempSegment(1, S2%)).z
        bZ4! = TempPoint(TempSegment(0, S2%)).z
    End If

    If c#(2, S1%) = 1 Or c#(2, S2%) = 1 Then
        ' ***** Si une droite est verticale
        If c#(2, S1%) = 1 And c#(2, S2%) = 1 Then
            ' ***** Parallèle?
            Exit Function
        ElseIf c#(2, S1%) = 1 Then
            LePoint.z = c#(0, S2%) * c#(0, S1%) + c#(1, S2%)
            LePoint.X = c#(0, S1%)
            
            ' ***** Sécante?
            'If (LePoint.X - bX3!) > 0.001 Or (LePoint.X - bX4!) > 0.001 Then ' Cas des lignes horizontales
            If LePoint.X < bX3! Then Exit Function
            If LePoint.X > bX4! Then Exit Function
            'End If
            If LePoint.z <= bZ1! Then Exit Function
            If LePoint.z >= bZ2! Then Exit Function
            If LePoint.z <= bZ3! Then Exit Function
            If LePoint.z >= bZ4! Then Exit Function
            Secante = True
                        
        Else
            LePoint.z = c#(0, S1%) * c#(0, S2%) + c#(1, S1%)
            LePoint.X = c#(0, S2%)
            
            ' ***** Sécante?
            'If (LePoint.X - bX1!) > 0.001 Or (LePoint.X - bX2!) > 0.001 Then ' Cas des lignes horizontales
            If LePoint.X < bX1! Then Exit Function
            If LePoint.X > bX2! Then Exit Function
            'End If
            If LePoint.z <= bZ1! Then Exit Function
            If LePoint.z >= bZ2! Then Exit Function
            If LePoint.z <= bZ3! Then Exit Function
            If LePoint.z >= bZ4! Then Exit Function
            Secante = True
            
        End If
    Else
        ' ***** Parallèle?
        If c#(0, S1%) = c#(0, S2%) Then Exit Function
        
        LePoint.X = (c#(1, S2%) - c#(1, S1%)) / (c#(0, S1%) - c#(0, S2%))
        LePoint.z = c#(0, S1%) * LePoint.X + c#(1, S1%)
        
'If s1% = 6 And s2% = 2 Then
'Debug.Print s1%; s2%; Format$(LePoint.x, "##.## "); Format$(LePoint.z, "##.## ");
'Debug.Print "["; bX1!; ","; bZ1!; "]["; bX2!; ","; bZ2!; "]";
'Debug.Print ":["; bX3!; ","; bZ3!; "]["; bX4!; ","; bZ4!; "]"
'End If
        ' ***** Sécante?
        If LePoint.X <= bX1! Then Exit Function
        If LePoint.X >= bX2! Then Exit Function
        If LePoint.X <= bX3! Then Exit Function
        If LePoint.X >= bX4! Then Exit Function
        If (LePoint.z - bZ1!) > 0.001 Or (LePoint.z - bZ2!) > 0.001 Then ' Cas des lignes horizontales
            If LePoint.z <= bZ1! Then Exit Function
            If LePoint.z >= bZ2! Then Exit Function
        End If
        If (LePoint.z - bZ3!) > 0.001 Or (LePoint.z - bZ4!) > 0.001 Then ' Cas des lignes horizontales
            If LePoint.z <= bZ3! Then Exit Function
            If LePoint.z >= bZ4! Then Exit Function
        End If
        Secante = True
    End If
End Function

'
' ********************************
' Initialise les valeurs de départ
' ********************************
'
Private Sub Class_Initialize()
    Call Raz
    PI# = 4 * Atn(1)
    DegRad! = Atn(1) / 45
End Sub

'
' *********************
' Lecture liste segment
' *********************
'
Public Property Get Liste_Segment%(a%, n%)
    Liste_Segment% = LSegments%(a%, n%)
End Property

'
' *******************
' Lecture liste point
' *******************
'
Public Property Get Liste_Point(n%) As D3DVECTOR
    Liste_Point = LPoints(n%)
End Property

'
' *********************
' Lecture nouveau point
' *********************
'
Public Property Get Nouveau_Point(n%) As D3DVECTOR
    Nouveau_Point = NPoints(n%)
End Property

'
' ***********************
' Lecture nouveau segment
' ***********************
'
Public Property Get Nouveau_Segment%(a%, n%)
    Nouveau_Segment = NSegments%(a%, n%)
End Property

'
' ********************************
' Génération de la forme du tunnel
' ********************************
'
Public Sub Genere(n%)
    'Call Cree_Couche(n%, 13)
    'Call Cree_Contour
    'Call Pack
End Sub

'
' *******************************************
' Réduction de la liste des segments
' Réduction de la liste des nouveaux segments
' *******************************************
'
Public Sub Pack()

End Sub

'
' ********************************
' Récupère la liste des rectangles
' ********************************
'
Public Sub Liste_Rectangle(n%, P1 As D3DVECTOR, P2 As D3DVECTOR, p3 As D3DVECTOR, p4 As D3DVECTOR)
     P1 = LRectangle(0, n%)
     P2 = LRectangle(1, n%)
     p3 = LRectangle(2, n%)
     p4 = LRectangle(3, n%)
End Sub

'
' ************************************
' Ajoute un rectangle dans la liste
' Utilisé par le générateur de contour
' ************************************
'
Public Sub Ajoute_Rectangle(P1 As D3DVECTOR, P2 As D3DVECTOR, p3 As D3DVECTOR, p4 As D3DVECTOR)
    NbLRectangle% = NbLRectangle% + 1
    ReDim Preserve LRectangle(3, NbLRectangle%)
    LRectangle(0, NbLRectangle%) = P1
    LRectangle(1, NbLRectangle%) = p3
    LRectangle(2, NbLRectangle%) = p4
    LRectangle(3, NbLRectangle%) = P2
End Sub

'
' ******************
' Efface les données
' ******************
'
Public Sub Raz()
    ReDim LPoints(0) As D3DVECTOR
    NbLSegment% = 0
    ReDim LSegments%(3, NbLSegment%)
    
    ReDim NPoints(0) As D3DVECTOR
    NbNSegment% = 0
    ReDim NSegments%(4, NbNSegment%)
    NbContour% = 0
    
    NbLRectangle% = 0
    ReDim LRectangle(3, NbLRectangle%) As D3DVECTOR
End Sub

'
' ************************************
' Tente de générer un contour à partir
' de la position de départ
' ************************************
'
Private Function Genere_Contour(Depart%)
    Dim i%, j%
    Dim AngleSegment%()
    Dim NbAngleSegment%
    Dim AngleActuel% ' Angle du segment en cours
    Dim S1%, P1%, P2%
    Dim n% ' Nombre d'itération pour sortie sur erreur

    For i% = 1 To NbNSegment%
        NSegments%(2, i%) = -1 ' NSegments%(2, i%) ' Par défaut non utilisé
    Next i%
    S1% = Depart% ' Segment de départ
    NSegments%(2, S1%) = 1 ' -NSegments%(2, S1%)
    P1% = NSegments%(0, S1%)
    P2% = NSegments%(1, S1%) ' Point de départ
    Do
        AngleActuel% = Calcul_Angle%(NPoints(NSegments%(0, S1%)).X, NPoints(NSegments%(0, S1%)).z, NPoints(NSegments%(1, S1%)).X, NPoints(NSegments%(1, S1%)).z)
'Debug.Print "Départ"; S1%; P2%; AngleActuel%
        NbAngleSegment% = 0
        ReDim AngleSegment%(2, NbAngleSegment%)
        For i% = 1 To NbNSegment%
            If NSegments%(2, i%) = -1 Then
                '
                ' ***** Test a l'endroit
                '
                If NSegments%(0, i%) = P2% Then
                    NbAngleSegment% = NbAngleSegment% + 1
                    ReDim Preserve AngleSegment%(2, NbAngleSegment%)
                    j% = Calcul_Angle%(NPoints(NSegments%(0, i%)).X, NPoints(NSegments%(0, i%)).z, NPoints(NSegments%(1, i%)).X, NPoints(NSegments%(1, i%)).z)
'Debug.Print "0"; j%
                    j% = -j% + AngleActuel% + 180
                    If j% > 360 Then j% = j% - 360
                    If j% < 0 Then j% = j% + 360
                    AngleSegment%(0, NbAngleSegment%) = i%
                    AngleSegment%(1, NbAngleSegment%) = j%
                    AngleSegment%(2, NbAngleSegment%) = 1
'Debug.Print NbAngleSegment%; AngleSegment%(0, NbAngleSegment%); AngleSegment%(1, NbAngleSegment%)
                End If
                '
                ' ***** Test a l'envers
                '
                'If NSegments%(1, i%) = P2% Then
                '    NbAngleSegment% = NbAngleSegment% + 1
                '    ReDim Preserve AngleSegment%(2, NbAngleSegment%)
                '    j% = Calcul_Angle%(NPoints(NSegments%(1, i%)).X, NPoints(NSegments%(1, i%)).z, NPoints(NSegments%(0, i%)).X, NPoints(NSegments%(0, i%)).z)
'Debug.Print "1"; j%
                '    j% = -j% + AngleActuel% + 180
                '    If j% > 360 Then j% = j% - 360
                '    If j% < 0 Then j% = j% + 360
                '    AngleSegment%(0, NbAngleSegment%) = i%
                '    AngleSegment%(1, NbAngleSegment%) = j%
                '    AngleSegment%(2, NbAngleSegment%) = 0
'Debug.Print NbAngleSegment%; AngleSegment%(0, NbAngleSegment%); AngleSegment%(1, NbAngleSegment%)
                'End If
            End If
        Next i%
        '
        If NbAngleSegment% > 0 Then
            j% = 1
            If NbAngleSegment% > 1 Then
                For i% = 2 To NbAngleSegment%
                    If AngleSegment(1, i%) < AngleSegment%(1, j%) Then j% = i%
                Next i%
            End If
            S1% = AngleSegment%(0, j%)
            P2% = NSegments%(AngleSegment%(2, j%), S1%) ' Nouveau point
        Else
'Debug.Print "Pas de bout!"
        End If
        '
        NSegments%(2, S1%) = 1 ' -NSegments%(2, S1%)
'Debug.Print NSegments%(0, S1%); NSegments%(1, S1%)
        If P2% = P1% Then Exit Do
        'If S1% = Depart% Then Exit Do ' Fin du parcours
        DoEvents

        ' En cas de boucle infinie on force la sortie
        n% = n% + 1
        If n% > NbNSegment% Then
Debug.Print "Sortie sur erreur"
            Genere_Contour = False ' Sortie sur erreur
            Exit Function
        End If
    Loop
    
    NbContour% = 0
    For i% = 1 To NbNSegment%
        If NSegments%(2, i%) <> -1 Then NbContour% = NbContour% + 1
    Next i%
    Genere_Contour = True
End Function

