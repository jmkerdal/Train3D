VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCatenaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Dim LRectangle() As D3DVECTOR
Public NbLRectangle%
Dim LPoints() As D3DVECTOR
Dim LSegments() As Integer
Public NbLSegment%
'
Dim NPoints() As D3DVECTOR
Dim NSegments() As Integer
Public NbNSegment%
Public NbContour% ' Nombre de segment du contour
'
Dim PI#
Dim DegRad As Single
Const PAS% = 50 ' Pas pour la création des cotés

'
' ********************************************
' Génération du Mesh à partir des informations
' ********************************************
'
Public Function Cree_Mesh(dxEngine As ClassDirectX) As Direct3DRMMeshBuilder3
    Dim i%, c&
    Dim Texture(0) As Direct3DRMTexture3
    Dim P() As D3DVECTOR
    Dim n%
    Dim pt(3) As D3DVECTOR
    '
    ReDim P(NbContour%) As D3DVECTOR
    Call dxConstruitMeshBuilder.Init
    n% = 0
    For i% = 1 To NbNSegment%
        If NSegments(2, i%) <> -1 Then
            n% = n% + 1
            P(n%) = NPoints(NSegments(0, i%))
            P(n%).Y = 30
            '
            pt(0) = NPoints(NSegments(1, i%))
            pt(3) = pt(0)
            pt(0).Y = 30
            '
            pt(1) = NPoints(NSegments(0, i%))
            pt(1).Y = 30
            pt(2) = pt(1): pt(2).Y = 0
            If NSegments(2, i%) <> 0 Then
                c& = &HFFFF00FF
                Call dxConstruitMeshBuilder.Add_Rectangle(pt(0), pt(1), pt(2), pt(3), 0, c&)
            End If
        End If
    Next i%
    '
    For i% = 1 To NbLRectangle%
        c& = dxEngine.dX7.CreateColorRGBA(Rnd(), Rnd(), Rnd(), 1)
        Call dxConstruitMeshBuilder.Add_Rectangle(LRectangle(0, i%), LRectangle(1, i%), LRectangle(2, i%), LRectangle(3, i%), 0, c&)
    Next i%
    '
    Set Cree_Mesh = dxEngine.dxD3Drm.CreateMeshBuilder
    Call dxConstruitMeshBuilder.Build(Cree_Mesh, Texture())
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
            If TempPoint(i%).X = Valeur.X Then
                If TempPoint(i%).z = Valeur.z Then
                    Ajoute_Point% = i%
                    Exit Function
                End If
            End If
        Next i%
    End If
    Ajoute_Point% = UBound(TempPoint()) + 1
    ReDim Preserve TempPoint(Ajoute_Point%)
    TempPoint(Ajoute_Point%) = Valeur
End Function

'
' ************************************************
' Crée le contour d'une voie à partir des segments
' ************************************************
'
Public Sub Cree_Contour()
    Dim i%, j%
    Dim Trouve As Boolean
    Dim P1%, P2%
    '
    ' ***** Elimine les bouts
    '
    For i% = 1 To NbLSegment%
        If LSegments%(2, i%) = 0 Then
            Trouve = False
            For j% = i% + 1 To NbLSegment%
                If LSegments%(2, j%) = 0 Then
                    If LPoints(LSegments%(0, i%)).X = LPoints(LSegments%(0, j%)).X _
                    And LPoints(LSegments%(0, i%)).z = LPoints(LSegments%(0, j%)).z _
                    And LPoints(LSegments%(1, i%)).X = LPoints(LSegments%(1, j%)).X _
                    And LPoints(LSegments%(1, i%)).z = LPoints(LSegments%(1, j%)).z Then
                        ' Double bout
                        Trouve = True
                        P1% = Ajoute_Point(NPoints(), LPoints(LSegments%(0, i%)))
                        P2% = Ajoute_Point(NPoints(), LPoints(LSegments%(1, i%)))
                        NbNSegment% = NbNSegment% + 1
                        ReDim Preserve NSegments%(2, NbNSegment%)
                        NSegments%(0, NbNSegment%) = P1%
                        NSegments%(1, NbNSegment%) = P2%
                        NSegments%(2, NbNSegment%) = 0 ' Non Affiché
                        LSegments%(2, j%) = -1 ' Plus possible à utiliser
                    End If
                    If LPoints(LSegments%(0, i%)).X = LPoints(LSegments%(1, j%)).X _
                    And LPoints(LSegments%(0, i%)).z = LPoints(LSegments%(1, j%)).z _
                    And LPoints(LSegments%(1, i%)).X = LPoints(LSegments%(0, j%)).X _
                    And LPoints(LSegments%(1, i%)).z = LPoints(LSegments%(0, j%)).z Then
                        ' Bout collé
                        Trouve = True
                        LSegments%(2, j%) = -1 ' Plus possible à utiliser
                    End If
                End If
            Next j%
            If Trouve = False Then
                ' Bout tout seul
                P1% = Ajoute_Point(NPoints(), LPoints(LSegments%(0, i%)))
                P2% = Ajoute_Point(NPoints(), LPoints(LSegments%(1, i%)))
                NbNSegment% = NbNSegment% + 1
                ReDim Preserve NSegments%(2, NbNSegment%)
                NSegments%(0, NbNSegment%) = P1%
                NSegments%(1, NbNSegment%) = P2%
                NSegments%(2, NbNSegment%) = 0 ' Non Affiché
            End If
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
                'Coef#(0, i%) = LPoints(LSegments%(0, i%)).x
                'Coef#(1, i%) = dz#
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
    For i% = 1 To NbLSegment%
        If LSegments%(2, i%) = 1 Then
            NbNouveauPoint% = 0
            ReDim NouveauPoint(NbNouveauPoint%)
            Trouve = False
            For j% = 1 To NbLSegment%
                If LSegments%(2, j%) = 1 Then
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
                        NouveauPoint(NbNouveauPoint%) = LePoint
                    End If
                End If
            Next j%
            '
            ' ***** Trie à bulle
            '
            For P1% = 1 To NbNouveauPoint% - 1
                For P2% = P1% + 1 To NbNouveauPoint%
                    If LPoints(LSegments%(0, i%)).X < LPoints(LSegments%(1, i%)).X Then
                        If NouveauPoint(P1%).X > NouveauPoint(P2%).X Then
                            LePoint = NouveauPoint(P1%)
                            NouveauPoint(P1%) = NouveauPoint(P2%)
                            NouveauPoint(P2%) = LePoint
                        End If
                    Else
                        If NouveauPoint(P1%).X < NouveauPoint(P2%).X Then
                            LePoint = NouveauPoint(P1%)
                            NouveauPoint(P1%) = NouveauPoint(P2%)
                            NouveauPoint(P2%) = LePoint
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
                ReDim Preserve NSegments%(2, NbNSegment%)
                NSegments%(0, NbNSegment%) = P1%
                NSegments%(1, NbNSegment%) = P2%
                NSegments%(2, NbNSegment%) = 1 ' A Afficher
'Debug.Print "Segment"; i%; j%; Format$(Coef#(0, i%), "##.## ");
'Debug.Print p1%; "=["; NPoints(p1%).x; NPoints(p1%).z; "]";
'Debug.Print p2%; "=["; NPoints(p2%).x; NPoints(p2%).z; "]"
            Next j%
        End If
    Next i%
    '
    ' ***** Génération du contour (enfin!)
    '
    Dim AngleSegment%()
    Dim NbAngleSegment%
    Dim AngleActuel% ' Angle du segment en cours
'Dim n%
    '
    For i% = 1 To NbNSegment%
        NSegments%(2, i%) = -NSegments%(2, i%) ' Par défaut non utilisé
    Next i%
    P1% = 1 ' Segment de départ
    Do
        P2% = NSegments%(1, P1%) ' Point de départ
        AngleActuel% = Calcul_Angle%(NPoints(NSegments%(0, P1%)).X, NPoints(NSegments%(0, P1%)).z, NPoints(NSegments%(1, P1%)).X, NPoints(NSegments%(1, P1%)).z)
'Debug.Print "Départ"; p1%; p2%; AngleActuel%
        NbAngleSegment% = 0
        ReDim AngleSegment%(1, NbAngleSegment%)
        For i% = 1 To NbNSegment%
            If NSegments%(0, i%) = P2% Then
                NbAngleSegment% = NbAngleSegment% + 1
                ReDim Preserve AngleSegment%(1, NbAngleSegment%)
                j% = Calcul_Angle%(NPoints(NSegments%(0, i%)).X, NPoints(NSegments%(0, i%)).z, NPoints(NSegments%(1, i%)).X, NPoints(NSegments%(1, i%)).z)
'Debug.Print j%
                j% = -j% + AngleActuel% + 180
                If j% > 360 Then j% = j% - 360
                If j% < 0 Then j% = j% + 360
                AngleSegment%(0, NbAngleSegment%) = i%
                AngleSegment%(1, NbAngleSegment%) = j%
'Debug.Print NbAngleSegment%; AngleSegment%(0, NbAngleSegment%); AngleSegment%(1, NbAngleSegment%)
            End If
        Next i%
        j% = 1
        If NbAngleSegment% > 1 Then
            For i% = 2 To NbAngleSegment%
                If AngleSegment(1, i%) < AngleSegment%(1, j%) Then j% = i%
            Next i%
        End If
        P1% = AngleSegment%(0, j%)
        NSegments%(2, P1%) = -NSegments%(2, P1%)
        If NSegments%(1, P1%) = 1 Then Exit Do ' Fin du parcours
        '
        DoEvents
        '
'n% = n% + 1
'If n% > NbNSegment% Then Exit Do
    Loop
    '
    NbContour% = 0
    For i% = 1 To NbNSegment%
        If NSegments%(2, i%) <> -1 Then
            NbContour% = NbContour% + 1
        End If
    Next i%
End Sub

'
' **********************************
' Crée les segments de départ à
' partir des informations de la voie
' **********************************
'
Public Sub Cree_Couche(No%, Largeur!)
    Dim i%, j%, k%, l%
    Dim n%, t!, X!, z!
    Dim nTraverse%, dTraverse!
    Dim v As D3DVECTOR
    Dim vt As Single
    Dim AxeY As D3DVECTOR
    Dim dX7 As New DirectX7
    Dim module!
    Dim NbLPoint%
    '
    AxeY.X = 0
    AxeY.Y = 1
    AxeY.z = 0
    '
    ' ***** Calcul des segments
    '
    Call Recalcule_Point(No%)
    NbLPoint% = 0
    NbLSegment% = 0
    NbLRectangle% = 0
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
                    '
                    NbLPoint% = NbLPoint% + 2
                    ReDim Preserve LPoints(NbLPoint%) As D3DVECTOR
                    '
                    LPoints(NbLPoint% - 1).X = 0
                    LPoints(NbLPoint% - 1).z = Largeur!
                    LPoints(NbLPoint%).X = 0
                    LPoints(NbLPoint%).z = -Largeur!
                    '
                    For l% = 0 To 1
                        If vt! <> 0 Then
                            module! = dX7.VectorModulus(LPoints(NbLPoint% - l%))
                            Call dX7.VectorRotate(LPoints(NbLPoint% - l%), LPoints(NbLPoint% - l%), AxeY, vt!)
                            Call dX7.VectorScale(LPoints(NbLPoint% - l%), LPoints(NbLPoint% - l%), module!)
                        End If
                        Call dX7.VectorAdd(LPoints(NbLPoint% - l%), LPoints(NbLPoint% - l%), v)
                    Next l%
                    '
                    If k% = 0 Then
                        NbLSegment% = NbLSegment% + 1
                        ReDim Preserve LSegments(2, NbLSegment%) As Integer
                        LSegments%(0, NbLSegment%) = NbLPoint%
                        LSegments%(1, NbLSegment%) = NbLPoint% - 1
                        LSegments%(2, NbLSegment%) = 0 ' Bout
                    Else
                        NbLSegment% = NbLSegment% + 2
                        ReDim Preserve LSegments(2, NbLSegment%) As Integer
                        LSegments%(0, NbLSegment% - 1) = NbLPoint% - 3
                        LSegments%(1, NbLSegment% - 1) = NbLPoint% - 1
                        LSegments%(2, NbLSegment% - 1) = 1 ' A tracer
                        LSegments%(0, NbLSegment%) = NbLPoint%
                        LSegments%(1, NbLSegment%) = NbLPoint% - 2
                        LSegments%(2, NbLSegment%) = 1 ' A tracer
                        '
    NbLRectangle% = NbLRectangle% + 1
    ReDim Preserve LRectangle(3, NbLRectangle%)
    LRectangle(0, NbLRectangle%) = LPoints(NbLPoint% - 3)
    LRectangle(0, NbLRectangle%).Y = 30
    LRectangle(1, NbLRectangle%) = LPoints(NbLPoint% - 1)
    LRectangle(1, NbLRectangle%).Y = 30
    LRectangle(2, NbLRectangle%) = LPoints(NbLPoint%)
    LRectangle(2, NbLRectangle%).Y = 30
    LRectangle(3, NbLRectangle%) = LPoints(NbLPoint% - 2)
    LRectangle(3, NbLRectangle%).Y = 30
    '
    ' ************************
    ' Positions intermédiaires
    ' ************************
    '
    'If k% <> nTraverse% + 1 Then
    '    NbLSegment% = NbLSegment% + 1
    '    ReDim Preserve LSegments(2, NbLSegment%) As Integer
    '    LSegments%(0, NbLSegment%) = NbLPoint% - 1
    '    LSegments%(1, NbLSegment%) = NbLPoint%
    '    LSegments%(2, NbLSegment%) = 1 ' A tracer
    'End If
                        '
                    End If
                Next k%
                '
                NbLSegment% = NbLSegment% + 1
                ReDim Preserve LSegments(2, NbLSegment%) As Integer
                LSegments%(0, NbLSegment%) = NbLPoint% - 1
                LSegments%(1, NbLSegment%) = NbLPoint%
                LSegments%(2, NbLSegment%) = 0 ' Non tracé
                '
            End If
        Next j%
    Next i%
    Set dX7 = Nothing
End Sub

'
' **************************************
' Test si les deux segments sont sécants
' **************************************
'
Private Function Secante(c#(), s1%, s2%, TempPoint() As D3DVECTOR, TempSegment%(), LePoint As D3DVECTOR) As Boolean
    Dim bX1!, bZ1!
    Dim bX2!, bZ2!
    Dim bX3!, bZ3!
    Dim bX4!, bZ4!
    '
    ' ***** Temporaire si droite verticale
    If c#(2, s1%) = 1 Or c#(2, s2%) = 1 Then Exit Function
    '
    ' ***** Parallèle?
    If c#(0, s1%) = c#(0, s2%) Then Exit Function
    '
    LePoint.X = (c#(1, s2%) - c#(1, s1%)) / (c#(0, s1%) - c#(0, s2%))
    LePoint.z = c#(0, s1%) * LePoint.X + c#(1, s1%)
    '
    If TempPoint(TempSegment(0, s1%)).X < TempPoint(TempSegment(1, s1%)).X Then
        bX1! = TempPoint(TempSegment(0, s1%)).X
        bX2! = TempPoint(TempSegment(1, s1%)).X
    Else
        bX1! = TempPoint(TempSegment(1, s1%)).X
        bX2! = TempPoint(TempSegment(0, s1%)).X
    End If
    If TempPoint(TempSegment(0, s1%)).z < TempPoint(TempSegment(1, s1%)).z Then
        bZ1! = TempPoint(TempSegment(0, s1%)).z
        bZ2! = TempPoint(TempSegment(1, s1%)).z
    Else
        bZ1! = TempPoint(TempSegment(1, s1%)).z
        bZ2! = TempPoint(TempSegment(0, s1%)).z
    End If
    If TempPoint(TempSegment(0, s2%)).X < TempPoint(TempSegment(1, s2%)).X Then
        bX3! = TempPoint(TempSegment(0, s2%)).X
        bX4! = TempPoint(TempSegment(1, s2%)).X
    Else
        bX3! = TempPoint(TempSegment(1, s2%)).X
        bX4! = TempPoint(TempSegment(0, s2%)).X
    End If
    If TempPoint(TempSegment(0, s2%)).z < TempPoint(TempSegment(1, s2%)).z Then
        bZ3! = TempPoint(TempSegment(0, s2%)).z
        bZ4! = TempPoint(TempSegment(1, s2%)).z
    Else
        bZ3! = TempPoint(TempSegment(1, s2%)).z
        bZ4! = TempPoint(TempSegment(0, s2%)).z
    End If
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
End Function

'
' ********************************
' Initialise les valeurs de départ
' ********************************
'
Private Sub Class_Initialize()
    ReDim LPoints(0) As D3DVECTOR
    ReDim LSegments(2, 0) As Integer
    ReDim NPoints(0) As D3DVECTOR
    ReDim NSegments(2, 0) As Integer
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
    Call Cree_Couche(n%, 13)
End Sub

'
' ********************************
' Récupère la liste des rectangles
' ********************************
'
Public Sub Liste_Rectangle(n%, P1 As D3DVECTOR, P2 As D3DVECTOR, P3 As D3DVECTOR, P4 As D3DVECTOR)
     P1 = LRectangle(0, n%)
     P2 = LRectangle(1, n%)
     P3 = LRectangle(2, n%)
     P4 = LRectangle(3, n%)
End Sub

