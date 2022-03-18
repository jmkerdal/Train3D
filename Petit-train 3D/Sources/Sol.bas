Attribute VB_Name = "Sol"
Option Explicit
'
Dim grille%()
Dim PointSol%()
Dim XGrille%, ZGrille%
Public PAS_CASE%
Public HMax%
Dim Rectangles() As D3DVECTOR
'
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'
' ***************
' Création du sol
' ***************
'
Public Sub Cree_Sol()
    Dim P(3) As D3DVECTOR ' Définition d'une face
    Dim v(1) As D3DVECTOR ' 2 vecteurs supplémentaires pour les cotés
    Dim c&
    Dim x1%, x2%, z1%, z2%
    Dim i%, X%, z%
    '
    If Command$ = Super$ Then
        ' ***** Pour les tests
        PAS_CASE% = 100
    Else
        PAS_CASE% = 25
    End If
    '
    x1% = xPlateauMin% - BORD%
    z1% = zPlateauMin% - BORD%
    XGrille% = xPlateauMax% - xPlateauMin% + 2 * BORD%
    ZGrille% = zPlateauMax% - zPlateauMin% + 2 * BORD%
    XGrille% = (XGrille% + PAS_CASE% / 2) / PAS_CASE%
    ZGrille% = (ZGrille% + PAS_CASE% / 2) / PAS_CASE%
    x2% = x1% + (XGrille% + 1) * PAS_CASE%
    z2% = z1% + (ZGrille% + 1) * PAS_CASE%
    '
    Set dxFond = dxVue.dxD3Drm.CreateMeshBuilder()
    Call dxConstruitMeshBuilder.Init
    '
    If ModeActuelle = ModeVisualisation Then
        '
        ' Coche les cases
        '
        ReDim grille%(XGrille%, ZGrille%)
        ReDim Rectangles(3, 0)
        For i% = 1 To UBound(Reseau())
            Call Coche_Voie(i%, x1%, z1%)
        Next i%
        For i% = 1 To UBound(ElementDecor())
            Call Coche_Decor(i%, x1%, z1%)
        Next i%
        Call Coche_Catenaire(x1%, z1%)
        '
        If ParamElevation = True Then
            '
            ' Calcul les élévations
            '
            Call Genere_Points
            '
            For X% = 0 To XGrille%
                For z% = 0 To ZGrille%
                    P(0).X = (X% * PAS_CASE% + x1%) / REDUCTION%
                    P(0).z = (-PAS_CASE% / 2.5 * PointSol%(X%, z% + 1) + 0.1) / REDUCTION%
                    P(0).Y = ((z% + 1) * PAS_CASE% + z1%) / REDUCTION%
                    '
                    P(1).X = ((X% + 1) * PAS_CASE% + x1%) / REDUCTION%
                    P(1).z = (-PAS_CASE% / 2.5 * PointSol%(X% + 1, z% + 1) + 0.1) / REDUCTION%
                    P(1).Y = ((z% + 1) * PAS_CASE% + z1%) / REDUCTION%
                    '
                    P(2).X = ((X% + 1) * PAS_CASE% + x1%) / REDUCTION%
                    P(2).z = (-PAS_CASE% / 2.5 * PointSol%(X% + 1, z%) + 0.1) / REDUCTION%
                    P(2).Y = (z% * PAS_CASE% + z1%) / REDUCTION%
                    '
                    P(3).X = (X% * PAS_CASE% + x1%) / REDUCTION%
                    P(3).z = (-PAS_CASE% / 2.5 * PointSol%(X%, z%) + 0.1) / REDUCTION%
                    P(3).Y = (z% * PAS_CASE% + z1%) / REDUCTION%
                    '
                    If grille%(X%, z%) = 1 Then
                        c& = dxVue.dX7.CreateColorRGBA(0.5, 0.5, 0.5, 1)
                        'c& = &HFF000000
                    Else
                        c& = dxVue.dX7.CreateColorRGBA(PointSol(X%, z%) / HMax% / 2 + 0.5, PointSol(X%, z%) / HMax% / 2 + 0.5, PointSol(X%, z%) / HMax% / 2 + 0.5, 1)
                    End If
                    Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(2), P(3), -1, c&)
                Next z%
            Next X%
        Else
            ' Pas d'élévation
            c& = &HFF00FF00
            P(0).X = x1% / REDUCTION%: P(0).Y = z2% / REDUCTION%: P(0).z = 0
            P(1).X = x2% / REDUCTION%: P(1).Y = z2% / REDUCTION%: P(1).z = 0
            P(2).X = x2% / REDUCTION%: P(2).Y = z1% / REDUCTION%: P(2).z = 0
            P(3).X = x1% / REDUCTION%: P(3).Y = z1% / REDUCTION%: P(3).z = 0
            Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(2), P(3), -1, c&)
        End If
        '
        ' ***** Ajoute les 4 cotés
        '
        c& = &HFF004F00
        P(0).X = x1% / REDUCTION%: P(0).Y = z2% / REDUCTION%: P(0).z = 0
        P(1).X = (x1% - 70) / REDUCTION%: P(1).Y = (z2% + 70) / REDUCTION%: P(1).z = 50 / REDUCTION%
        P(2).X = (x2% + 70) / REDUCTION%: P(2).Y = (z2% + 70) / REDUCTION%: P(2).z = 50 / REDUCTION%
        P(3).X = x2% / REDUCTION%: P(3).Y = z2% / REDUCTION%: P(3).z = 0
        Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(2), P(3), -1, c&)
        P(0).X = x1% / REDUCTION%: P(0).Y = z1% / REDUCTION%: P(0).z = 0
        P(1).X = x2% / REDUCTION%: P(1).Y = z1% / REDUCTION%: P(1).z = 0
        P(2).X = (x2% + 70) / REDUCTION%: P(2).Y = (z1% - 70) / REDUCTION%: P(2).z = 50 / REDUCTION%
        P(3).X = (x1% - 70) / REDUCTION%: P(3).Y = (z1% - 70) / REDUCTION%: P(3).z = 50 / REDUCTION%
        Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(2), P(3), -1, c&)
        P(0).X = x1% / REDUCTION%: P(0).Y = z2% / REDUCTION%: P(0).z = 0
        P(1).X = x1% / REDUCTION%: P(1).Y = z1% / REDUCTION%: P(1).z = 0
        P(2).X = (x1% - 70) / REDUCTION%: P(2).Y = (z1% - 70) / REDUCTION%: P(2).z = 50 / REDUCTION%
        P(3).X = (x1% - 70) / REDUCTION%: P(3).Y = (z2% + 70) / REDUCTION%: P(3).z = 50 / REDUCTION%
        Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(2), P(3), -1, c&)
        P(0).X = x2% / REDUCTION%: P(0).Y = z2% / REDUCTION%: P(0).z = 0
        P(1).X = (x2% + 70) / REDUCTION%: P(1).Y = (z2% + 70) / REDUCTION%: P(1).z = 50 / REDUCTION%
        P(2).X = (x2% + 70) / REDUCTION%: P(2).Y = (z1% - 70) / REDUCTION%: P(2).z = 50 / REDUCTION%
        P(3).X = x2% / REDUCTION%: P(3).Y = z1% / REDUCTION%: P(3).z = 0
        Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(2), P(3), -1, c&)
    Else
        ' Vue génération simplifié
        c& = &HFF004F00
        P(0).X = x1% / REDUCTION%: P(0).Y = z2% / REDUCTION%: P(0).z = 1 / REDUCTION%
        P(1).X = x2% / REDUCTION%: P(1).Y = z2% / REDUCTION%: P(1).z = 1 / REDUCTION%
        P(2).X = x2% / REDUCTION%: P(2).Y = z1% / REDUCTION%: P(2).z = 1 / REDUCTION%
        P(3).X = x1% / REDUCTION%: P(3).Y = z1% / REDUCTION%: P(3).z = 1 / REDUCTION%
        Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(2), P(3), -1, c&)
    End If
    '
    Call dxConstruitMeshBuilder.Build(dxFond, dxSolTexture())
    If ModeActuelle = ModeVisualisation Then
        Call dxVue.Wrap_Texture(dxFond, dxSolTexture(0), 5, (xPlateauMax% - xPlateauMin% + 2 * BORD%) \ 256, (zPlateauMax% - zPlateauMin% + 2 * BORD%) \ 256)
    End If
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
    '
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
    '
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

'
' *************************************
' Test les cases cochées par un segment
' *************************************
'
Public Sub Coche_Case(Ax1!, Ay1!, Ax2!, Ay2!)
    Dim xMin%, Ymin%, xMax%, Ymax%
    Dim X%, Y%
    Dim c&
    If Ax1! < Ax2! Then
        xMin% = Ax1! \ PAS_CASE%
        xMax% = Ax2! / PAS_CASE% + 0.5
    Else
        xMin% = Ax2! \ PAS_CASE%
        xMax% = Ax1! / PAS_CASE% + 0.5
    End If
    xMin% = xMin% * PAS_CASE%
    xMax% = xMax% * PAS_CASE%
    If Ay1! < Ay2! Then
        Ymin% = Ay1! \ PAS_CASE%
        Ymax% = Ay2! / PAS_CASE% + 0.5
    Else
        Ymin% = Ay2! \ PAS_CASE%
        Ymax% = Ay1! / PAS_CASE% + 0.5
    End If
    Ymin% = Ymin% * PAS_CASE%
    Ymax% = Ymax% * PAS_CASE%
    '
    ' On marque le point de départ
    grille%(Ax1! \ PAS_CASE%, Ay1! \ PAS_CASE%) = 1
    '
    For X% = xMin% To xMax% - PAS_CASE% Step PAS_CASE%
        For Y% = Ymin% To Ymax% - PAS_CASE% Step PAS_CASE%
            '
            If X% <> xMin% Then
                If SecanteV(Ax1!, Ay1!, Ax2!, Ay2!, Y%, Y% + PAS_CASE%, X%) = True Then
                    grille%(X% / PAS_CASE%, Y% / PAS_CASE%) = 1
                    grille%(Int(X% / PAS_CASE%) - 1, Y% / PAS_CASE%) = 1
                End If
            End If
            If Y% <> Ymin% Then
                If SecanteH(Ax1!, Ay1!, Ax2!, Ay2!, X%, X% + PAS_CASE%, Y%) = True Then
                    grille%(X% / PAS_CASE%, Y% / PAS_CASE%) = 1
                    grille%(X% / PAS_CASE%, Int(Y% / PAS_CASE%) - 1) = 1
                End If
            End If
        Next Y%
    Next X%
End Sub

'
' ***************************************
' Calcul des cases cochées par les décors
' ***************************************
'
Public Sub Coche_Decor(n%, dX%, dz%)
    Dim dxBox As D3DRMBOX
    Dim v(3) As D3DVECTOR
    Dim i%
    Dim l!
    Dim AxeY As D3DVECTOR
    '
    AxeY.X = 0: AxeY.Y = 1: AxeY.z = 0
    Call ListeDecor(ElementDecor(n%).NoDecor%).dxDecor.GetBox(dxBox)
    '
    v(0) = dxBox.Min
    v(1) = dxBox.Min: v(1).X = dxBox.Max.X
    v(2) = dxBox.Max
    v(3) = dxBox.Max: v(3).X = dxBox.Min.X
    '
    For i% = 0 To 3
        Call dxVue.dX7.VectorScale(v(i%), v(i%), REDUCTION%) ' Remet a la bonne taille
        l! = dxVue.dX7.VectorModulus(v(i%)) ' Tourne le rectangle
        Call dxVue.dX7.VectorRotate(v(i%), v(i%), AxeY, ElementDecor(n%).Angle% * DegRad!)
        Call dxVue.dX7.VectorScale(v(i%), v(i%), l!)
        v(i%).X = v(i%).X + ElementDecor(n%).Position.X - dX% ' positionne le rectangle
        v(i%).z = v(i%).z + ElementDecor(n%).Position.z - dz%
    Next i%
    Call Ajoute_Rectangle(v())
'Debug.Print v(0).X; ","; v(0).z; ","; v(1).X; ","; v(1).z; ","; v(2).X; ","; v(2).z; ","; v(3).X; ","; v(3).z
    '
    If ParamElevation = True Then
        Dim nb%
        Dim deltaX!, deltaZ!
        '
        deltaX! = v(2).X - v(1).X
        deltaZ! = v(2).z - v(1).z
        l! = Sqr(deltaX! ^ 2 + deltaZ! ^ 2)
        nb% = (l! \ PAS_CASE%) + 1
        deltaX! = deltaX! / nb%
        deltaZ! = deltaZ! / nb%
        '
        For i% = 0 To nb%
            Call Coche_Case(v(0).X, v(0).z, v(1).X, v(1).z)
            v(0).X = v(0).X + deltaX!
            v(0).z = v(0).z + deltaZ!
            v(1).X = v(1).X + deltaX!
            v(1).z = v(1).z + deltaZ!
        Next i%
    End If
End Sub

'
' ***************************************
' Coche les cases ou sont posés les voies
' ***************************************
'
Public Sub Coche_Voie(n%, dX%, dz%)
    Dim i%, j%, Pa(3) As D3DVECTOR
    Dim l!
    Dim AxeY As D3DVECTOR
    '
    'grille%((Reseau(n%).Position.x - dx%) / PAS_CASE%, (Reseau(n%).Position.z - dz%) / PAS_CASE%) = 1
    AxeY.X = 0: AxeY.Y = 1: AxeY.z = 0
    '
    For i% = 1 To FormeVoie(Reseau(n%).NoVoie%).NbLRectangle%
        Call FormeVoie(Reseau(n%).NoVoie%).Liste_Rectangle(i%, Pa(0), Pa(1), Pa(2), Pa(3))
'Debug.Print "a "; i%; Pa(0).X; Pa(0).z; "*"; Pa(1).X; Pa(1).z; "*"; Pa(2).X; Pa(2).z; "*"; Pa(3).X; Pa(3).z
        For j% = 0 To 3
            l! = dxVue.dX7.VectorModulus(Pa(j%))
            Call dxVue.dX7.VectorRotate(Pa(j%), Pa(j%), AxeY, Reseau(n%).Angle! * DegRad!)
            Call dxVue.dX7.VectorScale(Pa(j%), Pa(j%), l!)
            Pa(j%).X = Pa(j%).X + Reseau(n%).Position.X - dX%
            Pa(j%).z = Pa(j%).z + Reseau(n%).Position.z - dz%
        Next j%
'Debug.Print "b "; Pa(0).X; Pa(0).z; "*"; Pa(1).X; Pa(1).z; "*"; Pa(2).X; Pa(2).z; "*"; Pa(3).X; Pa(3).z
'Debug.Print Pa(0).X; ","; Pa(0).z; ","; Pa(1).X; ","; Pa(1).z; ","; Pa(2).X; ","; Pa(2).z; ","; Pa(3).X; ","; Pa(3).z
        If ParamElevation = True Then
            Call Coche_Case(Pa(0).X, Pa(0).z, Pa(1).X, Pa(1).z)
            'Call Coche_Case(Pa(1).X, Pa(1).z, Pa(2).X, Pa(2).z)
            Call Coche_Case(Pa(2).X, Pa(2).z, Pa(3).X, Pa(3).z)
            'Call Coche_Case(Pa(3).X, Pa(3).z, Pa(0).X, Pa(0).z)
            '
            ' Coche en croix
            Call Coche_Case(Pa(0).X, Pa(0).z, Pa(2).X, Pa(2).z)
            Call Coche_Case(Pa(1).X, Pa(1).z, Pa(3).X, Pa(3).z)
        End If
        Call Ajoute_Rectangle(Pa())
    Next i%
End Sub

'
' *******************************************
' Génére la hauteur des points pour le tunnel
' *******************************************
'
Public Sub Genere_Points()
    Dim i%, j%
    Call Rnd(-1)
    Call Randomize(0)
    '
    ' ***** Rempli la liste des points avec 99
    ' ***** sauf le contour
    '
    ReDim PointSol%(XGrille% + 1, ZGrille% + 1)
    For i% = 1 To ZGrille%
        For j% = 1 To XGrille%
            PointSol%(j%, i%) = 99
        Next j%
    Next i%
    '
    ' ***** Bloque les points ou la case est coché
    '
    For i% = 0 To ZGrille%
        For j% = 0 To XGrille%
                If grille%(j%, i%) = 1 Then
                    PointSol%(j%, i%) = 0
                    PointSol%(j% + 1, i%) = 0
                    PointSol%(j%, i% + 1) = 0
                    PointSol%(j% + 1, i% + 1) = 0
                End If
        Next j%
    Next i%
    '
    ' ***** Monte les points
    '
    Dim Trouve As Boolean
    Dim Pa%, pb%, pc%, pd%
    Dim pmin%
    HMax% = 0
    Do
        Trouve = False
        For i% = 1 To ZGrille%
            For j% = 1 To XGrille%
                If PointSol%(j%, i%) = 99 Then
                    Trouve = True
                    If i% = 0 Then
                        Pa% = 99
                    Else
                        Pa% = PointSol%(j%, i% - 1)
                    End If
                    If j% = XGrille% + 1 Then
                        pb% = 99
                    Else
                        pb% = PointSol%(j% + 1, i%)
                    End If
                    If i% = ZGrille% + 1 Then
                        pc% = 99
                    Else
                        pc% = PointSol%(j%, i% + 1)
                    End If
                    If j% = 0 Then
                        pd% = 99
                    Else
                        pd% = PointSol%(j% - 1, i%)
                    End If
                    '
                    pmin% = 99
                    If Pa% < pmin% Then pmin% = Pa%
                    If pb% < pmin% Then pmin% = pb%
                    If pc% < pmin% Then pmin% = pc%
                    If pd% < pmin% Then pmin% = pd%
                    '
                    If pmin% = HMax% Then
                        If Rnd * 10 > 10 - ParamHauteur% Then
                            PointSol%(j%, i%) = HMax% + 1
                        Else
                            PointSol%(j%, i%) = HMax%
                        End If
                    ElseIf pmin% < HMax% Then
                        PointSol%(j%, i%) = pmin% + 1
                    End If
                End If
            Next j%
        Next i%
        HMax% = HMax% + 1
    Loop While Trouve = True
End Sub

'
' **************************************
' Coche les cases ou sont les catenaires
' **************************************
'
Public Sub Coche_Catenaire(dX%, dz%)
    Dim i%, j%, k%
    Dim P(3) As D3DVECTOR ' 4 points de base
    Dim v(3) As D3DVECTOR ' 4 points calculés
    '
    P(0).X = 10 / REDUCTION%: P(0).z = -45 / REDUCTION%
    P(1).X = 10 / REDUCTION%: P(1).z = -25 / REDUCTION%
    P(2).X = -10 / REDUCTION%: P(2).z = -25 / REDUCTION%
    P(3).X = -10 / REDUCTION%: P(3).z = -45 / REDUCTION%
    '
    ' Mémorise le nombre de rectangle avant les catenaires
    Dim n%
    n% = UBound(Rectangles(), 2) - 1
    '
    For i% = 1 To UBound(Reseau())
        If Reseau(i%).NoVoie% <> 0 Then
            For j% = 0 To NbSegment%
                If ReseauElectrique = True Then
                    If Voie(Reseau(i%).NoVoie%).CatenaireSegment%(j%) <> 0 Then
                        '
                        For k% = 0 To 3
                            Call dxFrameVoie(j% + 3, i%).Transform(v(k%), P(k%))
                            v(k%).X = v(k%).X * REDUCTION% - dX%
                            v(k%).z = v(k%).z * REDUCTION% - dz%
                        Next k%
                        '
                        If Test_Intersection(v(), n%) = True Then
                            Reseau(i%).MetCatenaire(j%) = False
                        Else
                            Reseau(i%).MetCatenaire(j%) = True
                            Call Coche_Case(v(0).X, v(0).z, v(1).X, v(1).z)
                            'Call Coche_Case(v(1).X, v(1).z, v(2).X, v(2).z)
                            Call Coche_Case(v(2).X, v(2).z, v(3).X, v(3).z)
                            'Call Coche_Case(v(3).X, v(3).z, v(0).X, v(0).z)
                        End If
                    End If
                Else
                    Reseau(i%).MetCatenaire(j%) = False
                End If
            Next j%
        End If
    Next i%
End Sub

'
' ********************************************
' Ajoute un rectangle dans la liste
' pour pouvoir tester ensuite si un catenaire
' est sur une voie ou un décors, et l'éliminer
' ********************************************
'
Public Sub Ajoute_Rectangle(P() As D3DVECTOR)
    Dim n%, i%
    n% = UBound(Rectangles(), 2)
    ReDim Preserve Rectangles(3, n% + 1)
    For i% = 0 To 3
        Rectangles(i%, n%) = P(i%)
    Next i%
End Sub

'
' **********************************
' Test l'intersection d'un rectangle
' et d'un element dans la liste
' **********************************
'
Public Function Test_Intersection(Rec1() As D3DVECTOR, nb%) As Boolean
    Dim Polygone As Long ' Polygone object pointer
    Dim P1(5) As POINTAPI
    Dim P2(5) As POINTAPI
    Dim i%, n%
    '
    For i% = 0 To 3
        P1(i%).X = Rec1(i%).X
        P1(i%).Y = Rec1(i%).z
    Next i%
    '
    For n% = 0 To nb%
        For i% = 0 To 3
            P2(i%).X = Rectangles(i%, n%).X
            P2(i%).Y = Rectangles(i%, n%).z
        Next i%
        '
        Polygone = CreatePolygonRgn(P1(0), 4, 2)
        For i% = 0 To 3
            If PtInRegion(Polygone, P2(i%).X, P2(i%).Y) <> 0 Then
                Test_Intersection = True
                Exit For
            End If
        Next i%
        Call DeleteObject(Polygone)
        If Test_Intersection = True Then Exit Function
        '
        Polygone = CreatePolygonRgn(P2(0), 4, 2)
        For i% = 0 To 3
            If PtInRegion(Polygone, P1(i%).X, P1(i%).Y) <> 0 Then
                Test_Intersection = True
                Exit For
            End If
        Next i%
        Call DeleteObject(Polygone)
        If Test_Intersection = True Then Exit Function
    Next n%
End Function

