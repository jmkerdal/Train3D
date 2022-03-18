Attribute VB_Name = "Tunnel"
Option Explicit

Global ListeTunnel() As New ClassTunnel
Global FrameTunnel() As Direct3DRMFrame3
Global MeshTunnel() As Direct3DRMMeshBuilder3
Global NoTunnel% ' No du tunnel qui est pointé
Global SelectionTunnel% ' No du tunnel dont on veut déplacer un point

'
' ****************************************
' Génére un tunnel et l'attache a sa frame
' ****************************************
'
Private Sub Genere(No%)
    Dim i%, n%
    Dim i1%
    Dim c&
    Dim P(3) As D3DVECTOR
    '
    Set MeshTunnel(No%) = dxVue.dxD3Drm.CreateMeshBuilder()
    Call dxConstruitMeshBuilder.Init
    '
    If ListeTunnel(No%).Nb_Point% <> 0 Then
        If ListeTunnel(No%).Genere = True Then
            n% = ListeTunnel(No%).Nb_Triangle%()
            For i% = 0 To n% - 1
                For i1% = 0 To 2
                    P(i1%).X = ListeTunnel(No%).TriangleX(i1%, i%)
                    P(i1%).Y = 100
                    P(i1%).z = ListeTunnel(No%).TriangleZ(i1%, i%)
                Next i1%
                ' Dessus
                c& = dxVue.dX7.CreateColorRGBA(0, 0.6, 0, 1)
                Call dxConstruitMeshBuilder.Add_Triangle(P(0), P(1), P(2), 0, c&)
                ' Plafond
                P(0).Y = 85
                P(1).Y = 85
                P(2).Y = 85
                c& = dxVue.dX7.CreateColorRGBA(0.25, 0.25, 0.25, 1)
                Call dxConstruitMeshBuilder.Add_Triangle(P(0), P(2), P(1), 1, c&)
            Next i%
        End If
        '
        n% = ListeTunnel(No%).Nb_Point%()
        For i% = 0 To n% - 1
            i1% = (i% + 1) Mod n%
            P(0).X = ListeTunnel(No%).PositionX(i%)
            P(0).Y = 100
            P(0).z = ListeTunnel(No%).PositionZ(i%)
            P(1).X = ListeTunnel(No%).PositionX(i1%)
            P(1).Y = 100
            P(1).z = ListeTunnel(No%).PositionZ(i1%)
            If ListeTunnel(No%).Face(i%) = True Then
                P(0).Y = 85
                P(1).Y = 85
                P(2) = P(0)
                P(2).Y = 0
                P(3) = P(1)
                P(3).Y = 0
                ' Coté interne
                c& = dxVue.dX7.CreateColorRGBA(0.25, 0.25, 0.25, 1)
                Call dxConstruitMeshBuilder.Add_Rectangle(P(0), P(1), P(3), P(2), 1, c&)
                ' Coté externe
                P(0).Y = 100
                P(1).Y = 100
                P(2).X = ListeTunnel(No%).PositionX(i%) + ListeTunnel(No%).NormalX(i%) * 50
                P(2).z = ListeTunnel(No%).PositionZ(i%) + ListeTunnel(No%).NormalZ(i%) * 50
                P(3).X = ListeTunnel(No%).PositionX(i1%) + ListeTunnel(No%).NormalX(i1%) * 50
                P(3).z = ListeTunnel(No%).PositionZ(i1%) + ListeTunnel(No%).NormalZ(i1%) * 50
                c& = dxVue.dX7.CreateColorRGBA(0, 0.5, 0, 1)
                Call dxConstruitMeshBuilder.Add_Rectangle(P(1), P(0), P(2), P(3), 0, c&)
            Else
                ' Fronton
                P(2) = P(0)
                P(2).Y = 85
                P(3) = P(1)
                P(3).Y = 85
                c& = dxVue.dX7.CreateColorRGBA(0.5, 0.5, 0.5, 1)
                Call dxConstruitMeshBuilder.Add_Rectangle(P(1), P(0), P(2), P(3), 1, c&)
                ' Triangles coté
                P(2).Y = 0
                P(3).X = ListeTunnel(No%).PositionX(i%) + ListeTunnel(No%).NormalX(i%) * 50
                P(3).Y = 0
                P(3).z = ListeTunnel(No%).PositionZ(i%) + ListeTunnel(No%).NormalZ(i%) * 50
                c& = dxVue.dX7.CreateColorRGBA(0, 0.5, 0, 1)
                Call dxConstruitMeshBuilder.Add_Triangle(P(0), P(3), P(2), 0, c&)
                P(2) = P(1)
                P(2).Y = 0
                P(3).X = ListeTunnel(No%).PositionX(i1%) + ListeTunnel(No%).NormalX(i1%) * 50
                P(3).z = ListeTunnel(No%).PositionZ(i1%) + ListeTunnel(No%).NormalZ(i1%) * 50
                Call dxConstruitMeshBuilder.Add_Triangle(P(1), P(2), P(3), 0, c&)
            End If
        Next i%
        '
    End If
    Call dxConstruitMeshBuilder.Build(MeshTunnel(No%), dxSolTexture())
    Call MeshTunnel(No%).ScaleMesh(1 / REDUCTION%, 1 / REDUCTION%, 1 / REDUCTION%)
    Set FrameTunnel(No%) = dxVue.dxD3Drm.CreateFrame(Nothing)
    Call FrameTunnel(No%).AddVisual(MeshTunnel(No%))
End Sub

'
' ********************************************
' Crée un nouveau tunnel de 100x100 par défaut
' ********************************************
'
Public Sub Cree()
    Dim n%, i%, t%
    n% = UBound(ListeTunnel())
    ' ***** Cherche un tunnel vide
    t% = -1
    For i% = 0 To n% - 1
        If t% = -1 And ListeTunnel(i%).Nb_Point% = 0 Then
            t% = i%
        End If
    Next i%
    ' ***** Sinon on en crée un nouveau
    If t% = -1 Then
        ReDim Preserve ListeTunnel(n% + 1) As New ClassTunnel
        t% = n%
    End If
    '
    Call ListeTunnel(t%).Point_Ajoute(-50 + PosSouris.X * REDUCTION%, 50 + PosSouris.z * REDUCTION%)
    Call ListeTunnel(t%).Point_Ajoute(50 + PosSouris.X * REDUCTION%, 50 + PosSouris.z * REDUCTION%)
    Call ListeTunnel(t%).Point_Ajoute(50 + PosSouris.X * REDUCTION%, -50 + PosSouris.z * REDUCTION%)
    Call ListeTunnel(t%).Point_Ajoute(-50 + PosSouris.X * REDUCTION%, -50 + PosSouris.z * REDUCTION%)
    Call Initialisation.Vue_Decharge
    Call Vue.Calcule_Reseau
    Call Initialisation.Vue_Charge
End Sub

'
' **********************************
' Affiche un tunnel en surimpression
' **********************************
'
Public Sub Affiche(No%)
    Dim i%, i1%, n%
    Dim c&
    Dim P1(1) As D3DRMVECTOR4D
    Dim P2(1) As D3DVECTOR
    '
    n% = ListeTunnel(No%).Nb_Point%()
    For i% = 0 To n% - 1
        If ListeTunnel(No%).Face(i%) = True Then
            If ListeTunnel(No%).SegmentPointe% = i% Then
                c& = vbYellow
            Else
                c& = vbWhite
            End If
        Else
            If ListeTunnel(No%).SegmentPointe% = i% Then
                c& = RGB(255, 128, 64) ' Orange
            Else
                c& = RGB(64, 64, 64) ' Gris foncé
            End If
        End If
        '
        P2(0).X = ListeTunnel(No%).PositionX(i%) / REDUCTION%
        P2(0).Y = 0
        P2(0).z = ListeTunnel(No%).PositionZ(i%) / REDUCTION%
        Call dxVue.dxViewport.Transform(P1(0), P2(0))
        i1% = (i% + 1) Mod n%
        P2(1).X = ListeTunnel(No%).PositionX(i1%) / REDUCTION%
        P2(1).Y = 0
        P2(1).z = ListeTunnel(No%).PositionZ(i1%) / REDUCTION%
        Call dxVue.dxViewport.Transform(P1(1), P2(1))
        '
        Call dxVue.dxBack.SetForeColor(c&)
        Call dxVue.dxBack.DrawLine(P1(0).X, P1(0).Y, P1(1).X, P1(1).Y)
    Next i%
    '
    For i% = 0 To n% - 1
        If i% = ListeTunnel(No%).PointSelection% Then
            c& = vbRed
        Else
            If ListeTunnel(No%).PointPointe% = i% Then
                c& = vbCyan
            Else
                c& = vbBlue
            End If
        End If
        P2(0).X = (ListeTunnel(No%).PositionX(i%) - 3) / REDUCTION%
        P2(0).Y = 0
        P2(0).z = (ListeTunnel(No%).PositionZ(i%) + 3) / REDUCTION%
        Call dxVue.dxViewport.Transform(P1(0), P2(0))
        P2(1).X = (ListeTunnel(No%).PositionX(i%) + 3) / REDUCTION%
        P2(1).Y = 0
        P2(1).z = (ListeTunnel(No%).PositionZ(i%) - 3) / REDUCTION%
        Call dxVue.dxViewport.Transform(P1(1), P2(1))
        '
        Call dxVue.dxBack.SetForeColor(c&)
        Call dxVue.dxBack.DrawLine(P1(0).X, P1(0).Y, P1(1).X, P1(0).Y)
        Call dxVue.dxBack.DrawLine(P1(1).X, P1(0).Y, P1(1).X, P1(1).Y)
        Call dxVue.dxBack.DrawLine(P1(1).X, P1(1).Y, P1(0).X, P1(1).Y)
        Call dxVue.dxBack.DrawLine(P1(0).X, P1(1).Y, P1(0).X, P1(0).Y)
    Next i%
End Sub

'
' **************************
' Recherche le tunnel pointé
' **************************
'
Public Sub Tunnel_Pointe()
    Dim i%, n%
    n% = UBound(ListeTunnel())
    NoTunnel% = -1
    For i% = 0 To n% - 1
        If ListeTunnel(i%).Nb_Point% <> 0 Then
            If ListeTunnel(i%).Cherche_Pointe(PosSouris.X * REDUCTION%, -PosSouris.z * REDUCTION%) = True Then
                NoTunnel% = i%
            End If
        End If
    Next i%
End Sub

'
' ***************************
' Crée tous les mesh et frame
' ***************************
'
Public Sub Cree_Mesh()
    Dim n%, i%
    n% = UBound(ListeTunnel())
    ReDim FrameTunnel(n%) As Direct3DRMFrame3
    ReDim MeshTunnel(n%) As Direct3DRMMeshBuilder3
    For i% = 0 To n% - 1
        Call Tunnel.Genere(i%)
    Next i%
End Sub

'
' ********************************
' Détruit tous les meshs et frames
' ********************************
'
Public Sub Detruit_Mesh()
    Dim i%
    For i% = 0 To UBound(ListeTunnel()) - 1
        Call FrameTunnel(i%).DeleteVisual(MeshTunnel(i%))
        Set FrameTunnel(i%) = Nothing
        Set MeshTunnel(i%) = Nothing
    Next i%
End Sub

