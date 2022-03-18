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
'
Private Type TypeTunnel
    LaPosition As D3DVECTOR
    LaNormal As D3DVECTOR
    LaFace As Boolean
    Utilise As Boolean
End Type
'
Dim Tunnel() As TypeTunnel
Public Nb_Point%
Public PointSelection%
Public PointPointe%
Public SegmentPointe%
'
Dim PI#
Dim LeTriangle() As D3DVECTOR
Public Nb_Triangle%

Public Sub Point_Ajoute(ByVal X!, ByVal z!)
    ReDim Preserve Tunnel(Nb_Point% + 1) As TypeTunnel
    Tunnel(Nb_Point%).LaPosition.X = X!
    Tunnel(Nb_Point%).LaPosition.z = z!
    Tunnel(Nb_Point%).LaFace = True
    Nb_Point% = Nb_Point% + 1
End Sub

Public Sub Point_Deplace(n%, ByVal X!, ByVal z!)
    If n% = -1 Then Exit Sub
    Tunnel(n%).LaPosition.X = X!
    Tunnel(n%).LaPosition.z = z!
End Sub

'
' ************************************
' Cherche les éléments que l'on pointe
' ************************************
'
Public Function Cherche_Pointe(ByVal X!, ByVal Y!) As Boolean
    Dim i%
    Dim x1!, y1!
    Dim x2!, y2!
    Dim t!
    '
    SegmentPointe% = -1
    PointPointe% = -1
    For i% = 0 To Nb_Point% - 1
        x1! = Tunnel(i%).LaPosition.X
        y1! = -Tunnel(i%).LaPosition.z
        x2! = Tunnel((i% + 1) Mod Nb_Point%).LaPosition.X
        y2! = -Tunnel((i% + 1) Mod Nb_Point%).LaPosition.z
        '
        ' ***** Test pointe point
        '
        If Abs(x1! - X!) <= 4 And Abs(y1! - Y!) <= 4 Then
            PointPointe% = i%
            Cherche_Pointe = True
        End If
        '
        ' ***** Test pointe segment
        '
        If x1! > x2! Then
            t! = x1!
            x1! = x2!
            x2! = t!
        End If
        If y1! > y2! Then
            t! = y1!
            y1! = y2!
            y2! = t!
        End If
        If x1! <= X! And X! <= x2! And y1! <= Y! And Y! <= y2! Then
            SegmentPointe% = i%
            Cherche_Pointe = True
        End If
    Next i%
End Function

Public Sub Point_Insere()
    Dim i%
    ReDim Preserve Tunnel(Nb_Point% + 1) As TypeTunnel
    For i% = Nb_Point% - 1 To SegmentPointe% + 1 Step -1
        Tunnel(i% + 1) = Tunnel(i%)
    Next i%
    Tunnel(SegmentPointe% + 1).LaPosition.X = (Tunnel(SegmentPointe%).LaPosition.X + Tunnel((SegmentPointe% + 2) Mod (Nb_Point% + 1)).LaPosition.X) / 2
    Tunnel(SegmentPointe% + 1).LaPosition.z = (Tunnel(SegmentPointe%).LaPosition.z + Tunnel((SegmentPointe% + 2) Mod (Nb_Point% + 1)).LaPosition.z) / 2
    Tunnel(SegmentPointe% + 1).LaFace = True
    Nb_Point% = Nb_Point% + 1
End Sub

Public Sub Point_Supprime()
    Dim i%
    For i% = SegmentPointe% To Nb_Point% - 1
        Tunnel(i%) = Tunnel(i% + 1)
    Next i%
    ReDim Preserve Tunnel(Nb_Point% - 1) As TypeTunnel
    Nb_Point% = Nb_Point% - 1
End Sub

'
' **************************
' Inverse une face Coté/Bout
' **************************
'
Public Sub Inverse()
    Tunnel(SegmentPointe%).LaFace = Not Tunnel(SegmentPointe%).LaFace
End Sub

'
' *****************************
' Genere les triangles
' Calcul des normals aux points
' *****************************
'
Public Function Genere() As Boolean
    Dim i%, n%
    Dim i1%, i2%
    Dim Angle%, Angle1%, Angle2%
    Dim DegRad!
    '
    Do
        Genere = Genere_Triangle(n%)
        If Genere = True Then Exit Do
        n% = n% + 1
        If n% = Nb_Point% Then Exit Do
    Loop
    '
    DegRad! = Atn(1) / 45
    For i% = 0 To Nb_Point% - 1
        i1% = (i% + 1) Mod Nb_Point%
        i2% = (i% + 2) Mod Nb_Point%
        '
        Angle1% = Calcul_Angle%(Tunnel(i1%).LaPosition.X, Tunnel(i1%).LaPosition.z, Tunnel(i%).LaPosition.X, Tunnel(i%).LaPosition.z)
        Angle2% = Calcul_Angle%(Tunnel(i1%).LaPosition.X, Tunnel(i1%).LaPosition.z, Tunnel(i2%).LaPosition.X, Tunnel(i2%).LaPosition.z)
        If Tunnel(i%).LaFace = False And Tunnel(i1%).LaFace = True Then
            Angle% = Angle2% + 90
        ElseIf Tunnel(i1%).LaFace = False And Tunnel(i%).LaFace = True Then
            Angle% = Angle1% - 90
        Else
            Angle% = Angle1% - Angle2% ' Cherche le Delta entre Angle1 et Angle2
            If Angle% < 0 Then Angle% = Angle% + 360
'Debug.Print "a1"; Angle1%; "a2"; Angle2%; "Delta="; Angle%
            Angle% = Angle1% - Angle% / 2 ' Ajoute la moitié
        End If
        Tunnel(i1%).LaNormal.X = Cos(Angle% * DegRad!)
        Tunnel(i1%).LaNormal.z = Sin(Angle% * DegRad!)
'Debug.Print i1%; Tunnel(i1%).Normal.X; Tunnel(i1%).Normal.Y; Tunnel(i1%).Normal.z
    Next i%
End Function

'
' ************************************
' Génération de la liste des triangles
' ************************************
'
Public Function Genere_Triangle(Depart%) As Boolean
    Dim i%, q%, d%
    Dim P%(2)
    Dim a%(1)
    Dim Angle%
    '
    Nb_Triangle% = 0
    ReDim LeTriangle(2, Nb_Triangle%)
    q% = Nb_Point%
    For i% = 0 To Nb_Point% - 1
        Tunnel(i%).Utilise = False
    Next i%
    P%(0) = Depart%
    d% = P%(0)
    '
'Debug.Print "Départ en "; Depart%
    Do
        P%(1) = Point_Suivant%(P%(0))
        P%(2) = Point_Suivant%(P%(1))
        a%(0) = Calcul_Angle%(Tunnel(P%(0)).LaPosition.X, -Tunnel(P%(0)).LaPosition.z, Tunnel(P%(1)).LaPosition.X, -Tunnel(P%(1)).LaPosition.z)
        a%(1) = Calcul_Angle%(Tunnel(P%(1)).LaPosition.X, -Tunnel(P%(1)).LaPosition.z, Tunnel(P%(2)).LaPosition.X, -Tunnel(P%(2)).LaPosition.z)
        Angle% = a%(1) - a%(0)
        If Angle% > 360 Then Angle% = Angle% - 360
        If Angle% < 0 Then Angle% = Angle% + 360
'Debug.Print p%(0); p%(1); p%(2); Angle%
        If Angle% > 0 And Angle% < 180 Then
            ReDim Preserve LeTriangle(2, Nb_Triangle% + 1) As D3DVECTOR
            LeTriangle(0, Nb_Triangle%) = Tunnel(P%(0)).LaPosition
            LeTriangle(1, Nb_Triangle%) = Tunnel(P%(1)).LaPosition
            LeTriangle(2, Nb_Triangle%) = Tunnel(P%(2)).LaPosition
            Nb_Triangle% = Nb_Triangle% + 1
            '
            Tunnel(P%(1)).Utilise = True
            P%(0) = P%(2)
            d% = P%(2)
            q% = q% - 1
            If q% = 2 Then
                Genere_Triangle = True
                Exit Function ' Plus de triangle
            End If
        Else
            P%(0) = P%(1)
            If P%(0) = d% Then Exit Function ' Sortie sur erreur
        End If
        DoEvents
    Loop
End Function

'
' **************************************
' Recherche le point suivant non utilisé
' **************************************
'
Private Function Point_Suivant%(P%)
    Point_Suivant% = P%
    Do
        Point_Suivant% = (Point_Suivant% + 1) Mod Nb_Point%
        If Tunnel(Point_Suivant%).Utilise = False Then Exit Function
    Loop
End Function

'
' ************************************
' Retourne l'angle fait par le segment
' ************************************
'
Private Function Calcul_Angle%(ByVal x1!, ByVal z1!, ByVal x2!, ByVal z2!)
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

Private Sub Class_Initialize()
    Nb_Point% = 0
    ReDim Tunnel(Nb_Point%) As TypeTunnel
    Nb_Triangle% = 0
    ReDim LeTriangle(Nb_Triangle%) As D3DVECTOR
    PointSelection% = -1
    PointPointe% = -1
    SegmentPointe% = -1
    PI# = 4 * Atn(1)
End Sub

Public Property Get Face(n%) As Boolean
    Face = Tunnel(n%).LaFace
End Property

Public Property Let Face(n%, ByVal vNewValue As Boolean)
    Tunnel(n%).LaFace = vNewValue
End Property

Public Property Get PositionX!(n%)
    PositionX! = Tunnel(n%).LaPosition.X
End Property

Public Property Get PositionZ!(n%)
    PositionZ! = Tunnel(n%).LaPosition.z
End Property

Public Property Let PositionX(n%, ByVal vNewValue!)
    Tunnel(n%).LaPosition.X = vNewValue!
End Property

Public Property Let PositionZ(n%, ByVal vNewValue!)
    Tunnel(n%).LaPosition.z = vNewValue!
End Property

Public Property Get TriangleX!(i%, n%)
    TriangleX! = LeTriangle(i%, n%).X
End Property

Public Property Get TriangleZ!(i%, n%)
    TriangleZ! = LeTriangle(i%, n%).z
End Property

Public Property Let TriangleX(i%, n%, ByVal vNewValue!)
    LeTriangle(i%, n%).X = vNewValue!
End Property

Public Property Let TriangleZ(i%, n%, ByVal vNewValue!)
    LeTriangle(i%, n%).z = vNewValue!
End Property

Public Property Get NormalX!(n%)
    NormalX! = Tunnel(n%).LaNormal.X
End Property

Public Property Get NormalZ!(n%)
    NormalZ! = Tunnel(n%).LaNormal.z
End Property

Public Property Let NormalX(n%, ByVal vNewValue!)
    Tunnel(n%).LaNormal.X = vNewValue!
End Property

Public Property Let NormalZ(n%, ByVal vNewValue!)
    Tunnel(n%).LaNormal.z = vNewValue!
End Property

