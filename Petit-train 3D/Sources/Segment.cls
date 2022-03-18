VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
Dim SegmentLongueur As Single
Dim SegmentRayon As Single
Dim SegmentAngle As Single
Dim SegmentRotation As Single
'
Dim SegmentDroit As Boolean
Dim SegmentTaille As Single
'
Dim dX7 As New DirectX7
Dim DegRad As Single
Dim AxeY As D3DVECTOR

'
' ****************************
' Gestion d'un segment de rail
' ****************************
'
Private Sub Class_Initialize()
    DegRad! = Atn(1) / 45 ' Convertion Degrée en Radian
    AxeY.X = 0
    AxeY.Y = 1
    AxeY.z = 0
End Sub

Public Property Get Rayon() As Single
    Rayon! = SegmentRayon!
End Property

Public Property Let Rayon(r As Single)
    SegmentRayon! = r!
    Call Calcul_Taille
End Property

Public Property Get Longueur() As Single
    Longueur! = SegmentLongueur!
End Property

Public Property Let Longueur(l As Single)
    SegmentLongueur! = l!
    Call Calcul_Taille
End Property

Public Property Get Angle() As Single
    Angle! = SegmentAngle!
End Property

Public Property Let Angle(a As Single)
    SegmentAngle! = a!
    Call Calcul_Taille
End Property

Public Function Segment_Taille() As Single
    Segment_Taille! = Abs(SegmentTaille!)
End Function

Private Sub Calcul_Taille()
    If SegmentRayon! = 0 Or SegmentAngle! = 0 Then
        SegmentTaille! = SegmentLongueur!
        SegmentDroit = True
    Else
        SegmentTaille! = SegmentRayon! * SegmentAngle! * DegRad!
        SegmentDroit = False
    End If
End Sub

Public Function Point(Position As Single) As D3DVECTOR
    Dim P!, l!
    If SegmentDroit = True Then
        Point.X = Position!
        Point.z = 0
    Else
        If SegmentTaille! = 0 Then
            P! = 0
        Else
            P! = SegmentAngle! * DegRad! * Position! / SegmentTaille!
        End If
        Point.X = SegmentRayon! * Sin(P!)
        Point.z = SegmentRayon! * (Cos(P!) - 1)
    End If
    Point.Y = 0
    l! = dX7.VectorModulus(Point)
    Call dX7.VectorRotate(Point, Point, AxeY, SegmentRotation! * DegRad!)
    Call dX7.VectorScale(Point, Point, l!)
End Function

Public Function Segment_Droit() As Boolean
    Segment_Droit = SegmentDroit
End Function

Public Function Tangente(Position As Single) As D3DVECTOR
    Dim P!, l!
    If SegmentDroit = True Then
        Tangente.X = 1
        Tangente.z = 0
    Else
        If SegmentTaille! = 0 Then
            P! = 0
        Else
            P! = SegmentAngle! * DegRad! * Position! / SegmentTaille!
        End If
        If SegmentAngle! < 0 Then
            Tangente.X = -Cos(P!)
            Tangente.z = Sin(P!)
        Else
            Tangente.X = Cos(P!)
            Tangente.z = -Sin(P!)
        End If
    End If
    Tangente.Y = 0
    l! = dX7.VectorModulus(Tangente)
    Call dX7.VectorRotate(Tangente, Tangente, AxeY, SegmentRotation! * DegRad!)
    Call dX7.VectorScale(Tangente, Tangente, l!)
End Function

Public Property Get Rotation() As Single
    Rotation! = SegmentRotation!
End Property

Public Property Let Rotation(r As Single)
    SegmentRotation! = r!
End Property

Public Function Theta!(Position As Single)
    Dim P!, l!
    If SegmentDroit = True Then
        Theta! = 0
    Else
        If SegmentTaille! = 0 Then
            Theta! = 0
        Else
            Theta! = SegmentAngle! * Position! / SegmentTaille!
        End If
    End If
    Theta! = Theta! + SegmentRotation!
    If Theta! > 360 Then Theta! = Theta! - 360
End Function

