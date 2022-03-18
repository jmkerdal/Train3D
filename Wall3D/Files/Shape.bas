Attribute VB_Name = "dxShape"
Option Explicit
'
Type TypeShape
    X1 As Single
    Y1 As Single
    X2 As Single
    Y2 As Single
    Color As PALETTEENTRY
End Type
Global ListPoint() As TypeShape

'
' **********************************
' Add a new Shape to the MeshBuilder
' **********************************
'
Public Sub Add_Shape(Number%, List() As TypeShape, Angle%, Face%, Inversion As Boolean, Closed As Boolean)
    Dim p(3) As D3DVECTOR
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim i%, n%, pp%
    Dim TheStep!
    TheStep! = Angle% / Face%
    v2 = AxeX ' Initial vector start value
    For i% = 1 To Face%
        v1 = v2
        Call dxEngine.dX7.VectorRotate(v2, AxeX, AxeY, -i% * TheStep! / 180 * PI!)
        For n% = 0 To UBound(List())
            p(0).X = List(n%).X1! * v1.X
            p(0).Y = List(n%).Y1!
            p(0).z = List(n%).X1! * v1.z
            p(1).X = List(n%).X1! * v2.X
            p(1).Y = List(n%).Y1!
            p(1).z = List(n%).X1! * v2.z
            p(2).X = List(n%).X2! * v2.X
            p(2).Y = List(n%).Y2!
            p(2).z = List(n%).X2! * v2.z
            p(3).X = List(n%).X2! * v1.X
            p(3).Y = List(n%).Y2!
            p(3).z = List(n%).X2! * v1.z
            If p(2).X = p(3).X And p(2).Y = p(3).Y And p(2).z = p(3).z Then
                pp% = 3
            Else
                pp% = 4
            End If
            If p(0).X = p(1).X And p(0).Y = p(1).Y And p(0).z = p(1).z Then
                pp% = pp% - 1
                p(1) = p(2)
                p(2) = p(3)
            End If
            If pp% >= 3 Then
                Call Face_Add(Number%, List(n%).Color, 0, pp%, p(), Inversion)
                If Closed = True And Angle% <> 360 And (i% = 1 Or i% = Face%) Then
                    If i% = 1 Then
                        p(1) = p(3)
                        p(2).X = 0: p(2).Y = 0: p(2).z = 0
                    Else
                    'If i% = Face% Then
                        If pp% = 4 Then
                            p(0) = p(1)
                            p(1).X = 0: p(1).Y = 0: p(1).z = 0
                        Else
                            p(2) = p(1)
                            p(1) = p(0)
                            p(0) = p(2)
                            p(2).X = 0: p(2).Y = 0: p(2).z = 0
                        End If
                    End If
                    If p(0).Y <> p(1).Y Or p(1).Y <> p(2).Y Then
                        ' Only if the vertex is not horizontal
                        Call Face_Add(Number%, List(n%).Color, 0, 3, p(), Inversion)
                    End If
                End If
            End If
        Next n%
    Next i%
End Sub

'
' **********************
' Add a cube to the mesh
' **********************
'
Public Sub Add_Cube(n%, Color As PALETTEENTRY, Sens As Boolean)
    Dim p(3) As D3DVECTOR
    p(0).X = -0.5: p(0).Y = 0.5: p(0).z = 0.5
    p(1).X = 0.5: p(1).Y = 0.5: p(1).z = 0.5
    p(2).X = 0.5: p(2).Y = 0.5: p(2).z = -0.5
    p(3).X = -0.5: p(3).Y = 0.5: p(3).z = -0.5
    Call Tools3D.Face_Add(n%, Color, 0, 4, p(), Sens)
    p(0).X = -0.5: p(0).Y = 0.5: p(0).z = -0.5
    p(1).X = 0.5: p(1).Y = 0.5: p(1).z = -0.5
    p(2).X = 0.5: p(2).Y = -0.5: p(2).z = -0.5
    p(3).X = -0.5: p(3).Y = -0.5: p(3).z = -0.5
    Call Tools3D.Face_Add(n%, Color, 0, 4, p(), Sens)
    p(0).X = -0.5: p(0).Y = 0.5: p(0).z = 0.5
    p(1).X = -0.5: p(1).Y = 0.5: p(1).z = -0.5
    p(2).X = -0.5: p(2).Y = -0.5: p(2).z = -0.5
    p(3).X = -0.5: p(3).Y = -0.5: p(3).z = 0.5
    Call Tools3D.Face_Add(n%, Color, 0, 4, p(), Sens)
    p(0).X = 0.5: p(0).Y = 0.5: p(0).z = 0.5
    p(1).X = -0.5: p(1).Y = 0.5: p(1).z = 0.5
    p(2).X = -0.5: p(2).Y = -0.5: p(2).z = 0.5
    p(3).X = 0.5: p(3).Y = -0.5: p(3).z = 0.5
    Call Tools3D.Face_Add(n%, Color, 0, 4, p(), Sens)
    p(0).X = 0.5: p(0).Y = 0.5: p(0).z = -0.5
    p(1).X = 0.5: p(1).Y = 0.5: p(1).z = 0.5
    p(2).X = 0.5: p(2).Y = -0.5: p(2).z = 0.5
    p(3).X = 0.5: p(3).Y = -0.5: p(3).z = -0.5
    Call Tools3D.Face_Add(n%, Color, 0, 4, p(), Sens)
    p(0).X = -0.5: p(0).Y = -0.5: p(0).z = -0.5
    p(1).X = 0.5: p(1).Y = -0.5: p(1).z = -0.5
    p(2).X = 0.5: p(2).Y = -0.5: p(2).z = 0.5
    p(3).X = -0.5: p(3).Y = -0.5: p(3).z = 0.5
    Call Tools3D.Face_Add(n%, Color, 0, 4, p(), Sens)
End Sub

'
' ************
' Add a sphere
' ************
'
Public Sub Add_Sphere(n%, Faces%, Angle%, Color As PALETTEENTRY, Sens As Boolean, Closed As Boolean)
    Dim Va As D3DVECTOR
    Dim p(3) As D3DVECTOR
    Dim ni%, i!, j!
    Dim pp%
    Dim TheStep!, TheStep2!
    TheStep! = Angle% / Faces%
    TheStep2! = 180 / Faces%
    For ni% = 0 To Faces% - 1
        i! = ni% * Angle% / Faces%
        For j! = -90 To 90 - TheStep2! Step TheStep2!
            Call dxEngine.dX7.VectorRotate(Va, AxeX, AxeZ, j! / 180 * PI!)
            Call dxEngine.dX7.VectorRotate(p(0), Va, AxeY, -i! / 180 * PI!)
            Call dxEngine.dX7.VectorRotate(Va, AxeX, AxeZ, (j! + TheStep2!) / 180 * PI!)
            Call dxEngine.dX7.VectorRotate(p(1), Va, AxeY, -i! / 180 * PI!)
            Call dxEngine.dX7.VectorRotate(Va, AxeX, AxeZ, (j! + TheStep2!) / 180 * PI!)
            Call dxEngine.dX7.VectorRotate(p(2), Va, AxeY, -(i! + TheStep!) / 180 * PI!)
            Call dxEngine.dX7.VectorRotate(Va, AxeX, AxeZ, j! / 180 * PI!)
            Call dxEngine.dX7.VectorRotate(p(3), Va, AxeY, -(i! + TheStep!) / 180 * PI!)
            If j! <> -90 And j! + TheStep2! <> 90 Then
                pp% = 4
                Call Face_Add(n%, Color, 0, pp%, p(), Sens)
            Else
                pp% = 3
                If j! = -90 Then
                    p(0).X = 0: p(0).Y = -1: p(0).z = 0
                Else
                    p(2) = p(3)
                    p(1).X = 0: p(1).Y = 1: p(1).z = 0
                End If
                Call Face_Add(n%, Color, 0, pp%, p(), Sens)
            End If
            If Closed = True And Angle% <> 360 And (ni% = 0 Or ni% = Faces% - 1) Then
                If ni% = 0 Then
                    p(2) = p(1)
                    p(1).X = 0: p(1).Y = 0: p(1).z = 0
                Else
                    If pp% = 4 Then
                        p(1) = p(3)
                        p(0).X = 0: p(0).Y = 0: p(0).z = 0
                    Else
                        If j! = -90 Then
                            p(1) = p(2)
                            p(2).X = 0: p(2).Y = 0: p(2).z = 0
                        Else
                            p(0) = p(2)
                            p(2).X = 0: p(2).Y = 0: p(2).z = 0
                        End If
                    End If
                End If
                If p(0).Y <> p(1).Y Or p(1).Y <> p(2).Y Then
                    ' Only if the vertex is not horizontal
                    Call Face_Add(n%, Color, 0, 3, p(), Sens)
                End If
            End If
        Next j!
    Next ni%
End Sub

