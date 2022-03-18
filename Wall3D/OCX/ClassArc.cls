VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' *********************
' Draw a portion of arc
' *********************
'
Option Explicit
'
Dim ArcX%
Dim ArcY%
Dim ArcAngle1%
Dim ArcAngle2%
Dim ArcVertex%
Dim ArcRadius!
Dim ArcRed As Byte
Dim ArcGreen As Byte
Dim ArcBlue As Byte
Dim ArcAlpha As Byte
Dim ArcPoint() As D3DVECTOR
Dim vZ As D3DVECTOR

'
' *********************
' Test if the arc exist
' at X,Y position
' *********************
'
Public Function TestHit(ByVal X%, ByVal Y%) As Boolean
    Dim a!, b!
    Dim i%
    Dim p1 As D3DVECTOR
    Dim p2 As D3DVECTOR
    TestHit = False
    For i% = 1 To ArcVertex%
        If ArcPoint(i% - 1).X < ArcPoint(i%).X Then
            p1.X = ArcPoint(i% - 1).X
            p2.X = ArcPoint(i%).X
        Else
            p2.X = ArcPoint(i% - 1).X
            p1.X = ArcPoint(i%).X
        End If
        If ArcPoint(i% - 1).Y < ArcPoint(i%).Y Then
            p1.Y = ArcPoint(i% - 1).Y
            p2.Y = ArcPoint(i%).Y
        Else
            p2.Y = ArcPoint(i% - 1).Y
            p1.Y = ArcPoint(i%).Y
        End If
        If p1.X - 1 <= X% And X% <= p2.X + 1 Then
            If p1.Y - 1 <= Y% And Y% <= p2.Y + 1 Then
                If ArcPoint(i% - 1).X = ArcPoint(i%).X Then
                    TestHit = True
                Else
                    a! = (ArcPoint(i% - 1).Y - ArcPoint(i%).Y) / (ArcPoint(i% - 1).X - ArcPoint(i%).X)
                    b! = ArcPoint(i%).Y - a! * ArcPoint(i%).X
                    If Abs(a!) < 1 Then
                        If Abs((a! * X% + b! - Y%) * a!) < 1 Then TestHit = True
                    Else
                        If Abs((a! * X% + b! - Y%) / a!) < 1 Then TestHit = True
                    End If
                End If
            End If
        End If
    Next i%
End Function

'
' ********************************
' Retrieve the normal for a vertex
' ********************************
'
Public Sub GetNormal(Index%, Apply As D3DVECTOR, Normal As D3DVECTOR)
    If Index% < 1 Then Exit Sub
    If Index% > UBound(ArcPoint()) Then Exit Sub
    Dim BaseVector As D3DVECTOR
    Dim dX7 As New DirectX7
    Apply.X = (ArcPoint(Index%).X + ArcPoint(Index% - 1).X) / 2
    Apply.Y = (ArcPoint(Index%).Y + ArcPoint(Index% - 1).Y) / 2
    BaseVector.X = ArcPoint(Index%).X - ArcPoint(Index% - 1).X
    BaseVector.Y = ArcPoint(Index%).Y - ArcPoint(Index% - 1).Y
    Call dX7.VectorRotate(Normal, BaseVector, vZ, -2 * Atn(1))
    Set dX7 = Nothing
End Sub

'
' ********************************
' Get X value for this index point
' ********************************
'
Public Function GetX!(Index%)
    If Index% < 0 Then Exit Function
    If Index% > UBound(ArcPoint()) Then Exit Function
    GetX! = ArcPoint(Index%).X
End Function

'
' ********************************
' Get Y value for this index point
' ********************************
'
Public Function GetY!(Index%)
    If Index% < 0 Then Exit Function
    If Index% > UBound(ArcPoint()) Then Exit Function
    GetY! = ArcPoint(Index%).Y
End Function

Public Property Get X%()
    X% = ArcX%
End Property

Public Property Let X(n%)
    ArcX% = n%
    Call Generate
End Property

Public Property Get Y%()
    Y% = ArcY%
End Property

Public Property Let Y(n%)
    ArcY% = n%
    Call Generate
End Property

Public Property Get Angle1%()
    Angle1% = ArcAngle1%
End Property

Public Property Let Angle1(n%)
    ArcAngle1% = n%
    Call Generate
End Property

Public Property Get Angle2%()
    Angle2% = ArcAngle2%
End Property

Public Property Let Angle2(n%)
    ArcAngle2% = n%
    Call Generate
End Property

Public Property Get Vertex%()
    Vertex% = ArcVertex%
End Property

Public Property Let Vertex(n%)
    ArcVertex% = n%
    If ArcVertex% < 1 Then ArcVertex% = 1
    Call Generate
End Property

Public Property Get Red() As Byte
    Red = ArcRed
End Property

Public Property Let Red(n As Byte)
    ArcRed = n
End Property

Public Property Get Green() As Byte
    Green = ArcGreen
End Property

Public Property Let Green(n As Byte)
    ArcGreen = n
End Property

Public Property Get Blue() As Byte
    Blue = ArcBlue
End Property

Public Property Let Blue(n As Byte)
    ArcBlue = n
End Property

Public Property Get Alpha() As Byte
    Alpha = ArcAlpha
End Property

Public Property Let Alpha(n As Byte)
    ArcAlpha = n
End Property

'
' **************
' Initiale value
' **************
'
Private Sub Class_Initialize()
    ArcVertex = 1
    ReDim ArcPoint(ArcVertex%) As D3DVECTOR
    vZ.z = 1
End Sub

'
' **********************
' Generate list of point
' **********************
'
Private Sub Generate()
    Dim i%
    Dim PI#
    Dim Angle!
    ReDim ArcPoint(ArcVertex%) As D3DVECTOR
    If ArcAngle1% < ArcAngle2% Then Exit Sub
    PI# = 4 * Atn(1)
    For i% = 0 To ArcVertex%
        Angle! = ArcAngle1% + i% * (ArcAngle2% - ArcAngle1%) / ArcVertex%
        ArcPoint(i%).X = ArcX% + ArcRadius! * Cos(Angle! * PI# / 180)
        ArcPoint(i%).Y = ArcY% - ArcRadius! * Sin(Angle! * PI# / 180)
    Next i%
End Sub

Public Property Get Radius!()
    Radius! = ArcRadius!
End Property

Public Property Let Radius(n!)
    ArcRadius! = n!
    Call Generate
End Property

'
' **********************
' Draw the arc on screen
' **********************
'
Public Sub Draw(Zone As Object)
    Dim i%, ArcColor&
    ArcColor& = RGB(ArcRed, ArcGreen, ArcBlue)
    Zone.CurrentX = ArcPoint(0).X
    Zone.CurrentY = ArcPoint(0).Y
    For i% = 1 To ArcVertex%
        Zone.Line -(ArcPoint(i%).X, ArcPoint(i%).Y), ArcColor&
    Next i%
End Sub

