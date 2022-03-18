VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Curve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' *********************************************************
' This Class allow you to create and display curves
' I have test this module in a OCX but I have so much
' problem that I decide to use it as a Class only
' I think it will be usefull to be able to move
' a curve as the line control, but it's seems not possible.
' *********************************************************
'
Option Explicit
Dim CurveX1%, CurveY1%
Dim CurveX2%, CurveY2%
Dim CurveVertex% ' Number of vertex
Dim CurveRadius! ' Radius of the curve
Dim CurveRed As Byte
Dim CurveBlue As Byte
Dim CurveGreen As Byte
Dim CurveAlpha As Byte
Dim CurvePoint() As D3DVECTOR
Dim vX As D3DVECTOR
Dim vZ As D3DVECTOR

'
' ***************
' Get the start X
' ***************
'
Public Property Get X1%()
    X1% = CurveX1%
End Property

'
' *******
' Start X
' *******
'
Public Property Let X1(ByVal n%)
    CurveX1% = n%
    Call Generate
End Property

'
' *************
' Get the end Y
' *************
'
Public Property Get Y1%()
    Y1% = CurveY1%
End Property

'
' *******
' Start Y
' *******
'
Public Property Let Y1(ByVal n%)
    CurveY1% = n%
    Call Generate
End Property

'
' **********
' Get radius
' **********
'
Public Property Get Radius!()
    Radius! = CurveRadius!
End Property

'
' *************
' Change radius
' *************
'
Public Property Let Radius(ByVal n!)
    CurveRadius! = n!
    Call Generate
End Property

'
' ********************
' Get number of vertex
' ********************
'
Public Property Get Vertex%()
    Vertex% = CurveVertex%
End Property

'
' ***********************
' Change number of vertex
' ***********************
'
Public Property Let Vertex(ByVal n%)
    If n% < 1 Then n% = 1
    CurveVertex% = n%
    Call Generate
End Property

'
' ****************************************
' Redraw the form each time a value change
' ****************************************
'
Private Sub Generate()
    Dim i%
    Dim vResult1 As D3DVECTOR
    Dim vResult2 As D3DVECTOR
    Dim Ray!
    Dim Slant#
    Dim PI#
    Dim Start As D3DVECTOR
    Dim dX%, dY%
    Dim dX7 As New DirectX7
    If CurveVertex% = 0 Then Exit Sub ' Not ready
    If CurveX1% = CurveX2% And CurveY1% = CurveY2% Then Exit Sub ' Nothing to do
    '
    ' Intermediate calculation
    '
    PI# = 4 * Atn(1)
    dX% = CurveX2% - CurveX1%
    dY% = CurveY2% - CurveY1%
    If dX% = 0 Then
        If dY% < 0 Then
            Slant# = PI# / 2
        Else
            Slant# = 3 * PI# / 2
        End If
    Else
        Slant# = Atn(dY% / dX%)
    End If
    Ray! = Sqr(dX% ^ 2 + dY% ^ 2) / 2
    '
    ' Fill the array of point
    '
    ReDim CurvePoint(CurveVertex%) As D3DVECTOR
    For i% = 0 To CurveVertex%
        Call dX7.VectorRotate(vResult1, vX, vZ, PI# / CurveVertex% * i%)
        If dX% > 0 Then
            vResult1.X = vResult1.X * -Ray!
            vResult1.Y = vResult1.Y * -CurveRadius!
        Else
            vResult1.X = vResult1.X * Ray!
            vResult1.Y = vResult1.Y * CurveRadius!
        End If
        Call dX7.VectorRotate(vResult2, vResult1, vZ, Slant#)
        Call dX7.VectorScale(CurvePoint(i%), vResult2, dX7.VectorModulus(vResult1))
        If i% = 0 Then Start = CurvePoint(0)
        CurvePoint(i%).X = CurveX1% + CurvePoint(i%).X - Start.X
        CurvePoint(i%).Y = CurveY1% + CurvePoint(i%).Y - Start.Y
    Next i%
    Set dX7 = Nothing
End Sub

'
' ********************************
' Get X value for this index point
' ********************************
'
Public Function GetX!(Index%)
    If Index% < 0 Then Exit Function
    If Index% > UBound(CurvePoint()) Then Exit Function
    GetX! = CurvePoint(Index%).X
End Function

'
' ********************************
' Get Y value for this index point
' ********************************
'
Public Function GetY!(Index%)
    If Index% < 0 Then Exit Function
    If Index% > UBound(CurvePoint()) Then Exit Function
    GetY! = CurvePoint(Index%).Y
End Function

'
' ***********************
' Test if the curve exist
' at X,Y position
' ***********************
'
Public Function TestHit(ByVal X%, ByVal Y%) As Boolean
    Dim a!, b!
    Dim i%
    Dim p1 As D3DVECTOR
    Dim p2 As D3DVECTOR
    TestHit = False
    For i% = 1 To CurveVertex%
        If CurvePoint(i% - 1).X < CurvePoint(i%).X Then
            p1.X = CurvePoint(i% - 1).X
            p2.X = CurvePoint(i%).X
        Else
            p2.X = CurvePoint(i% - 1).X
            p1.X = CurvePoint(i%).X
        End If
        If CurvePoint(i% - 1).Y < CurvePoint(i%).Y Then
            p1.Y = CurvePoint(i% - 1).Y
            p2.Y = CurvePoint(i%).Y
        Else
            p2.Y = CurvePoint(i% - 1).Y
            p1.Y = CurvePoint(i%).Y
        End If
        If p1.X - 1 <= X% And X% <= p2.X + 1 Then
            If p1.Y - 1 <= Y% And Y% <= p2.Y + 1 Then
                If CurvePoint(i% - 1).X = CurvePoint(i%).X Then
                    TestHit = True
                Else
                    a! = (CurvePoint(i% - 1).Y - CurvePoint(i%).Y) / (CurvePoint(i% - 1).X - CurvePoint(i%).X)
                    b! = CurvePoint(i%).Y - a! * CurvePoint(i%).X
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
' ************************
' Draw the curve on screen
' ************************
'
Public Sub Draw(Zone As Object)
    Dim i%, TheColor&
    If CurveX1% = CurveX2% And CurveY1% = CurveY2% Then Exit Sub
    TheColor& = RGB(CurveRed, CurveGreen, CurveBlue)
    Zone.CurrentX = CurvePoint(0).X
    Zone.CurrentY = CurvePoint(0).Y
    For i% = 1 To CurveVertex%
        Zone.Line -(CurvePoint(i%).X, CurvePoint(i%).Y), TheColor&
    Next i%
End Sub

'
' **************
' Initiale value
' **************
'
Private Sub Class_Initialize()
    CurveX1 = 0
    CurveY1 = 0
    CurveX2 = 0
    CurveY2 = 0
    CurveVertex = 1
    CurveRadius = 0
    vX.X = 1
    vZ.z = 1
    ReDim CurvePoint(CurveVertex%) As D3DVECTOR
End Sub

'
' *******
' Start X
' *******
'
Public Property Get X2%()
    X2% = CurveX2%
End Property

'
' *****
' End X
' *****
'
Public Property Let X2(ByVal n%)
    CurveX2% = n%
    Call Generate
End Property

'
' *****
' End Y
' *****
'
Public Property Get Y2%()
    Y2% = CurveY2%
End Property

'
' *****
' End Y
' *****
'
Public Property Let Y2(ByVal n%)
    CurveY2% = n%
    Call Generate
End Property


'
' ***************
' Red color value
' ***************
'
Public Property Get Red() As Byte
    Red = CurveRed
End Property

Public Property Let Red(ByVal n As Byte)
    CurveRed = n
End Property

'
' *****************
' Green color value
' *****************
'
Public Property Get Green() As Byte
    Green = CurveGreen
End Property

Public Property Let Green(ByVal n As Byte)
    CurveGreen = n
End Property

'
' ****************
' Blue color value
' ****************
'
Public Property Get Blue() As Byte
    Blue = CurveBlue
End Property

Public Property Let Blue(ByVal n As Byte)
    CurveBlue = n
End Property


'
' *****************
' Alpha color value
' *****************
'
Public Property Get Alpha() As Byte
    Alpha = CurveAlpha
End Property

Public Property Let Alpha(ByVal n As Byte)
    CurveAlpha = n
End Property

'
' ********************************
' Retrieve the normal for a vertex
' ********************************
'
Public Sub GetNormal(Index%, Apply As D3DVECTOR, Normal As D3DVECTOR)
    If Index% < 1 Then Exit Sub
    If Index% > UBound(CurvePoint()) Then Exit Sub
    Dim BaseVector As D3DVECTOR
    Dim dX7 As New DirectX7
    Apply.X = (CurvePoint(Index%).X + CurvePoint(Index% - 1).X) / 2
    Apply.Y = (CurvePoint(Index%).Y + CurvePoint(Index% - 1).Y) / 2
    BaseVector.X = CurvePoint(Index%).X - CurvePoint(Index% - 1).X
    BaseVector.Y = CurvePoint(Index%).Y - CurvePoint(Index% - 1).Y
    Call dX7.VectorRotate(Normal, BaseVector, vZ, -2 * Atn(1))
    Set dX7 = Nothing
End Sub

