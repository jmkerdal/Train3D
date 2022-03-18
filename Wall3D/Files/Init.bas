Attribute VB_Name = "Init"
Option Explicit
'
Global dxEngine As New ClassDirectX
Global XFileLoadFrame As Direct3DRMFrame3
'
Global FrameRedSpot As Direct3DRMFrame3
Global RedSpot As Direct3DRMLight
Global WallPath$ ' Wall definition path
Global TexturePath$ ' Wall textures path
Global AnimPath$ ' Animation path
Global Orthographic As Boolean ' Orthographic projection
Global TheCamera As D3DLIGHT7
Global AxeX As D3DVECTOR
Global AxeY As D3DVECTOR
Global AxeZ As D3DVECTOR
Global TheRed As Byte
Global TheGreen As Byte
Global TheBlue As Byte
Global TheAlpha As Byte
Global Animation&() ' Animation array
Global NbAnimation% ' Number of animation in memory
Global NoAnimation% ' Frame animation already playing
Global ColorBack As PALETTEENTRY
Global LogFile%
Global ImportX$
Global ImportWall$
Global ExportX$
'
' ***** Shadow effect
'
Global dxShadowLight As Direct3DRMLight
Global dxShadowLightFrame As Direct3DRMFrame3
'
Global Const NbSegment% = 100
Global Segment(NbSegment%) As New Curve
Global NbFaceSegment%
Global StartX%
Global StartY%
'
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'
' *******************************
' Initial values for all segments
' *******************************
'
Public Sub Shape_Init()
    Dim i%
    ' Starting point
    StartX% = 8
    StartY% = 200
    ' Number of face
    NbFaceSegment% = 10
    ' Segments
    For i% = 1 To NbSegment%
        Segment(i%).X1 = StartX%
        Segment(i%).Y1 = StartY%
        Segment(i%).X2 = StartX%
        Segment(i%).Y2 = StartY%
        Segment(i%).Red = 255
        Segment(i%).Green = 255
        Segment(i%).Blue = 255
        Segment(i%).Alpha = 255
        Segment(i%).Vertex = 1
        Segment(i%).Radius = 0
    Next i%
    ' Default color
    Shape.FaceColor.Red = 255
    Shape.FaceColor.Green = 255
    Shape.FaceColor.Blue = 255
    Shape.FaceColor.Alpha = 255
End Sub

'
' **********************
' Create Axes for camera
' **********************
'
Public Sub Create_Axes()
    Dim p1 As D3DVECTOR
    Dim p2 As D3DVECTOR
    Dim p3 As D3DVECTOR
    Dim p4 As D3DVECTOR
    Dim Color&
    Dim Thick!
    '
    AxeX.X = 1: AxeX.Y = 0: AxeX.z = 0
    AxeY.X = 0: AxeY.Y = 1: AxeY.z = 0
    AxeZ.X = 0: AxeZ.Y = 0: AxeZ.z = 1
    Thick! = 0.03
    '
    Call dxConstruitMeshBuilder.Init
    Color& = &HFF00FFFF
    '
    p1.X = 0: p1.Y = Thick!: p1.z = Thick!
    p2.X = 0: p2.Y = Thick!: p2.z = -Thick!
    p3.X = 10: p3.Y = Thick!: p3.z = -Thick!
    p4.X = 10: p4.Y = Thick!: p4.z = Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = 10: p1.Y = Thick!: p1.z = Thick!
    p2.X = 10: p2.Y = Thick!: p2.z = -Thick!
    p3.X = 10: p3.Y = -Thick!: p3.z = -Thick!
    p4.X = 10: p4.Y = -Thick!: p4.z = Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = 0: p1.Y = Thick!: p1.z = Thick!
    p2.X = 10: p2.Y = Thick!: p2.z = Thick!
    p3.X = 10: p3.Y = -Thick!: p3.z = Thick!
    p4.X = 0: p4.Y = -Thick!: p4.z = Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = 0: p1.Y = Thick!: p1.z = -Thick!
    p2.X = 0: p2.Y = Thick!: p2.z = Thick!
    p3.X = 0: p3.Y = -Thick!: p3.z = Thick!
    p4.X = 0: p4.Y = -Thick!: p4.z = -Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = 10: p1.Y = Thick!: p1.z = -Thick!
    p2.X = 0: p2.Y = Thick!: p2.z = -Thick!
    p3.X = 0: p3.Y = -Thick!: p3.z = -Thick!
    p4.X = 10: p4.Y = -Thick!: p4.z = -Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = 10: p1.Y = -Thick!: p1.z = Thick!
    p2.X = 10: p2.Y = -Thick!: p2.z = -Thick!
    p3.X = 0: p3.Y = -Thick!: p3.z = -Thick!
    p4.X = 0: p4.Y = -Thick!: p4.z = Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = 0: p1.z = Thick!
    p2.X = Thick!: p2.Y = 0: p2.z = Thick!
    p3.X = Thick!: p3.Y = 10: p3.z = Thick!
    p4.X = -Thick!: p4.Y = 10: p4.z = Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = 10: p1.z = Thick!
    p2.X = Thick!: p2.Y = 10: p2.z = Thick!
    p3.X = Thick!: p3.Y = 10: p3.z = -Thick!
    p4.X = -Thick!: p4.Y = 10: p4.z = -Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = 0: p1.z = Thick!
    p2.X = -Thick!: p2.Y = 10: p2.z = Thick!
    p3.X = -Thick!: p3.Y = 10: p3.z = -Thick!
    p4.X = -Thick!: p4.Y = 0: p4.z = -Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = Thick!: p1.Y = 0: p1.z = Thick!
    p2.X = -Thick!: p2.Y = 0: p2.z = Thick!
    p3.X = -Thick!: p3.Y = 0: p3.z = -Thick!
    p4.X = Thick!: p4.Y = 0: p4.z = -Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = Thick!: p1.Y = 10: p1.z = Thick!
    p2.X = Thick!: p2.Y = 0: p2.z = Thick!
    p3.X = Thick!: p3.Y = 0: p3.z = -Thick!
    p4.X = Thick!: p4.Y = 10: p4.z = -Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = 10: p1.z = -Thick!
    p2.X = Thick!: p2.Y = 10: p2.z = -Thick!
    p3.X = Thick!: p3.Y = 0: p3.z = -Thick!
    p4.X = -Thick!: p4.Y = 0: p4.z = -Thick!
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = Thick!: p1.z = 10
    p2.X = Thick!: p2.Y = Thick!: p2.z = 10
    p3.X = Thick!: p3.Y = Thick!: p3.z = 0
    p4.X = -Thick!: p4.Y = Thick!: p4.z = 0
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = Thick!: p1.z = 0
    p2.X = Thick!: p2.Y = Thick!: p2.z = 0
    p3.X = Thick!: p3.Y = -Thick!: p3.z = 0
    p4.X = -Thick!: p4.Y = -Thick!: p4.z = 0
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = Thick!: p1.z = 10
    p2.X = -Thick!: p2.Y = Thick!: p2.z = 0
    p3.X = -Thick!: p3.Y = -Thick!: p3.z = 0
    p4.X = -Thick!: p4.Y = -Thick!: p4.z = 10
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = Thick!: p1.Y = Thick!: p1.z = 10
    p2.X = -Thick!: p2.Y = Thick!: p2.z = 10
    p3.X = -Thick!: p3.Y = -Thick!: p3.z = 10
    p4.X = Thick!: p4.Y = -Thick!: p4.z = 10
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = Thick!: p1.Y = Thick!: p1.z = 0
    p2.X = Thick!: p2.Y = Thick!: p2.z = 10
    p3.X = Thick!: p3.Y = -Thick!: p3.z = 10
    p4.X = Thick!: p4.Y = -Thick!: p4.z = 0
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    p1.X = -Thick!: p1.Y = -Thick!: p1.z = 0
    p2.X = Thick!: p2.Y = -Thick!: p2.z = 0
    p3.X = Thick!: p3.Y = -Thick!: p3.z = 10
    p4.X = -Thick!: p4.Y = -Thick!: p4.z = 10
    Call dxConstruitMeshBuilder.Add_Rectangle(p1, p2, p3, p4, -1, Color&)
    '
    Set DxWall(0).WallMeshBuilder = dxEngine.dxD3Drm.CreateMeshBuilder
    Call dxConstruitMeshBuilder.Build(DxWall(0).WallMeshBuilder, dxTexture())
End Sub

'
' **********************************
' Search for the first an last entry
' in array who mach the number
' adr% = Array adress
' z%   = Number to search
' MIN% = First entry of the array
' MAX% = Last entry of the array
' r1%  = return 1st value
' r2%  = return last value
' **********************************
'
Public Sub Find_List(adr%(), z%, MIN%, MAX%, r1%, r2%)
    Dim p1%, p2%, t%
    If MAX% = 0 Then
        r1% = 0
        r2% = 0
    Else
        '
        ' ***** Search one z% in adr%()
        '
        p1% = MIN%
        p2% = MAX%
        Do
            If adr%(p1%) = z% Then
                r1% = p1%
                Exit Do
            End If
            If adr%(p2%) = z% Then
                r1% = p2%
                Exit Do
            End If
            If p1% = p2% Or p1% + 1 = p2% Then
                r1% = 0
                r2% = 0
                Exit Do
            End If
            t% = (p1% + p2%) / 2
            If adr%(t%) < z% Then p1% = t% Else p2% = t%
        Loop
        '
        ' ***** Enlarge the search
        '
        If r1% = 0 Then
            If adr%(p1%) > z% Then
                r2% = p1%
            Else
                If adr%(p2%) < z% Then r2% = p2% + 1 Else r2% = p1% + 1
            End If
        Else
            r2% = r1%
            t% = -1
            Do
                If r1% = MIN% Then
                    t% = 0
                Else
                    If adr%(r1% - 1) = z% Then r1% = r1% - 1 Else t% = 0
                End If
            Loop Until t% = 0
            '
            t% = -1
            Do
                If r2% = MAX% Then
                    t% = 0
                Else
                    If adr%(r2% + 1) = z% Then r2% = r2% + 1 Else t% = 0
                End If
            Loop Until t% = 0
        End If
    End If
End Sub

'
' *******************************
' Search a place for z%
' adr% = array to use
' z%   = number to place
' min% = min array
' max% = max array
' *******************************
'
Function Find_Place%(adr%(), z%, MIN%, MAX%)
    Dim r1%, r2%
    Call Find_List(adr%(), z%, MIN%, MAX%, r1%, r2%)
    If r1% = 0 Then Find_Place% = 1 Else Find_Place% = r1%
End Function

