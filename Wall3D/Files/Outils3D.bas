Attribute VB_Name = "Tools3D"
Option Explicit
Global Const PI! = 3.14159265358979
'
' ***** Wall definition
'
Type TypeWall
    TheFile As String ' Extern X file description
    AnimationFile As String ' Extern X file animation
    Faces As Boolean ' Indicate if theirs some faces to show
    WrapType As Integer 'CONST_D3DRMWRAPTYPE
    WrapTexture As Integer
    WrapNx As Integer
    WrapNy As Integer
    ParentNo As Integer ' Parent frame number (default=0:scene)
    FrameName As String
End Type
Type TypeDxWall
    WallFrame As Direct3DRMFrame3
    WallMeshBuilder As Direct3DRMMeshBuilder3
    WallShadow As Direct3DRMShadow2
End Type
'
Global Const NbWall% = 256 ' Number of meshbuilder
Global DefWall(NbWall%) As TypeWall ' Meshbuilder definition
Global DxWall(NbWall%) As TypeDxWall ' Frames & MeshBuilders
Global MeshCopyDPosition&(11)
'
' ***** Textures
'
Type TypeTexture
    Name As String
    Transparency As Integer
    Color As Long
End Type
Global Const NbTexture% = 256 ' Max textures
Global Texture(NbTexture%) As TypeTexture
Global dxTexture(NbTexture%) As Direct3DRMTexture3 ' Texture array
'
' ***** Definition of faces
'
Type TypeMur3D
    NbPoint As Integer
    Point(3) As D3DVECTOR
    Texture As Integer
    Color As PALETTEENTRY
End Type
Global FaceIndex%()
Global DefFace() As TypeMur3D ' Faces array

'
' ************************************
' Retourne l'angle fait par le segment
' ************************************
'
Public Function Calcul_Angle%(x1!, z1!, x2!, z2!)
    Dim Tangente#
    If x1! = x2! Then
        If z2! > z1! Then
            Calcul_Angle% = 90
        Else
            Calcul_Angle% = 270
        End If
    Else
        Tangente# = (z2! - z1!) / (x2! - x1!)
        Calcul_Angle% = Atn(Tangente#) * 180 / PI!
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
' *********************
' Build the meshbuilder
' *********************
'
Public Sub Build_Wall(dX As ClassDirectX, n%, TmpFaceIndex%(), _
TmpDefFace() As TypeMur3D, TmpDefWall() As TypeWall, _
TmpDxWall() As TypeDxWall, TmpDxTexture() As Direct3DRMTexture3)
    '
    If TmpDefWall(n%).TheFile$ <> "" Then Exit Sub ' A load mesh exist already
    Dim i%, c&
    If n% = 0 Then Exit Sub
    Call dxConstruitMeshBuilder.Init
    TmpDefWall(n%).Faces = False
    For i% = 0 To UBound(TmpDefFace)
        If TmpFaceIndex%(i%) = n% Then
            TmpDefWall(n%).Faces = True
            c& = dX.dX7.CreateColorRGBA(TmpDefFace(i%).Color.red / 255, _
                                      TmpDefFace(i%).Color.green / 255, _
                                      TmpDefFace(i%).Color.blue / 255, _
                                      TmpDefFace(i%).Color.flags / 255)
            If TmpDefFace(i%).NbPoint% = 3 Then
                ' ***** Triangle
                Call dxConstruitMeshBuilder.Add_Triangle(TmpDefFace(i%).Point(0) _
                                   , TmpDefFace(i%).Point(1) _
                                   , TmpDefFace(i%).Point(2) _
                                   , TmpDefFace(i%).Texture% _
                                   , c&)
            Else
                ' ***** Rectangle
                Call dxConstruitMeshBuilder.Add_Rectangle(TmpDefFace(i%).Point(0) _
                                    , TmpDefFace(i%).Point(1) _
                                    , TmpDefFace(i%).Point(2) _
                                    , TmpDefFace(i%).Point(3) _
                                    , TmpDefFace(i%).Texture% _
                                    , c&)
            End If
        End If
    Next i%
    Set TmpDxWall(n%).WallMeshBuilder = dX.dxD3Drm.CreateMeshBuilder
    If TmpDefWall(n%).Faces = True Then
        Call dxConstruitMeshBuilder.Build(TmpDxWall(n%).WallMeshBuilder, TmpDxTexture())
        If TmpDefWall(n%).WrapTexture% <> 0 Then
            Dim TmpWrapType As CONST_D3DRMWRAPTYPE
            TmpWrapType = TmpDefWall(n%).WrapType
            Call dX.Wrap_Texture(TmpDxWall(n%).WallMeshBuilder, _
            TmpDxTexture(TmpDefWall(n%).WrapTexture%), TmpWrapType, _
            TmpDefWall(n%).WrapNx%, TmpDefWall(n%).WrapNy%)
        End If
    End If
End Sub

'
' *******************************************
' Build a MeshBuilder with a .wall definition
' *******************************************
'
Public Sub Build_MeshBuilder(dX As ClassDirectX, WallFile$, AnimFile$, _
MeshBuilders() As Direct3DRMMeshBuilder3, _
TdxTexture() As Direct3DRMTexture3, TTexture() As TypeTexture)
    Dim TFaceIndex%(), n%, i%, NbAnimation%
    Dim TDefFace() As TypeMur3D
    Dim TDefWall(NbWall%) As TypeWall
    Dim TDxWall(NbWall%) As TypeDxWall
    Dim TAnimation&()
    Dim TNbTexture%
    Dim TFrame As Direct3DRMFrame3
    '
    ' ***** Create a tempory master scene
    '
    Set TFrame = dX.dxD3Drm.CreateFrame(Nothing)
    For i% = 1 To NbWall%
        Set TDxWall(i%).WallFrame = dX.dxD3Drm.CreateFrame(TFrame)
    Next i%
    '
    Call Load_Wall_Definition(WallFile$, TFaceIndex%(), TDefFace(), TDefWall(), TAnimation&(), NbAnimation%, TNbTexture%, TTexture())
    '
    If TNbTexture% <> 0 Then
        For i% = 1 To TNbTexture%
            Call Tools3D.Textures_Update(dX, TdxTexture(i%), TTexture(i%))
        Next i%
    End If
    '
    ' ***** Build all MeshBuilder
    '
    ReDim MeshBuilders(NbAnimation%)
    For i% = 1 To NbWall%
        If TDefWall(i%).TheFile$ <> "" Then
            Set TDxWall(i%).WallMeshBuilder = dX.Load_MeshBuilder(TDefWall(i%).TheFile$)
            TDefWall(i%).Faces = True
        Else
            If TDefWall(i%).Faces = True Then _
            Call Build_Wall(dX, i%, TFaceIndex%(), TDefFace(), TDefWall(), TDxWall(), TdxTexture())
        End If
    Next i%
    '
    ' ***** Assembly Frames
    '
    For n% = 0 To NbAnimation%
        For i% = 1 To NbWall%
            If TDefWall(i%).Faces = True Then
                Call TDxWall(i%).WallFrame.AddVisual(TDxWall(i%).WallMeshBuilder)
                ' ***** Replace mesh
                If TDefWall(i%).ParentNo% = 0 Then
                    Call TFrame.AddChild(TDxWall(i%).WallFrame)
                    Call TDxWall(i%).WallFrame.SetPosition(TFrame, 0, 0, 0)
                    Call TDxWall(i%).WallFrame.SetOrientation(TFrame, 0, 0, 1, 0, 1, 0)
                Else
                    Call TDxWall(TDefWall(i%).ParentNo%).WallFrame.AddChild(TDxWall(i%).WallFrame)
                    Call TDxWall(i%).WallFrame.SetPosition(TDxWall(TDefWall(i%).ParentNo%).WallFrame, 0, 0, 0)
                    Call TDxWall(i%).WallFrame.SetOrientation(TDxWall(TDefWall(i%).ParentNo%).WallFrame, 0, 0, 1, 0, 1, 0)
                End If
                ' ***** Transform mesh
                Call TDxWall(i%).WallFrame.AddScale(D3DRMCOMBINE_AFTER, TAnimation&(i%, 6, n%) / 10 ^ 5, TAnimation&(i%, 7, n%) / 10 ^ 5, TAnimation&(i%, 8, n%) / 10 ^ 5)
                Call TDxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 1, 0, TAnimation&(i%, 3, n%) / 10 ^ 5 / 180 * PI!)
                Call TDxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 1, 0, 0, TAnimation&(i%, 4, n%) / 10 ^ 5 / 180 * PI!)
                Call TDxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 0, 1, TAnimation&(i%, 5, n%) / 10 ^ 5 / 180 * PI!)
                Call TDxWall(i%).WallFrame.AddTranslation(D3DRMCOMBINE_AFTER, TAnimation&(i%, 0, n%) / 10 ^ 5, TAnimation&(i%, 1, n%) / 10 ^ 5, TAnimation&(i%, 2, n%) / 10 ^ 5)
                Call TDxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 1, 0, TAnimation&(i%, 9, n%) / 10 ^ 5 / 180 * PI!)
                Call TDxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 1, 0, 0, TAnimation&(i%, 10, n%) / 10 ^ 5 / 180 * PI!)
                Call TDxWall(i%).WallFrame.AddRotation(D3DRMCOMBINE_AFTER, 0, 0, 1, TAnimation&(i%, 11, n%) / 10 ^ 5 / 180 * PI!)
            End If
        Next i%
        '
        ' ***** Get the final MeshBuilder
        '
        Set MeshBuilders(n%) = dX.dxD3Drm.CreateMeshBuilder
        Call MeshBuilders(n%).AddFrame(TFrame)
    Next n%
    '
    ' ***** Release objects
    '
    Set TFrame = Nothing
    For i% = 1 To NbWall%
        Set TDxWall(i%).WallMeshBuilder = Nothing
        Set TDxWall(i%).WallFrame = Nothing
    Next i%
End Sub

'
' **************************
' Erase 3D object at the end
' **************************
'
Public Sub Erase_3DObject()
    Dim i%
    For i% = 1 To NbTexture%
        Set dxTexture(i%) = Nothing
    Next i%
    For i% = 0 To NbWall%
        Set DxWall(i%).WallMeshBuilder = Nothing
    Next i%
End Sub

'
' *********************
' Load the texture list
' *********************
'
Public Sub Textures_Load(dX As ClassDirectX, FileTexture$)
    Dim n%, i%
    Dim canal%
    canal% = FreeFile()
    Call Key_Init(15)
    Call Open_File(FileTexture$, canal%, OPEN_BINARY)
    n% = Load_Integer%(canal%)
    For i% = 1 To NbTexture%
        Texture(i%).Name$ = Load_Chain$(canal%)
        Texture(i%).Transparency% = Load_Integer%(canal%)
        Texture(i%).Color& = Load_Long&(canal%)
        Call Textures_Update(dX, dxTexture(i%), Texture(i%))
    Next i%
    Close #canal%
End Sub

'
' ******************
' Load the wall list
' ******************
'
Public Sub Load_Wall_Definition(File$, TmpFaceIndex%(), _
TmpDefFace() As TypeMur3D, TmpDefWall() As TypeWall, _
TmpAnimation&(), TmpNbAnimation%, _
TmpNbTexture%, TmpTexture() As TypeTexture)
    Dim i%, j%, n%, f%
    f% = FreeFile()
    Open File$ For Binary Access Read As #f%
    Get #f%, , n%
    ReDim TmpFaceIndex%(n%)
    ReDim TmpDefFace(n%) As TypeMur3D
    Get #f%, , TmpFaceIndex%()
    Get #f%, , TmpDefFace()
    Get #f%, , TmpDefWall()
    Get #f%, , TmpNbAnimation%
    ReDim TmpAnimation&(NbWall%, 11, TmpNbAnimation%)
    Get #f%, , TmpAnimation&()
    Get #f%, , TmpNbTexture%
    If TmpNbTexture% <> 0 Then
        Get #f%, , TmpTexture()
    End If
    Close #f%
End Sub

'
' ************************
' Add a face to the object
' ************************
'
Public Sub Face_Add(NoMur%, Color As PALETTEENTRY, Texture%, nb%, P() As D3DVECTOR, Inversion As Boolean)
    Dim i%, t%
    t% = 0
    For i% = 1 To UBound(DefFace)
        If NoMur% >= FaceIndex%(i%) Then t% = i%
    Next i%
    t% = t% + 1
    '
    ReDim Preserve FaceIndex%(UBound(FaceIndex) + 1)
    ReDim Preserve DefFace(UBound(DefFace) + 1) As TypeMur3D
    For i% = UBound(DefFace) To t% + 1 Step -1
        DefFace(i%) = DefFace(i% - 1)
        FaceIndex%(i%) = FaceIndex%(i% - 1)
    Next i%
    '
    DefFace(t%).Color = Color
    DefFace(t%).Texture = Texture%
    FaceIndex%(t%) = NoMur%
    DefFace(t%).NbPoint = nb%
    DefFace(t%).Point(0) = P(0)
    If Inversion = True Then
        If nb% = 3 Then
            DefFace(t%).Point(1) = P(2)
            DefFace(t%).Point(2) = P(1)
        Else
            DefFace(t%).Point(1) = P(3)
            DefFace(t%).Point(2) = P(2)
            DefFace(t%).Point(3) = P(1)
        End If
    Else
        For i% = 1 To nb% - 1
            DefFace(t%).Point(i%) = P(i%)
        Next i%
    End If
End Sub

'
' ***********************
' Delete the current face
' ***********************
'
Public Sub Face_Delete(n%)
    Dim i%
    For i% = n% To UBound(DefFace) - 1
        DefFace(i%) = DefFace(i% + 1)
        FaceIndex%(i%) = FaceIndex%(i% + 1)
    Next i%
    ReDim Preserve FaceIndex%(UBound(FaceIndex) - 1)
    ReDim Preserve DefFace(UBound(DefFace) - 1) As TypeMur3D
End Sub

'
' *********************
' Save textures to file
' *********************
'
Public Sub Textures_Save(dxTexture$)
    Dim i%, n%, f%
    If dxTexture$ = "" Then Exit Sub
    n% = 0
    For i% = 1 To NbTexture%
        If Texture(i%).Name$ <> "" Then n% = i%
    Next i%
    Call Key_Init(15)
    f% = FreeFile()
    Open dxTexture$ For Output As #f%
    Call Save_Integer(f%, n%)
    For i% = 1 To NbTexture%
        Call Save_Chain(f%, Texture(i%).Name$)
        Call Save_Integer(f%, Texture(i%).Transparency)
        Call Save_Long(f%, Texture(i%).Color)
    Next i%
    Close #f%
End Sub

'
' ****************************
' Load a new texture from file
' and set the transparency
' ****************************
'
Public Sub Textures_Update(dX As ClassDirectX, dxTexture As Direct3DRMTexture3, DefTexture As TypeTexture)
    If DefTexture.Name$ <> "" Then
        Set dxTexture = dX.Load_Texture(DefTexture.Name$)
        If Not dxTexture Is Nothing Then
            If DefTexture.Transparency <> 0 Then
                Call dxTexture.SetDecalTransparency(True)
            Else
                Call dxTexture.SetDecalTransparency(False)
            End If
            Call dxTexture.SetDecalTransparentColor(DefTexture.Color)
        End If
    Else
        Set dxTexture = Nothing
    End If
End Sub

