VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDirectX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ******************************
' D3D Tools for easyer use
' © Kerdal Jean-Michel 1998-2000
' ******************************
'
Option Explicit
Public dX7 As New DirectX7
' ***** 2D Stuff
Public dxFront As DirectDrawSurface4
Public dxBack As DirectDrawSurface4
Dim SrcRECT As DxVBLib.RECT
Dim DesRECT As DxVBLib.RECT
Dim PIXEL As DDPIXELFORMAT
'
Enum EnumMode3D
    Mode3DClipper
    Mode3DSurface
End Enum
' ***** 3D Stuff
Public dxDirectDraw As DirectDraw4
Public dxD3Drm As Direct3DRM3
Public dxScene As Direct3DRMFrame3 ' Main Scene
Public dxCamera As Direct3DRMFrame3 ' The camera frame
Public dxLight As Direct3DRMLight ' Light attach to camera
Public dxAmbientLight As Direct3DRMLight ' Ambient light
Public dxAnimationSet As Direct3DRMAnimationSet2
Public dxClipper As DirectDrawClipper
Public dxDevice As Direct3DRMDevice3
Public dxViewport As Direct3DRMViewport2
'
Public ModeUsed As EnumMode3D
Public VideoDriver% ' Driver actually in use
Public dxEnum As Direct3DEnumDevices
Public TextureMaxWidth%
Public TextureMaxHeight%
'
' ***** Device quality
'
Enum EnumQuality
    QUALITE_NORMAL = D3DRMRENDER_GOURAUD
    QUALITE_VISION = D3DRMLIGHT_ON + D3DRMFILL_WIREFRAME + D3DSHADE_FLAT
End Enum
'
Const LR_LOADFROMFILE = &H10
Const LR_CREATEDIBSECTION = &H2000
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'
' *****************************
' Wrap a texture around a mesh
' Mode 4 and 5 are hiden method
' *****************************
'
Public Sub Wrap_Texture(MeshBuilder As Direct3DRMMeshBuilder3, dxTexture As Direct3DRMTexture3, WrapType As CONST_D3DRMWRAPTYPE, Optional ByVal nx%, Optional ByVal ny%)
    Dim dxBox As D3DRMBOX
    Dim Width!
    Dim Height!
    Dim dxWrap As Direct3DRMWrap
    '
    Call MeshBuilder.GetBox(dxBox)
    Height! = dxBox.Max.Y - dxBox.Min.Y
    Width! = dxBox.Max.X - dxBox.Min.X
    If nx% <= 0 Then nx% = 1
    If ny% <= 0 Then ny% = 1
    Select Case WrapType
    Case D3DRMWRAP_FLAT
        If Height! = 0 Then Exit Sub
        If Width! = 0 Then Exit Sub
        Set dxWrap = dxD3Drm.CreateWrap(D3DRMWRAP_FLAT, Nothing, _
        dxBox.Min.X, dxBox.Min.Y, dxBox.Min.z, _
        0, 0, 1, _
        0, 1, 0, _
        0, 0, _
        nx% / Width!, -ny% / Height!)
    Case D3DRMWRAP_CYLINDER
        If Height! = 0 Then Exit Sub
        Set dxWrap = dxD3Drm.CreateWrap(D3DRMWRAP_CYLINDER, Nothing, _
        0, 0, 0, _
        0, 1, 0, _
        0, 0, 1, _
        0, dxBox.Min.Y / Height!, _
        1, -1 / Height!)
    Case D3DRMWRAP_SPHERE
        Set dxWrap = dxD3Drm.CreateWrap(D3DRMWRAP_SPHERE, Nothing, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1)
    Case D3DRMWRAP_CHROME
        If Height! = 0 Then Exit Sub
        Set dxWrap = dxD3Drm.CreateWrap(D3DRMWRAP_CHROME, Nothing, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, dxBox.Min.Y / Height!, 1, -1 / Height)
    Case 4 ' Double face texturing (like a sheet of paper)
        If Height! = 0 Then Exit Sub
        If Width! = 0 Then Exit Sub
        Set dxWrap = dxD3Drm.CreateWrap(4, Nothing, _
        dxBox.Min.X, dxBox.Min.Y, dxBox.Min.z, _
        0, 0, 1, _
        0, 1, 0, _
        0, 0, _
        nx% / Width!, -ny% / Height!)
    Case 5 ' Cuboïde texture (like a dice)
        If Height! = 0 Then Exit Sub
        If Width! = 0 Then Exit Sub
        Set dxWrap = dxD3Drm.CreateWrap(5, Nothing, _
        dxBox.Min.X, dxBox.Min.Y, dxBox.Min.z, _
        0, 0, 1, _
        0, 1, 0, _
        0, 0, _
        nx% / Width!, ny% / Height!)
    End Select
    Call dxWrap.Apply(MeshBuilder)
    Call MeshBuilder.SetTexture(dxTexture)
    Set dxWrap = Nothing
End Sub

'
' **************************
' Load the X file and change
' the working directory
' **************************
'
Public Function Load_MeshBuilder(Name$) As Direct3DRMMeshBuilder3
    Set Load_MeshBuilder = dxD3Drm.CreateMeshBuilder
    Call Load_MeshBuilder.LoadFromFile(Find_Search_Path(Name$), 0, D3DRMLOAD_FROMFILE, Nothing, Nothing)
    Call Load_MeshBuilder.SetPerspective(D_TRUE)
End Function

'
' ************
' Load texture
' ************
'
Public Function Load_Texture(File$) As Direct3DRMTexture3
    Set Load_Texture = dxD3Drm.LoadTexture(File$)
End Function

'
' *****************
' Create interfaces
' and base elements
' *****************
'
Private Sub Class_Initialize()
    Set dxDirectDraw = dX7.DirectDraw4Create("")
    Set dxD3Drm = dX7.Direct3DRMCreate
    Call dxDirectDraw.SetCooperativeLevel(ByVal 0&, DDSCL_NORMAL Or DDSCL_NOWINDOWCHANGES)
    Set dxClipper = dxDirectDraw.CreateClipper(0)
    ' Main scene
    Set dxScene = dxD3Drm.CreateFrame(Nothing)
    ' Camera
    Set dxCamera = dxD3Drm.CreateFrame(dxScene)
    ' Spot attach to the camera
    Set dxLight = dxD3Drm.CreateLightRGB(D3DRMLIGHT_POINT, 0, 0, 0)
    Call dxCamera.AddLight(dxLight)
    ' Ambient light
    Set dxAmbientLight = dxD3Drm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 1, 1, 1)
    Call dxScene.AddLight(dxAmbientLight)
    ' Animation
    Set dxAnimationSet = dxD3Drm.CreateAnimationSet
    '
    ' **************************************************
    ' EnumDevices and chose the last one, the better one
    ' **************************************************
    '
    Dim dxDirect3D As Direct3D7
    Dim dxDirectDraw7 As DirectDraw7
    Dim dxDesc As D3DDEVICEDESC7
    '
    ' ***** Drivers
    '
    Set dxDirectDraw7 = dX7.DirectDrawCreate("")
    Set dxDirect3D = dxDirectDraw7.GetDirect3D()
    Set dxEnum = dxDirect3D.GetDevicesEnum()
    VideoDriver% = dxEnum.GetCount() ' The last by default
    Call dxEnum.GetDesc(VideoDriver%, dxDesc)
    TextureMaxHeight% = dxDesc.lMaxTextureHeight
    TextureMaxWidth% = dxDesc.lMaxTextureWidth
    Set dxDirect3D = Nothing
    Set dxDirectDraw7 = Nothing
End Sub

'
' *******************
' Objects destruction
' *******************
'
Private Sub Class_Terminate()
    If Not (dxEnum Is Nothing) Then Set dxEnum = Nothing
    If Not (dxBack Is Nothing) Then Set dxBack = Nothing
    If Not (dxFront Is Nothing) Then Set dxFront = Nothing
    If Not (dxViewport Is Nothing) Then Set dxViewport = Nothing
    If Not (dxDevice Is Nothing) Then Set dxDevice = Nothing
    If Not (dxAnimationSet Is Nothing) Then Set dxAnimationSet = Nothing
    If Not (dxAmbientLight Is Nothing) Then Set dxAmbientLight = Nothing
    If Not (dxLight Is Nothing) Then Set dxLight = Nothing
    If Not (dxCamera Is Nothing) Then Set dxCamera = Nothing
    If Not (dxScene Is Nothing) Then Set dxScene = Nothing
    If Not (dxClipper Is Nothing) Then Set dxClipper = Nothing
    If Not (dxD3Drm Is Nothing) Then Set dxD3Drm = Nothing
    If Not (dxDirectDraw Is Nothing) Then Set dxDirectDraw = Nothing
End Sub

'
' ********************************
' Create a new 3DRM with this zone
' ********************************
'
Public Sub Create_3DRM(Zone As Object, TX%, TY%, Mode As EnumMode3D)
    Const GDC_BitPerPixel As Long = &HC&
    Dim DDSD As DDSURFACEDESC2
    Dim DDSC As DDSCAPS2
    Dim Driver$
    ' ***** Set to nothing if recall this Sub
    If Not (dxViewport Is Nothing) Then Set dxViewport = Nothing
    ' ***** DirectDraw2 Buffers based
    If Not (dxBack Is Nothing) Then Set dxBack = Nothing
    If Not (dxFront Is Nothing) Then Set dxFront = Nothing
    ' ***** Create the 3DRM object
    Call dxClipper.SetHWnd(Zone.hWnd)
    ModeUsed = Mode ' Remember how we call this sub
    If VideoDriver% = 0 Then
        Driver$ = ""
    Else
        Driver$ = dxEnum.GetGuid(VideoDriver%)
        If InStr(UCase(Driver$), "IID_IDIRECT3D") = 0 Then
            Driver$ = dxEnum.GetGuid(VideoDriver% - 1)
            '"IID_IDirect3DTnLHalDevice" normally? Bug or not?
            ' I take the forward one for the moment
        End If
    End If
    '
    Select Case ModeUsed
    Case EnumMode3D.Mode3DClipper
        Set dxDevice = dxD3Drm.CreateDeviceFromClipper(dxClipper, Driver$, TX%, TY%)
    Case EnumMode3D.Mode3DSurface
        ' ***** Main buffer
        With DDSD
            .lFlags = DDSD_CAPS
            If InStr(UCase(Driver$), "HAL") = 0 Then
                .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
            Else
                .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_VIDEOMEMORY
            End If
        End With
        Set dxFront = dxDirectDraw.CreateSurface(DDSD)
        Call dxFront.SetClipper(dxClipper)
        '***** Back buffer
        With DDSD
            .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
            .lWidth = TX%
            .lHeight = TY%
            If InStr(UCase(Driver$), "HAL") = 0 Then
                .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_SYSTEMMEMORY ' Or DDSCAPS_OWNDC
            Else
                .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_VIDEOMEMORY ' Or DDSCAPS_OWNDC
            End If
        End With
        Set dxBack = dxDirectDraw.CreateSurface(DDSD)
        '***** Interface Device
        Set dxDevice = dxD3Drm.CreateDeviceFromSurface(Driver$, dxDirectDraw, dxBack, D3DRMDEVICE_DEFAULT)
    End Select
    Set dxViewport = dxD3Drm.CreateViewport(dxDevice, dxCamera, 0, 0, TX%, TY%)
    ' ***** Set the image quality
    With dxDevice
        .SetTextureQuality D3DRMTEXTURE_NEAREST ' D3DRMTEXTURE_NEAREST or D3DRMTEXTURE_LINEAR
        .SetDither 1 ' 0 or 1
        .SetQuality EnumQuality.QUALITE_NORMAL
        .SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY + _
                       D3DRMRENDERMODE_SORTEDTRANSPARENCY + _
                       D3DRMRENDERMODE_VIEWDEPENDENTSPECULAR
        Select Case GetDeviceCaps(Zone.hdc, GDC_BitPerPixel)
        Case 8
            .SetShades 4
            Call dxD3Drm.SetDefaultTextureShades(4)
        Case 16
            .SetShades 32
            Call dxD3Drm.SetDefaultTextureColors(64)
            Call dxD3Drm.SetDefaultTextureShades(32)
        Case 24, 32
            .SetShades 256
            Call dxD3Drm.SetDefaultTextureColors(64)
            Call dxD3Drm.SetDefaultTextureShades(256)
        End Select
    End With
    With dxViewport
        .SetBack 10000! ' Default 100
        .SetProjection D3DRMPROJECT_PERSPECTIVE 'D3DRMPROJECT_ORTHOGRAPHIC
        '.SetField (1)
        .SetUniformScaling True
    End With
    '
    Call dxScene.SetSceneFogColor(&HFFFFFF)
    Call dxScene.SetSceneFogMethod(D3DRMFOGMETHOD_ANY)
    Call dxScene.SetSceneFogMode(D3DRMFOG_EXPONENTIAL)
    Call dxScene.SetSceneFogParams(0, 70, 1)
    '
    Call dxScene.SetSceneFogEnable(D_FALSE)
End Sub

'
' ******************************************
' Calculate and show scene
' Copy: used to let you access to BackBuffer
' before copy to FrontBuffer (if exist)
' ******************************************
'
Public Sub Render(Copy As Boolean)
    On Error GoTo Abort
    If Copy = False Then
        With dxViewport
            Call .Clear(D3DRMCLEAR_ALL)
            Call .Render(dxScene)
        End With
        Call dxDevice.Update
    ElseIf ModeUsed = Mode3DSurface Then
        Call dX7.GetWindowRect(dxClipper.GetHWnd, DesRECT)
        Call dxFront.Blt(DesRECT, dxBack, SrcRECT, DDBLT_WAIT)
    End If
    Exit Sub
Abort:
    Call MsgBox("Error" + Str$(Err.Number) + vbCr + Err.Description, vbCritical + vbOKOnly)
End Sub

'
' *************************
' Load a .x file in a frame
' *************************
'
Public Sub Load_Frame(Frame As Direct3DRMFrame3, Name$)
    Call Frame.LoadFromFile(Find_Search_Path$(Name$), 0, D3DRMLOAD_FROMFILE, Nothing, Nothing)
End Sub

'
' ***************************
' Load a animation in a frame
' ***************************
'
Public Sub Load_Animation(Name$, dxFrame As Direct3DRMFrame3)
    Call dxAnimationSet.LoadFromFile(Find_Search_Path(Name$), 0, 0, Nothing, Nothing, dxFrame)
End Sub

'
' ********************
' Find the search path
' for .x file to load
' ********************
'
Private Function Find_Search_Path$(Name$)
    Dim i%, n%
    n% = 0
    For i% = 1 To Len(Name$)
        If Mid$(Name$, i%, 1) = "\" Then n% = i%
    Next i%
    Call dxD3Drm.SetSearchPath(Mid$(Name$, 1, n% - 1))
    Find_Search_Path$ = "." & Mid$(Name$, n%, Len(Name$) - n% + 1)
End Function

