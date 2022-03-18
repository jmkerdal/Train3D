VERSION 5.00
Begin VB.MDIForm MainWindow 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Wall3D"
   ClientHeight    =   3480
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5460
   Icon            =   "Principale.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MENU_File 
      Caption         =   "File"
      Begin VB.Menu MENU_Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MENU_Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MENU_New 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu MENU_Export_X 
         Caption         =   "Export .x meshbuilder"
      End
      Begin VB.Menu MENU_Minus0 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Load_Frame 
         Caption         =   "Load .x frame"
      End
      Begin VB.Menu MENU_Save_Frame 
         Caption         =   "Save .x frame"
      End
      Begin VB.Menu MENU_Minus1 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Screenshot 
         Caption         =   "Screenshot"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu MENU_Minus2 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu MENU_Textures 
      Caption         =   "Textures"
      Begin VB.Menu MENU_Load_Texture 
         Caption         =   "Open"
      End
      Begin VB.Menu MENU_Save_Texture 
         Caption         =   "Save"
      End
      Begin VB.Menu MENU_Minus3 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Show_Textures 
         Caption         =   "Show"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu MENU_Animation 
      Caption         =   "Animation"
      Begin VB.Menu MENU_Animation_Open 
         Caption         =   "Import"
      End
      Begin VB.Menu MENU_Animation_Save 
         Caption         =   "Export"
      End
      Begin VB.Menu MENU_Minus4 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Animation_Show 
         Caption         =   "Show"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MENU_Geometry 
      Caption         =   "Geometry"
      Begin VB.Menu MENU_Standard 
         Caption         =   "Standard"
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_Shape 
         Caption         =   "Shape"
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_Faces 
         Caption         =   "Faces"
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_Resize 
         Caption         =   "Resize all"
      End
      Begin VB.Menu MENU_Minus5 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Copy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu MENU_Paste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu MENU_Cut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MENU_Render 
      Caption         =   "Render"
      Begin VB.Menu MENU_Solid 
         Caption         =   "Solid"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu MENU_Orthographic 
         Caption         =   "Orthographic"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MENU_Back_Color 
         Caption         =   "Back color"
      End
      Begin VB.Menu MENU_Ambient_Light 
         Caption         =   "Ambient light"
      End
      Begin VB.Menu MENU_Spot_Light 
         Caption         =   "Spot light"
      End
      Begin VB.Menu MENU_Minus6 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_Driver 
         Caption         =   "No acceleration"
         Index           =   0
      End
      Begin VB.Menu MENU_Driver 
         Caption         =   "Driver1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Driver 
         Caption         =   "Driver2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_Driver 
         Caption         =   "Driver3"
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MENU_Import 
      Caption         =   "Import"
      Begin VB.Menu MENU_Import_MeshBuilder 
         Caption         =   "Import .x meshbuilder"
      End
      Begin VB.Menu MENU_Import_Wall 
         Caption         =   "Import .wall in meshbuilder"
      End
   End
   Begin VB.Menu MENU_Interrogation 
      Caption         =   "?"
      Begin VB.Menu MENU_About 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
'
Dim FaceCopy() As TypeMur3D
Dim MeshCopy As TypeWall

'
' ***********************
' Show the main vue first
' ***********************
'
Private Sub MDIForm_Load()
    Show
    '
    ' Organize form position at startup
    '
    Dim XOffset%
    Dim YOffset%
    XOffset% = Screen.TwipsPerPixelX * 16
    YOffset% = Screen.TwipsPerPixelY * 55
    '
    Call Textures.Move((MainWindow.Width - MainWindow.Left) - Textures.Width - XOffset%, (MainWindow.Height - MainWindow.Top) - Textures.Height - YOffset%)
    Call Axe_Camera.Move(0, (MainWindow.Height - MainWindow.Top) - Axe_Camera.Height - YOffset%)
    Call Transform.Move(Axe_Camera.Width, Axe_Camera.Top)
    Call Wall3D.Move((MainWindow.Width - MainWindow.Left) - Wall3D.Width - XOffset%, 0)
    Call View.Move(0, 0, Screen.Width / 2, Screen.Height / 2)
    '
    Dim i%, n%
    n% = dxEngine.dxEnum.GetCount()
    If n% > 3 Then n% = 3
    If n% > 0 Then
        For i% = 1 To n%
            MENU_Driver(i%).Visible = True
            MENU_Driver(i%).Caption = dxEngine.dxEnum.GetName(i%)
        Next i%
        MENU_Driver(n%).Checked = True
    End If
    Call Init.Shape_Init
    Call Textures.Show
    Call Wall3D.Show
    Call View.Show
End Sub

'
' ************************************
' Verify if you need to save something
' before exit
' ************************************
'
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i%
    For i% = 0 To NbWall%
        Set DxWall(NbWall%).WallFrame = Nothing
    Next i%
    Set XFileLoadFrame = Nothing
    Call Erase_3DObject
    Set dxEngine = Nothing
    End
End Sub

'
' *********************
' Show the "About form"
' *********************
'
Private Sub MENU_About_Click()
    Dim d$, r$
    d$ = "Create by: Kerdal Jean-Michel" + vbCr + _
        "2, place des Martyrs" + vbCr + _
        "92110 Clichy, France" + vbCr + _
        "Telephone: 01 49 68 83 67" + vbCr + _
        "mailto:jeanmichel.kerdal@free.fr"
    r$ = "© DemiCitron 1998-2000" + vbCr + "http://jeanmichel.kerdal.free.fr/"
    Call frmAbout.About(Me, "Direct3D wall editor", d$, "Version", r$, ".\Wall3D\Icones\Kerdal.gif")
End Sub

'
' ****************
' Resize all faces
' ****************
'
Private Sub MENU_Resize_Click()
    Dim i%, j%, n%
    Dim r$, Resize!
    '
    r$ = InputBox("Percent", "Resize all", "100")
    If r$ = "" Then Exit Sub
    Resize! = Val(r$)
    If Resize! = 0 Then Exit Sub
    '
    For n% = 0 To NbAnimation%
        For i% = 1 To NbWall%
            For j% = 0 To 2
                Animation&(i%, j%, n%) = Animation&(i%, j%, n%) * Resize! / 100
                Animation&(i%, j% + 6, n%) = Animation&(i%, j% + 6, n%) * Resize! / 100
            Next j%
        Next i%
    Next n%
    '
    DoEvents
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' *********************
' Show the texture form
' *********************
'
Private Sub MENU_Show_Textures_Click()
    Textures.Show
End Sub

'
' ********************
' Change ambient light
' ********************
'
Private Sub MENU_Ambient_Light_Click()
    Dim LightColor As Long
    LightColor = dxEngine.dxAmbientLight.GetColor
    Load ColorSelection
    ColorSelection.TheColor.Red = dxEngine.dX7.ColorGetRed(LightColor) * 255
    ColorSelection.TheColor.Green = dxEngine.dX7.ColorGetGreen(LightColor) * 255
    ColorSelection.TheColor.Blue = dxEngine.dX7.ColorGetBlue(LightColor) * 255
    ColorSelection.TheColor.ShowAlpha = False
    ColorSelection.Show vbModal
    Call dxEngine.dxAmbientLight.SetColorRGB(TheRed / 255, TheGreen / 255, TheBlue / 255)
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' **************
' Load animation
' **************
'
Private Sub MENU_Animation_Open_Click()
    Dim n%, i%, j%
    Dim File$, f%
    File$ = Open_Box$("Load project", AnimPath$, "All files *.anim|*.anim", BOX_LOAD, Wall3D.BoiteMur3D)
    If File$ = "" Then Exit Sub
    Unload FrmAnimation
    f% = FreeFile()
    Open File$ For Input As #f%
    '
    ' ****** Fill the animation array with default value
    '
    Input #f%, NbAnimation%
    ReDim Preserve Animation&(NbWall%, 11, NbAnimation%)
    For n% = 1 To NbAnimation%
        For i% = 1 To NbWall%
            For j% = 0 To 11
                Animation&(i%, j%, n%) = Animation&(i%, j%, 0)
            Next j%
        Next i%
    Next n%
    '
    ' ****** Load delta value
    '
    While Not (EOF(f%))
        Input #f%, n%
        Input #f%, i%
        For j% = 0 To 11
            Input #f%, Animation&(i%, j%, n%)
        Next j%
    Wend
    Close #f%
End Sub

'
' **************
' Save animation
' **************
'
Private Sub MENU_Animation_Save_Click()
    Dim n%, i%, j%
    Dim File$, f%
    Dim FileReturn As Boolean, ToSave As Boolean
    File$ = Open_Box$("Save project", AnimPath$, "All files *.anim|*.anim", BOX_SAVE, Wall3D.BoiteMur3D)
    If File$ = "" Then Exit Sub
    f% = FreeFile()
    Open File$ For Output As #f%
    Call Save_Number(f%, NbAnimation%, True)
    For n% = 1 To NbAnimation%
        For i% = 1 To NbWall%
            ToSave = False
            For j% = 0 To 11
                If Animation&(i%, j%, n%) <> Animation&(i%, j%, 0) Then ToSave = True
            Next j%
            If ToSave = True Then
                Call Save_Number(f%, n%, False)
                Call Save_Number(f%, i%, True)
                For j% = 0 To 11
                    If j% = 11 Then
                        FileReturn = True
                    Else
                        FileReturn = False
                    End If
                    Call Save_Number(f%, Animation&(i%, j%, n%), FileReturn)
                Next j%
            End If
        Next i%
    Next n%
    Close #f%
End Sub

'
' *******************
' Show animation form
' *******************
'
Private Sub MENU_Animation_Show_Click()
    FrmAnimation.Show
End Sub

'
' *****************
' Change back color
' *****************
'
Private Sub MENU_Back_Color_Click()
    Load ColorSelection
    ColorSelection.TheColor.Red = ColorBack.Red
    ColorSelection.TheColor.Green = ColorBack.Green
    ColorSelection.TheColor.Blue = ColorBack.Blue
    ColorSelection.TheColor.ShowAlpha = False
    ColorSelection.Show vbModal
    ColorBack.Red = TheRed
    ColorBack.Green = TheGreen
    ColorBack.Blue = TheBlue
    Call dxEngine.dxScene.SetSceneBackgroundRGB(ColorBack.Red / 255, ColorBack.Green / 255, ColorBack.Blue / 255)
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' *****************
' Load texture list
' *****************
'
Private Sub MENU_Load_Texture_Click()
    Dim File$, i%
    File$ = Open_Box$("Load textures", TexturePath$, "Texture list|*.tmur", BOX_LOAD, Wall3D.BoiteMur3D)
    If Exist(File$) = False Then Exit Sub
    Me.Enabled = False
    TexturePath$ = File$
    Call Tools3D.Textures_Load(dxEngine, TexturePath$)
    Call Textures.List_Update
    Me.Enabled = True
End Sub

'
' ***********************
' Copy the mesh in memory
' ***********************
'
Private Sub MENU_Copy_Click()
    Dim i%
    If Wall3D.No_Face.MIN <> 0 Then
        Call FacesCopy
    End If
End Sub

'
' *********************
' Copy and cut the mesh
' *********************
'
Private Sub MENU_Cut_Click()
    Dim i%
    If Wall3D.No_Face.MIN <> 0 Then
        Call FacesCopy
        For i% = Wall3D.No_Face.MAX To Wall3D.No_Face.MIN + 1 Step -1
            Call Tools3D.Face_Delete(i%)
        Next i%
        Call Tools3D.Face_Delete(Wall3D.No_Face.MIN)
        Call Wall3D.Update
    End If
End Sub

'
' *************************
' Select a new video driver
' *************************
'
Private Sub MENU_Driver_Click(Index As Integer)
    Dim i%
    For i% = 0 To 3
        If i% = Index Then
            MENU_Driver(i%).Checked = True
        Else
            MENU_Driver(i%).Checked = False
        End If
    Next i%
    dxEngine.VideoDriver% = Index
    Call dxEngine.Create_3DRM(View.shot, View.shot.ScaleWidth, View.shot.ScaleHeight, Mode3DSurface)
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' **********************
' Export complete X file
' **********************
'
Private Sub MENU_Export_X_Click()
    Dim File$
    Dim Axe As Boolean
    File$ = Open_Box$("Save X file", ExportX$, "DirectX X file|*.x", BOX_SAVE, Wall3D.BoiteMur3D)
    If File$ = "" Then Exit Sub
    ExportX$ = File$
    Dim TheMeshBuilder As Direct3DRMMeshBuilder3
    If Axe_Camera.ShowAxes.Value = vbChecked Then
        Axe = True
        Axe_Camera.ShowAxes.Value = vbUnchecked
    End If
    Call Wall3D.ReDraw(0, False)
    Set TheMeshBuilder = dxEngine.dxD3Drm.CreateMeshBuilder
    Call TheMeshBuilder.AddFrame(dxEngine.dxScene)
    Call TheMeshBuilder.Save(ExportX$, D3DRMXOF_TEXT, D3DRMXOFSAVE_ALL + D3DRMXOFSAVE_TEXTURETOPOLOGY) '+ D3DRMXOFSAVE_TEMPLATES ??? is it usefull?
    Call Wall3D.ReDraw(1, False)
    If Axe = True Then
        Axe_Camera.ShowAxes.Value = vbChecked
    End If
    Set TheMeshBuilder = Nothing
End Sub

'
' ********************************
' Create a complex mesh with faces
' ********************************
'
Private Sub MENU_Faces_Click()
    Call Faces.Show(vbModal)
End Sub

'
' ******************
' Load project faces
' Only for test
' ******************
'
Private Sub MENU_Import_Wall_Click()
    Dim File$, i%
    Dim TMeshBuilder() As Direct3DRMMeshBuilder3
    '
    File$ = Open_Box$("Load project", ImportWall$, "All files *.wall|*.wall", BOX_LOAD, Wall3D.BoiteMur3D)
    If Exist(File$) <> True Then Exit Sub
    ImportWall$ = File$
    '
    Call Tools3D.Build_MeshBuilder(dxEngine, ImportWall$, "", TMeshBuilder(), dxTexture(), Texture())
    Call Textures.List_Update
    '
    For i% = 0 To UBound(TMeshBuilder())
        Set DxWall(i% + 1).WallMeshBuilder = TMeshBuilder(i%)
        DefWall(i% + 1).Faces = True
        DefWall(i% + 1).TheFile$ = "For test only" + Str$(i% + 1)
        Set TMeshBuilder(i%) = Nothing
    Next i%
    '
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
    Call Wall3D.UpdateValue
End Sub

'
' *****************************
' Import a meshbuilder to array
' *****************************
'
Private Sub MENU_Import_MeshBuilder_Click()
    Dim File$, i%, j%, n%, n2%
    Dim TempMeshBuilder As Direct3DRMMeshBuilder3
    Dim TempFace As Direct3DRMFace2
    Dim TempVert() As D3DVECTOR
    Dim TempNorm As D3DVECTOR
    Dim TempColor As PALETTEENTRY, FaceColor As Long
    Dim TableColor() As Long, NColor%, TempFrame%
    '
    NColor% = 0
    ReDim TableColor(NColor%)
    File$ = Open_Box$("Import MeshBuilder", ImportX$, "All files *.x|*.x", BOX_LOAD, Wall3D.BoiteMur3D)
    If Exist(File$) <> True Then Exit Sub
    ImportX$ = File$
    Set TempMeshBuilder = dxEngine.dxD3Drm.CreateMeshBuilder
    Set TempMeshBuilder = dxEngine.Load_MeshBuilder(ImportX$)
    '
    n% = TempMeshBuilder.GetFaceCount
    For i% = 0 To n% - 1
        Set TempFace = TempMeshBuilder.GetFace(i%)
        n2% = TempFace.GetVertexCount
        ReDim TempVert(n2%) As D3DVECTOR
        For j% = 0 To n2% - 1
            Call TempFace.GetVertex(j%, TempVert(j%), TempNorm)
        Next j%
        '
        ' ***** Search if this color is already use in a frame
        '
        FaceColor = TempFace.GetColor
        TempFrame% = 0
        For j% = 1 To NColor%
            If TableColor(j%) = FaceColor Then TempFrame% = j%
        Next j%
        If TempFrame% = 0 Then
            NColor% = NColor% + 1
            ReDim Preserve TableColor(NColor%)
            TableColor(NColor%) = FaceColor
            TempFrame% = NColor%
        End If
        TempColor.Red = dxEngine.dX7.ColorGetRed(FaceColor) * 255
        TempColor.Green = dxEngine.dX7.ColorGetGreen(FaceColor) * 255
        TempColor.Blue = dxEngine.dX7.ColorGetBlue(FaceColor) * 255
        TempColor.flags = dxEngine.dX7.ColorGetAlpha(FaceColor) * 255
        '
        ' ***** Put this face in a frame
        '
        Call Tools3D.Face_Add(TempFrame%, TempColor, 0, n2%, TempVert(), False)
    Next i%
    '
    Set TempMeshBuilder = Nothing
    For i% = 1 To NColor%
        Wall3D.No_Mur = i%
        Call Wall3D.Update
    Next i%
End Sub

'
' ****************************
' Load a x file saved as frame
' ****************************
'
Private Sub MENU_Load_Frame_Click()
    Dim File$, i%
    File$ = Open_Box$("Load .x frame", "", "All files *.x|*.x", BOX_LOAD, Wall3D.BoiteMur3D)
    If Exist(File$) <> True Then Exit Sub
    If (XFileLoadFrame Is Nothing) = False Then
        Call dxEngine.dxScene.DeleteChild(XFileLoadFrame)
        Set XFileLoadFrame = Nothing
    End If
    Set XFileLoadFrame = dxEngine.dxD3Drm.CreateFrame(dxEngine.dxScene)
    Call dxEngine.Load_Frame(XFileLoadFrame, File$)
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' ***************
' Erase the scene
' ***************
'
Private Sub MENU_New_Click()
    If MsgBox("Erase scene?", vbYesNo + vbQuestion + vbDefaultButton2, "Wall3D") = vbYes Then
        WallPath$ = ""
        Caption = "Wall3D"
        Unload FrmAnimation
        Call Wall3D.InitScene
        If (XFileLoadFrame Is Nothing) = False Then
            Call dxEngine.dxScene.DeleteChild(XFileLoadFrame)
            Set XFileLoadFrame = Nothing
        End If
        Call Wall3D.ReDraw(0, False)
        Call Wall3D.ReDraw(1, False)
    End If
End Sub

'
' *******************
' Open a project file
' *******************
'
Private Sub MENU_Open_Click()
    Dim File$, i%
    Dim n%
    File$ = Open_Box$("Load project", WallPath$, "All files *.wall|*.wall", BOX_LOAD, Wall3D.BoiteMur3D)
    If File$ = "" Then Exit Sub
    If Exist(File$) <> True Then Exit Sub
    Wall3D.No_Mur = 0
    WallPath$ = File$
    Caption = File$ & " :Wall3D"
    Call Load_Wall_Definition(WallPath$, FaceIndex%(), DefFace(), DefWall(), Animation&(), NbAnimation%, n%, Texture())
    '
    If n% <> 0 Then
        For i% = 1 To n%
            Call Tools3D.Textures_Update(dxEngine, dxTexture(i%), Texture(i%))
        Next i%
        Call Textures.List_Update
    End If
    '
    For i% = 1 To UBound(DefWall())
        If DefWall(i%).TheFile$ <> "" Then
            Set DxWall(i%).WallMeshBuilder = dxEngine.Load_MeshBuilder(DefWall(i%).TheFile$)
            DefWall(i%).Faces = True
        Else
            Call Build_Wall(dxEngine, i%, FaceIndex%(), DefFace(), DefWall(), DxWall(), dxTexture())
        End If
    Next i%
    '
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
    Call Wall3D.UpdateValue
End Sub

'
' ***************************
' Change to orthographic mode
' ***************************
'
Private Sub MENU_Orthographic_Click()
    If Orthographic = False Then
        Orthographic = True
        MENU_Orthographic.Checked = True
        dxEngine.dxViewport.SetProjection D3DRMPROJECT_ORTHOGRAPHIC
    Else
        Orthographic = False
        MENU_Orthographic.Checked = False
        dxEngine.dxViewport.SetProjection D3DRMPROJECT_PERSPECTIVE
        dxEngine.dxViewport.SetField 0.5
    End If
    Call Axe_Camera.Position
End Sub

'
' ****************************
' Paste the mesh in a new mesh
' ****************************
'
Private Sub MENU_Paste_Click()
    Dim i%
    For i% = 0 To UBound(FaceCopy())
        Call Tools3D.Face_Add(Wall3D.No_Mur, FaceCopy(i%).Color, FaceCopy(i%).Texture, FaceCopy(i%).NbPoint, FaceCopy(i%).Point(), 0)
    Next i%
    For i% = 0 To 11
        Animation&(Wall3D.No_Mur, i%, 0) = MeshCopyDPosition(i%)
    Next i%
    DefWall(Wall3D.No_Mur) = MeshCopy
    Call Wall3D.Update
    Call Wall3D.UpdateValue
End Sub

'
' **************
' End of program
' **************
'
Private Sub MENU_Quit_Click()
    Unload Me
End Sub

'
' ******************
' Save textures list
' ******************
'
Private Sub MENU_Save_Texture_Click()
    Call Save_Texture
End Sub

'
' **********************
' Save a x file as frame
' **********************
'
Private Sub MENU_Save_Frame_Click()
    Dim File$, i%
    File$ = Open_Box$("Save .x frame", "", "All files *.x|*.x", BOX_SAVE, Wall3D.BoiteMur3D)
    If File$ = "" Then Exit Sub
    Call Wall3D.ReDraw(0, False)
    Call dxEngine.dxScene.Save(File$, D3DRMXOF_TEXT, D3DRMXOFSAVE_NORMALS _
    + D3DRMXOFSAVE_TEXTURECOORDINATES + D3DRMXOFSAVE_TEXTURENAMES + _
    D3DRMXOFSAVE_TEMPLATES + D3DRMXOFSAVE_TEXTURETOPOLOGY)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' ********************
' Save project to file
' ********************
'
Private Sub MENU_Save_Click()
    Dim File$, i%, j%, n%
    Dim f%, a$
    Dim Anim&(NbWall%, 11, 0)
    File$ = Open_Box$("Save project", WallPath$, "All files *.wall|*.wall", BOX_SAVE, Wall3D.BoiteMur3D)
    If File$ = "" Then Exit Sub
    WallPath$ = File$
    Caption = File$ & " :Wall3D"
    f% = FreeFile()
    Open File$ For Binary Access Write As #f%
    n% = UBound(DefFace())
    Put #f%, , n%
    Put #f%, , FaceIndex%()
    Put #f%, , DefFace()
    Put #f%, , DefWall()
    Put #f%, , NbAnimation%
    Put #f%, , Animation&()
    '
    Put #f%, , NbTexture%
    Put #f%, , Texture()
    '
    Close #f%
End Sub

'
' *************************************************
' Take a screenshot of the view and save it to disk
' *************************************************
'
Private Sub MENU_Screenshot_Click()
    Dim File$
    File$ = Open_Box$("Save scene to a bitmap file", "", "All files *.bmp|*.bmp", BOX_SAVE, Wall3D.BoiteMur3D)
    If File$ = "" Then Exit Sub
    View.shot.AutoRedraw = True
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
    ' Copy to a tempory picture and save to disk
    ' I've found nothing better than this method
    View.shot.ScaleWidth = View.ScaleWidth
    View.shot.ScaleHeight = View.ScaleHeight
    View.shot.Width = View.ScaleWidth
    View.shot.Height = View.ScaleHeight
    Call BitBlt(View.shot.hdc, 0, 0, View.shot.ScaleWidth, View.shot.ScaleHeight, View.hdc, 0, 0, vbSrcCopy)
    View.shot.Refresh
    Call SavePicture(View.shot.Image, File$)
    View.shot.AutoRedraw = False
End Sub

'
' *******************
' Call the shape form
' *******************
'
Private Sub MENU_Shape_Click()
    Call Shape.Show(vbModal)
End Sub

'
' ************************************
' True:  Draw full frame with textures
' False: Draw only vertex for speed
' ************************************
'
Private Sub MENU_Solid_Click()
    If MENU_Solid.Checked = False Then
        Call dxEngine.dxDevice.SetQuality(QUALITE_NORMAL)
        MENU_Solid.Checked = True
    Else
        Call dxEngine.dxDevice.SetQuality(QUALITE_VISION)
        MENU_Solid.Checked = False
    End If
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' **************************************
' Change spot light attach to the camera
' **************************************
'
Private Sub MENU_Spot_Light_Click()
    Dim LightColor As Long
    LightColor = dxEngine.dxLight.GetColor
    Load ColorSelection
    ColorSelection.TheColor.Red = dxEngine.dX7.ColorGetRed(LightColor) * 255
    ColorSelection.TheColor.Green = dxEngine.dX7.ColorGetGreen(LightColor) * 255
    ColorSelection.TheColor.Blue = dxEngine.dX7.ColorGetBlue(LightColor) * 255
    ColorSelection.TheColor.ShowAlpha = False
    ColorSelection.Show vbModal
    Call dxEngine.dxLight.SetColorRGB(TheRed / 255, TheGreen / 255, TheBlue / 255)
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' *******************************
' Call the standard geometry form
' *******************************
'
Private Sub MENU_Standard_Click()
    Call Geometry.Show(vbModal)
End Sub

'
' *******************
' Mesh and faces copy
' *******************
'
Public Sub FacesCopy()
    Dim i%
    MENU_Paste.Enabled = True
    ReDim FaceCopy(Wall3D.No_Face.MAX - Wall3D.No_Face.MIN) As TypeMur3D
    For i% = Wall3D.No_Face.MIN To Wall3D.No_Face.MAX
        FaceCopy(i% - Wall3D.No_Face.MIN) = DefFace(i%)
    Next i%
    For i% = 0 To 11
        MeshCopyDPosition(i%) = Animation&(Wall3D.No_Mur, i%, 0)
    Next i%
    MeshCopy = DefWall(Wall3D.No_Mur)
End Sub

'
' ******************
' Save textures list
' ******************
'
Private Sub Save_Texture()
    Dim File$
    File$ = Open_Box$("Save textures", TexturePath$, "Texture list|*.tmur", BOX_SAVE, Wall3D.BoiteMur3D)
    If File$ <> "" Then
        TexturePath$ = File$
        Call Tools3D.Textures_Save(TexturePath$)
    End If
End Sub

