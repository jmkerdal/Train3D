VERSION 5.00
Begin VB.Form View 
   BackColor       =   &H00FFFFFF&
   Caption         =   "View"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   Begin VB.PictureBox shot 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim ValidXY As Boolean
Dim LastX As Integer
Dim LastY As Integer

'
' ************************
' Command direct to screen
' ************************
'
Private Sub shot_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Choice%
    Dim Sens%, Offset%
    If KeyCode = vbKeyDelete Then
        If Wall3D.No_Face.MIN <> 0 Then
            Call Tools3D.Face_Delete(Wall3D.No_Face.Value)
            Wall3D.Update
        End If
    Else
        Choice% = -1
        Sens% = 1
        If Transform.Option(0) = True Then Offset% = 0
        If Transform.Option(1) = True Then Offset% = 3
        If Transform.Option(2) = True Then Offset% = 6
        If Transform.Option(3) = True Then Offset% = 9
        Select Case KeyCode
        Case vbKeyRight, vbKeyNumpad6
            Choice% = Offset%
            Sens% = 1
        Case vbKeyLeft, vbKeyNumpad4
            Choice% = Offset%
            Sens% = -1
        Case vbKeyUp, vbKeyNumpad8
            Choice% = Offset% + 1
            Sens% = 1
        Case vbKeyDown, vbKeyNumpad2
            Choice% = Offset% + 1
            Sens% = -1
        Case vbKeyPageUp, vbKeyNumpad9
            Choice% = Offset% + 2
            Sens% = 1
        Case vbKeyPageDown, vbKeyNumpad3
            Choice% = Offset% + 2
            Sens% = -1
        End Select
        If Choice% <> -1 Then
            Animation&(Wall3D.No_Mur, Choice%, NoAnimation%) = Animation&(Wall3D.No_Mur, Choice%, NoAnimation%) + Transform.StepValue! * Sens%
            Call Wall3D.Modify_Attibute(Choice%, Animation&(Wall3D.No_Mur, Choice%, NoAnimation%) + Transform.StepValue! * Sens%)
        End If
    End If
End Sub

'
' ************************************
' What Frame and Face has been hit?
' Code based from Nigel Thompson book
' "3D graphics programming for W95"
' translate to VB from C++
' ************************************
'
Private Sub shot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        ValidXY = True
        LastX = X
        LastY = Y
    End If
    If Button <> vbLeftButton Then Exit Sub
    '
    Dim dxPickArray As Direct3DRMPickArray
    Dim dxFrameArray As Direct3DRMFrameArray
    Dim dxFrame As Direct3DRMFrame3
    Dim dxD3DRMPickDesc As D3DRMPICKDESC
    Dim dxFaces As Direct3DRMFaceArray
    Dim dxFace As Direct3DRMFace2
    Dim i%, j%
    Dim NbVertex%
    Dim dxNormal As D3DVECTOR ' Unused, just for call
    Dim dxVertex(3) As D3DVECTOR
    Dim dxScreen(3) As D3DRMVECTOR4D
    Dim dxWorld(3) As D3DVECTOR
    Dim ScreenPoint(3) As POINTAPI ' Polygon create from face
    Dim Polygone As Long ' Polygone object pointer
    Dim FrameHit% ' Frame hit by the mouse cursor
    Dim FaceHit%, ZHit!, ZMoyen! ' Used to find the right face hit
    '
    ' ***** Redraw vue to selecting the face
    '
    Call Wall3D.ReDraw(0, False)
    '
    ' ***** Search the frame and the face hit
    '
    Set dxPickArray = dxEngine.dxViewport.Pick(X, Y)
    If dxPickArray.GetSize <> 0 Then
        Set dxFrameArray = dxPickArray.GetPickFrame(0, dxD3DRMPickDesc)
        Set dxFrame = dxFrameArray.GetElement(dxFrameArray.GetSize - 1)
        FrameHit% = dxFrame.GetAppData ' Now we known witch frame was hit
        If FrameHit% <> 0 Then
            '
            ' ***** Search the face hit in this frame
            '
            FaceHit% = -1
            Set dxFaces = DxWall(FrameHit%).WallMeshBuilder.GetFaces
            For i% = 0 To dxFaces.GetSize - 1
                Set dxFace = dxFaces.GetElement(i%)
                NbVertex% = dxFace.GetVertexCount
                ZMoyen! = 0
                For j% = 0 To NbVertex% - 1
                    '
                    ' ***** Projecting the face to screen
                    '
                    Call dxFace.GetVertex(j%, dxVertex(j%), dxNormal)
                    Call dxFrame.Transform(dxWorld(j%), dxVertex(j%))
                    Call dxEngine.dxViewport.Transform(dxScreen(j%), dxWorld(j%))
                    If dxScreen(j%).w <> 0 Then
                        ScreenPoint(j%).X = dxScreen(j%).X / dxScreen(j%).w
                        ScreenPoint(j%).Y = dxScreen(j%).Y / dxScreen(j%).w
                        ZMoyen! = ZMoyen! + dxScreen(j%).z / dxScreen(j%).w
                    Else
                        ScreenPoint(j%).X = 0
                        ScreenPoint(j%).Y = 0
                    End If
                Next j%
                ZMoyen! = ZMoyen! / NbVertex%
                '
                ' ***** Search if the mouse is inside the polygone
                '
                Polygone = CreatePolygonRgn(ScreenPoint(0), NbVertex%, 2)
                If PtInRegion(Polygone, X!, Y!) = 1 Then
                    If FaceHit% = -1 Then
                        FaceHit% = i%
                        ZHit! = ZMoyen!
                    Else
                        If ZMoyen! < ZHit! Then
                            FaceHit% = i%
                            ZHit! = ZMoyen!
                        End If
                    End If
                End If
                Call DeleteObject(Polygone)
            Next i%
        End If
    End If
    '
    ' ***** Release vue and objects
    '
    Call Wall3D.ReDraw(1, False)
    Set dxFace = Nothing
    Set dxFaces = Nothing
    Set dxFrame = Nothing
    Set dxFrameArray = Nothing
    Set dxPickArray = Nothing
    '
    ' ***** Update Frame and Face hit
    '
    Call Wall3D.ReDraw(2, True) ' Lock the update to avoid multiple refresh
    Wall3D.No_Mur = FrameHit%
    If FaceHit% <> -1 Then
        If Wall3D.No_Face.MIN <> 0 And DefWall(FrameHit%).TheFile = "" Then
            Wall3D.No_Face = Wall3D.No_Face.MIN + FaceHit%
        End If
    End If
    Call Wall3D.ReDraw(2, True) ' Unlock refresh level 2
End Sub

'
' ************************
' Rotate camera with mouse
' ************************
'
Private Sub shot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ValidXY = False Then Exit Sub
    '
    Dim Delta_x As Single, Delta_y As Single
    Dim Delta_r As Single, Radius As Single, Denom As Single, Angle As Single
    ' rotation axis in camcoords, worldcoords, sframecoords
    Dim axisC As D3DVECTOR
    Dim WorldCoord As D3DVECTOR
    Dim Axis As D3DVECTOR
    Dim Base As D3DVECTOR
    Dim Origin As D3DVECTOR
    '
    Delta_x = X - LastX
    Delta_y = Y - LastY
    LastX = X
    LastY = Y
    If Button = vbRightButton Then
        Delta_r = Sqr(Delta_x ^ 2 + Delta_y ^ 2)
        Radius = 50
        Denom = Sqr(Radius ^ 2 + Delta_r ^ 2)
        If (Delta_r = 0 Or Denom = 0) Then Exit Sub
        Angle = (Delta_r / Denom)
        axisC.X = (-Delta_y / Delta_r)
        axisC.Y = (-Delta_x / Delta_r)
        axisC.z = 0
        dxEngine.dxCamera.Transform WorldCoord, axisC
        dxEngine.dxScene.InverseTransform Axis, WorldCoord
        dxEngine.dxCamera.Transform WorldCoord, Origin
        dxEngine.dxScene.InverseTransform Base, WorldCoord
        Axis.X = Axis.X - Base.X
        Axis.Y = Axis.Y - Base.Y
        Axis.z = Axis.z - Base.z
        'Call Axe_Camera.Orientation(-Axis.Y * 5, -Axis.z * 5, -Axis.X * 5)
        Call dxEngine.dxCamera.AddRotation(D3DRMCOMBINE_AFTER, Axis.X, Axis.Y, Axis.z, -Angle)
    ElseIf Button = vbRightButton + vbLeftButton Then
        Call Axe_Camera.Traveling(3, Delta_y)
    End If
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

Private Sub shot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ValidXY = False
End Sub

'
' **************
' Redraw the vue
' **************
'
Private Sub shot_Paint()
    Call Wall3D.ReDraw(0, False)
    Call Wall3D.ReDraw(1, False)
End Sub

'
' ****************
' Redim the screen
' ****************
'
Private Sub Form_Resize()
    If dxEngine.dxViewport Is Nothing Then Exit Sub
    Me.shot.Height = View.ScaleHeight
    Me.shot.Width = View.ScaleWidth
    Me.shot.ScaleHeight = View.ScaleHeight
    Me.shot.ScaleHeight = View.ScaleHeight
    Call dxEngine.Create_3DRM(Me.shot, Me.shot.ScaleWidth, Me.shot.ScaleHeight, Mode3DSurface)
    Call Axe_Camera.Position
End Sub

